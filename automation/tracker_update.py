#!/usr/bin/env python3
"""
tracker_update.py - safe, VBA-preserving updater for
AI_ML_USA_Job_Search_Master_Tracker.xlsm (runs on the user's Mac via the Claude desktop VM).

What it guarantees (each one is checked at run time; any failure aborts WITHOUT touching the workbook):
  1. Refuses to run while Excel has the workbook open (Excel's "~$..." lock file present).
  2. Refuses to run while another updater run holds automation/run.lock (no overlapping writers).
  3. Backs up the workbook to backups/ before every write.
  4. Dedups every candidate against BOTH sheets (TO APPLY and APPLIED JOBS) and against the
     persistent reported-jobs DB (automation/seen_jobs.json) using: normalized URL,
     company+title+city, and company+requisition ID.
  5. Adds only new jobs, to TO APPLY, with defaults: Not Applied / Resume No / Cover Letter No / Date Applied blank.
  6. Never edits an existing row: every pre-existing record's 27 cells must be byte-identical after the
     run (rows may only change position, as the VBA's own sort/move rules do).
  7. Keeps the VBA project: xl/vbaProject.bin SHA-256 must be identical before and after.
  8. Keeps exactly the two visible sheets, column layout, and hides (never deletes) column M.
  9. Writes to a scratch copy first, re-opens and verifies it, then copies it over the original.

Usage:
  python3 tracker_update.py --folder <tracker folder> --jobs <new_jobs.json> [--dry-run]
  python3 tracker_update.py --folder <tracker folder> --check     # read-only health check

Output: one JSON report on stdout.
"""
import argparse, datetime as dt, hashlib, json, os, re, shutil, sys, tempfile, zipfile
from urllib.parse import urlsplit, parse_qsl, urlencode, urlunsplit

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import tracker_common as T   # noqa: E402
import engine                # noqa: E402
from openpyxl import load_workbook  # noqa: E402

WB_NAME = "AI_ML_USA_Job_Search_Master_Tracker.xlsm"
LOCK_STALE_MIN = 45
DATE_FIELDS = ("Job Posting Date", "Date Applied", "Follow-up Date")
DATETIME_FIELDS = ("Date and Time Discovered",)


# ----------------------------------------------------------------------------- keys
NP_MARK = "◆ Nonprofit"


def _norm(x):
    s = re.sub(r"\s*◆\s*nonprofit\b", " ", str(x or ""), flags=re.I)   # nonprofit mark never affects dedup
    return " ".join(s.strip().lower().split())


def norm_url(u):
    if not u:
        return ""
    u = str(u).strip()
    try:
        p = urlsplit(u)
        keep = [(k, v) for k, v in parse_qsl(p.query) if k.lower() in ("gh_jid", "jobid", "job_id", "id")]
        path = re.sub(r"/(application|apply)/?$", "", p.path.rstrip("/"))
        return urlunsplit((p.scheme.lower(), p.netloc.lower(), path.lower(), urlencode(keep), ""))
    except Exception:
        return _norm(u)


def keys_for(company, title, city, link, req=None):
    ks = set()
    if link:
        ks.add("url:" + norm_url(link))
    if company and title:
        ks.add("ctc:" + "|".join([_norm(company), _norm(title), _norm(city)]))
    if company and req:
        ks.add("req:" + _norm(company) + "|" + _norm(req))
    return ks


# ----------------------------------------------------------------------------- helpers
def sha_vba(path):
    with zipfile.ZipFile(path) as z:
        names = [n for n in z.namelist() if n.endswith("vbaProject.bin")]
        if not names:
            return None
        return hashlib.sha256(z.read(names[0])).hexdigest()


def excel_lock_present(folder):
    return [f for f in os.listdir(folder) if f.startswith("~$")]


def read_json(path, default):
    try:
        with open(path) as f:
            return json.load(f)
    except Exception:
        return default


def write_json(path, obj):
    with open(path, "w") as f:   # overwrite in place (no delete needed)
        json.dump(obj, f, indent=2, default=str)


def acquire_lock(auto_dir, now):
    p = os.path.join(auto_dir, "run.lock")
    cur = read_json(p, {"state": "FREE"})
    if cur.get("state") == "LOCKED":
        try:
            since = dt.datetime.fromisoformat(cur.get("since"))
            if (now - since).total_seconds() < LOCK_STALE_MIN * 60:
                return False, cur
        except Exception:
            pass
    write_json(p, {"state": "LOCKED", "since": now.isoformat(), "pid": os.getpid()})
    return True, cur


def release_lock(auto_dir):
    write_json(os.path.join(auto_dir, "run.lock"), {"state": "FREE", "since": dt.datetime.now().isoformat()})


def snapshot(wb):
    """All existing records keyed by Record Key -> tuple of the 27 cell values."""
    snap = {}
    for sheet in (T.SHEET_TOAPPLY, T.SHEET_APPLIED):
        for rec in engine.read_rows(wb[sheet]):
            snap[str(rec["vals"][T.COL_KEY - 1])] = (sheet, tuple(rec["vals"]), rec["link"])
    return snap


def workbook_keys(wb):
    ks = set()
    for sheet in (T.SHEET_TOAPPLY, T.SHEET_APPLIED):
        for rec in engine.read_rows(wb[sheet]):
            v = rec["vals"]
            link = rec["link"] or v[T.COL_LINK - 1]
            ks |= keys_for(v[T.COL_COMPANY - 1], v[T.COL_TITLE - 1], v[T.COL_CITY - 1], link)
    return ks


def coerce(job):
    out = dict(job)
    for f in DATE_FIELDS:
        v = out.get(f)
        if isinstance(v, str) and re.fullmatch(r"\d{4}-\d{2}-\d{2}", v):
            out[f] = dt.date.fromisoformat(v)
    for f in DATETIME_FIELDS:
        v = out.get(f)
        if isinstance(v, str) and v:
            try:
                out[f] = dt.datetime.fromisoformat(v)
            except ValueError:
                pass
    return out


def row_from_job(job, now):
    vals = [None] * T.NCOLS
    for header, i in T.COL.items():
        if header in job and job[header] is not None:
            vals[i - 1] = job[header]
    vals[T.COL_STATUS - 1] = "Not Applied"
    vals[T.COL_RESUME_UP - 1] = "No"
    vals[T.COL_CL_UP - 1] = "No"
    vals[T.COL_DATE_APPLIED - 1] = None
    for manual in ("Resume Version Used", "Cover Letter Version", "Recruiter / Contact Information",
                   "Follow-up Date", "Personal Notes"):
        vals[T.COL[manual] - 1] = None          # manual fields always start empty
    if not vals[T.COL_DISCOVERED - 1]:
        vals[T.COL_DISCOVERED - 1] = now.replace(microsecond=0)
    if not vals[T.COL["Job Availability"] - 1]:
        vals[T.COL["Job Availability"] - 1] = "Open"
    vals[T.COL_KEY - 1] = T.record_key(job.get("Company Name"), job.get("Job Title"),
                                       job.get("City / State"), job.get("Direct Application Link"))
    if job.get("_nonprofit") and vals[T.COL_COMPANY - 1] and NP_MARK not in str(vals[T.COL_COMPANY - 1]):
        vals[T.COL_COMPANY - 1] = f"{vals[T.COL_COMPANY - 1]} {NP_MARK}"   # visible nonprofit mark
    return {"vals": vals, "link": job.get("Direct Application Link")}


# ----------------------------------------------------------------------------- main
def run(folder, jobs_path=None, dry_run=False, check_only=False):
    now = dt.datetime.now()
    wb_path = os.path.join(folder, WB_NAME)
    auto_dir = os.path.join(folder, "automation")
    os.makedirs(auto_dir, exist_ok=True)
    rep = {"time": now.isoformat(timespec="seconds"), "workbook": wb_path, "dry_run": dry_run}

    if not os.path.exists(wb_path):
        rep.update(ok=False, error="workbook not found"); return rep
    locks = excel_lock_present(folder)
    if locks:
        rep.update(ok=False, error="Excel has the workbook open (lock file present); nothing written",
                   lock_files=locks); return rep

    vba_before = sha_vba(wb_path)
    wb = load_workbook(wb_path, keep_vba=True)
    rep["sheets"] = [(ws.title, ws.sheet_state) for ws in wb.worksheets]
    rep["vba_sha256"] = vba_before
    before = snapshot(wb)
    rep["existing_records"] = len(before)
    rep["existing_by_sheet"] = {s: sum(1 for v in before.values() if v[0] == s)
                                for s in (T.SHEET_TOAPPLY, T.SHEET_APPLIED)}
    if check_only:
        rep["ok"] = True; return rep

    got, cur = acquire_lock(auto_dir, now)
    if not got:
        rep.update(ok=False, error="another updater run holds the lock", lock=cur); return rep
    try:
        seen_path = os.path.join(auto_dir, "seen_jobs.json")
        seen = read_json(seen_path, {"keys": {}, "jobs": []})
        known = set(seen["keys"].keys()) | workbook_keys(wb)

        jobs = [coerce(j) for j in read_json(jobs_path, [])]
        new_recs, skipped, added_meta = [], [], []
        for j in jobs:
            ks = keys_for(j.get("Company Name"), j.get("Job Title"), j.get("City / State"),
                          j.get("Direct Application Link"), j.get("_req"))
            hit = ks & known
            if hit:
                skipped.append({"company": j.get("Company Name"), "title": j.get("Job Title"),
                                "matched": sorted(hit)[0]})
                continue
            known |= ks
            new_recs.append(row_from_job(j, now))
            added_meta.append((ks, j))
        rep["candidates"] = len(jobs)
        rep["added"] = len(new_recs)
        rep["skipped_duplicate"] = skipped

        if dry_run or not new_recs:
            rep["ok"] = True
            rep["written"] = False
            if not new_recs:
                rep["note"] = "no new jobs; workbook not modified"
            return rep

        # backup first
        bdir = os.path.join(folder, "backups")
        os.makedirs(bdir, exist_ok=True)
        bpath = os.path.join(bdir, f"tracker_{now.strftime('%Y%m%d_%H%M%S')}.xlsm")
        shutil.copy2(wb_path, bpath)
        rep["backup"] = bpath

        ws_p = wb[T.SHEET_TOAPPLY]
        rows = engine.read_rows(ws_p) + new_recs
        engine.write_rows(ws_p, rows)
        engine.reconcile_and_sort(wb)
        for sheet in (T.SHEET_TOAPPLY, T.SHEET_APPLIED):
            wb[sheet].column_dimensions["M"].hidden = True

        scratch = tempfile.mkdtemp(prefix="tracker_")
        tmp = os.path.join(scratch, WB_NAME)
        wb.save(tmp)

        # ---- verify the scratch copy before it goes anywhere near the original
        chk = load_workbook(tmp, keep_vba=True)
        vis = [ws.title for ws in chk.worksheets if ws.sheet_state == "visible"]
        assert vis == [T.SHEET_TOAPPLY, T.SHEET_APPLIED], f"visible sheets changed: {vis}"
        for sheet in (T.SHEET_TOAPPLY, T.SHEET_APPLIED):
            hdr = [chk[sheet].cell(T.HEADER_ROW, c).value for c in range(1, T.NCOLS + 1)]
            assert hdr == [c[0] for c in T.COLUMNS], f"{sheet} header/column positions changed"
        after = snapshot(chk)
        for k, (sheet, vals, link) in before.items():
            assert k in after, f"existing record lost: {k}"
            assert after[k][1] == vals, f"existing record modified: {k}"
            assert (after[k][2] or None) == (link or None), f"hyperlink changed: {k}"
        assert len(after) == len(before) + len(new_recs), "record count mismatch"
        vba_after = sha_vba(tmp)
        assert vba_after == vba_before, "VBA project changed or lost"
        rep["vba_preserved"] = True
        rep["existing_records_unchanged"] = len(before)

        shutil.copyfile(tmp, wb_path)                     # overwrite in place
        assert sha_vba(wb_path) == vba_before, "VBA check failed after copy"
        rep["written"] = True
        rep["total_records"] = len(after)
        rep["by_sheet"] = {s: sum(1 for v in after.values() if v[0] == s)
                           for s in (T.SHEET_TOAPPLY, T.SHEET_APPLIED)}

        # remember what was reported (persistent DB)
        for ks, j in added_meta:
            for k in ks:
                seen["keys"][k] = now.isoformat(timespec="seconds")
            seen["jobs"].append({"added": now.isoformat(timespec="seconds"),
                                 "company": j.get("Company Name"), "title": j.get("Job Title"),
                                 "city": j.get("City / State"), "link": j.get("Direct Application Link"),
                                 "posted": str(j.get("Job Posting Date")), "priority": j.get("Priority"),
                                 "score": j.get("Resume Match Score")})
        write_json(seen_path, seen)
        rep["ok"] = True
        return rep
    except AssertionError as e:
        rep.update(ok=False, written=False, error=f"verification failed, original untouched: {e}")
        return rep
    finally:
        release_lock(auto_dir)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--folder", required=True)
    ap.add_argument("--jobs")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--check", action="store_true")
    a = ap.parse_args()
    print(json.dumps(run(a.folder, a.jobs, a.dry_run, a.check), indent=2, default=str))
