#!/usr/bin/env python3
"""
exclude_jobs.py - move clearly ineligible jobs OUT of the TO APPLY sheet and INTO the exclusion log.

Safety (same guarantees as tracker_update.py; any failure aborts WITHOUT touching the workbook):
  * refuses while Excel has the workbook open, or while another run holds automation/run.lock
  * backs up the workbook to backups/ first
  * only removes TO APPLY rows whose status is still a pending status (never touches APPLIED JOBS,
    never touches a row you've started working on: In Progress / Ready to Apply rows are refused)
  * every other record must be byte-identical afterwards; VBA project SHA-256 must be unchanged
  * the removed rows (all 27 columns + link) are saved in automation/exclusions.json and the
    readable Exclusion_Log.xlsx BEFORE the workbook is replaced; their dedup keys stay in
    seen_jobs.json, so later runs never re-add them

Usage:
  python3 exclude_jobs.py --folder F --file exclude.json [--dry-run]
      exclude.json = [{"record_key": "<col AA value>", "reason": "<why ineligible>", "rule": "<rule name>"}]
  python3 exclude_jobs.py --folder F --log-candidates cands.json
      log jobs found by a run but never imported (cands.json = list of job dicts with workbook column names
      + "_exclusion_reason" + "_rule"); workbook untouched
"""
import argparse, datetime as dt, json, os, shutil, sys, tempfile
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, HERE)
import tracker_common as T, engine, tracker_update as U   # noqa
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment

LOG_JSON = "exclusions.json"
LOG_XLSX = "Exclusion_Log.xlsx"
SAFE_TO_REMOVE = {"Not Applied", "Saved", ""}


def load_log(auto):
    return U.read_json(os.path.join(auto, LOG_JSON), {"entries": []})


def save_log(folder, auto, log):
    U.write_json(os.path.join(auto, LOG_JSON), log)
    wb = Workbook(); ws = wb.active; ws.title = "EXCLUSION LOG"
    hdr = ["Date Excluded", "Rule", "Reason Excluded", "Source"] + [c[0] for c in T.COLUMNS[:-1]]
    ws.append(hdr)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFFFF", name="Arial"); c.fill = PatternFill("solid", fgColor="FF7F1D1D")
        c.alignment = Alignment(wrap_text=True, vertical="top")
    for e in log["entries"]:
        rec = e["record"]
        ws.append([e["excluded_at"], e.get("rule", ""), e["reason"], e.get("source", "")] +
                  [rec.get(c[0]) for c in T.COLUMNS[:-1]])
    widths = [18, 20, 50, 16, 8, 8, 14, 24, 32, 22, 16, 20, 16, 30, 30, 14, 20, 40, 16, 30, 30]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[ws.cell(1, i).column_letter].width = w
    ws.freeze_panes = "A2"
    wb.save(os.path.join(folder, LOG_XLSX))


def rec_dict(vals, link):
    d = {c[0]: (v.isoformat() if hasattr(v, "isoformat") else v) for c, v in zip(T.COLUMNS, vals)}
    if link: d["Direct Application Link"] = link
    return d


def log_candidates(folder, path):
    auto = os.path.join(folder, "automation"); log = load_log(auto)
    now = dt.datetime.now().isoformat(timespec="seconds")
    have = {e["record"].get("Direct Application Link") for e in log["entries"]}
    n = 0
    for j in U.read_json(path, []):
        if j.get("Direct Application Link") in have: continue
        log["entries"].append({"excluded_at": now, "rule": j.get("_rule", ""), "reason": j.get("_exclusion_reason", ""),
                               "source": "search candidate (never imported)",
                               "record": {k: v for k, v in j.items() if not k.startswith("_")}})
        n += 1
    save_log(folder, auto, log)
    return {"ok": True, "logged": n, "total_log_entries": len(log["entries"])}


def run(folder, path, dry):
    now = dt.datetime.now(); auto = os.path.join(folder, "automation")
    wbp = os.path.join(folder, U.WB_NAME)
    rep = {"time": now.isoformat(timespec="seconds"), "dry_run": dry}
    if U.excel_lock_present(folder):
        rep.update(ok=False, error="Excel has the workbook open; nothing changed"); return rep
    req = {r["record_key"]: r for r in U.read_json(path, [])}
    vba = U.sha_vba(wbp)
    wb = load_workbook(wbp, keep_vba=True)
    before = U.snapshot(wb)
    rows = engine.read_rows(wb[T.SHEET_TOAPPLY])
    keep, out = [], []
    for rec in rows:
        k = str(rec["vals"][T.COL_KEY - 1])
        if k in req:
            st = str(rec["vals"][T.COL_STATUS - 1] or "").strip()
            if st not in SAFE_TO_REMOVE:
                rep.update(ok=False, error=f"refusing: '{k}' has status '{st}' (you are working on it)"); return rep
            out.append(rec)
        else:
            keep.append(rec)
    missing = sorted(set(req) - {str(r["vals"][T.COL_KEY - 1]) for r in out})
    rep["to_remove"] = [{"company": r["vals"][3], "title": r["vals"][4], "reason": req[str(r["vals"][T.COL_KEY-1])]["reason"]} for r in out]
    rep["not_found"] = missing
    if dry or not out:
        rep.update(ok=True, written=False); return rep
    got, cur = U.acquire_lock(auto, now)
    if not got:
        rep.update(ok=False, error="another run holds the lock", lock=cur); return rep
    try:
        bdir = os.path.join(folder, "backups"); os.makedirs(bdir, exist_ok=True)
        bp = os.path.join(bdir, f"tracker_{now.strftime('%Y%m%d_%H%M%S')}_before_exclusion.xlsm")
        shutil.copy2(wbp, bp); rep["backup"] = bp
        engine.write_rows(wb[T.SHEET_TOAPPLY], keep)
        engine.reconcile_and_sort(wb)
        for s in (T.SHEET_TOAPPLY, T.SHEET_APPLIED):
            wb[s].column_dimensions["M"].hidden = True
        tmp = os.path.join(tempfile.mkdtemp(prefix="excl_"), U.WB_NAME); wb.save(tmp)
        chk = load_workbook(tmp, keep_vba=True)
        vis = [w.title for w in chk.worksheets if w.sheet_state == "visible"]
        assert vis == [T.SHEET_TOAPPLY, T.SHEET_APPLIED], f"sheets changed: {vis}"
        for s in (T.SHEET_TOAPPLY, T.SHEET_APPLIED):
            assert [chk[s].cell(T.HEADER_ROW, c).value for c in range(1, T.NCOLS + 1)] == [c[0] for c in T.COLUMNS]
        after = U.snapshot(chk)
        removed = {str(r["vals"][T.COL_KEY - 1]) for r in out}
        for k, (s, vals, link) in before.items():
            if k in removed:
                assert k not in after, f"not removed: {k}"; continue
            assert k in after and after[k][1] == vals and (after[k][2] or None) == (link or None), f"record changed: {k}"
        assert len(after) == len(before) - len(out)
        assert U.sha_vba(tmp) == vba, "VBA changed"
        # log FIRST, then replace the workbook
        log = load_log(auto)
        for r in out:
            k = str(r["vals"][T.COL_KEY - 1])
            log["entries"].append({"excluded_at": now.isoformat(timespec="seconds"), "rule": req[k].get("rule", ""),
                                   "reason": req[k]["reason"], "source": "removed from TO APPLY",
                                   "record": rec_dict(r["vals"], r["link"])})
        save_log(folder, auto, log)
        shutil.copyfile(tmp, wbp)
        assert U.sha_vba(wbp) == vba
        rep.update(ok=True, written=True, vba_preserved=True, removed=len(out),
                   remaining_to_apply=sum(1 for v in after.values() if v[0] == T.SHEET_TOAPPLY),
                   log_entries=len(log["entries"]))
        return rep
    except AssertionError as e:
        rep.update(ok=False, written=False, error=f"verification failed, original untouched: {e}"); return rep
    finally:
        U.release_lock(auto)


if __name__ == "__main__":
    ap = argparse.ArgumentParser(); ap.add_argument("--folder", required=True)
    ap.add_argument("--file"); ap.add_argument("--log-candidates"); ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()
    res = log_candidates(a.folder, a.log_candidates) if a.log_candidates else run(a.folder, a.file, a.dry_run)
    print(json.dumps(res, indent=2, default=str))
