#!/usr/bin/env python3
"""
maintain_tracker.py - small, verified maintenance edits to the master tracker.
It never adds, deletes or moves a record, and never touches a row you have started working on.

  python3 maintain_tracker.py --folder F --rerank FILE [--dry-run]
      FILE = [{"record_key" or "link": "...", "priority": "P3", "score": 64,
               "note": "text prepended to Missing Skills / Qualifications", "reason": "why"}]
      Changes ONLY Priority (A), Resume Match Score (B) and Missing Skills / Qualifications (Q), and only on
      TO APPLY rows that are untouched by you: status Not Applied / Saved / blank, Resume Uploaded and
      Cover Letter Uploaded = No or blank, and Resume Version, Cover Letter Version, Date Applied,
      Recruiter, Follow-up Date and Personal Notes all empty. Any other row is refused and left as is.

  python3 maintain_tracker.py --folder F --nonprofit-format [--dry-run]
      Adds the nonprofit highlight on both sheets (Company Name containing the nonprofit mark -> purple)
      and a one-line legend in row 3. Idempotent: does nothing if already present.

Safety (same guarantees as tracker_update.py; any failure aborts WITHOUT touching the workbook):
  refuses while Excel has the workbook open or another run holds automation/run.lock; backs up first;
  saves to a scratch copy and verifies there: exactly the 2 visible sheets, unchanged headers, every
  record identical except the allowed cells of the requested rows, same record count, same VBA hash;
  only then copies over the original and re-checks the VBA hash. Every change is logged in
  automation/maintenance_log.json.
"""
import argparse, datetime as dt, json, os, shutil, sys, tempfile
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, HERE)
import tracker_common as T, engine, tracker_update as U   # noqa
from openpyxl import load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font, PatternFill

NP_MARK = "◆ Nonprofit"
NP_LEGEND = ("◆ Nonprofit = nonprofit employer (purple). Universities, their affiliated nonprofits and nonprofit "
             "research institutes may be H-1B cap-exempt - confirm with HR.")
LOG = "maintenance_log.json"
PENDING_OK = {"Not Applied", "Saved", ""}
EDITABLE = (T.COL_PRIORITY, T.COL_SCORE, T.COL["Missing Skills / Qualifications"])
MANUAL_EMPTY = ("Resume Version Used", "Cover Letter Version", "Date Applied",
                "Recruiter / Contact Information", "Follow-up Date", "Personal Notes")


def untouched(vals):
    st = str(vals[T.COL_STATUS - 1] or "").strip()
    if st not in PENDING_OK:
        return False, f"status is '{st}'"
    for col in (T.COL_RESUME_UP, T.COL_CL_UP):
        if str(vals[col - 1] or "No").strip() != "No":
            return False, f"{T.COLUMNS[col - 1][0]} = {vals[col - 1]}"
    for name in MANUAL_EMPTY:
        v = vals[T.COL[name] - 1]
        if v not in (None, ""):
            return False, f"'{name}' has your entry"
    return True, ""


def has_np_rule(ws):
    return any("Nonprofit" in "".join(r.formula or []) for cf in ws.conditional_formatting for r in cf.rules)


def add_np_format(ws):
    rng = f"D{T.FIRST_DATA_ROW}:D{T.LAST_DATA_ROW}"
    rule = FormulaRule(formula=[f'ISNUMBER(SEARCH("{NP_MARK}",$D{T.FIRST_DATA_ROW}))'],
                       fill=PatternFill("solid", start_color="FFE9D5FF", end_color="FFE9D5FF"),
                       font=Font(bold=True, color="FF5B21B6"))
    rule.priority = 1   # above the P1/P2/P3 row fills (priorities 2-4), so the purple wins on column D
    ws.conditional_formatting.add(rng, rule)
    c = ws.cell(T.SPACER_ROW, 1)
    if c.value in (None, ""):
        c.value = NP_LEGEND
        c.font = Font(name=T.FONT_NAME, size=9, italic=True, color="FF5B21B6")
        ws.row_dimensions[T.SPACER_ROW].height = 15   # the spacer row is thin; make the legend readable


def run(folder, rerank_file=None, np_format=False, dry=False):
    now = dt.datetime.now(); auto = os.path.join(folder, "automation")
    wbp = os.path.join(folder, U.WB_NAME)
    rep = {"time": now.isoformat(timespec="seconds"), "dry_run": dry}
    if U.excel_lock_present(folder):
        rep.update(ok=False, written=False, error="Excel has the workbook open; nothing changed"); return rep
    vba = U.sha_vba(wbp)
    wb = load_workbook(wbp, keep_vba=True)
    before = U.snapshot(wb)
    changes, refused, not_found, expected = [], [], [], {}

    if rerank_file:
        req = U.read_json(rerank_file, [])
        rows = engine.read_rows(wb[T.SHEET_TOAPPLY])
        for item in req:
            hit = None
            for rec in rows:
                key = str(rec["vals"][T.COL_KEY - 1])
                link = rec["link"] or rec["vals"][T.COL_LINK - 1]
                if (item.get("record_key") and key == item["record_key"]) or \
                   (item.get("link") and U.norm_url(link) == U.norm_url(item["link"])):
                    hit = rec; break
            if not hit:
                not_found.append(item.get("record_key") or item.get("link")); continue
            v = hit["vals"]; key = str(v[T.COL_KEY - 1])
            ok, why = untouched(v)
            if not ok:
                refused.append({"company": v[3], "title": v[4], "why": why}); continue
            old = (v[T.COL_PRIORITY - 1], v[T.COL_SCORE - 1])
            q = T.COL["Missing Skills / Qualifications"] - 1
            if item.get("note") and not str(v[q] or "").startswith(item["note"]):
                v[q] = (item["note"] + " " + str(v[q] or "")).strip()
            v[T.COL_PRIORITY - 1] = item["priority"]; v[T.COL_SCORE - 1] = int(item["score"])
            expected[key] = tuple(v)
            changes.append({"record_key": key, "company": v[3], "title": v[4], "from": list(old),
                            "to": [item["priority"], int(item["score"])], "reason": item.get("reason", "")})
        if changes:
            engine.write_rows(wb[T.SHEET_TOAPPLY], rows)

    np_added = []
    if np_format:
        for s in (T.SHEET_TOAPPLY, T.SHEET_APPLIED):
            if not has_np_rule(wb[s]):
                add_np_format(wb[s]); np_added.append(s)

    rep.update(rerank=changes, refused=refused, not_found=not_found, nonprofit_format_added=np_added)
    if dry or not (changes or np_added):
        rep.update(ok=True, written=False); return rep

    got, cur = U.acquire_lock(auto, now)
    if not got:
        rep.update(ok=False, written=False, error="another run holds the lock", lock=cur); return rep
    try:
        bdir = os.path.join(folder, "backups"); os.makedirs(bdir, exist_ok=True)
        bp = os.path.join(bdir, f"tracker_{now.strftime('%Y%m%d_%H%M%S')}_before_maintenance.xlsm")
        shutil.copy2(wbp, bp); rep["backup"] = bp
        engine.reconcile_and_sort(wb)
        for s in (T.SHEET_TOAPPLY, T.SHEET_APPLIED):
            wb[s].column_dimensions["M"].hidden = True
        tmp = os.path.join(tempfile.mkdtemp(prefix="maint_"), U.WB_NAME); wb.save(tmp)
        chk = load_workbook(tmp, keep_vba=True)
        vis = [w.title for w in chk.worksheets if w.sheet_state == "visible"]
        assert vis == [T.SHEET_TOAPPLY, T.SHEET_APPLIED], f"sheets changed: {vis}"
        for s in (T.SHEET_TOAPPLY, T.SHEET_APPLIED):
            assert [chk[s].cell(T.HEADER_ROW, c).value for c in range(1, T.NCOLS + 1)] == [c[0] for c in T.COLUMNS]
            assert chk[s].column_dimensions["M"].hidden
            if np_format:
                assert has_np_rule(chk[s]), f"nonprofit format missing on {s}"
        after = U.snapshot(chk)
        assert len(after) == len(before), "record count changed"
        for k, (s, vals, link) in before.items():
            assert k in after, f"record lost: {k}"
            assert (after[k][2] or None) == (link or None), f"hyperlink changed: {k}"
            if k in expected:
                assert after[k][1] == expected[k], f"rerank not applied as requested: {k}"
                diff = [i + 1 for i, (a, b) in enumerate(zip(vals, after[k][1])) if a != b]
                assert set(diff) <= set(EDITABLE), f"cells outside A/B/Q changed on {k}: {diff}"
            else:
                assert after[k][1] == vals, f"record changed: {k}"
                assert after[k][0] == s, f"record moved sheets: {k}"
        assert U.sha_vba(tmp) == vba, "VBA changed"
        shutil.copyfile(tmp, wbp)
        assert U.sha_vba(wbp) == vba, "VBA check failed after copy"
        log = U.read_json(os.path.join(auto, LOG), {"entries": []})
        log["entries"].append({"time": now.isoformat(timespec="seconds"), "backup": os.path.basename(bp),
                               "rerank": changes, "nonprofit_format_added": np_added})
        U.write_json(os.path.join(auto, LOG), log)
        rep.update(ok=True, written=True, vba_preserved=True, records=len(after))
        return rep
    except AssertionError as e:
        rep.update(ok=False, written=False, error=f"verification failed, original untouched: {e}"); return rep
    finally:
        U.release_lock(auto)


if __name__ == "__main__":
    ap = argparse.ArgumentParser(); ap.add_argument("--folder", required=True)
    ap.add_argument("--rerank"); ap.add_argument("--nonprofit-format", action="store_true")
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()
    print(json.dumps(run(a.folder, a.rerank, a.nonprofit_format, a.dry_run), indent=2, default=str))
