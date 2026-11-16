#!/usr/bin/env python3
"""
record_application.py - add a job you applied to into APPLIED JOBS (status Applied), with the resume and
cover letter versions used. Same safety as tracker_update.py: refuses while Excel is open or the run lock is
held, backs up first, every existing row must stay byte-identical, VBA SHA-256 must be unchanged.
If the job is in the exclusion log, its saved record is reused and the log entry is annotated (never deleted).

  python3 record_application.py --folder F --file app.json [--dry-run]
  app.json = {"record_key": "<key from exclusions.json>" OR "job": {workbook column: value...},
              "set": {"Date Applied": "YYYY-MM-DD", "Resume Version Used": "...", ...}}
"""
import argparse, datetime as dt, json, os, shutil, sys, tempfile
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, HERE)
import tracker_common as T, engine, tracker_update as U   # noqa
from openpyxl import load_workbook
import exclude_jobs as X


def run(folder, path, dry):
    now = dt.datetime.now(); auto = os.path.join(folder, "automation"); wbp = os.path.join(folder, U.WB_NAME)
    rep = {"time": now.isoformat(timespec="seconds"), "dry_run": dry}
    if U.excel_lock_present(folder):
        rep.update(ok=False, error="Excel has the workbook open; close Excel and run again. Nothing changed."); return rep
    app = U.read_json(path, {})
    if app.get("existing_key"):
        return update_existing(folder, auto, wbp, app, now, rep, dry)
    log = X.load_log(auto); entry = None
    if app.get("record_key"):
        for e in log["entries"]:
            r = e["record"]
            if T.record_key(r.get("Company Name"), r.get("Job Title"), r.get("City / State"), r.get("Direct Application Link")) == app["record_key"]:
                entry = e
        if not entry:
            rep.update(ok=False, error="record_key not found in exclusion log"); return rep
        job = dict(entry["record"])
    else:
        job = dict(app["job"])
    job.update(app.get("set", {}))
    job["Application Status"] = "Applied"
    job = U.coerce(job)
    vals = [job.get(c[0]) for c in T.COLUMNS]
    vals[T.COL_KEY - 1] = T.record_key(job.get("Company Name"), job.get("Job Title"), job.get("City / State"), job.get("Direct Application Link"))
    newrec = {"vals": vals, "link": job.get("Direct Application Link")}
    vba = U.sha_vba(wbp); wb = load_workbook(wbp, keep_vba=True); before = U.snapshot(wb)
    if vals[T.COL_KEY - 1] in before:
        rep.update(ok=False, error="this job is already in the workbook"); return rep
    rep["adding"] = {"company": vals[3], "title": vals[4], "status": "Applied", "date_applied": str(job.get("Date Applied"))}
    if dry:
        rep.update(ok=True, written=False); return rep
    got, cur = U.acquire_lock(auto, now)
    if not got:
        rep.update(ok=False, error="another run holds the lock"); return rep
    try:
        bp = os.path.join(folder, "backups", f"tracker_{now.strftime('%Y%m%d_%H%M%S')}_before_application.xlsm")
        shutil.copy2(wbp, bp); rep["backup"] = bp
        ws = wb[T.SHEET_APPLIED]
        engine.write_rows(ws, engine.read_rows(ws) + [newrec])
        engine.reconcile_and_sort(wb)
        for s in (T.SHEET_TOAPPLY, T.SHEET_APPLIED):
            wb[s].column_dimensions["M"].hidden = True
        tmp = os.path.join(tempfile.mkdtemp(prefix="app_"), U.WB_NAME); wb.save(tmp)
        chk = load_workbook(tmp, keep_vba=True)
        assert [w.title for w in chk.worksheets if w.sheet_state == "visible"] == [T.SHEET_TOAPPLY, T.SHEET_APPLIED]
        after = U.snapshot(chk)
        for k, (s, v, l) in before.items():
            assert k in after and after[k][1] == v and (after[k][2] or None) == (l or None), f"record changed: {k}"
        assert after.get(vals[T.COL_KEY - 1], (None,))[0] == T.SHEET_APPLIED, "new row not on APPLIED JOBS"
        assert len(after) == len(before) + 1
        assert U.sha_vba(tmp) == vba, "VBA changed"
        shutil.copyfile(tmp, wbp); assert U.sha_vba(wbp) == vba
        if entry is not None:
            entry["later_action"] = f"Applied anyway on {job.get('Date Applied')} (moved to APPLIED JOBS {now.isoformat(timespec='minutes')})"
            X.save_log(folder, auto, log)
        rep.update(ok=True, written=True, vba_preserved=True,
                   by_sheet={s: sum(1 for x in after.values() if x[0] == s) for s in (T.SHEET_TOAPPLY, T.SHEET_APPLIED)})
        return rep
    except AssertionError as e:
        rep.update(ok=False, written=False, error=f"verification failed, original untouched: {e}"); return rep
    finally:
        U.release_lock(auto)


MANUAL = {"Application Status", "Resume Uploaded?", "Resume Version Used", "Cover Letter Uploaded?",
          "Cover Letter Version", "Date Applied", "Recruiter / Contact Information", "Follow-up Date", "Personal Notes"}


def update_existing(folder, auto, wbp, app, now, rep, dry):
    """Set your application fields on a job already in TO APPLY; the VBA-mirroring engine then moves it to APPLIED JOBS.
    Only the manual application fields may change; every other cell of that row and every other row must stay identical."""
    sets = dict(app.get("set", {})); sets["Application Status"] = "Applied"
    bad = set(sets) - MANUAL
    if bad:
        rep.update(ok=False, error=f"refusing to change non-application fields: {sorted(bad)}"); return rep
    sets = U.coerce(sets)
    vba = U.sha_vba(wbp); wb = load_workbook(wbp, keep_vba=True); before = U.snapshot(wb)
    key = app["existing_key"]
    if key not in before:
        rep.update(ok=False, error="job not found in workbook"); return rep
    if before[key][0] != T.SHEET_TOAPPLY:
        rep.update(ok=False, error="job is not on TO APPLY (already applied?)"); return rep
    ws = wb[T.SHEET_TOAPPLY]; rows = engine.read_rows(ws)
    for rec in rows:
        if str(rec["vals"][T.COL_KEY - 1]) == key:
            for f, v in sets.items():
                rec["vals"][T.COL[f] - 1] = v
            rep["adding"] = {"company": rec["vals"][3], "title": rec["vals"][4], "status": "Applied", "date_applied": str(sets.get("Date Applied"))}
    if dry:
        rep.update(ok=True, written=False); return rep
    got, cur = U.acquire_lock(auto, now)
    if not got:
        rep.update(ok=False, error="another run holds the lock"); return rep
    try:
        bp = os.path.join(folder, "backups", f"tracker_{now.strftime('%Y%m%d_%H%M%S')}_before_application.xlsm")
        shutil.copy2(wbp, bp); rep["backup"] = bp
        engine.write_rows(ws, rows); engine.reconcile_and_sort(wb)
        for s in (T.SHEET_TOAPPLY, T.SHEET_APPLIED):
            wb[s].column_dimensions["M"].hidden = True
        tmp = os.path.join(tempfile.mkdtemp(prefix="app_"), U.WB_NAME); wb.save(tmp)
        chk = load_workbook(tmp, keep_vba=True)
        assert [w.title for w in chk.worksheets if w.sheet_state == "visible"] == [T.SHEET_TOAPPLY, T.SHEET_APPLIED]
        after = U.snapshot(chk)
        assert len(after) == len(before)
        for k, (s, v, l) in before.items():
            assert k in after and (after[k][2] or None) == (l or None), f"record lost or link changed: {k}"
            if k == key:
                assert after[k][0] == T.SHEET_APPLIED, "job did not move to APPLIED JOBS"
                for i, (name, _w, _x) in enumerate(T.COLUMNS):
                    if name not in sets:
                        assert after[k][1][i] == v[i], f"non-application field changed: {name}"
            else:
                assert after[k][1] == v, f"record changed: {k}"
        assert U.sha_vba(tmp) == vba, "VBA changed"
        shutil.copyfile(tmp, wbp); assert U.sha_vba(wbp) == vba
        rep.update(ok=True, written=True, vba_preserved=True,
                   by_sheet={s: sum(1 for x in after.values() if x[0] == s) for s in (T.SHEET_TOAPPLY, T.SHEET_APPLIED)})
        return rep
    except AssertionError as e:
        rep.update(ok=False, written=False, error=f"verification failed, original untouched: {e}"); return rep
    finally:
        U.release_lock(auto)


if __name__ == "__main__":
    ap = argparse.ArgumentParser(); ap.add_argument("--folder", required=True); ap.add_argument("--file", required=True)
    ap.add_argument("--dry-run", action="store_true"); a = ap.parse_args()
    print(json.dumps(run(a.folder, a.file, a.dry_run), indent=2, default=str))
