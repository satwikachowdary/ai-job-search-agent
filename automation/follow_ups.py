#!/usr/bin/env python3
"""
follow_ups.py - READ-ONLY. Lists applications on APPLIED JOBS that need a follow-up today.
Due = status is Applied / Assessment / Interview AND (Follow-up Date is today or earlier,
      OR no Follow-up Date is set and it was applied 10+ days ago).
Never writes the workbook (safe even while Excel has it open: reads the saved file).

  python3 follow_ups.py --folder <tracker folder>      -> JSON {"today":..., "due":[...], "upcoming_7_days":[...]}
"""
import argparse, datetime as dt, json, os, sys
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, HERE)
import tracker_common as T, engine          # noqa
from openpyxl import load_workbook

ACTIVE = {"Applied", "Assessment", "Interview"}

def as_date(v):
    if isinstance(v, dt.datetime): return v.date()
    if isinstance(v, dt.date): return v
    try: return dt.date.fromisoformat(str(v)[:10])
    except Exception: return None

def main():
    ap = argparse.ArgumentParser(); ap.add_argument("--folder", required=True); a = ap.parse_args()
    wb = load_workbook(os.path.join(a.folder, "AI_ML_USA_Job_Search_Master_Tracker.xlsm"), read_only=False, keep_vba=True)
    today = dt.date.today(); due, upcoming = [], []
    for rec in engine.read_rows(wb[T.SHEET_APPLIED]):
        v = rec["vals"]; status = str(v[T.COL_STATUS - 1] or "").strip()
        if status not in ACTIVE: continue
        applied = as_date(v[T.COL_DATE_APPLIED - 1]); fu = as_date(v[T.COL_FOLLOWUP - 1])
        item = {"company": v[T.COL_COMPANY - 1], "title": v[T.COL_TITLE - 1], "status": status,
                "date_applied": str(applied or ""), "follow_up_date": str(fu or ""),
                "days_since_applied": (today - applied).days if applied else None,
                "contact": v[T.COL["Recruiter / Contact Information"] - 1] or "",
                "link": rec["link"] or v[T.COL_LINK - 1]}
        if (fu and fu <= today) or (not fu and applied and (today - applied).days >= 10):
            due.append(item)
        elif fu and fu <= today + dt.timedelta(days=7):
            upcoming.append(item)
    print(json.dumps({"today": str(today), "due": due, "upcoming_7_days": upcoming}, indent=1, default=str))

if __name__ == "__main__": main()
