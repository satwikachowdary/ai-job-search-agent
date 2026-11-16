#!/usr/bin/env python3
"""
run_state.py - persistent run state for the twice-daily job search (7:00 AM / 7:00 PM America/New_York).

State lives next to the workbook, on the Mac:  <tracker>/automation/state.json  and  run_log.md

  begin   -> decides what this run should do and prints JSON:
             now_et, allowed (False during quiet hours 23:00-06:00 ET unless --manual),
             last_success, missed_runs (scheduled 7AM/7PM slots skipped since then),
             catch_up (True if >= 1 slot was missed), search_posted_since (always today-14 days),
             window_bounds (A/B/C/D freshness windows as ET dates)
  finish  -> records a successful run (updates last_success) and appends to run_log.md
  fail    -> records an error; last_success stays put, so the next run catches up

Usage:
  python3 run_state.py --folder <tracker> begin [--manual] [--now 2026-11-11T07:00]
  python3 run_state.py --folder <tracker> finish --added 3 --summary "..."
  python3 run_state.py --folder <tracker> fail --error "..."
"""
import argparse, datetime as dt, json, os
from zoneinfo import ZoneInfo

ET = ZoneInfo("America/New_York")
SLOTS = (7, 19)            # scheduled hours, ET
QUIET_START, QUIET_END = 23, 6


def _p(folder, name):
    d = os.path.join(folder, "automation")
    os.makedirs(d, exist_ok=True)
    return os.path.join(d, name)


def load(folder):
    try:
        with open(_p(folder, "state.json")) as f:
            return json.load(f)
    except Exception:
        return {}


def save(folder, st):
    with open(_p(folder, "state.json"), "w") as f:
        json.dump(st, f, indent=2)


def log(folder, line):
    with open(_p(folder, "run_log.md"), "a") as f:
        f.write(line.rstrip() + "\n")


def missed_slots(last, now):
    """Scheduled 7AM/7PM ET slots strictly after `last` and strictly before the current slot."""
    if last is None:
        return 0
    n, d = 0, last.date()
    while d <= now.date():
        for h in SLOTS:
            t = dt.datetime(d.year, d.month, d.day, h, 0, tzinfo=ET)
            if last < t < now - dt.timedelta(minutes=90):
                n += 1
        d += dt.timedelta(days=1)
    return n


def begin(folder, manual=False, now=None):
    now = now or dt.datetime.now(ET)
    st = load(folder)
    last = dt.datetime.fromisoformat(st["last_success"]) if st.get("last_success") else None
    quiet = now.hour >= QUIET_START or now.hour < QUIET_END
    missed = missed_slots(last, now)
    today = now.date()
    out = {
        "now_et": now.isoformat(timespec="minutes"),
        "allowed": (not quiet) or manual,
        "reason": ("manual run" if manual else ("quiet hours (23:00-06:00 ET): no catch-up at night"
                                                if quiet else "scheduled window")),
        "last_success": st.get("last_success"),
        "missed_runs": missed,
        "catch_up": last is None or missed >= 1,
        "search_posted_since": str(today - dt.timedelta(days=14)),
        "window_bounds": {
            "A_last_24h": [str(today - dt.timedelta(days=1)), str(today)],
            "B_24_48h": [str(today - dt.timedelta(days=2))] * 2,
            "C_3_7_days": [str(today - dt.timedelta(days=7)), str(today - dt.timedelta(days=3))],
            "D_8_14_days": [str(today - dt.timedelta(days=14)), str(today - dt.timedelta(days=8))],
        },
        "last_error": st.get("last_error"),
    }
    st["last_attempt"] = now.isoformat(timespec="seconds")
    save(folder, st)
    return out


def finish(folder, added, summary, now=None):
    now = now or dt.datetime.now(ET)
    st = load(folder)
    st["last_success"] = now.isoformat(timespec="seconds")
    st["last_added"] = added
    st["last_error"] = None
    save(folder, st)
    log(folder, f"- {now:%Y-%m-%d %H:%M} ET  OK  added={added}  {summary}")
    return {"ok": True, "last_success": st["last_success"]}


def fail(folder, error, now=None):
    now = now or dt.datetime.now(ET)
    st = load(folder)
    st["last_error"] = {"time": now.isoformat(timespec="seconds"), "error": error}
    save(folder, st)
    log(folder, f"- {now:%Y-%m-%d %H:%M} ET  FAILED  {error}  (will retry at the next scheduled run)")
    return {"ok": True, "recorded_error": error}


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--folder", required=True)
    ap.add_argument("cmd", choices=["begin", "finish", "fail"])
    ap.add_argument("--manual", action="store_true")
    ap.add_argument("--now")
    ap.add_argument("--added", type=int, default=0)
    ap.add_argument("--summary", default="")
    ap.add_argument("--error", default="")
    a = ap.parse_args()
    now = dt.datetime.fromisoformat(a.now).replace(tzinfo=ET) if a.now else None
    if a.cmd == "begin":
        r = begin(a.folder, a.manual, now)
    elif a.cmd == "finish":
        r = finish(a.folder, a.added, a.summary, now)
    else:
        r = fail(a.folder, a.error, now)
    print(json.dumps(r, indent=2))
