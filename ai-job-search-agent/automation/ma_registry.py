#!/usr/bin/env python3
"""
ma_registry.py - Massachusetts employer registry helper (automation/ma_employers.json).

  python3 ma_registry.py --folder F tokens            -> JSON {greenhouse:[[name,token]..], ashby:[..], lever:[..]} for scan_ma_boards.js
  python3 ma_registry.py --folder F next-web --n 15   -> the N employers without a public feed checked longest ago (web-search rotation)
  python3 ma_registry.py --folder F record --file R   -> merge results JSON: list of {employer, category?, ma_area?, feeds?[{ats,token,careers_url}],
                                                         careers_site?, method, postings_total?, postings_in_ma?, relevant_openings?,
                                                         relevant_last_14_days?, outcome, note?}  (updates 'checked' = now ET)
  python3 ma_registry.py --folder F summary           -> counts (employers in registry, with feeds, checked today, with relevant openings)
Never deletes an employer; an employer with no openings stays in the registry so later runs revisit it.
"""
import argparse, datetime as dt, json, os
from zoneinfo import ZoneInfo
ET = ZoneInfo("America/New_York")

def path(folder): return os.path.join(folder, "automation", "ma_employers.json")
def load(folder):
    with open(path(folder)) as f: return json.load(f)
def save(folder, d):
    with open(path(folder), "w") as f: json.dump(d, f, indent=1)

def main():
    ap = argparse.ArgumentParser(); ap.add_argument("--folder", required=True)
    ap.add_argument("cmd", choices=["tokens", "next-web", "record", "summary"]); ap.add_argument("--n", type=int, default=15); ap.add_argument("--file")
    a = ap.parse_args(); d = load(a.folder); emps = {e["employer"]: e for e in d["employers"]}
    if a.cmd == "tokens":
        out = {"greenhouse": [], "ashby": [], "lever": []}
        for e in emps.values():
            for f in e.get("feeds", []):
                out[f["ats"]].append([e["employer"], f["token"]])
        print(json.dumps(out))
    elif a.cmd == "next-web":
        pool = [e for e in emps.values() if not e.get("feeds")]
        pool += [{"employer": n, "checked": "", "method": "not yet checked"} for n in d.get("candidates_no_public_feed_found_yet", []) if n not in emps]
        pool.sort(key=lambda e: e.get("checked") or "")
        print(json.dumps([{"employer": e["employer"], "careers_site": e.get("careers_site", ""), "last_checked": e.get("checked", "")} for e in pool[:a.n]], indent=1))
    elif a.cmd == "record":
        now = dt.datetime.now(ET).isoformat(timespec="minutes")
        with open(a.file) as f: res = json.load(f)
        for r in res:
            e = emps.setdefault(r["employer"], {"employer": r["employer"], "feeds": []})
            for k, v in r.items():
                if k == "feeds":
                    have = {(x["ats"], x["token"]) for x in e.get("feeds", [])}
                    e.setdefault("feeds", []).extend(x for x in v if (x["ats"], x["token"]) not in have)
                elif v is not None:
                    e[k] = v
            e["checked"] = now
            for alias in [r["employer"]] + r.get("aliases", []):          # once in the registry it is no longer a bare candidate
                if alias in d.get("candidates_no_public_feed_found_yet", []):
                    d["candidates_no_public_feed_found_yet"].remove(alias)
        d["employers"] = sorted(emps.values(), key=lambda e: e["employer"]); d["updated"] = now
        save(a.folder, d); print(json.dumps({"ok": True, "employers": len(emps)}))
    else:
        today = dt.datetime.now(ET).date().isoformat()
        print(json.dumps({"employers_in_registry": len(emps), "with_public_feed": sum(1 for e in emps.values() if e.get("feeds")),
                          "checked_today": sum(1 for e in emps.values() if str(e.get("checked", "")).startswith(today)),
                          "with_relevant_openings": sum(1 for e in emps.values() if e.get("relevant_openings")),
                          "not_yet_checked_candidates": len(d.get("candidates_no_public_feed_found_yet", []))}, indent=1))

if __name__ == "__main__": main()
