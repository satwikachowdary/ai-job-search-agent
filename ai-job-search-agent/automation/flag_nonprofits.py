#!/usr/bin/env python3
"""flag_nonprofits.py - record which Massachusetts registry employers / candidates are nonprofits (idempotent).
Adds a top-level "nonprofit_employers" map to automation/ma_employers.json and sets "nonprofit" and
"cap_exempt_possible" on matching registry entries. Never removes anything.
cap_exempt_possible is True only for higher-education institutions, nonprofits affiliated with one,
and nonprofit research organizations - and even then it is only a possibility to confirm with HR.
Usage: python3 flag_nonprofits.py <path to ma_employers.json>
"""
import json, sys, shutil

NP = {
    # registry employers
    "BCBS of Massachusetts": ("nonprofit health insurer", False, ""),
    "Beth Israel Lahey Health": ("nonprofit health system", True, "possible only for its Harvard-affiliated teaching hospital entities (e.g., BIDMC)"),
    "Boston University": ("private university", True, ""),
    "Dana-Farber Cancer Institute": ("nonprofit cancer center, Harvard-affiliated", True, ""),
    "Harvard University": ("private university", True, "some Harvard postings state they cannot sponsor; check each posting"),
    "Mass General Brigham": ("nonprofit health system, Harvard-affiliated teaching hospitals", True, ""),
    "MIT Lincoln Laboratory": ("MIT federally funded R&D center", True, "most roles need US citizenship/clearance"),
    "MITRE": ("nonprofit federally funded R&D center", True, "most roles need US citizenship/clearance"),
    "Northeastern University": ("private university", True, ""),
    # candidates not yet in the registry
    "Babson": ("private college", True, ""),
    "Bentley": ("private university", True, ""),
    "Brandeis": ("private university", True, ""),
    "Olin College": ("private college", True, ""),
    "Tufts University": ("private university", True, ""),
    "UMass": ("public university system (not-for-profit)", True, ""),
    "WPI": ("private university", True, ""),
    "MIT": ("private university", True, ""),
    "Boston Children's": ("nonprofit pediatric hospital, Harvard-affiliated", True, ""),
    "Broad Institute": ("nonprofit biomedical research institute (MIT/Harvard)", True, ""),
    "Whitehead Institute": ("nonprofit research institute, MIT-affiliated", True, ""),
    "Wyss Institute": ("research institute within Harvard University", True, ""),
    "Draper": ("nonprofit R&D laboratory", True, "many roles need US citizenship/clearance"),
    "Mass Robotics": ("nonprofit innovation hub", False, ""),
    "Point32Health": ("nonprofit health plan", False, ""),
}

p = sys.argv[1]
d = json.load(open(p))
shutil.copyfile(p, p + ".bak_before_nonprofit_flags")
m = d.setdefault("nonprofit_employers", {})
for name, (typ, cap, note) in NP.items():
    m[name] = {"type": typ, "cap_exempt_possible": cap, "note": note}
hit = 0
for e in d["employers"]:
    if e["employer"] in NP:
        typ, cap, note = NP[e["employer"]]
        e["nonprofit"] = True; e["nonprofit_type"] = typ; e["cap_exempt_possible"] = cap
        hit += 1
d["nonprofit_note"] = ("Nonprofit status recorded for the visible '◆ Nonprofit' mark. cap_exempt_possible means the employer "
                       "type (university, university-affiliated nonprofit, nonprofit research org) may be H-1B cap-exempt; "
                       "confirm with HR. It is not a sponsorship guarantee.")
json.dump(d, open(p, "w"), indent=1)
print(json.dumps({"ok": True, "nonprofits_recorded": len(m), "registry_entries_flagged": hit}))
