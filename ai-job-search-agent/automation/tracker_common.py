"""
Shared layout/config for AI_ML_USA_Job_Search_Master_Tracker.xlsm.

Single source of truth for: sheet names, column order, row layout,
dropdown value lists, status buckets, colors, and the move/sort/reconcile
logic. The builder, the logic test, and the scheduled-run updater all import
from here so the Python side and the VBA side stay in lockstep.
"""

# ---- Sheets ---------------------------------------------------------------
SHEET_TOAPPLY = "TO APPLY"
SHEET_APPLIED = "APPLIED JOBS"

# ---- Row layout (same on both sheets) -------------------------------------
# Row 1: banner | Row 2: summary metrics | Row 3: spacer
# Row 4: column headers (frozen) | Row 5+: job data
BANNER_ROW = 1
SUMMARY_ROW = 2
SPACER_ROW = 3
HEADER_ROW = 4
FIRST_DATA_ROW = 5
LAST_DATA_ROW = 1004          # 1000 data rows provisioned for DV/CF/filters

# ---- Columns (A..Z per spec, + AA hidden helper) --------------------------
# (header_text, width, wrap)
COLUMNS = [
    ("Priority", 9, False),                              # A
    ("Resume Match Score", 11, False),                   # B
    ("Application Status", 16, False),                   # C
    ("Company Name", 24, False),                         # D
    ("Job Title", 30, False),                            # E
    ("City / State", 20, False),                         # F
    ("Work Arrangement", 15, False),                     # G
    ("Salary / Salary Range", 20, False),                # H
    ("Sponsorship Status", 18, False),                   # I
    ("OPT / Work Authorization Information", 26, True),   # J
    ("Experience Required", 18, True),                   # K
    ("Job Posting Date", 14, False),                     # L
    ("Date and Time Discovered", 20, False),             # M
    ("Direct Application Link", 34, False),              # N
    ("Job Source", 16, False),                           # O
    ("Why This Job Matches My Resume", 44, True),        # P
    ("Missing Skills / Qualifications", 34, True),       # Q
    ("Resume Uploaded?", 14, False),                     # R
    ("Resume Version Used", 22, False),                  # S
    ("Cover Letter Uploaded?", 15, False),               # T
    ("Cover Letter Version", 22, False),                 # U
    ("Date Applied", 14, False),                         # V
    ("Recruiter / Contact Information", 28, True),       # W
    ("Follow-up Date", 14, False),                       # X
    ("Personal Notes", 34, True),                        # Y
    ("Job Availability", 14, False),                     # Z
    ("Record Key", 30, False),                           # AA (hidden helper)
]
NCOLS = len(COLUMNS)           # 27
LAST_COL_LETTER = "AA"

# 1-based column indexes by field (for code clarity)
COL = {name: i + 1 for i, (name, _w, _wrap) in enumerate(COLUMNS)}
COL_PRIORITY = COL["Priority"]                 # 1
COL_SCORE = COL["Resume Match Score"]          # 2
COL_STATUS = COL["Application Status"]          # 3
COL_COMPANY = COL["Company Name"]              # 4
COL_TITLE = COL["Job Title"]                   # 5
COL_CITY = COL["City / State"]                 # 6
COL_LINK = COL["Direct Application Link"]      # 14
COL_DISCOVERED = COL["Date and Time Discovered"]  # 13
COL_RESUME_UP = COL["Resume Uploaded?"]        # 18
COL_CL_UP = COL["Cover Letter Uploaded?"]      # 20
COL_DATE_APPLIED = COL["Date Applied"]         # 22
COL_FOLLOWUP = COL["Follow-up Date"]           # 24
COL_KEY = COL["Record Key"]                    # 27

# ---- Status buckets -------------------------------------------------------
STATUS_TOAPPLY = ["Not Applied", "Saved", "In Progress", "Ready to Apply"]
STATUS_APPLIED = ["Applied", "Assessment", "Interview",
                  "Rejected", "Offer", "Withdrawn After Applying"]
ALL_STATUSES = STATUS_TOAPPLY + STATUS_APPLIED   # dropdown offers all 10

# ---- Other dropdowns ------------------------------------------------------
DV_PRIORITY = ["P1", "P2", "P3"]
DV_YESNO = ["Yes", "No"]
DV_WORK = ["On-site", "Hybrid", "Remote"]
DV_SPONSOR = ["Available", "Not Available", "Unknown", "OPT/STEM OPT OK"]
DV_AVAIL = ["Open", "Closed", "Unknown"]

# ---- Colors (ARGB) --------------------------------------------------------
C_HEADER_BG = "FF1F3864"      # dark navy
C_HEADER_FG = "FFFFFFFF"
C_BANNER_BG = "FF2E5496"
C_BANNER_FG = "FFFFFFFF"
C_SUMMARY_BG = "FFEAF1FB"
C_P1 = "FFC6EFCE"             # light green row tint
C_P2 = "FFFFF2CC"            # light yellow row tint
C_P3 = "FFE7E6E6"            # light gray row tint
C_P1_CELL = "FF37B24D"        # strong green priority cell
C_P2_CELL = "FFF2C037"        # strong amber priority cell
C_P3_CELL = "FFADB5BD"        # strong gray priority cell
# status cell colors
C_STATUS = {
    "Not Applied":              ("FFFCE4E4", "FF842029"),
    "Saved":                    ("FFFFF3CD", "FF664D03"),
    "In Progress":              ("FFCFE2FF", "FF084298"),
    "Ready to Apply":           ("FFD1E7DD", "FF0F5132"),
    "Applied":                  ("FFD1E7DD", "FF0F5132"),
    "Assessment":               ("FFCFF4FC", "FF055160"),
    "Interview":                ("FFCFE2FF", "FF084298"),
    "Rejected":                 ("FFF8D7DA", "FF842029"),
    "Offer":                    ("FFFFE69C", "FF664D03"),
    "Withdrawn After Applying": ("FFE2E3E5", "FF41464B"),
}

FONT_NAME = "Arial"


# ===========================================================================
# Move / sort / reconcile logic (mirrors the VBA exactly)
# ===========================================================================
def record_key(company, title, city, link):
    """Stable identity key used for dedup. Normalized, case-insensitive."""
    def norm(x):
        return " ".join(str(x or "").strip().lower().split())
    return "|".join([norm(company), norm(title), norm(city), norm(link)])


def _priority_rank(p):
    return {"P1": 0, "P2": 1, "P3": 2}.get(str(p or "").strip().upper(), 3)


def sort_key_toapply(row):
    """P1->P3, then highest match score first."""
    score = row[COL_SCORE - 1]
    score = score if isinstance(score, (int, float)) else -1
    return (_priority_rank(row[COL_PRIORITY - 1]), -score)


def sort_key_applied(row):
    """Most recently applied first; fall back to score, then company."""
    import datetime as _dt
    d = row[COL_DATE_APPLIED - 1]
    # larger date = more recent = first -> use negative ordinal
    if isinstance(d, (_dt.datetime, _dt.date)):
        ord_ = d.toordinal() if isinstance(d, _dt.date) else d.date().toordinal()
        key_date = -ord_
    else:
        key_date = 1  # undated applied rows sink below dated ones
    score = row[COL_SCORE - 1]
    score = score if isinstance(score, (int, float)) else -1
    return (key_date, -score, str(row[COL_COMPANY - 1] or ""))
