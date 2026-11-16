"""
Python mirror of the VBA row-movement engine, operating on the real
workbook via openpyxl. Same rules as JobTracker.bas:
  - a TO APPLY row whose status is an 'applied' status moves to APPLIED JOBS
  - an APPLIED JOBS row whose status is a 'pending' status moves to TO APPLY
  - TO APPLY sorted P1->P3 then score desc; APPLIED sorted date-applied desc
  - all 27 columns (incl. hidden Record Key) + the hyperlink are preserved
Used by both the test harness and the scheduled-run updater.
"""
import datetime as dt
from openpyxl.utils import get_column_letter
import tracker_common as T


def read_rows(ws):
    """Return list of row-records: {'vals':[27], 'link':hyperlink_or_None}."""
    rows = []
    r = T.FIRST_DATA_ROW
    while True:
        company = ws.cell(r, T.COL_COMPANY).value
        # stop at first fully-empty company beyond the used block
        if company is None or str(company).strip() == "":
            # peek a little further in case of gaps; but our engine keeps data
            # contiguous, so an empty company means end of data.
            break
        vals = [ws.cell(r, c).value for c in range(1, T.NCOLS + 1)]
        link = ws.cell(r, T.COL_LINK).hyperlink
        link_target = link.target if link else None
        rows.append({"vals": vals, "link": link_target})
        r += 1
    return rows


def _clear_data(ws):
    r = T.FIRST_DATA_ROW
    while True:
        if ws.cell(r, T.COL_COMPANY).value in (None, ""):
            # clear a generous band to remove any stragglers
            break
        for c in range(1, T.NCOLS + 1):
            cell = ws.cell(r, c)
            cell.value = None
            cell.hyperlink = None
        r += 1
    # also wipe a safety margin below
    for rr in range(r, r + 5):
        for c in range(1, T.NCOLS + 1):
            ws.cell(rr, c).value = None
            ws.cell(rr, c).hyperlink = None


def write_rows(ws, rows):
    _clear_data(ws)
    from openpyxl.styles import Font
    for idx, rec in enumerate(rows):
        r = T.FIRST_DATA_ROW + idx
        for c in range(1, T.NCOLS + 1):
            ws.cell(r, c).value = rec["vals"][c - 1]
        if rec["link"]:
            cell = ws.cell(r, T.COL_LINK)
            cell.hyperlink = rec["link"]
            cell.font = Font(name=T.FONT_NAME, size=10, color="FF0563C1", underline="single")


def reconcile_and_sort(wb):
    """Mirror of Workbook_Open + post-change state: ensure every row is on the
    correct sheet, then sort both. Idempotent."""
    ws_p = wb[T.SHEET_TOAPPLY]
    ws_a = wb[T.SHEET_APPLIED]
    pending = read_rows(ws_p)
    applied = read_rows(ws_a)

    keep_p, keep_a = [], list(applied)
    for rec in pending:
        status = str(rec["vals"][T.COL_STATUS - 1] or "").strip()
        if status in T.STATUS_APPLIED:
            keep_a.append(rec)
        else:
            keep_p.append(rec)

    moved_a = []
    for rec in keep_a:
        status = str(rec["vals"][T.COL_STATUS - 1] or "").strip()
        if status in T.STATUS_TOAPPLY:
            keep_p.append(rec)
        else:
            moved_a.append(rec)

    keep_p.sort(key=lambda rec: T.sort_key_toapply(rec["vals"]))
    moved_a.sort(key=lambda rec: T.sort_key_applied(rec["vals"]))

    write_rows(ws_p, keep_p)
    write_rows(ws_a, moved_a)
    return len(keep_p), len(moved_a)
