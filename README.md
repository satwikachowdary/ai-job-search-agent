# ai-job-search-agent

A personal job-search automation system that scans employer applicant-tracking
systems (ATS) directly, scores postings against a candidate profile, and
maintains a single Excel workbook as the system of record — with pending and
applied applications kept on separate sheets that move automatically as you
update application status.

This repo holds the **engine** — the scanning, scoring, dedup, and
spreadsheet-maintenance code. It does not include any personal data: no
resume, no candidate profile, no job-search history, and no live tracker
workbook. Those are supplied locally by whoever runs it (see "Bring your own
profile" below).

## What it does

- **Scans ATS boards directly**, rather than scraping job boards: Greenhouse,
  Ashby, and Lever all publish read-only JSON APIs per employer
  (`boards-api.greenhouse.io/v1/boards/{token}/jobs`,
  `api.ashbyhq.com/posting-api/job-board/{org}`,
  `api.lever.co/v0/postings/{token}`). `scan_broad.js` / `scan_boards.js` /
  `scan_ma_boards.js` walk a configurable list of employer tokens
  (`sources.example.json`) and pull every open posting's title, location,
  posted date, and description.
- **Filters for eligibility** using regex rules you define: required
  citizenship/clearance, no-sponsorship statements, PhD-required, minimum
  years of experience, etc. Jobs that fail are logged to an exclusion file
  with the specific reason, never silently dropped.
- **Scores and prioritizes** every eligible posting against a weighted rubric
  (skills/project match, experience/education fit, role relevance, location,
  compensation/sponsorship) into P1/P2/P3 tiers.
- **Deduplicates** against everything already tracked, by normalized URL,
  by company+title+city, and by company+requisition-ID — so re-running the
  scan never creates duplicate rows.
- **Maintains one Excel workbook** (`.xlsm`) with two visible sheets — *To
  Apply* and *Applied Jobs* — using openpyxl plus a small VBA macro
  (not included here, since it's workbook-embedded) that moves a row to the
  other sheet the moment you change its Application Status. Every write goes
  through a lock-check → backup → scratch-file → verify sequence, so a
  crash or a concurrent edit in Excel never corrupts the workbook, and your
  manually-entered fields (notes, recruiter info, resume/cover-letter
  versions, application dates) are never overwritten by a later scan.
- **Supports a local-employer registry** (`ma_registry.py`) for tracking
  companies with no public ATS feed yet — flagging candidates to check
  manually and recording confirmed feeds once found — plus a watchlist for
  specific employers to flag on *any* new posting, in or out of your normal
  filters.

## Files

| File | Purpose |
|---|---|
| `automation/engine.py` | Shared scoring/formatting helpers used by the other scripts. |
| `automation/tracker_common.py` | Column schema, sheet layout, and lock/backup/verify helpers for the Excel workbook. |
| `automation/tracker_update.py` | Main entrypoint: takes a batch of newly-found jobs, dedups, scores, and writes them into *To Apply*. Supports `--dry-run`. |
| `automation/maintain_tracker.py` | Reconciles Application Status changes — moves rows between *To Apply* and *Applied Jobs*, refreshes summary counts and sort order, preserves manual fields. |
| `automation/exclude_jobs.py` | Logs ineligible postings with their specific exclusion reason, with a running candidate log for re-checking. |
| `automation/record_application.py` | Records a submitted application (resume/cover-letter version, date, notes) into *Applied Jobs*. |
| `automation/run_state.py` | Tracks last successful run time and appends a run log. |
| `automation/ma_registry.py` | Manages a regional/local employer registry: records confirmed ATS feeds, surfaces the next batch of unchecked candidates, flags authorization blockers. |
| `automation/flag_nonprofits.py` | Marks nonprofit employers in the tracker with distinct formatting. |
| `automation/follow_ups.py` | Surfaces applications due for a follow-up. |
| `automation/scan_broad.js` | Browser-executed scan across Greenhouse/Ashby/Lever boards with eligibility-flag extraction, for pasting into a browser console or running via a browser-automation tool. |
| `automation/scan_boards.js` / `scan_ma_boards.js` | Narrower/regional variants of the same scan. |
| `automation/scan_builtin.js` | Scans a job-board site's internal search API (e.g. Built In) in resumable batches, to stay under a browser-automation tool's per-call time limit. |
| `automation/sources.example.json` | Example config: lists of employer ATS tokens to scan, plus the title-include/exclude regex used for a first-pass relevance filter. Replace with your own target list. |

## Bring your own profile

This repo intentionally ships with no candidate data. To use it:

1. Write your own eligibility rules and scoring weights (there's no code
   dependency on a specific resume — the regex patterns and weights in
   `sources.example.json` and `engine.py` are meant to be edited).
2. Build your own `sources.json` with the ATS tokens for employers you want
   to track (most employer career pages that use Greenhouse, Ashby, or Lever
   expose their token in the application URL).
3. Create your own `.xlsm` workbook matching the column schema in
   `tracker_common.py`, with a VBA macro that watches the Application Status
   column and moves rows between sheets (or adapt `maintain_tracker.py`'s
   row-move logic to run the move in Python instead).
4. Run `tracker_update.py --dry-run` first, always — every write-path script
   in this repo supports a dry run and prints what it would change before
   touching the real file.

## Design notes

- **Never silently drop a result.** Every posting that gets filtered out is
  logged with the specific reason it failed (unmet eligibility rule, outside
  the freshness window, etc.), so nothing just disappears without a trace.
- **Posting-date freshness is verified, not estimated.** Dates come only
  from the employer's own API fields (e.g. Greenhouse's `first_published`,
  Ashby's `publishedAt`, Lever's `createdAt`); if a source doesn't expose a
  reliable date, postings from it are flagged as unverified rather than
  assumed fresh.
- **The workbook is the single source of truth.** Every script reads the
  current file, never assumes state from a prior run alone, and verifies its
  own write (checksum of unrelated cells, VBA project unchanged) before
  replacing the live file.

## License

MIT — do whatever you like with it.
