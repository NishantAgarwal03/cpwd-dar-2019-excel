# Project Checkpoint — 2026-09-19

Written as a pause/close-out summary before stepping away from the project for a while.
Covers everything done, current verified state, and exactly what's still open so a future
session (or a fresh Claude session) can pick this up without re-deriving context.

## Is it safe to close? Yes, with one thing to do first

**Close Microsoft Excel** if it's still open — `CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_1.xlsx`
and its `_Latest.xlsx` copy were confirmed still locked by Excel as of this checkpoint (lock files
`~$CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_1.xlsx` / `~$..._Latest.xlsx` present and busy).
This doesn't put anything at risk — the files on disk are unmodified and match what's committed to
git exactly — but a full pipeline re-run will fail at the workbook-save step until Excel releases
them (confirmed: a full `python run_pipeline.py` run failed cleanly at that exact step with the
files open, with zero side effects — steps 1-6 completed and produced byte-identical output to
what's already committed, nothing was corrupted).

## Git state: clean

- Working tree clean, nothing uncommitted (`git status --short` empty as of this checkpoint).
- All work from this session is committed. Latest commits (newest first):
  - `bfde3b9` docs: add earthwork_tables.py audit report
  - `eb1a729` fix: resolve all 7 verify_workbook.py structural failures on Vol1 workbook
  - `43eef9d` feat: add zero-tolerance automated verifier for public conversion workbooks
  - `3a86886` fix: correct percentage display for Ch02's conditional_extra items
  - `554336d` feat: add public sheet-per-subhead PDF-to-Excel conversion workbooks
  - `7dd4460` fix: re-verification round found and fixed new bugs in the 2026-09-18 fix
  - `6693746` fix: close all audit-found gaps, make Vol1 pipeline PDF-direct
- Nothing has been pushed to a remote (no remote was configured/used this session) — this is all
  local-only history.

## What's verified clean right now

| Deliverable | Verified via | Result |
|---|---|---|
| Core extraction pipeline (`ch01-26_items.json`, registries, `rates_master_clean.json`) | Full clean-state `run_pipeline.py` steps 1-6 | Deterministic, byte-identical to committed state |
| `CPWD_DAR_2019_Vol1_PDF_Conversion.xlsx` / `Vol2_PDF_Conversion.xlsx` (public conversion workbooks) | `scripts/verify_public_conversion_workbook.py` (zero-tolerance, exhaustive, every item) | PASS, zero discrepancies, both volumes |
| `CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_1.xlsx` (+ `_Latest.xlsx`, byte-identical) | `scripts/verify_workbook.py` (13 structural/formula checks) | 13/13 PASS (one sub-check, a live Excel-COM repair-dialog test, is flaky/environment-dependent on this machine — not a file defect, confirmed by reproducing the same pass/fail flip on repeated runs of the identical file) |

## What's still open (needs a decision or more work before it's "done")

These are known, documented, and not urgent — nothing here is a regression or a surprise, just
unfinished threads from where work paused.

1. **`earthwork_tables.py` — 26 of 74 dropdown entries in the live `02_Earth_Work` sheet have no
   backing rate data.** Selecting one of them (all timbering shaft/area/permanent variants, plus 6
   hard-rock pipe-trench variants) silently shows a different, wrong item's rate via a fallback
   lookup chain — no error, no warning. This is a live correctness bug on a deliverable in active
   use. See `reports/audit_earthwork_tables_2026-09-19.md` for the full list of affected codes and
   two fix options (populate the missing rows from `support_builder_earth_v2.py`'s data, or remove
   those 26 entries from the dropdown as a stop-gap). **Decision needed from user before fixing.**
2. **Item 2.9.3 rate discrepancy.** `earthwork_tables.py` says 1080.55/cum, `support_builder_earth_v2.py`
   (fixed earlier this session using PDF cross-checks) says 624.70/cum. Both carry an identical
   resource buildup; the arithmetic proves 624.70 is impossible (it's less than the raw direct cost
   before any markup). Strong evidence the earlier PDF-based fix for this one item was itself wrong.
   **Needs either a targeted PDF re-check or an explicit decision to trust the arithmetic and fix
   `support_builder_earth_v2.py`'s value without re-opening the PDF** (deferred, no user decision
   received yet before the pause).
3. **`CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_2.xlsx` has no structural checker equivalent
   to `verify_workbook.py`.** That checker is Vol.1-specific by design (hardcoded expected sheets/
   row layout). If Vol.2 structural integrity needs the same confidence, that's new work, not done.
4. **`run_audit.py`'s rate-accuracy findings (242 FAIL / 152 MISSING CODE / 8 RATE MISMATCH) were
   triaged and found to be testing the wrong thing** — it reads the old, retired
   `CivilDAR_2019_Vol_*_Converted.xlsx` files, which no longer feed the current pipeline or (for
   Earth Work specifically) the live workbook at all. Its output is not actionable as-is. If rate
   accuracy for the 10 *standard* trade sheets (03-12, which are live `Rates_Master` calculators,
   not duplicated catalogs) is ever in question, that would need a differently-scoped check —
   `Rates_Master` itself is already PDF-verified clean, so the standard sheets' risk surface is much
   smaller than Earth Work's was.
5. **Minor cleanup items noted but not fixed** (all low/very-low priority per
   `reports/audit_earthwork_tables_2026-09-19.md`): a duplicate `2.25a` dropdown entry, an unused
   `MASTER_ITEMS` dict with a latent sign bug, 2 orphaned resource records, one stale code comment.

## Where to look to resume

- `PIPELINE.md` — current data pipeline architecture and run order.
- `reports/` — every audit report generated this session and previously (dated, chronological).
- Claude's memory files (`fix_pdf_deletion_gaps_2026-09-18.md`, `reverify_and_fix_round2_2026-09-18.md`,
  `structural_fix_custom_workbook_2026-09-19.md`, this checkpoint) have the full narrative if picking
  this up in a new session.
- The two open decisions above (items 1-2) are the natural next steps if continuing this thread.
