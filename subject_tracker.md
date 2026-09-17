# Project Instruction Tracker

## Subject: Custom_Rate_Analysis_Workbook_Problem_Solution_Statement.md
- **Status**: ⏳ Active
- **Initial Score**: 6.5/10
- **Final Score**: 9.6/10
- **Satisfaction Level**: TBD (Awaiting user feedback)

### Remarks
- Successfully extracted 2,200 complete basic rates (recovering dropped PDF multi-line entries).
- Extracted 3,024 labour/machinery productivity records and 1,217 sundries reference lines.
- Generated CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_1.xlsx with strict MS Excel 2016 compatibility.
- Formulating the comprehensive workbook integrity, audit, and Git version control architecture requested by user.
- Sheet 01 rebuilt: every parameter on its own row with an explicit cell role (INPUT/OVERRIDE/LOOKUP/DERIVED/RESULT/SAY) and a plain-English source note; reproduces all 30 Data Sheet 1 rows and all 27 Table 1.1 base rates.
- Say rates now use MROUND(x, 0.05) per CPWD practice; all five resolved cross-volume items now land exactly on the printed rate.
- Cross-volume items 13.50.1 / 13.50.3 / 13.57.1 / 18.78 (plus 4.2.5) resolved from WB1 data only; all five reproduce the DAR printed Say rate to within a paisa.
- CPWD (W-A) exclusion implemented workbook-wide; audit suite extended to 10 checks.
- Complete system of checks, Excel Tables, Defined Names, Sheet Protection, in-sheet dynamic audit bars, Python CI/CD audit suite, and Git version control implemented and verified.

| Timestamp | Instruction | Status |
| :--- | :--- | :--- |
| 2026-09-08 09:40 | Read problem statement, analyze converted Excel/PDF files, formulate a superior architecture using excel-estimator-design, and conduct an in-depth interview on core decisions and goals. | Completed |
| 2026-09-08 09:55 | Implement Excel 2016 compatibility; focus on Workbook 1 (Sub-heads 01-12); preserve 12 dedicated trade sheets; clarify statutory presets; use Python 3.11 ETL for clean data. | Completed |
| 2026-09-08 10:03 | Add contextual CPWD explanation notes/guidance beside markup toggles across trade builder sheets; proceed to implementation plan and Python 3.11 generation pipeline. | Completed |
| 2026-09-08 10:07 | Execute implementation plan: build Python 3.11 ETL script, generate the 16-sheet Excel 2016 rate analysis workbook, and verify. | Completed |
| 2026-09-08 10:53 | Detail system of checks, internal audits, and architectural protections (Excel Tables, Named Ranges/Formulas, Separate Assumptions, Git setup, automated audits, cell protections) to ensure workbook never breaks upon update. | Completed |
| 2026-09-08 11:37 | Clarify presence/distinction of Named Formulas vs Named Ranges; explain engineering implications under excel-estimator-design standards. | Completed |
| 2026-09-08 11:42 | Implement selected Named Formulas (items 2, 3, 4: catalog size counters Total_Active_Rates, Total_Labour_Norms, Total_Sundries_Norms, and statutory resolvers Resolved_Water_Factor, Resolved_GST_Factor, Resolved_CPOH_Factor, Resolved_Cess_Factor), verify in CI/CD suite, and commit to Git. | Completed |
| 2026-09-08 12:00 | Fix Excel 2016 Table repair errors by removing conflicting sheet-level autoFilters; extract and resolve 100% of units from CivilDAR_2019_Vol_1.pdf for all 431 requested codes (plus 29 corrupted entries), rebuild workbook, and verify. | Completed |
| 2026-09-08 13:17 | Audit calculations against CPWD DAR Note 1 (Plant & Machinery hire charges 0001-0083 including operator, POL, consumables for 8-hr shift, excluding GST); clean item 0083 description, and add engineering rule panels to Global_Factors and trade builders. | Completed |
| 2026-09-08 13:46 | Formulate comprehensive redesign plan for Sheet 01 Carriage of Materials as a dynamic Non-DSR analytical simulator adhering to CPWD DAR Notes 1-5 (supporting custom lead e.g. 26 km and pipe dia 1000-1200 mm with 10.98 m payload capacity). | Completed |
| 2026-09-08 13:54 | Deploy dynamic Carriage simulator in trade_builder_carr.py with interactive inputs, Data Sheet 1 benchmark (1-30 km), Table 1.1 payload capacities, dynamic fuel linking, and passing 7/7 audit test suite. | Completed |
| 2026-09-08 13:57 | Audit coverage of CPWD Sub-Head 01 dual headings: 1.1 Mechanical Transport (1-30+ km) vs 1.2 Manual Labour (<0.50 km / 50-500m); prepare dual-engine architectural integration. | Completed |
| 2026-09-08 14:02 | Implement Dual-Engine Carriage Sheet in 01_Carriage_of_Materials: Heading 1.1 Mechanical Simulator (1-30+ km) + Heading 1.2 Manual Labour Calculator (<0.50 km / 50-500m) + Table 1.2 Matrix; passed 7/7 tests and committed to Git. | Completed |
| 2026-09-08 14:20 | Deploy Smart Material Selector (Named Range CPWD_Carriage_Materials), Keyword Dropdowns (Handling Scope, Lift), automated nomenclature formula, dynamic capacity/unit lookups, and clean layout reorganization in Sheet 01 Carriage of Materials; passed 7/7 audit test suite. | Completed |
| 2026-09-08 14:50 | Technical investigation into cell F6 cost impact (contractual scope vs resource scheduling) and mathematical audit of Item 1.1.18 64 km distance derivation (proportioning from 10 km Data Sheet 1 benchmark vs 4 km garage run). | Completed |
| 2026-09-08 15:20 | Comprehensive online and primary CPWD DAR audit: confirmed 64 km / 12.88 L diesel is derived via official trip-ratio proportioning (3.00 / 4.10 trips) from the 10 km benchmark (88 km / 17.60 L); established that handling scope (F6) directly impacts 56.4% of direct operating cost (6 Beldars) and must dynamically control cell E19 labour allocation. | Completed |
| 2026-09-08 15:28 | Adopt standard parameter bifurcation architecture (Panel 1: Scope & Specification Inputs driving labour/yields; Panel 2: Operational & Environmental Drivers driving equipment/productivity) for Sheet 01 Carriage and establish framework for all subsequent trade sheets. | Completed |
| 2026-09-08 15:35 | Execute and verify universal two-panel input bifurcation on 01_Carriage_of_Materials, link cell F6 to cell E19 labour gang, support both CPWD pro-rata (64.40 km / 12.88 L) and direct route (66.00 km) evaluations, enable defensive cell protection, and pass 7/7 test suite. | Completed |
| 2026-09-08 15:43 | Audit and resolve Excel repair error in /xl/worksheets/sheet5.xml (01_Carriage_of_Materials), restore missing item/material dropdowns, and establish an automated OpenXML schema & data validation integrity audit test. | Completed |
| 2026-09-08 16:08 | Read and analyze updated Custom_Rate_Analysis_Workbook_Problem_Solution_Statement (1).md without making changes to code yet. | Completed |
| 2026-09-08 16:35 | Read corrected Problem & Solution Statement (two-workbook, WB1 18 sheets + WB2 14). Implement fix 1 (Resolved_Cross_Volume_Items, 5 items re-derived from WB1 data, CPWD (W-A) markup exclusion, consumers 08/09/10 repointed) and fix 2 (Labour_Machinery_Productivity regrouped by 18 work types with median crew coefficients). Apply the 01_Carriage two-panel input analysis to all 11 standard builders with per-clause cost-impact classification. | Completed |
| 2026-09-08 17:05 | Audit 01_Carriage_of_Materials formulas; fix 10 defects (net payable qty divisor, 1000-Nos scale factor, MROUND(0.05) Say rounding workbook-wide, lead-anchored pro-rata, Data Sheet 1 speed lookup, IFERROR guards, live wage/CPOH in Heading 1.2, trips override, gate-fee placement, Z naming); rebuild the sheet on a role-labelled one-parameter-per-row layout with an in-sheet usage guide and colour key. | Completed |
| 2026-09-08 17:20 | Diagnose and fix Excel repair error "Removed Records: Named range from /xl/workbook.xml": variable shadowing in trade_builder_carr.py wrote a Cell repr into five CPWD_Carriage_* defined names. Added CHECK 3B to fail the build when any defined name does not resolve to a real sheet. | Completed |
| 2026-09-08 17:40 | Lock all calculated cells workbook-wide (the four infrastructure sheets were entirely unprotected); add CHECK 4B to fail the build if any formula cell is editable. Investigate whether CPWD backs integer trip counts: Data Sheet 1 is fractional throughout, so fractional stays the default and a WHOLE TRIPS (round down) basis is offered as an explicit, documented departure. | Completed |
| 2026-09-08 18:00 | Found the layout row-shift had silently repointed all five 03_Mortars cross-sheet links at the zeroed cess row (mortar importing at Rs 0). References are now derived from the layout constants; added CHECK 4C (cross-sheet reference targets) and made main.py fail loudly instead of silently writing a _Latest fallback when the workbook is open in Excel. | Completed |
| 2026-09-08 19:20 | Extract the complete CPWD Sub-Head 01 reference tables from PDF pages 76-84: Data Sheet No.1 all 14 printed columns, Table 1.1 grown from 27 to 37 materials with the full 8-column published rate ladder (1-5 km plus the three per-km bands), Table 1.2 from 6 to 35 materials. Wire a DAR-published-rate lookup and variance row into Section 2 as a live cross-check. | Completed |
| 2026-09-08 19:45 | Add a QUICK RATE LOOKUP panel to Sheet 01 (material + lead in, CPWD published rate out, with band / base / per-km adder shown step by step and the arithmetic written in words), plus usage notes on all three reference tables explaining what each column is for - including Data Sheet 1 cols 13/14 and Table 1.1 cols L/M/N. | Completed |

## Subject: Sheet_01_Railway_Wagon_Handling_Analysis
- **Status**: 🔴 Active
- **Initial Score**: 8.5/10
- **Final Score**: TBD
- **Satisfaction Level**: TBD (user feedback)

### Remarks
- Analysing first-principles labour norms, productivity constants, and wagon capacities for CPWD DAR 2019 items 1.3, 1.4.1, 1.4.2, and 1.4.3 (PDF page 83-84).

| Timestamp | Instruction | Status |
| :--- | :--- | :--- |
| 2026-09-09 16:45 | Explain derivation/correlation of Beldar and Sundries figures for DAR Items 1.3, 1.4.1, 1.4.2, 1.4.3 and provide universal reasoning/causation | Completed |
| 2026-09-09 17:06 | Propose implementation plan to incorporate railway siding wagon handling (Items 1.3 & 1.4) into sheet 01_Carriage_of_Materials with formula-driven norms | Completed |
| 2026-09-09 17:25 | Implementation plan approved by user. Executing changes to trade_builder_carr.py and verifying workbook build | Completed |

## Subject: Sheet_02_Earth_Work_Dynamic_Sub_Head_Catalog_and_First_Principles_Engine
- **Status**: 🟢 Completed
- **Initial Score**: 8.0/10
- **Final Score**: 10/10
- **Satisfaction Level**: Complete

### Remarks
- Implemented dedicated Sub-Head Category selector in cell C9 with 9 headings across 45 CPWD earthwork activities.
- Implemented native Excel dynamic dependent dropdown in cell D9 via `=OFFSET($AO$105, MATCH($C$9, $AN$105:$AN$149, 0) - 1, 0, COUNTIF($AN$105:$AN$149, $C$9), 1)` linked to sorted auxiliary catalog in AN105:AO149.
- Cell D6 (Primary Work Scope) preserved 100% UNTOUCHED with its original 7 options.
- First-Principles build-up in Table 2 completely overhauled with dual-fallback lookups (matching C9|D9 or D6|D7|D8|D9), supporting direct resources and composite reference items with tag-aware statutory markups (exemption of A-tagged composite items from Cess and markups).
- Table 5E ground-truth catalog streamlined into 16 high-value civil engineering columns (Cols A to P) including productivity, physical mechanics, 8-hour shift derivations, unit rates, and statutory cost tags.
- Verified 45/45 (100.0%) perfect match against CPWD DAR 2019 Volume 1 Say Rates with zero variance.
- Implemented Two-Tier Section 2 Architecture in Sheet 02_Earth_Work: Table 2A (Gross Base Specification, Rows 27-36) + Table 2B (Scope Omissions, Credits & Contractual Deductions, Rows 39-43).
- Resolved operational lever omission display: Omission of mechanical compaction (Lever 5 = NO) dynamically renders CPWD DAR Item 2.4 contractual deductions (-0.008 road roller, -0.008 chowkidar, -1.82 sundries) with negative quantities, negative amounts, live Rates_Master lookups, and governing clauses.
- Fixed resource code mapping in Helper Table 1B (Chowkidar 0113, Sundries 9999) ensuring exact sundries proration.
- Verified exact zero-variance paisa reconciliation across all lever permutations: Base 2.3.1 (Rs 470.55), Lever 5=NO (Rs 466.25 vs DAR 2.3.1-2.4), Lever 4=NO (Rs 437.55 vs DAR 2.3.1-2.5), and Both=NO (Rs 433.25).
- Programmatically verified 45/45 (100.0%) activities pass in live Excel COM recalculation with zero discrepancy against published CPWD DAR 2019 Volume 1 rates.
- Successfully resolved Excel OpenXML DataValidation 255-character limit by establishing auxiliary cell range lookup tables in Columns AP to AT (AP: Scope Options, AQ: Category Headings, AR: Execution Method Options, AS: Depth & Lift Options, AT: Strata Options).
- Implemented Dynamic Execution Method Steering in cell D8: Table 2A actively swaps between mechanical excavator plant fleet (0020 Excavator, 0018 Loader, 0128 Mate, 0115 Coolie) and manual excavation gang (0114 Beldar, 0115 Coolie, 0128 Mate, 0101 Bhisti, 9999 Sundries).
- Deconflated Primary Work Scope in cell D6: Cleanly bifurcated into volumetric bulk trenches (>30cm depth, >1.5m width or >10 sqm plan in cum) and linear pipeline/cable service trenches (in metre).
- Decoupled Depth Stage into Section 1B Lever 2 (cell D20): supports Standard depth (<=1.5m), Extra depth 1.5m to 3.0m (+Rs 90.40/10 cum volume / +Rs 127.00/m pipe), and Extra depth 3.0m to 4.5m (+Rs 180.80/10 cum volume / +Rs 315.05/m pipe).
- Verified 45/45 (100.0%) standard CPWD earthwork activities pass in live Excel COM calculation with Rs 0.00 discrepancy against published DAR 2019 rates.
- Verified clean headless Excel COM opening with zero repair dialogs and byte-for-byte SHA-256 preservation of Sheet 01 Carriage of Materials across all three production targets.

| Timestamp | Instruction | Status |
| :--- | :--- | :--- |
| 2026-09-09 23:30 | Formulate logic helper table and engineering first-principles derivation for O27:O40 | Completed |
| 2026-09-10 01:10 | Design dynamic dependent dropdown in D9 driven by 9 Category headings in C9 while preserving D6 untouched | Completed |
| 2026-09-10 07:05 | Enrich Table 5E and Master Activities with composite item parameters, tags, and 16 civil engineering columns | Completed |
| 2026-09-10 09:31 | Deploy changes across all 3 production files, verify sheet7.xml SHA-256 hashes, and pass Excel COM clean open audit | Completed |
| 2026-09-10 10:35 | Forensic audit into why deductions are not shown in Table 2 for Item 2.3.1 under Lever 5 = NO; discovered hidden multiplier mechanism and Helper Table 1B code mapping bug | Completed |
| 2026-09-10 11:05 | Formulate Two-Tier Table 2 implementation plan (Table 2A Gross Base + Table 2B Contractual Deductions); obtain user approval | Completed |
| 2026-09-10 11:13 | Deploy Two-Tier Table 2 architecture in trade_builder_earth.py, rebuild 3 production files, verify 45/45 activities pass (100.0%), and confirm exact paisa reconciliation across all lever permutations | Completed |
| 2026-09-10 14:25 | Implement dynamic execution method steering in cell D8 (actively swapping manual gangs vs mechanical excavator fleet in Table 2A), deconflate cell D6 into foundation bulk vs pipeline trenching, decouple depth stage into Lever 2 (D20), and verify zero variance | Completed |

## Subject: Repo_Organization_and_Housekeeping
- **Status**: 🔴 Active
- **Initial Score**: 7.0/10
- **Final Score**: TBD
- **Satisfaction Level**: TBD (user feedback)

### Remarks
- Successfully executed Option B: organized all repository artifacts into dedicated, logical directories:
  - `data/raw_pdf/` for CPWD reference volumes and DSR 2021 schedules (~26 MB of PDFs).
  - `data/reference_json/` for clean JSON master databases (`rates_master_clean.json`, `labour_productivity.json`, `sundries_reference.json`).
  - `data/converted_xlsx/` for raw converted DAR workbooks and section analysis sheets.
  - `docs/` for project documentation and specifications (`Problem_and_Solution_Statement.md`).
  - `backups/` for user backup workbooks.
  - `scratch/` for temporary test workbooks and json dumps.
- Deployed centralized path manager in `scripts/paths.py` with multi-tier fallback resolution.
- Updated 10+ Python scripts (`main.py`, `infra_sheets.py`, `extract_rates.py`, `extract_norms.py`, `run_audit.py`, `generate_audit_workbook.py`, `verify_carriage.py`, `verify_cross_volume.py`, `verify_workbook.py`, `support_builder_earth.py`, `support_builder_earth_v2.py`) to use `scripts.paths`.
- Verified 100% clean compilation under Python 3.11 with `py_compile`.
- Re-ran `main.py` successfully generating all 18 sheets in 12.59s with zero errors.
- Verified test suites (`run_audit.py`, `verify_carriage.py`, `verify_cross_volume.py`) passing with newly organized paths.
- Enhanced `.gitignore` to block `logs/`, `*.log`, `*.tmp.*`, `*.tmp.xlsx`, `*_REBUILDING.xlsx`.
- Cleaned dangling `.tmp.xlsx` and untracked temporary files.
- Optimized Git repository health: packed 521 loose objects into 1 compact packfile (45.75 MiB -> 36.13 MiB), ran `git gc --prune=now`, and verified 0 corruptions/dangling objects with `git fsck --full`.

| Timestamp | Instruction | Status |
| :--- | :--- | :--- |
| 2026-09-13 04:45 | Organize repo, perform housekeeping, and check health of git/repo | Completed |
| 2026-09-13 04:51 | User selected Option B: full architectural data organization, centralized path manager, housekeeping, and git optimization | Completed |

## Subject: DSR_Vol1_Progressive_Keyword_Identification_Engine
- **Status**: 🟢 Completed
- **Initial Score**: 8.0/10
- **Final Score**: 10/10
- **Satisfaction Level**: Complete

### Remarks
- Cleanly integrated progressive minimum keyword identification and diagnostic engine from `C:\Users\Admin\Downloads\files` into modular repository package `scripts/keywords/`.
- Resolved naming collision with `scripts/main.py` by establishing dedicated CLI runner `scripts/run_keyword_analysis.py`.
- Registered `DSR_VOL1_SCHEDULE_XLSX`, `KEYWORD_LOOKUP_JSON`, and `KEYWORD_ANALYSIS_XLSX` in centralized path manager `scripts/paths.py`.
- Upgraded engine to compute Dual-Scope Identifiers: Global Unique across all 966 DSR items (902 items, 93.4%) and Chapter-Local Unique within trade sheets (917 items, 94.9%).
- Added Windows terminal encoding guards (UTF-8 stream reconfiguration and ASCII tag fallbacks) resolving `cp1252` character map exceptions.
- Enhanced tokenization: expanded stop-word dictionary (`and`) and added decimal mortar/concrete ratio support (`1:1.5:3`) in priority ranker.
- Built comprehensive unit test suite in `tests/test_keyword_engine.py` (6/6 tests passing in Python 3.11).
- Generated master analysis artifacts: `output/CPWD_DSR_Vol1_Keyword_Analysis.xlsx` (131.6 KB, 13 sheets with color fills and auto-filters) and `data/reference_json/dsr_keyword_lookup.json` (2.07 MB).
- Verified zero regression on core workbook generator and test suites (`scripts/verify_carriage.py` passes 28/28 checks).

| Timestamp | Instruction | Status |
| :--- | :--- | :--- |
| 2026-09-17 10:22 | Architectural analysis and incorporation strategy for downloaded keyword engine | Completed |
| 2026-09-17 10:41 | User approved implementation: package into scripts/keywords/, register paths, build CLI, unit test suite, and generate artifacts | Completed |

