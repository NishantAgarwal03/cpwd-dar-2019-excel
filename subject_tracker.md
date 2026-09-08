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
