# Independent Verification Audit — rates_master_clean.json & Chapter 02 Earthwork

**Date:** 2026-09-17
**Scope:** `data/reference_json/rates_master_clean.json` (codes 0001–2399, PDF pages 11–72, 0-indexed) and Chapter 02 Earth Work (`scripts/support_builder_earth_v2.py` ITEMS list + `data/reference_json/earthwork_pdf_specs.json`) vs `data/raw_pdf/CivilDAR_2019_Vol_1.pdf`.
**Method:** Independent Python extraction using `pypdf` (not the project's own extraction scripts), written from scratch. Text was dumped page-by-page, parsed with an original regex-based line parser, and cross-checked by direct manual reading of the raw extracted text.

## Part 1 — rates_master_clean.json

**Sampling:** 893 codes in range 0001–2399 exist in the reference file. Sampled 50 codes: the first 5 (0001–0005), the last 5 (2395–2399), and ~40 evenly spaced codes across the full range (e.g. 0023, 0100, 0226, 0387, 0637, 0973, 1182, 1551, 1889, 2208, 2328 …). For every sampled code, my independent parser's rate/description was compared to `rates_master_clean.json`, and several (0755, 1182, 1889) were additionally manually re-read from the raw PDF text via `grep` to rule out parser artifacts.

**Result: 50/50 exact rate matches. Zero discrepancies found.**

Every sampled code's rate matched `rates_master_clean.json` to the cent, including edge cases with wrapped/multi-line descriptions (e.g. code 0755 "Decorative type louvers…", code 1889 "C.P. brass toilet paper holder…" — PDF text has broken words like "p aper" and "st andard" from PDF text-extraction spacing artifacts, but the numeric rate was unaffected and matched).

**Verdict for Part 1: rates_master_clean.json is a reliable, accurate replacement for the PDF's basic rate codes 0001–2399, based on this 50-code sample spanning the full page range including both extremes.** No evidence of extraction error was found in this slice.

## Part 2 — Chapter 02 Earth Work

The chapter lives on PDF pages 85–139 (0-indexed; printed page nos. 77–131), under "SUB HEAD : 2 - EARTH WORK", items 2.1 through 2.38.

**Sampling:** Independently transcribed and hand-verified 18 items spanning nearly every sub-section: 2.1.1, 2.2.1, 2.2.2(partial), 2.3.1, 2.6.1, 2.7.1, 2.7.2(partial), 2.9.1(partial), 2.10.1.1, 2.10.1.2, 2.16.1 (with full material/labour breakdown), 2.25, 2.25(a), 2.27, 2.28.1, 2.29.1, plus coverage-checked 2.35.3–2.38 (8 codes) against `earthwork_pdf_specs.json`.

### Discrepancies found

| # | Item/Code | PDF says | JSON/script says | PDF page | Type |
|---|---|---|---|---|---|
| 1 | 2.1.1 | Final rate (Cost of 1 sqm) = **92.55** | `dsr_rate` = **107.00** | 87 | Rate mismatch, +15.6% |
| 2 | 2.2.1 | Final rate (Cost of 1 cum) = **746.80** | `dsr_rate` = **862.70** | 87 | Rate mismatch, +15.5% |
| 3 | 2.3.1 | Final rate = **470.55** | `dsr_rate` = **543.40** | 88 | Rate mismatch, +15.5% |
| 4 | 2.6.1 | Final rate = **181.85** | `dsr_rate` = **205.45** | 89 | Rate mismatch, +13.0% |
| 5 | 2.7.1 | Final rate = **352.45** | `dsr_rate` = **412.95** | 90 | Rate mismatch, +17.2% |
| 6 | 2.10.1.1 | Final rate (per m) = **223.00** | `dsr_rate` = **255.55** | 94 | Rate mismatch, +14.6% |
| 7 | 2.10.1.2 | Final rate (per m) = **364.20** | `dsr_rate` = **417.35** | 94 | Rate mismatch, +14.6% |
| 8 | 2.25 | Final rate = **219.65** | `dsr_rate` = **253.95** | 125 | Rate mismatch, +15.6% |
| 9 | 2.27 | Final rate = **1953.05** | `dsr_rate` = **2161.20** | 126–127 | Rate mismatch, +10.7% |
| 10 | 2.28.1 | Final rate = **24.35** | `dsr_rate` = **28.15** | 127 | Rate mismatch, +15.6% |
| 11 | 2.29.1 | Final rate = **24.65** | `dsr_rate` = **28.50** | 127–128 | Rate mismatch, +15.6% |
| 12 | 2.16.1 material line, code 1198 | Qty **21.375** units of "**10 cudm**" @ rate **260.00** (= 0.21375 cum, amount 5557.50) | script uses unit **"cum"**, qty **0.1050**, rate **12000.00** (amount 1260.00) | 103–104 | Unit + qty + rate all wrong; line cost understated ~77% |
| 13 | 2.16.1 material line, code 1197 | Qty **7.50** × "10 cudm" @ 260.00 = 1950.00 | script: cum, qty 0.0358, rate 12000.00 = 429.60 | 104 | Same pattern as above |
| 14 | 2.16.1 material line, code 0302 | Qty **3.1875 metre** @ 40.00 = 127.50 | script: unit "each", qty 0.9000 @ 40.00 = 36.00 | 103 | Unit + qty wrong |
| 15 | 2.16.1 CARRIAGE, code 2204 | Qty **1.3125 cum** @ 118.59 = 155.65 | script: qty 0.1408 @ 118.59 = 16.70 | 104 | Qty wrong by ~9x |
| 16 | 2.16.1 LABOUR, Carpenter 2nd class | PDF 2.16.1 = 0.50 day (2.16.3 = 1.50 day) | script = 1.500 day | 104 | Matches a *different* sub-item (2.16.3), not 2.16.1 |
| 17 | 2.16.1 LABOUR, Beldar | PDF 2.16.1 = 1.00 day (2.16.2=2.00, 2.16.3=4.00) | script = 0.750 day | 104 | Does not match any of the three PDF depth-variants |
| 18 | earthwork_pdf_specs.json coverage | PDF contains items **2.35.3, 2.35.3.1, 2.35.4, 2.35.4.1, 2.35.5, 2.36, 2.37, 2.38** (pages 129–131) | These 8 codes are **absent** from `earthwork_pdf_specs.json`, although 2.35.3.1, 2.35.4.1, 2.36, 2.37, 2.38 are used as item IDs in the ITEMS list | 129–131 | Coverage gap — sheet headers render with blank description for these items |

### Items found clean (norms/rate matched PDF exactly)

2.2.1, 2.3.1, 2.6.1, 2.7.1, 2.10.1.1/1.2, 2.25, 2.25(a), 2.27, 2.28.1, 2.29.1 all had their **LABOUR/MACHINERY/MATERIAL section quantities and base wage/material rates** match the PDF's rate-analysis breakdown exactly — only the top-line `dsr_rate` field diverged. This means the underlying component data (which drives the sheet's cost buildup) is largely trustworthy for the excavation-family items sampled; it is the summary `dsr_rate` figure, and the close-timbering (2.16.x) items' component data, that are wrong.

## Verdicts

**rates_master_clean.json:** Reliable, 100%-accurate on this 50-code sample (spanning the full 0001–2399 range including both extremes). No PDF-deletion-blocking issue found in this slice.

**Chapter 02 Earthwork (support_builder_earth_v2.py ITEMS + earthwork_pdf_specs.json):** **NOT a reliable, 100%-accurate replacement for the PDF**, based on this sample. Two independent classes of error were found:
1. A **systematic ~11–17% overstatement** of the final `dsr_rate` field across every excavation-type item checked (11/11 items affected), even though the underlying labour/material norms were correct — indicating a bug in whatever formula/overhead assumptions produced `dsr_rate`, not a transcription slip.
2. **Materially fabricated or wrongly-sourced component data** for the close-timbering items (2.16.1 checked in full: 4 of its 6 line items have wrong units, quantities, and/or rates, understating true cost by roughly 45–55% overall for that item), plus an 8-code coverage gap in `earthwork_pdf_specs.json`.

Given these findings are concentrated in the derived/summary fields and one full item family (not evenly spread and not yet checked across all ~90 Ch02 items), the PDF for Chapter 02 **should not be deleted**; the earthwork support-sheet pipeline needs correction and a full item-by-item re-verification before this chapter's derived JSON/script data can be trusted as a PDF replacement.
