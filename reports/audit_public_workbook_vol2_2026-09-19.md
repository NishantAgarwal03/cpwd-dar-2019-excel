# Audit: CPWD_DAR_2019_Vol2_PDF_Conversion.xlsx (2026-09-19)

**Scope.** Exhaustive, code-driven audit of the public "PDF to Excel conversion" workbook for
Volume 2 (Ch13-26), checked purely against this repo's own already-verified source data
(`data/reference_json/ch13_items.json` ... `ch26_items.json`) and for internal structural/formula
soundness. No PDF was opened for this task, per instructions.

## Methodology

1. Read `scripts/generate_public_conversion_workbook.py` to get the exact expected layout of each
   item block: item header row (bold code + merged B:F desc), "Rate basis: per {unit} (source: cost
   for {basis:g} {unit})" note row, a `Code/Description/Unit/Quantity/Rate/Amount` column-header
   row, one row per resource with `=D{row}*E{row}` Amount formula, a "Resource Total" row with
   `=SUM(F{res_start}:F{res_end})`, a "SAY (Rs.) — final DAR rate" static-value row, then a blank
   separator. Volume 2 has no percentage-item branch and no chapter exclusions, so every item in
   every `ch{13-26}_items.json` was expected verbatim.
2. Wrote a Python/openpyxl script (`audit_vol2.py`) that parses every one of the 14 chapter sheets
   into item blocks by walking this exact row grammar, then for **every one of the 1697 source
   items**: matched by code, compared description, the basis-note text, every resource row
   (code/desc/unit/qty/rate, in order) and its Amount formula, and the Resource Total formula's
   range, plus the SAY value (float-tolerant compare). Also scanned every cell in every sheet for
   Excel error tokens, literal `"None"` text, raw Python repr artifacts, and Unicode replacement
   characters.
3. A second script (`audit_vol2_part2.py`) parsed the Cover & Index sheet: for each of the 14
   contents rows, confirmed the target sheet exists with the exact name, the hyperlink is exactly
   `#'sheet_name'!A1`, and the printed item count matches an independent recount of that sheet's
   actual item blocks. It also spot-checked number formats (15 items per chapter, or all items in
   chapters with fewer) on Quantity/Rate/Amount/SAY cells for non-"General" formatting.
4. Cross-reference resource rows (desc starting with `Rate as per SH:` or `Sub-analysis annexure:`,
   769 such rows in Vol 2) were not special-cased anywhere in the generator, so they were verified
   as part of the exhaustive resource-row check in step 2 (same code/desc/unit/qty/rate/formula
   comparison as any other resource row).

## Completeness-check summary (all 1697 items, all 14 sheets)

| Ch | Sheet | Items expected | Items found | Items fully matched |
|----|-------|----------------|-------------|----------------------|
| 13 | 13_Finishing | 164 | 164 | 164 |
| 14 | 14_Repairs_to_Buildings | 96 | 96 | 96 |
| 15 | 15_Dismantling_Demolishing | 92 | 92 | 92 |
| 16 | 16_Road_Work | 143 | 143 | 143 |
| 17 | 17_Sanitary_Installations | 217 | 217 | 217 |
| 18 | 18_Water_Supply | 556 | 556 | 556 |
| 19 | 19_Drainage | 123 | 123 | 123 |
| 20 | 20_Pile_Work | 38 | 38 | 38 |
| 21 | 21_Aluminium_Work | 42 | 42 | 42 |
| 22 | 22_Water_Proofing | 26 | 26 | 26 |
| 23 | 23_Rain_Water_Harvesting | 39 | 39 | 39 |
| 24 | 24_Heritage_Buildings | 8 | 8 | 8 |
| 25 | 25_Structural_Glazing | 7 | 7 | 7 |
| 26 | 26_New_Technologies | 146 | 146 | 146 |
| **Total** | | **1697** | **1697** | **1697** |

No missing items, no duplicate item codes, no extra/unexpected items in any sheet.

## Findings by check

1. **Completeness cross-check** — 0 discrepancies. Every source item appears exactly once with
   matching code, description, "Rate basis:" text, and SAY value (exact match, no floating-point
   drift found). Every resource row (code, desc, unit, qty, rate) matches the source in the same
   order.
2. **Formula integrity** — 0 discrepancies across all 1697 items. Every resource row's Amount
   formula is exactly `=D{row}*E{row}` on its own row; every item's Resource Total SUM range
   exactly spans that item's own resource rows with no off-by-one or cross-item overlap.
3. **Structural soundness** — 0 discrepancies. Every item header, "Rate basis:" note, column-header
   row, "Resource Total" label (where resources exist), and "SAY" label contains the expected text
   in all 14 sheets, including the 556-item Ch18 sheet. Cover & Index per-chapter item counts match
   actual sheet counts exactly for all 14 chapters; all 14 hyperlinks are exactly
   `#'sheet_name'!A1` pointing to sheets that exist with the exact expected name.
4. **Number formatting / rendering sanity** — 0 discrepancies. Spot-checked 15 items per chapter
   (or all items where a chapter has fewer than 15, e.g. Ch24/Ch25) — all Quantity/Rate/Amount/SAY
   cells carry non-"General" number formats. A full-cell scan of all 14 sheets found no Excel error
   values (`#REF!`, `#VALUE!`, `#DIV/0!`, etc.), no literal `"None"` text, no raw Python
   dict/tuple/list reprs, and no Unicode replacement characters anywhere in the workbook.
5. **Cross-reference resource rows** — 769 resource rows across Vol 2 have `desc` starting with
   `Rate as per SH:` or `Sub-analysis annexure:`. The generator has no special-case branch for
   these; they were verified identically to all other resource rows in check 1/2 and all render
   correctly (correct qty/rate, correct `=D*E` Amount formula, no corruption).

## Verdict

**Reliable, error-free.** Across all 14 sheets and 1697 items, the audit found **zero
discrepancies** in completeness, formula integrity, structural soundness, number formatting, or
cross-reference resource-row handling. `CPWD_DAR_2019_Vol2_PDF_Conversion.xlsx` faithfully and
mechanically reproduces `ch13_items.json` through `ch26_items.json` and is sound as a public
deliverable in its current form.
