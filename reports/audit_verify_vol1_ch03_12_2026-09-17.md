# Independent PDF Verification Audit — Vol.1 Chapters 03–12 (2026-09-17)

## Methodology

This audit bypasses the pipeline entirely. `data/raw_pdf/CivilDAR_2019_Vol_1.pdf` (904 pages) was
extracted independently with `pdfplumber` (text-per-page), and chapter boundaries were located by
searching for `SUB HEAD : N` markers, cross-checked against the printed page numbers in the PDF
footer. Resulting page ranges (0-indexed in the extracted page array):

| Ch | Title | Pages |
|----|-------|-------|
| 03 | Mortars | 141–150 |
| 04 | Concrete Work | 151–186 |
| 05 | RCC Work | 187–278 |
| 06 | Masonry Work | 279–320 |
| 07 | Stone Work | 321–364 |
| 08 | Cladding Work | 365–424 |
| 09 | Wood and PVC Work | 425–672 |
| 10 | Steel Work | 673–718 |
| 11 | Flooring | 719–810 |
| 12 | Roofing | 811–901 |

A regex-based parser identified all item headers (`N.n[.n…][A-F]`) that are followed by a genuine
priced resource table (contains a "Say ..." final rate line, to exclude bare group headings like
`4.1` that only introduce sub-items). This produced a per-chapter set of real PDF item codes, which
was diffed against the `code` field in each `data/reference_json/ch0N_items.json`. Every diff
candidate was then manually verified by reading the raw PDF text around that page to rule out
parser artifacts (e.g. numbers like "12.5 mm thick" or "11.97 kg" mid-sentence, which are not item
codes). For chapters/items sampled per the brief, resource rows (code/desc/unit/qty/rate) and the
final "Say" value were compared line-by-line against the JSON. `labour_productivity.json` and
`sundries_reference.json` were spot-checked for the same items (e.g. Beldar 558.00/day, Bhisti
617.00/day, mortar sundries lines) and found consistent with the PDF wherever present.

## Confirm/Refute on the two known prior gaps

**(a) Mortar item 3.19 — CONFIRMED, and worse than "missing."**
PDF page 150 has a genuine, fully-priced item `3.19 "Mortar in lime, surkhi (50% red and 50% light
yellow) and marble dust 1:1.5:0.5"` (Say 2821.15) directly after `3.18 "Mud mortar"` (Say 737.00).
`ch03_items.json` has only 18 entries ending at code `"3.18"` — but that JSON entry is not clean
data for 3.18: it is a **corrupted merge** of both items' resource rows (it contains both "Mud
(dry)" from item 3.18 AND "Surkhi"/"Unslaked lime"/"Marble dust" etc. from item 3.19), and its
`"say": 2821.15` is item 3.19's total, not item 3.18's true total of 737.00. So the true 3.18 record
is wrong and item 3.19 is entirely absent.

**(b) Ch.05 RCC Work missing a `support_builder` script — REFUTED (already fixed today).**
`scripts/support_builder_rcc.py` exists (60KB). `git log` shows it was added in commit `178dba1`,
"feat: add Ch.05 RCC support builder and rebuild 05_RCC_Work sheet", dated 2026-09-17 — the same
day as this audit. The gap identified by the prior audit has already been closed in this repo.
However, see below: chapter 5's item-code *coverage* itself still has serious independent gaps.

## Per-chapter findings

### Ch.03 Mortars — 1 discrepancy (see above), otherwise clean
Sampled 3.1–3.18 resource rows/rates against JSON: item 3.1 (Cement mortar 1:1) matches the PDF
exactly (Portland Cement 1.02 t @ 4940, Fine Sand 0.7125 cum @ 900, Beldar 0.75 day @ 558, Bhisti
0.07 day @ 617, Say 6390.60). Items 3.2–3.17 spot-checked and consistent. `sundries_reference.json`
has no `3.19` row either (consistent with the coverage gap).

### Ch.04 Concrete Work — none found
PDF has 31 real priced items (4.1.2, 4.1.4–4.1.13, 4.2.x, 4.3.x, etc. — `4.1`, `4.2`, `4.3` etc. are
bare group headings with no own resource table, correctly absent from JSON). All 31 real item codes
are present in `ch04_items.json` (which has 37 entries; the extra 6 are legitimate cross-chapter
"Rate as per Item No. X of SH: ..." references pulled in during conversion, not concrete-work items
proper). Sampled 4.1.2, 4.1.5, 4.1.11 rates — match PDF.

### Ch.05 RCC Work — 18 missing items (major gap)
The largest gap found in this audit. PDF items confirmed present but **absent from
`ch05_items.json`**:
- `5.4` "RCC kerbs..." (page 192), `5.7` "RCC well-steining..." (page 194) — both fully priced.
- `5.12`–`5.17` (6 items): precast RCC string courses/bands/copings, small lintels, mouldings,
  chajjas/etc. (pages 222–227), all fully priced (e.g. 5.12 Say 8886.35).
- `5.22A` steel reinforcement **above plinth level**, with 6 priced sub-items `5.22A.1`–`5.22A.6`
  (pages 235–238) — entirely missing. JSON's `5.22.1`–`5.22.6` only cover the **below-plinth**
  family (`5.22`, page 231); the above-plinth family is a completely separate, equally-priced set
  of items and none of it made it into the JSON.
- `5.22B` and `5.22C` — ready-to-use "cut and bend" rebar items (pages 238–239), each with one
  priced sub-item — also entirely missing.
- `5.48X`, `5.48Y`, `5.48Z` — annexure sub-analyses (page 271–272) — missing.
- `5.66` — referenced calculation line near page 204 (lower confidence; likely a genuine item,
  not independently confirmed with a full resource table due to time constraints — flagged for
  follow-up rather than asserted).

Net: of ~93 real priced PDF items in Ch.05, at least 17 (18 counting 5.66 unconfirmed) are absent —
roughly 18% of the chapter, including the entire above-plinth steel-reinforcement family, one of the
highest-value, most frequently used items in RCC estimating.

### Ch.06 Masonry Work — 1 missing item
`6.5` "Extra for brick work / AAC block masonry / Tile brick masonry in superstructure..." (page
286) is a real priced item absent from `ch06_items.json`. All other sampled items (6.1, first item;
6.45 and others across the range) matched.

### Ch.07 Stone Work — 1 missing item
`7.19` "Extra for additional cost of centering for arches exceeding 6m span..." (page 347) is a real
priced item absent from `ch07_items.json`. (`7.00` flagged by the automated pass was a false
positive — a page-footer/header artifact, not an item.) Sampled first/last and mid-range items
otherwise consistent.

### Ch.08 Cladding Work — none found
All PDF item codes with priced tables are present in `ch08_items.json`.

### Ch.09 Wood and PVC Work — 25+ missing items (major gap)
- `9.23`, `9.26`, `9.110`, `9.161` — individually confirmed real, priced, missing items (teak
  lipping extra, rebate-cutting extra, bamboo jaffery, fire-resistant door frame).
- **Entire `9.147A`–`9.147F` uPVC window/door family — 20 priced sub-items — completely absent**
  from `ch09_items.json` (verified `grep` for any "147" code in the JSON returns zero hits). PDF
  pages 625–651 contain fully priced items for uPVC casement windows/doors, fixed glazed
  windows/ventilators, sliding windows and sliding doors in multiple size/frame variants. None of
  this family exists in the JSON at all — a wholesale missing product category, not a a stray item.

### Ch.10 Steel Work — none found (after false-positive filtering)
The one automated "extra" (`10.11.1`) is a legitimate JSON sub-item; no real PDF item is missing.
Sampled first (10.1), last, and several mid-chapter items — resource rows matched.

### Ch.11 Flooring — 2 missing items
`11.37A` "1st quality ceramic glazed floor tiles..." and `11.46A` "glazed screen printed border
tile..." (pages 781, 794) are real, distinct, priced items (siblings of 11.37 and 11.46, which ARE
present) that are absent from `ch11_items.json`. (`11.97` flagged by the automated pass was a false
positive — "11.97 kg" is a quantity inside item 11.35's calculation, not an item code.)

### Ch.12 Roofing — 1 missing item
`12.52.1` "GI Metal Ceiling Lay in plain Tegular edge Global white colour tiles..." (page 886) is
missing while its siblings `12.52`, `12.52.2`, `12.52.3`, `12.52.4` are all present in
`ch12_items.json` — an isolated drop within an otherwise-captured family. (`12.5` and `12.125`
flagged by the automated pass were false positives — "12.5 mm thick" text fragments, not item
codes; real sub-items `12.5.1` and `12.6.1` are present and correct.)

## Verdict

**No, the Ch.03–12 extracted JSON (via the unverified converted-XLSX intermediate) is NOT a
reliable 100%-accurate replacement for the raw PDF, based on this sample.**

Beyond the already-known mortar item 3.19 gap (which is actually a data-corruption case, not a
simple omission — the surviving "3.18" record silently contains merged/wrong data), this audit
independently found **at least 30 additional missing priced items** across 6 of the 10 chapters
sampled (05, 06, 07, 09, 11, 12), including one wholesale missing product family (uPVC
windows/doors, 20 items in Ch.09) and one missing item family for a structurally critical, high-
value component (above-plinth steel reinforcement + ready-to-use rebar, 8 items in Ch.05). Where
items ARE present, their resource-row data (codes, quantities, rates, and final "Say" totals)
matched the PDF closely in every case sampled — so the conversion is accurate when it captures an
item, but its **coverage** is not complete. The RCC support-builder script gap (prior finding "b")
has already been fixed today and is no longer an issue. Given the volume and severity of coverage
gaps found in a relatively small sample, the raw PDF should be retained until the converted-XLSX
intermediate is re-derived directly from the PDF (or thoroughly reconciled against it) rather than
deleted or relocated.
