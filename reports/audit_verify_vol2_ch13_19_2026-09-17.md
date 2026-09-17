# Independent Verification Audit — CivilDAR_2019_Vol_2.pdf, Chapters 13–19

Date: 2026-09-17
Scope: `data/raw_pdf/CivilDAR_2019_Vol_2.pdf` (pages 0-indexed 9–855, Ch.13–19) vs `data/reference_json/ch13_items.json` … `ch19_items.json`.
Question under test: can `CivilDAR_2019_Vol_2.pdf` be deleted/relocated because the JSON is a 100%-accurate replacement?

## Methodology

This audit did **not** re-run or trust `scripts/vol2_chapter_extractor.py`'s own output. A separate script was written (`extract_pages.py`, pdfplumber, default settings) to dump raw per-page text for Ch.13–19 independently, and the resulting text files were read and compared by hand against the delivered JSON, line by line, for each sampled item: item code, description, unit, basis, "Say" rate, and every resource row (code/desc/unit/qty/rate).

Chapter boundary check: for every chapter, the first ~150 characters of the declared start page and the text of the declared end-1 page were printed and read. All seven `CHAPTER_PAGES` boundaries (13:9–109, 14:109–179, 15:179–227, 16:227–371, 17:371–485, 18:485–753, 19:753–855) line up correctly with "SUB HEAD : NN.0" headings — no boundary errors.

Sampling: Ch13 (10 items), Ch14 (10), Ch15 (9), Ch16 (16), Ch17 (10), Ch18 (18, proportional to its 268-page/547-item size), Ch19 (9) — evenly spread by index plus first/last, spot-checked against raw PDF text and page numbers. In addition, once a systemic pattern was suspected, every chapter's raw text was searched with regex to quantify how often that pattern occurs, giving population-level (not just sample-level) counts for the two most serious defects below.

## Findings

### 1. SEVERE — cross-reference lines mistaken for new item headers (data loss + corruption)

The extractor's item-header regex (`^13\.\d+(\.\d+)?\s+...`, one per chapter) matches **any** line starting with a chapter-numbered code — including internal cost-buildup lines like `"18.24 Rate as per item No. 18.24 Of SH: Water Supply"` that appear *inside* a later item's MATERIAL/LABOUR block as a cross-reference to an earlier item's rate. Because these lines also start with `NN.dd` at the left margin, they are treated as new item boundaries. This:
- **Deletes real items outright** (no "Say" value falls between two spurious header-matches, so the fragment is silently discarded), and
- **Injects garbage duplicate-coded items** with the wrong code, a description that is just the cross-reference text, and an incomplete resource list (only the rows after the false split point).

Confirmed concretely:
- **Ch18**: item **18.25.1** ("Providing and laying S&S C.I. standard specials … Up to 300 mm dia", Say ₹5,320.40, p.547) and **18.25.2** (Say ₹5,454.20, p.548), **18.26.1** (Say ₹7,862.50, p.548), **18.26.2** (Say ₹7,862.50, p.549), **18.27.1** ("100 mm dia pipe", p.549) are **completely absent** from `ch18_items.json` — confirmed by direct code search (0 hits for all five). In their place, the JSON contains dozens of corrupted duplicate entries under the wrong codes `"18.23"` (62 occurrences) and `"18.24"` (8 occurrences), each carrying a truncated description like `"Rate as per item No. 18.24 Of..."` instead of the real item text. Chapter-wide: the raw PDF text contains **125** lines matching this trap pattern in Ch18 alone; `ch18_items.json` contains **111 duplicate/garbage entries** across 10 distinct mis-captured codes out of 547 total "items" — roughly **20% of the file's entries are not real catalog items**, and an unknown but comparably-sized set of real items is missing.
- **Ch16**: item **16.10** ("Making bajri path including preparation of subgrade…", Say ₹172.45, p.240/241) is **completely absent**. It is split by two internal cross-refs (`16.3.8`, `16.3.10`) into three fragments; the first two are discarded (no Say line before the next false header), and the third survives mislabeled as code `16.3.10` with a wrong description and only 6 of the item's true ~10 resource rows. The same mechanism silently deletes items in the **16.13/16.14** range. Chapter-wide: 11 trap occurrences; codes `16.3.10` (5x) and `16.16` (2x) appear duplicated with garbage content in `ch16_items.json`.
- **Ch19**: 1 trap occurrence found (chapter mostly unaffected).
- **Ch13, Ch14, Ch15, Ch17**: 0 trap occurrences — these four chapters are **not** affected by this specific bug.

### 2. MODERATE/SEVERE — resource rows dropped when a material/labour description wraps onto 2+ lines

The resource-row regex requires the code, description, unit, quantity and rate to appear as one matchable span without a line break inside the description. Real DAR entries routinely wrap long material descriptions across 2–7 lines. When they do, the resource row is silently dropped from the JSON's `resources` array, even though the item's final "Say" rate (read independently from the "Say" line) still comes through correctly.

Confirmed: **17.81** ("floor mounted WC…", Say ₹15,457.95, p.1370) is missing resource code `1966` — the single largest cost component (₹9,500 material, ~61% of the item's total), captured in JSON with only 4 of 5 resource rows. **13.91** and **14.76** and **14.94** each are missing one resource row (a `9999` sundries/T&P line whose description wraps). **19.33** ("soak pit", Say ₹2,608.00) is missing 3 of its 7 true resource rows.

### 3. MODERATE — lower-case/abbreviated cross-reference phrasing not captured

`ref_pat` only matches the phrase `"Rate as per Item Number"` (capitalized). The PDF also uses the lower-case, abbreviated variant `"Rate as per item no."`, which is common in Ch19 and elsewhere. These cross-referenced resource rows are silently dropped. Confirmed in **19.33**: both `2.8.1` (earthwork, ₹436.48) and cross-chapter reference `16.8.1` (brick edging, ₹243.10) are missing from its resources.

### 4. MODERATE — "Details of cost for" (plural) not matched by the unit/basis regex

The regex expects singular `"Detail of cost for N unit"`; the PDF very frequently uses the plural `"Details of cost for N unit"`. When it does, both `unit` (defaults to `"nos"`) and `basis` (defaults to `1.0`) are wrong unless they happen to coincide. This is the single most widespread defect by volume — the plural phrasing is the **majority** form in most chapters: Ch14 66/95 (70%), Ch15 4/92 (4%), Ch16 131/140 (94%), Ch17 202/227 (89%), Ch18 544/556 (98%), Ch19 91/104 (88%); Ch13 is only 11/166 (7%) affected. Confirmed wrong `unit`/`basis` in **13.91** (basis should be 10, JSON shows 1.0), **14.76** (basis should be 10, JSON shows 1.0), **14.94** (unit should be "job", JSON shows "nos"), **16.3.8** (unit should be "cum", JSON shows "nos"). The headline "Say" rate is unaffected, but `unit`/`basis` metadata is unreliable for a large fraction of items in Ch14, 16, 17, 18, 19.

### What was NOT wrong
In every single sampled item across all 7 chapters (~80 items directly inspected against the raw PDF text), the **"Say" rate** — the number that actually matters for costing — matched the PDF exactly, including in the corrupted duplicate-code entries described in Finding 1 (the number itself was right, just filed under the wrong item code). Descriptions, units, and resource rows that were captured (not silently dropped) matched the PDF verbatim in every case checked.

## Per-chapter summary

| Ch | Items sampled | Trap-bug hits (pop.) | Confirmed missing items | Confirmed corrupted/dup items | Other issues |
|----|---|---|---|---|---|
| 13 | 10 | 0 | 0 | 0 | basis-default bug (11/166 lines, low impact) |
| 14 | 10 | 0 | 0 | 0 | basis/unit-default bug (66/95 lines), dropped resource rows |
| 15 | 9  | 0 | 0 | 0 | unit-label ambiguity on dual-unit item (15.20) |
| 16 | 16 | 11 | 16.10 confirmed missing (+ likely 16.13/16.14) | 7 (16.3.10 x5, 16.16 x2) | basis/unit-default bug (131/140 lines) |
| 17 | 10 | 0 | 0 | 0 | dropped largest resource row in 17.81; basis/unit-default bug (202/227 lines) |
| 18 | 18 | 125 | 18.25.1, 18.25.2, 18.26.1, 18.26.2, 18.27.1 confirmed missing (many more likely) | 111 garbage entries (~20% of file) | basis/unit-default bug (544/556 lines) |
| 19 | 9  | 1 | 0 confirmed | 0 | dropped resource rows (19.33), lowercase cross-ref not captured, basis/unit-default bug (91/104 lines) |

## Verdict

Is Ch13–19 extracted data a reliable 100%-accurate replacement for the raw PDF, based on this sample? **No.** The "Say" unit rate — the number most users actually need — proved accurate in every item sampled, so for simple rate lookups the JSON is largely trustworthy. But the JSON is demonstrably **not** a complete or fully-accurate substitute for the PDF: Chapter 18 has a confirmed, structural data-corruption bug that deletes real items (at minimum 5 confirmed, likely several dozen more given 125 trigger occurrences and 111 garbage entries already found) and replaces them with mislabeled duplicate entries; Chapter 16 has the same bug at smaller scale with at least one item (16.10) fully missing; resource-level detail (the itemized cost buildup) is unreliable across most of Ch14–19 due to silently dropped rows and wrong unit/basis defaults, most severely in Ch17 (17.81 is missing its single largest cost line) and Ch18/19. Chapters 13, 14, 15, and 17 are free of the item-deletion bug specifically, but still carry the resource-row and unit/basis defects. **The raw PDF should be retained** at least until Ch16 and Ch18 are re-extracted with a fixed parser (the header regex must not match cross-reference lines inside a block, the description regex must tolerate multi-line wraps, and the unit/basis and cross-reference regexes must accept the plural/lower-case phrasings documented above) and the fix is re-verified.
