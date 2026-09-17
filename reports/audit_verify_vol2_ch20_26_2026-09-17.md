# Independent Verification Audit — Vol.2 Ch.20–26 vs Raw PDF

Date: 2026-09-17
Scope: `data/raw_pdf/CivilDAR_2019_Vol_2.pdf` pages 855–1110 (0-indexed) vs `data/reference_json/ch{20..26}_items.json`
Question: can the raw PDF be deleted/relocated because the extracted JSON is a 100%-accurate replacement?

## Methodology

Independent Python/pdfplumber extraction was run directly against the raw PDF (not via `scripts/vol2_chapter_extractor.py`), dumping full page text per chapter to scratch files and reading it manually against each `ch{NN}_items.json`. For each chapter I verified: (a) the stated page range actually starts/ends at the `SUB HEAD` heading, (b) a spread of item codes (first, last, and 8–10 mid-range; exhaustive for Ch24/Ch25), checking code, description, unit, basis, "Say" rate, and every resource row, and (c) whether the PDF has any content past page 1110. I also wrote a second, independently-derived broad regex (different from the extractor's) to enumerate every resource-looking line (`CODE desc UNIT qty rate amount`) per chapter and diffed it against what the extractor's own resource regex would catch, to quantify systematic coverage loss rather than relying on one-off spot checks.

## Chapter boundary check

All seven boundaries are correct — the start page of each range contains the `SUB HEAD : NN.0` heading for that chapter, and the page immediately after each range's end already belongs to the next chapter's `SUB HEAD`. No boundary drift found.

## PDF total length check

The PDF has exactly **1110 pages** (indices 0–1109). Printed content ends on page index 1108 (printed page "1994", chapter 26's last item 26.81, `Say 24.20`). Page index 1109 is genuinely blank. **No content beyond page 1110 is being missed** — the pipeline's range correctly reaches the true end of the document.

## Headline "Say" rate accuracy

Every item's final "Say" unit rate that I checked (~20 items spread across all 7 chapters, including first/last items of each) matched the PDF exactly, character for character. This part of the extraction is reliable.

## Discrepancies found (systemic, not isolated typos)

### 1. Unit/basis wrongly defaulted for most items (severe)
The extractor's basis/unit regex only matches the phrase `"Detail of cost for <qty> <unit>"` (singular "Detail"). The PDF predominantly uses **"Details of cost for..."** (plural), plus variant phrasings like `"Details of cost for area 22.5x9.0=202.50 sqm"`. When the regex fails to match, the code silently defaults to `unit="nos", basis=1.0` — which is frequently wrong.

Measured impact (items with unit=nos & basis=1.0 in JSON, out of total items in chapter):
| Chapter | Total items | unit=nos/basis=1.0 (suspect default) |
|---|---|---|
| Ch20 | 34 | 34 (100%) |
| Ch21 | 38 | 32 (84%) |
| Ch22 | 26 | 22 (85%) |
| Ch23 | 31 | 31 (100%) |
| Ch24 | 8 | 1 (13%) |
| Ch25 | 7 | 7 (100%) |
| Ch26 | 111 | 79 (71%) |

Example: item **20.1.1** ("400 mm dia piles", p.858) — PDF says `Details of cost for 20 m length of pile` and `Cost of 1 metre pile 2244.08`. JSON records `unit: "nos", basis: 1.0`. The correct basis is per-metre. This mislabels the pricing basis for nearly every item in Ch20/23/25 and most of Ch21/22/26.

### 2. Resource rows dropped wholesale for material codes outside 0xxx/8xxx/9xxx (severe)
The resource-row regex only matches codes matching `0\d{3}` or `[89]\d{3}`. Material codes starting with 1, 2, 3, 4, 6, or 7 (very common — steel, cement, carriage, sand/aggregate, and most specialty/proprietary materials) are never captured at all.

Independently counted distinct resource codes per chapter (broad regex) vs. codes the extractor's own pattern would ever match:
| Chapter | Distinct resource codes (PDF) | Matched by extractor pattern | Missed |
|---|---|---|---|
| Ch20 | 13 | 10 | 3 (23%) |
| Ch21 | 23 | 17 | 6 (26%) |
| Ch22 | 33 | 25 | 8 (24%) |
| Ch23 | 32 | 9 | 23 (72%) |
| Ch24 | 21 | 9 | 12 (57%) |
| Ch25 | 26 | 4 | 22 (85%) |
| Ch26 | 99 | 69 | 30 (30%) |

Concrete example — item **24.2** ("double scaffolding system", p.973–975): PDF lists 11 resource lines (materials 7397, 4009, 7387, 1034, 7346, 7398, 7399, 2205, plus labour 0116, 0114, and 9999 Sundries). `ch24_items.json` for 24.2 contains only **3** resources (0116, 0114, 9999) — every material line (≈₹8,223 of the ₹39,040 pre-markup total) is missing, and `unit`/`basis` are also wrong (`nos`/1.0 instead of `sqm`/202.5, from "Details of cost for area 22.5x9.0=202.50 sqm").

### 3. Multi-line resource descriptions dropped even for valid 0/8/9 codes
The resource regex requires the whole `code + description + unit + qty + rate` to sit on one physical text line. When a description wraps across multiple lines (common for machinery/material items with long descriptions), the row is dropped even though its code would otherwise match.
Example: item **20.1.1** — resource `0024 Hire and running charges of hydraulic piling rig with power unit… day 0.36 30000.00 10800.00` (description wraps 4 lines) is completely absent from `ch20_items.json`'s resources for 20.1.1, despite code `0024` matching the intended pattern.

### 4. Cross-referenced sub-head rates missed by case/wording mismatch
The cross-reference pattern requires `"Rate as per Item Number"` (title case). The PDF frequently uses `"Rate as per item no"` (lower case, abbreviated). Example: item 20.1.1's line `5.33.1 Rate as per item no 5.33.1 of SH : RCC Work cum 2.51 7997.30 20073.22` — the single largest cost component of that item (52% of its pre-markup total) — is entirely absent from the JSON's resources.

### 5. Missing items / code collisions from nested sub-item numbering
The item-header regex only matches up to two dotted sub-levels after the chapter number (`NN.d+(.d+)?`), so genuine 4-level codes like `21.1.1.1` / `21.1.1.2` never match as their own headers. The first nested sub-item's data gets merged into the parent 3-level code's entry (e.g. `21.1.1`'s `desc` field in the JSON reads "For fixed portion 21.1.1.1 Anodised aluminium…", concatenating two different things), and the sibling sub-item **21.1.1.2 ("Powder coated aluminium…") is missing from `ch21_items.json` entirely** — a genuine coverage gap, not a value error.
Separately, a cross-reference sentence elsewhere in Ch21's text ("Rate as per Item Number 21.4.1 of SH:Aluminium Work…") spuriously matches the item-header pattern too, so **code `21.4.1` appears twice** in `ch21_items.json` — once as the real item, once as a bogus duplicate built from the cross-reference sentence.

## Summary discrepancy table

| Chapter | Boundary OK | Say-rate spot checks OK | Unit/basis systemic issue | Resource coverage gap | Missing/duplicate items found |
|---|---|---|---|---|---|
| Ch20 | Yes | Yes | Yes (100% default) | Yes (23% of codes + wrapped lines + cross-ref) | Not sampled beyond resources |
| Ch21 | Yes | Yes | Yes (84%) | Yes (26%) | Yes — 21.1.1.2 missing, 21.4.1 duplicated |
| Ch22 | Yes | Yes | Yes (85%) | Yes (24%) | Not exhaustively checked |
| Ch23 | Yes | Yes | Yes (100%) | Yes (72%) | Not exhaustively checked |
| Ch24 | Yes | Yes | Minor (13%) | Yes (57%), confirmed on item 24.2 | None found (exhaustive check) |
| Ch25 | Yes | Yes | Yes (100%) | Yes (85%) | Not exhaustively checked |
| Ch26 | Yes | Yes | Yes (71%) | Yes (30%) | Not exhaustively checked |

## Final verdict

**Is Ch20–26 extracted data a reliable 100%-accurate replacement for the raw PDF, based on this sample? No.** The chapter page boundaries are correct, the PDF's true end (page 1110/index 1109, blank) is fully within the captured range with nothing missed beyond it, and the headline "Say" unit rate for each item is accurate wherever it was captured. However, the JSON is **not** a safe substitute for the source PDF: (1) the `unit`/`basis` fields default incorrectly to `"nos"`/`1.0` for the majority of items in most chapters because the extractor's regex only recognizes the less-common singular phrasing "Detail of cost for"; (2) the `resources` breakdown arrays are systematically incomplete — anywhere from ~23% to ~85% of genuine material/resource line items per chapter are silently dropped because the resource regex only recognizes codes starting with 0, 8, or 9, cannot span multi-line descriptions, and misses lower-case/abbreviated cross-reference phrasing; and (3) at least one genuine missing item (21.1.1.2) and one duplicate/bogus item (21.4.1) were found from a nested-numbering blind spot in the item-header regex. Anyone relying on the JSON for cost breakdowns, unit verification, or a complete item listing would get materially wrong or incomplete answers even though the top-line rate looks right. **The original PDF should be retained** until the extractor is fixed (plural "Details of cost for" phrasing, full-digit resource codes, multi-line description handling, lower-case cross-reference phrasing, and full-depth item-code numbering) and Ch20–26 are re-extracted and re-verified.
