# CPWD DAR 2019 — Data Pipeline

## Canonical flow

```
raw_pdf/
  CivilDAR_2019_Vol_1.pdf ─── extract_rates.py ────────────────────► reference_json/rates_master_clean.json
                          └── vol1_chapter_extractor_pdf.py ───────► ch01_items.json, ch03_items.json … ch12_items.json
                                                                              │
                              support_builder_earth_v2.py (ch02 ITEMS)       │
                                           │                                 │
                                           └──────────────┬──────────────────┘
                                                           ▼
                                                  extract_norms.py
                                                           │
                                                           ▼
                                                 labour_productivity.json
                                                 sundries_reference.json

  CivilDAR_2019_Vol_2.pdf ─── vol2_chapter_extractor.py ───────────► ch13_items.json … ch26_items.json
```

`data/converted_xlsx/CivilDAR_2019_Vol_1_Converted.xlsx` is no longer read anywhere in this pipeline
(retired 2026-09-18 — it had no in-repo provenance and the old path built from it had material coverage
gaps; see `reports/audit_*_2026-09-17.md`). Every Vol 1 chapter now derives from the raw PDF, either via
`vol1_chapter_extractor_pdf.py` (Ch01, 03–12) or the hand-verified `support_builder_earth_v2.py` `ITEMS`
list (Ch02, `dsr_rate` literals cross-checked against the PDF 2026-09-18).

> Ch01 "Carriage of Materials" also contains large lead-vs-distance lookup tables with no per-item "Say"
> line; those are intentionally skipped — only its genuine priced cost-buildup items (e.g. 1.3, 1.4.1–1.4.3)
> are extracted.

```
ch02: support_builder_earth_v2.py (ITEMS list)  ┐
ch03–12: ch03_items.json … ch12_items.json      ├─ cpwd_item_parser.py ──► gang_registry_vol1.txt/.ini
                                                │                          productivity_registry_vol1.txt/.ini
ch13–26: ch13_items.json … ch26_items.json      ┘─ cpwd_item_parser_vol2.py ► gang_registry_vol2.txt/.ini
                                                                               productivity_registry_vol2.txt/.ini

rates_master_clean.json  ┐
labour_productivity.json ├─ main.py ──────────────────────────────► CPWD_DAR_2019_Vol1_Workbook.xlsx
sundries_reference.json  │
ch03–12_items.json       ┘

ch13–26_items.json ─────────── generate_vol2_workbook.py ─────────► CPWD_DAR_2019_Vol2_Workbook.xlsx
```

## Run order

Use `run_pipeline.py` — it handles dependencies in the correct order:

```
python run_pipeline.py --list           # show all steps
python run_pipeline.py                  # run everything
python run_pipeline.py --step 5         # rebuild Vol 1 registry only
python run_pipeline.py --from 5         # rebuild registries + workbooks
```

| Step | Script | Output |
|------|--------|--------|
| 1 | `extract_rates.py` | `rates_master_clean.json` |
| 2 | `vol1_chapter_extractor_pdf.py` | `ch01_items.json`, `ch03–12_items.json` |
| 3 | `extract_norms.py` | `labour_productivity.json`, `sundries_reference.json` (reads step 2's output + Ch02 `ITEMS`) |
| 4 | `vol2_chapter_extractor.py` | `ch13–26_items.json` |
| 5 | `cpwd_item_parser.py` | `gang_registry_vol1.*`, `productivity_registry_vol1.*` |
| 6 | `cpwd_item_parser_vol2.py` | `gang_registry_vol2.*`, `productivity_registry_vol2.*` |
| 7 | `main.py` | `Vol1_Workbook.xlsx` |
| 8 | `generate_vol2_workbook.py` | `Vol2_Workbook.xlsx` |

**Note:** step 3 depends on step 2's output, so `vol1_chapter_extractor_pdf.py` must run before
`extract_norms.py` — the reverse of the pre-2026-09-18 order, when both independently read the same
converted XLSX.

## Single source of truth per chapter

| Chapter | Canonical source | Used by |
|---------|-----------------|---------|
| Ch 01 Carriage of Materials | `ch01_items.json`, extracted directly from the PDF by `vol1_chapter_extractor_pdf.py` | `extract_norms.py` (step 3) |
| Ch 02 Earth Work | `support_builder_earth_v2.py` — `ITEMS` list (`dsr_rate` literals verified against PDF 2026-09-18) | registry (step 5), workbook (step 7), `extract_norms.py` (step 3) |
| Ch 03–12 Vol 1 | `ch03_items.json` … `ch12_items.json`, extracted directly from the PDF by `vol1_chapter_extractor_pdf.py` | registry (step 5), `extract_norms.py` (step 3) |
| Ch 13–26 Vol 2 | `ch13_items.json` … `ch26_items.json`, extracted directly from the PDF by `vol2_chapter_extractor.py` | registry (step 6), workbook (step 8) |

## Registry files

| File | Covers | Generator |
|------|--------|-----------|
| `data/gang_registry_vol1.txt` | Ch 02–12 | `cpwd_item_parser.py` |
| `data/productivity_registry_vol1.txt` | Ch 02–12 | `cpwd_item_parser.py` |
| `data/gang_registry_vol2.txt` | Ch 13–26 | `cpwd_item_parser_vol2.py` |
| `data/productivity_registry_vol2.txt` | Ch 13–26 | `cpwd_item_parser_vol2.py` |

INI versions of each file are identical in data — they exist for tooling that prefers INI format.

## Known unverifiable data (kept as-is, by user decision 2026-09-18)

Five Ch02 ".2" "Hard soil" variant items in `support_builder_earth_v2.py`'s `ITEMS` —
`2.1.2`, `2.2.2`, `2.3.2`, `2.6.2`, `2.8.2` — do not appear anywhere in `CivilDAR_2019_Vol_1.pdf`
(searched all 904 pages) or in `DSR_Vol1_UPDATED_DEC_2021.pdf` / `DSR_Vol2_UPDATED_DEC_2021.pdf`
(also in `data/raw_pdf/`). Those two DSR files confirmed DSR 2021 was the actual, undocumented
source of several *other* `dsr_rate` literals that were wrong before the 2026-09-18 fix (DSR 2021's
Chapter 2 values for 2.1.1/2.2.1/etc. exactly matched the pre-fix, incorrect numbers) — but DSR 2021's
Chapter 2 has the identical "All kinds of soil only, no Hard-soil variant" structure as DAR 2019, so it
is not the source of these five items either. Their rate/resource data is plausible (distinct
labour/machinery coefficients per item, not a fixed markup of the sibling item) and most likely traces
to an older CPWD schedule edition not present in this repo. Per explicit user decision, these five items
are kept as-is rather than removed, and are marked `# UNVERIFIABLE` in `support_builder_earth_v2.py`
directly above each one — treat them as unverified against any source in this repo, not as PDF-verified
like the rest of Ch02.
