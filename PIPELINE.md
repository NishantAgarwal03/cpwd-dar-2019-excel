# CPWD DAR 2019 — Data Pipeline

## Canonical flow

```
raw_pdf/
  CivilDAR_2019_Vol_1.pdf ─── extract_rates.py ──────────────────► reference_json/rates_master_clean.json
                          └── (manual conversion) ────────────────► converted_xlsx/CivilDAR_2019_Vol_1_Converted.xlsx
                                                                              │
                                           ┌──────────────────────────────────┤
                                           │                                  │
                              extract_norms.py                   vol1_chapter_extractor.py
                                           │                                  │
                                           ▼                                  ▼
                             labour_productivity.json           ch03_items.json … ch12_items.json
                             sundries_reference.json

  CivilDAR_2019_Vol_2.pdf ─── vol2_chapter_extractor.py ─────────► ch13_items.json … ch26_items.json
```

```
ch02: support_builder_earth_v2.py (ITEMS list)  ┐
ch03–12: ch03_items.json … ch12_items.json      ├─ cpwd_item_parser.py ──► gang_registry_vol1.txt/.ini
                                                │                          productivity_registry_vol1.txt/.ini
ch13–26: ch13_items.json … ch26_items.json      ┘─ cpwd_item_parser_vol2.py ► gang_registry_vol2.txt/.ini
                                                                               productivity_registry_vol2.txt/.ini

rates_master_clean.json  ┐
labour_productivity.json ├─ main.py ──────────────────────────────► CPWD_DAR_2019_Vol1_Workbook.xlsx
sundries_reference.json  │   + support_builder_*.py (teaching sheets)
ch03–12_items.json       ┘

ch13–26_items.json ─────────── generate_vol2_workbook.py ─────────► CPWD_DAR_2019_Vol2_Workbook.xlsx
                                + support_builder_vol2_all.py
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
| 2 | `extract_norms.py` | `labour_productivity.json`, `sundries_reference.json` |
| 3 | `vol1_chapter_extractor.py` | `ch03–12_items.json` |
| 4 | `vol2_chapter_extractor.py` | `ch13–26_items.json` |
| 5 | `cpwd_item_parser.py` | `gang_registry_vol1.*`, `productivity_registry_vol1.*` |
| 6 | `cpwd_item_parser_vol2.py` | `gang_registry_vol2.*`, `productivity_registry_vol2.*` |
| 7 | `main.py` | `Vol1_Workbook.xlsx` |
| 8 | `generate_vol2_workbook.py` | `Vol2_Workbook.xlsx` |

## Single source of truth per chapter

| Chapter | Canonical source | Used by |
|---------|-----------------|---------|
| Ch 02 Earth Work | `support_builder_earth_v2.py` — `ITEMS` list | registry (step 5), workbook (step 7) |
| Ch 03–12 Vol 1 | `ch03_items.json` … `ch12_items.json` | registry (step 5) |
| Ch 13–26 Vol 2 | `ch13_items.json` … `ch26_items.json` | registry (step 6), workbook (step 8) |

> **Note for Ch 03–12:** The `support_builder_*.py` teaching sheets for Ch 03–12 still contain
> their own inline data (the "rows" tuples are human-readable teaching text, not machine coefficients).
> These teaching sheets and the ch03–12 JSONs both derive from the same converted XLSX, but they are
> not cross-validated automatically. If you update a coefficient in a support builder, re-run
> `vol1_chapter_extractor.py` (step 3) to keep the JSONs in sync.

## Registry files

| File | Covers | Generator |
|------|--------|-----------|
| `data/gang_registry_vol1.txt` | Ch 02–12 | `cpwd_item_parser.py` |
| `data/productivity_registry_vol1.txt` | Ch 02–12 | `cpwd_item_parser.py` |
| `data/gang_registry_vol2.txt` | Ch 13–26 | `cpwd_item_parser_vol2.py` |
| `data/productivity_registry_vol2.txt` | Ch 13–26 | `cpwd_item_parser_vol2.py` |

INI versions of each file are identical in data — they exist for tooling that prefers INI format.

## Files to delete (pending explicit user confirmation)

These files exist but are superseded and should be removed to avoid confusion:

| File | Reason |
|------|--------|
| `scripts/generate_dar_workbook.py` | Stub — superseded by `main.py` |
| `scripts/build_full_workbook.py` | Stub — superseded by `main.py` |
| `scripts/audit_ch13_finishing.py` | Superseded by `vol2_chapter_audit.py` |
| `scripts/support_builder_earth.py` | v1 — superseded by `support_builder_earth_v2.py` |
| `data/gang_registry.txt` | Superseded by `gang_registry_vol1.txt` |
| `data/gang_registry.ini` | Superseded by `gang_registry_vol1.ini` |
| `data/gang_registry_all_trades.txt` | Identical to above |
| `data/gang_registry_all_trades.ini` | Identical to above |
| `data/productivity_registry.txt` | Superseded by `productivity_registry_vol1.txt` |
| `data/productivity_registry.ini` | Superseded by `productivity_registry_vol1.ini` |
| `data/productivity_registry_all_trades.txt` | Identical to above |
| `data/productivity_registry_all_trades.ini` | Identical to above |
