# Audit: `scripts/earthwork_tables.py` vs `scripts/support_builder_earth_v2.py` (2026-09-19)

## Scope & method

`earthwork_tables.py` (7,699 lines) had never been audited against anything. It is a pure-data
module (no imports) consumed by **only one** file: `scripts/trade_builder_earth.py`, which writes
it into the `02_Earth_Work` sheet of `CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_1.xlsx`.
`earthwork_composer_panel.py` / `earthwork_composer_resolver.py` / `earthwork_composer_catalogue.py`
do **not** import `earthwork_tables` at all — they drive a separate "custom rate builder" section
off `cpwd_rate_engine.py` / `support_builder_earth_v2.py` data. So `earthwork_tables.py` is the sole
source for the sheet's static lookup catalog (Table 4, Table 5E, the D9 item-picker, and the
`$AN$109:$AO$182` / `$Z$109:$AK$157` ranges).

Both files were imported as plain Python data modules and diffed programmatically (exhaustive,
not sampled), then cross-checked against the live workbook with `openpyxl` (`data_only=False`).

## 1. Structure map

| Structure | Rows | Role |
|---|---|---|
| `SCOPE_OPTIONS`, `STRATA_OPTIONS`, `METHOD_OPTIONS`, `LIFT_DEPTH_OPTIONS`, `TASK_OPTIONS`, `CATEGORY_HEADINGS` | small lists | Legacy/aux dropdown option lists |
| `HEADING_TASK_CATALOG` | 74 tuples `(category, task_spec, item_code)` | Drives the AN/AO columns → the guided D9 "Item Specification" dropdown |
| `L1_OPTIONS`, `L2_DATA`, `L3_DATA`, `L4_DATA` | cascade lists | D6→D7→D8→D9 guided decision-tree dropdown cascade |
| `BENCHMARK_PRODUCTIVITY_ROWS` | 27 dicts | Table 4 first-principles productivity matrix (gang/hours/logic), independent narrative table |
| `MASTER_ACTIVITIES` | **49 dicts**, keyed by `item_code` | The authoritative catalog: scope/strata/method/task, `batch_qty`, `batch_unit`, `published_say` (final rate), lookup/fallback keys. Written to cols Z–AK, rows 109–157 |
| `MASTER_ITEMS` | **49 dicts**, keyed by `item_code` string | A near-duplicate of `MASTER_ACTIVITIES` (code/desc/mode/batch/rate only) — **imported by `trade_builder_earth.py` but never referenced again (dead code)** |
| `TABLE_5E_RECORDS` | **198 dicts** | One row per resource line (labour/material/machinery) across all 49 items; written verbatim to cols A–Q, rows 109–306 |
| `_MANUAL_EXCAVATION_RECORDS`, `_CODE_TO_CAT_TASK`, `_PIPE_CODES`, `_COMPOSITE_CODES` | helper structures | Internal bookkeeping, not separately audited (derived/unused outside file) |

Everything is keyed by CPWD item code (`2.x.y`) even though the file's own docstring claims
"No CPWD Code Numbering in User Tables" — the codes are very much present and are the join key
used for all lookups.

## 2. Cross-reference vs `support_builder_earth_v2.py` ITEMS (exhaustive, all 49 overlapping codes)

49 of `earthwork_tables.py`'s 49 items have a matching id in `support_builder_earth_v2.py`'s 84
`ITEMS`. Of those 49:

### Confirmed real rate discrepancies (3)

| Code | earthwork_tables `published_say` | support_builder_earth_v2 `dsr_rate` | Diff |
|---|---|---|---|
| **2.9.3** (foundation trench, hard rock no-blast, manual chiselling) | **1080.55** | **624.70** | **Rs 455.85 (≈73%)** — see analysis below, this is the headline finding |
| 2.10.1.3 (pipe trench soil, dia >300mm) | 568.80 | 568.60 | Rs 0.20 |
| 2.12 (extra pipe trench depth 3–4.5m) | 315.05 | 314.95 | Rs 0.10 |

**2.9.3 deep-dive:** both files carry an *identical* 9-line resource buildup (same resource
codes, quantities, and unit rates — Hydraulic Excavator, Tipper, Rock Excavator/Breaker/
Chiseller, Blacksmith, Beldar, Coolie, Sundries), summing to an identical direct cost of
Rs 8,076.07 for the 10-cum batch (Rs 807.61/cum). A markup-consistency scan across all 44
`DIRECT`-mode items in `earthwork_tables.py` shows every single one applies essentially the same
~1.337–1.340× statutory markup factor (Water 1% + GST 14.05% + CPOH 15% + Cess 1%) between direct
cost and `published_say` — and 2.9.3 fits that pattern exactly (1.338×, giving 1080.55/cum).
`support_builder_earth_v2.py`'s 624.70 is **arithmetically impossible** for this same resource
buildup: 624.70 < the direct cost alone (807.61/cum), before any markup is even added. This means
one of the two files has a genuine, provable error for 2.9.3, and the resource-line evidence points
at `support_builder_earth_v2.py`'s 624.70 being the wrong value (possibly a stale/mistyped figure
that survived this session's 74-value fix pass), not at `earthwork_tables.py`. Flagging for
reconciliation rather than asserting a fix, per the instruction to treat `support_builder_earth_v2`
as ground truth — but the arithmetic strongly overrides that trust for this one code.

### Cosmetic-only differences (not rate bugs, verified via markup/derivation-batch analysis)

19 codes show `batch_qty`/`base_qty` or unit-spelling differences (`m` vs `metre`, `hole` vs
`each`, `tree` vs `each`, and batch sizes like `et batch_qty=1` vs `sb base_qty=180`). These are
**not** live discrepancies: `published_say` / `dsr_rate` are always the true *final per-unit* rate
in both files (confirmed both by the markup-ratio scan above and by tracing `trade_builder_earth.py`
— the D14 "Published Reference Say Rate" formula reads `AE` directly with no division by `AC`).
The differing qty fields just record each file's own "worked-example batch size" from the DSR
schedule (e.g. "cost for 180 m of pipe") and don't affect the rate actually shown. 2.24.1/2.24.2
(percentage extras) also match conceptually — `sb` just stores `dsr_rate=None` for percentage items
where `et` stores the percentage as text.

### Coverage gaps

- **Codes only in `support_builder_earth_v2.py` (35, unverifiable against the workbook because the item isn't offered there at all):**
  `2.1.2, 2.13.1.2, 2.13.1.3, 2.13.2.1, 2.13.2.2, 2.13.2.3, 2.13.3.1, 2.13.3.2, 2.13.3.3, 2.14, 2.15, 2.17.1, 2.17.2, 2.17.3, 2.18.1, 2.18.2, 2.18.3, 2.19, 2.2.2, 2.20.1, 2.20.2, 2.20.3, 2.21.1, 2.21.2, 2.21.3, 2.22.1, 2.22.2, 2.22.3, 2.23, 2.25(a), 2.28.2, 2.29.2, 2.3.2, 2.6.2, 2.8.2`
  (of these, `2.1.2, 2.2.2, 2.3.2, 2.6.2, 2.8.2` are the 5 accepted UNVERIFIABLE exceptions per your instructions and are excluded from concern here.)
- **Codes only in `earthwork_tables.py`:** none — every one of its 49 items has an `sb` counterpart.

## 3. Coverage gap analysis — the real headline problem

`HEADING_TASK_CATALOG` (the list that actually populates the D9 "Item Specification" dropdown a
user picks from) has **74 entries**, but `MASTER_ACTIVITIES` (the actual rate table) has only
**49**. **26 dropdown entries point at item codes that don't exist in `MASTER_ACTIVITIES` at all**:

```
2.13.1.2, 2.13.1.3, 2.13.2.1, 2.13.2.2, 2.13.2.3, 2.13.3.1, 2.13.3.2, 2.13.3.3,
2.17.1, 2.17.2, 2.17.3, 2.18.1, 2.18.2, 2.18.3, 2.19,
2.20.1, 2.20.2, 2.20.3, 2.21.1, 2.21.2, 2.21.3, 2.22.1, 2.22.2, 2.22.3, 2.23,
2.25a (duplicate entry, see §4)
```

This is not just a "missing data" gap — it is a **silent-wrong-answer bug**. I traced
`trade_builder_earth.py`'s `_master_lookup()` helper: the D14 rate cell uses a 3-tier
`IFERROR(INDEX/MATCH(...))` chain — (1) match `category|task`, (2) match
`scope|strata|method|task`, (3) match `scope|strata|method` only (task-agnostic fallback), and
only if all three fail does it fall back to a hardcoded literal default (e.g. `92.55`,
`"General surface cut..."`).

- Selecting a **timbering shaft/area/permanent** item (2.17.x, 2.18.x, 2.19, 2.20.x–2.23, 17 codes)
  from the dropdown: tiers 1–2 always fail (no such task exists in `MASTER_ACTIVITIES`). Tier 3
  (`scope|strata|method` = "Timbering & Shoring|All kinds of soil|Manual labor (depth ≤1.5m)")
  **does** match — against item **2.16.1** ("Close timbering depth ≤1.5 m"), because that
  `fallback_key` triplet exists. Result: the sheet **silently shows 2.16.1's rate and description**
  for whatever timbering item the user actually picked, with no warning.
- Selecting a **hard-rock pipe trench** item (2.13.2.x, 2.13.3.x, 6 codes): no matching
  `scope|strata|method` triplet exists anywhere in `MASTER_ACTIVITIES` for
  "Pipeline & Cable Trenching + Hard rock", so even tier 3 fails — the sheet falls all the way to
  the **hardcoded literal default** (surface-excavation-in-soil rate 92.55/sqm), a category-level
  wrong answer.
- 2.13.1.2/2.13.1.3 (ordinary-rock pipe trenches, other diameters) land on tier 3 and silently
  borrow 2.13.1.1's rate/description instead of their own.

I did not attempt to fix this — it's a design/data-completeness problem, not a value typo, and
needs a decision on scope (either populate the missing 26 `MASTER_ACTIVITIES` rows properly, using
`support_builder_earth_v2.py`'s resource buildups as source, or remove the 26 dead entries from
`HEADING_TASK_CATALOG` so the dropdown can't offer them).

## 4. Internal consistency findings (independent of `support_builder_earth_v2.py`)

1. **`MASTER_ITEMS` is dead code with a live sign bug.** `MASTER_ITEMS` is imported into
   `trade_builder_earth.py` (line 31) but never referenced again — grep confirms the only use of
   the name in that file is the import statement. It duplicates `MASTER_ACTIVITIES` but with a sign
   inversion for the two `DEDUCT`-mode items: `MASTER_ITEMS['2.4'].published_say = -4.3` vs
   `MASTER_ACTIVITIES` (the live copy) `= 4.3`; same for `2.5` (-33.0 vs 33.0). Currently harmless
   because the dict is unused, but it's a latent trap if anyone ever wires `MASTER_ITEMS` up later.
2. **Duplicate catalog entry.** `("3. Excavation – Soil / Earth", "Local earth supply, excavation & filling (mechanical)", "2.25a")` and `("6. Filling & Backfilling", "Local earth supply, excavation & filling (mechanical)", "2.25a")` are both present in `HEADING_TASK_CATALOG` — the same item code listed under two different category headings. Low-severity (it's at least the same underlying item both times, unlike the 26-code gap above), but it means neither of the two `cat_task_key` values actually generated (`"3. Excavation – Soil / Earth|Local earth supply..."` / `"6. Filling & Backfilling|Local earth supply..."`) necessarily equals whatever `cat_task_key` `MASTER_ACTIVITIES['2.25a']` actually has, so it likely also rides the tier-3 fallback rather than a clean tier-1 match. Worth a quick check/cleanup alongside item 3.
3. **Orphan resource records.** `TABLE_5E_RECORDS` contains 2 extra item codes not present in `MASTER_ACTIVITIES`: `2.6.1-M` and `2.8.1-M` (10 lines each — "Manual (Hand tools & labour gang)" method variants of 2.6.1/2.8.1). These are written into the sheet as extra Table-5E rows (they're part of the 198) but have no corresponding catalog entry, so they're inert reference-only rows, not reachable via the D9 dropdown. Not a bug, just unreferenced data.
4. **Stale row-count comment.** `T5E_ROWS = len(TABLE_5E_RECORDS)  # 178` — the comment says 178 but `len()` is computed dynamically and the actual live count is 198. Harmless (the code uses the dynamic value), but it's a tell that the file grew significantly after that comment was written and nothing re-verified the comment, consistent with "never audited."
5. **No TODO/FIXME/XXX/placeholder markers** were found anywhere in the file — no explicit authorial red flags were left behind.
6. **Markup-ratio scan (44 DIRECT-mode items):** all cluster tightly at 1.337×–1.340×, and all 8 `COMPOSITE`/`EXTRA_PCT`-mode items cluster at 1.000×–1.012× (as expected, since composite items reference already-marked-up base items rather than raw resources). No internal outliers apart from 2.9.3 being compared against the *external* sb figure (§2) — i.e., `earthwork_tables.py`'s own arithmetic is internally coherent throughout.
7. No duplicate `item_code` keys within `MASTER_ACTIVITIES`/`MASTER_ITEMS` (49 unique in each, sets identical).

## 5. Wiring sanity check (full population, not just spot-check)

Loaded the live `CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_1.xlsx` → `02_Earth_Work` sheet
with `openpyxl(data_only=False)` and compared every value `trade_builder_earth.py` is supposed to
have written against the corresponding Python source value:

- **All 49/49** `MASTER_ACTIVITIES` rows (cols AC/AD/AE, rows 109–157: `batch_qty`, `batch_unit`,
  `published_say`) match the sheet exactly. 0 mismatches.
- **All 198/198** `TABLE_5E_RECORDS` rows (cols F/K/N/O, rows 109–306: `res_code`, `qty_coeff`,
  `res_unit`, `rate`) match the sheet exactly. 0 mismatches.
- **All 74/74** `HEADING_TASK_CATALOG` rows (cols AN/AO, rows 109–182) match the sheet exactly.
  0 mismatches.

So the write path itself (`earthwork_tables.py` → `trade_builder_earth.py` → cell values) is a
faithful, bug-free 1:1 transcription — every literal value that exists in the Python data reaches
the sheet unchanged. **The problems are entirely upstream, in the data/coverage itself, not in the
writing code.**

## Verdict & priority

`earthwork_tables.py` is **better internally than expected** for a never-audited file — its own
arithmetic is self-consistent (markup ratios, resource buildups), there's no dead-typo debris, and
the write path to the live workbook is 100% faithful. But it has one real design defect that
actively produces wrong numbers for real users today, plus one confirmed rate error:

1. **HIGH — fix first.** 26 of 74 dropdown entries in `HEADING_TASK_CATALOG` (all timbering
   shaft/area/permanent-timbering variants, plus 6 hard-rock pipe-trench variants) have no backing
   `MASTER_ACTIVITIES` row. Selecting any of them from the guided D9 dropdown silently substitutes
   a wrong item's rate (or a generic default) with zero indication to the user. This is a live,
   user-facing correctness bug on a deliverable in active use. Fix by adding the 26 missing
   `MASTER_ACTIVITIES`/`TABLE_5E_RECORDS` entries (source data available in
   `support_builder_earth_v2.py`'s `ITEMS` for the same codes) or, as a stop-gap, removing those 26
   entries from `HEADING_TASK_CATALOG` so they can't be selected until real data exists.
2. **MEDIUM.** Item **2.9.3** — reconcile the Rs 455.85/cum (≈73%) gap between
   `earthwork_tables.py` (1080.55, arithmetically consistent with its own resource buildup) and
   `support_builder_earth_v2.py` (624.70, arithmetically impossible given the identical resource
   buildup both files carry). Evidence points to the `support_builder_earth_v2.py` value being
   the one in error.
2b. **LOW.** 2.10.1.3 (Rs 0.20) and 2.12 (Rs 0.10) rate deltas — trivial, likely rounding drift; fix opportunistically.
3. **LOW.** Duplicate `2.25a` entry in `HEADING_TASK_CATALOG`, and the two orphan `-M` resource
   records — cosmetic/dead data, no user-facing effect currently, clean up if convenient.
4. **VERY LOW / informational.** `MASTER_ITEMS`'s dead-but-sign-inverted duplicate of
   `MASTER_ACTIVITIES`, and the stale `# 178` row-count comment — no live effect; either delete
   `MASTER_ITEMS` (unused) or fix its sign and keep it in sync, and correct the comment.

Nothing in this audit requires touching `support_builder_earth_v2.py` (already independently
verified this session) except to flag item 2.9.3 back to that verification pass for reconciliation.
