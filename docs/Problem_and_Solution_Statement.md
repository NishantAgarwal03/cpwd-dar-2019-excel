# Custom Non-DSR Rate Analysis Workbooks — Problem & Solution Statement

**Base files studied:**

- `CivilDAR_2019_Vol_1_Converted.xlsx` — CPWD Delhi Analysis of Rates, 2019, Vol. 1 (sub-heads 01–12)
- `CivilDAR_2019_Vol_2_Converted.xlsx` — CPWD Delhi Analysis of Rates, 2019, Vol. 2 (sub-heads 13–26)

**Status:** Design document only — no workbook built yet.

---

## 0. What the base files actually teach us (grounding — nothing invented)

| #   | Fact                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           | Evidence                                                                                                                                                                                                     |
| --- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| 1   | The entire book's atomic rate catalogue (materials, labour day-rates, machinery hire — 2,154 unique codes) lives **only in Vol. 1's `00_Basic_Rates` sheet**. Vol. 2's own sheet named `00_Basic_Rates` is not a rate table at all — it's front-matter (preface/contents pages). Every Vol. 2 item still uses the *same* codes (verified: codes like `8300`, `1854`, `1881` used only in Vol. 2 items are present in Vol. 1's catalogue).                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                      | Vol.2 `00_Basic_Rates` = 47 rows of preface text. Vol.1 `00_Basic_Rates` = 2,154 codes, and spot-checked Vol.2-only codes all resolve there.                                                                 |
| 2   | The standard **item anatomy** (MATERIAL → LABOUR → optional MACHINERY → optional Sundries → W→X→Y→Z markup chain → final rate) holds across **both volumes** — confirmed identically in Vol. 2 sheets 14, 15, 17, 18, 19, 20, 21, 24, 25, 26, not just Vol. 1.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 | Spot-checked `14_Repairs_to_Buildings`, `18_Water_Supply`, `19_Drainage`, `20_Pile_Work`, `21_Aluminium_Work`, `22_Water_Proofing`, `24_Heritage_Buildings`, `25_Structural_Glazing`, `26_New_Technologies`. |
| 3   | `03_Mortars` (Vol.1) still breaks the pattern the same way in this larger picture: no markup chain, pure base-cost, feeding **many** Vol.2 sheets as a MATERIAL line, not just Vol.1 ones.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                     | Cross-reference count: `13_Finishing`→Mortars ×51, `14_Repairs`→Mortars ×13, `16_Road_Work`→Mortars ×3, `20_Pile_Work`→Mortars ×6, `22_Water_Proofing`→Mortars ×2.                                           |
| 4   | Two structural variants recur in Vol.2 too: `15_Dismantling_Demolishing` and `24_Heritage_Buildings` have **labour-only** items (no MATERIAL block) — same pattern as `02_Earth_Work`. `25_Structural_Glazing` item 25.1 has **no LABOUR block at all** (material + fabrication weight only) — a variant not seen anywhere in Vol.1.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           | `15_Dismantling` item 15.1, `24_Heritage` item 24.1 — LABOUR only. `25_Structural_Glazing` item 25.1 — MATERIAL → TOTAL directly.                                                                            |
| 5   | **Real cross-volume dependencies exist in both directions, and every one of them is traceable and resolvable:** <br>• **Vol.1 → Vol.2** (3 distinct items, all needed by Vol.1 builder sheets): `09_Wood_and_PVC_Work` needs Finishing items `13.50.1` (₹57.05/sqm), `13.50.3` (₹50.70/sqm); `10_Steel_Work` needs `13.50.3` too (same item, reused); `08_Cladding_Work` needs Water Supply item `18.78` (₹154.15/metre). **All 4 usages resolve to just 4 distinct external items**, and every one of them is built entirely from codes already in the shared Rates catalogue — `18.78` additionally uses Concrete Work item `4.2.5`, which is already in Vol.1. **Nothing here needs data that doesn't already exist in Vol.1.**<br>• **Vol.2 → Vol.1**: `13_Finishing`, `14_Repairs_to_Buildings`, `16_Road_Work`, `20_Pile_Work`, `22_Water_Proofing` pull Mortars (and `14_Repairs` also pulls one Wood/PVC item); `20_Pile_Work` also pulls RCC Work item `5.33.1`. This direction is fine to keep as a live dependency. | Full cross-reference matrix run against both files; breakdowns of `13.50.1`, `13.50.3`, `13.57.1`, `18.78` traced line-by-line and confirmed self-resolving.                                                 |

Facts 3, 4 and 5 directly shape the architecture and per-sheet statements below.

---

## 1. Master Problem Statement

> A quantity surveyor / estimator working from the CPWD DAR 2019 (both volumes) needs to price **non-scheduled ("non-DSR") items** across the *entire* book — not just the 12 sub-heads in Vol. 1 — while staying faithful to the book's own costing method: the same atomic material/labour/machinery rates, the same coefficient-based build-up, the same Sundries convention, and the same statutory markup chain used throughout. The book is split into two volumes for printing, but its underlying logic is **one shared system** — one rate catalogue, one markup convention, and a small number of real cross-references between the volumes (Vol.1 items borrowing rates from Vol.2's Finishing and Water Supply sub-heads; several Vol.2 sub-heads borrowing Vol.1's Mortars and RCC rates). Today there's no self-contained tool that lets someone assemble a *new* item in the *style* of any one of the book's 26 categories, using only inputs that already exist in the book, without both volumes needing to be open and cross-checked by hand — and any such tool has to respect the fact that Vol.1 must stand alone (it's the foundation), while Vol.2 legitimately depends on it.

## 2. Master Solution Statement — two-workbook architecture

**Two workbooks, one dependency direction: WB2 → WB1, never WB1 → WB2.**

### WB1 (built from Vol.1) — fully self-contained, 18 sheets

- **5 shared infrastructure sheets:**
  
  1. `Rates_Master` — the full 2,154-code catalogue (materials, labour, machinery), copied as static values.
  2. `Labour_Machinery_Productivity` — labour/machinery requirements mined from every item's LABOUR block, grouped by work type (e.g. "what does brickwork need").
  3. `Sundries_Reference` — every Sundries L.S. line found across the sheets, as a lookup to guide manual entry.
  4. `Global_Factors` — markup defaults (Water 1%, GST 0.1405, CPOH 15%, Cess 1%, Sundries ×2.00), editable, with an on/off flag per step per sheet (needed for Carriage's partial chain).
  5. **`Resolved_Cross_Volume_Items` (new)** — holds the 4 specific items Vol.1 needs from Vol.2 (`13.50.1`, `13.50.3`, `13.57.1` "Finishing" primer/preservative rates, and `18.78` "Water Supply" chase-cutting rate), each fully broken down (material/labour lines, all using WB1's own Rates_Master and, for `18.78`, WB1's own Concrete Work rate) and computed to its final rate — a **result**, not a link, so WB1 never needs Vol.2 or WB2 open. `06_Cladding_Work`, `09_Wood_and_PVC_Work` and `10_Steel_Work` point their relevant MATERIAL line at this sheet instead of re-deriving it.

- **12 builder sheets** (01–12), each a self-contained "compose a custom item in this category's style" UI, per the anatomy already established (item header → MATERIAL rows → LABOUR rows → Sundries → markup block → final rate → item library).

- Cross-sheet component picking (Mortars → Masonry/Stone/Cladding/Flooring; Resolved_Cross_Volume_Items → Cladding/Wood_PVC/Steel Work) works exactly as previously designed.

### WB2 (built from Vol.2) — same structure, live-referencing WB1

- **14 builder sheets**, one per Vol.2 sub-head: 13 (Finishing), 14 (Repairs to Buildings), 15 (Dismantling & Demolishing), 16 (Road Work), 17 (Sanitary Installations), 18 (Water Supply), 19 (Drainage), 20 (Pile Work), 21 (Aluminium Work), 22 (Water Proofing), 23 (Rain Water Harvesting), 24 (Heritage Buildings), 25 (Structural Glazing), 26 (New Technologies) — same builder anatomy as WB1's 12.
- **No duplicated infrastructure.** Per your decision, WB2 does **not** carry its own copies of `Rates_Master`, `Labour_Machinery_Productivity`, `Sundries_Reference` or `Global_Factors`. Its builder sheets' component pickers, coefficient lookups and markup defaults reference WB1's four infra sheets directly (external workbook links) — one source of truth, at the cost of WB2 needing WB1 physically present (same folder / both open) to fully calculate.
- Component picking additionally draws on WB1's `03_Mortars` output (needed by `13_Finishing`, `14_Repairs_to_Buildings`, `16_Road_Work`, `20_Pile_Work`, `22_Water_Proofing`), WB1's `09_Wood_and_PVC_Work` output (needed by `14_Repairs_to_Buildings`), and WB1's `05_RCC_Work` output (needed by `20_Pile_Work`) — all external references into WB1, matching the direction fact #5 confirmed is safe.
- WB2 does **not** need a `Resolved_Cross_Volume_Items`-style sheet of its own — since it's allowed to depend on WB1, it simply references WB1's builder outputs live rather than pre-resolving anything.

```
                         ┌─────────────── WB1 (self-contained) ───────────────┐
                         │  Rates_Master, Labour_Machinery_Productivity,       │
                         │  Sundries_Reference, Global_Factors,                │
                         │  Resolved_Cross_Volume_Items                        │
                         │  + 12 builder sheets (incl. 03_Mortars,             │
                         │    05_RCC_Work, 09_Wood_and_PVC_Work)               │
                         └───────────────────────┬─────────────────────────────┘
                                                  │  (live external references)
                                                  ▼
                         ┌─────────────── WB2 (depends on WB1) ───────────────┐
                         │  14 builder sheets (13, 14–26)                      │
                         └──────────────────────────────────────────────────────┘
```

---

## 3. WB1 — Per-sheet Problem & Solution Statements

### 01 — Carriage of Materials

**Problem:** Pricing carriage of a new material/waste type, or an uncovered lead/lift, by mechanical or manual means.
**Solution:** Two modes — (a) *Mechanical trip*, only CPOH active by default; (b) *Manual labour*, lead-distance matrix — both as previously designed.

### 02 — Earth Work

**Problem:** Pricing an earthwork operation at a labour mix or output not printed.
**Solution:** Labour-and-machinery-only builder, full markup chain.

### 03 — Mortars

**Problem:** A new mortar mix needed as a **building block** for other custom items, not a final priced item.
**Solution:** Stops at Cost of unit, markup chain off by default — feeds Masonry/Stone/Cladding/Flooring **and now several WB2 sheets** (Finishing, Repairs, Road Work, Pile Work, Water Proofing) via WB2's external reference.

### 04 — Concrete Work

**Problem:** Pricing a plain cement concrete mix/grade not printed.
**Solution:** Standard full-anatomy, self-contained builder — its output is also what WB1's `Resolved_Cross_Volume_Items` entry for `18.78` pulls from.

### 05 — RCC Work

**Problem:** Pricing an RCC item (excluding shuttering/reinforcement) not printed.
**Solution:** Standard full-anatomy builder, self-contained — its output is also referenced live by WB2's Pile Work sheet.

### 06 — Masonry Work

**Problem:** Pricing a masonry item in a mortar ratio/brick type not printed.
**Solution:** MATERIAL picker offers `03_Mortars` rates, as previously designed.

### 07 — Stone Work

**Problem:** Pricing a stone masonry item not printed.
**Solution:** Same shape as Masonry — Mortar-sheet rate as a MATERIAL option.

### 08 — Cladding Work

**Problem:** Pricing a cladding/veneer item not printed.
**Solution:** MATERIAL picker offers `03_Mortars` rates (bedding + pointing) **and now `Resolved_Cross_Volume_Items`' `18.78` rate** for chase-cutting-style components, fully resolved without touching Vol.2.

### 09 — Wood and PVC Work

**Problem:** Pricing a wood/PVC item not printed.
**Solution:** Standard self-contained builder, **plus its MATERIAL picker offers `Resolved_Cross_Volume_Items`' `13.50.1` and `13.50.3` primer rates** where a printed/priming finish is part of the item — resolved, no Vol.2 dependency.

### 10 — Steel Work

**Problem:** Pricing a structural steel item not printed.
**Solution:** Standard builder; **its MATERIAL picker offers `Resolved_Cross_Volume_Items`' `13.50.3` priming-coat rate** — same resolved item Wood/PVC uses, no separate re-derivation, no Vol.2 dependency. (This replaces the earlier "leave as placeholder" call — Vol.2 data is now in hand and the item resolves cleanly.)

### 11 — Flooring

**Problem:** Pricing a flooring item in a mortar ratio not printed.
**Solution:** Mortar-sheet rate as a MATERIAL option, as previously designed.

### 12 — Roofing

**Problem:** Pricing a roofing item not printed.
**Solution:** Standard self-contained builder, as previously designed.

---

## 4. WB2 — Per-sheet Problem & Solution Statements

### 13 — Finishing

**Problem:** Pricing a painting/polishing/finishing coat (primer, paint, French polish, etc.) not printed.
**Solution:** Standard builder — paint/primer MATERIAL lines from `Rates_Master`, Painter/Coolie LABOUR, brushes/Sundries; MATERIAL picker also offers `03_Mortars` (via WB1 reference) for items that finish over a plastered/mortar base coat.

### 14 — Repairs to Buildings

**Problem:** Pricing a repair/patch item (plaster patch, chowkhat repair, etc.) not printed.
**Solution:** Standard builder; MATERIAL picker offers `03_Mortars` (heavily used here) and `09_Wood_and_PVC_Work` outputs (both via WB1 reference).

### 15 — Dismantling and Demolishing

**Problem:** Pricing a dismantling/demolition operation not printed.
**Solution:** Labour-and-machinery-only builder (no MATERIAL block, matching the observed pattern), full markup chain.

### 16 — Road Work

**Problem:** Pricing a road-construction item (sub-base, WBM, bituminous layer, etc.) not printed.
**Solution:** Standard builder, machinery-heavy (roller/paver hire from `Rates_Master`); MATERIAL picker also offers `03_Mortars` and the sheet's own prior Road Work outputs (in-sheet library), matching the book's own internal self-referencing pattern.

### 17 — Sanitary Installations

**Problem:** Pricing a sanitary fixture installation (WC, wash basin, urinal, etc.) not printed.
**Solution:** Standard self-contained builder — fixture MATERIAL lines, Fitter/Mason LABOUR, cement/sand and carriage L.S. lines, full markup chain.

### 18 — Water Supply

**Problem:** Pricing a water-supply pipe/fitting installation not printed.
**Solution:** Standard builder with a wastage-% helper on pipe/fitting quantities (matching the book's "add 30% for fittings" pattern); MATERIAL picker also offers the sheet's own prior Water Supply outputs, matching the book's internal self-referencing pattern (this is also where `18.78`, feeding WB1's Cladding sheet, originates).

### 19 — Drainage

**Problem:** Pricing a drainage pipe-laying item not printed.
**Solution:** Standard self-contained builder — pipe + jointing-material MATERIAL lines with a breakage-allowance helper, Mason/Beldar/Bhisti LABOUR, full markup chain.

### 20 — Pile Work

**Problem:** Pricing a piling item (diameter/depth/method combination) not printed.
**Solution:** Builder whose MATERIAL picker offers `05_RCC_Work`'s rate (via WB1 reference) for the concrete-in-pile component, plus a MACHINERY block (piling rig/crane hire) alongside MATERIAL and LABOUR — the first WB2 sheet where MACHINERY is a first-class block rather than folded into materials.

### 21 — Aluminium Work

**Problem:** Pricing an aluminium fabrication item (frame/section combination) not printed.
**Solution:** Standard self-contained builder — weight-based aluminium-section MATERIAL lines with a wastage-% helper, anodising/coating as a MATERIAL line, Fitter/Skilled-Beldar LABOUR, full markup chain.

### 22 — Water Proofing

**Problem:** Pricing a water-proofing treatment (layered system) not printed.
**Solution:** Builder whose MATERIAL picker offers `03_Mortars` rates (bedding + slurry, via WB1 reference) alongside stone/membrane/compound MATERIAL lines, full markup chain.

### 23 — Rain Water Harvesting

**Problem:** Pricing a borewell/recharge-structure item not printed.
**Solution:** Standard self-contained builder, machinery-heavy (excavator/rig/tanker hire from `Rates_Master`), Beldar/Mistry LABOUR, full markup chain.

### 24 — Heritage Buildings

**Problem:** Pricing a heritage conservation/restoration operation not printed.
**Solution:** Labour-only builder (no MATERIAL block, matching the observed pattern, similar to Dismantling/Earth Work), full markup chain.

### 25 — Structural Glazing

**Problem:** Pricing a structural-glazing/ACP fabrication item not printed.
**Solution:** Standard builder, but with **LABOUR made optional** (the base book's own item 25.1 has none — fabrication priced by weight only) — a genuine structural variant, not an omission.

### 26 — New Technologies

**Problem:** Pricing a proprietary/new-material item (branded flooring, panel systems, etc.) not printed.
**Solution:** Standard self-contained builder — material-heavy per-unit-area MATERIAL lines with a wastage-% helper, full markup chain.

---

## 5. Decisions log

**Round 1 (WB1 only):**

1. Sundries — manual entry per item, guided by a dedicated `Sundries_Reference` sheet.
2. Rates infrastructure — one combined `Rates_Master` plus a separate `Labour_Machinery_Productivity` sheet by work type.
3. Steel Work → Finishing dependency — treated as a normal forward reference, assumed available.
4. `13_Finishing_Work` — left as a placeholder (**superseded in Round 2** — see below).

**Round 2 (WB1 + WB2):**
5. WB2 does not duplicate WB1's infra sheets — it live-references them (single source of truth; WB2 needs WB1 physically present).
6. The 4 resolved cross-volume items Vol.1 needs are centralized in one new WB1 sheet, `Resolved_Cross_Volume_Items`, rather than duplicated inside each consuming builder sheet.
7. With actual Vol.2 data in hand, all 4 of those items resolve cleanly using only WB1's own data (Rates_Master + Concrete Work) — the earlier "Finishing sheet is a placeholder/limitation" framing is dropped; Steel Work and Wood/PVC's Finishing dependency, and Cladding's Water Supply dependency, are now fully specified, not open.

No open items remain. Next step, whenever you want it, is the cell-by-cell layout for each sheet in both workbooks.
