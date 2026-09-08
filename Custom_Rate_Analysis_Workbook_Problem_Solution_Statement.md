# Custom Non-DSR Rate Analysis Workbook — Problem & Solution Statement

**Base file studied:** `CivilDAR_2019_Vol_1_Converted.xlsx` (CPWD Delhi Analysis of Rates, 2019, Vol. 1)
**Status:** Design document only — no workbook built yet, per your instruction.

---

## 0. What the base file actually teaches us (grounding — nothing invented)

Before writing problem/solution statements, here is the logic extracted directly from the base workbook. Everything downstream is built only on these observed facts.

| # | Fact | Evidence |
|---|------|----------|
| 1 | `00_Basic_Rates` is a flat catalogue of ~2,437 atomic inputs — materials, daily labour wage rates, machinery hire rates — each with a Code, Description, Unit, Rate. Every item in sheets 01–12 pulls its line items from these codes. | Codes like `0367` (Cement), `0114` (Beldar), `0002` (Concrete Mixer hire) recur identically across sheets. |
| 2 | The standard **item anatomy** in sheets 02, 04–12 is: MATERIAL block (code, description, unit, qty coefficient, rate, amount) → LABOUR block (same shape) → optional Sundries/L.S. lines → subtotal **W** → +1% Water charges → **X** → +GST (factor 0.1405) → **Y** → +15% CPOH → **Z** → +1% Cess → final **Cost of unit**. | Verified identically in `04_Concrete_Work`, `05_RCC_Work`, `06_Masonry_Work`, `07_Stone_Work`, `08_Cladding_Work`, `09_Wood_and_PVC_Work`, `11_Flooring`, `12_Roofing`. |
| 3 | **`03_Mortars` breaks the pattern**: its items stop at Material + Labour + Sundries → straight to "Cost of 1 cum" — **no markup chain at all**. Its rates (e.g. item 3.9, 3.11) are then pulled into Masonry, Stone, Cladding and Flooring items as a single MATERIAL line ("Rate as per Item No. 3.9 of SH: Mortar"), and the *consuming* item applies the full markup chain once, over the combined total. | Mortar item 3.1: Material+Labour+Sundries = Cost, no W/X/Y/Z rows. Masonry item 6.1.1 imports 3.9's rate as a material line, then runs its own full W→X→Y→Z chain. |
| 4 | **`01_Carriage_of_Materials` breaks the pattern differently**: mechanical-transport items apply **only +15% CPOH** (no Water charge, no GST, no Cess), and manual-labour carriage is modelled as a **lookup matrix** (labour-per-trip vs. lead distance) rather than the material+labour item template. | Item "Disposal of moorum..." row 17 goes straight from TOTAL to "Add 15% CPOH" — no Water/GST/Cess rows present. |
| 5 | Sundries amounts are **not** derived by one consistent formula from what's visible in-sheet (checked against material total, labour total, and combined total for several items — none matched a fixed %). The "L.S., Qty, Rate=2.00" pattern (2% multiplier) is applied but the *base* the 2% is taken on varies per item and isn't shown. | Cross-checked Concrete 4.1.2, Earth Work 2.2, Roofing 12.1.1 — no single formula reproduces all three Sundries values. |
| 6 | At least one item (`10_Steel_Work`, item 10.1) references a rate from a sheet **not present in this Volume-1 file at all** ("Rate as per Item Number 13.50.3 of SH: Finishing") — only the resulting number (₹50.70/sqm) is available, not its breakdown. | Row 12 of `10_Steel_Work`. |

Facts 3, 4 and 6 are structural exceptions, not noise — they directly shape the per-sheet solution statements below.

---

## 1. Master Problem Statement

> A quantity surveyor / estimator working from the CPWD DAR 2019 (Vol. 1) needs to price **non-scheduled ("non-DSR") items** — items that don't exist in the printed book — while staying faithful to the book's own costing method: the same atomic material/labour/machinery rates, the same coefficient-based build-up, the same Sundries convention, and the same statutory markup chain (Water charge → GST → CPOH → Cess) used throughout the book. Today this can only be done by manually re-deriving the book's logic in a blank sheet, which is slow and error-prone, and by definition can't be checked against the book once it's copied out. There is no self-contained tool that lets someone assemble a *new* item, in the *style* of any one of the book's 12 work categories, using only inputs that already exist in the book — without needing the book itself open at the same time.

## 2. Master Solution Statement (workbook architecture)

**Composition — 16 sheets total:**

- **4 shared infrastructure sheets** (data + rules, no item-building UI):
  1. `Rates_Master` — one combined, flat catalogue of every atomic rate — materials, labour day-rates, and machinery hire — exactly as `00_Basic_Rates` already keeps them together in the base book. This is the single source of truth for "what does X cost", with basic filter/search. It does **not** get a "build a custom item" UI — it has nothing to build, it's the parts bin.
  2. `Labour_Machinery_Productivity` — a *derived* reference, mined from the LABOUR blocks of every item across sheets 02, 04–12 (all values already exist in the base file — this sheet only re-groups them), organized by **work type** (Brickwork, Earthwork excavation, Concreting, Stone/Cladding fixing, Carpentry, Steel fixing, Roofing, Flooring) rather than by sheet. For each work type it lists: labour roles involved, their typical day-coefficient per unit of work, and any machinery normally paired with them — so you can look up "what labour does brickwork need" directly, then carry that team into the relevant builder sheet as a starting point (editable, not locked).
  3. `Sundries_Reference` — every Sundries L.S. line found across the 12 base sheets (item, work type, quantity, amount), kept as a lookup a user can consult when entering Sundries manually on a builder sheet, since fact #5 shows there's no single formula to auto-derive it.
  4. `Global_Factors` — the markup defaults found in the book: Water 1%, GST factor 0.1405, CPOH 15%, Cess 1%, Sundries multiplier 2.00 — plus a flag per factor for "applies by default" so a sheet like Carriage can switch off Water/GST/Cess the way the book itself does. Every value here is **editable**, pre-filled with the book's own defaults — nothing is hard-baked into the builder sheets themselves; they read from this sheet.

- **12 builder sheets**, one per base sheet (01–12), each a self-contained "compose a custom item in this category's style" UI.

**Cross-sheet component picking:** a builder sheet's component picker draws from three pools — raw `Rates_Master` lines, a suggested labour team from `Labour_Machinery_Productivity`, **and** the live final/base rate already computed on another builder sheet (mirroring how the book itself lets Masonry pull Mortar's rate, or Stone Work pull Mortar's rate). This makes the 12 sheets a dependency chain rather than 12 isolated tools:

```
Rates_Master ────────┬──────────────────────────────────────────────┐
                      │                                              │
               03_Mortars (base-cost only, no markup)                │
                      │                                              │
     ┌────────────────┼───────────────┬──────────────┐               │
06_Masonry      07_Stone_Work   08_Cladding    11_Flooring            │
                                                                      │
02_Earth_Work, 04_Concrete_Work, 05_RCC_Work, 09_Wood_PVC,           │
10_Steel_Work, 12_Roofing  ── pull straight from Rates_Master ───────┘
01_Carriage_of_Materials — its own two sub-patterns (see below)

Labour_Machinery_Productivity ── advisory input to every builder sheet (suggests a labour team; not a compute dependency)
```

**Common builder-sheet anatomy** (identical skeleton on all 12, so the workbook feels like one tool, not twelve):
1. Item header — description, output unit, output quantity basis (mirrors "Details of cost for 1 cum / 10 sqm / 100 Nos" in the book).
2. MATERIAL rows — add-a-row picker (`Rates_Master` code, or another builder sheet's computed rate), coefficient input, rate pulled automatically, amount computed.
3. LABOUR rows — same mechanic, restricted to labour/machinery codes, pre-populatable from the matching work type in `Labour_Machinery_Productivity`.
4. Sundries — a manual input line, guided by `Sundries_Reference`, not auto-derived, since the book gives no reproducible formula for it.
5. Markup block — reads its default steps and % from `Global_Factors`, shows W→X→Y→Z exactly as the book does, but each step has an on/off + override cell local to the item (needed for sheet 01, see below).
6. Final rate output, "Say" rounding, and a running library/list of every custom item built on that sheet — this is what makes the sheet self-contained after the base file is closed.

**Self-containment mechanic:** `Rates_Master`, `Labour_Machinery_Productivity`, `Sundries_Reference` and `Global_Factors` all hold *copied/derived values*, not links to the original workbook — so once built, the new workbook has zero dependency on `CivilDAR_2019_Vol_1_Converted.xlsx` ever being open again, satisfying your "no reference or need of base file" requirement.

---

## 3. Per-sheet Problem & Solution Statements

### 01 — Carriage of Materials
**Problem:** An estimator needs to price transporting a new material/waste type that isn't one of the book's listed carriage items, by mechanical or manual means, at a lead/lift the book doesn't cover.
**Solution:** Two builder modes on this one sheet, matching the two patterns found: (a) *Mechanical trip* mode — hire charges + fuel + labour per trip, divided by trips/capacity, with **only the CPOH step** active by default (Water/GST/Cess off, matching the book); (b) *Manual labour* mode — a small lead-distance table (base 50 m + additional 50 m increments) driven by Beldar/Coolie day-rates, reproducing the book's matrix rather than the standard item template.

### 02 — Earth Work
**Problem:** Pricing an earthwork operation (excavation, filling, dressing, compaction) at a labour mix or output rate not printed in the book.
**Solution:** A labour-and-machinery-only builder (no MATERIAL block required, matching every Earth Work item observed) — Beldar/Coolie/Bhisti/Chowkidar day-coefficients plus optional roller/machinery hire, full W→X→Y→Z markup chain, output per 100 sqm / 10 cum / cum as chosen.

### 03 — Mortars
**Problem:** A new mortar mix (a ratio, or a different sand/cement type) is needed as a **building block** for other custom items — not as a final priced item.
**Solution:** Builder stops at Material + Labour + Sundries → **Cost of unit, with the markup chain switched OFF by default** (matching the book's own convention), because this sheet's whole purpose is to feed a pre-markup rate into Masonry/Stone/Cladding/Flooring builders. Its output list is the primary "other sheet's rate" source those four builders will pick from.

### 04 — Concrete Work
**Problem:** Pricing a plain cement concrete mix/grade combination not in the printed schedule.
**Solution:** Standard full-anatomy builder — aggregate/sand/cement material lines with carriage sub-lines, Mason/Beldar/Bhisti labour, mixer/vibrator hire, Sundries, full W→X→Y→Z chain — mirroring the self-contained (non-referencing) structure every Concrete Work item uses.

### 05 — RCC Work
**Problem:** Pricing a reinforced-concrete item (grade/placement combination) not printed, excluding shuttering and reinforcement (as the book itself excludes them here).
**Solution:** Same anatomy as Concrete Work — this sheet does not reference the Concrete builder's output (RCC items in the book recompute material+labour themselves rather than importing a Concrete rate), so its picker defaults to raw Reference_Rates, consistent with what's observed.

### 06 — Masonry Work
**Problem:** Pricing a brick/block masonry item in a mortar ratio or brick type not printed.
**Solution:** Builder whose MATERIAL picker explicitly offers **03_Mortars' computed rates** as a component (this is the clearest example in the book of the cross-sheet pattern from your Q2 answer), alongside brick/carriage codes from Reference_Rates, Mason/Coolie/Bhisti labour, full markup chain.

### 07 — Stone Work
**Problem:** Pricing a stone masonry item (rubble, ashlar, etc.) at a stone type/mortar ratio not printed.
**Solution:** Same shape as Masonry — Mortar-sheet rate as a MATERIAL option, plus stone/bond-stone/carriage codes and a "backing concrete" style L.S. line if needed (the book uses a similar L.S. line for "Cement concrete 1:6:12" inside some Stone Work items), full markup chain.

### 08 — Cladding Work
**Problem:** Pricing a stone/marble cladding or veneer item in a slab size/finish not printed.
**Solution:** MATERIAL picker offers Mortar-sheet rates (both bedding mortar and pointing mortar, as the book uses two different Mortar items in the same Cladding item), stone/marble codes with a wastage-% helper on quantity, ornamental-work labour grades, Scaffolding as an optional Sundries-style line, full markup chain.

### 09 — Wood and PVC Work
**Problem:** Pricing a wood/PVC frame, shutter or fitting in a timber species/section size not printed.
**Solution:** Standard self-contained builder — timber-in-scantling + carriage material lines with a wastage-% helper, Carpenter/Beldar labour, full markup chain, output convertible between the item's working unit (e.g. per door) and a standard unit (cum) as the book itself does at the end of each item.

### 10 — Steel Work
**Problem:** Pricing a structural steel fabrication item (section/truss/plate combination) not printed.
**Solution:** Standard builder for the structural material/carriage/labour lines. Its picker also reserves a "priming coat, per `13_Finishing_Work`" component the same way Masonry reserves a Mortar-rate component — per your decision, treated as a normal cross-sheet dependency rather than a hard limitation, on the assumption that a Finishing builder sheet will be available. Since that sheet's underlying data isn't in this Volume-1 file yet, the reserved component starts out holding only the literal number the book does show (₹50.70/sqm) as an interim value, to be replaced by a live reference once `13_Finishing_Work` exists.

### 11 — Flooring
**Problem:** Pricing a flooring item (brick-on-edge, tile, etc.) in a mortar ratio not printed.
**Solution:** Same cross-sheet pattern as Masonry — Mortar-sheet rate as a MATERIAL option, brick/tile + carriage codes, Mason/Beldar/Coolie/Bhisti labour, full markup chain, per-sqm output.

### 12 — Roofing
**Problem:** Pricing a roofing item (sheeting, fixing hardware combination) not printed.
**Solution:** Standard self-contained builder — sheet/bolt/washer/hook material lines (each with the book's own weight/count derivation style as a helper, not a hidden formula), Mistry/Carpenter/Beldar/Painter labour including a priming-and-painting sub-section, full markup chain.

---

## 4. Decisions locked in (round 2)

1. **Sundries** — manual entry per item on the builder sheet, guided by a dedicated `Sundries_Reference` lookup sheet.
2. **Rates infrastructure** — one combined `Rates_Master` sheet (materials+labour+machinery together) plus a separate `Labour_Machinery_Productivity` sheet organized by work type, so you can see what labour a job like brickwork actually needs.
3. **Steel Work → Finishing dependency** — treated as a normal forward cross-sheet reference (like Masonry → Mortars), assumed available; interim literal value used until a `13_Finishing_Work` builder sheet exists.
4. **`13_Finishing_Work`** — left as a reserved, empty placeholder sheet (name/position held in the workbook) rather than built out now, until Finishing source data is brought in separately.

All open decisions are resolved. Next step: translate each per-sheet solution statement above into the actual sheet layout (cell-by-cell) before any formulas are written.
