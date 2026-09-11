# Earthwork first-principles learning design

## Purpose

Turn `02_support_earth_work` from a rate table with annotations into a teaching workbook. A student must be able to see what each labour or machinery coefficient represents, how it is calculated from a fixed CPWD norm, and which parts are published evidence versus engineering interpretation.

## Scope

- Edit only `02_support_earth_work`.
- Retain `02_Earth_Work` as the unmodified reference sheet.
- Preserve all existing resource quantities, rates, formula chains, project-quantity fields, and final rates.
- Apply the same learning structure to every eligible resource row in every item.

## Three-layer teaching model

### 1. Work-method context for each item

Each item starts with a compact explanation of the physical work, its CPWD standard batch, the fixed eight-hour shift, and relevant boundaries such as soil, lead, lift, depth, compaction, or equipment method. It answers: what operation does this rate analysis produce?

### 2. Visible derivation beside each resource

The I-column shows the full fixed-norm derivation for every labour and machinery resource. Labour follows:

`gang size × task-hours ÷ 8-hour shift = labour-days per standard batch`.

It also shows the inverse productivity where meaningful:

`standard batch ÷ resource-days = output per resource-day`.

Machinery follows:

`machine-hours required ÷ 8-hour shift = machine-days per standard batch`.

The visible derivation states the crew or plant role in the operation, rather than treating a coefficient as an unexplained input.

### 3. Quantity-cell learning card

Every D-column quantity cell receives a detailed Excel comment/Note. Its sections are:

1. **CPWD fixed norm / source evidence** — the published coefficient or source-supported record.
2. **Calculation** — the gang, activity hours, and eight-hour-shift conversion used to arrive at the quantity.
3. **Engineering interpretation for learning** — what the worker or machine does and why that allocation is plausible.
4. **Boundary conditions** — soil, lead, lift, method, machine, depth, or compaction conditions that make the norm applicable.
5. **When the norm changes** — clearly states which site changes require a new analysis.

If a detailed activity-hour reconstruction is not directly supported by the CPWD source, it is explicitly labelled **Teaching interpretation, not a published CPWD rule**.

## Gang-system explanation

Each multi-resource item includes a short explanation of how the gang works as one production system. It links roles such as excavation/cutting, trimming, collection, spoil handling, watering, supervision, and rolling so students understand why several allowances appear together.

## Presentation rules

- Use visible I-column derivations for the core learning path; comments provide the deeper audit trail.
- Label every claim either **CPWD fixed norm / source evidence** or **Engineering interpretation for learning**.
- Keep final unit rates visible when repeated markup rows are collapsed.
- Never overwrite the numeric D quantity with explanatory text.
- Never present a teaching reconstruction as an official CPWD published rule.

## Validation

Verify first, middle, and final item blocks for visible derivations, D-cell comments, source/interpreted labels, correct outline boundaries, and unchanged formulas/values. Confirm a student can trace one labour coefficient and one machinery coefficient from work method through gang-hours or machine-hours to the quantity used in the rate calculation.
