# Earthwork First-Principles Learning Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace coefficient-reading notes with source-labelled, 8-hour-shift and gang-based derivations for every eligible resource row.

**Architecture:** Extend the workbook editor with a resource-derivation model that consumes verified CPWD productivity records where available. Each visible I-column line expresses the calculation; each D-cell Note records the source norm, the calculation, the teaching interpretation, conditions, and when the norm changes. No reconstructed assumption is represented as an official CPWD rule.

**Tech Stack:** Bundled Python, `openpyxl`, `unittest`, Excel Notes/comments, XLSX export.

---

### Task 1: Model source-backed productivity derivations

**Files:** Modify `scripts/support_earth_learning_edit.py`; modify `tests/test_support_earth_learning_edit.py`.

- [ ] Add failing tests for a representative labour norm and machinery norm. Assert the derivation data has `source_norm`, `role`, `gang_or_machine`, `task_hours`, `shift_hours=8`, `coefficient`, `unit`, `batch_quantity`, `batch_unit`, and `interpretation_label`.

- [ ] Run `scripts/run_support_earth_learning_test.ps1`; expect failure because no first-principles derivation model exists.

- [ ] Implement `build_derivation_catalog(earth_sheet)` and `derive_resource_norm(resource, item_context, catalog)`. Only create a derivation when a unique source key matches the support item/resource. Preserve source evidence verbatim enough to audit it. For a labour derivation calculate `gang_size * task_hours / 8`; for machinery calculate `machine_hours / 8`; assert the result equals the existing D coefficient within 0.001.

- [ ] If source gives only a final coefficient, return a structured record with `source_norm` and no invented hours. Its output label must be `Teaching interpretation, not a published CPWD rule`.

- [ ] Re-run focused tests; expect PASS. Commit `feat: model earthwork first-principles norms`.

### Task 2: Render the three-layer learning experience

**Files:** Modify `scripts/support_earth_learning_edit.py`; modify `tests/test_support_earth_learning_edit.py`.

- [ ] Add failing assertions for item 2.1.1 and one machinery item: I-column contains `8-hour shift`, the visible arithmetic equals the D quantity, and the D-cell Note includes headings `CPWD fixed norm / source evidence`, `Calculation`, `Engineering interpretation for learning`, `Boundary conditions`, and `When the norm changes`.

- [ ] Implement the item work-method context line: physical operation, standard batch, 8-hour shift, and applicable method constraints. Add a short gang-system explanation for multi-resource items.

- [ ] Write visible I-cell text in this exact shape when derivable: `CPWD fixed norm: [role]. Calculation: [gang/machine] × [task-hours] ÷ 8-hour shift = [D quantity] [unit] per [batch]. Equivalent productivity: [batch ÷ D quantity].` When no direct derivation exists, state the fixed coefficient then use the explicit teaching-interpretation label.

- [ ] Create/update the D-cell Note with the five required labelled sections. Never overwrite D numeric values, formulas, or number formats. Use a pale-blue wrapped I-cell style and preserve A:H formats.

- [ ] Re-run tests; expect PASS for visible derivations, Notes, labels, and unchanged quantities. Commit `feat: teach earthwork norms from first principles`.

### Task 3: Verify source integrity and deliver

**Files:** Modify only `outputs/earthwork-first-principles/CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_1_Latest.xlsx`.

- [ ] Add regression tests covering first, middle, and final items: all non-target sheets unchanged; `02_Earth_Work` values, formulas, styles, and dimensions unchanged; A:H support values/formulas/styles unchanged except approved description/Note changes; D Notes and I derivations exist for every eligible row; markup groups exclude final unit-rate rows.

- [ ] Run `node container_tools/mark_artifact_operation_started.mjs --operation-kind edit --expected-output-count 1 --output-format xlsx` exactly once immediately before the final tracked save.

- [ ] Export the one workbook to `outputs/earthwork-first-principles/CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_1_Latest.xlsx`, then run the complete tests against both source and exported output.

- [ ] Scan formulas for literal Excel error tokens. If no recalculation engine is available, state that formulas are preserved and Excel will recalculate on open; do not claim a cached-value calculation check.

- [ ] Commit code/tests only with `feat: add first-principles earthwork learning`; do not commit the user workbook.

## Self-review

- Tasks 1–2 distinguish fixed CPWD evidence from instructional reconstruction and make the eight-hour/gang logic inspectable.
- Task 3 protects all original calculations and reference sheets while validating first, middle, and final examples.
