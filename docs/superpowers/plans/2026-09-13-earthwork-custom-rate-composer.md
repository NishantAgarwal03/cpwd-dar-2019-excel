# Earthwork Custom-Rate Composer Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a controlled-keyword custom unit-rate composer above `02_support_earth_work`.

**Architecture:** A workbook-owned keyword catalogue resolves a base item plus deduction, add-on, conditional-extra, and overlap-warning relationships. The composer builds an independent resource schedule and shows the closest CPWD item only as a benchmark.

**Tech Stack:** Bundled Python, `openpyxl`, `unittest`, XLSX export.

---

### Task 1: Correct the 2.24 catalogue definition

**Files:** Modify `scripts/support_earth_learning_edit.py`; modify `tests/test_support_earth_learning_edit.py`.

- [ ] Add failing tests asserting that 2.24.1 resolves to a 20% difficult-condition extra for water/liquid mud and 2.24.2 resolves to a 25% foul-position extra, both measured per metre depth; neither may be timbering-only.
- [ ] Run `python -m unittest tests.test_support_earth_learning_edit`; expect failure against the present mapping.
- [ ] Implement a catalogue record with `relationship_type`, `base_scope`, `percent`, `condition`, and `measurement_basis`; update the supported-sheet presentation to the official definitions.
- [ ] Re-run tests; expect PASS. Commit `fix: correct earthwork difficult-condition extras`.

### Task 2: Implement controlled keyword resolution

**Files:** Modify `scripts/support_earth_learning_edit.py`; modify `tests/test_support_earth_learning_edit.py`.

- [ ] Add failing resolver tests:

```python
assert resolve(["banking", "no power roller", "no watering"]).relationships == ["base", "deduct", "deduct"]
assert resolve(["rough excavation", "under water"]).relationships == ["base", "conditional_extra"]
assert "overlap" in resolve(["surface excavation", "grass clearing"]).warnings
```

- [ ] Implement `resolve_keywords(keywords)` using only controlled catalogue entries. Return the base item, deductions, additions, conditional extras, incompatibilities, warnings, and a composed plain-language scope.
- [ ] Re-run tests; expect PASS. Commit `feat: resolve earthwork custom-rate keywords`.

### Task 3: Add the workbook composer above the catalogue

**Files:** Modify `scripts/support_earth_learning_edit.py`; modify `tests/test_support_earth_learning_edit.py`.

- [ ] Add failing tests that the top composer contains controlled keyword cells, resolved relationship rows, composed scope, included/excluded resources, custom unit rate, CPWD benchmark, and variance; catalogue rows below remain intact.
- [ ] Implement the top-sheet layout and validation lists. Populate the custom schedule from resolver output, calculate deductions as negative schedule impacts, and calculate conditional extras as `selected percentage × qualifying base rate`. Keep project quantities out of the unit-rate calculation.
- [ ] Re-run tests; expect PASS. Commit `feat: add earthwork custom rate composer`.

### Task 4: Validate the three agreed cases and export

**Files:** Modify only `outputs/earthwork-custom-composer/CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_1_Latest.xlsx`.

- [ ] Test banking plus no-power-roller/no-watering deductions; rough excavation plus under-water conditional extra; and surface excavation plus grass-clearing overlap warning.
- [ ] Immediately before final save, run `node container_tools/mark_artifact_operation_started.mjs --operation-kind edit --expected-output-count 1 --output-format xlsx` exactly once.
- [ ] Export one workbook, preserve the sheet-8 formula compatibility fix, scan for literal formula errors, and run all unit tests against source and output.
- [ ] Commit code/tests only with `feat: compose custom earthwork unit rates`; do not commit the workbook.

## Self-review

- Tasks 1–2 implement correct controlled classification.
- Task 3 keeps the composer on the existing support sheet and makes custom calculations independent of benchmark coefficients.
- Task 4 proves each user-supplied composition case and preserves the repaired workbook.
