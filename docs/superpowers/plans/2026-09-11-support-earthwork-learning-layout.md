# Support Earthwork Learning Layout Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make every item in `02_support_earth_work` a compact, self-explanatory rate-analysis lesson while preserving its calculations.

**Architecture:** A focused editor reads the existing support sheet and `02_Earth_Work`, discovers all 84 item blocks, writes visible batch and productivity notes in column I, and sets Excel row outlines for recurring statutory markup calculations. It never modifies `02_Earth_Work`.

**Tech Stack:** Node.js, `@oai/artifact-tool`, Node assertions, XLSX output.

---

## File structure

- Create `scripts/support_earth_learning_edit.mjs`: imports the existing workbook, performs targeted learning-layout edits, recalculates, checks, and exports.
- Create `tests/support_earth_learning_edit.test.mjs`: regression tests on a disposable workbook output.
- Modify `CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_1_Latest.xlsx`: only `02_support_earth_work`.
- Preserve `scripts/support_builder_earth_v2.py`: it documents original construction and is not run or edited.

### Task 1: Discover and classify item blocks

**Files:** Create `scripts/support_earth_learning_edit.mjs`; create `tests/support_earth_learning_edit.test.mjs`.

- [ ] Write a failing discovery test.

```js
import assert from "node:assert/strict";
import { loadWorkbook, collectItemBlocks } from "../scripts/support_earth_learning_edit.mjs";
const wb = await loadWorkbook(process.env.SUPPORT_EARTH_INPUT);
const blocks = await collectItemBlocks(wb.worksheets.getItem("02_support_earth_work"));
assert.equal(blocks.length, 84);
assert.deepEqual(blocks.slice(0, 3).map(x => x.itemId), ["2.1.1", "2.1.2", "2.2.1"]);
assert.equal(blocks.at(-1).itemId, "2.38");
```

- [ ] Run `node tests/support_earth_learning_edit.test.mjs`; expect failure because the editor module is absent.

- [ ] Implement `loadWorkbook` using `FileBlob.load` and `SpreadsheetFile.importXlsx`, and implement block discovery from column-A values matching `/^Item ([^|\s]+)/`. Define every block end as two rows before the following item heading, or the end of the used range for item 2.38. Do not alter or export the workbook yet.

- [ ] Re-run the test; expect PASS with 84 blocks.

- [ ] Commit: `git add scripts/support_earth_learning_edit.mjs tests/support_earth_learning_edit.test.mjs` then `git commit -m "test: map support earthwork item blocks"`.

### Task 2: Add visible basis and productivity teaching notes

**Files:** Modify `scripts/support_earth_learning_edit.mjs`; modify `tests/support_earth_learning_edit.test.mjs`.

- [ ] Add failing assertions:

```js
assert.match(String(support.getRange("A4").values[0][0]), /100 sqm.*normal/i);
assert.match(String(support.getRange("I7").values[0][0]), /Beldar.*6\.8.*100 sqm/i);
assert.match(String(support.getRange("I8").values[0][0]), /Coolie.*5\.6.*100 sqm/i);
assert.match(String(support.getRange("A42").values[0][0]), /10 cum.*standard batch/i);
assert.match(String(support.getRange("A470").values[0][0]), /percentage/i);
```

- [ ] Run the test; expect FAIL because column-I notes and revised batch basis are absent.

- [ ] Implement `describeBatchBasis(baseQty, unit)` with these exact cases:

```js
if (baseQty === 100 && unit === "sqm") return "Cost details are normalised to 100 sqm, the CPWD standard output batch. Divide the batch cost by 100 for the rate per sqm; multiply that rate by the actual measured area.";
if (baseQty === 10 && unit === "cum") return "Cost details are normalised to 10 cum, the CPWD standard output batch. Divide the batch cost by 10 for the rate per cum; multiply that rate by the actual excavated or filled volume.";
if (unit === "%") return "This is a percentage extra over its referenced base item, not an independently measured resource batch. Apply the stated percentage to the applicable base-item rate.";
return `Cost details use the CPWD standard batch of ${baseQty} ${unit}. Divide the batch cost by ${baseQty} for the unit rate, then multiply by the measured project quantity.`;
```

- [ ] Replace each item’s batch-description cell (the row immediately following its Item heading) with the existing label plus this explanation. For each non-section resource row, write to column I: `Productivity basis: [resource] = [quantity] [unit] per [base quantity] [base unit]. This coefficient is the crew/plant time allowance used to produce the standard batch; revise it only when the site condition, crew method, lead, lift, or material condition differs.` Add matching operational context from the resource/productivity catalog on `02_Earth_Work` when an exact item-resource record matches. Do not change D/E/F/H formulas or values.

- [ ] Set I width to 58. Use `Learning note / productivity basis` in every item’s header row, with small wrapped, top-aligned dark text on a restrained pale-blue fill. Preserve all A:H formatting and column widths.

- [ ] Re-run the assertions; expect PASS. Commit with message `feat: explain earthwork batch and productivity basis`.

### Task 3: Group repeated statutory markup rows

**Files:** Modify `scripts/support_earth_learning_edit.mjs`; modify `tests/support_earth_learning_edit.test.mjs`.

- [ ] Add failing tests that identify item 2.1.1 markup rows as 10 through 17, and verify equivalent detected groups for items 2.2.1 and the final standard-rate item.

- [ ] Run the test; expect FAIL because no outline metadata exists.

- [ ] For every standard-rate item, locate the contiguous rows starting at `Add: Water charges` and ending at `Total cost for` (inclusive). Assign one common outline level to that row range and enable summary-below behavior. Leave groups expanded by default. This lets students inspect one calculation or collapse all recurring calculation groups using the outline-level control. Do not group item headings, resource rows, the final per-unit/Say rate, DSR comparison, or project-quantity row. Do not create empty groups for pure-reference or percentage items.

- [ ] Export a disposable copy and assert outline metadata for the first, middle, and final standard-rate groups. Re-run tests; expect PASS. Commit with message `feat: group repeated earthwork markup rows`.

### Task 4: Export and verify the deliverable

**Files:** Modify only the exported workbook at `outputs/support-earthwork-learning/CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_1_Latest.xlsx`.

- [ ] Immediately before the first write, run `node container_tools/mark_artifact_operation_started.mjs --operation-kind edit --expected-output-count 1 --output-format xlsx` exactly once.

- [ ] Run the editor with the current workbook as input and the above output path. Call `workbook.recalculate()` once before export.

- [ ] Inspect values and formulas in `02_support_earth_work!A3:I20`, `A41:I64`, and the final item block. Confirm F-column formulas and H-column project-amount formulas match the input. Scan `02_support_earth_work` for `#REF!|#DIV/0!|#VALUE!|#NAME\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!`; report pre-existing errors but do not repair unrelated content.

- [ ] Render `A1:I64`, a middle sample, and the final sample at normal zoom. Confirm readable I-column notes, no clipped content, and usable expanded/collapsed outline states. Confirm `02_Earth_Work` was not changed.

- [ ] Run `SUPPORT_EARTH_INPUT="outputs/support-earthwork-learning/CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_1_Latest.xlsx" node tests/support_earth_learning_edit.test.mjs`; expect PASS for all 84 items.

- [ ] Commit only editor and test code with `git commit -m "feat: add earthwork learning guidance"`. Do not commit the user’s workbook unless separately requested. Deliver the single exported `.xlsx`.

## Self-review

- Task 2 covers batch normalisation and editable productivity interpretation for every item.
- Task 3 resolves the user’s example rows 10–18 as the repetitive markup chain, keeping final per-unit/Say outcomes visible.
- Task 4 preserves formulas and the source sheet, checks errors and visual usability, and produces one workbook artifact.
