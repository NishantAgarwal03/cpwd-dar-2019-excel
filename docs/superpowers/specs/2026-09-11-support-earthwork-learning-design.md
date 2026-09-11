# Support earthwork learning design

## Purpose

Make the `02_support_earth_work` worksheet a usable teaching aid for CPWD earthwork rate analysis. Students should understand both the accepted calculation basis and how to adapt the analysis to a changed quantity or site condition.

## Scope

- Edit only `02_support_earth_work` in the supplied workbook.
- Keep `02_earth_work` unchanged as the technical reference for explanations and coefficients.
- Apply the same student-facing structure to every supported item on `02_support_earth_work`.

## Student-facing layout

Each item keeps its headline calculation visible and adds the following teaching elements:

1. `A4` explains why the rate is normalised to 100 sqm and how the resulting rate scales to another quantity.
2. `D7:D8` state the productivity coefficient and its calculation rule or basis.
3. `I7:I8` contain concise, visible notes explaining the coefficient, its source or operational meaning, and how a student may recalculate or vary it for different site conditions.
4. Repeated arithmetic rows corresponding to rows 10–18 are placed in a common Excel outline level. Students can expand one item’s workings when reviewing the math, or collapse the shared detail level across the sheet.

## Interaction and calculation flow

The visible worksheet leads students through: 100-sqm normalisation, productivity basis, and the resulting rate. Expanded detail supplies the repeated arithmetic for audit or practice without making it the default reading view. Clearly identified inputs remain editable so a student can test alternative quantities or conditions while preserving the original worked example.

## Source and preservation rules

- Derive explanations from `02_earth_work` and do not alter that sheet.
- Preserve existing item formulas, values, formatting conventions, and workbook features unless an edit is required for the teaching layout.
- Do not move repeated calculations to another worksheet.

## Validation

For representative first, middle, and final items, confirm that the visible explanation is present, values and formulas remain intact, the outline expands and collapses correctly, and the sheet stays legible when detail is collapsed and expanded. Scan affected ranges for formula errors and save a single edited workbook output.
