# Earthwork custom-rate composer design

## Purpose

Add a controlled description composer to `02_support_earth_work` so a user can create an auditable custom unit rate for an amended work description. The composer selects and combines CPWD base, deduction, add-on, and conditional-extra relationships; it does not merely copy a close published rate.

## User workflow

The user selects controlled keywords in natural description order. They do not need to choose a base work family first. Example:

`banking` → `all kinds of soil` → `50 m lead` → `1.5 m lift` → `no power roller` → `no watering`.

The composer then resolves the compatible base item and any permitted relationships, blocks incompatible selections, states the composed scope in plain language, and builds a custom unit-rate analysis.

## Relationship rules

- **Base item:** establishes the standard work method and common measurement basis.
- **Deduction:** removes a named allowance from the base method. Example: banking rate less power-roller and watering deductions.
- **Add-on:** adds a distinct operation that is demonstrably outside the base scope.
- **Conditional percentage extra:** applies only where the selected difficult condition qualifies. It is shown as a percentage applied to the named base, not as a normal resource row.
- **Overlap warning:** blocks or warns when a selected operation is already included in the base scope, such as clearing already included in surface excavation.

## Official item correction

The supported catalogue must correct item 2.24:

- 2.24.1: 20% extra for work in/under water or liquid mud, including pumping, measured per metre depth for the qualifying quantity.
- 2.24.2: 25% extra for work in/under foul position, including pumping, measured per metre depth for the qualifying quantity.

These extras apply to qualifying work items and are not timbering-only. The prior timbering-only mapping in `02_support_earth_work` is erroneous.

## Workbook layout

Place the composer above the existing item schedule on `02_support_earth_work`. It contains:

1. controlled keyword selectors;
2. resolved base/deduction/add-on/conditional relationship list;
3. composed work description;
4. included and excluded resource/productivity schedule;
5. custom unit rate and a separate CPWD benchmark/variance display.

The existing schedule stays below as the reference catalogue.

## Calculation principle

The custom schedule owns the calculation. A close CPWD item is a benchmark only: its resource coefficients are not silently copied into the custom rate. Every included/deducted resource is traceable to the selected operation and first-principles productivity rule.

## Validation

Test the three agreed cases:

1. Banking base with no power roller and no watering resolves the base plus both deductions.
2. Rough excavation/banking with a difficult-condition keyword applies the correct 2.24 percentage relationship only as a conditional extra.
3. Surface excavation plus grass/rubbish clearing raises an overlap warning when the base scope already includes that work.

Verify formula continuity, keyword compatibility, relationship traceability, and unchanged catalogue rows below the composer.
