"""Controlled relationships for composing CPWD earthwork custom rates.

The catalogue deliberately separates an official percentage extra from the
base work it qualifies.  It is not a resource schedule and must never be
treated as a timbering-only allowance.
"""

from __future__ import annotations

from typing import Any


_DEPTH_BASIS = "metre depth from sub-soil water level to the centre of gravity of qualifying work"

# This taxonomy is intentionally broader than the initial controlled-keyword
# resolver.  It marks the CPWD Earth Work base families already visible in the
# support sheet and makes clear where later resource/productivity extensions
# belong without inventing productivity norms in the composer itself.
BASE_WORK_FAMILIES: tuple[str, ...] = (
    "Excavation",
    "Banking/Embankment",
    "Filling",
    "Timbering/Shoring/Planking",
    "Site Clearance/Surface Preparation",
    "Chemical Anti-Termite Treatment",
    "Difficult-condition extras",
)

EARTHWORK_COMPOSER_CATALOGUE: dict[str, dict[str, Any]] = {
    "2.24.1": {
        "item_code": "2.24.1",
        "relationship_type": "conditional_extra",
        "base_scope": "each applicable earthwork item",
        "percent": 20,
        "condition": "in or under water and/or liquid mud, including pumping out water as required",
        "measurement_basis": _DEPTH_BASIS,
    },
    "2.24.2": {
        "item_code": "2.24.2",
        "relationship_type": "conditional_extra",
        "base_scope": "each applicable earthwork item",
        "percent": 25,
        "condition": "in or under foul position, including pumping out water as required",
        "measurement_basis": _DEPTH_BASIS,
    },
}


def _visible_text(item: dict[str, Any]) -> str:
    """Render the official relationship in a student-readable reference row."""
    return (
        f"Item {item['item_code']}: {item['percent']}% extra over {item['base_scope']} "
        f"for work {item['condition']}. Apply only to the qualifying quantity. "
        f"Measurement: {item['measurement_basis']}."
    )


def calculate_difficult_condition_extra(item_code: str, selected_base_rate: float) -> dict[str, Any]:
    """Resolve 2.24 against the selected base rate, never as a flat rupee line.

    The resulting rate is for one unit of *qualifying* work.  Quantity/depth
    allocation remains a measurement decision outside this unit-rate resolver.
    """
    try:
        item = EARTHWORK_COMPOSER_CATALOGUE[item_code]
    except KeyError as error:
        raise ValueError(f"Unknown difficult-condition item: {item_code}") from error
    if item["relationship_type"] != "conditional_extra":
        raise ValueError(f"Item {item_code} is not a difficult-condition percentage extra")
    if not isinstance(selected_base_rate, (int, float)) or isinstance(selected_base_rate, bool):
        raise TypeError("selected_base_rate must be numeric")
    if selected_base_rate < 0:
        raise ValueError("selected_base_rate cannot be negative")
    percent = item["percent"]
    return {
        **item,
        "selected_base_rate": selected_base_rate,
        "extra_rate": selected_base_rate * percent / 100,
    }


def apply_difficult_condition_reference(sheet) -> None:
    """Insert or replace only the two visible 2.24 reference lines.

    The support schedule is a catalogue/reference sheet.  This function leaves
    every unrelated item and rate untouched, while making the composer-facing
    relationship readable in column A.  A later composer renderer can use the
    structured catalogue rather than parsing this display text.
    """
    if "Sub-Head 2.0 — EARTH WORK" in str(sheet.cell(1, 1).value or ""):
        return
    targets = {f"Item {code}": item for code, item in EARTHWORK_COMPOSER_CATALOGUE.items()}
    found: set[str] = set()
    for row in range(1, sheet.max_row + 1):
        value = str(sheet.cell(row, 1).value or "")
        for prefix, item in targets.items():
            if value.startswith(prefix):
                sheet.cell(row, 1).value = _visible_text(item)
                _rewrite_difficult_condition_block(sheet, row, item)
                found.add(item["item_code"])
                break

    for code, item in EARTHWORK_COMPOSER_CATALOGUE.items():
        if code not in found:
            sheet.append([_visible_text(item)])


def _rewrite_difficult_condition_block(sheet, heading_row: int, item: dict[str, Any]) -> None:
    """Correct every display value in the existing compact 2.24 item block.

    The legacy support sheet reserves six rows for an item.  Values are
    replaced in place so all pre-existing styles, row heights, merges and
    outline settings remain untouched.
    """
    percent = item["percent"]
    condition = item["condition"]
    measurement = item["measurement_basis"]
    sheet.cell(heading_row + 1, 1).value = (
        f"Conditional extra for qualifying work: {percent}% of the qualifying selected base rate; "
        f"not a flat Rs rate. {condition.capitalize()}. Measurement: {measurement}."
    )
    sheet.cell(heading_row + 3, 1).value = (
        f"Composer rule: select the applicable base item first, then add {percent}% only for the quantity "
        f"executed {condition}. Depth is measured from the sub-soil water level to the centre of gravity "
        "of the qualifying work. No separate labour/material resource line is added."
    )
    sheet.cell(heading_row + 4, 1).value = "Composer percentage calculation (not a flat Rs rate)"
    sheet.cell(heading_row + 4, 6).value = f"{percent}% of the qualifying selected base rate"
