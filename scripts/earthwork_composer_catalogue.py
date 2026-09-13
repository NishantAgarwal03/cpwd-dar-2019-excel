"""Controlled relationships for composing CPWD earthwork custom rates.

The catalogue deliberately separates an official percentage extra from the
base work it qualifies.  It is not a resource schedule and must never be
treated as a timbering-only allowance.
"""

from __future__ import annotations

from typing import Any


_DEPTH_BASIS = "metre depth from sub-soil water level to the centre of gravity of qualifying work"

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


def apply_difficult_condition_reference(sheet) -> None:
    """Insert or replace only the two visible 2.24 reference lines.

    The support schedule is a catalogue/reference sheet.  This function leaves
    every unrelated item and rate untouched, while making the composer-facing
    relationship readable in column A.  A later composer renderer can use the
    structured catalogue rather than parsing this display text.
    """
    targets = {f"Item {code}": item for code, item in EARTHWORK_COMPOSER_CATALOGUE.items()}
    found: set[str] = set()
    for row in range(1, sheet.max_row + 1):
        value = str(sheet.cell(row, 1).value or "")
        for prefix, item in targets.items():
            if value.startswith(prefix):
                sheet.cell(row, 1).value = _visible_text(item)
                found.add(item["item_code"])
                break

    for code, item in EARTHWORK_COMPOSER_CATALOGUE.items():
        if code not in found:
            sheet.append([_visible_text(item)])
