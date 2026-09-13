"""Resolve controlled earthwork keywords into a custom-rate composition.

This module deliberately does not calculate rates or alter a worksheet.  It
turns a student's ordered, controlled selections into the base work and any
deduction, addition or qualifying percentage relationship that a later
workbook renderer can display and calculate.
"""

from __future__ import annotations

from collections.abc import Iterable
from typing import Any

from earthwork_composer_catalogue import EARTHWORK_COMPOSER_CATALOGUE


_COMPONENTS: tuple[dict[str, Any], ...] = (
    {
        "item_code": "2.1.1",
        "relationship_type": "base",
        "keywords": ("surface excavation", "all kinds of soil"),
        "requires_all": ("surface excavation", "all kinds of soil"),
        "description": "Surface excavation, all kinds of soil.",
    },
    {
        "item_code": "2.2",
        "relationship_type": "base",
        "keywords": ("rough excavation and banking",),
        "requires_all": ("rough excavation and banking",),
        "description": "Rough excavation and banking in layers.",
    },
    {
        "item_code": "2.3.1",
        "relationship_type": "base",
        "keywords": ("banking excavated earth", "all kinds of soil"),
        "requires_all": ("banking excavated earth", "all kinds of soil"),
        "description": "Banking excavated earth in layers, all kinds of soil.",
    },
    {
        "item_code": "2.4",
        "relationship_type": "deduction",
        "keywords": ("no power roller",),
        "requires_all": ("no power roller",),
        "description": "Deduct where the required 8-tonne power-roller operation is omitted.",
    },
    {
        "item_code": "2.5",
        "relationship_type": "deduction",
        "keywords": ("no watering",),
        "requires_all": ("no watering",),
        "description": "Deduct where watering of excavated earth for banking is omitted.",
    },
    {
        "item_code": "2.32",
        "relationship_type": "addition",
        "keywords": ("clearing grass",),
        "requires_all": ("clearing grass",),
        "description": "Clearing grass and removal of rubbish.",
    },
)

_CONDITIONAL_KEYWORDS = {
    "water or liquid mud": "2.24.1",
    "foul position": "2.24.2",
}

_BANKING_BASES = frozenset(("2.2", "2.3.1"))

_OVERLAP_RULES = {
    frozenset(("2.1.1", "2.32")): (
        "Review scope before pricing items 2.1.1 and 2.32 together: surface excavation may already "
        "remove grass/rubbish within the excavation area. Add item 2.32 only for a separately defined area "
        "or operation, so the same clearing is not paid twice."
    ),
}


def resolve_selected_keywords(selected_keywords: Iterable[str]) -> dict[str, Any]:
    """Resolve ordered controlled selections into an auditable composition.

    The input is a sequence of approved keyword labels, in the same order that
    the user selected them.  Within each resulting component category that
    order is retained.  Exactly one base must resolve; 2.24 additions are
    linked to that qualifying base rather than emitted as stand-alone rupee
    items.
    """
    keywords = _normalise_keywords(selected_keywords)
    _validate_controlled_keywords(keywords)
    _reject_duplicate_condition_selections(keywords)
    matches = [component for component in _COMPONENTS if _matches(component, keywords)]
    base_matches = [component for component in matches if component["relationship_type"] == "base"]
    if len(base_matches) != 1:
        raise ValueError("Select exactly one recognised earthwork base item before composing additions.")

    base = _clean_component(base_matches[0])
    deductions = _components_of_type(matches, "deduction", keywords)
    additions = _components_of_type(matches, "addition", keywords)
    conditional_extras = _conditional_extras(keywords, base["item_code"])
    selected_codes = {base["item_code"], *(item["item_code"] for item in deductions),
                      *(item["item_code"] for item in additions),
                      *(item["item_code"] for item in conditional_extras)}
    warnings = [warning for codes, warning in _OVERLAP_RULES.items() if codes.issubset(selected_codes)]
    incompatibilities = _incompatibilities(base["item_code"], deductions, conditional_extras)
    return {
        "base": base,
        "deductions": deductions,
        "additions": additions,
        "conditional_extras": conditional_extras,
        "overlap_warnings": warnings,
        "incompatibilities": incompatibilities,
        "composed_scope": _composed_scope(base, deductions, additions, conditional_extras),
    }


def _normalise_keywords(selected_keywords: Iterable[str]) -> list[str]:
    if isinstance(selected_keywords, (str, bytes)):
        raise TypeError("selected_keywords must be an ordered iterable of keyword labels")
    normalised: list[str] = []
    for keyword in selected_keywords:
        if not isinstance(keyword, str):
            raise TypeError("Each selected keyword must be text")
        value = keyword.strip().casefold()
        if not value:
            raise ValueError("Selected keywords cannot be blank")
        normalised.append(value)
    return normalised


def _validate_controlled_keywords(keywords: list[str]) -> None:
    allowed = {
        keyword.casefold()
        for component in _COMPONENTS
        for keyword in component["keywords"]
    } | set(_CONDITIONAL_KEYWORDS)
    unknown = [keyword for keyword in keywords if keyword not in allowed]
    if unknown:
        raise ValueError(f"Unknown controlled keyword label(s): {', '.join(unknown)}")


def _reject_duplicate_condition_selections(keywords: list[str]) -> None:
    duplicates = [keyword for keyword in _CONDITIONAL_KEYWORDS if keywords.count(keyword) > 1]
    if duplicates:
        raise ValueError(f"Duplicate controlled keyword selection: {', '.join(duplicates)}")


def _matches(component: dict[str, Any], keywords: list[str]) -> bool:
    return all(required.casefold() in keywords for required in component["requires_all"])


def _clean_component(component: dict[str, Any]) -> dict[str, Any]:
    return {key: value for key, value in component.items() if key not in {"keywords", "requires_all"}}


def _components_of_type(matches: list[dict[str, Any]], relationship_type: str, keywords: list[str]) -> list[dict[str, Any]]:
    selected = [item for item in matches if item["relationship_type"] == relationship_type]
    return [
        _clean_component(component)
        for component in sorted(selected, key=lambda item: min(keywords.index(key.casefold()) for key in item["keywords"]))
    ]


def _conditional_extras(keywords: list[str], base_item_code: str) -> list[dict[str, Any]]:
    extras: list[dict[str, Any]] = []
    for keyword in keywords:
        item_code = _CONDITIONAL_KEYWORDS.get(keyword)
        if item_code is None:
            continue
        item = EARTHWORK_COMPOSER_CATALOGUE[item_code]
        extras.append({
            "item_code": item_code,
            "relationship_type": item["relationship_type"],
            "percent": item["percent"],
            "condition": item["condition"],
            "measurement_basis": item["measurement_basis"],
            "qualifying_base_item": base_item_code,
        })
    return extras


def _incompatibilities(
    base_item_code: str, deductions: list[dict[str, Any]], conditional_extras: list[dict[str, Any]]
) -> list[str]:
    messages: list[str] = []
    for deduction in deductions:
        if base_item_code not in _BANKING_BASES:
            messages.append(
                f"Item {deduction['item_code']} is a banking deduction and is not compatible with base "
                f"item {base_item_code}. Select a banking base before applying it."
            )
    selected_extra_codes = {item["item_code"] for item in conditional_extras}
    if {"2.24.1", "2.24.2"}.issubset(selected_extra_codes):
        messages.append(
            "Items 2.24.1 and 2.24.2 require separate qualifying quantities; do not combine their "
            "percentage extras on the same quantity without an engineer's measurement decision."
        )
    return messages


def _composed_scope(
    base: dict[str, Any],
    deductions: list[dict[str, Any]],
    additions: list[dict[str, Any]],
    conditional_extras: list[dict[str, Any]],
) -> str:
    phrases = [f"Base {base['item_code']}: {base['description']}"]
    phrases.extend(f"Deduct {item['item_code']}: {item['description']}" for item in deductions)
    phrases.extend(f"Add {item['item_code']}: {item['description']}" for item in additions)
    phrases.extend(
        f"Conditional {item['item_code']}: {item['percent']}% extra only on the qualifying quantity "
        f"of base {item['qualifying_base_item']}."
        for item in conditional_extras
    )
    return " ".join(phrases)
