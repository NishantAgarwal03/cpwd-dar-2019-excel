"""Source-backed first-principles productivity derivations for earthwork.

This module deliberately reads the verified source table in ``02_Earth_Work``
without editing it.  It exposes a small, auditable model for the later workbook
renderer to use when explaining support-sheet resource coefficients.
"""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
import re
from typing import Any


SHIFT_HOURS = 8.0
INTERPRETATION = "Teaching interpretation, not a published CPWD rule"
SOURCE_EVIDENCE = "CPWD fixed norm / source evidence"


@dataclass(frozen=True)
class _SourceRecord:
    key: str
    code: str
    role: str
    unit: str
    coefficient: float | None
    activity: str
    source_norm: str


def _as_number(value: Any) -> float | None:
    return float(value) if isinstance(value, (int, float)) and not isinstance(value, bool) else None


def _canonical_code(value: Any) -> str:
    """Compare CPWD codes safely whether Excel preserved the leading zeroes."""
    text = str(value).strip()
    if text.isdigit():
        return str(int(text))
    return text


def _source_rows(earth_sheet):
    """Yield structured records from the source table's A:Q layout.

    The defined data table begins at row 109 today, but scanning all rows makes
    the reader robust to a future shift while recognizing only rows that carry
    both a source lookup key (L) and a resource code (F).
    """
    for row in earth_sheet.iter_rows(min_col=1, max_col=17, values_only=True):
        key, code = row[11], row[5]
        if not isinstance(key, str) or not key.strip() or code in (None, ""):
            continue
        yield _SourceRecord(
            key=key.strip(),
            code=_canonical_code(code),
            role=str(row[12] or "").strip(),
            unit=str(row[13] or "").strip(),
            coefficient=_as_number(row[10]),
            activity=str(row[7] or "").strip(),
            source_norm=str(row[8] or "").strip(),
        )


def build_derivation_catalog(earth_sheet) -> dict[str, tuple[_SourceRecord, ...]]:
    """Index source productivity records by their support-resource lookup key.

    A tuple is retained rather than silently selecting a duplicate: callers can
    require a unique mapping and avoid presenting ambiguous evidence as fact.
    """
    catalog: dict[str, list[_SourceRecord]] = defaultdict(list)
    for record in _source_rows(earth_sheet):
        catalog[record.key].append(record)
    return {key: tuple(records) for key, records in catalog.items()}


_GANG_PATTERN = re.compile(
    r"(?:Gang|Crew|Plant)\s*:\s*(?P<size>\d+(?:\.\d+)?)\s+[^×]*?×\s*"
    r"(?P<hours>\d+(?:\.\d+)?)\s*h(?:ours?)?",
    re.IGNORECASE,
)


def _parse_effort(source_norm: str) -> tuple[float, float] | None:
    match = _GANG_PATTERN.search(source_norm)
    if not match:
        return None
    return float(match.group("size")), float(match.group("hours"))


def _is_machine(record: _SourceRecord) -> bool:
    evidence = f"{record.role} {record.source_norm}".lower()
    return "plant:" in evidence or any(word in evidence for word in (
        "roller", "excavator", "loader", "tipper", "machine",
    ))


def derive_resource_norm(resource: dict[str, Any], item_context: dict[str, Any], catalog):
    """Return a derivation only where one source row uniquely verifies it.

    ``resource`` supplies the support-sheet code, coefficient and unit;
    ``item_context`` supplies the already-resolved source key and standard
    batch.  This strict contract keeps matching explicit and prevents inferred
    gang sizes or task hours from being passed off as CPWD evidence.
    """
    source_key = item_context.get("source_key")
    candidates = tuple(catalog.get(source_key, ()))
    resource_code = _canonical_code(resource.get("code", ""))
    candidates = tuple(record for record in candidates if record.code == resource_code)
    if len(candidates) != 1:
        return None

    record = candidates[0]
    supplied_coefficient = _as_number(resource.get("coefficient"))
    if supplied_coefficient is None or record.coefficient is None:
        return None
    if abs(supplied_coefficient - record.coefficient) > 0.001:
        return None

    result = {
        "source_norm": "\n".join(part for part in (record.activity, record.source_norm) if part),
        "role": record.role or str(resource.get("description", "")).strip(),
        "gang_or_machine": None,
        "task_hours": None,
        "shift_hours": int(SHIFT_HOURS),
        "coefficient": supplied_coefficient,
        "unit": str(resource.get("unit") or record.unit),
        "batch_quantity": item_context.get("batch_quantity"),
        "batch_unit": item_context.get("batch_unit"),
        "interpretation_label": INTERPRETATION,
    }
    effort = _parse_effort(record.source_norm)
    if effort is None:
        return result

    gang_or_machine, task_hours = effort
    calculated = task_hours / SHIFT_HOURS if _is_machine(record) else gang_or_machine * task_hours / SHIFT_HOURS
    if abs(calculated - supplied_coefficient) > 0.001:
        # The source text is not a validated derivation; retain no reconstructed
        # hours even though it superficially looks like one.
        return result

    result.update({
        "gang_or_machine": gang_or_machine,
        "task_hours": task_hours,
        "interpretation_label": SOURCE_EVIDENCE,
    })
    if _is_machine(record):
        result["machine_hours"] = task_hours
    return result
