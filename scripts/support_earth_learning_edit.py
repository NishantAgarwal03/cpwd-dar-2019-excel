"""Source-backed first-principles productivity derivations for earthwork.

This module deliberately reads the verified source table in ``02_Earth_Work``
without editing it.  It exposes a small, auditable model for the later workbook
renderer to use when explaining support-sheet resource coefficients.
"""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
import re
import textwrap
from typing import Any
import xml.etree.ElementTree as ET
import zipfile

from openpyxl.comments import Comment
from openpyxl.styles import Alignment, PatternFill


SHIFT_HOURS = 8.0
INTERPRETATION = "Teaching interpretation, not a published CPWD rule"
RECONSTRUCTION = "Engineering teaching reconstruction, not a published CPWD rule"
SOURCE_EVIDENCE = "CPWD fixed norm / source evidence"
FORMULA_LITERAL_REPLACEMENTS = str.maketrans({"—": "-", "→": "->", "›": ">"})


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


def export_ascii_formula_workbook(source_path, output_path, sheet_xml_path="xl/worksheets/sheet8.xml"):
    """Copy a workbook while sanitizing only non-ASCII literal text in formulas.

    The workbook is copied entry-for-entry rather than saved through an XLSX
    writer.  That preserves all non-target worksheet XML (including visible
    labels) byte-for-byte while changing only ``<f>`` nodes on sheet8.
    """
    source_path = str(source_path)
    output_path = str(output_path)
    with zipfile.ZipFile(source_path, "r") as source:
        if sheet_xml_path not in source.namelist():
            raise ValueError(f"Workbook does not contain {sheet_xml_path}")
        # An in-place final export avoids leaving a second workbook artefact.
        # Read every entry before reopening the same path for writing.
        entries = [(entry, source.read(entry.filename)) for entry in source.infolist()]
        if source_path == output_path:
            source.close()
        with zipfile.ZipFile(output_path, "w") as output:
            for entry, payload in entries:
                if entry.filename == sheet_xml_path:
                    root = ET.fromstring(payload)
                    for formula in root.findall(".//{*}f"):
                        formula.text = (formula.text or "").translate(FORMULA_LITERAL_REPLACEMENTS)
                    payload = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                output.writestr(entry, payload)


_ITEM_NUMBER = re.compile(r"^Item\s+(?P<number>[0-9.]+)")
_BATCH = re.compile(r"for\s+(?P<quantity>[0-9.]+)\s+(?P<unit>[A-Za-z]+)", re.IGNORECASE)
_PALE_BLUE = PatternFill("solid", fgColor="DDEBF7")
DEFAULT_SOURCE_MATCH = {
    "2.1.1": "General surface cut (≤30 cm deep)",
    "2.2.1": "Full cycle: rough excavation + banking + roll",
    "2.3.1": "Banking & rolling only (excavation excluded)",
}


def _display_number(value: float | int) -> str:
    return f"{value:.12f}".rstrip("0").rstrip(".")


def estimated_card_lines(text: str, column_width: float) -> int:
    """Conservatively estimate wrapped Excel lines for a detailed I-card.

    Excel's column-width unit is font-dependent.  Using 0.85 characters per
    width unit deliberately errs on the safe side for the workbook's standard
    font, while preserving explicit paragraph breaks as visual lines.
    """
    capacity = max(12, int(float(column_width) * 0.85))
    lines = 0
    for paragraph in str(text or "").splitlines() or [""]:
        lines += max(1, len(textwrap.wrap(paragraph, width=capacity, break_long_words=True,
                                          break_on_hyphens=False)))
    return lines


def _card_height(text: str, column_width: float) -> float:
    return estimated_card_lines(text, column_width) * 15 + 8


def _support_items(sheet):
    """Yield each fixed support-sheet item and its resource rows.

    The support sheet intentionally has a repeated, print-friendly layout.
    Detection relies only on its existing item headings and numeric resource
    quantities, so calculation rows and final rates cannot be mistaken for
    resources.
    """
    starts = []
    for row in range(1, sheet.max_row + 1):
        value = sheet.cell(row, 1).value
        if isinstance(value, str) and _ITEM_NUMBER.match(value):
            starts.append(row)
    for index, start in enumerate(starts):
        end = starts[index + 1] - 1 if index + 1 < len(starts) else sheet.max_row
        header = str(sheet.cell(start, 1).value)
        number = _ITEM_NUMBER.match(header).group("number")
        batch_row = start + 1
        batch_text = str(sheet.cell(batch_row, 1).value or "")
        batch = _BATCH.search(batch_text)
        batch_quantity = float(batch.group("quantity")) if batch else 1.0
        batch_unit = batch.group("unit") if batch else "unit"
        resources = []
        section = ""
        for row in range(start + 1, end + 1):
            code, role, unit, coefficient = (sheet.cell(row, column).value for column in range(1, 5))
            # The support sheets use all-caps section labels (MATERIAL,
            # CARRIAGE and LABOUR).  Retain this nearby context instead of
            # trying to infer a material from its numeric coefficient.
            if (isinstance(code, str) and code.strip().upper() in {"MATERIAL", "CARRIAGE", "LABOUR"}
                    and role in (None, "") and coefficient is None):
                section = code.strip().upper()
                continue
            if code not in (None, "") and isinstance(role, str) and _as_number(coefficient) is not None:
                resources.append((row, {
                    "code": code, "description": role, "unit": unit,
                    "coefficient": coefficient, "section": section,
                }))
        yield {
            "number": number, "start": start, "batch_row": batch_row, "header": header,
            "batch_quantity": batch_quantity, "batch_unit": batch_unit, "resources": resources,
        }


def _record_for_resource(resource, item_number, catalog, source_key_by_item):
    """Select only a source row that independently agrees with the fixed norm."""
    source_key = source_key_by_item.get(item_number)
    if source_key:
        source_keys = [key for key in catalog if key == source_key or source_key in key]
        for candidate_key in source_keys:
            result = derive_resource_norm(resource, {
                "source_key": candidate_key,
                "batch_quantity": None,
                "batch_unit": None,
            }, catalog)
            if result is not None:
                return result

    code = _canonical_code(resource["code"])
    coefficient = _as_number(resource["coefficient"])
    matches = [
        record for records in catalog.values() for record in records
        if record.code == code and record.coefficient is not None
        and coefficient is not None and abs(record.coefficient - coefficient) <= 0.001
    ]
    # Repeated source records are valid only when their evidence agrees; choose
    # no record when the associated method differs, rather than inventing one.
    evidence = {(record.role, record.source_norm) for record in matches}
    if len(evidence) != 1 or not matches:
        return None
    record = matches[0]
    return derive_resource_norm(resource, {
        "source_key": record.key, "batch_quantity": None, "batch_unit": None,
    }, catalog)


def _visible_derivation(resource, derivation, batch_quantity, batch_unit):
    coefficient = _display_number(resource["coefficient"])
    batch = f"{_display_number(batch_quantity)} {batch_unit}"
    role = str(resource["description"])
    unit = str(resource["unit"])
    if _is_material_resource(resource):
        return _material_derivation(resource, batch_quantity, batch_unit, concise=True)
    if derivation and derivation["task_hours"] is not None:
        actor = _display_number(derivation["gang_or_machine"])
        hours = _display_number(derivation["task_hours"])
        return (
            f"CPWD fixed norm: {role}. Calculation: {actor} × {hours} ÷ 8-hour shift = "
            f"{coefficient} {unit} per {batch}. Equivalent productivity: {batch} ÷ {coefficient} {unit}."
        )
    gang, actor = _reconstruction_gang(resource)
    hours = float(resource["coefficient"]) * SHIFT_HOURS / gang
    return (
        f"CPWD fixed norm: {role}. {RECONSTRUCTION}. Calculation: {actor} × {_display_number(hours)} "
        f"task-hours ÷ 8-hour shift = {coefficient} {unit} per {batch}. "
        f"Equivalent productivity: {batch} ÷ {coefficient} {unit}."
    )


def _reconstruction_gang(resource):
    """Choose a transparent pedagogic allocation, never evidence, for each norm."""
    role = str(resource["description"]).lower()
    machine_words = ("roller", "excavator", "loader", "tipper", "breaker", "driller", "machine")
    if any(word in role for word in machine_words):
        return 1, "1 machine"
    if any(word in role for word in ("beldar", "coolie")):
        return 4, "4-person labour gang"
    if str(resource.get("unit", "")).lower() == "day":
        return 1, "1-person specialist gang"
    return 1, "1 resource-handling allocation"


def _is_material_resource(resource):
    """Keep physical inputs out of the labour/machine shift-day model."""
    role = str(resource.get("description", "")).lower()
    section = str(resource.get("section", "")).upper()
    machine_words = ("roller", "excavator", "loader", "tipper", "breaker", "driller", "machine")
    labour_words = ("beldar", "coolie", "carpenter", "mason", "bhishti", "bhisti", "chowkidar", "mate", "helper")
    if any(word in role for word in machine_words):
        return False
    if ((section == "LABOUR" and any(word in role for word in labour_words))
            or (str(resource.get("unit", "")).lower() == "day" and "sundries" not in role)):
        return False
    # MATERIAL and CARRIAGE are physical-consumption lines.  Any remaining
    # non-worker/non-machine resource (including consumables and LS sundries)
    # must not be explained with fictional task-hours.
    return section in {"MATERIAL", "CARRIAGE"} or not any(word in role for word in labour_words)


def _material_physical_basis(resource, batch_quantity, batch_unit):
    """Explain a material coefficient through measurable consumption, not time.

    A support row often preserves only the published coefficient.  This helper
    uses dimensions actually stated in its description when available and says
    plainly when the source does not contain enough geometry to reconstruct the
    original CPWD take-off.
    """
    coefficient = float(resource["coefficient"])
    role = str(resource["description"])
    unit = str(resource.get("unit", ""))
    batch = f"{_display_number(batch_quantity)} {batch_unit}"
    thickness = re.search(r"(\d+(?:\.\d+)?)\s*mm\s*thick", role, re.IGNORECASE)
    if unit.lower() == "cum" and thickness:
        metres = float(thickness.group(1)) / 1000
        coverage = coefficient / metres
        return (
            f"Known geometry: the stated {thickness.group(1)} mm thickness converts the fixed volume to "
            f"{_display_number(coverage)} sqm of material face ({_display_number(coefficient)} cum ÷ "
            f"{_display_number(metres)} m). The source does not state the timber layout, reuse cycle or wastage."
        )
    if unit.lower() == "each":
        per_each = batch_quantity / coefficient if coefficient else 0
        return (
            f"Count/coverage basis: {_display_number(coefficient)} each is allocated to {batch}, equivalent to one "
            f"unit for every {_display_number(per_each)} {batch_unit}. The source does not state spacing, reuse or wastage."
        )
    if unit.lower() == "cum":
        return (
            f"Volume-consumption basis: {_display_number(coefficient)} cum is the fixed material volume for {batch}. "
            "The source does not provide section dimensions, installed length, coverage or wastage from which to rebuild that volume."
        )
    return (
        f"Physical-consumption basis: the fixed allowance is {_display_number(coefficient)} {unit} for {batch}. "
        "The source does not provide the count, geometry, coverage or wastage detail needed to rebuild the allowance."
    )


def _material_derivation(resource, batch_quantity, batch_unit, concise=False):
    coefficient = _display_number(resource["coefficient"])
    batch = f"{_display_number(batch_quantity)} {batch_unit}"
    body = _material_physical_basis(resource, batch_quantity, batch_unit)
    if concise:
        return (
            f"CPWD fixed material coefficient: {resource['description']}. Material-consumption basis: "
            f"{coefficient} {resource['unit']} per {batch}. {body} {RECONSTRUCTION}"
        )
    return (
        "CPWD fixed material coefficient / source evidence\n"
        "The support sheet fixes the stated material coefficient; no unique source-row geometry was available.\n\n"
        "Material-consumption basis\n"
        f"{body}\n\n"
        "Calculation\n"
        f"Fixed consumption = {coefficient} {resource['unit']} per {batch}. This is a quantity take-off/allowance, not a labour productivity calculation.\n\n"
        "Engineering interpretation for learning\n"
        f"{RECONSTRUCTION}. The physical explanation makes the published quantity auditable but does not claim an unpublished CPWD layout, coverage or wastage rule.\n\n"
        "Boundary conditions\n"
        "Use only for the stated material, section, output batch and method. Check reuse, cutting loss, spacing, member size and carriage separately when the item specifies them.\n\n"
        "When the norm changes\n"
        "Revise the quantity only when the material specification, section, coverage, count, reuse, wastage, carriage or CPWD item scope changes."
    )


def _learning_note(resource, derivation, batch_quantity, batch_unit):
    coefficient = _display_number(resource["coefficient"])
    batch = f"{_display_number(batch_quantity)} {batch_unit}"
    if _is_material_resource(resource):
        return _material_derivation(resource, batch_quantity, batch_unit)
    source = derivation["source_norm"] if derivation else "No unique source-row match was available for this support-sheet coefficient."
    if derivation and derivation["task_hours"] is not None:
        actor = _display_number(derivation["gang_or_machine"])
        hours = _display_number(derivation["task_hours"])
        calculation = f"{actor} × {hours} task-hours ÷ 8-hour shift = {coefficient} {resource['unit']} per {batch}."
        basis = f"{actor} source-backed gang/machine allocation; each assigned task-hour is converted to an 8-hour shift-day."
        interpretation = (
            f"The resource allocation represents its assigned activity within the work method. "
            f"The inverse gives {_display_number(batch_quantity / float(resource['coefficient']))} {batch_unit} per {resource['unit']}."
        )
    else:
        gang, actor = _reconstruction_gang(resource)
        hours = float(resource["coefficient"]) * SHIFT_HOURS / gang
        calculation = (
            f"{actor} × {_display_number(hours)} task-hours ÷ 8-hour shift = "
            f"{coefficient} {resource['unit']} per {batch}."
        )
        basis = f"{actor}; task-hours are selected solely so the teaching reconstruction reproduces the fixed coefficient exactly."
        interpretation = (
            f"{RECONSTRUCTION}. This allocation is chosen only to make the published coefficient auditable; "
            "it is not source evidence of an actual CPWD gang or task-hour norm."
        )
    return (
        f"CPWD fixed norm / source evidence\n{source}\n\n"
        f"Role, activity and gang/machine basis\n{resource['description']}: {basis}\n\n"
        f"Calculation\n{calculation}\n\n"
        f"Engineering interpretation for learning\n{interpretation}\n\n"
        f"Boundary conditions\nUse only for the stated operation, output batch, lead/lift, material and method. The teaching shift is fixed at 8 hours.\n\n"
        f"When the norm changes\nRevise the resource norm only when the CPWD item scope, site condition, lead/lift, material, method or specified output batch changes."
    )


def apply_markup_outline(sheet):
    """Group only the repeat compounding calculation beneath every item.

    The group begins with Water charges and ends at Total cost.  The actual
    rate, rounded Say rate, and project rate stay at outline level zero, so a
    student sees the answer while retaining one-click access to the working.
    """
    for item in _support_items(sheet):
        start = end = None
        for row in range(item["start"], sheet.max_row + 1):
            text = str(sheet.cell(row, 1).value or "")
            if row > item["start"] and text.startswith("Item "):
                break
            if text.startswith("Add: Water charges"):
                start = row
            if start is not None and text.startswith("Total cost for"):
                end = row
                break
        if start is None or end is None:
            continue
        for row in range(start, end + 1):
            dimension = sheet.row_dimensions[row]
            dimension.outlineLevel = 1
            dimension.hidden = False
        # Explicitly safeguard the answer rows after the grouped calculation.
        for row in range(end + 1, min(end + 4, sheet.max_row + 1)):
            sheet.row_dimensions[row].outlineLevel = 0
    sheet.sheet_properties.outlinePr.summaryBelow = True


def apply_first_principles_learning(workbook, source_key_by_item: dict[str, str] | None = None):
    """Add the approved three-layer explanation without touching A:H values/styles.

    ``source_key_by_item`` is an explicit mapping for cases such as item 2.1.1;
    all other rows fall back only to unambiguous matching fixed source evidence.
    """
    source_key_by_item = {**DEFAULT_SOURCE_MATCH, **(source_key_by_item or {})}
    support = workbook["02_support_earth_work"]
    catalog = build_derivation_catalog(workbook["02_Earth_Work"])
    support.column_dimensions["I"].width = 110
    for item in _support_items(support):
        records = []
        for row, resource in item["resources"]:
            derivation = _record_for_resource(resource, item["number"], catalog, source_key_by_item)
            # The worksheet is the primary learning surface: retain the full
            # source/interpretation card in I, and reserve the D Note for a
            # quick arithmetic check beside the fixed number.
            card = _learning_note(resource, derivation, item["batch_quantity"], item["batch_unit"])
            support.cell(row, 9).value = card
            support.cell(row, 9).fill = _PALE_BLUE
            support.cell(row, 9).alignment = Alignment(wrap_text=True, vertical="top")
            support.row_dimensions[row].height = _card_height(card, support.column_dimensions["I"].width)
            support.cell(row, 4).comment = Comment(
                _visible_derivation(resource, derivation, item["batch_quantity"], item["batch_unit"]), "CPWD learning guide"
            )
            records.append(str(resource["description"]))
        if records:
            method = (
                f"Work method: {item['header'].split('|', 1)[-1].strip()}. Standard batch: "
                f"{_display_number(item['batch_quantity'])} {item['batch_unit']}; teaching basis: fixed 8-hour shift. "
                "Method constraints remain those stated in the CPWD item. "
            )
            if len(records) > 1:
                method += "Gang system: " + ", ".join(records) + " work as coordinated resources; each line records its own share of the same output."
            support.cell(item["start"], 9).value = method
            support.cell(item["start"], 9).fill = _PALE_BLUE
            support.cell(item["start"], 9).alignment = Alignment(wrap_text=True, vertical="top")
    apply_markup_outline(support)
    return workbook
