"""CPWD Published Productivity Registry.

Maps CPWD resource coefficients to transparent productivity metrics:
- Labour: Beldar/Coolie/Bhisti -> published coefficient -> derived sqm/day or cum/day
- Machinery: Excavator/Roller -> published coefficient -> derived cum/hour or cum/shift
- Materials: Physical consumption and sundries

Provides export to INI/text format for human inspection and auditing.
"""

from __future__ import annotations

import configparser
from dataclasses import dataclass
import json
from pathlib import Path
import re
from typing import Any


SHIFT_HOURS = 8.0


@dataclass(frozen=True)
class ProductivityEntry:
    subhead: str
    item_code: str
    activity: str
    resource_code: str
    resource_name: str
    resource_type: str  # "Labour" | "Machinery" | "Material"
    coefficient: float
    unit: str
    batch_quantity: float
    batch_unit: str
    daily_productivity: float | None
    hourly_productivity: float | None
    metric_unit: str
    effort_hours: float | None
    derivation: str
    source_citation: str

    def to_dict(self) -> dict[str, Any]:
        return {
            "subhead": self.subhead,
            "item_code": self.item_code,
            "activity": self.activity,
            "resource_code": self.resource_code,
            "resource_name": self.resource_name,
            "resource_type": self.resource_type,
            "coefficient": self.coefficient,
            "unit": self.unit,
            "batch_quantity": self.batch_quantity,
            "batch_unit": self.batch_unit,
            "daily_productivity": self.daily_productivity,
            "hourly_productivity": self.hourly_productivity,
            "metric_unit": self.metric_unit,
            "effort_hours": self.effort_hours,
            "derivation": self.derivation,
            "source_citation": self.source_citation,
        }


# Published CPWD benchmark entries for Earth Work (Pilot)
_EARTHWORK_BENCHMARKS: tuple[dict[str, Any], ...] = (
    # 2.1.1 Surface excavation ordinary soil (100 sqm)
    {
        "subhead": "02_Earth_Work",
        "item_code": "2.1.1",
        "activity": "Surface excavation <= 30cm depth in ordinary soil",
        "resource_code": "0114",
        "resource_name": "Beldar",
        "resource_type": "Labour",
        "coefficient": 6.80,
        "unit": "day",
        "batch_quantity": 100.0,
        "batch_unit": "sqm",
        "gang_size": 4.0,
        "task_hours": 13.6,
        "source_citation": "CPWD DAR 2019 Item 2.1.1",
    },
    {
        "subhead": "02_Earth_Work",
        "item_code": "2.1.1",
        "activity": "Surface excavation <= 30cm depth in ordinary soil",
        "resource_code": "0115",
        "resource_name": "Coolie",
        "resource_type": "Labour",
        "coefficient": 5.60,
        "unit": "day",
        "batch_quantity": 100.0,
        "batch_unit": "sqm",
        "gang_size": 4.0,
        "task_hours": 11.2,
        "source_citation": "CPWD DAR 2019 Item 2.1.1",
    },
    # 2.2 / 2.2.1 Rough excavation and banking (10 cum)
    {
        "subhead": "02_Earth_Work",
        "item_code": "2.2",
        "activity": "Rough excavation & banking in layers",
        "resource_code": "0114",
        "resource_name": "Beldar",
        "resource_type": "Labour",
        "coefficient": 5.90,
        "unit": "day",
        "batch_quantity": 10.0,
        "batch_unit": "cum",
        "gang_size": 4.0,
        "task_hours": 11.8,
        "source_citation": "CPWD DAR 2019 Item 2.2.1",
    },
    {
        "subhead": "02_Earth_Work",
        "item_code": "2.2",
        "activity": "Rough excavation & banking in layers",
        "resource_code": "0115",
        "resource_name": "Coolie",
        "resource_type": "Labour",
        "coefficient": 3.60,
        "unit": "day",
        "batch_quantity": 10.0,
        "batch_unit": "cum",
        "gang_size": 3.0,
        "task_hours": 9.6,
        "source_citation": "CPWD DAR 2019 Item 2.2.1",
    },
    {
        "subhead": "02_Earth_Work",
        "item_code": "2.2",
        "activity": "Rough excavation & banking in layers",
        "resource_code": "0101",
        "resource_name": "Bhisti",
        "resource_type": "Labour",
        "coefficient": 0.40,
        "unit": "day",
        "batch_quantity": 10.0,
        "batch_unit": "cum",
        "gang_size": 1.0,
        "task_hours": 3.2,
        "source_citation": "CPWD DAR 2019 Item 2.2.1",
    },
    {
        "subhead": "02_Earth_Work",
        "item_code": "2.2",
        "activity": "Rough excavation & banking in layers",
        "resource_code": "0003",
        "resource_name": "Hire charges of Diesel Road Roller - 8 to 10 tonne",
        "resource_type": "Machinery",
        "coefficient": 0.008,
        "unit": "day",
        "batch_quantity": 10.0,
        "batch_unit": "cum",
        "machine_hours": 0.064,
        "source_citation": "CPWD DAR 2019 Item 2.2.1",
    },
    {
        "subhead": "02_Earth_Work",
        "item_code": "2.2",
        "activity": "Rough excavation & banking in layers",
        "resource_code": "0113",
        "resource_name": "Chowkidar",
        "resource_type": "Labour",
        "coefficient": 0.008,
        "unit": "day",
        "batch_quantity": 10.0,
        "batch_unit": "cum",
        "gang_size": 1.0,
        "task_hours": 0.064,
        "source_citation": "CPWD DAR 2019 Item 2.2.1",
    },
    {
        "subhead": "02_Earth_Work",
        "item_code": "2.2",
        "activity": "Rough excavation & banking in layers",
        "resource_code": "EQUIP-0.5T",
        "resource_name": "½-tonne hand roller / wooden or steel rammers (durmats)",
        "resource_type": "Machinery",
        "coefficient": 1.100,
        "unit": "day",
        "batch_quantity": 10.0,
        "batch_unit": "cum",
        "machine_hours": 8.80,
        "source_citation": "CPWD DAR 2019 Item 2.2 Specification & Beldar Allocation",
    },
    # 2.3.1 Banking excavated earth in layers (10 cum)
    {
        "subhead": "02_Earth_Work",
        "item_code": "2.3.1",
        "activity": "Banking excavated earth in layers not exceeding 20 cm",
        "resource_code": "0114",
        "resource_name": "Beldar",
        "resource_type": "Labour",
        "coefficient": 2.20,
        "unit": "day",
        "batch_quantity": 10.0,
        "batch_unit": "cum",
        "gang_size": 2.0,
        "task_hours": 8.8,
        "source_citation": "CPWD DAR 2019 Item 2.3.1",
    },
    {
        "subhead": "02_Earth_Work",
        "item_code": "2.3.1",
        "activity": "Banking excavated earth in layers not exceeding 20 cm",
        "resource_code": "0115",
        "resource_name": "Coolie",
        "resource_type": "Labour",
        "coefficient": 3.60,
        "unit": "day",
        "batch_quantity": 10.0,
        "batch_unit": "cum",
        "gang_size": 3.0,
        "task_hours": 9.6,
        "source_citation": "CPWD DAR 2019 Item 2.3.1",
    },
    {
        "subhead": "02_Earth_Work",
        "item_code": "2.3.1",
        "activity": "Banking excavated earth in layers not exceeding 20 cm",
        "resource_code": "0101",
        "resource_name": "Bhisti",
        "resource_type": "Labour",
        "coefficient": 0.40,
        "unit": "day",
        "batch_quantity": 10.0,
        "batch_unit": "cum",
        "gang_size": 1.0,
        "task_hours": 3.2,
        "source_citation": "CPWD DAR 2019 Item 2.3.1",
    },
    {
        "subhead": "02_Earth_Work",
        "item_code": "2.3.1",
        "activity": "Banking excavated earth in layers not exceeding 20 cm",
        "resource_code": "0003",
        "resource_name": "Hire charges of Diesel Road Roller - 8 to 10 tonne",
        "resource_type": "Machinery",
        "coefficient": 0.008,
        "unit": "day",
        "batch_quantity": 10.0,
        "batch_unit": "cum",
        "machine_hours": 0.064,
        "source_citation": "CPWD DAR 2019 Item 2.3.1",
    },
    {
        "subhead": "02_Earth_Work",
        "item_code": "2.3.1",
        "activity": "Banking excavated earth in layers not exceeding 20 cm",
        "resource_code": "0113",
        "resource_name": "Chowkidar",
        "resource_type": "Labour",
        "coefficient": 0.008,
        "unit": "day",
        "batch_quantity": 10.0,
        "batch_unit": "cum",
        "gang_size": 1.0,
        "task_hours": 0.064,
        "source_citation": "CPWD DAR 2019 Item 2.3.1",
    },
    {
        "subhead": "02_Earth_Work",
        "item_code": "2.3.1",
        "activity": "Banking excavated earth in layers not exceeding 20 cm",
        "resource_code": "EQUIP-0.5T",
        "resource_name": "½-tonne hand roller / wooden or steel rammers (durmats)",
        "resource_type": "Machinery",
        "coefficient": 1.100,
        "unit": "day",
        "batch_quantity": 10.0,
        "batch_unit": "cum",
        "machine_hours": 8.80,
        "source_citation": "CPWD DAR 2019 Item 2.3.1 Specification & Beldar Allocation",
    },
    # 2.4 Deduct for not rolling (10 cum)
    {
        "subhead": "02_Earth_Work",
        "item_code": "2.4",
        "activity": "Deduct for not rolling with 8-tonne power roller",
        "resource_code": "0003",
        "resource_name": "Hire charges of Diesel Road Roller - 8 to 10 tonne",
        "resource_type": "Machinery",
        "coefficient": -0.008,
        "unit": "day",
        "batch_quantity": 10.0,
        "batch_unit": "cum",
        "machine_hours": -0.064,
        "source_citation": "CPWD DAR 2019 Item 2.4",
    },
    {
        "subhead": "02_Earth_Work",
        "item_code": "2.4",
        "activity": "Deduct for not rolling with 8-tonne power roller",
        "resource_code": "0113",
        "resource_name": "Chowkidar",
        "resource_type": "Labour",
        "coefficient": -0.008,
        "unit": "day",
        "batch_quantity": 10.0,
        "batch_unit": "cum",
        "gang_size": -1.0,
        "task_hours": -0.064,
        "source_citation": "CPWD DAR 2019 Item 2.4",
    },
    # 2.5 Deduct for not watering (10 cum)
    {
        "subhead": "02_Earth_Work",
        "item_code": "2.5",
        "activity": "Deduct for not watering excavated earth for banking",
        "resource_code": "0101",
        "resource_name": "Bhisti",
        "resource_type": "Labour",
        "coefficient": -0.40,
        "unit": "day",
        "batch_quantity": 10.0,
        "batch_unit": "cum",
        "gang_size": -1.0,
        "task_hours": -3.2,
        "source_citation": "CPWD DAR 2019 Item 2.5",
    },
    # 2.6.1 Hydraulic Excavator (10 cum)
    {
        "subhead": "02_Earth_Work",
        "item_code": "2.6.1",
        "activity": "Hydraulic excavation in foundation trenches",
        "resource_code": "0020",
        "resource_name": "Hydraulic Excavator (0.9 cum) with driver and fuel",
        "resource_type": "Machinery",
        "coefficient": 0.041,
        "unit": "day",
        "batch_quantity": 10.0,
        "batch_unit": "cum",
        "machine_hours": 0.328,
        "source_citation": "CPWD DAR 2019 Item 2.6.1",
    },
    {
        "subhead": "02_Earth_Work",
        "item_code": "2.6.1",
        "activity": "Hydraulic excavation in foundation trenches",
        "resource_code": "0018",
        "resource_name": "Loader",
        "resource_type": "Machinery",
        "coefficient": 0.041,
        "unit": "day",
        "batch_quantity": 10.0,
        "batch_unit": "cum",
        "machine_hours": 0.328,
        "source_citation": "CPWD DAR 2019 Item 2.6.1",
    },
    # 2.32 Clearing grass (100 sqm)
    {
        "subhead": "02_Earth_Work",
        "item_code": "2.32",
        "activity": "Clearing grass and removal of rubbish",
        "resource_code": "0114",
        "resource_name": "Beldar",
        "resource_type": "Labour",
        "coefficient": 0.80,
        "unit": "day",
        "batch_quantity": 100.0,
        "batch_unit": "sqm",
        "gang_size": 2.0,
        "task_hours": 3.2,
        "source_citation": "CPWD DAR 2019 Item 2.32",
    },
    {
        "subhead": "02_Earth_Work",
        "item_code": "2.32",
        "activity": "Clearing grass and removal of rubbish",
        "resource_code": "0115",
        "resource_name": "Coolie",
        "resource_type": "Labour",
        "coefficient": 0.67,
        "unit": "day",
        "batch_quantity": 100.0,
        "batch_unit": "sqm",
        "gang_size": 2.0,
        "task_hours": 2.68,
        "source_citation": "CPWD DAR 2019 Item 2.32",
    },
    # 2.25 Filling available excavated earth in plinth / trenches (10 cum)
    {
        "subhead": "02_Earth_Work",
        "item_code": "2.25",
        "activity": "Filling in plinth and trenches in 20 cm layers",
        "resource_code": "EQUIP-DURMAT",
        "resource_name": "Wooden or steel rammers (durmats) for layer compaction",
        "resource_type": "Machinery",
        "coefficient": 0.800,
        "unit": "day",
        "batch_quantity": 10.0,
        "batch_unit": "cum",
        "machine_hours": 6.40,
        "source_citation": "CPWD DAR 2019 Item 2.25 Specification & Beldar Allocation",
    },
    {
        "subhead": "02_Earth_Work",
        "item_code": "2.25",
        "activity": "Filling in plinth and trenches in 20 cm layers",
        "resource_code": "0114",
        "resource_name": "Beldar",
        "resource_type": "Labour",
        "coefficient": 0.80,
        "unit": "day",
        "batch_quantity": 10.0,
        "batch_unit": "cum",
        "gang_size": 1.0,
        "task_hours": 6.4,
        "source_citation": "CPWD DAR 2019 Item 2.25",
    },
    {
        "subhead": "02_Earth_Work",
        "item_code": "2.25",
        "activity": "Filling in plinth and trenches in 20 cm layers",
        "resource_code": "0115",
        "resource_name": "Coolie",
        "resource_type": "Labour",
        "coefficient": 0.85,
        "unit": "day",
        "batch_quantity": 10.0,
        "batch_unit": "cum",
        "gang_size": 1.0,
        "task_hours": 6.8,
        "source_citation": "CPWD DAR 2019 Item 2.25",
    },
    {
        "subhead": "02_Earth_Work",
        "item_code": "2.25",
        "activity": "Filling in plinth and trenches in 20 cm layers",
        "resource_code": "0101",
        "resource_name": "Bhisti",
        "resource_type": "Labour",
        "coefficient": 0.20,
        "unit": "day",
        "batch_quantity": 10.0,
        "batch_unit": "cum",
        "gang_size": 1.0,
        "task_hours": 1.6,
        "source_citation": "CPWD DAR 2019 Item 2.25",
    },
)


class CPWDProductivityRegistry:
    """Central registry of published CPWD productivity norms and derivations."""

    def __init__(self) -> None:
        self._entries: dict[tuple[str, str, str], ProductivityEntry] = {}
        self._load_earthwork_benchmarks()

    def _load_earthwork_benchmarks(self) -> None:
        for bench in _EARTHWORK_BENCHMARKS:
            entry = self._build_entry(bench)
            key = (entry.subhead, entry.item_code, entry.resource_code)
            self._entries[key] = entry

    @staticmethod
    def _build_entry(raw: dict[str, Any]) -> ProductivityEntry:
        subhead = raw["subhead"]
        item_code = raw["item_code"]
        activity = raw["activity"]
        resource_code = raw["resource_code"]
        resource_name = raw["resource_name"]
        resource_type = raw["resource_type"]
        coefficient = float(raw["coefficient"])
        unit = raw["unit"]
        batch_quantity = float(raw["batch_quantity"])
        batch_unit = raw["batch_unit"]
        source_citation = raw["source_citation"]

        abs_coeff = abs(coefficient)
        daily_productivity: float | None = None
        hourly_productivity: float | None = None
        effort_hours: float | None = None
        metric_unit = f"{batch_unit}/day"

        if resource_type == "Labour":
            if abs_coeff > 0:
                daily_productivity = round(batch_quantity / abs_coeff, 2)
                hourly_productivity = round(daily_productivity / SHIFT_HOURS, 2)
                metric_unit = f"{batch_unit}/worker-day"
                if raw.get("task_hours") is not None and raw.get("gang_size") is not None:
                    effort_hours = float(raw["task_hours"]) * float(raw["gang_size"])
                elif raw.get("task_hours") is not None:
                    effort_hours = float(raw["task_hours"])
                else:
                    effort_hours = round(abs_coeff * SHIFT_HOURS, 2)
            is_deduct = coefficient < 0
            prefix = "Deduction: " if is_deduct else ""
            derivation = (
                f"{prefix}{batch_quantity} {batch_unit} / {abs_coeff:.4f} worker-days = "
                f"{daily_productivity} {metric_unit} ({hourly_productivity} {batch_unit}/hr across {int(SHIFT_HOURS)}h shift)"
            )

        elif resource_type == "Machinery":
            if abs_coeff > 0:
                m_hr = raw.get("machine_hours")
                if m_hr is not None:
                    machine_hours = float(m_hr)
                else:
                    machine_hours = abs_coeff * SHIFT_HOURS
                abs_machine_hours = abs(machine_hours)
                effort_hours = abs_machine_hours
                if abs_machine_hours > 0:
                    hourly_productivity = round(batch_quantity / abs_machine_hours, 2)
                daily_productivity = round(hourly_productivity * SHIFT_HOURS, 2) if hourly_productivity else None
                metric_unit = f"{batch_unit}/machine-hour"
            is_deduct = coefficient < 0
            prefix = "Deduction: " if is_deduct else ""
            derivation = (
                f"{prefix}{batch_quantity} {batch_unit} / {abs_coeff:.4f} machine-days "
                f"({abs(effort_hours or 0):.3f} h) = {hourly_productivity} {metric_unit} "
                f"({daily_productivity} {batch_unit}/shift)"
            )

        else:  # Material
            derivation = f"Direct consumption: {coefficient} {unit} per {batch_quantity} {batch_unit}"
            metric_unit = f"{unit}/{batch_unit}"

        return ProductivityEntry(
            subhead=subhead,
            item_code=item_code,
            activity=activity,
            resource_code=resource_code,
            resource_name=resource_name,
            resource_type=resource_type,
            coefficient=coefficient,
            unit=unit,
            batch_quantity=batch_quantity,
            batch_unit=batch_unit,
            daily_productivity=daily_productivity,
            hourly_productivity=hourly_productivity,
            metric_unit=metric_unit,
            effort_hours=effort_hours,
            derivation=derivation,
            source_citation=source_citation,
        )

    def load_from_json(self, json_path: str | Path) -> int:
        """Load external productivity records from labour_productivity.json."""
        path = Path(json_path)
        if not path.exists():
            return 0
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        added = 0
        for item in data:
            subhead = item.get("subhead") or ""
            item_no = item.get("item_no") or ""
            code = str(item.get("code") or "").strip()
            coeff = item.get("coefficient")
            if not subhead or not item_no or not code or coeff is None:
                continue
            key = (subhead, item_no, code)
            if key in self._entries:
                continue
            try:
                coeff_f = float(coeff)
            except (ValueError, TypeError):
                continue
            basis_str = str(item.get("basis") or "1 unit").strip()
            parts = basis_str.split()
            try:
                batch_qty = float(parts[0]) if parts else 1.0
            except ValueError:
                batch_qty = 1.0
            batch_unit = parts[1] if len(parts) > 1 else "unit"

            entry = self._build_entry({
                "subhead": subhead,
                "item_code": item_no,
                "activity": item.get("item_desc") or "",
                "resource_code": code,
                "resource_name": item.get("description") or "",
                "resource_type": item.get("type") or "Labour",
                "coefficient": coeff_f,
                "unit": item.get("unit") or "day",
                "batch_quantity": batch_qty,
                "batch_unit": batch_unit,
                "source_citation": f"CPWD DAR 2019 {subhead} Item {item_no}",
            })
            self._entries[key] = entry
            added += 1
        return added

    def lookup(self, subhead: str, item_code: str, resource_code: str) -> ProductivityEntry | None:
        """Look up a productivity entry by subhead, item_code, and resource_code."""
        entry = self._entries.get((subhead, item_code, resource_code))
        if entry is not None:
            return entry
        norm_code = resource_code.lstrip("0") or "0"
        for (sh, it, rc), val in self._entries.items():
            if sh == subhead and it == item_code and (rc == resource_code or rc.lstrip("0") == norm_code):
                return val
        return None

    def all_entries(self) -> list[ProductivityEntry]:
        return list(self._entries.values())

    def export_to_ini(self, output_path: str | Path) -> Path:
        """Export the registry to a clean, human-readable INI file."""
        target = Path(output_path)
        target.parent.mkdir(parents=True, exist_ok=True)
        config = configparser.ConfigParser(interpolation=None)

        used_sections: set[str] = set()
        for entry in self._entries.values():
            safe_name = re.sub(r"[^a-zA-Z0-9_-]", "_", entry.resource_name[:20].strip())
            base_sec = f"{entry.subhead}.{entry.item_code}.{entry.resource_code}_{safe_name}"
            sec = base_sec
            counter = 2
            while sec in used_sections:
                sec = f"{base_sec}_{counter}"
                counter += 1
            used_sections.add(sec)

            config[sec] = {
                "subhead": entry.subhead,
                "item_code": entry.item_code,
                "activity": entry.activity.replace("\n", " ").replace("\r", " ").strip(),
                "resource_code": entry.resource_code,
                "resource_name": entry.resource_name.replace("\n", " ").replace("\r", " ").strip(),
                "resource_type": entry.resource_type,
                "coefficient": f"{entry.coefficient:.4f}",
                "unit": entry.unit,
                "batch_quantity": str(entry.batch_quantity),
                "batch_unit": entry.batch_unit,
                "daily_productivity": f"{entry.daily_productivity} {entry.metric_unit}" if entry.daily_productivity is not None else "N/A",
                "hourly_productivity": f"{entry.hourly_productivity} {entry.batch_unit}/hr" if entry.hourly_productivity is not None else "N/A",
                "effort_hours": f"{entry.effort_hours:.2f} h" if entry.effort_hours is not None else "N/A",
                "derivation": entry.derivation.replace("\n", " ").replace("\r", " ").strip(),
                "source_citation": entry.source_citation.replace("\n", " ").replace("\r", " ").strip(),
            }

        with open(target, "w", encoding="utf-8") as f:
            config.write(f)
        return target

    def export_to_txt(self, output_path: str | Path) -> Path:
        """Export the registry to a cleanly formatted text report."""
        target = Path(output_path)
        target.parent.mkdir(parents=True, exist_ok=True)
        lines = [
            "=" * 90,
            "CPWD DAR 2019 - PUBLISHED PRODUCTIVITY REGISTRY",
            "Transparent derivations: resource x work activity -> published coefficient -> productivity",
            "=" * 90,
            "",
        ]
        current_subhead = ""
        for entry in sorted(self._entries.values(), key=lambda e: (e.subhead, e.item_code, e.resource_code)):
            if entry.subhead != current_subhead:
                current_subhead = entry.subhead
                lines.append(f"\n# Sub-Head: {current_subhead}")
                lines.append("-" * 80)
            lines.append(f"[{entry.item_code}] {entry.activity}")
            lines.append(f"  Resource: [{entry.resource_code}] {entry.resource_name} ({entry.resource_type})")
            lines.append(f"  Published Coefficient: {entry.coefficient:.4f} {entry.unit} per {entry.batch_quantity} {entry.batch_unit}")
            if entry.daily_productivity is not None:
                lines.append(f"  Derived Productivity: {entry.daily_productivity} {entry.metric_unit} | {entry.hourly_productivity} {entry.batch_unit}/hr")
            lines.append(f"  Derivation: {entry.derivation}")
            lines.append(f"  Evidence: {entry.source_citation}")
            lines.append("")

        with open(target, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
        return target


# Global shared singleton
PRODUCTIVITY_REGISTRY = CPWDProductivityRegistry()
