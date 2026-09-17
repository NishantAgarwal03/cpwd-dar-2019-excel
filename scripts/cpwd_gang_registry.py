"""CPWD Gang Composition Registry.

Defines the labour, machinery, equipment and material mix for CPWD work items:
- Activity -> labour + machine + equipment mix -> output basis
- Deductions: subtraction of omitted equipment/labour from base gang
- Additions: combination of distinct operations
- Output export to INI and TXT for human audit and inspection.
"""

from __future__ import annotations

import configparser
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


@dataclass
class GangResource:
    code: str
    name: str
    category: str  # "Machinery" | "Labour" | "Material" | "Sundries"
    coefficient: float  # quantity per batch
    unit: str
    rate: float
    gang_count: float | None = None
    task_hours: float | None = None

    @property
    def amount(self) -> float:
        return round(self.coefficient * self.rate, 2)

    def copy(self, coefficient: float | None = None) -> GangResource:
        return GangResource(
            code=self.code,
            name=self.name,
            category=self.category,
            coefficient=self.coefficient if coefficient is None else coefficient,
            unit=self.unit,
            rate=self.rate,
            gang_count=self.gang_count,
            task_hours=self.task_hours,
        )


@dataclass
class ItemGang:
    item_code: str
    description: str
    batch_quantity: float
    batch_unit: str
    dsr_rate: float | None
    machinery: list[GangResource] = field(default_factory=list)
    labour: list[GangResource] = field(default_factory=list)
    materials: list[GangResource] = field(default_factory=list)
    sundries: list[GangResource] = field(default_factory=list)
    notes: str = ""

    @property
    def direct_cost_w(self) -> float:
        return sum(
            res.amount
            for res in [*self.machinery, *self.labour, *self.materials, *self.sundries]
        )


@dataclass
class ComposedGang:
    base_item: ItemGang
    net_machinery: list[GangResource]
    net_labour: list[GangResource]
    net_materials: list[GangResource]
    net_sundries: list[GangResource]
    deducted_resources: list[GangResource]
    added_resources: list[GangResource]
    batch_quantity: float
    batch_unit: str

    @property
    def all_active_resources(self) -> list[GangResource]:
        return [
            res
            for res in [*self.net_machinery, *self.net_labour, *self.net_materials, *self.net_sundries]
            if res.coefficient > 0.0001
        ]

    @property
    def direct_cost_w(self) -> float:
        return round(sum(res.amount for res in self.all_active_resources), 2)


# Ground truth CPWD items for Earth Work Pilot
_EARTHWORK_ITEMS: tuple[dict[str, Any], ...] = (
    # 2.1.1 Surface excavation ordinary soil (100 sqm)
    {
        "item_code": "2.1.1",
        "description": "Surface dressing of ground / surface excavation <= 30cm in ordinary soil",
        "batch_quantity": 100.0,
        "batch_unit": "sqm",
        "dsr_rate": 107.00,
        "labour": [
            {"code": "0114", "name": "Beldar", "unit": "day", "coefficient": 6.80, "rate": 558.00, "gang_count": 4.0, "task_hours": 13.6},
            {"code": "0115", "name": "Coolie", "unit": "day", "coefficient": 5.60, "rate": 558.00, "gang_count": 4.0, "task_hours": 11.2},
        ],
    },
    # 2.2 Rough excavation and banking in layers (10 cum)
    {
        "item_code": "2.2",
        "description": "Earth work in rough excavation, banking excavated earth in layers not exceeding 20cm in depth, breaking clods, watering, rolling each layer with ½ tonne roller or wooden or steel rammers, and rolling every 3rd and top-most layer with power roller of minimum 8 tonnes and dressing up in embankments for roads, flood banks, marginal banks and guide banks or filling up ground depressions, lead upto 50 m and lift upto 1.5 m : All kinds of soil",
        "batch_quantity": 10.0,
        "batch_unit": "cum",
        "dsr_rate": 746.80,
        "machinery": [
            {"code": "0003", "name": "Hire charges of Diesel Road Roller - 8 to 10 tonne", "unit": "day", "coefficient": 0.008, "rate": 3000.00, "task_hours": 0.064},
            {"code": "EQUIP-0.5T", "name": "½-tonne hand roller / wooden or steel rammers (durmats)", "unit": "day", "coefficient": 1.100, "rate": 0.00, "task_hours": 8.80},
        ],
        "labour": [
            {"code": "0114", "name": "Beldar", "unit": "day", "coefficient": 5.90, "rate": 558.00, "gang_count": 4.0, "task_hours": 11.8},
            {"code": "0115", "name": "Coolie", "unit": "day", "coefficient": 3.60, "rate": 558.00, "gang_count": 3.0, "task_hours": 9.6},
            {"code": "0101", "name": "Bhisti", "unit": "day", "coefficient": 0.40, "rate": 617.00, "gang_count": 1.0, "task_hours": 3.2},
            {"code": "0113", "name": "Chowkidar", "unit": "day", "coefficient": 0.008, "rate": 558.00, "gang_count": 1.0, "task_hours": 0.064},
        ],
        "sundries": [
            {"code": "9999", "name": "Sundries", "unit": "LS", "coefficient": 2.73, "rate": 2.00},
        ],
    },
    # 2.3.1 Banking excavated earth in layers (10 cum)
    {
        "item_code": "2.3.1",
        "description": "Banking excavated earth in layers not exceeding 20 cm in depth, breaking clods, watering, rolling each layer with ½ tonne roller, or wooden or steel rammers, and rolling every 3rd and top-most layer with power roller of minimum 8 tonnes and dressing up, in embankments for roads, flood banks, marginal banks, and guide banks etc., lead upto 50 m and lift upto 1.5 m : All kinds of soil",
        "batch_quantity": 10.0,
        "batch_unit": "cum",
        "dsr_rate": 543.40,
        "machinery": [
            {"code": "0003", "name": "Hire charges of Diesel Road Roller - 8 to 10 tonne", "unit": "day", "coefficient": 0.008, "rate": 3000.00, "task_hours": 0.064},
            {"code": "EQUIP-0.5T", "name": "½-tonne hand roller / wooden or steel rammers (durmats)", "unit": "day", "coefficient": 1.100, "rate": 0.00, "task_hours": 8.80},
        ],
        "labour": [
            {"code": "0114", "name": "Beldar", "unit": "day", "coefficient": 2.20, "rate": 558.00, "gang_count": 2.0, "task_hours": 8.8},
            {"code": "0115", "name": "Coolie", "unit": "day", "coefficient": 3.60, "rate": 558.00, "gang_count": 3.0, "task_hours": 9.6},
            {"code": "0101", "name": "Bhisti", "unit": "day", "coefficient": 0.40, "rate": 617.00, "gang_count": 1.0, "task_hours": 3.2},
            {"code": "0113", "name": "Chowkidar", "unit": "day", "coefficient": 0.008, "rate": 558.00, "gang_count": 1.0, "task_hours": 0.064},
        ],
        "sundries": [
            {"code": "9999", "name": "Sundries", "unit": "LS", "coefficient": 2.73, "rate": 2.00},
        ],
    },
    # 2.4 Deduct for not rolling with 8-tonne power roller (10 cum)
    {
        "item_code": "2.4",
        "description": "Deduct for not rolling with power roller of minimum 8 tonnes for banking excavated earth in layers not exceeding 20 cm in depth.",
        "batch_quantity": 10.0,
        "batch_unit": "cum",
        "dsr_rate": 42.95,
        "machinery": [
            {"code": "0003", "name": "Hire charges of Diesel Road Roller - 8 to 10 tonne", "unit": "day", "coefficient": 0.008, "rate": 3000.00, "task_hours": 0.064},
        ],
        "labour": [
            {"code": "0113", "name": "Chowkidar", "unit": "day", "coefficient": 0.008, "rate": 558.00, "gang_count": 1.0, "task_hours": 0.064},
        ],
        "sundries": [
            {"code": "9999", "name": "Sundries", "unit": "LS", "coefficient": 1.82, "rate": 2.00},
        ],
        "notes": "DEDUCTION item: subtracts road roller, chowkidar, and partial sundries",
    },
    # 2.5 Deduct for not watering excavated earth for banking (10 cum)
    {
        "item_code": "2.5",
        "description": "Deduct for not watering the excavated earth for banking.",
        "batch_quantity": 10.0,
        "batch_unit": "cum",
        "dsr_rate": 326.95,
        "labour": [
            {"code": "0101", "name": "Bhisti", "unit": "day", "coefficient": 0.40, "rate": 617.00, "gang_count": 1.0, "task_hours": 3.2},
        ],
        "notes": "DEDUCTION item: subtracts watering bhishti",
    },
    # 2.32 Clearing grass and removal of rubbish (100 sqm)
    {
        "item_code": "2.32",
        "description": "Clearing grass and removal of the rubbish up to a distance of 50 m outside the periphery of the area cleared.",
        "batch_quantity": 100.0,
        "batch_unit": "sqm",
        "dsr_rate": 10.20,
        "labour": [
            {"code": "0114", "name": "Beldar", "unit": "day", "coefficient": 0.80, "rate": 558.00, "gang_count": 2.0, "task_hours": 3.2},
            {"code": "0115", "name": "Coolie", "unit": "day", "coefficient": 0.67, "rate": 558.00, "gang_count": 2.0, "task_hours": 2.68},
        ],
        "notes": "ADDITION item",
    },
    # 2.25 Filling in plinth / trenches in 20 cm layers (10 cum)
    {
        "item_code": "2.25",
        "description": "Filling available excavated earth (excluding rock) in trenches, plinth, sides of foundations etc. in layers not exceeding 20cm in depth, consolidating each deposited layer by ramming and watering, lead up to 50 m and lift upto 1.5 m.",
        "batch_quantity": 10.0,
        "batch_unit": "cum",
        "dsr_rate": 232.85,
        "machinery": [
            {"code": "EQUIP-DURMAT", "name": "Wooden or steel rammers (durmats) for layer compaction", "unit": "day", "coefficient": 0.800, "rate": 0.00, "task_hours": 6.40},
        ],
        "labour": [
            {"code": "0114", "name": "Beldar (spreading & durmat ramming)", "unit": "day", "coefficient": 0.80, "rate": 558.00, "gang_count": 1.0, "task_hours": 6.40},
            {"code": "0115", "name": "Coolie (basket haulage <=50m)", "unit": "day", "coefficient": 0.85, "rate": 558.00, "gang_count": 1.0, "task_hours": 6.80},
            {"code": "0101", "name": "Bhisti (layer watering to OMC)", "unit": "day", "coefficient": 0.20, "rate": 617.00, "gang_count": 1.0, "task_hours": 1.60},
        ],
        "notes": "Manual trench/plinth backfilling with durmat rammers",
    },
)


class CPWDGangRegistry:
    """Registry of CPWD Gang Compositions and dynamic composition engine."""

    def __init__(self) -> None:
        self._items: dict[str, ItemGang] = {}
        self._load_earthwork_gangs()

    def _load_earthwork_gangs(self) -> None:
        for raw in _EARTHWORK_ITEMS:
            code = raw["item_code"]
            machinery = [
                GangResource(
                    code=r["code"],
                    name=r["name"],
                    category="Machinery",
                    coefficient=float(r["coefficient"]),
                    unit=r["unit"],
                    rate=float(r["rate"]),
                    task_hours=r.get("task_hours"),
                )
                for r in raw.get("machinery", [])
            ]
            labour = [
                GangResource(
                    code=r["code"],
                    name=r["name"],
                    category="Labour",
                    coefficient=float(r["coefficient"]),
                    unit=r["unit"],
                    rate=float(r["rate"]),
                    gang_count=r.get("gang_count"),
                    task_hours=r.get("task_hours"),
                )
                for r in raw.get("labour", [])
            ]
            materials = [
                GangResource(
                    code=r["code"],
                    name=r["name"],
                    category="Material",
                    coefficient=float(r["coefficient"]),
                    unit=r["unit"],
                    rate=float(r["rate"]),
                )
                for r in raw.get("materials", [])
            ]
            sundries = [
                GangResource(
                    code=r["code"],
                    name=r["name"],
                    category="Sundries",
                    coefficient=float(r["coefficient"]),
                    unit=r["unit"],
                    rate=float(r["rate"]),
                )
                for r in raw.get("sundries", [])
            ]

            self._items[code] = ItemGang(
                item_code=code,
                description=raw["description"],
                batch_quantity=float(raw["batch_quantity"]),
                batch_unit=raw["batch_unit"],
                dsr_rate=raw.get("dsr_rate"),
                machinery=machinery,
                labour=labour,
                materials=materials,
                sundries=sundries,
                notes=raw.get("notes", ""),
            )

    def get_item(self, item_code: str) -> ItemGang | None:
        return self._items.get(item_code)

    def all_items(self) -> list[ItemGang]:
        return list(self._items.values())

    def compose_gang(
        self,
        base_item_code: str,
        deduction_codes: list[str] | tuple[str, ...],
        addition_codes: list[str] | tuple[str, ...],
    ) -> ComposedGang:
        """Compose a net gang by subtracting deductions and adding additions to base."""
        base = self._items.get(base_item_code)
        if base is None:
            raise ValueError(f"Unknown base item code: {base_item_code}")

        # Clone base resources
        net_machinery = {res.code: res.copy() for res in base.machinery}
        net_labour = {res.code: res.copy() for res in base.labour}
        net_materials = {res.code: res.copy() for res in base.materials}
        net_sundries = {res.code: res.copy() for res in base.sundries}

        deducted_resources: list[GangResource] = []
        added_resources: list[GangResource] = []

        # Process deductions
        for d_code in deduction_codes:
            deduct_item = self._items.get(d_code)
            if deduct_item is None:
                continue

            for res in deduct_item.machinery:
                if res.code in net_machinery:
                    rem = max(0.0, net_machinery[res.code].coefficient - res.coefficient)
                    net_machinery[res.code] = net_machinery[res.code].copy(coefficient=round(rem, 4))
                    deducted_resources.append(res.copy())

            for res in deduct_item.labour:
                if res.code in net_labour:
                    rem = max(0.0, net_labour[res.code].coefficient - res.coefficient)
                    net_labour[res.code] = net_labour[res.code].copy(coefficient=round(rem, 4))
                    deducted_resources.append(res.copy())

            for res in deduct_item.sundries:
                if res.code in net_sundries:
                    rem = max(0.0, net_sundries[res.code].coefficient - res.coefficient)
                    net_sundries[res.code] = net_sundries[res.code].copy(coefficient=round(rem, 4))
                    deducted_resources.append(res.copy())

            for res in deduct_item.materials:
                if res.code in net_materials:
                    rem = max(0.0, net_materials[res.code].coefficient - res.coefficient)
                    net_materials[res.code] = net_materials[res.code].copy(coefficient=round(rem, 4))
                    deducted_resources.append(res.copy())

        # Process additions (if compatible measurement basis)
        for a_code in addition_codes:
            add_item = self._items.get(a_code)
            if add_item is None:
                continue
            # When batch units match, combine directly; otherwise add as separate line
            scale = 1.0
            if add_item.batch_unit == base.batch_unit:
                scale = base.batch_quantity / add_item.batch_quantity

            for res in add_item.labour:
                scaled_coeff = round(res.coefficient * scale, 4)
                if res.code in net_labour:
                    new_coeff = round(net_labour[res.code].coefficient + scaled_coeff, 4)
                    net_labour[res.code] = net_labour[res.code].copy(coefficient=new_coeff)
                else:
                    net_labour[res.code] = res.copy(coefficient=scaled_coeff)
                added_resources.append(res.copy(coefficient=scaled_coeff))

            for res in add_item.machinery:
                scaled_coeff = round(res.coefficient * scale, 4)
                if res.code in net_machinery:
                    new_coeff = round(net_machinery[res.code].coefficient + scaled_coeff, 4)
                    net_machinery[res.code] = net_machinery[res.code].copy(coefficient=new_coeff)
                else:
                    net_machinery[res.code] = res.copy(coefficient=scaled_coeff)
                added_resources.append(res.copy(coefficient=scaled_coeff))

        return ComposedGang(
            base_item=base,
            net_machinery=list(net_machinery.values()),
            net_labour=list(net_labour.values()),
            net_materials=list(net_materials.values()),
            net_sundries=list(net_sundries.values()),
            deducted_resources=deducted_resources,
            added_resources=added_resources,
            batch_quantity=base.batch_quantity,
            batch_unit=base.batch_unit,
        )

    def export_to_ini(self, output_path: str | Path) -> Path:
        """Export gang registry to clean INI file for user inspection (pure physical gang mix)."""
        target = Path(output_path)
        target.parent.mkdir(parents=True, exist_ok=True)
        config = configparser.ConfigParser(interpolation=None)

        for item in self._items.values():
            section = f"Gang.Item_{item.item_code}"
            mach_str = ", ".join(
                f"{r.code}:{r.name.replace(':', '-')}:{r.coefficient:.4f}{r.unit}" + (f"({r.task_hours}h)" if r.task_hours else "")
                for r in item.machinery
            ) or "None"
            lab_str = ", ".join(
                f"{r.code}:{r.name.replace(':', '-')}:{r.coefficient:.4f}{r.unit}" + (f"(gang_{r.gang_count}x{r.task_hours}h)" if r.gang_count and r.task_hours else "")
                for r in item.labour
            ) or "None"
            mat_str = ", ".join(f"{r.code}:{r.name.replace(':', '-')}:{r.coefficient:.4f}{r.unit}" for r in item.materials) or "None"
            sun_str = ", ".join(f"{r.code}:{r.name.replace(':', '-')}:{r.coefficient:.4f}{r.unit}" for r in item.sundries) or "None"

            clean_desc = (item.description or "").replace("\n", " ").replace("\r", " ").strip()
            clean_notes = (item.notes or "Standard CPWD gang specification").replace("\n", " ").replace("\r", " ").strip()

            config[section] = {
                "item_code": item.item_code,
                "description": clean_desc,
                "batch_quantity": str(item.batch_quantity),
                "batch_unit": item.batch_unit,
                "machinery_mix": mach_str,
                "labour_mix": lab_str,
                "material_mix": mat_str,
                "sundries_mix": sun_str,
                "notes": clean_notes,
            }

        with open(target, "w", encoding="utf-8") as f:
            config.write(f)
        return target

    def export_to_txt(self, output_path: str | Path) -> Path:
        """Export gang registry to cleanly formatted text report (pure physical gang mix)."""
        target = Path(output_path)
        target.parent.mkdir(parents=True, exist_ok=True)
        lines = [
            "=" * 90,
            "CPWD DAR 2019 - GANG COMPOSITION REGISTRY",
            "Physical resource mix: Activity -> labour + machine + equipment mix -> output basis",
            "=" * 90,
            "",
        ]
        for item in self._items.values():
            lines.append(f"[{item.item_code}] {item.description}")
            lines.append(f"  Standard Batch Output: {item.batch_quantity} {item.batch_unit}")
            lines.append("  Physical Gang Composition (Labour, Machinery & Materials):")
            for m in item.machinery:
                hours_info = f" (Plant time: {m.task_hours} machine-hrs)" if m.task_hours else ""
                lines.append(f"    [Plant/Machinery] {m.code} {m.name}: {m.coefficient:.4f} {m.unit}{hours_info}")
            for l in item.labour:
                gang_info = f" (Gang of {l.gang_count} for {l.task_hours}h = {l.gang_count * l.task_hours:.2f} man-hrs)" if l.gang_count and l.task_hours else ""
                lines.append(f"    [Labour]          {l.code} {l.name}: {l.coefficient:.4f} {l.unit}{gang_info}")
            for mat in item.materials:
                lines.append(f"    [Material]        {mat.code} {mat.name}: {mat.coefficient:.4f} {mat.unit}")
            for s in item.sundries:
                lines.append(f"    [Sundries]        {s.code} {s.name}: {s.coefficient:.4f} {s.unit}")
            if item.notes:
                lines.append(f"  Notes: {item.notes}")
            lines.append("")

        with open(target, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
        return target


# Global shared singleton
GANG_REGISTRY = CPWDGangRegistry()
