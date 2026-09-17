"""CPWD Item Parser — Volume 1 (Ch.02–12).

Single source of truth for the Vol 1 gang and productivity registries.

Data flow (matches Vol 2 pattern):
    Ch 02:    support_builder_earth_v2.ITEMS   → ingest_earthwork_items()
    Ch 03-12: data/reference_json/ch0N_items.json → ingest_vol1_trades_items()
              (JSON files produced by vol1_chapter_extractor.py from the converted XLSX)

Outputs (single pair — no duplicate _all_trades twin):
    data/gang_registry_vol1.txt / .ini
    data/productivity_registry_vol1.txt / .ini

Run:
    python scripts/cpwd_item_parser.py
"""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from scripts.cpwd_gang_registry import GANG_REGISTRY, GangResource, ItemGang
from scripts.cpwd_productivity_registry import PRODUCTIVITY_REGISTRY, SHIFT_HOURS
import scripts.support_builder_earth_v2 as _earth_v2

_JSON_DIR = Path(__file__).resolve().parent.parent / "data" / "reference_json"

# ── Chapter → subhead label ───────────────────────────────────────────────────
_TRADE_SUBHEADS: dict[int, str] = {
    3:  "03_Mortars",
    4:  "04_Concrete_Work",
    5:  "05_RCC_Work",
    6:  "06_Masonry_Work",
    7:  "07_Stone_Work",
    8:  "08_Cladding_Work",
    9:  "09_Wood_and_PVC_Work",
    10: "10_Steel_Work",
    11: "11_Flooring",
    12: "12_Roofing",
}

# ── Gang effort annotations (Ch 02) ──────────────────────────────────────────
_GANG_EFFORT_MAP: dict[tuple[str, str], tuple[float, float]] = {
    ("2.1.1", "0114"): (4.0, 13.60), ("2.1.1", "0115"): (4.0, 11.20),
    ("2.1.2", "0114"): (4.0, 17.20), ("2.1.2", "0115"): (4.0, 14.20),
    ("2.2.1", "0114"): (4.0, 11.80), ("2.2.1", "0115"): (3.0,  9.60),
    ("2.2.1", "0101"): (1.0,  3.20), ("2.2.1", "0113"): (1.0,  0.064),
    ("2.2",   "0114"): (4.0, 11.80), ("2.2",   "0115"): (3.0,  9.60),
    ("2.2",   "0101"): (1.0,  3.20), ("2.2",   "0113"): (1.0,  0.064),
    ("2.3.1", "0114"): (2.0,  8.80), ("2.3.1", "0115"): (3.0,  9.60),
    ("2.3.1", "0101"): (1.0,  3.20), ("2.3.1", "0113"): (1.0,  0.064),
    ("2.4",   "0113"): (1.0,  0.064),
    ("2.5",   "0101"): (1.0,  3.20),
    ("2.25",  "0114"): (1.0,  6.40), ("2.25",  "0115"): (1.0,  6.80),
    ("2.25",  "0101"): (1.0,  1.60),
    ("2.32",  "0114"): (2.0,  3.20), ("2.32",  "0115"): (2.0,  2.68),
}
_MACHINERY_HOURS_MAP: dict[tuple[str, str], float] = {
    ("2.2",   "0003"): 0.064, ("2.2",   "EQUIP-0.5T"): 8.80,
    ("2.2.1", "0003"): 0.064, ("2.2.1", "EQUIP-0.5T"): 8.80,
    ("2.3.1", "0003"): 0.064, ("2.3.1", "EQUIP-0.5T"): 8.80,
    ("2.4",   "0003"): -0.064,
    ("2.25",  "EQUIP-DURMAT"): 6.40,
    ("2.6.1", "0020"): 0.328, ("2.6.1", "0018"): 0.328,
    ("2.6.2", "0020"): 0.504, ("2.6.2", "0018"): 0.504,
    ("2.7.1", "0020"): 0.504, ("2.7.1", "0017"): 0.504,
    ("2.7.2", "0020"): 1.000, ("2.7.2", "0017"): 1.000,
    ("2.7.3", "0020"): 1.000, ("2.7.3", "0017"): 1.000,
}

_SKIP_CODES = {"9999", "9977", "9988"}


# ── Resource classification (same rule as Vol 2 parser) ──────────────────────
def _classify(code: str, unit: str = "", rate: float = 0.0) -> str | None:
    code = code.strip()
    if code in _SKIP_CODES:
        return "Sundries"
    if unit.upper() in ("L.S.", "LS") and rate <= 2.0:
        return "Sundries"
    try:
        c = int(code)
        if 1 <= c <= 99:
            return "Machinery"
        if 100 <= c <= 199:
            return "Labour"
    except ValueError:
        pass
    return "Material"


def _register_item(
    item_code: str,
    desc: str,
    basis: float,
    batch_unit: str,
    resources: list[dict[str, Any]],
    subhead: str,
    notes: str,
) -> None:
    """Build GangResource lists, write to GANG_REGISTRY and PRODUCTIVITY_REGISTRY."""
    machinery_list: list[GangResource] = []
    labour_list:    list[GangResource] = []
    material_list:  list[GangResource] = []
    sundries_list:  list[GangResource] = []

    for r in resources:
        code  = str(r.get("code", "")).strip()
        name  = str(r.get("name") or r.get("desc", "")).strip()
        unit  = str(r.get("unit", "day")).strip()
        coeff = float(r.get("coefficient") if "coefficient" in r else r.get("qty", 0))
        rate  = float(r.get("rate", 0.0))
        cat   = r.get("category") or _classify(code, unit, rate)
        if cat is None:
            continue

        task_hours: float | None = None
        gang_count: float | None = None

        if cat == "Labour":
            gs, th = _GANG_EFFORT_MAP.get((item_code, code), (None, None))
            task_hours = th if th is not None else round(abs(coeff) * SHIFT_HOURS, 4)
            gang_count = gs if gs is not None else 1.0
        elif cat == "Machinery":
            mh = _MACHINERY_HOURS_MAP.get((item_code, code))
            task_hours = mh if mh is not None else round(abs(coeff) * SHIFT_HOURS, 4)

        res = GangResource(
            code=code, name=name, category=cat,
            coefficient=coeff, unit=unit, rate=rate,
            gang_count=gang_count, task_hours=task_hours,
        )
        if cat == "Machinery":
            machinery_list.append(res)
        elif cat == "Labour":
            labour_list.append(res)
        elif cat == "Sundries":
            sundries_list.append(res)
        else:
            material_list.append(res)

    item_gang = ItemGang(
        item_code=item_code, description=desc,
        batch_quantity=basis, batch_unit=batch_unit, dsr_rate=None,
        machinery=machinery_list, labour=labour_list,
        materials=material_list, sundries=sundries_list,
        notes=notes,
    )
    GANG_REGISTRY._items[item_code] = item_gang

    for res in [*machinery_list, *labour_list]:
        if res.coefficient != 0:
            entry = PRODUCTIVITY_REGISTRY._build_entry({
                "subhead": subhead,
                "item_code": item_code,
                "activity": desc,
                "resource_code": res.code,
                "resource_name": res.name,
                "resource_type": res.category,
                "coefficient": res.coefficient,
                "unit": res.unit,
                "batch_quantity": basis,
                "batch_unit": batch_unit,
                "machine_hours": res.task_hours if res.category == "Machinery" else None,
                "task_hours":    res.task_hours if res.category == "Labour"    else None,
                "gang_size":     res.gang_count,
                "source_citation": f"CPWD DAR 2019 {subhead} Item {item_code}",
            })
            key = (entry.subhead, entry.item_code, entry.resource_code)
            PRODUCTIVITY_REGISTRY._entries[key] = entry


# ── Chapter 02: Earth Work ─────────────────────────────────────────────────────
def ingest_earthwork_items() -> int:
    """Ingest all Earth Work items from support_builder_earth_v2.ITEMS."""
    count = 0
    for raw in _earth_v2.ITEMS:
        item_code = raw["id"]
        desc      = raw.get("desc", "")
        batch_unit = raw.get("unit", "cum")
        basis     = float(raw.get("base_qty", 10.0))
        notes     = raw.get("notes", "CPWD DAR 2019 Specification")

        resources: list[dict[str, Any]] = []
        has_0003 = False

        for section_name, rows in raw.get("sections", []):
            for r in rows:
                code, r_desc, r_unit, coeff, rate = str(r[0]), str(r[1]), str(r[2]), float(r[3]), float(r[4])
                cat = {"MACHINERY": "Machinery", "LABOUR": "Labour",
                       "MATERIAL": "Material"}.get(section_name, "Sundries")
                if cat == "Material" and (code == "9999" or "sundries" in r_desc.lower()):
                    cat = "Sundries"
                resources.append({"code": code, "name": r_desc, "unit": r_unit,
                                   "coefficient": coeff, "rate": rate, "category": cat})
                if code == "0003":
                    has_0003 = True

        # Append implicit equipment for banking/rough-excavation items
        if has_0003 and item_code in ("2.2.1", "2.2.2", "2.3.1", "2.3.2"):
            resources.append({"code": "EQUIP-0.5T",
                               "name": "½-tonne hand roller / wooden or steel rammers (durmats)",
                               "unit": "day", "coefficient": 1.100, "rate": 0.0, "category": "Machinery"})
        if item_code == "2.25":
            resources.append({"code": "EQUIP-DURMAT",
                               "name": "Wooden or steel rammers (durmats) for layer compaction",
                               "unit": "day", "coefficient": 0.800, "rate": 0.0, "category": "Machinery"})

        _register_item(item_code, desc, basis, batch_unit, resources, "02_Earth_Work", notes)
        count += 1
    return count


# ── Chapters 03-12: JSON-driven ───────────────────────────────────────────────
def ingest_vol1_trades_items() -> int:
    """Ingest Ch.03–12 from ch03_items.json … ch12_items.json."""
    total = 0
    for ch, subhead in sorted(_TRADE_SUBHEADS.items()):
        json_path = _JSON_DIR / f"ch{ch:02d}_items.json"
        if not json_path.exists():
            print(f"  [SKIP] {json_path.name} not found — run vol1_chapter_extractor.py first")
            continue
        items: list[dict] = json.load(open(json_path, encoding="utf-8", errors="replace"))
        # De-duplicate variant items that share a code (same logic as Vol 2 parser)
        from collections import Counter
        counts = Counter(str(it.get("code", "")) for it in items)
        needs_suffix = {c for c, n in counts.items() if n > 1}
        seen: dict[str, int] = {}
        for item in items:
            raw_code = str(item.get("code", "")).strip()
            if not raw_code:
                continue
            if raw_code in needs_suffix:
                seen[raw_code] = seen.get(raw_code, 0) + 1
                item_code = f"{raw_code}.{seen[raw_code]}"
            else:
                item_code = raw_code
            _register_item(
                item_code=item_code,
                desc=f"CPWD DAR 2019 {subhead} — Item {item_code}",
                basis=float(item.get("basis", 1.0)),
                batch_unit=str(item.get("unit", "nos")),
                resources=item.get("resources", []),
                subhead=subhead,
                notes=f"CPWD DAR 2019 {subhead} specification",
            )
            total += 1
        print(f"  Ch {ch:02d} ({subhead}): {len(items)} items ingested", flush=True)
    return total


# ── Build everything and export ───────────────────────────────────────────────
def build_vol1_registries(out_dir: str | Path = "data") -> tuple[Path, Path, Path, Path]:
    """Ingest all Vol 1 items and write the four registry files."""
    print("Ingesting Vol 1 items (Ch.02–12) …", flush=True)
    n_earth = ingest_earthwork_items()
    print(f"  Ch 02 (02_Earth_Work): {n_earth} items ingested", flush=True)
    n_trades = ingest_vol1_trades_items()
    n_total = n_earth + n_trades
    n_prod = len(PRODUCTIVITY_REGISTRY._entries)
    print(f"  → {n_total} items total, {n_prod} productivity entries", flush=True)

    out = Path(out_dir)
    g_txt = GANG_REGISTRY.export_to_txt(str(out / "gang_registry_vol1.txt"))
    g_ini = GANG_REGISTRY.export_to_ini(str(out / "gang_registry_vol1.ini"))
    p_txt = PRODUCTIVITY_REGISTRY.export_to_txt(str(out / "productivity_registry_vol1.txt"))
    p_ini = PRODUCTIVITY_REGISTRY.export_to_ini(str(out / "productivity_registry_vol1.ini"))
    print(f"Exported:\n  {g_txt}\n  {g_ini}\n  {p_txt}\n  {p_ini}", flush=True)
    return g_txt, g_ini, p_txt, p_ini


if __name__ == "__main__":
    build_vol1_registries()
