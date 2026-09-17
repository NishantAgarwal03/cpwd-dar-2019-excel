"""CPWD Item Parser — Volume 2 (Ch. 13-26).

Reads ch13_items.json … ch26_items.json from data/reference_json/ and writes:
  data/gang_registry_vol2.txt / .ini
  data/productivity_registry_vol2.txt / .ini

Run directly:
    python scripts/cpwd_item_parser_vol2.py
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from scripts.cpwd_gang_registry import CPWDGangRegistry, GangResource, ItemGang
from scripts.cpwd_productivity_registry import CPWDProductivityRegistry

SHIFT_HOURS = 8.0

# ── Chapter metadata ──────────────────────────────────────────────────────────
CHAPTERS: dict[int, tuple[str, str]] = {
    13: ("13_Finishing",    "13_Finishing_Works"),
    14: ("14_Repairs",      "14_Repair_of_Buildings"),
    15: ("15_Dismantling",  "15_Dismantling_Works"),
    16: ("16_Roads",        "16_Roads_and_Pavements"),
    17: ("17_Drainage",     "17_Drainage"),
    18: ("18_WSI",          "18_Water_Supply_Sanitary"),
    19: ("19_Drainage2",    "19_Drainage_Sanitary"),
    20: ("20_Piling",       "20_Piling_Work"),
    21: ("21_Precast",      "21_Precast_Concrete"),
    22: ("22_WaterProof",   "22_Water_Proofing"),
    23: ("23_Horticulture", "23_Horticulture"),
    24: ("24_GlassAlum",    "24_Glass_Aluminium"),
    25: ("25_Misc",         "25_Miscellaneous"),
    26: ("26_Carriage",     "26_Carriage_Miscellaneous"),
}

_SKIP_CODES = {"9999", "9977", "9988"}
_JSON_DIR = Path(__file__).resolve().parent.parent / "data" / "reference_json"


def _classify(code: str, desc: str, unit: str, rate: float) -> str | None:
    """Return 'Labour' / 'Machinery' / 'Material' / 'Sundries', or None to skip."""
    code = code.strip()
    if code in _SKIP_CODES:
        return "Sundries"
    if unit.upper() in ("L.S.", "LS") and rate <= 2.0:
        return "Sundries"
    try:
        code_int = int(code)
        if 1 <= code_int <= 99:
            return "Machinery"
        if 100 <= code_int <= 199:
            return "Labour"
    except ValueError:
        pass
    return "Material"


def ingest_vol2_items(
    gang_reg: CPWDGangRegistry,
    prod_reg: CPWDProductivityRegistry,
) -> tuple[int, int]:
    """Ingest all Ch. 13–26 items. Returns (items_count, productivity_entries_count)."""
    total_items = 0

    for ch_num, (short_name, subhead) in CHAPTERS.items():
        json_path = _JSON_DIR / f"ch{ch_num}_items.json"
        if not json_path.exists():
            print(f"  [SKIP] {json_path.name} not found", flush=True)
            continue

        with open(json_path, encoding="utf-8", errors="replace") as f:
            items: list[dict] = json.load(f)

        # Build per-code counters to de-duplicate variants with the same code
        code_counts: dict[str, int] = {}
        for item in items:
            raw_code = str(item.get("code", "")).strip()
            code_counts[raw_code] = code_counts.get(raw_code, 0) + 1
        code_seen: dict[str, int] = {}
        needs_suffix = {c for c, n in code_counts.items() if n > 1}

        for item in items:
            raw_code = str(item.get("code", "")).strip()
            if not raw_code:
                continue
            # Append variant index for duplicate codes
            if raw_code in needs_suffix:
                code_seen[raw_code] = code_seen.get(raw_code, 0) + 1
                item_code = f"{raw_code}.{code_seen[raw_code]}"
            else:
                item_code = raw_code

            basis = float(item.get("basis", 10.0))
            batch_unit = str(item.get("unit", "")).strip() or "nos"
            resources_raw: list[dict] = item.get("resources", [])

            machinery_list: list[GangResource] = []
            labour_list: list[GangResource] = []
            material_list: list[GangResource] = []
            sundries_list: list[GangResource] = []

            for r in resources_raw:
                r_code = str(r.get("code", "")).strip()
                r_desc = str(r.get("desc", "")).strip()
                r_unit = str(r.get("unit", "")).strip()
                r_qty = float(r.get("qty", 0.0))
                r_rate = float(r.get("rate", 0.0))

                cat = _classify(r_code, r_desc, r_unit, r_rate)
                if cat is None:
                    continue

                task_hours: float | None = None
                gang_count: float | None = None
                if cat in ("Labour", "Machinery"):
                    task_hours = round(abs(r_qty) * SHIFT_HOURS, 4)
                    if cat == "Labour":
                        gang_count = 1.0

                res = GangResource(
                    code=r_code,
                    name=r_desc,
                    category=cat,
                    coefficient=r_qty,
                    unit=r_unit,
                    rate=r_rate,
                    gang_count=gang_count,
                    task_hours=task_hours,
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
                item_code=item_code,
                description=f"CPWD DAR 2019 Vol 2 {subhead} — Item {item_code}",
                batch_quantity=basis,
                batch_unit=batch_unit,
                dsr_rate=item.get("say"),
                machinery=machinery_list,
                labour=labour_list,
                materials=material_list,
                sundries=sundries_list,
                notes=f"CPWD DAR 2019 {subhead} specification",
            )
            gang_reg._items[item_code] = item_gang

            for res in [*machinery_list, *labour_list]:
                if res.coefficient != 0:
                    raw_bench = {
                        "subhead": subhead,
                        "item_code": item_code,
                        "activity": item_gang.description,
                        "resource_code": res.code,
                        "resource_name": res.name,
                        "resource_type": res.category,
                        "coefficient": res.coefficient,
                        "unit": res.unit,
                        "batch_quantity": basis,
                        "batch_unit": batch_unit,
                        "machine_hours": res.task_hours if res.category == "Machinery" else None,
                        "task_hours": res.task_hours if res.category == "Labour" else None,
                        "gang_size": res.gang_count,
                        "source_citation": f"CPWD DAR 2019 Vol 2 Item {item_code}",
                    }
                    entry = prod_reg._build_entry(raw_bench)
                    key = (entry.subhead, entry.item_code, entry.resource_code)
                    prod_reg._entries[key] = entry

            total_items += 1

        print(f"  Ch {ch_num:02d} ({subhead}): {len(items)} items ingested", flush=True)

    return total_items, len(prod_reg._entries)


def build_vol2_registries(out_dir: str | Path = "data") -> tuple[Path, Path, Path, Path]:
    """Build and export all four Vol 2 registry files. Returns (g_txt, g_ini, p_txt, p_ini)."""
    gang_reg = CPWDGangRegistry()
    prod_reg = CPWDProductivityRegistry()
    # Clear Vol 1 earthwork data that both __init__ methods pre-load
    gang_reg._items.clear()
    prod_reg._entries.clear()

    print("Ingesting Vol 2 items (Ch. 13–26) …", flush=True)
    n_items, n_entries = ingest_vol2_items(gang_reg, prod_reg)
    print(f"  → {n_items} items, {n_entries} productivity entries", flush=True)

    out = Path(out_dir)
    g_txt = gang_reg.export_to_txt(str(out / "gang_registry_vol2.txt"))
    g_ini = gang_reg.export_to_ini(str(out / "gang_registry_vol2.ini"))
    p_txt = prod_reg.export_to_txt(str(out / "productivity_registry_vol2.txt"))
    p_ini = prod_reg.export_to_ini(str(out / "productivity_registry_vol2.ini"))

    print(f"Exported:\n  {g_txt}\n  {g_ini}\n  {p_txt}\n  {p_ini}", flush=True)
    return g_txt, g_ini, p_txt, p_ini


if __name__ == "__main__":
    build_vol2_registries()
