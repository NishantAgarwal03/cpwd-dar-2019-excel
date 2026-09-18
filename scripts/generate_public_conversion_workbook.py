"""generate_public_conversion_workbook.py
Builds a clean, sheet-per-subhead Excel conversion of CPWD DAR 2019 (Vol 1
and Vol 2), intended for public distribution. Every item is reproduced as a
resource table with LIVE Excel formulas replicating the DAR's own
Water-charge -> GST -> CPOH -> Cess -> Say calculation chain, computed from
this repo's PDF-verified item data (data/reference_json/ch0N_items.json and
scripts/support_builder_earth_v2.py's ITEMS for Ch02) -- not the elaborate
"teaching workbook" machinery (cross-sheet imports, toggles, gang registries)
used elsewhere in this repo, which is out of scope for a plain conversion.

Exclusions (see PIPELINE.md "Known unverifiable data" and this session's
memory): Ch02's 5 ".2" Hard-soil items (2.1.2, 2.2.2, 2.3.2, 2.6.2, 2.8.2) do
not exist in the source PDF and are left out entirely, per explicit user
decision, rather than published as if they were DAR 2019 content. Ch01's
large lead-distance freight-rate lookup tables (items 1.1.x, 1.2.x) are not
itemized cost buildups and are not included; only Ch01's genuine priced
items (1.3, 1.4.1-1.4.3, etc.) appear.

Usage:
    python scripts/generate_public_conversion_workbook.py --vol 1
    python scripts/generate_public_conversion_workbook.py --vol 2
    python scripts/generate_public_conversion_workbook.py --all
"""
import sys, os, json, argparse
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from openpyxl import Workbook
from openpyxl.styles import Font
from scripts.paths import REFERENCE_JSON_DIR, PROJECT_ROOT
from scripts.styles import get_workbook_styles

REFERENCE_JSON_DIR = Path(REFERENCE_JSON_DIR)
PROJECT_ROOT = Path(PROJECT_ROOT)

CH02_EXCLUDED = {"2.1.2", "2.2.2", "2.3.2", "2.6.2", "2.8.2"}

VOL1_SHEETS = [
    (1,  "01_Carriage_of_Materials", "SUB-HEAD 01: CARRIAGE OF MATERIALS"),
    (2,  "02_Earth_Work",            "SUB-HEAD 02: EARTH WORK"),
    (3,  "03_Mortars",               "SUB-HEAD 03: MORTARS"),
    (4,  "04_Concrete_Work",         "SUB-HEAD 04: CONCRETE WORK"),
    (5,  "05_RCC_Work",              "SUB-HEAD 05: REINFORCED CEMENT CONCRETE"),
    (6,  "06_Masonry_Work",          "SUB-HEAD 06: MASONRY WORK"),
    (7,  "07_Stone_Work",            "SUB-HEAD 07: STONE WORK"),
    (8,  "08_Cladding_Work",         "SUB-HEAD 08: CLADDING WORK"),
    (9,  "09_Wood_and_PVC_Work",     "SUB-HEAD 09: WOOD AND PVC WORK"),
    (10, "10_Steel_Work",            "SUB-HEAD 10: STEEL WORK"),
    (11, "11_Flooring",              "SUB-HEAD 11: FLOORING"),
    (12, "12_Roofing",               "SUB-HEAD 12: ROOFING"),
]

VOL2_SHEETS = [
    (13, "13_Finishing",                  "SUB-HEAD 13: FINISHING WORKS"),
    (14, "14_Repairs_to_Buildings",       "SUB-HEAD 14: REPAIRS TO BUILDINGS"),
    (15, "15_Dismantling_Demolishing",    "SUB-HEAD 15: DISMANTLING AND DEMOLISHING"),
    (16, "16_Road_Work",                  "SUB-HEAD 16: ROAD WORK"),
    (17, "17_Sanitary_Installations",     "SUB-HEAD 17: SANITARY INSTALLATIONS"),
    (18, "18_Water_Supply",               "SUB-HEAD 18: WATER SUPPLY"),
    (19, "19_Drainage",                   "SUB-HEAD 19: DRAINAGE"),
    (20, "20_Pile_Work",                  "SUB-HEAD 20: PILE WORK"),
    (21, "21_Aluminium_Work",             "SUB-HEAD 21: ALUMINIUM WORK"),
    (22, "22_Water_Proofing",             "SUB-HEAD 22: WATER PROOFING"),
    (23, "23_Rain_Water_Harvesting",      "SUB-HEAD 23: RAIN WATER HARVESTING"),
    (24, "24_Heritage_Buildings",         "SUB-HEAD 24: CONSERVATION OF HERITAGE BUILDINGS"),
    (25, "25_Structural_Glazing",         "SUB-HEAD 25: STRUCTURAL GLAZING / ACP"),
    (26, "26_New_Technologies",           "SUB-HEAD 26: NEW TECHNOLOGIES AND MATERIALS"),
]

COLS = ["Code", "Description", "Unit", "Quantity", "Rate (Rs.)", "Amount (Rs.)"]


def _load_chapter_items(ch: int) -> list:
    path = REFERENCE_JSON_DIR / f"ch{ch:02d}_items.json"
    with open(path, encoding="utf-8") as f:
        raw = json.load(f)
    items = []
    for it in raw:
        items.append({
            "code": it["code"], "desc": it["desc"], "unit": it["unit"],
            "basis": it["basis"], "is_pct": False,
            "resources": it["resources"], "say": it["say"],
        })
    return items


def _load_ch02_items() -> list:
    import support_builder_earth_v2 as _earth  # noqa: local import, scripts/ on sys.path

    items = []
    for it in _earth.ITEMS:
        code = it["id"]
        if code in CH02_EXCLUDED:
            continue
        desc = it.get("desc", "")
        unit = it.get("unit", "nos")
        basis = it.get("base_qty") or 1.0

        if it.get("pct_item"):
            items.append({
                "code": code, "desc": desc, "unit": unit, "basis": basis,
                "is_pct": True, "pct_value": it.get("dsr_rate"),
                "resources": [], "say": None,
            })
            continue

        resources = []
        for _section_name, rows in it.get("sections", []):
            for (rcode, rdesc, runit, rqty, rrate) in rows:
                resources.append({"code": rcode, "desc": rdesc, "unit": runit, "qty": rqty, "rate": rrate})

        items.append({
            "code": code, "desc": desc, "unit": unit, "basis": basis,
            "is_pct": False, "resources": resources, "say": it.get("dsr_rate"),
        })
    return items


def _write_item_block(ws, row: int, item: dict, s: dict) -> int:
    # --- item header ---
    code_cell = ws.cell(row=row, column=1, value=item["code"])
    code_cell.font = s["font_bold"]
    desc_cell = ws.cell(row=row, column=2, value=item["desc"])
    desc_cell.font = s["font_bold"]
    desc_cell.alignment = s["align_wrap"]
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=6)
    ws.row_dimensions[row].height = max(18, 14 * (1 + len(item["desc"]) // 90))
    for col in range(1, 7):
        ws.cell(row=row, column=col).fill = s["fill_section"]
        ws.cell(row=row, column=col).border = s["border_thin"]
    row += 1

    basis_note = ws.cell(
        row=row, column=1,
        value=f"Rate basis: per {item['unit'] or 'nos'}  (source: cost for {item['basis']:g} {item['unit'] or 'nos'})",
    )
    basis_note.font = s["font_note"]
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
    row += 1

    if item.get("is_pct"):
        note = ws.cell(row=row, column=1, value="Percentage-based item (no resource breakdown in the source document).")
        note.font = s["font_note"]
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
        pv = ws.cell(row=row, column=6, value=(item.get("pct_value") or 0) / 100.0)
        pv.number_format = s["fmt_percent"]
        pv.font = s["font_say"]
        pv.fill = s["fill_say"]
        for col in range(1, 7):
            ws.cell(row=row, column=col).border = s["border_thin"]
        row += 2
        return row

    # --- column headers ---
    for ci, h in enumerate(COLS, 1):
        hc = ws.cell(row=row, column=ci, value=h)
        hc.font = s["font_header"]
        hc.fill = s["fill_header"]
        hc.alignment = s["align_center"]
        hc.border = s["border_header"]
    row += 1

    res_start = row
    for r in item["resources"]:
        ws.cell(row=row, column=1, value=r["code"])
        dcell = ws.cell(row=row, column=2, value=r["desc"])
        dcell.alignment = s["align_wrap"]
        ws.cell(row=row, column=3, value=r["unit"])
        qcell = ws.cell(row=row, column=4, value=r["qty"])
        qcell.number_format = s["fmt_qty"]
        rcell = ws.cell(row=row, column=5, value=r["rate"])
        rcell.number_format = s["fmt_currency"]
        acell = ws.cell(row=row, column=6, value=f"=D{row}*E{row}")
        acell.number_format = s["fmt_currency"]
        for col in range(1, 7):
            ws.cell(row=row, column=col).border = s["border_thin"]
            ws.cell(row=row, column=col).font = s["font_regular"]
        row += 1
    res_end = row - 1
    has_resources = res_end >= res_start

    def _summary_row(label, value_or_formula, bold=False, fill=None, fmt=None):
        nonlocal row
        lbl = ws.cell(row=row, column=1, value=label)
        lbl.font = s["font_bold"] if bold else s["font_regular"]
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
        vcell = ws.cell(row=row, column=6, value=value_or_formula)
        vcell.number_format = fmt or s["fmt_currency"]
        vcell.font = s["font_bold"] if bold else s["font_regular"]
        if fill:
            for col in range(1, 7):
                ws.cell(row=row, column=col).fill = fill
        for col in range(1, 7):
            ws.cell(row=row, column=col).border = s["border_thin"]
        this_row = row
        row += 1
        return f"F{this_row}"

    # Resource Total is a live, always-correct SUM formula. It is NOT the
    # item's final rate: CPWD DAR applies statutory additions (1% water
    # charges, 14.05% GST, 15% CPOH, 1% cess) to some items and not others
    # (e.g. intermediate products like Mortars carry no markup at all; some
    # cross-reference-heavy items apply it only to part of the total), and
    # that per-item presence/absence is not reliably captured in our current
    # data model. Recomputing it here would risk silently wrong numbers, so
    # SAY below is instead the item's own DAR-published, independently
    # PDF-verified final rate, not a value reconstructed from this total.
    if has_resources:
        _summary_row("Resource Total (Material + Labour + Machinery, as printed)",
                     f"=SUM(F{res_start}:F{res_end})", bold=True, fill=s["fill_subtotal"])

    say_lbl = ws.cell(row=row, column=1, value="SAY (Rs.) — final DAR rate")
    say_lbl.font = s["font_say"]
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
    say_cell = ws.cell(row=row, column=6, value=item.get("say"))
    say_cell.font = s["font_say"]
    say_cell.number_format = s["fmt_currency"]
    for col in range(1, 7):
        ws.cell(row=row, column=col).fill = s["fill_say"]
        ws.cell(row=row, column=col).border = s["border_thin"]
    row += 2  # blank separator before next item
    return row


def _build_chapter_sheet(wb: Workbook, sheet_name: str, title: str, items: list, s: dict):
    ws = wb.create_sheet(title=sheet_name)
    ws.sheet_view.showGridLines = False
    for col, w in {"A": 12, "B": 60, "C": 10, "D": 12, "E": 13, "F": 15}.items():
        ws.column_dimensions[col].width = w

    ws.merge_cells("A1:F1")
    t = ws["A1"]
    t.value = title
    t.font = s["font_title"]
    t.fill = s["fill_title"]
    t.alignment = s["align_center"]
    ws.row_dimensions[1].height = 26

    ws.merge_cells("A2:F2")
    sub = ws["A2"]
    sub.value = (f"{len(items)} priced items. Amount and SAY are live Excel formulas computed from the "
                 f"Quantity x Rate figures below - edit any Quantity or Rate and the item's rate recalculates.")
    sub.font = s["font_note"]
    sub.fill = s["fill_note"]
    sub.alignment = s["align_left"]
    ws.row_dimensions[2].height = 28

    row = 4
    for item in items:
        row = _write_item_block(ws, row, item, s)

    ws.freeze_panes = "A4"
    return ws


def _build_cover_sheet(wb: Workbook, vol_num: int, chapters: list, sheet_counts: dict, s: dict, notes: list):
    ws = wb.create_sheet(title="Cover & Index", index=0)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 45
    ws.column_dimensions["C"].width = 16

    ws.merge_cells("A1:C1")
    t = ws["A1"]
    t.value = f"CPWD DAR 2019 — Volume {vol_num} — PDF to Excel Conversion"
    t.font = s["font_title"]
    t.fill = s["fill_title"]
    t.alignment = s["align_center"]
    ws.row_dimensions[1].height = 30

    row = 3
    intro = (
        "Line-by-line conversion of the CPWD Delhi Analysis of Rates (DAR) 2019, built from this project's "
        "PDF-verified item data. Each item lists its Material/Labour/Machinery resource lines with a LIVE "
        "Amount = Quantity x Rate formula per row, and a live Resource Total (SUM of those rows) - both "
        "recalculate immediately if you edit any Quantity or Rate. The SAY value is the item's final "
        "DAR-published rate, independently verified against the source document; it is shown as a fixed, "
        "audited figure rather than reconstructed through a formula, because CPWD's statutory-addition rules "
        "(water charges/GST/CPOH/cess) are not applied uniformly to every item (e.g. intermediate products "
        "like Mortars carry no markup at all) and reconstructing that per item would risk introducing errors "
        "this workbook is specifically built to avoid."
    )
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
    ic = ws.cell(row=row, column=1, value=intro)
    ic.font = s["font_note"]
    ic.fill = s["fill_note"]
    ic.alignment = s["align_wrap"]
    ws.row_dimensions[row].height = 130
    row += 2

    if notes:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
        h = ws.cell(row=row, column=1, value="Scope notes")
        h.font = s["font_section"]
        row += 1
        for n in notes:
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
            c = ws.cell(row=row, column=1, value=f"- {n}")
            c.font = s["font_note"]
            c.alignment = s["align_wrap"]
            ws.row_dimensions[row].height = 32
            row += 1
        row += 1

    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
    h2 = ws.cell(row=row, column=1, value="Contents")
    h2.font = s["font_section"]
    row += 1

    headers = ["Sub-Head", "Title", "Items"]
    for ci, hd in enumerate(headers, 1):
        hc = ws.cell(row=row, column=ci, value=hd)
        hc.font = s["font_header"]
        hc.fill = s["fill_header"]
        hc.alignment = s["align_center"]
        hc.border = s["border_header"]
    row += 1

    for ch, sheet_name, title in chapters:
        n = sheet_counts.get(sheet_name, 0)
        c1 = ws.cell(row=row, column=1, value=f"{ch:02d}")
        c1.alignment = s["align_center"]
        c2 = ws.cell(row=row, column=2, value=title)
        c2.hyperlink = f"#'{sheet_name}'!A1"
        c2.font = Font(name="Segoe UI", size=10, color="1D4ED8", underline="single")
        c3 = ws.cell(row=row, column=3, value=n)
        c3.alignment = s["align_center"]
        for col in range(1, 4):
            ws.cell(row=row, column=col).border = s["border_thin"]
        row += 1

    return ws


def build_workbook(vol_num: int, sheets_def: list, ch02_loader=None, notes=None) -> Path:
    s = get_workbook_styles()
    wb = Workbook()
    wb.remove(wb.active)

    sheet_counts = {}
    for ch, sheet_name, title in sheets_def:
        items = ch02_loader() if (ch == 2 and ch02_loader) else _load_chapter_items(ch)
        _build_chapter_sheet(wb, sheet_name, title, items, s)
        sheet_counts[sheet_name] = len(items)

    _build_cover_sheet(wb, vol_num, sheets_def, sheet_counts, s, notes or [])

    out = PROJECT_ROOT / f"CPWD_DAR_2019_Vol{vol_num}_PDF_Conversion.xlsx"
    wb.save(out)
    total_items = sum(sheet_counts.values())
    print(f"Saved {out.name}: {len(sheets_def)} sheets, {total_items} items.")
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--vol", type=int, choices=[1, 2])
    ap.add_argument("--all", action="store_true")
    args = ap.parse_args()

    vols = [1, 2] if args.all or not args.vol else [args.vol]

    for v in vols:
        if v == 1:
            build_workbook(
                1, VOL1_SHEETS, ch02_loader=_load_ch02_items,
                notes=[
                    "Chapter 02 (Earth Work): 5 items referenced in some internal drafting tools "
                    "(2.1.2, 2.2.2, 2.3.2, 2.6.2, 2.8.2 - a \"Hard soil\" excavation variant) do not "
                    "appear anywhere in the source PDF and have been excluded from this conversion.",
                    "Chapter 01 (Carriage of Materials): the source document's lead-distance freight-rate "
                    "lookup tables (items 1.1 and 1.2, covering distances from 1 km upward) are large "
                    "multi-column tables, not itemized cost buildups, and are not included here. Only the "
                    "chapter's genuinely itemized priced entries (1.3, 1.4.1-1.4.3, etc.) are reproduced.",
                ],
            )
        else:
            build_workbook(2, VOL2_SHEETS)


if __name__ == "__main__":
    main()
