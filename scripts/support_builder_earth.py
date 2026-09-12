"""
support_builder_earth.py
Creates / updates the sheet '02_support_earth_work' in the main workbook.
Covers CPWD DAR 2019 Vol-1 Sub-Head 2.0 Earth Work (pages 88-140).
Formulas: limited to SUM and basic arithmetic (*, +) — no INDEX/MATCH/VLOOKUP.
Run: python scripts/support_builder_earth.py
"""

import os
import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────────────────────────────────────
# Config
# ─────────────────────────────────────────────────────────────────────────────
from scripts.paths import WB_VOL1_LATEST_FILE

WB_PATH = WB_VOL1_LATEST_FILE

SHEET_NAME = "02_support_earth_work"

# ─────────────────────────────────────────────────────────────────────────────
# Style helpers
# ─────────────────────────────────────────────────────────────────────────────
def _font(bold=False, size=10, color="000000", italic=False):
    return Font(name="Calibri", bold=bold, size=size, color=color, italic=italic)

def _fill(hex_color):
    return PatternFill(fill_type="solid", fgColor=hex_color)

def _border(style="thin"):
    s = Side(style=style)
    return Border(left=s, right=s, top=s, bottom=s)

def _align(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

# Theme colours
C_HEAD1    = "1F4E79"   # dark blue  – section titles (white text)
C_HEAD2    = "2E75B6"   # med blue   – sub-section
C_HEAD3    = "BDD7EE"   # light blue – column headers
C_CAT_BG   = "D6E4F0"   # very light – category rows
C_ALTROW   = "F5FAFF"   # alt row bg
C_DEDUCT   = "FFF2CC"   # yellow     – deduct items
C_EXTRA    = "E2EFDA"   # light green – extra / add-on items
C_TOTAL    = "DDEBF7"   # blue tint  – total rows
C_BASIC_BG = "F4E6FF"   # lavender   – basic rates table

INR = '₹'          # ₹  (used only in header strings, not formula literals)

def _set(ws, row, col, value, font=None, fill=None, align=None, border=None,
         num_format=None):
    cell = ws.cell(row=row, column=col, value=value)
    if font:    cell.font    = font
    if fill:    cell.fill    = fill
    if align:   cell.alignment = align
    if border:  cell.border  = border
    if num_format: cell.number_format = num_format
    return cell


# ─────────────────────────────────────────────────────────────────────────────
# DATA
# ─────────────────────────────────────────────────────────────────────────────

# ── Basic Rates (page 95) ─────────────────────────────────────────────────────
BASIC_RATES = [
    ("0003", "Hire charges of Diesel Road Roller – 8 to 10 tonne",             "day",     3000.00),
    ("0017", "Hire and running charges of tipper",                              "day",     3750.00),
    ("0018", "Hire and running charges of loader",                              "day",     6000.00),
    ("0020", "Hydraulic Excavator (3D) with driver and fuel",                   "day",     7000.00),
    ("0101", "Bhisti",                                                          "day",      714.00),
    ("0103", "Blacksmith 2nd class",                                            "day",      714.00),
    ("0112", "Carpenter 2nd class",                                             "day",      714.00),
    ("0113", "Chowkidar",                                                       "day",      645.00),
    ("0114", "Beldar",                                                          "day",      645.00),
    ("0115", "Coolie",                                                          "day",      645.00),
    ("0124", "Mason (brick layer) 2nd class",                                   "day",      714.00),
    ("0128", "Mate",                                                            "day",      714.00),
    ("0132", "Rock Excavator",                                                  "day",      645.00),
    ("0133", "Rock Breaker",                                                    "day",      645.00),
    ("0134", "Rock Hole Driller",                                               "day",      645.00),
    ("0135", "Stone Chiseller",                                                 "day",      714.00),
    ("0302", "Safeda ballies 125 mm diameter",                                  "metre",     40.00),
    ("0325", "Blasting powder",                                                 "kg",        40.00),
    ("0326", "Blasting fuse (fuse wire)",                                       "each",      15.00),
    ("0771", "Kerosene oil",                                                    "litre",     50.00),
    ("1197", "Second class kail wood in scantling",                             "10 cudm",  260.00),
    ("1198", "Second class kail wood in planks",                                "10 cudm",  260.00),
    ("1235", "Diesel",                                                          "litre",     80.87),
    ("1980", "Fly ash",                                                         "cum",       11.00),
    ("2204", "Carriage of Timber",                                              "cum",      187.35),
    ("2262", "Carriage of Flyash",                                              "cum",      163.93),
    ("2335", "Carriage of sand",                                                "cum",      163.93),
    ("2342", "Carriage of Solvent / Diesel",                                    "quintal",   16.39),
    ("6501", "Sand zone V (Local)",                                             "cum",     1300.00),
    ("7022", "Chlorpyriphos 20% E.C.",                                          "litre",    150.00),
    ("9999", "Sundries",                                                        "L.S.",       2.12),
]

# ── Schedule of Rates – Earth Work items 2.1 to 2.38 ─────────────────────────
# Each entry: (item_code, description, unit, dar_rate, note_flag)
# note_flag: '' | 'DEDUCT' | 'EXTRA'
EW_ITEMS = [
    # ── 2.1  Surface Excavation ──────────────────────────────────────────────
    ("CAT", "2.1  Earth Work in Surface Excavation (≤30 cm depth, width >1.5 m, area >10 sqm)",
                                                                        "", None, ""),
    ("2.1.1",  "All kinds of soil  [disposal upto 50 m lead, 1.5 m lift]",
                                                                     "sqm",  107.00, ""),

    # ── 2.2  Banking – Rough Excavation ─────────────────────────────────────
    ("CAT", "2.2  Earth Work – Banking / Embankment (rough excavation, roller-compacted)",
                                                                        "", None, ""),
    ("2.2.1",  "All kinds of soil  [lead ≤50 m, lift ≤1.5 m]",
                                                                     "cum",  862.70, ""),

    # ── 2.3  Banking – Compaction only ──────────────────────────────────────
    ("CAT", "2.3  Banking Excavated Earth (without excavation, roller-compacted)",
                                                                        "", None, ""),
    ("2.3.1",  "All kinds of soil  [lead ≤50 m, lift ≤1.5 m]",
                                                                     "cum",  543.40, ""),

    # ── 2.4  Deduct – no power roller ───────────────────────────────────────
    ("2.4",    "Deduct for not rolling with power roller ≥8 t (banking)",
                                                                     "cum",    4.40, "DEDUCT"),

    # ── 2.5  Deduct – no watering ────────────────────────────────────────────
    ("2.5",    "Deduct for not watering excavated earth for banking",
                                                                     "cum",   38.20, "DEDUCT"),

    # ── 2.6  Open Area Excavation – Soil ────────────────────────────────────
    ("CAT", "2.6  Earth Work – Open Area Excavation by Mechanical / Manual Means (depth >30 cm)",
                                                                        "", None, ""),
    ("2.6.1",  "All kinds of soil",                                  "cum",  205.45, ""),

    # ── 2.7  Open Area Excavation – Rock ────────────────────────────────────
    ("CAT", "2.7  Open Area Excavation – Rock (same conditions as 2.6)",
                                                                        "", None, ""),
    ("2.7.1",  "Ordinary rock",                                      "cum",  412.95, ""),
    ("2.7.2",  "Hard rock (requiring blasting)",                     "cum",  711.35, ""),
    ("2.7.3",  "Hard rock (blasting prohibited)",                    "cum", 1184.30, ""),

    # ── 2.8  Foundation Trench – Soil ────────────────────────────────────────
    ("CAT", "2.8  Excavation in Foundation Trenches / Drains (width ≤1.5 m, plan ≤10 sqm) – Soil",
                                                                        "", None, ""),
    ("2.8.1",  "All kinds of soil  [lift ≤1.5 m, lead ≤50 m]",      "cum",  286.85, ""),

    # ── 2.9  Foundation Trench – Rock ────────────────────────────────────────
    ("CAT", "2.9  Excavation in Foundation Trenches / Drains – Rock",
                                                                        "", None, ""),
    ("2.9.1",  "Ordinary rock",                                      "cum",  523.50, ""),
    ("2.9.2",  "Hard rock (requiring blasting)",                     "cum",  846.25, ""),
    ("2.9.3",  "Hard rock (blasting prohibited)",                    "cum", 1258.60, ""),

    # ── 2.10  Pipe / Cable Trenches – Soil ──────────────────────────────────
    ("CAT", "2.10  Excavating Trenches for Pipes / Cables – Soil (depth ≤1.5 m, lead ≤50 m)",
                                                                        "", None, ""),
    ("2.10.1.1", "All kinds of soil – Pipes/cables ≤80 mm dia",    "metre",  255.55, ""),
    ("2.10.1.2", "All kinds of soil – Pipes/cables >80 mm, ≤300 mm dia",
                                                                    "metre",  417.35, ""),
    ("2.10.1.3", "All kinds of soil – Pipes/cables >300 mm, ≤600 mm dia",
                                                                    "metre",  651.55, ""),

    # ── 2.11  Extra depth >1.5 m to 3 m – Soil trench ───────────────────────
    ("2.11",   "Extra for pipe/cable trench in soil – depth >1.5 m to 3 m",
                                                                    "metre",  127.00, "EXTRA"),

    # ── 2.12  Extra depth >3 m to 4.5 m – Soil trench ───────────────────────
    ("2.12",   "Extra for pipe/cable trench in soil – depth >3 m to 4.5 m",
                                                                    "metre",  314.95, "EXTRA"),

    # ── 2.13  Pipe / Cable Trenches – Rock ──────────────────────────────────
    ("CAT", "2.13  Excavating Trenches for Pipes / Cables – Rock (depth ≤1.5 m, lead ≤50 m)",
                                                                        "", None, ""),
    # 2.13.1 Ordinary rock
    ("2.13.1.1", "Ordinary rock – Pipes/cables ≤80 mm dia",        "metre",  376.95, ""),
    ("2.13.1.2", "Ordinary rock – Pipes/cables >80 mm, ≤300 mm dia",
                                                                    "metre",  933.35, ""),
    ("2.13.1.3", "Ordinary rock – Pipes/cables >300 mm, ≤600 mm dia",
                                                                    "metre", 1074.00, ""),
    # 2.13.2 Hard rock – blasting required
    ("2.13.2.1", "Hard rock (blasting) – Pipes/cables ≤80 mm dia","metre",   531.85, ""),
    ("2.13.2.2", "Hard rock (blasting) – Pipes/cables >80 mm, ≤300 mm dia",
                                                                    "metre", 1316.90, ""),
    ("2.13.2.3", "Hard rock (blasting) – Pipes/cables >300 mm, ≤600 mm dia",
                                                                    "metre", 1515.20, ""),
    # 2.13.3 Hard rock – blasting prohibited
    ("2.13.3.1", "Hard rock (no blasting) – Pipes/cables ≤80 mm dia",
                                                                    "metre",  726.65, ""),
    ("2.13.3.2", "Hard rock (no blasting) – Pipes/cables >80 mm, ≤300 mm dia",
                                                                    "metre", 1799.35, ""),
    ("2.13.3.3", "Hard rock (no blasting) – Pipes/cables >300 mm, ≤600 mm dia",
                                                                    "metre", 2070.50, ""),

    # ── 2.14  Extra depth >1.5 m to 3 m – Rock trench ───────────────────────
    ("2.14",   "Extra for pipe/cable trench in rock – depth >1.5 m to 3 m",
                                                                    "metre",  103.75, "EXTRA"),

    # ── 2.15  Extra depth >3 m to 4.5 m – Rock trench ───────────────────────
    ("2.15",   "Extra for pipe/cable trench in rock – depth >3 m to 4.5 m",
                                                                    "metre",  256.15, "EXTRA"),

    # ── 2.16  Close Timbering – Trenches ────────────────────────────────────
    ("CAT", "2.16  Close Timbering in Trenches (face area measured)",
                                                                        "", None, ""),
    ("2.16.1", "Depth ≤1.5 m",                                      "sqm",  132.90, ""),
    ("2.16.2", "Depth >1.5 m to 3 m",                               "sqm",  145.55, ""),
    ("2.16.3", "Depth >3 m to 4.5 m",                               "sqm",  174.00, ""),

    # ── 2.17  Close Timbering – Shafts / Wells / Manholes ───────────────────
    ("CAT", "2.17  Close Timbering – Shafts, Wells, Cesspits, Manholes (face area)",
                                                                        "", None, ""),
    ("2.17.1", "Depth ≤1.5 m",                                      "sqm",  142.50, ""),
    ("2.17.2", "Depth >1.5 m to 3 m",                               "sqm",  169.35, ""),
    ("2.17.3", "Depth >3 m to 4.5 m",                               "sqm",  197.60, ""),

    # ── 2.18  Close Timbering – Open Areas ──────────────────────────────────
    ("CAT", "2.18  Close Timbering Over Areas (face area)",           "", None, ""),
    ("2.18.1", "Depth ≤1.5 m",                                      "sqm",  119.10, ""),
    ("2.18.2", "Depth >1.5 m to 3 m",                               "sqm",  134.20, ""),
    ("2.18.3", "Depth >3 m to 4.5 m",                               "sqm",  149.95, ""),

    # ── 2.19  Extra – Permanent Close Timbering ──────────────────────────────
    ("2.19",   "Extra for close timbering materials left permanently in position",
                                                                     "sqm", 1596.50, "EXTRA"),

    # ── 2.20  Open Timbering – Trenches ─────────────────────────────────────
    ("CAT", "2.20  Open Timbering in Trenches (face area)", "", None, ""),
    ("2.20.1", "Depth ≤1.5 m",                                      "sqm",   68.55, ""),
    ("2.20.2", "Depth >1.5 m to 3 m",                               "sqm",   76.40, ""),
    ("2.20.3", "Depth >3 m to 4.5 m",                               "sqm",   89.35, ""),

    # ── 2.21  Open Timbering – Shafts / Wells ────────────────────────────────
    ("CAT", "2.21  Open Timbering – Shafts, Wells, Cesspits, Manholes (face area)",
                                                                        "", None, ""),
    ("2.21.1", "Depth ≤1.5 m",                                      "sqm",   62.35, ""),
    ("2.21.2", "Depth >1.5 m to 3 m",                               "sqm",   74.95, ""),
    ("2.21.3", "Depth >3 m to 4.5 m",                               "sqm",   91.60, ""),

    # ── 2.22  Open Timbering – Open Areas ───────────────────────────────────
    ("CAT", "2.22  Open Timbering Over Areas (face area)", "", None, ""),
    ("2.22.1", "Depth ≤1.5 m",                                      "sqm",   42.40, ""),
    ("2.22.2", "Depth >1.5 m to 3 m",                               "sqm",   50.80, ""),
    ("2.22.3", "Depth >3 m to 4.5 m",                               "sqm",   64.30, ""),

    # ── 2.23  Extra – Permanent Open Timbering ──────────────────────────────
    ("2.23",   "Extra for open timbering materials left permanently in position",
                                                                     "sqm",  822.05, "EXTRA"),

    # ── 2.24  Extra – Water / Foul Position ─────────────────────────────────
    ("CAT", "2.24  Extra Rates for Works in Difficult Conditions", "", None, ""),
    ("2.24.1", "In or under water / liquid mud (incl. pumping) – % over basic",
                                                                   "metre depth", None, "EXTRA"),
    ("2.24.2", "In or under foul position (incl. pumping) – % over basic",
                                                                   "metre depth", None, "EXTRA"),

    # ── 2.25  Filling – available excavated earth ─────────────────────────────
    ("CAT", "2.25  Filling Works", "", None, ""),
    ("2.25",   "Filling available excavated earth in trenches/plinth (layers ≤20 cm, compacted)",
                                                                     "cum",  253.95, ""),
    ("2.25a",  "Excavating, supplying & filling local earth by mechanical transport (lead ≤5 km)",
                                                                     "cum",  368.65, ""),

    # ── 2.26  Extra lift ─────────────────────────────────────────────────────
    ("CAT", "2.26  Extra for Each Additional Lift of 1.5 m (or part thereof)",
                                                                        "", None, ""),
    ("2.26.1", "All kinds of soil",                                  "cum",  104.50, "EXTRA"),
    ("2.26.2", "Ordinary or hard rock",                              "cum",  187.40, "EXTRA"),

    # ── 2.27  Sand filling ───────────────────────────────────────────────────
    ("2.27",   "Supplying & filling plinth with sand under floors (watered, rammed, dressed)",
                                                                     "cum", 2161.20, ""),

    # ── 2.28  Surface dressing ───────────────────────────────────────────────
    ("CAT", "2.28  Surface Dressing of Ground", "", None, ""),
    ("2.28.1", "All kinds of soil  [remove vegetation & inequalities ≤15 cm, lead ≤50 m]",
                                                                     "sqm",   28.15, ""),

    # ── 2.29  Ploughing ──────────────────────────────────────────────────────
    ("CAT", "2.29  Ploughing Existing Ground (15 cm to 25 cm, watering)", "", None, ""),
    ("2.29.1", "All kinds of soil",                                  "sqm",   28.50, ""),

    # ── 2.30  Excavating holes ───────────────────────────────────────────────
    ("CAT", "2.30  Excavating Holes (>0.10 cum to ≤0.50 cum)", "", None, ""),
    ("2.30.1", "All kinds of soil",                                  "each",  89.90, ""),
    ("2.30.2", "Ordinary rock",                                      "each", 160.90, ""),
    ("2.30.3", "Hard rock (requiring blasting)",                     "each", 257.70, ""),
    ("2.30.4", "Hard rock (blasting prohibited)",                    "each", 381.40, ""),

    # ── 2.31  Clearing jungle ────────────────────────────────────────────────
    ("2.31",   "Clearing jungle incl. uprooting vegetation, grass, brushwood, trees ≤30 cm girth",
                                                                     "sqm",   14.50, ""),

    # ── 2.32  Clearing grass ─────────────────────────────────────────────────
    ("2.32",   "Clearing grass and removal of rubbish (50 m periphery)",
                                                                     "sqm",    7.40, ""),

    # ── 2.33  Felling trees ──────────────────────────────────────────────────
    ("CAT", "2.33  Felling Trees (incl. cutting, roots removal, stacking)", "", None, ""),
    ("2.33.1", "Girth >30 cm to 60 cm",                             "each",  439.25, ""),
    ("2.33.2", "Girth >60 cm to 120 cm",                            "each", 1957.15, ""),
    ("2.33.3", "Girth >120 cm to 240 cm",                           "each", 9084.05, ""),
    ("2.33.4", "Girth above 240 cm",                                "each", 18198.70, ""),

    # ── 2.34  Chemical emulsion supply ──────────────────────────────────────
    ("CAT", "2.34  Supplying Chemical Emulsion (Anti-Termite)", "", None, ""),
    ("2.34.1", "Chlorpyriphos emulsifiable concentrate 20% (in sealed containers)",
                                                                    "litre",  200.90, ""),

    # ── 2.35  Post-constructional anti-termite treatment ─────────────────────
    ("CAT", "2.35  Post-Constructional Anti-Termite Treatment – Diluting & Injecting",
                                                                        "", None, ""),
    ("2.35.1.1", "Along external wall (no apron) @ 7.5 L/sqm, depth 300 mm – Chlorpyriphos 1%",
                                                                    "metre",   32.30, ""),
    ("2.35.2.1", "Along external wall below concrete/masonry apron @ 2.25 L/m – Chlorpyriphos 1%",
                                                                    "metre",   44.70, ""),
    ("2.35.3.1", "Soil under existing floors @ 1 L/hole, 300 mm apart – Chlorpyriphos 1%",
                                                                     "sqm",  256.15, ""),
    ("2.35.4.1", "Existing masonry @ 1 L/hole, 300 mm interval, 45 deg – Chlorpyriphos 1%",
                                                                    "metre",   35.75, ""),
    ("2.35.5",   "Points of contact of woodwork, 0.5 L/hole, 6 mm dia, 150 mm c/c, 45 deg",
                                                                    "metre",  257.55, ""),

    # ── 2.36  Extra levelling ─────────────────────────────────────────────────
    ("2.36",   "Extra for levelling & neatly dressing disposed soil",
                                                                     "cum",   76.70, "EXTRA"),

    # ── 2.37  Fly ash supply ──────────────────────────────────────────────────
    ("2.37",   "Supply & stacking of fly ash (IRC-58) at site incl. carriage (any lead)",
                                                                     "cum",  234.05, ""),

    # ── 2.38  Fly ash + Earth filling ────────────────────────────────────────
    ("2.38",   "Filling with fly ash & earth in trenches/embankment (layers ≤15 cm, compacted, "
               "intermediate earth layers, top/side earth ≥30 cm)",
                                                                     "cum",  234.05, ""),
]

# ─────────────────────────────────────────────────────────────────────────────
# Builder
# ─────────────────────────────────────────────────────────────────────────────
def build_sheet(wb):
    # Remove existing sheet if present
    if SHEET_NAME in wb.sheetnames:
        del wb[SHEET_NAME]

    # Insert after '02_Earth_Work'
    ref_idx = wb.sheetnames.index("02_Earth_Work") if "02_Earth_Work" in wb.sheetnames else 0
    ws = wb.create_sheet(SHEET_NAME, ref_idx + 1)

    # ── Column widths ────────────────────────────────────────────────────────
    # A=item code, B=Description, C=Unit, D=DAR Rate, E=Project Qty, F=Amount
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 70
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 14
    ws.column_dimensions["F"].width = 16

    # ── Row 1 – workbook banner ──────────────────────────────────────────────
    ws.merge_cells("A1:F1")
    _set(ws, 1, 1,
         "CPWD DAR 2019 Vol-1  |  Sub-Head 2.0 Earth Work  |  Schedule of Rates — Support Sheet",
         font=_font(bold=True, size=12, color="FFFFFF"),
         fill=_fill(C_HEAD1),
         align=_align("center"))
    ws.row_dimensions[1].height = 22

    # ── Row 2 – reference line ───────────────────────────────────────────────
    ws.merge_cells("A2:F2")
    _set(ws, 2, 1,
         "Reference: CPWD DAR 2019, Vol-1, Pages 94–101  |  Basic Rates as on 01.04.2021  |  "
         "Rates inclusive of GST on Works Contract @ 12%, 15% CP & OH and 1% Cess",
         font=_font(italic=True, size=9, color="1F4E79"),
         fill=_fill("EBF3FB"),
         align=_align("center"))
    ws.row_dimensions[2].height = 16

    row = 3  # current write row

    # ═════════════════════════════════════════════════════════════════════════
    # SECTION A — BASIC RATES
    # ═════════════════════════════════════════════════════════════════════════
    # Section header
    ws.merge_cells(f"A{row}:F{row}")
    _set(ws, row, 1, "SECTION A — BASIC RATES (Sub-Head 2.0 Earth Work, Page 95)",
         font=_font(bold=True, size=11, color="FFFFFF"),
         fill=_fill(C_HEAD1),
         align=_align("center"))
    ws.row_dimensions[row].height = 20
    row += 1

    # Column headers
    headers_br = ["Code No.", "Description", "Unit", "Basic Rate (Rs.)", "", ""]
    for ci, h in enumerate(headers_br, 1):
        _set(ws, row, ci, h,
             font=_font(bold=True, color="FFFFFF"),
             fill=_fill(C_HEAD2),
             align=_align("center"),
             border=_border())
    ws.row_dimensions[row].height = 16
    row += 1

    br_start = row
    for code, desc, unit, rate in BASIC_RATES:
        bg = C_BASIC_BG if (row - br_start) % 2 == 0 else "FFFFFF"
        _set(ws, row, 1, code,  font=_font(), fill=_fill(bg), align=_align("center"), border=_border())
        _set(ws, row, 2, desc,  font=_font(), fill=_fill(bg), align=_align("left", wrap=True), border=_border())
        _set(ws, row, 3, unit,  font=_font(), fill=_fill(bg), align=_align("center"), border=_border())
        _set(ws, row, 4, rate,  font=_font(), fill=_fill(bg), align=_align("right"), border=_border(),
             num_format='#,##0.00')
        ws.merge_cells(f"E{row}:F{row}")
        row += 1
    br_end = row - 1

    # Spacer
    row += 1

    # ═════════════════════════════════════════════════════════════════════════
    # SECTION B — SCHEDULE OF RATES
    # ═════════════════════════════════════════════════════════════════════════
    ws.merge_cells(f"A{row}:F{row}")
    _set(ws, row, 1,
         "SECTION B — SCHEDULE OF RATES  |  Earth Work Items 2.1 to 2.38  (Pages 96–101)",
         font=_font(bold=True, size=11, color="FFFFFF"),
         fill=_fill(C_HEAD1),
         align=_align("center"))
    ws.row_dimensions[row].height = 20
    row += 1

    # Instructions row
    ws.merge_cells(f"A{row}:F{row}")
    _set(ws, row, 1,
         "Enter Project Quantity in column E  →  Amount (column F) = DAR Rate x Project Qty   "
         "|  Yellow = Deduct items   |  Green = Extra / Add-on items",
         font=_font(italic=True, size=9, color="1F4E79"),
         fill=_fill("EBF3FB"),
         align=_align("left", wrap=True))
    ws.row_dimensions[row].height = 14
    row += 1

    # Column headers
    col_hdrs = ["Item Code", "Description", "Unit", "DAR Rate (Rs.)",
                "Project Qty", "Amount (Rs.)"]
    for ci, h in enumerate(col_hdrs, 1):
        _set(ws, row, ci, h,
             font=_font(bold=True, color="FFFFFF"),
             fill=_fill(C_HEAD2),
             align=_align("center", wrap=True),
             border=_border())
    ws.row_dimensions[row].height = 18
    row += 1

    # Track amount cells grouped by category for SUM
    cat_groups = {}     # cat_label -> [f_cell_addr, ...]
    current_cat = "Misc"
    item_rows = []      # list of (row, flag) for rate items

    for entry in EW_ITEMS:
        code, desc, unit, rate, flag = entry

        if code == "CAT":
            # Category header row
            ws.merge_cells(f"A{row}:F{row}")
            _set(ws, row, 1, desc,
                 font=_font(bold=True, size=10, color="1F4E79"),
                 fill=_fill(C_CAT_BG),
                 align=_align("left", wrap=True))
            ws.row_dimensions[row].height = 18
            current_cat = desc
            cat_groups[current_cat] = []
            row += 1
            continue

        # Determine row background
        if flag == "DEDUCT":
            bg = C_DEDUCT
        elif flag == "EXTRA":
            bg = C_EXTRA
        else:
            bg = C_ALTROW if (row % 2 == 0) else "FFFFFF"

        _set(ws, row, 1, code, font=_font(), fill=_fill(bg),
             align=_align("center"), border=_border())
        _set(ws, row, 2, desc, font=_font(), fill=_fill(bg),
             align=_align("left", wrap=True), border=_border())
        _set(ws, row, 3, unit, font=_font(), fill=_fill(bg),
             align=_align("center"), border=_border())

        if rate is not None:
            _set(ws, row, 4, rate, font=_font(), fill=_fill(bg),
                 align=_align("right"), border=_border(), num_format='#,##0.00')
            # E: user input quantity (blank, no formula)
            e_cell = ws.cell(row=row, column=5)
            e_cell.fill = _fill("FFFDE7")   # light amber – user input
            e_cell.border = _border()
            e_cell.alignment = _align("right")
            e_cell.number_format = '#,##0.000'
            # F: Amount = D * E
            f_addr = f"F{row}"
            f_cell = ws.cell(row=row, column=6)
            f_cell.value = f"=D{row}*E{row}"
            f_cell.font  = _font()
            f_cell.fill  = _fill(bg)
            f_cell.border = _border()
            f_cell.alignment = _align("right")
            f_cell.number_format = '#,##0.00'
            if current_cat in cat_groups:
                cat_groups[current_cat].append(f"F{row}")
        else:
            # Percentage items (2.24.1, 2.24.2) – show note instead
            pct_note = "20% add-on" if "20%" in desc else "25% add-on"
            _set(ws, row, 4, pct_note, font=_font(italic=True, color="7F7F7F"),
                 fill=_fill(bg), align=_align("center"), border=_border())
            _set(ws, row, 5, "—", font=_font(color="7F7F7F"),
                 fill=_fill(bg), align=_align("center"), border=_border())
            _set(ws, row, 6, "—", font=_font(color="7F7F7F"),
                 fill=_fill(bg), align=_align("center"), border=_border())

        item_rows.append((row, flag))
        row += 1

    # Spacer
    row += 1

    # ═════════════════════════════════════════════════════════════════════════
    # SECTION C — SUMMARY TOTALS
    # ═════════════════════════════════════════════════════════════════════════
    ws.merge_cells(f"A{row}:F{row}")
    _set(ws, row, 1, "SECTION C — PROJECT AMOUNT SUMMARY",
         font=_font(bold=True, size=11, color="FFFFFF"),
         fill=_fill(C_HEAD1),
         align=_align("center"))
    ws.row_dimensions[row].height = 20
    row += 1

    # Column headers for summary
    for ci, h in enumerate(["Category", "", "", "", "Sub-Total Amount (Rs.)", ""], 1):
        _set(ws, row, ci, h,
             font=_font(bold=True, color="FFFFFF"),
             fill=_fill(C_HEAD2),
             align=_align("center"),
             border=_border())
    ws.row_dimensions[row].height = 16
    row += 1

    grand_cells = []
    for cat_label, f_cells in cat_groups.items():
        if not f_cells:
            continue
        ws.merge_cells(f"A{row}:D{row}")
        _set(ws, row, 1, cat_label,
             font=_font(), fill=_fill(C_TOTAL),
             align=_align("left", wrap=True), border=_border())
        ws.merge_cells(f"E{row}:F{row}")
        sum_formula = "=SUM(" + ",".join(f_cells) + ")"
        _set(ws, row, 5, sum_formula,
             font=_font(bold=True), fill=_fill(C_TOTAL),
             align=_align("right"), border=_border(), num_format='#,##0.00')
        grand_cells.append(f"E{row}")
        row += 1

    # Grand total
    ws.merge_cells(f"A{row}:D{row}")
    _set(ws, row, 1, "GRAND TOTAL — Sub-Head 2.0 Earth Work",
         font=_font(bold=True, size=11), fill=_fill(C_HEAD3),
         align=_align("left"), border=_border())
    ws.merge_cells(f"E{row}:F{row}")
    if grand_cells:
        grand_formula = "=SUM(" + ",".join(grand_cells) + ")"
    else:
        grand_formula = 0
    _set(ws, row, 5, grand_formula,
         font=_font(bold=True, size=11), fill=_fill(C_HEAD3),
         align=_align("right"), border=_border(), num_format='#,##0.00')
    ws.row_dimensions[row].height = 20
    row += 1

    # ── Freeze panes ─────────────────────────────────────────────────────────
    ws.freeze_panes = "A5"

    print(f"  Sheet '{SHEET_NAME}' written — {row-1} rows used.")
    return ws


# ─────────────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────────────
def main():
    print(f"Loading workbook: {WB_PATH}")
    wb = openpyxl.load_workbook(WB_PATH)
    print(f"  Current sheets: {wb.sheetnames}")

    build_sheet(wb)

    import shutil, os
    tmp = WB_PATH.replace(".xlsx", "_REBUILDING.xlsx")
    wb.save(tmp)
    # Try atomic replace; fall back to copy if main file is locked
    try:
        os.replace(tmp, WB_PATH)
        print(f"Saved (replaced): {WB_PATH}")
    except PermissionError:
        print(f"[WARN] Main file locked by Excel. Saved as: {tmp}")
        print("  Close the workbook in Excel, then rename/copy manually.")
    print(f"  Sheets now: {wb.sheetnames}")


if __name__ == "__main__":
    main()
