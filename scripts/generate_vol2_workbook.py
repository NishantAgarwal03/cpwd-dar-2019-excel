# -*- coding: utf-8 -*-
"""
Generate CPWD DAR 2019 Volume 2 Custom Rate Analysis Workbook.

Architecture:
  WB2 (this file's output) references WB1 for all Rates_Master data and
  Global_Factors.  Both workbooks must be in the same directory for Excel
  to resolve the external links automatically.

  WB1 filename: CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_1.xlsx
  WB2 filename: CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_2.xlsx

Named ranges defined in WB2 that mirror WB1 named ranges:
  Factor_Water     → WB1 Global_Factors!$F$12
  Factor_GST       → WB1 Global_Factors!$F$13
  Factor_CPOH      → WB1 Global_Factors!$F$14
  Factor_Cess      → WB1 Global_Factors!$F$15
  Factor_Sundries  → WB1 Global_Factors!$F$16

  These are defined as workbook-level defined names in WB2 using the external
  workbook link syntax so they resolve without any infra sheets in WB2.

14 trade-builder sheets are built in Vol. 2 order (sub-heads 13 – 26).
"""

import os
import sys
import time

sys.path.insert(0, os.path.abspath('.'))

import openpyxl
from openpyxl.workbook.defined_name import DefinedName

from scripts.styles import get_workbook_styles
from scripts.vol2_configs import get_vol2_configs
from scripts.trade_builder_vol2 import build_vol2_trade

# ---------------------------------------------------------------------------
# WB1 external link constants
# ---------------------------------------------------------------------------
WB1_FILE = 'CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_1.xlsx'
WB2_FILE = 'CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_2.xlsx'

# External reference string template — Excel uses [filename]Sheet!range
def _wb1_ext(sheet, cell):
    return f"'[{WB1_FILE}]{sheet}'!{cell}"


def _define_wb2_factor_names(wb):
    """
    Create workbook-level named ranges in WB2 pointing at WB1's Global_Factors
    sheet.  Excel resolves these when both files are open in the same folder.
    """
    factors = [
        ('Factor_Water',    _wb1_ext('Global_Factors', '$F$12')),
        ('Factor_GST',      _wb1_ext('Global_Factors', '$F$13')),
        ('Factor_CPOH',     _wb1_ext('Global_Factors', '$F$14')),
        ('Factor_Cess',     _wb1_ext('Global_Factors', '$F$15')),
        ('Factor_Sundries', _wb1_ext('Global_Factors', '$F$16')),
    ]
    for name, refers_to in factors:
        dn = DefinedName(name=name, attr_text=refers_to)
        wb.defined_names.add(dn)
    print('  Registered 5 WB1→WB2 factor name mirrors '
          '(Factor_Water / GST / CPOH / Cess / Sundries)')


def build_vol2_cover(wb, styles):
    """
    Build a compact Cover / Instructions sheet as the first sheet of WB2.
    Explains the two-workbook architecture and the external-link dependency.
    """
    ws = wb.create_sheet(title='Vol_2_Cover')
    ws.views.sheetView[0].showGridLines = False

    # Title
    ws.merge_cells('A1:I1')
    c = ws.cell(row=1, column=1)
    c.value = (
        'CPWD DAR 2019  |  CUSTOM RATE ANALYSIS WORKBOOK — VOLUME 2 (Sub-Heads 13–26)\n'
        'Companion to Vol. 1  |  Do NOT rename or move either workbook file'
    )
    c.font      = styles['font_title']
    c.fill      = styles['fill_title']
    c.alignment = styles['align_center']
    ws.row_dimensions[1].height = 40

    # Architecture note
    ws.merge_cells('A3:I10')
    arch = ws.cell(row=3, column=1)
    arch.value = (
        'HOW TO USE THIS WORKBOOK\n\n'
        '1. BOTH WORKBOOKS MUST BE OPEN IN THE SAME FOLDER.\n'
        f'   Vol. 1: {WB1_FILE}\n'
        f'   Vol. 2: {WB2_FILE} (this file)\n\n'
        '2. EXTERNAL LINKS: All rate lookups in this workbook (description, unit, basic rate)\n'
        '   fetch live data from Vol. 1 Rates_Master sheet. The formula in each Code cell reads:\n'
        "       =IF(B28=\"\",\"\", IFERROR(INDEX('[WB1]Rates_Master'!$C:$C, MATCH(B28,'[WB1]Rates_Master'!$A:$A,0)), \"Custom\"))\n"
        '   Excel resolves this automatically when both workbooks are open. If you see #REF! or\n'
        '   #VALUE! errors, open Vol. 1 and then press Ctrl+Alt+F9 to force recalculation.\n\n'
        '3. NAMED RANGES: Factor_Water, Factor_GST, Factor_CPOH, Factor_Cess, Factor_Sundries\n'
        '   are defined in this workbook and point to Vol. 1 Global_Factors cells.\n'
        '   To change a factor, edit it in Vol. 1 — do not override it here.\n\n'
        '4. A-TAG RULE (CPWD W-A convention): Any MATERIAL line whose rate you import from\n'
        '   Vol. 1 builder sheets (e.g. a mortar rate from 03_Mortars, or an RCC rate from\n'
        '   05_RCC_Work) already carries Water, GST, CPOH and Cess. Tag that line "A" in\n'
        '   column I. The markup chain automatically subtracts A-total from each markup base\n'
        '   so the import is not taxed twice.\n\n'
        '5. CROSS-VOLUME REFERENCES by sub-head:\n'
        '   13 Finishing      → Vol.1 03_Mortars (items 3.2, 3.4, 3.6, 3.8–3.12, 3.16)\n'
        '   14 Repairs        → Vol.1 03_Mortars (3.4, 3.9, 3.18); 09_Wood_and_PVC (8.23)\n'
        '   16 Road Work      → Vol.1 03_Mortars; 04_Concrete; 05_RCC_Work (various)\n'
        '   20 Pile Work      → Vol.1 05_RCC_Work (item 5.33.1)\n'
        '   22 Water Proofing → Vol.1 03_Mortars (3.8, 3.9, 3.10)\n'
        '   24 Heritage       → Vol.1 03_Mortars (3.19)\n'
        '   All other sheets  → self-contained (Rates_Master codes only)\n\n'
        '6. MACHINERY SHEETS: Sub-heads 15, 16, 20, 23, 26 have an extra machinery block\n'
        '   (Section 4b) between the Labour subtotal and the Sundries line.\n'
        '   Plant codes 0001-0083 in Rates_Master include operator, fuel and lubricants\n'
        '   for one 8-hour shift per CPWD Note 1.'
    )
    arch.font      = styles['font_note']
    arch.fill      = styles['fill_note']
    arch.alignment = styles['align_wrap']
    arch.border    = styles['border_thin']
    ws.row_dimensions[3].height = 380

    for col, w in {'A': 18, 'B': 16, 'C': 46, 'D': 14, 'E': 16,
                   'F': 16, 'G': 18, 'H': 42, 'I': 14}.items():
        ws.column_dimensions[col].width = w

    print('  Built Vol_2_Cover sheet')


def generate_vol2_workbook():
    t0 = time.time()

    print('=' * 70)
    print('CPWD DAR 2019 VOLUME 2 CUSTOM RATE ANALYSIS WORKBOOK GENERATOR')
    print('Sub-Heads 13 – 26  |  Companion to WB1 (Vol. 1)')
    print('=' * 70)

    wb = openpyxl.Workbook()
    default_sheet = wb.active

    styles  = get_workbook_styles()
    configs = get_vol2_configs()

    # 1. Cover sheet
    print('--- Building Cover / Architecture sheet ---')
    build_vol2_cover(wb, styles)

    # 2. Register WB1 factor mirrors as WB2 named ranges
    print('--- Registering WB1 factor mirrors ---')
    _define_wb2_factor_names(wb)

    # 3. Remove openpyxl default sheet
    if default_sheet in wb.worksheets:
        wb.remove(default_sheet)

    # 4. Build 14 trade-builder sheets in Vol. 2 order
    print('--- Generating 14 Vol. 2 Trade Builder Sheets ---')
    sheet_order = [
        '13_Finishing',
        '14_Repairs_to_Buildings',
        '15_Dismantling_Demolishing',
        '16_Road_Work',
        '17_Sanitary_Installations',
        '18_Water_Supply',
        '19_Drainage',
        '20_Pile_Work',
        '21_Aluminium_Work',
        '22_Water_Proofing',
        '23_Rain_Water_Harvesting',
        '24_Heritage_Buildings',
        '25_Structural_Glazing',
        '26_New_Technologies',
    ]

    for key in sheet_order:
        cfg = configs[key]
        build_vol2_trade(wb, cfg, styles)

    # 5. Save
    out_path = WB2_FILE
    print(f'Saving workbook to {out_path} ...')
    try:
        wb.save(out_path)
    except PermissionError:
        alt_path = WB2_FILE.replace('.xlsx', '_Latest.xlsx')
        wb.save(alt_path)
        print('=' * 70)
        print('BUILD DID NOT UPDATE THE MAIN WORKBOOK')
        print(f"  '{out_path}' is open in Excel and could not be overwritten.")
        print(f"  The new build was written to '{alt_path}' instead.")
        print('  Close Excel, delete the ~$ lock file, re-run this script.')
        print('=' * 70)
        elapsed = time.time() - t0
        size    = os.path.getsize(alt_path) / (1024 * 1024)
        print(f'Wrote {len(wb.sheetnames)} sheets in {elapsed:.2f}s ({size:.2f} MB)')
        raise SystemExit(1)

    elapsed       = time.time() - t0
    file_size_mb  = os.path.getsize(out_path) / (1024 * 1024)
    print(f'SUCCESS! Generated {len(wb.sheetnames)} sheets in '
          f'{elapsed:.2f}s ({file_size_mb:.2f} MB)')
    print(f'Sheet names: {wb.sheetnames}')
    print('=' * 70)
    print('REMINDER: open both WB1 and WB2 in the same folder in Excel to')
    print('          resolve all external Rates_Master and Global_Factors links.')
    print('=' * 70)


if __name__ == '__main__':
    generate_vol2_workbook()
