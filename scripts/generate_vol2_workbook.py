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
WB2_FILE = 'CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_2.xltx'

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
    print('  Registered 5 WB1->WB2 factor name mirrors '
          '(Factor_Water / GST / CPOH / Cess / Sundries)')


def build_vol2_local_factors(wb, styles):
    """
    Build Vol2_Local_Factors — the ONLY local infra sheet in WB2.

    Contains sub-head-specific allowances that are NOT in WB1's Global_Factors:
      - 18_Water_Supply: 30% fittings & wastage on pipe cost (DAR std, p.488)
      - 19_Drainage: 10% breakage allowance on pipe quantity (DAR std, p.756)
      - 26_New_Technologies: 7% wastage on sheet/board products (DAR guidance)
      - 17_Sanitary: 5% jointing & breakage allowance (DAR std)

    The builder sheets reference these as named ranges (Vol2_WS_Fittings etc.)
    so users can see and override them in one place.

    All WB1 global factors (Water/GST/CPOH/Cess/Sundries) live in WB1 ONLY
    and are NOT duplicated here.  Factor mirrors are WB2 named ranges that
    resolve to WB1 Global_Factors cells.
    """
    ws = wb.create_sheet(title='Vol2_Local_Factors')
    ws.views.sheetView[0].showGridLines = True

    # Title
    ws.merge_cells('A1:I1')
    t = ws.cell(row=1, column=1)
    t.value = ('VOL. 2 LOCAL FACTORS  |  Sub-head-specific allowances NOT present in WB1\n'
               'DO NOT duplicate WB1 Global_Factors here — edit those in WB1 only.')
    t.font      = styles['font_title']
    t.fill      = styles['fill_title']
    t.alignment = styles['align_center']
    ws.row_dimensions[1].height = 32

    # Architecture note
    ws.merge_cells('A2:I3')
    n = ws.cell(row=2, column=1)
    n.value = ('PRINCIPLE: WB2 contains NO duplicate of WB1 infrastructure.  '
               'Global factors (Water 1%, GST 14.05%, CPOH 15%, Cess 1%, Sundries multiplier 2.00) '
               'are read from WB1 Global_Factors via the workbook-level named ranges '
               'Factor_Water / Factor_GST / Factor_CPOH / Factor_Cess / Factor_Sundries defined '
               'in this workbook.  The table below holds ONLY allowances that are unique to '
               'specific Vol.2 sub-heads and have no equivalent in WB1.')
    n.font      = styles['font_note']
    n.fill      = styles['fill_note']
    n.alignment = styles['align_wrap']
    n.border    = styles['border_thin']
    ws.row_dimensions[2].height = 52

    # Column headers
    headers = ['Named Range', 'Sub-head', 'Allowance Description',
               'DAR Reference', 'Current Value', 'Unit / Basis',
               'How to override', 'CPWD Guidance', 'Status']
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=4, column=col, value=h)
        c.font      = styles['font_header']
        c.fill      = styles['fill_header']
        c.alignment = styles['align_center']
        c.border    = styles['border_header']
    ws.row_dimensions[4].height = 22

    # Data rows — sub-head-specific factors
    local_factors = [
        ('Vol2_WS_Fittings', '18_Water_Supply',
         'Fittings & wastage on pipe cost (P)',
         'DAR 2019 p.488, Note below item 18.1.1',
         0.30, '% of pipe material sub-total',
         'Change col E value; re-run generator to embed in formulas',
         '30 x P / 100 — standard CPWD allowance for pipe fittings, joints, '
         'clamps and cutting/jointing wastage. Applies to ALL pipe materials '
         'in this sub-head unless item text specifies otherwise.',
         'ACTIVE'),
        ('Vol2_DR_Breakage', '19_Drainage',
         'Breakage allowance on SW/RCC pipe quantity',
         'DAR 2019 p.756, Note below item 19.1.1',
         0.10, '% of pipe quantity (lengths)',
         'Change col E value; re-run generator',
         '10% of pipe quantity added for SW pipes; 5% for RCC/CI pipes. '
         'Applied to BOTH the pipe quantity and its carriage line.',
         'ACTIVE'),
        ('Vol2_NT_Wastage', '26_New_Technologies',
         'Wastage on sheet/board products (laminate, ACP, EPS)',
         'DAR 2019 guidance note, p.980',
         0.07, '% of product area (sqm)',
         'Change col E value; re-run generator',
         '7% of net area ordered to allow for cutting, trimming and diagonal '
         'laying. Reduce to 5% for simple rectangular rooms; increase to 10% '
         'for complex floor plans.',
         'ACTIVE'),
        ('Vol2_SI_Jointing', '17_Sanitary_Installations',
         'Jointing & breakage allowance on sanitary ware',
         'DAR 2019 standard practice, p.374',
         0.05, '% of fixture cost',
         'Change col E value; re-run generator',
         '5% allowance for cement, sand, and minor breakage risk during '
         'unloading and installation of vitreous china fittings.',
         'ACTIVE'),
    ]

    for row_idx, (name, sh, desc, ref, val, basis, override, guidance, status) in \
            enumerate(local_factors, 5):
        ws.cell(row=row_idx, column=1, value=name).font  = styles['font_bold']
        ws.cell(row=row_idx, column=1).alignment          = styles['align_center']
        ws.cell(row=row_idx, column=2, value=sh).font    = styles['font_note']
        ws.cell(row=row_idx, column=2).alignment          = styles['align_center']
        ws.cell(row=row_idx, column=3, value=desc).alignment = styles['align_left']
        ws.cell(row=row_idx, column=4, value=ref).font   = styles['font_note']
        ws.cell(row=row_idx, column=4).alignment          = styles['align_wrap']
        cv = ws.cell(row=row_idx, column=5, value=val)
        cv.number_format = '0.00%'
        cv.font          = styles['font_bold']
        cv.fill          = styles['fill_input']
        cv.alignment     = styles['align_center']
        ws.cell(row=row_idx, column=6, value=basis).font = styles['font_note']
        ws.cell(row=row_idx, column=6).alignment          = styles['align_center']
        ws.cell(row=row_idx, column=7, value=override).font = styles['font_note']
        ws.cell(row=row_idx, column=7).alignment          = styles['align_wrap']
        ws.cell(row=row_idx, column=8, value=guidance).font = styles['font_note']
        ws.cell(row=row_idx, column=8).alignment          = styles['align_wrap']
        ws.cell(row=row_idx, column=9, value=status).font = styles['font_bold']
        ws.cell(row=row_idx, column=9).alignment          = styles['align_center']
        ws.row_dimensions[row_idx].height = 46
        from openpyxl.styles import Border, Side
        thin = styles['border_thin']
        for c in range(1, 10):
            ws.cell(row=row_idx, column=c).border = thin

    # WB1 factor mirror documentation (read-only reference, not editable here)
    ws.merge_cells('A10:I10')
    sep = ws.cell(row=10, column=1)
    sep.value = ('─── WB1 GLOBAL FACTORS (read-only reference — edit in WB1 only; '
                 'WB2 named ranges automatically pull the live values) ───')
    sep.font      = styles['font_header']
    sep.fill      = styles['fill_header']
    sep.alignment = styles['align_center']
    ws.row_dimensions[10].height = 20

    wb1_refs = [
        ('Factor_Water',    'All sub-heads', 'Water charges for curing & site use', 'Global_Factors!F12', 0.01,    '% of (W-A)'),
        ('Factor_GST',      'All sub-heads', 'GST on Works Contract',               'Global_Factors!F13', 0.1405,  '% of (X-A)'),
        ('Factor_CPOH',     'All sub-heads', 'Contractor Profit & Overheads',       'Global_Factors!F14', 0.15,    '% of (Y-A)'),
        ('Factor_Cess',     'All sub-heads', 'Labour Welfare Cess (BOCW)',           'Global_Factors!F15', 0.01,    '% of (Z-A)'),
        ('Factor_Sundries', 'All sub-heads', 'Sundries Cost Index Multiplier',      'Global_Factors!F16', 2.00,    'multiplier on L.S. base'),
    ]
    for row_idx, (name, sh, desc, src, val, basis) in enumerate(wb1_refs, 11):
        ws.cell(row=row_idx, column=1, value=name).font  = styles['font_bold']
        ws.cell(row=row_idx, column=1).alignment          = styles['align_center']
        ws.cell(row=row_idx, column=2, value=sh).font    = styles['font_note']
        ws.cell(row=row_idx, column=2).alignment          = styles['align_center']
        ws.cell(row=row_idx, column=3, value=desc).alignment = styles['align_left']
        ws.cell(row=row_idx, column=4, value=f"WB1 {src} (external link)").font = styles['font_note']
        ws.cell(row=row_idx, column=4).alignment          = styles['align_center']
        cv = ws.cell(row=row_idx, column=5, value=val)
        cv.number_format = '0.00%' if val < 1 else '0.00'
        cv.font          = styles['font_regular']
        cv.fill          = styles['fill_subtotal']
        cv.alignment     = styles['align_center']
        ws.cell(row=row_idx, column=6, value=basis).font = styles['font_note']
        ws.cell(row=row_idx, column=6).alignment          = styles['align_center']
        ws.cell(row=row_idx, column=7,
                value='Edit in WB1 Global_Factors — do NOT change here').font = styles['font_note']
        ws.cell(row=row_idx, column=7).alignment = styles['align_center']
        ws.cell(row=row_idx, column=9,
                value='WB1 only').fill = styles['fill_note']
        ws.cell(row=row_idx, column=9).alignment = styles['align_center']
        ws.cell(row=row_idx, column=9).font = styles['font_note']
        ws.row_dimensions[row_idx].height = 22
        for c in range(1, 10):
            ws.cell(row=row_idx, column=c).border = styles['border_thin']

    # Column widths
    for col, w in {'A': 22, 'B': 22, 'C': 38, 'D': 28, 'E': 14,
                   'F': 20, 'G': 32, 'H': 46, 'I': 14}.items():
        ws.column_dimensions[col].width = w

    ws.freeze_panes = 'A5'
    print('  Built Vol2_Local_Factors sheet')


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
        'Template File  |  Open → Work → Save As <Project Name>.xlsx  |  Do NOT rename this file'
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
        '★  TEMPLATE — THIS FILE OPENS FRESH EACH TIME  ★\n'
        '   Nothing you type here is saved to the template. To keep your work:\n'
        '   File → Save As → give it a project/item name.\n\n'
        '1. BOTH WORKBOOKS MUST BE OPEN IN THE SAME FOLDER.\n'
        f'   Vol. 1 must be saved as: {WB1_FILE}\n'
        '   (Open the Vol. 1 template, then File → Save As with that exact filename.)\n'
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

    # 2. Vol2 local factors (sub-head-specific rates NOT in WB1)
    print('--- Building Vol2_Local_Factors sheet ---')
    build_vol2_local_factors(wb, styles)

    # 3. Register WB1 factor mirrors as WB2 named ranges
    print('--- Registering WB1 factor mirrors ---')
    _define_wb2_factor_names(wb)

    # 4. Remove openpyxl default sheet
    if default_sheet in wb.worksheets:
        wb.remove(default_sheet)

    # 5. Build 14 trade-builder sheets in Vol. 2 order
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
        alt_path = WB2_FILE.replace('.xltx', '_Latest.xltx')
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
