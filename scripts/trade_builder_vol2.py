# -*- coding: utf-8 -*-
"""
Builder for WB2 (Vol. 2) trade sheets.

Differences from trade_builder.py (WB1):
  - No infra sheets in WB2; all Rates_Master and Global_Factors references
    are prefixed with the WB1 external workbook name:
       '[CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_1.xlsx]SheetName'
  - Named ranges in WB2 (Factor_Water, Factor_GST, etc.) resolve to WB1 via
    external links — see generate_vol2_workbook.py for the definitions.
  - Code-description/unit/rate LOOKUP formulas reference WB1 Rates_Master via
    the external-link string WB1_RATES_MASTER_EXT.
  - Optional MACHINERY block: when config['has_machinery'] is True an extra
    section is inserted between LABOUR and SUNDRIES using the same row anatomy
    as the LABOUR block but headed "MACHINERY & PLANT HIRE".
  - No Master_Codes data-validation (Excel cannot use an external-workbook
    reference in a DV formula); the code cells still receive a descriptive
    prompt so users know what to type.
  - Wastage-% helper row (for sub-head 18 Water Supply 30% fittings rule and
    sub-head 19 Drainage 10% breakage) is emitted when
    config['driver_inputs'] contains a 'Fittings & wastage %' entry.
"""

import openpyxl
from openpyxl.styles import Protection
from openpyxl.worksheet.datavalidation import DataValidation

from scripts.trade_layout import (
    R_TITLE, R_LEGEND, R_GUIDE,
    R_P1_HEAD, R_P1_FIRST, R_P1_LAST,
    R_META, R_NOMEN,
    R_P2_FIRST, R_P2_LAST, R_KEY,
    R_AUDIT,
    R_MAT_HEAD, R_MAT_COLS, R_MAT_FIRST, R_MAT_LAST, R_MAT_SUB, R_A_TOTAL,
    R_LAB_HEAD, R_LAB_COLS, R_LAB_FIRST, R_LAB_LAST, R_LAB_SUB,
    R_SUN_HEAD, R_SUN,
    R_MU_HEAD, R_MU_COLS, R_W, R_X1, R_X, R_Y1, R_Y, R_Z1, R_Z, R_Z2,
    R_COST, R_RATE, R_SAY,
    R_SAY_NOTE,
    R_LIB_HEAD, R_LIB_NOTE, R_LIB_COLS, R_LIB_FIRST, R_LIB_LAST,
    LAST_COL, SAY_RULE_NOTE,
    section_bar, col_headers, border_row, build_guide,
    build_scope_panel, build_meta_and_nomenclature, build_driver_panel,
)

# ---------------------------------------------------------------------------
# WB1 external reference strings
# ---------------------------------------------------------------------------
WB1_FILE = 'CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_1.xlsx'

def _ext(sheet, cell_or_range):
    """Return an Excel external reference formula fragment (no leading =)."""
    return f"'[{WB1_FILE}]{sheet}'!{cell_or_range}"

# Full lookup path for Rates_Master columns
RM_CODES_COL  = _ext('Rates_Master', '$A:$A')
RM_DESC_COL   = _ext('Rates_Master', '$C:$C')
RM_UNIT_COL   = _ext('Rates_Master', '$D:$D')
RM_RATE_COL   = _ext('Rates_Master', '$E:$E')

# Named ranges are defined in WB2 pointing at WB1 (see generate_vol2_workbook.py)
# and can be used as formula references normally, e.g. =Factor_Sundries.

# ---------------------------------------------------------------------------
# Column header specs (identical to WB1)
# ---------------------------------------------------------------------------
TBL_HEADERS = [
    'Line', 'Code / Source', 'Description / Specification', 'Unit',
    'Quantity / Coeff', 'Basic Rate (Rs)', 'Amount (Rs)',
    'Source Reference / Remarks', 'Markup Tag',
]

MU_HEADERS = [
    'Step', 'Component / Markup Description', 'Apply? (YES/NO)',
    'Basis Applied', 'Base Amount (Rs)', 'Factor / %', 'Amount (Rs)',
    'CPWD Statutory Rule & Guidance Note', 'A-Rule',
]

LIB_HEADERS = [
    'Item Code', 'Item Nomenclature / Specification', 'Unit', 'Output Basis',
    'Direct Cost W (Rs)', 'Markups Applied', 'Unit Rate (Rs)', 'Say Rate (Rs)',
    'A-Total (Rs)',
]


# ---------------------------------------------------------------------------
# Row helper: one resource line (material, labour, or machinery)
# ---------------------------------------------------------------------------
def _resource_row(ws, r, idx, spec, styles, kind):
    """Emit one MATERIAL, LABOUR or MACHINERY line using WB1 external refs."""
    fallback = {
        'material': 'Custom / Intermediate Material',
        'labour':   'Custom Labour',
        'machinery':'Custom Plant / Machinery',
    }[kind]

    ws.cell(row=r, column=1, value=idx).alignment = styles['align_center']

    # Col B — code (INPUT)
    c_code = ws.cell(row=r, column=2, value=(spec['code'] if spec else ''))
    c_code.alignment = styles['align_center']
    c_code.font     = styles['font_bold']
    c_code.fill     = styles['fill_input']
    # No DV list (external workbook refs not allowed in Excel DV formulas);
    # the cell's note explains this to the user.

    # Col C — description (LOOKUP via WB1 external ref)
    c_desc = ws.cell(row=r, column=3)
    if spec and spec.get('custom_desc'):
        c_desc.value = spec['custom_desc']
    else:
        c_desc.value = (
            f'=IF(B{r}="","", IFERROR(INDEX({RM_DESC_COL}, '
            f'MATCH(B{r}, {RM_CODES_COL}, 0)), "{fallback}"))'
        )
    c_desc.alignment = styles['align_left']
    c_desc.fill      = styles['fill_lookup']

    # Col D — unit (LOOKUP)
    c_unit = ws.cell(row=r, column=4)
    if spec and spec.get('custom_unit'):
        c_unit.value = spec['custom_unit']
    else:
        c_unit.value = (
            f'=IF(B{r}="","", IFERROR(INDEX({RM_UNIT_COL}, '
            f'MATCH(B{r}, {RM_CODES_COL}, 0)), ""))'
        )
    c_unit.alignment = styles['align_center']
    c_unit.fill      = styles['fill_lookup']

    # Col E — quantity / coefficient (INPUT)
    c_qty = ws.cell(row=r, column=5, value=(spec['coeff'] if spec else None))
    c_qty.alignment    = styles['align_right']
    c_qty.number_format = styles['fmt_qty']
    c_qty.fill         = styles['fill_input']

    # Col F — basic rate (LOOKUP from WB1)
    c_rate = ws.cell(row=r, column=6)
    if spec and spec.get('custom_rate_formula'):
        c_rate.value = spec['custom_rate_formula']
    elif spec and spec.get('custom_rate') is not None:
        c_rate.value = spec['custom_rate']
    else:
        c_rate.value = (
            f'=IF(B{r}="","", IFERROR(INDEX({RM_RATE_COL}, '
            f'MATCH(B{r}, {RM_CODES_COL}, 0)), 0))'
        )
    c_rate.alignment    = styles['align_right']
    c_rate.number_format = styles['fmt_currency']
    c_rate.fill         = styles['fill_lookup']

    # Col G — amount (DERIVED)
    c_amt = ws.cell(row=r, column=7,
                    value=f'=IF(OR(B{r}="", E{r}=""), 0, ROUND(E{r} * F{r}, 2))')
    c_amt.alignment    = styles['align_right']
    c_amt.number_format = styles['fmt_currency']
    c_amt.fill         = styles['fill_calc']

    # Col H — remarks / note
    c_rem = ws.cell(row=r, column=8, value=(spec.get('note', spec.get('evidence', '')) if spec else ''))
    c_rem.alignment = styles['align_wrap']
    c_rem.font      = styles['font_note']

    # Col I — A-tag (INPUT for materials only)
    c_tag = ws.cell(row=r, column=9,
                    value=('A' if (spec and spec.get('a_tag')) else ''))
    c_tag.alignment = styles['align_center']
    c_tag.font      = styles['font_bold']
    c_tag.fill      = (styles['fill_say'] if (spec and spec.get('a_tag'))
                       else styles['fill_input'])

    border_row(ws, r, styles)
    ws.row_dimensions[r].height = 20


# ---------------------------------------------------------------------------
# Machinery block constants (appended after LABOUR, before SUNDRIES)
# We use a fixed 6-row machinery block (same MAT_ROWS count = 10 for labour).
# Rows are placed immediately after R_LAB_SUB.
# ---------------------------------------------------------------------------
MACH_ROWS = 6          # up to 6 plant lines


def _build_machinery_block(ws, config, styles, r_mach_head, dv_atag):
    """
    Emit a MACHINERY & PLANT HIRE block starting at r_mach_head.
    Returns the row number of the machinery subtotal row.
    """
    section_bar(
        ws, r_mach_head,
        '4b. MACHINERY & PLANT HIRE  '
        '(plant codes 0001-0083 include operator + fuel + lubricants per CPWD Note 1; '
        'specialist rigs priced at hire-invoice rate — enter as custom_rate)',
        styles,
    )
    col_headers(ws, r_mach_head + 1, TBL_HEADERS, styles)

    r_mach_first = r_mach_head + 2
    r_mach_last  = r_mach_first + MACH_ROWS - 1

    mach_specs = config.get('machinery_defaults', [])
    for i in range(MACH_ROWS):
        r = r_mach_first + i
        spec = mach_specs[i] if i < len(mach_specs) else None
        _resource_row(ws, r, i + 1, spec, styles, 'machinery')
        # Machinery lines do not accept A-tag; lock column 9
        ws.cell(row=r, column=9).fill = styles['fill_calc']

    r_mach_sub = r_mach_last + 1
    ws.merge_cells(start_row=r_mach_sub, start_column=1,
                   end_row=r_mach_sub,   end_column=6)
    ws.cell(row=r_mach_sub, column=1,
            value='Subtotal Machinery & Plant (Rs):').font = styles['font_bold']
    ws.cell(row=r_mach_sub, column=1).alignment = styles['align_right']
    cm = ws.cell(row=r_mach_sub, column=7,
                 value=f'=SUM(G{r_mach_first}:G{r_mach_last})')
    cm.font          = styles['font_bold']
    cm.alignment     = styles['align_right']
    cm.number_format = styles['fmt_currency']
    for c in range(1, LAST_COL + 1):
        ws.cell(row=r_mach_sub, column=c).fill   = styles['fill_subtotal']
        ws.cell(row=r_mach_sub, column=c).border = styles['border_double_bottom']
    ws.row_dimensions[r_mach_sub].height = 22

    return r_mach_sub


# ---------------------------------------------------------------------------
# Main builder
# ---------------------------------------------------------------------------
def build_vol2_trade(wb, config, styles):
    """Build one Vol-2 trade-builder sheet and attach it to wb."""
    ws = wb.create_sheet(title=config['sheet_name'])
    ws.views.sheetView[0].showGridLines = True

    has_machinery = config.get('has_machinery', False)

    # -----------------------------------------------------------------------
    # Header, legend, guide
    # -----------------------------------------------------------------------
    ws.merge_cells(start_row=R_TITLE, start_column=1,
                   end_row=R_TITLE, end_column=LAST_COL)
    c_t = ws.cell(row=R_TITLE, column=1)
    c_t.value     = config['trade_title']
    c_t.font      = styles['font_title']
    c_t.fill      = styles['fill_title']
    c_t.alignment = styles['align_center']
    ws.row_dimensions[R_TITLE].height = 30

    ws.cell(row=R_LEGEND, column=1, value='LEGEND:').font = styles['font_legend']
    ws.cell(row=R_LEGEND, column=1).alignment = styles['align_center']
    legend_items = [
        (2, 'INPUT - you type it',      styles['fill_input'],    styles['font_bold']),
        (3, 'OVERRIDE - optional',      styles['fill_override'], styles['font_bold']),
        (4, 'LOOKUP - fetched (WB1)',   styles['fill_lookup'],   styles['font_regular']),
        (5, 'DERIVED - calculated',     styles['fill_calc'],     styles['font_regular']),
        (6, 'RESULT - subtotal',        styles['fill_result'],   styles['font_bold']),
        (7, 'SAY - final rate',         styles['fill_say'],      styles['font_bold']),
        (8, 'CPWD rule / note',         styles['fill_note'],     styles['font_note']),
        (9, 'Not editable',             styles['fill_subtotal'], styles['font_regular']),
    ]
    for col, text, fill, font in legend_items:
        c = ws.cell(row=R_LEGEND, column=col, value=text)
        c.fill      = fill
        c.font      = font
        c.border    = styles['border_thin']
        c.alignment = styles['align_center']
    ws.row_dimensions[R_LEGEND].height = 20

    ws.merge_cells(start_row=R_GUIDE, start_column=1,
                   end_row=R_GUIDE, end_column=LAST_COL)
    c_g = ws.cell(row=R_GUIDE, column=1)
    c_g.value     = config['trade_guidance']
    c_g.font      = styles['font_note']
    c_g.fill      = styles['fill_note']
    c_g.alignment = styles['align_wrap']
    c_g.border    = styles['border_thin']
    ws.row_dimensions[R_GUIDE].height = 30

    # -----------------------------------------------------------------------
    # Data validations (those that DO NOT need external-workbook formulas)
    # -----------------------------------------------------------------------
    dv_yesno = DataValidation(type='list', formula1='"YES,NO"', allow_blank=False)
    ws.add_data_validation(dv_yesno)

    dv_atag = DataValidation(
        type='list', formula1='"A,"', allow_blank=True,
        prompt=('Enter A only if this line imports a rate that already includes '
                'Water, GST, CPOH and Cess. A-tagged lines are excluded from every '
                'markup base per the CPWD (W-A) convention.'),
        promptTitle='CPWD A-tag',
    )
    ws.add_data_validation(dv_atag)

    # Code cells: no list DV (external workbook refs not allowed in DV formulas).
    # Instead, add an input prompt so the user knows what is expected.
    dv_code_prompt = DataValidation(
        type='whole', operator='greaterThan', formula1='0',
        allow_blank=True,
        showErrorMessage=False,
        prompt=('Type a Rates_Master code (numeric, e.g. 0123) from WB1, or enter '
                'a Vol.1 item reference such as 3.4 for Mortars. '
                'The LOOKUP columns fetch description, unit and rate from WB1 automatically '
                'when both workbooks are open in the same folder.'),
        promptTitle='Rates_Master code (WB1)',
    )
    ws.add_data_validation(dv_code_prompt)

    # -----------------------------------------------------------------------
    # Panels 1, 2 (scope inputs and driver inputs)
    # -----------------------------------------------------------------------
    dv_registry = []
    build_scope_panel(ws, config, styles, dv_registry)
    build_meta_and_nomenclature(ws, config, styles)
    build_driver_panel(ws, config, styles, dv_registry)

    for options, cell in dv_registry:
        formula = '"' + ','.join(options) + '"'
        if len(formula) <= 255:
            dv = DataValidation(type='list', formula1=formula,
                                allow_blank=True, showErrorMessage=False)
            ws.add_data_validation(dv)
            dv.add(cell)

    # -----------------------------------------------------------------------
    # Audit bar
    # -----------------------------------------------------------------------
    ws.cell(row=R_AUDIT, column=1, value='AUDIT STATUS:').font = styles['font_white_bold']
    ws.cell(row=R_AUDIT, column=1).fill      = styles['fill_header']
    ws.cell(row=R_AUDIT, column=1).alignment = styles['align_center']

    b = ws.cell(row=R_AUDIT, column=2,
                value=f'=IF(AND(D{R_AUDIT}="OK", F{R_AUDIT}="OK", H{R_AUDIT}="OK"), '
                      f'"[PASS] ALL CHECKS OK", "[ALERT] CHECKS FAILED")')
    b.font      = styles['font_result']
    b.fill      = styles['fill_result']
    b.alignment = styles['align_center']

    ws.cell(row=R_AUDIT, column=3, value='Output Qty Check:').font = styles['font_note']
    ws.cell(row=R_AUDIT, column=3).alignment = styles['align_right']
    ws.cell(row=R_AUDIT, column=4,
            value=f'=IF(D{R_META}>0, "OK", "ERR: Qty<=0")').alignment = styles['align_center']

    ws.cell(row=R_AUDIT, column=5, value='Direct Cost Check:').font = styles['font_note']
    ws.cell(row=R_AUDIT, column=5).alignment = styles['align_right']
    ws.cell(row=R_AUDIT, column=6,
            value=f'=IF(G{R_W}>0, "OK", "ERR: W<=0")').alignment = styles['align_center']

    ws.cell(row=R_AUDIT, column=7, value='Markup Flow Check:').font = styles['font_note']
    ws.cell(row=R_AUDIT, column=7).alignment = styles['align_right']
    ws.cell(row=R_AUDIT, column=8,
            value=f'=IF(AND(G{R_W}<=G{R_X}, G{R_X}<=G{R_Y}, G{R_Y}<=G{R_Z}, G{R_A_TOTAL}<=G{R_W}), '
                  f'"OK", "ERR: Markups Broken")').alignment = styles['align_center']

    ws.cell(row=R_AUDIT, column=9,
            value=f'=IF(G{R_A_TOTAL}>0, "A-rule active", "No A imports")')
    ws.cell(row=R_AUDIT, column=9).alignment = styles['align_center']
    ws.cell(row=R_AUDIT, column=9).font      = styles['font_note']

    for c in (4, 6, 8):
        ws.cell(row=R_AUDIT, column=c).font = styles['font_bold']
    border_row(ws, R_AUDIT, styles)
    ws.row_dimensions[R_AUDIT].height = 22

    # -----------------------------------------------------------------------
    # Section 3: MATERIAL block
    # -----------------------------------------------------------------------
    section_bar(
        ws, R_MAT_HEAD,
        config.get('material_section_title',
                   '3. MATERIAL COMPONENT BUILD-UP  (tag column I with "A" for any line whose rate '
                   'already includes statutory markups — see CPWD (W-A) convention)'),
        styles,
    )
    col_headers(ws, R_MAT_COLS, TBL_HEADERS, styles)

    materials = config.get('default_materials', [])
    for i in range(R_MAT_LAST - R_MAT_FIRST + 1):
        r = R_MAT_FIRST + i
        spec = materials[i] if i < len(materials) else None
        _resource_row(ws, r, i + 1, spec, styles, 'material')
        dv_atag.add(ws.cell(row=r, column=9))
        dv_code_prompt.add(ws.cell(row=r, column=2))

    # Material subtotal
    ws.merge_cells(start_row=R_MAT_SUB, start_column=1,
                   end_row=R_MAT_SUB, end_column=6)
    ws.cell(row=R_MAT_SUB, column=1,
            value='Subtotal Materials (Rs):').font = styles['font_bold']
    ws.cell(row=R_MAT_SUB, column=1).alignment = styles['align_right']
    cm = ws.cell(row=R_MAT_SUB, column=7,
                 value=f'=SUM(G{R_MAT_FIRST}:G{R_MAT_LAST})')
    cm.font          = styles['font_bold']
    cm.alignment     = styles['align_right']
    cm.number_format = styles['fmt_currency']
    for c in range(1, LAST_COL + 1):
        ws.cell(row=R_MAT_SUB, column=c).fill   = styles['fill_subtotal']
        ws.cell(row=R_MAT_SUB, column=c).border = styles['border_double_bottom']
    ws.row_dimensions[R_MAT_SUB].height = 22

    # A-total row
    ws.merge_cells(start_row=R_A_TOTAL, start_column=1,
                   end_row=R_A_TOTAL, end_column=6)
    ws.cell(row=R_A_TOTAL, column=1,
            value='of which "A" - imported lines already carrying statutory markups:').font = styles['font_bold']
    ws.cell(row=R_A_TOTAL, column=1).alignment = styles['align_right']
    ca = ws.cell(row=R_A_TOTAL, column=7,
                 value=f'=ROUND(SUMIF($I${R_MAT_FIRST}:$I${R_MAT_LAST}, "A", '
                       f'$G${R_MAT_FIRST}:$G${R_MAT_LAST}), 2)')
    ca.font          = styles['font_bold']
    ca.alignment     = styles['align_right']
    ca.number_format = styles['fmt_currency']
    ws.cell(row=R_A_TOTAL, column=8,
            value=('CPWD (W-A) convention: this amount is included in W and every subtotal, '
                   'but is subtracted from each markup base so an already-marked-up import '
                   'is not marked up twice. For Vol.2 items importing WB1 mortar or RCC rates.')).font = styles['font_note']
    ws.cell(row=R_A_TOTAL, column=8).alignment = styles['align_wrap']
    for c in range(1, LAST_COL + 1):
        ws.cell(row=R_A_TOTAL, column=c).fill = styles['fill_subtotal']
    border_row(ws, R_A_TOTAL, styles)
    ws.row_dimensions[R_A_TOTAL].height = 30

    # -----------------------------------------------------------------------
    # Section 4: LABOUR block
    # -----------------------------------------------------------------------
    section_bar(
        ws, R_LAB_HEAD,
        '4. LABOUR & MACHINERY COMPONENT BUILD-UP  '
        '(crew suggestions: see the work-type summary on Labour_Machinery_Productivity in WB1. '
        'Plant codes 0001-0083 already include operator, fuel and lubricants for an 8-hour shift '
        'per CPWD Note 1)',
        styles,
    )
    col_headers(ws, R_LAB_COLS, TBL_HEADERS, styles)

    labour = config.get('default_labour', [])
    for i in range(R_LAB_LAST - R_LAB_FIRST + 1):
        r = R_LAB_FIRST + i
        spec = labour[i] if i < len(labour) else None
        _resource_row(ws, r, i + 1, spec, styles, 'labour')
        ws.cell(row=r, column=9).fill = styles['fill_calc']
        dv_code_prompt.add(ws.cell(row=r, column=2))

    # Labour subtotal
    ws.merge_cells(start_row=R_LAB_SUB, start_column=1,
                   end_row=R_LAB_SUB, end_column=6)
    ws.cell(row=R_LAB_SUB, column=1,
            value='Subtotal Labour & Plant Hire (Rs):').font = styles['font_bold']
    ws.cell(row=R_LAB_SUB, column=1).alignment = styles['align_right']
    cl = ws.cell(row=R_LAB_SUB, column=7,
                 value=f'=SUM(G{R_LAB_FIRST}:G{R_LAB_LAST})')
    cl.font          = styles['font_bold']
    cl.alignment     = styles['align_right']
    cl.number_format = styles['fmt_currency']
    for c in range(1, LAST_COL + 1):
        ws.cell(row=R_LAB_SUB, column=c).fill   = styles['fill_subtotal']
        ws.cell(row=R_LAB_SUB, column=c).border = styles['border_double_bottom']
    ws.row_dimensions[R_LAB_SUB].height = 22

    # -----------------------------------------------------------------------
    # Optional: Section 4b MACHINERY block
    # -----------------------------------------------------------------------
    r_mach_sub = None
    if has_machinery:
        r_mach_head  = R_LAB_SUB + 1
        r_mach_sub   = _build_machinery_block(ws, config, styles, r_mach_head, dv_atag)
        r_sun_head   = r_mach_sub + 1
        r_sun        = r_sun_head + 1
    else:
        r_sun_head   = R_LAB_SUB + 1
        r_sun        = r_sun_head + 1

    # -----------------------------------------------------------------------
    # Section 5: SUNDRIES
    # -----------------------------------------------------------------------
    section_bar(
        ws, r_sun_head,
        '5. SUNDRIES & LUMP-SUM ALLOWANCES  '
        '(manual entry — the DAR gives no reproducible formula; '
        'use Sundries_Reference in WB1 to find a comparable item)',
        styles,
    )

    # Build the W formula to include machinery if present
    if has_machinery and r_mach_sub is not None:
        w_formula = f'=G{R_MAT_SUB} + G{R_LAB_SUB} + G{r_mach_sub} + G{r_sun}'
    else:
        w_formula = f'=G{R_MAT_SUB} + G{R_LAB_SUB} + G{r_sun}'

    ws.cell(row=r_sun, column=1, value='1').alignment = styles['align_center']
    ws.cell(row=r_sun, column=2, value='9999').alignment = styles['align_center']
    ws.cell(row=r_sun, column=2).font = styles['font_bold']
    ws.cell(row=r_sun, column=3,
            value='Sundries (guided by Sundries_Reference sheet in WB1)')
    ws.cell(row=r_sun, column=4, value='L.S.').alignment = styles['align_center']
    cs = ws.cell(row=r_sun, column=5,
                 value=config.get('default_sundries_base', 0.0))
    cs.alignment    = styles['align_right']
    cs.number_format = '0.00'
    cs.fill         = styles['fill_input']
    cf = ws.cell(row=r_sun, column=6, value='=IFERROR(Factor_Sundries, 2.00)')
    cf.alignment    = styles['align_right']
    cf.number_format = '0.00'
    cf.fill         = styles['fill_lookup']
    cg = ws.cell(row=r_sun, column=7, value=f'=ROUND(E{r_sun} * F{r_sun}, 2)')
    cg.alignment    = styles['align_right']
    cg.number_format = styles['fmt_currency']
    cg.font         = styles['font_bold']
    ws.cell(row=r_sun, column=8,
            value=('Base L.S. allowance × Cost Index Multiplier (from WB1 Global_Factors '
                   'via Factor_Sundries named range)')).font = styles['font_note']
    border_row(ws, r_sun, styles)
    ws.row_dimensions[r_sun].height = 22

    # -----------------------------------------------------------------------
    # Section 6: Markup chain
    # (rows must follow immediately after sundries; the row numbers below
    # are the standard layout constants from trade_layout.py which embed the
    # assumption that sundries is at R_SUN.  For machinery sheets those
    # constants are still used — we just make W include the machinery total.)
    # -----------------------------------------------------------------------
    r_mu_head = r_sun + 2
    r_mu_cols = r_mu_head + 1

    # We re-use the standard layout row numbers for the markup rows if this is
    # a no-machinery sheet (they'll coincide with the defaults).  For machinery
    # sheets the actual rows differ; we compute them dynamically.
    if has_machinery and r_mach_sub is not None:
        # Offset everything from r_sun
        mu_offset = r_sun - R_SUN
        r_w_  = R_W  + mu_offset
        r_x1_ = R_X1 + mu_offset
        r_x_  = R_X  + mu_offset
        r_y1_ = R_Y1 + mu_offset
        r_y_  = R_Y  + mu_offset
        r_z1_ = R_Z1 + mu_offset
        r_z_  = R_Z  + mu_offset
        r_z2_ = R_Z2 + mu_offset
        r_cost_ = R_COST + mu_offset
        r_rate_ = R_RATE + mu_offset
        r_say_  = R_SAY  + mu_offset
        r_say_note_ = R_SAY_NOTE + mu_offset
        r_lib_head_ = R_LIB_HEAD + mu_offset
        r_lib_note_ = R_LIB_NOTE + mu_offset
        r_lib_cols_ = R_LIB_COLS + mu_offset
        r_lib_first_ = R_LIB_FIRST + mu_offset
        r_lib_last_  = R_LIB_LAST  + mu_offset
    else:
        r_w_ = R_W; r_x1_ = R_X1; r_x_ = R_X
        r_y1_ = R_Y1; r_y_ = R_Y; r_z1_ = R_Z1; r_z_ = R_Z; r_z2_ = R_Z2
        r_cost_ = R_COST; r_rate_ = R_RATE; r_say_ = R_SAY
        r_say_note_ = R_SAY_NOTE
        r_lib_head_ = R_LIB_HEAD; r_lib_note_ = R_LIB_NOTE
        r_lib_cols_ = R_LIB_COLS; r_lib_first_ = R_LIB_FIRST
        r_lib_last_  = R_LIB_LAST
        r_mu_head    = R_MU_HEAD
        r_mu_cols    = R_MU_COLS

    section_bar(ws, r_mu_head,
                '6. STATUTORY MARKUPS & FINAL RATE DERIVATION '
                '(CPWD DAR METHOD, WITH THE (W-A) EXCLUSION)',
                styles)
    col_headers(ws, r_mu_cols, MU_HEADERS, styles)

    toggles = config.get('toggles',
                         {'water': 'YES', 'gst': 'YES', 'cpoh': 'YES', 'cess': 'YES'})
    notes   = config.get('toggle_notes', {})

    markup_rows = [
        (r_w_,  'W',    'Base Direct Production Cost (Materials + Labour + Sundries)',
         None, 'Direct Sum', None, None, w_formula,
         'Total direct cost of production (W); A-tagged imports included, exactly as the book totals it.',
         'includes A'),
        (r_x1_, 'X1',   'Add Water Charges', toggles['water'], 'On (W - A)',
         f'=G{r_w_} - G{R_A_TOTAL}', '=IFERROR(Factor_Water, 0.01)',
         f'=IF(C{r_x1_}="YES", ROUND(E{r_x1_} * F{r_x1_}, 2), 0)',
         notes.get('water', 'CPWD Standard: 1% for curing & site water. Toggle NO if dry item.'),
         'net of A'),
        (r_x_,  'X',    'Subtotal "X" (W + Water Charges)',
         None, 'W + Water', None, None, f'=G{r_w_} + G{r_x1_}',
         'Compounded base for GST calculation.', 'includes A'),
        (r_y1_, 'Y1',   'Add GST on Works Contract', toggles['gst'], 'On (X - A)',
         f'=G{r_x_} - G{R_A_TOTAL}', '=IFERROR(Factor_GST, 0.1405)',
         f'=IF(C{r_y1_}="YES", ROUND(E{r_y1_} * F{r_y1_}, 2), 0)',
         notes.get('gst', 'CPWD DAR 2019 factor 0.1405 (works contract). Toggle NO if tax exempt.'),
         'net of A'),
        (r_y_,  'Y',    'Subtotal "Y" (X + GST)',
         None, 'X + GST', None, None, f'=G{r_x_} + G{r_y1_}',
         'Compounded base for Contractor Profit & Overheads.', 'includes A'),
        (r_z1_, 'Z1',   'Add Contractor Profit & Overheads (15% CPOH)', toggles['cpoh'], 'On (Y - A)',
         f'=G{r_y_} - G{R_A_TOTAL}', '=IFERROR(Factor_CPOH, 0.15)',
         f'=IF(C{r_z1_}="YES", ROUND(E{r_z1_} * F{r_z1_}, 2), 0)',
         notes.get('cpoh', 'Standard CPWD 15% allowance for site overheads, head-office costs & margin.'),
         'net of A'),
        (r_z_,  'Z',    'Subtotal "Z" (Y + CPOH)',
         None, 'Y + CPOH', None, None, f'=G{r_y_} + G{r_z1_}',
         'Compounded base for Labour Welfare Cess.', 'includes A'),
        (r_z2_, 'Z2',   'Add Labour Welfare Cess (1% BOCW Cess)', toggles['cess'], 'On (Z - A)',
         f'=G{r_z_} - G{R_A_TOTAL}', '=IFERROR(Factor_Cess, 0.01)',
         f'=IF(C{r_z2_}="YES", ROUND(E{r_z2_} * F{r_z2_}, 2), 0)',
         notes.get('cess', 'Statutory 1% Building and Other Construction Workers Welfare Cess.'),
         'net of A'),
        (r_cost_, 'Cost',
         f'="Total Cost of Batch (" & TEXT(D{R_META}, "0.00") & " " & F{R_META} & "):"',
         None, 'Z + Cess', None, None, f'=G{r_z_} + G{r_z2_}',
         'Total evaluated cost for the batch size in the Panel 1 header row.', '-'),
        (r_rate_, 'Rate',
         f'="Analyzed Unit Rate per 1.00 " & F{R_META} & ":"',
         None, 'Cost / Batch Qty', None, None, f'=ROUND(G{r_cost_} / D{R_META}, 2)',
         'Calculated cost per standard unit of measurement.', '-'),
        (r_say_,  'SAY',
         f'="OFFICIAL SAY RATE (Rs per " & F{R_META} & "):"',
         None, 'MROUND to Rs 0.05', None, None, f'=MROUND(G{r_rate_}, 0.05)',
         ('OFFICIAL CPWD ROUNDED RATE FOR TENDER SCHEDULES & ESTIMATES. '
          'CPWD quotes Say rates to the nearest 5 paise — see the note directly below this row.'),
         '-'),
    ]

    for r, step, desc, toggle, basis, base_f, factor_f, amount_f, note, arule in markup_rows:
        ws.cell(row=r, column=1, value=step).font = styles['font_bold']
        ws.cell(row=r, column=1).alignment = styles['align_center']
        ws.cell(row=r, column=2, value=desc).alignment = styles['align_left']
        ws.cell(row=r, column=3,
                value=(toggle if toggle else '-')).alignment = styles['align_center']
        ws.cell(row=r, column=4, value=basis).alignment = styles['align_center']

        cb = ws.cell(row=r, column=5, value=(base_f if base_f else '-'))
        cb.alignment = styles['align_right']
        if base_f:
            cb.number_format = styles['fmt_currency']

        cfa = ws.cell(row=r, column=6, value=(factor_f if factor_f else '-'))
        cfa.alignment = styles['align_right']
        if factor_f:
            cfa.number_format = styles['fmt_percent']

        cv = ws.cell(row=r, column=7, value=amount_f)
        cv.alignment    = styles['align_right']
        cv.number_format = styles['fmt_currency']
        cv.font         = styles['font_bold']

        cn = ws.cell(row=r, column=8, value=note)
        cn.alignment = styles['align_wrap']
        cn.font      = styles['font_note']

        ct = ws.cell(row=r, column=9, value=arule)
        ct.alignment = styles['align_center']
        ct.font      = styles['font_note']

        border_row(ws, r, styles)

        if r in (r_w_, r_x_, r_y_, r_z_):
            for c in range(1, LAST_COL + 1):
                ws.cell(row=r, column=c).fill = styles['fill_subtotal']
        elif r in (r_x1_, r_y1_, r_z1_, r_z2_):
            ws.cell(row=r, column=3).fill = styles['fill_input']
            ws.cell(row=r, column=3).font = styles['font_bold']
            dv_yesno.add(ws.cell(row=r, column=3))
        elif r == r_cost_:
            for c in range(1, LAST_COL + 1):
                ws.cell(row=r, column=c).fill = styles['fill_result']
        elif r == r_rate_:
            for c in range(1, LAST_COL + 1):
                ws.cell(row=r, column=c).fill = styles['fill_subtotal']
        elif r == r_say_:
            for c in range(1, LAST_COL + 1):
                ws.cell(row=r, column=c).fill = styles['fill_say']
            ws.cell(row=r, column=2).font = styles['font_say']
            ws.cell(row=r, column=7).font = styles['font_say']

        ws.row_dimensions[r].height = 30 if r == r_say_ else 24

    # Say-rate footnote
    ws.merge_cells(start_row=r_say_note_, start_column=1,
                   end_row=r_say_note_, end_column=LAST_COL)
    sn = ws.cell(row=r_say_note_, column=1)
    sn.value     = SAY_RULE_NOTE
    sn.font      = styles['font_note']
    sn.fill      = styles['fill_note']
    sn.alignment = styles['align_wrap']
    border_row(ws, r_say_note_, styles)
    ws.row_dimensions[r_say_note_].height = 44

    # -----------------------------------------------------------------------
    # Section 7: Running custom items library
    # -----------------------------------------------------------------------
    section_bar(ws, r_lib_head_,
                '7. RUNNING CUSTOM NON-DSR ITEMS LIBRARY FOR THIS TRADE', styles)
    ws.merge_cells(start_row=r_lib_note_, start_column=1,
                   end_row=r_lib_note_, end_column=LAST_COL)
    ci = ws.cell(row=r_lib_note_, column=1)
    ci.value = ('Log all completed custom rate analyses for this trade below. Items logged here '
                'maintain an internal project audit trail and can be referenced from other builder sheets.')
    ci.font      = styles['font_note']
    ci.fill      = styles['fill_note']
    ci.alignment = styles['align_wrap']
    ws.row_dimensions[r_lib_note_].height = 20

    col_headers(ws, r_lib_cols_, LIB_HEADERS, styles)

    sn_prefix = config['sheet_name'][:2]
    sample_lib = config.get('sample_library', [])
    for i in range(r_lib_last_ - r_lib_first_ + 1):
        r    = r_lib_first_ + i
        item = sample_lib[i] if i < len(sample_lib) else None
        ws.cell(row=r, column=1,
                value=(item['code'] if item else f'C-{sn_prefix}.{i+1:02d}')).alignment = styles['align_center']
        ws.cell(row=r, column=2,
                value=(item['desc'] if item else '(Available for new custom item)')).alignment = styles['align_left']
        ws.cell(row=r, column=3,
                value=(item['unit'] if item else config['default_basis_unit'])).alignment = styles['align_center']
        ws.cell(row=r, column=4,
                value=(item['basis'] if item else config['default_basis_qty'])).alignment = styles['align_center']
        cw = ws.cell(row=r, column=5, value=(item['w'] if item else None))
        cw.alignment    = styles['align_right']
        cw.number_format = styles['fmt_currency']
        ws.cell(row=r, column=6,
                value=(item['markups'] if item else 'Full W→X→Y→Z')).alignment = styles['align_center']
        cr2 = ws.cell(row=r, column=7, value=(item['rate'] if item else None))
        cr2.alignment    = styles['align_right']
        cr2.number_format = styles['fmt_currency']
        cy = ws.cell(row=r, column=8, value=(item['say'] if item else None))
        cy.alignment    = styles['align_right']
        cy.number_format = styles['fmt_currency']
        cy.font         = styles['font_bold']
        cat = ws.cell(row=r, column=9, value=(item.get('a_total') if item else None))
        cat.alignment    = styles['align_right']
        cat.number_format = styles['fmt_currency']

        row_fill = styles['fill_subtotal'] if r % 2 == 0 else styles['fill_calc']
        for c in range(1, LAST_COL + 1):
            cell = ws.cell(row=r, column=c)
            cell.fill   = row_fill
            cell.border = styles['border_thin']
            if c not in (1, 8):
                cell.font = styles['font_regular']
        ws.row_dimensions[r].height = 20

    # -----------------------------------------------------------------------
    # Column widths
    # -----------------------------------------------------------------------
    for col, w in {'A': 34, 'B': 26, 'C': 46, 'D': 16, 'E': 18,
                   'F': 18, 'G': 20, 'H': 44, 'I': 14}.items():
        ws.column_dimensions[col].width = w

    # -----------------------------------------------------------------------
    # Sheet protection (editable cells unlocked)
    # -----------------------------------------------------------------------
    for r in range(1, ws.max_row + 1):
        for c in range(1, LAST_COL + 1):
            ws.cell(row=r, column=c).protection = Protection(locked=True)

    unlocked = [
        f'B{R_META}', f'D{R_META}', f'F{R_META}', f'I{R_NOMEN}',
        f'E{r_sun}',
        f'C{r_x1_}', f'C{r_y1_}', f'C{r_z1_}', f'C{r_z2_}',
    ]
    for ref in unlocked:
        ws[ref].protection = Protection(locked=False)

    for r in range(R_P1_FIRST, R_P1_LAST + 1):
        ws.cell(row=r, column=2).protection = Protection(locked=False)
    for r in range(R_P2_FIRST, R_P2_LAST + 1):
        ws.cell(row=r, column=2).protection = Protection(locked=False)
    for r in range(R_MAT_FIRST, R_MAT_LAST + 1):
        for c in (2, 5, 8, 9):
            ws.cell(row=r, column=c).protection = Protection(locked=False)
    for r in range(R_LAB_FIRST, R_LAB_LAST + 1):
        for c in (2, 5, 8):
            ws.cell(row=r, column=c).protection = Protection(locked=False)
    for r in range(r_lib_first_, r_lib_last_ + 1):
        for c in range(1, LAST_COL + 1):
            ws.cell(row=r, column=c).protection = Protection(locked=False)

    if has_machinery and r_mach_sub is not None:
        r_mach_first = r_mach_sub - MACH_ROWS
        for r in range(r_mach_first, r_mach_first + MACH_ROWS):
            for c in (2, 5, 8):
                ws.cell(row=r, column=c).protection = Protection(locked=False)

    ws.protection.sheet = True
    ws.freeze_panes = 'A4'
    print(f'  Built Vol-2 sheet: {config["sheet_name"]}')
