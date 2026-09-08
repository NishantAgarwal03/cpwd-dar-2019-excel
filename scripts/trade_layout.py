# -*- coding: utf-8 -*-
"""
Shared row map and panel renderers for the trade builder sheets.

Every builder sheet now opens with the same two-panel input analysis that was
first developed on 01_Carriage_of_Materials:

  PANEL 1 - SCOPE & SPECIFICATION INPUTS
      One row per clause of the item nomenclature. Each row states, explicitly,
      whether that clause actually moves cost and whether the DAR 2019 data
      supports it:

        DRIVES COST         a listed choice changes a code, coefficient or
                            imported rate on this sheet, and the DAR prints
                            different figures for the different choices.
        NOMENCLATURE ONLY   the clause changes the wording of the item but the
                            DAR prices one rate for every variant, so there is
                            no differential to apply.
        USER-SUPPLIED COST  the clause has a real cost but the DAR gives no
                            figure for it; the user must enter an amount.

  PANEL 2 - OPERATIONAL & DERIVATION DRIVERS
      The quantitative levers (wastage %, thickness, lead, plant productivity)
      that scale the coefficients rather than describe the scope.

Below the panels the anatomy is unchanged apart from column I, which carries
the CPWD "A" tag: a material line whose rate already includes statutory
markups is tagged A and excluded from every markup base (W-A, X-A, Y-A, Z-A).
"""

# --- Row map ---------------------------------------------------------------
R_TITLE = 1
R_LEGEND = 2
R_GUIDE = 3

R_P1_HEAD = 5
R_P1_COLS = 6
R_P1_FIRST = 7
P1_ROWS = 6
R_P1_LAST = R_P1_FIRST + P1_ROWS - 1          # 12

R_META = 13                                    # item code / basis / unit / sub-head
R_NOMEN = 14                                   # assembled nomenclature

R_P2_HEAD = 15
R_P2_COLS = 16
R_P2_FIRST = 17
P2_ROWS = 4
R_P2_LAST = R_P2_FIRST + P2_ROWS - 1          # 20
R_KEY = R_P2_LAST + 1                          # 21 - cost-impact key strip

R_AUDIT = 22

R_MAT_HEAD = 23
R_MAT_COLS = 24
R_MAT_FIRST = 25
MAT_ROWS = 10
R_MAT_LAST = R_MAT_FIRST + MAT_ROWS - 1       # 34
R_MAT_SUB = 35
R_A_TOTAL = 36

R_LAB_HEAD = 37
R_LAB_COLS = 38
R_LAB_FIRST = 39
LAB_ROWS = 10
R_LAB_LAST = R_LAB_FIRST + LAB_ROWS - 1       # 48
R_LAB_SUB = 49

R_SUN_HEAD = 51
R_SUN = 52

R_MU_HEAD = 54
R_MU_COLS = 55
R_W = 56
R_X1 = 57
R_X = 58
R_Y1 = 59
R_Y = 60
R_Z1 = 61
R_Z = 62
R_Z2 = 63
R_COST = 64
R_RATE = 65
R_SAY = 66

R_LIB_HEAD = 69
R_LIB_NOTE = 70
R_LIB_COLS = 71
R_LIB_FIRST = 72
LIB_ROWS = 8
R_LIB_LAST = R_LIB_FIRST + LIB_ROWS - 1       # 79

LAST_COL = 9  # column I

IMPACT_FILLS = {
    'DRIVES COST': 'fill_result',
    'NOMENCLATURE ONLY': 'fill_subtotal',
    'USER-SUPPLIED COST': 'fill_input',
}

COST_IMPACT_KEY = (
    'COST-IMPACT KEY - "DRIVES COST" = the DAR 2019 prints different figures for the listed choices, and '
    'changing the selection changes a code, coefficient or imported rate on this sheet. "NOMENCLATURE ONLY" = '
    'the clause changes the wording of the item but the DAR prices one rate for every variant, so there is no '
    'differential data to apply. "USER-SUPPLIED COST" = the clause has a real cost but the DAR gives no figure '
    'for it, so enter the amount yourself (normally as a Sundries / L.S. line).'
)


def section_bar(ws, row, text, styles, last_col=LAST_COL, height=24):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=last_col)
    c = ws.cell(row=row, column=1)
    c.value = text
    c.font = styles['font_white_bold']
    c.fill = styles['fill_header']
    c.alignment = styles['align_left']
    ws.row_dimensions[row].height = height
    return c


def col_headers(ws, row, headers, styles, height=24):
    for ci, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=ci, value=h)
        c.font = styles['font_header']
        c.fill = styles['fill_header']
        c.alignment = styles['align_center']
        c.border = styles['border_header']
    ws.row_dimensions[row].height = height


def border_row(ws, row, styles, last_col=LAST_COL):
    for c in range(1, last_col + 1):
        ws.cell(row=row, column=c).border = styles['border_thin']


def build_scope_panel(ws, config, styles, dv_registry):
    """PANEL 1 - one row per nomenclature clause, with its cost-impact class."""
    section_bar(ws, R_P1_HEAD, '1. PANEL 1 - SCOPE & SPECIFICATION INPUTS  '
                               '(what is being built, and which clauses actually move cost)', styles)

    col_headers(ws, R_P1_COLS,
                ['Nomenclature Clause / Dimension', 'Selected Value', 'Cost Impact',
                 'Drives', 'What changes when this selection changes, and the DAR 2019 evidence for it',
                 '', '', '', ''], styles, height=30)
    ws.merge_cells(start_row=R_P1_COLS, start_column=5, end_row=R_P1_COLS, end_column=LAST_COL)

    dims = config.get('scope_inputs', [])
    for i in range(P1_ROWS):
        r = R_P1_FIRST + i
        d = dims[i] if i < len(dims) else None

        if d is None:
            ws.cell(row=r, column=1, value='(spare - add a further scope clause here)').font = styles['font_note']
            ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=LAST_COL)
            border_row(ws, r, styles)
            ws.row_dimensions[r].height = 20
            continue

        ws.cell(row=r, column=1, value=d['label']).font = styles['font_bold']
        ws.cell(row=r, column=1).alignment = styles['align_wrap']

        cv = ws.cell(row=r, column=2, value=d['default'])
        cv.fill = styles['fill_input']
        cv.font = styles['font_bold']
        cv.alignment = styles['align_wrap']
        if d.get('options'):
            dv_registry.append((d['options'], cv))

        ci = ws.cell(row=r, column=3, value=d['impact'])
        ci.alignment = styles['align_center']
        ci.font = styles['font_bold']
        ci.fill = styles[IMPACT_FILLS.get(d['impact'], 'fill_calc')]

        cd = ws.cell(row=r, column=4, value=d.get('drives', '-'))
        cd.alignment = styles['align_center']
        cd.font = styles['font_note']

        ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=LAST_COL)
        cn = ws.cell(row=r, column=5, value=d['evidence'])
        cn.font = styles['font_note']
        cn.alignment = styles['align_wrap']

        border_row(ws, r, styles)
        ws.row_dimensions[r].height = 32


def build_meta_and_nomenclature(ws, config, styles):
    """Item code / batch basis / output unit / sub-head, then the assembled name."""
    pairs = [
        ('A', 'Custom Item Code:', 'B', config['default_item_code'], None),
        ('C', 'Batch Output Basis:', 'D', config['default_basis_qty'], '0.00'),
        ('E', 'Output Unit:', 'F', config['default_basis_unit'], None),
        ('G', 'Trade Sub-Head:', 'H', config['sheet_name'].replace('_', ' '), None),
    ]
    for lab_col, lab, val_col, val, fmt in pairs:
        lc = ws[f'{lab_col}{R_META}']
        lc.value = lab
        lc.font = styles['font_bold']
        vc = ws[f'{val_col}{R_META}']
        vc.value = val
        vc.font = styles['font_bold']
        vc.alignment = styles['align_center']
        vc.fill = styles['fill_input'] if val_col != 'H' else styles['fill_lookup']
        if fmt:
            vc.number_format = fmt
    ws[f'I{R_META}'] = 'Output Unit Note:'
    ws[f'I{R_META}'].font = styles['font_note']
    border_row(ws, R_META, styles)
    ws.row_dimensions[R_META].height = 22

    # Assembled nomenclature, with a manual override that wins when filled.
    ws[f'A{R_NOMEN}'] = 'Assembled Item Nomenclature:'
    ws[f'A{R_NOMEN}'].font = styles['font_bold']
    ws.merge_cells(start_row=R_NOMEN, start_column=2, end_row=R_NOMEN, end_column=7)
    nc = ws.cell(row=R_NOMEN, column=2)
    nc.value = config.get('nomenclature_formula') or config.get('default_item_desc', '')
    nc.font = styles['font_regular']
    nc.fill = styles['fill_lookup']
    nc.alignment = styles['align_wrap']

    ws[f'H{R_NOMEN}'] = 'Manual Override:'
    ws[f'H{R_NOMEN}'].font = styles['font_note']
    ov = ws[f'I{R_NOMEN}']
    ov.value = None
    ov.fill = styles['fill_input']
    ov.alignment = styles['align_wrap']
    border_row(ws, R_NOMEN, styles)
    ws.row_dimensions[R_NOMEN].height = 40


def build_driver_panel(ws, config, styles, dv_registry):
    """PANEL 2 - the quantitative levers that scale coefficients."""
    section_bar(ws, R_P2_HEAD, '2. PANEL 2 - OPERATIONAL & DERIVATION DRIVERS  '
                               '(quantities and allowances that scale the build-up below)', styles)
    col_headers(ws, R_P2_COLS,
                ['Driver', 'Value', 'Unit', 'Applies To',
                 'How this driver is used, and the DAR 2019 basis for the default', '', '', '', ''],
                styles, height=26)
    ws.merge_cells(start_row=R_P2_COLS, start_column=5, end_row=R_P2_COLS, end_column=LAST_COL)

    drivers = config.get('driver_inputs', [])
    for i in range(P2_ROWS):
        r = R_P2_FIRST + i
        d = drivers[i] if i < len(drivers) else None

        if d is None:
            ws.cell(row=r, column=1, value='(spare - add a further driver here)').font = styles['font_note']
            ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=LAST_COL)
            border_row(ws, r, styles)
            ws.row_dimensions[r].height = 20
            continue

        ws.cell(row=r, column=1, value=d['label']).font = styles['font_bold']
        ws.cell(row=r, column=1).alignment = styles['align_wrap']

        cv = ws.cell(row=r, column=2, value=d['default'])
        cv.fill = styles['fill_input']
        cv.font = styles['font_bold']
        cv.alignment = styles['align_center']
        if d.get('number_format'):
            cv.number_format = d['number_format']
        if d.get('options'):
            dv_registry.append((d['options'], cv))

        ws.cell(row=r, column=3, value=d.get('unit', '')).alignment = styles['align_center']
        ws.cell(row=r, column=4, value=d.get('applies_to', '-')).alignment = styles['align_center']
        ws.cell(row=r, column=4).font = styles['font_note']

        ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=LAST_COL)
        cn = ws.cell(row=r, column=5, value=d['note'])
        cn.font = styles['font_note']
        cn.alignment = styles['align_wrap']

        border_row(ws, r, styles)
        ws.row_dimensions[r].height = 30

    # Cost-impact key sits directly under Panel 2 so both panels read together.
    ws.merge_cells(start_row=R_KEY, start_column=1, end_row=R_KEY, end_column=LAST_COL)
    k = ws.cell(row=R_KEY, column=1)
    k.value = COST_IMPACT_KEY
    k.font = styles['font_note']
    k.fill = styles['fill_note']
    k.alignment = styles['align_wrap']
    border_row(ws, R_KEY, styles)
    ws.row_dimensions[R_KEY].height = 44
