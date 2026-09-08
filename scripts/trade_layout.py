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

R_GUIDE_HEAD = 4                               # 'HOW TO USE THIS SHEET' bar
R_GUIDE_STEPS = 5                              # the four working steps
R_GUIDE_ROLES = 6                              # what each cell colour means
R_P1_HEAD = 8
R_P1_COLS = 9
R_P1_FIRST = 10
P1_ROWS = 6
R_P1_LAST = R_P1_FIRST + P1_ROWS - 1          # 15

R_META = 16                                    # item code / basis / unit / sub-head
R_NOMEN = 17                                   # assembled nomenclature

R_P2_HEAD = 18
R_P2_COLS = 19
R_P2_FIRST = 20
P2_ROWS = 4
R_P2_LAST = R_P2_FIRST + P2_ROWS - 1          # 23
R_KEY = R_P2_LAST + 1                          # 24 - cost-impact key strip

R_AUDIT = 25

R_MAT_HEAD = 26
R_MAT_COLS = 27
R_MAT_FIRST = 28
MAT_ROWS = 10
R_MAT_LAST = R_MAT_FIRST + MAT_ROWS - 1       # 37
R_MAT_SUB = 38
R_A_TOTAL = 39

R_LAB_HEAD = 40
R_LAB_COLS = 41
R_LAB_FIRST = 42
LAB_ROWS = 10
R_LAB_LAST = R_LAB_FIRST + LAB_ROWS - 1       # 51
R_LAB_SUB = 52

R_SUN_HEAD = 54
R_SUN = 55

R_MU_HEAD = 57
R_MU_COLS = 58
R_W = 59
R_X1 = 60
R_X = 61
R_Y1 = 62
R_Y = 63
R_Z1 = 64
R_Z = 65
R_Z2 = 66
R_COST = 67
R_RATE = 68
R_SAY = 69
R_SAY_NOTE = 70                                # why MROUND(x, 0.05)

R_LIB_HEAD = 72
R_LIB_NOTE = 73
R_LIB_COLS = 74
R_LIB_FIRST = 75
LIB_ROWS = 8
R_LIB_LAST = R_LIB_FIRST + LIB_ROWS - 1       # 82

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
    ov.fill = styles['fill_override']
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


# ---------------------------------------------------------------------------
# Cell-role vocabulary. Every value on a builder sheet is exactly one of these,
# and the fill colour is the only thing you need to read to know which.
# ---------------------------------------------------------------------------
ROLE_LEGEND = [
    ('INPUT',    'fill_input',    'You type it. Yellow.'),
    ('OVERRIDE', 'fill_override', 'Optional. Blank = use the value above it. Orange.'),
    ('LOOKUP',   'fill_lookup',   'Fetched from a reference sheet. Blue. Do not type here.'),
    ('DERIVED',  'fill_calc',     'Calculated from the cells above. White. Do not type here.'),
    ('RESULT',   'fill_result',   'A subtotal you are meant to read. Green.'),
    ('SAY',      'fill_say',      'The final rate you quote. Gold.'),
]

GUIDE_STEPS = (
    'HOW TO USE THIS SHEET  -  '
    'STEP 1: fill in PANEL 1 to say WHAT is being built; each row tells you whether that choice '
    'actually changes the cost.   '
    'STEP 2: set the quantities in PANEL 2.   '
    'STEP 3: enter the MATERIAL and LABOUR lines - type a code in the yellow "Code / Source" cell '
    'and the description, unit and rate fill themselves in.   '
    'STEP 4: read the gold SAY rate at the bottom. Everything between Step 3 and the SAY rate is '
    'calculated for you - if a cell is not yellow or orange, do not type in it.'
)

GUIDE_ROLES = (
    'WHAT THE COLOURS MEAN  -  '
    'YELLOW = INPUT, you type it.   '
    'ORANGE = OVERRIDE, optional; leave it blank to use the looked-up value, or type here to force '
    'your own.   '
    'BLUE = LOOKUP, fetched from Rates_Master or another reference sheet.   '
    'WHITE = DERIVED, calculated from the cells above it.   '
    'GREEN = a subtotal to read.   '
    'GOLD = the final SAY rate.   '
    'The sheet is protected, so only the yellow and orange cells will let you type at all.'
)

SAY_RULE_NOTE = (
    'WHY THE SAY RATE IS ROUNDED TO 5 PAISE:  CPWD DAR 2019 quotes every "Say" rate to the nearest '
    'Rs 0.05, not to two decimals - e.g. item 1.3 computes 104.89 and is printed as 104.90; item '
    '1.4.1 computes 155.47 and is printed as 155.45; item 10.1 computes 86.04 and is printed as '
    '86.05. This sheet therefore uses MROUND(rate, 0.05) so the analysed rate matches the book '
    'exactly instead of sitting a paisa away from it. Change the 0.05 only if your circle office '
    'issues a different rounding rule.'
)


def build_guide(ws, styles, last_col=LAST_COL):
    """Render the 'how to use' bar and the cell-role colour key."""
    section_bar(ws, R_GUIDE_HEAD, 'HOW TO USE THIS SHEET  (read this before typing anything)',
                styles, last_col=last_col, height=22)

    ws.merge_cells(start_row=R_GUIDE_STEPS, start_column=1,
                   end_row=R_GUIDE_STEPS, end_column=last_col)
    c = ws.cell(row=R_GUIDE_STEPS, column=1)
    c.value = GUIDE_STEPS
    c.font = styles['font_note']
    c.fill = styles['fill_note']
    c.alignment = styles['align_wrap']
    border_row(ws, R_GUIDE_STEPS, styles, last_col)
    ws.row_dimensions[R_GUIDE_STEPS].height = 46

    ws.merge_cells(start_row=R_GUIDE_ROLES, start_column=1,
                   end_row=R_GUIDE_ROLES, end_column=last_col)
    c2 = ws.cell(row=R_GUIDE_ROLES, column=1)
    c2.value = GUIDE_ROLES
    c2.font = styles['font_note']
    c2.fill = styles['fill_note']
    c2.alignment = styles['align_wrap']
    border_row(ws, R_GUIDE_ROLES, styles, last_col)
    ws.row_dimensions[R_GUIDE_ROLES].height = 46


def role_cell(ws, row, col, role, styles, value=None, number_format=None,
              align='align_center'):
    """Write a cell and paint it with its role colour in one call."""
    fill_key = dict((r[0], r[1]) for r in ROLE_LEGEND)[role]
    c = ws.cell(row=row, column=col)
    if value is not None:
        c.value = value
    c.fill = styles[fill_key]
    c.font = styles['font_say'] if role == 'SAY' else styles['font_bold']
    c.alignment = styles[align]
    if number_format:
        c.number_format = number_format
    c.border = styles['border_thin']
    return c
