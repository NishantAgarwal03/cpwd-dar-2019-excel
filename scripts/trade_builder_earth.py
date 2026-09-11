# -*- coding: utf-8 -*-
"""
CPWD DAR 2019 Sub-Head 02: Earth Work (First-Principles Rate Analysis & Productivity Engine)

Architecture Highlights:
1. 100% Civil Engineering Parameter-Driven: Driven by Primary Work Scope (D6), Ground Strata (D7),
   Execution Method & Depth (D8), and Task Specification (D9). Completely eliminates CPWD item code numbering.
2. Architecture 3: Parametric Operational Breakdown across 5 core civil engineering levers
   (Loosening/Digging, Handling & Haulage, Spreading & Dressing, Watering to OMC, and Compaction).
3. Section 4: Table 4 First-Principles Benchmark Matrix with dedicated Resource Code column (Col C)
   and no CPWD item numbering.
4. Section 6: Table 5E Overhaul with 15 rich, well-thought civil engineering columns. Completely eliminates
   clutter columns: Key (A|N), Item Code, A Flag, and CPOH Base.
5. Live 5-step compounding statutory markups (Water 1%, GST 14.05%, CPOH 15%, Cess 1%) linked to Global_Factors.
"""

from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side, Protection
from scripts.trade_builder import add_trade_header_and_legend
from scripts.earthwork_tables import (
    SCOPE_OPTIONS,
    STRATA_OPTIONS,
    METHOD_OPTIONS,
    LIFT_DEPTH_OPTIONS,
    TASK_OPTIONS,
    CATEGORY_HEADINGS,
    HEADING_TASK_CATALOG,
    BENCHMARK_PRODUCTIVITY_ROWS,
    TABLE_5E_RECORDS,
    MASTER_ACTIVITIES,
    MASTER_ITEMS,
    # Decision-tree cascade data (L1→L2→L3→L4)
    L1_OPTIONS,
    L2_DATA,
    L3_DATA,
    L4_DATA,
)


# -----------------------------------------------------------------------------
# Row Coordinate Constants
# -----------------------------------------------------------------------------
# Section 1: Rate Analysis Configuration & Operational Profile
R_SEC1_HEAD = 4
R_SEC1_COLS = 5
R_SCOPE = 6         # D6: Primary Work Scope [INPUT]
R_STRATA = 7        # D7: Ground Strata / Material [INPUT]
R_METHOD = 8        # D8: Execution Method & Depth Condition [INPUT]
R_TASK = 9          # D9: Operation / Task Specification [INPUT]
R_BATCH_QTY = 10    # D10: Standard Output Batch Qty [LOOKUP]
R_BATCH_UNIT = 11   # D11: Standard Output Measurement Unit [LOOKUP]
R_CUSTOM_CODE = 12  # D12: Custom Project Rate Reference Code [INPUT]
R_NOMENCLATURE = 13 # D13: Official Work Specification / Nomenclature [LOOKUP]
R_DAR_PUB = 14      # D14: Published Benchmark Reference Say Rate [LOOKUP]
R_AUDIT = 15        # D15: Audit Verification Status [RESULT]

# Section 1B: Parametric Operational Decomposition (Architecture 3: 5 Work Levers)
R_SEC1B_HEAD = 17
R_SEC1B_COLS = 18
R_OP_LOOSENING = 19   # Lever 1: Method of Loosening or Digging
R_OP_HANDLING = 20    # Lever 2: Handling & Haulage
R_OP_SPREADING = 21   # Lever 3: Spreading & Dressing
R_OP_WATERING = 22    # Lever 4: Watering (Compaction / OMC)
R_OP_COMPACTION = 23  # Lever 5: Compaction Mechanism

# Section 2A: Base Specification Resource Build-Up (Table 2A: Gross Scope)
R_SEC2A_HEAD = 25
R_SEC2A_COLS = 26
R_RES_FIRST = 27
RES_ROWS = 10
R_RES_LAST = R_RES_FIRST + RES_ROWS - 1  # 36
R_W_GROSS = 37                          # Gross Direct Production Cost W_gross

# Section 2B: Scope Omissions, Credits & Contractual Deductions (Table 2B)
R_SEC2B_HEAD = 38
R_DED_FIRST = 39
DED_ROWS = 5
R_DED_LAST = R_DED_FIRST + DED_ROWS - 1  # 43
R_W_DEDUCT = 44                         # Subtotal Scope Deductions W_deduct

# Net Direct Production Cost & Tagged Totals
R_W_SUB = 45                            # Net Direct Production Cost W = W_gross + W_deduct
R_A_SUB = 46                            # A-tagged Composite Subtotal
R_BLANK_SEP = 47                        # Blank separation row

# Section 3: Statutory Markups
R_SEC3_HEAD = 48
R_SEC3_COLS = 49
R_W = 50
R_X1 = 51
R_X = 52
R_Y1 = 53
R_Y = 54
R_Z1 = 55
R_Z = 56
R_Z2 = 57
R_COST = 58
R_RATE = 59
R_SAY = 60
R_DAR_REF = 61
R_DAR_DIFF = 62

# Section 4: First-Principles Productivity Benchmark Matrix (Table 4)
R_SEC4_HEAD = 64
R_SEC4_BANNER = 65
R_SEC4_COLS = 66
R_BM_FIRST = 67
BM_ROWS = len(BENCHMARK_PRODUCTIVITY_ROWS)  # 26
R_BM_LAST = R_BM_FIRST + BM_ROWS - 1        # 92

# Section 5: Running Custom Library
R_SEC5_HEAD = 94
R_SEC5_BANNER = 95
R_SEC5_COLS = 96
R_LIB_FIRST = 97
LIB_ROWS = 8
R_LIB_LAST = R_LIB_FIRST + LIB_ROWS - 1     # 104

# Section 6: Table 5E Ground-Truth Norms Catalog (178 records)
R_SEC6_HEAD = 106
R_SEC6_BANNER = 107
R_SEC6_COLS = 108
R_T5E_FIRST = 109
T5E_ROWS = len(TABLE_5E_RECORDS)            # 178
R_T5E_LAST = R_T5E_FIRST + T5E_ROWS - 1     # 286

# Master Activities Table Coordinates (Cols AA to AK)
R_ACT_FIRST = 109
R_ACT_LAST = R_ACT_FIRST + len(MASTER_ACTIVITIES) - 1  # 157

# Category-Task Catalog Coordinates (Cols AN to AO)
R_CAT_FIRST = 109
R_CAT_LAST = R_CAT_FIRST + len(HEADING_TASK_CATALOG) - 1  # 153

LAST_COL = 15  # Column O (Visible printable columns A to O)



def _apply_param_row(ws, r, styles, label, val_formula, unit, role, note, number_format=None, dv=None):
    """Formats a parameter row in Section 1 (Cols A-O)."""
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    lbl_cell = ws.cell(row=r, column=1, value=label)
    lbl_cell.font = styles['font_bold']
    lbl_cell.alignment = styles['align_left']
    
    val_cell = ws.cell(row=r, column=4, value=val_formula)
    val_cell.alignment = styles['align_center']
    val_cell.font = styles['font_say'] if role == 'SAY' else styles['font_bold']
    
    role_fill = {
        'INPUT': styles['fill_input'],
        'LOOKUP': styles['fill_lookup'],
        'OVERRIDE': styles['fill_override'],
        'DERIVED': styles['fill_calc'],
        'RESULT': styles['fill_result'],
        'SAY': styles['fill_say']
    }.get(role, styles['fill_calc'])
    val_cell.fill = role_fill
    
    if number_format:
        val_cell.number_format = number_format
    if role in ('INPUT', 'OVERRIDE'):
        val_cell.protection = Protection(locked=False)
    if dv:
        dv.add(val_cell)
        
    unit_cell = ws.cell(row=r, column=5, value=unit)
    unit_cell.alignment = styles['align_center']
    unit_cell.font = styles['font_regular']
    
    role_cell = ws.cell(row=r, column=6, value=role)
    role_cell.alignment = styles['align_center']
    role_cell.font = styles['font_bold']
    role_cell.fill = role_fill
    
    ws.merge_cells(start_row=r, start_column=7, end_row=r, end_column=LAST_COL)
    note_cell = ws.cell(row=r, column=7, value=note)
    note_cell.font = styles['font_note']
    note_cell.alignment = styles['align_wrap']
    
    for c in range(1, LAST_COL + 1):
        ws.cell(row=r, column=c).border = styles['border_thin']
    ws.row_dimensions[r].height = 24


def build_earthwork_trade(wb, config, styles):
    sn = config.get('sheet_name', '02_Earth_Work')
    sheet_idx = None
    if sn in wb.sheetnames:
        sheet_idx = wb.sheetnames.index(sn)
        wb.remove(wb[sn])
    if sheet_idx is not None:
        ws = wb.create_sheet(title=sn, index=sheet_idx)
    else:
        ws = wb.create_sheet(title=sn)
        
    config.setdefault('trade_title', 'CPWD DAR 2019 — SUB-HEAD 02: EARTH WORK (FIRST-PRINCIPLES SIMULATOR)')
    config.setdefault('trade_guidance', 'CPWD DAR Sub-head 02: Earthwork operations (surface excavation, foundation trenches, embankment, backfilling, site clearance). Driven by civil engineering parameters with 5-step compounding statutory markups.')
    
    ws.views.sheetView[0].showGridLines = True
    add_trade_header_and_legend(ws, config, styles)
    
    # -------------------------------------------------------------
    # Column Widths (15 Visible Columns: A to O)
    # -------------------------------------------------------------
    col_widths = {
        'A': 10,  # Line / Step / Scope Code
        'B': 26,  # Primary Scope / Res Code / Markup Step
        'C': 34,  # Strata / Sub-Task / Res Name / Resource Code
        'D': 34,  # Method / Res Role / Selected Value
        'E': 14,  # Batch Qty / Unit / Base Amount
        'F': 16,  # Category / Gang Size / Day Coeff
        'G': 16,  # Unit Rate / Duration H / Factor %
        'H': 18,  # Line Amount / Total MH / Running Total
        'I': 14,  # Markup Tag / Shift Hours / Role
        'J': 16,  # Cost Share % / Day Coeff (D)
        'K': 18,  # Daily Productivity Norm (P)
        'L': 16,  # Unit Specific Effort (SE)
        'M': 20,  # Hourly Pace (Pace)
        'N': 44,  # CPWD Field Specification Guidance & Notes
        'O': 56,  # Engineering First-Principles 8-Hour Derivation Logic
        
        # Auxiliary, Helper Tables & Dedicated Audit Table (P to W)
        'P': 28,  # Helper 1B Res Code / Helper 2 Base Physical Norm
        'Q': 28,  # Helper 1B Description / Helper 2 Lever Factor
        'R': 24,  # Helper 1B Governing Lever / Helper 2 Net Calculated Qty
        'S': 36,  # Helper 1B Multiplier / Audit Col S: Governing Operational Precedent
        'T': 55,  # Audit Col T: Live Precedent & Calculation Trace
        'U': 52,  # Audit Col U: Engineering First-Principles (Site Mechanics)
        'V': 52,  # Audit Col V: CPWD 8-Hour Shift Logic (Mathematical Derivation)
        'W': 52,  # Audit Col W: Contractual Omission & Deduction Audit Clause
        'J': 52,  # Table 5E Col J: Category, Task & Method Line Key (C9|D9|D8|Line)
        'K': 16,  # Table 5E Col K: Catalog Norm Input Qty
        'L': 45,  # Table 5E Col L: Category & Task Line Key Fallback (C9|D9|Line)
        'Q': 55,  # Table 5E Col Q: Internal Activity Scope Key (D6|D7|D8|D9|Line)
        'Z': 45,  # Master Category & Task Key (C9|D9)
        'AA': 45, # Master Activity Full Key
        'AB': 35, # Master Activity Fallback Key
        'AC': 12, # Master Batch Qty
        'AD': 12, # Master Batch Unit
        'AE': 14, # Master Published Say
        'AF': 60, # Master Full Nomenclature
        'AG': 30, # Master Loosening
        'AH': 30, # Master Handling
        'AI': 30, # Master Spreading
        'AJ': 30, # Master Watering
        'AK': 30, # Master Compaction
        'AL': 45, # Master Task Options
        'AN': 36, # Category / Sub-Head Heading (Source for C9 & D9 Dependent Range)
        'AO': 58, # Task / Operation Specification Options (D9 Dropdown Source)
        # Decision-tree cascade lookup tables (L1→L2→L3→L4)
        'AU': 52, # L2 key (L1 value repeated per L2 entry)
        'AV': 62, # L2 value (sub-scope / geometry label)
        'AW': 72, # L3 key (L1|L2 concatenated)
        'AX': 52, # L3 value (ground material / location / girth)
        'AY': 88, # L4 key (L1|L2|L3 concatenated)
        'AZ': 62, # L4 value (terminal task name = HEADING_TASK_CATALOG entry)
    }
    for col_l, w in col_widths.items():
        ws.column_dimensions[col_l].width = w

    # -------------------------------------------------------------
    # Data Validations (Dropdowns)
    # -------------------------------------------------------------
    dv_yesno = DataValidation(type='list', formula1='"YES,NO"', allow_blank=False)
    ws.add_data_validation(dv_yesno)

    # ── L1: Primary Work Nature (D6) — static list in AP column ─────────────
    r_l1_end = R_CAT_FIRST + len(L1_OPTIONS) - 1
    dv_scope = DataValidation(type='list', formula1=f'=$AP${R_CAT_FIRST}:$AP${r_l1_end}', allow_blank=False)
    ws.add_data_validation(dv_scope)

    # ── L2: Sub-scope / Geometry (D7) — cascades from D6 ────────────────────
    # Key col AU (47), value col AV (48).  Row range 109 to 109+len(L2_DATA)-1.
    _l2_end = R_CAT_FIRST + len(L2_DATA) - 1
    dv_d7 = DataValidation(
        type='list',
        formula1=(
            f'=OFFSET($AV${R_CAT_FIRST},'
            f'MATCH($D${R_SCOPE},$AU${R_CAT_FIRST}:$AU${_l2_end},0)-1,'
            f'0,COUNTIF($AU${R_CAT_FIRST}:$AU${_l2_end},$D${R_SCOPE}),1)'
        ),
        allow_blank=True,
    )
    ws.add_data_validation(dv_d7)

    # ── L3: Ground Material / Location / Girth (D8) — cascades from D6+D7 ───
    # Key col AW (49), value col AX (50).
    _l3_end = R_CAT_FIRST + len(L3_DATA) - 1
    dv_d8 = DataValidation(
        type='list',
        formula1=(
            f'=OFFSET($AX${R_CAT_FIRST},'
            f'MATCH($D${R_SCOPE}&"|"&$D${R_STRATA},$AW${R_CAT_FIRST}:$AW${_l3_end},0)-1,'
            f'0,COUNTIF($AW${R_CAT_FIRST}:$AW${_l3_end},$D${R_SCOPE}&"|"&$D${R_STRATA}),1)'
        ),
        allow_blank=True,
    )
    ws.add_data_validation(dv_d8)

    # ── L4: Terminal Item Specification (D9) — cascades from D6+D7+D8 ───────
    # Key col AY (51), value col AZ (52).
    _l4_end = R_CAT_FIRST + len(L4_DATA) - 1
    dv_task = DataValidation(
        type='list',
        formula1=(
            f'=OFFSET($AZ${R_CAT_FIRST},'
            f'MATCH($D${R_SCOPE}&"|"&$D${R_STRATA}&"|"&$D${R_METHOD},'
            f'$AY${R_CAT_FIRST}:$AY${_l4_end},0)-1,'
            f'0,COUNTIF($AY${R_CAT_FIRST}:$AY${_l4_end},'
            f'$D${R_SCOPE}&"|"&$D${R_STRATA}&"|"&$D${R_METHOD}),1)'
        ),
        allow_blank=True,
    )
    ws.add_data_validation(dv_task)

    # ── Legacy aux DVs (retained for Section 1B levers and audit references) ─
    r_depth_end = R_CAT_FIRST + len(LIFT_DEPTH_OPTIONS) - 1
    dv_depth = DataValidation(type='list', formula1=f'=$AS${R_CAT_FIRST}:$AS${r_depth_end}', allow_blank=False)
    ws.add_data_validation(dv_depth)

    r_cat_end = R_CAT_FIRST + len(CATEGORY_HEADINGS) - 1
    dv_cat = DataValidation(type='list', formula1=f'=$AQ${R_CAT_FIRST}:$AQ${r_cat_end}', allow_blank=False)
    ws.add_data_validation(dv_cat)


    # Master activities ranges for lookups (Cols Z, AA, AB)
    act_cat_key_rng = f"$Z${R_ACT_FIRST}:$Z${R_ACT_LAST}"
    act_full_key_rng = f"$AA${R_ACT_FIRST}:$AA${R_ACT_LAST}"
    act_fb_key_rng = f"$AB${R_ACT_FIRST}:$AB${R_ACT_LAST}"

    cat_key_expr = f'$C$9&"|"&$D${R_TASK}'
    full_key_expr = f'$D${R_SCOPE}&"|"&$D${R_STRATA}&"|"&$D${R_METHOD}&"|"&$D${R_TASK}'
    fb_key_expr = f'$D${R_SCOPE}&"|"&$D${R_STRATA}&"|"&$D${R_METHOD}'

    def _master_lookup(col_letter, default_val):
        target_rng = f"${col_letter}${R_ACT_FIRST}:${col_letter}${R_ACT_LAST}"
        default_str = f'"{default_val}"' if isinstance(default_val, str) else str(default_val)
        return (
            f'=IFERROR(INDEX({target_rng}, MATCH({cat_key_expr}, {act_cat_key_rng}, 0)), '
            f'IFERROR(INDEX({target_rng}, MATCH({full_key_expr}, {act_full_key_rng}, 0)), '
            f'IFERROR(INDEX({target_rng}, MATCH({fb_key_expr}, {act_fb_key_rng}, 0)), {default_str})))'
        )

    # -------------------------------------------------------------
    # SECTION 1: RATE ANALYSIS CONFIGURATION & OPERATIONAL PROFILE
    # -------------------------------------------------------------
    ws.merge_cells(f'A{R_SEC1_HEAD}:O{R_SEC1_HEAD}')
    s1_title = ws[f'A{R_SEC1_HEAD}']
    s1_title.value = '1. RATE ANALYSIS CONFIGURATION — GUIDED DECISION TREE  ▶  D6 Primary Nature → D7 Sub-scope → D8 Ground Material → D9 Item Specification'
    s1_title.font = styles['font_white_bold']
    s1_title.fill = styles['fill_header']
    s1_title.alignment = styles['align_left']
    ws.row_dimensions[R_SEC1_HEAD].height = 26
    
    s1_headers = ['Parameter', '', '', 'Selected Value / Parameter Input', 'Unit', 'Cell Role', 'CPWD DAR 2019 Engineering Specification & Practical Field Guidance Notes']
    for ci, h in enumerate(s1_headers, 1):
        c = ws.cell(row=R_SEC1_COLS, column=ci, value=h)
        c.font = styles['font_header']
        c.fill = styles['fill_header']
        c.alignment = styles['align_center']
        c.border = styles['border_header']
    ws.merge_cells(f'A{R_SEC1_COLS}:C{R_SEC1_COLS}')
    ws.merge_cells(f'G{R_SEC1_COLS}:O{R_SEC1_COLS}')
    ws.row_dimensions[R_SEC1_COLS].height = 24
    
    _apply_param_row(ws, R_SCOPE, styles,
                     'Primary Work Nature',
                     'Excavation',
                     '—', 'INPUT',
                     'STEP 1 of 4 — Select the primary nature of the earthwork: Excavation · Banking/Embankment · Filling · Timbering · Site Clearance · Chemical Anti-Termite · Extra/Add-on. '
                     'Your choice here drives the Sub-scope dropdown (D7) → Ground Material (D8) → Item Specification (D9) in a guided decision tree.',
                     dv=dv_scope)
    _apply_param_row(ws, R_STRATA, styles,
                     'Sub-scope / Geometry',
                     'Surface (≤30 cm depth, width >1.5 m, area >10 sqm)',
                     '—', 'INPUT',
                     'STEP 2 of 4 — Sub-scope options shown depend on Step 1 (D6). '
                     'For Excavation: Surface / Open Area / Foundation Trench / Pipe Trench / Isolated Hole. '
                     'For Timbering: Close or Open. For Filling: material type. Reselect D7 whenever D6 changes.',
                     dv=dv_d7)
    _apply_param_row(ws, R_METHOD, styles,
                     'Ground Material / Location / Girth',
                     'All kinds of soil',
                     '—', 'INPUT',
                     'STEP 3 of 4 — Material, location, or girth options shown depend on Steps 1+2 (D6+D7). '
                     'Excavation paths: All soil / Ordinary rock / Hard rock blasting / Hard rock no blast. '
                     'Timbering paths: In trenches / In shafts & wells / Over areas. '
                     'Tree felling: girth band. Shows N/A when Step 3 is not needed for this scope.',
                     dv=dv_d8)

    # Row 9: C9 (auto-derived category) + D9 (terminal item — L4 cascade)
    ws.merge_cells(f'A{R_TASK}:B{R_TASK}')
    lbl_9 = ws.cell(row=R_TASK, column=1, value='Item Specification (Terminal Selection):')
    lbl_9.font = styles['font_bold']
    lbl_9.alignment = styles['align_left']

    # C9: Auto-derived category heading (formula, not a dropdown).
    # Looks up D9 task name in the AO catalog and returns the matching AN category.
    # This feeds _master_lookup which uses C9|D9 as the primary lookup key.
    _ao_lookup_end = R_CAT_FIRST + len(HEADING_TASK_CATALOG) - 1
    c9_cell = ws.cell(row=R_TASK, column=3,
                      value=f'=IFERROR(INDEX($AN${R_CAT_FIRST}:$AN${_ao_lookup_end},'
                            f'MATCH($D${R_TASK},$AO${R_CAT_FIRST}:$AO${_ao_lookup_end},0)),"—")')
    c9_cell.alignment = styles['align_center']
    c9_cell.font = styles['font_note']
    c9_cell.fill = styles['fill_lookup']
    # C9 is locked (formula-derived); do NOT add dv_cat here

    # D9: Terminal item specification — L4 cascade driven by D6 + D7 + D8.
    d9_cell = ws.cell(row=R_TASK, column=4, value='General surface cut (≤30 cm deep)')
    d9_cell.alignment = styles['align_center']
    d9_cell.font = styles['font_bold']
    d9_cell.fill = styles['fill_input']
    d9_cell.protection = Protection(locked=False)
    dv_task.add(d9_cell)

    # E9: Unit
    e9_cell = ws.cell(row=R_TASK, column=5, value='—')
    e9_cell.alignment = styles['align_center']
    e9_cell.font = styles['font_regular']

    # F9: Role label
    f9_cell = ws.cell(row=R_TASK, column=6, value='INPUT')
    f9_cell.alignment = styles['align_center']
    f9_cell.font = styles['font_bold']
    f9_cell.fill = styles['fill_input']

    # G9:O9: Guidance note with stale-value detection
    ws.merge_cells(f'G{R_TASK}:O{R_TASK}')
    g9_cell = ws.cell(row=R_TASK, column=7,
                      value=(
                          f'=IF(COUNTIF($AO${R_CAT_FIRST}:$AO${_ao_lookup_end},$D${R_TASK})=0,'
                          f'"[STALE ITEM] D9 value ["&$D${R_TASK}&"] not found in catalog. '
                          f'Reselect D9 from the dropdown after confirming D6, D7, D8.",'
                          f'"STEP 4 of 4 — Select the terminal item specification. '
                          f'Options shown are filtered to exactly those valid for your D6→D7→D8 path. '
                          f'Active path: ["&$D${R_SCOPE}&"] › ["&$D${R_STRATA}&"] › ["&$D${R_METHOD}&"]. '
                          f'Selected item: ["&$D${R_TASK}&"]. C9 category is auto-derived.")'
                      ))
    g9_cell.font = styles['font_note']
    g9_cell.alignment = styles['align_left']
    for c in range(1, LAST_COL + 1):
        ws.cell(row=R_TASK, column=c).border = styles['border_thin']
    ws.row_dimensions[R_TASK].height = 22

    # Batch Qty & Unit looked up from Master Activities Table
    _apply_param_row(ws, R_BATCH_QTY, styles, 'Standard Output Batch Quantity',
                     _master_lookup('AC', 100),
                     'units', 'LOOKUP',
                     'Standard output batch volume or area defined in CPWD DAR 2019 (100 sqm, 10 cum, 1 m, 1 tree, 10 holes).', number_format='0.00')

    _apply_param_row(ws, R_BATCH_UNIT, styles, 'Standard Output Measurement Unit',
                     _master_lookup('AD', 'sqm'),
                     '—', 'LOOKUP',
                     'Billing measurement unit for work output (cum, sqm, metre, tree, hole).')
    
    _apply_param_row(ws, R_CUSTOM_CODE, styles, 'Custom Project Rate Reference Code', 'C-02.01', '—', 'INPUT',
                     'Project-specific non-DSR reference code when saving custom rate variations.')
    
    # Official Nomenclature
    ws.merge_cells(f'A{R_NOMENCLATURE}:C{R_NOMENCLATURE}')
    ws.cell(row=R_NOMENCLATURE, column=1, value='Official Work Specification Nomenclature:').font = styles['font_bold']
    ws.merge_cells(f'D{R_NOMENCLATURE}:O{R_NOMENCLATURE}')
    nom_cell = ws.cell(row=R_NOMENCLATURE, column=4,
                       value=_master_lookup('AF', 'Earth Work Operation'))
    nom_cell.font = styles['font_regular']
    nom_cell.fill = styles['fill_lookup']
    nom_cell.alignment = styles['align_wrap']
    for c in range(1, LAST_COL + 1):
        ws.cell(row=R_NOMENCLATURE, column=c).border = styles['border_thin']
    ws.row_dimensions[R_NOMENCLATURE].height = 40
    
    # Published Benchmark Reference Say Rate
    _apply_param_row(ws, R_DAR_PUB, styles, 'Published Reference Say Rate (Benchmark)',
                     f'=IF(AND(OR($D$9="Open areas & wide foundation trenches", $D$9="Foundation trenches & drain trenches"), NOT(ISNUMBER(SEARCH("Mechanical", $D${R_METHOD})))), "Non-DSR Custom Synthesis", ' + _master_lookup('AE', 92.55)[1:] + ')',
                     f'=D{R_BATCH_UNIT}', 'LOOKUP',
                     'Official CPWD DAR 2019 Volume 1 published rate for identical scope and strata.', number_format=styles['fmt_currency'])

    # Audit Status Bar
    ws.merge_cells(f'A{R_AUDIT}:C{R_AUDIT}')
    ws.cell(row=R_AUDIT, column=1, value='AUDIT VERIFICATION STATUS:').font = styles['font_white_bold']
    ws.cell(row=R_AUDIT, column=1).fill = styles['fill_header']
    ws.cell(row=R_AUDIT, column=1).alignment = styles['align_center']
    
    ws.cell(row=R_AUDIT, column=4, value=f'=IF(ISNUMBER(D{R_DAR_PUB}), IF(AND(H{R_DAR_DIFF}<=0.05, H{R_DAR_DIFF}>=-0.05), "[PASS] 100% DAR MATCH", "[ALERT] VARIANCE DETECTED"), "[PASS] FIRST-PRINCIPLES DERIVED")')
    ws.cell(row=R_AUDIT, column=4).font = styles['font_result']
    ws.cell(row=R_AUDIT, column=4).fill = styles['fill_result']

    
    ws.cell(row=R_AUDIT, column=6, value=f'=IF(H{R_W_SUB}>0, "OK", "ERR: W<=0")').font = styles['font_bold']
    ws.cell(row=R_AUDIT, column=6).alignment = styles['align_center']
    
    ws.merge_cells(f'G{R_AUDIT}:O{R_AUDIT}')
    ws.cell(row=R_AUDIT, column=7, value='Live variance check against published CPWD DAR 2019 Volume 1 Say Rate. Target variance: ₹ 0.00.').font = styles['font_note']
    for c in range(1, LAST_COL + 1):
        ws.cell(row=R_AUDIT, column=c).border = styles['border_thin']
    ws.row_dimensions[R_AUDIT].height = 24

    # -------------------------------------------------------------
    # ROW 16: ANALYSIS MODE FLAG & ITEM LINKAGE STATUS BAR
    # (Sits between Section 1 and Section 1B — no row constants shift)
    # -------------------------------------------------------------
    dv_mode = DataValidation(type='list', formula1='"DSR,CUSTOM"', allow_blank=False)
    ws.add_data_validation(dv_mode)

    ws.merge_cells('A16:C16')
    r16_lbl = ws.cell(row=16, column=1, value='Analysis Mode / Item Linkage Status:')
    r16_lbl.font = styles['font_bold']
    r16_lbl.alignment = styles['align_left']

    d16 = ws.cell(row=16, column=4, value='DSR')
    d16.font = styles['font_bold']
    d16.fill = styles['fill_input']
    d16.alignment = styles['align_center']
    d16.protection = Protection(locked=False)
    dv_mode.add(d16)

    ws.cell(row=16, column=5, value='mode').alignment = styles['align_center']
    ws.cell(row=16, column=5).font = styles['font_regular']

    f16 = ws.cell(row=16, column=6, value='INPUT')
    f16.font = styles['font_bold']
    f16.fill = styles['fill_input']
    f16.alignment = styles['align_center']

    ws.merge_cells('G16:O16')
    lnk_status = ws.cell(row=16, column=7,
        value=(
            f'=IF($D$16="CUSTOM",'
            f'"[CUSTOM MODE] Table 2A rows unlocked: press Delete on any B/F/G cell to clear its formula, '
            f'then type Resource Code (Col B), Qty (Col F), Rate (Col G) directly. '
            f'Col H amounts and all Section 3 markups calculate automatically. Tag this run in D12.",'
            f'"[DSR MODE] Lever Linkage: "'
            f'&IF(OR(ISNUMBER(SEARCH("Embankment",$D${R_SCOPE})),ISNUMBER(SEARCH("banking",$D${R_TASK})),ISNUMBER(SEARCH("rolling",$D${R_TASK}))),'
            f'"Watering (D22)="&$D${R_OP_WATERING}&" | Rolling (D23)="&$D${R_OP_COMPACTION}&" - LEVERS ACTIVE: deduction fires in Table 2B when set to NO.",'
            f'"Watering (D22): [NA] | Rolling (D23): [NA] - toggles have no rate effect for Excavation / Clearance scope.")'
            f'&" | Depth-Lift (D20): "'
            f'&IF(ISNUMBER(SEARCH("1.5 m to 3.0 m",$D${R_OP_HANDLING})),"ACTIVE - 1 extra lift stage (DAR 2.26.1 / 2.11)",'
            f'IF(ISNUMBER(SEARCH("3.0 m to 4.5 m",$D${R_OP_HANDLING})),"ACTIVE - 2 extra lift stages (DAR 2.26.1 / 2.12)","STANDARD <=1.5 m - no extra lift"))'
            f'&" | Strata: "&$D${R_STRATA})'
        )
    )
    lnk_status.font = styles['font_note']
    lnk_status.alignment = styles['align_wrap']
    lnk_status.fill = styles['fill_note']

    for c in range(1, LAST_COL + 1):
        ws.cell(row=16, column=c).border = styles['border_thin']
    ws.row_dimensions[16].height = 30

    # -------------------------------------------------------------
    # SECTION 1B: PARAMETRIC OPERATIONAL DECOMPOSITION (ARCHITECTURE 3)
    # -------------------------------------------------------------
    ws.merge_cells(f'A{R_SEC1B_HEAD}:O{R_SEC1B_HEAD}')
    s1b_title = ws[f'A{R_SEC1B_HEAD}']
    s1b_title.value = '1B. PARAMETRIC OPERATIONAL SPECIFICATION PROFILE (5 CORE CIVIL WORK LEVERS)'
    s1b_title.font = styles['font_white_bold']
    s1b_title.fill = styles['fill_header']
    s1b_title.alignment = styles['align_left']
    ws.row_dimensions[R_SEC1B_HEAD].height = 24
    
    s1b_cols = ['Operational Work Lever', '', '', 'Active Execution Mechanism / Field Technique', 'CPWD Standard Field Specification & Practical Engineering Scope Notes', '', '', '', '', '', '', '', '', '', '']
    for ci, h in enumerate(s1b_cols, 1):
        c = ws.cell(row=R_SEC1B_COLS, column=ci, value=h)
        c.font = styles['font_header']
        c.fill = styles['fill_header']
        c.alignment = styles['align_center']
        c.border = styles['border_header']
    ws.merge_cells(f'A{R_SEC1B_COLS}:C{R_SEC1B_COLS}')
    ws.merge_cells(f'E{R_SEC1B_COLS}:O{R_SEC1B_COLS}')
    ws.row_dimensions[R_SEC1B_COLS].height = 24
    
    # Rows 19 to 23: 5 Levers
    # Lever 1: Loosening
    ws.merge_cells(f'A{R_OP_LOOSENING}:C{R_OP_LOOSENING}')
    ws.cell(row=R_OP_LOOSENING, column=1, value='1. Loosening / Digging Mode').font = styles['font_bold']
    ws.cell(row=R_OP_LOOSENING, column=4, value=_master_lookup('AG', 'Standard Digging')).font = styles['font_bold']
    ws.cell(row=R_OP_LOOSENING, column=4).fill = styles['fill_lookup']
    ws.merge_cells(f'E{R_OP_LOOSENING}:O{R_OP_LOOSENING}')
    ws.cell(row=R_OP_LOOSENING, column=5, value='Loosening of soil/rock crust to required formation level or trench invert.').font = styles['font_note']
    ws.cell(row=R_OP_LOOSENING, column=5).alignment = styles['align_wrap']
    ws.row_dimensions[R_OP_LOOSENING].height = 22

    # Lever 2: Handling & Haulage / Depth Stage
    ws.merge_cells(f'A{R_OP_HANDLING}:C{R_OP_HANDLING}')
    ws.cell(row=R_OP_HANDLING, column=1, value='2. Excavation Depth & Lift Stage').font = styles['font_bold']
    d20_cell = ws.cell(row=R_OP_HANDLING, column=4, value='Depth ≤ 1.5 m (Standard lift upto 1.5 m)')
    d20_cell.font = styles['font_bold']
    d20_cell.fill = styles['fill_input']
    d20_cell.alignment = styles['align_center']
    d20_cell.protection = Protection(locked=False)
    dv_depth.add(d20_cell)
    ws.merge_cells(f'E{R_OP_HANDLING}:O{R_OP_HANDLING}')
    ws.cell(row=R_OP_HANDLING, column=5,
            value='ADDITIVE EXTRA-DEPTH LEVER: Select the stage that covers your deepest excavation point. '
                  'The base D9 task is always included; this lever ADDS the extra-depth cost on top. '
                  'Stage 1 - Standard (<=1.5 m): no add-on; DAR base rate only. '
                  'Stage 2 - Depth >1.5 m to 3.0 m: adds Extra lift (DAR 2.26.1) for soil or Extra pipe trench allowance (DAR 2.11) for pipe trenches. '
                  'Stage 3 - Depth >3.0 m to 4.5 m: adds both Extra lift stages (DAR 2.26.1 x2) or Extra pipe trench allowance (DAR 2.12). '
                  'These items (2.11, 2.12, 2.26.1) are ADDITIVE — they appear in Table 2B Section L, not in D9.'
           ).font = styles['font_note']
    ws.cell(row=R_OP_HANDLING, column=5).alignment = styles['align_wrap']
    ws.row_dimensions[R_OP_HANDLING].height = 24


    # Lever 3: Layer Placement & Spreading
    ws.merge_cells(f'A{R_OP_SPREADING}:C{R_OP_SPREADING}')
    ws.cell(row=R_OP_SPREADING, column=1, value='3. Layer Placement & Spreading').font = styles['font_bold']
    ws.cell(row=R_OP_SPREADING, column=4, value=_master_lookup('AI', 'Standard Spreading')).font = styles['font_bold']
    ws.cell(row=R_OP_SPREADING, column=4).fill = styles['fill_lookup']
    ws.merge_cells(f'E{R_OP_SPREADING}:O{R_OP_SPREADING}')
    ws.cell(row=R_OP_SPREADING, column=5, value='Depositing excavated earth in uniform layers ≤20 cm depth, breaking clods and shaping slopes.').font = styles['font_note']
    ws.cell(row=R_OP_SPREADING, column=5).alignment = styles['align_wrap']
    ws.row_dimensions[R_OP_SPREADING].height = 22

    # Lever 4: Compaction Watering to OMC (Interactive Dropdown D22)
    ws.merge_cells(f'A{R_OP_WATERING}:C{R_OP_WATERING}')
    ws.cell(row=R_OP_WATERING, column=1, value='4. Compaction Watering to OMC').font = styles['font_bold']
    c_water = ws.cell(row=R_OP_WATERING, column=4, value='YES')
    c_water.font = styles['font_bold']
    c_water.fill = styles['fill_input']
    c_water.alignment = styles['align_center']
    c_water.protection = Protection(locked=False)
    dv_yesno.add(c_water)
    ws.merge_cells(f'E{R_OP_WATERING}:O{R_OP_WATERING}')
    c_w_note = ws.cell(row=R_OP_WATERING, column=5,
                       value=f'=IF(OR(ISNUMBER(SEARCH("Embankment",$D${R_SCOPE})),ISNUMBER(SEARCH("banking",$D${R_TASK})),ISNUMBER(SEARCH("rolling",$D${R_TASK}))),'
                             f'IF($D${R_OP_WATERING}="YES",'
                             f'"[LEVER ACTIVE - INCLUDED] Watering to OMC is in this item rate (0.40 Bhishti-day / 10 cum). Full rate applies; no deduction in Table 2B. '
                             f'Purpose: moisten loose embankment layers to Optimum Moisture Content before rolling.",'
                             f'"[LEVER ACTIVE - DEDUCTION] Watering omitted by contractor (DAR Item 2.5). '
                             f'Deduction = -0.40 Bhishti-day / 10 cum @ Rs 617 = Rs 246.80 direct cost. '
                             f'Compounded: +GST 14.05% (Rs 34.68) + CPOH 15% (Rs 42.22) + Cess 1% (Rs 3.24) = Rs 326.94 / 10 cum. '
                             f'Rate reduction ~Rs 33.00/cum. See Table 2B Row D4 for live deduction amount."),'
                             f'"[LEVER INACTIVE] This toggle has NO rate effect for the selected scope (Surface Excavation / Foundation Trench / Jungle Clearance / Timbering). '
                             f'Watering to OMC is not a specified item for excavation tasks - no Bhishti norm applies and no deduction exists. '
                             f'Switch to Embankment or Banking scope to activate this lever.")')
    c_w_note.font = styles['font_note']
    c_w_note.alignment = styles['align_wrap']
    ws.row_dimensions[R_OP_WATERING].height = 26

    # Lever 5: Power Roller Compaction (Interactive Dropdown D23)
    ws.merge_cells(f'A{R_OP_COMPACTION}:C{R_OP_COMPACTION}')
    ws.cell(row=R_OP_COMPACTION, column=1, value='5. Compaction Mechanism').font = styles['font_bold']
    c_roll = ws.cell(row=R_OP_COMPACTION, column=4, value='YES')
    c_roll.font = styles['font_bold']
    c_roll.fill = styles['fill_input']
    c_roll.alignment = styles['align_center']
    c_roll.protection = Protection(locked=False)
    dv_yesno.add(c_roll)
    ws.merge_cells(f'E{R_OP_COMPACTION}:O{R_OP_COMPACTION}')
    c_r_note = ws.cell(row=R_OP_COMPACTION, column=5,
                       value=f'=IF(OR(ISNUMBER(SEARCH("Embankment",$D${R_SCOPE})),ISNUMBER(SEARCH("banking",$D${R_TASK})),ISNUMBER(SEARCH("rolling",$D${R_TASK}))),'
                             f'IF($D${R_OP_COMPACTION}="YES",'
                             f'"[LEVER ACTIVE - INCLUDED] Power Roller Compaction is in this item rate (0.008 roller-day / 10 cum). Full rate applies; no deduction in Table 2B. '
                             f'Standard: 8-10 tonne diesel road roller consolidating compacted embankment layers at 1250 cum/8-hr shift.",'
                             f'"[LEVER ACTIVE - DEDUCTION] Power rolling omitted by contractor (DAR Item 2.4). '
                             f'Deduction = -0.008 Roller-day/10 cum @ Rs 3000 (Rs 24.00) + 0.008 Chowkidar-day (Rs 4.46) + 1.82 Sundries (Rs 3.64) = Rs 32.10 direct. '
                             f'Compounded (GST 14.05% + CPOH 15% + Cess 1%) = Rs 42.95/10 cum. Rate reduction ~Rs 4.30/cum. '
                             f'See Table 2B Rows D1-D3 for live deduction amounts."),'
                             f'"[LEVER INACTIVE] This toggle has NO rate effect for the selected scope (Surface Excavation / Foundation Trench / Jungle Clearance / Timbering). '
                             f'Power roller compaction is not a specified operation for excavation tasks - the 0.008 roller-day norm and DAR Item 2.4 do not apply here. '
                             f'Switch to Embankment or Banking scope to activate this lever.")')
    c_r_note.font = styles['font_note']
    c_r_note.alignment = styles['align_wrap']
    ws.row_dimensions[R_OP_COMPACTION].height = 26
    
    for r in range(R_OP_LOOSENING, R_OP_COMPACTION + 1):
        for c in range(1, LAST_COL + 1):
            ws.cell(row=r, column=c).border = styles['border_thin']

    # -------------------------------------------------------------
    # HELPER TABLE 1B: OPERATIONAL LEVER MULTIPLIER RULES (Cols P to S, Rows 17 to 22)
    # -------------------------------------------------------------
    ws.merge_cells('P17:S17')
    h1b_title = ws['P17']
    h1b_title.value = 'HELPER TABLE 1B: OPERATIONAL LEVER MULTIPLIER RULES (Lookup Source for Table 2 Col Q)'
    h1b_title.font = styles['font_white_bold']
    h1b_title.fill = styles['fill_header']
    h1b_title.alignment = styles['align_left']
    
    h1b_cols = ['Resource Code', 'Resource Description & Trade Role', 'Governing Civil Lever', 'Active Multiplier']
    for ci, h in enumerate(h1b_cols, 16):  # Col P is 16
        c = ws.cell(row=18, column=ci, value=h)
        c.font = styles['font_header']
        c.fill = styles['fill_header']
        c.alignment = styles['align_center']
        c.border = styles['border_header']
        
    lever_rules = [
        ('0101', 'Bhishti (Compaction Watering to OMC)', f'Lever 4: Compaction Watering ($D${R_OP_WATERING})', f'=IF($D${R_OP_WATERING}="NO", 0, 1)'),
        ('0003', '10 Tonne Road Roller (Compaction)', f'Lever 5: Power Roller Compaction ($D${R_OP_COMPACTION})', f'=IF($D${R_OP_COMPACTION}="NO", 0, 1)'),
        ('0113', 'Chowkidar (Road Roller Guard)', f'Lever 5: Power Roller Compaction ($D${R_OP_COMPACTION})', f'=IF($D${R_OP_COMPACTION}="NO", 0, 1)'),
        ('9999', 'Sundries (Road Roller)', f'Lever 5: Power Roller Compaction ($D${R_OP_COMPACTION})', f'=IF($D${R_OP_COMPACTION}="NO", 0, 1)'),
    ]
    for idx, (rc, desc, lev, mult_form) in enumerate(lever_rules):
        hr = 19 + idx
        c_rc = ws.cell(row=hr, column=16, value=rc)
        c_rc.alignment = styles['align_center']
        c_rc.font = styles['font_bold']
        c_rc.fill = styles['fill_lookup']
        
        c_desc = ws.cell(row=hr, column=17, value=desc)
        c_desc.alignment = styles['align_left']
        c_desc.font = styles['font_regular']
        
        c_lev = ws.cell(row=hr, column=18, value=lev)
        c_lev.alignment = styles['align_left']
        c_lev.font = styles['font_note']
        
        c_m = ws.cell(row=hr, column=19, value=mult_form)
        c_m.alignment = styles['align_right']
        c_m.font = styles['font_bold']
        c_m.fill = styles['fill_calc']
        c_m.number_format = '0.00'
        
        for c in range(16, 20):
            ws.cell(row=hr, column=c).border = styles['border_thin']

    # -------------------------------------------------------------
    # SECTION 2A: BASE SPECIFICATION RESOURCE BUILD-UP (GROSS SCOPE)
    # -------------------------------------------------------------
    ws.merge_cells(f'A{R_SEC2A_HEAD}:O{R_SEC2A_HEAD}')
    s2_title = ws[f'A{R_SEC2A_HEAD}']
    s2_title.value = '2A. BASE SPECIFICATION RESOURCE BUILD-UP (Gross Scope; Unmodified CPWD DAR 2019 Norms)'
    s2_title.font = styles['font_white_bold']
    s2_title.fill = styles['fill_header']
    s2_title.alignment = styles['align_left']
    ws.row_dimensions[R_SEC2A_HEAD].height = 24
    
    s2_cols = [
        'Line', 'Resource Code', 'Resource Name & Description', 'Resource Category',
        'Unit', 'Norm Input Qty', 'Unit Rate (₹)', 'Line Amount (₹)',
        'Include? / Apply', 'Cost Share %', 'Daily Productivity Norm', 'Specific Effort',
        'Hourly Pace', 'Field Guidance Note', 'Engineering First-Principles & 8-Hour Shift Logic'
    ]
    for ci, h in enumerate(s2_cols, 1):
        c = ws.cell(row=R_SEC2A_COLS, column=ci, value=h)
        c.font = styles['font_header']
        c.fill = styles['fill_header']
        c.alignment = styles['align_center']
        c.border = styles['border_header']
    ws.row_dimensions[R_SEC2A_COLS].height = 24
    
    # Helper Table 2 Header in Cols P to R
    ws.merge_cells('P25:R25')
    h2_title = ws['P25']
    h2_title.value = 'HELPER TABLE 2: FIRST-PRINCIPLES QUANTITY RESOLVER (Calculates Table 2A Col F)'
    h2_title.font = styles['font_white_bold']
    h2_title.fill = styles['fill_header']
    h2_title.alignment = styles['align_left']
    
    h2_cols = ['Base Physical Norm (D10 / Kr)', 'Base Scope Status', 'Gross Base Quantity']
    for ci, h in enumerate(h2_cols, 16):  # Col P is 16
        c = ws.cell(row=R_SEC2A_COLS, column=ci, value=h)
        c.font = styles['font_header']
        c.fill = styles['fill_header']
        c.alignment = styles['align_center']
        c.border = styles['border_header']
    
    t5e_method_rng = f"$J${R_T5E_FIRST}:$J${R_T5E_LAST}"
    t5e_cat_rng = f"$L${R_T5E_FIRST}:$L${R_T5E_LAST}"
    t5e_full_rng = f"$Q${R_T5E_FIRST}:$Q${R_T5E_LAST}"
    
    # Audit Table Header in Cols S to W (Cols 19 to 23)
    ws.merge_cells('S25:W25')
    audit_title = ws['S25']
    audit_title.value = 'DEDICATED ENGINEERING LOGIC, PRECEDENTS & AUDIT TABLE (Precedents, Site Mechanics & Shift Derivation)'
    audit_title.font = styles['font_white_bold']
    audit_title.fill = styles['fill_header']
    audit_title.alignment = styles['align_left']
    
    audit_cols = [
        'Governing Operational Precedent',
        'Live Precedent & Calculation Trace',
        'Engineering First-Principles (Physical Mechanics)',
        'CPWD 8-Hour Shift Logic (Mathematical Derivation)',
        'Contractual Omission & Deduction Audit Clause'
    ]
    for ci, h in enumerate(audit_cols, 19):  # Col S is 19
        c = ws.cell(row=R_SEC2A_COLS, column=ci, value=h)
        c.font = styles['font_header']
        c.fill = styles['fill_header']
        c.alignment = styles['align_center']
        c.border = styles['border_header']
    
    # 10 Resource Rows for Table 2A (Rows 27 to 36)
    for line_idx in range(1, RES_ROWS + 1):
        r = R_RES_FIRST + line_idx - 1
        
        ws.cell(row=r, column=1, value=line_idx).alignment = styles['align_center']
        ws.cell(row=r, column=1).font = styles['font_bold']
        
        method_expr = 'IF(ISNUMBER(SEARCH("Mechanical",$D$8)),"Mechanical (Hydraulic Excavator 0.9 cum)","Manual labor (depth <=1.5m)")'
        line_key_method = f'$C$9&"|"&$D$9&"|"&{method_expr}&"|"&{line_idx}'
        method_exists_key = f'$C$9&"|"&$D$9&"|"&{method_expr}&"|1"'
        line_key_cat = f'$C$9&"|"&$D$9&"|"&{line_idx}'
        line_key_full = f'$D${R_SCOPE}&"|"&$D${R_STRATA}&"|"&$D${R_METHOD}&"|"&$D${R_TASK}&"|"&{line_idx}'

        def _t5e_line_lookup(col_letter, default_val='""'):
            rng = f"${col_letter}${R_T5E_FIRST}:${col_letter}${R_T5E_LAST}"
            # Tier 1: Full scope key D6|D7|D8|D9|line — picks up strata-specific records first
            # Tier 2: Method key C9|D9|method|line — primary match for existing data
            # Tier 3: Category key C9|D9|line — broadest fallback
            # This order means D7 (Ground Strata) drives the lookup whenever strata-specific
            # records exist in Table 5E (Col Q); falls back gracefully to Tier 2 & 3.
            return (
                f'IFERROR(INDEX({rng}, MATCH({line_key_full}, {t5e_full_rng}, 0)), '
                f'IFERROR(INDEX({rng}, MATCH({line_key_method}, {t5e_method_rng}, 0)), '
                f'IFERROR(INDEX({rng}, MATCH({line_key_cat}, {t5e_cat_rng}, 0)), {default_val})))'
            )

        
        # Resource Code in Col B — DSR: auto-lookup; CUSTOM: blank so user can type directly
        c_code = ws.cell(row=r, column=2, value=f'=IF($D$16="DSR",{_t5e_line_lookup("F")},"")')
        c_code.alignment = styles['align_center']
        c_code.font = styles['font_bold']
        c_code.fill = styles['fill_lookup']
        c_code.protection = Protection(locked=False)  # Unlocked for CUSTOM mode entry
        
        # Resource Name in Col C
        name_fallback = _t5e_line_lookup('M', '"Custom Resource"')
        c_name = ws.cell(row=r, column=3,
                         value=f'=IF(B{r}="","", IFERROR(INDEX(Rates_Master!$C:$C, MATCH(B{r}, Rates_Master!$A:$A, 0)), {name_fallback}))')
        c_name.alignment = styles['align_left']
        c_name.font = styles['font_regular']
        
        # Category in Col D
        c_cat = ws.cell(row=r, column=4,
                        value=f'=IF(B{r}="","", IFERROR(INDEX(Rates_Master!$B:$B, MATCH(B{r}, Rates_Master!$A:$A, 0)), "Labour"))')
        c_cat.alignment = styles['align_center']
        c_cat.font = styles['font_note']
        
        # Unit in Col E
        unit_fallback = _t5e_line_lookup('N', '"day"')
        c_unit = ws.cell(row=r, column=5,
                         value=f'=IF(B{r}="","", IFERROR(INDEX(Rates_Master!$D:$D, MATCH(B{r}, Rates_Master!$A:$A, 0)), {unit_fallback}))')
        c_unit.alignment = styles['align_center']
        c_unit.font = styles['font_regular']
        
        # Productivity in Col K
        c_prod = ws.cell(row=r, column=11, value='=' + _t5e_line_lookup('G'))
        c_prod.alignment = styles['align_right']
        c_prod.font = styles['font_note']
        c_prod.number_format = '0.00'
        
        # Helper Table 2 calculations in Cols P, Q, R:
        raw_qty_lookup = _t5e_line_lookup('K', '0')
        c_base_q = ws.cell(row=r, column=16,
                           value=f'=IF(B{r}="","", IF(AND(ISNUMBER(K{r}), K{r}>0), ROUND(D${R_BATCH_QTY}/K{r}, 5), {raw_qty_lookup}))')
        c_base_q.alignment = styles['align_right']
        c_base_q.font = styles['font_regular']
        c_base_q.fill = styles['fill_calc']
        c_base_q.number_format = '0.000'
        
        c_fac_q = ws.cell(row=r, column=17,
                          value=f'=IF(B{r}="","", "GROSS BASE")')
        c_fac_q.alignment = styles['align_center']
        c_fac_q.font = styles['font_note']
        c_fac_q.fill = styles['fill_lookup']
        
        c_net_q = ws.cell(row=r, column=18,
                          value=f'=IF(B{r}="","", P{r})')
        c_net_q.alignment = styles['align_right']
        c_net_q.font = styles['font_bold']
        c_net_q.fill = styles['fill_result']
        c_net_q.number_format = '0.000'

        # Norm Input Qty in Col F — unlocked so user can override in CUSTOM mode
        c_qty = ws.cell(row=r, column=6, value=f'=IF(B{r}="","", R{r})')
        c_qty.alignment = styles['align_right']
        c_qty.font = styles['font_bold']
        c_qty.fill = styles['fill_input']
        c_qty.number_format = '0.000'
        c_qty.protection = Protection(locked=False)  # Unlocked for CUSTOM mode direct entry

        # Rate in Col G — unlocked so user can override in CUSTOM mode
        rate_fallback = _t5e_line_lookup('O', '0')
        c_rate = ws.cell(row=r, column=7,
                         value=f'=IF(B{r}="","", IFERROR(INDEX(Rates_Master!$E:$E, MATCH(B{r}, Rates_Master!$A:$A, 0)), {rate_fallback}))')
        c_rate.alignment = styles['align_right']
        c_rate.font = styles['font_regular']
        c_rate.fill = styles['fill_lookup']
        c_rate.number_format = styles['fmt_currency']
        c_rate.protection = Protection(locked=False)  # Unlocked for CUSTOM mode direct entry
        
        # Line Amount in Col H (=ROUND(F*G, 2))
        c_amt = ws.cell(row=r, column=8,
                        value=f'=IF(OR(B{r}="", F{r}=""), 0, ROUND(F{r} * G{r}, 2))')
        c_amt.alignment = styles['align_right']
        c_amt.font = styles['font_bold']
        c_amt.fill = styles['fill_calc']
        c_amt.number_format = styles['fmt_currency']
        
        # Include? / Tag in Col I
        tag_lookup = _t5e_line_lookup('P', '"W"')
        c_tag = ws.cell(row=r, column=9, value=f'=IF(B{r}="","", {tag_lookup})')
        c_tag.alignment = styles['align_center']
        c_tag.font = styles['font_bold']
        c_tag.fill = styles['fill_lookup']
        
        # Cost Share % in Col J (=ROUND(H/H$37*100, 1))
        c_pct = ws.cell(row=r, column=10,
                        value=f'=IF(H${R_W_GROSS}>0, ROUND(H{r} / H${R_W_GROSS} * 100, 1), 0)')
        c_pct.alignment = styles['align_right']
        c_pct.font = styles['font_regular']
        c_pct.fill = styles['fill_calc']
        c_pct.number_format = '0.0"%"'
        
        # Specific Effort in Col L
        c_se = ws.cell(row=r, column=12,
                       value=f'=IF(AND(ISNUMBER(F{r}), F{r}>0, D${R_BATCH_QTY}>0), ROUND(F{r} * 8 / D${R_BATCH_QTY}, 4), "")')
        c_se.alignment = styles['align_right']
        c_se.font = styles['font_note']
        c_se.number_format = '0.0000'
        
        # Hourly Pace in Col M
        c_pace = ws.cell(row=r, column=13,
                         value=f'=IF(AND(ISNUMBER(F{r}), F{r}>0, D${R_BATCH_QTY}>0), ROUND(D${R_BATCH_QTY} / (F{r} * 8), 2), "")')
        c_pace.alignment = styles['align_right']
        c_pace.font = styles['font_note']
        c_pace.number_format = '0.00'
        
        # Field Note in Col N
        c_note = ws.cell(row=r, column=14,
                         value=f'=IF(B{r}="","", "CPWD Standard Earthwork Production Resource")')
        c_note.alignment = styles['align_left']
        c_note.font = styles['font_note']
        
        # Derivation Summary in Col O
        c_logic = ws.cell(row=r, column=15, value=f'=IF(B{r}="","", V{r})')
        c_logic.alignment = styles['align_left']
        c_logic.font = styles['font_note']
        
        # Dedicated Audit Columns (Cols S to W, Rows 27 to 36)
        prec_form = f'=IF(B{r}="","", "Core Scope Baseline (Mandatory Execution)")'
        c_audit_s = ws.cell(row=r, column=19, value=prec_form)
        c_audit_s.alignment = styles['align_left']
        c_audit_s.font = styles['font_note']

        c_audit_t = ws.cell(row=r, column=20,
                            value=f'=IF(B{r}="","", "Batch Q=" & TEXT(D${R_BATCH_QTY}, "0.00") & " " & $D${R_BATCH_UNIT} & " | Norm K=" & TEXT(K{r}, "0.00") & " | Gross Qty F=" & TEXT(F{r}, "0.000") & " " & E{r} & " | Rate G=Rs " & TEXT(G{r}, "0.00") & " | Amt H=Rs " & TEXT(H{r}, "0.00"))')
        c_audit_t.alignment = styles['align_left']
        c_audit_t.font = styles['font_note']

        c_audit_u = ws.cell(row=r, column=21,
                            value=f'=IF(B{r}="","", ' + _t5e_line_lookup('H') + ')')
        c_audit_u.alignment = styles['align_wrap']
        c_audit_u.font = styles['font_note']

        c_audit_v = ws.cell(row=r, column=22,
                            value=f'=IF(B{r}="","", ' + _t5e_line_lookup('I') + ')')
        c_audit_v.alignment = styles['align_wrap']
        c_audit_v.font = styles['font_note']

        c_audit_w = ws.cell(row=r, column=23,
                            value=f'=IF(B{r}="","", "BASE SPECIFICATION: Published baseline scope for active item.")')
        c_audit_w.alignment = styles['align_wrap']
        c_audit_w.font = styles['font_note']

        for c in range(16, 24):
            ws.cell(row=r, column=c).border = styles['border_thin']
        
        for c in range(1, LAST_COL + 1):
            ws.cell(row=r, column=c).border = styles['border_thin']
        ws.row_dimensions[r].height = 22
        
    # Subtotal Gross Direct Cost W_gross (Row 37)
    ws.merge_cells(f'A{R_W_GROSS}:G{R_W_GROSS}')
    ws.cell(row=R_W_GROSS, column=1, value='Gross Direct Production Cost W_gross = Σ Amount (Base Specification):').font = styles['font_bold']
    ws.cell(row=R_W_GROSS, column=1).alignment = styles['align_right']
    
    c_wg = ws.cell(row=R_W_GROSS, column=8, value=f'=ROUND(SUM(H{R_RES_FIRST}:H{R_RES_LAST}), 2)')
    c_wg.font = styles['font_bold']
    c_wg.alignment = styles['align_right']
    c_wg.number_format = styles['fmt_currency']
    c_wg.fill = styles['fill_subtotal']
    
    ws.cell(row=R_W_GROSS, column=9, value='W_gross').alignment = styles['align_center']
    ws.merge_cells(f'J{R_W_GROSS}:O{R_W_GROSS}')
    ws.cell(row=R_W_GROSS, column=10, value='Direct resource build-up at full published baseline scope before contractual deductions.').font = styles['font_note']
    for c in range(1, LAST_COL + 1):
        ws.cell(row=R_W_GROSS, column=c).fill = styles['fill_subtotal']
        ws.cell(row=R_W_GROSS, column=c).border = styles['border_thin']
    ws.row_dimensions[R_W_GROSS].height = 24

    # -------------------------------------------------------------
    # SECTION 2B: SCOPE OMISSIONS, CREDITS & CONTRACTUAL DEDUCTIONS
    # -------------------------------------------------------------
    ws.merge_cells(f'A{R_SEC2B_HEAD}:O{R_SEC2B_HEAD}')
    s2b_title = ws[f'A{R_SEC2B_HEAD}']
    s2b_title.value = '2B. SCOPE OMISSIONS, CREDITS & DEDUCTIONS (CPWD Contractual Clauses & Active Levers)'
    s2b_title.font = styles['font_white_bold']
    s2b_title.fill = styles['fill_header']
    s2b_title.alignment = styles['align_left']
    ws.row_dimensions[R_SEC2B_HEAD].height = 24

    # Helper Table 2B Header in Cols P to R
    ws.merge_cells(f'P{R_SEC2B_HEAD}:R{R_SEC2B_HEAD}')
    h2b_title = ws[f'P{R_SEC2B_HEAD}']
    h2b_title.value = 'HELPER TABLE 2B: DEDUCTION RESOLVER (Calculates Table 2B Col F)'
    h2b_title.font = styles['font_white_bold']
    h2b_title.fill = styles['fill_header']
    h2b_title.alignment = styles['align_left']

    # Dedicated Deduction Rows (Rows 39 to 43)
    deduction_rows_spec = [
        {
            'line': 'D1',
            'code': '0003',
            'desc_fallback': 'Hire charges of Diesel Road Roller - 8 to 10 tonne',
            'cat': 'Plant & Machinery',
            'unit': 'day',
            'qty_formula': f'=IF(AND(OR(ISNUMBER(SEARCH("banking", $D${R_TASK})), ISNUMBER(SEARCH("rolling", $D${R_TASK}))), $D${R_OP_COMPACTION}="NO"), ROUND(-0.008 * ($D${R_BATCH_QTY}/10), 5), 0)',
            'rate_fallback': 3000.0,
            'tag': 'W',
            'norm_str': '1250.00 cum/day',
            'clause': 'CPWD DAR 2019 Item 2.4',
            'logic': 'Deduct for not rolling with power roller of min 8t: -0.008 roller-day / 10 cum.',
            'prec': f'="Lever 5: Power Rolling ($D${R_OP_COMPACTION} = " & $D${R_OP_COMPACTION} & ")"',
            'trace': f'="Batch Q=" & TEXT(D${R_BATCH_QTY}, "0.00") & " " & $D${R_BATCH_UNIT} & " | Deduct Qty=" & TEXT(F39, "0.000") & " " & E39 & " | Rate=Rs " & TEXT(G39, "0.00") & " | Amt=Rs " & TEXT(H39, "0.00")',
            'mech': 'Mechanical compaction by 8-10t roller omitted at site',
            'shift': '-0.008 roller-day reversal (1 roller consolidating 1250 cum/8-hr day)',
            'audit_w': f'="OMITTED DEDUCTION (DAR Item 2.4): 0.008 Roller @ Rs " & TEXT(G39, "0.00") & " = Rs " & TEXT(ABS(H39), "0.00") & " direct credit."'
        },
        {
            'line': 'D2',
            'code': '0113',
            'desc_fallback': 'Chowkidar',
            'cat': 'Labour',
            'unit': 'day',
            'qty_formula': f'=IF(AND(OR(ISNUMBER(SEARCH("banking", $D${R_TASK})), ISNUMBER(SEARCH("rolling", $D${R_TASK}))), $D${R_OP_COMPACTION}="NO"), ROUND(-0.008 * ($D${R_BATCH_QTY}/10), 5), 0)',
            'rate_fallback': 558.0,
            'tag': 'W',
            'norm_str': '1250.00 cum/day',
            'clause': 'CPWD DAR 2019 Item 2.4',
            'logic': 'Omission of power roller chowkidar: -0.008 chowkidar-day / 10 cum.',
            'prec': f'="Lever 5: Power Rolling ($D${R_OP_COMPACTION} = " & $D${R_OP_COMPACTION} & ")"',
            'trace': f'="Batch Q=" & TEXT(D${R_BATCH_QTY}, "0.00") & " " & $D${R_BATCH_UNIT} & " | Deduct Qty=" & TEXT(F40, "0.000") & " " & E40 & " | Rate=Rs " & TEXT(G40, "0.00") & " | Amt=Rs " & TEXT(H40, "0.00")',
            'mech': 'Equipment security guard omitted alongside road roller',
            'shift': '-0.008 chowkidar-day reversal (1 guard per roller shift)',
            'audit_w': f'="OMITTED DEDUCTION (DAR Item 2.4): 0.008 Chowkidar @ Rs " & TEXT(G40, "0.00") & " = Rs " & TEXT(ABS(H40), "0.00") & " direct credit."'
        },
        {
            'line': 'D3',
            'code': '9999',
            'desc_fallback': 'Sundries',
            'cat': 'Sundries',
            'unit': 'L.S.',
            'qty_formula': f'=IF(AND(OR(ISNUMBER(SEARCH("banking", $D${R_TASK})), ISNUMBER(SEARCH("rolling", $D${R_TASK}))), $D${R_OP_COMPACTION}="NO"), ROUND(-1.820 * ($D${R_BATCH_QTY}/10), 5), 0)',
            'rate_fallback': 2.0,
            'tag': 'W',
            'norm_str': 'Fixed allowance',
            'clause': 'CPWD DAR 2019 Item 2.4',
            'logic': 'Sundries allowance deduction for power roller: -1.82 L.S. / 10 cum.',
            'prec': f'="Lever 5: Power Rolling ($D${R_OP_COMPACTION} = " & $D${R_OP_COMPACTION} & ")"',
            'trace': f'="Batch Q=" & TEXT(D${R_BATCH_QTY}, "0.00") & " " & $D${R_BATCH_UNIT} & " | Deduct Qty=" & TEXT(F41, "0.000") & " " & E41 & " | Rate=Rs " & TEXT(G41, "0.00") & " | Amt=Rs " & TEXT(H41, "0.00")',
            'mech': 'Lubricants, flags, markers and minor consumables for roller',
            'shift': '-1.82 L.S. allowance deduction per CPWD DAR Item 2.4 standard',
            'audit_w': f'="OMITTED DEDUCTION (DAR Item 2.4): 1.82 Sundries = Rs " & TEXT(ABS(H41), "0.00") & " direct credit."'
        },
        {
            'line': 'D4',
            'code': '0101',
            'desc_fallback': 'Bhisti',
            'cat': 'Labour',
            'unit': 'day',
            'qty_formula': f'=IF(AND(OR(ISNUMBER(SEARCH("banking", $D${R_TASK})), ISNUMBER(SEARCH("rolling", $D${R_TASK})), ISNUMBER(SEARCH("Embankment", $D${R_SCOPE}))), $D${R_OP_WATERING}="NO"), ROUND(-0.400 * ($D${R_BATCH_QTY}/10), 5), 0)',
            'rate_fallback': 617.0,
            'tag': 'W',
            'norm_str': '25.00 cum/day',
            'clause': 'CPWD DAR 2019 Item 2.5',
            'logic': 'Deduct for not watering excavated earth for banking: -0.40 bhisti-day / 10 cum.',
            'prec': f'="Lever 4: Compaction Watering ($D${R_OP_WATERING} = " & $D${R_OP_WATERING} & ")"',
            'trace': f'="Batch Q=" & TEXT(D${R_BATCH_QTY}, "0.00") & " " & $D${R_BATCH_UNIT} & " | Deduct Qty=" & TEXT(F42, "0.000") & " " & E42 & " | Rate=Rs " & TEXT(G42, "0.00") & " | Amt=Rs " & TEXT(H42, "0.00")',
            'mech': 'Watering loose layers to achieve Optimum Moisture Content (OMC) omitted',
            'shift': '-0.40 bhisti-day reversal (1 bhisti watering 25 cum/8-hr day)',
            'audit_w': f'="OMITTED DEDUCTION (DAR Item 2.5): 0.40 Bhishti @ Rs " & TEXT(G42, "0.00") & " = Rs " & TEXT(ABS(H42), "0.00") & " direct credit."'
        },
        {
            'line': 'L1',
            'code': f'=IF(OR(ISNUMBER(SEARCH("Extra lift", $D$9)), ISNUMBER(SEARCH("Extra for pipe trench", $D$9))), "", IF(ISNUMBER(SEARCH("1.5 m to 3.0 m", $D$20)), IF($D$11="m", "2.11", "2.26.1"), IF(ISNUMBER(SEARCH("3.0 m to 4.5 m", $D$20)), IF($D$11="m", "2.12", "2.26.1"), "")))',
            'desc_fallback': 'Extra Depth / Lift Stage Allowance',
            'cat': 'Extra Depth Lift',
            'unit': 'day',
            'qty_formula': f'=IF(OR(B43="", ISNUMBER(SEARCH("Extra lift", $D$9)), ISNUMBER(SEARCH("Extra for pipe trench", $D$9))), 0, IF(ISNUMBER(SEARCH("1.5 m to 3.0 m", $D$20)), IF($D$11="m", 1.0, ROUND(1.0 * ($D$10/10), 4)), IF(ISNUMBER(SEARCH("3.0 m to 4.5 m", $D$20)), IF($D$11="m", 1.0, ROUND(2.0 * ($D$10/10), 4)), 0)))',
            'rate_fallback': 90.40,
            'tag': 'W',
            'norm_str': '1.5 m lift stage',
            'clause': 'CPWD DAR 2.26 / 2.11 / 2.12',
            'logic': 'Extra depth lift allowance beyond standard 1.5 m depth limit.',
            'prec': f'="Lever 2: Depth & Lift Stage (" & $D$20 & ")"',
            'trace': f'="Batch Q=" & TEXT($D$10, "0.00") & " " & $D$11 & " | Extra Depth Qty=" & TEXT(F43, "0.000") & " | Rate=Rs " & TEXT(G43, "0.00") & " | Amt=Rs " & TEXT(H43, "0.00")',
            'mech': 'Manual relay lifting and staging for deep trench excavation beyond 1.5 m limit.',
            'shift': '1.5 m lift stage allocation per CPWD DAR Item 2.26 / 2.11 / 2.12.',
            'audit_w': f'="EXTRA DEPTH LIFT: Added Rs " & TEXT(H43, "0.00") & " for depth exceeding 1.5 m."'
        }
    ]

    for d_idx, dspec in enumerate(deduction_rows_spec):
        dr = R_DED_FIRST + d_idx
        
        ws.cell(row=dr, column=1, value=dspec['line']).alignment = styles['align_center']
        ws.cell(row=dr, column=1).font = styles['font_bold']

        c_dcode = ws.cell(row=dr, column=2, value=dspec['code'])
        c_dcode.alignment = styles['align_center']
        c_dcode.font = styles['font_bold']
        c_dcode.fill = styles['fill_lookup']

        if dr == 43:
            c_dname = ws.cell(row=dr, column=3,
                              value=f'=IF(B{dr}="","", IF(B{dr}="2.11", "Extra for excavating pipe trenches depth 1.5m to 3.0m (CPWD DAR 2.11)", IF(B{dr}="2.12", "Extra for excavating pipe trenches depth 3.0m to 4.5m (CPWD DAR 2.12)", IF(ISNUMBER(SEARCH("3.0 m to 4.5 m", $D$20)), "Extra lift for additional depth 3.0m to 4.5m (2 stages x CPWD DAR 2.26.1)", "Extra lift for additional depth 1.5m to 3.0m (1 stage x CPWD DAR 2.26.1)"))))')
            c_dunit = ws.cell(row=dr, column=5,
                              value=f'=IF(B{dr}="","", IF(B{dr}="2.26.1", "10 cum", "metre"))')
            c_drate = ws.cell(row=dr, column=7,
                              value=f'=IF(B{dr}="","", IF(B{dr}="2.11", 127.00, IF(B{dr}="2.12", 315.05, 90.40)))')
            c_dtag = ws.cell(row=dr, column=9,
                             value=f'=IF(B{dr}="","", IF(OR(B{dr}="2.11", B{dr}="2.12"), "A", "W"))')
        else:
            fb_str = f'"{dspec["desc_fallback"]}"'
            c_dname = ws.cell(row=dr, column=3,
                              value=f'=IF(B{dr}="","", IFERROR(INDEX(Rates_Master!$C:$C, MATCH(B{dr}, Rates_Master!$A:$A, 0)), {fb_str}))')
            c_dunit = ws.cell(row=dr, column=5, value=dspec['unit'])
            c_drate = ws.cell(row=dr, column=7,
                              value=f'=IF(B{dr}="","", IFERROR(INDEX(Rates_Master!$E:$E, MATCH(B{dr}, Rates_Master!$A:$A, 0)), {dspec["rate_fallback"]}))')
            c_dtag = ws.cell(row=dr, column=9, value=dspec['tag'])
        c_dname.alignment = styles['align_left']
        c_dname.font = styles['font_regular']

        c_dcat = ws.cell(row=dr, column=4, value=dspec['cat'])
        c_dcat.alignment = styles['align_center']
        c_dcat.font = styles['font_note']

        c_dunit.alignment = styles['align_center']
        c_dunit.font = styles['font_regular']

        # Helper 2B (Cols P to R)
        c_dbase_q = ws.cell(row=dr, column=16, value=0)
        c_dbase_q.alignment = styles['align_right']
        c_dbase_q.font = styles['font_regular']
        c_dbase_q.fill = styles['fill_calc']
        c_dbase_q.number_format = '0.000'

        c_dfac_q = ws.cell(row=dr, column=17, value='DEDUCTION' if dr < 43 else 'EXTRA DEPTH')
        c_dfac_q.alignment = styles['align_center']
        c_dfac_q.font = styles['font_note']
        c_dfac_q.fill = styles['fill_lookup']

        c_dnet_q = ws.cell(row=dr, column=18, value=f'=F{dr}')
        c_dnet_q.alignment = styles['align_right']
        c_dnet_q.font = styles['font_bold']
        c_dnet_q.fill = styles['fill_result']
        c_dnet_q.number_format = '0.000'

        # Norm Input Qty in Col F (negative or positive or zero)
        c_dqty = ws.cell(row=dr, column=6, value=dspec['qty_formula'])
        c_dqty.alignment = styles['align_right']
        c_dqty.font = styles['font_bold']
        c_dqty.fill = styles['fill_input']
        c_dqty.number_format = '0.000;[Red]-0.000;0.000'

        # Rate in Col G
        c_drate.alignment = styles['align_right']
        c_drate.font = styles['font_regular']
        c_drate.fill = styles['fill_lookup']
        c_drate.number_format = styles['fmt_currency']

        # Line Amount in Col H (=ROUND(F*G, 2))
        c_damt = ws.cell(row=dr, column=8,
                         value=f'=IF(OR(B{dr}="", F{dr}=0), 0, ROUND(F{dr} * G{dr}, 2))')
        c_damt.alignment = styles['align_right']
        c_damt.font = styles['font_bold']
        c_damt.fill = styles['fill_calc']
        c_damt.number_format = '₹#,##0.00;[Red]-₹#,##0.00;₹0.00'

        c_dtag.alignment = styles['align_center']
        c_dtag.font = styles['font_bold']
        c_dtag.fill = styles['fill_lookup']


        # Cost Share % against Gross Cost
        c_dpct = ws.cell(row=dr, column=10,
                         value=f'=IF(H${R_W_GROSS}>0, ROUND(H{dr} / H${R_W_GROSS} * 100, 1), 0)')
        c_dpct.alignment = styles['align_right']
        c_dpct.font = styles['font_regular']
        c_dpct.fill = styles['fill_calc']
        c_dpct.number_format = '0.0"%"'

        c_dprod = ws.cell(row=dr, column=11, value=dspec['norm_str'])
        c_dprod.alignment = styles['align_right']
        c_dprod.font = styles['font_note']

        c_dse = ws.cell(row=dr, column=12, value='—')
        c_dse.alignment = styles['align_center']
        c_dse.font = styles['font_note']

        c_dpace = ws.cell(row=dr, column=13, value='—')
        c_dpace.alignment = styles['align_center']
        c_dpace.font = styles['font_note']

        c_dclause = ws.cell(row=dr, column=14, value=dspec['clause'])
        c_dclause.alignment = styles['align_left']
        c_dclause.font = styles['font_note']

        c_dlogic = ws.cell(row=dr, column=15, value=dspec['logic'])
        c_dlogic.alignment = styles['align_left']
        c_dlogic.font = styles['font_note']

        # Audit Columns S to W
        c_daudit_s = ws.cell(row=dr, column=19, value=dspec['prec'])
        c_daudit_s.alignment = styles['align_left']
        c_daudit_s.font = styles['font_note']

        c_daudit_t = ws.cell(row=dr, column=20, value=dspec['trace'])
        c_daudit_t.alignment = styles['align_left']
        c_daudit_t.font = styles['font_note']

        c_daudit_u = ws.cell(row=dr, column=21, value=dspec['mech'])
        c_daudit_u.alignment = styles['align_left']
        c_daudit_u.font = styles['font_note']

        c_daudit_v = ws.cell(row=dr, column=22, value=dspec['shift'])
        c_daudit_v.alignment = styles['align_left']
        c_daudit_v.font = styles['font_note']

        c_daudit_w = ws.cell(row=dr, column=23, value=dspec['audit_w'])
        c_daudit_w.alignment = styles['align_left']
        c_daudit_w.font = styles['font_note']

        for c in range(16, 24):
            ws.cell(row=dr, column=c).border = styles['border_thin']
        
        for c in range(1, LAST_COL + 1):
            ws.cell(row=dr, column=c).border = styles['border_thin']
        ws.row_dimensions[dr].height = 22

    # Subtotal Scope Deductions W_deduct (Row 44)
    ws.merge_cells(f'A{R_W_DEDUCT}:G{R_W_DEDUCT}')
    ws.cell(row=R_W_DEDUCT, column=1, value='Total Scope Deductions W_deduct = Σ Deductions & Scope Credits:').font = styles['font_bold']
    ws.cell(row=R_W_DEDUCT, column=1).alignment = styles['align_right']
    
    c_wd = ws.cell(row=R_W_DEDUCT, column=8, value=f'=ROUND(SUM(H{R_DED_FIRST}:H{R_DED_LAST}), 2)')
    c_wd.font = styles['font_bold']
    c_wd.alignment = styles['align_right']
    c_wd.number_format = '₹#,##0.00;[Red]-₹#,##0.00;₹0.00'
    c_wd.fill = styles['fill_subtotal']
    
    ws.cell(row=R_W_DEDUCT, column=9, value='W_ded').alignment = styles['align_center']
    ws.merge_cells(f'J{R_W_DEDUCT}:O{R_W_DEDUCT}')
    ws.cell(row=R_W_DEDUCT, column=10, value='Contractual deductions applied via Section 1B operational levers (e.g. DAR Items 2.4 & 2.5).').font = styles['font_note']
    for c in range(1, LAST_COL + 1):
        ws.cell(row=R_W_DEDUCT, column=c).fill = styles['fill_subtotal']
        ws.cell(row=R_W_DEDUCT, column=c).border = styles['border_thin']
    ws.row_dimensions[R_W_DEDUCT].height = 24

    # Net Direct Production Cost W (Row 45)
    ws.merge_cells(f'A{R_W_SUB}:G{R_W_SUB}')
    ws.cell(row=R_W_SUB, column=1, value='Net Direct Production Cost W = W_gross + W_deduct (all active resources):').font = styles['font_bold']
    ws.cell(row=R_W_SUB, column=1).alignment = styles['align_right']
    
    c_w = ws.cell(row=R_W_SUB, column=8, value=f'=ROUND(H{R_W_GROSS} + H{R_W_DEDUCT}, 2)')
    c_w.font = styles['font_bold']
    c_w.alignment = styles['align_right']
    c_w.number_format = styles['fmt_currency']
    c_w.fill = styles['fill_subtotal']
    
    ws.cell(row=R_W_SUB, column=9, value='W').alignment = styles['align_center']
    ws.merge_cells(f'J{R_W_SUB}:O{R_W_SUB}')
    ws.cell(row=R_W_SUB, column=10, value='Base for statutory markup calculations before exclusion of pre-marked reference items.').font = styles['font_note']
    for c in range(1, LAST_COL + 1):
        ws.cell(row=R_W_SUB, column=c).fill = styles['fill_subtotal']
        ws.cell(row=R_W_SUB, column=c).border = styles['border_thin']
    ws.row_dimensions[R_W_SUB].height = 24
    
    # A-tagged Total (Row 46)
    ws.merge_cells(f'A{R_A_SUB}:G{R_A_SUB}')
    ws.cell(row=R_A_SUB, column=1, value='A-tagged Total (pre-marked composite reference items excluded from markups):').font = styles['font_bold']
    ws.cell(row=R_A_SUB, column=1).alignment = styles['align_right']
    
    c_a = ws.cell(row=R_A_SUB, column=8, value=f'=ROUND(SUMIF(I{R_RES_FIRST}:I{R_RES_LAST}, "A", H{R_RES_FIRST}:H{R_RES_LAST}) + SUMIF(I{R_DED_FIRST}:I{R_DED_LAST}, "A", H{R_DED_FIRST}:H{R_DED_LAST}), 2)')
    c_a.font = styles['font_bold']
    c_a.alignment = styles['align_right']
    c_a.number_format = styles['fmt_currency']
    c_a.fill = styles['fill_subtotal']
    
    ws.cell(row=R_A_SUB, column=9, value='A').alignment = styles['align_center']
    ws.merge_cells(f'J{R_A_SUB}:O{R_A_SUB}')
    ws.cell(row=R_A_SUB, column=10, value='Uses robust SUMIF across Table 2A and 2B to guarantee error-free evaluation even with empty slots.').font = styles['font_note']
    for c in range(1, LAST_COL + 1):
        ws.cell(row=R_A_SUB, column=c).fill = styles['fill_subtotal']
        ws.cell(row=R_A_SUB, column=c).border = styles['border_double_bottom']
    ws.row_dimensions[R_A_SUB].height = 24

    # Blank separation row 47
    ws.row_dimensions[R_BLANK_SEP].height = 12

    # -------------------------------------------------------------
    # SECTION 3: STATUTORY MARKUP CHAIN
    # -------------------------------------------------------------
    ws.merge_cells(f'A{R_SEC3_HEAD}:O{R_SEC3_HEAD}')
    s3_title = ws[f'A{R_SEC3_HEAD}']
    s3_title.value = '3. STATUTORY MARKUPS & FINAL RATE DERIVATION (CPWD DAR Standard Compounded Markup Chain)'
    s3_title.font = styles['font_white_bold']
    s3_title.fill = styles['fill_header']
    s3_title.alignment = styles['align_left']
    ws.row_dimensions[R_SEC3_HEAD].height = 24
    
    s3_cols = ['Step', 'Component / Markup Description', 'Apply?', 'Basis Applied', 'Base Amount (₹)', 'Factor / %', 'Added (₹)', 'Running Total (₹)', 'Cell Role', 'CPWD Statutory Guidance Rule & Practical Field Note', '', '', '', '', '']
    for ci, h in enumerate(s3_cols, 1):
        c = ws.cell(row=R_SEC3_COLS, column=ci, value=h)
        c.font = styles['font_header']
        c.fill = styles['fill_header']
        c.alignment = styles['align_center']
        c.border = styles['border_header']
    ws.merge_cells(f'J{R_SEC3_COLS}:O{R_SEC3_COLS}')
    ws.row_dimensions[R_SEC3_COLS].height = 24
    
    mu_chain = [
        (R_W, 'W', 'Direct Production Cost W (all resources)', '—', 'Direct Sum',
         f'=H{R_W_SUB}', '—', '=0', f'=H{R_W_SUB}', 'DERIVED',
         'Sum of all direct labour, machinery and sundries from Section 2.'),
        (R_X1, 'X1', 'Add: Water Charges (1% on W-A)', 'YES', 'On (W - A)',
         f'=H{R_W} - H{R_A_SUB}', '=Factor_Water', f'=IF(C{R_X1}="YES", ROUND(E{R_X1} * F{R_X1}, 2), 0)', f'=H{R_W} + G{R_X1}', 'INPUT',
         'CPWD standard: 1% for site compaction watering, trench dampening and dust control. Live-linked to Factor_Water.'),
        (R_X, 'X', 'Subtotal X (W + Water Charges)', '—', 'W + Water',
         f'=H{R_X1}', '—', '=0', f'=H{R_X1}', 'DERIVED',
         'Compounded base passed to Works Contract GST calculation.'),
        (R_Y1, 'Y1', 'Add: GST on Works Contract (14.05% on X-A)', 'YES', 'On (X - A)',
         f'=H{R_X} - H{R_A_SUB}', '=Factor_GST', f'=IF(C{R_Y1}="YES", ROUND(E{R_Y1} * F{R_Y1}, 2), 0)', f'=H{R_X} + G{R_Y1}', 'INPUT',
         'CPWD DAR statutory works contract tax factor (0.1405). Live-linked to Factor_GST on Global_Factors.'),
        (R_Y, 'Y', 'Subtotal Y (X + GST)', '—', 'X + GST',
         f'=H{R_Y1}', '—', '=0', f'=H{R_Y1}', 'DERIVED',
         'Compounded base passed to Contractor Profit & Overheads calculation.'),
        (R_Z1, 'Z1', 'Add: Contractor Profit & Overheads (15% CPOH on Y-A)', 'YES', 'On (Y - A)',
         f'=H{R_Y} - H{R_A_SUB}', '=Factor_CPOH', f'=IF(C{R_Z1}="YES", ROUND(E{R_Z1} * F{R_Z1}, 2), 0)', f'=H{R_Y} + G{R_Z1}', 'INPUT',
         'Standard 15% allowance for contractor overheads, plant depreciation and profit. Live-linked to Factor_CPOH.'),
        (R_Z, 'Z', 'Subtotal Z (Y + CPOH)', '—', 'Y + CPOH',
         f'=H{R_Z1}', '—', '=0', f'=H{R_Z1}', 'DERIVED',
         'Compounded base passed to Building and Other Construction Workers Welfare Cess calculation.'),
        (R_Z2, 'Z2', 'Add: Building and Other Construction Workers Welfare Cess (1% on Z - A)', 'YES', 'On (Z - A)',
         f'=H{R_Z} - H{R_A_SUB}', '=Factor_Cess', f'=IF(C{R_Z2}="YES", ROUND(E{R_Z2} * F{R_Z2}, 2), 0)', f'=H{R_Z} + G{R_Z2}', 'INPUT',
         'Statutory 1% BOCW Welfare Cess levied on overall cost excluding pre-marked composite items. Live-linked to Factor_Cess.'),
        (R_COST, 'Cost', f'="Cost for " & D{R_BATCH_QTY} & " " & D{R_BATCH_UNIT}', '—', 'Standard Output Batch',
         f'=H{R_Z2}', '—', '=0', f'=H{R_Z2}', 'RESULT',
         'Fully compounded cost for the standard DAR batch quantity before unit-rate derivation.'),
        (R_RATE, 'Rate', f'="Rate per " & D{R_BATCH_UNIT} & " (unrounded)"', '—', f'=D{R_BATCH_UNIT}',
         f'=H{R_COST} / D{R_BATCH_QTY}', '—', '=0', f'=H{R_COST} / D{R_BATCH_QTY}', 'DERIVED',
         'Exact rate per unit output (Cost / Batch Qty) carried into statutory rounding.'),
        (R_SAY, 'SAY', f'="FINAL ANALYSED SAY RATE (Rs per " & D{R_BATCH_UNIT} & ")"', '—', 'MROUND 0.05',
         f'=H{R_RATE}', '—', '=0', f'=MROUND(H{R_RATE}, 0.05)', 'SAY',
         'Final published billing rate rounded to the nearest 5 paise in accordance with CPWD DAR Section 0 conventions.'),
        (R_DAR_REF, 'REF', f'="Published Benchmark Say Rate (Rs per " & D{R_BATCH_UNIT} & "):"', '—', 'Benchmark',
         f'=IF(ISNUMBER(D{R_DAR_PUB}), D{R_DAR_PUB}, "-")', '—', '=0', f'=IF(ISNUMBER(D{R_DAR_PUB}), D{R_DAR_PUB}, "-")', 'LOOKUP',
         'Official CPWD DAR 2019 Volume 1 published rate for verification.'),
        (R_DAR_DIFF, 'DIFF', 'Audit Variance (Analysed Say - Published Say):', '—', 'Say - Ref',
         f'=IF(ISNUMBER(D{R_DAR_PUB}), H{R_SAY} - H{R_DAR_REF}, 0)', '—', '=0', f'=IF(ISNUMBER(D{R_DAR_PUB}), ROUND(H{R_SAY} - H{R_DAR_REF}, 2), 0)', 'RESULT',
         'Discrepancy audit against published benchmark. Target = ₹ 0.00.')

    ]
    
    for r, step, desc, apply, basis, base, fac, added, tot, role, note in mu_chain:
        ws.cell(row=r, column=1, value=step).alignment = styles['align_center']
        ws.cell(row=r, column=1).font = styles['font_bold']
        
        ws.cell(row=r, column=2, value=desc).alignment = styles['align_left']
        ws.cell(row=r, column=2).font = styles['font_bold'] if role in ('RESULT', 'SAY') else styles['font_regular']
        
        c_apply = ws.cell(row=r, column=3, value=apply)
        c_apply.alignment = styles['align_center']
        if apply in ('YES', 'NO'):
            dv_yesno.add(c_apply)
            c_apply.fill = styles['fill_input']
            c_apply.font = styles['font_bold']
            c_apply.protection = Protection(locked=False)
            
        ws.cell(row=r, column=4, value=basis).alignment = styles['align_center']
        
        c_base = ws.cell(row=r, column=5, value=base)
        c_base.alignment = styles['align_right']
        c_base.number_format = styles['fmt_currency']
        
        c_fac = ws.cell(row=r, column=6, value=fac)
        c_fac.alignment = styles['align_right']
        if r in (R_X1, R_Y1, R_Z1, R_Z2):
            c_fac.number_format = styles['fmt_percent']
            c_fac.fill = styles['fill_lookup']
            
        c_add = ws.cell(row=r, column=7, value=added)
        c_add.alignment = styles['align_right']
        c_add.number_format = styles['fmt_currency']
        
        c_tot = ws.cell(row=r, column=8, value=tot)
        c_tot.alignment = styles['align_right']
        c_tot.number_format = styles['fmt_currency']
        c_tot.font = styles['font_say'] if role == 'SAY' else (styles['font_bold'] if role == 'RESULT' else styles['font_regular'])
        
        c_role = ws.cell(row=r, column=9, value=role)
        c_role.alignment = styles['align_center']
        c_role.font = styles['font_bold']
        
        ws.merge_cells(f'J{r}:O{r}')
        c_note = ws.cell(row=r, column=10, value=note)
        c_note.font = styles['font_note']
        c_note.alignment = styles['align_wrap']
        
        if role == 'SAY':
            ws.row_dimensions[r].height = 30
            for c in range(1, LAST_COL + 1):
                ws.cell(row=r, column=c).fill = styles['fill_say']
                ws.cell(row=r, column=c).border = styles['border_thin']
        elif role == 'RESULT':
            ws.row_dimensions[r].height = 24
            for c in range(1, LAST_COL + 1):
                ws.cell(row=r, column=c).fill = styles['fill_result']
                ws.cell(row=r, column=c).border = styles['border_thin']
        elif r in (R_X, R_Y, R_Z):
            ws.row_dimensions[r].height = 22
            for c in range(1, LAST_COL + 1):
                ws.cell(row=r, column=c).fill = styles['fill_subtotal']
                ws.cell(row=r, column=c).border = styles['border_thin']
        else:
            ws.row_dimensions[r].height = 22
            for c in range(1, LAST_COL + 1):
                ws.cell(row=r, column=c).border = styles['border_thin']

    # -------------------------------------------------------------
    # SECTION 4: FIRST-PRINCIPLES LABOUR & MACHINERY BENCHMARK MATRIX (TABLE 4)
    # -------------------------------------------------------------
    ws.merge_cells(f'A{R_SEC4_HEAD}:O{R_SEC4_HEAD}')
    s4_title = ws[f'A{R_SEC4_HEAD}']
    s4_title.value = '4. FIRST-PRINCIPLES LABOUR & MACHINERY PRODUCTIVITY BENCHMARK MATRIX (8-Hour Working Shift Basis)'
    s4_title.font = styles['font_white_bold']
    s4_title.fill = styles['fill_header']
    s4_title.alignment = styles['align_left']
    ws.row_dimensions[R_SEC4_HEAD].height = 24
    
    ws.merge_cells(f'A{R_SEC4_BANNER}:O{R_SEC4_BANNER}')
    s4_banner = ws[f'A{R_SEC4_BANNER}']
    s4_banner.value = (
        'FIRST-PRINCIPLES DERIVATION MECHANICS: Shows how physical labour and equipment operations in an 8-hour shift translate into CPWD norm factors. '
        'Standard Gang (G) × Task Duration (H) = Total Man-Hours (MH) → Day Coeff (D) = MH / 8.00 h → '
        'Daily Output (P) = Batch Q / D → Specific Effort (SE) = MH / Q → Hourly Pace = P / 8.00. '
        'Every calculation cell is a live Excel formula. No CPWD code numbering; Resource Code is placed in its own dedicated column (Col C).'
    )
    s4_banner.font = styles['font_note']
    s4_banner.fill = styles['fill_note']
    s4_banner.alignment = styles['align_wrap']
    ws.row_dimensions[R_SEC4_BANNER].height = 36
    
    bm_cols = [
        'Civil Work Scope / Operation', 'Operation / Sub-Task', 'Resource Code', 'Resource Trade / Fleet Role',
        'Batch Qty (Q)', 'Unit', 'Gang Size (G)', 'Duration (H hrs)',
        'Total Man-Hours (MH)', 'Shift (h)', 'Day Coeff (D)', 'Daily Output (P)',
        'Unit Effort (SE)', 'Hourly Pace', 'Engineering Task Breakdown & First-Principles 8-Hour Logic'
    ]
    for ci, h in enumerate(bm_cols, 1):
        c = ws.cell(row=R_SEC4_COLS, column=ci, value=h)
        c.font = styles['font_header']
        c.fill = styles['fill_header']
        c.alignment = styles['align_center']
        c.border = styles['border_header']
    ws.row_dimensions[R_SEC4_COLS].height = 24
    
    for idx, row_data in enumerate(BENCHMARK_PRODUCTIVITY_ROWS):
        r = R_BM_FIRST + idx
        
        ws.cell(row=r, column=1, value=row_data['scope']).alignment = styles['align_left']
        ws.cell(row=r, column=1).font = styles['font_bold']
        
        ws.cell(row=r, column=2, value=row_data['task']).alignment = styles['align_left']
        
        # Dedicated Resource Code in Col C
        c_rcode = ws.cell(row=r, column=3, value=row_data['res_code'])
        c_rcode.alignment = styles['align_center']
        c_rcode.font = styles['font_bold']
        c_rcode.fill = styles['fill_lookup']
        
        ws.cell(row=r, column=4, value=row_data['trade']).alignment = styles['align_left']
        ws.cell(row=r, column=4).font = styles['font_bold']
        
        c_q = ws.cell(row=r, column=5, value=row_data['qty'])
        c_q.alignment = styles['align_right']
        c_q.number_format = '0.00'
        c_q.fill = styles['fill_lookup']
        
        ws.cell(row=r, column=6, value=row_data['unit']).alignment = styles['align_center']
        
        c_g = ws.cell(row=r, column=7, value=row_data['gang'])
        c_g.alignment = styles['align_right']
        c_g.number_format = '0'
        c_g.fill = styles['fill_input']
        c_g.protection = Protection(locked=False)
        
        c_h = ws.cell(row=r, column=8, value=row_data['hours'])
        c_h.alignment = styles['align_right']
        c_h.number_format = '0.000'
        c_h.fill = styles['fill_input']
        c_h.protection = Protection(locked=False)
        
        c_mh = ws.cell(row=r, column=9, value=f'=ROUND(G{r} * H{r}, 3)')
        c_mh.alignment = styles['align_right']
        c_mh.number_format = '0.000'
        c_mh.fill = styles['fill_calc']
        c_mh.font = styles['font_bold']
        
        c_shift = ws.cell(row=r, column=10, value=8.00)
        c_shift.alignment = styles['align_center']
        c_shift.number_format = '0.00'
        c_shift.fill = styles['fill_lookup']
        
        c_d = ws.cell(row=r, column=11, value=f'=ROUND(I{r} / J{r}, 4)')
        c_d.alignment = styles['align_right']
        c_d.number_format = '0.000'
        c_d.fill = styles['fill_calc']
        c_d.font = styles['font_bold']
        
        c_p = ws.cell(row=r, column=12, value=f'=IFERROR(ROUND(E{r} / K{r}, 2), "")')
        c_p.alignment = styles['align_right']
        c_p.number_format = '0.00'
        c_p.fill = styles['fill_calc']
        
        c_se = ws.cell(row=r, column=13, value=f'=IFERROR(ROUND(I{r} / E{r}, 4), "")')
        c_se.alignment = styles['align_right']
        c_se.number_format = '0.0000'
        c_se.fill = styles['fill_calc']
        
        c_pace = ws.cell(row=r, column=14, value=f'=IFERROR(ROUND(L{r} / J{r}, 2), "")')
        c_pace.alignment = styles['align_right']
        c_pace.number_format = '0.00'
        c_pace.fill = styles['fill_calc']
        
        c_lg = ws.cell(row=r, column=15, value=row_data['logic'])
        c_lg.alignment = styles['align_left']
        c_lg.font = styles['font_note']
        
        for c in range(1, LAST_COL + 1):
            ws.cell(row=r, column=c).border = styles['border_thin']
        ws.row_dimensions[r].height = 24

    # -------------------------------------------------------------
    # SECTION 5: RUNNING CUSTOM NON-DSR ITEMS LIBRARY
    # -------------------------------------------------------------
    ws.merge_cells(f'A{R_SEC5_HEAD}:O{R_SEC5_HEAD}')
    s5_title = ws[f'A{R_SEC5_HEAD}']
    s5_title.value = '5. RUNNING CUSTOM NON-DSR ITEMS LIBRARY FOR EARTH WORK'
    s5_title.font = styles['font_white_bold']
    s5_title.fill = styles['fill_header']
    s5_title.alignment = styles['align_left']
    ws.row_dimensions[R_SEC5_HEAD].height = 24
    
    ws.merge_cells(f'A{R_SEC5_BANNER}:O{R_SEC5_BANNER}')
    s5_banner = ws[f'A{R_SEC5_BANNER}']
    s5_banner.value = 'Log custom non-DSR earthwork rate analyses below. Results generated above can be archived here for tender schedules.'
    s5_banner.font = styles['font_note']
    s5_banner.fill = styles['fill_note']
    s5_banner.alignment = styles['align_wrap']
    ws.row_dimensions[R_SEC5_BANNER].height = 20
    
    lib_cols = ['Custom Code', 'Item Nomenclature / Scope', 'Unit', 'Batch', 'Mode', 'CPOH Base', 'Direct Cost W (₹)', 'Unit Rate (₹)', 'Say Rate (₹)', 'Status', 'Technical Basis & Remarks', '', '', '', '']
    for ci, h in enumerate(lib_cols, 1):
        c = ws.cell(row=R_SEC5_COLS, column=ci, value=h)
        c.font = styles['font_header']
        c.fill = styles['fill_header']
        c.alignment = styles['align_center']
        c.border = styles['border_header']
    ws.merge_cells(f'K{R_SEC5_COLS}:O{R_SEC5_COLS}')
    ws.row_dimensions[R_SEC5_COLS].height = 24
    
    sample_custom = [
        ('C-02.01', 'Earth work surface excavation in soil ≤30cm depth (DAR baseline)', 'sqm', 100, 'DIRECT', 'W', 6919.20, 92.55, 92.55, 'COMPLETE', 'Standard DAR surface excavation benchmark'),
        ('C-02.02', 'Earth work rough excavation, banking, watering & rolling (DAR baseline)', 'cum', 10, 'DIRECT', 'W', 5581.72, 746.80, 746.80, 'COMPLETE', 'Standard DAR embankment benchmark'),
        ('C-02.03', 'Hydraulic excavator 0.9 cum area excavation (DAR baseline)', 'cum', 10, 'DIRECT', 'W', 1358.00, 181.85, 181.85, 'COMPLETE', 'Standard DAR excavator benchmark'),
        ('C-02.04', '(Available for custom earthwork specification)', 'cum', 10, 'DIRECT', 'W', None, None, None, 'DRAFT', 'Custom item slot'),
        ('C-02.05', '(Available for custom earthwork specification)', 'cum', 10, 'DIRECT', 'W', None, None, None, 'DRAFT', 'Custom item slot'),
        ('C-02.06', '(Available for custom earthwork specification)', 'sqm', 100, 'DIRECT', 'W', None, None, None, 'DRAFT', 'Custom item slot'),
        ('C-02.07', '(Available for custom earthwork specification)', 'cum', 10, 'DIRECT', 'W', None, None, None, 'DRAFT', 'Custom item slot'),
        ('C-02.08', '(Available for custom earthwork specification)', 'm', 100, 'DIRECT', 'W', None, None, None, 'DRAFT', 'Custom item slot'),
    ]
    for idx, (ccode, cdesc, cunit, cqty, cmode, ccpoh, cw, crate, csay, cstat, crem) in enumerate(sample_custom):
        r = R_LIB_FIRST + idx
        ws.cell(row=r, column=1, value=ccode).alignment = styles['align_center']
        ws.cell(row=r, column=2, value=cdesc).alignment = styles['align_left']
        ws.cell(row=r, column=3, value=cunit).alignment = styles['align_center']
        ws.cell(row=r, column=4, value=cqty).alignment = styles['align_right']
        ws.cell(row=r, column=5, value=cmode).alignment = styles['align_center']
        ws.cell(row=r, column=6, value=ccpoh).alignment = styles['align_center']
        
        c_w = ws.cell(row=r, column=7, value=cw)
        c_w.alignment = styles['align_right']
        c_w.number_format = styles['fmt_currency']
        
        c_r = ws.cell(row=r, column=8, value=crate)
        c_r.alignment = styles['align_right']
        c_r.number_format = styles['fmt_currency']
        
        c_s = ws.cell(row=r, column=9, value=csay)
        c_s.alignment = styles['align_right']
        c_s.number_format = styles['fmt_currency']
        c_s.font = styles['font_bold']
        
        ws.cell(row=r, column=10, value=cstat).alignment = styles['align_center']
        
        ws.merge_cells(f'K{r}:O{r}')
        ws.cell(row=r, column=11, value=crem).alignment = styles['align_left']
        ws.cell(row=r, column=11).font = styles['font_note']
        
        for c in range(1, LAST_COL + 1):
            ws.cell(row=r, column=c).border = styles['border_thin']
        ws.row_dimensions[r].height = 20

    # -------------------------------------------------------------
    # SECTION 6: TABLE 5E — GROUND-TRUTH RESOURCE NORMS CATALOG (178 records)
    # -------------------------------------------------------------
    ws.merge_cells(f'A{R_SEC6_HEAD}:Q{R_SEC6_HEAD}')
    s6_title = ws[f'A{R_SEC6_HEAD}']
    s6_title.value = f'TABLE 5E — GROUND-TRUTH RESOURCE NORMS & ACTIVITY SPECIFICATION CATALOG ({T5E_ROWS} Verified Records)'
    s6_title.font = styles['font_white_bold']
    s6_title.fill = styles['fill_header']
    s6_title.alignment = styles['align_left']
    ws.row_dimensions[R_SEC6_HEAD].height = 24
    
    ws.merge_cells(f'A{R_SEC6_BANNER}:Q{R_SEC6_BANNER}')
    s6_banner = ws[f'A{R_SEC6_BANNER}']
    s6_banner.value = (
        'Ground-truth reference norms transcribed from official CPWD DAR 2019 Volume 1 (Sub-Head 02: Earth Work, PDF pages 88-140). '
        'Streamlined into 17 civil engineering columns with separated First-Principles, 8-Hour Shift Derivations, and dynamic method keys.'
    )
    s6_banner.font = styles['font_note']
    s6_banner.fill = styles['fill_note']
    s6_banner.alignment = styles['align_wrap']
    ws.row_dimensions[R_SEC6_BANNER].height = 24
    
    t5e_cols = [
        'Primary Work Scope', 'Ground Strata / Material', 'Execution Method & Depth Condition',
        'Operation / Task Specification', 'Line No', 'Resource Code',
        'Daily Productivity Norm (P)', 'Engineering First-Principles (Site Mechanics)',
        'CPWD 8-Hour Shift Logic (Mathematical Derivation)', 'Category, Task & Method Line Key',
        'Catalog Norm Input Qty', 'Category & Task Line Key (Fallback)',
        'Resource Description / Sub-Analysis Name', 'Resource Unit',
        'Basic / Sub-Analysis Rate (₹)', 'Cost Tag (W/A)', 'Internal Activity Scope Key'
    ]
    for ci, h in enumerate(t5e_cols, 1):
        c = ws.cell(row=R_SEC6_COLS, column=ci, value=h)
        c.font = styles['font_header']
        c.fill = styles['fill_header']
        c.alignment = styles['align_center']
        c.border = styles['border_header']
    ws.row_dimensions[R_SEC6_COLS].height = 24
    
    # Write Table 5E records (Cols A to Q)
    for idx, rec in enumerate(TABLE_5E_RECORDS):
        r = R_T5E_FIRST + idx
        
        ws.cell(row=r, column=1, value=rec['scope']).alignment = styles['align_left']
        ws.cell(row=r, column=1).font = styles['font_bold']
        
        ws.cell(row=r, column=2, value=rec['strata']).alignment = styles['align_left']
        ws.cell(row=r, column=3, value=rec['method']).alignment = styles['align_left']
        ws.cell(row=r, column=4, value=rec['task_spec']).alignment = styles['align_left']
        
        ws.cell(row=r, column=5, value=rec.get('line_no', 1)).alignment = styles['align_center']
        ws.cell(row=r, column=5).font = styles['font_bold']
        
        # Resource Code in Col F
        c_rc = ws.cell(row=r, column=6, value=rec['res_code'])
        c_rc.alignment = styles['align_center']
        c_rc.font = styles['font_bold']
        c_rc.fill = styles['fill_lookup']
        
        # Daily Productivity Norm (P) in Col G
        prod_val = rec.get('productivity_val')
        if prod_val is None and rec.get('res_unit') == 'day' and rec.get('qty_coeff', 0) > 0:
            prod_val = rec['batch_qty'] / rec['qty_coeff']
            
        if prod_val is not None:
            c_pn = ws.cell(row=r, column=7, value=prod_val)
            c_pn.alignment = styles['align_right']
            b_u = rec.get('batch_unit', 'unit')
            c_pn.number_format = f'0.00 "{b_u}/day"'
        else:
            c_pn = ws.cell(row=r, column=7, value="")
            c_pn.alignment = styles['align_center']
        c_pn.font = styles['font_note']
        
        # Physical Engineering First-Principles in Col H
        c_fp = ws.cell(row=r, column=8, value=rec.get('first_principles', ''))
        c_fp.alignment = styles['align_left']
        c_fp.font = styles['font_note']
        
        # CPWD 8-Hour Shift Logic in Col I
        c_sl = ws.cell(row=r, column=9, value=rec.get('shift_logic', ''))
        c_sl.alignment = styles['align_left']
        c_sl.font = styles['font_note']
        
        # Category, Task & Method Line Key in Col J (10)
        c_cat_k = ws.cell(row=r, column=10, value=rec.get('cat_task_method_line_key', ''))
        c_cat_k.alignment = styles['align_left']
        c_cat_k.font = styles['font_note']

        # Catalog Norm Input Qty in Col K (11)
        c_cq = ws.cell(row=r, column=11, value=rec['qty_coeff'])
        c_cq.alignment = styles['align_right']
        c_cq.number_format = '0.000'
        c_cq.font = styles['font_regular']

        # Category & Task Line Key (Fallback) in Col L (12)
        c_cat_fb = ws.cell(row=r, column=12, value=rec.get('cat_task_line_key', ''))
        c_cat_fb.alignment = styles['align_left']
        c_cat_fb.font = styles['font_note']

        # Resource Description in Col M (13)
        ws.cell(row=r, column=13, value=rec.get('res_name', '')).alignment = styles['align_left']
        ws.cell(row=r, column=13).font = styles['font_note']

        # Resource Unit in Col N (14)
        ws.cell(row=r, column=14, value=rec.get('res_unit', '')).alignment = styles['align_center']
        ws.cell(row=r, column=14).font = styles['font_note']

        # Basic / Sub-Analysis Rate in Col O (15)
        c_rate_val = ws.cell(row=r, column=15, value=rec.get('rate', 0.0))
        c_rate_val.alignment = styles['align_right']
        c_rate_val.font = styles['font_regular']
        c_rate_val.number_format = styles['fmt_currency']

        # Cost Tag in Col P (16)
        c_tag_val = ws.cell(row=r, column=16, value=rec.get('tag', 'W'))
        c_tag_val.alignment = styles['align_center']
        c_tag_val.font = styles['font_bold']

        # Internal Activity Key in Col Q (17)
        c_key = ws.cell(row=r, column=17, value=rec.get('internal_key', ''))
        c_key.alignment = styles['align_left']
        c_key.font = styles['font_note']
        
        for c in range(1, 18):
            ws.cell(row=r, column=c).border = styles['border_thin']
        ws.row_dimensions[r].height = 22


    # -------------------------------------------------------------
    # Master Activities Catalog (Cols Z to AK, Rows 105 to 153)
    # -------------------------------------------------------------
    act_headers = [
        'Master Category & Task Key', 'Master Activity Full Key', 'Master Fallback Key',
        'Batch Qty', 'Batch Unit', 'Published Say Rate (₹)', 'Official Nomenclature',
        'Loosening Lever', 'Handling Lever', 'Spreading Lever', 'Watering Lever', 'Compaction Lever'
    ]
    for ci, h in enumerate(act_headers, 26):  # Col Z is 26
        c = ws.cell(row=R_SEC6_COLS, column=ci, value=h)
        c.font = styles['font_header']
        c.fill = styles['fill_header']
        c.alignment = styles['align_center']
        
    for idx, act in enumerate(MASTER_ACTIVITIES):
        r = R_ACT_FIRST + idx
        ws.cell(row=r, column=26, value=act.get('cat_task_key', ''))  # Col Z
        ws.cell(row=r, column=27, value=act['lookup_key'])            # Col AA
        ws.cell(row=r, column=28, value=act['fallback_key'])          # Col AB
        ws.cell(row=r, column=29, value=act['batch_qty'])             # Col AC
        ws.cell(row=r, column=30, value=act['batch_unit'])            # Col AD
        ws.cell(row=r, column=31, value=act['published_say'])         # Col AE
        ws.cell(row=r, column=32, value=act['nomenclature'])          # Col AF
        ws.cell(row=r, column=33, value=act['loosening'])             # Col AG
        ws.cell(row=r, column=34, value=act['handling'])              # Col AH
        ws.cell(row=r, column=35, value=act['spreading'])             # Col AI
        ws.cell(row=r, column=36, value=act['watering'])              # Col AJ
        ws.cell(row=r, column=37, value=act['compaction'])            # Col AK

    # Column AL: Task / Operation Specification Options (Legacy)
    c_al_hdr = ws.cell(row=R_SEC6_COLS, column=38, value='Task / Operation Specification Options (Legacy)')
    c_al_hdr.font = styles['font_header']
    c_al_hdr.fill = styles['fill_header']
    c_al_hdr.alignment = styles['align_center']

    for idx, t in enumerate(TASK_OPTIONS):
        r = R_ACT_FIRST + idx
        c_t = ws.cell(row=r, column=38, value=t)
        c_t.alignment = styles['align_left']
        c_t.font = styles['font_regular']

    # Columns AN & AO: Category / Sub-Head Heading & Task Options Catalog (Source for C9 & D9 Dependent Range)
    c_an_hdr = ws.cell(row=R_SEC6_COLS, column=40, value='Category / Sub-Head Heading (AN)')
    c_an_hdr.font = styles['font_header']
    c_an_hdr.fill = styles['fill_header']
    c_an_hdr.alignment = styles['align_center']

    c_ao_hdr = ws.cell(row=R_SEC6_COLS, column=41, value='Task / Operation Specification (AO)')
    c_ao_hdr.font = styles['font_header']
    c_ao_hdr.fill = styles['fill_header']
    c_ao_hdr.alignment = styles['align_center']

    for idx, (cat_name, t_name, code) in enumerate(HEADING_TASK_CATALOG):
        r = R_CAT_FIRST + idx
        c_an = ws.cell(row=r, column=40, value=cat_name)
        c_an.alignment = styles['align_left']
        c_an.font = styles['font_regular']
        
        c_ao = ws.cell(row=r, column=41, value=t_name)
        c_ao.alignment = styles['align_left']
        c_ao.font = styles['font_regular']

    # Columns AP to AT: Legacy Dropdown Options Auxiliary Lists
    aux_configs = [
        (42, 'L1 Primary Work Nature (AP) — D6 Decision Tree', L1_OPTIONS),   # AP replaces old SCOPE_OPTIONS
        (43, 'Category / Sub-Head Headings (AQ)', CATEGORY_HEADINGS),
        (44, 'Execution Method Options (AR) — legacy', METHOD_OPTIONS),
        (45, 'Excavation Depth & Lift Options (AS) — legacy', LIFT_DEPTH_OPTIONS),
        (46, 'Ground Strata Options (AT) — legacy', STRATA_OPTIONS),
    ]
    for col_idx, hdr_title, opt_list in aux_configs:
        c_hdr = ws.cell(row=R_SEC6_COLS, column=col_idx, value=hdr_title)
        c_hdr.font = styles['font_header']
        c_hdr.fill = styles['fill_header']
        c_hdr.alignment = styles['align_center']
        for opt_idx, opt_val in enumerate(opt_list):
            r_opt = R_CAT_FIRST + opt_idx
            c_val = ws.cell(row=r_opt, column=col_idx, value=opt_val)
            c_val.alignment = styles['align_left']
            c_val.font = styles['font_regular']

    # ── Columns AU–AZ: Decision-Tree Cascade Lookup Tables ─────────────────
    # AU (47) = L2 key (L1 value); AV (48) = L2 label
    # AW (49) = L3 key (L1|L2);   AX (50) = L3 label
    # AY (51) = L4 key (L1|L2|L3); AZ (52) = L4 label (terminal task name)
    tree_col_headers = [
        (47, 'L2 Key — L1 Value (AU)'),
        (48, 'L2 Value — Sub-scope / Geometry (AV)'),
        (49, 'L3 Key — L1|L2 (AW)'),
        (50, 'L3 Value — Ground Material / Location (AX)'),
        (51, 'L4 Key — L1|L2|L3 (AY)'),
        (52, 'L4 Value — Terminal Item Spec / Task Name (AZ)'),
    ]
    for col_idx, hdr_title in tree_col_headers:
        c_hdr = ws.cell(row=R_SEC6_COLS, column=col_idx, value=hdr_title)
        c_hdr.font = styles['font_header']
        c_hdr.fill = styles['fill_header']
        c_hdr.alignment = styles['align_center']

    # Write L2_DATA  — (key, value) pairs → cols AU, AV
    for idx, (key_val, label_val) in enumerate(L2_DATA):
        r = R_CAT_FIRST + idx
        c_k = ws.cell(row=r, column=47, value=key_val)
        c_k.alignment = styles['align_left']
        c_k.font = styles['font_regular']
        c_v = ws.cell(row=r, column=48, value=label_val)
        c_v.alignment = styles['align_left']
        c_v.font = styles['font_regular']

    # Write L3_DATA  — (key, value) pairs → cols AW, AX
    for idx, (key_val, label_val) in enumerate(L3_DATA):
        r = R_CAT_FIRST + idx
        c_k = ws.cell(row=r, column=49, value=key_val)
        c_k.alignment = styles['align_left']
        c_k.font = styles['font_regular']
        c_v = ws.cell(row=r, column=50, value=label_val)
        c_v.alignment = styles['align_left']
        c_v.font = styles['font_regular']

    # Write L4_DATA  — (key, value) pairs → cols AY, AZ
    for idx, (key_val, label_val) in enumerate(L4_DATA):
        r = R_CAT_FIRST + idx
        c_k = ws.cell(row=r, column=51, value=key_val)
        c_k.alignment = styles['align_left']
        c_k.font = styles['font_regular']
        c_v = ws.cell(row=r, column=52, value=label_val)
        c_v.alignment = styles['align_left']
        c_v.font = styles['font_regular']

    ws.protection.sheet = True
    ws.freeze_panes = 'A4'
    print('Built Earth Work trade sheet.')
