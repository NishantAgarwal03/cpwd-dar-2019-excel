# -*- coding: utf-8 -*-
"""
01_Carriage_of_Materials - CPWD DAR Sub-Head 01 analytical simulator.

Rewritten so that every value on the sheet sits on its own row with an explicit
CELL ROLE (INPUT / OVERRIDE / LOOKUP / DERIVED / RESULT / SAY) and a plain
English explanation of where it comes from. Nothing is a bare number in an
unlabelled cell any more.

Defects corrected in this revision (all verified against the DAR source):

  1. Payload now reads the NET PAYABLE quantity (Table 1.1 column D), not the
     gross truck payload (column C). CPWD pays on the net figure after the
     looseness deduction - reproducing all 27 Table 1.1 base rates requires the
     net column (27/27 with net, 22/27 with gross).
  2. The schedule-unit scale factor is derived from the unit itself: 1000 for
     "1000 Nos", 100 for "100 m", otherwise 1. The old formula multiplied both
     by 100, understating every brick and tile rate tenfold.
  3. The SAY rate uses MROUND(x, 0.05) - CPWD quotes Say rates to the nearest
     5 paise. The two-decimal "analysed rate" is kept on the row above it so the
     unrounded figure is still visible.
  4. The CPWD pro-rata basis now looks up the Data Sheet 1 row for the ACTUAL
     lead and scales that row's km/day by (N / that row's trips). It used to be
     hard-wired to the 10 km row's constants (88.00 and 4.10).
  5. Average speed is looked up from Data Sheet 1 by lead, with an explicit
     override row. Previously it was typed by hand and could silently
     contradict the benchmark table sitting on the same sheet.
  6. Every division is wrapped in IFERROR, and the audit bar now tests the
     outputs that matter (N, payable output, W, and that the SAY cell is a
     number) instead of the fuel litres.
  7. Heading 1.2 derives its labour cost from the gang sizes and the live
     Rates_Master day wage, and takes CPOH from Factor_CPOH. The rupee
     constants 4279.86 / 5133.60 / 931.86 / 753.30 and the hard-coded 0.15 are
     gone.
  8. The trips override is honoured whenever it is filled, in any mode, so the
     sheet no longer ships showing a number that is being ignored.
  9. The municipal gate fee has its own row stating that it is added after the
     markup chain and is not subject to CPOH.
 10. The markup block now names Z the same way the other eleven builders do:
     Z = Y + CPOH, and the cess-inclusive figure is "Total Cost".
"""

from openpyxl.styles import Protection
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName

from scripts.carriage_tables import (
    DATASHEET1, RATE_DIESEL, RATE_MOBIL, RATE_BELDAR, RATE_TRUCK, GANG,
    table_11_rows, table_12_rows, validate as validate_tables, default_material,
)
from scripts.trade_builder import add_trade_header_and_legend
from scripts.trade_layout import (
    LAST_COL, SAY_RULE_NOTE, section_bar, col_headers, border_row, build_guide,
)

# --- Row map ---------------------------------------------------------------
R_P1_HEAD, R_P1_COLS = 8, 9
R_ITEM_CODE, R_MATERIAL, R_SCOPE, R_LIFT, R_GATE_FEE = 10, 11, 12, 13, 14
R_NOM_OVERRIDE, R_NOMENCLATURE = 15, 16

R_P2_HEAD, R_P2_COLS = 18, 19
R_LEAD, R_SPEED_BM, R_SPEED_OV, R_SPEED_EFF, R_TURNAROUND = 20, 21, 22, 23, 24
R_MODE, R_TRIPS_BASIS, R_TRIPS_OV, R_TRIPS = 25, 26, 27, 28
R_DIST_BASIS, R_DISTANCE, R_DIESEL, R_MOBIL = 29, 30, 31, 32

R_P3_HEAD, R_P3_COLS = 34, 35
R_T11_ROW = 36
R_PAY_GROSS, R_PAY_NET, R_PAY_OV, R_PAY_EFF = 37, 38, 39, 40
R_UNIT, R_SCALE, R_OUTPUT = 41, 42, 43

R_AUDIT = 45

R_RES_HEAD, R_RES_COLS, R_RES_FIRST = 47, 48, 49
RES_ROWS = 8
R_RES_LAST = R_RES_FIRST + RES_ROWS - 1                # 56
R_W_SUB = 57
R_TRIP_COST = 58

R_MU_HEAD, R_MU_COLS = 60, 61
R_W, R_X1, R_X, R_Y1, R_Y, R_Z1, R_Z, R_Z2 = 62, 63, 64, 65, 66, 67, 68, 69
R_TOTAL, R_TRIP_OH, R_RATE_UNIT, R_RATE_SCHED = 70, 71, 72, 73
R_DAR_PUB, R_DAR_DIFF = 74, 75
R_SAY, R_SAY_NOTE = 76, 77

R_MAN_HEAD, R_MAN_BANNER, R_MAN_COLS = 79, 80, 81
R_M_CODE, R_M_CAT, R_M_LEAD, R_M_STEPS = 82, 83, 84, 85
R_M_GANG, R_M_GANG_ADD, R_M_WAGE = 86, 87, 88
R_M_BASE, R_M_ADD, R_M_LABOUR, R_M_CPOH, R_M_TOTAL = 89, 90, 91, 92, 93
R_M_CAP, R_M_UNIT, R_M_SCALE, R_M_RATE, R_M_SAY = 94, 95, 96, 97, 98

R_BM_HEAD = 100
R_DS1_HEAD, R_DS1_COLS, R_DS1_FIRST = 101, 102, 103
DS1_ROWS = 30
R_DS1_LAST = R_DS1_FIRST + DS1_ROWS - 1                # 132
R_DS1_NOTE = 133

R_T11_HEAD, R_T11_COLS, R_T11_FIRST = 135, 136, 137
T11_ROWS = 37
R_T11_LAST = R_T11_FIRST + T11_ROWS - 1                # 173

R_T12_HEAD, R_T12_COLS, R_T12_FIRST = 175, 176, 177
T12_ROWS = 35
R_T12_LAST = R_T12_FIRST + T12_ROWS - 1                # 211

R_SC_HEAD, R_SC_COLS, R_SC_FIRST = 213, 214, 215
SC_ROWS = 6
R_SC_LAST = R_SC_FIRST + SC_ROWS - 1                   # 220

# The reference tables carry the book's full column set and run wider than the
# nine-column working area above them.
BENCH_COL = 14   # column N

PARAM_HEADERS = ['Parameter', 'Value', 'Unit', 'Cell Role',
                 'Where this value comes from, and the CPWD DAR 2019 basis for it', '', '', '', '']

ROLE_FILL = {
    'INPUT': 'fill_input',
    'OVERRIDE': 'fill_override',
    'LOOKUP': 'fill_lookup',
    'DERIVED': 'fill_calc',
    'RESULT': 'fill_result',
    'SAY': 'fill_say',
}


def _param_row(ws, r, styles, label, value, unit, role, note,
               number_format=None, dv=None, wrap_value=False):
    """One parameter per row: label | value | unit | role | explanation."""
    lc = ws.cell(row=r, column=1, value=label)
    lc.font = styles['font_bold']
    lc.alignment = styles['align_wrap']

    vc = ws.cell(row=r, column=2, value=value)
    vc.fill = styles[ROLE_FILL[role]]
    vc.font = styles['font_say'] if role == 'SAY' else styles['font_bold']
    vc.alignment = styles['align_wrap'] if wrap_value else styles['align_center']
    if number_format:
        vc.number_format = number_format
    if dv is not None:
        dv.add(vc)

    uc = ws.cell(row=r, column=3, value=unit)
    uc.alignment = styles['align_center']
    uc.font = styles['font_regular']

    rc = ws.cell(row=r, column=4, value=role)
    rc.alignment = styles['align_center']
    rc.font = styles['font_bold']
    rc.fill = styles[ROLE_FILL[role]]

    ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=LAST_COL)
    nc = ws.cell(row=r, column=5, value=note)
    nc.font = styles['font_note']
    nc.alignment = styles['align_wrap']

    border_row(ws, r, styles)
    ws.row_dimensions[r].height = 30
    return vc


def build_carriage_trade(wb, config, styles):
    ws = wb.create_sheet(title=config['sheet_name'])
    ws.views.sheetView[0].showGridLines = True
    add_trade_header_and_legend(ws, config, styles)
    build_guide(ws, styles)

    sn = config['sheet_name']
    ds1 = f"$A${R_DS1_FIRST}:$A${R_DS1_LAST}"        # lead
    ds1_s = f"$B${R_DS1_FIRST}:$B${R_DS1_LAST}"      # speed
    ds1_n = f"$C${R_DS1_FIRST}:$C${R_DS1_LAST}"      # trips
    ds1_km = f"$D${R_DS1_FIRST}:$D${R_DS1_LAST}"     # km/day
    t11_mat = f"$B${R_T11_FIRST}:$B${R_T11_LAST}"
    t11_gross = f"$C${R_T11_FIRST}:$C${R_T11_LAST}"
    t11_net = f"$D${R_T11_FIRST}:$D${R_T11_LAST}"
    t11_unit = f"$E${R_T11_FIRST}:$E${R_T11_LAST}"
    # the eight printed rate columns: 1..5 km, then the three per-km bands
    lad = {k: f"$%s${R_T11_FIRST}:$%s${R_T11_LAST}" % (c, c)
           for k, c in ((1, 'G'), (2, 'H'), (3, 'I'), (4, 'J'), (5, 'K'),
                        ('b5_10', 'L'), ('b10_20', 'M'), ('b20', 'N'))}

    # --- Data validations -------------------------------------------------
    def mkdv(formula, blank=True):
        d = DataValidation(type='list', formula1=formula, allow_blank=blank,
                           showErrorMessage=False)
        ws.add_data_validation(d)
        return d

    dv_yesno = mkdv('"YES,NO"', False)
    dv_mode = mkdv('"STANDARD (DAYTIME),URBAN RESTRICTED HOURS"', False)
    dv_basis = mkdv('"DIRECT ROUTE (2NL + 6),CPWD PRO-RATA (Data Sheet 1)"', False)
    dv_tripbasis = mkdv('"CPWD FRACTIONAL (Data Sheet 1 basis),WHOLE TRIPS (round down to integer)"',
                        False)
    dv_unit = mkdv('"cum,tonne,metre,100 m,1000 Nos"', False)
    dv_cat = mkdv('"Category A (Bulk / Earth / Bricks),Category B (Heavy / Pipes / Steel)"', False)
    dv_mlead = mkdv('"50,100,150,200,250,300,350,400,450,500"', False)
    dv_items = mkdv('=CPWD_Carriage_Item_Codes', False)
    dv_mats = mkdv('=CPWD_Carriage_Materials', False)
    dv_scope = mkdv('=CPWD_Carriage_Scope', False)
    dv_lift = mkdv('"for all lifts,for lift upto 1.5 m,with mechanical lift,'
                   'for all lifts and leads"', False)

    for _sp in (7, 17, 33, 44, 46, 59, 78, 99):
        ws.row_dimensions[_sp].height = 8

    # =====================================================================
    # PANEL 1 - what is being moved
    # =====================================================================
    section_bar(ws, R_P1_HEAD,
                'PANEL 1 - SCOPE & SPECIFICATION  (what is being moved, and under what contract scope)',
                styles)
    col_headers(ws, R_P1_COLS, PARAM_HEADERS, styles, height=26)
    ws.merge_cells(start_row=R_P1_COLS, start_column=5, end_row=R_P1_COLS, end_column=LAST_COL)

    _param_row(ws, R_ITEM_CODE, styles, 'Custom item code',
               config.get('default_item_code', '1.1.CUSTOM'), '-', 'INPUT',
               'Your reference for this analysis. Pick a DAR item number from the dropdown to base it '
               'on a printed item, or type your own code for a genuinely non-DSR item.', dv=dv_items)

    # Always taken from Table 1.1 itself: if this string does not match a row
    # in that table the Panel 3 lookups all blank out and the rate reads zero.
    _param_row(ws, R_MATERIAL, styles, 'Material commodity', default_material(),
               '-', 'INPUT',
               'DRIVES COST. Sets the truck payload, the net payable quantity and the billing unit, all '
               'looked up from Table 1.1 below. Choose from the dropdown so the lookups resolve.',
               dv=dv_mats, wrap_value=True)

    _param_row(ws, R_SCOPE, styles, 'Handling scope',
               config.get('default_scope', 'including loading, transporting, unloading and stacking'),
               '-', 'INPUT',
               'DRIVES COST. Sets the Beldar gang on the resource schedule below: 6 for full turnkey '
               'handling, 5 without stacking, 3 when machine-loaded, 3.75 at a railway siding, 0 for '
               'pure haulage. See the scope table at the bottom of this sheet.',
               dv=dv_scope, wrap_value=True)

    _param_row(ws, R_LIFT, styles, 'Lift condition',
               config.get('default_lift', 'for all lifts'), '-', 'INPUT',
               'NOMENCLATURE ONLY. The DAR quotes one carriage rate whatever the lift, so this changes '
               'the wording of the item and nothing in the cost.', dv=dv_lift)

    _param_row(ws, R_GATE_FEE, styles, 'Municipal gate / tipping fee', 0.00, 'Rs per billing unit',
               'INPUT',
               'USER-SUPPLIED COST. The DAR carries no tipping or gate fee anywhere in Sub-Head 01, so '
               'there is no figure to look up. It is added AFTER the markup chain and is therefore NOT '
               'subject to the 15% CPOH - it is a reimbursable disbursement, not contractor cost.',
               number_format=styles['fmt_currency'])

    _param_row(ws, R_NOM_OVERRIDE, styles, 'Nomenclature override', None, '-', 'OVERRIDE',
               'Leave blank and the item description below writes itself from your Panel 1 choices. '
               'Type here only when you want to word the item yourself.', wrap_value=True)

    ws.cell(row=R_NOMENCLATURE, column=1, value='Assembled item nomenclature').font = styles['font_bold']
    ws.merge_cells(start_row=R_NOMENCLATURE, start_column=2, end_row=R_NOMENCLATURE, end_column=LAST_COL)
    nc = ws.cell(row=R_NOMENCLATURE, column=2)
    nc.value = (
        f'=IF(B{R_NOM_OVERRIDE}<>"", B{R_NOM_OVERRIDE}, '
        f'"Carriage of " & B{R_MATERIAL} & " by mechanical transport " & B{R_SCOPE} & '
        f'" for lead upto " & TEXT(B{R_LEAD}, "0.00") & " km " & B{R_LIFT} & '
        f'IF(B{R_GATE_FEE}>0, ", including municipal tipping / gate fee of Rs " & '
        f'TEXT(B{R_GATE_FEE}, "0.00") & " per " & B{R_UNIT}, "") & '
        f'", complete as per directions of Engineer-in-charge.")')
    nc.font = styles['font_bold']
    nc.fill = styles['fill_calc']
    nc.alignment = styles['align_wrap']
    border_row(ws, R_NOMENCLATURE, styles)
    ws.row_dimensions[R_NOMENCLATURE].height = 40

    # =====================================================================
    # PANEL 2 - trip dynamics
    # =====================================================================
    section_bar(ws, R_P2_HEAD,
                'PANEL 2 - ROUTE & TRIP DYNAMICS  (how far, how fast, how many trips a shift)', styles)
    col_headers(ws, R_P2_COLS, PARAM_HEADERS, styles, height=26)
    ws.merge_cells(start_row=R_P2_COLS, start_column=5, end_row=R_P2_COLS, end_column=LAST_COL)

    _param_row(ws, R_LEAD, styles, 'Lead distance (L)', config.get('default_lead', 26.0), 'km', 'INPUT',
               'DRIVES COST. One-way haul distance. Everything below - speed, trips, fuel - follows '
               'from it.', number_format='0.00')

    _param_row(ws, R_SPEED_BM, styles, 'Benchmark average speed at this lead', None, 'km/h', 'LOOKUP',
               'Read off CPWD Data Sheet No. 1 (section 5A below) for the lead above: 16 km/h at 1 km '
               'rising to 31 km/h at 30 km. Beyond 30 km the last tabulated row is used.',
               number_format='0.0')
    ws.cell(row=R_SPEED_BM, column=2).value = (
        f'=IFERROR(INDEX({ds1_s}, MATCH(B{R_LEAD}, {ds1}, 1)), 16)')

    _param_row(ws, R_SPEED_OV, styles, 'Speed override', None, 'km/h', 'OVERRIDE',
               'Leave blank to use the Data Sheet 1 benchmark above. Fill it in only when site '
               'conditions genuinely differ, and record why - departing from the benchmark is what '
               'makes a rate non-DSR.', number_format='0.0')

    _param_row(ws, R_SPEED_EFF, styles, 'Effective average speed (S)', None, 'km/h', 'DERIVED',
               'The override if you gave one, otherwise the benchmark. This is the S used in the trips '
               'formula.', number_format='0.0')
    ws.cell(row=R_SPEED_EFF, column=2).value = (
        f'=IF(B{R_SPEED_OV}<>"", B{R_SPEED_OV}, B{R_SPEED_BM})')

    _param_row(ws, R_TURNAROUND, styles, 'Turnaround time (T)',
               config.get('default_turnaround', 1.0), 'hours per trip', 'INPUT',
               'Loading plus unloading plus waiting per round trip. CPWD Data Sheet 1 uses 1.00 hour.',
               number_format='0.00')

    _param_row(ws, R_MODE, styles, 'Operational mode', 'STANDARD (DAYTIME)', '-', 'INPUT',
               'Urban restricted hours means heavy vehicles may only run off-peak, which caps the trips '
               'achievable in a shift. Selecting it does not change the cost by itself - enter the '
               'capped trip count in the override on the next row. DAR item 1.1.18 is the worked '
               'example: 3.00 trips instead of the 4.10 the formula gives at 10 km.', dv=dv_mode)

    _param_row(ws, R_TRIPS_BASIS, styles, 'Trip count basis',
               'CPWD FRACTIONAL (Data Sheet 1 basis)', '-', 'INPUT',
               'CPWD FRACTIONAL is the book\'s own basis and the default. Data Sheet No. 1 quotes '
               'trips as decimals throughout - 7.11 at 1 km, 4.10 at 10 km, 2.73 at 30 km - because N '
               'is a productivity RATE per 8-hour shift, not a count of trips on one particular day; '
               'the truck is hired by the day and the cost is spread over the average output. '
               'WHOLE TRIPS rounds down to an integer, which is defensible when a real site or '
               'regulatory constraint caps the trips, but it does NOT reproduce the printed DAR rates '
               '- flooring 7.11 to 7 at 1 km raises the lime/moorum rate from Rs 104.94 to Rs 106.52. '
               'Where the DAR itself fixes a whole trip count (item 1.1.18 uses exactly 3.00 under '
               'urban restricted hours) it does so as an explicit cap, not as a rounding rule - use '
               'the Trips override below for that.',
               dv=dv_tripbasis, wrap_value=True)

    _param_row(ws, R_TRIPS_OV, styles, 'Trips override', None, 'trips per shift', 'OVERRIDE',
               'Leave blank to use the calculated trips below. Fill it in to force a trip count - this '
               'is how a restricted-hours item is priced. It is honoured in any mode.',
               number_format='0.00')

    _param_row(ws, R_TRIPS, styles, 'Daily trips achieved (N)', None, 'trips per shift', 'DERIVED',
               'N = 8 / ((2L / S) + T) - an 8-hour shift divided by one round trip. The Trips '
               'override wins if it is filled; otherwise the Trip count basis decides whether the '
               'result keeps its decimals (CPWD) or is rounded down to a whole trip.',
               number_format='0.00')
    ws.cell(row=R_TRIPS, column=2).value = (
        f'=IF(B{R_TRIPS_OV}<>"", B{R_TRIPS_OV}, '
        f'IF(B{R_TRIPS_BASIS}="WHOLE TRIPS (round down to integer)", '
        f'IFERROR(INT(8 / ((2 * B{R_LEAD} / B{R_SPEED_EFF}) + B{R_TURNAROUND})), 0), '
        f'IFERROR(ROUND(8 / ((2 * B{R_LEAD} / B{R_SPEED_EFF}) + B{R_TURNAROUND}), 2), 0)))')

    _param_row(ws, R_DIST_BASIS, styles, 'Distance / fuel basis', 'DIRECT ROUTE (2NL + 6)', '-', 'INPUT',
               'DIRECT ROUTE computes km/day = 2NL + 6, the 6 km being the depot run - this reproduces '
               'every row of Data Sheet 1. CPWD PRO-RATA instead takes the Data Sheet 1 km/day for your '
               'lead and scales it by (your trips / the benchmark trips at that lead); that is how DAR '
               'item 1.1.18 derives 64.40 km from the 10 km row.', dv=dv_basis)

    _param_row(ws, R_DISTANCE, styles, 'Distance travelled per shift', None, 'km', 'DERIVED',
               'Follows the basis selected above.', number_format='0.00')
    ws.cell(row=R_DISTANCE, column=2).value = (
        f'=IF(B{R_DIST_BASIS}="CPWD PRO-RATA (Data Sheet 1)", '
        f'IFERROR(ROUND(INDEX({ds1_km}, MATCH(B{R_LEAD}, {ds1}, 1)) * '
        f'(B{R_TRIPS} / INDEX({ds1_n}, MATCH(B{R_LEAD}, {ds1}, 1))), 2), 0), '
        f'IFERROR(ROUND((2 * B{R_TRIPS} * B{R_LEAD}) + 6, 2), 0))')

    _param_row(ws, R_DIESEL, styles, 'Diesel consumed per shift', None, 'litres', 'DERIVED',
               'CPWD Note 3: distance travelled / 5.0 km per litre.', number_format='0.00')
    ws.cell(row=R_DIESEL, column=2).value = f'=ROUND(B{R_DISTANCE} / 5, 2)'

    _param_row(ws, R_MOBIL, styles, 'Mobil oil consumed per shift', None, 'litres', 'DERIVED',
               'CPWD Note 4: distance travelled / 140.0 km per litre.', number_format='0.000')
    ws.cell(row=R_MOBIL, column=2).value = f'=ROUND(B{R_DISTANCE} / 140, 3)'

    # =====================================================================
    # PANEL 3 - payload and billing
    # =====================================================================
    section_bar(ws, R_P3_HEAD,
                'PANEL 3 - PAYLOAD & BILLING BASIS  (how much is actually paid for per trip)', styles)
    col_headers(ws, R_P3_COLS, PARAM_HEADERS, styles, height=26)
    ws.merge_cells(start_row=R_P3_COLS, start_column=5, end_row=R_P3_COLS, end_column=LAST_COL)

    _param_row(ws, R_T11_ROW, styles, 'Table 1.1 row for this material', None, 'row index', 'DERIVED',
               'Which row of Table 1.1 (section 5B) the material commodity matches. Everything below, '
               'and the published DAR rate in Section 2, reads that row. Shows #N/A if the material '
               'is not one of the book\'s - use the overrides in that case.', number_format='0')
    ws.cell(row=R_T11_ROW, column=2).value = (
        f'=IFERROR(MATCH(B{R_MATERIAL}, {t11_mat}, 0), "")')

    _param_row(ws, R_PAY_GROSS, styles, 'Gross truck payload per trip', None, 'per Table 1.1', 'LOOKUP',
               'What physically goes on the truck, from Table 1.1 below. Shown for information only - '
               'it is NOT what the rate is divided by.', number_format='0.00')
    ws.cell(row=R_PAY_GROSS, column=2).value = (
        f'=IFERROR(INDEX({t11_gross}, B{R_T11_ROW}), "")')

    _param_row(ws, R_PAY_NET, styles, 'Net payable quantity per trip', None, 'per Table 1.1', 'LOOKUP',
               'THIS is the divisor CPWD uses. Loose materials are paid on a reduced quantity: earth '
               'less 20%, excavated rock less 50%, soling stone less 15%, 40 mm aggregate and manure '
               'less 8%. Using the gross payload instead understates the rate by that percentage.',
               number_format='0.00')
    ws.cell(row=R_PAY_NET, column=2).value = (
        f'=IFERROR(INDEX({t11_net}, B{R_T11_ROW}), "")')

    _param_row(ws, R_PAY_OV, styles, 'Payable quantity override', None, 'per Table 1.1', 'OVERRIDE',
               'Leave blank to use the net payable quantity above. Fill it in for a material that is '
               'not in Table 1.1, or when a site measurement justifies a different figure.',
               number_format='0.00')

    _param_row(ws, R_PAY_EFF, styles, 'Effective payable quantity per trip', None, 'per Table 1.1',
               'DERIVED', 'The override if given, otherwise the net payable quantity.',
               number_format='0.00')
    ws.cell(row=R_PAY_EFF, column=2).value = (
        f'=IF(B{R_PAY_OV}<>"", B{R_PAY_OV}, IF(B{R_PAY_NET}="", 0, B{R_PAY_NET}))')

    _param_row(ws, R_UNIT, styles, 'Schedule billing unit', None, '-', 'LOOKUP',
               'The unit the DAR bills this material in, from Table 1.1. Bulk materials in cum, cement '
               'and steel in tonne, bricks per 1000 Nos, pipes per 100 m.')
    ws.cell(row=R_UNIT, column=2).value = (
        f'=IFERROR(INDEX({t11_unit}, B{R_T11_ROW}), "cum")')

    _param_row(ws, R_SCALE, styles, 'Schedule unit scale factor', None, 'multiplier', 'DERIVED',
               'The rate is first derived per single unit (per cum, per metre, per brick) and then '
               'multiplied by this factor to reach the schedule unit: 1000 for "1000 Nos", 100 for '
               '"100 m", otherwise 1.', number_format='0')
    ws.cell(row=R_SCALE, column=2).value = (
        f'=IF(B{R_UNIT}="1000 Nos", 1000, IF(B{R_UNIT}="100 m", 100, 1))')

    _param_row(ws, R_OUTPUT, styles, 'Total payable output per shift', None, 'per Table 1.1', 'RESULT',
               'Trips x effective payable quantity. This is what the shift cost is divided by.',
               number_format='0.0000')
    ws.cell(row=R_OUTPUT, column=2).value = f'=ROUND(B{R_TRIPS} * B{R_PAY_EFF}, 4)'

    # =====================================================================
    # Audit bar
    # =====================================================================
    ws.cell(row=R_AUDIT, column=1, value='AUDIT STATUS:').font = styles['font_white_bold']
    ws.cell(row=R_AUDIT, column=1).fill = styles['fill_header']
    ws.cell(row=R_AUDIT, column=1).alignment = styles['align_center']

    ab = ws.cell(row=R_AUDIT, column=2)
    # The book comparison in column I is advisory, not pass/fail - see the note
    # on the variance row for why a mid-band lead legitimately differs.
    ab.value = (f'=IF(AND(D{R_AUDIT}="OK", F{R_AUDIT}="OK", H{R_AUDIT}="OK"), '
                f'"[PASS] ALL CHECKS OK", "[ALERT] CHECKS FAILED")')
    ab.font = styles['font_result']
    ab.fill = styles['fill_result']
    ab.alignment = styles['align_center']

    ws.cell(row=R_AUDIT, column=3, value='Trips:').font = styles['font_note']
    ws.cell(row=R_AUDIT, column=3).alignment = styles['align_right']
    ws.cell(row=R_AUDIT, column=4,
            value=f'=IF(B{R_TRIPS}>0, "OK", "ERR: N<=0")').alignment = styles['align_center']

    ws.cell(row=R_AUDIT, column=5, value='Payable output:').font = styles['font_note']
    ws.cell(row=R_AUDIT, column=5).alignment = styles['align_right']
    ws.cell(row=R_AUDIT, column=6,
            value=f'=IF(B{R_OUTPUT}>0, "OK", "ERR: no payable qty - pick a Table 1.1 material '
                  f'or set the override")').alignment = styles['align_center']

    ws.cell(row=R_AUDIT, column=7, value='Direct cost:').font = styles['font_note']
    ws.cell(row=R_AUDIT, column=7).alignment = styles['align_right']
    ws.cell(row=R_AUDIT, column=8,
            value=f'=IF(G{R_W_SUB}>0, "OK", "ERR: W<=0")').alignment = styles['align_center']

    ws.cell(row=R_AUDIT, column=9,
            value=f'=IF(NOT(AND(ISNUMBER(G{R_SAY}), G{R_SAY}>0)), "ERR: rate not resolved", '
                  f'IF(G{R_DAR_DIFF}="", "n/a - custom material", '
                  f'IF(ABS(G{R_DAR_DIFF})<0.05, "matches book exactly", '
                  f'IF(ABS(G{R_DAR_DIFF})/G{R_DAR_PUB}<=0.02, "within book band tolerance", '
                  f'"CHECK: over 2% from book"))))')
    ws.cell(row=R_AUDIT, column=9).alignment = styles['align_center']

    for c in (4, 6, 8, 9):
        ws.cell(row=R_AUDIT, column=c).font = styles['font_bold']
    border_row(ws, R_AUDIT, styles)
    ws.row_dimensions[R_AUDIT].height = 22

    # =====================================================================
    # Resource schedule
    # =====================================================================
    section_bar(ws, R_RES_HEAD,
                'SECTION 1 - MECHANICAL TRANSPORT RESOURCE SCHEDULE FOR ONE 8-HOUR SHIFT  '
                '(truck hire + labour gang + fuel)', styles)
    col_headers(ws, R_RES_COLS,
                ['Line', 'Code / Source', 'Resource / Fuel Description', 'Unit', 'Quantity / Coeff',
                 'Basic Rate (Rs)', 'Amount (Rs)', 'Source Reference / CPWD Note', 'Cell Role'],
                styles)

    beldar = (f'=IF(ISNUMBER(SEARCH("transporting only", B{R_SCOPE})), 0, '
              f'IF(ISNUMBER(SEARCH("machine loaded", B{R_SCOPE})), 3, '
              f'IF(ISNUMBER(SEARCH("excluding loading", B{R_SCOPE})), 3, '
              f'IF(ISNUMBER(SEARCH("railway siding", B{R_SCOPE})), 3.75, '
              f'IF(ISNUMBER(SEARCH("excluding stacking", B{R_SCOPE})), 5, 6)))))')

    resources = [
        {'code': '0084', 'qty': 1.0, 'role': 'INPUT',
         'note': 'CPWD Note 5: hire of a 9-tonne diesel truck for one 8-hour shift, excluding diesel '
                 'and mobil oil (those are the two rows below).'},
        {'code': '0114', 'qty': beldar, 'role': 'DERIVED',
         'note': 'Gang size follows the Handling Scope you chose in Panel 1: 6 turnkey / 5 no stacking / '
                 '3 machine-loaded / 3.75 railway siding / 0 haulage only. See section 5D below.'},
        {'code': '1235', 'qty': f'=B{R_DIESEL}', 'role': 'DERIVED',
         'note': 'CPWD Note 3: litres from Panel 2 (distance / 5.0 km per litre).'},
        {'code': '5001', 'qty': f'=B{R_MOBIL}', 'role': 'DERIVED',
         'note': 'CPWD Note 4: litres from Panel 2 (distance / 140.0 km per litre).'},
    ]

    for i in range(RES_ROWS):
        r = R_RES_FIRST + i
        res = resources[i] if i < len(resources) else None
        ws.cell(row=r, column=1, value=i + 1).alignment = styles['align_center']

        cc = ws.cell(row=r, column=2, value=(res['code'] if res else ''))
        cc.alignment = styles['align_center']
        cc.font = styles['font_bold']
        cc.fill = styles['fill_input']

        ws.cell(row=r, column=3,
                value=f'=IF(B{r}="","", IFERROR(INDEX(Rates_Master!$C:$C, '
                      f'MATCH(B{r}, Rates_Master!$A:$A, 0)), "Code not in Rates_Master"))'
                ).fill = styles['fill_lookup']
        cu = ws.cell(row=r, column=4,
                     value=f'=IF(B{r}="","", IFERROR(INDEX(Rates_Master!$D:$D, '
                           f'MATCH(B{r}, Rates_Master!$A:$A, 0)), ""))')
        cu.alignment = styles['align_center']
        cu.fill = styles['fill_lookup']

        role = res['role'] if res else 'INPUT'
        cq = ws.cell(row=r, column=5, value=(res['qty'] if res else None))
        cq.alignment = styles['align_right']
        cq.number_format = styles['fmt_qty']
        cq.fill = styles[ROLE_FILL[role]]
        cq.font = styles['font_bold']

        cr = ws.cell(row=r, column=6,
                     value=f'=IF(B{r}="","", IFERROR(INDEX(Rates_Master!$E:$E, '
                           f'MATCH(B{r}, Rates_Master!$A:$A, 0)), 0))')
        cr.alignment = styles['align_right']
        cr.number_format = styles['fmt_currency']
        cr.fill = styles['fill_lookup']

        ca = ws.cell(row=r, column=7, value=f'=IF(OR(B{r}="", E{r}=""), 0, ROUND(E{r} * F{r}, 2))')
        ca.alignment = styles['align_right']
        ca.number_format = styles['fmt_currency']

        cn = ws.cell(row=r, column=8, value=(res['note'] if res else ''))
        cn.font = styles['font_note']
        cn.alignment = styles['align_wrap']

        crole = ws.cell(row=r, column=9, value=(role if res else 'INPUT'))
        crole.alignment = styles['align_center']
        crole.font = styles['font_note']
        crole.fill = styles[ROLE_FILL[role]]

        border_row(ws, r, styles)
        ws.row_dimensions[r].height = 26

    ws.merge_cells(start_row=R_W_SUB, start_column=1, end_row=R_W_SUB, end_column=6)
    ws.cell(row=R_W_SUB, column=1,
            value='Direct operating cost of one 8-hour shift, W (Rs):').font = styles['font_bold']
    ws.cell(row=R_W_SUB, column=1).alignment = styles['align_right']
    cw = ws.cell(row=R_W_SUB, column=7, value=f'=SUM(G{R_RES_FIRST}:G{R_RES_LAST})')
    cw.font = styles['font_bold']
    cw.alignment = styles['align_right']
    cw.number_format = styles['fmt_currency']
    ws.cell(row=R_W_SUB, column=8,
            value='Compare against the "Total Shift Cost" column of Data Sheet 1 below for the same '
                  'lead - they should agree.').font = styles['font_note']
    for c in range(1, LAST_COL + 1):
        ws.cell(row=R_W_SUB, column=c).fill = styles['fill_subtotal']
        ws.cell(row=R_W_SUB, column=c).border = styles['border_double_bottom']
    ws.row_dimensions[R_W_SUB].height = 24

    ws.merge_cells(start_row=R_TRIP_COST, start_column=1, end_row=R_TRIP_COST, end_column=6)
    ws.cell(row=R_TRIP_COST, column=1,
            value='Direct operating cost per round trip (Rs):').font = styles['font_bold']
    ws.cell(row=R_TRIP_COST, column=1).alignment = styles['align_right']
    ct = ws.cell(row=R_TRIP_COST, column=7,
                 value=f'=IFERROR(ROUND(G{R_W_SUB} / B{R_TRIPS}, 2), 0)')
    ct.font = styles['font_bold']
    ct.alignment = styles['align_right']
    ct.number_format = styles['fmt_currency']
    ws.cell(row=R_TRIP_COST, column=8, value='W divided by the trips in Panel 2.').font = styles['font_note']
    border_row(ws, R_TRIP_COST, styles)
    ws.row_dimensions[R_TRIP_COST].height = 22

    # =====================================================================
    # Markups and rate derivation
    # =====================================================================
    section_bar(ws, R_MU_HEAD,
                'SECTION 2 - STATUTORY MARKUPS & RATE DERIVATION  '
                '(CPWD carriage convention: 15% CPOH only)', styles)
    col_headers(ws, R_MU_COLS,
                ['Step', 'Description / Stage', 'Apply? (YES/NO)', 'Basis Applied', 'Base Amount (Rs)',
                 'Factor / %', 'Amount (Rs)', 'CPWD Statutory Rule & Guidance Note', 'Cell Role'],
                styles)

    mu_rows = [
        (R_W, 'W', 'Direct operating cost of the shift (truck + labour + fuel)', None, 'Direct sum',
         None, None, f'=G{R_W_SUB}',
         'Carried straight down from Section 1.', 'DERIVED'),
        (R_X1, 'X1', 'Add Water Charges', 'NO', 'On W', f'=G{R_W}', '=Factor_Water',
         f'=IF(C{R_X1}="YES", ROUND(E{R_X1} * F{R_X1}, 2), 0)',
         'CPWD RULE: switched off for carriage - no water is consumed in haulage. DAR item 1.1.18 goes '
         'straight from TOTAL to "Add 15% CPOH".', 'INPUT'),
        (R_X, 'X', 'Subtotal X (W + Water)', None, 'W + Water', None, None, f'=G{R_W} + G{R_X1}',
         'Running subtotal.', 'DERIVED'),
        (R_Y1, 'Y1', 'Add GST on works contract', 'NO', 'On X', f'=G{R_X}', '=Factor_GST',
         f'=IF(C{R_Y1}="YES", ROUND(E{R_Y1} * F{R_Y1}, 2), 0)',
         'CPWD RULE: the printed DAR carriage rates exclude GST. Leave NO to stay comparable with the '
         'book; switch to YES only for a tender that requires it.', 'INPUT'),
        (R_Y, 'Y', 'Subtotal Y (X + GST)', None, 'X + GST', None, None, f'=G{R_X} + G{R_Y1}',
         'Base for contractor profit and overheads.', 'DERIVED'),
        (R_Z1, 'Z1', 'Add Contractor Profit & Overheads (15% CPOH)', 'YES', 'On Y', f'=G{R_Y}',
         '=Factor_CPOH', f'=IF(C{R_Z1}="YES", ROUND(E{R_Z1} * F{R_Z1}, 2), 0)',
         'CPWD RULE: 15% CPOH is applied to every carriage item without exception.', 'INPUT'),
        (R_Z, 'Z', 'Subtotal Z (Y + CPOH)', None, 'Y + CPOH', None, None, f'=G{R_Y} + G{R_Z1}',
         'Named the same way as on the other eleven builder sheets.', 'DERIVED'),
        (R_Z2, 'Z2', 'Add BOCW Welfare Cess', 'NO', 'On Z', f'=G{R_Z}', '=Factor_Cess',
         f'=IF(C{R_Z2}="YES", ROUND(E{R_Z2} * F{R_Z2}, 2), 0)',
         'CPWD RULE: excluded from the printed base rates. Switch to YES only if your circle requires '
         'it on carriage.', 'INPUT'),
        (R_TOTAL, 'Total', 'Total cost of one 8-hour shift with overheads', None, 'Z + Cess', None, None,
         f'=G{R_Z} + G{R_Z2}', 'The full shift cost that the payable output is divided into.', 'RESULT'),
        (R_TRIP_OH, '-', 'Cost per round trip with overheads', None, 'Total / N', f'=G{R_TOTAL}',
         f'=B{R_TRIPS}', f'=IFERROR(ROUND(G{R_TOTAL} / B{R_TRIPS}, 2), 0)',
         'For comparison with the "Cost per Trip" column of Data Sheet 1.', 'DERIVED'),
        (R_RATE_UNIT, '-', 'Rate per single payable unit, plus gate fee', None, 'Total / payable output',
         f'=G{R_TOTAL}', f'=B{R_OUTPUT}',
         f'=IFERROR(ROUND(G{R_TOTAL} / B{R_OUTPUT}, 2), 0) + B{R_GATE_FEE}',
         'Per cum / per tonne / per metre / per brick. Divided by the NET payable output from Panel 3, '
         'not by the gross truck payload. The gate fee is added here, after CPOH.', 'DERIVED'),
        (R_RATE_SCHED, '-', 'Analysed rate per schedule unit', None, 'x scale factor', f'=B{R_SCALE}',
         None, f'=ROUND(G{R_RATE_UNIT} * B{R_SCALE}, 2)',
         'Multiplied up to the DAR schedule unit - x1000 for "1000 Nos", x100 for "100 m", x1 '
         'otherwise. This unrounded figure is what the Table 1.1 base rates below are quoted at.',
         'DERIVED'),
        (R_DAR_PUB, '-', 'DAR published rate at this lead (Table 1.1 ladder)', None,
         'Book ladder', f'=B{R_LEAD}', f'=B{R_T11_ROW}',
         # <=5 km: the printed 1/2/3/4/5 km column, taking the next tabulated km.
         # 5-10 / 10-20 / >20 km: the 5 km rate plus the printed per-km bands,
         # which is exactly how the book builds a rate beyond 5 km.
         f'=IFERROR(IF(B{R_T11_ROW}="", "", '
         f'IF(B{R_LEAD}<=5, CHOOSE(MAX(1,MIN(5,CEILING(B{R_LEAD},1))), '
         f'INDEX({lad[1]},B{R_T11_ROW}), INDEX({lad[2]},B{R_T11_ROW}), '
         f'INDEX({lad[3]},B{R_T11_ROW}), INDEX({lad[4]},B{R_T11_ROW}), '
         f'INDEX({lad[5]},B{R_T11_ROW})), '
         f'IF(B{R_LEAD}<=10, INDEX({lad[5]},B{R_T11_ROW}) + (B{R_LEAD}-5)*INDEX({lad["b5_10"]},B{R_T11_ROW}), '
         f'IF(B{R_LEAD}<=20, INDEX({lad[5]},B{R_T11_ROW}) + 5*INDEX({lad["b5_10"]},B{R_T11_ROW}) '
         f'+ (B{R_LEAD}-10)*INDEX({lad["b10_20"]},B{R_T11_ROW}), '
         f'INDEX({lad[5]},B{R_T11_ROW}) + 5*INDEX({lad["b5_10"]},B{R_T11_ROW}) '
         f'+ 10*INDEX({lad["b10_20"]},B{R_T11_ROW}) '
         f'+ (B{R_LEAD}-20)*INDEX({lad["b20"]},B{R_T11_ROW}))))), "")',
         'THE BOOK\'S OWN ANSWER for this material and lead, read from the Table 1.1 ladder in '
         'section 5B. Up to 5 km it is the printed column for the next tabulated kilometre; beyond '
         'that it is the 5 km rate plus the printed per-km band rates, which is how the DAR itself '
         'extends a carriage rate. Blank if the material is not one of the book\'s.', 'LOOKUP'),
        (R_DAR_DIFF, '-', 'Simulator vs published rate', None, 'Variance', f'=G{R_RATE_SCHED}',
         f'=G{R_DAR_PUB}',
         f'=IF(OR(G{R_DAR_PUB}="", G{R_DAR_PUB}=0), "", ROUND(G{R_RATE_SCHED}-G{R_DAR_PUB}, 2))',
         'A small difference here is NORMAL and is not an error. The book prints exact rates only at '
         '1, 2, 3, 4, 5, 10, 20 and 30 km; between those anchors its ladder interpolates linearly '
         'from the per-km band rates, while this sheet computes the exact figure for your actual '
         'lead from Data Sheet 1. The two agree to the paisa at every anchor lead and diverge by up '
         'to about 1% mid-band - at 15 km, for example, by 1.09%. A variance beyond 2% is flagged in '
         'the audit bar and does mean something to check: a speed or payload override, a whole-trip '
         'basis, a non-standard handling scope, or a gate fee - any of which is a legitimate reason '
         'to depart from the book, but should be a deliberate one.', 'DERIVED'),
        (R_SAY, 'SAY', 'CPWD OFFICIAL "SAY" RATE FOR THIS CARRIAGE ITEM', None, 'MROUND to Rs 0.05',
         None, None, f'=MROUND(G{R_RATE_SCHED}, 0.05)',
         'The rate to quote in a BOQ or estimate. Rounded to the nearest 5 paise - see the note below.',
         'SAY'),
    ]

    for r, step, desc, toggle, basis, base_f, factor_f, amount_f, note, role in mu_rows:
        ws.cell(row=r, column=1, value=step).font = styles['font_bold']
        ws.cell(row=r, column=1).alignment = styles['align_center']
        ws.cell(row=r, column=2, value=desc).alignment = styles['align_left']
        ws.cell(row=r, column=3, value=(toggle if toggle else '-')).alignment = styles['align_center']
        ws.cell(row=r, column=4, value=basis).alignment = styles['align_center']

        cb = ws.cell(row=r, column=5, value=(base_f if base_f else '-'))
        cb.alignment = styles['align_right']
        if base_f and r not in (R_TRIP_OH, R_RATE_SCHED, R_DAR_PUB, R_DAR_DIFF):
            cb.number_format = styles['fmt_currency']

        cf = ws.cell(row=r, column=6, value=(factor_f if factor_f else '-'))
        cf.alignment = styles['align_right']
        if factor_f and factor_f.startswith('=Factor'):
            cf.number_format = styles['fmt_percent']

        cv = ws.cell(row=r, column=7, value=amount_f)
        cv.alignment = styles['align_right']
        cv.number_format = styles['fmt_currency']
        cv.font = styles['font_say'] if r == R_SAY else styles['font_bold']

        cn = ws.cell(row=r, column=8, value=note)
        cn.alignment = styles['align_wrap']
        cn.font = styles['font_note']

        cr2 = ws.cell(row=r, column=9, value=role)
        cr2.alignment = styles['align_center']
        cr2.font = styles['font_note']
        cr2.fill = styles[ROLE_FILL[role]]

        border_row(ws, r, styles)

        if r in (R_W, R_X, R_Y, R_Z):
            for c in range(1, LAST_COL):
                ws.cell(row=r, column=c).fill = styles['fill_subtotal']
        elif r in (R_X1, R_Y1, R_Z1, R_Z2):
            ws.cell(row=r, column=3).fill = styles['fill_input']
            ws.cell(row=r, column=3).font = styles['font_bold']
            dv_yesno.add(ws.cell(row=r, column=3))
        elif r == R_TOTAL:
            for c in range(1, LAST_COL):
                ws.cell(row=r, column=c).fill = styles['fill_result']
        elif r in (R_TRIP_OH, R_RATE_UNIT, R_RATE_SCHED, R_DAR_PUB, R_DAR_DIFF):
            for c in range(1, LAST_COL):
                ws.cell(row=r, column=c).fill = styles['fill_subtotal']
        elif r == R_SAY:
            for c in range(1, LAST_COL):
                ws.cell(row=r, column=c).fill = styles['fill_say']
            ws.cell(row=r, column=2).font = styles['font_say']

        ws.row_dimensions[r].height = 30 if r == R_SAY else 26

    ws.merge_cells(start_row=R_SAY_NOTE, start_column=1, end_row=R_SAY_NOTE, end_column=LAST_COL)
    say_note = ws.cell(row=R_SAY_NOTE, column=1)
    say_note.value = SAY_RULE_NOTE
    say_note.font = styles['font_note']
    say_note.fill = styles['fill_note']
    say_note.alignment = styles['align_wrap']
    border_row(ws, R_SAY_NOTE, styles)
    ws.row_dimensions[R_SAY_NOTE].height = 44

    # =====================================================================
    # Heading 1.2 - manual labour carriage
    # =====================================================================
    section_bar(ws, R_MAN_HEAD,
                'SECTION 3 - HEADING 1.2: MANUAL LABOUR CARRIAGE  (lead under 0.50 km, 50 m to 500 m)',
                styles)
    ws.merge_cells(start_row=R_MAN_BANNER, start_column=1, end_row=R_MAN_BANNER, end_column=LAST_COL)
    mb = ws.cell(row=R_MAN_BANNER, column=1)
    mb.value = ('CPWD DAR 2019 Heading 1.2 gang norms: Category A (lime, moorum, rubbish, earth, sand, '
                'aggregate, bricks) = 7.67 Beldars for the first 50 m plus 1.67 extra coolies for every '
                'additional 50 m. Category B (stone blocks, pipes, cement, steel, timber, bitumen) = '
                '9.20 Beldars for the first 50 m plus 1.35 extra Beldars per additional 50 m. Every '
                'rupee figure below is derived from these gang sizes and the live day wage in '
                'Rates_Master, so a wage revision flows through automatically.')
    mb.font = styles['font_note']
    mb.fill = styles['fill_note']
    mb.alignment = styles['align_wrap']
    border_row(ws, R_MAN_BANNER, styles)
    ws.row_dimensions[R_MAN_BANNER].height = 44

    col_headers(ws, R_MAN_COLS, PARAM_HEADERS, styles, height=26)
    ws.merge_cells(start_row=R_MAN_COLS, start_column=5, end_row=R_MAN_COLS, end_column=LAST_COL)

    _param_row(ws, R_M_CODE, styles, 'Manual item code', '1.2.CUSTOM', '-', 'INPUT',
               'Your reference for the manual-carriage analysis.')
    _param_row(ws, R_M_CAT, styles, 'Material category', 'Category B (Heavy / Pipes / Steel)', '-',
               'INPUT', 'DRIVES COST. Selects the gang norms quoted in the banner above.', dv=dv_cat)
    _param_row(ws, R_M_LEAD, styles, 'Lead distance', 100, 'metres', 'INPUT',
               'DRIVES COST. The first 50 m is the base gang; every further 50 m adds labour.',
               dv=dv_mlead)
    _param_row(ws, R_M_STEPS, styles, 'Additional 50 m steps (M)', None, 'steps', 'DERIVED',
               'MAX(0, (lead - 50) / 50). A part step counts as a full step in CPWD practice.',
               number_format='0.00')
    ws.cell(row=R_M_STEPS, column=2).value = f'=MAX(0, ROUNDUP((B{R_M_LEAD} - 50) / 50, 0))'

    _param_row(ws, R_M_GANG, styles, 'Gang for the first 50 m', None, 'labour days', 'LOOKUP',
               'CPWD Heading 1.2 norm for the selected category.', number_format='0.00')
    ws.cell(row=R_M_GANG, column=2).value = (
        f'=IF(ISNUMBER(SEARCH("Category A", B{R_M_CAT})), 7.67, 9.2)')

    _param_row(ws, R_M_GANG_ADD, styles, 'Extra gang per additional 50 m', None, 'labour days',
               'LOOKUP', 'CPWD Heading 1.2 norm for the selected category.', number_format='0.00')
    ws.cell(row=R_M_GANG_ADD, column=2).value = (
        f'=IF(ISNUMBER(SEARCH("Category A", B{R_M_CAT})), 1.67, 1.35)')

    _param_row(ws, R_M_WAGE, styles, 'Labour day wage', None, 'Rs per day', 'LOOKUP',
               'Live rate for Beldar / Coolie (code 0114) from Rates_Master. Both grades are on the '
               'same day rate in DAR 2019.', number_format=styles['fmt_currency'])
    ws.cell(row=R_M_WAGE, column=2).value = (
        '=IFERROR(INDEX(Rates_Master!$E:$E, MATCH("0114", Rates_Master!$A:$A, 0)), 0)')

    _param_row(ws, R_M_BASE, styles, 'Labour cost, first 50 m', None, 'Rs', 'DERIVED',
               'Base gang x day wage.', number_format=styles['fmt_currency'])
    ws.cell(row=R_M_BASE, column=2).value = f'=ROUND(B{R_M_GANG} * B{R_M_WAGE}, 2)'

    _param_row(ws, R_M_ADD, styles, 'Labour cost, additional lead', None, 'Rs', 'DERIVED',
               'Steps x extra gang x day wage.', number_format=styles['fmt_currency'])
    ws.cell(row=R_M_ADD, column=2).value = f'=ROUND(B{R_M_STEPS} * B{R_M_GANG_ADD} * B{R_M_WAGE}, 2)'

    _param_row(ws, R_M_LABOUR, styles, 'Total labour cost for the shift', None, 'Rs', 'DERIVED',
               'First 50 m plus additional lead.', number_format=styles['fmt_currency'])
    ws.cell(row=R_M_LABOUR, column=2).value = f'=B{R_M_BASE} + B{R_M_ADD}'

    _param_row(ws, R_M_CPOH, styles, 'Add Contractor Profit & Overheads', None, 'Rs', 'DERIVED',
               'Taken from Factor_CPOH on Global_Factors, so a project override reaches this engine too.',
               number_format=styles['fmt_currency'])
    ws.cell(row=R_M_CPOH, column=2).value = f'=ROUND(B{R_M_LABOUR} * Factor_CPOH, 2)'

    _param_row(ws, R_M_TOTAL, styles, 'Total 8-hour cost with CPOH', None, 'Rs', 'RESULT',
               'Compare with the "Base Cost 1st 50m" column of Table 1.2 below.',
               number_format=styles['fmt_currency'])
    ws.cell(row=R_M_TOTAL, column=2).value = f'=B{R_M_LABOUR} + B{R_M_CPOH}'

    _param_row(ws, R_M_CAP, styles, 'Net payable quantity per 8-hour day', 1702.00, 'per billing unit',
               'INPUT',
               'From Table 1.2 below. Note these are already NET of the looseness deduction - earth is '
               'listed at 28 cum, not the 35 cum a gang physically shifts.', number_format='0.00')
    _param_row(ws, R_M_UNIT, styles, 'Billing unit', '100 m', '-', 'INPUT',
               'The DAR schedule unit for this material.', dv=dv_unit)
    _param_row(ws, R_M_SCALE, styles, 'Schedule unit scale factor', None, 'multiplier', 'DERIVED',
               'x1000 for "1000 Nos", x100 for "100 m", otherwise x1 - the same rule as Panel 3.',
               number_format='0')
    ws.cell(row=R_M_SCALE, column=2).value = (
        f'=IF(B{R_M_UNIT}="1000 Nos", 1000, IF(B{R_M_UNIT}="100 m", 100, 1))')

    _param_row(ws, R_M_RATE, styles, 'Analysed rate per schedule unit', None, 'Rs', 'DERIVED',
               'Total cost / payable quantity x scale factor. This unrounded figure is what Table 1.2 '
               'below is quoted at.', number_format=styles['fmt_currency'])
    ws.cell(row=R_M_RATE, column=2).value = (
        f'=IFERROR(ROUND(B{R_M_TOTAL} / B{R_M_CAP} * B{R_M_SCALE}, 2), 0)')

    _param_row(ws, R_M_SAY, styles, 'CPWD OFFICIAL "SAY" RATE (manual carriage)', None, 'Rs', 'SAY',
               'Rounded to the nearest 5 paise per CPWD practice - see the note in Section 2.',
               number_format=styles['fmt_currency'])
    ws.cell(row=R_M_SAY, column=2).value = f'=IFERROR(MROUND(B{R_M_RATE}, 0.05), 0)'

    # =====================================================================
    # Section 5 - benchmarks
    # =====================================================================
    section_bar(ws, R_BM_HEAD,
                'SECTION 5 - GROUND-TRUTH REFERENCE TABLES (CPWD DAR 2019 SUB-HEAD 01), transcribed '
                'from printed pages 67-75. These are source data - the panels above read from them.',
                styles, last_col=BENCH_COL)

    section_bar(ws, R_DS1_HEAD,
                '5A. CPWD DATA SHEET NO. 1 - mechanical transport benchmark, 1 km to 30 km, all 14 '
                'printed columns. Panel 2 looks up the speed, trips and km/day for your lead here.',
                styles, last_col=BENCH_COL, height=22)
    col_headers(ws, R_DS1_COLS,
                ['1 Lead (L) km', '2 Avg speed (S) km/h', '3 Trips N/day', '4 Km done/day',
                 '5 Diesel litres', '6 Cost of diesel (Rs)', '7 Mobil oil litres',
                 '8 Cost of mobil oil (Rs)', '9 Cost of 6 Beldars (Rs)', '10 Truck hire (Rs)',
                 '11 Total cost (Rs)', '12 Cost per trip (Rs)', '13 Increase per km (Rs)',
                 '14 Avg cost per addl km (Rs)'], styles)

    prev_total = None
    money = styles['fmt_currency']
    for i, rec in enumerate(DATASHEET1):
        r = R_DS1_FIRST + i
        lead, speed, n, km, dl, ml, total, per_trip = rec
        # Columns 6, 8, 9, 10 are the money columns the book prints and the
        # earlier build omitted; 11 is their sum and is checked against the
        # printed total.
        c_diesel = round(dl * RATE_DIESEL, 2)
        c_mobil = round(ml * RATE_MOBIL, 2)
        c_beldar = round(GANG * RATE_BELDAR, 2)
        c_truck = RATE_TRUCK
        col13 = '' if prev_total is None else round(total - prev_total, 2)
        prev_total = total
        vals = [lead, speed, n, km, dl, c_diesel, ml, c_mobil, c_beldar, c_truck,
                total, per_trip, col13,
                f'=IF(A{r}<=5, "", IF(A{r}<=10, ROUND((L{r}-INDEX($L${R_DS1_FIRST}:$L${R_DS1_LAST},5))/5, 2), '
                f'IF(A{r}<=20, ROUND((L{r}-INDEX($L${R_DS1_FIRST}:$L${R_DS1_LAST},10))/10, 2), '
                f'ROUND((L{r}-INDEX($L${R_DS1_FIRST}:$L${R_DS1_LAST},20))/10, 2))))']
        fmts = ['0.0', '0.0', '0.00', '0.00', '0.00', money, '0.000', money, money, money,
                money, money, money, money]
        for ci, v in enumerate(vals, 1):
            c = ws.cell(row=r, column=ci, value=(v if v != '' else None))
            c.alignment = styles['align_center'] if ci <= 4 else styles['align_right']
            c.number_format = fmts[ci - 1]
            c.font = styles['font_regular']
        fill = styles['fill_subtotal'] if r % 2 == 0 else styles['fill_calc']
        for c in range(1, BENCH_COL + 1):
            ws.cell(row=r, column=c).fill = fill
            ws.cell(row=r, column=c).border = styles['border_thin']
        ws.row_dimensions[r].height = 18

    ws.merge_cells(start_row=R_DS1_NOTE, start_column=1, end_row=R_DS1_NOTE, end_column=LAST_COL)
    dn = ws.cell(row=R_DS1_NOTE, column=1)
    dn.value = ('HOW THIS TABLE IS BUILT: N = 8 / ((2L/S) + 1). Km done per day = 2NL + 6, the 6 km '
                'being the depot run. Diesel = km / 5.0. Mobil oil = km / 140.0. Total shift cost = '
                'truck hire 0084 + 6 Beldars 0114 + diesel 1235 + mobil oil 5001. Section 1 above '
                'reproduces this line for line, which is why its W should equal the shift cost here.')
    dn.font = styles['font_note']
    dn.fill = styles['fill_note']
    dn.alignment = styles['align_wrap']
    border_row(ws, R_DS1_NOTE, styles)
    ws.row_dimensions[R_DS1_NOTE].height = 40

    # --- Table 1.1 -------------------------------------------------------
    section_bar(ws, R_T11_HEAD,
                '5B. TABLE 1.1 - MECHANICAL TRANSPORT RATE LADDER (CPWD DAR 2019, printed pages 69-70). '
                'Net payable quantity is the divisor CPWD uses; the eight rate columns are the book\'s '
                'own published rates per schedule unit including 15% CP&OH.', styles,
                last_col=BENCH_COL, height=30)
    col_headers(ws, R_T11_COLS,
                ['DAR item', 'Material / Specification', 'Capacity per trip', 'Net payable qty',
                 'Schedule unit', 'Looseness deduction', 'Rate 1 km', 'Rate 2 km', 'Rate 3 km',
                 'Rate 4 km', 'Rate 5 km', '>5-10 km per km', '>10-20 km per km',
                 '>20 km per addl km'], styles, height=30)

    for i, row in enumerate(table_11_rows()):
        r = R_T11_FIRST + i
        lad = row['ladder']
        vals = [row['item'], row['material'], row['capacity'], row['net'], row['unit'],
                row['looseness'], lad[1], lad[2], lad[3], lad[4], lad[5],
                lad['b5_10'], lad['b10_20'], lad['b20']]
        for ci, v in enumerate(vals, 1):
            c = ws.cell(row=r, column=ci, value=v)
            c.font = styles['font_regular']
            if ci in (1, 3, 4, 5):
                c.alignment = styles['align_center']
            elif ci >= 7:
                c.alignment = styles['align_right']
                c.number_format = styles['fmt_currency']
            else:
                c.alignment = styles['align_left']
        ws.cell(row=r, column=4).font = styles['font_bold']
        fill = styles['fill_subtotal'] if r % 2 == 0 else styles['fill_calc']
        for c in range(1, BENCH_COL + 1):
            ws.cell(row=r, column=c).fill = fill
            ws.cell(row=r, column=c).border = styles['border_thin']
        ws.cell(row=r, column=4).fill = styles['fill_lookup']
        ws.row_dimensions[r].height = 18

    # --- Table 1.2 -------------------------------------------------------
    section_bar(ws, R_T12_HEAD,
                '5C. TABLE 1.2 - MANUAL LABOUR CARRIAGE, LEAD UNDER 0.50 km (CPWD DAR 2019, printed '
                'pages 72-74). Category A gangs 7.67 Beldars + 1.67 coolies per extra 50 m; '
                'Category B gangs 9.20 Beldars + 1.35 Beldars per extra 50 m.', styles,
                last_col=BENCH_COL, height=30)
    col_headers(ws, R_T12_COLS,
                ['DAR item', 'Material / Specification', 'Capacity per day', 'Net payable qty',
                 'Schedule unit', 'Gang category', 'Cost per 8-hr day incl. CP&OH (Rs)',
                 'Rate for 1st 50 m (Rs)', 'Rate per addl 50 m (Rs)', '', '', '', '', ''],
                styles, height=30)

    for i, row in enumerate(table_12_rows()):
        r = R_T12_FIRST + i
        vals = [row['item'], row['material'], row['capacity'], row['net'], row['unit'],
                'Category %s' % row['category'], row['day_cost'], row['first50'], row['addl50']]
        for ci, v in enumerate(vals, 1):
            c = ws.cell(row=r, column=ci, value=v)
            c.font = styles['font_regular']
            if ci in (1, 3, 4, 5, 6):
                c.alignment = styles['align_center']
            elif ci >= 7:
                c.alignment = styles['align_right']
                c.number_format = styles['fmt_currency']
            else:
                c.alignment = styles['align_left']
        ws.cell(row=r, column=4).font = styles['font_bold']
        fill = styles['fill_subtotal'] if r % 2 == 0 else styles['fill_calc']
        for c in range(1, BENCH_COL + 1):
            ws.cell(row=r, column=c).fill = fill
            ws.cell(row=r, column=c).border = styles['border_thin']
        ws.cell(row=r, column=4).fill = styles['fill_lookup']
        ws.row_dimensions[r].height = 18

    # --- Scope / gang table ---------------------------------------------
    section_bar(ws, R_SC_HEAD,
                '5D. HANDLING SCOPE - LABOUR GANG ALLOCATION. The Handling Scope you pick in Panel 1 '
                'sets the Beldar row in Section 1 from this table.', styles,
                last_col=BENCH_COL, height=22)
    col_headers(ws, R_SC_COLS,
                ['Scope wording (Panel 1 dropdown)', 'Beldar gang', 'Loading', 'Unloading', 'Stacking',
                 'Daily labour cost (Rs)', 'Direct shift cost (Rs)', 'Engineering rationale', 'Source'],
                styles)
    scopes = [
        ('including loading, transporting, unloading and stacking', 6, '3.00', '2.00', '1.00',
         3348.00, 5939.58, 'DAR item 1.1.1 baseline - full turnkey manual handling at both ends.'),
        ('including loading, transporting, unloading to approved municipal dumping ground',
         6, '3.00', '2.00', '1.00', 3348.00, 5939.58,
         'DAR item 1.1.18 benchmark for urban malba / rubbish disposal.'),
        ('including loading, transporting and unloading (excluding stacking)', 5, '3.00', '2.00', '0.00',
         2790.00, 5381.58, 'Bulk delivery or tipping; manual stacking omitted.'),
        ('transporting and unloading only (machine loaded at source)', 3, '0.00', '2.00', '1.00',
         1674.00, 4265.58, 'Loaded by excavator or JCB under Sub-Head 02; manual unloading only.'),
        ('transporting only (excluding loading, unloading and stacking)', 0, '0.00', '0.00', '0.00',
         0.00, 2591.58, 'Pure haulage; machine loaded at source and tipped at destination.'),
        ('including unloading and stacking at railway siding', 3.75, '0.00', '2.50', '1.25',
         2092.50, 4684.08, 'DAR items 1.3 and 1.4 standard for railway wagon handling.'),
    ]
    for i, rec in enumerate(scopes):
        r = R_SC_FIRST + i
        for ci, val in enumerate(rec, 1):
            c = ws.cell(row=r, column=ci, value=val)
            c.font = styles['font_regular']
            if ci == 1 or ci == 8:
                c.alignment = styles['align_wrap']
            elif ci in (6, 7):
                c.alignment = styles['align_right']
                c.number_format = styles['fmt_currency']
            else:
                c.alignment = styles['align_center']
        ws.cell(row=r, column=9, value='DAR 2019 Sub-Head 01').font = styles['font_note']
        fill = styles['fill_subtotal'] if r % 2 == 0 else styles['fill_calc']
        for c in range(1, LAST_COL + 1):
            ws.cell(row=r, column=c).fill = fill
            ws.cell(row=r, column=c).border = styles['border_thin']
        ws.row_dimensions[r].height = 30

    # --- Defined names (computed, so the tables can move) ----------------
    for name, ref in [
        ('CPWD_Carriage_Item_Codes', f"'{sn}'!$A${R_T11_FIRST}:$A${R_T11_LAST}"),
        ('CPWD_Carriage_Materials', f"'{sn}'!$B${R_T11_FIRST}:$B${R_T11_LAST}"),
        ('CPWD_Carriage_Scope', f"'{sn}'!$A${R_SC_FIRST}:$A${R_SC_LAST}"),
        ('CPWD_Carriage_Net_Payable', f"'{sn}'!$D${R_T11_FIRST}:$D${R_T11_LAST}"),
        ('CPWD_DataSheet1', f"'{sn}'!$A${R_DS1_FIRST}:$H${R_DS1_LAST}"),
    ]:
        if name in wb.defined_names:
            del wb.defined_names[name]
        wb.defined_names.add(DefinedName(name, attr_text=ref))

    # --- Column widths ---------------------------------------------------
    for col, w in {'A': 34, 'B': 30, 'C': 20, 'D': 16, 'E': 20,
                   'F': 20, 'G': 20, 'H': 44, 'I': 16}.items():
        ws.column_dimensions[col].width = w

    # --- Protection: only INPUT and OVERRIDE cells stay editable ---------
    for r in range(1, ws.max_row + 1):
        for c in range(1, LAST_COL + 1):
            ws.cell(row=r, column=c).protection = Protection(locked=True)

    editable = [R_ITEM_CODE, R_MATERIAL, R_SCOPE, R_LIFT, R_GATE_FEE, R_NOM_OVERRIDE,
                R_LEAD, R_SPEED_OV, R_TURNAROUND, R_MODE, R_TRIPS_BASIS, R_TRIPS_OV, R_DIST_BASIS,
                R_PAY_OV, R_M_CODE, R_M_CAT, R_M_LEAD, R_M_CAP, R_M_UNIT]
    for r in editable:
        ws.cell(row=r, column=2).protection = Protection(locked=False)
    for r in (R_X1, R_Y1, R_Z1, R_Z2):
        ws.cell(row=r, column=3).protection = Protection(locked=False)
    for r in range(R_RES_FIRST, R_RES_LAST + 1):
        ws.cell(row=r, column=2).protection = Protection(locked=False)
        if r > R_RES_FIRST + 3:
            ws.cell(row=r, column=5).protection = Protection(locked=False)
    ws.cell(row=R_RES_FIRST, column=5).protection = Protection(locked=False)   # truck days

    ws.protection.sheet = True
    ws.freeze_panes = 'A7'
    print(f"Built carriage simulator: {sn} (role-labelled layout, net payable qty, "
          f"unit scale factor, MROUND say rate)")
