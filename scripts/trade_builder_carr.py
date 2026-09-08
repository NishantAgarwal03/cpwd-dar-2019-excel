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
R_MODE, R_TRIPS_OV, R_TRIPS, R_DIST_BASIS, R_DISTANCE = 25, 26, 27, 28, 29
R_DIESEL, R_MOBIL = 30, 31

R_P3_HEAD, R_P3_COLS = 33, 34
R_PAY_GROSS, R_PAY_NET, R_PAY_OV, R_PAY_EFF = 35, 36, 37, 38
R_UNIT, R_SCALE, R_OUTPUT = 39, 40, 41

R_AUDIT = 43

R_RES_HEAD, R_RES_COLS, R_RES_FIRST = 45, 46, 47
RES_ROWS = 8
R_RES_LAST = R_RES_FIRST + RES_ROWS - 1                # 54
R_W_SUB = 55
R_TRIP_COST = 56

R_MU_HEAD, R_MU_COLS = 58, 59
R_W, R_X1, R_X, R_Y1, R_Y, R_Z1, R_Z, R_Z2 = 60, 61, 62, 63, 64, 65, 66, 67
R_TOTAL, R_TRIP_OH, R_RATE_UNIT, R_RATE_SCHED, R_SAY, R_SAY_NOTE = 68, 69, 70, 71, 72, 73

R_MAN_HEAD, R_MAN_BANNER, R_MAN_COLS = 75, 76, 77
R_M_CODE, R_M_CAT, R_M_LEAD, R_M_STEPS = 78, 79, 80, 81
R_M_GANG, R_M_GANG_ADD, R_M_WAGE = 82, 83, 84
R_M_BASE, R_M_ADD, R_M_LABOUR, R_M_CPOH, R_M_TOTAL = 85, 86, 87, 88, 89
R_M_CAP, R_M_UNIT, R_M_SCALE, R_M_RATE, R_M_SAY = 90, 91, 92, 93, 94

R_BM_HEAD = 96
R_DS1_HEAD, R_DS1_COLS, R_DS1_FIRST = 97, 98, 99
DS1_ROWS = 30
R_DS1_LAST = R_DS1_FIRST + DS1_ROWS - 1                # 128
R_DS1_NOTE = 129

R_T11_HEAD, R_T11_COLS, R_T11_FIRST = 131, 132, 133
T11_ROWS = 27
R_T11_LAST = R_T11_FIRST + T11_ROWS - 1                # 159

R_T12_HEAD, R_T12_COLS, R_T12_FIRST = 161, 162, 163
T12_ROWS = 6
R_T12_LAST = R_T12_FIRST + T12_ROWS - 1                # 168

R_SC_HEAD, R_SC_COLS, R_SC_FIRST = 170, 171, 172
SC_ROWS = 6
R_SC_LAST = R_SC_FIRST + SC_ROWS - 1                   # 177

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

    # --- Data validations -------------------------------------------------
    def mkdv(formula, blank=True):
        d = DataValidation(type='list', formula1=formula, allow_blank=blank,
                           showErrorMessage=False)
        ws.add_data_validation(d)
        return d

    dv_yesno = mkdv('"YES,NO"', False)
    dv_mode = mkdv('"STANDARD (DAYTIME),URBAN RESTRICTED HOURS"', False)
    dv_basis = mkdv('"DIRECT ROUTE (2NL + 6),CPWD PRO-RATA (Data Sheet 1)"', False)
    dv_unit = mkdv('"cum,tonne,metre,100 m,1000 Nos"', False)
    dv_cat = mkdv('"Category A (Bulk / Earth / Bricks),Category B (Heavy / Pipes / Steel)"', False)
    dv_mlead = mkdv('"50,100,150,200,250,300,350,400,450,500"', False)
    dv_items = mkdv('=CPWD_Carriage_Item_Codes', False)
    dv_mats = mkdv('=CPWD_Carriage_Materials', False)
    dv_scope = mkdv('=CPWD_Carriage_Scope', False)
    dv_lift = mkdv('"for all lifts,for lift upto 1.5 m,with mechanical lift,'
                   'for all lifts and leads"', False)

    ws.row_dimensions[7].height = 8
    ws.row_dimensions[17].height = 8
    ws.row_dimensions[32].height = 8
    ws.row_dimensions[42].height = 8
    ws.row_dimensions[44].height = 8
    ws.row_dimensions[57].height = 8
    ws.row_dimensions[74].height = 8
    ws.row_dimensions[95].height = 8

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

    _param_row(ws, R_MATERIAL, styles, 'Material commodity',
               config.get('default_material', 'R.C.C./C.I./Steel pipes 1000, 1100 & 1200 mm dia'),
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

    _param_row(ws, R_TRIPS_OV, styles, 'Trips override', None, 'trips per shift', 'OVERRIDE',
               'Leave blank to use the calculated trips below. Fill it in to force a trip count - this '
               'is how a restricted-hours item is priced. It is honoured in any mode.',
               number_format='0.00')

    _param_row(ws, R_TRIPS, styles, 'Daily trips achieved (N)', None, 'trips per shift', 'DERIVED',
               'N = 8 / ((2L / S) + T) - an 8-hour shift divided by one round trip. Overridden by the '
               'row above when that is filled.', number_format='0.00')
    ws.cell(row=R_TRIPS, column=2).value = (
        f'=IF(B{R_TRIPS_OV}<>"", B{R_TRIPS_OV}, '
        f'IFERROR(ROUND(8 / ((2 * B{R_LEAD} / B{R_SPEED_EFF}) + B{R_TURNAROUND}), 2), 0))')

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

    _param_row(ws, R_PAY_GROSS, styles, 'Gross truck payload per trip', None, 'per Table 1.1', 'LOOKUP',
               'What physically goes on the truck, from Table 1.1 below. Shown for information only - '
               'it is NOT what the rate is divided by.', number_format='0.00')
    ws.cell(row=R_PAY_GROSS, column=2).value = (
        f'=IFERROR(INDEX({t11_gross}, MATCH(B{R_MATERIAL}, {t11_mat}, 0)), "")')

    _param_row(ws, R_PAY_NET, styles, 'Net payable quantity per trip', None, 'per Table 1.1', 'LOOKUP',
               'THIS is the divisor CPWD uses. Loose materials are paid on a reduced quantity: earth '
               'less 20%, excavated rock less 50%, soling stone less 15%, 40 mm aggregate and manure '
               'less 8%. Using the gross payload instead understates the rate by that percentage.',
               number_format='0.00')
    ws.cell(row=R_PAY_NET, column=2).value = (
        f'=IFERROR(INDEX({t11_net}, MATCH(B{R_MATERIAL}, {t11_mat}, 0)), "")')

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
        f'=IFERROR(INDEX({t11_unit}, MATCH(B{R_MATERIAL}, {t11_mat}, 0)), "cum")')

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
    ab.value = (f'=IF(AND(D{R_AUDIT}="OK", F{R_AUDIT}="OK", H{R_AUDIT}="OK", I{R_AUDIT}="OK"), '
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
            value=f'=IF(AND(ISNUMBER(G{R_SAY}), G{R_SAY}>0), "OK", "ERR: rate not resolved")')
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
        if base_f and r not in (R_TRIP_OH, R_RATE_SCHED):
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
        elif r in (R_TRIP_OH, R_RATE_UNIT, R_RATE_SCHED):
            for c in range(1, LAST_COL):
                ws.cell(row=r, column=c).fill = styles['fill_subtotal']
        elif r == R_SAY:
            for c in range(1, LAST_COL):
                ws.cell(row=r, column=c).fill = styles['fill_say']
            ws.cell(row=r, column=2).font = styles['font_say']

        ws.row_dimensions[r].height = 30 if r == R_SAY else 26

    ws.merge_cells(start_row=R_SAY_NOTE, start_column=1, end_row=R_SAY_NOTE, end_column=LAST_COL)
    sn = ws.cell(row=R_SAY_NOTE, column=1)
    sn.value = SAY_RULE_NOTE
    sn.font = styles['font_note']
    sn.fill = styles['fill_note']
    sn.alignment = styles['align_wrap']
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
                'SECTION 5 - GROUND-TRUTH REFERENCE TABLES (CPWD DAR 2019 SUB-HEAD 01). '
                'These are source data - the panels above read from them. Do not edit.', styles)

    section_bar(ws, R_DS1_HEAD,
                '5A. CPWD DATA SHEET NO. 1 - mechanical transport benchmark, 1 km to 30 km. '
                'Panel 2 looks up the speed, trips and km/day for your lead here.', styles, height=22)
    col_headers(ws, R_DS1_COLS,
                ['Lead (L) km', 'Avg Speed (S) km/h', 'Trips (N)/day', 'Km done/day',
                 'Diesel (litres)', 'Mobil oil (litres)', 'Total shift cost (Rs)',
                 'Cost per trip (Rs)', 'Source'], styles)

    datasheet1 = [
        (1.0, 16.0, 7.11, 20.22, 4.04, 0.144, 5190.30, 730.00),
        (2.0, 17.0, 6.48, 31.92, 6.38, 0.228, 5388.75, 831.60),
        (3.0, 17.5, 5.96, 41.76, 8.35, 0.298, 5555.60, 932.15),
        (4.0, 18.0, 5.54, 50.32, 10.06, 0.359, 5700.50, 1028.97),
        (5.0, 18.5, 5.19, 57.90, 11.58, 0.414, 5829.54, 1123.23),
        (6.0, 19.0, 4.90, 64.80, 12.96, 0.463, 5946.41, 1213.55),
        (7.0, 19.5, 4.66, 71.24, 14.25, 0.509, 6055.72, 1299.51),
        (8.0, 20.0, 4.44, 77.04, 15.41, 0.550, 6153.89, 1386.01),
        (9.0, 20.5, 4.26, 82.68, 16.54, 0.591, 6249.86, 1467.10),
        (10.0, 21.0, 4.10, 88.00, 17.60, 0.629, 6339.74, 1546.28),
        (11.0, 21.5, 3.95, 92.90, 18.58, 0.664, 6422.79, 1626.02),
        (12.0, 22.0, 3.83, 97.92, 19.58, 0.699, 6507.32, 1699.04),
        (13.0, 22.5, 3.71, 102.46, 20.49, 0.732, 6584.60, 1774.82),
        (14.0, 23.0, 3.61, 107.08, 21.42, 0.765, 6663.35, 1845.80),
        (15.0, 23.5, 3.51, 111.30, 22.26, 0.795, 6734.54, 1918.67),
        (16.0, 24.0, 3.43, 115.76, 23.15, 0.827, 6810.04, 1985.43),
        (17.0, 24.5, 3.35, 119.90, 23.98, 0.856, 6880.17, 2053.78),
        (18.0, 25.0, 3.28, 124.08, 24.82, 0.886, 6951.36, 2119.32),
        (19.0, 25.5, 3.21, 127.98, 25.60, 0.914, 7017.51, 2186.14),
        (20.0, 26.0, 3.15, 132.00, 26.40, 0.943, 7085.45, 2249.35),
        (21.0, 26.5, 3.09, 135.78, 27.16, 0.970, 7149.81, 2313.85),
        (22.0, 27.0, 3.04, 139.76, 27.95, 0.998, 7216.70, 2373.91),
        (23.0, 27.5, 2.99, 143.54, 28.71, 1.025, 7281.07, 2435.14),
        (24.0, 28.0, 2.95, 147.60, 29.52, 1.054, 7349.73, 2491.43),
        (25.0, 28.5, 2.90, 151.00, 30.20, 1.079, 7407.59, 2554.34),
        (26.0, 29.0, 2.86, 154.72, 30.94, 1.105, 7470.17, 2611.95),
        (27.0, 29.5, 2.83, 158.82, 31.76, 1.134, 7539.57, 2664.16),
        (28.0, 30.0, 2.79, 162.24, 32.45, 1.159, 7598.17, 2723.36),
        (29.0, 30.5, 2.76, 166.08, 33.22, 1.186, 7663.26, 2776.54),
        (30.0, 31.0, 2.73, 169.80, 33.96, 1.213, 7726.16, 2830.10),
    ]
    fmts = ['0.0', '0.0', '0.00', '0.00', '0.00', '0.000',
            styles['fmt_currency'], styles['fmt_currency']]
    for i, rec in enumerate(datasheet1):
        r = R_DS1_FIRST + i
        for ci, val in enumerate(rec, 1):
            c = ws.cell(row=r, column=ci, value=val)
            c.alignment = styles['align_center'] if ci <= 3 else styles['align_right']
            c.number_format = fmts[ci - 1]
            c.font = styles['font_regular']
        ws.cell(row=r, column=9, value='DAR 2019 Data Sheet 1').font = styles['font_note']
        fill = styles['fill_subtotal'] if r % 2 == 0 else styles['fill_calc']
        for c in range(1, LAST_COL + 1):
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
                '5B. TABLE 1.1 - MATERIAL PAYLOAD, NET PAYABLE QUANTITY AND SCHEDULE UNIT. '
                'Panel 3 looks up the material you chose here. The "Net payable" column is the one '
                'the rate is divided by.', styles, height=22)
    col_headers(ws, R_T11_COLS,
                ['DAR Code', 'Material / Trade Specification', 'Gross truck payload / trip',
                 'Net payable qty / trip', 'Schedule unit', 'Looseness deduction applied',
                 'DAR base rate @ 1 km (Rs)', 'Additional rate > 20 km (Rs)', 'Source'], styles)

    table11 = [
        ('1.1.1', 'Lime, moorum, building rubbish, malba', 8, 8, 'cum',
         'Nil - full 8.00 cum payable', 104.94, 8.35),
        ('1.1.2', 'Earth (excavated soil / good earth)', 8, 6.4, 'cum',
         '20% deduction for looseness (net 6.40 cum)', 131.17, 10.44),
        ('1.1.3', 'Manure or sludge', 8, 7.36, 'cum',
         '8% deduction for looseness (net 7.36 cum)', 114.06, 9.08),
        ('1.1.4', 'Excavated rock', 8, 4, 'cum',
         '50% deduction for voids / looseness (net 4.00 cum)', 209.88, 16.70),
        ('1.1.5', 'Sand, stone aggregate below 40 mm', 8, 8, 'cum',
         'Nil - standard density, full 8.00 cum payable', 104.94, 8.35),
        ('1.1.6', 'Stone aggregate 40 mm nominal size & above', 8, 7.36, 'cum',
         '8% deduction for voids in coarse aggregate', 114.06, 9.08),
        ('1.1.7', 'Soling stone & masonry stone', 8, 6.8, 'cum',
         '15% deduction for stack voids (net 6.80 cum)', 123.46, 9.82),
        ('1.1.8', 'Bricks (standard modular / conventional)', 3000, 3000, '1000 Nos',
         'Nil - 3,000 bricks per 9-tonne truck load', 279.83, 22.26),
        ('1.1.9', 'Brick tiles / Allahabad roofing tiles', 5000, 5000, '1000 Nos',
         'Nil - 5,000 tiles per 9-tonne truck load', 167.90, 13.36),
        ('1.1.10', 'Cement, stone blocks, Kota stone slabs', 9, 9, 'tonne',
         'Nil - rated truck payload 9 tonne', 93.28, 7.42),
        ('1.1.11', 'Steel bars, structural sections & fabric', 9, 9, 'tonne',
         'Nil - rated truck payload 9 tonne', 93.28, 7.42),
        ('1.1.12', 'Timber (scantlings / logs)', 7, 7, 'cum',
         'Nil - volume limit on timber body 7.00 cum', 119.93, 9.54),
        ('1.1.13', 'Tar, bitumen in drums', 8, 8, 'tonne',
         'Nil - packed drum loading capacity 8 tonne', 104.94, 8.35),
        ('1.1.14', 'Steam coal', 7, 7, 'tonne',
         'Nil - bulk density restriction 7 tonne', 119.93, 9.54),
        ('1.1.15.1', 'S.W. pipes 100 mm dia', 600, 600, '100 m', 'Nil - 600 m per load', 139.92, 11.13),
        ('1.1.15.2', 'S.W. pipes 150 mm dia', 300, 300, '100 m', 'Nil - 300 m per load', 279.83, 22.26),
        ('1.1.16.1', 'R.C.C. / C.I. pipes 100 mm dia', 366, 366, '100 m', 'Nil - 366 m per load', 229.37, 18.25),
        ('1.1.16.3', 'R.C.C. / C.I. pipes 150 mm dia', 219.6, 219.6, '100 m', 'Nil - 219.60 m per load', 382.29, 30.42),
        ('1.1.16.4', 'R.C.C. / C.I. pipes 200 mm dia', 135, 135, '100 m', 'Nil - 135 m per load', 621.85, 49.48),
        ('1.1.16.5', 'R.C.C. / C.I. pipes 250 mm dia', 95, 95, '100 m', 'Nil - 95 m per load', 883.68, 70.31),
        ('1.1.16.6', 'R.C.C. / C.I. pipes 300 mm dia', 76.86, 76.86, '100 m', 'Nil - 76.86 m per load', 1092.25, 86.90),
        ('1.1.16.7', 'R.C.C. / C.I. pipes 350 mm dia', 54.9, 54.9, '100 m', 'Nil - 54.90 m per load', 1529.14, 121.66),
        ('1.1.16.8', 'R.C.C. / C.I. pipes 400 mm dia', 40.26, 40.26, '100 m', 'Nil - 40.26 m per load', 2085.20, 165.90),
        ('1.1.16.9', 'R.C.C. / C.I. pipes 450 & 500 mm dia', 32.94, 32.94, '100 m', 'Nil - 32.94 m per load', 2548.57, 202.77),
        ('1.1.16.10', 'R.C.C. / C.I. pipes 600, 700, 750 & 800 mm dia', 21.96, 21.96, '100 m', 'Nil - 21.96 m per load', 3822.86, 304.15),
        ('1.1.16.11', 'R.C.C. / C.I. pipes 900 mm dia', 14.64, 14.64, '100 m', 'Nil - 14.64 m per load', 5734.29, 456.23),
        ('1.1.17.12', 'R.C.C./C.I./Steel pipes 1000, 1100 & 1200 mm dia', 10.98, 10.98, '100 m',
         'Nil - 10.98 m per load (heavy large bore)', 7645.72, 608.31),
    ]
    for i, rec in enumerate(table11):
        r = R_T11_FIRST + i
        for ci, val in enumerate(rec, 1):
            c = ws.cell(row=r, column=ci, value=val)
            c.font = styles['font_regular']
            if ci in (1, 3, 4, 5):
                c.alignment = styles['align_center']
            elif ci in (7, 8):
                c.alignment = styles['align_right']
                c.number_format = styles['fmt_currency']
            else:
                c.alignment = styles['align_left']
        ws.cell(row=r, column=4).font = styles['font_bold']
        ws.cell(row=r, column=9, value='DAR 2019 Sub-Head 01').font = styles['font_note']
        fill = styles['fill_subtotal'] if r % 2 == 0 else styles['fill_calc']
        for c in range(1, LAST_COL + 1):
            ws.cell(row=r, column=c).fill = fill
            ws.cell(row=r, column=c).border = styles['border_thin']
        ws.cell(row=r, column=4).fill = styles['fill_lookup']
        ws.row_dimensions[r].height = 18

    # --- Table 1.2 -------------------------------------------------------
    section_bar(ws, R_T12_HEAD,
                '5C. TABLE 1.2 - MANUAL LABOUR CARRIAGE REFERENCE (lead under 0.50 km). '
                'Capacities here are already net of looseness.', styles, height=22)
    col_headers(ws, R_T12_COLS,
                ['Item No.', 'Material / Specification', 'Net payable qty / day', 'Billing unit',
                 'Cost for 8 hours incl. CPOH (Rs)', 'Rate for 1st 50 m (Rs)',
                 'Rate per additional 50 m (Rs)', 'CPWD gang norm', 'Source'], styles)
    table12 = [
        ('1.2.1', 'Lime, moorum, building rubbish', 35, 'cum', 4921.84, 140.62, 30.62,
         '7.67 Beldars 1st 50 m, 1.67 addl coolie / 50 m'),
        ('1.2.2', 'Earth (already net of 20% looseness)', 28, 'cum', 4921.84, 175.78, 38.27,
         '7.67 Beldars 1st 50 m, 1.67 addl coolie / 50 m'),
        ('1.2.8', 'Bricks (standard modular)', 15000, '1000 Nos', 4921.84, 328.12, 71.44,
         '7.67 Beldars 1st 50 m, 1.67 addl coolie / 50 m'),
        ('1.2.11', 'Stone blocks, G.I., C.I. pipes below 100 mm', 46, 'tonne', 5903.64, 128.34, 18.83,
         '9.20 Beldars 1st 50 m, 1.35 addl Beldar / 50 m'),
        ('1.2.12', 'Cement in bags', 57.99, 'tonne', 5903.64, 101.80, 14.94,
         '9.20 Beldars 1st 50 m, 1.35 addl Beldar / 50 m'),
        ('1.2.17.1', 'R.C.C. / C.I. pipes 100 mm dia', 1702, '100 m', 5903.64, 346.86, 50.90,
         '9.20 Beldars 1st 50 m, 1.35 addl Beldar / 50 m'),
    ]
    for i, rec in enumerate(table12):
        r = R_T12_FIRST + i
        for ci, val in enumerate(rec, 1):
            c = ws.cell(row=r, column=ci, value=val)
            c.font = styles['font_regular']
            if ci in (1, 3, 4):
                c.alignment = styles['align_center']
            elif ci in (5, 6, 7):
                c.alignment = styles['align_right']
                c.number_format = styles['fmt_currency']
            else:
                c.alignment = styles['align_left']
        ws.cell(row=r, column=9, value='DAR 2019 Table 1.2').font = styles['font_note']
        fill = styles['fill_subtotal'] if r % 2 == 0 else styles['fill_calc']
        for c in range(1, LAST_COL + 1):
            ws.cell(row=r, column=c).fill = fill
            ws.cell(row=r, column=c).border = styles['border_thin']
        ws.row_dimensions[r].height = 18

    # --- Scope / gang table ---------------------------------------------
    section_bar(ws, R_SC_HEAD,
                '5D. HANDLING SCOPE - LABOUR GANG ALLOCATION. The Handling Scope you pick in Panel 1 '
                'sets the Beldar row in Section 1 from this table.', styles, height=22)
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
                R_LEAD, R_SPEED_OV, R_TURNAROUND, R_MODE, R_TRIPS_OV, R_DIST_BASIS,
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
