# -*- coding: utf-8 -*-
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from scripts.trade_builder import add_trade_header_and_legend

def build_carriage_trade(wb, config, styles):
    ws = wb.create_sheet(title=config['sheet_name'])
    ws.views.sheetView[0].showGridLines = True
    add_trade_header_and_legend(ws, config, styles)
    
    # Data Validations
    dv_yesno = DataValidation(type='list', formula1='"YES,NO"', allow_blank=False)
    ws.add_data_validation(dv_yesno)
    
    dv_env_mode = DataValidation(type='list', formula1='"STANDARD (DAYTIME),URBAN RESTRICTED HOURS"', allow_blank=False)
    ws.add_data_validation(dv_env_mode)
    
    dv_dist_basis = DataValidation(type='list', formula1='"CPWD PRO-RATA (Item 1.1.18),DIRECT ROUTE (2NL + 6)"', allow_blank=False)
    ws.add_data_validation(dv_dist_basis)
    
    dv_unit = DataValidation(type='list', formula1='"cum,metre,100 m,tonne,1000 Nos"', allow_blank=False)
    ws.add_data_validation(dv_unit)
    
    dv_manual_cat = DataValidation(type='list', formula1='"Category A (Bulk / Earth / Bricks),Category B (Heavy / Pipes / Steel)"', allow_blank=False)
    ws.add_data_validation(dv_manual_cat)
    
    dv_manual_lead = DataValidation(type='list', formula1='"50,100,150,200,250,300,350,400,450,500"', allow_blank=False)
    ws.add_data_validation(dv_manual_lead)

    # Keyword Data Validations for Smart Selector
    dv_materials = DataValidation(type='list', formula1='=CPWD_Carriage_Materials', allow_blank=False)
    ws.add_data_validation(dv_materials)

    dv_scope = DataValidation(
        type='list',
        formula1='"including loading, transporting, unloading and stacking,including loading, transporting, unloading to approved municipal dumping ground,including loading, transporting and unloading (excluding stacking),transporting and unloading only (machine loaded / excluding loading),transporting only (excluding loading, unloading and stacking),including unloading and stacking at railway siding"',
        allow_blank=False
    )
    ws.add_data_validation(dv_scope)

    dv_lift = DataValidation(
        type='list',
        formula1='"for all lifts,for lift upto 1.5 m,with mechanical lift,for all lifts and leads"',
        allow_blank=False
    )
    ws.add_data_validation(dv_lift)

    # Spacer Row 4
    ws.row_dimensions[4].height = 10

    # =========================================================================
    # PANEL 1: SCOPE & SPECIFICATION INPUTS (Rows 5-8)
    # =========================================================================
    ws.merge_cells('A5:H5')
    p1_head = ws['A5']
    p1_head.value = 'PANEL 1: SCOPE & SPECIFICATION INPUTS (What is being moved & contractual handling scope)'
    p1_head.font = styles['font_white_bold']
    p1_head.fill = styles['fill_header']
    p1_head.alignment = styles['align_left']
    ws.row_dimensions[5].height = 24

    # Row 6: Primary Scope Dropdowns
    ws['A6'] = 'Custom Item Code:'
    ws['A6'].font = styles['font_bold']
    ws['A6'].alignment = styles['align_right']
    ws['B6'] = config.get('default_item_code', '1.1.CUSTOM')
    ws['B6'].fill = styles['fill_input']
    ws['B6'].font = styles['font_bold']
    ws['B6'].alignment = styles['align_center']

    ws['C6'] = 'Material Commodity:'
    ws['C6'].font = styles['font_bold']
    ws['C6'].alignment = styles['align_right']
    ws['D6'] = config.get('default_material', 'R.C.C./C.I./Steel pipes 1000, 1100 & 1200 mm dia')
    ws['D6'].fill = styles['fill_input']
    ws['D6'].font = styles['font_bold']
    ws['D6'].alignment = styles['align_left']
    dv_materials.add(ws['D6'])

    ws['E6'] = 'Handling Scope:'
    ws['E6'].font = styles['font_bold']
    ws['E6'].alignment = styles['align_right']
    ws['F6'] = config.get('default_scope', 'including loading, transporting, unloading and stacking')
    ws['F6'].fill = styles['fill_input']
    ws['F6'].font = styles['font_bold']
    ws['F6'].alignment = styles['align_left']
    dv_scope.add(ws['F6'])

    ws['G6'] = 'Lift Condition:'
    ws['G6'].font = styles['font_bold']
    ws['G6'].alignment = styles['align_right']
    ws['H6'] = config.get('default_lift', 'for all lifts')
    ws['H6'].fill = styles['fill_input']
    ws['H6'].font = styles['font_bold']
    ws['H6'].alignment = styles['align_center']
    dv_lift.add(ws['H6'])

    for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']:
        ws[f'{col}6'].border = styles['border_thin']
    ws.row_dimensions[6].height = 24

    # Row 7: Tipping Fee Surcharge & Manual Override
    ws['A7'] = 'Municipal Gate Fee:'
    ws['A7'].font = styles['font_bold']
    ws['A7'].alignment = styles['align_right']
    ws['B7'] = 0.00
    ws['B7'].fill = styles['fill_input']
    ws['B7'].font = styles['font_bold']
    ws['B7'].alignment = styles['align_center']
    ws['B7'].number_format = styles['fmt_currency']

    ws['C7'] = 'Manual Override (Opt):'
    ws['C7'].font = styles['font_bold']
    ws['C7'].alignment = styles['align_right']
    ws['C7'].border = styles['border_thin']

    ws.merge_cells('D7:H7')
    ws['D7'] = None
    ws['D7'].font = styles['font_regular']
    ws['D7'].fill = styles['fill_input']
    ws['D7'].alignment = styles['align_wrap']
    ws['D7'].border = styles['border_thin']
    ws.row_dimensions[7].height = 24

    # Row 8: Assembled Item Nomenclature
    ws['A8'] = 'Assembled Nomenclature:'
    ws['A8'].font = styles['font_bold']
    ws['A8'].alignment = styles['align_right']
    ws['A8'].fill = styles['fill_note']
    ws['A8'].border = styles['border_thin']

    ws.merge_cells('B8:H8')
    ws['B8'] = '=IF(D7<>"","" & D7,"Carriage of " & D6 & " by mechanical transport " & F6 & " for lead upto " & TEXT(B11, "0.00") & " km " & H6 & IF(B7>0, ", including municipal tipping royalty/gate fee of Rs " & TEXT(B7, "0.00") & " per " & B13, "") & ", complete as per directions of Engineer-in-charge.")'
    ws['B8'].font = styles['font_bold']
    ws['B8'].fill = styles['fill_note']
    ws['B8'].alignment = styles['align_wrap']
    ws['B8'].border = styles['border_thin']
    ws.row_dimensions[8].height = 38

    # Spacer Row 9
    ws.row_dimensions[9].height = 10

    # =========================================================================
    # PANEL 2: OPERATIONAL & TRIP DYNAMICS (Rows 10-14)
    # =========================================================================
    ws.merge_cells('A10:H10')
    p2_head = ws['A10']
    p2_head.value = 'PANEL 2: OPERATIONAL & TRIP DYNAMICS (Site logistics, environmental mode & equipment cycles)'
    p2_head.font = styles['font_white_bold']
    p2_head.fill = styles['fill_header']
    p2_head.alignment = styles['align_left']
    ws.row_dimensions[10].height = 24

    # Row 11: Route Parameters & Standard Payload
    ws['A11'] = 'Lead Distance (L):'
    ws['A11'].font = styles['font_bold']
    ws['A11'].alignment = styles['align_right']
    ws['B11'] = config.get('default_lead', 26.0)
    ws['B11'].fill = styles['fill_input']
    ws['B11'].font = styles['font_bold']
    ws['B11'].alignment = styles['align_center']
    ws['B11'].number_format = '0.00 "km"'

    ws['C11'] = 'Average Speed (S):'
    ws['C11'].font = styles['font_bold']
    ws['C11'].alignment = styles['align_right']
    ws['D11'] = config.get('default_speed', 29.0)
    ws['D11'].fill = styles['fill_input']
    ws['D11'].font = styles['font_bold']
    ws['D11'].alignment = styles['align_center']
    ws['D11'].number_format = '0.00 "km/h"'

    ws['E11'] = 'Turnaround Time (T):'
    ws['E11'].font = styles['font_bold']
    ws['E11'].alignment = styles['align_right']
    ws['F11'] = config.get('default_turnaround', 1.0)
    ws['F11'].fill = styles['fill_input']
    ws['F11'].font = styles['font_bold']
    ws['F11'].alignment = styles['align_center']
    ws['F11'].number_format = '0.00 "hrs"'

    ws['G11'] = 'Standard Payload (C):'
    ws['G11'].font = styles['font_bold']
    ws['G11'].alignment = styles['align_right']
    ws['H11'] = '=IFERROR(INDEX($C$88:$C$114, MATCH(D6, $B$88:$B$114, 0)), 10.98)'
    ws['H11'].fill = styles['fill_calc']
    ws['H11'].font = styles['font_bold']
    ws['H11'].alignment = styles['align_center']
    ws['H11'].number_format = '0.00'

    for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']:
        ws[f'{col}11'].border = styles['border_thin']
    ws.row_dimensions[11].height = 22

    # Row 12: Operational Environment & Distance Basis
    ws['A12'] = 'Operational Mode:'
    ws['A12'].font = styles['font_bold']
    ws['A12'].alignment = styles['align_right']
    ws['B12'] = 'STANDARD (DAYTIME)'
    ws['B12'].fill = styles['fill_input']
    ws['B12'].font = styles['font_bold']
    ws['B12'].alignment = styles['align_center']
    dv_env_mode.add(ws['B12'])

    ws['C12'] = 'Fixed Trips Override:'
    ws['C12'].font = styles['font_bold']
    ws['C12'].alignment = styles['align_right']
    ws['D12'] = 3.00
    ws['D12'].fill = styles['fill_input']
    ws['D12'].font = styles['font_bold']
    ws['D12'].alignment = styles['align_center']
    ws['D12'].number_format = '0.00'

    ws['E12'] = 'Distance / Fuel Basis:'
    ws['E12'].font = styles['font_bold']
    ws['E12'].alignment = styles['align_right']
    ws['F12'] = 'DIRECT ROUTE (2NL + 6)'
    ws['F12'].fill = styles['fill_input']
    ws['F12'].font = styles['font_bold']
    ws['F12'].alignment = styles['align_center']
    dv_dist_basis.add(ws['F12'])

    ws['G12'] = 'Payload Override (C):'
    ws['G12'].font = styles['font_bold']
    ws['G12'].alignment = styles['align_right']
    ws['H12'] = None
    ws['H12'].fill = styles['fill_input']
    ws['H12'].font = styles['font_bold']
    ws['H12'].alignment = styles['align_center']
    ws['H12'].number_format = '0.00'

    for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']:
        ws[f'{col}12'].border = styles['border_thin']
    ws.row_dimensions[12].height = 22

    # Row 13: Vehicle Dynamics & Daily Outputs
    ws['A13'] = 'Billing / Basis Unit:'
    ws['A13'].font = styles['font_bold']
    ws['A13'].alignment = styles['align_right']
    ws['B13'] = '=IFERROR(INDEX($E$88:$E$114, MATCH(D6, $B$88:$B$114, 0)), "cum")'
    ws['B13'].fill = styles['fill_calc']
    ws['B13'].font = styles['font_bold']
    ws['B13'].alignment = styles['align_center']

    ws['C13'] = 'Daily Operational Trips (N):'
    ws['C13'].font = styles['font_bold']
    ws['C13'].alignment = styles['align_right']
    ws['D13'] = '=IF(B12="URBAN RESTRICTED HOURS", D12, ROUND(8 / ((2 * B11 / D11) + F11), 2))'
    ws['D13'].fill = styles['fill_subtotal']
    ws['D13'].font = styles['font_bold']
    ws['D13'].alignment = styles['align_center']
    ws['D13'].number_format = '0.00 "trips"'

    ws['E13'] = 'Daily Distance Travelled:'
    ws['E13'].font = styles['font_bold']
    ws['E13'].alignment = styles['align_right']
    ws['F13'] = '=IF(F12="CPWD PRO-RATA (Item 1.1.18)", ROUND(88.00 * (D13 / 4.10), 2), ROUND((2 * D13 * B11) + 6.0, 2))'
    ws['F13'].fill = styles['fill_subtotal']
    ws['F13'].font = styles['font_bold']
    ws['F13'].alignment = styles['align_center']
    ws['F13'].number_format = '0.00 "km"'

    ws['G13'] = 'Total Daily Hauled Output:'
    ws['G13'].font = styles['font_bold']
    ws['G13'].alignment = styles['align_right']
    ws['H13'] = '=ROUND(D13 * IF(H12>0, H12, H11), 2)'
    ws['H13'].fill = styles['fill_result']
    ws['H13'].font = styles['font_bold']
    ws['H13'].alignment = styles['align_center']
    ws['H13'].number_format = '0.00'

    for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']:
        ws[f'{col}13'].border = styles['border_thin']
    ws.row_dimensions[13].height = 22

    # Row 14: Fuel Consumption & Real-Time In-Sheet Audit Bar
    ws['A14'] = 'AUDIT STATUS:'
    ws['A14'].font = styles['font_white_bold']
    ws['A14'].fill = styles['fill_header']
    ws['A14'].alignment = styles['align_center']

    ws['B14'] = '=IF(AND(D14="OK", F14>0, H14>0), "[PASS] ALL CHECKS OK", "[ALERT] CHECKS FAILED")'
    ws['B14'].font = styles['font_result']
    ws['B14'].fill = styles['fill_result']
    ws['B14'].alignment = styles['align_center']

    ws['C14'] = 'Daily Trips Check:'
    ws['C14'].font = styles['font_note']
    ws['C14'].alignment = styles['align_right']
    ws['D14'] = '=IF(D13>0, "OK", "ERR: N<=0")'
    ws['D14'].font = styles['font_bold']
    ws['D14'].alignment = styles['align_center']

    ws['E14'] = 'Evaluated Diesel (Litres):'
    ws['E14'].font = styles['font_bold']
    ws['E14'].alignment = styles['align_right']
    ws['F14'] = '=ROUND(F13 / 5.0, 2)'
    ws['F14'].fill = styles['fill_subtotal']
    ws['F14'].font = styles['font_bold']
    ws['F14'].alignment = styles['align_center']
    ws['F14'].number_format = '0.00 "L"'

    ws['G14'] = 'Evaluated Mobil Oil (Litres):'
    ws['G14'].font = styles['font_bold']
    ws['G14'].alignment = styles['align_right']
    ws['H14'] = '=ROUND(F13 / 140.0, 3)'
    ws['H14'].fill = styles['fill_subtotal']
    ws['H14'].font = styles['font_bold']
    ws['H14'].alignment = styles['align_center']
    ws['H14'].number_format = '0.000 "L"'

    for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']:
        ws[f'{col}14'].border = styles['border_thin']
    ws.row_dimensions[14].height = 22

    # Spacer Row 15
    ws.row_dimensions[15].height = 10

    # =========================================================================
    # SECTION 1: MECHANICAL TRANSPORT RESOURCE SCHEDULE (Rows 16-27)
    # =========================================================================
    ws.merge_cells('A16:H16')
    c_sec1 = ws['A16']
    c_sec1.value = '1. MECHANICAL TRANSPORT TRIP BUILDER (Truck Hire + Labour + Fuel Linked to CPWD Mileage Formulas)'
    c_sec1.font = styles['font_white_bold']
    c_sec1.fill = styles['fill_header']
    c_sec1.alignment = styles['align_left']
    ws.row_dimensions[16].height = 24

    tbl_headers = ['Line', 'Code / Source', 'Resource / Fuel Description', 'Unit', 'Quantity / Coeff', 'Basic Rate (Rs)', 'Amount (Rs)', 'Source Reference / CPWD Notes']
    for c_idx, h in enumerate(tbl_headers, 1):
        cell = ws.cell(row=17, column=c_idx, value=h)
        cell.font = styles['font_header']
        cell.fill = styles['fill_header']
        cell.alignment = styles['align_center']
        cell.border = styles['border_header']
    ws.row_dimensions[17].height = 24

    # Dynamic Beldar Formula driving labour gang based on Handling Scope in F6
    beldar_scope_formula = '=IF(ISNUMBER(SEARCH("transporting only", F6)), 0, IF(ISNUMBER(SEARCH("machine loaded", F6)), 3, IF(ISNUMBER(SEARCH("excluding loading", F6)), 3, IF(ISNUMBER(SEARCH("railway siding", F6)), 3.75, IF(ISNUMBER(SEARCH("excluding stacking", F6)), 5, 6)))))'

    mech_resources = [
        {'code': '0084', 'qty_formula': 1.0, 'note': 'CPWD Note 5: Hire charges of Diesel Truck (9-tonne) excluding diesel & mobil oil (Code 0084)'},
        {'code': '0114', 'qty_formula': beldar_scope_formula, 'note': 'CPWD Note 5: Labour gang scaled by Panel 1 Scope: 6 (Turnkey), 5 (No Stack), 3 (Machine Loaded), 0 (Haulage Only)'},
        {'code': '1235', 'qty_formula': '=F14', 'note': 'CPWD Note 3: Diesel consumed = Distance travelled / 5.0 km/L (linked to cell F14)'},
        {'code': '5001', 'qty_formula': '=H14', 'note': 'CPWD Note 4: Mobil oil consumed = Distance travelled / 140.0 km/L (linked to cell H14)'}
    ]

    for idx in range(8):
        r = 18 + idx
        res = mech_resources[idx] if idx < len(mech_resources) else None

        ws.cell(row=r, column=1, value=idx+1).alignment = styles['align_center']

        c_code = ws.cell(row=r, column=2, value=res['code'] if res else '')
        c_code.alignment = styles['align_center']
        c_code.font = styles['font_bold']
        c_code.fill = styles['fill_input']

        c_desc = ws.cell(row=r, column=3, value=f'=IF(B{r}="","", IFERROR(INDEX(Rates_Master!$C:$C, MATCH(B{r}, Rates_Master!$A:$A, 0)), "Custom Resource Line"))')
        c_desc.alignment = styles['align_left']
        c_desc.fill = styles['fill_lookup']

        c_unit = ws.cell(row=r, column=4, value=f'=IF(B{r}="","", IFERROR(INDEX(Rates_Master!$D:$D, MATCH(B{r}, Rates_Master!$A:$A, 0)), ""))')
        c_unit.alignment = styles['align_center']
        c_unit.fill = styles['fill_lookup']

        c_qty = ws.cell(row=r, column=5)
        if res:
            c_qty.value = res['qty_formula']
        else:
            c_qty.value = None
        c_qty.alignment = styles['align_right']
        c_qty.number_format = styles['fmt_qty']
        if r in [19, 20, 21]:
            c_qty.fill = styles['fill_calc']
            c_qty.font = styles['font_bold']
        else:
            c_qty.fill = styles['fill_input']

        c_rate = ws.cell(row=r, column=6, value=f'=IF(B{r}="","", IFERROR(INDEX(Rates_Master!$E:$E, MATCH(B{r}, Rates_Master!$A:$A, 0)), 0))')
        c_rate.alignment = styles['align_right']
        c_rate.number_format = styles['fmt_currency']
        c_rate.fill = styles['fill_lookup']

        c_amt = ws.cell(row=r, column=7, value=f'=IF(OR(B{r}="", E{r}=""), 0, ROUND(E{r} * F{r}, 2))')
        c_amt.alignment = styles['align_right']
        c_amt.number_format = styles['fmt_currency']
        c_amt.fill = styles['fill_calc']

        c_rem = ws.cell(row=r, column=8, value=res['note'] if res else '')
        c_rem.alignment = styles['align_left']
        c_rem.font = styles['font_note']

        for col in range(1, 9):
            ws.cell(row=r, column=col).border = styles['border_thin']
        ws.row_dimensions[r].height = 20

    # Row 26: Direct Daily Operating Cost (W)
    ws.merge_cells('A26:E26')
    ws['A26'] = 'Direct Daily Operating Cost (W) (Rs):'
    ws['A26'].font = styles['font_bold']
    ws['A26'].alignment = styles['align_right']
    ws['A26'].fill = styles['fill_subtotal']

    c_dirsum = ws['G26']
    c_dirsum.value = '=SUM(G18:G25)'
    c_dirsum.font = styles['font_bold']
    c_dirsum.alignment = styles['align_right']
    c_dirsum.number_format = styles['fmt_currency']
    c_dirsum.fill = styles['fill_subtotal']
    for c in ['A26', 'B26', 'C26', 'D26', 'E26', 'F26', 'G26', 'H26']:
        ws[c].border = styles['border_double_bottom']
    ws.row_dimensions[26].height = 22

    # Row 27: Operating Cost per Single Round Trip
    ws.merge_cells('A27:E27')
    ws['A27'] = 'Operating Cost per Single Round Trip (Rs):'
    ws['A27'].font = styles['font_bold']
    ws['A27'].alignment = styles['align_right']
    ws['A27'].fill = styles['fill_subtotal']

    c_tripsum = ws['G27']
    c_tripsum.value = '=ROUND(G26 / D13, 2)'
    c_tripsum.font = styles['font_bold']
    c_tripsum.alignment = styles['align_right']
    c_tripsum.number_format = styles['fmt_currency']
    c_tripsum.fill = styles['fill_result']

    ws['H27'] = 'Evaluated operating cost per single round trip (= W / N)'
    ws['H27'].font = styles['font_note']
    ws['H27'].alignment = styles['align_left']

    for c in ['A27', 'B27', 'C27', 'D27', 'E27', 'F27', 'G27', 'H27']:
        ws[c].border = styles['border_thin']
    ws.row_dimensions[27].height = 22

    # Spacer Row 28
    ws.row_dimensions[28].height = 10

    # =========================================================================
    # SECTION 2: STATUTORY MARKUPS & UNIT RATE DERIVATION (Rows 29-42)
    # =========================================================================
    ws.merge_cells('A29:H29')
    c_sec2 = ws['A29']
    c_sec2.value = '2. STATUTORY MARKUPS & UNIT RATE DERIVATION (CPWD CARRIAGE CONVENTION: 15% CPOH ONLY)'
    c_sec2.font = styles['font_white_bold']
    c_sec2.fill = styles['fill_header']
    c_sec2.alignment = styles['align_left']
    ws.row_dimensions[29].height = 24

    stat_headers = ['Item', 'Description / Stage', 'Apply? (YES/NO)', 'Basis Applied', 'Base Amount (Rs)', 'Factor / %', 'Amount (Rs)', 'CPWD Statutory Rule & Guidance Note']
    for c_idx, h in enumerate(stat_headers, 1):
        cell = ws.cell(row=30, column=c_idx, value=h)
        cell.font = styles['font_header']
        cell.fill = styles['fill_header']
        cell.alignment = styles['align_center']
        cell.border = styles['border_header']
    ws.row_dimensions[30].height = 24

    # Row 31: Base Direct Cost (W)
    ws['A31'] = 'W'
    ws['B31'] = 'Direct Daily Operating Cost (Truck + Labour + Fuel)'
    ws['C31'] = '-'
    ws['D31'] = 'Direct Sum'
    ws['E31'] = '-'
    ws['F31'] = '-'
    ws['G31'] = '=G26'
    ws['H31'] = 'Total direct operating shift cost before statutory overheads'

    # Row 32: Water Charges (NO by default per CPWD)
    ws['A32'] = 'X1'
    ws['B32'] = 'Add Water Charges (1% on W)'
    ws['C32'] = 'NO'
    ws['D32'] = 'On "W"'
    ws['E32'] = '=G31'
    ws['F32'] = '=Factor_Water'
    ws['G32'] = '=IF(C32="YES", ROUND(E32 * F32, 2), 0)'
    ws['H32'] = 'CPWD RULE: Omitted (NO) in Carriage items because no water is consumed in haulage.'

    # Row 33: Subtotal (X)
    ws['A33'] = 'X'
    ws['B33'] = 'Subtotal "X" (W + Water Charges)'
    ws['C33'] = '-'
    ws['D33'] = 'W + Water'
    ws['E33'] = '-'
    ws['F33'] = '-'
    ws['G33'] = '=G31 + G32'
    ws['H33'] = 'Base for GST (if applicable)'

    # Row 34: GST (NO by default in DAR base analysis)
    ws['A34'] = 'Y1'
    ws['B34'] = 'Add GST (14.05% Works Contract Tax on X)'
    ws['C34'] = 'NO'
    ws['D34'] = 'On "X"'
    ws['E34'] = '=G33'
    ws['F34'] = '=Factor_GST'
    ws['G34'] = '=IF(C34="YES", ROUND(E34 * F34, 2), 0)'
    ws['H34'] = 'CPWD RULE: Base DAR 2019 carriage rates exclude GST. Keep NO for official DSR comparison.'

    # Row 35: Subtotal (Y)
    ws['A35'] = 'Y'
    ws['B35'] = 'Subtotal "Y" (X + GST)'
    ws['C35'] = '-'
    ws['D35'] = 'X + GST'
    ws['E35'] = '-'
    ws['F35'] = '-'
    ws['G35'] = '=G33 + G34'
    ws['H35'] = 'Base for Contractor Profit & Overheads (CPOH)'

    # Row 36: Contractor Profit & Overheads (15% on Y)
    ws['A36'] = 'Z1'
    ws['B36'] = 'Add Contractor Profit & Overheads (15% on Y)'
    ws['C36'] = 'YES'
    ws['D36'] = 'On "Y"'
    ws['E36'] = '=G35'
    ws['F36'] = '=Factor_CPOH'
    ws['G36'] = '=IF(C36="YES", ROUND(E36 * F36, 2), 0)'
    ws['H36'] = 'CPWD RULE: Standard 15% CPOH is universally applied to all carriage items.'

    # Row 37: BOCW Cess (NO by default in base DAR)
    ws['A37'] = 'Z2'
    ws['B37'] = 'Add BOCW Welfare Cess (1% on Y + CPOH)'
    ws['C37'] = 'NO'
    ws['D37'] = 'On "Y + CPOH"'
    ws['E37'] = '=G35 + G36'
    ws['F37'] = '=Factor_Cess'
    ws['G37'] = '=IF(C37="YES", ROUND(E37 * F37, 2), 0)'
    ws['H37'] = 'CPWD RULE: Cess is excluded in base rates; toggle YES only if project requires local cess.'

    # Row 38: Grand Total Daily Operating Cost with Overheads (Z)
    ws['A38'] = 'Z'
    ws['B38'] = 'Grand Total Daily Operating Cost with Overheads'
    ws['C38'] = '-'
    ws['D38'] = 'Y + CPOH + Cess'
    ws['E38'] = '-'
    ws['F38'] = '-'
    ws['G38'] = '=G35 + G36 + G37'
    ws['H38'] = 'Total operating cost for 1 shift of 8 hours including 15% contractor profit & overheads'

    # Row 39: Cost per Single Round Trip with Overheads
    ws['A39'] = '-'
    ws['B39'] = 'Cost per Single Round Trip with Overheads'
    ws['C39'] = '-'
    ws['D39'] = 'Z / N Trips'
    ws['E39'] = '=G38'
    ws['F39'] = '=D13'
    ws['G39'] = '=ROUND(G38 / D13, 2)'
    ws['H39'] = 'Evaluated rate per single trip with overheads (= Total Cost Z / N)'

    # Row 40: Analyzed Unit Rate per Billing Unit
    ws['A40'] = '-'
    ws['B40'] = 'Analyzed Unit Rate per Billing Unit'
    ws['C40'] = '-'
    ws['D40'] = 'Z / Total Output + Fee'
    ws['E40'] = '=G38'
    ws['F40'] = '=H13'
    ws['G40'] = '=ROUND((G38 / H13) + B7, 2)'
    ws['H40'] = 'Derived cost per individual billing unit including optional municipal tipping/gate fee'

    # Row 41: Rate per Schedule Output Unit (100 m / 1000 Nos)
    ws['A41'] = '-'
    ws['B41'] = 'Rate per Schedule Output Unit (100 m / 1000 Nos)'
    ws['C41'] = '-'
    ws['D41'] = 'Basis Unit'
    ws['E41'] = '=B13'
    ws['F41'] = '-'
    ws['G41'] = '=IF(OR(B13="100 m", B13="1000 Nos"), ROUND(G40 * 100, 2), G40)'
    ws['H41'] = 'Schedule rate formatted to CPWD DAR basis (e.g. per 100m for pipes, per 1000 for bricks)'

    # Row 42: OFFICIAL CPWD SAY RATE CALLOUT
    ws['A42'] = 'SAY'
    ws['B42'] = 'CPWD OFFICIAL "SAY" RATE FOR CARRIAGE (Linked to Contract Estimate):'
    ws['C42'] = '-'
    ws['D42'] = 'Rounded per CPWD'
    ws['E42'] = '-'
    ws['F42'] = '-'
    ws['G42'] = '=IF(OR(B13="100 m", B13="1000 Nos"), ROUND(G41, 0), ROUND(G40, 2))'
    ws['H42'] = 'Final analytical rate ready for integration into BOQ and project cost estimates'

    for r in range(31, 43):
        ws.cell(row=r, column=1).alignment = styles['align_center']
        ws.cell(row=r, column=1).font = styles['font_bold']

        ws.cell(row=r, column=2).alignment = styles['align_left']
        ws.cell(row=r, column=3).alignment = styles['align_center']
        ws.cell(row=r, column=4).alignment = styles['align_center']

        c_base = ws.cell(row=r, column=5)
        c_base.alignment = styles['align_right']
        if r in [32, 34, 36, 37]:
            c_base.number_format = styles['fmt_currency']

        c_fac = ws.cell(row=r, column=6)
        c_fac.alignment = styles['align_right']
        if r in [32, 34, 36, 37]:
            c_fac.number_format = styles['fmt_percent']

        c_val = ws.cell(row=r, column=7)
        c_val.alignment = styles['align_right']
        c_val.number_format = styles['fmt_currency']
        c_val.font = styles['font_bold']

        c_note = ws.cell(row=r, column=8)
        c_note.alignment = styles['align_wrap']
        c_note.font = styles['font_note']

        for c in range(1, 9):
            ws.cell(row=r, column=c).border = styles['border_thin']

        if r in [31, 33, 35, 37]:
            for c in range(1, 8):
                ws.cell(row=r, column=c).fill = styles['fill_subtotal']
        elif r in [32, 34, 36, 37]:
            ws.cell(row=r, column=3).fill = styles['fill_input']
            ws.cell(row=r, column=3).font = styles['font_bold']
            dv_yesno.add(ws.cell(row=r, column=3))
        elif r == 38:
            for c in range(1, 8):
                ws.cell(row=r, column=c).fill = styles['fill_result']
        elif r in [39, 40, 41]:
            for c in range(1, 8):
                ws.cell(row=r, column=c).fill = styles['fill_subtotal']
        elif r == 42:
            ws.row_dimensions[42].height = 30
            for c in range(1, 8):
                ws.cell(row=r, column=c).fill = styles['fill_say']
            ws['B42'].font = styles['font_say']
            ws['G42'].font = styles['font_say']

        ws.row_dimensions[r].height = 24 if r != 42 else 30

    # Spacer Row 43
    ws.row_dimensions[43].height = 10

    # =========================================================================
    # SECTION 3: HEADING 1.2 MANUAL LABOUR CARRIAGE CALCULATOR (< 0.50 KM) (Rows 44-49)
    # =========================================================================
    ws.merge_cells('A44:H44')
    c_sec3 = ws['A44']
    c_sec3.value = '3. HEADING 1.2: MANUAL LABOUR CARRIAGE CALCULATOR (For Lead Less Than 0.50 km / 50 m to 500 m)'
    c_sec3.font = styles['font_white_bold']
    c_sec3.fill = styles['fill_header']
    c_sec3.alignment = styles['align_left']
    ws.row_dimensions[44].height = 24

    ws.merge_cells('A45:H45')
    c_mbanner = ws['A45']
    c_mbanner.value = 'CPWD DAR 2019 Item 1.2 Standards: Category A (Bulk Materials) = 7.67 Beldars 1st 50m (+1.67 Coolies/addl 50m). Category B (Heavy/Pipes/Steel) = 9.20 Beldars 1st 50m (+1.35 Beldars/addl 50m). All rates include 15% CPOH.'
    c_mbanner.font = styles['font_note']
    c_mbanner.fill = styles['fill_note']
    c_mbanner.alignment = styles['align_left']
    ws.row_dimensions[45].height = 22

    # Row 46: Inputs
    ws['A46'] = 'Manual Item Code:'
    ws['A46'].font = styles['font_bold']
    ws['B46'] = '1.2.CUSTOM'
    ws['B46'].fill = styles['fill_input']
    ws['B46'].font = styles['font_bold']
    ws['B46'].alignment = styles['align_center']

    ws['C46'] = 'Material Category:'
    ws['C46'].font = styles['font_bold']
    ws['D46'] = 'Category B (Heavy / Pipes / Steel)'
    ws['D46'].fill = styles['fill_input']
    ws['D46'].font = styles['font_bold']
    ws['D46'].alignment = styles['align_center']
    dv_manual_cat.add(ws['D46'])

    ws['E46'] = 'Lead Distance (Metres):'
    ws['E46'].font = styles['font_bold']
    ws['F46'] = 100
    ws['F46'].fill = styles['fill_input']
    ws['F46'].font = styles['font_bold']
    ws['F46'].alignment = styles['align_center']
    dv_manual_lead.add(ws['F46'])

    ws['G46'] = 'Addl 50m Steps (M):'
    ws['G46'].font = styles['font_bold']
    ws['H46'] = '=MAX(0, (F46 - 50) / 50)'
    ws['H46'].fill = styles['fill_subtotal']
    ws['H46'].font = styles['font_bold']
    ws['H46'].alignment = styles['align_center']
    ws['H46'].number_format = '0'

    for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']:
        ws[f'{col}46'].border = styles['border_thin']
    ws.row_dimensions[46].height = 22

    # Row 47: Manual Specification
    ws['A47'] = 'Manual Specification:'
    ws['A47'].font = styles['font_bold']
    ws['A47'].border = styles['border_thin']
    ws.merge_cells('B47:H47')
    ws['B47'] = 'Carriage by manual labour including loading, unloading and stacking for lead upto 100 metres complete as per directions of Engineer-in-charge.'
    ws['B47'].font = styles['font_regular']
    ws['B47'].fill = styles['fill_input']
    ws['B47'].alignment = styles['align_wrap']
    ws['B47'].border = styles['border_thin']
    ws.row_dimensions[47].height = 26

    # Row 48: Capacity & Norms
    ws['A48'] = '8-Hour Output Capacity:'
    ws['A48'].font = styles['font_bold']
    ws['B48'] = 1702.00
    ws['B48'].fill = styles['fill_input']
    ws['B48'].font = styles['font_bold']
    ws['B48'].alignment = styles['align_center']
    ws['B48'].number_format = '0.00'

    ws['C48'] = 'Output / Billing Unit:'
    ws['C48'].font = styles['font_bold']
    ws['D48'] = '100 m'
    ws['D48'].fill = styles['fill_input']
    ws['D48'].font = styles['font_bold']
    ws['D48'].alignment = styles['align_center']
    dv_unit.add(ws['D48'])

    ws['E48'] = 'Base Labour 1st 50m (Rs):'
    ws['E48'].font = styles['font_bold']
    ws['F48'] = '=IF(ISNUMBER(SEARCH("Category A", D46)), 4279.86, 5133.60)'
    ws['F48'].fill = styles['fill_subtotal']
    ws['F48'].font = styles['font_bold']
    ws['F48'].alignment = styles['align_right']
    ws['F48'].number_format = styles['fmt_currency']

    ws['G48'] = 'Addl Labour / 50m (Rs):'
    ws['G48'].font = styles['font_bold']
    ws['H48'] = '=IF(ISNUMBER(SEARCH("Category A", D46)), 931.86, 753.30)'
    ws['H48'].fill = styles['fill_subtotal']
    ws['H48'].font = styles['font_bold']
    ws['H48'].alignment = styles['align_right']
    ws['H48'].number_format = styles['fmt_currency']

    for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']:
        ws[f'{col}48'].border = styles['border_thin']
    ws.row_dimensions[48].height = 22

    # Row 49: Cost Derivation
    ws['A49'] = 'Total Labour Cost (Rs):'
    ws['A49'].font = styles['font_bold']
    ws['B49'] = '=F48 + (H46 * H48)'
    ws['B49'].fill = styles['fill_subtotal']
    ws['B49'].font = styles['font_bold']
    ws['B49'].alignment = styles['align_right']
    ws['B49'].number_format = styles['fmt_currency']

    ws['C49'] = 'Add 15% CPOH (Rs):'
    ws['C49'].font = styles['font_bold']
    ws['D49'] = '=ROUND(B49 * 0.15, 2)'
    ws['D49'].fill = styles['fill_subtotal']
    ws['D49'].font = styles['font_bold']
    ws['D49'].alignment = styles['align_right']
    ws['D49'].number_format = styles['fmt_currency']

    ws['E49'] = 'Total 8-Hr Cost with CPOH:'
    ws['E49'].font = styles['font_bold']
    ws['F49'] = '=B49 + D49'
    ws['F49'].fill = styles['fill_result']
    ws['F49'].font = styles['font_bold']
    ws['F49'].alignment = styles['align_right']
    ws['F49'].number_format = styles['fmt_currency']

    ws['G49'] = 'Analyzed Unit Rate (Rs):'
    ws['G49'].font = styles['font_bold']
    ws['H49'] = '=IF(OR(D48="100 m", D48="1000 Nos"), ROUND((F49 / B48) * 100, 2), ROUND(F49 / B48, 2))'
    ws['H49'].fill = styles['fill_say']
    ws['H49'].font = styles['font_say']
    ws['H49'].alignment = styles['align_right']
    ws['H49'].number_format = styles['fmt_currency']

    for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']:
        ws[f'{col}49'].border = styles['border_thin']
    ws.row_dimensions[49].height = 26

    # Spacer Row 50
    ws.row_dimensions[50].height = 12

    # =========================================================================
    # SECTION 4: GROUND-TRUTH ENGINEERING REFERENCE BENCHMARKS (Rows 51+)
    # =========================================================================
    ws.merge_cells('A51:H51')
    c_sec4_main = ws['A51']
    c_sec4_main.value = '4. GROUND-TRUTH ENGINEERING REFERENCE BENCHMARKS (CPWD DAR 2019 SUB-HEAD 01)'
    c_sec4_main.font = styles['font_white_bold']
    c_sec4_main.fill = styles['fill_header']
    c_sec4_main.alignment = styles['align_left']
    ws.row_dimensions[51].height = 24

    # --- 4A: CPWD DATA SHEET NO. 1 (Rows 52-84) ---
    ws.merge_cells('A52:H52')
    c_sec4a = ws['A52']
    c_sec4a.value = '4A. CPWD DAR 2019 DATA SHEET NO. 1 GROUND-TRUTH BENCHMARK (1 km to 30 km Mechanical Transport Reference)'
    c_sec4a.font = styles['font_white_bold']
    c_sec4a.fill = styles['fill_header']
    c_sec4a.alignment = styles['align_left']
    ws.row_dimensions[52].height = 22

    bench_headers = ['Lead (L) (km)', 'Avg Speed (S) (km/h)', 'Trips (N) / Day', 'Km Done / Day', 'Diesel Qty (Litres)', 'Mobil Oil (Litres)', 'Total Shift Cost (Rs)', 'Cost per Trip (Rs)']
    for c_idx, h in enumerate(bench_headers, 1):
        cell = ws.cell(row=53, column=c_idx, value=h)
        cell.font = styles['font_header']
        cell.fill = styles['fill_header']
        cell.alignment = styles['align_center']
        cell.border = styles['border_header']
    ws.row_dimensions[53].height = 24

    datasheet1_records = [
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
        (30.0, 31.0, 2.73, 169.80, 33.96, 1.213, 7726.16, 2830.10)
    ]

    for idx, rec in enumerate(datasheet1_records):
        r = 54 + idx
        ws.cell(row=r, column=1, value=rec[0]).alignment = styles['align_center']
        ws.cell(row=r, column=1).number_format = '0.0 "km"'

        ws.cell(row=r, column=2, value=rec[1]).alignment = styles['align_center']
        ws.cell(row=r, column=2).number_format = '0.0'

        ws.cell(row=r, column=3, value=rec[2]).alignment = styles['align_center']
        ws.cell(row=r, column=3).number_format = '0.00'

        ws.cell(row=r, column=4, value=rec[3]).alignment = styles['align_right']
        ws.cell(row=r, column=4).number_format = '0.00'

        ws.cell(row=r, column=5, value=rec[4]).alignment = styles['align_right']
        ws.cell(row=r, column=5).number_format = '0.00'

        ws.cell(row=r, column=6, value=rec[5]).alignment = styles['align_right']
        ws.cell(row=r, column=6).number_format = '0.000'

        ws.cell(row=r, column=7, value=rec[6]).alignment = styles['align_right']
        ws.cell(row=r, column=7).number_format = styles['fmt_currency']

        c_bcost = ws.cell(row=r, column=8, value=rec[7])
        c_bcost.alignment = styles['align_right']
        c_bcost.number_format = styles['fmt_currency']
        c_bcost.font = styles['font_bold']

        row_fill = styles['fill_subtotal'] if r % 2 == 0 else styles['fill_calc']
        for c in range(1, 9):
            cell = ws.cell(row=r, column=c)
            cell.fill = row_fill
            cell.border = styles['border_thin']
            if c != 8:
                cell.font = styles['font_regular']
        ws.row_dimensions[r].height = 20

    ws.merge_cells('A84:H84')
    c_bnote = ws['A84']
    c_bnote.value = 'NOTES: N=8/((2L/S)+1); Km Done=(2NL)+6.0; Diesel=Km/5.0; Mobil Oil=Km/140.0; Cost=Diesel+Oil+Rs 3348 (6 Beldars)+Rs 1500 (Truck); Cost/Trip=Total Cost/N.'
    c_bnote.font = styles['font_note']
    c_bnote.fill = styles['fill_note']
    c_bnote.alignment = styles['align_left']
    c_bnote.border = styles['border_thin']
    ws.row_dimensions[84].height = 20

    # Spacer Row 85
    ws.row_dimensions[85].height = 10

    # --- 4B: CPWD TABLE 1.1 MATERIAL PAYLOAD CAPACITIES MATRIX (Rows 86-114) ---
    ws.merge_cells('A86:H86')
    c_sec4b = ws['A86']
    c_sec4b.value = '4B. CPWD DAR STANDARD MATERIAL PAYLOAD CAPACITIES MATRIX (Table 1.1 Technical Reference & Dropdown Source)'
    c_sec4b.font = styles['font_white_bold']
    c_sec4b.fill = styles['fill_header']
    c_sec4b.alignment = styles['align_left']
    ws.row_dimensions[86].height = 22

    cap_headers = ['DAR Code', 'Material / Trade Specification', 'Truck Payload / Trip', 'Net Payable Qty', 'Schedule Unit', 'Looseness Allowance / Technical Norms', 'Base Rate 1 km (Rs)', 'Addl Rate >20 km (Rs)']
    for c_idx, h in enumerate(cap_headers, 1):
        cell = ws.cell(row=87, column=c_idx, value=h)
        cell.font = styles['font_header']
        cell.fill = styles['fill_header']
        cell.alignment = styles['align_center']
        cell.border = styles['border_header']
    ws.row_dimensions[87].height = 24

    capacities_table = [
        ('1.1.1', 'Lime, moorum, building rubbish, malba', 8.00, 8.00, 'cum', 'Nil looseness deduction; full 8.0 cum payable', 104.94, 8.35),
        ('1.1.2', 'Earth (excavated soil / good earth)', 8.00, 6.40, 'cum', '20% deduction for looseness (Net = 6.40 cum)', 131.17, 10.44),
        ('1.1.3', 'Manure or sludge', 8.00, 7.36, 'cum', '8% deduction for looseness (Net = 7.36 cum)', 114.06, 9.08),
        ('1.1.4', 'Excavated rock', 8.00, 4.00, 'cum', '50% deduction for voids/looseness (Net = 4.00 cum)', 209.88, 16.70),
        ('1.1.5', 'Sand, stone aggregate below 40 mm', 8.00, 8.00, 'cum', 'Nil looseness deduction; standard density 8.0 cum', 104.94, 8.35),
        ('1.1.6', 'Stone aggregate 40 mm nominal size & above', 8.00, 7.36, 'cum', '8% deduction for voids in coarse aggregate', 114.06, 9.08),
        ('1.1.7', 'Soling stone & masonry stone', 8.00, 6.80, 'cum', '15% deduction for stack voids (Net = 6.80 cum)', 123.46, 9.82),
        ('1.1.8', 'Bricks (standard modular / conventional)', 3000.00, 3000.00, '1000 Nos', '3,000 Bricks per 9-tonne truck load', 279.83, 22.26),
        ('1.1.9', 'Brick tiles / Allahabad roofing tiles', 5000.00, 5000.00, '1000 Nos', '5,000 Tiles per 9-tonne truck load', 167.90, 13.36),
        ('1.1.10', 'Cement, stone blocks, Kota stone slabs', 9.00, 9.00, 'tonne', 'Rated maximum truck payload capacity = 9 tonnes', 93.28, 7.42),
        ('1.1.11', 'Steel bars, structural sections & fabric', 9.00, 9.00, 'tonne', 'Rated maximum truck payload capacity = 9 tonnes', 93.28, 7.42),
        ('1.1.12', 'Timber (scantlings / logs)', 7.00, 7.00, 'cum', 'Volume limit on timber body = 7.0 cum', 119.93, 9.54),
        ('1.1.13', 'Tar, bitumen in drums', 8.00, 8.00, 'tonne', 'Packed drum loading capacity = 8 tonnes', 104.94, 8.35),
        ('1.1.14', 'Steam coal', 7.00, 7.00, 'tonne', 'Bulk density volume restriction = 7 tonnes', 119.93, 9.54),
        ('1.1.15.1', 'S.W. pipes 100 mm dia', 600.00, 600.00, '100 m', 'Payload capacity = 600 metres (6 x 100 m)', 139.92, 11.13),
        ('1.1.15.2', 'S.W. pipes 150 mm dia', 300.00, 300.00, '100 m', 'Payload capacity = 300 metres (3 x 100 m)', 279.83, 22.26),
        ('1.1.16.1', 'R.C.C. / C.I. pipes 100 mm dia', 366.00, 366.00, '100 m', 'Payload capacity = 366 metres', 229.37, 18.25),
        ('1.1.16.3', 'R.C.C. / C.I. pipes 150 mm dia', 219.60, 219.60, '100 m', 'Payload capacity = 219.60 metres', 382.29, 30.42),
        ('1.1.16.4', 'R.C.C. / C.I. pipes 200 mm dia', 135.00, 135.00, '100 m', 'Payload capacity = 135 metres', 621.85, 49.48),
        ('1.1.16.5', 'R.C.C. / C.I. pipes 250 mm dia', 95.00, 95.00, '100 m', 'Payload capacity = 95 metres', 883.68, 70.31),
        ('1.1.16.6', 'R.C.C. / C.I. pipes 300 mm dia', 76.86, 76.86, '100 m', 'Payload capacity = 76.86 metres', 1092.25, 86.90),
        ('1.1.16.7', 'R.C.C. / C.I. pipes 350 mm dia', 54.90, 54.90, '100 m', 'Payload capacity = 54.90 metres', 1529.14, 121.66),
        ('1.1.16.8', 'R.C.C. / C.I. pipes 400 mm dia', 40.26, 40.26, '100 m', 'Payload capacity = 40.26 metres', 2085.20, 165.90),
        ('1.1.16.9', 'R.C.C. / C.I. pipes 450 & 500 mm dia', 32.94, 32.94, '100 m', 'Payload capacity = 32.94 metres', 2548.57, 202.77),
        ('1.1.16.10', 'R.C.C. / C.I. pipes 600, 700, 750 & 800 mm dia', 21.96, 21.96, '100 m', 'Payload capacity = 21.96 metres', 3822.86, 304.15),
        ('1.1.16.11', 'R.C.C. / C.I. pipes 900 mm dia', 14.64, 14.64, '100 m', 'Payload capacity = 14.64 metres', 5734.29, 456.23),
        ('1.1.17.12', 'R.C.C./C.I./Steel pipes 1000, 1100 & 1200 mm dia', 10.98, 10.98, '100 m', 'Payload capacity = 10.98 metres (heavy large bore)', 7645.72, 608.31)
    ]

    for idx, mat in enumerate(capacities_table):
        r = 88 + idx
        ws.cell(row=r, column=1, value=mat[0]).alignment = styles['align_center']
        ws.cell(row=r, column=1).font = styles['font_bold']

        ws.cell(row=r, column=2, value=mat[1]).alignment = styles['align_left']

        ws.cell(row=r, column=3, value=mat[2]).alignment = styles['align_right']
        ws.cell(row=r, column=3).number_format = '0.00'

        ws.cell(row=r, column=4, value=mat[3]).alignment = styles['align_right']
        ws.cell(row=r, column=4).number_format = '0.00'

        ws.cell(row=r, column=5, value=mat[4]).alignment = styles['align_center']

        ws.cell(row=r, column=6, value=mat[5]).alignment = styles['align_left']
        ws.cell(row=r, column=6).font = styles['font_note']

        ws.cell(row=r, column=7, value=mat[6]).alignment = styles['align_right']
        ws.cell(row=r, column=7).number_format = styles['fmt_currency']

        ws.cell(row=r, column=8, value=mat[7]).alignment = styles['align_right']
        ws.cell(row=r, column=8).number_format = styles['fmt_currency']

        row_fill = styles['fill_subtotal'] if r % 2 == 0 else styles['fill_calc']
        for c in range(1, 9):
            cell = ws.cell(row=r, column=c)
            cell.fill = row_fill
            cell.border = styles['border_thin']
            if c not in [1, 7, 8]:
                cell.font = styles['font_regular']
        ws.row_dimensions[r].height = 20

    # Named Range for Materials
    wb.defined_names.add(DefinedName('CPWD_Carriage_Materials', attr_text=f"'{config['sheet_name']}'!$B$88:$B$114"))

    # Spacer Row 115
    ws.row_dimensions[115].height = 10

    # --- 4C: CPWD TABLE 1.2 MANUAL LABOUR REFERENCE (Rows 116-125) ---
    ws.merge_cells('A116:H116')
    c_sec4c = ws['A116']
    c_sec4c.value = '4C. CPWD DAR TABLE 1.2 MANUAL LABOUR REFERENCE (Labour Norms & Rates for Lead < 0.50 km)'
    c_sec4c.font = styles['font_white_bold']
    c_sec4c.fill = styles['fill_header']
    c_sec4c.alignment = styles['align_left']
    ws.row_dimensions[116].height = 22

    manual_headers = ['Item No.', 'Material / Specification', 'Capacity / Day', 'Unit', 'Base Cost 1st 50m (Rs)', 'Cost / 1st 50m (Rs)', 'Addl Cost / 50m (Rs)', 'CPWD Labour Norms']
    for c_idx, h in enumerate(manual_headers, 1):
        cell = ws.cell(row=117, column=c_idx, value=h)
        cell.font = styles['font_header']
        cell.fill = styles['fill_header']
        cell.alignment = styles['align_center']
        cell.border = styles['border_header']
    ws.row_dimensions[117].height = 24

    manual_records = [
        ('1.2.1', 'Lime, moorum, building rubbish', 35.0, 'cum', 4921.84, 140.62, 30.62, '7.67 Beldars 1st 50m, 1.67 addl coolie/50m'),
        ('1.2.2', 'Earth (20% looseness deduction)', 28.0, 'cum', 4921.84, 175.78, 38.27, '7.67 Beldars 1st 50m, 1.67 addl coolie/50m'),
        ('1.2.8', 'Bricks (standard modular)', 15000.0, '1000 Nos', 4921.84, 328.12, 71.44, '7.67 Beldars 1st 50m, 1.67 addl coolie/50m'),
        ('1.2.11', 'Stone blocks, G.I., C.I., pipes <100mm', 46.0, 'tonne', 5903.64, 128.34, 18.83, '9.20 Beldars 1st 50m, 1.35 addl coolie/50m'),
        ('1.2.12', 'Cement in bags', 57.99, 'tonne', 5903.64, 101.80, 14.94, '9.20 Beldars 1st 50m, 1.35 addl coolie/50m'),
        ('1.2.17.1', 'R.C.C./C.I. pipes 100mm dia', 1702.0, '100 m', 5903.64, 346.86, 50.90, '9.20 Beldars 1st 50m, 1.35 addl coolie/50m')
    ]

    for idx, m_rec in enumerate(manual_records):
        r = 118 + idx
        ws.cell(row=r, column=1, value=m_rec[0]).alignment = styles['align_center']
        ws.cell(row=r, column=1).font = styles['font_bold']

        ws.cell(row=r, column=2, value=m_rec[1]).alignment = styles['align_left']

        ws.cell(row=r, column=3, value=m_rec[2]).alignment = styles['align_right']
        ws.cell(row=r, column=3).number_format = '0.00'

        ws.cell(row=r, column=4, value=m_rec[3]).alignment = styles['align_center']

        ws.cell(row=r, column=5, value=m_rec[4]).alignment = styles['align_right']
        ws.cell(row=r, column=5).number_format = styles['fmt_currency']

        ws.cell(row=r, column=6, value=m_rec[5]).alignment = styles['align_right']
        ws.cell(row=r, column=6).number_format = styles['fmt_currency']

        ws.cell(row=r, column=7, value=m_rec[6]).alignment = styles['align_right']
        ws.cell(row=r, column=7).number_format = styles['fmt_currency']

        ws.cell(row=r, column=8, value=m_rec[7]).alignment = styles['align_left']
        ws.cell(row=r, column=8).font = styles['font_note']

        row_fill = styles['fill_subtotal'] if r % 2 == 0 else styles['fill_calc']
        for c in range(1, 9):
            cell = ws.cell(row=r, column=c)
            cell.fill = row_fill
            cell.border = styles['border_thin']
        ws.row_dimensions[r].height = 20

    # Spacer Row 124
    ws.row_dimensions[124].height = 10

    # --- 4D: HANDLING SCOPE LABOUR GANG ALLOCATION MATRIX (Rows 125-134) ---
    ws.merge_cells('A125:H125')
    c_sec4d = ws['A125']
    c_sec4d.value = '4D. CPWD HANDLING SCOPE LABOUR GANG ALLOCATION MATRIX (Engineering Rationale for Cell E19)'
    c_sec4d.font = styles['font_white_bold']
    c_sec4d.fill = styles['fill_header']
    c_sec4d.alignment = styles['align_left']
    ws.row_dimensions[125].height = 22

    scope_headers = ['Scope Option (Panel 1 Cell F6)', 'Beldar Gang', 'Loading Labour', 'Unloading Labour', 'Stacking Labour', 'Daily Labour Cost (Rs)', 'Direct Shift Cost (Rs)', 'Engineering Operational Rationale']
    for c_idx, h in enumerate(scope_headers, 1):
        cell = ws.cell(row=126, column=c_idx, value=h)
        cell.font = styles['font_header']
        cell.fill = styles['fill_header']
        cell.alignment = styles['align_center']
        cell.border = styles['border_header']
    ws.row_dimensions[126].height = 24

    scope_records = [
        ('including loading, transporting, unloading and stacking', 6.00, '3.00 Beldars', '2.00 Beldars', '1.00 Beldar', 3348.00, 5939.58, 'CPWD Item 1.1.1 baseline turnkey manual handling.'),
        ('including loading, transporting, unloading to approved municipal dumping ground', 6.00, '3.00 Beldars', '2.00 Beldars', '1.00 Beldar', 3348.00, 5939.58, 'CPWD Item 1.1.18 benchmark for urban malba/rubbish disposal.'),
        ('including loading, transporting and unloading (excluding stacking)', 5.00, '3.00 Beldars', '2.00 Beldars', '0.00 Beldars', 2790.00, 5381.58, 'Bulk delivery / dumping; manual stacking omitted.'),
        ('transporting and unloading only (machine loaded / excluding loading)', 3.00, '0.00 Beldars', '2.00 Beldars', '1.00 Beldar', 1674.00, 4265.58, 'Loaded by excavator/JCB (Sub-Head 02); manual loading excluded.'),
        ('transporting only (excluding loading, unloading and stacking)', 0.00, '0.00 Beldars', '0.00 Beldars', '0.00 Beldars', 0.00, 2591.58, 'Pure haulage; machine loaded at source and tipper unloaded at site.'),
        ('including unloading and stacking at railway siding', 3.75, '0.00 Beldars', '2.50 Beldars', '1.25 Beldars', 2092.50, 4684.08, 'CPWD Item 1.3 / 1.4 standard for railway wagon godown siding.')
    ]

    for idx, s_rec in enumerate(scope_records):
        r = 127 + idx
        ws.cell(row=r, column=1, value=s_rec[0]).alignment = styles['align_left']
        ws.cell(row=r, column=2, value=s_rec[1]).alignment = styles['align_center']
        ws.cell(row=r, column=2).number_format = '0.00'
        ws.cell(row=r, column=2).font = styles['font_bold']

        ws.cell(row=r, column=3, value=s_rec[2]).alignment = styles['align_center']
        ws.cell(row=r, column=4, value=s_rec[3]).alignment = styles['align_center']
        ws.cell(row=r, column=5, value=s_rec[4]).alignment = styles['align_center']

        c_scost = ws.cell(row=r, column=6, value=s_rec[5])
        c_scost.alignment = styles['align_right']
        c_scost.number_format = styles['fmt_currency']

        c_dcost = ws.cell(row=r, column=7, value=s_rec[6])
        c_dcost.alignment = styles['align_right']
        c_dcost.number_format = styles['fmt_currency']
        c_dcost.font = styles['font_bold']

        c_srat = ws.cell(row=r, column=8, value=s_rec[7])
        c_srat.alignment = styles['align_left']
        c_srat.font = styles['font_note']

        row_fill = styles['fill_subtotal'] if r % 2 == 0 else styles['fill_calc']
        for c in range(1, 9):
            cell = ws.cell(row=r, column=c)
            cell.fill = row_fill
            cell.border = styles['border_thin']
        ws.row_dimensions[r].height = 20

    # Column Widths
    col_widths = {
        'A': 24,
        'B': 24,
        'C': 26,
        'D': 24,
        'E': 26,
        'F': 24,
        'G': 24,
        'H': 46
    }
    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width

    # Lock all cells by default, then unlock user editable cells
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            ws.cell(row=r, column=c).protection = Protection(locked=True)

    # Unlock Panel 1 inputs
    ws['B6'].protection = Protection(locked=False) # Item Code
    ws['D6'].protection = Protection(locked=False) # Material Commodity
    ws['F6'].protection = Protection(locked=False) # Handling Scope
    ws['H6'].protection = Protection(locked=False) # Lift Condition
    ws['B7'].protection = Protection(locked=False) # Municipal Gate Fee
    ws['D7'].protection = Protection(locked=False) # Manual Override

    # Unlock Panel 2 inputs
    ws['B11'].protection = Protection(locked=False) # Lead Distance (L)
    ws['D11'].protection = Protection(locked=False) # Average Speed (S)
    ws['F11'].protection = Protection(locked=False) # Turnaround Time (T)
    ws['B12'].protection = Protection(locked=False) # Operational Mode
    ws['D12'].protection = Protection(locked=False) # Fixed Trips Override
    ws['F12'].protection = Protection(locked=False) # Distance Basis
    ws['H12'].protection = Protection(locked=False) # Payload Override

    # Unlock Section 1 custom resource lines (rows 22-25)
    for r in range(22, 26):
        ws.cell(row=r, column=2).protection = Protection(locked=False) # Code
        ws.cell(row=r, column=5).protection = Protection(locked=False) # Qty

    # Unlock Section 2 statutory toggles
    ws['C32'].protection = Protection(locked=False)
    ws['C34'].protection = Protection(locked=False)
    ws['C36'].protection = Protection(locked=False)
    ws['C37'].protection = Protection(locked=False)

    # Unlock Section 3 Manual Labour inputs
    ws['B46'].protection = Protection(locked=False)
    ws['D46'].protection = Protection(locked=False)
    ws['F46'].protection = Protection(locked=False)
    ws['B47'].protection = Protection(locked=False)
    ws['B48'].protection = Protection(locked=False)
    ws['D48'].protection = Protection(locked=False)

    ws.protection.sheet = True
    ws.freeze_panes = 'A4'
