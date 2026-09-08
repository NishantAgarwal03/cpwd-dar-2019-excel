# -*- coding: utf-8 -*-
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.worksheet.datavalidation import DataValidation
from scripts.trade_builder import add_trade_header_and_legend

def build_carriage_trade(wb, config, styles):
    ws = wb.create_sheet(title=config['sheet_name'])
    ws.views.sheetView[0].showGridLines = True
    add_trade_header_and_legend(ws, config, styles)
    
    # Data Validations
    dv_yesno = DataValidation(type='list', formula1='"YES,NO"', allow_blank=False)
    ws.add_data_validation(dv_yesno)
    
    dv_mode = DataValidation(type='list', formula1='"FORMULA,FIXED TRIPS"', allow_blank=False)
    ws.add_data_validation(dv_mode)
    
    dv_unit = DataValidation(type='list', formula1='"metre,100 m,cum,tonne,1000 Nos"', allow_blank=False)
    ws.add_data_validation(dv_unit)
    
    dv_manual_cat = DataValidation(type='list', formula1='"Category A (Bulk / Earth / Bricks),Category B (Heavy / Pipes / Steel)"', allow_blank=False)
    ws.add_data_validation(dv_manual_cat)
    
    dv_manual_lead = DataValidation(type='list', formula1='"50,100,150,200,250,300,350,400,450,500"', allow_blank=False)
    ws.add_data_validation(dv_manual_lead)

    # =========================================================================
    # PART A: HEADING 1.1 — MECHANICAL TRANSPORT SIMULATOR (>= 1.0 KM)
    # =========================================================================
    # Row 5: Primary Parameters
    ws['A5'] = 'Custom Item Code:'
    ws['A5'].font = styles['font_bold']
    ws['B5'] = config.get('default_item_code', '1.1.CUSTOM')
    ws['B5'].fill = styles['fill_input']
    ws['B5'].font = styles['font_bold']
    ws['B5'].alignment = styles['align_center']
    
    ws['C5'] = 'Lead Distance (L):'
    ws['C5'].font = styles['font_bold']
    ws['D5'] = config.get('default_lead', 26.0)
    ws['D5'].fill = styles['fill_input']
    ws['D5'].font = styles['font_bold']
    ws['D5'].alignment = styles['align_center']
    ws['D5'].number_format = '0.00 "km"'
    
    ws['E5'] = 'Average Speed (S):'
    ws['E5'].font = styles['font_bold']
    ws['F5'] = config.get('default_speed', 29.0)
    ws['F5'].fill = styles['fill_input']
    ws['F5'].font = styles['font_bold']
    ws['F5'].alignment = styles['align_center']
    ws['F5'].number_format = '0.00 "km/h"'
    
    ws['G5'] = 'Turnaround Time (T):'
    ws['G5'].font = styles['font_bold']
    ws['H5'] = config.get('default_turnaround', 1.0)
    ws['H5'].fill = styles['fill_input']
    ws['H5'].font = styles['font_bold']
    ws['H5'].alignment = styles['align_center']
    ws['H5'].number_format = '0.00 "hrs"'
    
    for c in ['A5', 'B5', 'C5', 'D5', 'E5', 'F5', 'G5', 'H5']:
        ws[c].border = styles['border_thin']
    ws.row_dimensions[5].height = 22

    # Row 6: Item Nomenclature
    ws['A6'] = 'Item Nomenclature / Specification:'
    ws['A6'].font = styles['font_bold']
    ws['A6'].border = styles['border_thin']
    
    ws.merge_cells('B6:H6')
    ws['B6'] = config.get('default_item_desc', 'Transport of pipes/material by mechanical transport...')
    ws['B6'].font = styles['font_regular']
    ws['B6'].fill = styles['fill_input']
    ws['B6'].alignment = styles['align_wrap']
    ws['B6'].border = styles['border_thin']
    ws.row_dimensions[6].height = 36

    # Row 7: Payload, Unit & Core Outputs (N Trips, Km Done)
    ws['A7'] = 'Payload Capacity / Trip (C):'
    ws['A7'].font = styles['font_bold']
    ws['B7'] = config.get('default_capacity', 10.98)
    ws['B7'].fill = styles['fill_input']
    ws['B7'].font = styles['font_bold']
    ws['B7'].alignment = styles['align_center']
    ws['B7'].number_format = '0.00'
    
    ws['C7'] = 'Output / Billing Unit:'
    ws['C7'].font = styles['font_bold']
    ws['D7'] = config.get('default_basis_unit', 'metre')
    ws['D7'].fill = styles['fill_input']
    ws['D7'].font = styles['font_bold']
    ws['D7'].alignment = styles['align_center']
    dv_unit.add(ws['D7'])
    
    ws['E7'] = 'Daily Calculated Trips (N):'
    ws['E7'].font = styles['font_bold']
    ws['F7'] = '=IF(D8="FIXED TRIPS", F8, ROUND(8 / ((2 * D5 / F5) + H5), 2))'
    ws['F7'].fill = styles['fill_subtotal']
    ws['F7'].font = styles['font_bold']
    ws['F7'].alignment = styles['align_center']
    ws['F7'].number_format = '0.00 "trips"'
    
    ws['G7'] = 'Daily Distance Travelled:'
    ws['G7'].font = styles['font_bold']
    ws['H7'] = '=ROUND((2 * F7 * D5) + B8, 2)'
    ws['H7'].fill = styles['fill_subtotal']
    ws['H7'].font = styles['font_bold']
    ws['H7'].alignment = styles['align_center']
    ws['H7'].number_format = '0.00 "km"'
    
    for c in ['A7', 'B7', 'C7', 'D7', 'E7', 'F7', 'G7', 'H7']:
        ws[c].border = styles['border_thin']
    ws.row_dimensions[7].height = 22

    # Row 8: Garage Allowance & Mode Toggles
    ws['A8'] = 'Garage Allowance (km):'
    ws['A8'].font = styles['font_bold']
    ws['B8'] = 6.00
    ws['B8'].fill = styles['fill_input']
    ws['B8'].font = styles['font_bold']
    ws['B8'].alignment = styles['align_center']
    ws['B8'].number_format = '0.00 "km"'
    
    ws['C8'] = 'Trip Mode Toggle:'
    ws['C8'].font = styles['font_bold']
    ws['D8'] = 'FORMULA'
    ws['D8'].fill = styles['fill_input']
    ws['D8'].font = styles['font_bold']
    ws['D8'].alignment = styles['align_center']
    dv_mode.add(ws['D8'])
    
    ws['E8'] = 'Fixed Trips Override:'
    ws['E8'].font = styles['font_bold']
    ws['F8'] = 3.00
    ws['F8'].fill = styles['fill_input']
    ws['F8'].font = styles['font_bold']
    ws['F8'].alignment = styles['align_center']
    ws['F8'].number_format = '0.00'
    
    ws['G8'] = 'Diesel Mileage Norm:'
    ws['G8'].font = styles['font_bold']
    ws['H8'] = 5.00
    ws['H8'].fill = styles['fill_input']
    ws['H8'].font = styles['font_bold']
    ws['H8'].alignment = styles['align_center']
    ws['H8'].number_format = '0.00 "km/L"'
    
    for c in ['A8', 'B8', 'C8', 'D8', 'E8', 'F8', 'G8', 'H8']:
        ws[c].border = styles['border_thin']
    ws.row_dimensions[8].height = 22

    # Row 9: Derived Fuel & Total Output
    ws['A9'] = 'Mobil Oil Mileage Norm:'
    ws['A9'].font = styles['font_bold']
    ws['B9'] = 140.00
    ws['B9'].fill = styles['fill_input']
    ws['B9'].font = styles['font_bold']
    ws['B9'].alignment = styles['align_center']
    ws['B9'].number_format = '0.00 "km/L"'
    
    ws['C9'] = 'Evaluated Diesel (Litres):'
    ws['C9'].font = styles['font_bold']
    ws['D9'] = '=ROUND(H7 / H8, 2)'
    ws['D9'].fill = styles['fill_subtotal']
    ws['D9'].font = styles['font_bold']
    ws['D9'].alignment = styles['align_center']
    ws['D9'].number_format = '0.00 "L"'
    
    ws['E9'] = 'Evaluated Mobil Oil (Litres):'
    ws['E9'].font = styles['font_bold']
    ws['F9'] = '=ROUND(H7 / B9, 3)'
    ws['F9'].fill = styles['fill_subtotal']
    ws['F9'].font = styles['font_bold']
    ws['F9'].alignment = styles['align_center']
    ws['F9'].number_format = '0.000 "L"'
    
    ws['G9'] = 'Total Daily Hauled Output:'
    ws['G9'].font = styles['font_bold']
    ws['H9'] = '=ROUND(F7 * B7, 2)'
    ws['H9'].fill = styles['fill_result']
    ws['H9'].font = styles['font_bold']
    ws['H9'].alignment = styles['align_center']
    ws['H9'].number_format = '0.00'
    
    for c in ['A9', 'B9', 'C9', 'D9', 'E9', 'F9', 'G9', 'H9']:
        ws[c].border = styles['border_thin']
    ws.row_dimensions[9].height = 22

    # Row 10: REAL-TIME IN-SHEET AUDIT STATUS PANEL
    ws['A10'] = 'AUDIT STATUS:'
    ws['A10'].font = styles['font_white_bold']
    ws['A10'].fill = styles['fill_header']
    ws['A10'].alignment = styles['align_center']
    
    ws['B10'] = '=IF(AND(D10="OK", F10="OK", H10="OK"), "[PASS] ALL CHECKS OK", "[ALERT] CHECKS FAILED")'
    ws['B10'].font = styles['font_result']
    ws['B10'].fill = styles['fill_result']
    ws['B10'].alignment = styles['align_center']
    
    ws['C10'] = 'Daily Trips Check:'
    ws['C10'].font = styles['font_note']
    ws['C10'].alignment = styles['align_right']
    
    ws['D10'] = '=IF(F7>0, "OK", "ERR: N<=0")'
    ws['D10'].font = styles['font_bold']
    ws['D10'].alignment = styles['align_center']
    
    ws['E10'] = 'Lead Distance Check:'
    ws['E10'].font = styles['font_note']
    ws['E10'].alignment = styles['align_right']
    
    ws['F10'] = '=IF(D5>0, "OK", "ERR: L<=0")'
    ws['F10'].font = styles['font_bold']
    ws['F10'].alignment = styles['align_center']
    
    ws['G10'] = 'Daily Cost Check:'
    ws['G10'].font = styles['font_note']
    ws['G10'].alignment = styles['align_right']
    
    ws['H10'] = '=IF(G22>0, "OK", "ERR: Cost<=0")'
    ws['H10'].font = styles['font_bold']
    ws['H10'].alignment = styles['align_center']
    
    for c in ['A10', 'B10', 'C10', 'D10', 'E10', 'F10', 'G10', 'H10']:
        ws[c].border = styles['border_thin']
    ws.row_dimensions[10].height = 22

    # Spacer
    ws.row_dimensions[11].height = 10

    # =========================================================================
    # SECTION 1: MECHANICAL TRANSPORT RESOURCE BUILDER (Rows 12 to 23)
    # =========================================================================
    ws.merge_cells('A12:H12')
    c_sec1 = ws['A12']
    c_sec1.value = '1. MECHANICAL TRANSPORT TRIP BUILDER (Truck Hire + Labour + Fuel Linked to CPWD Mileage Formulas)'
    c_sec1.font = styles['font_white_bold']
    c_sec1.fill = styles['fill_header']
    c_sec1.alignment = styles['align_left']
    ws.row_dimensions[12].height = 24
    
    tbl_headers = ['Line', 'Code / Source', 'Resource / Fuel Description', 'Unit', 'Quantity / Coeff', 'Basic Rate (Rs)', 'Amount (Rs)', 'Source Reference / CPWD Notes']
    for c_idx, h in enumerate(tbl_headers, 1):
        cell = ws.cell(row=13, column=c_idx, value=h)
        cell.font = styles['font_header']
        cell.fill = styles['fill_header']
        cell.alignment = styles['align_center']
        cell.border = styles['border_header']
    ws.row_dimensions[13].height = 24
    
    mech_resources = [
        {'code': '0084', 'qty_formula': 1.0, 'note': 'CPWD Note 5: Hire charges of Diesel Truck (9-tonne) excluding diesel & mobil oil (Code 0084)'},
        {'code': '0114', 'qty_formula': 6.0, 'note': 'CPWD Note 5: Loading & unloading labour based on 1 gang of 6 Beldars per truck (Code 0114)'},
        {'code': '1235', 'qty_formula': '=D9', 'note': 'CPWD Note 3: Diesel consumed = Distance travelled / 5.0 km/L (linked to cell D9)'},
        {'code': '5001', 'qty_formula': '=F9', 'note': 'CPWD Note 4: Mobil oil consumed = Distance travelled / 140.0 km/L (linked to cell F9)'}
    ]
    
    for idx in range(8):
        r = 14 + idx
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
        if r in [16, 17]:
            c_qty.fill = styles['fill_calc']
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
        
    # Row 22: Direct Daily Operating Cost (W)
    ws.merge_cells('A22:E22')
    ws['A22'] = 'Direct Daily Operating Cost (W) (Rs):'
    ws['A22'].font = styles['font_bold']
    ws['A22'].alignment = styles['align_right']
    ws['A22'].fill = styles['fill_subtotal']
    
    c_dirsum = ws['G22']
    c_dirsum.value = '=SUM(G14:G21)'
    c_dirsum.font = styles['font_bold']
    c_dirsum.alignment = styles['align_right']
    c_dirsum.number_format = styles['fmt_currency']
    c_dirsum.fill = styles['fill_subtotal']
    for c in ['A22', 'B22', 'C22', 'D22', 'E22', 'F22', 'G22', 'H22']:
        ws[c].border = styles['border_double_bottom']
    ws.row_dimensions[22].height = 22
    
    # Row 23: Operating Cost per Single Round Trip
    ws.merge_cells('A23:E23')
    ws['A23'] = 'Operating Cost per Single Round Trip (Rs):'
    ws['A23'].font = styles['font_bold']
    ws['A23'].alignment = styles['align_right']
    ws['A23'].fill = styles['fill_subtotal']
    
    c_tripsum = ws['G23']
    c_tripsum.value = '=ROUND(G22 / F7, 2)'
    c_tripsum.font = styles['font_bold']
    c_tripsum.alignment = styles['align_right']
    c_tripsum.number_format = styles['fmt_currency']
    c_tripsum.fill = styles['fill_result']
    
    ws['H23'] = 'Evaluated operating cost per single round trip (= W / N)'
    ws['H23'].font = styles['font_note']
    ws['H23'].alignment = styles['align_left']
    
    for c in ['A23', 'B23', 'C23', 'D23', 'E23', 'F23', 'G23', 'H23']:
        ws[c].border = styles['border_thin']
    ws.row_dimensions[23].height = 22

    # Spacer
    ws.row_dimensions[24].height = 10

    # =========================================================================
    # SECTION 2: STATUTORY MARKUPS & UNIT RATE EVALUATION (Rows 25 to 38)
    # =========================================================================
    ws.merge_cells('A25:H25')
    c_sec2 = ws['A25']
    c_sec2.value = '2. STATUTORY MARKUPS & UNIT RATE DERIVATION (CPWD CARRIAGE CONVENTION: 15% CPOH ONLY)'
    c_sec2.font = styles['font_white_bold']
    c_sec2.fill = styles['fill_header']
    c_sec2.alignment = styles['align_left']
    ws.row_dimensions[25].height = 24
    
    stat_headers = ['Item', 'Description / Stage', 'Apply? (YES/NO)', 'Basis Applied', 'Base Amount (Rs)', 'Factor / %', 'Amount (Rs)', 'CPWD Statutory Rule & Guidance Note']
    for c_idx, h in enumerate(stat_headers, 1):
        cell = ws.cell(row=26, column=c_idx, value=h)
        cell.font = styles['font_header']
        cell.fill = styles['fill_header']
        cell.alignment = styles['align_center']
        cell.border = styles['border_header']
    ws.row_dimensions[26].height = 24
    
    # Row 27: Base Direct Cost (W)
    ws['A27'] = 'W'
    ws['B27'] = 'Direct Daily Operating Cost (Truck + Labour + Fuel)'
    ws['C27'] = '-'
    ws['D27'] = 'Direct Sum'
    ws['E27'] = '-'
    ws['F27'] = '-'
    ws['G27'] = '=G22'
    ws['H27'] = 'Total direct operating shift cost before statutory overheads'
    
    # Row 28: Water Charges (NO by default per CPWD)
    ws['A28'] = 'X1'
    ws['B28'] = 'Add Water Charges (1% on W)'
    ws['C28'] = 'NO'
    ws['D28'] = 'On "W"'
    ws['E28'] = '=G27'
    ws['F28'] = '=Factor_Water'
    ws['G28'] = '=IF(C28="YES", ROUND(E28 * F28, 2), 0)'
    ws['H28'] = 'CPWD RULE: Omitted (NO) in Carriage items because no water is consumed in haulage.'
    
    # Row 29: Subtotal (X)
    ws['A29'] = 'X'
    ws['B29'] = 'Subtotal "X" (W + Water Charges)'
    ws['C29'] = '-'
    ws['D29'] = 'W + Water'
    ws['E29'] = '-'
    ws['F29'] = '-'
    ws['G29'] = '=G27 + G28'
    ws['H29'] = 'Base for GST (if applicable)'
    
    # Row 30: GST (NO by default in DAR)
    ws['A30'] = 'Y1'
    ws['B30'] = 'Add GST on Transport Freight'
    ws['C30'] = 'NO'
    ws['D30'] = 'On "X"'
    ws['E30'] = '=G29'
    ws['F30'] = '=Factor_GST'
    ws['G30'] = '=IF(C30="YES", ROUND(E30 * F30, 2), 0)'
    ws['H30'] = 'CPWD RULE: Omitted (NO) in DAR Carriage items; tax on vehicle freight is handled separately.'
    
    # Row 31: Subtotal (Y)
    ws['A31'] = 'Y'
    ws['B31'] = 'Subtotal "Y" (X + GST)'
    ws['C31'] = '-'
    ws['D31'] = 'X + GST'
    ws['E31'] = '-'
    ws['F31'] = '-'
    ws['G31'] = '=G29 + G30'
    ws['H31'] = 'Base for Contractor Profit & Overheads'
    
    # Row 32: CPOH (YES by default, 15%)
    ws['A32'] = 'Z1'
    ws['B32'] = 'Add Contractor Profit & Overheads (15% CPOH)'
    ws['C32'] = 'YES'
    ws['D32'] = 'On "Y"'
    ws['E32'] = '=G31'
    ws['F32'] = '=Factor_CPOH'
    ws['G32'] = '=IF(C32="YES", ROUND(E32 * F32, 2), 0)'
    ws['H32'] = 'CPWD RULE: 15% CPOH is added on Carriage direct operating cost per DAR 2019 Item 1.1.'
    
    # Row 33: Subtotal (Z)
    ws['A33'] = 'Z'
    ws['B33'] = 'Subtotal "Z" (Y + CPOH)'
    ws['C33'] = '-'
    ws['D33'] = 'Y + CPOH'
    ws['E33'] = '-'
    ws['F33'] = '-'
    ws['G33'] = '=G31 + G32'
    ws['H33'] = 'Base for Labour Welfare Cess'
    
    # Row 34: BOCW Cess (NO by default)
    ws['A34'] = 'Z2'
    ws['B34'] = 'Add Labour Welfare Cess (1% BOCW Cess)'
    ws['C34'] = 'NO'
    ws['D34'] = 'On "Z"'
    ws['E34'] = '=G33'
    ws['F34'] = '=Factor_Cess'
    ws['G34'] = '=IF(C34="YES", ROUND(E34 * F34, 2), 0)'
    ws['H34'] = 'CPWD RULE: Omitted (NO) in standard Carriage items. Toggle YES if contract requires cess on haulage.'
    
    # Row 35: Total Daily Cost with Markups
    ws['A35'] = 'Cost'
    ws['B35'] = 'Total Daily Carriage Cost (including CPOH):'
    ws['C35'] = '-'
    ws['D35'] = 'Z + Cess'
    ws['E35'] = '-'
    ws['F35'] = '-'
    ws['G35'] = '=G33 + G34'
    ws['H35'] = 'Total evaluated carriage cost for 1 full working shift (8 hours)'
    
    # Row 36: Analyzed Rate per 1.00 Unit
    ws['A36'] = 'Rate'
    ws['B36'] = '="Analyzed Carriage Rate per 1.00 " & D7 & ":"'
    ws['C36'] = '-'
    ws['D36'] = 'Cost / Daily Qty'
    ws['E36'] = '-'
    ws['F36'] = '-'
    ws['G36'] = '=ROUND(G35 / H9, 2)'
    ws['H36'] = 'Derived carriage rate per base unit (= Total Cost / Daily Output Qty in cell H9)'
    
    # Row 37: Analyzed Rate per 100 Units (Schedule Unit for Pipes/Tiles)
    ws['A37'] = 'Rate/100'
    ws['B37'] = '="Analyzed Carriage Rate per 100 " & D7 & " (Schedule Rate):"'
    ws['C37'] = '-'
    ws['D37'] = 'Rate * 100'
    ws['E37'] = '-'
    ws['F37'] = '-'
    ws['G37'] = '=IF(D7="metre", ROUND(G36 * 100, 2), IF(D7="cum", G36, ROUND(G36 * 100, 2)))'
    ws['H37'] = 'Schedule rate for 100 m / 1000 Nos (as per CPWD DAR pipe schedule unit)'
    
    # Row 38: OFFICIAL SAY RATE
    ws['A38'] = 'SAY'
    ws['B38'] = '="OFFICIAL SAY RATE (Rs per " & D7 & "):"'
    ws['C38'] = '-'
    ws['D38'] = 'Final Say'
    ws['E38'] = '-'
    ws['F38'] = '-'
    ws['G38'] = '=ROUND(G36, 2)'
    ws['H38'] = 'OFFICIAL CPWD ROUNDED RATE FOR TENDER SCHEDULES & HAULAGE SPECIFICATIONS'
    
    for r in range(27, 39):
        ws.cell(row=r, column=1).alignment = styles['align_center']
        ws.cell(row=r, column=1).font = styles['font_bold']
        ws.cell(row=r, column=2).alignment = styles['align_left']
        ws.cell(row=r, column=3).alignment = styles['align_center']
        ws.cell(row=r, column=4).alignment = styles['align_center']
        
        c_base = ws.cell(row=r, column=5)
        c_base.alignment = styles['align_right']
        if r in [28, 30, 32, 34]:
            c_base.number_format = styles['fmt_currency']
            
        c_fac = ws.cell(row=r, column=6)
        c_fac.alignment = styles['align_right']
        if r in [28, 30, 32, 34]:
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
            
        if r in [27, 29, 31, 33]:
            for c in range(1, 8):
                ws.cell(row=r, column=c).fill = styles['fill_subtotal']
        elif r in [28, 30, 32, 34]:
            ws.cell(row=r, column=3).fill = styles['fill_input']
            ws.cell(row=r, column=3).font = styles['font_bold']
            dv_yesno.add(ws.cell(row=r, column=3))
        elif r == 35:
            for c in range(1, 8):
                ws.cell(row=r, column=c).fill = styles['fill_result']
        elif r in [36, 37]:
            for c in range(1, 8):
                ws.cell(row=r, column=c).fill = styles['fill_subtotal']
        elif r == 38:
            ws.row_dimensions[38].height = 30
            for c in range(1, 8):
                ws.cell(row=r, column=c).fill = styles['fill_say']
            ws['B38'].font = styles['font_say']
            ws['G38'].font = styles['font_say']
            
        ws.row_dimensions[r].height = 24 if r != 38 else 30

    # Spacer
    ws.row_dimensions[39].height = 10

    # =========================================================================
    # SECTION 3: CPWD DATA SHEET NO. 1 GROUND-TRUTH BENCHMARK (Rows 40 to 73)
    # =========================================================================
    ws.merge_cells('A40:H40')
    c_sec3 = ws['A40']
    c_sec3.value = '3. CPWD DAR 2019 DATA SHEET NO. 1 GROUND-TRUTH BENCHMARK (1 km to 30 km Mechanical Transport Reference)'
    c_sec3.font = styles['font_white_bold']
    c_sec3.fill = styles['fill_header']
    c_sec3.alignment = styles['align_left']
    ws.row_dimensions[40].height = 24
    
    bench_headers = ['Lead (L) (km)', 'Avg Speed (S) (km/h)', 'Trips (N) / Day', 'Km Done / Day', 'Diesel Qty (Litres)', 'Mobil Oil (Litres)', 'Total Shift Cost (Rs)', 'Cost per Trip (Rs)']
    for c_idx, h in enumerate(bench_headers, 1):
        cell = ws.cell(row=41, column=c_idx, value=h)
        cell.font = styles['font_header']
        cell.fill = styles['fill_header']
        cell.alignment = styles['align_center']
        cell.border = styles['border_header']
    ws.row_dimensions[41].height = 24
    
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
        r = 42 + idx
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
        
    # Benchmark Footer Notes
    ws.merge_cells('A72:H72')
    c_bnote = ws['A72']
    c_bnote.value = 'NOTES: N=8/((2L/S)+1); Km Done=(2NL)+6.0; Diesel=Km/5.0; Mobil Oil=Km/140.0; Cost=Diesel+Oil+Rs 3348 (6 Beldars)+Rs 1500 (Truck); Cost/Trip=Total Cost/N.'
    c_bnote.font = styles['font_note']
    c_bnote.fill = styles['fill_note']
    c_bnote.alignment = styles['align_left']
    c_bnote.border = styles['border_thin']
    ws.row_dimensions[72].height = 20

    # Spacer
    ws.row_dimensions[73].height = 10

    # =========================================================================
    # SECTION 4: CPWD STANDARD MATERIAL PAYLOAD CAPACITIES MATRIX (Rows 74 to 105)
    # =========================================================================
    ws.merge_cells('A74:H74')
    c_sec4 = ws['A74']
    c_sec4.value = '4. CPWD DAR STANDARD MATERIAL PAYLOAD CAPACITIES MATRIX (Table 1.1 Technical Reference)'
    c_sec4.font = styles['font_white_bold']
    c_sec4.fill = styles['fill_header']
    c_sec4.alignment = styles['align_left']
    ws.row_dimensions[74].height = 24
    
    cap_headers = ['DAR Code', 'Material / Trade Specification', 'Truck Payload / Trip', 'Net Payable Qty', 'Schedule Unit', 'Looseness Allowance / Technical Norms', 'Base Rate 1 km (Rs)', 'Addl Rate >20 km (Rs)']
    for c_idx, h in enumerate(cap_headers, 1):
        cell = ws.cell(row=75, column=c_idx, value=h)
        cell.font = styles['font_header']
        cell.fill = styles['fill_header']
        cell.alignment = styles['align_center']
        cell.border = styles['border_header']
    ws.row_dimensions[75].height = 24
    
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
        r = 76 + idx
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

    # Spacer
    r_spacer1 = 76 + len(capacities_table)
    ws.row_dimensions[r_spacer1].height = 12

    # =========================================================================
    # PART B: HEADING 1.2 — MANUAL LABOUR CARRIAGE BUILDER & DISTANCE MATRIX (< 0.50 KM)
    # =========================================================================
    r_sec5 = r_spacer1 + 1
    ws.merge_cells(f'A{r_sec5}:H{r_sec5}')
    c_sec5 = ws[f'A{r_sec5}']
    c_sec5.value = '5. HEADING 1.2: MANUAL LABOUR CARRIAGE CALCULATOR (For Lead Less Than 0.50 km / 50 m to 500 m)'
    c_sec5.font = styles['font_white_bold']
    c_sec5.fill = styles['fill_header']
    c_sec5.alignment = styles['align_left']
    ws.row_dimensions[r_sec5].height = 24
    
    r_mbanner = r_sec5 + 1
    ws.merge_cells(f'A{r_mbanner}:H{r_mbanner}')
    c_mbanner = ws[f'A{r_mbanner}']
    c_mbanner.value = 'CPWD DAR 2019 Item 1.2 Standards: Category A (Bulk Materials) = 7.67 Beldars 1st 50m (+1.67 Coolies/addl 50m). Category B (Heavy/Pipes/Steel) = 9.20 Beldars 1st 50m (+1.35 Beldars/addl 50m). All rates include 15% CPOH.'
    c_mbanner.font = styles['font_note']
    c_mbanner.fill = styles['fill_note']
    c_mbanner.alignment = styles['align_left']
    ws.row_dimensions[r_mbanner].height = 22
    
    # Manual Calculator Input Rows
    r_minp1 = r_mbanner + 1
    ws[f'A{r_minp1}'] = 'Manual Item Code:'
    ws[f'A{r_minp1}'].font = styles['font_bold']
    ws[f'B{r_minp1}'] = '1.2.CUSTOM'
    ws[f'B{r_minp1}'].fill = styles['fill_input']
    ws[f'B{r_minp1}'].font = styles['font_bold']
    ws[f'B{r_minp1}'].alignment = styles['align_center']
    
    ws[f'C{r_minp1}'] = 'Material Category:'
    ws[f'C{r_minp1}'].font = styles['font_bold']
    ws[f'D{r_minp1}'] = 'Category B (Heavy / Pipes / Steel)'
    ws[f'D{r_minp1}'].fill = styles['fill_input']
    ws[f'D{r_minp1}'].font = styles['font_bold']
    ws[f'D{r_minp1}'].alignment = styles['align_center']
    dv_manual_cat.add(ws[f'D{r_minp1}'])
    
    ws[f'E{r_minp1}'] = 'Lead Distance (Metres):'
    ws[f'E{r_minp1}'].font = styles['font_bold']
    ws[f'F{r_minp1}'] = 100
    ws[f'F{r_minp1}'].fill = styles['fill_input']
    ws[f'F{r_minp1}'].font = styles['font_bold']
    ws[f'F{r_minp1}'].alignment = styles['align_center']
    dv_manual_lead.add(ws[f'F{r_minp1}'])
    
    ws[f'G{r_minp1}'] = 'Addl 50m Steps (M):'
    ws[f'G{r_minp1}'].font = styles['font_bold']
    ws[f'H{r_minp1}'] = f'=MAX(0, (F{r_minp1} - 50) / 50)'
    ws[f'H{r_minp1}'].fill = styles['fill_subtotal']
    ws[f'H{r_minp1}'].font = styles['font_bold']
    ws[f'H{r_minp1}'].alignment = styles['align_center']
    ws[f'H{r_minp1}'].number_format = '0'
    
    for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']:
        ws[f'{col}{r_minp1}'].border = styles['border_thin']
    ws.row_dimensions[r_minp1].height = 22
    
    r_minp2 = r_minp1 + 1
    ws[f'A{r_minp2}'] = 'Manual Item Specification:'
    ws[f'A{r_minp2}'].font = styles['font_bold']
    ws[f'A{r_minp2}'].border = styles['border_thin']
    ws.merge_cells(f'B{r_minp2}:H{r_minp2}')
    ws[f'B{r_minp2}'] = 'Carriage by manual labour including loading, unloading and stacking for lead upto 100 metres complete as per directions of Engineer-in-charge.'
    ws[f'B{r_minp2}'].font = styles['font_regular']
    ws[f'B{r_minp2}'].fill = styles['fill_input']
    ws[f'B{r_minp2}'].alignment = styles['align_wrap']
    ws[f'B{r_minp2}'].border = styles['border_thin']
    ws.row_dimensions[r_minp2].height = 28
    
    r_minp3 = r_minp2 + 1
    ws[f'A{r_minp3}'] = '8-Hour Output Capacity:'
    ws[f'A{r_minp3}'].font = styles['font_bold']
    ws[f'B{r_minp3}'] = 1702.00 # e.g. 100mm RCC pipe capacity per 8 hrs
    ws[f'B{r_minp3}'].fill = styles['fill_input']
    ws[f'B{r_minp3}'].font = styles['font_bold']
    ws[f'B{r_minp3}'].alignment = styles['align_center']
    ws[f'B{r_minp3}'].number_format = '0.00'
    
    ws[f'C{r_minp3}'] = 'Output / Billing Unit:'
    ws[f'C{r_minp3}'].font = styles['font_bold']
    ws[f'D{r_minp3}'] = '100 m'
    ws[f'D{r_minp3}'].fill = styles['fill_input']
    ws[f'D{r_minp3}'].font = styles['font_bold']
    ws[f'D{r_minp3}'].alignment = styles['align_center']
    dv_unit.add(ws[f'D{r_minp3}'])
    
    ws[f'E{r_minp3}'] = 'Base Labour 1st 50m (Rs):'
    ws[f'E{r_minp3}'].font = styles['font_bold']
    ws[f'F{r_minp3}'] = f'=IF(ISNUMBER(SEARCH("Category A", D{r_minp1})), 4279.86, 5133.60)'
    ws[f'F{r_minp3}'].fill = styles['fill_subtotal']
    ws[f'F{r_minp3}'].font = styles['font_bold']
    ws[f'F{r_minp3}'].alignment = styles['align_right']
    ws[f'F{r_minp3}'].number_format = styles['fmt_currency']
    
    ws[f'G{r_minp3}'] = 'Addl Labour / 50m (Rs):'
    ws[f'G{r_minp3}'].font = styles['font_bold']
    ws[f'H{r_minp3}'] = f'=IF(ISNUMBER(SEARCH("Category A", D{r_minp1})), 931.86, 753.30)'
    ws[f'H{r_minp3}'].fill = styles['fill_subtotal']
    ws[f'H{r_minp3}'].font = styles['font_bold']
    ws[f'H{r_minp3}'].alignment = styles['align_right']
    ws[f'H{r_minp3}'].number_format = styles['fmt_currency']
    
    for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']:
        ws[f'{col}{r_minp3}'].border = styles['border_thin']
    ws.row_dimensions[r_minp3].height = 22
    
    r_minp4 = r_minp3 + 1
    ws[f'A{r_minp4}'] = 'Total Labour Cost (Rs):'
    ws[f'A{r_minp4}'].font = styles['font_bold']
    ws[f'B{r_minp4}'] = f'=F{r_minp3} + (H{r_minp1} * H{r_minp3})'
    ws[f'B{r_minp4}'].fill = styles['fill_subtotal']
    ws[f'B{r_minp4}'].font = styles['font_bold']
    ws[f'B{r_minp4}'].alignment = styles['align_right']
    ws[f'B{r_minp4}'].number_format = styles['fmt_currency']
    
    ws[f'C{r_minp4}'] = 'Add 15% CPOH (Rs):'
    ws[f'C{r_minp4}'].font = styles['font_bold']
    ws[f'D{r_minp4}'] = f'=ROUND(B{r_minp4} * 0.15, 2)'
    ws[f'D{r_minp4}'].fill = styles['fill_subtotal']
    ws[f'D{r_minp4}'].font = styles['font_bold']
    ws[f'D{r_minp4}'].alignment = styles['align_right']
    ws[f'D{r_minp4}'].number_format = styles['fmt_currency']
    
    ws[f'E{r_minp4}'] = 'Total 8-Hr Cost with CPOH:'
    ws[f'E{r_minp4}'].font = styles['font_bold']
    ws[f'F{r_minp4}'] = f'=B{r_minp4} + D{r_minp4}'
    ws[f'F{r_minp4}'].fill = styles['fill_result']
    ws[f'F{r_minp4}'].font = styles['font_bold']
    ws[f'F{r_minp4}'].alignment = styles['align_right']
    ws[f'F{r_minp4}'].number_format = styles['fmt_currency']
    
    ws[f'G{r_minp4}'] = 'Analyzed Unit Rate (Rs):'
    ws[f'G{r_minp4}'].font = styles['font_bold']
    ws[f'H{r_minp4}'] = f'=IF(OR(D{r_minp3}="100 m", D{r_minp3}="1000 Nos"), ROUND((F{r_minp4} / B{r_minp3}) * 100, 2), ROUND(F{r_minp4} / B{r_minp3}, 2))'
    ws[f'H{r_minp4}'].fill = styles['fill_say']
    ws[f'H{r_minp4}'].font = styles['font_say']
    ws[f'H{r_minp4}'].alignment = styles['align_right']
    ws[f'H{r_minp4}'].number_format = styles['fmt_currency']
    
    for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']:
        ws[f'{col}{r_minp4}'].border = styles['border_thin']
    ws.row_dimensions[r_minp4].height = 26

    # Spacer
    r_spacer2 = r_minp4 + 1
    ws.row_dimensions[r_spacer2].height = 10

    # =========================================================================
    # SECTION 6: CPWD TABLE 1.2 MANUAL LABOUR GROUND-TRUTH MATRIX (Rows 117+)
    # =========================================================================
    r_sec6 = r_spacer2 + 1
    ws.merge_cells(f'A{r_sec6}:H{r_sec6}')
    c_sec6 = ws[f'A{r_sec6}']
    c_sec6.value = '6. CPWD DAR TABLE 1.2 MANUAL LABOUR GROUND-TRUTH MATRIX (< 0.50 km Standards)'
    c_sec6.font = styles['font_white_bold']
    c_sec6.fill = styles['fill_header']
    c_sec6.alignment = styles['align_left']
    ws.row_dimensions[r_sec6].height = 24
    
    man_headers = ['DAR Code', 'Material Specification', '8-Hour Capacity', 'Net Payable Qty', 'Unit of Rate', 'Cost for 1st 50 m (Rs)', 'Cost for Addl 50 m (Rs)', 'CPWD DAR Labour Composition']
    r_manhead = r_sec6 + 1
    for c_idx, h in enumerate(man_headers, 1):
        cell = ws.cell(row=r_manhead, column=c_idx, value=h)
        cell.font = styles['font_header']
        cell.fill = styles['fill_header']
        cell.alignment = styles['align_center']
        cell.border = styles['border_header']
    ws.row_dimensions[r_manhead].height = 24
    
    manual_table_data = [
        ('1.2.1', 'Lime, moorum, building rubbish, malba', 35.00, 35.00, 'cum', 140.62, 30.62, '7.67 Beldars @ Rs 558 + 15% CPOH (Addl: 1.67 Coolies)'),
        ('1.2.2', 'Earth (excavated soil / good earth)', 35.00, 28.00, 'cum', 175.78, 38.27, '20% looseness deduction (Net = 28 cum)'),
        ('1.2.3', 'Manure or sludge', 35.00, 32.20, 'cum', 152.85, 33.28, '8% looseness deduction (Net = 32.2 cum)'),
        ('1.2.4', 'Excavated rock', 35.00, 17.50, 'cum', 281.25, 61.24, '50% voids deduction (Net = 17.5 cum)'),
        ('1.2.5', 'Sand, stone aggregate below 40 mm', 28.00, 28.00, 'cum', 175.78, 38.27, '28 cum density norm (Nil looseness)'),
        ('1.2.6', 'Stone aggregate 40 mm & above', 28.00, 25.90, 'cum', 190.03, 41.38, '8% voids deduction (Net = 25.9 cum)'),
        ('1.2.7', 'Soling stone & masonry stone', 28.00, 23.80, 'cum', 206.80, 45.03, '15% voids deduction (Net = 23.8 cum)'),
        ('1.2.8', 'Bricks (conventional / modular)', 15000.0, 15000.0, '1000 Nos', 328.12, 71.44, '15,000 Bricks per 8 hours manual gang'),
        ('1.2.9', 'Brick tiles / roofing tiles', 24000.0, 24000.0, '1000 Nos', 205.08, 44.65, '24,000 Tiles per 8 hours manual gang'),
        ('1.2.10', 'Steam coal', 30.00, 30.00, 'tonne', 164.06, 35.72, '30 tonnes per 8 hours manual gang'),
        ('1.2.11', 'Stone blocks, pipes <100mm, heavy items', 46.00, 46.00, 'tonne', 128.34, 18.83, 'Category B: 9.20 Beldars 1st 50m (+1.35 Beldars addl)'),
        ('1.2.12', 'Cement in bags', 57.99, 57.99, 'tonne', 101.80, 14.94, '57.99 tonnes (approx 1,160 bags) per shift'),
        ('1.2.13', 'Steel bars & structural steel', 27.00, 27.00, 'tonne', 218.65, 32.09, '27 tonnes steel shifting per 8 hours shift'),
        ('1.2.14', 'Timber scantlings & logs', 42.00, 42.00, 'cum', 140.56, 20.63, '42 cum timber manual haulage per shift'),
        ('1.2.15', 'Tar, bitumen in drums', 46.00, 46.00, 'tonne', 128.34, 18.83, '46 tonnes drum handling per shift'),
        ('1.2.16.1', 'S.W. pipes 100 mm dia', 2298.0, 2298.0, '100 m', 256.90, 37.70, '2,298 metres manual carrying per shift'),
        ('1.2.16.2', 'S.W. pipes 150 mm dia', 1398.0, 1398.0, '100 m', 422.29, 61.97, '1,398 metres manual carrying per shift'),
        ('1.2.16.3', 'S.W. pipes 200 mm dia', 999.0, 999.0, '100 m', 590.95, 86.72, '999 metres manual carrying per shift'),
        ('1.2.16.5', 'S.W. pipes 250 mm dia', 600.0, 600.0, '100 m', 983.94, 144.38, '600 metres manual carrying per shift'),
        ('1.2.17.1', 'R.C.C. pipes 100 mm dia', 1702.0, 1702.0, '100 m', 346.86, 50.90, '1,702 metres manual carrying per shift'),
        ('1.2.17.2', 'R.C.C. pipes 125 mm dia', 1391.0, 1391.0, '100 m', 424.42, 62.28, '1,391 metres manual carrying per shift'),
        ('1.2.17.3', 'R.C.C. pipes 150 mm dia', 1208.0, 1208.0, '100 m', 488.71, 71.71, '1,208 metres manual carrying per shift'),
        ('1.2.17.4', 'R.C.C. pipes 200 mm dia', 805.0, 805.0, '100 m', 733.37, 107.61, '805 metres manual carrying per shift'),
        ('1.2.17.5', 'R.C.C. pipes 250 mm dia', 458.0, 458.0, '100 m', 1289.00, 189.15, '458 metres manual carrying per shift'),
        ('1.2.17.6', 'R.C.C. pipes 300 mm dia', 366.0, 366.0, '100 m', 1613.02, 236.69, '366 metres manual carrying per shift'),
        ('1.2.17.10', 'R.C.C. pipes 600, 700, 750 & 800 mm dia', 150.0, 150.0, '100 m', 3935.76, 577.53, '150 metres manual carrying per shift')
    ]
    
    for idx, mrow in enumerate(manual_table_data):
        r = r_manhead + 1 + idx
        ws.cell(row=r, column=1, value=mrow[0]).alignment = styles['align_center']
        ws.cell(row=r, column=1).font = styles['font_bold']
        
        ws.cell(row=r, column=2, value=mrow[1]).alignment = styles['align_left']
        
        ws.cell(row=r, column=3, value=mrow[2]).alignment = styles['align_right']
        ws.cell(row=r, column=3).number_format = '0.00'
        
        ws.cell(row=r, column=4, value=mrow[3]).alignment = styles['align_right']
        ws.cell(row=r, column=4).number_format = '0.00'
        
        ws.cell(row=r, column=5, value=mrow[4]).alignment = styles['align_center']
        
        c_mcost1 = ws.cell(row=r, column=6, value=mrow[5])
        c_mcost1.alignment = styles['align_right']
        c_mcost1.number_format = styles['fmt_currency']
        
        c_mcost2 = ws.cell(row=r, column=7, value=mrow[6])
        c_mcost2.alignment = styles['align_right']
        c_mcost2.number_format = styles['fmt_currency']
        
        ws.cell(row=r, column=8, value=mrow[7]).alignment = styles['align_left']
        ws.cell(row=r, column=8).font = styles['font_note']
        
        row_fill = styles['fill_subtotal'] if r % 2 == 0 else styles['fill_calc']
        for c in range(1, 9):
            cell = ws.cell(row=r, column=c)
            cell.fill = row_fill
            cell.border = styles['border_thin']
            if c not in [1, 6, 7]:
                cell.font = styles['font_regular']
        ws.row_dimensions[r].height = 20

    # Spacer
    r_spacer3 = r_manhead + 1 + len(manual_table_data)
    ws.row_dimensions[r_spacer3].height = 12

    # =========================================================================
    # SECTION 7: RUNNING CUSTOM NON-DSR CARRIAGE LIBRARY (Rows 146+)
    # =========================================================================
    r_sec7 = r_spacer3 + 1
    ws.merge_cells(f'A{r_sec7}:H{r_sec7}')
    c_sec7 = ws[f'A{r_sec7}']
    c_sec7.value = '7. RUNNING CUSTOM NON-DSR ITEMS LIBRARY FOR CARRIAGE (Mechanical & Manual)'
    c_sec7.font = styles['font_white_bold']
    c_sec7.fill = styles['fill_header']
    c_sec7.alignment = styles['align_left']
    ws.row_dimensions[r_sec7].height = 24
    
    r_linst = r_sec7 + 1
    ws.merge_cells(f'A{r_linst}:H{r_linst}')
    c_linst = ws[f'A{r_linst}']
    c_linst.value = 'Log and save completed custom haulage rate analyses (Mechanical or Manual) below for immediate reference across project estimates.'
    c_linst.font = styles['font_note']
    c_linst.fill = styles['fill_note']
    c_linst.alignment = styles['align_left']
    ws.row_dimensions[r_linst].height = 20
    
    r_llibhead = r_linst + 1
    lib_headers = ['Item Code', 'Item Nomenclature / Specification', 'Unit', 'Output Basis', 'Direct Cost W (Rs)', 'Markups Applied', 'Unit Rate (Rs)', 'Say Rate (Rs)']
    for c_idx, h in enumerate(lib_headers, 1):
        cell = ws.cell(row=r_llibhead, column=c_idx, value=h)
        cell.font = styles['font_header']
        cell.fill = styles['fill_header']
        cell.alignment = styles['align_center']
        cell.border = styles['border_header']
    ws.row_dimensions[r_llibhead].height = 24
    
    combined_lib = [
        {'code': '1.1.18', 'desc': 'Disposal of building malba by mechanical transport - lead 10 km (Restricted urban 3 trips, 8 cum/trip)', 'unit': 'cum', 'basis': 24.0, 'w': 5939.58, 'markups': '15% CPOH only', 'rate': 284.60, 'say': 284.60},
        {'code': '1.1.17.12-VAR', 'desc': 'Transport of 1000, 1100 & 1200 mm dia pipes - lead 26 km (N=2.86 trips, Speed 29 km/h, Payload 10.98 m/trip)', 'unit': 'metre', 'basis': 31.40, 'w': 7470.17, 'markups': '15% CPOH only', 'rate': 273.56, 'say': 273.60},
        {'code': '1.1.1', 'desc': 'Carriage of Lime, moorum, building rubbish by mechanical transport - lead 5 km (N=5.19 trips, 8 cum/trip)', 'unit': 'cum', 'basis': 41.52, 'w': 5829.54, 'markups': '15% CPOH only', 'rate': 161.46, 'say': 161.50},
        {'code': '1.2.1', 'desc': 'Manual carriage of lime, moorum, rubbish for lead upto 50 metres (Category A)', 'unit': 'cum', 'basis': 35.0, 'w': 4279.86, 'markups': '15% CPOH only', 'rate': 140.62, 'say': 140.60},
        {'code': '1.2.13', 'desc': 'Manual carriage of steel bars & structural steel for lead upto 50 metres (Category B)', 'unit': 'tonne', 'basis': 27.0, 'w': 5133.60, 'markups': '15% CPOH only', 'rate': 218.65, 'say': 218.65},
        {'code': '1.2.17.1', 'desc': 'Manual carriage of R.C.C. pipes 100 mm dia for lead upto 100 metres (Category B)', 'unit': '100 m', 'basis': 17.02, 'w': 5886.90, 'markups': '15% CPOH only', 'rate': 397.76, 'say': 397.75}
    ]
    
    for idx in range(12):
        r = r_llibhead + 1 + idx
        item = combined_lib[idx] if idx < len(combined_lib) else None
        
        ws.cell(row=r, column=1, value=item['code'] if item else f"C-01.{idx+1:02d}").alignment = styles['align_center']
        ws.cell(row=r, column=2, value=item['desc'] if item else '(Available for new custom carriage analysis)').alignment = styles['align_left']
        ws.cell(row=r, column=3, value=item['unit'] if item else 'cum').alignment = styles['align_center']
        ws.cell(row=r, column=4, value=item['basis'] if item else 1.0).alignment = styles['align_center']
        
        c_w = ws.cell(row=r, column=5, value=item['w'] if item else None)
        c_w.alignment = styles['align_right']
        c_w.number_format = styles['fmt_currency']
        
        ws.cell(row=r, column=6, value=item['markups'] if item else '15% CPOH only').alignment = styles['align_center']
        
        c_rate = ws.cell(row=r, column=7, value=item['rate'] if item else None)
        c_rate.alignment = styles['align_right']
        c_rate.number_format = styles['fmt_currency']
        
        c_say = ws.cell(row=r, column=8, value=item['say'] if item else None)
        c_say.alignment = styles['align_right']
        c_say.number_format = styles['fmt_currency']
        c_say.font = styles['font_bold']
        
        row_fill = styles['fill_subtotal'] if r % 2 == 0 else styles['fill_calc']
        for c in range(1, 9):
            cell = ws.cell(row=r, column=c)
            cell.fill = row_fill
            cell.border = styles['border_thin']
            if c not in [1, 8]:
                cell.font = styles['font_regular']
        ws.row_dimensions[r].height = 20
        
    # Column Widths
    ws.column_dimensions['A'].width = 14
    ws.column_dimensions['B'].width = 30
    ws.column_dimensions['C'].width = 44
    ws.column_dimensions['D'].width = 18
    ws.column_dimensions['E'].width = 22
    ws.column_dimensions['F'].width = 18
    ws.column_dimensions['G'].width = 22
    ws.column_dimensions['H'].width = 48
    
    # =========================================================================
    # CELL PROTECTION CONFIGURATION (MS Excel 2016)
    # =========================================================================
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            ws.cell(row=r, column=c).protection = Protection(locked=True)
            
    # Part A Unlocked Input Cells
    ws['B5'].protection = Protection(locked=False)
    ws['D5'].protection = Protection(locked=False)
    ws['F5'].protection = Protection(locked=False)
    ws['H5'].protection = Protection(locked=False)
    
    ws['B6'].protection = Protection(locked=False)
    
    ws['B7'].protection = Protection(locked=False)
    ws['D7'].protection = Protection(locked=False)
    
    ws['B8'].protection = Protection(locked=False)
    ws['D8'].protection = Protection(locked=False)
    ws['F8'].protection = Protection(locked=False)
    ws['H8'].protection = Protection(locked=False)
    
    ws['B9'].protection = Protection(locked=False)
    
    # Section 1 Unlocked inputs
    for r in range(14, 22):
        ws.cell(row=r, column=2).protection = Protection(locked=False) # Code
        if r not in [16, 17]: # Row 16 (Diesel) and Row 17 (Oil) are formula-driven locked
            ws.cell(row=r, column=5).protection = Protection(locked=False) # Qty
        ws.cell(row=r, column=8).protection = Protection(locked=False) # Remarks
        
    # Section 2 Toggles
    ws['C28'].protection = Protection(locked=False)
    ws['C30'].protection = Protection(locked=False)
    ws['C32'].protection = Protection(locked=False)
    ws['C34'].protection = Protection(locked=False)
    
    # Part B Unlocked Input Cells
    ws[f'B{r_minp1}'].protection = Protection(locked=False)
    ws[f'D{r_minp1}'].protection = Protection(locked=False)
    ws[f'F{r_minp1}'].protection = Protection(locked=False)
    ws[f'B{r_minp2}'].protection = Protection(locked=False)
    ws[f'B{r_minp3}'].protection = Protection(locked=False)
    ws[f'D{r_minp3}'].protection = Protection(locked=False)
    
    # Section 7 Library rows
    for r in range(r_llibhead + 1, r_llibhead + 13):
        for c in range(1, 9):
            ws.cell(row=r, column=c).protection = Protection(locked=False)
            
    ws.protection.sheet = True
    ws.freeze_panes = 'A4'
    print('Built Dual-Engine Carriage Sheet (1.1 Mechanical & 1.2 Manual Labour).')
