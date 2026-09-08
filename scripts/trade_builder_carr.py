# -*- coding: utf-8 -*-
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.worksheet.datavalidation import DataValidation
from scripts.trade_builder import add_trade_header_and_legend

def build_carriage_trade(wb, config, styles):
    ws = wb.create_sheet(title=config['sheet_name'])
    ws.views.sheetView[0].showGridLines = True
    add_trade_header_and_legend(ws, config, styles)
    
    # Item Metadata (Rows 5-6)
    ws['A5'] = 'Custom Item Code:'
    ws['A5'].font = styles['font_bold']
    ws['B5'] = config['default_item_code']
    ws['B5'].fill = styles['fill_input']
    ws['B5'].font = styles['font_bold']
    ws['B5'].alignment = styles['align_center']
    
    ws['C5'] = 'Batch Output Basis:'
    ws['C5'].font = styles['font_bold']
    ws['D5'] = config['default_basis_qty']
    ws['D5'].fill = styles['fill_input']
    ws['D5'].font = styles['font_bold']
    ws['D5'].alignment = styles['align_center']
    ws['D5'].number_format = '0.00'
    
    ws['E5'] = 'Output Unit:'
    ws['E5'].font = styles['font_bold']
    ws['F5'] = config['default_basis_unit']
    ws['F5'].fill = styles['fill_input']
    ws['F5'].font = styles['font_bold']
    ws['F5'].alignment = styles['align_center']
    
    ws['G5'] = 'Trade Sub-Head:'
    ws['G5'].font = styles['font_bold']
    ws['H5'] = '01 Carriage of Materials'
    ws['H5'].font = styles['font_bold']
    ws['H5'].fill = styles['fill_lookup']
    ws['H5'].alignment = styles['align_center']
    
    for c in ['A5', 'B5', 'C5', 'D5', 'E5', 'F5', 'G5', 'H5']:
        ws[c].border = styles['border_thin']
    ws.row_dimensions[5].height = 22
    
    ws['A6'] = 'Item Nomenclature / Full Description:'
    ws['A6'].font = styles['font_bold']
    ws['A6'].border = styles['border_thin']
    
    ws.merge_cells('B6:H6')
    ws['B6'] = config['default_item_desc']
    ws['B6'].font = styles['font_regular']
    ws['B6'].fill = styles['fill_input']
    ws['B6'].alignment = styles['align_wrap']
    ws['B6'].border = styles['border_thin']
    ws.row_dimensions[6].height = 36
    
    dv_yesno = DataValidation(type='list', formula1='"YES,NO"', allow_blank=False)
    ws.add_data_validation(dv_yesno)
    
    
    # REAL-TIME IN-SHEET AUDIT PANEL (Row 7)
    ws['A7'] = 'AUDIT STATUS:'
    ws['A7'].font = styles['font_white_bold']
    ws['A7'].fill = styles['fill_header']
    ws['A7'].alignment = styles['align_center']
    
    ws['B7'] = '=IF(AND(D7="OK", F7="OK"), "[PASS] ALL CHECKS OK", "[ALERT] CHECKS FAILED")'
    ws['B7'].font = styles['font_result']
    ws['B7'].fill = styles['fill_result']
    ws['B7'].alignment = styles['align_center']
    
    ws['C7'] = 'Output Qty Check:'
    ws['C7'].font = styles['font_note']
    ws['C7'].alignment = styles['align_right']
    
    ws['D7'] = '=IF(D5>0, "OK", "ERR: Qty<=0")'
    ws['D7'].font = styles['font_bold']
    ws['D7'].alignment = styles['align_center']
    
    ws['E7'] = 'Trip Cost Check:'
    ws['E7'].font = styles['font_note']
    ws['E7'].alignment = styles['align_right']
    
    ws['F7'] = '=IF(G22>0, "OK", "ERR: Trip Cost<=0")'
    ws['F7'].font = styles['font_bold']
    ws['F7'].alignment = styles['align_center']
    
    ws['G7'] = 'CPOH Applied:'
    ws['G7'].font = styles['font_note']
    ws['G7'].alignment = styles['align_right']
    
    ws['H7'] = '=IF(C27="YES", "CPOH 15% ACTIVE", "CPOH OFF")'
    ws['H7'].font = styles['font_bold']
    ws['H7'].alignment = styles['align_center']
    
    for c in ['A7', 'B7', 'C7', 'D7', 'E7', 'F7', 'G7', 'H7']:
        ws[c].border = styles['border_thin']
    ws.row_dimensions[7].height = 22
    
    # SECTION 1: MECHANICAL TRANSPORT BUILDER (Rows 8 to 20)
    ws.merge_cells('A8:H8')
    c_sec1 = ws['A8']
    c_sec1.value = '1. MECHANICAL TRANSPORT TRIP BUILDER (Vehicle Hire + Fuel + Loading/Unloading Labour)'
    c_sec1.font = styles['font_white_bold']
    c_sec1.fill = styles['fill_header']
    c_sec1.alignment = styles['align_left']
    ws.row_dimensions[8].height = 24
    
    tbl_headers = ['Line', 'Code / Source', 'Resource / Fuel Description', 'Unit', 'Quantity / Coeff', 'Basic Rate (₹)', 'Amount (₹)', 'Source Reference / Remarks']
    for c_idx, h in enumerate(tbl_headers, 1):
        cell = ws.cell(row=9, column=c_idx, value=h)
        cell.font = styles['font_header']
        cell.fill = styles['fill_header']
        cell.alignment = styles['align_center']
        cell.border = styles['border_header']
    ws.row_dimensions[9].height = 24
    
    mech_items = [
        {'code': '0084', 'coeff': 1.0, 'note': 'Hire charges of Diesel Truck - 9 tonne excluding diesel & mobile oil'},
        {'code': '0114', 'coeff': 6.0, 'note': 'Beldar (loading & unloading labour)'},
        {'code': '1235', 'coeff': 12.88, 'note': 'Diesel oil for mechanical transport trip'},
        {'code': '5001', 'coeff': 0.46, 'note': 'Mobil oil'}
    ]
    
    for idx in range(8):
        r = 10 + idx
        item = mech_items[idx] if idx < len(mech_items) else None
        
        ws.cell(row=r, column=1, value=idx+1).alignment = styles['align_center']
        
        c_code = ws.cell(row=r, column=2, value=item['code'] if item else '')
        c_code.alignment = styles['align_center']
        c_code.font = styles['font_bold']
        c_code.fill = styles['fill_input']
        
        c_desc = ws.cell(row=r, column=3, value=f'=IF(B{r}="","", IFERROR(INDEX(Rates_Master!$C:$C, MATCH(B{r}, Rates_Master!$A:$A, 0)), "Custom Transport Line"))')
        c_desc.alignment = styles['align_left']
        c_desc.fill = styles['fill_lookup']
        
        c_unit = ws.cell(row=r, column=4, value=f'=IF(B{r}="","", IFERROR(INDEX(Rates_Master!$D:$D, MATCH(B{r}, Rates_Master!$A:$A, 0)), ""))')
        c_unit.alignment = styles['align_center']
        c_unit.fill = styles['fill_lookup']
        
        c_qty = ws.cell(row=r, column=5, value=item['coeff'] if item else None)
        c_qty.alignment = styles['align_right']
        c_qty.number_format = styles['fmt_qty']
        c_qty.fill = styles['fill_input']
        
        c_rate = ws.cell(row=r, column=6, value=f'=IF(B{r}="","", IFERROR(INDEX(Rates_Master!$E:$E, MATCH(B{r}, Rates_Master!$A:$A, 0)), 0))')
        c_rate.alignment = styles['align_right']
        c_rate.number_format = styles['fmt_currency']
        c_rate.fill = styles['fill_lookup']
        
        c_amt = ws.cell(row=r, column=7, value=f'=IF(OR(B{r}="", E{r}=""), 0, ROUND(E{r} * F{r}, 2))')
        c_amt.alignment = styles['align_right']
        c_amt.number_format = styles['fmt_currency']
        c_amt.fill = styles['fill_calc']
        
        c_rem = ws.cell(row=r, column=8, value=item['note'] if item else '')
        c_rem.alignment = styles['align_left']
        c_rem.font = styles['font_note']
        
        for col in range(1, 9):
            ws.cell(row=r, column=col).border = styles['border_thin']
        ws.row_dimensions[r].height = 20
        
    # Subtotal Direct Trip Cost (Row 18)
    ws.merge_cells('A18:E18')
    ws['A18'] = 'Direct Trip Operating Cost (W) (₹):'
    ws['A18'].font = styles['font_bold']
    ws['A18'].alignment = styles['align_right']
    ws['A18'].fill = styles['fill_subtotal']
    
    c_dirsum = ws['G18']
    c_dirsum.value = '=SUM(G10:G17)'
    c_dirsum.font = styles['font_bold']
    c_dirsum.alignment = styles['align_right']
    c_dirsum.number_format = styles['fmt_currency']
    c_dirsum.fill = styles['fill_subtotal']
    for c in ['A18', 'B18', 'C18', 'D18', 'E18', 'F18', 'G18', 'H18']:
        ws[c].border = styles['border_double_bottom']
    ws.row_dimensions[18].height = 22
    
    # SECTION 2: STATUTORY MARKUP (CPOH ONLY PER CPWD DAR) (Rows 20 to 32)
    ws.merge_cells('A20:H20')
    c_sec2 = ws['A20']
    c_sec2.value = '2. STATUTORY MARKUPS (CPWD CARRIAGE CONVENTION: 15% CPOH ONLY; WATER/GST/CESS OFF)'
    c_sec2.font = styles['font_white_bold']
    c_sec2.fill = styles['fill_header']
    c_sec2.alignment = styles['align_left']
    ws.row_dimensions[20].height = 24
    
    for c_idx, h in enumerate(tbl_headers[:2] + ['Apply? (YES/NO)', 'Basis Applied', 'Base Amount (₹)', 'Factor / %', 'Amount (₹)', 'CPWD Statutory Rule & Guidance Note'], 1):
        cell = ws.cell(row=21, column=c_idx, value=h)
        cell.font = styles['font_header']
        cell.fill = styles['fill_header']
        cell.alignment = styles['align_center']
        cell.border = styles['border_header']
    ws.row_dimensions[21].height = 24
    
    # Row 22: Base Direct Cost (W)
    ws['A22'] = 'W'
    ws['B22'] = 'Direct Trip Operating Cost (Vehicle + Labour + Fuel)'
    ws['C22'] = '—'
    ws['D22'] = 'Direct Sum'
    ws['E22'] = '—'
    ws['F22'] = '—'
    ws['G22'] = '=G18'
    ws['H22'] = 'Total direct operating trip cost before overheads'
    
    # Row 23: Water Charges (NO by default)
    ws['A23'] = 'X1'
    ws['B23'] = 'Add Water Charges (1% on W)'
    ws['C23'] = 'NO'
    ws['D23'] = 'On "W"'
    ws['E23'] = '=G22'
    ws['F23'] = '=Factor_Water'
    ws['G23'] = '=IF(C23="YES", ROUND(E23 * F23, 2), 0)'
    ws['H23'] = 'CPWD RULE: Omitted (NO) in Carriage items because no water is consumed in transport.'
    
    # Row 24: Subtotal (X)
    ws['A24'] = 'X'
    ws['B24'] = 'Subtotal "X" (W + Water Charges)'
    ws['C24'] = '—'
    ws['D24'] = 'W + Water'
    ws['E24'] = '—'
    ws['F24'] = '—'
    ws['G24'] = '=G22 + G23'
    ws['H24'] = 'Base for GST (if applicable)'
    
    # Row 25: GST (NO by default)
    ws['A25'] = 'Y1'
    ws['B25'] = 'Add GST on Transport'
    ws['C25'] = 'NO'
    ws['D25'] = 'On "X"'
    ws['E25'] = '=G24'
    ws['F25'] = '=Factor_GST'
    ws['G25'] = '=IF(C25="YES", ROUND(E25 * F25, 2), 0)'
    ws['H25'] = 'CPWD RULE: Omitted (NO) in DAR Carriage items; tax on vehicle freight is handled separately.'
    
    # Row 26: Subtotal (Y)
    ws['A26'] = 'Y'
    ws['B26'] = 'Subtotal "Y" (X + GST)'
    ws['C26'] = '—'
    ws['D26'] = 'X + GST'
    ws['E26'] = '—'
    ws['F26'] = '—'
    ws['G26'] = '=G24 + G25'
    ws['H26'] = 'Base for Contractor Profit & Overheads'
    
    # Row 27: CPOH (YES by default, 15%)
    ws['A27'] = 'Z1'
    ws['B27'] = 'Add Contractor Profit & Overheads (15% CPOH)'
    ws['C27'] = 'YES'
    ws['D27'] = 'On "Y"'
    ws['E27'] = '=G26'
    ws['F27'] = '=Factor_CPOH'
    ws['G27'] = '=IF(C27="YES", ROUND(E27 * F27, 2), 0)'
    ws['H27'] = 'CPWD RULE: 15% CPOH is added on Carriage direct cost per DAR 2019 item 1.1.'
    
    # Row 28: Subtotal (Z)
    ws['A28'] = 'Z'
    ws['B28'] = 'Subtotal "Z" (Y + CPOH)'
    ws['C28'] = '—'
    ws['D28'] = 'Y + CPOH'
    ws['E28'] = '—'
    ws['F28'] = '—'
    ws['G28'] = '=G26 + G27'
    ws['H28'] = 'Base for Labour Cess'
    
    # Row 29: BOCW Cess (NO by default)
    ws['A29'] = 'Z2'
    ws['B29'] = 'Add Labour Welfare Cess (1% BOCW Cess)'
    ws['C29'] = 'NO'
    ws['D29'] = 'On "Z"'
    ws['E29'] = '=G28'
    ws['F29'] = '=Factor_Cess'
    ws['G29'] = '=IF(C29="YES", ROUND(E29 * F29, 2), 0)'
    ws['H29'] = 'CPWD RULE: Omitted (NO) in standard Carriage items. Toggle YES if contract requires cess on haulage.'
    
    # Row 30: Cost of Specified Batch
    ws['A30'] = 'Cost'
    ws['B30'] = '="Total Carriage Cost for " & TEXT(D5, "0.00") & " " & F5 & ":"'
    ws['C30'] = '—'
    ws['D30'] = 'Z + Cess'
    ws['E30'] = '—'
    ws['F30'] = '—'
    ws['G30'] = '=G28 + G29'
    ws['H30'] = 'Total evaluated carriage cost for batch specified in Row 5'
    
    # Row 31: Analyzed Unit Rate
    ws['A31'] = 'Rate'
    ws['B31'] = '="Analyzed Carriage Rate per 1.00 " & F5 & ":"'
    ws['C31'] = '—'
    ws['D31'] = 'Cost / Batch Qty'
    ws['E31'] = '—'
    ws['F31'] = '—'
    ws['G31'] = '=ROUND(G30 / D5, 2)'
    ws['H31'] = 'Derived carriage rate per cubic metre (or unit)'
    
    # Row 32: SAY RATE
    ws['A32'] = 'SAY'
    ws['B32'] = '="OFFICIAL SAY RATE (₹ per " & F5 & "):"'
    ws['C32'] = '—'
    ws['D32'] = 'Final Say'
    ws['E32'] = '—'
    ws['F32'] = '—'
    ws['G32'] = '=ROUND(G31, 2)'
    ws['H32'] = 'OFFICIAL CPWD ROUNDED RATE FOR TENDER SCHEDULES & CARRIAGE SPECIFICATIONS'
    
    for r in range(22, 33):
        ws.cell(row=r, column=1).alignment = styles['align_center']
        ws.cell(row=r, column=1).font = styles['font_bold']
        ws.cell(row=r, column=2).alignment = styles['align_left']
        ws.cell(row=r, column=3).alignment = styles['align_center']
        ws.cell(row=r, column=4).alignment = styles['align_center']
        
        c_base = ws.cell(row=r, column=5)
        c_base.alignment = styles['align_right']
        if r in [23, 25, 27, 29]:
            c_base.number_format = styles['fmt_currency']
            
        c_fac = ws.cell(row=r, column=6)
        c_fac.alignment = styles['align_right']
        if r in [23, 25, 27, 29]:
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
            
        if r in [22, 24, 26, 28]:
            for c in range(1, 8):
                ws.cell(row=r, column=c).fill = styles['fill_subtotal']
        elif r in [23, 25, 27, 29]:
            ws.cell(row=r, column=3).fill = styles['fill_input']
            ws.cell(row=r, column=3).font = styles['font_bold']
            dv_yesno.add(ws.cell(row=r, column=3))
        elif r == 30:
            for c in range(1, 8):
                ws.cell(row=r, column=c).fill = styles['fill_result']
        elif r == 31:
            for c in range(1, 8):
                ws.cell(row=r, column=c).fill = styles['fill_subtotal']
        elif r == 32:
            ws.row_dimensions[32].height = 30
            for c in range(1, 8):
                ws.cell(row=r, column=c).fill = styles['fill_say']
            ws['B32'].font = styles['font_say']
            ws['G32'].font = styles['font_say']
            
        ws.row_dimensions[r].height = 24 if r != 32 else 30
        
    # SECTION 3: MANUAL LABOUR LEAD REFERENCE MATRIX (Rows 35 to 48)
    ws.merge_cells('A35:H35')
    c_sec3 = ws['A35']
    c_sec3.value = '3. MANUAL LABOUR LEAD REFERENCE MATRIX (CPWD DAR 2019 Item 1.2 Standards)'
    c_sec3.font = styles['font_white_bold']
    c_sec3.fill = styles['fill_header']
    c_sec3.alignment = styles['align_left']
    ws.row_dimensions[35].height = 24
    
    lead_headers = ['Lead Distance', 'Beldars / Day', 'Coolies / Day', 'Labour Cost (₹)', 'Add 15% CPOH (₹)', 'Total Rate / Day', 'Lead Range Basis', 'CPWD DAR Norms & Labour Build-up']
    for c_idx, h in enumerate(lead_headers, 1):
        cell = ws.cell(row=36, column=c_idx, value=h)
        cell.font = styles['font_header']
        cell.fill = styles['fill_header']
        cell.alignment = styles['align_center']
        cell.border = styles['border_header']
    ws.row_dimensions[36].height = 24
    
    lead_matrix = [
        ('1st 50 Metres (Base)', 7.67, 0.00, '=ROUND(B37*558, 2)', '=ROUND(D37*0.15, 2)', '=D37+E37', 'Lead upto 50 m', 'CPWD Norm: 7.67 Beldars @ ₹558/day works 8 hours carrying standard load'),
        ('Addl. 50 m (upto 100 m)', 0.00, 1.67, '=ROUND(C38*558, 2)', '=ROUND(D38*0.15, 2)', '=D38+E38', 'Lead 50 m to 100 m', '1.67 Coolies added per additional 50 m lead increment'),
        ('Addl. 50 m (upto 150 m)', 0.00, 1.67, '=ROUND(C39*558, 2)', '=ROUND(D39*0.15, 2)', '=D39+E39', 'Lead 100 m to 150 m', '1.67 Coolies added per additional 50 m lead increment'),
        ('Addl. 50 m (upto 200 m)', 0.00, 1.67, '=ROUND(C40*558, 2)', '=ROUND(D40*0.15, 2)', '=D40+E40', 'Lead 150 m to 200 m', '1.67 Coolies added per additional 50 m lead increment'),
        ('Addl. 50 m (upto 250 m)', 0.00, 1.67, '=ROUND(C41*558, 2)', '=ROUND(D41*0.15, 2)', '=D41+E41', 'Lead 200 m to 250 m', '1.67 Coolies added per additional 50 m lead increment'),
        ('Addl. 50 m (upto 300 m)', 0.00, 1.67, '=ROUND(C42*558, 2)', '=ROUND(D42*0.15, 2)', '=D42+E42', 'Lead 250 m to 300 m', '1.67 Coolies added per additional 50 m lead increment'),
        ('Addl. 50 m (upto 350 m)', 0.00, 1.67, '=ROUND(C43*558, 2)', '=ROUND(D43*0.15, 2)', '=D43+E43', 'Lead 300 m to 350 m', '1.67 Coolies added per additional 50 m lead increment'),
        ('Addl. 50 m (upto 400 m)', 0.00, 1.67, '=ROUND(C44*558, 2)', '=ROUND(D44*0.15, 2)', '=D44+E44', 'Lead 350 m to 400 m', '1.67 Coolies added per additional 50 m lead increment'),
        ('Addl. 50 m (upto 450 m)', 0.00, 1.67, '=ROUND(C45*558, 2)', '=ROUND(D45*0.15, 2)', '=D45+E45', 'Lead 400 m to 450 m', '1.67 Coolies added per additional 50 m lead increment'),
        ('Addl. 50 m (upto 500 m)', 0.00, 1.67, '=ROUND(C46*558, 2)', '=ROUND(D46*0.15, 2)', '=D46+E46', 'Lead 450 m to 500 m', '1.67 Coolies added per additional 50 m lead increment')
    ]
    
    for idx, row_data in enumerate(lead_matrix):
        r = 37 + idx
        ws.cell(row=r, column=1, value=row_data[0]).alignment = styles['align_left']
        ws.cell(row=r, column=1).font = styles['font_bold']
        
        c_beld = ws.cell(row=r, column=2, value=row_data[1])
        c_beld.alignment = styles['align_right']
        c_beld.number_format = '0.00'
        
        c_cool = ws.cell(row=r, column=3, value=row_data[2])
        c_cool.alignment = styles['align_right']
        c_cool.number_format = '0.00'
        
        c_lcost = ws.cell(row=r, column=4, value=row_data[3])
        c_lcost.alignment = styles['align_right']
        c_lcost.number_format = styles['fmt_currency']
        
        c_cpoh = ws.cell(row=r, column=5, value=row_data[4])
        c_cpoh.alignment = styles['align_right']
        c_cpoh.number_format = styles['fmt_currency']
        
        c_tot = ws.cell(row=r, column=6, value=row_data[5])
        c_tot.alignment = styles['align_right']
        c_tot.number_format = styles['fmt_currency']
        c_tot.font = styles['font_bold']
        
        ws.cell(row=r, column=7, value=row_data[6]).alignment = styles['align_center']
        
        c_rem = ws.cell(row=r, column=8, value=row_data[7])
        c_rem.alignment = styles['align_wrap']
        c_rem.font = styles['font_note']
        
        row_fill = styles['fill_subtotal'] if r % 2 == 0 else styles['fill_calc']
        for c in range(1, 9):
            cell = ws.cell(row=r, column=c)
            cell.fill = row_fill
            cell.border = styles['border_thin']
            if c not in [1, 6]:
                cell.font = styles['font_regular']
        ws.row_dimensions[r].height = 20
        
    # SECTION 4: RUNNING CUSTOM CARRIAGE ITEMS LIBRARY (Rows 49+)
    ws.merge_cells('A49:H49')
    c_sec4 = ws['A49']
    c_sec4.value = '4. RUNNING CUSTOM NON-DSR ITEMS LIBRARY FOR CARRIAGE'
    c_sec4.font = styles['font_white_bold']
    c_sec4.fill = styles['fill_header']
    c_sec4.alignment = styles['align_left']
    ws.row_dimensions[49].height = 24
    
    ws.merge_cells('A50:H50')
    c_inst = ws['A50']
    c_inst.value = 'Log all completed custom carriage rate analyses below.'
    c_inst.font = styles['font_note']
    c_inst.fill = styles['fill_note']
    ws.row_dimensions[50].height = 20
    
    lib_headers = ['Item Code', 'Item Nomenclature / Specification', 'Unit', 'Output Basis', 'Direct Cost W (₹)', 'Markups Applied', 'Unit Rate (₹)', 'Say Rate (₹)']
    for c_idx, h in enumerate(lib_headers, 1):
        cell = ws.cell(row=51, column=c_idx, value=h)
        cell.font = styles['font_header']
        cell.fill = styles['fill_header']
        cell.alignment = styles['align_center']
        cell.border = styles['border_header']
    ws.row_dimensions[51].height = 24
    
    sample_lib = config.get('sample_library', [])
    for idx in range(8):
        r = 52 + idx
        item = sample_lib[idx] if idx < len(sample_lib) else None
        
        ws.cell(row=r, column=1, value=item['code'] if item else f"C-01.0{idx+1}").alignment = styles['align_center']
        ws.cell(row=r, column=2, value=item['desc'] if item else '(Available for new custom carriage item)').alignment = styles['align_left']
        ws.cell(row=r, column=3, value=item['unit'] if item else config['default_basis_unit']).alignment = styles['align_center']
        ws.cell(row=r, column=4, value=item['basis'] if item else config['default_basis_qty']).alignment = styles['align_center']
        
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
        
    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 24
    ws.column_dimensions['C'].width = 48
    ws.column_dimensions['D'].width = 16
    ws.column_dimensions['E'].width = 18
    ws.column_dimensions['F'].width = 18
    ws.column_dimensions['G'].width = 20
    ws.column_dimensions['H'].width = 48
    
    
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            ws.cell(row=r, column=c).protection = Protection(locked=True)
            
    ws['B5'].protection = Protection(locked=False)
    ws['D5'].protection = Protection(locked=False)
    ws['F5'].protection = Protection(locked=False)
    ws['B6'].protection = Protection(locked=False)
    ws['C23'].protection = Protection(locked=False)
    ws['C25'].protection = Protection(locked=False)
    ws['C27'].protection = Protection(locked=False)
    ws['C29'].protection = Protection(locked=False)
    
    for r in range(10, 18):
        ws.cell(row=r, column=2).protection = Protection(locked=False)
        ws.cell(row=r, column=5).protection = Protection(locked=False)
        ws.cell(row=r, column=8).protection = Protection(locked=False)
    for r in range(52, 60):
        for c in range(1, 9):
            ws.cell(row=r, column=c).protection = Protection(locked=False)
            
    ws.protection.sheet = True
    ws.freeze_panes = 'A4'
    print('Built Carriage trade sheet.')

