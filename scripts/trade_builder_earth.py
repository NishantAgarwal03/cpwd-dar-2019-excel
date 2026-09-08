# -*- coding: utf-8 -*-
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.worksheet.datavalidation import DataValidation
from scripts.trade_builder import add_trade_header_and_legend

def build_earthwork_trade(wb, config, styles):
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
    ws['H5'] = '02 Earth Work'
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
    
    ws['B7'] = '=IF(AND(D7="OK", F7="OK", H7="OK"), "[PASS] ALL CHECKS OK", "[ALERT] CHECKS FAILED")'
    ws['B7'].font = styles['font_result']
    ws['B7'].fill = styles['fill_result']
    ws['B7'].alignment = styles['align_center']
    
    ws['C7'] = 'Output Qty Check:'
    ws['C7'].font = styles['font_note']
    ws['C7'].alignment = styles['align_right']
    
    ws['D7'] = '=IF(D5>0, "OK", "ERR: Qty<=0")'
    ws['D7'].font = styles['font_bold']
    ws['D7'].alignment = styles['align_center']
    
    ws['E7'] = 'Direct Cost Check:'
    ws['E7'].font = styles['font_note']
    ws['E7'].alignment = styles['align_right']
    
    ws['F7'] = '=IF(G27>0, "OK", "ERR: W<=0")'
    ws['F7'].font = styles['font_bold']
    ws['F7'].alignment = styles['align_center']
    
    ws['G7'] = 'Markup Flow Check:'
    ws['G7'].font = styles['font_note']
    ws['G7'].alignment = styles['align_right']
    
    ws['H7'] = '=IF(AND(G27<=G29, G29<=G31, G31<=G33), "OK", "ERR: Markups Broken")'
    ws['H7'].font = styles['font_bold']
    ws['H7'].alignment = styles['align_center']
    
    for c in ['A7', 'B7', 'C7', 'D7', 'E7', 'F7', 'G7', 'H7']:
        ws[c].border = styles['border_thin']
    ws.row_dimensions[7].height = 22
    
    # SECTION 1: LABOUR & MACHINERY (Rows 8 to 20)
    ws.merge_cells('A8:H8')
    c_sec1 = ws['A8']
    c_sec1.value = '1. LABOUR & MACHINERY COMPONENT BUILD-UP (Note: Equipment codes 0001-0083 include driver, fuel & lubricants per 8-hr shift)'
    c_sec1.font = styles['font_white_bold']
    c_sec1.fill = styles['fill_header']
    c_sec1.alignment = styles['align_left']
    ws.row_dimensions[8].height = 24
    
    tbl_headers = ['Line', 'Code / Source', 'Labour / Equipment Description', 'Unit', 'Day Coeff / Qty', 'Basic Rate (₹)', 'Amount (₹)', 'Source Reference / Remarks']
    for c_idx, h in enumerate(tbl_headers, 1):
        cell = ws.cell(row=9, column=c_idx, value=h)
        cell.font = styles['font_header']
        cell.fill = styles['fill_header']
        cell.alignment = styles['align_center']
        cell.border = styles['border_header']
    ws.row_dimensions[9].height = 24
    
    labour = config.get('default_labour', [])
    for idx in range(10):
        r = 10 + idx
        lab = labour[idx] if idx < len(labour) else None
        
        ws.cell(row=r, column=1, value=idx+1).alignment = styles['align_center']
        
        c_code = ws.cell(row=r, column=2, value=lab['code'] if lab else '')
        c_code.alignment = styles['align_center']
        c_code.font = styles['font_bold']
        c_code.fill = styles['fill_input']
        
        c_desc = ws.cell(row=r, column=3, value=f'=IF(B{r}="","", IFERROR(INDEX(Rates_Master!$C:$C, MATCH(B{r}, Rates_Master!$A:$A, 0)), "Custom Resource"))')
        c_desc.alignment = styles['align_left']
        c_desc.fill = styles['fill_lookup']
        
        c_unit = ws.cell(row=r, column=4, value=f'=IF(B{r}="","", IFERROR(INDEX(Rates_Master!$D:$D, MATCH(B{r}, Rates_Master!$A:$A, 0)), ""))')
        c_unit.alignment = styles['align_center']
        c_unit.fill = styles['fill_lookup']
        
        c_qty = ws.cell(row=r, column=5, value=lab['coeff'] if lab else None)
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
        
        c_rem = ws.cell(row=r, column=8, value=lab['note'] if lab else '')
        c_rem.alignment = styles['align_left']
        c_rem.font = styles['font_note']
        
        for col in range(1, 9):
            ws.cell(row=r, column=col).border = styles['border_thin']
        ws.row_dimensions[r].height = 20
        
    # Subtotal Labour (Row 20)
    ws.merge_cells('A20:E20')
    ws['A20'] = 'Subtotal Labour & Plant Hire (₹):'
    ws['A20'].font = styles['font_bold']
    ws['A20'].alignment = styles['align_right']
    ws['A20'].fill = styles['fill_subtotal']
    
    c_labsum = ws['G20']
    c_labsum.value = '=SUM(G10:G19)'
    c_labsum.font = styles['font_bold']
    c_labsum.alignment = styles['align_right']
    c_labsum.number_format = styles['fmt_currency']
    c_labsum.fill = styles['fill_subtotal']
    for c in ['A20', 'B20', 'C20', 'D20', 'E20', 'F20', 'G20', 'H20']:
        ws[c].border = styles['border_double_bottom']
    ws.row_dimensions[20].height = 22
    
    # SECTION 2: SUNDRIES (Rows 22 to 23)
    ws.merge_cells('A22:H22')
    c_sec2 = ws['A22']
    c_sec2.value = '2. SUNDRIES & LUMP-SUM ALLOWANCES'
    c_sec2.font = styles['font_white_bold']
    c_sec2.fill = styles['fill_header']
    c_sec2.alignment = styles['align_left']
    ws.row_dimensions[22].height = 24
    
    ws['A23'] = '1'
    ws['A23'].alignment = styles['align_center']
    ws['B23'] = '9999'
    ws['B23'].alignment = styles['align_center']
    ws['B23'].font = styles['font_bold']
    ws['C23'] = 'Sundries (guided by Sundries_Reference sheet)'
    ws['D23'] = 'L.S.'
    ws['D23'].alignment = styles['align_center']
    
    ws['E23'] = config.get('default_sundries_base', 0.0)
    ws['E23'].alignment = styles['align_right']
    ws['E23'].number_format = '0.00'
    ws['E23'].fill = styles['fill_input']
    
    ws['F23'] = '=Factor_Sundries'
    ws['F23'].alignment = styles['align_right']
    ws['F23'].number_format = '0.00'
    ws['F23'].fill = styles['fill_lookup']
    
    ws['G23'] = '=ROUND(E23 * F23, 2)'
    ws['G23'].alignment = styles['align_right']
    ws['G23'].number_format = styles['fmt_currency']
    ws['G23'].fill = styles['fill_calc']
    ws['G23'].font = styles['font_bold']
    
    ws['H23'] = 'Cost Index Multiplier (2.00) applied on base allowance'
    ws['H23'].font = styles['font_note']
    for c in ['A23', 'B23', 'C23', 'D23', 'E23', 'F23', 'G23', 'H23']:
        ws[c].border = styles['border_thin']
    ws.row_dimensions[23].height = 22
    
    # SECTION 3: STATUTORY MARKUP CHAIN (Rows 25 to 37)
    ws.merge_cells('A25:H25')
    c_sec3 = ws['A25']
    c_sec3.value = '3. STATUTORY MARKUPS & FINAL RATE DERIVATION (CPWD DAR METHOD)'
    c_sec3.font = styles['font_white_bold']
    c_sec3.fill = styles['fill_header']
    c_sec3.alignment = styles['align_left']
    ws.row_dimensions[25].height = 24
    
    mu_headers = ['Step', 'Component / Markup Description', 'Apply? (YES/NO)', 'Basis Applied', 'Base Amount (₹)', 'Factor / %', 'Amount (₹)', 'CPWD Statutory Rule & Guidance Note']
    for c_idx, h in enumerate(mu_headers, 1):
        cell = ws.cell(row=26, column=c_idx, value=h)
        cell.font = styles['font_header']
        cell.fill = styles['fill_header']
        cell.alignment = styles['align_center']
        cell.border = styles['border_header']
    ws.row_dimensions[26].height = 24
    
    toggles = config.get('toggles', {'water': 'YES', 'gst': 'YES', 'cpoh': 'YES', 'cess': 'YES'})
    notes = config.get('toggle_notes', {})
    
    # Row 27: Base Direct Cost (W)
    ws['A27'] = 'W'
    ws['B27'] = 'Base Direct Production Cost (Labour + Plant + Sundries)'
    ws['C27'] = '—'
    ws['D27'] = 'Direct Sum'
    ws['E27'] = '—'
    ws['F27'] = '—'
    ws['G27'] = '=G20 + G23'
    ws['H27'] = 'Total direct cost of production (W) before statutory overheads'
    
    # Row 28: Water Charges
    ws['A28'] = 'X1'
    ws['B28'] = 'Add Water Charges (1% on W)'
    ws['C28'] = toggles['water']
    ws['D28'] = 'On "W"'
    ws['E28'] = '=G27'
    ws['F28'] = '=Factor_Water'
    ws['G28'] = '=IF(C28="YES", ROUND(E28 * F28, 2), 0)'
    ws['H28'] = notes.get('water', 'CPWD Standard: 1% for site compaction water and dust suppression.')
    
    # Row 29: Subtotal (X)
    ws['A29'] = 'X'
    ws['B29'] = 'Subtotal "X" (W + Water Charges)'
    ws['C29'] = '—'
    ws['D29'] = 'W + Water'
    ws['E29'] = '—'
    ws['F29'] = '—'
    ws['G29'] = '=G27 + G28'
    ws['H29'] = 'Compounded base for GST calculation'
    
    # Row 30: GST
    ws['A30'] = 'Y1'
    ws['B30'] = 'Add GST on Works Contract'
    ws['C30'] = toggles['gst']
    ws['D30'] = 'On "X"'
    ws['E30'] = '=G29'
    ws['F30'] = '=Factor_GST'
    ws['G30'] = '=IF(C30="YES", ROUND(E30 * F30, 2), 0)'
    ws['H30'] = notes.get('gst', 'CPWD DAR factor 0.1405 (works contract tax).')
    
    # Row 31: Subtotal (Y)
    ws['A31'] = 'Y'
    ws['B31'] = 'Subtotal "Y" (X + GST)'
    ws['C31'] = '—'
    ws['D31'] = 'X + GST'
    ws['E31'] = '—'
    ws['F31'] = '—'
    ws['G31'] = '=G29 + G30'
    ws['H31'] = 'Compounded base for Contractor Profit & Overheads'
    
    # Row 32: CPOH
    ws['A32'] = 'Z1'
    ws['B32'] = 'Add Contractor Profit & Overheads (15% CPOH)'
    ws['C32'] = toggles['cpoh']
    ws['D32'] = 'On "Y"'
    ws['E32'] = '=G31'
    ws['F32'] = '=Factor_CPOH'
    ws['G32'] = '=IF(C32="YES", ROUND(E32 * F32, 2), 0)'
    ws['H32'] = notes.get('cpoh', 'Standard CPWD 15% allowance for site overheads, equipment depreciation & margin.')
    
    # Row 33: Subtotal (Z)
    ws['A33'] = 'Z'
    ws['B33'] = 'Subtotal "Z" (Y + CPOH)'
    ws['C33'] = '—'
    ws['D33'] = 'Y + CPOH'
    ws['E33'] = '—'
    ws['F33'] = '—'
    ws['G33'] = '=G31 + G32'
    ws['H33'] = 'Compounded base for Labour Welfare Cess'
    
    # Row 34: BOCW Cess
    ws['A34'] = 'Z2'
    ws['B34'] = 'Add Labour Welfare Cess (1% BOCW Cess)'
    ws['C34'] = toggles['cess']
    ws['D34'] = 'On "Z"'
    ws['E34'] = '=G33'
    ws['F34'] = '=Factor_Cess'
    ws['G34'] = '=IF(C34="YES", ROUND(E34 * F34, 2), 0)'
    ws['H34'] = notes.get('cess', 'Statutory 1% Building and Other Construction Workers Welfare Cess.')
    
    # Row 35: Cost of Specified Output
    ws['A35'] = 'Cost'
    ws['B35'] = '="Total Cost of Batch (" & TEXT(D5, "0.00") & " " & F5 & "):"'
    ws['C35'] = '—'
    ws['D35'] = 'Z + Cess'
    ws['E35'] = '—'
    ws['F35'] = '—'
    ws['G35'] = '=G33 + G34'
    ws['H35'] = 'Total cost for batch specified in row 5'
    
    # Row 36: Cost per 1.00 Unit
    ws['A36'] = 'Rate'
    ws['B36'] = '="Analyzed Unit Rate per 1.00 " & F5 & ":"'
    ws['C36'] = '—'
    ws['D36'] = 'Cost / Batch Qty'
    ws['E36'] = '—'
    ws['F36'] = '—'
    ws['G36'] = '=ROUND(G35 / D5, 2)'
    ws['H36'] = 'Base rate per single unit of measurement'
    
    # Row 37: OFFICIAL SAY RATE
    ws['A37'] = 'SAY'
    ws['B37'] = '="OFFICIAL SAY RATE (₹ per " & F5 & "):"'
    ws['C37'] = '—'
    ws['D37'] = 'Final Say'
    ws['E37'] = '—'
    ws['F37'] = '—'
    ws['G37'] = '=ROUND(G36, 2)'
    ws['H37'] = 'OFFICIAL CPWD ROUNDED RATE FOR TENDER SCHEDULES & ESTIMATES'
    
    for r in range(27, 38):
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
        elif r == 36:
            for c in range(1, 8):
                ws.cell(row=r, column=c).fill = styles['fill_subtotal']
        elif r == 37:
            ws.row_dimensions[37].height = 30
            for c in range(1, 8):
                ws.cell(row=r, column=c).fill = styles['fill_say']
            ws['B37'].font = styles['font_say']
            ws['G37'].font = styles['font_say']
            
        ws.row_dimensions[r].height = 24 if r != 37 else 30
        
    # SECTION 4: RUNNING CUSTOM ITEMS LIBRARY (Rows 40+)
    ws.merge_cells('A40:H40')
    c_sec4 = ws['A40']
    c_sec4.value = '4. RUNNING CUSTOM NON-DSR ITEMS LIBRARY FOR EARTH WORK'
    c_sec4.font = styles['font_white_bold']
    c_sec4.fill = styles['fill_header']
    c_sec4.alignment = styles['align_left']
    ws.row_dimensions[40].height = 24
    
    ws.merge_cells('A41:H41')
    c_inst = ws['A41']
    c_inst.value = 'Log all completed custom rate analyses for Earth Work below.'
    c_inst.font = styles['font_note']
    c_inst.fill = styles['fill_note']
    ws.row_dimensions[41].height = 20
    
    lib_headers = ['Item Code', 'Item Nomenclature / Specification', 'Unit', 'Output Basis', 'Direct Cost W (₹)', 'Markups Applied', 'Unit Rate (₹)', 'Say Rate (₹)']
    for c_idx, h in enumerate(lib_headers, 1):
        cell = ws.cell(row=42, column=c_idx, value=h)
        cell.font = styles['font_header']
        cell.fill = styles['fill_header']
        cell.alignment = styles['align_center']
        cell.border = styles['border_header']
    ws.row_dimensions[42].height = 24
    
    sample_lib = config.get('sample_library', [])
    for idx in range(8):
        r = 43 + idx
        item = sample_lib[idx] if idx < len(sample_lib) else None
        
        ws.cell(row=r, column=1, value=item['code'] if item else f"C-02.0{idx+1}").alignment = styles['align_center']
        ws.cell(row=r, column=2, value=item['desc'] if item else '(Available for new custom earthwork item)').alignment = styles['align_left']
        ws.cell(row=r, column=3, value=item['unit'] if item else config['default_basis_unit']).alignment = styles['align_center']
        ws.cell(row=r, column=4, value=item['basis'] if item else config['default_basis_qty']).alignment = styles['align_center']
        
        c_w = ws.cell(row=r, column=5, value=item['w'] if item else None)
        c_w.alignment = styles['align_right']
        c_w.number_format = styles['fmt_currency']
        
        ws.cell(row=r, column=6, value=item['markups'] if item else 'Full W->X->Y->Z').alignment = styles['align_center']
        
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
        
    ws.column_dimensions['A'].width = 10
    ws.column_dimensions['B'].width = 24
    ws.column_dimensions['C'].width = 50
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
    ws['E23'].protection = Protection(locked=False)
    ws['C28'].protection = Protection(locked=False)
    ws['C30'].protection = Protection(locked=False)
    ws['C32'].protection = Protection(locked=False)
    ws['C34'].protection = Protection(locked=False)
    
    for r in range(10, 20):
        ws.cell(row=r, column=2).protection = Protection(locked=False)
        ws.cell(row=r, column=5).protection = Protection(locked=False)
        ws.cell(row=r, column=8).protection = Protection(locked=False)
    for r in range(43, 51):
        for c in range(1, 9):
            ws.cell(row=r, column=c).protection = Protection(locked=False)
            
    ws.protection.sheet = True
    ws.freeze_panes = 'A4'
    print('Built Earth Work trade sheet.')

