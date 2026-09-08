# -*- coding: utf-8 -*-
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.worksheet.datavalidation import DataValidation

def add_trade_header_and_legend(ws, config, styles):
    # Title
    ws.merge_cells('A1:H1')
    c_title = ws['A1']
    c_title.value = config['trade_title']
    c_title.font = styles['font_title']
    c_title.fill = styles['fill_title']
    c_title.alignment = styles['align_center']
    ws.row_dimensions[1].height = 30
    
    # Legend
    ws['A2'] = 'LEGEND:'
    ws['A2'].font = styles['font_legend']
    ws['A2'].alignment = styles['align_center']
    
    legend_items = [
        ('B2', 'User Input (Editable)', styles['fill_input'], styles['font_bold']),
        ('C2', 'Auto Lookup / Reference', styles['fill_lookup'], styles['font_regular']),
        ('D2', 'Calculated Formula', styles['fill_calc'], styles['font_regular']),
        ('E2', 'Final Rate / Say', styles['fill_say'], styles['font_say']),
        ('F2', 'CPWD Rule / Note', styles['fill_note'], styles['font_note'])
    ]
    for cell_ref, text, fill, font in legend_items:
        c = ws[cell_ref]
        c.value = text
        c.fill = fill
        c.font = font
        c.border = styles['border_thin']
        c.alignment = styles['align_center']
    ws.row_dimensions[2].height = 20
    
    # Guidance Note
    ws.merge_cells('A3:H3')
    c_guide = ws['A3']
    c_guide.value = config['trade_guidance']
    c_guide.font = styles['font_note']
    c_guide.fill = styles['fill_note']
    c_guide.alignment = styles['align_wrap']
    c_guide.border = styles['border_thin']
    ws.row_dimensions[3].height = 26

def build_standard_trade(wb, config, styles):
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
    ws['H5'] = config['sheet_name'].replace('_', ' ')
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
    
    # Validation for YES/NO toggles
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
    
    ws['F7'] = '=IF(G41>0, "OK", "ERR: W<=0")'
    ws['F7'].font = styles['font_bold']
    ws['F7'].alignment = styles['align_center']
    
    ws['G7'] = 'Markup Flow Check:'
    ws['G7'].font = styles['font_note']
    ws['G7'].alignment = styles['align_right']
    
    ws['H7'] = '=IF(AND(G41<=G43, G43<=G45, G45<=G47), "OK", "ERR: Markups Broken")'
    ws['H7'].font = styles['font_bold']
    ws['H7'].alignment = styles['align_center']
    
    for c in ['A7', 'B7', 'C7', 'D7', 'E7', 'F7', 'G7', 'H7']:
        ws[c].border = styles['border_thin']
    ws.row_dimensions[7].height = 22
    
    # SECTION 1: MATERIALS (Rows 8 to 20)
    ws.merge_cells('A8:H8')
    c_sec1 = ws['A8']
    c_sec1.value = '1. MATERIAL COMPONENT BUILD-UP'
    c_sec1.font = styles['font_white_bold']
    c_sec1.fill = styles['fill_header']
    c_sec1.alignment = styles['align_left']
    ws.row_dimensions[8].height = 24
    
    tbl_headers = ['Line', 'Code / Source', 'Material Description', 'Unit', 'Quantity / Coeff', 'Basic Rate (₹)', 'Amount (₹)', 'Source Reference / Remarks']
    for c_idx, h in enumerate(tbl_headers, 1):
        cell = ws.cell(row=9, column=c_idx, value=h)
        cell.font = styles['font_header']
        cell.fill = styles['fill_header']
        cell.alignment = styles['align_center']
        cell.border = styles['border_header']
    ws.row_dimensions[9].height = 24
    
    materials = config.get('default_materials', [])
    for idx in range(10):
        r = 10 + idx
        mat = materials[idx] if idx < len(materials) else None
        
        ws.cell(row=r, column=1, value=idx+1).alignment = styles['align_center']
        
        c_code = ws.cell(row=r, column=2, value=mat['code'] if mat else '')
        c_code.alignment = styles['align_center']
        c_code.font = styles['font_bold']
        c_code.fill = styles['fill_input']
        
        c_desc = ws.cell(row=r, column=3)
        if mat and mat.get('custom_desc'):
            c_desc.value = mat['custom_desc']
        else:
            c_desc.value = f'=IF(B{r}="","", IFERROR(INDEX(Rates_Master!$C:$C, MATCH(B{r}, Rates_Master!$A:$A, 0)), "Custom / Intermediate Material"))'
        c_desc.alignment = styles['align_left']
        c_desc.fill = styles['fill_lookup']
        
        c_unit = ws.cell(row=r, column=4)
        if mat and mat.get('custom_unit'):
            c_unit.value = mat['custom_unit']
        else:
            c_unit.value = f'=IF(B{r}="","", IFERROR(INDEX(Rates_Master!$D:$D, MATCH(B{r}, Rates_Master!$A:$A, 0)), ""))'
        c_unit.alignment = styles['align_center']
        c_unit.fill = styles['fill_lookup']
        
        c_qty = ws.cell(row=r, column=5, value=mat['coeff'] if mat else None)
        c_qty.alignment = styles['align_right']
        c_qty.number_format = styles['fmt_qty']
        c_qty.fill = styles['fill_input']
        
        c_rate = ws.cell(row=r, column=6)
        if mat and mat.get('custom_rate_formula'):
            c_rate.value = mat['custom_rate_formula']
        elif mat and mat.get('custom_rate'):
            c_rate.value = mat['custom_rate']
        else:
            c_rate.value = f'=IF(B{r}="","", IFERROR(INDEX(Rates_Master!$E:$E, MATCH(B{r}, Rates_Master!$A:$A, 0)), 0))'
        c_rate.alignment = styles['align_right']
        c_rate.number_format = styles['fmt_currency']
        c_rate.fill = styles['fill_lookup']
        
        c_amt = ws.cell(row=r, column=7, value=f'=IF(OR(B{r}="", E{r}=""), 0, ROUND(E{r} * F{r}, 2))')
        c_amt.alignment = styles['align_right']
        c_amt.number_format = styles['fmt_currency']
        c_amt.fill = styles['fill_calc']
        
        c_rem = ws.cell(row=r, column=8, value=mat['note'] if mat else '')
        c_rem.alignment = styles['align_left']
        c_rem.font = styles['font_note']
        
        for col in range(1, 9):
            ws.cell(row=r, column=col).border = styles['border_thin']
        ws.row_dimensions[r].height = 20
        
    # Subtotal Materials (Row 20)
    ws.merge_cells('A20:E20')
    ws['A20'] = 'Subtotal Materials (₹):'
    ws['A20'].font = styles['font_bold']
    ws['A20'].alignment = styles['align_right']
    ws['A20'].fill = styles['fill_subtotal']
    
    c_matsum = ws['G20']
    c_matsum.value = '=SUM(G10:G19)'
    c_matsum.font = styles['font_bold']
    c_matsum.alignment = styles['align_right']
    c_matsum.number_format = styles['fmt_currency']
    c_matsum.fill = styles['fill_subtotal']
    for c in ['A20', 'B20', 'C20', 'D20', 'E20', 'F20', 'G20', 'H20']:
        ws[c].border = styles['border_double_bottom']
    ws.row_dimensions[20].height = 22
    
    # SECTION 2: LABOUR & MACHINERY (Rows 22 to 34)
    ws.merge_cells('A22:H22')
    c_sec2 = ws['A22']
    c_sec2.value = '2. LABOUR & MACHINERY COMPONENT BUILD-UP (Note: Plant codes 0001-0083 include operator, fuel & lubricants per 8-hr shift)'
    c_sec2.font = styles['font_white_bold']
    c_sec2.fill = styles['fill_header']
    c_sec2.alignment = styles['align_left']
    ws.row_dimensions[22].height = 24
    
    for c_idx, h in enumerate(tbl_headers, 1):
        cell = ws.cell(row=23, column=c_idx, value=h)
        cell.font = styles['font_header']
        cell.fill = styles['fill_header']
        cell.alignment = styles['align_center']
        cell.border = styles['border_header']
    ws.row_dimensions[23].height = 24
    
    labour = config.get('default_labour', [])
    for idx in range(10):
        r = 24 + idx
        lab = labour[idx] if idx < len(labour) else None
        
        ws.cell(row=r, column=1, value=idx+1).alignment = styles['align_center']
        
        c_code = ws.cell(row=r, column=2, value=lab['code'] if lab else '')
        c_code.alignment = styles['align_center']
        c_code.font = styles['font_bold']
        c_code.fill = styles['fill_input']
        
        c_desc = ws.cell(row=r, column=3, value=f'=IF(B{r}="","", IFERROR(INDEX(Rates_Master!$C:$C, MATCH(B{r}, Rates_Master!$A:$A, 0)), "Custom Labour/Plant"))')
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
        
    # Subtotal Labour (Row 34)
    ws.merge_cells('A34:E34')
    ws['A34'] = 'Subtotal Labour & Plant Hire (₹):'
    ws['A34'].font = styles['font_bold']
    ws['A34'].alignment = styles['align_right']
    ws['A34'].fill = styles['fill_subtotal']
    
    c_labsum = ws['G34']
    c_labsum.value = '=SUM(G24:G33)'
    c_labsum.font = styles['font_bold']
    c_labsum.alignment = styles['align_right']
    c_labsum.number_format = styles['fmt_currency']
    c_labsum.fill = styles['fill_subtotal']
    for c in ['A34', 'B34', 'C34', 'D34', 'E34', 'F34', 'G34', 'H34']:
        ws[c].border = styles['border_double_bottom']
    ws.row_dimensions[34].height = 22
    
    # SECTION 3: SUNDRIES (Rows 36 to 37)
    ws.merge_cells('A36:H36')
    c_sec3 = ws['A36']
    c_sec3.value = '3. SUNDRIES & LUMP-SUM ALLOWANCES'
    c_sec3.font = styles['font_white_bold']
    c_sec3.fill = styles['fill_header']
    c_sec3.alignment = styles['align_left']
    ws.row_dimensions[36].height = 24
    
    ws['A37'] = '1'
    ws['A37'].alignment = styles['align_center']
    ws['B37'] = '9999'
    ws['B37'].alignment = styles['align_center']
    ws['B37'].font = styles['font_bold']
    ws['C37'] = 'Sundries (guided by Sundries_Reference sheet)'
    ws['D37'] = 'L.S.'
    ws['D37'].alignment = styles['align_center']
    
    ws['E37'] = config.get('default_sundries_base', 0.0)
    ws['E37'].alignment = styles['align_right']
    ws['E37'].number_format = '0.00'
    ws['E37'].fill = styles['fill_input']
    
    ws['F37'] = '=Factor_Sundries'
    ws['F37'].alignment = styles['align_right']
    ws['F37'].number_format = '0.00'
    ws['F37'].fill = styles['fill_lookup']
    
    ws['G37'] = '=ROUND(E37 * F37, 2)'
    ws['G37'].alignment = styles['align_right']
    ws['G37'].number_format = styles['fmt_currency']
    ws['G37'].fill = styles['fill_calc']
    ws['G37'].font = styles['font_bold']
    
    ws['H37'] = 'Cost Index Multiplier (2.00) applied on base PAR-2012 allowance'
    ws['H37'].font = styles['font_note']
    
    for c in ['A37', 'B37', 'C37', 'D37', 'E37', 'F37', 'G37', 'H37']:
        ws[c].border = styles['border_thin']
    ws.row_dimensions[37].height = 22
    
    # SECTION 4: STATUTORY MARKUP CHAIN (Rows 39 to 51)
    ws.merge_cells('A39:H39')
    c_sec4 = ws['A39']
    c_sec4.value = '4. STATUTORY MARKUPS & FINAL RATE DERIVATION (CPWD DAR METHOD)'
    c_sec4.font = styles['font_white_bold']
    c_sec4.fill = styles['fill_header']
    c_sec4.alignment = styles['align_left']
    ws.row_dimensions[39].height = 24
    
    mu_headers = ['Step', 'Component / Markup Description', 'Apply? (YES/NO)', 'Basis Applied', 'Base Amount (₹)', 'Factor / %', 'Amount (₹)', 'CPWD Statutory Rule & Guidance Note']
    for c_idx, h in enumerate(mu_headers, 1):
        cell = ws.cell(row=40, column=c_idx, value=h)
        cell.font = styles['font_header']
        cell.fill = styles['fill_header']
        cell.alignment = styles['align_center']
        cell.border = styles['border_header']
    ws.row_dimensions[40].height = 24
    
    toggles = config.get('toggles', {'water': 'YES', 'gst': 'YES', 'cpoh': 'YES', 'cess': 'YES'})
    notes = config.get('toggle_notes', {})
    
    # Row 41: Base Direct Cost (W)
    ws['A41'] = 'W'
    ws['B41'] = 'Base Direct Production Cost (Materials + Labour + Sundries)'
    ws['C41'] = '—'
    ws['D41'] = 'Direct Sum'
    ws['E41'] = '—'
    ws['F41'] = '—'
    ws['G41'] = '=G20 + G34 + G37'
    ws['H41'] = 'Total direct cost of production (W) before statutory overheads'
    
    # Row 42: Water Charges
    ws['A42'] = 'X1'
    ws['B42'] = 'Add Water Charges (1% on W)'
    ws['C42'] = toggles['water']
    ws['D42'] = 'On "W"'
    ws['E42'] = '=G41'
    ws['F42'] = '=Factor_Water'
    ws['G42'] = '=IF(C42="YES", ROUND(E42 * F42, 2), 0)'
    ws['H42'] = notes.get('water', 'CPWD Standard: 1% for curing & site water. Toggle NO if dry item or water supplied free.')
    
    # Row 43: Subtotal (X)
    ws['A43'] = 'X'
    ws['B43'] = 'Subtotal "X" (W + Water Charges)'
    ws['C43'] = '—'
    ws['D43'] = 'W + Water'
    ws['E43'] = '—'
    ws['F43'] = '—'
    ws['G43'] = '=G41 + G42'
    ws['H43'] = 'Compounded base for GST calculation'
    
    # Row 44: GST
    ws['A44'] = 'Y1'
    ws['B44'] = 'Add GST on Works Contract'
    ws['C44'] = toggles['gst']
    ws['D44'] = 'On "X"'
    ws['E44'] = '=G43'
    ws['F44'] = '=Factor_GST'
    ws['G44'] = '=IF(C44="YES", ROUND(E44 * F44, 2), 0)'
    ws['H44'] = notes.get('gst', 'CPWD DAR 2019 factor 0.1405 (works contract). Toggle NO if tax exempt.')
    
    # Row 45: Subtotal (Y)
    ws['A45'] = 'Y'
    ws['B45'] = 'Subtotal "Y" (X + GST)'
    ws['C45'] = '—'
    ws['D45'] = 'X + GST'
    ws['E45'] = '—'
    ws['F45'] = '—'
    ws['G45'] = '=G43 + G44'
    ws['H45'] = 'Compounded base for Contractor Profit & Overheads'
    
    # Row 46: CPOH
    ws['A46'] = 'Z1'
    ws['B46'] = 'Add Contractor Profit & Overheads (15% CPOH)'
    ws['C46'] = toggles['cpoh']
    ws['D46'] = 'On "Y"'
    ws['E46'] = '=G45'
    ws['F46'] = '=Factor_CPOH'
    ws['G46'] = '=IF(C46="YES", ROUND(E46 * F46, 2), 0)'
    ws['H46'] = notes.get('cpoh', 'Standard CPWD 15% allowance for site overheads, head office costs & margin.')
    
    # Row 47: Subtotal (Z)
    ws['A47'] = 'Z'
    ws['B47'] = 'Subtotal "Z" (Y + CPOH)'
    ws['C47'] = '—'
    ws['D47'] = 'Y + CPOH'
    ws['E47'] = '—'
    ws['F47'] = '—'
    ws['G47'] = '=G45 + G46'
    ws['H47'] = 'Compounded base for Labour Welfare Cess'
    
    # Row 48: BOCW Cess
    ws['A48'] = 'Z2'
    ws['B48'] = 'Add Labour Welfare Cess (1% BOCW Cess)'
    ws['C48'] = toggles['cess']
    ws['D48'] = 'On "Z"'
    ws['E48'] = '=G47'
    ws['F48'] = '=Factor_Cess'
    ws['G48'] = '=IF(C48="YES", ROUND(E48 * F48, 2), 0)'
    ws['H48'] = notes.get('cess', 'Statutory 1% Building and Other Construction Workers Welfare Cess.')
    
    # Row 49: Cost of Specified Output
    ws['A49'] = 'Cost'
    ws['B49'] = '="Total Cost of Batch (" & TEXT(D5, "0.00") & " " & F5 & "):"'
    ws['C49'] = '—'
    ws['D49'] = 'Z + Cess'
    ws['E49'] = '—'
    ws['F49'] = '—'
    ws['G49'] = '=G47 + G48'
    ws['H49'] = 'Total evaluated cost for the batch size in Row 5'
    
    # Row 50: Cost per 1.00 Unit
    ws['A50'] = 'Rate'
    ws['B50'] = '="Analyzed Unit Rate per 1.00 " & F5 & ":"'
    ws['C50'] = '—'
    ws['D50'] = 'Cost / Batch Qty'
    ws['E50'] = '—'
    ws['F50'] = '—'
    ws['G50'] = '=ROUND(G49 / D5, 2)'
    ws['H50'] = 'Calculated cost per standard unit of measurement'
    
    # Row 51: OFFICIAL SAY RATE
    ws['A51'] = 'SAY'
    ws['B51'] = '="OFFICIAL SAY RATE (₹ per " & F5 & "):"'
    ws['C51'] = '—'
    ws['D51'] = 'Final Say'
    ws['E51'] = '—'
    ws['F51'] = '—'
    ws['G51'] = '=ROUND(G50, 2)'
    ws['H51'] = 'OFFICIAL CPWD ROUNDED RATE FOR TENDER SCHEDULES & ESTIMATES'
    
    # Formatting markup rows
    for r in range(41, 52):
        ws.cell(row=r, column=1).alignment = styles['align_center']
        ws.cell(row=r, column=1).font = styles['font_bold']
        ws.cell(row=r, column=2).alignment = styles['align_left']
        ws.cell(row=r, column=3).alignment = styles['align_center']
        ws.cell(row=r, column=4).alignment = styles['align_center']
        
        c_base = ws.cell(row=r, column=5)
        c_base.alignment = styles['align_right']
        if r in [42, 44, 46, 48]:
            c_base.number_format = styles['fmt_currency']
            
        c_fac = ws.cell(row=r, column=6)
        c_fac.alignment = styles['align_right']
        if r in [42, 44, 46, 48]:
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
            
        if r in [41, 43, 45, 47]:
            for c in range(1, 8):
                ws.cell(row=r, column=c).fill = styles['fill_subtotal']
        elif r in [42, 44, 46, 48]:
            ws.cell(row=r, column=3).fill = styles['fill_input']
            ws.cell(row=r, column=3).font = styles['font_bold']
            dv_yesno.add(ws.cell(row=r, column=3))
        elif r == 49:
            for c in range(1, 8):
                ws.cell(row=r, column=c).fill = styles['fill_result']
        elif r == 50:
            for c in range(1, 8):
                ws.cell(row=r, column=c).fill = styles['fill_subtotal']
        elif r == 51:
            ws.row_dimensions[51].height = 30
            for c in range(1, 8):
                ws.cell(row=r, column=c).fill = styles['fill_say']
            ws['B51'].font = styles['font_say']
            ws['G51'].font = styles['font_say']
            
        ws.row_dimensions[r].height = 24 if r != 51 else 30
        
    # SECTION 5: RUNNING CUSTOM ITEMS LIBRARY (Rows 54+)
    ws.merge_cells('A54:H54')
    c_sec5 = ws['A54']
    c_sec5.value = '5. RUNNING CUSTOM NON-DSR ITEMS LIBRARY FOR THIS TRADE'
    c_sec5.font = styles['font_white_bold']
    c_sec5.fill = styles['fill_header']
    c_sec5.alignment = styles['align_left']
    ws.row_dimensions[54].height = 24
    
    ws.merge_cells('A55:H55')
    c_inst = ws['A55']
    c_inst.value = 'Log all completed custom rate analyses for this trade below. Items logged here maintain an internal project audit trail and can be referenced in composite items.'
    c_inst.font = styles['font_note']
    c_inst.fill = styles['fill_note']
    c_inst.alignment = styles['align_wrap']
    ws.row_dimensions[55].height = 20
    
    lib_headers = ['Item Code', 'Item Nomenclature / Specification', 'Unit', 'Output Basis', 'Direct Cost W (₹)', 'Markups Applied', 'Unit Rate (₹)', 'Say Rate (₹)']
    for c_idx, h in enumerate(lib_headers, 1):
        cell = ws.cell(row=56, column=c_idx, value=h)
        cell.font = styles['font_header']
        cell.fill = styles['fill_header']
        cell.alignment = styles['align_center']
        cell.border = styles['border_header']
    ws.row_dimensions[56].height = 24
    
    sample_lib = config.get('sample_library', [])
    for idx in range(8):
        r = 57 + idx
        item = sample_lib[idx] if idx < len(sample_lib) else None
        
        ws.cell(row=r, column=1, value=item['code'] if item else f"C-{config['sheet_name'][:2]}.0{idx+1}").alignment = styles['align_center']
        ws.cell(row=r, column=2, value=item['desc'] if item else '(Available for new custom item)').alignment = styles['align_left']
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
        
    # Column dimensions
    ws.column_dimensions['A'].width = 10
    ws.column_dimensions['B'].width = 24
    ws.column_dimensions['C'].width = 50
    ws.column_dimensions['D'].width = 16
    ws.column_dimensions['E'].width = 18
    ws.column_dimensions['F'].width = 18
    ws.column_dimensions['G'].width = 20
    ws.column_dimensions['H'].width = 48
    
    
    # Unlock user editable yellow cells and lock calculations
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            # Default is locked
            cell.protection = Protection(locked=True)
            
    # Unprotect yellow input cells
    ws['B5'].protection = Protection(locked=False)
    ws['D5'].protection = Protection(locked=False)
    ws['F5'].protection = Protection(locked=False)
    ws['B6'].protection = Protection(locked=False)
    ws['E37'].protection = Protection(locked=False)
    ws['C42'].protection = Protection(locked=False)
    ws['C44'].protection = Protection(locked=False)
    ws['C46'].protection = Protection(locked=False)
    ws['C48'].protection = Protection(locked=False)
    
    for r in range(10, 20):
        ws.cell(row=r, column=2).protection = Protection(locked=False) # code
        ws.cell(row=r, column=5).protection = Protection(locked=False) # qty
        ws.cell(row=r, column=8).protection = Protection(locked=False) # note
    for r in range(24, 34):
        ws.cell(row=r, column=2).protection = Protection(locked=False) # code
        ws.cell(row=r, column=5).protection = Protection(locked=False) # qty
        ws.cell(row=r, column=8).protection = Protection(locked=False) # note
    for r in range(57, 65):
        for c in range(1, 9):
            ws.cell(row=r, column=c).protection = Protection(locked=False) # running library
            
    ws.protection.sheet = True
    ws.freeze_panes = 'A4'
    print(f"Built standard trade sheet: {config['sheet_name']}")

