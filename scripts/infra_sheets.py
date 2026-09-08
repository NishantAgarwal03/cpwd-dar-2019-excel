# -*- coding: utf-8 -*-
import openpyxl
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.table import Table, TableStyleInfo
import json

def build_rates_master(wb, styles):
    ws = wb.create_sheet(title='Rates_Master')
    ws.views.sheetView[0].showGridLines = True
    
    headers = ['Code', 'Category', 'Description / Specification', 'Unit', 'Basic Rate (Rs)']
    for col_idx, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=col_idx, value=h)
        c.font = styles['font_header']
        c.fill = styles['fill_header']
        c.alignment = styles['align_center']
        c.border = styles['border_header']
    ws.row_dimensions[1].height = 28
    
    with open('rates_master_clean.json', 'r', encoding='utf-8') as rf:
        rates = json.load(rf)
        
    for r_idx, it in enumerate(rates, 2):
        c_code = ws.cell(row=r_idx, column=1, value=it['code'])
        c_cat = ws.cell(row=r_idx, column=2, value=it['category'])
        c_desc = ws.cell(row=r_idx, column=3, value=it['desc'])
        c_unit = ws.cell(row=r_idx, column=4, value=it['unit'])
        c_rate = ws.cell(row=r_idx, column=5, value=it['rate'])
        
        c_code.font = styles['font_bold']
        c_code.alignment = styles['align_center']
        c_cat.font = styles['font_regular']
        c_cat.alignment = styles['align_left']
        c_desc.font = styles['font_regular']
        c_desc.alignment = styles['align_left']
        c_unit.font = styles['font_regular']
        c_unit.alignment = styles['align_center']
        c_rate.font = styles['font_bold']
        c_rate.alignment = styles['align_right']
        c_rate.number_format = styles['fmt_currency']
        
        row_fill = styles['fill_subtotal'] if r_idx % 2 == 0 else styles['fill_calc']
        for cell in [c_code, c_cat, c_desc, c_unit, c_rate]:
            cell.fill = row_fill
            cell.border = styles['border_thin']
        ws.row_dimensions[r_idx].height = 20
        
    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 34
    ws.column_dimensions['C'].width = 65
    ws.column_dimensions['D'].width = 16
    ws.column_dimensions['E'].width = 18
    
    ws.freeze_panes = 'A2'
    
    wb.defined_names.add(DefinedName('Master_Codes', attr_text="'Rates_Master'!$A$2:$A$" + str(len(rates)+1)))
    wb.defined_names.add(DefinedName('Master_Rates_Table', attr_text="'Rates_Master'!$A$2:$E$" + str(len(rates)+1)))
    wb.defined_names.add(DefinedName('Total_Active_Rates', attr_text="COUNTA('Rates_Master'!$A:$A)-1"))
    
    tab = Table(displayName="tbl_RatesMaster", ref=f"A1:E{len(rates)+1}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
    ws.add_table(tab)
    print(f'Rates_Master built: {len(rates)} records with Table and Named Ranges.')

def build_global_factors(wb, styles):
    ws = wb.create_sheet(title='Global_Factors')
    ws.views.sheetView[0].showGridLines = True
    
    # Title Banner
    ws.merge_cells('A1:G1')
    c_title = ws['A1']
    c_title.value = 'CPWD DELHI ANALYSIS OF RATES (DAR 2019) — GLOBAL STATUTORY FACTORS & PROJECT SETTINGS'
    c_title.font = styles['font_title']
    c_title.fill = styles['fill_title']
    c_title.alignment = styles['align_center']
    ws.row_dimensions[1].height = 32
    
    # Legend
    ws['A2'] = 'LEGEND:'
    ws['A2'].font = styles['font_legend']
    ws['B2'] = 'Standard Default Value (CPWD DAR 2019)'
    ws['B2'].fill = styles['fill_subtotal']
    ws['B2'].font = styles['font_regular']
    ws['C2'] = 'Project Specific Override (User Input)'
    ws['C2'].fill = styles['fill_input']
    ws['C2'].font = styles['font_bold']
    ws['D2'] = 'Effective Value Used Across Workbook'
    ws['D2'].fill = styles['fill_result']
    ws['D2'].font = styles['font_bold']
    for col in ['A', 'B', 'C', 'D']:
        ws[f'{col}2'].border = styles['border_thin']
    ws.row_dimensions[2].height = 22
    
    # Section 1: Project Metadata
    ws['A4'] = '1. PROJECT & CONTRACT PARAMETERS'
    ws['A4'].font = styles['font_section']
    ws.row_dimensions[4].height = 24
    
    meta = [
        ('Project Name / Scope of Work:', 'General Construction & Maintenance Estimates (Non-DSR Items)'),
        ('Project Location / Zone:', 'Delhi & National Capital Region (CPWD Circle)'),
        ('Base Rates Schedule:', 'CPWD Delhi Analysis of Rates (DAR 2019, Vol. 1)'),
        ('Location Cost Index:', 118.0)
    ]
    for idx, item in enumerate(meta, 5):
        ws[f'A{idx}'] = item[0]
        ws[f'A{idx}'].font = styles['font_bold']
        ws[f'A{idx}'].border = styles['border_thin']
        ws[f'B{idx}'] = item[1]
        ws[f'B{idx}'].font = styles['font_regular']
        ws[f'B{idx}'].fill = styles['fill_input']
        ws[f'B{idx}'].border = styles['border_thin']
        ws.row_dimensions[idx].height = 20
        
    # Section 2: Statutory Factors Table
    ws['A10'] = '2. STATUTORY MARKUPS & OVERHEAD COEFFICIENTS'
    ws['A10'].font = styles['font_section']
    ws.row_dimensions[10].height = 24
    
    tbl_headers = ['S.No', 'Statutory Markup Component', 'CPWD Default Rate', 'Basis of Application', 'Project Override Rate', 'Effective Rate Applied', 'CPWD DAR 2019 Rule & Statutory Basis']
    for c_idx, h in enumerate(tbl_headers, 1):
        cell = ws.cell(row=11, column=c_idx, value=h)
        cell.font = styles['font_header']
        cell.fill = styles['fill_header']
        cell.alignment = styles['align_center']
        cell.border = styles['border_header']
    ws.row_dimensions[11].height = 26
    
    factors = [
        (1, 'Water Charges', 0.01, 'Applied on Direct Cost (W)', None, '=IF(E12<>"", E12, C12)', 'CPWD Standard: 1.0% on Direct Cost (W) for water consumed in construction/curing. Toggleable on trade sheets.'),
        (2, 'Goods and Services Tax (GST)', 0.1405, 'Applied on Compounded Base (X = W + Water)', None, '=IF(E13<>"", E13, C13)', 'CPWD DAR 2019 multiplying factor 0.1405 (14.05% effective works contract rate). Override here for 18% (0.18) or 12% (0.12).'),
        (3, "Contractor's Profit & Overheads (CPOH)", 0.15, 'Applied on Compounded Base (Y = X + GST)', None, '=IF(E14<>"", E14, C14)', 'CPWD DAR Standard: 15.0% for site management, head office overheads, tools, and contractor margin.'),
        (4, 'Building & Other Construction Workers Cess (BOCW)', 0.01, 'Applied on Compounded Base (Z = Y + CPOH)', None, '=IF(E15<>"", E15, C15)', 'Statutory Labour Welfare Cess @ 1.0% per BOCW Welfare Cess Act, 1996.'),
        (5, 'Sundries Cost Index Multiplier', 2.00, 'Applied on Base L.S. Allowance (PAR-2012 Base)', None, '=IF(E16<>"", E16, C16)', 'CPWD DAR 2019 Preface: Multiplier of 2.00 on base sundries allowance reflecting Delhi Cost Index (118 vs PAR-2012 base 100).')
    ]
    
    for r_idx, row_data in enumerate(factors, 12):
        ws.cell(row=r_idx, column=1, value=row_data[0]).alignment = styles['align_center']
        ws.cell(row=r_idx, column=2, value=row_data[1]).alignment = styles['align_left']
        c_def = ws.cell(row=r_idx, column=3, value=row_data[2])
        c_def.alignment = styles['align_right']
        c_def.number_format = styles['fmt_percent'] if row_data[0] != 5 else '0.00'
        c_def.fill = styles['fill_subtotal']
        
        ws.cell(row=r_idx, column=4, value=row_data[3]).alignment = styles['align_left']
        
        c_over = ws.cell(row=r_idx, column=5, value=row_data[4])
        c_over.alignment = styles['align_right']
        c_over.number_format = styles['fmt_percent'] if row_data[0] != 5 else '0.00'
        c_over.fill = styles['fill_input']
        
        c_eff = ws.cell(row=r_idx, column=6, value=row_data[5])
        c_eff.alignment = styles['align_right']
        c_eff.number_format = styles['fmt_percent'] if row_data[0] != 5 else '0.00'
        c_eff.fill = styles['fill_result']
        c_eff.font = styles['font_bold']
        
        c_rem = ws.cell(row=r_idx, column=7, value=row_data[6])
        c_rem.alignment = styles['align_wrap']
        c_rem.font = styles['font_note']
        
        for c in range(1, 8):
            ws.cell(row=r_idx, column=c).border = styles['border_thin']
        ws.row_dimensions[r_idx].height = 36
        
    # Section 3: Subhead Presets Table
    ws['A18'] = '3. CPWD SUB-HEAD STATUTORY PRESETS & TRADE CONVENTIONS'
    ws['A18'].font = styles['font_section']
    ws.row_dimensions[18].height = 24
    
    subhead_headers = ['Sub-Head Code', 'Trade Name', 'Water (1%)', 'GST (14.05%)', 'CPOH (15%)', 'Cess (1%)', 'CPWD Trade Statutory Convention & Rationale']
    for c_idx, h in enumerate(subhead_headers, 1):
        cell = ws.cell(row=19, column=c_idx, value=h)
        cell.font = styles['font_header']
        cell.fill = styles['fill_header']
        cell.alignment = styles['align_center']
        cell.border = styles['border_header']
    ws.row_dimensions[19].height = 24
    
    subhead_presets = [
        ('01', 'Carriage of Materials', 'NO', 'NO', 'YES', 'NO', 'Mechanical transport applies 15% CPOH only. Water & Cess omitted (no water used). GST handled separately in freight.'),
        ('02', 'Earth Work', 'YES', 'YES', 'YES', 'YES', 'Standard 5-step compounding chain applied across labour & earthmoving equipment.'),
        ('03', 'Mortars', 'NO', 'NO', 'NO', 'NO', 'Intermediate building block. Stops at direct cost (W). Markup chain is OFF to prevent double taxation when imported into Masonry/Flooring.'),
        ('04', 'Concrete Work', 'YES', 'YES', 'YES', 'YES', 'Standard 5-step compounding chain applied over materials, labour, mixer & vibrator hire.'),
        ('05', 'RCC Work', 'YES', 'YES', 'YES', 'YES', 'Standard 5-step compounding chain. Shuttering and reinforcement are measured & priced separately in CPWD.'),
        ('06', 'Masonry Work', 'YES', 'YES', 'YES', 'YES', 'Imports un-marked Mortar rates as a Material line, then applies full markup chain once over the combined total.'),
        ('07', 'Stone Work', 'YES', 'YES', 'YES', 'YES', 'Imports Mortar rates as Material line; applies full markup chain over stone, mortar, dressing and laying labour.'),
        ('08', 'Cladding Work', 'YES', 'YES', 'YES', 'YES', 'Imports both bedding mortar and pointing mortar rates; applies full markup chain over stone veneer, mortar, scaffolding.'),
        ('09', 'Wood & PVC Work', 'YES', 'YES', 'YES', 'YES', 'Applies full markup chain over timber in scantling (with 5% wastage allowance) and carpentry labour.'),
        ('10', 'Steel Work', 'YES', 'YES', 'YES', 'YES', 'Applies markups over (W - Priming Coat), because priming coat (Item 13.50.3) already contains full statutory markups under Finishing.'),
        ('11', 'Flooring', 'YES', 'YES', 'YES', 'YES', 'Imports Mortar/Concrete bed rate; applies full markup chain over tiles, stone, laying labour and polishing.'),
        ('12', 'Roofing', 'YES', 'YES', 'YES', 'YES', 'Applies full markup chain over roofing sheets, fixing J-hooks/bolts, limpet washers, and carpenter/fixing labour.')
    ]
    
    for r_idx, row_data in enumerate(subhead_presets, 20):
        for c_idx in range(1, 8):
            cell = ws.cell(row=r_idx, column=c_idx, value=row_data[c_idx-1])
            cell.font = styles['font_regular']
            cell.border = styles['border_thin']
            if c_idx in [1, 3, 4, 5, 6]:
                cell.alignment = styles['align_center']
                if row_data[c_idx-1] == 'YES':
                    cell.fill = styles['fill_result']
                    cell.font = styles['font_bold']
                elif row_data[c_idx-1] == 'NO':
                    cell.fill = styles['fill_subtotal']
            elif c_idx == 7:
                cell.alignment = styles['align_wrap']
                cell.font = styles['font_note']
            else:
                cell.alignment = styles['align_left']
                cell.font = styles['font_bold']
        ws.row_dimensions[r_idx].height = 26
        
    ws.column_dimensions['A'].width = 16
    ws.column_dimensions['B'].width = 30
    ws.column_dimensions['C'].width = 18
    ws.column_dimensions['D'].width = 32
    ws.column_dimensions['E'].width = 22
    ws.column_dimensions['F'].width = 22
    ws.column_dimensions['G'].width = 65
    
    ws.freeze_panes = 'A3'
    
    wb.defined_names.add(DefinedName('Factor_Water', attr_text="'Global_Factors'!$F$12"))
    wb.defined_names.add(DefinedName('Factor_GST', attr_text="'Global_Factors'!$F$13"))
    wb.defined_names.add(DefinedName('Factor_CPOH', attr_text="'Global_Factors'!$F$14"))
    wb.defined_names.add(DefinedName('Factor_Cess', attr_text="'Global_Factors'!$F$15"))
    wb.defined_names.add(DefinedName('Factor_Sundries', attr_text="'Global_Factors'!$F$16"))
    # Named Formulas for direct statutory resolution
    wb.defined_names.add(DefinedName('Resolved_Water_Factor', attr_text="IF('Global_Factors'!$E$12<>\"\",'Global_Factors'!$E$12,'Global_Factors'!$C$12)"))
    wb.defined_names.add(DefinedName('Resolved_GST_Factor', attr_text="IF('Global_Factors'!$E$13<>\"\",'Global_Factors'!$E$13,'Global_Factors'!$C$13)"))
    wb.defined_names.add(DefinedName('Resolved_CPOH_Factor', attr_text="IF('Global_Factors'!$E$14<>\"\",'Global_Factors'!$E$14,'Global_Factors'!$C$14)"))
    wb.defined_names.add(DefinedName('Resolved_Cess_Factor', attr_text="IF('Global_Factors'!$E$15<>\"\",'Global_Factors'!$E$15,'Global_Factors'!$C$15)"))
    print('Global_Factors built with Statutory Factor Named Ranges and Named Formulas.')

def build_labour_productivity(wb, styles):
    ws = wb.create_sheet(title='Labour_Machinery_Productivity')
    ws.views.sheetView[0].showGridLines = True
    
    headers = ['Sub-Head', 'DAR Item No', 'Item Nomenclature', 'Output Basis', 'Resource Type', 'Code', 'Resource Description', 'Unit', 'Day Coeff / Qty', 'Basic Rate (Rs)']
    for col_idx, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=col_idx, value=h)
        c.font = styles['font_header']
        c.fill = styles['fill_header']
        c.alignment = styles['align_center']
        c.border = styles['border_header']
    ws.row_dimensions[1].height = 28
    
    with open('labour_productivity.json', 'r', encoding='utf-8') as pf:
        records = json.load(pf)
        
    for r_idx, rec in enumerate(records, 2):
        ws.cell(row=r_idx, column=1, value=rec['subhead']).alignment = styles['align_left']
        ws.cell(row=r_idx, column=2, value=rec['item_no']).alignment = styles['align_center']
        ws.cell(row=r_idx, column=3, value=rec['item_desc']).alignment = styles['align_left']
        ws.cell(row=r_idx, column=4, value=rec['basis']).alignment = styles['align_center']
        ws.cell(row=r_idx, column=5, value=rec['type']).alignment = styles['align_center']
        c_code = ws.cell(row=r_idx, column=6, value=rec['code'])
        c_code.alignment = styles['align_center']
        c_code.font = styles['font_bold']
        ws.cell(row=r_idx, column=7, value=rec['description']).alignment = styles['align_left']
        ws.cell(row=r_idx, column=8, value=rec['unit']).alignment = styles['align_center']
        
        c_coeff = ws.cell(row=r_idx, column=9, value=rec['coefficient'])
        c_coeff.alignment = styles['align_right']
        c_coeff.number_format = styles['fmt_qty']
        
        c_rate = ws.cell(row=r_idx, column=10, value=rec['rate'])
        c_rate.alignment = styles['align_right']
        c_rate.number_format = styles['fmt_currency']
        
        row_fill = styles['fill_subtotal'] if r_idx % 2 == 0 else styles['fill_calc']
        for c in range(1, 11):
            cell = ws.cell(row=r_idx, column=c)
            cell.fill = row_fill
            cell.border = styles['border_thin']
            if c not in [6, 9, 10]:
                cell.font = styles['font_regular']
        ws.row_dimensions[r_idx].height = 20
        
    ws.column_dimensions['A'].width = 22
    ws.column_dimensions['B'].width = 14
    ws.column_dimensions['C'].width = 45
    ws.column_dimensions['D'].width = 18
    ws.column_dimensions['E'].width = 16
    ws.column_dimensions['F'].width = 12
    ws.column_dimensions['G'].width = 40
    ws.column_dimensions['H'].width = 12
    ws.column_dimensions['I'].width = 18
    ws.column_dimensions['J'].width = 16
    
    ws.freeze_panes = 'A2'
    
    tab = Table(displayName="tbl_Productivity", ref=f"A1:J{len(records)+1}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
    ws.add_table(tab)
    wb.defined_names.add(DefinedName('Total_Labour_Norms', attr_text="COUNTA('Labour_Machinery_Productivity'!$A:$A)-1"))
    print(f'Labour_Machinery_Productivity built: {len(records)} records with Table and Named Formula.')

def build_sundries_reference(wb, styles):
    ws = wb.create_sheet(title='Sundries_Reference')
    ws.views.sheetView[0].showGridLines = True
    
    headers = ['Sub-Head', 'DAR Item No', 'Item Nomenclature', 'Output Basis', 'Sundries Description', 'Base L.S. Quantity', 'Index Multiplier', 'Total Amount (Rs)']
    for col_idx, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=col_idx, value=h)
        c.font = styles['font_header']
        c.fill = styles['fill_header']
        c.alignment = styles['align_center']
        c.border = styles['border_header']
    ws.row_dimensions[1].height = 28
    
    with open('sundries_reference.json', 'r', encoding='utf-8') as sf:
        records = json.load(sf)
        
    for r_idx, rec in enumerate(records, 2):
        ws.cell(row=r_idx, column=1, value=rec['subhead']).alignment = styles['align_left']
        ws.cell(row=r_idx, column=2, value=rec['item_no']).alignment = styles['align_center']
        ws.cell(row=r_idx, column=3, value=rec['item_desc']).alignment = styles['align_left']
        ws.cell(row=r_idx, column=4, value=rec['basis']).alignment = styles['align_center']
        ws.cell(row=r_idx, column=5, value=rec['description']).alignment = styles['align_left']
        
        c_base = ws.cell(row=r_idx, column=6, value=rec['base_ls'])
        c_base.alignment = styles['align_right']
        c_base.number_format = '0.00'
        
        c_mul = ws.cell(row=r_idx, column=7, value=rec['multiplier'])
        c_mul.alignment = styles['align_right']
        c_mul.number_format = '0.00'
        
        c_amt = ws.cell(row=r_idx, column=8, value=rec['amount'])
        c_amt.alignment = styles['align_right']
        c_amt.number_format = styles['fmt_currency']
        c_amt.font = styles['font_bold']
        
        row_fill = styles['fill_subtotal'] if r_idx % 2 == 0 else styles['fill_calc']
        for c in range(1, 9):
            cell = ws.cell(row=r_idx, column=c)
            cell.fill = row_fill
            cell.border = styles['border_thin']
            if c != 8:
                cell.font = styles['font_regular']
        ws.row_dimensions[r_idx].height = 20
        
    ws.column_dimensions['A'].width = 22
    ws.column_dimensions['B'].width = 14
    ws.column_dimensions['C'].width = 45
    ws.column_dimensions['D'].width = 18
    ws.column_dimensions['E'].width = 35
    ws.column_dimensions['F'].width = 18
    ws.column_dimensions['G'].width = 16
    ws.column_dimensions['H'].width = 18
    
    ws.freeze_panes = 'A2'
    
    tab = Table(displayName="tbl_SundriesRef", ref=f"A1:H{len(records)+1}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
    ws.add_table(tab)
    wb.defined_names.add(DefinedName('Total_Sundries_Norms', attr_text="COUNTA('Sundries_Reference'!$A:$A)-1"))
    print(f'Sundries_Reference built: {len(records)} records with Table and Named Formula.')
