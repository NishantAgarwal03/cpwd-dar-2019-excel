# -*- coding: utf-8 -*-
import openpyxl
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.table import Table, TableStyleInfo
import json
from openpyxl.styles import Protection
from scripts.work_types import classify, summarise, WORK_TYPE_GROUP, WORK_TYPE_INDEX
from scripts.paths import RATES_MASTER_JSON, LABOUR_PRODUCTIVITY_JSON, SUNDRIES_REFERENCE_JSON


def lock_sheet(ws, editable_cells=(), editable_cols=(), first_data_row=2):
    """Lock every cell, then reopen only what is genuinely meant to be edited.

    `editable_cells`  - explicit coordinates (e.g. project parameters).
    `editable_cols`   - column letters unlocked for every data row, used for the
                        rate column on Rates_Master so rates can still be revised
                        without exposing the code/description lookup keys.
    """
    for row in ws.iter_rows():
        for c in row:
            c.protection = Protection(locked=True)
    for ref in editable_cells:
        ws[ref].protection = Protection(locked=False)
    for col in editable_cols:
        for r in range(first_data_row, ws.max_row + 1):
            ws[f'{col}{r}'].protection = Protection(locked=False)
    ws.protection.sheet = True

def build_vol1_cover(wb, styles):
    """
    Cover / Instructions sheet for Vol. 1.
    Explains template usage, sheet layout, and how Vol. 2 links back.
    """
    ws = wb.create_sheet(title='Vol_1_Cover')
    ws.views.sheetView[0].showGridLines = False

    # Title banner
    ws.merge_cells('A1:I1')
    c = ws.cell(row=1, column=1)
    c.value = (
        'CPWD DAR 2019  |  CUSTOM RATE ANALYSIS WORKBOOK — VOLUME 1 (Sub-Heads 01–12)\n'
        'Template File  |  Open → Work → Save As <Project Name>.xlsx  |  Do NOT rename this file'
    )
    c.font      = styles['font_title']
    c.fill      = styles['fill_title']
    c.alignment = styles['align_center']
    ws.row_dimensions[1].height = 40

    # Template usage note
    ws.merge_cells('A3:I5')
    t = ws.cell(row=3, column=1)
    t.value = (
        '★  TEMPLATE — THIS FILE OPENS FRESH EACH TIME  ★\n'
        'Nothing you type here is saved to the template. To keep your work:\n'
        '   File → Save As → give it a project/item name (e.g. "ItemAnalysis_CC_Flooring.xlsx").'
    )
    t.font      = styles['font_header']
    t.fill      = styles['fill_note']
    t.alignment = styles['align_wrap']
    t.border    = styles['border_thin']
    ws.row_dimensions[3].height = 55

    # How-to note
    ws.merge_cells('A7:I16')
    n = ws.cell(row=7, column=1)
    n.value = (
        'HOW TO USE THIS WORKBOOK\n\n'
        '1. SHEET LAYOUT (each trade sheet 01–12 has the same structure):\n'
        '   Panel 1  — Scope & Item Identification (yellow INPUT cells)\n'
        '   Section 2 — Material Block  (enter code → rate auto-fills from Rates_Master)\n'
        '   Section 3 — Labour Block    (same lookup pattern)\n'
        '   Section 4 — Sundries, Markup Chain, Say Rate\n\n'
        '2. YELLOW CELLS are user inputs. All other cells are calculated — do not overtype them.\n\n'
        '3. RATES_MASTER: 2,200 CPWD DAR 2019 codes with basic rates. Type the code in Column B\n'
        '   of the material/labour block; description, unit and rate fill automatically.\n'
        '   To override a rate, type it directly in Column F of that row.\n\n'
        '4. GLOBAL FACTORS (Water Tax, GST, CPOH, Cess, Sundries%) are in the Global_Factors\n'
        '   sheet. Change them there once; all trade sheets update automatically.\n\n'
        '5. IF YOU USE VOL. 2 (Sub-Heads 13–26):\n'
        '   Before opening Vol. 2, save THIS file as:\n'
        '       CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_1.xlsx\n'
        '   in the SAME FOLDER as Vol. 2. Vol. 2 external links depend on that exact filename.\n\n'
        '6. REFERENCE SHEETS (read-only):\n'
        '   Rates_Master, Global_Factors, Labour_Machinery_Productivity, Sundries_Reference,\n'
        '   Resolved_Cross_Volume_Items — do not edit these; regenerate from source if needed.'
    )
    n.font      = styles['font_note']
    n.fill      = styles['fill_note']
    n.alignment = styles['align_wrap']
    n.border    = styles['border_thin']
    ws.row_dimensions[7].height = 320

    for col, w in {'A': 18, 'B': 16, 'C': 46, 'D': 14, 'E': 16,
                   'F': 16, 'G': 18, 'H': 42, 'I': 14}.items():
        ws.column_dimensions[col].width = w

    # Static instructions only, nothing here is meant to be edited.
    lock_sheet(ws)

    print('  Built Vol_1_Cover sheet')


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
    
    with open(RATES_MASTER_JSON, 'r', encoding='utf-8') as rf:
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
    
    # Column E (Basic Rate) stays editable so rates can be revised; the code,
    # category, description and unit columns are locked because every INDEX/MATCH
    # on every builder sheet depends on column A matching exactly.
    lock_sheet(ws, editable_cols=('E',))

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
        
    ws.merge_cells('D5:G8')
    c_pm_note = ws['D5']
    c_pm_note.value = "CPWD PLANT & MACHINERY HIRE CONVENTION (Item Codes 0001 to 0083):\n1. Hire charges include services of operating staff, lubricating oil, diesel/petrol/kerosene, and all consumables for running the machinery, and exclude GST.\n2. Hire charges are on per-day basis for a single shift of eight working hours.\n3. Do not add duplicate fuel or operator lines unless using un-fueled codes (e.g. 0084 onwards)."
    c_pm_note.font = styles['font_note']
    c_pm_note.fill = styles['fill_note']
    c_pm_note.alignment = styles['align_wrap']
    c_pm_note.border = styles['border_thin']
        
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
        ('10', 'Steel Work', 'YES', 'YES', 'YES', 'YES', 'Applies the CPWD (W-A) rule: the priming coat (resolved item 13.50.3) already carries Water, GST, CPOH and Cess, so it is tagged A and every markup is computed on (W-A), (X-A), (Y-A), (Z-A). Verified against DAR item 10.1.'),
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
    
    # Only the project-parameter cells and the override column are editable.
    # The "Effective Rate Applied" column holds the resolver formulas that every
    # builder sheet reads through Factor_Water / Factor_GST / Factor_CPOH /
    # Factor_Cess / Factor_Sundries, so it must not be overtypeable.
    lock_sheet(ws, editable_cells=('B5', 'B6', 'B7', 'B8',
                                   'E12', 'E13', 'E14', 'E15', 'E16'))

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
    """Work-type organised labour & machinery productivity reference.

    Section 1 rolls every mined LABOUR/MACHINERY line up by WORK TYPE (per the
    Problem & Solution Statement 2.2) so a user can look up "what crew does
    brickwork need" without knowing which sub-head sheet it was printed on.
    Section 2 keeps the full record-level detail behind it, now carrying the
    same Work Type and Trade Group columns so it can be filtered the same way.
    """
    ws = wb.create_sheet(title='Labour_Machinery_Productivity')
    ws.views.sheetView[0].showGridLines = True

    with open(LABOUR_PRODUCTIVITY_JSON, 'r', encoding='utf-8') as pf:
        records = json.load(pf)

    for rec in records:
        rec['work_type'] = classify(rec['subhead'], rec.get('item_desc'), rec.get('description'))

    summary = summarise(records)

    # ------------------------------------------------------------------
    # Section 1: Work-type crew summary
    # ------------------------------------------------------------------
    ws.merge_cells('A1:L1')
    t = ws['A1']
    t.value = ('CPWD DAR 2019 - LABOUR & MACHINERY PRODUCTIVITY, ORGANISED BY WORK TYPE '
               '(Section 1 = crew summary per work type / Section 2 = full source detail)')
    t.font = styles['font_title']
    t.fill = styles['fill_title']
    t.alignment = styles['align_center']
    ws.row_dimensions[1].height = 30

    ws.merge_cells('A2:L2')
    g = ws['A2']
    g.value = ('Advisory reference only - nothing here is a compute dependency. "Typical Coeff" is the median '
               'day-coefficient observed across every DAR item of that work type, at the output basis shown; '
               'Low/High give the range actually printed in the book, and Observations says how many item lines '
               'back it. Carry a crew into a builder sheet as a starting point, then edit the coefficients to '
               'suit the item being priced.')
    g.font = styles['font_note']
    g.fill = styles['fill_note']
    g.alignment = styles['align_wrap']
    ws.row_dimensions[2].height = 40

    ws.merge_cells('A4:L4')
    s1 = ws['A4']
    s1.value = '1. CREW SUMMARY BY WORK TYPE (median day-coefficient per output basis)'
    s1.font = styles['font_white_bold']
    s1.fill = styles['fill_header']
    s1.alignment = styles['align_left']
    ws.row_dimensions[4].height = 24

    sum_headers = ['Work Type', 'Trade Group', 'Resource Type', 'Code', 'Resource Description', 'Unit',
                   'Typical Coeff', 'Low', 'High', 'Observations', 'Output Basis', 'Basic Rate (Rs)']
    for ci, h in enumerate(sum_headers, 1):
        c = ws.cell(row=5, column=ci, value=h)
        c.font = styles['font_header']
        c.fill = styles['fill_header']
        c.alignment = styles['align_center']
        c.border = styles['border_header']
    ws.row_dimensions[5].height = 26

    sum_first = 6
    prev_wt = None
    for i, s in enumerate(summary):
        r = sum_first + i
        new_group = (s['work_type'] != prev_wt)
        prev_wt = s['work_type']

        ws.cell(row=r, column=1, value=s['work_type']).alignment = styles['align_left']
        ws.cell(row=r, column=2, value=s['trade_group']).alignment = styles['align_center']
        ws.cell(row=r, column=3, value=s['res_type']).alignment = styles['align_center']
        cc = ws.cell(row=r, column=4, value=s['code'])
        cc.alignment = styles['align_center']
        cc.font = styles['font_bold']
        ws.cell(row=r, column=5, value=s['description']).alignment = styles['align_left']
        ws.cell(row=r, column=6, value=s['unit']).alignment = styles['align_center']

        ct = ws.cell(row=r, column=7, value=s['typical'])
        ct.alignment = styles['align_right']
        ct.number_format = styles['fmt_qty']
        ct.font = styles['font_bold']
        ct.fill = styles['fill_result']

        for col, key in ((8, 'low'), (9, 'high')):
            cx = ws.cell(row=r, column=col, value=s[key])
            cx.alignment = styles['align_right']
            cx.number_format = styles['fmt_qty']

        ws.cell(row=r, column=10, value=s['observations']).alignment = styles['align_center']
        ws.cell(row=r, column=11, value=s['basis']).alignment = styles['align_center']
        cr = ws.cell(row=r, column=12, value=s['rate'])
        cr.alignment = styles['align_right']
        cr.number_format = styles['fmt_currency']

        for c in range(1, 13):
            cell = ws.cell(row=r, column=c)
            cell.border = styles['border_thin']
            if c != 7:
                cell.fill = styles['fill_lookup'] if new_group else styles['fill_calc']
                cell.font = styles['font_regular']
        if new_group:
            ws.cell(row=r, column=1).font = styles['font_bold']
        ws.cell(row=r, column=4).font = styles['font_bold']
        ws.row_dimensions[r].height = 20

    sum_last = sum_first + len(summary) - 1

    # ------------------------------------------------------------------
    # Section 2: full detail
    # ------------------------------------------------------------------
    sec2 = sum_last + 2
    ws.merge_cells(start_row=sec2, start_column=1, end_row=sec2, end_column=12)
    s2 = ws.cell(row=sec2, column=1)
    s2.value = '2. FULL SOURCE DETAIL (every LABOUR / MACHINERY line mined from the base volumes)'
    s2.font = styles['font_white_bold']
    s2.fill = styles['fill_header']
    s2.alignment = styles['align_left']
    ws.row_dimensions[sec2].height = 24

    det_headers = ['Work Type', 'Trade Group', 'Sub-Head', 'DAR Item No', 'Item Nomenclature', 'Output Basis',
                   'Resource Type', 'Code', 'Resource Description', 'Unit', 'Day Coeff / Qty', 'Basic Rate (Rs)']
    hdr_row = sec2 + 1
    for ci, h in enumerate(det_headers, 1):
        c = ws.cell(row=hdr_row, column=ci, value=h)
        c.font = styles['font_header']
        c.fill = styles['fill_header']
        c.alignment = styles['align_center']
        c.border = styles['border_header']
    ws.row_dimensions[hdr_row].height = 26

    records.sort(key=lambda x: (WORK_TYPE_INDEX[x['work_type']], x['subhead'], x['item_no'] or ''))

    det_first = hdr_row + 1
    for i, rec in enumerate(records):
        r = det_first + i
        ws.cell(row=r, column=1, value=rec['work_type']).alignment = styles['align_left']
        ws.cell(row=r, column=2, value=WORK_TYPE_GROUP[rec['work_type']]).alignment = styles['align_center']
        ws.cell(row=r, column=3, value=rec['subhead']).alignment = styles['align_left']
        ws.cell(row=r, column=4, value=rec['item_no']).alignment = styles['align_center']
        ws.cell(row=r, column=5, value=rec['item_desc']).alignment = styles['align_left']
        ws.cell(row=r, column=6, value=rec['basis']).alignment = styles['align_center']
        ws.cell(row=r, column=7, value=rec['type']).alignment = styles['align_center']
        c_code = ws.cell(row=r, column=8, value=rec['code'])
        c_code.alignment = styles['align_center']
        c_code.font = styles['font_bold']
        ws.cell(row=r, column=9, value=rec['description']).alignment = styles['align_left']
        ws.cell(row=r, column=10, value=rec['unit']).alignment = styles['align_center']

        c_coeff = ws.cell(row=r, column=11, value=rec['coefficient'])
        c_coeff.alignment = styles['align_right']
        c_coeff.number_format = styles['fmt_qty']

        c_rate = ws.cell(row=r, column=12, value=rec['rate'])
        c_rate.alignment = styles['align_right']
        c_rate.number_format = styles['fmt_currency']

        row_fill = styles['fill_subtotal'] if r % 2 == 0 else styles['fill_calc']
        for c in range(1, 13):
            cell = ws.cell(row=r, column=c)
            cell.fill = row_fill
            cell.border = styles['border_thin']
            if c not in [8, 11, 12]:
                cell.font = styles['font_regular']
        ws.row_dimensions[r].height = 20

    det_last = det_first + len(records) - 1

    widths = {'A': 38, 'B': 20, 'C': 24, 'D': 14, 'E': 45, 'F': 16,
              'G': 15, 'H': 12, 'I': 40, 'J': 12, 'K': 16, 'L': 16}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    ws.freeze_panes = 'A6'

    # Advisory reference only - nothing here is an input, so the whole sheet is
    # read-only. Filter and sort still work through the Excel Table.
    lock_sheet(ws)

    tab = Table(displayName="tbl_ProductivityDetail", ref="A%d:L%d" % (hdr_row, det_last))
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
    ws.add_table(tab)

    wb.defined_names.add(DefinedName(
        'Productivity_Summary',
        attr_text="'Labour_Machinery_Productivity'!$A$%d:$L$%d" % (sum_first, sum_last)))
    wb.defined_names.add(DefinedName(
        'Productivity_Work_Types',
        attr_text="'Labour_Machinery_Productivity'!$A$%d:$A$%d" % (sum_first, sum_last)))
    wb.defined_names.add(DefinedName(
        'Total_Labour_Norms',
        attr_text="COUNTA('Labour_Machinery_Productivity'!$C$%d:$C$%d)" % (det_first, det_last)))
    wb.defined_names.add(DefinedName(
        'Total_Work_Type_Crews',
        attr_text="ROWS('Labour_Machinery_Productivity'!$A$%d:$A$%d)" % (sum_first, sum_last)))

    n_wt = len(set(s['work_type'] for s in summary))
    print('Labour_Machinery_Productivity built: %d work types, %d crew-summary rows, %d detail records.'
          % (n_wt, len(summary), len(records)))

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
    
    with open(SUNDRIES_REFERENCE_JSON, 'r', encoding='utf-8') as sf:
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
    
    # Lookup reference only - read-only, like the productivity sheet.
    lock_sheet(ws)

    tab = Table(displayName="tbl_SundriesRef", ref=f"A1:H{len(records)+1}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
    ws.add_table(tab)
    wb.defined_names.add(DefinedName('Total_Sundries_Norms', attr_text="COUNTA('Sundries_Reference'!$A:$A)-1"))
    print(f'Sundries_Reference built: {len(records)} records with Table and Named Formula.')
