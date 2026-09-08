# -*- coding: utf-8 -*-
"""
Builds the `Resolved_Cross_Volume_Items` infrastructure sheet for WB1.

Per the Problem & Solution Statement (Round 2, decisions 6 and 7), Vol.1 builder
sheets need four items that are printed in Vol.2. Rather than leaving them as
literals or as live links into a Vol.2 workbook, each one is re-derived here
line-by-line from WB1's own `Rates_Master` catalogue and computed to its final
marked-up rate. The consuming builder sheets (08, 09, 10) then reference the
resolved *result*, so WB1 never needs Vol.2 present.

Items resolved:
  13.50.1  Priming coat, ready mixed pink/grey primer (wood)     -> 09 Wood & PVC
  13.50.3  Priming coat, red oxide zinc chromate primer (steel)  -> 09, 10
  13.57.1  Oil type wood preservative, new work, two+ coats      -> 09
  4.2.5    Cement concrete 1:3:6, 20 mm aggregate (WB1 internal) -> feeds 18.78
  18.78    Making chases 7.5x7.5 cm in walls, made good          -> 08 Cladding

CPWD "W - A" convention
-----------------------
Where an item imports a rate that ALREADY carries its own statutory markups,
the book tags that line "A" and excludes it from every markup base:
    Water on (W - A), GST on (X - A), CPOH on (Y - A), Cess on (Z - A)
Verified against Vol.1 item 10.1 and Vol.2 items 18.78 / 18.79 / 18.2.x.
Imports of PRE-markup rates (e.g. 03_Mortars) are NOT tagged and take the full
chain -- verified against Vol.1 item 6.1.1.
"""

from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import Protection


# ---------------------------------------------------------------------------
# Source data - every line transcribed from the converted DAR volumes.
# 'a_tag' True marks a line whose rate already contains statutory markups.
# ---------------------------------------------------------------------------

RESOLVED_ITEMS = [
    {
        'key': 'XV_4_2_5',
        'item_no': '4.2.5',
        'source_sh': '04 Concrete Work (Vol.1)',
        'nomenclature': ('Providing and laying in position cement concrete of specified grade excluding the cost of '
                         'centering and shuttering - 1:3:6 (1 cement : 3 coarse sand (zone-III) : 6 graded stone '
                         'aggregate 20 mm nominal size)'),
        'basis_qty': 1.0,
        'basis_unit': 'cum',
        'consumed_by': 'Feeds resolved item 18.78 below (not a Vol.2 item - re-derived here so 18.78 is self-contained)',
        'book_w': 5998.05,
        'book_say': 8025.00,
        'toggles': {'water': 'YES', 'gst': 'YES', 'cpoh': 'YES', 'cess': 'YES'},
        'materials': [
            ('0295', 0.70, 'Stone Aggregate (single size) 20 mm nominal size', False),
            ('0297', 0.24, 'Stone Aggregate (single size) 10 mm nominal size', False),
            ('2202', 0.94, 'Carriage of Stone aggregate below 40 mm nominal size', False),
            ('0982', 0.47, 'Coarse sand (zone III)', False),
            ('2203', 0.47, 'Carriage of Coarse sand', False),
            ('0367', 0.22, 'Portland Cement (0.15674 cum = 0.22 tonne)', False),
            ('2209', 0.22, 'Carriage of Cement', False),
        ],
        'labour': [
            ('0114', 0.90, 'Beldar'),
            ('0115', 0.78, 'Coolie'),
            ('0101', 0.70, 'Bhisti'),
            ('0123', 0.06, 'Mason (brick layer) 1st class'),
            ('0124', 0.06, 'Mason (brick layer) 2nd class'),
            ('0002', 0.07, 'Hire charges of Concrete Mixer 0.25 to 0.40 cum with hopper'),
            ('0012', 0.07, 'Vibrator (needle type 40 mm)'),
            ('0115', 1.88, 'Coolie - extra labour lifting material upto floor V level (0.75 x 2.5 = 1.88)'),
        ],
        'sundries': [
            (114.40, 'Scaffolding (17.60 x 2.5 avg.)'),
            (13.52, 'Sundries'),
        ],
    },
    {
        'key': 'XV_13_50_1',
        'item_no': '13.50.1',
        'source_sh': '13 Finishing (Vol.2)',
        'nomenclature': ('Applying priming coat with ready mixed pink or Grey primer of approved brand and '
                         'manufacture on wood work (hard and soft wood)'),
        'basis_qty': 10.0,
        'basis_unit': 'sqm',
        'consumed_by': '09_Wood_and_PVC_Work (primer on wood work)',
        'book_w': 426.48,
        'book_say': 57.05,
        'toggles': {'water': 'YES', 'gst': 'YES', 'cpoh': 'YES', 'cess': 'YES'},
        'materials': [
            ('0823', 0.75, 'Pink primer (for wood)', False),
        ],
        'labour': [
            ('0131', 0.25, 'Painter'),
            ('0115', 0.25, 'Coolie'),
        ],
        'sundries': [
            (2.73, 'Putty'),
            (0.39, 'Carriage of material'),
            (5.33, 'Brushes, sand paper etc.'),
            (10.79, 'Sundries'),
        ],
    },
    {
        'key': 'XV_13_50_3',
        'item_no': '13.50.3',
        'source_sh': '13 Finishing (Vol.2)',
        'nomenclature': ('Applying priming coat with ready mixed red oxide zinc chromate primer of approved brand '
                         'and manufacture on steel / iron work'),
        'basis_qty': 10.0,
        'basis_unit': 'sqm',
        'consumed_by': '10_Steel_Work (priming coat), 09_Wood_and_PVC_Work (steel fittings)',
        'book_w': 378.90,
        'book_say': 50.70,
        'toggles': {'water': 'YES', 'gst': 'YES', 'cpoh': 'YES', 'cess': 'YES'},
        'materials': [
            ('4202', 0.54, 'Red oxide zinc chromate primer', False),
        ],
        'labour': [
            ('0131', 0.24, 'Painter'),
            ('0115', 0.24, 'Coolie'),
        ],
        'sundries': [
            (0.52, 'Carriage of material'),
            (10.79, 'Brushes, sand paper including sundries'),
        ],
    },
    {
        'key': 'XV_13_57_1',
        'item_no': '13.57.1',
        'source_sh': '13 Finishing (Vol.2)',
        'nomenclature': ('Painting with oil type wood preservative of approved brand and manufacture - new work '
                         '(two or more coats)'),
        'basis_qty': 10.0,
        'basis_unit': 'sqm',
        'consumed_by': '09_Wood_and_PVC_Work (preservative treatment on timber)',
        'book_w': 332.71,
        'book_say': 44.50,
        'toggles': {'water': 'YES', 'gst': 'YES', 'cpoh': 'YES', 'cess': 'YES'},
        'materials': [
            ('0859', 1.00, 'Oil type wood preservative', False),
        ],
        'labour': [
            ('0131', 0.15, 'Painter'),
            ('0115', 0.15, 'Coolie'),
        ],
        'sundries': [
            (0.52, 'Carriage of material'),
            (4.16, 'Brushes etc.'),
            (3.90, 'Sundries'),
        ],
    },
    {
        'key': 'XV_18_78',
        'item_no': '18.78',
        'source_sh': '18 Water Supply (Vol.2)',
        'nomenclature': ('Making chases up to 7.5 x 7.5 cm in walls including making good and finishing the same '
                         'with cement concrete 1:3:6'),
        'basis_qty': 10.0,
        'basis_unit': 'metre',
        'consumed_by': '08_Cladding_Work (chase cutting for concealed fixings / services)',
        'book_w': 1233.25,
        'book_say': 154.15,
        'toggles': {'water': 'YES', 'gst': 'YES', 'cpoh': 'YES', 'cess': 'YES'},
        'materials': [
            # A-tagged: 4.2.5 is imported at its already-marked-up rate.
            ('4.2.5', 0.04, 'Cement concrete 1:3:6 - 0.075 x 0.075 x 10 m = 0.05625 cum, less 33% for pipe '
                            '= 0.03769 cum, say 0.04 cum. Rate resolved above.', True),
        ],
        'labour': [
            ('0123', 0.25, 'Mason (brick layer) 1st class'),
            ('0124', 0.25, 'Mason (brick layer) 2nd class'),
            ('0114', 1.00, 'Beldar'),
        ],
        'sundries': [],
    },
]

BLOCK_HEADERS = ['Line', 'Code / Source', 'Description / Specification', 'Unit',
                 'Quantity / Coeff', 'Basic Rate (Rs)', 'Amount (Rs)',
                 'Source Reference / Remarks', 'Markup Tag']


def _style_row(ws, r, styles, cols=9):
    for c in range(1, cols + 1):
        ws.cell(row=r, column=c).border = styles['border_thin']


def _build_block(ws, item, start_row, styles, dv_yesno, prior_rate_cells):
    """Emit one fully-resolved item block. Returns (next_free_row, say_cell)."""
    r = start_row

    # --- Block heading -----------------------------------------------------
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9)
    c = ws.cell(row=r, column=1)
    c.value = f"ITEM {item['item_no']}  -  {item['source_sh']}  -  {item['nomenclature']}"
    c.font = styles['font_white_bold']
    c.fill = styles['fill_header']
    c.alignment = styles['align_wrap']
    ws.row_dimensions[r].height = 32
    r += 1

    # --- Basis line --------------------------------------------------------
    ws.cell(row=r, column=1, value='Basis:').font = styles['font_bold']
    ws.cell(row=r, column=2, value=item['basis_qty']).alignment = styles['align_center']
    ws.cell(row=r, column=2).number_format = '0.00'
    ws.cell(row=r, column=3, value=item['basis_unit']).alignment = styles['align_left']
    ws.cell(row=r, column=4, value='Consumed by:').font = styles['font_bold']
    ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=9)
    cc = ws.cell(row=r, column=5, value=item['consumed_by'])
    cc.font = styles['font_note']
    cc.alignment = styles['align_wrap']
    _style_row(ws, r, styles)
    ws.row_dimensions[r].height = 20
    r += 1

    # --- Column headers ----------------------------------------------------
    for ci, h in enumerate(BLOCK_HEADERS, 1):
        hc = ws.cell(row=r, column=ci, value=h)
        hc.font = styles['font_header']
        hc.fill = styles['fill_header']
        hc.alignment = styles['align_center']
        hc.border = styles['border_header']
    ws.row_dimensions[r].height = 22
    r += 1

    first_line = r
    line_no = 0

    # --- MATERIAL lines ----------------------------------------------------
    for code, coeff, note, a_tag in item['materials']:
        line_no += 1
        ws.cell(row=r, column=1, value=line_no).alignment = styles['align_center']
        cd = ws.cell(row=r, column=2, value=code)
        cd.alignment = styles['align_center']
        cd.font = styles['font_bold']
        cd.fill = styles['fill_input']

        if a_tag:
            # Imported item rate - resolved earlier on this same sheet.
            ws.cell(row=r, column=3, value=f"Rate as per resolved item no. {code} above").alignment = styles['align_left']
            ws.cell(row=r, column=4, value='cum').alignment = styles['align_center']
            src = prior_rate_cells.get(code)
            ws.cell(row=r, column=6, value=(f'={src}' if src else 0)).alignment = styles['align_right']
        else:
            ws.cell(row=r, column=3,
                    value=f'=IFERROR(INDEX(Rates_Master!$C:$C, MATCH(B{r}, Rates_Master!$A:$A, 0)), "Code not found in Rates_Master")')
            ws.cell(row=r, column=4,
                    value=f'=IFERROR(INDEX(Rates_Master!$D:$D, MATCH(B{r}, Rates_Master!$A:$A, 0)), "")').alignment = styles['align_center']
            ws.cell(row=r, column=6,
                    value=f'=IFERROR(INDEX(Rates_Master!$E:$E, MATCH(B{r}, Rates_Master!$A:$A, 0)), 0)').alignment = styles['align_right']

        ws.cell(row=r, column=3).fill = styles['fill_lookup']
        ws.cell(row=r, column=4).fill = styles['fill_lookup']
        ws.cell(row=r, column=6).fill = styles['fill_lookup']
        ws.cell(row=r, column=6).number_format = styles['fmt_currency']

        cq = ws.cell(row=r, column=5, value=coeff)
        cq.alignment = styles['align_right']
        cq.number_format = styles['fmt_qty']
        cq.fill = styles['fill_input']

        ca = ws.cell(row=r, column=7, value=f'=ROUND(E{r} * F{r}, 2)')
        ca.alignment = styles['align_right']
        ca.number_format = styles['fmt_currency']

        cn = ws.cell(row=r, column=8, value=note)
        cn.font = styles['font_note']
        cn.alignment = styles['align_wrap']

        ct = ws.cell(row=r, column=9, value=('A' if a_tag else ''))
        ct.alignment = styles['align_center']
        ct.font = styles['font_bold']
        if a_tag:
            ct.fill = styles['fill_say']
        _style_row(ws, r, styles)
        ws.row_dimensions[r].height = 20
        r += 1

    # --- LABOUR lines ------------------------------------------------------
    for code, coeff, note in item['labour']:
        line_no += 1
        ws.cell(row=r, column=1, value=line_no).alignment = styles['align_center']
        cd = ws.cell(row=r, column=2, value=code)
        cd.alignment = styles['align_center']
        cd.font = styles['font_bold']
        cd.fill = styles['fill_input']
        ws.cell(row=r, column=3,
                value=f'=IFERROR(INDEX(Rates_Master!$C:$C, MATCH(B{r}, Rates_Master!$A:$A, 0)), "Code not found in Rates_Master")').fill = styles['fill_lookup']
        cu = ws.cell(row=r, column=4,
                     value=f'=IFERROR(INDEX(Rates_Master!$D:$D, MATCH(B{r}, Rates_Master!$A:$A, 0)), "")')
        cu.alignment = styles['align_center']
        cu.fill = styles['fill_lookup']
        cq = ws.cell(row=r, column=5, value=coeff)
        cq.alignment = styles['align_right']
        cq.number_format = styles['fmt_qty']
        cq.fill = styles['fill_input']
        cr = ws.cell(row=r, column=6,
                     value=f'=IFERROR(INDEX(Rates_Master!$E:$E, MATCH(B{r}, Rates_Master!$A:$A, 0)), 0)')
        cr.alignment = styles['align_right']
        cr.number_format = styles['fmt_currency']
        cr.fill = styles['fill_lookup']
        ca = ws.cell(row=r, column=7, value=f'=ROUND(E{r} * F{r}, 2)')
        ca.alignment = styles['align_right']
        ca.number_format = styles['fmt_currency']
        cn = ws.cell(row=r, column=8, value=note)
        cn.font = styles['font_note']
        cn.alignment = styles['align_wrap']
        ws.cell(row=r, column=9, value='').alignment = styles['align_center']
        _style_row(ws, r, styles)
        ws.row_dimensions[r].height = 20
        r += 1

    # --- SUNDRIES / L.S. lines --------------------------------------------
    for base, note in item['sundries']:
        line_no += 1
        ws.cell(row=r, column=1, value=line_no).alignment = styles['align_center']
        cd = ws.cell(row=r, column=2, value='9999')
        cd.alignment = styles['align_center']
        cd.font = styles['font_bold']
        ws.cell(row=r, column=3, value=note).alignment = styles['align_left']
        ws.cell(row=r, column=4, value='L.S.').alignment = styles['align_center']
        cq = ws.cell(row=r, column=5, value=base)
        cq.alignment = styles['align_right']
        cq.number_format = '0.00'
        cq.fill = styles['fill_input']
        cr = ws.cell(row=r, column=6, value='=Factor_Sundries')
        cr.alignment = styles['align_right']
        cr.number_format = '0.00'
        cr.fill = styles['fill_lookup']
        ca = ws.cell(row=r, column=7, value=f'=ROUND(E{r} * F{r}, 2)')
        ca.alignment = styles['align_right']
        ca.number_format = styles['fmt_currency']
        cn = ws.cell(row=r, column=8, value='Base L.S. allowance x Cost Index Multiplier (Global_Factors)')
        cn.font = styles['font_note']
        ws.cell(row=r, column=9, value='').alignment = styles['align_center']
        _style_row(ws, r, styles)
        ws.row_dimensions[r].height = 20
        r += 1

    last_line = r - 1

    # --- W, A, and the markup chain ---------------------------------------
    w_row = r
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    ws.cell(row=r, column=1, value='TOTAL DIRECT COST "W" (all lines above):').font = styles['font_bold']
    ws.cell(row=r, column=1).alignment = styles['align_right']
    cw = ws.cell(row=r, column=7, value=f'=ROUND(SUM(G{first_line}:G{last_line}), 2)')
    cw.font = styles['font_bold']
    cw.number_format = styles['fmt_currency']
    cw.alignment = styles['align_right']
    ws.cell(row=r, column=8,
            value=f"Book value of W for item {item['item_no']} = Rs {item['book_w']:,.2f}").font = styles['font_note']
    for c in range(1, 10):
        ws.cell(row=r, column=c).fill = styles['fill_subtotal']
    _style_row(ws, r, styles)
    r += 1

    a_row = r
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    ws.cell(row=r, column=1, value='Less "A" - imported lines already carrying statutory markups:').font = styles['font_bold']
    ws.cell(row=r, column=1).alignment = styles['align_right']
    ca = ws.cell(row=r, column=7, value=f'=ROUND(SUMIF($I${first_line}:$I${last_line}, "A", $G${first_line}:$G${last_line}), 2)')
    ca.font = styles['font_bold']
    ca.number_format = styles['fmt_currency']
    ca.alignment = styles['align_right']
    ws.cell(row=r, column=8,
            value='CPWD convention: markups are computed on (W-A), (X-A), (Y-A), (Z-A). A line tagged "A" '
                  'is an imported rate that already includes Water, GST, CPOH and Cess.').font = styles['font_note']
    ws.cell(row=r, column=8).alignment = styles['align_wrap']
    for c in range(1, 10):
        ws.cell(row=r, column=c).fill = styles['fill_subtotal']
    _style_row(ws, r, styles)
    ws.row_dimensions[r].height = 28
    r += 1

    tg = item['toggles']
    chain_start = r
    # Explicit row map so every base and subtotal points at the right cell.
    r_x1, r_x, r_y1, r_y, r_z1, r_z, r_z2 = (chain_start + i for i in range(7))
    chain = [
        (r_x1, 'X1', 'Add Water Charges on (W - A)', tg['water'],
         f'=G{w_row} - G{a_row}', '=Factor_Water',
         'CPWD: 1% water charges, computed on the direct cost net of A-tagged imports.'),
        (r_x, 'X', 'Subtotal "X" (W + Water Charges)', None,
         None, None, f'=G{w_row} + G{r_x1}'),
        (r_y1, 'Y1', 'Add GST on (X - A)', tg['gst'],
         f'=G{r_x} - G{a_row}', '=Factor_GST',
         'CPWD DAR 2019 works contract multiplying factor 0.1405, net of A.'),
        (r_y, 'Y', 'Subtotal "Y" (X + GST)', None,
         None, None, f'=G{r_x} + G{r_y1}'),
        (r_z1, 'Z1', 'Add Contractor Profit & Overheads on (Y - A)', tg['cpoh'],
         f'=G{r_y} - G{a_row}', '=Factor_CPOH',
         'Standard 15% CPOH, net of A.'),
        (r_z, 'Z', 'Subtotal "Z" (Y + CPOH)', None,
         None, None, f'=G{r_y} + G{r_z1}'),
        (r_z2, 'Z2', 'Add BOCW Welfare Cess on (Z - A)', tg['cess'],
         f'=G{r_z} - G{a_row}', '=Factor_Cess',
         'Statutory 1% BOCW Welfare Cess, net of A.'),
    ]

    for rw, step, desc, toggle, base_f, factor_f, tail in chain:
        ws.cell(row=rw, column=1, value=step).font = styles['font_bold']
        ws.cell(row=rw, column=1).alignment = styles['align_center']
        ws.cell(row=rw, column=3, value=desc).alignment = styles['align_left']
        if toggle is not None:
            ct = ws.cell(row=rw, column=2, value=toggle)
            ct.alignment = styles['align_center']
            ct.font = styles['font_bold']
            ct.fill = styles['fill_input']
            dv_yesno.add(ct)
            cb = ws.cell(row=rw, column=5, value=base_f)
            cb.number_format = styles['fmt_currency']
            cb.alignment = styles['align_right']
            cf = ws.cell(row=rw, column=6, value=factor_f)
            cf.number_format = styles['fmt_percent']
            cf.alignment = styles['align_right']
            ws.cell(row=rw, column=7, value=f'=IF(B{rw}="YES", ROUND(E{rw} * F{rw}, 2), 0)')
            ws.cell(row=rw, column=8, value=tail).font = styles['font_note']
        else:
            ws.cell(row=rw, column=2, value='-').alignment = styles['align_center']
            ws.cell(row=rw, column=7, value=tail)
            ws.cell(row=rw, column=8, value='Running compounded subtotal').font = styles['font_note']
            for c in range(1, 10):
                ws.cell(row=rw, column=c).fill = styles['fill_subtotal']
        ws.cell(row=rw, column=7).number_format = styles['fmt_currency']
        ws.cell(row=rw, column=7).alignment = styles['align_right']
        ws.cell(row=rw, column=7).font = styles['font_bold']
        ws.cell(row=rw, column=8).alignment = styles['align_wrap']
        _style_row(ws, rw, styles)
        ws.row_dimensions[rw].height = 22

    r = chain_start + 7
    z_row = r_z
    cess_row = r_z2

    # Cost of basis quantity
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    ws.cell(row=r, column=1,
            value=f"Cost of {item['basis_qty']:.2f} {item['basis_unit']}:").font = styles['font_bold']
    ws.cell(row=r, column=1).alignment = styles['align_right']
    cc = ws.cell(row=r, column=7, value=f'=ROUND(G{z_row} + G{cess_row}, 2)')
    cc.number_format = styles['fmt_currency']
    cc.alignment = styles['align_right']
    cc.font = styles['font_bold']
    for c in range(1, 10):
        ws.cell(row=r, column=c).fill = styles['fill_result']
    _style_row(ws, r, styles)
    cost_row = r
    r += 1

    # Say rate
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    ws.cell(row=r, column=1,
            value=f"RESOLVED RATE per 1.00 {item['basis_unit']}  (SAY):").font = styles['font_say']
    ws.cell(row=r, column=1).alignment = styles['align_right']
    cs = ws.cell(row=r, column=7, value=f"=ROUND(G{cost_row} / {item['basis_qty']}, 2)")
    cs.number_format = styles['fmt_currency']
    cs.alignment = styles['align_right']
    cs.font = styles['font_say']
    ws.cell(row=r, column=8,
            value=f"CPWD DAR 2019 printed Say rate for this item = Rs {item['book_say']:,.2f} "
                  f"per {item['basis_unit']}. Any difference means a Rates_Master rate has been "
                  f"edited away from the 2019 base.").font = styles['font_note']
    ws.cell(row=r, column=8).alignment = styles['align_wrap']
    for c in range(1, 10):
        ws.cell(row=r, column=c).fill = styles['fill_say']
    _style_row(ws, r, styles)
    ws.row_dimensions[r].height = 30
    say_row = r
    r += 2

    return r, f"'Resolved_Cross_Volume_Items'!$G${say_row}", say_row


def build_resolved_cross_volume(wb, styles):
    ws = wb.create_sheet(title='Resolved_Cross_Volume_Items')
    ws.views.sheetView[0].showGridLines = True

    ws.merge_cells('A1:I1')
    t = ws['A1']
    t.value = 'CPWD DAR 2019 - RESOLVED CROSS-VOLUME ITEMS (Vol.2 items re-derived from WB1 data only)'
    t.font = styles['font_title']
    t.fill = styles['fill_title']
    t.alignment = styles['align_center']
    ws.row_dimensions[1].height = 30

    ws['A2'] = 'LEGEND:'
    ws['A2'].font = styles['font_legend']
    for ref, txt, fill, font in [
        ('B2', 'User Input (Editable)', styles['fill_input'], styles['font_bold']),
        ('C2', 'Auto Lookup from Rates_Master', styles['fill_lookup'], styles['font_regular']),
        ('D2', 'Calculated', styles['fill_calc'], styles['font_regular']),
        ('E2', 'Resolved Rate', styles['fill_say'], styles['font_bold']),
        ('F2', 'CPWD Rule / Note', styles['fill_note'], styles['font_note']),
    ]:
        c = ws[ref]
        c.value, c.fill, c.font = txt, fill, font
        c.alignment = styles['align_center']
        c.border = styles['border_thin']

    ws.merge_cells('A3:I3')
    g = ws['A3']
    g.value = ('Vol.1 builder sheets 08, 09 and 10 need four items printed in Vol.2. Each is re-derived here '
               'line-by-line from this workbook\'s own Rates_Master catalogue and computed to a final rate, so '
               'WB1 never needs Vol.2 open. Consuming sheets reference the RESULT (see summary below), and tag '
               'that line "A" so it is excluded from their own markup chain per the CPWD (W-A) convention.')
    g.font = styles['font_note']
    g.fill = styles['fill_note']
    g.alignment = styles['align_wrap']
    ws.row_dimensions[3].height = 40

    dv_yesno = DataValidation(type='list', formula1='"YES,NO"', allow_blank=False)
    ws.add_data_validation(dv_yesno)

    # --- Section 1: summary (filled in after blocks are built) -------------
    ws.merge_cells('A5:I5')
    s1 = ws['A5']
    s1.value = '1. RESOLVED RATE SUMMARY (referenced by the consuming builder sheets)'
    s1.font = styles['font_white_bold']
    s1.fill = styles['fill_header']
    s1.alignment = styles['align_left']
    ws.row_dimensions[5].height = 24

    sum_headers = ['Item No', 'Source Sub-Head', 'Nomenclature', 'Unit', 'Basis',
                   'DAR 2019 Printed Rate', 'Resolved Rate (live)', 'Defined Name', 'Consumed By']
    for ci, h in enumerate(sum_headers, 1):
        c = ws.cell(row=6, column=ci, value=h)
        c.font = styles['font_header']
        c.fill = styles['fill_header']
        c.alignment = styles['align_center']
        c.border = styles['border_header']
    ws.row_dimensions[6].height = 24

    summary_first = 7
    summary_last = summary_first + len(RESOLVED_ITEMS) - 1

    # --- Section 2: the blocks --------------------------------------------
    sec_row = summary_last + 2
    ws.merge_cells(start_row=sec_row, start_column=1, end_row=sec_row, end_column=9)
    s2 = ws.cell(row=sec_row, column=1)
    s2.value = '2. FULL LINE-BY-LINE BREAKDOWNS (every rate traced to Rates_Master)'
    s2.font = styles['font_white_bold']
    s2.fill = styles['fill_header']
    s2.alignment = styles['align_left']
    ws.row_dimensions[sec_row].height = 24

    r = sec_row + 1
    rate_cells = {}
    say_rows = {}
    for item in RESOLVED_ITEMS:
        r, ref, say_row = _build_block(ws, item, r, styles, dv_yesno, rate_cells)
        rate_cells[item['item_no']] = ref
        say_rows[item['item_no']] = say_row

    # --- Fill summary rows -------------------------------------------------
    for idx, item in enumerate(RESOLVED_ITEMS):
        rr = summary_first + idx
        ws.cell(row=rr, column=1, value=item['item_no']).font = styles['font_bold']
        ws.cell(row=rr, column=1).alignment = styles['align_center']
        ws.cell(row=rr, column=2, value=item['source_sh']).alignment = styles['align_left']
        ws.cell(row=rr, column=3, value=item['nomenclature']).alignment = styles['align_wrap']
        ws.cell(row=rr, column=4, value=item['basis_unit']).alignment = styles['align_center']
        cb = ws.cell(row=rr, column=5, value=item['basis_qty'])
        cb.alignment = styles['align_center']
        cb.number_format = '0.00'
        cp = ws.cell(row=rr, column=6, value=item['book_say'])
        cp.alignment = styles['align_right']
        cp.number_format = styles['fmt_currency']
        cl = ws.cell(row=rr, column=7, value=f"=G{say_rows[item['item_no']]}")
        cl.alignment = styles['align_right']
        cl.number_format = styles['fmt_currency']
        cl.font = styles['font_say']
        cl.fill = styles['fill_say']
        ws.cell(row=rr, column=8, value=item['key']).alignment = styles['align_center']
        ws.cell(row=rr, column=8).font = styles['font_bold']
        ws.cell(row=rr, column=9, value=item['consumed_by']).alignment = styles['align_wrap']
        ws.cell(row=rr, column=9).font = styles['font_note']
        _style_row(ws, rr, styles)
        ws.row_dimensions[rr].height = 34

        wb.defined_names.add(DefinedName(
            item['key'],
            attr_text=f"'Resolved_Cross_Volume_Items'!$G${say_rows[item['item_no']]}"))

    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 26
    ws.column_dimensions['C'].width = 56
    ws.column_dimensions['D'].width = 14
    ws.column_dimensions['E'].width = 16
    ws.column_dimensions['F'].width = 18
    ws.column_dimensions['G'].width = 20
    ws.column_dimensions['H'].width = 60
    ws.column_dimensions['I'].width = 12

    # Lock everything except the yellow coefficient / code / toggle cells.
    for rr in range(1, ws.max_row + 1):
        for cc in range(1, 10):
            ws.cell(row=rr, column=cc).protection = Protection(locked=True)
    for rr in range(sec_row, ws.max_row + 1):
        f = ws.cell(row=rr, column=2).fill
        if f is not None and f.start_color is not None and f.start_color.rgb == styles['fill_input'].start_color.rgb:
            ws.cell(row=rr, column=2).protection = Protection(locked=False)
        f5 = ws.cell(row=rr, column=5).fill
        if f5 is not None and f5.start_color is not None and f5.start_color.rgb == styles['fill_input'].start_color.rgb:
            ws.cell(row=rr, column=5).protection = Protection(locked=False)
    ws.protection.sheet = True
    ws.freeze_panes = 'A4'

    print(f"Built Resolved_Cross_Volume_Items ({len(RESOLVED_ITEMS)} items resolved, {ws.max_row} rows)")
    return ws
