"""
support_builder_cladding.py
Rebuilds the existing '08_Cladding_Work' worksheet with First-Principles Resource,
Productivity, Gang, Machinery and Material Analysis per CPWD DAR 2019 Vol 1:
  Item Code | Labour / Machine / Material | Work done | Condition / When used | Category | Productivity | Quantity
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from pathlib import Path
import os, sys, shutil

# Styling definitions matching standard
C_TITLE_BG   = "1A1A2E"   # Navy dark
C_SUBTITLE_BG= "E8F4FD"   # Soft blue
C_HDR_PARENT = "1F4E79"   # Dark blue
C_HDR_SUB    = "2E75B6"   # Mid blue
C_TH_BG      = "1F4E79"   # Table header dark blue
C_WHITE      = "FFFFFF"
C_ALT_ROW    = "F9FBFD"

def fill(hex_col):
    if not hex_col:
        return PatternFill(fill_type=None)
    return PatternFill(fill_type="solid", fgColor=hex_col)

def thin_border():
    s = Side(style="thin", color="AAAAAA")
    return Border(left=s, right=s, top=s, bottom=s)

AL_C = Alignment(horizontal="center", vertical="center")
AL_L = Alignment(horizontal="left",   vertical="center", wrap_text=True)
AL_R = Alignment(horizontal="right",  vertical="center")

CLADDING_ITEMS = [
    # ─── 8.1 MARBLE WALL LINING WITH CRAMPS ─────────────────────────────────────
    {
        "id": "8.1.1.1",
        "parent_title": "8.1 MARBLE WORK (GANG SAW CUT) WALL LINING UP TO FLOOR V LEVEL",
        "title": "8.1.1.1 Raj Nagar Plain white / Udaipur green marble wall lining (slab area <= 0.50 sqm) (1 sqm basis)",
        "unit": "sqm", "base_qty": 1.0,
        "rows": [
            ("Raj Nagar White / Udaipur Green Marble slab 18mm", "Gang saw cut mirror polished marble slabs (area <= 0.50 sqm per slab) (1.20 sqm net incl. 20% cutting wastage)", "Premium marble slabs conforming to IS:1130; selected matching veins, mirror polished face", "Material", "1.20 sqm (1.200 sqm/sqm)", "1 sqm"),
            ("Carriage of Marble slabs", "Careful carriage, crate packing, and vertical hoisting of polished slabs to floor V", "Carriage allowance", "Material", "0.06 tonne", "1 sqm"),
            ("Gunmetal / Stainless steel cramps", "Specially shaped cramps (75x25x6mm) anchored into backing masonry with brass pins (4 nos./sqm)", "Non-corrosive cramps conforming to IS:3316 holding slabs securely to backing wall", "Material", "4.00 nos (4.00 nos/sqm)", "1 sqm"),
            ("Portland Cement", "Hydraulic binder decomposed from backing mortar 1:3 (REF#3.8) (0.029 cum mortar x 0.51 t/cum = 0.015 t)", "OPC-43 Grade cement; backing cavity filling behind marble slab", "Material", "0.015 tonne (0.015 t/sqm)", "1 sqm"),
            ("Coarse sand", "Zone III coarse sand decomposed from backing mortar 1:3 (REF#3.8) (0.029 cum x 1.07 = 0.031 cum)", "Clean natural river sand Zone III", "Material", "0.031 cum (0.031 cum/sqm)", "1 sqm"),
            ("White Portland cement", "Pure white cement for non-staining joint pointing and slurry backing (0.003 tonne)", "White cement conforming to IS:8042 (REF#3.15)", "Material", "0.003 tonne (0.003 t/sqm)", "1 sqm"),
            ("Carriage of Cement & Sand", "Haulage of cement and sand to mixing platform", "Carriage allowance", "Material", "0.05 cum/t equiv", "1 sqm"),
            ("Stone Mason (Marble fixer)", "Aligning slabs, drilling pin holes, inserting cramps into wall sockets, plumbing, and joint pointing", "Master marble fixer; 0.90 day = 7.20 man-hrs per sqm", "Labour", "7.20 man-hrs (0.14 sqm/man-hr)", "1 sqm"),
            ("Beldar", "Mixing backing mortar in small batches, filling cavity behind slab in lifts, and washing face", "Labour crew; 0.70 day = 5.60 man-hrs per sqm", "Labour", "5.60 man-hrs (0.18 sqm/man-hr)", "1 sqm"),
            ("Coolie", "Handling marble slabs with padded gloves, holding slabs in place until cramps set", "Handling crew; 0.40 day = 3.20 man-hrs per sqm", "Labour", "3.20 man-hrs (0.31 sqm/man-hr)", "1 sqm"),
            ("Bhisti", "Moist curing backing mortar and joints for 14 days", "Watering crew; 0.20 day = 1.60 man-hrs per sqm", "Labour", "1.60 man-hrs (0.62 sqm/man-hr)", "1 sqm"),
            ("Sundries", "Drill bits for pin holes, Araldite epoxy adhesive, plaster of Paris temporary setting dots, and carborundum stone", "Sundries allowance (L.S. 25.35 x cost index)", "Equipment", "L.S. allowance", "1 sqm"),
            ("REF#3.8 (Cement mortar 1:3)", "Referenced backing mortar: 0.029 cum cement mortar 1:3 per 1 sqm marble cladding", "Mortar scope executed per CPWD REF#3.8 specification", "Reference", "0.029 cum (0.029 cum/sqm)", "1 sqm"),
            ("REF#3.15 (White cement mortar 1:2)", "Referenced joint pointing mortar: 0.003 cum white cement mortar per 1 sqm cladding", "Mortar scope executed per CPWD REF#3.15 specification", "Reference", "0.003 cum (0.003 cum/sqm)", "1 sqm"),
        ]
    },
    {
        "id": "8.1.1.2",
        "parent_title": "8.1 MARBLE WORK (GANG SAW CUT) WALL LINING UP TO FLOOR V LEVEL",
        "title": "8.1.1.2 Raj Nagar Plain white / Udaipur green marble wall lining (large slab area > 0.50 sqm) (1 sqm basis)",
        "unit": "sqm", "base_qty": 1.0,
        "rows": [
            ("Raj Nagar White / Udaipur Green Marble slab 18mm", "Large-format gang saw cut marble slabs (area > 0.50 sqm per slab) (1.20 sqm net incl. cutting wastage)", "Selected architectural marble slabs with book-matched veining", "Material", "1.20 sqm (1.200 sqm/sqm)", "1 sqm"),
            ("Carriage of Marble slabs", "Crate transport and crane/hoist handling of heavy slabs", "Carriage allowance", "Material", "0.06 tonne", "1 sqm"),
            ("Gunmetal / Stainless steel cramps", "Heavy cramps (75x25x6mm) anchored with brass dowel pins (6 nos./sqm for large slabs)", "Heavy-duty structural stone cramps", "Material", "6.00 nos (6.00 nos/sqm)", "1 sqm"),
            ("Portland Cement", "Hydraulic binder decomposed from backing mortar 1:3 (REF#3.8) (0.015 tonne)", "OPC-43 Grade cement; backing cavity filling", "Material", "0.015 tonne (0.015 t/sqm)", "1 sqm"),
            ("Coarse sand", "Zone III coarse sand decomposed from backing mortar 1:3 (REF#3.8) (0.031 cum)", "Clean natural river sand Zone III", "Material", "0.031 cum (0.031 cum/sqm)", "1 sqm"),
            ("White Portland cement", "Pure white cement for non-staining joints (0.003 tonne)", "White cement conforming to IS:8042", "Material", "0.003 tonne (0.003 t/sqm)", "1 sqm"),
            ("Carriage of Cement & Sand", "Haulage of cement and sand to mixing platform", "Carriage allowance", "Material", "0.05 cum/t equiv", "1 sqm"),
            ("Stone Mason (Marble fixer)", "Handling large slabs, plumbing, adjusting cramp alignment, and seamless joint matching", "Master marble fixer; 1.05 day = 8.40 man-hrs per sqm", "Labour", "8.40 man-hrs (0.12 sqm/man-hr)", "1 sqm"),
            ("Beldar", "Mixing backing mortar, filling cavity behind slab in stages, and surface cleaning", "Labour crew; 0.80 day = 6.40 man-hrs per sqm", "Labour", "6.40 man-hrs (0.16 sqm/man-hr)", "1 sqm"),
            ("Coolie", "Two-man handling of large marble slabs with suction lifters to avoid corner fracture", "Handling crew; 0.60 day = 4.80 man-hrs per sqm", "Labour", "4.80 man-hrs (0.21 sqm/man-hr)", "1 sqm"),
            ("Bhisti", "Moist curing backing mortar and joints for 14 days", "Watering crew; 0.20 day = 1.60 man-hrs per sqm", "Labour", "1.60 man-hrs (0.62 sqm/man-hr)", "1 sqm"),
            ("Sundries", "Drill bits, Araldite epoxy, vacuum suction pads, PoP dots, and rubbing carborundum", "Sundries allowance (L.S. 27.30 x cost index)", "Equipment", "L.S. allowance", "1 sqm"),
            ("REF#3.8 (Cement mortar 1:3)", "Referenced backing mortar: 0.029 cum cement mortar 1:3 per 1 sqm cladding", "Mortar scope executed per CPWD REF#3.8 specification", "Reference", "0.029 cum (0.029 cum/sqm)", "1 sqm"),
            ("REF#3.15 (White cement mortar 1:2)", "Referenced joint pointing mortar: 0.003 cum white cement mortar per 1 sqm cladding", "Mortar scope executed per CPWD REF#3.15 specification", "Reference", "0.003 cum (0.003 cum/sqm)", "1 sqm"),
        ]
    },

    # ─── 8.2 GRANITE & MARBLE WALL CLADDING OVER 12MM MORTAR BED ────────────────
    {
        "id": "8.2.2.1",
        "parent_title": "8.2 GRANITE STONE CLADDING OVER MORTAR BED UP TO FLOOR V LEVEL",
        "title": "8.2.2.1 Granite stone cladding 18mm thick of any colour (slab area <= 0.50 sqm) over 12mm mortar bed 1:3 (1 sqm)",
        "unit": "sqm", "base_qty": 1.0,
        "rows": [
            ("Granite stone slabs 18mm thick", "Gang saw cut, mirror polished granite stone slabs of approved color and shade (1.05 sqm net incl. 5% cutting wastage)", "Premium granite slabs conforming to IS:3316; water absorption < 0.5%, compressive strength > 100 N/mm2", "Material", "1.05 sqm (1.050 sqm/sqm)", "1 sqm"),
            ("Carriage of Granite slabs", "Haulage and vertical shifting of polished granite slabs up to floor V", "Carriage allowance", "Material", "0.055 tonne", "1 sqm"),
            ("Portland Cement", "Hydraulic binder decomposed from 0.024 cum of cement mortar 1:3 (REF#3.8) (0.012 tonne) + neat cement slurry (0.003 t)", "OPC-43 Grade cement; bed mortar and slurry coat on slab back", "Material", "0.015 tonne (0.015 t/sqm)", "1 sqm"),
            ("Coarse sand", "Zone III coarse sand decomposed from 0.024 cum of cement mortar 1:3 (REF#3.8) (0.026 cum)", "Clean natural river sand Zone III", "Material", "0.026 cum (0.026 cum/sqm)", "1 sqm"),
            ("Matching epoxy / pigment paste", "UV-resistant colored epoxy grout matching granite shade for flush joint filling (joint thickness <= 1.5 mm)", "Two-component epoxy resin grout", "Material", "0.20 kg (0.200 kg/sqm)", "1 sqm"),
            ("Carriage of Cement & Sand", "Haulage of cement and sand to mixing platform", "Carriage allowance", "Material", "0.04 cum/t equiv", "1 sqm"),
            ("Stone Mason (Granite fixer)", "Applying 12mm mortar bed on wall, buttering slab with slurry, fixing granite true to line, tapping with rubber mallet", "Skilled granite fixer; 0.72 day = 5.76 man-hrs per sqm", "Labour", "5.76 man-hrs (0.17 sqm/man-hr)", "1 sqm"),
            ("Beldar", "Mixing bedding mortar, roughing backing wall with chisel, and washing granite backs", "Labour crew; 0.40 day = 3.20 man-hrs per sqm", "Labour", "3.20 man-hrs (0.31 sqm/man-hr)", "1 sqm"),
            ("Coolie", "Carrying granite slabs, mortar pans, and shifting scaffolding", "Handling crew; 0.30 day = 2.40 man-hrs per sqm", "Labour", "2.40 man-hrs (0.42 sqm/man-hr)", "1 sqm"),
            ("Bhisti", "Wetting brick/concrete backing wall and curing cladding for 14 days", "Watering crew; 0.15 day = 1.20 man-hrs per sqm", "Labour", "1.20 man-hrs (0.83 sqm/man-hr)", "1 sqm"),
            ("Sundries", "Rubber mallets, spirit levels, tile spacers, edge masking tape, and diamond blade cutting allowance", "Sundries allowance (L.S. 19.50 x cost index)", "Equipment", "L.S. allowance", "1 sqm"),
            ("REF#3.8 (Cement mortar 1:3)", "Referenced bedding mortar: 0.024 cum cement mortar 1:3 per 1 sqm granite cladding", "Mortar scope executed per CPWD REF#3.8 specification", "Reference", "0.024 cum (0.024 cum/sqm)", "1 sqm"),
        ]
    },
    {
        "id": "8.2.2.2",
        "parent_title": "8.2 GRANITE STONE CLADDING OVER MORTAR BED UP TO FLOOR V LEVEL",
        "title": "8.2.2.2 Granite stone cladding 18mm thick of any colour (large slab area > 0.50 sqm) over 12mm mortar bed 1:3 (1 sqm)",
        "unit": "sqm", "base_qty": 1.0,
        "rows": [
            ("Granite stone slabs 18mm thick", "Large-format gang saw cut, mirror polished granite slabs of approved color (1.05 sqm net)", "Premium granite slabs conforming to IS:3316; mirror polished, uniform shade", "Material", "1.05 sqm (1.050 sqm/sqm)", "1 sqm"),
            ("Carriage of Granite slabs", "Haulage and hoisting of large granite slabs up to floor V", "Carriage allowance", "Material", "0.055 tonne", "1 sqm"),
            ("Portland Cement", "Hydraulic binder decomposed from 0.024 cum of cement mortar 1:3 (REF#3.8) (0.012 tonne) + neat slurry (0.003 t)", "OPC-43 Grade cement; bed mortar and slurry coat", "Material", "0.015 tonne (0.015 t/sqm)", "1 sqm"),
            ("Coarse sand", "Zone III coarse sand decomposed from 0.024 cum of cement mortar 1:3 (REF#3.8) (0.026 cum)", "Clean natural river sand Zone III", "Material", "0.026 cum (0.026 cum/sqm)", "1 sqm"),
            ("Matching epoxy / pigment paste", "UV-resistant colored epoxy grout matching granite shade", "Two-component epoxy resin grout", "Material", "0.20 kg (0.200 kg/sqm)", "1 sqm"),
            ("Carriage of Cement & Sand", "Haulage of materials to mixing platform", "Carriage allowance", "Material", "0.04 cum/t equiv", "1 sqm"),
            ("Stone Mason (Granite fixer)", "Applying 12mm mortar bed, buttering slab with slurry, fixing large granite true to line, plumbing", "Skilled granite fixer; 0.85 day = 6.80 man-hrs per sqm", "Labour", "6.80 man-hrs (0.15 sqm/man-hr)", "1 sqm"),
            ("Beldar", "Mixing bedding mortar, roughing backing wall with chisel, and washing slab backs", "Labour crew; 0.50 day = 4.00 man-hrs per sqm", "Labour", "4.00 man-hrs (0.25 sqm/man-hr)", "1 sqm"),
            ("Coolie", "Two-man handling of heavy granite slabs with suction grips and positioning against wall", "Handling crew; 0.45 day = 3.60 man-hrs per sqm", "Labour", "3.60 man-hrs (0.28 sqm/man-hr)", "1 sqm"),
            ("Bhisti", "Wetting wall and curing cladding for 14 days", "Watering crew; 0.15 day = 1.20 man-hrs per sqm", "Labour", "1.20 man-hrs (0.83 sqm/man-hr)", "1 sqm"),
            ("Sundries", "Rubber mallets, suction grips, spirit levels, tile spacers, and diamond cutting blades", "Sundries allowance (L.S. 22.10 x cost index)", "Equipment", "L.S. allowance", "1 sqm"),
            ("REF#3.8 (Cement mortar 1:3)", "Referenced bedding mortar: 0.024 cum cement mortar 1:3 per 1 sqm cladding", "Mortar scope executed per CPWD REF#3.8 specification", "Reference", "0.024 cum (0.024 cum/sqm)", "1 sqm"),
        ]
    },

    # ─── 8.3 EDGE MOULDING TO MARBLE & GRANITE COUNTERS ─────────────────────────
    {
        "id": "8.3.1",
        "parent_title": "8.3 EDGE MOULDING TO STONE COUNTERS & FASCIA",
        "title": "8.3.1 Edge moulding to 18 mm thick marble stone counters, vanities and fascia (10 m run)",
        "unit": "metre", "base_qty": 10.0,
        "rows": [
            ("Edge profiling & polishing machine", "Variable speed edge profiling machine with diamond router profile wheels (full bull-nose / half bull-nose)", "Machine hire; 1.50 hr per 10 m edge profiling; output = 6.67 m/hr", "Machine", "1.50 machine-hrs (6.67 m/hr)", "10 m"),
            ("Diamond profiling bits & polishing pads", "Set of electroplated diamond profile router bits and resin-bond polishing pads (grit 50 to 3000)", "Tool wear and replacement allowance for 10 m run", "Equipment", "Tool allowance", "10 m"),
            ("Stone Mason (Polisher)", "Routing edge profile to uniform radius, wet grinding, and progressive diamond pad polishing to mirror lustre", "Skilled marble edge polisher; 0.25 day = 2.00 man-hrs per 10 m", "Labour", "2.00 man-hrs (5.00 m/man-hr)", "10 m"),
            ("Beldar", "Supplying water jet during routing, cleaning slurry, holding guide templates, and wiping dry", "Helper crew; 0.20 day = 1.60 man-hrs per 10 m", "Labour", "1.60 man-hrs (6.25 m/man-hr)", "10 m"),
            ("Sundries", "Oxalic acid / tin oxide rubbing powder, felt buffs, and edge masking tape", "Sundries allowance (L.S. 15.60 x cost index)", "Equipment", "L.S. allowance", "10 m"),
        ]
    },
    {
        "id": "8.3.2",
        "parent_title": "8.3 EDGE MOULDING TO STONE COUNTERS & FASCIA",
        "title": "8.3.2 Edge moulding to 18 mm thick granite stone counters, vanities and fascia (10 m run)",
        "unit": "metre", "base_qty": 10.0,
        "rows": [
            ("Edge profiling & polishing machine", "Heavy-duty electric edge router machine running diamond vacuum-brazed profiling wheels on hard granite", "Machine hire; 2.50 hr per 10 m granite profiling; output = 4.00 m/hr", "Machine", "2.50 machine-hrs (4.00 m/hr)", "10 m"),
            ("Diamond profiling bits & granite pads", "High-grade diamond vacuum-brazed router bits and 7-step wet granite polishing pads (grit 50 to 3000)", "Specialised granite diamond tooling allowance", "Equipment", "Tool allowance", "10 m"),
            ("Stone Mason (Polisher)", "Routing granite edge to true profile, step-by-step wet diamond polishing to high-gloss mirror finish", "Skilled granite polisher; 0.40 day = 3.20 man-hrs per 10 m", "Labour", "3.20 man-hrs (3.12 m/man-hr)", "10 m"),
            ("Beldar", "Supplying continuous cooling water flow, vacuuming granite slurry, and drying edge for gloss inspection", "Helper crew; 0.35 day = 2.80 man-hrs per 10 m", "Labour", "2.80 man-hrs (3.57 m/man-hr)", "10 m"),
            ("Sundries", "Diamond polishing paste, buffing wheels, protective masking tape, and gloss wax", "Sundries allowance (L.S. 23.40 x cost index)", "Equipment", "L.S. allowance", "10 m"),
        ]
    },

    # ─── 8.6 MIRROR POLISHING ON STONE WORK ─────────────────────────────────────
    {
        "id": "8.6",
        "parent_title": "8.6 MIRROR POLISHING ON STONE WORK",
        "title": "8.6 Mirror polishing on marble work / granite work / stone work where permitted (10 sqm)",
        "unit": "sqm", "base_qty": 10.0,
        "rows": [
            ("Electric rotary stone polishing machine", "Heavy floor/wall rotary polishing machine with orbital head (1.0 machine-hr per 10 sqm)", "Machine hire; output = 10.0 sqm/hr", "Machine", "1.00 machine-hrs (10.00 sqm/hr)", "10 sqm"),
            ("Carborundum & diamond polishing stones", "Set of carborundum abrasives (grits 60, 120, 320) and diamond polishing pucks (grits 400 to 3000)", "Abrasive stone consumption allowance", "Material", "Set of polishing pucks", "10 sqm"),
            ("Oxalic acid / Tin oxide polish powder", "High-purity chemical polishing powder reacting with calcium carbonate to produce glass-like mirror lustre (0.50 kg)", "Chemical polish compound", "Material", "0.50 kg (0.050 kg/sqm)", "10 sqm"),
            ("Stone Mason (Polisher)", "Operating rotary machine, checking surface flatness, applying slurry, and hand-buffing corners", "Skilled stone polisher; 0.35 day = 2.80 man-hrs per 10 sqm", "Labour", "2.80 man-hrs (3.57 sqm/man-hr)", "10 sqm"),
            ("Beldar", "Supplying water, spreading abrasive slurry, scraping dirty slurry, and final wet washing", "Labour crew; 0.30 day = 2.40 man-hrs per 10 sqm", "Labour", "2.40 man-hrs (4.17 sqm/man-hr)", "10 sqm"),
            ("Bhisti", "Supplying abundant water and final surface flush washing", "Watering crew; 0.15 day = 1.20 man-hrs per 10 sqm", "Labour", "1.20 man-hrs (8.33 sqm/man-hr)", "10 sqm"),
            ("Sundries", "Hessian rags, felt buffing pads, rubber squeegees, and protective masking", "Sundries allowance (L.S. 13.00 x cost index)", "Equipment", "L.S. allowance", "10 sqm"),
        ]
    },

    # ─── 8.7 CRAMPS FOR STONE CLADDING ──────────────────────────────────────────
    {
        "id": "8.7.2",
        "parent_title": "8.7 CRAMPS FOR STONE CLADDING",
        "title": "8.7.2 Providing and fixing Stainless Steel cramps in RCC/brick walls for stone veneer work (1 kg basis)",
        "unit": "kg", "base_qty": 1.0,
        "rows": [
            ("Stainless Steel cramps grade 304", "Fabricated SS 304 flat bar cramps (25x5mm or 25x6mm) with slotted holes and anchor pins (1.05 kg net)", "Grade AISI-304 stainless steel; non-magnetic, corrosion-proof stone veneer support", "Material", "1.05 kg (1.050 kg/kg)", "1 kg"),
            ("Carriage of Stainless Steel cramps", "Handling and transport of cramps to installation points", "Carriage allowance", "Material", "0.001 tonne", "1 kg"),
            ("Fast setting epoxy / polymer grout", "Epoxy resin mortar for anchoring cramp tangs into drilled masonry/RCC holes (0.15 kg)", "High-strength chemical anchoring grout", "Material", "0.15 kg (0.150 kg/kg)", "1 kg"),
            ("Rotary hammer drill machine", "Electric hammer drill with carbide masonry bits for drilling cramp anchor holes", "Drill machine hire allowance", "Equipment", "Drill allowance", "1 kg"),
            ("Blacksmith / Fitter", "Drilling anchor holes in concrete/brick wall, aligning cramp slots, inserting pins, and packing epoxy", "Skilled fitter; 0.08 day = 0.64 man-hrs per kg", "Labour", "0.64 man-hrs (1.56 kg/man-hr)", "1 kg"),
            ("Beldar", "Holding cramp alignment, mixing epoxy grout, and cleaning drill dust with hand blower", "Helper; 0.08 day = 0.64 man-hrs per kg", "Labour", "0.64 man-hrs (1.56 kg/man-hr)", "1 kg"),
            ("Sundries", "Carbide drill bits, blower bulbs, brass alignment pins, and cleaning solvents", "Sundries allowance (L.S. 3.90 x cost index)", "Equipment", "L.S. allowance", "1 kg"),
        ]
    },
]

def write_item_block(ws, item, start_row):
    r = start_row
    item_code = item["id"]
    base_qty = item["base_qty"]
    unit = item["unit"]
    qty_str = f"{int(base_qty) if base_qty == int(base_qty) else base_qty} {unit}"

    # 1. Parent Header row
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
    c_p = ws.cell(row=r, column=1, value=item["parent_title"])
    c_p.font = Font(name="Calibri", size=10, bold=True, color=C_WHITE)
    c_p.fill = fill(C_HDR_PARENT)
    c_p.alignment = AL_L
    ws.row_dimensions[r].height = 24
    for col in range(1, 8):
        ws.cell(row=r, column=col).border = thin_border()
    r += 1

    # 2. Item Header row
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
    c_i = ws.cell(row=r, column=1, value=item["title"])
    c_i.font = Font(name="Calibri", size=10, bold=True, color=C_WHITE)
    c_i.fill = fill(C_HDR_SUB)
    c_i.alignment = AL_L
    ws.row_dimensions[r].height = 24
    for col in range(1, 8):
        ws.cell(row=r, column=col).border = thin_border()
    r += 1

    # 3. Spacing row
    ws.row_dimensions[r].height = 6
    r += 1

    # 4. Table Column Header (7 Columns)
    headers = [
        "Item Code",
        "Labour / Machine / Material",
        "Work done",
        "Condition / When used",
        "Category",
        "Productivity",
        "Quantity"
    ]
    for col_idx, h_text in enumerate(headers, start=1):
        c_th = ws.cell(row=r, column=col_idx, value=h_text)
        c_th.font = Font(name="Calibri", size=9.5, bold=True, color=C_WHITE)
        c_th.fill = fill(C_TH_BG)
        c_th.alignment = AL_C
        c_th.border = thin_border()
    ws.row_dimensions[r].height = 22
    r += 1

    # 5. Data Rows
    for row_idx, data_row in enumerate(item["rows"]):
        disp_name, work_done, condition, category, prod_str, row_qty = data_row
        row_fill = fill(C_ALT_ROW if row_idx % 2 == 1 else C_WHITE)

        max_len = max(len(work_done), len(condition))
        ws.row_dimensions[r].height = max(20, min(80, (max_len // 45 + 1) * 14))

        # Col A: Item Code
        c1 = ws.cell(row=r, column=1, value=item_code)
        c1.font = Font(name="Calibri", size=9, bold=True)
        c1.alignment = AL_C
        c1.fill = row_fill
        c1.border = thin_border()

        # Col B: Labour / Machine / Material (Clean Name, NO code)
        c2 = ws.cell(row=r, column=2, value=disp_name)
        c2.font = Font(name="Calibri", size=9, bold=(category in ("Machine", "Equipment", "Material", "Reference")))
        c2.alignment = AL_L
        c2.fill = row_fill
        c2.border = thin_border()

        # Col C: Work done
        c3 = ws.cell(row=r, column=3, value=work_done)
        c3.font = Font(name="Calibri", size=9)
        c3.alignment = AL_L
        c3.fill = row_fill
        c3.border = thin_border()

        # Col D: Condition / When used
        c4 = ws.cell(row=r, column=4, value=condition)
        c4.font = Font(name="Calibri", size=9)
        c4.alignment = AL_L
        c4.fill = row_fill
        c4.border = thin_border()

        # Col E: Category
        c5 = ws.cell(row=r, column=5, value=category)
        c5.font = Font(name="Calibri", size=9)
        c5.alignment = AL_C
        c5.fill = row_fill
        c5.border = thin_border()

        # Col F: Productivity
        c6 = ws.cell(row=r, column=6, value=prod_str)
        c6.font = Font(name="Calibri", size=9, bold=True)
        c6.alignment = AL_R
        c6.fill = row_fill
        c6.border = thin_border()

        # Col G: Quantity
        c7 = ws.cell(row=r, column=7, value=row_qty)
        c7.font = Font(name="Calibri", size=9)
        c7.alignment = AL_C
        c7.fill = row_fill
        c7.border = thin_border()

        r += 1

    # Blank row separating items
    ws.row_dimensions[r].height = 12
    r += 1
    return r

def build_cladding_sheet(ws):
    if ws.views.sheetView:
        ws.views.sheetView[0].showGridLines = True

    # Sheet title spanning A1:G1
    ws.merge_cells("A1:G1")
    t1 = ws.cell(row=1, column=1, value="Sub-Head 8.0 — CLADDING WORK  |  First-Principles Resource, Work & Gang Analysis")
    t1.font = Font(name="Calibri", size=13, bold=True, color=C_WHITE)
    t1.fill = fill(C_TITLE_BG)
    t1.alignment = AL_C
    ws.row_dimensions[1].height = 28
    for col in range(1, 8):
        ws.cell(row=1, column=col).border = thin_border()

    # Subtitle spanning A2:G2
    ws.merge_cells("A2:G2")
    t2 = ws.cell(row=2, column=1, value="Evidence: CPWD DAR 2019 Vol 1  |  Format: Item Code • Labour / Machine / Material • Work done • Condition / When used • Category • Productivity • Standard Batch Quantity")
    t2.font = Font(name="Calibri", size=9, italic=True, color="333333")
    t2.fill = fill(C_SUBTITLE_BG)
    t2.alignment = AL_L
    ws.row_dimensions[2].height = 18
    for col in range(1, 8):
        ws.cell(row=2, column=col).border = thin_border()

    # Spacing row
    ws.row_dimensions[3].height = 8

    current_row = 4
    for item in CLADDING_ITEMS:
        current_row = write_item_block(ws, item, current_row)

    # Column widths for 7 columns (A to G)
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 34
    ws.column_dimensions["C"].width = 54
    ws.column_dimensions["D"].width = 52
    ws.column_dimensions["E"].width = 16
    ws.column_dimensions["F"].width = 28
    ws.column_dimensions["G"].width = 16

    ws.sheet_view.showGridLines = True
    ws.freeze_panes = "A4"

def rebuild_cladding_in_workbook(file_path):
    print(f"Opening {file_path} ...")
    wb = openpyxl.load_workbook(file_path)
    if "08_Cladding_Work" not in wb.sheetnames:
        raise ValueError(f"'08_Cladding_Work' sheet not found in {file_path}")

    pos = wb.sheetnames.index("08_Cladding_Work")
    del wb["08_Cladding_Work"]
    ws = wb.create_sheet("08_Cladding_Work", pos)
    ws.sheet_properties.tabColor = "2E75B6"
    build_cladding_sheet(ws)

    temp_path = file_path.replace(".xlsx", "_TMP_CLADDING.xlsx")
    wb.save(temp_path)
    wb.close()
    os.replace(temp_path, file_path)
    print(f"Successfully updated {file_path} -> sheet '08_Cladding_Work' ({len(CLADDING_ITEMS)} items).")

def main():
    repo_root = Path(__file__).resolve().parents[1]
    main_wb = repo_root / "CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_1_Latest.xlsx"
    output_wb = repo_root / "outputs" / "earthwork-custom-rate-composer" / "CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_1_Latest.xlsx"

    if main_wb.exists():
        rebuild_cladding_in_workbook(str(main_wb))
    if output_wb.exists():
        rebuild_cladding_in_workbook(str(output_wb))

if __name__ == "__main__":
    main()
