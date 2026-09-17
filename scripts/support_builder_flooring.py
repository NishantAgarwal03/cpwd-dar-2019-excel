"""
support_builder_flooring.py
Rebuilds the existing '11_Flooring' worksheet with First-Principles Resource,
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

FLOORING_ITEMS = [
    # ─── 11.1 BRICK ON EDGE FLOORING ────────────────────────────────────────────
    {
        "id": "11.1",
        "parent_title": "11.1 BRICK ON EDGE FLOORING",
        "title": "11.1 Brick on edge flooring with bricks class 7.5 on 12 mm bed of cement mortar 1:4 (10 sqm basis)",
        "unit": "10 sqm", "base_qty": 10.0,
        "rows": [
            ("Common burnt clay bricks class 7.5", "Selected non-modular bricks laid on edge in herringbone, diagonal or parallel pattern with 10mm joints (550 nos. incl. wastage)", "Conforming to IS:1077; compressive strength >= 7.5 N/mm2", "Material", "550.00 nos (55.00 nos/sqm)", "10 sqm"),
            ("Carriage of Bricks", "Haulage and stacking of bricks near laying area", "Carriage allowance", "Material", "550.00 nos", "10 sqm"),
            ("Portland Cement", "Hydraulic binder decomposed from 0.170 cum bedding/jointing mortar 1:4 (REF#3.9) (0.170 x 0.380 t = 0.065 t)", "OPC-43 Grade cement; bed and joint mortar", "Material", "0.065 tonne (0.0065 t/sqm)", "10 sqm"),
            ("Coarse sand", "Zone III sand decomposed from 0.170 cum mortar 1:4 (REF#3.9) (0.170 x 1.07 = 0.182 cum)", "Clean natural river sand", "Material", "0.182 cum (0.0182 cum/sqm)", "10 sqm"),
            ("Carriage of Cement & Sand", "Haulage to mixing platform", "Carriage allowance", "Material", "0.20 cum/t equiv", "10 sqm"),
            ("Mason (brick layer) 2nd class", "Laying 12mm mortar bed, placing bricks on edge true to slope/camber, striking joints flush, and washing surface", "Skilled floor mason; 0.80 day = 6.40 man-hrs per 10 sqm", "Labour", "6.40 man-hrs (1.56 sqm/man-hr)", "10 sqm"),
            ("Beldar", "Mixing mortar 1:4, soaking bricks in water tank, carrying bricks and mortar pans to mason", "Labour crew; 1.40 days = 11.20 man-hrs per 10 sqm", "Labour", "11.20 man-hrs (0.89 sqm/man-hr)", "10 sqm"),
            ("Bhisti", "Water soaking bricks prior to laying and continuous wet curing for 10 days", "Watering crew; 1.04 days = 8.32 man-hrs per 10 sqm", "Labour", "8.32 man-hrs (1.20 sqm/man-hr)", "10 sqm"),
            ("REF#3.9 (Cement mortar 1:4)", "Referenced bedding & jointing mortar: 0.170 cum CM 1:4 per 10 sqm brick floor", "Mortar scope executed per CPWD REF#3.9 specification", "Reference", "0.170 cum (0.017 cum/sqm)", "10 sqm"),
        ]
    },

    # ─── 11.3 CEMENT CONCRETE FLOORING 1:2:4 ────────────────────────────────────
    {
        "id": "11.3.1",
        "parent_title": "11.3 CEMENT CONCRETE FLOORING (IPS FLOORING)",
        "title": "11.3.1 Cement concrete flooring 1:2:4, 40 mm thick with 20 mm nominal size stone aggregate (10 sqm basis)",
        "unit": "10 sqm", "base_qty": 10.0,
        "rows": [
            ("Portland Cement", "Hydraulic binder for 40mm thick concrete layer + neat cement floating finish @ 2.2 kg/sqm (0.170 tonne)", "OPC-43 Grade conforming to IS:269 / IS:8112", "Material", "0.170 tonne (0.017 t/sqm)", "10 sqm"),
            ("Coarse sand (zone III)", "Clean graded river sand for concrete mix (0.178 cum)", "Fine aggregate conforming to IS:383", "Material", "0.178 cum (0.0178 cum/sqm)", "10 sqm"),
            ("Stone aggregate 20mm nominal size", "Single size 20mm broken granite/trap coarse aggregate (0.267 cum)", "Coarse aggregate conforming to IS:383", "Material", "0.267 cum (0.0267 cum/sqm)", "10 sqm"),
            ("Stone aggregate 10mm nominal size", "Single size 10mm stone chippings for dense void filling (0.089 cum)", "Coarse aggregate conforming to IS:383", "Material", "0.089 cum (0.0089 cum/sqm)", "10 sqm"),
            ("Carriage of Materials", "Haulage of cement, sand and aggregate to mixer platform", "Carriage allowance", "Material", "0.55 cum/t equiv", "10 sqm"),
            ("Mason (brick layer) 2nd class", "Laying concrete in alternate bays with glass/aluminium dividing strips, screeding with straight edge, and trowelling neat cement slurry finish", "Skilled floor mason; 0.80 day = 6.40 man-hrs per 10 sqm", "Labour", "6.40 man-hrs (1.56 sqm/man-hr)", "10 sqm"),
            ("Beldar", "Batching aggregates, feeding mixer, transporting concrete in barrows, tamping with heavy wooden thappies, and surface floating", "Labour crew; 1.40 days = 11.20 man-hrs per 10 sqm", "Labour", "11.20 man-hrs (0.89 sqm/man-hr)", "10 sqm"),
            ("Bhisti", "Flooding floor panels with pond curing for 14 days", "Watering crew; 1.04 days = 8.32 man-hrs per 10 sqm", "Labour", "8.32 man-hrs (1.20 sqm/man-hr)", "10 sqm"),
            ("Sundries & Dividing strips", "Glass strips (40x4mm) / aluminium strips, formwork wooden battens, straight edges, and trowels", "Sundries allowance", "Equipment", "L.S. allowance", "10 sqm"),
        ]
    },
    {
        "id": "11.3.2",
        "parent_title": "11.3 CEMENT CONCRETE FLOORING (IPS FLOORING)",
        "title": "11.3.2 Cement concrete flooring 1:2:4, 50 mm thick with 20 mm nominal size stone aggregate (10 sqm basis)",
        "unit": "10 sqm", "base_qty": 10.0,
        "rows": [
            ("Portland Cement", "Hydraulic binder for 50mm thick concrete slab + neat cement slurry trowel finish (0.210 tonne)", "OPC-43 Grade conforming to IS:269 / IS:8112", "Material", "0.210 tonne (0.021 t/sqm)", "10 sqm"),
            ("Coarse sand (zone III)", "Clean natural river sand Zone III (0.222 cum)", "Fine aggregate conforming to IS:383", "Material", "0.222 cum (0.0222 cum/sqm)", "10 sqm"),
            ("Stone aggregate 20mm nominal size", "Single size 20mm hard stone aggregate (0.334 cum)", "Coarse aggregate conforming to IS:383", "Material", "0.334 cum (0.0334 cum/sqm)", "10 sqm"),
            ("Stone aggregate 10mm nominal size", "Single size 10mm stone aggregate (0.111 cum)", "Coarse aggregate conforming to IS:383", "Material", "0.111 cum (0.0111 cum/sqm)", "10 sqm"),
            ("Carriage of Materials", "Transport and staging of materials to floor area", "Carriage allowance", "Material", "0.68 cum/t equiv", "10 sqm"),
            ("Mason (brick layer) 2nd class", "Leveling, laying alternate panels, compacting, straight-edge screening, applying neat cement slurry and steel trowelling", "Skilled floor mason; 0.90 day = 7.20 man-hrs per 10 sqm", "Labour", "7.20 man-hrs (1.39 sqm/man-hr)", "10 sqm"),
            ("Beldar", "Operating mixer, wheeling concrete, tamping, and making ponding bunds", "Labour crew; 1.60 days = 12.80 man-hrs per 10 sqm", "Labour", "12.80 man-hrs (0.78 sqm/man-hr)", "10 sqm"),
            ("Bhisti", "Pond curing floor continuously for 14 days", "Watering crew; 1.04 days = 8.32 man-hrs per 10 sqm", "Labour", "8.32 man-hrs (1.20 sqm/man-hr)", "10 sqm"),
            ("Sundries & Dividing strips", "Glass dividing strips (50x4mm), edge screed timber, and smoothing floats", "Sundries allowance", "Equipment", "L.S. allowance", "10 sqm"),
        ]
    },

    # ─── 11.6 CEMENT PLASTER SKIRTING ───────────────────────────────────────────
    {
        "id": "11.6",
        "parent_title": "11.6 CEMENT PLASTER SKIRTING",
        "title": "11.6 Cement plaster skirting up to 30 cm height with CM 1:3 (1 cement : 3 sand) & neat cement float (10 sqm basis)",
        "unit": "10 sqm", "base_qty": 10.0,
        "rows": [
            ("Portland Cement", "Hydraulic binder decomposed from 0.180 cum backing mortar 1:3 (REF#3.8) (0.092 t) + neat cement floating coat (0.044 t) = 0.136 tonne", "OPC-43 Grade cement; skirting wall coating", "Material", "0.136 tonne (0.0136 t/sqm)", "10 sqm"),
            ("Coarse sand (zone III)", "Zone III sand decomposed from 0.180 cum mortar 1:3 (REF#3.8) (0.180 x 1.07 = 0.193 cum)", "Clean natural river sand", "Material", "0.193 cum (0.0193 cum/sqm)", "10 sqm"),
            ("Carriage of Cement & Sand", "Haulage to plastering stations", "Carriage allowance", "Material", "0.22 cum/t equiv", "10 sqm"),
            ("Mason (brick layer) 1st class", "Raking brick joints, applying 18mm mortar coat true to plumb, forming top pencil groove/junction, and floating neat cement with steel trowel", "Master plaster mason; 1.25 days = 10.00 man-hrs per 10 sqm", "Labour", "10.00 man-hrs (1.00 sqm/man-hr)", "10 sqm"),
            ("Beldar", "Mixing mortar 1:3, hacking masonry base, carrying mortar hods, and cleaning floor splash", "Labour crew; 1.50 days = 12.00 man-hrs per 10 sqm", "Labour", "12.00 man-hrs (0.83 sqm/man-hr)", "10 sqm"),
            ("Bhisti", "Water spraying and curing skirting continuously for 14 days", "Watering crew; 0.75 day = 6.00 man-hrs per 10 sqm", "Labour", "6.00 man-hrs (1.67 sqm/man-hr)", "10 sqm"),
            ("Sundries", "Grooving tools, straight edges, plumb bobs, trowels, and scaffolding boards", "Sundries allowance", "Equipment", "L.S. allowance", "10 sqm"),
            ("REF#3.8 (Cement mortar 1:3)", "Referenced skirting mortar: 0.180 cum CM 1:3 per 10 sqm skirting", "Mortar scope executed per CPWD REF#3.8 specification", "Reference", "0.180 cum (0.018 cum/sqm)", "10 sqm"),
        ]
    },

    # ─── 11.9 MARBLE CHIPS (TERRAZZO) FLOORING ──────────────────────────────────
    {
        "id": "11.9.1",
        "parent_title": "11.9 MARBLE CHIPS (TERRAZZO) FLOORING",
        "title": "11.9.1 40 mm thick marble chips flooring with 34mm under layer 1:2:4 and 6mm topping of marble chips (10 sqm basis)",
        "unit": "10 sqm", "base_qty": 10.0,
        "rows": [
            ("Marble chips (graded sizes 1 to 4)", "White and black selected marble chips graded 2mm to 6mm size for terrazzo wearing coat (87.2 kg)", "Conforming to IS:2114; hard crystalline marble chips free from dust", "Material", "0.872 quintal (0.087 q/sqm)", "10 sqm"),
            ("Portland Cement (ordinary & white)", "Cement for 34mm under layer 1:2:4 (0.109 t) + cement for 6mm marble chip topping (0.020 t) + slurry (0.020 t) = 0.149 tonne", "OPC-43 and white cement for terrazzo matrix", "Material", "0.149 tonne (0.0149 t/sqm)", "10 sqm"),
            ("Stone aggregate (12.5mm & 10mm)", "Broken hard stone aggregate for 34mm under layer (0.303 cum)", "Coarse aggregate conforming to IS:383", "Material", "0.303 cum (0.0303 cum/sqm)", "10 sqm"),
            ("Coarse sand (zone III)", "Natural sand for under layer concrete (0.151 cum)", "Fine aggregate conforming to IS:383", "Material", "0.151 cum (0.0151 cum/sqm)", "10 sqm"),
            ("Carriage of Materials", "Transport of marble chips, aggregates, cement and sand", "Carriage allowance", "Material", "0.50 cum/t equiv", "10 sqm"),
            ("Mason (terrazzo specialist)", "Laying 34mm under layer in bays with glass strips, placing 6mm marble topping, rolling with steel roller, and trowelling level", "Master terrazzo mason; 1.00 day = 8.00 man-hrs per 10 sqm", "Labour", "8.00 man-hrs (1.25 sqm/man-hr)", "10 sqm"),
            ("Floor Polisher / Machine operator", "1st grinding (carborundum 60-grit), grouting pinholes with cement, 2nd grinding (120-grit), 3rd grinding (320-grit) & oxalic acid wash", "Polishing technician; 1.80 days = 14.40 man-hrs per 10 sqm", "Labour", "14.40 man-hrs (0.69 sqm/man-hr)", "10 sqm"),
            ("Terrazzo floor grinding machine", "Heavy-duty electric floor grinding machine with rotary grinding heads and water cooling attachment", "Continuous floor grinding and carborundum polishing", "Machinery", "1.00 machine-day (10.0 sqm/mach-day)", "10 sqm"),
            ("Beldar", "Batching, washing marble chips, removing polishing sludge, sweeping and scrubbing", "Labour crew; 2.50 days = 20.00 man-hrs per 10 sqm", "Labour", "20.00 man-hrs (0.50 sqm/man-hr)", "10 sqm"),
            ("Bhisti", "Water curing terrazzo under layer and topping between grindings for 14 days", "Watering crew; 1.20 days = 9.60 man-hrs per 10 sqm", "Labour", "9.60 man-hrs (1.04 sqm/man-hr)", "10 sqm"),
            ("Sundries & Polishing materials", "Carborundum blocks (60, 120, 320 grit), oxalic acid powder, floor wax polish, and glass dividing strips 40x4mm", "Polishing supplies", "Equipment", "L.S. allowance", "10 sqm"),
        ]
    },

    # ─── 11.23 MARBLE STONE FLOORING ────────────────────────────────────────────
    {
        "id": "11.23.1",
        "parent_title": "11.23 MARBLE STONE FLOORING",
        "title": "11.23.1 Makrana white marble stone flooring 18 mm thick over 20 mm bed of cement mortar 1:4 (10 sqm basis)",
        "unit": "10 sqm", "base_qty": 10.0,
        "rows": [
            ("Makrana White marble slabs 18mm", "Selected white marble slabs (area >= 0.30 sqm each) gang saw cut and semi-polished (11.50 sqm net incl. 15% cutting/dressing wastage)", "Conforming to IS:1130; white crystalline marble with natural plain grey veins", "Material", "11.50 sqm (1.150 sqm/sqm)", "10 sqm"),
            ("Carriage of Marble slabs", "Careful transport, crate handling and shifting of slabs to floor area", "Carriage allowance", "Material", "0.55 tonne", "10 sqm"),
            ("Portland Cement", "Hydraulic binder decomposed from 0.224 cum bed mortar 1:4 (REF#3.9) (0.085 t) + neat cement bedding/jointing slurry @ 5.0 kg/sqm (0.050 t) = 0.135 tonne", "OPC-43 Grade cement; bedding and slurry bond", "Material", "0.135 tonne (0.0135 t/sqm)", "10 sqm"),
            ("Coarse sand (zone III)", "Zone III coarse sand decomposed from 0.224 cum bed mortar 1:4 (REF#3.9) (0.224 x 1.07 = 0.240 cum)", "Clean natural river sand", "Material", "0.240 cum (0.0240 cum/sqm)", "10 sqm"),
            ("White cement & pigment for joints", "White Portland cement conforming to IS:8042 mixed with matching pigment for hairline joints (joint width <= 1.0mm)", "Joint grout", "Material", "0.010 tonne (0.0010 t/sqm)", "10 sqm"),
            ("Stone Mason (Marble layer)", "Spreading 20mm mortar bed, buttering slab back with slurry, laying marble true to slope, tapping with wooden mallet, and pointing joints", "Master marble layer; 1.25 days = 10.00 man-hrs per 10 sqm", "Labour", "10.00 man-hrs (1.00 sqm/man-hr)", "10 sqm"),
            ("Floor Polisher / Machine operator", "Machine grinding in 3 cuts with carborundum stones (80, 150, 320 grit), tin oxide / oxalic acid rubbing, and buffing to mirror finish", "Polishing technician; 1.50 days = 12.00 man-hrs per 10 sqm", "Labour", "12.00 man-hrs (0.83 sqm/man-hr)", "10 sqm"),
            ("Rotary floor polishing machine", "Electric rotary floor polishing machine with planetary grinding heads and water feed", "Polishing marble floor to mirror finish", "Machinery", "0.80 machine-day (12.5 sqm/mach-day)", "10 sqm"),
            ("Beldar", "Handling marble slabs, mixing bedding mortar, carrying slurry, removing slurry paste during grinding, and cleaning", "Labour crew; 2.20 days = 17.60 man-hrs per 10 sqm", "Labour", "17.60 man-hrs (0.57 sqm/man-hr)", "10 sqm"),
            ("Bhisti", "Water curing mortar bed and slabs for 14 days", "Watering crew; 0.90 day = 7.20 man-hrs per 10 sqm", "Labour", "7.20 man-hrs (1.39 sqm/man-hr)", "10 sqm"),
            ("Sundries & Polishing stones", "Carborundum blocks, oxalic acid, hessian cloth, wooden mallets, and spirit level", "Sundries allowance", "Equipment", "L.S. allowance", "10 sqm"),
            ("REF#3.9 (Cement mortar 1:4)", "Referenced bedding mortar: 0.224 cum CM 1:4 per 10 sqm marble flooring", "Mortar scope executed per CPWD REF#3.9 specification", "Reference", "0.224 cum (0.0224 cum/sqm)", "10 sqm"),
        ]
    },

    # ─── 11.26 KOTA STONE FLOORING ──────────────────────────────────────────────
    {
        "id": "11.26.1",
        "parent_title": "11.26 KOTA STONE SLAB FLOORING",
        "title": "11.26.1 Kota stone slab flooring 25 mm thick over 20 mm bed of cement mortar 1:4 (10 sqm basis)",
        "unit": "10 sqm", "base_qty": 10.0,
        "rows": [
            ("Kota stone slabs 25mm (semi-polished)", "Natural greenish blue / brown fine grained limestone slabs (size 55x55cm or 60x60cm) (11.50 sqm net incl. 15% wastage)", "Conforming to IS:1124 & IS:1130; hard, non-porous, wear resistant limestone", "Material", "11.50 sqm (1.150 sqm/sqm)", "10 sqm"),
            ("Carriage of Kota stone", "Haulage and handling of heavy Kota stone slabs (0.67 tonne per 10 sqm)", "Carriage allowance", "Material", "0.670 tonne", "10 sqm"),
            ("Portland Cement", "Hydraulic binder decomposed from 0.224 cum bed mortar 1:4 (REF#3.9) (0.085 t) + neat cement slurry @ 6.4 kg/sqm (0.064 t) = 0.149 tonne", "OPC-43 Grade cement; bed and slurry", "Material", "0.149 tonne (0.0149 t/sqm)", "10 sqm"),
            ("Coarse sand (zone III)", "Zone III sand decomposed from 0.224 cum bed mortar 1:4 (REF#3.9) (0.224 x 1.07 = 0.240 cum)", "Clean natural river sand", "Material", "0.240 cum (0.0240 cum/sqm)", "10 sqm"),
            ("Stone Mason (Kota stone layer)", "Laying 20mm mortar bed, slurry wash on base and slab back, laying Kota stone true to line/slope, jointing with grey cement slurry", "Skilled stone floor mason; 1.00 day = 8.00 man-hrs per 10 sqm", "Labour", "8.00 man-hrs (1.25 sqm/man-hr)", "10 sqm"),
            ("Floor Polisher / Machine operator", "Coarse grinding with carborundum stone (60-grit), pinhole filling, fine grinding (120-grit) and brown/green stone polishing", "Polishing technician; 1.40 days = 11.20 man-hrs per 10 sqm", "Labour", "11.20 man-hrs (0.89 sqm/man-hr)", "10 sqm"),
            ("Floor grinding machine", "Heavy electric terrazzo/stone grinding machine", "Grinding semi-polished Kota stone to uniform smooth surface", "Machinery", "0.80 machine-day (12.5 sqm/mach-day)", "10 sqm"),
            ("Beldar", "Carrying Kota slabs, mixing mortar, carrying cement slurry, scrubbing and scraping grinding sludge", "Labour crew; 2.00 days = 16.00 man-hrs per 10 sqm", "Labour", "16.00 man-hrs (0.625 sqm/man-hr)", "10 sqm"),
            ("Bhisti", "Water curing bed and stone floor for 14 days", "Watering crew; 0.90 day = 7.20 man-hrs per 10 sqm", "Labour", "7.20 man-hrs (1.39 sqm/man-hr)", "10 sqm"),
            ("Sundries & Carborundum blocks", "Carborundum blocks, oxalic acid rubbing powder, wooden mallets, and string lines", "Sundries allowance", "Equipment", "L.S. allowance", "10 sqm"),
            ("REF#3.9 (Cement mortar 1:4)", "Referenced bedding mortar: 0.224 cum CM 1:4 per 10 sqm Kota stone flooring", "Mortar scope executed per CPWD REF#3.9 specification", "Reference", "0.224 cum (0.0224 cum/sqm)", "10 sqm"),
        ]
    },

    # ─── 11.27 KOTA STONE IN RISERS & SKIRTING ──────────────────────────────────
    {
        "id": "11.27",
        "parent_title": "11.27 KOTA STONE IN RISERS, SKIRTING & DADO",
        "title": "11.27 Kota stone slabs 20 mm thick in risers of steps, skirting & dado over 12 mm bed of CM 1:3 (10 sqm basis)",
        "unit": "10 sqm", "base_qty": 10.0,
        "rows": [
            ("Kota stone slabs 20mm (pre-polished)", "Selected single piece Kota slabs for steps, risers and skirting up to 2m height (11.00 sqm net incl. 10% cutting wastage)", "Conforming to IS:1124; edges machine cut and top edge rounded / chamfered", "Material", "11.00 sqm (1.100 sqm/sqm)", "10 sqm"),
            ("Carriage of Kota stone", "Haulage and vertical shifting to floor levels", "Carriage allowance", "Material", "0.550 tonne", "10 sqm"),
            ("Portland Cement", "Hydraulic binder decomposed from 0.120 cum backing mortar 1:3 (REF#3.8) (0.061 t) + neat cement slurry @ 4.4 kg/sqm (0.044 t) = 0.105 tonne", "OPC-43 Grade cement; backing mortar and slurry coat", "Material", "0.105 tonne (0.0105 t/sqm)", "10 sqm"),
            ("Coarse sand (zone III)", "Zone III sand decomposed from 0.120 cum backing mortar 1:3 (REF#3.8) (0.120 x 1.07 = 0.128 cum)", "Clean natural river sand", "Material", "0.128 cum (0.0128 cum/sqm)", "10 sqm"),
            ("Stone Mason (Skirting specialist)", "Fixing slabs vertical against backing wall with mortar bed and slurry, aligning top edge with spirit level, and hairline joint pointing", "Skilled mason; 1.40 days = 11.20 man-hrs per 10 sqm", "Labour", "11.20 man-hrs (0.89 sqm/man-hr)", "10 sqm"),
            ("Beldar", "Hacking concrete/brick backing wall, mixing mortar, carrying slabs, and cleaning face", "Labour crew; 1.50 days = 12.00 man-hrs per 10 sqm", "Labour", "12.00 man-hrs (0.83 sqm/man-hr)", "10 sqm"),
            ("Bhisti", "Water spraying and curing skirting/risers for 14 days", "Watering crew; 0.60 day = 4.80 man-hrs per 10 sqm", "Labour", "4.80 man-hrs (2.08 sqm/man-hr)", "10 sqm"),
            ("Sundries & Edge polishing", "Nosing edge polishing, chamfering carborundum, and PoP temporary dots", "Sundries allowance", "Equipment", "L.S. allowance", "10 sqm"),
            ("REF#3.8 (Cement mortar 1:3)", "Referenced backing mortar: 0.120 cum CM 1:3 per 10 sqm skirting/risers", "Mortar scope executed per CPWD REF#3.8 specification", "Reference", "0.120 cum (0.0120 cum/sqm)", "10 sqm"),
        ]
    },

    # ─── 11.37 CERAMIC GLAZED TILES ─────────────────────────────────────────────
    {
        "id": "11.37",
        "parent_title": "11.37 CERAMIC GLAZED FLOOR TILES",
        "title": "11.37 Providing & laying Ceramic glazed floor tiles 300x300 mm over 20 mm bed of CM 1:4 with white cement pointing (10 sqm basis)",
        "unit": "10 sqm", "base_qty": 10.0,
        "rows": [
            ("Ceramic glazed floor tiles 300x300mm", "First quality glazed ceramic tiles conforming to IS:15622 (Group B IIa) of approved shade (10.25 sqm net incl. 2.5% wastage & breakage)", "Non-skid ceramic floor tiles; water absorption 3-6%, abrasion resistant", "Material", "10.25 sqm (1.025 sqm/sqm)", "10 sqm"),
            ("Carriage of Ceramic tiles", "Careful transport and carton handling of ceramic tiles", "Carriage allowance", "Material", "0.180 tonne", "10 sqm"),
            ("Portland Cement", "Hydraulic binder decomposed from 0.240 cum bed mortar 1:4 (REF#3.9) (0.091 t) + neat cement slurry @ 3.3 kg/sqm (0.033 t) = 0.124 tonne", "OPC-43 Grade cement; bedding and slurry coat", "Material", "0.124 tonne (0.0124 t/sqm)", "10 sqm"),
            ("Coarse sand (zone III)", "Zone III sand decomposed from 0.240 cum bed mortar 1:4 (REF#3.9) (0.240 x 1.07 = 0.257 cum)", "Clean natural river sand", "Material", "0.257 cum (0.0257 cum/sqm)", "10 sqm"),
            ("White Portland cement & pigment", "White cement (IS:8042) mixed with matching shade pigment for flush joint pointing (0.015 tonne)", "Tile joint grout", "Material", "0.015 tonne (0.0015 t/sqm)", "10 sqm"),
            ("Tile Mason (brick layer 1st class)", "Spreading 20mm mortar bed, pouring cement slurry, fixing tiles true to slope with spacer crosses, tapping with rubber mallet, and joint pointing", "Skilled tile mason; 2.00 days = 16.00 man-hrs per 10 sqm", "Labour", "16.00 man-hrs (0.625 sqm/man-hr)", "10 sqm"),
            ("Beldar", "Soaking tiles, mixing mortar 1:4, carrying tile boxes, wiping excess grout with sponge, and cleaning face", "Labour crew; 2.00 days = 16.00 man-hrs per 10 sqm", "Labour", "16.00 man-hrs (0.625 sqm/man-hr)", "10 sqm"),
            ("Bhisti", "Water curing mortar bed and tiled surface for 10 days", "Watering crew; 0.60 day = 4.80 man-hrs per 10 sqm", "Labour", "4.80 man-hrs (2.08 sqm/man-hr)", "10 sqm"),
            ("Sundries", "Plastic tile spacers (2mm), diamond wheel tile cutter, rubber floats, sponge, and masking tape", "Sundries allowance", "Equipment", "L.S. allowance", "10 sqm"),
            ("REF#3.9 (Cement mortar 1:4)", "Referenced bedding mortar: 0.240 cum CM 1:4 per 10 sqm ceramic tile flooring", "Mortar scope executed per CPWD REF#3.9 specification", "Reference", "0.240 cum (0.0240 cum/sqm)", "10 sqm"),
        ]
    },

    # ─── 11.41 VITRIFIED FLOOR TILES ────────────────────────────────────────────
    {
        "id": "11.41.2",
        "parent_title": "11.41 VITRIFIED FLOOR TILES",
        "title": "11.41.2 Providing & laying Vitrified floor tiles 600x600 mm over 20 mm bed of CM 1:4 with white cement pointing (10 sqm basis)",
        "unit": "10 sqm", "base_qty": 10.0,
        "rows": [
            ("Vitrified floor tiles 600x600mm", "First quality double charged / glazed vitrified tiles conforming to IS:15622 (water absorption < 0.08%) (10.25 sqm incl. 2.5% wastage)", "Full-body / nano-polished vitrified tiles; high breaking strength >= 1300 N, stain and scratch resistant", "Material", "10.25 sqm (1.025 sqm/sqm)", "10 sqm"),
            ("Carriage of Vitrified tiles", "Crate transport and handling to floor levels", "Carriage allowance", "Material", "0.220 tonne", "10 sqm"),
            ("Portland Cement", "Hydraulic binder decomposed from 0.240 cum bed mortar 1:4 (REF#3.9) (0.091 t) + neat cement slurry @ 3.3 kg/sqm (0.033 t) = 0.124 tonne", "OPC-43 Grade cement; bed and slurry", "Material", "0.124 tonne (0.0124 t/sqm)", "10 sqm"),
            ("Coarse sand (zone III)", "Zone III sand decomposed from 0.240 cum bed mortar 1:4 (REF#3.9) (0.240 x 1.07 = 0.257 cum)", "Clean natural river sand", "Material", "0.257 cum (0.0257 cum/sqm)", "10 sqm"),
            ("White cement & epoxy/matching grout", "White cement mixed with matching pigment or polymer grout for joint filling (joint width <= 2.0mm)", "Joint grout", "Material", "0.012 tonne (0.0012 t/sqm)", "10 sqm"),
            ("Tile Mason (vitrified specialist)", "Laying 20mm mortar bed, leveling, pouring cement slurry, laying large format 600x600mm tiles with suction cups, tapping with rubber mallet, aligning lippage", "Master tile installer; 2.00 days = 16.00 man-hrs per 10 sqm", "Labour", "16.00 man-hrs (0.625 sqm/man-hr)", "10 sqm"),
            ("Beldar", "Mixing mortar, carrying large tiles, cleaning joints with nylon brush, washing with damp sponge", "Labour crew; 2.00 days = 16.00 man-hrs per 10 sqm", "Labour", "16.00 man-hrs (0.625 sqm/man-hr)", "10 sqm"),
            ("Bhisti", "Water curing mortar bed and tiled surface for 10 days", "Watering crew; 0.60 day = 4.80 man-hrs per 10 sqm", "Labour", "4.80 man-hrs (2.08 sqm/man-hr)", "10 sqm"),
            ("Sundries", "Tile levelling clips & wedges, suction lifters, water-fed diamond blade cutter, and rubber floats", "Sundries allowance", "Equipment", "L.S. allowance", "10 sqm"),
            ("REF#3.9 (Cement mortar 1:4)", "Referenced bedding mortar: 0.240 cum CM 1:4 per 10 sqm vitrified tile flooring", "Mortar scope executed per CPWD REF#3.9 specification", "Reference", "0.240 cum (0.0240 cum/sqm)", "10 sqm"),
        ]
    },

    # ─── 11.56 POLISHED GRANITE STONE FLOORING ──────────────────────────────────
    {
        "id": "11.56",
        "parent_title": "11.56 POLISHED GRANITE STONE FLOORING",
        "title": "11.56 Providing & laying Polished Granite stone flooring 18 mm thick over 20 mm bed of CM 1:4 with pigment pointing (10 sqm basis)",
        "unit": "10 sqm", "base_qty": 10.0,
        "rows": [
            ("Polished Granite stone slabs 18mm", "Selected gang saw cut, mirror polished granite stone slabs of approved colour and pattern (11.00 sqm net incl. 10% cutting wastage)", "Premium granite conforming to IS:3316; compressive strength > 100 N/mm2, water absorption < 0.5%", "Material", "11.00 sqm (1.100 sqm/sqm)", "10 sqm"),
            ("Carriage of Granite slabs", "Careful transport and mechanical lifting of heavy granite slabs", "Carriage allowance", "Material", "0.580 tonne", "10 sqm"),
            ("Portland Cement", "Hydraulic binder decomposed from 0.240 cum bed mortar 1:4 (REF#3.9) (0.091 t) + neat cement slurry @ 4.4 kg/sqm (0.044 t) = 0.135 tonne", "OPC-43 Grade cement; bedding and slurry coat", "Material", "0.135 tonne (0.0135 t/sqm)", "10 sqm"),
            ("Coarse sand (zone III)", "Zone III sand decomposed from 0.240 cum bed mortar 1:4 (REF#3.9) (0.240 x 1.07 = 0.257 cum)", "Clean natural river sand", "Material", "0.257 cum (0.0257 cum/sqm)", "10 sqm"),
            ("Matching epoxy / pigment paste", "UV-resistant resin grout matching granite shade for joint pointing (joint <= 1.5mm)", "Joint grout", "Material", "2.00 kg (0.200 kg/sqm)", "10 sqm"),
            ("Stone Mason (Granite specialist)", "Spreading 20mm mortar bed, slurry wash, setting granite slabs true to pattern and slope, tapping with heavy rubber mallet, and flush pointing", "Master granite installer; 1.50 days = 12.00 man-hrs per 10 sqm", "Labour", "12.00 man-hrs (0.833 sqm/man-hr)", "10 sqm"),
            ("Floor Polisher / Machine operator", "Final edge rubbing, seamless joint levelling with diamond resin pads, and buffing to high gloss sheen", "Finishing specialist; 0.80 day = 6.40 man-hrs per 10 sqm", "Labour", "6.40 man-hrs (1.56 sqm/man-hr)", "10 sqm"),
            ("Beldar", "Handling heavy slabs with suction handles, mixing bedding mortar, carrying slurry, and cleaning face", "Labour crew; 2.20 days = 17.60 man-hrs per 10 sqm", "Labour", "17.60 man-hrs (0.568 sqm/man-hr)", "10 sqm"),
            ("Bhisti", "Water curing mortar bed and granite floor for 14 days", "Watering crew; 0.80 day = 6.40 man-hrs per 10 sqm", "Labour", "6.40 man-hrs (1.56 sqm/man-hr)", "10 sqm"),
            ("Sundries & Diamond pads", "Diamond resin polishing pads, rubber mallets, edge masking tape, and diamond saw blade cutting allowance", "Sundries allowance", "Equipment", "L.S. allowance", "10 sqm"),
            ("REF#3.9 (Cement mortar 1:4)", "Referenced bedding mortar: 0.240 cum CM 1:4 per 10 sqm granite flooring", "Mortar scope executed per CPWD REF#3.9 specification", "Reference", "0.240 cum (0.0240 cum/sqm)", "10 sqm"),
        ]
    },
]

def write_item_block(ws, item, start_row):
    r = start_row

    # Parent category banner
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
    c_parent = ws.cell(row=r, column=1, value=item["parent_title"])
    c_parent.font = Font(name="Calibri", size=11, bold=True, color=C_WHITE)
    c_parent.fill = fill(C_HDR_PARENT)
    c_parent.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[r].height = 24
    for col in range(1, 8):
        ws.cell(row=r, column=col).border = thin_border()
    r += 1

    # Item title banner
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
    c_sub = ws.cell(row=r, column=1, value=f"{item['id']}  —  {item['title']}")
    c_sub.font = Font(name="Calibri", size=10, bold=True, color=C_WHITE)
    c_sub.fill = fill(C_HDR_SUB)
    c_sub.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[r].height = 22
    for col in range(1, 8):
        ws.cell(row=r, column=col).border = thin_border()
    r += 1

    # Table Header Row (7 standard columns)
    headers = [
        "Item Code",
        "Labour / Machine / Material",
        "Work done",
        "Condition / When used",
        "Category",
        "Productivity",
        "Quantity"
    ]
    for col_idx, h in enumerate(headers, 1):
        c = ws.cell(row=r, column=col_idx, value=h)
        c.font = Font(name="Calibri", size=9, bold=True, color=C_WHITE)
        c.fill = fill(C_TH_BG)
        c.alignment = AL_C
        c.border = thin_border()
    ws.row_dimensions[r].height = 20
    r += 1

    # Data rows
    for row_idx, data in enumerate(item["rows"]):
        bg = C_ALT_ROW if (row_idx % 2 == 1) else C_WHITE
        row_fill = fill(bg)

        c1 = ws.cell(row=r, column=1, value=item["id"])
        c1.alignment = AL_C
        c1.font = Font(name="Calibri", size=9)
        c1.fill = row_fill
        c1.border = thin_border()

        c2 = ws.cell(row=r, column=2, value=data[0])
        c2.alignment = AL_L
        c2.font = Font(name="Calibri", size=9, bold=(data[3] in ["Labour", "Machinery"]))
        c2.fill = row_fill
        c2.border = thin_border()

        c3 = ws.cell(row=r, column=3, value=data[1])
        c3.alignment = AL_L
        c3.font = Font(name="Calibri", size=8.5)
        c3.fill = row_fill
        c3.border = thin_border()

        c4 = ws.cell(row=r, column=4, value=data[2])
        c4.alignment = AL_L
        c4.font = Font(name="Calibri", size=8.5)
        c4.fill = row_fill
        c4.border = thin_border()

        c5 = ws.cell(row=r, column=5, value=data[3])
        c5.alignment = AL_C
        c5.font = Font(name="Calibri", size=9)
        c5.fill = row_fill
        c5.border = thin_border()

        c6 = ws.cell(row=r, column=6, value=data[4])
        c6.alignment = AL_L
        c6.font = Font(name="Calibri", size=8.5)
        c6.fill = row_fill
        c6.border = thin_border()

        c7 = ws.cell(row=r, column=7, value=data[5])
        c7.alignment = AL_C
        c7.font = Font(name="Calibri", size=9)
        c7.fill = row_fill
        c7.border = thin_border()

        ws.row_dimensions[r].height = 22
        r += 1

    ws.row_dimensions[r].height = 8
    r += 1
    return r

def build_flooring_sheet(ws):
    if ws.views.sheetView:
        ws.views.sheetView[0].showGridLines = True

    # Sheet title spanning A1:G1
    ws.merge_cells("A1:G1")
    t1 = ws.cell(row=1, column=1, value="Sub-Head 11.0 — FLOORING  |  First-Principles Resource, Work & Gang Analysis")
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
    for item in FLOORING_ITEMS:
        current_row = write_item_block(ws, item, current_row)

    # Column widths for 7 columns (A to G)
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 36
    ws.column_dimensions["C"].width = 54
    ws.column_dimensions["D"].width = 52
    ws.column_dimensions["E"].width = 16
    ws.column_dimensions["F"].width = 28
    ws.column_dimensions["G"].width = 16

    ws.sheet_view.showGridLines = True
    ws.freeze_panes = "A4"

def rebuild_flooring_in_workbook(file_path):
    print(f"Opening {file_path} ...")
    wb = openpyxl.load_workbook(file_path)
    if "11_Flooring" not in wb.sheetnames:
        raise ValueError(f"'11_Flooring' sheet not found in {file_path}")

    pos = wb.sheetnames.index("11_Flooring")
    del wb["11_Flooring"]
    ws = wb.create_sheet("11_Flooring", pos)
    ws.sheet_properties.tabColor = "2E75B6"
    build_flooring_sheet(ws)

    temp_path = file_path.replace(".xlsx", "_TMP_FLOORING.xlsx")
    wb.save(temp_path)
    wb.close()
    os.replace(temp_path, file_path)
    print(f"Successfully updated {file_path} -> sheet '11_Flooring' ({len(FLOORING_ITEMS)} items).")

def main():
    repo_root = Path(__file__).resolve().parents[1]
    main_wb = repo_root / "CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_1_Latest.xlsx"
    output_wb = repo_root / "outputs" / "earthwork-custom-rate-composer" / "CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_1_Latest.xlsx"

    if main_wb.exists():
        rebuild_flooring_in_workbook(str(main_wb))
    if output_wb.exists():
        rebuild_flooring_in_workbook(str(output_wb))

if __name__ == "__main__":
    main()
