"""
support_builder_roofing.py
Rebuilds the existing '12_Roofing' worksheet with First-Principles Resource,
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

ROOFING_ITEMS = [
    # ─── 12.1 CORRUGATED G.S. SHEET ROOFING ────────────────────────────────────
    {
        "id": "12.1.1",
        "parent_title": "12.1 CORRUGATED GALVANISED STEEL (C.G.S.) SHEET ROOFING",
        "title": "12.1.1 Corrugated G.S. sheet roofing 1.00 mm thick with zinc coating 275 gm/sqm (10 sqm basis)",
        "unit": "10 sqm", "base_qty": 10.0,
        "rows": [
            ("Galvanised steel corrugated sheets 1.0mm", "CGS sheets with corrugation pitch 75mm and depth 19mm conforming to IS:277 Class 3 (1.26 quintals incl. 5% end/side lap wastage)", "Heavy-duty zinc coated corrugated roofing sheets; non-combustible weather barrier", "Material", "1.260 quintal (0.126 q/sqm)", "10 sqm"),
            ("Carriage of G.I. sheets", "Transport, loading, unloading and staging sheets onto roof trusses", "Carriage allowance", "Material", "0.126 tonne", "10 sqm"),
            ("G.I. 'J' or 'L' hooks 8mm & seam bolts 6mm", "Galvanized 8mm hook bolts with nuts, bitumen washers and GI limp washers (25 nos. hooks + 48 nos. seam bolts per 10 sqm)", "Fastening hardware conforming to IS:730", "Material", "73.00 nos (7.30 nos/sqm)", "10 sqm"),
            ("Blacksmith 1st class", "Drilling holes through sheet crowns with hand drill, fitting J-hooks to steel purlins, seam bolting side laps, and tightening nuts", "Skilled roofing craftsman; 0.40 day = 3.20 man-hrs per 10 sqm", "Labour", "3.20 man-hrs (3.125 sqm/man-hr)", "10 sqm"),
            ("Beldar", "Hoisting sheets to roof eaves, carrying along cat-ladders, positioning on purlins with 15cm end lap and 1.5 corrugation side lap", "Roofing handling crew; 0.60 day = 4.80 man-hrs per 10 sqm", "Labour", "4.80 man-hrs (2.083 sqm/man-hr)", "10 sqm"),
            ("Mistry", "Checking roof alignment, overhang at eaves (<= 300mm), camber and watertight seam fastening", "Supervisory roofer; 0.10 day = 0.80 man-hrs per 10 sqm", "Labour", "0.80 man-hrs (12.50 sqm/man-hr)", "10 sqm"),
            ("Sundries", "Hacksaw blades, bitumen washer mastic, safety roof ladders, scaffolding ropes, and safety harnesses", "Sundries allowance", "Equipment", "L.S. allowance", "10 sqm"),
        ]
    },

    # ─── 12.4 RIDGES AND HIPS IN G.S. SHEET ─────────────────────────────────────
    {
        "id": "12.4.1",
        "parent_title": "12.4 RIDGES OR HIPS IN PLAIN G.S. SHEET",
        "title": "12.4.1 Plain G.S. sheet ridges or hips 60 cm overall width, 0.80 mm thick with seam bolts & washers (10 m run basis)",
        "unit": "10 m", "base_qty": 10.0,
        "rows": [
            ("Galvanised steel plain sheets 0.80mm", "Plain zinc coated sheet cut and bent along ridge apex to 60 cm total girth with 22.5cm end laps (0.58 quintal incl. wastage)", "Conforming to IS:277 Class 3; watertight capping along roof ridge/hip line", "Material", "0.580 quintal (0.058 q/m)", "10 m"),
            ("Carriage of G.I. sheets", "Haulage and staging to roof ridge level", "Carriage allowance", "Material", "0.058 tonne", "10 m"),
            ("G.I. seam bolts 6x25mm & washers", "Galvanized seam bolts with bitumen washers and curved limp washers securing ridge wings to corrugated sheet crowns", "Connecting hardware", "Material", "28.00 nos (2.80 nos/m)", "10 m"),
            ("Blacksmith 1st class", "Bending ridge wings to roof pitch, positioning over apex, drilling through ridge and corrugation crowns, and bolt fixing", "Skilled craftsman; 1.20 days = 9.60 man-hrs per 10 m", "Labour", "9.60 man-hrs (1.042 m/man-hr)", "10 m"),
            ("Mistry", "Directing ridge line true to string, ensuring symmetric overhang on both roof slopes", "Roof foreman; 0.40 day = 3.20 man-hrs per 10 m", "Labour", "3.20 man-hrs (3.125 m/man-hr)", "10 m"),
            ("Beldar", "Passing ridge pieces along cat-ladders, holding in position against wind gusts during fastening", "Labour crew; 0.60 day = 4.80 man-hrs per 10 m", "Labour", "4.80 man-hrs (2.083 m/man-hr)", "10 m"),
            ("Sundries", "Bituminous sealing compound along laps, drill bits, and safety lines", "Sundries allowance", "Equipment", "L.S. allowance", "10 m"),
        ]
    },

    # ─── 12.7 SEMI-CIRCULAR G.S. SHEET EAVES GUTTERS ────────────────────────────
    {
        "id": "12.7.1",
        "parent_title": "12.7 PLAIN G.S. SHEET EAVES GUTTERS",
        "title": "12.7.1 Semi-circular plain G.S. sheet gutters 15 cm wide, 45 cm girth, 0.80 mm thick with 50x3 mm brackets (10 m run basis)",
        "unit": "10 m", "base_qty": 10.0,
        "rows": [
            ("Galvanised steel plain sheets 0.80mm", "Cold-formed semi-circular gutter trough with beaded edges (45 cm total developed width) (0.35 quintal)", "Conforming to IS:277; rainwater drainage trough along roof eaves", "Material", "0.350 quintal (0.035 q/m)", "10 m"),
            ("M.S. flat iron brackets 50x3mm", "Galvanized flat iron brackets bent to gutter radius and screwed to wooden/steel fascia or rafters at 1.2m spacing", "Structural support brackets", "Material", "0.080 quintal (0.008 q/m)", "10 m"),
            ("G.I. bolts, nuts & seam rivets", "6mm seam bolts, copper/tin rivets and soldered joints at gutter section overlaps", "Fasteners", "Material", "25.00 nos (2.50 nos/m)", "10 m"),
            ("Carriage of G.I. materials", "Haulage of gutter sections and brackets to site", "Carriage allowance", "Material", "0.043 tonne", "10 m"),
            ("Blacksmith 1st class", "Fixing brackets to 1 in 100 slope, bedding gutter in brackets, lapping 100mm, seam riveting, and soldering watertight joints", "Skilled gutter craftsman; 0.55 day = 4.40 man-hrs per 10 m", "Labour", "4.40 man-hrs (2.27 m/man-hr)", "10 m"),
            ("Beldar", "Erecting scaffolding along eaves, hoisting gutter lengths, holding during bracket clamping", "Labour crew; 0.55 day = 4.40 man-hrs per 10 m", "Labour", "4.40 man-hrs (2.27 m/man-hr)", "10 m"),
            ("Sundries", "Soldering zinc chloride flux, tin-lead solder wire, blow lamp gas, and zinc chromate touch-up", "Sundries allowance", "Equipment", "L.S. allowance", "10 m"),
        ]
    },

    # ─── 12.16 MUD PHASKA TERRACING ─────────────────────────────────────────────
    {
        "id": "12.16",
        "parent_title": "12.16 MUD PHASKA TERRACING WITH TILE PAVING",
        "title": "12.16 Mud phaska 10 cm thick with damped brick earth, gobri leaping & flat tile paving grouted in CM 1:3 (10 sqm basis)",
        "unit": "10 sqm", "base_qty": 10.0,
        "rows": [
            ("Selected brick earth", "Clean cohesive earth free from kankar, soluble salts and organic roots (1.25 cum incl. compaction shrinkage)", "Soil conforming to CPWD specification; thermal insulation layer on flat roof", "Material", "1.250 cum (0.125 cum/sqm)", "10 sqm"),
            ("Bhusa (cut straw)", "Chopped wheat/paddy straw mixed @ 35 kg per cum of earth for reinforcement and cracking resistance (44 kg)", "Fibrous binder", "Material", "0.440 quintal (0.044 q/sqm)", "10 sqm"),
            ("Gobri leaping (clay & cowdung mix)", "Plastering 25mm mud mortar with cowdung slurry 1:1 forming smooth uniform bed", "Bedding slurry", "Material", "L.S. batch", "10 sqm"),
            ("Flat tile bricks (FPStiles)", "Clay roofing tile bricks (size 229x114x38mm) laid flat with 6mm joints (320 nos. per 10 sqm)", "Conforming to IS:2690; durable weather and wear barrier over mud phaska", "Material", "320.00 nos (32.00 nos/sqm)", "10 sqm"),
            ("Portland Cement", "Hydraulic binder decomposed from 0.035 cum joint grouting mortar 1:3 (REF#3.8) (0.018 t)", "OPC-43 Grade cement; waterproof joint grouting", "Material", "0.018 tonne (0.0018 t/sqm)", "10 sqm"),
            ("Fine sand", "Clean fine sand decomposed from 0.035 cum grouting mortar 1:3 (REF#3.8) (0.037 cum)", "Fine aggregate", "Material", "0.037 cum (0.0037 cum/sqm)", "10 sqm"),
            ("Integral waterproofing compound", "Powder/liquid waterproofing admixture mixed @ 2% by weight of cement (0.36 kg)", "Conforming to IS:2645", "Material", "0.360 kg (0.036 kg/sqm)", "10 sqm"),
            ("Mason (brick layer)", "Leveling mud phaska to slope (1 in 40), laying tile bricks true to line, and flush grouting joints with cement mortar 1:3", "Skilled mason; 0.70 day = 5.60 man-hrs per 10 sqm", "Labour", "5.60 man-hrs (1.79 sqm/man-hr)", "10 sqm"),
            ("Beldar", "Mixing earth with bhusa, rotting for 7 days, carrying to roof, ramming with thappies to 10cm density, and shifting tiles", "Labour crew; 2.20 days = 17.60 man-hrs per 10 sqm", "Labour", "17.60 man-hrs (0.57 sqm/man-hr)", "10 sqm"),
            ("Bhisti", "Watering earth during pugging and curing tiled terrace for 10 days", "Watering crew; 0.80 day = 6.40 man-hrs per 10 sqm", "Labour", "6.40 man-hrs (1.56 sqm/man-hr)", "10 sqm"),
            ("REF#3.8 (Cement mortar 1:3)", "Referenced joint grouting mortar: 0.035 cum CM 1:3 per 10 sqm terrace", "Mortar scope executed per CPWD REF#3.8 specification", "Reference", "0.035 cum (0.0035 cum/sqm)", "10 sqm"),
        ]
    },

    # ─── 12.21 CEMENT CONCRETE GOLA ─────────────────────────────────────────────
    {
        "id": "12.21",
        "parent_title": "12.21 CEMENT CONCRETE GOLA AT ROOF-WALL JUNCTION",
        "title": "12.21 Providing gola 75x75 mm in cement concrete 1:2:4 finished with cement mortar 1:3 (10 m run basis)",
        "unit": "10 m", "base_qty": 10.0,
        "rows": [
            ("Portland Cement", "Hydraulic binder decomposed from 0.056 cum CC 1:2:4 (REF#4.1.3) (0.018 t) + CM 1:3 finish (0.005 t) = 0.023 tonne", "OPC-43 Grade cement; triangular gola filling fillet angle", "Material", "0.023 tonne (0.0023 t/m)", "10 m"),
            ("Coarse sand (zone III)", "Zone III coarse sand decomposed from 0.056 cum CC 1:2:4 (REF#4.1.3) (0.025 cum) + finishing mortar (0.005 cum) = 0.030 cum", "Clean natural river sand", "Material", "0.030 cum (0.0030 cum/m)", "10 m"),
            ("Stone aggregate 10mm and down", "Broken stone aggregate decomposed from 0.056 cum CC 1:2:4 (REF#4.1.3) (0.051 cum)", "Coarse aggregate conforming to IS:383", "Material", "0.051 cum (0.0051 cum/m)", "10 m"),
            ("Carriage of Materials", "Haulage of materials to terrace junction", "Carriage allowance", "Material", "0.08 tonne equiv", "10 m"),
            ("Mason (brick layer 1st class)", "Raking parapet masonry groove 75x75mm, laying concrete 1:2:4 triangular fillet, and finishing with curved steel float with CM 1:3", "Skilled mason; 0.30 day = 2.40 man-hrs per 10 m", "Labour", "2.40 man-hrs (4.17 m/man-hr)", "10 m"),
            ("Beldar", "Hacking wall angle, mixing concrete and mortar batches, and cleaning work area", "Labour crew; 0.30 day = 2.40 man-hrs per 10 m", "Labour", "2.40 man-hrs (4.17 m/man-hr)", "10 m"),
            ("Bhisti", "Water spraying and curing gola for 10 days", "Watering crew; 0.15 day = 1.20 man-hrs per 10 m", "Labour", "1.20 man-hrs (8.33 m/man-hr)", "10 m"),
            ("REF#4.1.3 (Cement concrete 1:2:4)", "Referenced concrete: 0.056 cum CC 1:2:4 per 10 m gola", "Concrete scope executed per CPWD REF#4.1.3 specification", "Reference", "0.056 cum (0.0056 cum/m)", "10 m"),
            ("REF#3.8 (Cement mortar 1:3)", "Referenced finish: 0.005 cum CM 1:3 per 10 m gola", "Mortar scope executed per CPWD REF#3.8 specification", "Reference", "0.005 cum (0.0005 cum/m)", "10 m"),
        ]
    },

    # ─── 12.22 MAKING KHURRAS FOR RAINWATER OUTLETS ─────────────────────────────
    {
        "id": "12.22",
        "parent_title": "12.22 RAINWATER KHURRAS AT PARAPET OUTLETS",
        "title": "12.22 Making khurras 45x45 cm with 5 cm CC 1:2:4 over PVC sheet 400 micron & CM 1:3 plaster (10 nos basis)",
        "unit": "10 nos", "base_qty": 10.0,
        "rows": [
            ("P.V.C. sheet 400 micron (1m x 1m)", "Waterproof unplasticised PVC membrane 400 micron thickness laid under concrete khurra (10 sqm for 10 nos)", "Impervious flashing membrane preventing seepage around spout opening", "Material", "10.00 sqm (1.000 sqm/khurra)", "10 nos"),
            ("Portland Cement", "Hydraulic binder decomposed from 0.100 cum CC 1:2:4 (REF#4.1.3) (0.032 t) + 12mm CM 1:3 plaster & neat cement coat (0.012 t) = 0.044 tonne", "OPC-43 Grade cement; depression base and smooth throat", "Material", "0.044 tonne (0.0044 t/khurra)", "10 nos"),
            ("Coarse sand (zone III)", "Zone III sand decomposed from 0.100 cum CC 1:2:4 (REF#4.1.3) (0.044 cum) + plaster (0.015 cum) = 0.059 cum", "Clean natural river sand", "Material", "0.059 cum (0.0059 cum/khurra)", "10 nos"),
            ("Stone aggregate 20mm & 10mm", "Graded coarse stone aggregate decomposed from 0.100 cum CC 1:2:4 (REF#4.1.3) (0.089 cum)", "Coarse aggregate conforming to IS:383", "Material", "0.089 cum (0.0089 cum/khurra)", "10 nos"),
            ("Carriage of Materials", "Haulage of materials to rainwater spout locations", "Carriage allowance", "Material", "0.15 tonne equiv", "10 nos"),
            ("Mason (brick layer 1st class)", "Forming 45x45 cm depression sloped towards spout pipe, laying CC 1:2:4, applying 12mm CM 1:3 plaster, rounding corners with neat cement trowelling", "Skilled mason; 0.40 day = 3.20 man-hrs per 10 nos", "Labour", "3.20 man-hrs (3.125 nos/man-hr)", "10 nos"),
            ("Beldar", "Cutting depression in roof slab/parapet base, mixing concrete and mortar, and carrying pans", "Labour crew; 0.60 day = 4.80 man-hrs per 10 nos", "Labour", "4.80 man-hrs (2.083 nos/man-hr)", "10 nos"),
            ("Bhisti", "Water curing khurras for 10 days", "Watering crew; 0.20 day = 1.60 man-hrs per 10 nos", "Labour", "1.60 man-hrs (6.25 nos/man-hr)", "10 nos"),
            ("REF#4.1.3 (Cement concrete 1:2:4)", "Referenced concrete: 0.100 cum CC 1:2:4 per 10 khurras", "Concrete scope executed per CPWD REF#4.1.3 specification", "Reference", "0.100 cum (0.010 cum/khurra)", "10 nos"),
            ("REF#3.8 (Cement mortar 1:3)", "Referenced plaster: 0.015 cum CM 1:3 per 10 khurras", "Mortar scope executed per CPWD REF#3.8 specification", "Reference", "0.015 cum (0.0015 cum/khurra)", "10 nos"),
        ]
    },

    # ─── 12.41 RIGID PVC RAIN WATER PIPES ───────────────────────────────────────
    {
        "id": "12.41.2",
        "parent_title": "12.41 RIGID PVC RAIN WATER PIPES",
        "title": "12.41.2 Providing & fixing on wall face unplasticised Rigid PVC rain water pipes 110 mm dia (10 m run basis)",
        "unit": "10 m", "base_qty": 10.0,
        "rows": [
            ("uPVC single socketed rain water pipe 110mm", "Unplasticised PVC pipe (working pressure 4 kg/sqcm) conforming to IS:13592 Type A for rainwater drainage (10 m run)", "UV-stabilized rigid PVC pipe; corrosion-free smooth hydraulic flow", "Material", "10.00 metre (1.000 m/m)", "10 m"),
            ("EPDM rubber sealing ring 110mm", "Elastomeric sealing rings for socket expansion joints accommodating thermal expansion", "Joint seal", "Material", "2.00 nos (0.20 no/m)", "10 m"),
            ("uPVC pipe clips & holder bats", "Heavy-duty uPVC pipe spacer clips anchored to wall with wooden plugs and brass screws at 1.8m centres", "Fixing hardware", "Material", "6.00 nos (0.60 no/m)", "10 m"),
            ("Carriage of pipes & accessories", "Haulage and staging of light uPVC pipes to wall elevations", "Carriage allowance", "Material", "10.00 metre", "10 m"),
            ("Fitter (grade 1)", "Setting out vertical drop with plumb line, drilling wall plug holes, fixing clips, lubricating socket rings, pushing pipes home, checking verticality", "Skilled pipe installer; 0.38 day = 3.04 man-hrs per 10 m", "Labour", "3.04 man-hrs (3.29 m/man-hr)", "10 m"),
            ("Beldar", "Erecting staging/scaffolding, holding pipes during clip fastening, assisting installer", "Labour crew; 0.75 day = 6.00 man-hrs per 10 m", "Labour", "6.00 man-hrs (1.67 m/man-hr)", "10 m"),
            ("Bandhani", "Rigging safety scaffolding and fall arrest systems on external building facades", "Scaffolding specialist; 0.18 day = 1.44 man-hrs per 10 m", "Labour", "1.44 man-hrs (6.94 m/man-hr)", "10 m"),
            ("Sundries", "PVC solvent cement, lubricant grease, screws, rawl plugs, and scaffolding ropes", "Sundries allowance", "Equipment", "L.S. allowance", "10 m"),
        ]
    },

    # ─── 12.50 PRECOATED GALVANISED PROFILE SHEETS (PPGI) ───────────────────────
    {
        "id": "12.50",
        "parent_title": "12.50 PRECOATED GALVANISED IRON (PPGI) PROFILE SHEETS",
        "title": "12.50 Providing & fixing precoated galvanised iron (PPGI) profile sheets 0.50 mm thick with self-drilling screws (10 sqm basis)",
        "unit": "10 sqm", "base_qty": 10.0,
        "rows": [
            ("Precoated Galvanised Iron profile sheets 0.50mm", "PPGI trapezoidal profile sheets 0.50 mm TCT, steel grade 240 MPa, zinc 120 g/sqm, 5-7 micron epoxy primer, 18 micron polyester coat with 25 micron guard film (10.50 sqm incl. 5% lap wastage)", "Conforming to IS:277 & IS:14246; modern aesthetic roofing with high tensile strength and UV color retention", "Material", "10.50 sqm (1.050 sqm/sqm)", "10 sqm"),
            ("Carriage of Profile sheets", "Careful transport and mechanical crane/winch staging of long sheets up to 12m length", "Carriage allowance", "Material", "0.065 tonne", "10 sqm"),
            ("Self-drilling / self-tapping screws 5.5x55mm", "Corrosion-resistant hex-head self-drilling screws with bonded EPDM sealing washers (60 nos. per 10 sqm)", "Fasteners conforming to AS:3566 Class 3", "Material", "60.00 nos (6.00 nos/sqm)", "10 sqm"),
            ("Carpenter / Fitter", "Aligning profile sheets on steel purlins, power screwing fasteners at crests/troughs with electric torque driver, cutting flashing rebates", "Skilled roofing technician; 0.30 day = 2.40 man-hrs per 10 sqm", "Labour", "2.40 man-hrs (4.17 sqm/man-hr)", "10 sqm"),
            ("Beldar", "Hoisting continuous length sheets, carrying with edge gloves, peeling protective film, and holding sheets during fastening", "Roofing crew; 0.60 day = 4.80 man-hrs per 10 sqm", "Labour", "4.80 man-hrs (2.08 sqm/man-hr)", "10 sqm"),
            ("Sundries", "Magnetic nut setters, silicone neutral sealant at end overlaps, safety lifelines, and edge protection netting", "Sundries allowance", "Equipment", "L.S. allowance", "10 sqm"),
        ]
    },

    # ─── 12.51 PRECOATED STEEL SHEET RIDGES & FLASHINGS ─────────────────────────
    {
        "id": "12.51.1",
        "parent_title": "12.51 PRECOATED STEEL SHEET RIDGES & FLASHINGS",
        "title": "12.51.1 Precoated galvanised steel plain ridges (500-600 mm) 0.50 mm thick with self-drilling fasteners (10 m run basis)",
        "unit": "10 m", "base_qty": 10.0,
        "rows": [
            ("Precoated galvanised steel plain ridges 0.50mm", "Factory folded 500-600 mm wide plain ridge matching profile sheet colour and coating (10.5 m run incl. 5% overlap wastage)", "Conforming to IS:14246; weather-sealed apex capping", "Material", "10.50 metre (1.050 m/m)", "10 m"),
            ("Self-drilling screws with EPDM washers", "Fasteners fixing ridge apron to crest of profile sheets at every alternate pitch", "Fasteners", "Material", "30.00 nos (3.00 nos/m)", "10 m"),
            ("Carriage of Ridge pieces", "Haulage and staging to roof ridge line", "Carriage allowance", "Material", "0.035 tonne", "10 m"),
            ("Carpenter / Fitter", "Setting ridge line, profiling closures to fit trapezoidal flute gaps, driving self-tapping fasteners, and sealing with butyl tape", "Skilled technician; 0.27 day = 2.16 man-hrs per 10 m", "Labour", "2.16 man-hrs (4.63 m/man-hr)", "10 m"),
            ("Beldar", "Passing ridge pieces along apex catwalks, holding firmly in wind during screw driving", "Labour crew; 0.80 day = 6.40 man-hrs per 10 m", "Labour", "6.40 man-hrs (1.56 m/man-hr)", "10 m"),
            ("Sundries", "Closed-cell profile foam fillers, butyl sealant tape, and touch-up aerosol paint", "Sundries allowance", "Equipment", "L.S. allowance", "10 m"),
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

def build_roofing_sheet(ws):
    if ws.views.sheetView:
        ws.views.sheetView[0].showGridLines = True

    # Sheet title spanning A1:G1
    ws.merge_cells("A1:G1")
    t1 = ws.cell(row=1, column=1, value="Sub-Head 12.0 — ROOFING  |  First-Principles Resource, Work & Gang Analysis")
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
    for item in ROOFING_ITEMS:
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

def rebuild_roofing_in_workbook(file_path):
    print(f"Opening {file_path} ...")
    wb = openpyxl.load_workbook(file_path)
    if "12_Roofing" not in wb.sheetnames:
        raise ValueError(f"'12_Roofing' sheet not found in {file_path}")

    pos = wb.sheetnames.index("12_Roofing")
    del wb["12_Roofing"]
    ws = wb.create_sheet("12_Roofing", pos)
    ws.sheet_properties.tabColor = "2E75B6"
    build_roofing_sheet(ws)

    temp_path = file_path.replace(".xlsx", "_TMP_ROOFING.xlsx")
    wb.save(temp_path)
    wb.close()
    os.replace(temp_path, file_path)
    print(f"Successfully updated {file_path} -> sheet '12_Roofing' ({len(ROOFING_ITEMS)} items).")

def main():
    repo_root = Path(__file__).resolve().parents[1]
    main_wb = repo_root / "CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_1_Latest.xlsx"
    output_wb = repo_root / "outputs" / "earthwork-custom-rate-composer" / "CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_1_Latest.xlsx"

    if main_wb.exists():
        rebuild_roofing_in_workbook(str(main_wb))
    if output_wb.exists():
        rebuild_roofing_in_workbook(str(output_wb))

if __name__ == "__main__":
    main()
