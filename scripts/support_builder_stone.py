"""
support_builder_stone.py
Rebuilds the existing '07_Stone_Work' worksheet with First-Principles Resource,
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

STONE_ITEMS = [
    # ─── 7.1 RANDOM RUBBLE (R.R.) MASONRY IN FOUNDATION & PLINTH ────────────────
    {
        "id": "7.1.1",
        "parent_title": "7.1 RANDOM RUBBLE (R.R.) MASONRY IN FOUNDATION & PLINTH",
        "title": "7.1.1 Random rubble masonry with hard stone in foundation and plinth in cement mortar 1:6",
        "unit": "cum", "base_qty": 1.0,
        "rows": [
            ("Hard stone boulders / rubble", "Sound, tough, unweathered angular quartzite/granite stones from quarry (1.00 cum net solid quarry volume)", "Selected hard stone rubble with flat bed faces; free from cracks, decay, and clay coatings", "Material", "1.00 cum (1.000 cum/cum)", "1 cum"),
            ("Bond stone / through stones", "Long header stones extending at least 2/3rd through wall thickness or full width for walls <= 40 cm (0.07 cum)", "Conforming to CPWD specification; provided @ 1 bond stone every 0.5 sqm face area", "Material", "0.07 cum (0.070 cum/cum)", "1 cum"),
            ("Stone spalls / chips", "Small angular stone fragments used to wedge and pack spaces between rubble boulders in hearting (0.09 cum)", "Clean stone spalls broken from parent stone during dressing", "Material", "0.09 cum (0.090 cum/cum)", "1 cum"),
            ("Carriage of Stone & Boulder", "Mechanical haulage, unloading and stacking of rubble stones near foundation trenches (1.16 cum)", "Transport from quarry to site storage yard", "Material", "1.16 cum (1.160 cum/cum)", "1 cum"),
            ("Portland Cement", "Hydraulic binder decomposed from 0.33 cum of cement mortar 1:6 (REF#3.11) (0.083 tonne = 1.66 bags)", "OPC-43 Grade cement; 0.33 cum mortar x 0.25 t/cum = 0.0825 t", "Material", "0.083 tonne (0.083 t/cum)", "1 cum"),
            ("Coarse sand", "Zone III coarse sand decomposed from 0.33 cum of cement mortar 1:6 (REF#3.11) (0.353 cum)", "Clean natural river sand Zone III; 0.33 cum mortar x 1.07 cum/cum = 0.353 cum", "Material", "0.353 cum (0.353 cum/cum)", "1 cum"),
            ("Carriage of Cement & Sand", "Haulage of cement and sand to mortar mixing platform near trench", "Carriage allowance decomposed from REF#3.11", "Material", "0.43 cum/t equiv", "1 cum"),
            ("Mechanical Mortar Mixer", "Machine batching and mixing 1:6 mortar paste (decomposed from REF#3.11)", "Mixer operational hire; 0.33 cum mortar x 0.15 hr/cum = 0.050 hr", "Machine", "0.050 machine-hrs (20.0 cum/hr)", "1 cum"),
            ("Stone Mason 1st & 2nd class", "Hammer dressing face stones, knocking off bush projections, bedding in thick mortar, and placing through stones", "Skilled stone mason crew; 1.50 day = 12.00 man-hrs per cum", "Labour", "12.00 man-hrs (0.08 cum/man-hr)", "1 cum"),
            ("Beldar", "Breaking spalls, washing quarry dust from stones, mixing mortar, and filling hearting spaces", "Labour crew; 1.80 day = 14.40 man-hrs per cum", "Labour", "14.40 man-hrs (0.07 cum/man-hr)", "1 cum"),
            ("Coolie", "Hauling heavy boulders and bond stones from stacks to foundation trench; carrying mortar pans", "Heavy haulage crew; 0.90 day = 7.20 man-hrs per cum", "Labour", "7.20 man-hrs (0.14 cum/man-hr)", "1 cum"),
            ("Bhisti", "Watering stones before laying, filling water barrels, and curing massive stone walls for 14 days", "Watering crew; 0.70 day = 5.60 man-hrs per cum", "Labour", "5.60 man-hrs (0.18 cum/man-hr)", "1 cum"),
            ("Sundries", "Spall hammers, scabbing picks, mason lines, plumb bobs, and crowbars", "Sundries allowance (L.S. 8.19 x cost index)", "Equipment", "L.S. allowance", "1 cum"),
            ("REF#3.11 (Cement mortar 1:6)", "Referenced intermediate mortar: 0.33 cum cement mortar 1:6 per 1 cum random rubble masonry", "Mortar scope executed per CPWD REF#3.11 specification", "Reference", "0.33 cum (0.330 cum/cum)", "1 cum"),
        ]
    },

    # ─── 7.2 RANDOM RUBBLE (R.R.) MASONRY IN SUPERSTRUCTURE ─────────────────────
    {
        "id": "7.2.1",
        "parent_title": "7.2 RANDOM RUBBLE (R.R.) MASONRY IN SUPERSTRUCTURE UP TO FLOOR V",
        "title": "7.2.1 Random rubble masonry with hard stone in superstructure above plinth up to floor V level in cement mortar 1:6",
        "unit": "cum", "base_qty": 1.0,
        "rows": [
            ("Hard stone boulders / rubble", "Sound quarry stones selected for superstructure walls (1.00 cum solid)", "Selected hard stone rubble; uniform color tone and sound structure", "Material", "1.00 cum (1.000 cum/cum)", "1 cum"),
            ("Bond stone / through stones", "Header stones extending full wall thickness or overlapping headers (0.07 cum)", "Conforming to CPWD specification; provided @ 1 per 0.5 sqm face", "Material", "0.07 cum (0.070 cum/cum)", "1 cum"),
            ("Stone spalls / chips", "Angular stone fragments for packing interior voids (0.09 cum)", "Clean stone spalls broken from parent stone", "Material", "0.09 cum (0.090 cum/cum)", "1 cum"),
            ("Carriage of Stone & Boulder", "Haulage of stone boulders to site storage yard (1.16 cum)", "Carriage from quarry to site", "Material", "1.16 cum (1.160 cum/cum)", "1 cum"),
            ("Portland Cement", "Hydraulic binder decomposed from 0.33 cum of cement mortar 1:6 (REF#3.11) (0.083 tonne = 1.66 bags)", "OPC-43 Grade cement; 0.33 cum mortar x 0.25 t/cum = 0.0825 t", "Material", "0.083 tonne (0.083 t/cum)", "1 cum"),
            ("Coarse sand", "Zone III coarse sand decomposed from 0.33 cum of cement mortar 1:6 (REF#3.11) (0.353 cum)", "Clean natural river sand Zone III", "Material", "0.353 cum (0.353 cum/cum)", "1 cum"),
            ("Carriage of Cement & Sand", "Haulage of cement and sand to mortar mixing platform", "Carriage allowance", "Material", "0.43 cum/t equiv", "1 cum"),
            ("Mechanical Mortar Mixer", "Machine batching and mixing 1:6 mortar paste", "Mixer operational hire; output = 20.0 cum/hr", "Machine", "0.050 machine-hrs", "1 cum"),
            ("Heavy double scaffolding", "Heavy-duty steel tubular scaffolding with safety platforms up to floor V level", "Scaffolding hire, erection and dismantling for heavy stone masonry", "Equipment", "Scaffolding allowance", "1 cum"),
            ("Stone Mason 1st & 2nd class", "Hammer dressing face stones, checking vertical plumb, maintaining corner quoin stones", "Skilled stone mason crew; 1.65 day = 13.20 man-hrs per cum", "Labour", "13.20 man-hrs (0.08 cum/man-hr)", "1 cum"),
            ("Beldar", "Washing stones, mixing mortar, hoisting stone buckets, and serving masons on staging", "Labour crew; 2.20 day = 17.60 man-hrs per cum", "Labour", "17.60 man-hrs (0.06 cum/man-hr)", "1 cum"),
            ("Coolie", "Hoisting heavy stones and mortar pans with gin wheels/ropes to upper floor levels up to floor V", "Hoisting crew; 1.30 day = 10.40 man-hrs per cum", "Labour", "10.40 man-hrs (0.10 cum/man-hr)", "1 cum"),
            ("Bhisti", "Watering stones, wet curing vertical stone walls with hessian wraps for 14 days", "Watering crew; 0.80 day = 6.40 man-hrs per cum", "Labour", "6.40 man-hrs (0.16 cum/man-hr)", "1 cum"),
            ("Sundries", "Gin wheels, heavy hemp ropes, stone lifting clamps, crowbars, and plumb bobs", "Sundries allowance (L.S. 10.92 x cost index)", "Equipment", "L.S. allowance", "1 cum"),
            ("REF#3.11 (Cement mortar 1:6)", "Referenced intermediate mortar: 0.33 cum cement mortar 1:6 per 1 cum superstructure masonry", "Mortar scope executed per CPWD REF#3.11 specification", "Reference", "0.33 cum (0.330 cum/cum)", "1 cum"),
        ]
    },

    # ─── 7.6 & 7.7 COURSED RUBBLE (C.R.) MASONRY ────────────────────────────────
    {
        "id": "7.6.1",
        "parent_title": "7.6 COURSED RUBBLE (C.R.) MASONRY (FIRST SORT) IN FOUNDATION & PLINTH",
        "title": "7.6.1 Coursed rubble masonry (first sort) with hard stone in foundation and plinth in cement mortar 1:6",
        "unit": "cum", "base_qty": 1.0,
        "rows": [
            ("Hard stone boulders (first sort)", "Quarry stones hammer dressed on beds and joints to give truly level courses 15-30 cm high (1.21 cum net)", "First sort C.R. masonry: face stones hammer dressed, bushings not exceeding 25 mm", "Material", "1.21 cum (1.210 cum/cum)", "1 cum"),
            ("Through stones / bond stones", "Full-width bond stones dressed to course height (0.07 cum)", "Conforming to CPWD specification; provided @ 1 bond stone per 0.5 sqm face area", "Material", "0.07 cum (0.070 cum/cum)", "1 cum"),
            ("Stone spalls / chips", "Angular chips for packing bed voids and hearting (0.09 cum)", "Clean stone spalls", "Material", "0.09 cum (0.090 cum/cum)", "1 cum"),
            ("Carriage of Stone & Boulder", "Mechanical haulage and unloading of dressed stones at site (1.37 cum)", "Carriage allowance", "Material", "1.37 cum (1.370 cum/cum)", "1 cum"),
            ("Portland Cement", "Hydraulic binder decomposed from 0.30 cum of cement mortar 1:6 (REF#3.11) (0.075 tonne = 1.50 bags)", "OPC-43 Grade cement; 0.30 cum mortar x 0.25 t/cum = 0.075 t", "Material", "0.075 tonne (0.075 t/cum)", "1 cum"),
            ("Coarse sand", "Zone III coarse sand decomposed from 0.30 cum of cement mortar 1:6 (REF#3.11) (0.321 cum)", "Clean natural river sand Zone III", "Material", "0.321 cum (0.321 cum/cum)", "1 cum"),
            ("Carriage of Cement & Sand", "Haulage of materials to mixing platform", "Carriage allowance", "Material", "0.40 cum/t equiv", "1 cum"),
            ("Mechanical Mortar Mixer", "Machine batching and mixing 1:6 mortar paste", "Mixer operational hire; output = 22.2 cum/hr", "Machine", "0.045 machine-hrs", "1 cum"),
            ("Stone Mason 1st & 2nd class", "Careful dressing, setting stones in horizontal courses of uniform height, striking joints", "Skilled stone mason; 1.80 day = 14.40 man-hrs per cum", "Labour", "14.40 man-hrs (0.07 cum/man-hr)", "1 cum"),
            ("Beldar", "Chiseling dressing arrises, washing stones, mixing mortar, and bedding spalls", "Labour crew; 1.90 day = 15.20 man-hrs per cum", "Labour", "15.20 man-hrs (0.07 cum/man-hr)", "1 cum"),
            ("Coolie", "Hauling dressed course stones from dressing yard to foundation trench", "Handling crew; 0.95 day = 7.60 man-hrs per cum", "Labour", "7.60 man-hrs (0.13 cum/man-hr)", "1 cum"),
            ("Bhisti", "Watering stones and curing thick walls for 14 days", "Watering crew; 0.70 day = 5.60 man-hrs per cum", "Labour", "5.60 man-hrs (0.18 cum/man-hr)", "1 cum"),
            ("Sundries", "Dressing hammers, chisels, straight edges, spirit levels, and plumb bobs", "Sundries allowance (L.S. 9.55 x cost index)", "Equipment", "L.S. allowance", "1 cum"),
            ("REF#3.11 (Cement mortar 1:6)", "Referenced intermediate mortar: 0.30 cum cement mortar 1:6 per 1 cum coursed rubble masonry", "Mortar scope executed per CPWD REF#3.11 specification", "Reference", "0.30 cum (0.300 cum/cum)", "1 cum"),
        ]
    },
    {
        "id": "7.7.1",
        "parent_title": "7.7 COURSED RUBBLE (C.R.) MASONRY (SECOND SORT) IN FOUNDATION & PLINTH",
        "title": "7.7.1 Coursed rubble masonry (second sort) with hard stone in foundation and plinth in cement mortar 1:6",
        "unit": "cum", "base_qty": 1.0,
        "rows": [
            ("Hard stone boulders (second sort)", "Quarry stones dressed to rough level courses; two stones allowed to make up course height (1.10 cum)", "Second sort C.R. masonry: bush projections <= 40 mm, stones less finely dressed", "Material", "1.10 cum (1.100 cum/cum)", "1 cum"),
            ("Through stones / bond stones", "Full-width bond stones (0.07 cum)", "Conforming to CPWD specification; provided @ 1 bond stone per 0.5 sqm face", "Material", "0.07 cum (0.070 cum/cum)", "1 cum"),
            ("Stone spalls / chips", "Angular chips for packing bed voids and hearting (0.09 cum)", "Clean stone spalls", "Material", "0.09 cum (0.090 cum/cum)", "1 cum"),
            ("Carriage of Stone & Boulder", "Mechanical haulage and unloading of stones (1.26 cum)", "Carriage allowance", "Material", "1.26 cum (1.260 cum/cum)", "1 cum"),
            ("Portland Cement", "Hydraulic binder decomposed from 0.30 cum of cement mortar 1:6 (REF#3.11) (0.075 tonne = 1.50 bags)", "OPC-43 Grade cement; 0.30 cum mortar x 0.25 t/cum = 0.075 t", "Material", "0.075 tonne (0.075 t/cum)", "1 cum"),
            ("Coarse sand", "Zone III coarse sand decomposed from 0.30 cum of cement mortar 1:6 (REF#3.11) (0.321 cum)", "Clean natural river sand Zone III", "Material", "0.321 cum (0.321 cum/cum)", "1 cum"),
            ("Carriage of Cement & Sand", "Haulage of materials to mixing platform", "Carriage allowance", "Material", "0.40 cum/t equiv", "1 cum"),
            ("Mechanical Mortar Mixer", "Machine batching and mixing 1:6 mortar paste", "Mixer operational hire; output = 22.2 cum/hr", "Machine", "0.045 machine-hrs", "1 cum"),
            ("Stone Mason 1st & 2nd class", "Hammer dressing, setting stones in level courses, and filling joint spaces", "Skilled stone mason; 1.65 day = 13.20 man-hrs per cum", "Labour", "13.20 man-hrs (0.08 cum/man-hr)", "1 cum"),
            ("Beldar", "Breaking spalls, washing stones, mixing mortar, and bedding stones", "Labour crew; 1.80 day = 14.40 man-hrs per cum", "Labour", "14.40 man-hrs (0.07 cum/man-hr)", "1 cum"),
            ("Coolie", "Hauling stones from dressing yard to trench", "Handling crew; 0.90 day = 7.20 man-hrs per cum", "Labour", "7.20 man-hrs (0.14 cum/man-hr)", "1 cum"),
            ("Bhisti", "Watering stones and curing masonry for 14 days", "Watering crew; 0.70 day = 5.60 man-hrs per cum", "Labour", "5.60 man-hrs (0.18 cum/man-hr)", "1 cum"),
            ("Sundries", "Dressing hammers, chisels, straight edges, and plumb bobs", "Sundries allowance", "Equipment", "L.S. allowance", "1 cum"),
            ("REF#3.11 (Cement mortar 1:6)", "Referenced intermediate mortar: 0.30 cum cement mortar 1:6 per 1 cum coursed rubble masonry", "Mortar scope executed per CPWD REF#3.11 specification", "Reference", "0.30 cum (0.300 cum/cum)", "1 cum"),
        ]
    },
    {
        "id": "7.8.1",
        "parent_title": "7.8 COURSED RUBBLE (C.R.) MASONRY IN SUPERSTRUCTURE",
        "title": "7.8.1 Coursed rubble masonry (first sort) with hard stone in superstructure up to floor V in cement mortar 1:6",
        "unit": "cum", "base_qty": 1.0,
        "rows": [
            ("Hard stone boulders (first sort)", "Quarry stones dressed to uniform course height (1.21 cum net)", "First sort C.R. masonry for superstructure walls up to floor V level", "Material", "1.21 cum (1.210 cum/cum)", "1 cum"),
            ("Through stones / bond stones", "Full-width bond stones dressed to course height (0.07 cum)", "Provided @ 1 bond stone per 0.5 sqm face", "Material", "0.07 cum (0.070 cum/cum)", "1 cum"),
            ("Stone spalls / chips", "Angular chips for packing interior voids (0.09 cum)", "Clean stone spalls", "Material", "0.09 cum (0.090 cum/cum)", "1 cum"),
            ("Carriage of Stone & Boulder", "Mechanical carriage of dressed stones (1.37 cum)", "Carriage allowance", "Material", "1.37 cum (1.370 cum/cum)", "1 cum"),
            ("Portland Cement", "Hydraulic binder decomposed from 0.30 cum of cement mortar 1:6 (REF#3.11) (0.075 tonne = 1.50 bags)", "OPC-43 Grade cement; 0.30 cum mortar x 0.25 t/cum = 0.075 t", "Material", "0.075 tonne (0.075 t/cum)", "1 cum"),
            ("Coarse sand", "Zone III coarse sand decomposed from 0.30 cum of cement mortar 1:6 (REF#3.11) (0.321 cum)", "Clean natural river sand Zone III", "Material", "0.321 cum (0.321 cum/cum)", "1 cum"),
            ("Carriage of Cement & Sand", "Haulage of materials to mixing platform", "Carriage allowance", "Material", "0.40 cum/t equiv", "1 cum"),
            ("Mechanical Mortar Mixer", "Machine batching and mixing 1:6 mortar paste", "Mixer operational hire; output = 22.2 cum/hr", "Machine", "0.045 machine-hrs", "1 cum"),
            ("Heavy double scaffolding", "Heavy-duty steel scaffolding up to floor V level", "Scaffolding hire allowance", "Equipment", "Scaffolding allowance", "1 cum"),
            ("Stone Mason 1st & 2nd class", "Setting dressed stones in true courses, checking face plumb, maintaining straight bond", "Skilled stone mason; 1.95 day = 15.60 man-hrs per cum", "Labour", "15.60 man-hrs (0.06 cum/man-hr)", "1 cum"),
            ("Beldar", "Washing stones, mixing mortar, hoisting stones, and serving masons", "Labour crew; 2.30 day = 18.40 man-hrs per cum", "Labour", "18.40 man-hrs (0.05 cum/man-hr)", "1 cum"),
            ("Coolie", "Hoisting heavy dressed stones and mortar pans to staging up to floor V", "Hoisting crew; 1.40 day = 11.20 man-hrs per cum", "Labour", "11.20 man-hrs (0.09 cum/man-hr)", "1 cum"),
            ("Bhisti", "Watering stones and curing masonry walls for 14 days", "Watering crew; 0.80 day = 6.40 man-hrs per cum", "Labour", "6.40 man-hrs (0.16 cum/man-hr)", "1 cum"),
            ("Sundries", "Gin wheels, heavy ropes, chisels, straight edges, and spirit levels", "Sundries allowance (L.S. 12.28 x cost index)", "Equipment", "L.S. allowance", "1 cum"),
            ("REF#3.11 (Cement mortar 1:6)", "Referenced intermediate mortar: 0.30 cum cement mortar 1:6 per 1 cum superstructure coursed rubble masonry", "Mortar scope executed per CPWD REF#3.11 specification", "Reference", "0.30 cum (0.300 cum/cum)", "1 cum"),
        ]
    },

    # ─── 7.12 ASHLAR STONE WORK ─────────────────────────────────────────────────
    {
        "id": "7.12.1.1",
        "parent_title": "7.12 STONE WORK IN PLAIN ASHLAR UP TO FLOOR V LEVEL",
        "title": "7.12.1.1 Stone work in plain ashlar one face dressed with Red sand stone up to floor V level (10 cudm / 0.01 cum)",
        "unit": "cudm", "base_qty": 10.0,
        "rows": [
            ("Red sand stone rough blocks", "Quarry cut red sand stone blocks dressed on exposed face with fine chisel dressing (1.333 cudm per 10 cudm finished)", "Conforming to IS:3622 / IS:1121; fine chisel dressed face, true sharp arrises, beds plane and square", "Material", "1.333 cudm (0.133 cudm/cudm)", "10 cudm"),
            ("Carriage of Red sand stone", "Haulage and careful handling of sand stone blocks to site dressing yard", "Carriage allowance", "Material", "0.031 tonne", "10 cudm"),
            ("Portland Cement", "Hydraulic binder decomposed from thin jointing mortar 1:3 (REF#3.8) and neat cement slurry backing", "OPC-43 Grade cement; thin joint bedding and pointing", "Material", "0.003 tonne", "10 cudm"),
            ("Marble dust / fine sand", "Fine marble dust / sand for non-staining thin ashlar joints (joint thickness <= 5 mm)", "Clean fine filler", "Material", "0.003 cum", "10 cudm"),
            ("Stone Mason (Ashlar carver / fixer)", "Fine chisel dressing (chisel draft 25 mm wide around face), setting to razor line, plumbing, and pointing", "Master ashlar mason; 0.088 day = 0.70 man-hrs per 10 cudm", "Labour", "0.70 man-hrs (14.3 cudm/man-hr)", "10 cudm"),
            ("Coolie / Helper", "Handling dressed ashlar blocks with canvas slings to avoid chipping corners", "Handling crew; 0.015 day = 0.12 man-hrs per 10 cudm", "Labour", "0.12 man-hrs (83.3 cudm/man-hr)", "10 cudm"),
            ("Bhisti", "Moist curing ashlar joint pointing", "Watering allowance; 0.005 day = 0.04 man-hrs", "Labour", "0.04 man-hrs", "10 cudm"),
            ("Sundries", "Point chisels, claw tools, carborundum stones for edge rubbing, lead plugs, and brass cramps", "Sundries allowance (L.S. 2.73 x cost index)", "Equipment", "L.S. allowance", "10 cudm"),
        ]
    },
    {
        "id": "7.12.1.2",
        "parent_title": "7.12 STONE WORK IN PLAIN ASHLAR UP TO FLOOR V LEVEL",
        "title": "7.12.1.2 Stone work in plain ashlar one face dressed with White sand stone up to floor V level (10 cudm / 0.01 cum)",
        "unit": "cudm", "base_qty": 10.0,
        "rows": [
            ("White sand stone rough blocks", "Quarry cut white sand stone blocks fine chisel dressed on face (1.333 cudm per 10 cudm finished)", "Conforming to IS:3622; uniform white/cream texture, fine chisel dressed face, square beds", "Material", "1.333 cudm (0.133 cudm/cudm)", "10 cudm"),
            ("Carriage of White sand stone", "Haulage and careful handling of white sand stone blocks", "Carriage allowance", "Material", "0.031 tonne", "10 cudm"),
            ("White Portland Cement", "White cement for non-staining jointing mortar 1:3 (REF#3.16)", "White cement conforming to IS:8042", "Material", "0.003 tonne", "10 cudm"),
            ("White marble dust", "Fine white marble dust for thin white ashlar joints", "Clean white marble dust", "Material", "0.003 cum", "10 cudm"),
            ("Stone Mason (Ashlar carver / fixer)", "Fine chisel dressing face, setting to razor line, plumbing, and white cement pointing", "Master ashlar mason; 0.088 day = 0.70 man-hrs per 10 cudm", "Labour", "0.70 man-hrs (14.3 cudm/man-hr)", "10 cudm"),
            ("Coolie / Helper", "Handling dressed white ashlar blocks with clean padding", "Handling crew; 0.015 day = 0.12 man-hrs per 10 cudm", "Labour", "0.12 man-hrs (83.3 cudm/man-hr)", "10 cudm"),
            ("Bhisti", "Moist curing ashlar joint pointing", "Watering allowance", "Labour", "0.04 man-hrs", "10 cudm"),
            ("Sundries", "Point chisels, claw tools, carborundum rubbing stones, brass cramps, and protective wrapping", "Sundries allowance", "Equipment", "L.S. allowance", "10 cudm"),
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

def build_stone_sheet(ws):
    if ws.views.sheetView:
        ws.views.sheetView[0].showGridLines = True

    # Sheet title spanning A1:G1
    ws.merge_cells("A1:G1")
    t1 = ws.cell(row=1, column=1, value="Sub-Head 7.0 — STONE WORK  |  First-Principles Resource, Work & Gang Analysis")
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
    for item in STONE_ITEMS:
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

def rebuild_stone_in_workbook(file_path):
    print(f"Opening {file_path} ...")
    wb = openpyxl.load_workbook(file_path)
    if "07_Stone_Work" not in wb.sheetnames:
        raise ValueError(f"'07_Stone_Work' sheet not found in {file_path}")

    pos = wb.sheetnames.index("07_Stone_Work")
    del wb["07_Stone_Work"]
    ws = wb.create_sheet("07_Stone_Work", pos)
    ws.sheet_properties.tabColor = "2E75B6"
    build_stone_sheet(ws)

    temp_path = file_path.replace(".xlsx", "_TMP_STONE.xlsx")
    wb.save(temp_path)
    wb.close()
    os.replace(temp_path, file_path)
    print(f"Successfully updated {file_path} -> sheet '07_Stone_Work' ({len(STONE_ITEMS)} items).")

def main():
    repo_root = Path(__file__).resolve().parents[1]
    main_wb = repo_root / "CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_1_Latest.xlsx"
    output_wb = repo_root / "outputs" / "earthwork-custom-rate-composer" / "CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_1_Latest.xlsx"

    if main_wb.exists():
        rebuild_stone_in_workbook(str(main_wb))
    if output_wb.exists():
        rebuild_stone_in_workbook(str(output_wb))

if __name__ == "__main__":
    main()
