"""
support_builder_masonry.py
Rebuilds the existing '06_Masonry_Work' worksheet with First-Principles Resource,
Productivity, Gang, Machinery and Material Analysis per CPWD DAR 2019 Vol 1:
  Item Code | Labour / Machine / Material | Work done | Condition / When used | Category | Productivity | Quantity
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from pathlib import Path
import os, sys, shutil

# Colours matching standard
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

MASONRY_ITEMS = [
    # ─── 6.1 BRICK WORK IN FOUNDATION & PLINTH (F.P.S. BRICKS) ─────────────────
    {
        "id": "6.1.1",
        "parent_title": "6.1 BRICK WORK IN FOUNDATION & PLINTH WITH F.P.S. BRICKS",
        "title": "6.1.1 Brick work in foundation and plinth with F.P.S. bricks class 7.5 in cement mortar 1:4",
        "unit": "cum", "base_qty": 1.0,
        "rows": [
            ("Common burnt clay F.P.S. bricks class 7.5", "Standard burnt clay building bricks conforming to IS:1077 (500 nos. per 1 cum finished masonry; 494 net + wastage)", "Class 7.5 bricks having minimum compressive strength 7.5 N/mm2; soaked in water prior to laying", "Material", "494.0 nos (494 nos/cum)", "1 cum"),
            ("Carriage of Bricks", "Mechanical transport, loading, unloading and careful stacking in stacks of 1000/2000 at site", "Carriage from brick kiln to site unloading area", "Material", "494.0 nos (494 nos/cum)", "1 cum"),
            ("Portland Cement", "Hydraulic binder decomposed from 0.25 cum of cement mortar 1:4 (REF#3.9) (0.095 tonne = 1.9 bags)", "OPC-43 Grade cement conforming to IS:269; 0.25 cum mortar x 0.38 t/cum = 0.095 t", "Material", "0.095 tonne (0.095 t/cum)", "1 cum"),
            ("Coarse sand", "Zone III coarse sand decomposed from 0.25 cum of cement mortar 1:4 (REF#3.9) (0.268 cum)", "Clean natural river coarse sand; 0.25 cum mortar x 1.07 cum/cum = 0.268 cum", "Material", "0.268 cum (0.268 cum/cum)", "1 cum"),
            ("Carriage of Cement & Sand", "Haulage of cement and sand to mortar mixing platform near foundation trenches", "Carriage allowance decomposed from REF#3.9", "Material", "0.36 cum/t equiv", "1 cum"),
            ("Mechanical Mortar Mixer", "Machine batching and mixing 1:4 mortar paste (decomposed from REF#3.9)", "Mixer operational hire; 0.25 cum mortar x 0.15 hr/cum = 0.038 hr", "Machine", "0.038 machine-hrs (26.3 cum/hr)", "1 cum"),
            ("Mason (brick layer) 1st & 2nd class", "Laying bricks in English bond, true to line, plumb and level; maintaining uniform 10-12 mm joint thickness", "Skilled brick layer; 0.85 day = 6.80 man-hrs per cum", "Labour", "6.80 man-hrs (0.15 cum/man-hr)", "1 cum"),
            ("Beldar", "Soaking bricks in water tanks, transporting bricks to trench, and feeding mortar to masons", "Labour crew; 0.90 day = 7.20 man-hrs per cum", "Labour", "7.20 man-hrs (0.14 cum/man-hr)", "1 cum"),
            ("Coolie", "Manual carriage of bricks and mortar along foundation trenches; shifting mortar pans", "Unskilled helper; 0.45 day = 3.60 man-hrs per cum", "Labour", "3.60 man-hrs (0.28 cum/man-hr)", "1 cum"),
            ("Bhisti", "Flooding soaking tanks, filling brick wetting trenches, and curing finished brickwork for 14 days", "Watering crew; 0.40 day = 3.20 man-hrs per cum", "Labour", "3.20 man-hrs (0.31 cum/man-hr)", "1 cum"),
            ("Sundries", "Plumb bobs, spirit levels, mason lines, mortar boards, and soak pit maintenance", "Sundries allowance (L.S. 5.46 x cost index)", "Equipment", "L.S. allowance", "1 cum"),
            ("REF#3.9 (Cement mortar 1:4)", "Referenced intermediate mortar: 0.25 cum cement mortar 1:4 per 1 cum brickwork", "Mortar scope executed per CPWD REF#3.9 specification", "Reference", "0.25 cum (0.250 cum/cum)", "1 cum"),
        ]
    },
    {
        "id": "6.1.2",
        "parent_title": "6.1 BRICK WORK IN FOUNDATION & PLINTH WITH F.P.S. BRICKS",
        "title": "6.1.2 Brick work in foundation and plinth with F.P.S. bricks class 7.5 in cement mortar 1:6",
        "unit": "cum", "base_qty": 1.0,
        "rows": [
            ("Common burnt clay F.P.S. bricks class 7.5", "Standard burnt clay building bricks conforming to IS:1077 (494 net + wastage)", "Class 7.5 bricks; soaked in water prior to laying in foundation steps", "Material", "494.0 nos (494 nos/cum)", "1 cum"),
            ("Carriage of Bricks", "Mechanical haulage and stacking of bricks at site", "Carriage from kiln to site", "Material", "494.0 nos (494 nos/cum)", "1 cum"),
            ("Portland Cement", "Hydraulic binder decomposed from 0.25 cum of cement mortar 1:6 (REF#3.11) (0.063 tonne = 1.25 bags)", "OPC-43 Grade cement; 0.25 cum mortar x 0.25 t/cum = 0.063 t", "Material", "0.063 tonne (0.063 t/cum)", "1 cum"),
            ("Coarse sand", "Zone III coarse sand decomposed from 0.25 cum of cement mortar 1:6 (REF#3.11) (0.268 cum)", "Clean natural river sand; 0.25 cum mortar x 1.07 cum/cum = 0.268 cum", "Material", "0.268 cum (0.268 cum/cum)", "1 cum"),
            ("Carriage of Cement & Sand", "Carriage of cement and sand to mortar mixing platform", "Carriage allowance decomposed from REF#3.11", "Material", "0.33 cum/t equiv", "1 cum"),
            ("Mechanical Mortar Mixer", "Machine mixing 1:6 mortar paste (decomposed from REF#3.11)", "Mixer hire; 0.25 cum mortar x 0.15 hr/cum = 0.038 hr", "Machine", "0.038 machine-hrs (26.3 cum/hr)", "1 cum"),
            ("Mason (brick layer) 1st & 2nd class", "Laying bricks in English bond in foundation steps, checking plumb and alignment", "Skilled brick layer; 0.85 day = 6.80 man-hrs per cum", "Labour", "6.80 man-hrs (0.15 cum/man-hr)", "1 cum"),
            ("Beldar", "Soaking bricks, wheeling mortar, and serving bricklayers", "Labour crew; 0.90 day = 7.20 man-hrs per cum", "Labour", "7.20 man-hrs (0.14 cum/man-hr)", "1 cum"),
            ("Coolie", "Hauling bricks and mortar along foundation trenches", "Unskilled helper; 0.45 day = 3.60 man-hrs per cum", "Labour", "3.60 man-hrs (0.28 cum/man-hr)", "1 cum"),
            ("Bhisti", "Tank soaking and foundation masonry curing for 14 days", "Watering crew; 0.40 day = 3.20 man-hrs per cum", "Labour", "3.20 man-hrs (0.31 cum/man-hr)", "1 cum"),
            ("Sundries", "Plumb bobs, spirit levels, mason lines, mortar boards, and soak pit maintenance", "Sundries allowance", "Equipment", "L.S. allowance", "1 cum"),
            ("REF#3.11 (Cement mortar 1:6)", "Referenced intermediate mortar: 0.25 cum cement mortar 1:6 per 1 cum brickwork", "Mortar scope executed per CPWD REF#3.11 specification", "Reference", "0.25 cum (0.250 cum/cum)", "1 cum"),
        ]
    },

    # ─── 6.2 BRICK WORK IN FOUNDATION & PLINTH (MODULAR BRICKS) ─────────────────
    {
        "id": "6.2.1",
        "parent_title": "6.2 BRICK WORK IN FOUNDATION & PLINTH WITH MODULAR BRICKS",
        "title": "6.2.1 Brick work in foundation and plinth with Modular bricks class 7.5 in cement mortar 1:4",
        "unit": "cum", "base_qty": 1.0,
        "rows": [
            ("Common burnt clay Modular bricks class 7.5", "Modular burnt clay bricks (190x90x90 mm nominal size per IS:1077) (487 net + wastage)", "Standard modular size saving 12% mortar volume compared to non-modular bricks", "Material", "487.0 nos (487 nos/cum)", "1 cum"),
            ("Carriage of Bricks", "Mechanical carriage and careful stacking of modular bricks at site", "Carriage from kiln to site", "Material", "487.0 nos (487 nos/cum)", "1 cum"),
            ("Portland Cement", "Hydraulic binder decomposed from 0.22 cum of cement mortar 1:4 (REF#3.9) (0.084 tonne = 1.67 bags)", "OPC-43 Grade cement; 0.22 cum mortar x 0.38 t/cum = 0.084 t", "Material", "0.084 tonne (0.084 t/cum)", "1 cum"),
            ("Coarse sand", "Zone III coarse sand decomposed from 0.22 cum of cement mortar 1:4 (REF#3.9) (0.235 cum)", "Clean natural river sand; 0.22 cum mortar x 1.07 cum/cum = 0.235 cum", "Material", "0.235 cum (0.235 cum/cum)", "1 cum"),
            ("Carriage of Cement & Sand", "Haulage of materials to mortar platform", "Carriage allowance", "Material", "0.32 cum/t equiv", "1 cum"),
            ("Mechanical Mortar Mixer", "Machine mixing 1:4 mortar paste", "Mixer operational hire; 0.22 cum mortar x 0.15 hr/cum = 0.033 hr", "Machine", "0.033 machine-hrs (30.3 cum/hr)", "1 cum"),
            ("Mason (brick layer) 1st & 2nd class", "Laying modular bricks in true courses, checking bond and uniform 10 mm joints", "Skilled brick layer; 0.85 day = 6.80 man-hrs per cum", "Labour", "6.80 man-hrs (0.15 cum/man-hr)", "1 cum"),
            ("Beldar", "Soaking modular bricks in tanks and feeding mortar", "Labour crew; 0.90 day = 7.20 man-hrs per cum", "Labour", "7.20 man-hrs (0.14 cum/man-hr)", "1 cum"),
            ("Coolie", "Hauling modular bricks and mortar pans", "Unskilled helper; 0.45 day = 3.60 man-hrs per cum", "Labour", "3.60 man-hrs (0.28 cum/man-hr)", "1 cum"),
            ("Bhisti", "Tank soaking and curing masonry for 14 days", "Watering crew; 0.40 day = 3.20 man-hrs per cum", "Labour", "3.20 man-hrs (0.31 cum/man-hr)", "1 cum"),
            ("Sundries", "Plumb bobs, mason lines, and mortar boards", "Sundries allowance", "Equipment", "L.S. allowance", "1 cum"),
            ("REF#3.9 (Cement mortar 1:4)", "Referenced intermediate mortar: 0.22 cum cement mortar 1:4 per 1 cum modular brickwork", "Mortar scope executed per CPWD REF#3.9 specification", "Reference", "0.22 cum (0.220 cum/cum)", "1 cum"),
        ]
    },
    {
        "id": "6.2.2",
        "parent_title": "6.2 BRICK WORK IN FOUNDATION & PLINTH WITH MODULAR BRICKS",
        "title": "6.2.2 Brick work in foundation and plinth with Modular bricks class 7.5 in cement mortar 1:6",
        "unit": "cum", "base_qty": 1.0,
        "rows": [
            ("Common burnt clay Modular bricks class 7.5", "Modular burnt clay bricks (190x90x90 mm) (487 net + wastage)", "Standard modular size conforming to IS:1077", "Material", "487.0 nos (487 nos/cum)", "1 cum"),
            ("Carriage of Bricks", "Mechanical carriage of modular bricks to site", "Carriage from kiln to site", "Material", "487.0 nos (487 nos/cum)", "1 cum"),
            ("Portland Cement", "Hydraulic binder decomposed from 0.22 cum of cement mortar 1:6 (REF#3.11) (0.055 tonne = 1.10 bags)", "OPC-43 Grade cement; 0.22 cum mortar x 0.25 t/cum = 0.055 t", "Material", "0.055 tonne (0.055 t/cum)", "1 cum"),
            ("Coarse sand", "Zone III coarse sand decomposed from 0.22 cum of cement mortar 1:6 (REF#3.11) (0.235 cum)", "Clean natural river sand; 0.22 cum mortar x 1.07 cum/cum = 0.235 cum", "Material", "0.235 cum (0.235 cum/cum)", "1 cum"),
            ("Carriage of Cement & Sand", "Haulage of materials to mortar platform", "Carriage allowance", "Material", "0.29 cum/t equiv", "1 cum"),
            ("Mechanical Mortar Mixer", "Machine mixing 1:6 mortar paste", "Mixer operational hire; output = 30.3 cum/hr", "Machine", "0.033 machine-hrs (30.3 cum/hr)", "1 cum"),
            ("Mason (brick layer) 1st & 2nd class", "Laying modular bricks in English bond in foundation steps", "Skilled brick layer; 0.85 day = 6.80 man-hrs per cum", "Labour", "6.80 man-hrs (0.15 cum/man-hr)", "1 cum"),
            ("Beldar", "Soaking modular bricks in tanks and feeding mortar", "Labour crew; 0.90 day = 7.20 man-hrs per cum", "Labour", "7.20 man-hrs (0.14 cum/man-hr)", "1 cum"),
            ("Coolie", "Hauling modular bricks and mortar pans", "Unskilled helper; 0.45 day = 3.60 man-hrs per cum", "Labour", "3.60 man-hrs (0.28 cum/man-hr)", "1 cum"),
            ("Bhisti", "Tank soaking and curing masonry for 14 days", "Watering crew; 0.40 day = 3.20 man-hrs per cum", "Labour", "3.20 man-hrs (0.31 cum/man-hr)", "1 cum"),
            ("Sundries", "Plumb bobs, mason lines, and mortar boards", "Sundries allowance", "Equipment", "L.S. allowance", "1 cum"),
            ("REF#3.11 (Cement mortar 1:6)", "Referenced intermediate mortar: 0.22 cum cement mortar 1:6 per 1 cum modular brickwork", "Mortar scope executed per CPWD REF#3.11 specification", "Reference", "0.22 cum (0.220 cum/cum)", "1 cum"),
        ]
    },

    # ─── 6.4 BRICK WORK IN SUPERSTRUCTURE UP TO FLOOR V LEVEL ───────────────────
    {
        "id": "6.4.1",
        "parent_title": "6.4 BRICK WORK IN SUPERSTRUCTURE UP TO FLOOR V LEVEL (F.P.S. BRICKS)",
        "title": "6.4.1 Brick work in superstructure above plinth up to floor V level with F.P.S. bricks in cement mortar 1:4",
        "unit": "cum", "base_qty": 1.0,
        "rows": [
            ("Common burnt clay F.P.S. bricks class 7.5", "Standard bricks for superstructure walls up to floor V level (494 net + wastage)", "Clean selected bricks with sharp arrises and true faces", "Material", "494.0 nos (494 nos/cum)", "1 cum"),
            ("Carriage of Bricks", "Haulage and stacking of bricks at site", "Carriage from kiln to site", "Material", "494.0 nos (494 nos/cum)", "1 cum"),
            ("Portland Cement", "Hydraulic binder decomposed from 0.25 cum of cement mortar 1:4 (REF#3.9) (0.095 tonne = 1.9 bags)", "OPC-43 Grade cement; 0.25 cum mortar x 0.38 t/cum = 0.095 t", "Material", "0.095 tonne (0.095 t/cum)", "1 cum"),
            ("Coarse sand", "Zone III coarse sand decomposed from 0.25 cum of cement mortar 1:4 (REF#3.9) (0.268 cum)", "Clean natural river sand Zone III", "Material", "0.268 cum (0.268 cum/cum)", "1 cum"),
            ("Carriage of Cement & Sand", "Haulage of cement and sand to mortar mixing platform", "Carriage allowance", "Material", "0.36 cum/t equiv", "1 cum"),
            ("Mechanical Mortar Mixer", "Machine batching and mixing mortar paste", "Mixer operational hire; output = 26.3 cum/hr", "Machine", "0.038 machine-hrs (26.3 cum/hr)", "1 cum"),
            ("Double steel / timber scaffolding", "Independent double scaffolding with steel tubular pipes, clamps and wooden planks up to floor V", "Scaffolding hire, erection, safety guardrails, and dismantling", "Equipment", "Scaffolding allowance", "1 cum"),
            ("Mason (brick layer) 1st & 2nd class", "Laying bricks in superstructure walls, plumb and line, raking joints 10mm for plaster key", "Skilled brick layer; 0.95 day = 7.60 man-hrs per cum", "Labour", "7.60 man-hrs (0.13 cum/man-hr)", "1 cum"),
            ("Beldar", "Soaking bricks, hoisting mortar to staging, serving masons, and cleaning scaffold", "Labour crew; 1.20 day = 9.60 man-hrs per cum", "Labour", "9.60 man-hrs (0.10 cum/man-hr)", "1 cum"),
            ("Coolie", "Hoisting bricks by pulley/head-loads to upper floor stages up to floor V level", "Hoisting crew; 0.75 day = 6.00 man-hrs per cum", "Labour", "6.00 man-hrs (0.17 cum/man-hr)", "1 cum"),
            ("Bhisti", "Tank soaking and wall curing by wetting burlap/spray curing for 14 days", "Watering crew; 0.50 day = 4.00 man-hrs per cum", "Labour", "4.00 man-hrs (0.25 cum/man-hr)", "1 cum"),
            ("Sundries", "Hoisting ropes, pulleys, buckets, scaffolding ties, and plumb bobs", "Sundries allowance (L.S. 8.19 x cost index)", "Equipment", "L.S. allowance", "1 cum"),
            ("REF#3.9 (Cement mortar 1:4)", "Referenced intermediate mortar: 0.25 cum cement mortar 1:4 per 1 cum superstructure brickwork", "Mortar scope executed per CPWD REF#3.9 specification", "Reference", "0.25 cum (0.250 cum/cum)", "1 cum"),
        ]
    },
    {
        "id": "6.4.2",
        "parent_title": "6.4 BRICK WORK IN SUPERSTRUCTURE UP TO FLOOR V LEVEL (F.P.S. BRICKS)",
        "title": "6.4.2 Brick work in superstructure above plinth up to floor V level with F.P.S. bricks in cement mortar 1:6",
        "unit": "cum", "base_qty": 1.0,
        "rows": [
            ("Common burnt clay F.P.S. bricks class 7.5", "Standard bricks for superstructure walls (494 net + wastage)", "Clean selected bricks conforming to IS:1077", "Material", "494.0 nos (494 nos/cum)", "1 cum"),
            ("Carriage of Bricks", "Haulage and stacking of bricks at site", "Carriage from kiln to site", "Material", "494.0 nos (494 nos/cum)", "1 cum"),
            ("Portland Cement", "Hydraulic binder decomposed from 0.25 cum of cement mortar 1:6 (REF#3.11) (0.063 tonne = 1.25 bags)", "OPC-43 Grade cement; 0.25 cum mortar x 0.25 t/cum = 0.063 t", "Material", "0.063 tonne (0.063 t/cum)", "1 cum"),
            ("Coarse sand", "Zone III coarse sand decomposed from 0.25 cum of cement mortar 1:6 (REF#3.11) (0.268 cum)", "Clean natural river sand Zone III", "Material", "0.268 cum (0.268 cum/cum)", "1 cum"),
            ("Carriage of Cement & Sand", "Haulage of cement and sand to mortar mixing platform", "Carriage allowance", "Material", "0.33 cum/t equiv", "1 cum"),
            ("Mechanical Mortar Mixer", "Machine batching and mixing mortar paste", "Mixer operational hire; output = 26.3 cum/hr", "Machine", "0.038 machine-hrs (26.3 cum/hr)", "1 cum"),
            ("Double steel / timber scaffolding", "Independent double scaffolding with steel tubular pipes, clamps and wooden planks up to floor V", "Scaffolding hire, erection, safety guardrails, and dismantling", "Equipment", "Scaffolding allowance", "1 cum"),
            ("Mason (brick layer) 1st & 2nd class", "Laying bricks in superstructure walls, plumb and line, raking joints 10mm for plaster key", "Skilled brick layer; 0.95 day = 7.60 man-hrs per cum", "Labour", "7.60 man-hrs (0.13 cum/man-hr)", "1 cum"),
            ("Beldar", "Soaking bricks, hoisting mortar to staging, serving masons, and cleaning scaffold", "Labour crew; 1.20 day = 9.60 man-hrs per cum", "Labour", "9.60 man-hrs (0.10 cum/man-hr)", "1 cum"),
            ("Coolie", "Hoisting bricks by pulley/head-loads to upper floor stages up to floor V level", "Hoisting crew; 0.75 day = 6.00 man-hrs per cum", "Labour", "6.00 man-hrs (0.17 cum/man-hr)", "1 cum"),
            ("Bhisti", "Tank soaking and wall curing by wetting burlap/spray curing for 14 days", "Watering crew; 0.50 day = 4.00 man-hrs per cum", "Labour", "4.00 man-hrs (0.25 cum/man-hr)", "1 cum"),
            ("Sundries", "Hoisting ropes, pulleys, buckets, scaffolding ties, and plumb bobs", "Sundries allowance", "Equipment", "L.S. allowance", "1 cum"),
            ("REF#3.11 (Cement mortar 1:6)", "Referenced intermediate mortar: 0.25 cum cement mortar 1:6 per 1 cum superstructure brickwork", "Mortar scope executed per CPWD REF#3.11 specification", "Reference", "0.25 cum (0.250 cum/cum)", "1 cum"),
        ]
    },

    # ─── 6.6 CAVITY WALLS ───────────────────────────────────────────────────────
    {
        "id": "6.6",
        "parent_title": "6.6 CAVITY WALL CONSTRUCTION",
        "title": "6.6 Extra for forming cavity 5 cm to 11.5 cm wide in cavity walls with galvanized M.S. ties (10 sqm)",
        "unit": "sqm", "base_qty": 10.0,
        "rows": [
            ("Galvanized MS wall ties / butterfly ties", "Galvanized mild steel butterfly/twist wall ties (200-250 mm long) connecting inner and outer leaf (45 nos.)", "Conforming to IS:1361; spaced at 450 mm horizontal & 450 mm vertical (4.5 ties/sqm)", "Material", "45.0 nos (4.50 nos/sqm)", "10 sqm"),
            ("Timber cavity cleaning battens & ropes", "Removable timber cavity battens suspended by cords to catch mortar droppings", "Batten hire and cavity cleaning allowance", "Equipment", "Batten allowance", "10 sqm"),
            ("Mason (brick layer)", "Placing wall ties with drip pointing downward, keeping cavity clear of mortar bridges, and checking width", "Skilled mason; 0.25 day = 2.00 man-hrs per 10 sqm", "Labour", "2.00 man-hrs (5.00 sqm/man-hr)", "10 sqm"),
            ("Beldar", "Lifting cavity battens periodically to remove mortar droppings before setting, keeping weep holes clear", "Helper crew; 0.25 day = 2.00 man-hrs per 10 sqm", "Labour", "2.00 man-hrs (5.00 sqm/man-hr)", "10 sqm"),
            ("Sundries", "Plastic cavity spacer discs, weep hole pipes, and cleaning rods", "Sundries allowance", "Equipment", "L.S. allowance", "10 sqm"),
        ]
    },

    # ─── 6.12 & 6.13 HALF BRICK MASONRY (10 SQM BASIS) ─────────────────────────
    {
        "id": "6.12.1",
        "parent_title": "6.12 HALF BRICK MASONRY IN SUPERSTRUCTURE UP TO FLOOR V LEVEL",
        "title": "6.12.1 Half brick masonry in superstructure up to floor V level with F.P.S. bricks in cement mortar 1:3 (10 sqm)",
        "unit": "sqm", "base_qty": 10.0,
        "rows": [
            ("Common burnt clay F.P.S. bricks class 7.5", "Standard burnt clay bricks laid on bed as stretchers (565 nos. per 10 sqm half-brick wall)", "Conforming to IS:1077 (550 net + wastage = 565 nos.)", "Material", "565.0 nos (56.5 nos/sqm)", "10 sqm"),
            ("Carriage of Bricks", "Mechanical haulage and stacking of bricks at site", "Carriage allowance", "Material", "565.0 nos (56.5 nos/sqm)", "10 sqm"),
            ("Portland Cement", "Hydraulic binder decomposed from 0.28 cum of cement mortar 1:3 (REF#3.8) (0.143 tonne = 2.86 bags)", "OPC-43 Grade cement; 0.28 cum mortar x 0.51 t/cum = 0.143 t", "Material", "0.143 tonne (0.0143 t/sqm)", "10 sqm"),
            ("Coarse sand", "Zone III coarse sand decomposed from 0.28 cum of cement mortar 1:3 (REF#3.8) (0.300 cum)", "Clean natural river sand Zone III", "Material", "0.300 cum (0.0300 cum/sqm)", "10 sqm"),
            ("Carriage of Cement & Sand", "Haulage of cement and sand to mixing platform", "Carriage allowance", "Material", "0.45 cum/t equiv", "10 sqm"),
            ("Mechanical Mortar Mixer", "Machine batching and mixing 1:3 mortar paste", "Mixer operational hire; 0.28 cum mortar x 0.15 hr/cum = 0.042 hr", "Machine", "0.042 machine-hrs", "10 sqm"),
            ("Single / double scaffolding", "Scaffolding staging for partition wall construction up to floor V", "Scaffolding hire allowance", "Equipment", "Scaffolding allowance", "10 sqm"),
            ("Mason (brick layer) 1st & 2nd class", "Laying half-brick partition true to plumb, maintaining straight bond, raking joints both faces", "Skilled brick layer; 1.05 day = 8.40 man-hrs per 10 sqm", "Labour", "8.40 man-hrs (1.19 sqm/man-hr)", "10 sqm"),
            ("Beldar", "Soaking bricks, hoisting mortar to staging, serving masons", "Labour crew; 1.15 day = 9.20 man-hrs per 10 sqm", "Labour", "9.20 man-hrs (1.09 sqm/man-hr)", "10 sqm"),
            ("Coolie", "Carrying bricks and mortar pans to partition location up to floor V", "Handling crew; 0.70 day = 5.60 man-hrs per 10 sqm", "Labour", "5.60 man-hrs (1.79 sqm/man-hr)", "10 sqm"),
            ("Bhisti", "Tank soaking and double-sided wall spray curing for 14 days", "Watering crew; 0.40 day = 3.20 man-hrs per 10 sqm", "Labour", "3.20 man-hrs (3.12 sqm/man-hr)", "10 sqm"),
            ("Sundries", "Plumb bobs, straight edges, mason lines, and scaffolding clamps", "Sundries allowance (L.S. 13.52 x cost index)", "Equipment", "L.S. allowance", "10 sqm"),
            ("REF#3.8 (Cement mortar 1:3)", "Referenced intermediate mortar: 0.28 cum cement mortar 1:3 per 10 sqm half-brick wall", "Mortar scope executed per CPWD REF#3.8 specification", "Reference", "0.28 cum (0.028 cum/sqm)", "10 sqm"),
        ]
    },
    {
        "id": "6.12.2",
        "parent_title": "6.12 HALF BRICK MASONRY IN SUPERSTRUCTURE UP TO FLOOR V LEVEL",
        "title": "6.12.2 Half brick masonry in superstructure up to floor V level with F.P.S. bricks in cement mortar 1:4 (10 sqm)",
        "unit": "sqm", "base_qty": 10.0,
        "rows": [
            ("Common burnt clay F.P.S. bricks class 7.5", "Standard bricks laid as stretchers (565 nos. per 10 sqm half-brick wall)", "Conforming to IS:1077 (565 nos.)", "Material", "565.0 nos (56.5 nos/sqm)", "10 sqm"),
            ("Carriage of Bricks", "Mechanical carriage of bricks to site", "Carriage allowance", "Material", "565.0 nos (56.5 nos/sqm)", "10 sqm"),
            ("Portland Cement", "Hydraulic binder decomposed from 0.28 cum of cement mortar 1:4 (REF#3.9) (0.106 tonne = 2.12 bags)", "OPC-43 Grade cement; 0.28 cum mortar x 0.38 t/cum = 0.106 t", "Material", "0.106 tonne (0.0106 t/sqm)", "10 sqm"),
            ("Coarse sand", "Zone III coarse sand decomposed from 0.28 cum of cement mortar 1:4 (REF#3.9) (0.300 cum)", "Clean natural river sand Zone III", "Material", "0.300 cum (0.0300 cum/sqm)", "10 sqm"),
            ("Carriage of Cement & Sand", "Haulage of cement and sand to mixing platform", "Carriage allowance", "Material", "0.41 cum/t equiv", "10 sqm"),
            ("Mechanical Mortar Mixer", "Machine batching and mixing 1:4 mortar paste", "Mixer operational hire; 0.28 cum mortar x 0.15 hr/cum = 0.042 hr", "Machine", "0.042 machine-hrs", "10 sqm"),
            ("Single / double scaffolding", "Scaffolding staging for partition wall construction up to floor V", "Scaffolding hire allowance", "Equipment", "Scaffolding allowance", "10 sqm"),
            ("Mason (brick layer) 1st & 2nd class", "Laying half-brick partition true to plumb, maintaining straight bond, raking joints both faces", "Skilled brick layer; 1.05 day = 8.40 man-hrs per 10 sqm", "Labour", "8.40 man-hrs (1.19 sqm/man-hr)", "10 sqm"),
            ("Beldar", "Soaking bricks, hoisting mortar to staging, serving masons", "Labour crew; 1.15 day = 9.20 man-hrs per 10 sqm", "Labour", "9.20 man-hrs (1.09 sqm/man-hr)", "10 sqm"),
            ("Coolie", "Carrying bricks and mortar pans to partition location up to floor V", "Handling crew; 0.70 day = 5.60 man-hrs per 10 sqm", "Labour", "5.60 man-hrs (1.79 sqm/man-hr)", "10 sqm"),
            ("Bhisti", "Tank soaking and double-sided wall spray curing for 14 days", "Watering crew; 0.40 day = 3.20 man-hrs per 10 sqm", "Labour", "3.20 man-hrs (3.12 sqm/man-hr)", "10 sqm"),
            ("Sundries", "Plumb bobs, straight edges, mason lines, and scaffolding clamps", "Sundries allowance", "Equipment", "L.S. allowance", "10 sqm"),
            ("REF#3.9 (Cement mortar 1:4)", "Referenced intermediate mortar: 0.28 cum cement mortar 1:4 per 10 sqm half-brick wall", "Mortar scope executed per CPWD REF#3.9 specification", "Reference", "0.28 cum (0.028 cum/sqm)", "10 sqm"),
        ]
    },
    {
        "id": "6.13.1",
        "parent_title": "6.13 HALF BRICK MASONRY IN FOUNDATION & PLINTH",
        "title": "6.13.1 Half brick masonry in foundation and plinth with F.P.S. bricks in cement mortar 1:3 (10 sqm)",
        "unit": "sqm", "base_qty": 10.0,
        "rows": [
            ("Common burnt clay F.P.S. bricks class 7.5", "Standard bricks laid as stretchers in plinth/honeycomb steps (565 nos.)", "Conforming to IS:1077 (565 nos.)", "Material", "565.0 nos (56.5 nos/sqm)", "10 sqm"),
            ("Carriage of Bricks", "Mechanical carriage of bricks to site", "Carriage allowance", "Material", "565.0 nos (56.5 nos/sqm)", "10 sqm"),
            ("Portland Cement", "Hydraulic binder decomposed from 0.28 cum of cement mortar 1:3 (REF#3.8) (0.143 tonne = 2.86 bags)", "OPC-43 Grade cement; 0.28 cum mortar x 0.51 t/cum = 0.143 t", "Material", "0.143 tonne (0.0143 t/sqm)", "10 sqm"),
            ("Coarse sand", "Zone III coarse sand decomposed from 0.28 cum of cement mortar 1:3 (REF#3.8) (0.300 cum)", "Clean natural river sand Zone III", "Material", "0.300 cum (0.0300 cum/sqm)", "10 sqm"),
            ("Carriage of Cement & Sand", "Haulage of cement and sand to mixing platform", "Carriage allowance", "Material", "0.45 cum/t equiv", "10 sqm"),
            ("Mechanical Mortar Mixer", "Machine batching and mixing 1:3 mortar paste", "Mixer operational hire; output = 23.8 cum/hr", "Machine", "0.042 machine-hrs", "10 sqm"),
            ("Mason (brick layer) 1st & 2nd class", "Laying half-brick wall in plinth/foundation trenches, checking plumb and alignment", "Skilled brick layer; 0.95 day = 7.60 man-hrs per 10 sqm", "Labour", "7.60 man-hrs (1.32 sqm/man-hr)", "10 sqm"),
            ("Beldar", "Soaking bricks, wheeling mortar, serving bricklayers", "Labour crew; 1.05 day = 8.40 man-hrs per 10 sqm", "Labour", "8.40 man-hrs (1.19 sqm/man-hr)", "10 sqm"),
            ("Coolie", "Hauling bricks and mortar along trenches", "Handling crew; 0.50 day = 4.00 man-hrs per 10 sqm", "Labour", "4.00 man-hrs (2.50 sqm/man-hr)", "10 sqm"),
            ("Bhisti", "Tank soaking and curing masonry for 14 days", "Watering crew; 0.35 day = 2.80 man-hrs per 10 sqm", "Labour", "2.80 man-hrs (3.57 sqm/man-hr)", "10 sqm"),
            ("Sundries", "Plumb bobs, straight edges, and mason lines", "Sundries allowance", "Equipment", "L.S. allowance", "10 sqm"),
            ("REF#3.8 (Cement mortar 1:3)", "Referenced intermediate mortar: 0.28 cum cement mortar 1:3 per 10 sqm half-brick wall", "Mortar scope executed per CPWD REF#3.8 specification", "Reference", "0.28 cum (0.028 cum/sqm)", "10 sqm"),
        ]
    },
    {
        "id": "6.13.2",
        "parent_title": "6.13 HALF BRICK MASONRY IN FOUNDATION & PLINTH",
        "title": "6.13.2 Half brick masonry in foundation and plinth with F.P.S. bricks in cement mortar 1:4 (10 sqm)",
        "unit": "sqm", "base_qty": 10.0,
        "rows": [
            ("Common burnt clay F.P.S. bricks class 7.5", "Standard bricks laid as stretchers in plinth/foundation (565 nos.)", "Conforming to IS:1077 (565 nos.)", "Material", "565.0 nos (56.5 nos/sqm)", "10 sqm"),
            ("Carriage of Bricks", "Mechanical carriage of bricks to site", "Carriage allowance", "Material", "565.0 nos (56.5 nos/sqm)", "10 sqm"),
            ("Portland Cement", "Hydraulic binder decomposed from 0.28 cum of cement mortar 1:4 (REF#3.9) (0.106 tonne = 2.12 bags)", "OPC-43 Grade cement; 0.28 cum mortar x 0.38 t/cum = 0.106 t", "Material", "0.106 tonne (0.0106 t/sqm)", "10 sqm"),
            ("Coarse sand", "Zone III coarse sand decomposed from 0.28 cum of cement mortar 1:4 (REF#3.9) (0.300 cum)", "Clean natural river sand Zone III", "Material", "0.300 cum (0.0300 cum/sqm)", "10 sqm"),
            ("Carriage of Cement & Sand", "Haulage of cement and sand to mixing platform", "Carriage allowance", "Material", "0.41 cum/t equiv", "10 sqm"),
            ("Mechanical Mortar Mixer", "Machine batching and mixing 1:4 mortar paste", "Mixer operational hire; output = 23.8 cum/hr", "Machine", "0.042 machine-hrs", "10 sqm"),
            ("Mason (brick layer) 1st & 2nd class", "Laying half-brick wall in plinth/foundation trenches", "Skilled brick layer; 0.95 day = 7.60 man-hrs per 10 sqm", "Labour", "7.60 man-hrs (1.32 sqm/man-hr)", "10 sqm"),
            ("Beldar", "Soaking bricks, wheeling mortar, serving bricklayers", "Labour crew; 1.05 day = 8.40 man-hrs per 10 sqm", "Labour", "8.40 man-hrs (1.19 sqm/man-hr)", "10 sqm"),
            ("Coolie", "Hauling bricks and mortar along trenches", "Handling crew; 0.50 day = 4.00 man-hrs per 10 sqm", "Labour", "4.00 man-hrs (2.50 sqm/man-hr)", "10 sqm"),
            ("Bhisti", "Tank soaking and curing masonry for 14 days", "Watering crew; 0.35 day = 2.80 man-hrs per 10 sqm", "Labour", "2.80 man-hrs (3.57 sqm/man-hr)", "10 sqm"),
            ("Sundries", "Plumb bobs, straight edges, and mason lines", "Sundries allowance", "Equipment", "L.S. allowance", "10 sqm"),
            ("REF#3.9 (Cement mortar 1:4)", "Referenced intermediate mortar: 0.28 cum cement mortar 1:4 per 10 sqm half-brick wall", "Mortar scope executed per CPWD REF#3.9 specification", "Reference", "0.28 cum (0.028 cum/sqm)", "10 sqm"),
        ]
    },

    # ─── 6.15 REINFORCEMENT IN HALF BRICK MASONRY ───────────────────────────────
    {
        "id": "6.15",
        "parent_title": "6.15 REINFORCEMENT IN HALF BRICK MASONRY",
        "title": "6.15 Extra for providing and placing in position 2 Nos 6mm dia M.S. bars at every third course (100 m run)",
        "unit": "metre", "base_qty": 100.0,
        "rows": [
            ("M.S. round bars 6 mm dia (IS:432)", "Mild steel reinforcement bars embedded in bed joints (0.044 quintal = 44.4 kg per 100 m pair)", "Conforming to IS:432 Grade I; nominal weight 0.222 kg/m x 2 bars = 0.444 kg/m", "Material", "44.40 kg (0.444 kg/m)", "100 m"),
            ("Carriage of Steel", "Haulage and handling of steel bars from store to work floor", "Carriage allowance", "Material", "0.044 quintal", "100 m"),
            ("Blacksmith (2nd class) / Fitter", "Straightening bars, cutting to lengths, hooking ends at junctions, and placing true in mortar bed", "Skilled steel fixer; 0.12 day = 0.96 man-hrs per 100 m", "Labour", "0.96 man-hrs (104.2 m/man-hr)", "100 m"),
            ("Beldar", "Holding bars during embedding and fully encasing in mortar without contact with bricks", "Helper crew; 0.12 day = 0.96 man-hrs per 100 m", "Labour", "0.96 man-hrs (104.2 m/man-hr)", "100 m"),
            ("Sundries", "Bar benders, hacksaw blades, wire brushes, and rust cleaning", "Sundries allowance (L.S. 1.82 x cost index)", "Equipment", "L.S. allowance", "100 m"),
        ]
    },

    # ─── 6.23 HONEYCOMB BRICK WORK ──────────────────────────────────────────────
    {
        "id": "6.23",
        "parent_title": "6.23 HONEY-COMB BRICK WORK",
        "title": "6.23 Honey-comb brick work 10 / 11.4 cm thick with F.P.S. bricks in cement mortar 1:4 (10 sqm)",
        "unit": "sqm", "base_qty": 10.0,
        "rows": [
            ("Common burnt clay F.P.S. bricks class 7.5", "Standard bricks laid with open perforations/apertures for architectural screening and ventilation (350 nos.)", "Bricks laid on edge/bed with regular open spaces conforming to architectural drawing", "Material", "350.0 nos (35.0 nos/sqm)", "10 sqm"),
            ("Carriage of Bricks", "Mechanical haulage and stacking of bricks", "Carriage allowance", "Material", "350.0 nos (35.0 nos/sqm)", "10 sqm"),
            ("Portland Cement", "Hydraulic binder decomposed from 0.12 cum of cement mortar 1:4 (REF#3.9) (0.046 tonne = 0.92 bag)", "OPC-43 Grade cement; 0.12 cum mortar x 0.38 t/cum = 0.046 t", "Material", "0.046 tonne (0.0046 t/sqm)", "10 sqm"),
            ("Coarse sand", "Zone III coarse sand decomposed from 0.12 cum of cement mortar 1:4 (REF#3.9) (0.128 cum)", "Clean natural river sand Zone III", "Material", "0.128 cum (0.0128 cum/sqm)", "10 sqm"),
            ("Carriage of Cement & Sand", "Haulage of materials to mixing platform", "Carriage allowance", "Material", "0.18 cum/t equiv", "10 sqm"),
            ("Mechanical Mortar Mixer", "Machine mixing mortar paste", "Mixer operational hire; output = 55.5 cum/hr", "Machine", "0.018 machine-hrs", "10 sqm"),
            ("Scaffolding staging", "Scaffolding for honeycomb screen construction", "Scaffolding allowance", "Equipment", "Scaffolding allowance", "10 sqm"),
            ("Mason (brick layer) 1st & 2nd class", "Setting bricks with uniform open voids, maintaining exact spacing, cleaning mortar droppings from holes", "Skilled mason; 1.25 day = 10.00 man-hrs per 10 sqm", "Labour", "10.00 man-hrs (1.00 sqm/man-hr)", "10 sqm"),
            ("Beldar", "Soaking bricks, hoisting mortar, cleaning void apertures", "Labour crew; 1.10 day = 8.80 man-hrs per 10 sqm", "Labour", "8.80 man-hrs (1.14 sqm/man-hr)", "10 sqm"),
            ("Coolie", "Carrying bricks and mortar pans to staging", "Handling crew; 0.60 day = 4.80 man-hrs per 10 sqm", "Labour", "4.80 man-hrs (2.08 sqm/man-hr)", "10 sqm"),
            ("Bhisti", "Tank soaking and spray curing honeycomb screen on both faces", "Watering crew; 0.35 day = 2.80 man-hrs per 10 sqm", "Labour", "2.80 man-hrs (3.57 sqm/man-hr)", "10 sqm"),
            ("Sundries", "Wooden spacing gauges, plumb bobs, and clean pointing tools", "Sundries allowance", "Equipment", "L.S. allowance", "10 sqm"),
            ("REF#3.9 (Cement mortar 1:4)", "Referenced intermediate mortar: 0.12 cum cement mortar 1:4 per 10 sqm honeycomb screen", "Mortar scope executed per CPWD REF#3.9 specification", "Reference", "0.12 cum (0.012 cum/sqm)", "10 sqm"),
        ]
    },

    # ─── 6.38 & 6.47 AUTOCLAVED AERATED CONCRETE (AAC) BLOCKS ───────────────────
    {
        "id": "6.38",
        "parent_title": "6.38 AUTOCLAVED AERATED CONCRETE (AAC) BLOCKS MASONRY",
        "title": "6.38 Autoclaved aerated concrete (AAC) blocks masonry with cement mortar 1:4",
        "unit": "cum", "base_qty": 1.0,
        "rows": [
            ("Autoclaved aerated concrete (AAC) blocks", "High-precision lightweight steam-cured cellular concrete blocks conforming to IS:2185 Part 3 (1.00 cum net)", "Grade 1 AAC blocks (density 551-650 kg/cum); thermal insulating, lightweight block masonry", "Material", "1.00 cum (1.000 cum/cum)", "1 cum"),
            ("Carriage of AAC blocks", "Careful mechanical transport, palletised unloading, and protected stacking to avoid edge damage", "Carriage allowance", "Material", "1.00 cum (1.000 cum/cum)", "1 cum"),
            ("Portland Cement", "Hydraulic binder decomposed from 0.15 cum of cement mortar 1:4 (REF#3.9) (0.057 tonne = 1.14 bags)", "OPC-43 Grade cement; 0.15 cum mortar x 0.38 t/cum = 0.057 t", "Material", "0.057 tonne (0.057 t/cum)", "1 cum"),
            ("Coarse sand", "Zone III coarse sand decomposed from 0.15 cum of cement mortar 1:4 (REF#3.9) (0.161 cum)", "Clean natural river sand Zone III", "Material", "0.161 cum (0.161 cum/cum)", "1 cum"),
            ("Carriage of Cement & Sand", "Haulage of materials to mixing platform", "Carriage allowance", "Material", "0.22 cum/t equiv", "1 cum"),
            ("Mechanical Mortar Mixer", "Machine mixing thin mortar paste", "Mixer operational hire; output = 43.5 cum/hr", "Machine", "0.023 machine-hrs", "1 cum"),
            ("Scaffolding staging", "Staging and working platforms up to floor V", "Scaffolding allowance", "Equipment", "Scaffolding allowance", "1 cum"),
            ("Mason (brick layer) 1st & 2nd class", "Laying large-format AAC blocks, buttering thin bed joints, checking alignment with spirit level and plumb", "Skilled mason; 0.65 day = 5.20 man-hrs per cum", "Labour", "5.20 man-hrs (0.19 cum/man-hr)", "1 cum"),
            ("Beldar", "Moistening block surfaces with brush, handling blocks, mixing mortar, and serving masons", "Labour crew; 0.70 day = 5.60 man-hrs per cum", "Labour", "5.60 man-hrs (0.18 cum/man-hr)", "1 cum"),
            ("Coolie", "Careful handling of lightweight blocks to work face without breaking corners", "Handling crew; 0.40 day = 3.20 man-hrs per cum", "Labour", "3.20 man-hrs (0.31 cum/man-hr)", "1 cum"),
            ("Bhisti", "Moist curing mortar joints with fine spray for 10 days", "Watering crew; 0.30 day = 2.40 man-hrs per cum", "Labour", "2.40 man-hrs (0.42 cum/man-hr)", "1 cum"),
            ("Sundries", "Hand saw for cutting blocks, notched trowels, rubber mallets, and spirit levels", "Sundries allowance (L.S. 2.73 x cost index)", "Equipment", "L.S. allowance", "1 cum"),
            ("REF#3.9 (Cement mortar 1:4)", "Referenced intermediate mortar: 0.15 cum cement mortar 1:4 per 1 cum AAC blockwork", "Mortar scope executed per CPWD REF#3.9 specification", "Reference", "0.15 cum (0.150 cum/cum)", "1 cum"),
        ]
    },
    {
        "id": "6.47",
        "parent_title": "6.47 AUTOCLAVED AERATED CONCRETE (AAC) BLOCKS WITH THIN-BED ADHESIVE",
        "title": "6.47 Autoclaved aerated concrete (AAC) blocks masonry with thin-bed polymer block adhesive",
        "unit": "cum", "base_qty": 1.0,
        "rows": [
            ("Autoclaved aerated concrete (AAC) blocks", "Precision factory-cut AAC blocks conforming to IS:2185 Part 3 (1.00 cum net)", "High-dimensional-accuracy AAC blocks for thin-joint application (joint thickness 2-3 mm)", "Material", "1.00 cum (1.000 cum/cum)", "1 cum"),
            ("Carriage of AAC blocks", "Palletised carriage and careful manual handling of AAC blocks", "Carriage allowance", "Material", "1.00 cum (1.000 cum/cum)", "1 cum"),
            ("Polymer modified thin-bed block adhesive", "Factory-premixed dry mortar adhesive conforming to IS:15477 / ASTM C1660 (30 kg per 1 cum masonry)", "Ready-to-use adhesive mixed with water; high tensile adhesion and thermal bridge elimination", "Material", "30.00 kg (30.0 kg/cum)", "1 cum"),
            ("Carriage of Block adhesive", "Handling and transport of 30 kg adhesive bags to work floor", "Carriage allowance", "Material", "0.03 tonne", "1 cum"),
            ("Hand paddle mixer (electric)", "Electric hand drill mixer with paddle attachment for high-shear adhesive slurry mixing", "Paddle mixer hire allowance", "Equipment", "Mixer allowance", "1 cum"),
            ("Scaffolding staging", "Staging platforms for partition wall construction", "Scaffolding allowance", "Equipment", "Scaffolding allowance", "1 cum"),
            ("Mason (brick layer) 1st & 2nd class", "Applying thin adhesive with notched trowel (3mm comb), setting blocks, tapping with rubber mallet, plumbing", "Skilled mason; 0.36 day = 2.88 man-hrs per cum", "Labour", "2.88 man-hrs (0.35 cum/man-hr)", "1 cum"),
            ("Beldar", "Electric paddle mixing of adhesive in buckets, handling blocks, and keeping blocks clean of dust", "Labour crew; 0.45 day = 3.60 man-hrs per cum", "Labour", "3.60 man-hrs (0.28 cum/man-hr)", "1 cum"),
            ("Coolie", "Carrying blocks and adhesive buckets to work face", "Handling crew; 0.25 day = 2.00 man-hrs per cum", "Labour", "2.00 man-hrs (0.50 cum/man-hr)", "1 cum"),
            ("Sundries", "Notched trowels (3mm), rubber mallets, block rasp, spirit levels, and adhesive buckets (no curing needed)", "Sundries allowance (L.S. 2.73 x cost index)", "Equipment", "L.S. allowance", "1 cum"),
        ]
    },

    # ─── 6.44 BRICK EDGING TO PLINTH PROTECTION ─────────────────────────────────
    {
        "id": "6.44",
        "parent_title": "6.44 BRICK EDGING TO PLINTH PROTECTION",
        "title": "6.44 Brick edging 7cm wide 11.4 cm deep to plinth protection with F.P.S. bricks in cement mortar 1:4 (100 m run)",
        "unit": "metre", "base_qty": 100.0,
        "rows": [
            ("Common burnt clay F.P.S. bricks class 7.5", "Standard bricks laid on edge as protective retaining toe kerb along apron (46 nos. per 10 m = 460 nos./100m)", "Class 7.5 bricks embedded 11.4 cm deep in ground to retain plinth apron ballast", "Material", "460.0 nos (4.60 nos/m)", "100 m"),
            ("Carriage of Bricks", "Mechanical haulage and unloading along perimeter trenches", "Carriage allowance", "Material", "460.0 nos (4.60 nos/m)", "100 m"),
            ("Portland Cement", "Hydraulic binder decomposed from 0.036 cum of cement mortar 1:4 (REF#3.9) (0.014 tonne = 0.28 bag)", "OPC-43 Grade cement; 0.036 cum mortar x 0.38 t/cum = 0.014 t", "Material", "0.014 tonne (0.00014 t/m)", "100 m"),
            ("Coarse sand", "Zone III coarse sand decomposed from 0.036 cum of cement mortar 1:4 (REF#3.9) (0.039 cum)", "Clean natural river sand Zone III", "Material", "0.039 cum (0.00039 cum/m)", "100 m"),
            ("Carriage of Cement & Sand", "Haulage of cement and sand to perimeter trenches", "Carriage allowance", "Material", "0.05 cum/t equiv", "100 m"),
            ("Mason (brick layer)", "Excavating 7x11.4cm shallow trench, bedding bricks on edge, aligning to string, and pointing joints", "Skilled mason; 0.50 day = 4.00 man-hrs per 100 m", "Labour", "4.00 man-hrs (25.0 m/man-hr)", "100 m"),
            ("Beldar / Coolie", "Soaking bricks, wheeling mortar along trench, backfilling earth against outer brick edge, and ramming", "Labour crew; 0.80 day = 6.40 man-hrs per 100 m", "Labour", "6.40 man-hrs (15.6 m/man-hr)", "100 m"),
            ("Bhisti", "Watering brick edging joints and surrounding soil", "Watering crew; 0.25 day = 2.00 man-hrs per 100 m", "Labour", "2.00 man-hrs (50.0 m/man-hr)", "100 m"),
            ("Sundries", "Rammers, strings, pointing trowels, and stakes", "Sundries allowance", "Equipment", "L.S. allowance", "100 m"),
            ("REF#3.9 (Cement mortar 1:4)", "Referenced intermediate mortar: 0.036 cum cement mortar 1:4 per 100 m brick edging", "Mortar scope executed per CPWD REF#3.9 specification", "Reference", "0.036 cum (0.00036 cum/m)", "100 m"),
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

def build_masonry_sheet(ws):
    if ws.views.sheetView:
        ws.views.sheetView[0].showGridLines = True

    # Sheet title spanning A1:G1
    ws.merge_cells("A1:G1")
    t1 = ws.cell(row=1, column=1, value="Sub-Head 6.0 — MASONRY WORK  |  First-Principles Resource, Work & Gang Analysis")
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
    for item in MASONRY_ITEMS:
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

def rebuild_masonry_in_workbook(file_path):
    print(f"Opening {file_path} ...")
    wb = openpyxl.load_workbook(file_path)
    if "06_Masonry_Work" not in wb.sheetnames:
        raise ValueError(f"'06_Masonry_Work' sheet not found in {file_path}")

    pos = wb.sheetnames.index("06_Masonry_Work")
    del wb["06_Masonry_Work"]
    ws = wb.create_sheet("06_Masonry_Work", pos)
    ws.sheet_properties.tabColor = "2E75B6"
    build_masonry_sheet(ws)

    temp_path = file_path.replace(".xlsx", "_TMP_MASONRY.xlsx")
    wb.save(temp_path)
    wb.close()
    os.replace(temp_path, file_path)
    print(f"Successfully updated {file_path} -> sheet '06_Masonry_Work' ({len(MASONRY_ITEMS)} items).")

def main():
    repo_root = Path(__file__).resolve().parents[1]
    main_wb = repo_root / "CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_1_Latest.xlsx"
    output_wb = repo_root / "outputs" / "earthwork-custom-rate-composer" / "CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_1_Latest.xlsx"

    if main_wb.exists():
        rebuild_masonry_in_workbook(str(main_wb))
    if output_wb.exists():
        rebuild_masonry_in_workbook(str(output_wb))

if __name__ == "__main__":
    main()
