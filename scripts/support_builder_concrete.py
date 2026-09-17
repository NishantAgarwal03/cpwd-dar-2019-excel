"""
support_builder_concrete.py
Rebuilds the existing '04_Concrete_Work' worksheet with First-Principles Resource,
Productivity, Gang, Machinery and Material Analysis per CPWD DAR 2019 Vol 1:
  Item Code | Labour / Machine / Material | Work done | Condition / When used | Category | Productivity | Quantity
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from pathlib import Path
import os, sys, shutil

# Colours
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

CONCRETE_ITEMS = [
    # ─── 4.1 PLAIN CEMENT CONCRETE UP TO PLINTH LEVEL ───────────────────────────
    {
        "id": "4.1.2",
        "parent_title": "4.1 CEMENT CONCRETE UP TO PLINTH LEVEL (EXCL. SHUTTERING)",
        "title": "4.1.2 Cement concrete 1:1.5:3 (1 cement : 1.5 coarse sand : 3 graded stone aggregate 20 mm)",
        "unit": "cum", "base_qty": 1.0,
        "rows": [
            ("Stone Aggregate 20 mm", "Coarse aggregate single size 20 mm providing structural skeleton (0.57 cum)", "Clean crushed granite/basalt aggregate conforming to IS:383", "Material", "0.57 cum (0.570 cum/cum)", "1 cum"),
            ("Stone Aggregate 10 mm", "Coarse aggregate single size 10 mm filling voids between 20 mm stones (0.28 cum)", "Graded 10 mm crushed stone aggregate for dense matrix packing", "Material", "0.28 cum (0.280 cum/cum)", "1 cum"),
            ("Carriage of Stone aggregate below 40 mm", "Mechanical haulage, unloading and stacking of 20mm & 10mm aggregates at site (0.85 cum)", "Transport from crusher plant to site batching yard", "Material", "0.85 cum (0.850 cum/cum)", "1 cum"),
            ("Coarse sand", "Zone III coarse sand filling interstitial voids between aggregates (0.425 cum)", "Clean natural river coarse sand conforming to IS:383 Zone III", "Material", "0.425 cum (0.425 cum/cum)", "1 cum"),
            ("Carriage of Coarse sand", "Mechanical transport, unloading and stacking of coarse sand at site (0.425 cum)", "Carriage from sand pit to site mixing plant", "Material", "0.425 cum (0.425 cum/cum)", "1 cum"),
            ("Portland Cement", "Hydraulic binding matrix providing compressive strength and cohesion (0.40 tonne)", "Ordinary Portland Cement 43 Grade conforming to IS:269 / IS:8112", "Material", "0.40 tonne (0.400 t/cum)", "1 cum"),
            ("Carriage of Cement", "Handling, mechanical transport and unloading of cement bags from store to mixer (0.40 tonne)", "Standard site haulage allowance for cement bags", "Material", "0.40 tonne (0.400 t/cum)", "1 cum"),
            ("Concrete Mixer 0.28 to 0.40 cum", "Machine batching and mechanical wet mixing to achieve uniform plastic workability", "Mixer operational hire; cycle duration ~18 min/cum; output = 1.79 cum/hr", "Machine", "0.56 machine-hrs (1.79 cum/hr)", "1 cum"),
            ("Needle Vibrator", "Mechanical internal needle vibration to consolidate concrete and eliminate entrapped air voids", "Needle vibrator hire; output = 1.79 cum/hr", "Machine", "0.56 machine-hrs (1.79 cum/hr)", "1 cum"),
            ("Mason (average)", "Leveling concrete surface, establishing grade pegs, screeding, and float finishing", "Skilled mason; 0.10 day = 0.80 man-hrs per cum", "Labour", "0.80 man-hrs (1.25 cum/man-hr)", "1 cum"),
            ("Beldar", "Measuring aggregates, feeding mixer hopper, transporting, spreading, and consolidating in layers", "Operational labour crew; 1.63 day = 13.04 man-hrs per cum", "Labour", "13.04 man-hrs (0.08 cum/man-hr)", "1 cum"),
            ("Coolie", "Carrying concrete pans, shifting aggregates, and assisting batching crew", "Unskilled haulage labour; 0.70 day = 5.60 man-hrs per cum", "Labour", "5.60 man-hrs (0.18 cum/man-hr)", "1 cum"),
            ("Bhisti", "Adding measured water during mixing, cleaning mixer drum, and watering concrete for initial cure", "Water supply and curing crew; 0.70 day = 5.60 man-hrs per cum", "Labour", "5.60 man-hrs (0.18 cum/man-hr)", "1 cum"),
            ("Sundries", "Vibrator fuel, screed boards, floats, measuring farmas, and platform upkeep", "Sundries and minor tools allowance (L.S. 14.30 x cost index)", "Equipment", "L.S. allowance", "1 cum"),
        ]
    },
    {
        "id": "4.1.3",
        "parent_title": "4.1 CEMENT CONCRETE UP TO PLINTH LEVEL (EXCL. SHUTTERING)",
        "title": "4.1.3 Cement concrete 1:2:4 (1 cement : 2 coarse sand : 4 graded stone aggregate 20 mm)",
        "unit": "cum", "base_qty": 1.0,
        "rows": [
            ("Stone Aggregate 20 mm", "Coarse aggregate single size 20 mm providing structural bulk skeleton (0.57 cum)", "Clean crushed stone aggregate conforming to IS:383 nominal 20 mm", "Material", "0.57 cum (0.570 cum/cum)", "1 cum"),
            ("Stone Aggregate 10 mm", "Coarse aggregate single size 10 mm grading interlocking voids (0.28 cum)", "Clean crushed stone aggregate 10 mm for void filling", "Material", "0.28 cum (0.280 cum/cum)", "1 cum"),
            ("Carriage of Stone aggregate below 40 mm", "Mechanical carriage and unloading of 20mm & 10mm stone aggregates (0.85 cum)", "Carriage from crushing plant to batching platform", "Material", "0.85 cum (0.850 cum/cum)", "1 cum"),
            ("Coarse sand", "Zone III coarse sand matrix (0.425 cum)", "Clean natural river sand Zone III conforming to CPWD specification", "Material", "0.425 cum (0.425 cum/cum)", "1 cum"),
            ("Carriage of Coarse sand", "Mechanical haulage and unloading of coarse sand at site (0.425 cum)", "Carriage of coarse sand", "Material", "0.425 cum (0.425 cum/cum)", "1 cum"),
            ("Portland Cement", "Hydraulic binding paste matrix (0.32 tonne = 6.4 bags)", "Standard 1:2:4 mix cement allowance conforming to IS:269", "Material", "0.32 tonne (0.320 t/cum)", "1 cum"),
            ("Carriage of Cement", "Handling and transport of cement bags to site batching point (0.32 tonne)", "Carriage of bagged cement", "Material", "0.32 tonne (0.320 t/cum)", "1 cum"),
            ("Concrete Mixer 0.28 to 0.40 cum", "Machine batching and thorough wet mixing of 1:2:4 concrete", "Mixer operational hire; cycle ~18 min/cum; output = 1.79 cum/hr", "Machine", "0.56 machine-hrs (1.79 cum/hr)", "1 cum"),
            ("Needle Vibrator", "Compacting placed concrete to achieve dense honey-comb-free mass", "Needle vibrator operational hire; output = 1.79 cum/hr", "Machine", "0.56 machine-hrs (1.79 cum/hr)", "1 cum"),
            ("Mason (average)", "Leveling concrete surface, screeding, checking levels and float finishing", "Mason labour allowance; 0.10 day = 0.80 man-hrs per cum", "Labour", "0.80 man-hrs (1.25 cum/man-hr)", "1 cum"),
            ("Beldar", "Measuring ingredients, feeding mixer drum, transporting, and spreading in layers", "Labour crew; 1.63 day = 13.04 man-hrs per cum", "Labour", "13.04 man-hrs (0.08 cum/man-hr)", "1 cum"),
            ("Coolie", "Carrying concrete pans and moving materials around the work area", "Unskilled labour; 0.70 day = 5.60 man-hrs per cum", "Labour", "5.60 man-hrs (0.18 cum/man-hr)", "1 cum"),
            ("Bhisti", "Water batching, drum washing, and initial hydration watering", "Watering crew; 0.70 day = 5.60 man-hrs per cum", "Labour", "5.60 man-hrs (0.18 cum/man-hr)", "1 cum"),
            ("Sundries", "Farma boxes, wheelbarrows, float tools, and testing cube moulds", "Sundries allowance (L.S. 14.30 x cost index)", "Equipment", "L.S. allowance", "1 cum"),
        ]
    },
    {
        "id": "4.1.4",
        "parent_title": "4.1 CEMENT CONCRETE UP TO PLINTH LEVEL (EXCL. SHUTTERING)",
        "title": "4.1.4 Cement concrete 1:2:4 (1 cement : 2 coarse sand : 4 graded stone aggregate 40 mm)",
        "unit": "cum", "base_qty": 1.0,
        "rows": [
            ("Stone Aggregate 40 mm", "Coarse aggregate single size 40 mm for mass concrete foundations (0.57 cum)", "Clean graded crushed stone 40 mm nominal size", "Material", "0.57 cum (0.570 cum/cum)", "1 cum"),
            ("Stone Aggregate 20 mm", "Coarse aggregate single size 20 mm filling large aggregate voids (0.28 cum)", "Graded crushed stone 20 mm nominal size", "Material", "0.28 cum (0.280 cum/cum)", "1 cum"),
            ("Carriage of Stone aggregate 40 mm", "Haulage and unloading of 40 mm stone aggregate (0.57 cum)", "Carriage of 40 mm aggregate", "Material", "0.57 cum (0.570 cum/cum)", "1 cum"),
            ("Carriage of Stone aggregate below 40 mm", "Haulage and unloading of 20 mm stone aggregate (0.28 cum)", "Carriage of 20 mm aggregate", "Material", "0.28 cum (0.280 cum/cum)", "1 cum"),
            ("Coarse sand", "Zone III coarse sand matrix (0.425 cum)", "Clean natural river sand Zone III", "Material", "0.425 cum (0.425 cum/cum)", "1 cum"),
            ("Carriage of Coarse sand", "Carriage of coarse sand to site platform (0.425 cum)", "Sand haulage", "Material", "0.425 cum (0.425 cum/cum)", "1 cum"),
            ("Portland Cement", "Hydraulic binding cement matrix (0.32 tonne)", "OPC-43 Grade cement conforming to IS:269", "Material", "0.32 tonne (0.320 t/cum)", "1 cum"),
            ("Carriage of Cement", "Transport of cement bags to mixing yard (0.32 tonne)", "Carriage of cement", "Material", "0.32 tonne (0.320 t/cum)", "1 cum"),
            ("Concrete Mixer 0.28 to 0.40 cum", "Machine batching and mixing 40 mm aggregate concrete", "Mixer operational hire; output = 1.79 cum/hr", "Machine", "0.56 machine-hrs (1.79 cum/hr)", "1 cum"),
            ("Needle / Surface Vibrator", "Compacting large aggregate concrete in foundation footings", "Vibrator hire; output = 1.79 cum/hr", "Machine", "0.56 machine-hrs (1.79 cum/hr)", "1 cum"),
            ("Mason (average)", "Leveling and finishing surface to designated invert/plinth levels", "Mason labour; 0.10 day = 0.80 man-hrs per cum", "Labour", "0.80 man-hrs (1.25 cum/man-hr)", "1 cum"),
            ("Beldar", "Measuring ingredients, charging mixer, shoveling and laying concrete", "Labour crew; 1.63 day = 13.04 man-hrs per cum", "Labour", "13.04 man-hrs (0.08 cum/man-hr)", "1 cum"),
            ("Coolie", "Carrying concrete pans and moving coarse aggregate piles", "Unskilled labour; 0.70 day = 5.60 man-hrs per cum", "Labour", "5.60 man-hrs (0.18 cum/man-hr)", "1 cum"),
            ("Bhisti", "Water batching, drum cleaning, and foundation watering", "Watering crew; 0.70 day = 5.60 man-hrs per cum", "Labour", "5.60 man-hrs (0.18 cum/man-hr)", "1 cum"),
            ("Sundries", "Sundry tools, farma boxes, wheelbarrows, and testing apparatus", "Sundries allowance", "Equipment", "L.S. allowance", "1 cum"),
        ]
    },
    {
        "id": "4.1.5",
        "parent_title": "4.1 CEMENT CONCRETE UP TO PLINTH LEVEL (EXCL. SHUTTERING)",
        "title": "4.1.5 Cement concrete 1:3:6 (1 cement : 3 coarse sand : 6 graded stone aggregate 20 mm)",
        "unit": "cum", "base_qty": 1.0,
        "rows": [
            ("Stone Aggregate 20 mm", "Coarse aggregate single size 20 mm for lean foundation base (0.60 cum)", "Clean crushed stone aggregate 20 mm nominal size", "Material", "0.60 cum (0.600 cum/cum)", "1 cum"),
            ("Stone Aggregate 10 mm", "Coarse aggregate single size 10 mm for matrix interlock (0.30 cum)", "Clean crushed stone aggregate 10 mm nominal size", "Material", "0.30 cum (0.300 cum/cum)", "1 cum"),
            ("Carriage of Stone aggregate below 40 mm", "Haulage and unloading of 20mm & 10mm aggregates (0.90 cum)", "Aggregate carriage to site platform", "Material", "0.90 cum (0.900 cum/cum)", "1 cum"),
            ("Coarse sand", "Zone III coarse sand matrix (0.45 cum)", "Clean coarse sand conforming to IS:383 Zone III", "Material", "0.45 cum (0.450 cum/cum)", "1 cum"),
            ("Carriage of Coarse sand", "Mechanical haulage and unloading of coarse sand (0.45 cum)", "Carriage of coarse sand", "Material", "0.45 cum (0.450 cum/cum)", "1 cum"),
            ("Portland Cement", "Hydraulic binder for 1:3:6 lean structural sub-base (0.22 tonne = 4.4 bags)", "OPC-43 Grade cement conforming to IS:269", "Material", "0.22 tonne (0.220 t/cum)", "1 cum"),
            ("Carriage of Cement", "Transport of cement bags to mixing yard (0.22 tonne)", "Carriage of cement", "Material", "0.22 tonne (0.220 t/cum)", "1 cum"),
            ("Concrete Mixer 0.28 to 0.40 cum", "Machine batching and thorough wet mixing of 1:3:6 concrete", "Mixer operational hire; cycle ~18 min/cum; output = 1.79 cum/hr", "Machine", "0.56 machine-hrs (1.79 cum/hr)", "1 cum"),
            ("Needle Vibrator", "Compacting sub-base concrete bed to uniform density", "Vibrator hire; output = 1.79 cum/hr", "Machine", "0.56 machine-hrs (1.79 cum/hr)", "1 cum"),
            ("Mason (average)", "Leveling concrete bed, grading slopes, and screed finishing", "Mason labour; 0.10 day = 0.80 man-hrs per cum", "Labour", "0.80 man-hrs (1.25 cum/man-hr)", "1 cum"),
            ("Beldar", "Measuring ingredients, feeding mixer drum, and laying concrete in trenches", "Labour crew; 1.63 day = 13.04 man-hrs per cum", "Labour", "13.04 man-hrs (0.08 cum/man-hr)", "1 cum"),
            ("Coolie", "Hauling concrete pans and assisting batching operations", "Unskilled labour; 0.70 day = 5.60 man-hrs per cum", "Labour", "5.60 man-hrs (0.18 cum/man-hr)", "1 cum"),
            ("Bhisti", "Water batching, mixer drum cleanup, and curing bed wetting", "Watering crew; 0.70 day = 5.60 man-hrs per cum", "Labour", "5.60 man-hrs (0.18 cum/man-hr)", "1 cum"),
            ("Sundries", "Sundry tools, measuring boxes, pans, and platform upkeep", "Sundries allowance", "Equipment", "L.S. allowance", "1 cum"),
        ]
    },
    {
        "id": "4.1.6",
        "parent_title": "4.1 CEMENT CONCRETE UP TO PLINTH LEVEL (EXCL. SHUTTERING)",
        "title": "4.1.6 Cement concrete 1:3:6 (1 cement : 3 coarse sand : 6 graded stone aggregate 40 mm)",
        "unit": "cum", "base_qty": 1.0,
        "rows": [
            ("Stone Aggregate 40 mm", "Coarse aggregate 40 mm for mass foundation PCC bed (0.60 cum)", "Clean crushed stone aggregate 40 mm nominal size", "Material", "0.60 cum (0.600 cum/cum)", "1 cum"),
            ("Stone Aggregate 20 mm", "Coarse aggregate 20 mm filling large aggregate voids (0.30 cum)", "Clean crushed stone aggregate 20 mm nominal size", "Material", "0.30 cum (0.300 cum/cum)", "1 cum"),
            ("Carriage of Stone aggregate 40 mm", "Haulage and unloading of 40 mm aggregate (0.60 cum)", "Carriage of 40 mm aggregate", "Material", "0.60 cum (0.600 cum/cum)", "1 cum"),
            ("Carriage of Stone aggregate below 40 mm", "Haulage and unloading of 20 mm aggregate (0.30 cum)", "Carriage of 20 mm aggregate", "Material", "0.30 cum (0.300 cum/cum)", "1 cum"),
            ("Coarse sand", "Zone III coarse sand matrix (0.45 cum)", "Clean natural river sand Zone III", "Material", "0.45 cum (0.450 cum/cum)", "1 cum"),
            ("Carriage of Coarse sand", "Carriage of coarse sand to site platform (0.45 cum)", "Sand haulage", "Material", "0.45 cum (0.450 cum/cum)", "1 cum"),
            ("Portland Cement", "Hydraulic binder for mass foundation concrete (0.22 tonne)", "OPC-43 Grade cement conforming to IS:269", "Material", "0.22 tonne (0.220 t/cum)", "1 cum"),
            ("Carriage of Cement", "Transport of cement bags to mixing yard (0.22 tonne)", "Carriage of cement", "Material", "0.22 tonne (0.220 t/cum)", "1 cum"),
            ("Concrete Mixer 0.28 to 0.40 cum", "Machine batching and mixing 40 mm lean concrete", "Mixer operational hire; output = 1.79 cum/hr", "Machine", "0.56 machine-hrs (1.79 cum/hr)", "1 cum"),
            ("Needle Vibrator", "Compacting mass foundation concrete in trenches", "Vibrator hire; output = 1.79 cum/hr", "Machine", "0.56 machine-hrs (1.79 cum/hr)", "1 cum"),
            ("Mason (average)", "Leveling concrete foundation beds to exact invert levels", "Mason labour; 0.10 day = 0.80 man-hrs per cum", "Labour", "0.80 man-hrs (1.25 cum/man-hr)", "1 cum"),
            ("Beldar", "Measuring ingredients, charging mixer, shoveling and laying concrete", "Labour crew; 1.63 day = 13.04 man-hrs per cum", "Labour", "13.04 man-hrs (0.08 cum/man-hr)", "1 cum"),
            ("Coolie", "Carrying concrete pans and moving coarse aggregates", "Unskilled labour; 0.70 day = 5.60 man-hrs per cum", "Labour", "5.60 man-hrs (0.18 cum/man-hr)", "1 cum"),
            ("Bhisti", "Water batching, drum cleaning, and foundation curing", "Watering crew; 0.70 day = 5.60 man-hrs per cum", "Labour", "5.60 man-hrs (0.18 cum/man-hr)", "1 cum"),
            ("Sundries", "Sundry tools, farma boxes, wheelbarrows, and testing tools", "Sundries allowance", "Equipment", "L.S. allowance", "1 cum"),
        ]
    },
    {
        "id": "4.1.8",
        "parent_title": "4.1 CEMENT CONCRETE UP TO PLINTH LEVEL (EXCL. SHUTTERING)",
        "title": "4.1.8 Cement concrete 1:4:8 (1 cement : 4 coarse sand : 8 graded stone aggregate 40 mm)",
        "unit": "cum", "base_qty": 1.0,
        "rows": [
            ("Stone Aggregate 40 mm", "Coarse aggregate 40 mm for mass levelling course (0.63 cum)", "Clean crushed stone aggregate 40 mm nominal size", "Material", "0.63 cum (0.630 cum/cum)", "1 cum"),
            ("Stone Aggregate 20 mm", "Coarse aggregate 20 mm for void packing (0.31 cum)", "Clean crushed stone aggregate 20 mm nominal size", "Material", "0.31 cum (0.310 cum/cum)", "1 cum"),
            ("Carriage of Stone aggregate 40 mm", "Haulage and unloading of 40 mm aggregate (0.63 cum)", "Carriage of 40 mm aggregate", "Material", "0.63 cum (0.630 cum/cum)", "1 cum"),
            ("Carriage of Stone aggregate below 40 mm", "Haulage and unloading of 20 mm aggregate (0.31 cum)", "Carriage of 20 mm aggregate", "Material", "0.31 cum (0.310 cum/cum)", "1 cum"),
            ("Coarse sand", "Zone III coarse sand matrix (0.47 cum)", "Clean natural river sand Zone III", "Material", "0.47 cum (0.470 cum/cum)", "1 cum"),
            ("Carriage of Coarse sand", "Carriage of coarse sand to site platform (0.47 cum)", "Sand haulage", "Material", "0.47 cum (0.470 cum/cum)", "1 cum"),
            ("Portland Cement", "Hydraulic binder for 1:4:8 lean mud-mat concrete (0.17 tonne = 3.4 bags)", "OPC-43 Grade cement conforming to IS:269", "Material", "0.17 tonne (0.170 t/cum)", "1 cum"),
            ("Carriage of Cement", "Transport of cement bags to mixing yard (0.17 tonne)", "Carriage of cement", "Material", "0.17 tonne (0.170 t/cum)", "1 cum"),
            ("Concrete Mixer 0.28 to 0.40 cum", "Machine batching and mixing 1:4:8 lean concrete", "Mixer operational hire; output = 1.79 cum/hr", "Machine", "0.56 machine-hrs (1.79 cum/hr)", "1 cum"),
            ("Needle Vibrator", "Compacting sub-base lean concrete", "Vibrator hire; output = 1.79 cum/hr", "Machine", "0.56 machine-hrs (1.79 cum/hr)", "1 cum"),
            ("Mason (average)", "Leveling screed course under footing rafts and flooring bases", "Mason labour; 0.10 day = 0.80 man-hrs per cum", "Labour", "0.80 man-hrs (1.25 cum/man-hr)", "1 cum"),
            ("Beldar", "Measuring ingredients, charging mixer, shoveling and laying concrete", "Labour crew; 1.63 day = 13.04 man-hrs per cum", "Labour", "13.04 man-hrs (0.08 cum/man-hr)", "1 cum"),
            ("Coolie", "Carrying concrete pans and moving coarse aggregates", "Unskilled labour; 0.70 day = 5.60 man-hrs per cum", "Labour", "5.60 man-hrs (0.18 cum/man-hr)", "1 cum"),
            ("Bhisti", "Water batching, drum cleaning, and mud-mat curing", "Watering crew; 0.70 day = 5.60 man-hrs per cum", "Labour", "5.60 man-hrs (0.18 cum/man-hr)", "1 cum"),
            ("Sundries", "Sundry tools, farma boxes, wheelbarrows, and testing tools", "Sundries allowance", "Equipment", "L.S. allowance", "1 cum"),
        ]
    },
    {
        "id": "4.1.10",
        "parent_title": "4.1 CEMENT CONCRETE UP TO PLINTH LEVEL (EXCL. SHUTTERING)",
        "title": "4.1.10 Cement concrete 1:5:10 (1 cement : 5 coarse sand : 10 graded stone aggregate 40 mm)",
        "unit": "cum", "base_qty": 1.0,
        "rows": [
            ("Stone Aggregate 40 mm", "Coarse aggregate 40 mm for lean mass foundations (0.60 cum)", "Clean crushed stone aggregate 40 mm nominal size", "Material", "0.60 cum (0.600 cum/cum)", "1 cum"),
            ("Stone Aggregate 20 mm", "Coarse aggregate 20 mm for void filling (0.30 cum)", "Clean crushed stone aggregate 20 mm nominal size", "Material", "0.30 cum (0.300 cum/cum)", "1 cum"),
            ("Carriage of Stone aggregate 40 mm", "Haulage and unloading of 40 mm aggregate (0.60 cum)", "Carriage of 40 mm aggregate", "Material", "0.60 cum (0.600 cum/cum)", "1 cum"),
            ("Carriage of Stone aggregate below 40 mm", "Haulage and unloading of 20 mm aggregate (0.30 cum)", "Carriage of 20 mm aggregate", "Material", "0.30 cum (0.300 cum/cum)", "1 cum"),
            ("Coarse sand", "Zone III coarse sand matrix (0.45 cum)", "Clean natural river sand Zone III", "Material", "0.45 cum (0.450 cum/cum)", "1 cum"),
            ("Carriage of Coarse sand", "Carriage of coarse sand to site platform (0.45 cum)", "Sand haulage", "Material", "0.45 cum (0.450 cum/cum)", "1 cum"),
            ("Portland Cement", "Hydraulic binder for 1:5:10 mass concrete (0.13 tonne = 2.6 bags)", "OPC-43 Grade cement conforming to IS:269", "Material", "0.13 tonne (0.130 t/cum)", "1 cum"),
            ("Carriage of Cement", "Transport of cement bags to mixing yard (0.13 tonne)", "Carriage of cement", "Material", "0.13 tonne (0.130 t/cum)", "1 cum"),
            ("Concrete Mixer 0.28 to 0.40 cum", "Machine batching and mixing 1:5:10 lean concrete", "Mixer operational hire; output = 1.79 cum/hr", "Machine", "0.56 machine-hrs (1.79 cum/hr)", "1 cum"),
            ("Needle Vibrator", "Compacting sub-base lean concrete", "Vibrator hire; output = 1.79 cum/hr", "Machine", "0.56 machine-hrs (1.79 cum/hr)", "1 cum"),
            ("Mason (average)", "Leveling screed course under mass masonry foundations", "Mason labour; 0.10 day = 0.80 man-hrs per cum", "Labour", "0.80 man-hrs (1.25 cum/man-hr)", "1 cum"),
            ("Beldar", "Measuring ingredients, charging mixer, shoveling and laying concrete", "Labour crew; 1.63 day = 13.04 man-hrs per cum", "Labour", "13.04 man-hrs (0.08 cum/man-hr)", "1 cum"),
            ("Coolie", "Carrying concrete pans and moving coarse aggregates", "Unskilled labour; 0.70 day = 5.60 man-hrs per cum", "Labour", "5.60 man-hrs (0.18 cum/man-hr)", "1 cum"),
            ("Bhisti", "Water batching, drum cleaning, and mud-mat curing", "Watering crew; 0.70 day = 5.60 man-hrs per cum", "Labour", "5.60 man-hrs (0.18 cum/man-hr)", "1 cum"),
            ("Sundries", "Sundry tools, farma boxes, wheelbarrows, and testing tools", "Sundries allowance", "Equipment", "L.S. allowance", "1 cum"),
        ]
    },
    {
        "id": "4.1.11",
        "parent_title": "4.1 CEMENT CONCRETE UP TO PLINTH LEVEL (EXCL. SHUTTERING)",
        "title": "4.1.11 Cement concrete 1:5:10 (1 cement : 5 fine sand : 10 graded stone aggregate 40 mm)",
        "unit": "cum", "base_qty": 1.0,
        "rows": [
            ("Stone Aggregate 40 mm", "Coarse aggregate 40 mm for mass foundation bed (0.60 cum)", "Clean crushed stone aggregate 40 mm nominal size", "Material", "0.60 cum (0.600 cum/cum)", "1 cum"),
            ("Stone Aggregate 20 mm", "Coarse aggregate 20 mm for void filling (0.30 cum)", "Clean crushed stone aggregate 20 mm nominal size", "Material", "0.30 cum (0.300 cum/cum)", "1 cum"),
            ("Carriage of Stone aggregate 40 mm", "Haulage and unloading of 40 mm aggregate (0.60 cum)", "Carriage of 40 mm aggregate", "Material", "0.60 cum (0.600 cum/cum)", "1 cum"),
            ("Carriage of Stone aggregate below 40 mm", "Haulage and unloading of 20 mm aggregate (0.30 cum)", "Carriage of 20 mm aggregate", "Material", "0.30 cum (0.300 cum/cum)", "1 cum"),
            ("Fine sand", "Zone IV fine sand matrix filler (0.45 cum)", "Clean fine sand conforming to IS:383 Zone IV", "Material", "0.45 cum (0.450 cum/cum)", "1 cum"),
            ("Carriage of Fine sand", "Mechanical transport and unloading of fine sand (0.45 cum)", "Carriage of fine sand", "Material", "0.45 cum (0.450 cum/cum)", "1 cum"),
            ("Portland Cement", "Hydraulic binder for 1:5:10 lean concrete (0.13 tonne)", "OPC-43 Grade cement conforming to IS:269", "Material", "0.13 tonne (0.130 t/cum)", "1 cum"),
            ("Carriage of Cement", "Transport of cement bags to mixing yard (0.13 tonne)", "Carriage of cement", "Material", "0.13 tonne (0.130 t/cum)", "1 cum"),
            ("Concrete Mixer 0.28 to 0.40 cum", "Machine batching and mixing 1:5:10 fine sand concrete", "Mixer operational hire; output = 1.79 cum/hr", "Machine", "0.56 machine-hrs (1.79 cum/hr)", "1 cum"),
            ("Needle Vibrator", "Compacting sub-base lean concrete", "Vibrator hire; output = 1.79 cum/hr", "Machine", "0.56 machine-hrs (1.79 cum/hr)", "1 cum"),
            ("Mason (average)", "Leveling screed course under mass masonry foundations", "Mason labour; 0.10 day = 0.80 man-hrs per cum", "Labour", "0.80 man-hrs (1.25 cum/man-hr)", "1 cum"),
            ("Beldar", "Measuring ingredients, charging mixer, shoveling and laying concrete", "Labour crew; 1.63 day = 13.04 man-hrs per cum", "Labour", "13.04 man-hrs (0.08 cum/man-hr)", "1 cum"),
            ("Coolie", "Carrying concrete pans and moving aggregates", "Unskilled labour; 0.70 day = 5.60 man-hrs per cum", "Labour", "5.60 man-hrs (0.18 cum/man-hr)", "1 cum"),
            ("Bhisti", "Water batching, drum cleaning, and foundation curing", "Watering crew; 0.70 day = 5.60 man-hrs per cum", "Labour", "5.60 man-hrs (0.18 cum/man-hr)", "1 cum"),
            ("Sundries", "Sundry tools, farma boxes, wheelbarrows, and testing tools", "Sundries allowance", "Equipment", "L.S. allowance", "1 cum"),
        ]
    },
    {
        "id": "4.1.12",
        "parent_title": "4.1 CEMENT CONCRETE UP TO PLINTH LEVEL (EXCL. SHUTTERING)",
        "title": "4.1.12 Cement concrete 1:2:3.5:9 (1 cement : 2 fly ash : 3.5 coarse sand : 9 graded stone aggregate 40 mm)",
        "unit": "cum", "base_qty": 1.0,
        "rows": [
            ("Stone Aggregate 40 mm", "Coarse aggregate 40 mm for fly-ash blend foundation concrete (0.57 cum)", "Clean crushed stone aggregate 40 mm nominal size", "Material", "0.57 cum (0.570 cum/cum)", "1 cum"),
            ("Stone Aggregate 20 mm", "Coarse aggregate 20 mm for void packing (0.28 cum)", "Clean crushed stone aggregate 20 mm nominal size", "Material", "0.28 cum (0.280 cum/cum)", "1 cum"),
            ("Carriage of Stone aggregate 40 mm", "Haulage of 40 mm aggregate (0.57 cum)", "Carriage of 40 mm aggregate", "Material", "0.57 cum (0.570 cum/cum)", "1 cum"),
            ("Carriage of Stone aggregate below 40 mm", "Haulage of 20 mm aggregate (0.28 cum)", "Carriage of 20 mm aggregate", "Material", "0.28 cum (0.280 cum/cum)", "1 cum"),
            ("Coarse sand", "Zone III coarse sand matrix (0.33 cum)", "Clean natural river sand Zone III", "Material", "0.33 cum (0.330 cum/cum)", "1 cum"),
            ("Carriage of Coarse sand", "Carriage of coarse sand (0.33 cum)", "Sand carriage", "Material", "0.33 cum (0.330 cum/cum)", "1 cum"),
            ("Fly ash", "Pulverized fuel ash (pozzolana) replacing part of binder and fine aggregate (0.19 tonne)", "Conforming to IS:3812 Part 1 for pozzolanic concrete blends", "Material", "0.19 tonne (0.190 t/cum)", "1 cum"),
            ("Carriage of Fly ash", "Haulage and protected handling of fly ash to batching yard (0.19 tonne)", "Carriage of fly ash", "Material", "0.19 tonne (0.190 t/cum)", "1 cum"),
            ("Portland Cement", "Hydraulic binder in fly-ash blended concrete (0.13 tonne)", "OPC-43 Grade cement conforming to IS:269", "Material", "0.13 tonne (0.130 t/cum)", "1 cum"),
            ("Carriage of Cement", "Transport of cement bags to mixing yard (0.13 tonne)", "Carriage of cement", "Material", "0.13 tonne (0.130 t/cum)", "1 cum"),
            ("Concrete Mixer 0.28 to 0.40 cum", "Intensive mixing to ensure uniform pozzolanic fly ash blending", "Mixer operational hire; output = 1.79 cum/hr", "Machine", "0.56 machine-hrs (1.79 cum/hr)", "1 cum"),
            ("Needle Vibrator", "Compacting blended foundation concrete", "Vibrator hire; output = 1.79 cum/hr", "Machine", "0.56 machine-hrs (1.79 cum/hr)", "1 cum"),
            ("Mason (average)", "Leveling and finishing concrete sub-base", "Mason labour; 0.10 day = 0.80 man-hrs per cum", "Labour", "0.80 man-hrs (1.25 cum/man-hr)", "1 cum"),
            ("Beldar", "Measuring ingredients, charging mixer, shoveling and laying concrete", "Labour crew; 1.63 day = 13.04 man-hrs per cum", "Labour", "13.04 man-hrs (0.08 cum/man-hr)", "1 cum"),
            ("Coolie", "Carrying concrete pans and moving aggregates", "Unskilled labour; 0.70 day = 5.60 man-hrs per cum", "Labour", "5.60 man-hrs (0.18 cum/man-hr)", "1 cum"),
            ("Bhisti", "Water batching, drum cleaning, and curing", "Watering crew; 0.70 day = 5.60 man-hrs per cum", "Labour", "5.60 man-hrs (0.18 cum/man-hr)", "1 cum"),
            ("Sundries", "Sundry tools, farma boxes, wheelbarrows, and testing tools", "Sundries allowance", "Equipment", "L.S. allowance", "1 cum"),
        ]
    },
    {
        "id": "4.1.13",
        "parent_title": "4.1 CEMENT CONCRETE UP TO PLINTH LEVEL (EXCL. SHUTTERING)",
        "title": "4.1.13 Cement concrete 1:2.5:4:11 (1 cement : 2.5 fly ash : 4 coarse sand : 11 graded stone aggregate 40 mm)",
        "unit": "cum", "base_qty": 1.0,
        "rows": [
            ("Stone Aggregate 40 mm", "Coarse aggregate 40 mm for lean fly-ash foundation concrete (0.57 cum)", "Clean crushed stone aggregate 40 mm nominal size", "Material", "0.57 cum (0.570 cum/cum)", "1 cum"),
            ("Stone Aggregate 20 mm", "Coarse aggregate 20 mm for void packing (0.28 cum)", "Clean crushed stone aggregate 20 mm nominal size", "Material", "0.28 cum (0.280 cum/cum)", "1 cum"),
            ("Carriage of Stone aggregate 40 mm", "Haulage of 40 mm aggregate (0.57 cum)", "Carriage of 40 mm aggregate", "Material", "0.57 cum (0.570 cum/cum)", "1 cum"),
            ("Carriage of Stone aggregate below 40 mm", "Haulage of 20 mm aggregate (0.28 cum)", "Carriage of 20 mm aggregate", "Material", "0.28 cum (0.280 cum/cum)", "1 cum"),
            ("Coarse sand", "Zone III coarse sand matrix (0.31 cum)", "Clean natural river sand Zone III", "Material", "0.31 cum (0.310 cum/cum)", "1 cum"),
            ("Carriage of Coarse sand", "Carriage of coarse sand (0.31 cum)", "Sand carriage", "Material", "0.31 cum (0.310 cum/cum)", "1 cum"),
            ("Fly ash", "Pulverized fuel ash (pozzolana) blend (0.19 tonne)", "Conforming to IS:3812 Part 1", "Material", "0.19 tonne (0.190 t/cum)", "1 cum"),
            ("Carriage of Fly ash", "Haulage of fly ash to batching yard (0.19 tonne)", "Carriage of fly ash", "Material", "0.19 tonne (0.190 t/cum)", "1 cum"),
            ("Portland Cement", "Hydraulic binder in lean fly-ash concrete (0.10 tonne)", "OPC-43 Grade cement conforming to IS:269", "Material", "0.10 tonne (0.100 t/cum)", "1 cum"),
            ("Carriage of Cement", "Transport of cement bags to mixing yard (0.10 tonne)", "Carriage of cement", "Material", "0.10 tonne (0.100 t/cum)", "1 cum"),
            ("Concrete Mixer 0.28 to 0.40 cum", "Intensive mixing of high fly-ash lean blend", "Mixer operational hire; output = 1.79 cum/hr", "Machine", "0.56 machine-hrs (1.79 cum/hr)", "1 cum"),
            ("Needle Vibrator", "Compacting sub-base lean concrete", "Vibrator hire; output = 1.79 cum/hr", "Machine", "0.56 machine-hrs (1.79 cum/hr)", "1 cum"),
            ("Mason (average)", "Leveling and finishing concrete sub-base", "Mason labour; 0.10 day = 0.80 man-hrs per cum", "Labour", "0.80 man-hrs (1.25 cum/man-hr)", "1 cum"),
            ("Beldar", "Measuring ingredients, charging mixer, shoveling and laying concrete", "Labour crew; 1.63 day = 13.04 man-hrs per cum", "Labour", "13.04 man-hrs (0.08 cum/man-hr)", "1 cum"),
            ("Coolie", "Carrying concrete pans and moving aggregates", "Unskilled labour; 0.70 day = 5.60 man-hrs per cum", "Labour", "5.60 man-hrs (0.18 cum/man-hr)", "1 cum"),
            ("Bhisti", "Water batching, drum cleaning, and curing", "Watering crew; 0.70 day = 5.60 man-hrs per cum", "Labour", "5.60 man-hrs (0.18 cum/man-hr)", "1 cum"),
            ("Sundries", "Sundry tools, farma boxes, wheelbarrows, and testing tools", "Sundries allowance", "Equipment", "L.S. allowance", "1 cum"),
        ]
    },

    # ─── 4.2 CONCRETE IN WALLS & RETAINING WALLS UP TO FLOOR V LEVEL ─────────────
    {
        "id": "4.2.2",
        "parent_title": "4.2 CEMENT CONCRETE IN WALLS & RETAINING WALLS UP TO FLOOR V LEVEL",
        "title": "4.2.2 Concrete in walls 1:1.5:3 (1 cement : 1.5 coarse sand : 3 graded stone aggregate 20 mm)",
        "unit": "cum", "base_qty": 1.0,
        "rows": [
            ("Stone Aggregate 20 mm", "Coarse aggregate 20 mm for retaining wall core (0.57 cum)", "Crushed stone aggregate 20 mm nominal size", "Material", "0.57 cum (0.570 cum/cum)", "1 cum"),
            ("Stone Aggregate 10 mm", "Coarse aggregate 10 mm for void packing (0.28 cum)", "Crushed stone aggregate 10 mm nominal size", "Material", "0.28 cum (0.280 cum/cum)", "1 cum"),
            ("Carriage of Stone aggregate below 40 mm", "Haulage of 20mm & 10mm aggregates (0.85 cum)", "Carriage to site platform", "Material", "0.85 cum (0.850 cum/cum)", "1 cum"),
            ("Coarse sand", "Zone III coarse sand matrix (0.425 cum)", "Clean natural river sand Zone III", "Material", "0.425 cum (0.425 cum/cum)", "1 cum"),
            ("Carriage of Coarse sand", "Carriage of coarse sand (0.425 cum)", "Sand carriage", "Material", "0.425 cum (0.425 cum/cum)", "1 cum"),
            ("Portland Cement", "Hydraulic binder for retaining wall concrete (0.40 tonne)", "OPC-43 Grade cement conforming to IS:269", "Material", "0.40 tonne (0.400 t/cum)", "1 cum"),
            ("Carriage of Cement", "Transport of cement bags to mixing point (0.40 tonne)", "Carriage of cement", "Material", "0.40 tonne (0.400 t/cum)", "1 cum"),
            ("Concrete Mixer 0.28 to 0.40 cum", "Machine batching and mixing concrete", "Mixer operational hire; output = 1.79 cum/hr", "Machine", "0.56 machine-hrs (1.79 cum/hr)", "1 cum"),
            ("Needle Vibrator", "Compacting concrete in narrow wall forms between vertical reinforcement/ties", "Needle vibrator hire; output = 1.79 cum/hr", "Machine", "0.56 machine-hrs (1.79 cum/hr)", "1 cum"),
            ("Mason (average)", "Controlling wall lift placement, checking vertical alignments, and screeding tops", "Mason labour; 0.18 day = 1.44 man-hrs per cum", "Labour", "1.44 man-hrs (0.69 cum/man-hr)", "1 cum"),
            ("Beldar", "Feeding mixer, hoisting pans, placing in narrow forms in layers not exceeding 30 cm", "Labour crew; 2.24 day = 17.92 man-hrs per cum", "Labour", "17.92 man-hrs (0.06 cum/man-hr)", "1 cum"),
            ("Coolie", "Hoisting and handling concrete pans to staging platforms up to floor V level", "Haulage crew; 1.63 day = 13.04 man-hrs per cum", "Labour", "13.04 man-hrs (0.08 cum/man-hr)", "1 cum"),
            ("Bhisti", "Water batching, cleaning mixer drum, and curing vertical walls with hessian cloth", "Watering crew; 0.70 day = 5.60 man-hrs per cum", "Labour", "5.60 man-hrs (0.18 cum/man-hr)", "1 cum"),
            ("Sundries", "Hoisting ropes, pulleys, buckets, scaffolding clamps, and vibrator spares", "Sundries allowance", "Equipment", "L.S. allowance", "1 cum"),
        ]
    },
    {
        "id": "4.2.3",
        "parent_title": "4.2 CEMENT CONCRETE IN WALLS & RETAINING WALLS UP TO FLOOR V LEVEL",
        "title": "4.2.3 Concrete in walls 1:2:4 (1 cement : 2 coarse sand : 4 graded stone aggregate 20 mm)",
        "unit": "cum", "base_qty": 1.0,
        "rows": [
            ("Stone Aggregate 20 mm", "Coarse aggregate 20 mm for retaining wall core (0.57 cum)", "Crushed stone aggregate 20 mm nominal size", "Material", "0.57 cum (0.570 cum/cum)", "1 cum"),
            ("Stone Aggregate 10 mm", "Coarse aggregate 10 mm for void packing (0.28 cum)", "Crushed stone aggregate 10 mm nominal size", "Material", "0.28 cum (0.280 cum/cum)", "1 cum"),
            ("Carriage of Stone aggregate below 40 mm", "Haulage of aggregates (0.85 cum)", "Carriage to site platform", "Material", "0.85 cum (0.850 cum/cum)", "1 cum"),
            ("Coarse sand", "Zone III coarse sand matrix (0.425 cum)", "Clean natural river sand Zone III", "Material", "0.425 cum (0.425 cum/cum)", "1 cum"),
            ("Carriage of Coarse sand", "Carriage of coarse sand (0.425 cum)", "Sand carriage", "Material", "0.425 cum (0.425 cum/cum)", "1 cum"),
            ("Portland Cement", "Hydraulic binder for retaining wall concrete (0.32 tonne)", "OPC-43 Grade cement conforming to IS:269", "Material", "0.32 tonne (0.320 t/cum)", "1 cum"),
            ("Carriage of Cement", "Transport of cement bags to mixing point (0.32 tonne)", "Carriage of cement", "Material", "0.32 tonne (0.320 t/cum)", "1 cum"),
            ("Concrete Mixer 0.28 to 0.40 cum", "Machine batching and mixing concrete", "Mixer operational hire; output = 1.79 cum/hr", "Machine", "0.56 machine-hrs (1.79 cum/hr)", "1 cum"),
            ("Needle Vibrator", "Compacting concrete in narrow wall forms", "Needle vibrator hire; output = 1.79 cum/hr", "Machine", "0.56 machine-hrs (1.79 cum/hr)", "1 cum"),
            ("Mason (average)", "Controlling wall lift placement and checking vertical alignment", "Mason labour; 0.18 day = 1.44 man-hrs per cum", "Labour", "1.44 man-hrs (0.69 cum/man-hr)", "1 cum"),
            ("Beldar", "Feeding mixer, hoisting pans, placing in narrow forms in layers", "Labour crew; 2.24 day = 17.92 man-hrs per cum", "Labour", "17.92 man-hrs (0.06 cum/man-hr)", "1 cum"),
            ("Coolie", "Hoisting and handling concrete pans to staging platforms up to floor V level", "Haulage crew; 1.63 day = 13.04 man-hrs per cum", "Labour", "13.04 man-hrs (0.08 cum/man-hr)", "1 cum"),
            ("Bhisti", "Water batching, mixer cleanup, and curing vertical walls with wet hessian", "Watering crew; 0.70 day = 5.60 man-hrs per cum", "Labour", "5.60 man-hrs (0.18 cum/man-hr)", "1 cum"),
            ("Sundries", "Hoisting ropes, pulleys, buckets, scaffolding clamps, and vibrator spares", "Sundries allowance", "Equipment", "L.S. allowance", "1 cum"),
        ]
    },
    {
        "id": "4.2.5",
        "parent_title": "4.2 CEMENT CONCRETE IN WALLS & RETAINING WALLS UP TO FLOOR V LEVEL",
        "title": "4.2.5 Concrete in walls 1:3:6 (1 cement : 3 coarse sand : 6 graded stone aggregate 20 mm)",
        "unit": "cum", "base_qty": 1.0,
        "rows": [
            ("Stone Aggregate 20 mm", "Coarse aggregate 20 mm for retaining wall mass core (0.60 cum)", "Crushed stone aggregate 20 mm nominal size", "Material", "0.60 cum (0.600 cum/cum)", "1 cum"),
            ("Stone Aggregate 10 mm", "Coarse aggregate 10 mm for void packing (0.30 cum)", "Crushed stone aggregate 10 mm nominal size", "Material", "0.30 cum (0.300 cum/cum)", "1 cum"),
            ("Carriage of Stone aggregate below 40 mm", "Haulage of aggregates (0.90 cum)", "Carriage to site platform", "Material", "0.90 cum (0.900 cum/cum)", "1 cum"),
            ("Coarse sand", "Zone III coarse sand matrix (0.45 cum)", "Clean natural river sand Zone III", "Material", "0.45 cum (0.450 cum/cum)", "1 cum"),
            ("Carriage of Coarse sand", "Carriage of coarse sand (0.45 cum)", "Sand carriage", "Material", "0.45 cum (0.450 cum/cum)", "1 cum"),
            ("Portland Cement", "Hydraulic binder for mass wall concrete (0.22 tonne)", "OPC-43 Grade cement conforming to IS:269", "Material", "0.22 tonne (0.220 t/cum)", "1 cum"),
            ("Carriage of Cement", "Transport of cement bags to mixing point (0.22 tonne)", "Carriage of cement", "Material", "0.22 tonne (0.220 t/cum)", "1 cum"),
            ("Concrete Mixer 0.28 to 0.40 cum", "Machine batching and mixing concrete", "Mixer operational hire; output = 1.79 cum/hr", "Machine", "0.56 machine-hrs (1.79 cum/hr)", "1 cum"),
            ("Needle Vibrator", "Compacting concrete in mass wall forms", "Needle vibrator hire; output = 1.79 cum/hr", "Machine", "0.56 machine-hrs (1.79 cum/hr)", "1 cum"),
            ("Mason (average)", "Controlling wall lift placement and checking vertical alignment", "Mason labour; 0.18 day = 1.44 man-hrs per cum", "Labour", "1.44 man-hrs (0.69 cum/man-hr)", "1 cum"),
            ("Beldar", "Feeding mixer, hoisting pans, placing in wall forms in layers", "Labour crew; 2.24 day = 17.92 man-hrs per cum", "Labour", "17.92 man-hrs (0.06 cum/man-hr)", "1 cum"),
            ("Coolie", "Hoisting and handling concrete pans to staging platforms up to floor V level", "Haulage crew; 1.63 day = 13.04 man-hrs per cum", "Labour", "13.04 man-hrs (0.08 cum/man-hr)", "1 cum"),
            ("Bhisti", "Water batching, mixer cleanup, and curing vertical walls with wet hessian", "Watering crew; 0.70 day = 5.60 man-hrs per cum", "Labour", "5.60 man-hrs (0.18 cum/man-hr)", "1 cum"),
            ("Sundries", "Hoisting ropes, pulleys, buckets, scaffolding clamps, and vibrator spares", "Sundries allowance", "Equipment", "L.S. allowance", "1 cum"),
        ]
    },
    {
        "id": "4.2.8",
        "parent_title": "4.2 CEMENT CONCRETE IN WALLS & RETAINING WALLS UP TO FLOOR V LEVEL",
        "title": "4.2.8 Concrete in walls 1:5:10 (1 cement : 5 coarse sand : 10 graded stone aggregate 40 mm)",
        "unit": "cum", "base_qty": 1.0,
        "rows": [
            ("Stone Aggregate 40 mm", "Coarse aggregate 40 mm for mass gravity retaining wall (0.60 cum)", "Crushed stone aggregate 40 mm nominal size", "Material", "0.60 cum (0.600 cum/cum)", "1 cum"),
            ("Stone Aggregate 20 mm", "Coarse aggregate 20 mm for void packing (0.30 cum)", "Crushed stone aggregate 20 mm nominal size", "Material", "0.30 cum (0.300 cum/cum)", "1 cum"),
            ("Carriage of Stone aggregate 40 mm", "Haulage of 40 mm aggregate (0.60 cum)", "Carriage of 40 mm aggregate", "Material", "0.60 cum (0.600 cum/cum)", "1 cum"),
            ("Carriage of Stone aggregate below 40 mm", "Haulage of 20 mm aggregate (0.30 cum)", "Carriage of 20 mm aggregate", "Material", "0.30 cum (0.300 cum/cum)", "1 cum"),
            ("Coarse sand", "Zone III coarse sand matrix (0.45 cum)", "Clean natural river sand Zone III", "Material", "0.45 cum (0.450 cum/cum)", "1 cum"),
            ("Carriage of Coarse sand", "Carriage of coarse sand (0.45 cum)", "Sand carriage", "Material", "0.45 cum (0.450 cum/cum)", "1 cum"),
            ("Portland Cement", "Hydraulic binder for mass gravity walls (0.13 tonne)", "OPC-43 Grade cement conforming to IS:269", "Material", "0.13 tonne (0.130 t/cum)", "1 cum"),
            ("Carriage of Cement", "Transport of cement bags to mixing point (0.13 tonne)", "Carriage of cement", "Material", "0.13 tonne (0.130 t/cum)", "1 cum"),
            ("Concrete Mixer 0.28 to 0.40 cum", "Machine batching and mixing concrete", "Mixer operational hire; output = 1.79 cum/hr", "Machine", "0.56 machine-hrs (1.79 cum/hr)", "1 cum"),
            ("Needle Vibrator", "Compacting mass gravity concrete", "Needle vibrator hire; output = 1.79 cum/hr", "Machine", "0.56 machine-hrs (1.79 cum/hr)", "1 cum"),
            ("Mason (average)", "Controlling wall lift placement and checking vertical alignment", "Mason labour; 0.18 day = 1.44 man-hrs per cum", "Labour", "1.44 man-hrs (0.69 cum/man-hr)", "1 cum"),
            ("Beldar", "Feeding mixer, hoisting pans, placing in wall forms in layers", "Labour crew; 2.24 day = 17.92 man-hrs per cum", "Labour", "17.92 man-hrs (0.06 cum/man-hr)", "1 cum"),
            ("Coolie", "Hoisting and handling concrete pans to staging platforms up to floor V level", "Haulage crew; 1.63 day = 13.04 man-hrs per cum", "Labour", "13.04 man-hrs (0.08 cum/man-hr)", "1 cum"),
            ("Bhisti", "Water batching, mixer cleanup, and curing vertical walls with wet hessian", "Watering crew; 0.70 day = 5.60 man-hrs per cum", "Labour", "5.60 man-hrs (0.18 cum/man-hr)", "1 cum"),
            ("Sundries", "Hoisting ropes, pulleys, buckets, scaffolding clamps, and vibrator spares", "Sundries allowance", "Equipment", "L.S. allowance", "1 cum"),
        ]
    },

    # ─── 4.3 CENTERING & SHUTTERING ─────────────────────────────────────────────
    {
        "id": "4.3.1",
        "parent_title": "4.3 CENTERING & SHUTTERING (INCLUDING STRUTTING, PROPPING & REMOVAL)",
        "title": "4.3.1 Centering and shuttering for foundations, footings, bases for columns (10 sqm)",
        "unit": "sqm", "base_qty": 10.0,
        "rows": [
            ("Shuttering plywood / timber planking", "Form surface contacting wet concrete; resisting hydrostatic concrete pressure (0.027 cum timber)", "12 mm water-proof shuttering plywood or 25 mm thick timber planks (6-use lifespan)", "Material", "0.027 cum (0.0027 cum/sqm)", "10 sqm"),
            ("Timber battens / runners", "Horizontal and vertical stiffening battens supporting form face (0.025 cum timber)", "Sal/hardwood battens 50x50mm and 75x50mm (10-use lifespan)", "Material", "0.025 cum (0.0025 cum/sqm)", "10 sqm"),
            ("Timber props / struts", "Lateral bracing and vertical inclined props preventing lateral displacement (0.030 cum timber)", "Sal ballies 80-100mm dia or wooden struts (15-use lifespan)", "Material", "0.030 cum (0.0030 cum/sqm)", "10 sqm"),
            ("Nails, spikes and binding wire", "Securing timber battens, corner clamps, and tie-wires (0.40 kg)", "Standard MS wire nails and galvanized binding wire", "Material", "0.40 kg (0.040 kg/sqm)", "10 sqm"),
            ("Shuttering mould release oil", "Chemical barrier coat preventing concrete adhesion and ensuring clean form release (0.60 litre)", "Approved non-staining mineral mould release oil", "Material", "0.60 litre (0.060 L/sqm)", "10 sqm"),
            ("Carpenter (1st & 2nd class)", "Measuring, cutting forms, erecting shuttering, plumbing, bracing, and striking forms after 24-48 hrs", "Skilled carpentry crew; 0.40 day = 3.20 man-hrs per 10 sqm", "Labour", "3.20 man-hrs (3.12 sqm/man-hr)", "10 sqm"),
            ("Beldar / Coolie", "Carrying formwork materials, holding props during alignment, de-shuttering, cleaning and oiling panels", "Helper crew; 0.40 day = 3.20 man-hrs per 10 sqm", "Labour", "3.20 man-hrs (3.12 sqm/man-hr)", "10 sqm"),
            ("Sundries", "Chalk line, plumb bob, spirit level, crowbars, and timber preservative treatment", "Sundries allowance", "Equipment", "L.S. allowance", "10 sqm"),
        ]
    },
    {
        "id": "4.3.2",
        "parent_title": "4.3 CENTERING & SHUTTERING (INCLUDING STRUTTING, PROPPING & REMOVAL)",
        "title": "4.3.2 Centering and shuttering for retaining walls, return walls, walls (any thickness) (10 sqm)",
        "unit": "sqm", "base_qty": 10.0,
        "rows": [
            ("Shuttering plywood / steel plates", "Vertical form panels resisting lateral hydrostatic wet concrete pressure (0.035 cum timber)", "High-density 12 mm shuttering plywood or MS steel plates (6-use lifespan)", "Material", "0.035 cum (0.0035 cum/sqm)", "10 sqm"),
            ("Timber walers and strongbacks", "Heavy horizontal walers (100x75mm) distributing lateral forces to tie rods (0.032 cum timber)", "Sal wood heavy walers resisting bulging deflection (10-use lifespan)", "Material", "0.032 cum (0.0032 cum/sqm)", "10 sqm"),
            ("Props and inclined push-pull struts", "Tubular steel props or sal ballies providing lateral stability and vertical plumb (0.040 cum timber)", "Heavy-duty adjustable struts supporting vertical wall forms", "Material", "0.040 cum (0.0040 cum/sqm)", "10 sqm"),
            ("Form ties, through-bolts and PVC sleeves", "MS through tie-rods with PVC sleeves preventing wall spread under vibration (0.80 kg)", "MS tie-rods with cone washers preventing slurry leakage", "Material", "0.80 kg (0.080 kg/sqm)", "10 sqm"),
            ("Shuttering mould release oil", "Applying barrier coat on vertical form faces prior to concrete pouring (0.80 litre)", "Non-staining mineral release agent", "Material", "0.80 litre (0.080 L/sqm)", "10 sqm"),
            ("Carpenter (1st & 2nd class)", "Fabricating wall panels, fixing tie rods, plumbing to vertical tolerance, and safe de-shuttering", "Skilled carpentry crew; 0.65 day = 5.20 man-hrs per 10 sqm", "Labour", "5.20 man-hrs (1.92 sqm/man-hr)", "10 sqm"),
            ("Beldar / Coolie", "Lifting panels, tightening tie-bolt nuts, assisting alignment, stripping forms, and scraping mortar", "Helper crew; 0.65 day = 5.20 man-hrs per 10 sqm", "Labour", "5.20 man-hrs (1.92 sqm/man-hr)", "10 sqm"),
            ("Sundries", "Scaffolding clamps, turnbuckles, sealing foam tape, and cleaning brushes", "Sundries allowance", "Equipment", "L.S. allowance", "10 sqm"),
        ]
    },
    {
        "id": "4.3.3",
        "parent_title": "4.3 CENTERING & SHUTTERING (INCLUDING STRUTTING, PROPPING & REMOVAL)",
        "title": "4.3.3 Centering and shuttering for columns, piers, abutments, pillars, posts and struts (10 sqm)",
        "unit": "sqm", "base_qty": 10.0,
        "rows": [
            ("Shuttering plywood / steel column forms", "Four-sided column box resisting intense bursting pressure during vibrated pouring (0.042 cum timber)", "12 mm film-faced plywood or heavy fabricated steel column casings (8-use lifespan)", "Material", "0.042 cum (0.0042 cum/sqm)", "10 sqm"),
            ("Column yokes / steel clamps", "Heavy horizontal column clamps (yokes) spaced at close intervals (0.038 cum timber)", "Adjustable MS column clamps or heavy timber yokes resisting lateral bursting forces", "Material", "0.038 cum (0.0038 cum/sqm)", "10 sqm"),
            ("Diagonal bracing props", "Four-way diagonal push-pull props holding column plumb in orthogonal directions (0.045 cum timber)", "Steel telescopic props or heavy sal ballies (15-use lifespan)", "Material", "0.045 cum (0.0045 cum/sqm)", "10 sqm"),
            ("Clamping bolts, wedges and foam tape", "High-tensile clamping bolts and foam tape on corner joints preventing cement slurry leakage (1.00 kg)", "Clamping hardware and gasket tape", "Material", "1.00 kg (0.100 kg/sqm)", "10 sqm"),
            ("Shuttering mould release oil", "Coating internal column faces before rebar cage closure (0.80 litre)", "Non-staining release agent", "Material", "0.80 litre (0.080 L/sqm)", "10 sqm"),
            ("Carpenter (1st & 2nd class)", "Box assembly, plumbing both faces with optical/plumb bobs, bolting yokes, and striking after 24 hrs", "Skilled carpentry crew; 0.80 day = 6.40 man-hrs per 10 sqm", "Labour", "6.40 man-hrs (1.56 sqm/man-hr)", "10 sqm"),
            ("Beldar / Coolie", "Hoisting box panels, tightening yoke wedges, holding props, and cleaning panels after stripping", "Helper crew; 0.80 day = 6.40 man-hrs per 10 sqm", "Labour", "6.40 man-hrs (1.56 sqm/man-hr)", "10 sqm"),
            ("Sundries", "Corner fillets (chamfer strips), plumb bobs, wrenches, and platform staging", "Sundries allowance", "Equipment", "L.S. allowance", "10 sqm"),
        ]
    },

    # ─── 4.4 TO 4.9 PRECAST & KERB CONCRETE ITEMS (DECOMPOSED) ───────────────────
    {
        "id": "4.4.1",
        "parent_title": "4.4 CONCRETE IN KERBS, STEPS AT OR NEAR GROUND LEVEL",
        "title": "4.4.1 Cement concrete in kerbs, steps 1:1.5:3 (1 cement : 1.5 coarse sand : 3 stone aggregate 20 mm)",
        "unit": "cum", "base_qty": 1.0,
        "rows": [
            ("Stone Aggregate 20 mm & 10 mm", "Coarse aggregate component decomposed from base PCC mix (REF#4.1.2) (0.85 cum)", "Graded 20mm (0.57 cum) + 10mm (0.28 cum) stone aggregate", "Material", "0.85 cum (0.850 cum/cum)", "1 cum"),
            ("Coarse sand", "Zone III coarse sand matrix decomposed from base PCC mix (REF#4.1.2) (0.425 cum)", "Clean natural river sand Zone III", "Material", "0.425 cum (0.425 cum/cum)", "1 cum"),
            ("Portland Cement", "Hydraulic binder decomposed from base PCC mix (REF#4.1.2) (0.40 tonne)", "OPC-43 Grade cement conforming to IS:269", "Material", "0.40 tonne (0.400 t/cum)", "1 cum"),
            ("Carriage of Aggregates & Cement", "Haulage of stone aggregate, coarse sand and cement to kerb laying site", "Combined carriage decomposed from REF#4.1.2", "Material", "1.675 cum/t equiv", "1 cum"),
            ("Concrete Mixer 0.28 to 0.40 cum", "Machine batching and mixing concrete (REF#4.1.2)", "Mixer operational hire; output = 1.79 cum/hr", "Machine", "0.56 machine-hrs (1.79 cum/hr)", "1 cum"),
            ("Needle / Plate Vibrator", "Compacting in situ concrete in kerbs and step moulds", "Vibrator hire; output = 1.79 cum/hr", "Machine", "0.56 machine-hrs (1.79 cum/hr)", "1 cum"),
            ("Mason (average)", "Forming kerb profiles, nosings of steps, rounded edges, and true alignment", "Combined REF#4.1.2 (0.10 d) + extra kerb shaping (0.15 d) = 0.25 day/cum", "Labour", "2.00 man-hrs (0.50 cum/man-hr)", "1 cum"),
            ("Beldar", "Measuring, mixer loading, shoveling, placing and ramming in kerb forms", "Combined REF#4.1.2 (1.63 d) + extra kerb handling (0.40 d) = 2.03 day/cum", "Labour", "16.24 man-hrs (0.06 cum/man-hr)", "1 cum"),
            ("Coolie", "Carrying concrete pans, handling materials, and assisting kerb alignment", "Labour allowance; 1.00 day = 8.00 man-hrs per cum", "Labour", "8.00 man-hrs (0.125 cum/man-hr)", "1 cum"),
            ("Bhisti", "Water batching, tool washing, and continuous kerb curing with wet gunny bags", "Watering crew; 0.70 day = 5.60 man-hrs per cum", "Labour", "5.60 man-hrs (0.18 cum/man-hr)", "1 cum"),
            ("REF#4.1.2 (Base PCC 1:1.5:3)", "Referenced base concrete formulation: Concrete 1:1.5:3 up to plinth level", "Base PCC scope executed per CPWD REF#4.1.2 specification", "Reference", "1.00 cum (1.000 cum/cum)", "1 cum"),
        ]
    },
    {
        "id": "4.5.1",
        "parent_title": "4.5 PRECAST CONCRETE STRING COURSES, COPINGS, BED BLOCKS UP TO FLOOR V",
        "title": "4.5.1 Precast concrete 1:1.5:3 in string courses, copings, sills, steps up to floor V level",
        "unit": "cum", "base_qty": 1.0,
        "rows": [
            ("Stone Aggregate 20 mm & 10 mm", "Coarse aggregate decomposed from base PCC mix (REF#4.1.2) (0.85 cum)", "Graded 20mm + 10mm crushed stone aggregate", "Material", "0.85 cum (0.850 cum/cum)", "1 cum"),
            ("Coarse sand", "Zone III coarse sand decomposed from base PCC (REF#4.1.2) & bedding mortar (REF#3.8) (0.50 cum)", "Coarse sand for precast concrete and 1:3 setting mortar", "Material", "0.50 cum (0.500 cum/cum)", "1 cum"),
            ("Portland Cement", "Cement for precast casting (0.40 t), 1:3 setting mortar (0.05 t), and 6mm face plaster (0.02 t)", "Total cement consumed in casting, hoisting, setting and plastering", "Material", "0.47 tonne (0.470 t/cum)", "1 cum"),
            ("Carriage of Materials", "Haulage of cement, aggregates, and precast units to installation points up to floor V", "Carriage and vertical shifting allowance", "Material", "1.85 cum/t equiv", "1 cum"),
            ("Concrete Mixer & Vibrator", "Batch mixing concrete and vibrating precast moulds on vibrating table", "Machine hire decomposed from REF#4.1.2; output = 1.79 cum/hr", "Machine", "0.56 machine-hrs (1.79 cum/hr)", "1 cum"),
            ("Precast casting moulds & shuttering", "Wooden/steel forms for casting copings, sills, and string courses with drip throats (REF#4.3)", "Mould hire, assembly, oiling, and dismantling", "Equipment", "Mould allowance", "1 cum"),
            ("Mason (average)", "Setting precast units true to line and level with 1:3 mortar and applying 6mm face plaster", "Skilled mason; 0.50 day = 4.00 man-hrs per cum", "Labour", "4.00 man-hrs (0.25 cum/man-hr)", "1 cum"),
            ("Beldar", "Casting, de-moulding, yard curing, hoisting to upper floors, and mixing setting mortar", "Labour crew; 2.80 day = 22.40 man-hrs per cum", "Labour", "22.40 man-hrs (0.04 cum/man-hr)", "1 cum"),
            ("Coolie", "Hoisting precast elements with pulleys/ropes up to floor V level and holding during setting", "Hoisting crew; 1.50 day = 12.00 man-hrs per cum", "Labour", "12.00 man-hrs (0.08 cum/man-hr)", "1 cum"),
            ("Bhisti", "Water-tank immersion curing in precast yard and post-installation joint watering", "Watering crew; 0.80 day = 6.40 man-hrs per cum", "Labour", "6.40 man-hrs (0.16 cum/man-hr)", "1 cum"),
            ("REF#4.1.2 (Base PCC 1:1.5:3)", "Referenced base precast concrete: 1 cum finished precast volume", "Base concrete scope executed per CPWD REF#4.1.2 specification", "Reference", "1.00 cum (1.000 cum/cum)", "1 cum"),
            ("REF#3.8 (Cement mortar 1:3)", "Referenced setting mortar for bedding and vertical jointing (0.10 cum)", "Mortar scope executed per CPWD REF#3.8 specification", "Reference", "0.10 cum (0.100 cum/cum)", "1 cum"),
        ]
    },
    {
        "id": "4.6.1",
        "parent_title": "4.6 PRECAST CONCRETE KERBS & EDGINGS AT OR NEAR GROUND LEVEL",
        "title": "4.6.1 Precast concrete kerbs, edgings 1:1.5:3 set with cement mortar 1:3",
        "unit": "cum", "base_qty": 1.0,
        "rows": [
            ("Stone Aggregate 20 mm & 10 mm", "Coarse aggregate decomposed from base PCC mix (REF#4.1.2) (0.85 cum)", "Graded 20mm + 10mm crushed stone aggregate", "Material", "0.85 cum (0.850 cum/cum)", "1 cum"),
            ("Coarse sand", "Coarse sand for casting concrete (0.425 cum) and setting mortar 1:3 (0.08 cum)", "Clean natural river sand Zone III", "Material", "0.505 cum (0.505 cum/cum)", "1 cum"),
            ("Portland Cement", "Cement for precast casting (0.40 tonne) and 1:3 bedding mortar (0.04 tonne)", "OPC-43 Grade cement conforming to IS:269", "Material", "0.44 tonne (0.440 t/cum)", "1 cum"),
            ("Carriage of Materials", "Haulage of cement, aggregates and shifting precast kerb stones along road alignment", "Carriage allowance", "Material", "1.80 cum/t equiv", "1 cum"),
            ("Concrete Mixer & Vibrating Table", "Machine mixing and compacting precast kerb moulds", "Machine hire decomposed from REF#4.1.2", "Machine", "0.56 machine-hrs (1.79 cum/hr)", "1 cum"),
            ("Precast steel gang moulds", "Precision steel moulds for chamfered road kerbs and edgings", "Mould hire and maintenance allowance", "Equipment", "Mould allowance", "1 cum"),
            ("Mason (average)", "Setting precast kerbs to string alignment, road curves, gradients, and pointing joints", "Skilled mason; 0.45 day = 3.60 man-hrs per cum", "Labour", "3.60 man-hrs (0.28 cum/man-hr)", "1 cum"),
            ("Beldar", "Casting, pond curing, shifting, trench bedding, and mixing joint mortar", "Labour crew; 2.50 day = 20.00 man-hrs per cum", "Labour", "20.00 man-hrs (0.05 cum/man-hr)", "1 cum"),
            ("Coolie", "Manual shifting and positioning heavy kerb blocks along roadside trench", "Handling labour; 1.20 day = 9.60 man-hrs per cum", "Labour", "9.60 man-hrs (0.10 cum/man-hr)", "1 cum"),
            ("Bhisti", "Tank curing in yard and curing joint mortar along kerb lines", "Watering crew; 0.75 day = 6.00 man-hrs per cum", "Labour", "6.00 man-hrs (0.17 cum/man-hr)", "1 cum"),
            ("REF#4.1.2 (Base PCC 1:1.5:3)", "Referenced base precast concrete volume: 1 cum finished kerb units", "Base concrete scope executed per CPWD REF#4.1.2 specification", "Reference", "1.00 cum (1.000 cum/cum)", "1 cum"),
            ("REF#3.8 (Cement mortar 1:3)", "Referenced setting mortar for kerb foundation and jointing (0.08 cum)", "Jointing mortar scope executed per CPWD REF#3.8 specification", "Reference", "0.08 cum (0.080 cum/cum)", "1 cum"),
        ]
    },
    {
        "id": "4.7.1",
        "parent_title": "4.7 PRECAST CONCRETE SOLID BLOCKS UP TO FLOOR V LEVEL",
        "title": "4.7.1 Precast solid blocks 1:1.5:3 set with cement mortar 1:3 and 6mm plaster face",
        "unit": "cum", "base_qty": 1.0,
        "rows": [
            ("Stone Aggregate 20 mm & 10 mm", "Coarse aggregate decomposed from base PCC mix (REF#4.1.2) (0.85 cum)", "Graded crushed stone aggregate", "Material", "0.85 cum (0.850 cum/cum)", "1 cum"),
            ("Coarse sand", "Coarse sand for concrete casting (0.425 cum) and 1:3 jointing mortar (0.10 cum)", "Clean natural river sand Zone III", "Material", "0.525 cum (0.525 cum/cum)", "1 cum"),
            ("Portland Cement", "Cement for block casting (0.40 t), 1:3 jointing mortar (0.05 t), and 6mm face plaster (0.02 t)", "OPC-43 Grade cement conforming to IS:269", "Material", "0.47 tonne (0.470 t/cum)", "1 cum"),
            ("Carriage of Materials", "Haulage of materials and hoisting solid blocks up to floor V level", "Carriage allowance", "Material", "1.85 cum/t equiv", "1 cum"),
            ("Concrete Mixer & Block Machine", "Mixing and consolidating solid blocks in gang moulds", "Machine hire allowance", "Machine", "0.56 machine-hrs (1.79 cum/hr)", "1 cum"),
            ("Mason (average)", "Laying solid blocks in plumb courses, filling vertical joints, and finishing exposed plaster", "Skilled mason; 0.55 day = 4.40 man-hrs per cum", "Labour", "4.40 man-hrs (0.23 cum/man-hr)", "1 cum"),
            ("Beldar", "Casting, handling, hoisting, and mixing mortar", "Labour crew; 2.80 day = 22.40 man-hrs per cum", "Labour", "22.40 man-hrs (0.04 cum/man-hr)", "1 cum"),
            ("Coolie", "Carrying blocks, assisting masons, and shifting scaffolding", "Handling crew; 1.50 day = 12.00 man-hrs per cum", "Labour", "12.00 man-hrs (0.08 cum/man-hr)", "1 cum"),
            ("Bhisti", "Watering block masonry courses and curing plaster", "Watering crew; 0.80 day = 6.40 man-hrs per cum", "Labour", "6.40 man-hrs (0.16 cum/man-hr)", "1 cum"),
            ("REF#4.1.2 (Base PCC 1:1.5:3)", "Referenced base concrete volume: 1 cum finished solid blocks", "Base concrete scope executed per CPWD REF#4.1.2 specification", "Reference", "1.00 cum (1.000 cum/cum)", "1 cum"),
            ("REF#3.8 (Cement mortar 1:3)", "Referenced setting mortar for block masonry (0.10 cum)", "Mortar scope executed per CPWD REF#3.8 specification", "Reference", "0.10 cum (0.100 cum/cum)", "1 cum"),
        ]
    },
    {
        "id": "4.8.1",
        "parent_title": "4.8 PRECAST CONCRETE HOLLOW BLOCKS UP TO FLOOR V LEVEL",
        "title": "4.8.1 Precast hollow blocks 1:1.5:3 set with cement mortar 1:3 and 6mm plaster face",
        "unit": "cum", "base_qty": 1.0,
        "rows": [
            ("Stone Aggregate 20 mm & 10 mm", "Coarse aggregate for hollow block web and shell walls (0.55 cum)", "Graded crushed stone aggregate (net solid volume ~0.65 cum/cum gross)", "Material", "0.55 cum (0.550 cum/cum)", "1 cum"),
            ("Coarse sand", "Coarse sand for hollow block casting (0.28 cum) and jointing mortar (0.08 cum)", "Clean natural river sand Zone III", "Material", "0.36 cum (0.360 cum/cum)", "1 cum"),
            ("Portland Cement", "Cement for hollow blocks (0.26 t), setting mortar (0.04 t), and 6mm plaster face (0.02 t)", "OPC-43 Grade cement conforming to IS:269", "Material", "0.32 tonne (0.320 t/cum)", "1 cum"),
            ("Carriage of Materials", "Haulage of materials and hoisting lightweight hollow blocks up to floor V", "Carriage allowance", "Material", "1.30 cum/t equiv", "1 cum"),
            ("Concrete Mixer & Hollow Block Machine", "Mixing and vibrating hollow blocks with core extractors", "Machine hire allowance", "Machine", "0.56 machine-hrs (1.79 cum/hr)", "1 cum"),
            ("Mason (average)", "Laying hollow blocks with mortar on face shells, checking cavities, and plastering faces", "Skilled mason; 0.50 day = 4.00 man-hrs per cum", "Labour", "4.00 man-hrs (0.25 cum/man-hr)", "1 cum"),
            ("Beldar", "Casting, handling, hoisting, and mixing mortar", "Labour crew; 2.40 day = 19.20 man-hrs per cum", "Labour", "19.20 man-hrs (0.05 cum/man-hr)", "1 cum"),
            ("Coolie", "Carrying blocks to masons and scaffolding platforms", "Handling crew; 1.20 day = 9.60 man-hrs per cum", "Labour", "9.60 man-hrs (0.10 cum/man-hr)", "1 cum"),
            ("Bhisti", "Curing hollow block walls and wetting plaster", "Watering crew; 0.75 day = 6.00 man-hrs per cum", "Labour", "6.00 man-hrs (0.17 cum/man-hr)", "1 cum"),
            ("REF#4.1.2 (Base PCC 1:1.5:3)", "Referenced base concrete volume for solid web/shell volume (~0.65 cum)", "Base concrete scope executed per CPWD REF#4.1.2 specification", "Reference", "0.65 cum (0.650 cum/cum)", "1 cum"),
            ("REF#3.8 (Cement mortar 1:3)", "Referenced setting mortar for hollow block bedding (0.08 cum)", "Mortar scope executed per CPWD REF#3.8 specification", "Reference", "0.08 cum (0.080 cum/cum)", "1 cum"),
        ]
    },
    {
        "id": "4.9",
        "parent_title": "4.9 PRECAST CONCRETE BOLLARDS",
        "title": "4.9 Precasting and placing in position 125 mm dia Bollards 600 mm high in PCC 1:3:6",
        "unit": "each", "base_qty": 1.0,
        "rows": [
            ("PCC 1:3:6 Concrete (REF#4.1.5)", "Finished concrete volume per cylindrical bollard: pi/4 x 0.125^2 x 0.60m = 0.0074 cum", "Concrete 1:3:6 decomposed into aggregates, coarse sand, and cement", "Material", "0.0074 cum (0.0074 cum/each)", "1 each"),
            ("Cylindrical bollard moulds", "Fabricated steel cylindrical moulds (125mm dia x 600mm height) with rounded dome cap", "Mould hire and oiling allowance", "Equipment", "Mould allowance", "1 each"),
            ("Concrete Mixer & Needle Vibrator", "Mixing small batch concrete and vibrating slender bollard column", "Machine hire decomposed from REF#4.1.5", "Machine", "0.01 machine-hrs", "1 each"),
            ("Mason (average)", "Placing in pit, plumbing vertically, embedding base, and neat cement float dome finish", "Skilled mason; 0.05 day = 0.40 man-hrs per bollard", "Labour", "0.40 man-hrs (2.50 each/man-hr)", "1 each"),
            ("Beldar / Coolie", "Digging foundation socket pit, casting bollard, carrying to location, and backfill ramming", "Labour crew; 0.12 day = 0.96 man-hrs per bollard", "Labour", "0.96 man-hrs (1.04 each/man-hr)", "1 each"),
            ("Bhisti", "Watering foundation concrete and curing bollard", "Watering allowance; 0.03 day = 0.24 man-hrs per bollard", "Labour", "0.24 man-hrs (4.17 each/man-hr)", "1 each"),
            ("REF#4.1.5 (Base PCC 1:3:6)", "Referenced base concrete formulation: Concrete 1:3:6 (20mm)", "Base concrete scope executed per CPWD REF#4.1.5 specification", "Reference", "0.0074 cum", "1 each"),
        ]
    },

    # ─── 4.10 & 4.11 DAMP-PROOF COURSE (DPC) ─────────────────────────────────────
    {
        "id": "4.10",
        "parent_title": "4.10 DAMP-PROOF COURSE (DPC) 40 MM THICK",
        "title": "4.10 Damp-proof course 40mm thick with cement concrete 1:2:4 (1 cement : 2 coarse sand : 4 stone aggregate 12.5mm)",
        "unit": "sqm", "base_qty": 10.0,
        "rows": [
            ("Stone Aggregate 12.5 mm & 10 mm", "Coarse aggregate decomposed from base PCC mix (REF#4.1.3) (0.34 cum for 0.40 cum net concrete)", "Clean graded stone aggregate 12.5mm nominal size for thin 40mm layer", "Material", "0.34 cum (0.034 cum/sqm)", "10 sqm"),
            ("Coarse sand", "Zone III coarse sand matrix decomposed from REF#4.1.3 (0.17 cum)", "Clean natural river sand Zone III", "Material", "0.17 cum (0.017 cum/sqm)", "10 sqm"),
            ("Portland Cement", "Hydraulic binder decomposed from REF#4.1.3 (0.128 tonne = 2.56 bags)", "OPC-43 Grade cement conforming to IS:269", "Material", "0.128 tonne (0.0128 t/sqm)", "10 sqm"),
            ("Carriage of Aggregates & Cement", "Haulage and handling of materials along plinth wall trenches", "Carriage allowance decomposed from REF#4.1.3", "Material", "0.55 cum/t equiv", "10 sqm"),
            ("Concrete Mixer 0.28 to 0.40 cum", "Machine mixing 1:2:4 dense concrete batch for DPC layer", "Mixer operational hire; output = 1.79 cum/hr (0.40 cum concrete = 0.22 hr)", "Machine", "0.22 machine-hrs (45.45 sqm/hr)", "10 sqm"),
            ("Timber side stop battens", "Wooden side shuttering battens (40mm deep) pegged along plinth wall edges", "Batten hire and nail allowance", "Equipment", "Batten allowance", "10 sqm"),
            ("Mason (average)", "Setting side battens, spreading concrete, compacting with wooden rammers, and neat cement floating", "Skilled mason; 0.22 day = 1.76 man-hrs per 10 sqm", "Labour", "1.76 man-hrs (5.68 sqm/man-hr)", "10 sqm"),
            ("Beldar", "Measuring ingredients, feeding mixer drum, wheeling concrete along plinth, and ramming", "Labour crew; 0.70 day = 5.60 man-hrs per 10 sqm", "Labour", "5.60 man-hrs (1.79 sqm/man-hr)", "10 sqm"),
            ("Coolie", "Carrying concrete pans and moving materials along foundation walls", "Handling crew; 0.35 day = 2.80 man-hrs per 10 sqm", "Labour", "2.80 man-hrs (3.57 sqm/man-hr)", "10 sqm"),
            ("Bhisti", "Water batching, wetting plinth brickwork before DPC laying, and moist curing for 7 days", "Watering crew; 0.35 day = 2.80 man-hrs per 10 sqm", "Labour", "2.80 man-hrs (3.57 sqm/man-hr)", "10 sqm"),
            ("REF#4.1.3 (Base PCC 1:2:4)", "Referenced base concrete volume: 10 sqm x 0.04m = 0.40 cum concrete", "Base concrete scope executed per CPWD REF#4.1.3 specification", "Reference", "0.40 cum (0.040 cum/sqm)", "10 sqm"),
        ]
    },
    {
        "id": "4.11",
        "parent_title": "4.11 DAMP-PROOF COURSE (DPC) 50 MM THICK",
        "title": "4.11 Damp-proof course 50mm thick with cement concrete 1:2:4 (1 cement : 2 coarse sand : 4 stone aggregate 20mm)",
        "unit": "sqm", "base_qty": 10.0,
        "rows": [
            ("Stone Aggregate 20 mm & 10 mm", "Coarse aggregate decomposed from base PCC mix (REF#4.1.3) (0.425 cum for 0.50 cum net concrete)", "Clean graded stone aggregate 20mm nominal size for 50mm layer", "Material", "0.425 cum (0.0425 cum/sqm)", "10 sqm"),
            ("Coarse sand", "Zone III coarse sand matrix decomposed from REF#4.1.3 (0.213 cum)", "Clean natural river sand Zone III", "Material", "0.213 cum (0.0213 cum/sqm)", "10 sqm"),
            ("Portland Cement", "Hydraulic binder decomposed from REF#4.1.3 (0.160 tonne = 3.20 bags)", "OPC-43 Grade cement conforming to IS:269", "Material", "0.160 tonne (0.0160 t/sqm)", "10 sqm"),
            ("Carriage of Aggregates & Cement", "Haulage and handling of materials along plinth wall trenches", "Carriage allowance decomposed from REF#4.1.3", "Material", "0.70 cum/t equiv", "10 sqm"),
            ("Concrete Mixer 0.28 to 0.40 cum", "Machine mixing 1:2:4 dense concrete batch for 50mm DPC", "Mixer operational hire; output = 1.79 cum/hr (0.50 cum concrete = 0.28 hr)", "Machine", "0.28 machine-hrs (35.71 sqm/hr)", "10 sqm"),
            ("Timber side stop battens", "Wooden side shuttering battens (50mm deep) pegged along plinth wall edges", "Batten hire and nail allowance", "Equipment", "Batten allowance", "10 sqm"),
            ("Mason (average)", "Setting side battens, spreading concrete, compacting with rammers, and trowel finishing", "Skilled mason; 0.26 day = 2.08 man-hrs per 10 sqm", "Labour", "2.08 man-hrs (4.81 sqm/man-hr)", "10 sqm"),
            ("Beldar", "Measuring ingredients, feeding mixer drum, wheeling concrete along plinth, and ramming", "Labour crew; 0.85 day = 6.80 man-hrs per 10 sqm", "Labour", "6.80 man-hrs (1.47 sqm/man-hr)", "10 sqm"),
            ("Coolie", "Carrying concrete pans and moving materials along foundation walls", "Handling crew; 0.40 day = 3.20 man-hrs per 10 sqm", "Labour", "3.20 man-hrs (3.12 sqm/man-hr)", "10 sqm"),
            ("Bhisti", "Water batching, wetting plinth brickwork before DPC laying, and moist curing for 7 days", "Watering crew; 0.40 day = 3.20 man-hrs per 10 sqm", "Labour", "3.20 man-hrs (3.12 sqm/man-hr)", "10 sqm"),
            ("REF#4.1.3 (Base PCC 1:2:4)", "Referenced base concrete volume: 10 sqm x 0.05m = 0.50 cum concrete", "Base concrete scope executed per CPWD REF#4.1.3 specification", "Reference", "0.50 cum (0.050 cum/sqm)", "10 sqm"),
        ]
    },

    # ─── 4.12 TO 4.18 ADD-ONS, BITUMEN, LIFTS & FIBRES ───────────────────────────
    {
        "id": "4.12",
        "parent_title": "4.12 WATERPROOFING COMPOUND IN CONCRETE",
        "title": "4.12 Extra for providing and mixing water proofing material in cement concrete work (per 50 kg cement bag)",
        "unit": "per bag", "base_qty": 1.0,
        "rows": [
            ("Integral cement water proofing compound", "Hydrophobic liquid or powder admixture blocking capillary pore channels in concrete", "Conforming to IS:2645; dosed at 1.00 kg per 50 kg cement bag (2% by weight of cement)", "Material", "1.00 kg (1.000 kg/bag)", "1 bag"),
            ("Beldar", "Measuring admixture in calibrated flask, pre-dispersing into mixing water before adding to mixer drum", "Admixture dispensing and batching allowance", "Labour", "0.05 man-hrs (20.0 bags/man-hr)", "1 bag"),
        ]
    },
    {
        "id": "4.13",
        "parent_title": "4.13 BITUMEN COATING ON DPC",
        "title": "4.13 Providing and applying a coat of residual petroleum bitumen VG-10 on DPC surface @ 1.7 kg/sqm",
        "unit": "sqm", "base_qty": 10.0,
        "rows": [
            ("Paving bitumen VG-10", "Penetration/viscosity grade residual bitumen forming continuous waterproof membrane (17 kg)", "Industrial paving bitumen grade VG-10 conforming to IS:73", "Material", "17.00 kg (1.700 kg/sqm)", "10 sqm"),
            ("Kerosene oil", "Primer solvent/cutter used for surface cleaning and brush workability (1.23 litre)", "Clean technical kerosene oil", "Material", "1.23 litre (0.123 L/sqm)", "10 sqm"),
            ("Coal (steam) / fuel", "Heating fuel for bitumen boiling boiler/tar kettle (3.5 kg coal)", "Steam coal for bitumen heating kettle", "Material", "3.50 kg (0.350 kg/sqm)", "10 sqm"),
            ("Carriage of Tar bitumen", "Transport and handling of bitumen drums to heating station", "Carriage allowance", "Material", "0.017 tonne", "10 sqm"),
            ("Bitumen boiler / heating kettle", "Heating kettle, squeegees, mopping brushes, and thermometer", "Kettle equipment hire allowance", "Equipment", "Equipment allowance", "10 sqm"),
            ("Beldar", "Heating bitumen in boiler, monitoring boiling temperature (165-175°C), and carrying hot tar buckets", "Hot bitumen heating and application crew; 0.12 day = 0.96 man-hrs per 10 sqm", "Labour", "0.96 man-hrs (10.42 sqm/man-hr)", "10 sqm"),
            ("Coolie", "Cleaning DPC concrete surface of dust with wire brushes and assisting application", "Helper crew; 0.07 day = 0.56 man-hrs per 10 sqm", "Labour", "0.56 man-hrs (17.86 sqm/man-hr)", "10 sqm"),
            ("Sundries", "Coir brushes, gloves, safety boots, and bucket allowances", "Sundries allowance", "Equipment", "L.S. allowance", "10 sqm"),
        ]
    },
    {
        "id": "4.14",
        "parent_title": "4.14 STAGE LIFT EXTRA ABOVE FLOOR V LEVEL",
        "title": "4.14 Extra for concrete work in superstructure above floor V level for each four floors or part thereof",
        "unit": "cum", "base_qty": 1.0,
        "rows": [
            ("Coolie", "Vertical manual and hoist handling of concrete buckets/pans for additional height above floor V", "Stage lift labour allowance; 1.50 day = 12.00 man-hrs per cum", "Labour", "12.00 man-hrs (0.08 cum/man-hr)", "1 cum"),
            ("Builder's hoist / crane", "Material hoist / passenger-material hoist running charges for vertical transit of concrete", "Vertical lift machinery allowance", "Machine", "Vertical hoist allowance", "1 cum"),
        ]
    },
    {
        "id": "4.15",
        "parent_title": "4.15 CONCRETE IN WATER / LIQUID MUD CONDITIONS",
        "title": "4.15 Extra for laying concrete in or under water and/or liquid mud including pumping/bailing out water",
        "unit": "cum", "base_qty": 1.0,
        "rows": [
            ("Pumpset capacity 4000 to 5000 litres/hr", "Continuous dewatering pumpset running hire to keep excavation pit dry during tremie/chute concrete pour", "Pumpset operational hire (0.375 day = 3.0 machine-hrs per cum)", "Machine", "3.00 machine-hrs (0.33 cum/hr)", "1 cum"),
            ("Beldar", "Continuous bailing out water, guiding tremie pipes/chutes, building cofferdam mud dykes, and rapid placing", "Dewatering and mud placing crew; 4.00 day = 32.00 man-hrs per cum", "Labour", "32.00 man-hrs (0.03 cum/man-hr)", "1 cum"),
            ("Tremie pipes, suction hoses & canvas chutes", "Watertight tremie pipe sections, hopper funnel, and discharge hoses", "Specialised sub-aqueous placing equipment", "Equipment", "Equipment allowance", "1 cum"),
        ]
    },
    {
        "id": "4.16",
        "parent_title": "4.16 CONCRETE IN FOUL POSITIONS",
        "title": "4.16 Extra for laying concrete in or under foul positions (sewers, manholes, cesspools)",
        "unit": "cum", "base_qty": 1.0,
        "rows": [
            ("Mason (average)", "Working inside foul septic/sewer environments under gas masks, setting levels, and floating surfaces", "Specialised foul environment mason; 0.04 day = 0.32 man-hrs per cum", "Labour", "0.32 man-hrs (3.12 cum/man-hr)", "1 cum"),
            ("Beldar", "Descending into manhole shafts, shoveling, placing and compacting concrete under hazardous foul conditions", "Foul position labour crew; 0.25 day = 2.00 man-hrs per cum", "Labour", "2.00 man-hrs (0.50 cum/man-hr)", "1 cum"),
            ("Coolie", "Surface safety watch, hoisting spoil, lowering concrete buckets, and emergency lifeline tending", "Safety and haulage helper; 0.15 day = 1.20 man-hrs per cum", "Labour", "1.20 man-hrs (0.83 cum/man-hr)", "1 cum"),
            ("Safety gas masks, blower & harnesses", "Gas detectors, oxygen escape breathing sets, ventilation blower, and safety harness ropes", "Safety gear allowance", "Equipment", "Safety allowance", "1 cum"),
        ]
    },
    {
        "id": "4.17",
        "parent_title": "4.17 PLINTH PROTECTION 50 MM THICK OVER 75 MM BRICK BALLAST",
        "title": "4.17 Making plinth protection 50mm thick of cement concrete 1:3:6 over 75mm dry brick ballast bed (10 sqm)",
        "unit": "sqm", "base_qty": 10.0,
        "rows": [
            ("Brick Aggregate (Single size) 40 mm", "Dry brick ballast bed underlayer 75 mm nominal thickness (0.75 cum)", "Clean broken brick aggregate 40 mm size consolidating subgrade under apron", "Material", "0.75 cum (0.075 cum/sqm)", "10 sqm"),
            ("Carriage of Brick aggregate", "Haulage and dumping of brick aggregate along building perimeter (0.75 cum)", "Carriage of brick aggregate", "Material", "0.75 cum (0.075 cum/sqm)", "10 sqm"),
            ("Fine sand", "Fine sand spread over dry brick ballast to fill voids before concrete pour (0.06 cum)", "Clean fine sand filling ballast inter-spaces", "Material", "0.06 cum (0.006 cum/sqm)", "10 sqm"),
            ("Carriage of Fine sand", "Carriage of fine sand (0.06 cum)", "Carriage of fine sand", "Material", "0.06 cum (0.006 cum/sqm)", "10 sqm"),
            ("Stone Aggregate 20 mm & 10 mm", "Coarse aggregate for 50mm PCC 1:3:6 top apron slab (0.45 cum decomposed from REF#4.1.5)", "Graded 20mm + 10mm crushed stone aggregate", "Material", "0.45 cum (0.045 cum/sqm)", "10 sqm"),
            ("Coarse sand", "Zone III coarse sand matrix for 50mm PCC top slab (0.225 cum decomposed from REF#4.1.5)", "Clean coarse sand Zone III", "Material", "0.225 cum (0.0225 cum/sqm)", "10 sqm"),
            ("Portland Cement", "Cement for 50mm PCC 1:3:6 apron slab (0.11 t) plus neat cement slurry floating & 75mm wall riser", "OPC-43 Grade cement conforming to IS:269", "Material", "0.14 tonne (0.014 t/sqm)", "10 sqm"),
            ("Carriage of Stone aggregate & Cement", "Haulage of materials along building perimeter", "Carriage allowance", "Material", "0.80 cum/t equiv", "10 sqm"),
            ("Concrete Mixer 0.28 to 0.40 cum", "Machine mixing 1:3:6 concrete batch (0.50 cum net volume = 0.28 hr)", "Mixer operational hire; output = 1.79 cum/hr", "Machine", "0.28 machine-hrs (35.71 sqm/hr)", "10 sqm"),
            ("Mason (average)", "Setting outer slope battens, screeding apron to 1:48 outward fall, and trowel neat cement finish", "Skilled mason; 0.45 day = 3.60 man-hrs per 10 sqm", "Labour", "3.60 man-hrs (2.78 sqm/man-hr)", "10 sqm"),
            ("Beldar", "Spreading and hand-ramming 75mm brick ballast, blinding with sand, placing concrete, and ramming", "Labour crew; 1.40 day = 11.20 man-hrs per 10 sqm", "Labour", "11.20 man-hrs (0.89 sqm/man-hr)", "10 sqm"),
            ("Coolie", "Carrying brick ballast, sand, concrete pans, and shifting slope forms along plinth", "Handling crew; 0.80 day = 6.40 man-hrs per 10 sqm", "Labour", "6.40 man-hrs (1.56 sqm/man-hr)", "10 sqm"),
            ("Bhisti", "Watering brick ballast bed before concrete pour and ponding/curing apron for 14 days", "Watering crew; 0.45 day = 3.60 man-hrs per 10 sqm", "Labour", "3.60 man-hrs (2.78 sqm/man-hr)", "10 sqm"),
            ("REF#4.1.5 (Base PCC 1:3:6)", "Referenced base concrete formulation: 10 sqm x 0.05m = 0.50 cum concrete", "Base concrete scope executed per CPWD REF#4.1.5 specification", "Reference", "0.50 cum (0.050 cum/sqm)", "10 sqm"),
        ]
    },
    {
        "id": "4.18",
        "parent_title": "4.18 SYNTHETIC POLYESTER FIBRE IN CONCRETE",
        "title": "4.18 Extra for addition of synthetic Polyester triangular fibre of length 12mm @ 125 gm per 50 kg cement bag",
        "unit": "per bag", "base_qty": 1.0,
        "rows": [
            ("Synthetic polyester triangular fibre", "Virgin polyester triangular micro-fibres dispersing throughout concrete matrix to control plastic shrinkage cracks", "Conforming to MORTH / CPWD specification; 12mm length dosed @ 0.125 kg per 50 kg cement bag", "Material", "0.125 kg (0.125 kg/bag)", "1 bag"),
            ("Beldar", "Opening fibre packets, teasing clumps, and distributing evenly into mixer drum during dry mix cycle", "Fibre dispensing and mixing labour allowance", "Labour", "0.04 man-hrs (25.0 bags/man-hr)", "1 bag"),
        ]
    },

    # ─── 4.19 READY MIXED CONCRETE (RMC) WITH FLY ASH ───────────────────────────
    {
        "id": "4.19.1.1",
        "parent_title": "4.19 READY MIXED CONCRETE (RMC) WITH FLY ASH — UP TO PLINTH LEVEL",
        "title": "4.19.1.1 RMC M-15 grade plain cement concrete with fly ash up to plinth level (cement considered @ 240 kg/cum)",
        "unit": "cum", "base_qty": 1.0,
        "rows": [
            ("Stone Aggregate 40 mm", "Coarse aggregate 40 mm batched at computerised RMC batching plant (0.65 cum)", "Clean graded crushed stone aggregate 40 mm nominal size", "Material", "0.65 cum (0.650 cum/cum)", "1 cum"),
            ("Stone Aggregate 20 mm", "Coarse aggregate 20 mm batched at RMC plant (0.24 cum)", "Clean graded crushed stone aggregate 20 mm nominal size", "Material", "0.24 cum (0.240 cum/cum)", "1 cum"),
            ("Carriage of Stone aggregate", "Haulage of 40mm and 20mm aggregates to RMC plant (0.89 cum)", "Carriage allowance", "Material", "0.89 cum (0.890 cum/cum)", "1 cum"),
            ("Coarse sand", "Zone III coarse sand batched at RMC plant (0.37 cum)", "Clean natural river sand Zone III", "Material", "0.37 cum (0.370 cum/cum)", "1 cum"),
            ("Carriage of Coarse sand", "Haulage of coarse sand to RMC plant (0.37 cum)", "Carriage allowance", "Material", "0.37 cum (0.370 cum/cum)", "1 cum"),
            ("Fly ash", "Industrial pozzolanic fly ash blended in RMC mix (0.06 tonne = 60 kg)", "Conforming to IS:3812 Part 1", "Material", "0.06 tonne (0.060 t/cum)", "1 cum"),
            ("Carriage of Fly ash", "Bulk tanker haulage of fly ash to RMC silo (0.06 tonne)", "Carriage allowance", "Material", "0.06 tonne (0.060 t/cum)", "1 cum"),
            ("Portland Cement", "OPC-43 Grade cement batched in RMC silo (0.24 tonne = 240 kg)", "Conforming to IS:269 / IS:8112", "Material", "0.24 tonne (0.240 t/cum)", "1 cum"),
            ("Carriage of Cement", "Bulk tanker transport of cement to RMC plant silo (0.24 tonne)", "Carriage allowance", "Material", "0.24 tonne (0.240 t/cum)", "1 cum"),
            ("Chemical Plasticizer / Retarder", "High-range water-reducing and retarding admixture maintaining 2-hour slump retention during transit", "Conforming to IS:9103 dosed @ 0.5% by weight of binder (1.5 kg)", "Material", "1.50 kg (1.500 kg/cum)", "1 cum"),
            ("Transit Mixer 6.0 cum capacity", "Haulage of wet batched concrete from RMC plant to site under continuous agitation", "Transit mixer hire; 6 cum payload; transit distance up to 15 km; output = 3.0 cum/hr", "Machine", "0.33 machine-hrs (3.00 cum/hr)", "1 cum"),
            ("Concrete Pump / Placer 30 cum/hr", "Pumping concrete through pipeline and flexible hose boom directly to foundation point", "Concrete pump hire allowance; output = 15.0 cum/hr", "Machine", "0.067 machine-hrs (15.0 cum/hr)", "1 cum"),
            ("Needle / Surface Vibrator", "Mechanical consolidation of pumped concrete", "Vibrator hire; output = 2.0 cum/hr", "Machine", "0.50 machine-hrs (2.00 cum/hr)", "1 cum"),
            ("Mason (average)", "Setting grade stakes, leveling pumped concrete, screeding, and float finishing", "Mason labour; 0.08 day = 0.64 man-hrs per cum", "Labour", "0.64 man-hrs (1.56 cum/man-hr)", "1 cum"),
            ("Beldar", "Handling pump delivery hose, spreading pumped concrete, and operating needle vibrator", "Labour crew; 0.60 day = 4.80 man-hrs per cum", "Labour", "4.80 man-hrs (0.21 cum/man-hr)", "1 cum"),
            ("Coolie", "Guiding transit mixer, coupling pipeline clamps, washing pump hopper and pipeline", "Helper crew; 0.40 day = 3.20 man-hrs per cum", "Labour", "3.20 man-hrs (0.31 cum/man-hr)", "1 cum"),
            ("Bhisti", "Supplying wash water and initial curing", "Watering crew; 0.30 day = 2.40 man-hrs per cum", "Labour", "2.40 man-hrs (0.42 cum/man-hr)", "1 cum"),
            ("Sundries", "Pipeline washouts, sponge balls, slurry primers, and testing cubes", "Sundries allowance", "Equipment", "L.S. allowance", "1 cum"),
        ]
    },
    {
        "id": "4.19.1.2",
        "parent_title": "4.19 READY MIXED CONCRETE (RMC) WITH FLY ASH — UP TO PLINTH LEVEL",
        "title": "4.19.1.2 RMC M-10 grade plain cement concrete with fly ash up to plinth level (cement considered @ 220 kg/cum)",
        "unit": "cum", "base_qty": 1.0,
        "rows": [
            ("Stone Aggregate 40 mm", "Coarse aggregate 40 mm batched at RMC plant (0.65 cum)", "Crushed stone aggregate 40 mm", "Material", "0.65 cum (0.650 cum/cum)", "1 cum"),
            ("Stone Aggregate 20 mm", "Coarse aggregate 20 mm batched at RMC plant (0.24 cum)", "Crushed stone aggregate 20 mm", "Material", "0.24 cum (0.240 cum/cum)", "1 cum"),
            ("Carriage of Stone aggregate", "Haulage of aggregates to RMC plant (0.89 cum)", "Carriage allowance", "Material", "0.89 cum (0.890 cum/cum)", "1 cum"),
            ("Coarse sand", "Zone III coarse sand batched at RMC plant (0.37 cum)", "Clean natural river sand Zone III", "Material", "0.37 cum (0.370 cum/cum)", "1 cum"),
            ("Carriage of Coarse sand", "Haulage of coarse sand to RMC plant (0.37 cum)", "Carriage allowance", "Material", "0.37 cum (0.370 cum/cum)", "1 cum"),
            ("Fly ash", "Industrial pozzolanic fly ash blend (0.06 tonne = 60 kg)", "Conforming to IS:3812 Part 1", "Material", "0.06 tonne (0.060 t/cum)", "1 cum"),
            ("Carriage of Fly ash", "Haulage of fly ash to RMC silo (0.06 tonne)", "Carriage allowance", "Material", "0.06 tonne (0.060 t/cum)", "1 cum"),
            ("Portland Cement", "OPC-43 Grade cement batched in RMC silo (0.22 tonne = 220 kg)", "Conforming to IS:269", "Material", "0.22 tonne (0.220 t/cum)", "1 cum"),
            ("Carriage of Cement", "Transport of cement to RMC plant silo (0.22 tonne)", "Carriage allowance", "Material", "0.22 tonne (0.220 t/cum)", "1 cum"),
            ("Chemical Plasticizer / Retarder", "Water-reducing retarding admixture maintaining workability (1.4 kg)", "Conforming to IS:9103", "Material", "1.40 kg (1.400 kg/cum)", "1 cum"),
            ("Transit Mixer 6.0 cum capacity", "Haulage of wet concrete from batching plant to site under agitation", "Transit mixer hire; output = 3.0 cum/hr", "Machine", "0.33 machine-hrs (3.00 cum/hr)", "1 cum"),
            ("Concrete Pump / Placer 30 cum/hr", "Pumping concrete directly to foundation bed", "Concrete pump hire; output = 15.0 cum/hr", "Machine", "0.067 machine-hrs (15.0 cum/hr)", "1 cum"),
            ("Needle / Surface Vibrator", "Mechanical consolidation of pumped concrete", "Vibrator hire; output = 2.0 cum/hr", "Machine", "0.50 machine-hrs (2.00 cum/hr)", "1 cum"),
            ("Mason (average)", "Leveling pumped concrete, screeding, and float finishing", "Mason labour; 0.08 day = 0.64 man-hrs per cum", "Labour", "0.64 man-hrs (1.56 cum/man-hr)", "1 cum"),
            ("Beldar", "Handling delivery hose, spreading and vibrating concrete", "Labour crew; 0.60 day = 4.80 man-hrs per cum", "Labour", "4.80 man-hrs (0.21 cum/man-hr)", "1 cum"),
            ("Coolie", "Guiding transit mixer, coupling pipes, and washing equipment", "Helper crew; 0.40 day = 3.20 man-hrs per cum", "Labour", "3.20 man-hrs (0.31 cum/man-hr)", "1 cum"),
            ("Bhisti", "Wash water supply and initial curing", "Watering crew; 0.30 day = 2.40 man-hrs per cum", "Labour", "2.40 man-hrs (0.42 cum/man-hr)", "1 cum"),
            ("Sundries", "Washout sponges, pipeline clamps, and cube testing", "Sundries allowance", "Equipment", "L.S. allowance", "1 cum"),
        ]
    },
    {
        "id": "4.19.2.1",
        "parent_title": "4.19 READY MIXED CONCRETE (RMC) WITH FLY ASH — ABOVE PLINTH TO FLOOR V",
        "title": "4.19.2.1 RMC M-15 grade plain cement concrete with fly ash above plinth up to floor V level",
        "unit": "cum", "base_qty": 1.0,
        "rows": [
            ("Stone Aggregate 40 mm & 20 mm", "Graded stone aggregate 40mm (0.65 cum) and 20mm (0.24 cum) batched at RMC plant", "Clean crushed stone aggregate", "Material", "0.89 cum (0.890 cum/cum)", "1 cum"),
            ("Coarse sand", "Zone III coarse sand batched at RMC plant (0.37 cum)", "Clean natural river sand Zone III", "Material", "0.37 cum (0.370 cum/cum)", "1 cum"),
            ("Fly ash", "Pozzolanic fly ash blend (0.06 tonne = 60 kg)", "Conforming to IS:3812 Part 1", "Material", "0.06 tonne (0.060 t/cum)", "1 cum"),
            ("Portland Cement", "OPC-43 Grade cement batched in RMC silo (0.24 tonne = 240 kg)", "Conforming to IS:269", "Material", "0.24 tonne (0.240 t/cum)", "1 cum"),
            ("Carriage of Materials", "Bulk haulage of cement, fly ash, sand and aggregates to RMC batching plant", "Carriage allowance", "Material", "1.56 cum/t equiv", "1 cum"),
            ("Chemical Plasticizer / Retarder", "High-range plasticizer maintaining slump retention (1.5 kg)", "Conforming to IS:9103", "Material", "1.50 kg (1.500 kg/cum)", "1 cum"),
            ("Transit Mixer 6.0 cum capacity", "Haulage from RMC plant to site under agitation", "Transit mixer hire; output = 3.0 cum/hr", "Machine", "0.33 machine-hrs (3.00 cum/hr)", "1 cum"),
            ("Concrete Boom Pump 30 cum/hr", "Boom pumping vertically through articulated boom up to floor V level (15-20 m vertical head)", "High-pressure concrete pump hire; output = 12.0 cum/hr", "Machine", "0.083 machine-hrs (12.0 cum/hr)", "1 cum"),
            ("Needle Vibrator", "Compacting pumped concrete in elevated forms", "Vibrator hire; output = 2.0 cum/hr", "Machine", "0.50 machine-hrs (2.00 cum/hr)", "1 cum"),
            ("Mason (average)", "Controlling elevated pour, leveling, and screeding to deck benchmarks", "Mason labour; 0.12 day = 0.96 man-hrs per cum", "Labour", "0.96 man-hrs (1.04 cum/man-hr)", "1 cum"),
            ("Beldar", "Handling boom hose end, spreading, and operating vibrator on elevated floor", "Labour crew; 0.80 day = 6.40 man-hrs per cum", "Labour", "6.40 man-hrs (0.16 cum/man-hr)", "1 cum"),
            ("Coolie", "Safety watch, guiding boom, coupling riser pipes, and washing hopper", "Helper crew; 0.50 day = 4.00 man-hrs per cum", "Labour", "4.00 man-hrs (0.25 cum/man-hr)", "1 cum"),
            ("Bhisti", "Water supply and initial deck curing", "Watering crew; 0.35 day = 2.80 man-hrs per cum", "Labour", "2.80 man-hrs (0.36 cum/man-hr)", "1 cum"),
            ("Sundries", "Elevated safety barriers, hose slings, washout sponges, and testing cubes", "Sundries allowance", "Equipment", "L.S. allowance", "1 cum"),
        ]
    },
    {
        "id": "4.19.2.2",
        "parent_title": "4.19 READY MIXED CONCRETE (RMC) WITH FLY ASH — ABOVE PLINTH TO FLOOR V",
        "title": "4.19.2.2 RMC M-10 grade plain cement concrete with fly ash above plinth up to floor V level",
        "unit": "cum", "base_qty": 1.0,
        "rows": [
            ("Stone Aggregate 40 mm & 20 mm", "Graded stone aggregate 40mm (0.65 cum) and 20mm (0.24 cum) batched at RMC plant", "Clean crushed stone aggregate", "Material", "0.89 cum (0.890 cum/cum)", "1 cum"),
            ("Coarse sand", "Zone III coarse sand batched at RMC plant (0.37 cum)", "Clean natural river sand Zone III", "Material", "0.37 cum (0.370 cum/cum)", "1 cum"),
            ("Fly ash", "Pozzolanic fly ash blend (0.06 tonne = 60 kg)", "Conforming to IS:3812 Part 1", "Material", "0.06 tonne (0.060 t/cum)", "1 cum"),
            ("Portland Cement", "OPC-43 Grade cement batched in RMC silo (0.22 tonne = 220 kg)", "Conforming to IS:269", "Material", "0.22 tonne (0.220 t/cum)", "1 cum"),
            ("Carriage of Materials", "Bulk haulage of materials to RMC batching plant", "Carriage allowance", "Material", "1.54 cum/t equiv", "1 cum"),
            ("Chemical Plasticizer / Retarder", "High-range plasticizer maintaining slump retention (1.4 kg)", "Conforming to IS:9103", "Material", "1.40 kg (1.400 kg/cum)", "1 cum"),
            ("Transit Mixer 6.0 cum capacity", "Haulage from RMC plant to site under agitation", "Transit mixer hire; output = 3.0 cum/hr", "Machine", "0.33 machine-hrs (3.00 cum/hr)", "1 cum"),
            ("Concrete Boom Pump 30 cum/hr", "Boom pumping vertically through articulated boom up to floor V level", "High-pressure concrete pump hire; output = 12.0 cum/hr", "Machine", "0.083 machine-hrs (12.0 cum/hr)", "1 cum"),
            ("Needle Vibrator", "Compacting pumped concrete in elevated forms", "Vibrator hire; output = 2.0 cum/hr", "Machine", "0.50 machine-hrs (2.00 cum/hr)", "1 cum"),
            ("Mason (average)", "Controlling elevated pour, leveling, and screeding to deck benchmarks", "Mason labour; 0.12 day = 0.96 man-hrs per cum", "Labour", "0.96 man-hrs (1.04 cum/man-hr)", "1 cum"),
            ("Beldar", "Handling boom hose end, spreading, and operating vibrator on elevated floor", "Labour crew; 0.80 day = 6.40 man-hrs per cum", "Labour", "6.40 man-hrs (0.16 cum/man-hr)", "1 cum"),
            ("Coolie", "Safety watch, guiding boom, coupling riser pipes, and washing hopper", "Helper crew; 0.50 day = 4.00 man-hrs per cum", "Labour", "4.00 man-hrs (0.25 cum/man-hr)", "1 cum"),
            ("Bhisti", "Water supply and initial deck curing", "Watering crew; 0.35 day = 2.80 man-hrs per cum", "Labour", "2.80 man-hrs (0.36 cum/man-hr)", "1 cum"),
            ("Sundries", "Elevated safety barriers, hose slings, washout sponges, and testing cubes", "Sundries allowance", "Equipment", "L.S. allowance", "1 cum"),
        ]
    },

    # ─── 4.20 READY MIXED CONCRETE (RMC) WITHOUT FLY ASH ────────────────────────
    {
        "id": "4.20.1.1",
        "parent_title": "4.20 READY MIXED CONCRETE (RMC) WITHOUT FLY ASH — UP TO PLINTH LEVEL",
        "title": "4.20.1.1 RMC M-15 grade plain cement concrete (pure cement) up to plinth level (cement considered @ 240 kg/cum)",
        "unit": "cum", "base_qty": 1.0,
        "rows": [
            ("Stone Aggregate 40 mm", "Coarse aggregate 40 mm batched at RMC plant (0.65 cum)", "Clean graded crushed stone aggregate 40 mm nominal size", "Material", "0.65 cum (0.650 cum/cum)", "1 cum"),
            ("Stone Aggregate 20 mm", "Coarse aggregate 20 mm batched at RMC plant (0.24 cum)", "Clean graded crushed stone aggregate 20 mm nominal size", "Material", "0.24 cum (0.240 cum/cum)", "1 cum"),
            ("Carriage of Stone aggregate", "Haulage of 40mm and 20mm aggregates to RMC plant (0.89 cum)", "Carriage allowance", "Material", "0.89 cum (0.890 cum/cum)", "1 cum"),
            ("Coarse sand", "Zone III coarse sand batched at RMC plant (0.47 cum)", "Clean natural river sand Zone III", "Material", "0.47 cum (0.470 cum/cum)", "1 cum"),
            ("Carriage of Coarse sand", "Haulage of coarse sand to RMC plant (0.47 cum)", "Carriage allowance", "Material", "0.47 cum (0.470 cum/cum)", "1 cum"),
            ("Portland Cement", "Pure OPC-43 Grade cement batched in RMC silo (0.24 tonne = 240 kg)", "Conforming to IS:269", "Material", "0.24 tonne (0.240 t/cum)", "1 cum"),
            ("Carriage of Cement", "Bulk tanker haulage of cement to RMC plant silo (0.24 tonne)", "Carriage allowance", "Material", "0.24 tonne (0.240 t/cum)", "1 cum"),
            ("Chemical Plasticizer / Retarder", "Admixture for workability and slump retention during transit (1.5 kg)", "Conforming to IS:9103", "Material", "1.50 kg (1.500 kg/cum)", "1 cum"),
            ("Transit Mixer 6.0 cum capacity", "Haulage of wet concrete from batching plant to site under agitation", "Transit mixer hire; output = 3.0 cum/hr", "Machine", "0.33 machine-hrs (3.00 cum/hr)", "1 cum"),
            ("Concrete Pump / Placer 30 cum/hr", "Pumping concrete directly into foundation trench/bed", "Concrete pump hire; output = 15.0 cum/hr", "Machine", "0.067 machine-hrs (15.0 cum/hr)", "1 cum"),
            ("Needle / Surface Vibrator", "Consolidating pumped concrete", "Vibrator hire; output = 2.0 cum/hr", "Machine", "0.50 machine-hrs (2.00 cum/hr)", "1 cum"),
            ("Mason (average)", "Setting grade stakes, screeding, and float finishing", "Mason labour; 0.08 day = 0.64 man-hrs per cum", "Labour", "0.64 man-hrs (1.56 cum/man-hr)", "1 cum"),
            ("Beldar", "Handling hose end, spreading, and operating vibrator", "Labour crew; 0.60 day = 4.80 man-hrs per cum", "Labour", "4.80 man-hrs (0.21 cum/man-hr)", "1 cum"),
            ("Coolie", "Guiding mixer, coupling pipes, and cleaning pump", "Helper crew; 0.40 day = 3.20 man-hrs per cum", "Labour", "3.20 man-hrs (0.31 cum/man-hr)", "1 cum"),
            ("Bhisti", "Water supply and initial foundation curing", "Watering crew; 0.30 day = 2.40 man-hrs per cum", "Labour", "2.40 man-hrs (0.42 cum/man-hr)", "1 cum"),
            ("Sundries", "Pipeline washouts, sponge balls, and testing cubes", "Sundries allowance", "Equipment", "L.S. allowance", "1 cum"),
        ]
    },
    {
        "id": "4.20.1.2",
        "parent_title": "4.20 READY MIXED CONCRETE (RMC) WITHOUT FLY ASH — UP TO PLINTH LEVEL",
        "title": "4.20.1.2 RMC M-10 grade plain cement concrete (pure cement) up to plinth level (cement considered @ 220 kg/cum)",
        "unit": "cum", "base_qty": 1.0,
        "rows": [
            ("Stone Aggregate 40 mm", "Coarse aggregate 40 mm batched at RMC plant (0.65 cum)", "Clean graded crushed stone aggregate 40 mm nominal size", "Material", "0.65 cum (0.650 cum/cum)", "1 cum"),
            ("Stone Aggregate 20 mm", "Coarse aggregate 20 mm batched at RMC plant (0.24 cum)", "Clean graded crushed stone aggregate 20 mm nominal size", "Material", "0.24 cum (0.240 cum/cum)", "1 cum"),
            ("Carriage of Stone aggregate", "Haulage of 40mm and 20mm aggregates to RMC plant (0.89 cum)", "Carriage allowance", "Material", "0.89 cum (0.890 cum/cum)", "1 cum"),
            ("Coarse sand", "Zone III coarse sand batched at RMC plant (0.47 cum)", "Clean natural river sand Zone III", "Material", "0.47 cum (0.470 cum/cum)", "1 cum"),
            ("Carriage of Coarse sand", "Haulage of coarse sand to RMC plant (0.47 cum)", "Carriage allowance", "Material", "0.47 cum (0.470 cum/cum)", "1 cum"),
            ("Portland Cement", "Pure OPC-43 Grade cement batched in RMC silo (0.22 tonne = 220 kg)", "Conforming to IS:269", "Material", "0.22 tonne (0.220 t/cum)", "1 cum"),
            ("Carriage of Cement", "Bulk tanker haulage of cement to RMC plant silo (0.22 tonne)", "Carriage allowance", "Material", "0.22 tonne (0.220 t/cum)", "1 cum"),
            ("Chemical Plasticizer / Retarder", "Admixture for workability and slump retention during transit (1.4 kg)", "Conforming to IS:9103", "Material", "1.40 kg (1.400 kg/cum)", "1 cum"),
            ("Transit Mixer 6.0 cum capacity", "Haulage of wet concrete from batching plant to site under agitation", "Transit mixer hire; output = 3.0 cum/hr", "Machine", "0.33 machine-hrs (3.00 cum/hr)", "1 cum"),
            ("Concrete Pump / Placer 30 cum/hr", "Pumping concrete directly into foundation trench/bed", "Concrete pump hire; output = 15.0 cum/hr", "Machine", "0.067 machine-hrs (15.0 cum/hr)", "1 cum"),
            ("Needle / Surface Vibrator", "Consolidating pumped concrete", "Vibrator hire; output = 2.0 cum/hr", "Machine", "0.50 machine-hrs (2.00 cum/hr)", "1 cum"),
            ("Mason (average)", "Setting grade stakes, screeding, and float finishing", "Mason labour; 0.08 day = 0.64 man-hrs per cum", "Labour", "0.64 man-hrs (1.56 cum/man-hr)", "1 cum"),
            ("Beldar", "Handling hose end, spreading, and operating vibrator", "Labour crew; 0.60 day = 4.80 man-hrs per cum", "Labour", "4.80 man-hrs (0.21 cum/man-hr)", "1 cum"),
            ("Coolie", "Guiding mixer, coupling pipes, and cleaning pump", "Helper crew; 0.40 day = 3.20 man-hrs per cum", "Labour", "3.20 man-hrs (0.31 cum/man-hr)", "1 cum"),
            ("Bhisti", "Water supply and initial foundation curing", "Watering crew; 0.30 day = 2.40 man-hrs per cum", "Labour", "2.40 man-hrs (0.42 cum/man-hr)", "1 cum"),
            ("Sundries", "Pipeline washouts, sponge balls, and testing cubes", "Sundries allowance", "Equipment", "L.S. allowance", "1 cum"),
        ]
    },
    {
        "id": "4.20.2.1",
        "parent_title": "4.20 READY MIXED CONCRETE (RMC) WITHOUT FLY ASH — ABOVE PLINTH TO FLOOR V",
        "title": "4.20.2.1 RMC M-15 grade plain cement concrete (pure cement) above plinth up to floor V level",
        "unit": "cum", "base_qty": 1.0,
        "rows": [
            ("Stone Aggregate 40 mm & 20 mm", "Graded stone aggregate 40mm (0.65 cum) and 20mm (0.24 cum) batched at RMC plant", "Clean crushed stone aggregate", "Material", "0.89 cum (0.890 cum/cum)", "1 cum"),
            ("Coarse sand", "Zone III coarse sand batched at RMC plant (0.47 cum)", "Clean natural river sand Zone III", "Material", "0.47 cum (0.470 cum/cum)", "1 cum"),
            ("Portland Cement", "Pure OPC-43 Grade cement batched in RMC silo (0.24 tonne = 240 kg)", "Conforming to IS:269", "Material", "0.24 tonne (0.240 t/cum)", "1 cum"),
            ("Carriage of Materials", "Bulk haulage of cement, sand and aggregates to RMC batching plant", "Carriage allowance", "Material", "1.60 cum/t equiv", "1 cum"),
            ("Chemical Plasticizer / Retarder", "High-range plasticizer maintaining slump retention (1.5 kg)", "Conforming to IS:9103", "Material", "1.50 kg (1.500 kg/cum)", "1 cum"),
            ("Transit Mixer 6.0 cum capacity", "Haulage from RMC plant to site under agitation", "Transit mixer hire; output = 3.0 cum/hr", "Machine", "0.33 machine-hrs (3.00 cum/hr)", "1 cum"),
            ("Concrete Boom Pump 30 cum/hr", "Boom pumping vertically through articulated boom up to floor V level", "High-pressure concrete pump hire; output = 12.0 cum/hr", "Machine", "0.083 machine-hrs (12.0 cum/hr)", "1 cum"),
            ("Needle Vibrator", "Compacting pumped concrete in elevated forms", "Vibrator hire; output = 2.0 cum/hr", "Machine", "0.50 machine-hrs (2.00 cum/hr)", "1 cum"),
            ("Mason (average)", "Controlling elevated pour, leveling, and screeding to deck benchmarks", "Mason labour; 0.12 day = 0.96 man-hrs per cum", "Labour", "0.96 man-hrs (1.04 cum/man-hr)", "1 cum"),
            ("Beldar", "Handling boom hose end, spreading, and operating vibrator on elevated floor", "Labour crew; 0.80 day = 6.40 man-hrs per cum", "Labour", "6.40 man-hrs (0.16 cum/man-hr)", "1 cum"),
            ("Coolie", "Safety watch, guiding boom, coupling riser pipes, and washing hopper", "Helper crew; 0.50 day = 4.00 man-hrs per cum", "Labour", "4.00 man-hrs (0.25 cum/man-hr)", "1 cum"),
            ("Bhisti", "Water supply and initial deck curing", "Watering crew; 0.35 day = 2.80 man-hrs per cum", "Labour", "2.80 man-hrs (0.36 cum/man-hr)", "1 cum"),
            ("Sundries", "Elevated safety barriers, hose slings, washout sponges, and testing cubes", "Sundries allowance", "Equipment", "L.S. allowance", "1 cum"),
        ]
    },
    {
        "id": "4.20.2.2",
        "parent_title": "4.20 READY MIXED CONCRETE (RMC) WITHOUT FLY ASH — ABOVE PLINTH TO FLOOR V",
        "title": "4.20.2.2 RMC M-10 grade plain cement concrete (pure cement) above plinth up to floor V level",
        "unit": "cum", "base_qty": 1.0,
        "rows": [
            ("Stone Aggregate 40 mm & 20 mm", "Graded stone aggregate 40mm (0.65 cum) and 20mm (0.24 cum) batched at RMC plant", "Clean crushed stone aggregate", "Material", "0.89 cum (0.890 cum/cum)", "1 cum"),
            ("Coarse sand", "Zone III coarse sand batched at RMC plant (0.47 cum)", "Clean natural river sand Zone III", "Material", "0.47 cum (0.470 cum/cum)", "1 cum"),
            ("Portland Cement", "Pure OPC-43 Grade cement batched in RMC silo (0.22 tonne = 220 kg)", "Conforming to IS:269", "Material", "0.22 tonne (0.220 t/cum)", "1 cum"),
            ("Carriage of Materials", "Bulk haulage of cement, sand and aggregates to RMC batching plant", "Carriage allowance", "Material", "1.58 cum/t equiv", "1 cum"),
            ("Chemical Plasticizer / Retarder", "High-range plasticizer maintaining slump retention (1.4 kg)", "Conforming to IS:9103", "Material", "1.40 kg (1.400 kg/cum)", "1 cum"),
            ("Transit Mixer 6.0 cum capacity", "Haulage from RMC plant to site under agitation", "Transit mixer hire; output = 3.0 cum/hr", "Machine", "0.33 machine-hrs (3.00 cum/hr)", "1 cum"),
            ("Concrete Boom Pump 30 cum/hr", "Boom pumping vertically through articulated boom up to floor V level", "High-pressure concrete pump hire; output = 12.0 cum/hr", "Machine", "0.083 machine-hrs (12.0 cum/hr)", "1 cum"),
            ("Needle Vibrator", "Compacting pumped concrete in elevated forms", "Vibrator hire; output = 2.0 cum/hr", "Machine", "0.50 machine-hrs (2.00 cum/hr)", "1 cum"),
            ("Mason (average)", "Controlling elevated pour, leveling, and screeding to deck benchmarks", "Mason labour; 0.12 day = 0.96 man-hrs per cum", "Labour", "0.96 man-hrs (1.04 cum/man-hr)", "1 cum"),
            ("Beldar", "Handling boom hose end, spreading, and operating vibrator on elevated floor", "Labour crew; 0.80 day = 6.40 man-hrs per cum", "Labour", "6.40 man-hrs (0.16 cum/man-hr)", "1 cum"),
            ("Coolie", "Safety watch, guiding boom, coupling riser pipes, and washing hopper", "Helper crew; 0.50 day = 4.00 man-hrs per cum", "Labour", "4.00 man-hrs (0.25 cum/man-hr)", "1 cum"),
            ("Bhisti", "Water supply and initial deck curing", "Watering crew; 0.35 day = 2.80 man-hrs per cum", "Labour", "2.80 man-hrs (0.36 cum/man-hr)", "1 cum"),
            ("Sundries", "Elevated safety barriers, hose slings, washout sponges, and testing cubes", "Sundries allowance", "Equipment", "L.S. allowance", "1 cum"),
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

def build_concrete_sheet(ws):
    # Configure sheet view
    if ws.views.sheetView:
        ws.views.sheetView[0].showGridLines = True

    # Sheet title spanning A1:G1
    ws.merge_cells("A1:G1")
    t1 = ws.cell(row=1, column=1, value="Sub-Head 4.0 — CONCRETE WORK  |  First-Principles Resource, Work & Gang Analysis")
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
    for item in CONCRETE_ITEMS:
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

def rebuild_concrete_in_workbook(file_path):
    print(f"Opening {file_path} ...")
    wb = openpyxl.load_workbook(file_path)
    if "04_Concrete_Work" not in wb.sheetnames:
        raise ValueError(f"'04_Concrete_Work' sheet not found in {file_path}")

    pos = wb.sheetnames.index("04_Concrete_Work")
    del wb["04_Concrete_Work"]
    ws = wb.create_sheet("04_Concrete_Work", pos)
    ws.sheet_properties.tabColor = "2E75B6"
    build_concrete_sheet(ws)

    temp_path = file_path.replace(".xlsx", "_TMP_CONCRETE.xlsx")
    wb.save(temp_path)
    wb.close()
    os.replace(temp_path, file_path)
    print(f"Successfully updated {file_path} -> sheet '04_Concrete_Work' ({len(CONCRETE_ITEMS)} items).")

def main():
    repo_root = Path(__file__).resolve().parents[1]
    main_wb = repo_root / "CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_1_Latest.xlsx"
    output_wb = repo_root / "outputs" / "earthwork-custom-rate-composer" / "CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_1_Latest.xlsx"

    if main_wb.exists():
        rebuild_concrete_in_workbook(str(main_wb))
    if output_wb.exists():
        rebuild_concrete_in_workbook(str(output_wb))

if __name__ == "__main__":
    main()
