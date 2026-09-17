"""
support_builder_mortars.py
Rebuilds the existing '03_Mortars' worksheet with First-Principles Resource,
Productivity, Gang and Material Analysis per CPWD DAR 2019 Vol 1:
  Item Code | Labour / Machine / Material | Work done | Condition / When used | Category | Productivity | Quantity
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from pathlib import Path
import os, sys, shutil

# Styling definitions matching 02_support_earth_work
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

# Master Mortar Items Definition (Items 3.1 to 3.18)
MORTAR_ITEMS = [
    {
        "id": "3.1",
        "parent_title": "3.0 MORTARS — CEMENT MORTARS WITH FINE SAND",
        "title": "3.1 Cement mortar 1:1 (1 cement : 1 fine sand)",
        "unit": "cum",
        "base_qty": 1.0,
        "rows": [
            ("Portland Cement", "Primary cementitious hydraulic binder providing high-strength paste matrix (0.7175 cum = 1.02 tonne)", "Very rich cement mortar for specialised pointing, urgent waterproofing fillets, high-strength patch repairs, and heavy-duty wearing beds", "Material", "1.02 tonne (1.020 t/cum)", "1 cum"),
            ("Carriage of Cement", "Loading, mechanical transport and unloading of cement bags from store/depot to site batching point", "Haulage of cement bags within site lead", "Material", "1.02 tonne (1.020 t/cum)", "1 cum"),
            ("Fine sand", "Inert fine aggregate providing matrix density and dimensional stability (0.7175 cum)", "Zone IV / fine sand conforming to CPWD specification; silt content <= 3%", "Material", "0.72 cum (0.718 cum/cum)", "1 cum"),
            ("Carriage of Fine sand", "Mechanical transport, haulage, unloading and stacking of fine sand at site mixing platform", "Carriage from source/stockpile to batching yard", "Material", "0.72 cum (0.718 cum/cum)", "1 cum"),
            ("Mechanical Mortar Mixer", "Machine batching and thorough wet mixing of cement and fine sand to homogeneous consistency", "Mechanical mixer hire; cycle duration ~9 min per 1 cum batch; output = 6.67 cum/hr", "Machine", "0.15 machine-hrs (6.67 cum/hr)", "1 cum"),
            ("Beldar", "Measuring sand and cement, loading into mixer drum/mixing platform, handling ingredients and discharging wet mortar", "Manual batching, feeding, handling, and transferring mixed mortar into mortar pans", "Labour", "6.00 man-hrs (0.17 cum/man-hr)", "1 cum"),
            ("Bhisti", "Adding measured water during mixing cycle to maintain specified water-cement ratio and washing equipment", "Water conditioning, watering during mixing and cleaning mixing drum/platform", "Labour", "0.56 man-hrs (1.79 cum/man-hr)", "1 cum"),
            ("Sundries", "Hand mixing pans, shovels, water buckets, platform cleaning, and testing allowances", "Sundries and minor tools allowance (L.S. 13.52 x cost index)", "Equipment", "L.S. allowance", "1 cum"),
        ]
    },
    {
        "id": "3.2",
        "parent_title": "3.0 MORTARS — CEMENT MORTARS WITH FINE SAND",
        "title": "3.2 Cement mortar 1:2 (1 cement : 2 fine sand)",
        "unit": "cum",
        "base_qty": 1.0,
        "rows": [
            ("Portland Cement", "Primary cementitious hydraulic binder providing dense impermeable paste (0.475 cum = 0.68 tonne)", "Rich waterproof mortar for damp-proof courses (DPC), water-retaining structures, pointing, and thin precast units", "Material", "0.68 tonne (0.680 t/cum)", "1 cum"),
            ("Carriage of Cement", "Loading, transport and unloading of cement bags to site batching point", "Standard site haulage allowance for cement bags", "Material", "0.68 tonne (0.680 t/cum)", "1 cum"),
            ("Fine sand", "Inert fine aggregate providing body and preventing excessive shrinkage cracks (0.950 cum)", "Fine sand conforming to CPWD specifications for waterproofing/pointing beds", "Material", "0.95 cum (0.950 cum/cum)", "1 cum"),
            ("Carriage of Fine sand", "Mechanical haulage, unloading and stacking of fine sand at site platform", "Carriage from local supply yard to mixing plant", "Material", "0.95 cum (0.950 cum/cum)", "1 cum"),
            ("Mechanical Mortar Mixer", "Machine batching and thorough wet mixing to achieve uniform plastic workability", "Mechanical mixer hire; cycle duration ~9 min per 1 cum batch; output = 6.67 cum/hr", "Machine", "0.15 machine-hrs (6.67 cum/hr)", "1 cum"),
            ("Beldar", "Measuring ingredients, loading mixer, discharging and loading mixed mortar into mortar pans", "Mixing and handling crew; 0.75 day per cum batch", "Labour", "6.00 man-hrs (0.17 cum/man-hr)", "1 cum"),
            ("Bhisti", "Supplying controlled water for mixing and keeping mixing trays moist and washed", "Water supply and mixing conditioning; 0.07 day per cum batch", "Labour", "0.56 man-hrs (1.79 cum/man-hr)", "1 cum"),
            ("Sundries", "Small tools, shovels, water hoses, measuring boxes, and platform maintenance", "Sundries and minor equipment allowance (L.S. 13.52 x cost index)", "Equipment", "L.S. allowance", "1 cum"),
        ]
    },
    {
        "id": "3.3",
        "parent_title": "3.0 MORTARS — CEMENT MORTARS WITH FINE SAND",
        "title": "3.3 Cement mortar 1:3 (1 cement : 3 fine sand)",
        "unit": "cum",
        "base_qty": 1.0,
        "rows": [
            ("Portland Cement", "Hydraulic binder for high-bond structural bedding (0.357 cum = 0.51 tonne)", "High-strength mortar for load-bearing brickwork, damp environments, external rich plaster, and stone cladding", "Material", "0.51 tonne (0.510 t/cum)", "1 cum"),
            ("Carriage of Cement", "Transport of cement bags from site godown to mortar mixing station", "Site haulage of bagged cement", "Material", "0.51 tonne (0.510 t/cum)", "1 cum"),
            ("Fine sand", "Fine aggregate filler conforming to CPWD standards (1.07 cum)", "Clean screened fine sand free from organic matter and excessive silt", "Material", "1.07 cum (1.070 cum/cum)", "1 cum"),
            ("Carriage of Fine sand", "Carriage and stacking of fine sand at site batching platform", "Bulk haulage of sand to site", "Material", "1.07 cum (1.070 cum/cum)", "1 cum"),
            ("Mechanical Mortar Mixer", "Mechanical mixing of 1:3 mortar to prevent segregation and ensure uniform hydration", "Mixer operational time; output = 6.67 cum/hr", "Machine", "0.15 machine-hrs (6.67 cum/hr)", "1 cum"),
            ("Beldar", "Measuring sand and cement boxes, feeding mixer drum, and discharging into distribution pans", "Crew labour allowance; 0.75 day = 6.00 man-hrs per cum", "Labour", "6.00 man-hrs (0.17 cum/man-hr)", "1 cum"),
            ("Bhisti", "Measuring and adding mixing water and washing mixer drum after discharge", "Watering and equipment cleaning; 0.07 day = 0.56 man-hrs per cum", "Labour", "0.56 man-hrs (1.79 cum/man-hr)", "1 cum"),
            ("Sundries", "Sundry tools, wooden measuring farma, mixing sheets, and iron pans", "Sundries and tools allowance", "Equipment", "L.S. allowance", "1 cum"),
        ]
    },
    {
        "id": "3.4",
        "parent_title": "3.0 MORTARS — CEMENT MORTARS WITH FINE SAND",
        "title": "3.4 Cement mortar 1:4 (1 cement : 4 fine sand)",
        "unit": "cum",
        "base_qty": 1.0,
        "rows": [
            ("Portland Cement", "Standard hydraulic binder for exterior/interior plaster and reinforced brickwork (0.268 cum = 0.38 tonne)", "Standard structural mortar for general brickwork, ceiling plaster, and exterior plastering", "Material", "0.38 tonne (0.380 t/cum)", "1 cum"),
            ("Carriage of Cement", "Site handling and carriage of cement bags to mixing point", "Haulage of bagged cement", "Material", "0.38 tonne (0.380 t/cum)", "1 cum"),
            ("Fine sand", "Fine aggregate providing optimum grain interlock (1.07 cum)", "Clean fine sand for plaster and mortar bedding", "Material", "1.07 cum (1.070 cum/cum)", "1 cum"),
            ("Carriage of Fine sand", "Carriage, unloading and stacking of fine sand at mixing yard", "Bulk sand transport", "Material", "1.07 cum (1.070 cum/cum)", "1 cum"),
            ("Mechanical Mortar Mixer", "Uniform mechanical mixing for homogeneous plaster/bedding mortar", "Mixer hire allowance; output = 6.67 cum/hr", "Machine", "0.15 machine-hrs (6.67 cum/hr)", "1 cum"),
            ("Beldar", "Batching with farmas, charging mixer drum, wheeling and discharging mixed mortar", "Mixing crew; 0.75 day per cum batch", "Labour", "6.00 man-hrs (0.17 cum/man-hr)", "1 cum"),
            ("Bhisti", "Water batching, maintaining slump, and drum cleanup", "Watering crew; 0.07 day per cum batch", "Labour", "0.56 man-hrs (1.79 cum/man-hr)", "1 cum"),
            ("Sundries", "Measuring farmas, shovels, pans, platform maintenance", "Sundries allowance", "Equipment", "L.S. allowance", "1 cum"),
        ]
    },
    {
        "id": "3.5",
        "parent_title": "3.0 MORTARS — CEMENT MORTARS WITH FINE SAND",
        "title": "3.5 Cement mortar 1:5 (1 cement : 5 fine sand)",
        "unit": "cum",
        "base_qty": 1.0,
        "rows": [
            ("Portland Cement", "Hydraulic binder for general masonry and internal plastering (0.214 cum = 0.31 tonne)", "General masonry mortar for superstructure walls, internal partitions, and undercoat plastering", "Material", "0.31 tonne (0.310 t/cum)", "1 cum"),
            ("Carriage of Cement", "Handling and transport of cement bags to batching yard", "Site transport of cement", "Material", "0.31 tonne (0.310 t/cum)", "1 cum"),
            ("Fine sand", "Screened fine sand matrix filler (1.07 cum)", "Standard fine sand for internal brickwork and plaster", "Material", "1.07 cum (1.070 cum/cum)", "1 cum"),
            ("Carriage of Fine sand", "Carriage and dumping of fine sand at site platform", "Sand transport to mixing location", "Material", "1.07 cum (1.070 cum/cum)", "1 cum"),
            ("Mechanical Mortar Mixer", "Batch mixing to ensure cement dispersion throughout lean aggregate volume", "Mixer operational hire; output = 6.67 cum/hr", "Machine", "0.15 machine-hrs (6.67 cum/hr)", "1 cum"),
            ("Beldar", "Measuring sand and cement boxes, charging drum, and discharging into iron pans", "Labour allowance; 0.75 day = 6.00 man-hrs per cum", "Labour", "6.00 man-hrs (0.17 cum/man-hr)", "1 cum"),
            ("Bhisti", "Watering batch and washing tools and mixing platform", "Watering allowance; 0.07 day = 0.56 man-hrs per cum", "Labour", "0.56 man-hrs (1.79 cum/man-hr)", "1 cum"),
            ("Sundries", "Hand shovels, mixing trays, farma boxes, and cleanup allowance", "Sundries allowance", "Equipment", "L.S. allowance", "1 cum"),
        ]
    },
    {
        "id": "3.6",
        "parent_title": "3.0 MORTARS — CEMENT MORTARS WITH FINE SAND",
        "title": "3.6 Cement mortar 1:6 (1 cement : 6 fine sand)",
        "unit": "cum",
        "base_qty": 1.0,
        "rows": [
            ("Portland Cement", "Economical hydraulic binder for mass masonry and non-structural walls (0.178 cum = 0.25 tonne)", "Lean mortar widely used for non-load bearing brick masonry, boundary walls, and mass brickwork", "Material", "0.25 tonne (0.250 t/cum)", "1 cum"),
            ("Carriage of Cement", "Transport of cement bags to site mixing platform", "Standard carriage allowance for cement", "Material", "0.25 tonne (0.250 t/cum)", "1 cum"),
            ("Fine sand", "Fine aggregate filler (1.07 cum)", "Clean fine sand conforming to CPWD specifications", "Material", "1.07 cum (1.070 cum/cum)", "1 cum"),
            ("Carriage of Fine sand", "Mechanical transport and unloading of sand at site", "Sand carriage to batching platform", "Material", "1.07 cum (1.070 cum/cum)", "1 cum"),
            ("Mechanical Mortar Mixer", "Mechanical mixing ensuring thorough dispersion of lean cement content", "Mixer hire; output = 6.67 cum/hr", "Machine", "0.15 machine-hrs (6.67 cum/hr)", "1 cum"),
            ("Beldar", "Measuring, mixer loading, discharging and distributing mortar", "Labour allowance; 0.75 day = 6.00 man-hrs per cum", "Labour", "6.00 man-hrs (0.17 cum/man-hr)", "1 cum"),
            ("Bhisti", "Controlled water supply during mixing and washing pans", "Watering allowance; 0.07 day = 0.56 man-hrs per cum", "Labour", "0.56 man-hrs (1.79 cum/man-hr)", "1 cum"),
            ("Sundries", "Sundries, measuring boxes, pans, and platform upkeep", "Sundries allowance", "Equipment", "L.S. allowance", "1 cum"),
        ]
    },
    {
        "id": "3.7",
        "parent_title": "3.0 MORTARS — CEMENT MORTARS WITH COARSE SAND",
        "title": "3.7 Cement mortar 1:2 (1 cement : 2 coarse sand)",
        "unit": "cum",
        "base_qty": 1.0,
        "rows": [
            ("Portland Cement", "High-strength hydraulic binder for structural bedding and waterproof linings (0.476 cum = 0.68 tonne)", "Rich mortar with coarse aggregate for heavy load-bearing brickwork, damp-proof courses, and external rough plaster", "Material", "0.68 tonne (0.680 t/cum)", "1 cum"),
            ("Carriage of Cement", "Site handling and carriage of cement bags to mixing yard", "Carriage of bagged cement", "Material", "0.68 tonne (0.680 t/cum)", "1 cum"),
            ("Coarse sand", "Graded coarse sand (Zone III) providing high interlocking compressive strength (0.95 cum)", "Zone III coarse sand conforming to IS:383 for structural mortars", "Material", "0.95 cum (0.950 cum/cum)", "1 cum"),
            ("Carriage of Coarse sand", "Mechanical haulage and unloading of coarse sand at site platform", "Carriage of coarse sand to batching yard", "Material", "0.95 cum (0.950 cum/cum)", "1 cum"),
            ("Mechanical Mortar Mixer", "Mechanical batch mixing for dense coarse-aggregate mortar paste", "Mixer operational hire; output = 6.67 cum/hr", "Machine", "0.15 machine-hrs (6.67 cum/hr)", "1 cum"),
            ("Beldar", "Measuring coarse sand and cement, charging mixer, handling and discharging mixed mortar", "Mixing and handling crew; 0.75 day = 6.00 man-hrs per cum", "Labour", "6.00 man-hrs (0.17 cum/man-hr)", "1 cum"),
            ("Bhisti", "Water batching, maintaining workability, and cleaning mixer drum", "Watering crew; 0.07 day = 0.56 man-hrs per cum", "Labour", "0.56 man-hrs (1.79 cum/man-hr)", "1 cum"),
            ("Sundries", "Farma boxes, shovels, water buckets, and cleaning allowance", "Sundries allowance", "Equipment", "L.S. allowance", "1 cum"),
        ]
    },
    {
        "id": "3.8",
        "parent_title": "3.0 MORTARS — CEMENT MORTARS WITH COARSE SAND",
        "title": "3.8 Cement mortar 1:3 (1 cement : 3 coarse sand)",
        "unit": "cum",
        "base_qty": 1.0,
        "rows": [
            ("Portland Cement", "Hydraulic binder for heavy-duty structural masonry and cladding bedding (0.357 cum = 0.51 tonne)", "Specified as primary bedding mortar for stone cladding (REF#8.1.1), heavy stone masonry, and foundation plinth walls", "Material", "0.51 tonne (0.510 t/cum)", "1 cum"),
            ("Carriage of Cement", "Carriage of cement bags from site store to batching plant", "Cement transport within site lead", "Material", "0.51 tonne (0.510 t/cum)", "1 cum"),
            ("Coarse sand", "Graded coarse sand (Zone III) ensuring high shear resistance and low shrinkage (1.07 cum)", "Zone III coarse sand free from organic matter and clay lumps", "Material", "1.07 cum (1.070 cum/cum)", "1 cum"),
            ("Carriage of Coarse sand", "Haulage and stacking of coarse sand at site mixing platform", "Bulk coarse sand transport", "Material", "1.07 cum (1.070 cum/cum)", "1 cum"),
            ("Mechanical Mortar Mixer", "Thorough mechanical mixing to achieve uniform paste coating over coarse grains", "Mixer operational hire; output = 6.67 cum/hr", "Machine", "0.15 machine-hrs (6.67 cum/hr)", "1 cum"),
            ("Beldar", "Measuring sand and cement boxes, feeding mixer drum, and discharging into iron pans", "Labour allowance; 0.75 day = 6.00 man-hrs per cum", "Labour", "6.00 man-hrs (0.17 cum/man-hr)", "1 cum"),
            ("Bhisti", "Adding measured water during mixing and washing drum after discharge", "Watering allowance; 0.07 day = 0.56 man-hrs per cum", "Labour", "0.56 man-hrs (1.79 cum/man-hr)", "1 cum"),
            ("Sundries", "Sundry tools, measuring boxes, pans, and platform maintenance", "Sundries allowance", "Equipment", "L.S. allowance", "1 cum"),
        ]
    },
    {
        "id": "3.9",
        "parent_title": "3.0 MORTARS — CEMENT MORTARS WITH COARSE SAND",
        "title": "3.9 Cement mortar 1:4 (1 cement : 4 coarse sand)",
        "unit": "cum",
        "base_qty": 1.0,
        "rows": [
            ("Portland Cement", "Standard hydraulic binder for general structural brickwork and flooring bed (0.268 cum = 0.38 tonne)", "The benchmark mortar of CPWD DAR; widely consumed by Masonry (REF#6.1.1), Flooring (REF#11.1), and Road works", "Material", "0.38 tonne (0.380 t/cum)", "1 cum"),
            ("Carriage of Cement", "Loading, mechanical carriage and unloading of cement bags at site", "Site carriage of cement bags", "Material", "0.38 tonne (0.380 t/cum)", "1 cum"),
            ("Coarse sand", "Zone III coarse sand providing structural matrix for brick joints (1.07 cum)", "Clean coarse sand conforming to IS:383 Zone III", "Material", "1.07 cum (1.070 cum/cum)", "1 cum"),
            ("Carriage of Coarse sand", "Haulage, unloading and stacking of coarse sand at site platform", "Carriage of coarse sand", "Material", "1.07 cum (1.070 cum/cum)", "1 cum"),
            ("Mechanical Mortar Mixer", "Mechanical mixing of 1:4 mix ensuring complete hydration and plasticity", "Mixer operational hire; output = 6.67 cum/hr", "Machine", "0.15 machine-hrs (6.67 cum/hr)", "1 cum"),
            ("Beldar", "Measuring ingredients with wooden farmas, feeding drum, and discharging into mortar pans", "Mixing crew; 0.75 day = 6.00 man-hrs per cum", "Labour", "6.00 man-hrs (0.17 cum/man-hr)", "1 cum"),
            ("Bhisti", "Water batching, maintaining specified consistency, and cleaning platform", "Watering crew; 0.07 day = 0.56 man-hrs per cum", "Labour", "0.56 man-hrs (1.79 cum/man-hr)", "1 cum"),
            ("Sundries", "Measuring farmas, shovels, mortar trays, wheelbarrows, and testing tools", "Sundries allowance (L.S. 13.52 x cost index)", "Equipment", "L.S. allowance", "1 cum"),
        ]
    },
    {
        "id": "3.10",
        "parent_title": "3.0 MORTARS — CEMENT MORTARS WITH COARSE SAND",
        "title": "3.10 Cement mortar 1:5 (1 cement : 5 coarse sand)",
        "unit": "cum",
        "base_qty": 1.0,
        "rows": [
            ("Portland Cement", "Hydraulic binder for load-bearing and partition masonry walls (0.214 cum = 0.31 tonne)", "Commonly specified for superstructure brickwork in multi-storey buildings and external brick facing", "Material", "0.31 tonne (0.310 t/cum)", "1 cum"),
            ("Carriage of Cement", "Transport of cement bags from store to batching yard", "Site transport of bagged cement", "Material", "0.31 tonne (0.310 t/cum)", "1 cum"),
            ("Coarse sand", "Zone III coarse sand providing interlocking aggregate skeleton (1.07 cum)", "Clean coarse sand conforming to IS:383", "Material", "1.07 cum (1.070 cum/cum)", "1 cum"),
            ("Carriage of Coarse sand", "Mechanical haulage and unloading of coarse sand at site", "Carriage of coarse sand", "Material", "1.07 cum (1.070 cum/cum)", "1 cum"),
            ("Mechanical Mortar Mixer", "Mechanical mixing to ensure uniform cement film coating on all sand grains", "Mixer hire allowance; output = 6.67 cum/hr", "Machine", "0.15 machine-hrs (6.67 cum/hr)", "1 cum"),
            ("Beldar", "Measuring, drum charging, discharging, and loading pans", "Labour allowance; 0.75 day = 6.00 man-hrs per cum", "Labour", "6.00 man-hrs (0.17 cum/man-hr)", "1 cum"),
            ("Bhisti", "Supplying mixing water and washing mixer drum after discharge", "Watering allowance; 0.07 day = 0.56 man-hrs per cum", "Labour", "0.56 man-hrs (1.79 cum/man-hr)", "1 cum"),
            ("Sundries", "Sundry tools, measuring boxes, pans, and platform upkeep", "Sundries allowance", "Equipment", "L.S. allowance", "1 cum"),
        ]
    },
    {
        "id": "3.11",
        "parent_title": "3.0 MORTARS — CEMENT MORTARS WITH COARSE SAND",
        "title": "3.11 Cement mortar 1:6 (1 cement : 6 coarse sand)",
        "unit": "cum",
        "base_qty": 1.0,
        "rows": [
            ("Portland Cement", "Economical hydraulic binder for mass brickwork and foundations (0.178 cum = 0.25 tonne)", "Standard economical mortar for foundation brickwork (REF#6.1.2), compound walls, and mass stone masonry", "Material", "0.25 tonne (0.250 t/cum)", "1 cum"),
            ("Carriage of Cement", "Site handling and carriage of cement bags to mixing yard", "Carriage of bagged cement", "Material", "0.25 tonne (0.250 t/cum)", "1 cum"),
            ("Coarse sand", "Zone III coarse sand matrix filler (1.07 cum)", "Standard coarse sand conforming to CPWD specifications", "Material", "1.07 cum (1.070 cum/cum)", "1 cum"),
            ("Carriage of Coarse sand", "Haulage and stacking of coarse sand at site platform", "Carriage of coarse sand", "Material", "1.07 cum (1.070 cum/cum)", "1 cum"),
            ("Mechanical Mortar Mixer", "Mechanical batch mixing ensuring uniform dispersion of lean cement content", "Mixer operational hire; output = 6.67 cum/hr", "Machine", "0.15 machine-hrs (6.67 cum/hr)", "1 cum"),
            ("Beldar", "Measuring sand and cement boxes, charging drum, and discharging wet mortar", "Labour allowance; 0.75 day = 6.00 man-hrs per cum", "Labour", "6.00 man-hrs (0.17 cum/man-hr)", "1 cum"),
            ("Bhisti", "Water batching, maintaining slump, and drum cleanup", "Watering allowance; 0.07 day = 0.56 man-hrs per cum", "Labour", "0.56 man-hrs (1.79 cum/man-hr)", "1 cum"),
            ("Sundries", "Hand shovels, mixing trays, farma boxes, and site cleanup", "Sundries allowance", "Equipment", "L.S. allowance", "1 cum"),
        ]
    },
    {
        "id": "3.12",
        "parent_title": "3.0 MORTARS — CEMENT MORTARS WITH STONE DUST & MARBLE DUST",
        "title": "3.12 Cement mortar 1:2 (1 cement : 2 stone dust)",
        "unit": "cum",
        "base_qty": 1.0,
        "rows": [
            ("Portland Cement", "Hydraulic binder for high-density crushed stone mortar (0.475 cum = 0.68 tonne)", "Heavy-duty wear-resistant mortar used in quarry zones where river sand is scarce; excellent interlock for stone bedding", "Material", "0.68 tonne (0.680 t/cum)", "1 cum"),
            ("Carriage of Cement", "Transport of cement bags to site batching yard", "Standard carriage of cement bags", "Material", "0.68 tonne (0.680 t/cum)", "1 cum"),
            ("Stone dust", "Crushed stone crusher screenings passing 4.75 mm with high angular friction (0.95 cum)", "Stone dust conforming to IS:383 grading; free from excess silt/dust fines", "Material", "0.95 cum (0.950 cum/cum)", "1 cum"),
            ("Carriage of Stone dust", "Mechanical haulage and unloading of stone dust at site", "Carriage of stone dust from crushing plant", "Material", "0.95 cum (0.950 cum/cum)", "1 cum"),
            ("Mechanical Mortar Mixer", "Intensive mechanical mixing required to disperse crushed stone angular fines", "Mixer hire; output = 6.67 cum/hr", "Machine", "0.15 machine-hrs (6.67 cum/hr)", "1 cum"),
            ("Beldar", "Measuring stone dust and cement, feeding drum, and discharging into pans", "Labour allowance; 0.75 day = 6.00 man-hrs per cum", "Labour", "6.00 man-hrs (0.17 cum/man-hr)", "1 cum"),
            ("Bhisti", "Water supply and mixing conditioning for angular stone dust matrix", "Watering allowance; 0.07 day = 0.56 man-hrs per cum", "Labour", "0.56 man-hrs (1.79 cum/man-hr)", "1 cum"),
            ("Sundries", "Sundry tools, farma boxes, and platform maintenance", "Sundries allowance", "Equipment", "L.S. allowance", "1 cum"),
        ]
    },
    {
        "id": "3.13",
        "parent_title": "3.0 MORTARS — CEMENT MORTARS WITH STONE DUST & MARBLE DUST",
        "title": "3.13 Cement mortar 1:2 (1 cement : 2 marble dust)",
        "unit": "cum",
        "base_qty": 1.0,
        "rows": [
            ("Portland Cement", "Hydraulic binder for dense ornamental and marble cladding bedding (0.475 cum = 0.68 tonne)", "Dense non-staining mortar used for backing and jointing marble slabs, decorative mouldings, and specialised pointing", "Material", "0.68 tonne (0.680 t/cum)", "1 cum"),
            ("Carriage of Cement", "Site handling of cement bags to mixing yard", "Carriage of bagged cement", "Material", "0.68 tonne (0.680 t/cum)", "1 cum"),
            ("Marble dust/ powder", "Finely pulverized marble powder passing 300 micron sieve (0.95 cum)", "Pure white/gray marble dust free from organic matter and clay fines", "Material", "0.95 cum (0.950 cum/cum)", "1 cum"),
            ("Carriage of Marble dust and marble chips", "Haulage and careful handling of bagged marble dust", "Carriage from marble processing units", "Material", "0.95 cum (0.950 cum/cum)", "1 cum"),
            ("Mechanical Mortar Mixer", "Mechanical mixing to prevent balling of fine marble powder", "Mixer operational hire; output = 6.67 cum/hr", "Machine", "0.15 machine-hrs (6.67 cum/hr)", "1 cum"),
            ("Beldar", "Measuring marble dust and cement, charging mixer, and discharging into clean pans", "Labour allowance; 0.75 day = 6.00 man-hrs per cum", "Labour", "6.00 man-hrs (0.17 cum/man-hr)", "1 cum"),
            ("Bhisti", "Water batching, maintaining consistency, and washing pans", "Watering allowance; 0.07 day = 0.56 man-hrs per cum", "Labour", "0.56 man-hrs (1.79 cum/man-hr)", "1 cum"),
            ("Sundries", "Clean iron pans, measuring boxes, and platform cleanup", "Sundries allowance", "Equipment", "L.S. allowance", "1 cum"),
        ]
    },
    {
        "id": "3.14",
        "parent_title": "3.0 MORTARS — CEMENT MORTARS WITH STONE DUST & MARBLE DUST",
        "title": "3.14 Cement mortar 1:5 (1 cement : 5 marble dust)",
        "unit": "cum",
        "base_qty": 1.0,
        "rows": [
            ("Portland Cement", "Hydraulic binder for decorative masonry and marble stone pointing (0.214 cum = 0.31 tonne)", "Economical marble dust mortar for internal decorative plaster finishes and architectural stone joints", "Material", "0.31 tonne (0.310 t/cum)", "1 cum"),
            ("Carriage of Cement", "Transport of cement bags to mixing point", "Standard carriage of cement bags", "Material", "0.31 tonne (0.310 t/cum)", "1 cum"),
            ("Marble dust/ powder", "Pulverized marble stone dust (1.07 cum)", "Clean marble dust conforming to CPWD architectural specifications", "Material", "1.07 cum (1.070 cum/cum)", "1 cum"),
            ("Carriage of Marble dust and marble chips", "Haulage and stacking of marble dust at site", "Carriage of marble dust", "Material", "1.07 cum (1.070 cum/cum)", "1 cum"),
            ("Mechanical Mortar Mixer", "Mechanical mixing ensuring uniform dispersion throughout marble dust volume", "Mixer operational hire; output = 6.67 cum/hr", "Machine", "0.15 machine-hrs (6.67 cum/hr)", "1 cum"),
            ("Beldar", "Measuring, drum charging, discharging, and distributing mortar", "Labour allowance; 0.75 day = 6.00 man-hrs per cum", "Labour", "6.00 man-hrs (0.17 cum/man-hr)", "1 cum"),
            ("Bhisti", "Water addition and mixer washing", "Watering allowance; 0.07 day = 0.56 man-hrs per cum", "Labour", "0.56 man-hrs (1.79 cum/man-hr)", "1 cum"),
            ("Sundries", "Sundry tools, farma boxes, and site cleanup", "Sundries allowance", "Equipment", "L.S. allowance", "1 cum"),
        ]
    },
    {
        "id": "3.15",
        "parent_title": "3.0 MORTARS — WHITE CEMENT MORTARS",
        "title": "3.15 White cement mortar 1:2 (1 white cement : 2 marble dust)",
        "unit": "cum",
        "base_qty": 1.0,
        "rows": [
            ("White Cement", "High-purity white Portland cement binder (0.475 cum = 0.68 tonne)", "Architectural mortar specified for jointing white marble flooring (REF#11.23), wall cladding (REF#8.1.1), and terrazzo chips", "Material", "0.68 tonne (0.680 t/cum)", "1 cum"),
            ("Carriage of Cement", "Careful carriage and protected storage of white cement bags to prevent discoloration", "Haulage of bagged white cement", "Material", "0.68 tonne (0.680 t/cum)", "1 cum"),
            ("Marble dust/ powder", "Pure white pulverized marble powder (0.95 cum)", "Selected pure white marble dust free from yellowing iron impurities", "Material", "0.95 cum (0.950 cum/cum)", "1 cum"),
            ("Carriage of Marble dust and marble chips", "Protected carriage and stacking of white marble dust", "Carriage of white marble dust", "Material", "0.95 cum (0.950 cum/cum)", "1 cum"),
            ("Mechanical Mortar Mixer", "Mechanical mixing in clean dedicated mixer drum to avoid gray cement contamination", "Mixer operational hire; output = 6.67 cum/hr", "Machine", "0.15 machine-hrs (6.67 cum/hr)", "1 cum"),
            ("Beldar", "Careful measuring, feeding clean drum, and discharging into clean white mortar pans", "Labour allowance; 0.75 day = 6.00 man-hrs per cum", "Labour", "6.00 man-hrs (0.17 cum/man-hr)", "1 cum"),
            ("Bhisti", "Potable clear water addition and drum washing to preserve stark white matrix", "Watering allowance; 0.07 day = 0.56 man-hrs per cum", "Labour", "0.56 man-hrs (1.79 cum/man-hr)", "1 cum"),
            ("Sundries", "Stainless/galvanized pans, clean farma boxes, and clean tools allowance", "Sundries allowance (L.S. 13.52 x cost index)", "Equipment", "L.S. allowance", "1 cum"),
        ]
    },
    {
        "id": "3.16",
        "parent_title": "3.0 MORTARS — WHITE CEMENT MORTARS",
        "title": "3.16 White cement mortar 1:3 (1 white cement : 3 marble dust)",
        "unit": "cum",
        "base_qty": 1.0,
        "rows": [
            ("White Cement", "Pure white Portland cement binder (0.357 cum = 0.51 tonne)", "Architectural mortar for pointing white marble stone joints, terrazzo skirting, and ornamental relief mouldings", "Material", "0.51 tonne (0.510 t/cum)", "1 cum"),
            ("Carriage of Cement", "Protected transport and handling of white cement bags", "Carriage of white cement", "Material", "0.51 tonne (0.510 t/cum)", "1 cum"),
            ("Marble dust/ powder", "Pure white pulverized marble powder (1.07 cum)", "White marble dust free from stains and organic fines", "Material", "1.07 cum (1.070 cum/cum)", "1 cum"),
            ("Carriage of Marble dust and marble chips", "Protected transport of marble dust", "Carriage of white marble dust", "Material", "1.07 cum (1.070 cum/cum)", "1 cum"),
            ("Mechanical Mortar Mixer", "Mechanical mixing in clean drum to ensure uniform white color tone", "Mixer operational hire; output = 6.67 cum/hr", "Machine", "0.15 machine-hrs (6.67 cum/hr)", "1 cum"),
            ("Beldar", "Measuring ingredients, feeding mixer, and discharging into clean pans", "Labour allowance; 0.75 day = 6.00 man-hrs per cum", "Labour", "6.00 man-hrs (0.17 cum/man-hr)", "1 cum"),
            ("Bhisti", "Clear water addition and cleaning mixing equipment", "Watering allowance; 0.07 day = 0.56 man-hrs per cum", "Labour", "0.56 man-hrs (1.79 cum/man-hr)", "1 cum"),
            ("Sundries", "Clean tools, measuring farmas, and contamination protection", "Sundries allowance", "Equipment", "L.S. allowance", "1 cum"),
        ]
    },
    {
        "id": "3.17",
        "parent_title": "3.0 MORTARS — WHITE CEMENT MORTARS",
        "title": "3.17 White cement mortar 1:5 (1 white cement : 5 marble dust)",
        "unit": "cum",
        "base_qty": 1.0,
        "rows": [
            ("White Cement", "Pure white Portland cement binder for decorative plaster beds (0.214 cum = 0.31 tonne)", "Economical white mortar for internal architectural stone jointing, white plaster undercoats, and light terrazzo bedding", "Material", "0.31 tonne (0.310 t/cum)", "1 cum"),
            ("Carriage of Cement", "Protected transport and handling of white cement bags", "Carriage of white cement", "Material", "0.31 tonne (0.310 t/cum)", "1 cum"),
            ("Marble dust/ powder", "Pure white pulverized marble powder (1.07 cum)", "White marble dust conforming to CPWD specifications", "Material", "1.07 cum (1.070 cum/cum)", "1 cum"),
            ("Carriage of Marble dust and marble chips", "Carriage and protected stacking of marble dust", "Carriage of marble dust", "Material", "1.07 cum (1.070 cum/cum)", "1 cum"),
            ("Mechanical Mortar Mixer", "Thorough mechanical mixing ensuring uniform white color dispersion", "Mixer operational hire; output = 6.67 cum/hr", "Machine", "0.15 machine-hrs (6.67 cum/hr)", "1 cum"),
            ("Beldar", "Measuring, charging drum, discharging, and distributing mixed mortar", "Labour allowance; 0.75 day = 6.00 man-hrs per cum", "Labour", "6.00 man-hrs (0.17 cum/man-hr)", "1 cum"),
            ("Bhisti", "Water addition and drum cleaning", "Watering allowance; 0.07 day = 0.56 man-hrs per cum", "Labour", "0.56 man-hrs (1.79 cum/man-hr)", "1 cum"),
            ("Sundries", "Clean tools, farmas, and contamination prevention", "Sundries allowance", "Equipment", "L.S. allowance", "1 cum"),
        ]
    },
    {
        "id": "3.18",
        "parent_title": "3.0 MORTARS — MUD MORTAR",
        "title": "3.18 Mud mortar (clay mortar)",
        "unit": "cum",
        "base_qty": 1.0,
        "rows": [
            ("Mud (dry)", "Selected dry clayey earth/loam excavated locally, pulverised and sieved to remove stones and grass roots (1.08 cum)", "Traditional ecological and economical binder used for temporary masonry, rural brickwork, and adobe structures", "Material", "1.08 cum (1.080 cum/cum)", "1 cum"),
            ("Beldar", "Excavating clay, pulverising clods, sieving, puddling earth with feet/spades, and kneading into uniform plastic mortar", "Manual earth processing, puddling, kneading, and handling; 0.63 day = 5.04 man-hrs per cum", "Labour", "5.04 man-hrs (0.20 cum/man-hr)", "1 cum"),
            ("Bhisti", "Watering clay in pit, soaking earth overnight, adding water during puddling and kneading to attain soft plastic consistency", "Thorough soaking, moistening, and water supply; 0.315 day = 2.52 man-hrs per cum", "Labour", "2.52 man-hrs (0.40 cum/man-hr)", "1 cum"),
            ("Sundries", "Wooden puddling spades, sieves, shallow mixing pits, and mortar distribution baskets", "Sundries and tools allowance (L.S. 6.45 x cost index)", "Equipment", "L.S. allowance", "1 cum"),
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
        c2.font = Font(name="Calibri", size=9, bold=(category in ("Machine", "Equipment", "Material")))
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

def build_mortars_sheet(ws):
    # Configure fresh worksheet
    if ws.views.sheetView:
        ws.views.sheetView[0].showGridLines = True

    # Sheet title spanning A1:G1
    ws.merge_cells("A1:G1")
    t1 = ws.cell(row=1, column=1, value="Sub-Head 3.0 — MORTARS  |  First-Principles Resource, Work & Gang Analysis")
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
    for item in MORTAR_ITEMS:
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

def rebuild_mortars_in_workbook(file_path):
    print(f"Opening {file_path} ...")
    wb = openpyxl.load_workbook(file_path)
    if "03_Mortars" not in wb.sheetnames:
        raise ValueError(f"'03_Mortars' sheet not found in {file_path}")

    pos = wb.sheetnames.index("03_Mortars")
    del wb["03_Mortars"]
    ws = wb.create_sheet("03_Mortars", pos)
    ws.sheet_properties.tabColor = "2E75B6"
    build_mortars_sheet(ws)

    temp_path = file_path.replace(".xlsx", "_TMP_MORTARS.xlsx")
    wb.save(temp_path)
    wb.close()
    os.replace(temp_path, file_path)
    print(f"Successfully updated {file_path} -> sheet '03_Mortars' ({len(MORTAR_ITEMS)} items).")

def main():
    repo_root = Path(__file__).resolve().parents[1]
    main_wb = repo_root / "CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_1_Latest.xlsx"
    output_wb = repo_root / "outputs" / "earthwork-custom-rate-composer" / "CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_1_Latest.xlsx"

    if main_wb.exists():
        rebuild_mortars_in_workbook(str(main_wb))
    if output_wb.exists():
        rebuild_mortars_in_workbook(str(output_wb))

if __name__ == "__main__":
    main()
