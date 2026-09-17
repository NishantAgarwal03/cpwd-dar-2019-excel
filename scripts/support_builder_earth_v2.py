"""
support_builder_earth_v2.py
Rebuilds 02_support_earth_work sheet with First-Principles Teaching Format:
  Item Code | Labour / Machine / Material | Work done | Condition / When used | Category | Productivity | Quantity
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from pathlib import Path
import os, sys, shutil, json, re
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir)))
from scripts.paths import WB_VOL1_LATEST_FILE
import scripts.earthwork_teaching_knowledge as earthwork_teaching_knowledge

# Colours
C_ITEM_HDR   = "1F4E79"   # dark blue
C_SEC_HDR    = "2E75B6"   # mid blue
C_WHITE      = "FFFFFF"

def fill(hex_col):
    if hex_col is None:
        return PatternFill(fill_type=None)
    return PatternFill(fill_type="solid", fgColor=hex_col)

def thin_border():
    s = Side(style="thin", color="AAAAAA")
    return Border(left=s, right=s, top=s, bottom=s)

AL_C = Alignment(horizontal="center", vertical="center")
AL_L = Alignment(horizontal="left",   vertical="center", wrap_text=True)
AL_R = Alignment(horizontal="right",  vertical="center")

_SPECS_PATH = Path(__file__).resolve().parents[1] / "data" / "reference_json" / "earthwork_pdf_specs.json"
if _SPECS_PATH.exists():
    with open(_SPECS_PATH, encoding="utf-8") as _f:
        PDF_SPECS = json.load(_f)
else:
    PDF_SPECS = {}

def get_item_headers(code: str) -> list[str]:
    parts = code.split(".")
    standalone = [
        "2.4", "2.5", "2.11", "2.12", "2.14", "2.15", "2.19", "2.23",
        "2.25", "2.25(a)", "2.27", "2.31", "2.32", "2.36", "2.37", "2.38"
    ]
    if code in standalone:
        full_desc = PDF_SPECS.get(code, "")
        return [f"{code} {full_desc}".strip()]
    if len(parts) == 3:
        p_code = f"{parts[0]}.{parts[1]}"
        p_desc = PDF_SPECS.get(p_code, "")
        c_desc = PDF_SPECS.get(code, "")
        lines = []
        if p_desc:
            lines.append(f"{p_code} {p_desc}".strip())
        lines.append(f"{code} {c_desc}".strip() if c_desc else code)
        return lines
    if len(parts) >= 4:
        p2_code = f"{parts[0]}.{parts[1]}"
        p3_code = f"{parts[0]}.{parts[1]}.{parts[2]}"
        p2_desc = PDF_SPECS.get(p2_code, "")
        p3_desc = PDF_SPECS.get(p3_code, "")
        c_desc = PDF_SPECS.get(code, "")
        lines = []
        if p2_desc:
            lines.append(f"{p2_code} {p2_desc}".strip())
        sub_desc = f"{p3_desc} - {c_desc}".strip(" -")
        lines.append(f"{code} {sub_desc}".strip())
        return lines
    return [f"{code} {PDF_SPECS.get(code, '')}".strip()]


ITEMS = [

# ─── 2.1 SURFACE DRESSING ────────────────────────────────────────────────────
{
 "id":"2.1.1","desc":"Earth work in surface excavation not exceeding 30 cm in depth but exceeding 1.5 m in width as well as 10 sqm on plan including getting out and disposal of excavated earth upto 50 m and lift upto 1.5 m, as directed by Engineer-in- Charge: All kinds of soil",
 "unit":"sqm","base_qty":100,"dsr_rate":92.55,
 "sections":[
  ("LABOUR",[
   ("0114","Beldar","day",6.80,558.00),
   ("0115","Coolie","day",5.60,558.00),
  ]),
 ],
},
# UNVERIFIABLE (investigated 2026-09-18): this "Hard soil" variant does not exist in
# CivilDAR_2019_Vol_1.pdf (searched all 904 pages, zero "hard soil" hits) nor in
# DSR_Vol1_UPDATED_DEC_2021.pdf / DSR_Vol2_UPDATED_DEC_2021.pdf. dsr_rate and the
# resource coefficients below cannot be traced to any source document in this repo.
# Kept as-is per explicit user decision; do not treat as PDF-verified.
{
 "id":"2.1.2","desc":"Earth work in surface excavation not exceeding 30 cm in depth but exceeding 1.5 m in width as well as 10 sqm on plan including getting out and disposal of excavated earth upto 50 m and lift upto 1.5 m, as directed by Engineer-in- Charge:",
 "unit":"sqm","base_qty":100,"dsr_rate":134.55,
 "sections":[
  ("LABOUR",[
   ("0114","Beldar","day",8.60,558.00),
   ("0115","Coolie","day",7.10,558.00),
  ]),
 ],
},

# ─── 2.2 EARTHWORK IN EXCAVATION (ORDINARY SOIL, LEAD ≤50m) ─────────────────
{
 "id":"2.2.1","desc":"Earth work in rough excavation, banking excavated earth in layers not exceeding 20cm in depth, breaking clods, watering, rolling each layer with ½ tonne roller or wooden or steel rammers, and rolling every 3rd and top-most layer with power roller of minimum 8 tonnes and dressing up in embankments for roads, flood banks, marginal banks and guide banks or filling up ground depressions, lead upto 50 m and lift upto 1.5 m : All kinds of soil",
 "unit":"cum","base_qty":10,"dsr_rate":746.80,
 "sections":[
  ("MACHINERY",[
   ("0003","Diesel Road Roller 8-10t (for compaction)","day",0.008,3000.00),
  ]),
  ("LABOUR",[
   ("0114","Beldar","day",5.90,558.00),
   ("0115","Coolie","day",3.60,558.00),
   ("0101","Bhisti","day",0.40,617.00),
   ("0113","Chowkidar","day",0.008,558.00),
  ]),
  ("MATERIAL",[
   ("9999","Sundries","LS",2.73,2.00),
  ]),
 ],
},
# UNVERIFIABLE Hard-soil variant - see note above 2.1.2. Kept as-is per user decision.
{
 "id":"2.2.2","desc":"Earth work in rough excavation, banking excavated earth in layers not exceeding 20cm in depth, breaking clods, watering, rolling each layer with ½ tonne roller or wooden or steel rammers, and rolling every 3rd and top-most layer with power roller of minimum 8 tonnes and dressing up in embankments for roads, flood banks, marginal banks and guide banks or filling up ground depressions, lead upto 50 m and lift upto 1.5 m : Hard soil",
 "unit":"cum","base_qty":10,"dsr_rate":1049.65,
 "sections":[
  ("MACHINERY",[
   ("0003","Diesel Road Roller 8-10t","day",0.008,3000.00),
  ]),
  ("LABOUR",[
   ("0114","Beldar","day",7.50,558.00),
   ("0115","Coolie","day",4.50,558.00),
   ("0101","Bhisti","day",0.40,617.00),
   ("0113","Chowkidar","day",0.008,558.00),
  ]),
  ("MATERIAL",[
   ("9999","Sundries","LS",3.64,2.00),
  ]),
 ],
},

# ─── 2.3 EARTHWORK IN EXCAVATION (ORDINARY SOIL, LEAD ≤50m, shallow) ────────
{
 "id":"2.3.1","desc":"Banking excavated earth in layers not exceeding 20 cm in depth, breaking clods, watering, rolling each layer with ½ tonne roller, or wooden or steel rammers, and rolling every 3rd and top-most layer with power roller of minimum 8 tonnes and dressing up, in embankments for roads, flood banks, marginal banks, and guide banks etc., lead upto 50 m and lift upto 1.5 m : All kinds of soil",
 "unit":"cum","base_qty":10,"dsr_rate":470.55,
 "sections":[
  ("MACHINERY",[
   ("0003","Diesel Road Roller 8-10t","day",0.008,3000.00),
  ]),
  ("LABOUR",[
   ("0114","Beldar","day",2.20,558.00),
   ("0115","Coolie","day",3.60,558.00),
   ("0101","Bhisti","day",0.40,617.00),
   ("0113","Chowkidar","day",0.008,558.00),
  ]),
  ("MATERIAL",[
   ("9999","Sundries","LS",2.73,2.00),
  ]),
 ],
},
# UNVERIFIABLE Hard-soil variant - see note above 2.1.2. Kept as-is per user decision.
{
 "id":"2.3.2","desc":"Banking excavated earth in layers not exceeding 20 cm in depth, breaking clods, watering, rolling each layer with ½ tonne roller, or wooden or steel rammers, and rolling every 3rd and top-most layer with power roller of minimum 8 tonnes and dressing up, in embankments for roads, flood banks, marginal banks, and guide banks etc., lead upto 50 m and lift upto 1.5 m : Hard soil",
 "unit":"cum","base_qty":10,"dsr_rate":661.60,
 "sections":[
  ("MACHINERY",[
   ("0003","Diesel Road Roller 8-10t","day",0.008,3000.00),
  ]),
  ("LABOUR",[
   ("0114","Beldar","day",3.20,558.00),
   ("0115","Coolie","day",4.50,558.00),
   ("0101","Bhisti","day",0.40,617.00),
   ("0113","Chowkidar","day",0.008,558.00),
  ]),
  ("MATERIAL",[
   ("9999","Sundries","LS",3.64,2.00),
  ]),
 ],
},

# ─── 2.4 DEDUCT – Roller & Chowkidar ─────────────────────────────────────────
{
 "id":"2.4","desc":"Deduct for not rolling with power roller of minimum 8 tonnes for banking excavated earth in layers not exceeding 20 cm in depth.",
 "unit":"cum","base_qty":10,"dsr_rate":4.30,
 "sections":[
  ("MACHINERY",[
   ("0003","Diesel Road Roller 8-10t","day",0.008,3000.00),
  ]),
  ("LABOUR",[
   ("0113","Chowkidar","day",0.008,558.00),
  ]),
  ("MATERIAL",[
   ("9999","Sundries","LS",1.82,2.00),
  ]),
 ],
 "notes":"DEDUCT item",
},

# ─── 2.5 DEDUCT – Bhisti ─────────────────────────────────────────────────────
{
 "id":"2.5","desc":"Deduct for not watering the excavated earth for banking",
 "unit":"cum","base_qty":10,"dsr_rate":33.00,
 "sections":[
  ("LABOUR",[
   ("0101","Bhisti","day",0.40,617.00),
  ]),
 ],
 "notes":"DEDUCT item",
},

# ─── 2.6 HYDRAULIC EXCAVATION ─────────────────────────────────────────────────
{
 "id":"2.6.1","desc":"Earth work in excavation by mechanical means (Hydraulic excavator)/manual means over areas (exceeding 30 cm in depth, 1.5 m in width as well as 10 sqm on plan) including getting out and disposal of excavated earth lead upto 50 m and lift upto 1.5 m, as directed by Engineer-in-charge. - All kinds of soil",
 "unit":"cum","base_qty":10,"dsr_rate":181.85,
 "sections":[
  ("MACHINERY",[
   ("0020","Hydraulic Excavator 0.9 cum","day",0.041,7000.00),
   ("0018","Loader","day",0.041,5000.00),
  ]),
  ("LABOUR",[
   ("0128","Mate","day",0.32,617.00),
   ("0115","Coolie","day",1.20,558.00),
  ]),
 ],
},
# UNVERIFIABLE Hard-soil variant - see note above 2.1.2. Kept as-is per user decision.
{
 "id":"2.6.2","desc":"Earth work in excavation by mechanical means (Hydraulic excavator)/manual means over areas (exceeding 30 cm in depth, 1.5 m in width as well as 10 sqm on plan) including getting out and disposal of excavated earth lead upto 50 m and lift upto 1.5 m, as directed by Engineer-in-charge. - Hard soil",
 "unit":"cum","base_qty":10,"dsr_rate":261.10,
 "sections":[
  ("MACHINERY",[
   ("0020","Hydraulic Excavator 0.9 cum","day",0.063,7000.00),
   ("0018","Loader","day",0.063,5000.00),
  ]),
  ("LABOUR",[
   ("0128","Mate","day",0.48,617.00),
   ("0115","Coolie","day",1.80,558.00),
  ]),
 ],
},

# ─── 2.7 ROCK EXCAVATION ─────────────────────────────────────────────────────
{
 "id":"2.7.1","desc":"Earth work in excavation by mechanical means (Hydraulic excavator)/manual means over areas (exceeding 30 cm in depth, 1.5 m in width as well as 10 sqm on plan) including getting out and disposal of excavated earth lead upto 50 m and lift upto 1.5 m, as directed by Engineer-in-charge. - Ordinary rock",
 "unit":"cum","base_qty":10,"dsr_rate":352.45,
 "sections":[
  ("MACHINERY",[
   ("0020","Hydraulic Excavator 0.9 cum","day",0.063,7000.00),
   ("0017","Tipper 6 cum","day",0.063,1700.00),
  ]),
  ("LABOUR",[
   ("0132","Rock Excavator","day",0.705,558.00),
   ("0133","Rock Breaker","day",1.590,558.00),
   ("0134","Rock Hole Driller","day",0.355,558.00),
   ("0114","Beldar","day",0.50,558.00),
   ("0115","Coolie","day",0.55,558.00),
  ]),
  ("MATERIAL",[
   ("9999","Sundries (blasting materials)","LS",10.79,2.00),
  ]),
 ],
},
{
 "id":"2.7.2","desc":"Earth work in excavation by mechanical means (Hydraulic excavator)/manual means over areas (exceeding 30 cm in depth, 1.5 m in width as well as 10 sqm on plan) including getting out and disposal of excavated earth lead upto 50 m and lift upto 1.5 m, as directed by Engineer-in-charge. - Hard rock (requiring blasting)",
 "unit":"cum","base_qty":10,"dsr_rate":609.65,
 "sections":[
  ("MACHINERY",[
   ("0020","Hydraulic Excavator 0.9 cum","day",0.125,7000.00),
   ("0017","Tipper 6 cum","day",0.125,1700.00),
  ]),
  ("LABOUR",[
   ("0132","Rock Excavator","day",1.060,558.00),
   ("0133","Rock Breaker","day",2.825,558.00),
   ("0134","Rock Hole Driller","day",0.885,558.00),
   ("0114","Beldar","day",0.45,558.00),
   ("0115","Coolie","day",0.55,558.00),
  ]),
  ("MATERIAL",[
   ("","Blasting powder","kg",3.93,40.00),
   ("","Fuse","metre",4.00,15.00),
   ("9999","Sundries","LS",16.12,2.00),
  ]),
 ],
},
{
 "id":"2.7.3","desc":"Earth work in excavation by mechanical means (Hydraulic excavator)/manual means over areas (exceeding 30 cm in depth, 1.5 m in width as well as 10 sqm on plan) including getting out and disposal of excavated earth lead upto 50 m and lift upto 1.5 m, as directed by Engineer-in-charge. - Hard rock (blasting prohibited)",
 "unit":"cum","base_qty":10,"dsr_rate":1016.20,
 "sections":[
  ("MACHINERY",[
   ("0020","Hydraulic Excavator 0.9 cum","day",0.125,7000.00),
   ("0017","Tipper 6 cum","day",0.125,1700.00),
  ]),
  ("LABOUR",[
   ("0132","Rock Excavator","day",2.470,558.00),
   ("0133","Rock Breaker","day",6.000,558.00),
   ("0135","Stone Chiseller","day",1.060,617.00),
   ("0103","Blacksmith 2nd class","day",0.175,679.00),
   ("0114","Beldar","day",0.75,558.00),
   ("0115","Coolie","day",1.00,558.00),
  ]),
  ("MATERIAL",[
   ("9999","Sundries","LS",16.12,2.00),
  ]),
 ],
},

# ─── 2.8 EARTHWORK IN EXCAVATION FOR PIPE TRENCHES ───────────────────────────
{
 "id":"2.8.1","desc":"Earth work in excavation by mechanical means (Hydraulic excavator) / manual means in foundation trenches or drains (not exceeding 1.5 m in width or 10 sqm on plan), including dressing of sides and ramming of bottoms, lift upto 1.5 m, including getting out the excavated soil and disposal of surplus excavated soil as directed, within a lead of 50 m. - All kinds of soil.",
 "unit":"cum","base_qty":10,"dsr_rate":252.30,
 "sections":[
  ("MACHINERY",[
   ("0020","Hydraulic Excavator 0.9 cum","day",0.04125,7000.00),
   ("0018","Loader","day",0.04125,5000.00),
  ]),
  ("LABOUR",[
   ("0128","Mate","day",0.40,617.00),
   ("0115","Coolie","day",2.05,558.00),
  ]),
 ],
},
# UNVERIFIABLE Hard-soil variant - see note above 2.1.2. Kept as-is per user decision.
{
 "id":"2.8.2","desc":"Earth work in excavation by mechanical means (Hydraulic excavator) / manual means in foundation trenches or drains (not exceeding 1.5 m in width or 10 sqm on plan), including dressing of sides and ramming of bottoms, lift upto 1.5 m, including getting out the excavated soil and disposal of surplus excavated soil as directed, within a lead of 50 m. - Hard soil",
 "unit":"cum","base_qty":10,"dsr_rate":347.10,
 "sections":[
  ("MACHINERY",[
   ("0020","Hydraulic Excavator 0.9 cum","day",0.063,7000.00),
   ("0018","Loader","day",0.063,5000.00),
  ]),
  ("LABOUR",[
   ("0128","Mate","day",0.60,617.00),
   ("0115","Coolie","day",3.00,558.00),
  ]),
 ],
},

# ─── 2.9 PIPE TRENCHES IN ROCK ───────────────────────────────────────────────
{
 "id":"2.9.1","desc":"Excavation work by mechanical means (Hydraulic excavator)/ manual means in foundation trenches or drains (not exceeding 1.5m in width or 10 sqm on plan), including dressing of sides and ramming of bottoms, lift upto 1.5 m, including getting out the excavated soil and disposal of surplus excavated soils as directed, within a lead of 50 m. - Ordinary rock",
 "unit":"cum","base_qty":10,"dsr_rate":448.15,
 "sections":[
  ("MACHINERY",[
   ("0020","Hydraulic Excavator 0.9 cum","day",0.0625,7000.00),
   ("0017","Tipper 6 cum","day",0.0625,1700.00),
  ]),
  ("LABOUR",[
   ("0132","Rock Excavator","day",0.885,558.00),
   ("0133","Rock Breaker","day",1.765,558.00),
   ("0134","Rock Hole Driller","day",0.530,558.00),
   ("0114","Beldar","day",0.50,558.00),
   ("0115","Coolie","day",1.30,558.00),
  ]),
  ("MATERIAL",[
   ("9999","Sundries","LS",13.52,2.00),
  ]),
 ],
},
{
 "id":"2.9.2","desc":"Excavation work by mechanical means (Hydraulic excavator)/ manual means in foundation trenches or drains (not exceeding 1.5m in width or 10 sqm on plan), including dressing of sides and ramming of bottoms, lift upto 1.5 m, including getting out the excavated soil and disposal of surplus excavated soils as directed, within a lead of 50 m. - Hard rock (requiring blasting)",
 "unit":"cum","base_qty":10,"dsr_rate":729.00,
 "sections":[
  ("MACHINERY",[
   ("0020","Hydraulic Excavator 0.9 cum","day",0.125,7000.00),
   ("0017","Tipper 6 cum","day",0.125,1700.00),
  ]),
  ("LABOUR",[
   ("0132","Rock Excavator","day",1.240,558.00),
   ("0133","Rock Breaker","day",3.000,558.00),
   ("0134","Rock Hole Driller","day",1.060,558.00),
   ("0114","Beldar","day",0.50,558.00),
   ("0115","Coolie","day",1.30,558.00),
  ]),
  ("MATERIAL",[
   ("","Blasting powder","kg",6.42,40.00),
   ("","Fuse","metre",7.00,15.00),
   ("9999","Sundries","LS",18.85,2.00),
  ]),
 ],
},
{
 "id":"2.9.3","desc":"Excavation work by mechanical means (Hydraulic excavator)/ manual means in foundation trenches or drains (not exceeding 1.5m in width or 10 sqm on plan), including dressing of sides and ramming of bottoms, lift upto 1.5 m, including getting out the excavated soil and disposal of surplus excavated soils as directed, within a lead of 50 m. - Hard rock (blasting prohibited)",
 "unit":"cum","base_qty":10,"dsr_rate":624.70,
 "sections":[
  ("MACHINERY",[
   ("0020","Hydraulic Excavator 0.9 cum","day",0.125,7000.00),
   ("0017","Tipper 6 cum","day",0.125,1700.00),
  ]),
  ("LABOUR",[
   ("0132","Rock Excavator","day",2.650,558.00),
   ("0133","Rock Breaker","day",6.175,558.00),
   ("0135","Stone Chiseller","day",1.060,617.00),
   ("0103","Blacksmith 2nd class","day",0.175,679.00),
   ("0114","Beldar","day",0.75,558.00),
   ("0115","Coolie","day",1.50,558.00),
  ]),
  ("MATERIAL",[
   ("9999","Sundries","LS",17.94,2.00),
  ]),
 ],
},

# ─── 2.10 PIPE TRENCHES COMPOSITE (EXCAV + FILLING) ─────────────────────────
# W-A pattern: REF sub-items are excluded from Water/GST/CPOH/Cess chain
{
 "id":"2.10.1.1","desc":"Excavating trenches of required width for pipes, cables, etc including excavation for sockets, and dressing of sides, ramming of bottoms, depth upto 1.5 m, including getting out the excavated soil, and then returning the soil as required, in layers not exceeding 20 cm in depth, including consolidating each deposited layer by ramming, watering, etc. and disposing of surplus excavated soil as directed, within a lead of 50 m : - All kinds of soil - Pipes, cables etc, not exceeding 80 mm dia.",
 "unit":"metre","base_qty":180,"dsr_rate":223.00,
 "ref_pattern":True,
 "sections":[
  ("REFERENCE ITEMS (excluded from multiplier chain)",[
   ("REF#2.8.1","Excavation pipe trench ordinary soil (85.05 cum)","cum",85.05,252.30),
   ("REF#2.25","Filling in trenches (85.05 cum)","cum",85.05,219.65),
  ]),
 ],
 "notes":"W-A pattern: entire cost = sum of referenced items × base qty. No separate Water/GST/CPOH applied on top.",
},
{
 "id":"2.10.1.2","desc":"Excavating trenches of required width for pipes, cables, etc including excavation for sockets, and dressing of sides, ramming of bottoms, depth upto 1.5 m, including getting out the excavated soil, and then returning the soil as required, in layers not exceeding 20 cm in depth, including consolidating each deposited layer by ramming, watering, etc. and disposing of surplus excavated soil as directed, within a lead of 50 m : - All kinds of soil - Pipes, cables etc. exceeding 80 mm dia. but not exceeding 300 mm dia",
 "unit":"metre","base_qty":110,"dsr_rate":315.05,
 "ref_pattern":True,
 "sections":[
  ("REFERENCE ITEMS",[
   ("REF#2.8.1","Excavation pipe trench ordinary soil (84.89 cum)","cum",84.89,252.30),
   ("REF#2.25","Filling in trenches (84.89 cum)","cum",84.89,219.65),
  ]),
 ],
},
{
 "id":"2.10.1.3","desc":"Excavating trenches of required width for pipes, cables, etc including excavation for sockets, and dressing of sides, ramming of bottoms, depth upto 1.5 m, including getting out the excavated soil, and then returning the soil as required, in layers not exceeding 20 cm in depth, including consolidating each deposited layer by ramming, watering, etc. and disposing of surplus excavated soil as directed, within a lead of 50 m : - All kinds of soil - Pipes, cables etc. exceeding 300 mm dia but not exceeding 600 mm",
 "unit":"metre","base_qty":60,"dsr_rate":568.60,
 "ref_pattern":True,
 "sections":[
  ("REFERENCE ITEMS",[
   ("REF#2.8.1","Excavation pipe trench ordinary soil (72.29 cum)","cum",72.29,252.30),
   ("REF#2.25","Filling in trenches (72.29 cum)","cum",72.29,219.65),
  ]),
 ],
},

# ─── 2.11 EXTRA DEPTH (%) ─────────────────────────────────────────────────────
{
 "id":"2.11","desc":"Extra for excavating trenches for pipes, cables etc. in all kinds of soil for depth exceeding 1.5 m, but not exceeding 3 m. (Rate is over corresponding basic item for depth upto 1.5 metre).",
 "unit":"%","base_qty":1,"dsr_rate":127.00,
 "pct_item":True,
 "sections":[],
 "notes":"Percentage-based extra: 127% over base rate of Item 2.10. No resource rows. DAR computed value = 127.00%.",
},
{
 "id":"2.12","desc":"Extra for excavating trenches for pipes, cables, etc, in all kinds of soil for depth exceeding 3 m in depth, but not exceeding 4.5 m. (Rate is over corresponding basic item for depth upto 1.5 metre.)",
 "unit":"%","base_qty":1,"dsr_rate":314.95,
 "pct_item":True,
 "sections":[],
 "notes":"Percentage-based extra: 315.05% over base rate of Item 2.10. DAR computed value = 315.05% (DSR 2021: 314.95%).",
},

# ─── 2.13 ROCK PIPE TRENCHES (COMPOSITE) ─────────────────────────────────────
{
 "id":"2.13.1.1","desc":"Excavating trenches of required width for pipes, cables, etc, including excavation for sockets, depth upto 1.5 m, including getting out the excavated materials, returning the soil as required in layers not exceeding 20 cm in depth, including consolidating each deposited layers by ramming, watering etc., stacking serviceable material for measurements and disposal of unserviceable material as directed, within a lead of 50 m : - Ordinary rock : - Pipes, cables etc. not exceeding 80 mm dia",
 "unit":"metre","base_qty":180,"dsr_rate":323.85,
 "ref_pattern":True,
 "sections":[
  ("REFERENCE ITEMS (A - excluded from multipliers)",[
   ("REF#2.9.1","Excavation ordinary rock (85.05 cum)","cum",85.05,448.15),
   ("REF#2.25","Filling in trenches (85.05 cum)","cum",85.05,219.65),
  ]),
  ("LABOUR (subject to Water/GST/CPOH/Cess)",[
   ("0114","Beldar (extra for rock dressing)","day",1.80,558.00),
  ]),
 ],
},
{
 "id":"2.13.1.2","desc":"Excavating trenches of required width for pipes, cables, etc, including excavation for sockets, depth upto 1.5 m, including getting out the excavated materials, returning the soil as required in layers not exceeding 20 cm in depth, including consolidating each deposited layers by ramming, watering etc., stacking serviceable material for measurements and disposal of unserviceable material as directed, within a lead of 50 m : - Ordinary rock : - Pipes, cables etc. exceeding 80 mm dia but not exceeding 300 mm dia",
 "unit":"metre","base_qty":80,"dsr_rate":801.85,
 "ref_pattern":True,
 "sections":[
  ("REFERENCE ITEMS (A)",[
   ("REF#2.9.1","Excavation ordinary rock (93.60 cum)","cum",93.60,448.15),
   ("REF#2.25","Filling in trenches (93.60 cum)","cum",93.60,219.65),
  ]),
  ("LABOUR",[
   ("0114","Beldar (extra dressing)","day",1.80,558.00),
  ]),
 ],
},
{
 "id":"2.13.1.3","desc":"Excavating trenches of required width for pipes, cables, etc, including excavation for sockets, depth upto 1.5 m, including getting out the excavated materials, returning the soil as required in layers not exceeding 20 cm in depth, including consolidating each deposited layers by ramming, watering etc., stacking serviceable material for measurements and disposal of unserviceable material as directed, within a lead of 50 m : - Ordinary rock : - Pipes, cables exceeding 300 mm dia but not exceeding 600 mm dia",
 "unit":"metre","base_qty":30,"dsr_rate":922.65,
 "ref_pattern":True,
 "sections":[
  ("REFERENCE ITEMS (A)",[
   ("REF#2.9.1","Excavation ordinary rock (40.399 cum)","cum",40.399,448.15),
   ("REF#2.25","Filling in trenches (40.399 cum)","cum",40.399,219.65),
  ]),
  ("LABOUR",[
   ("0114","Beldar (extra dressing)","day",0.94,558.00),
  ]),
 ],
},
{
 "id":"2.13.2.1","desc":"Excavating trenches of required width for pipes, cables, etc, including excavation for sockets, depth upto 1.5 m, including getting out the excavated materials, returning the soil as required in layers not exceeding 20 cm in depth, including consolidating each deposited layers by ramming, watering etc., stacking serviceable material for measurements and disposal of unserviceable material as directed, within a lead of 50 m : - Hard rock (requiring blasting) - Pipes, cables etc. not exceeding 80 mm dia",
 "unit":"metre","base_qty":180,"dsr_rate":458.60,
 "ref_pattern":True,
 "sections":[
  ("REFERENCE ITEMS (A)",[
   ("REF#2.9.2","Excavation hard rock blasting (85.05 cum)","cum",85.05,729.00),
   ("REF#2.25","Filling in trenches (85.05 cum)","cum",85.05,219.65),
  ]),
  ("LABOUR",[
   ("0114","Beldar (rock ramming)","day",1.80,558.00),
  ]),
 ],
},
{
 "id":"2.13.2.2","desc":"Excavating trenches of required width for pipes, cables, etc, including excavation for sockets, depth upto 1.5 m, including getting out the excavated materials, returning the soil as required in layers not exceeding 20 cm in depth, including consolidating each deposited layers by ramming, watering etc., stacking serviceable material for measurements and disposal of unserviceable material as directed, within a lead of 50 m : - Hard rock (requiring blasting) - Pipes, cables etc. exceeding 80 mm dia but not exceeding 300 mm dia",
 "unit":"metre","base_qty":80,"dsr_rate":255.60,
 "ref_pattern":True,
 "sections":[
  ("REFERENCE ITEMS (A)",[
   ("REF#2.9.2","Excavation hard rock blasting (93.60 cum)","cum",93.60,729.00),
   ("REF#2.25","Filling in trenches (93.60 cum)","cum",93.60,219.65),
  ]),
  ("LABOUR",[
   ("0114","Beldar (rock ramming)","day",1.80,558.00),
  ]),
 ],
},
{
 "id":"2.13.2.3","desc":"Excavating trenches of required width for pipes, cables, etc, including excavation for sockets, depth upto 1.5 m, including getting out the excavated materials, returning the soil as required in layers not exceeding 20 cm in depth, including consolidating each deposited layers by ramming, watering etc., stacking serviceable material for measurements and disposal of unserviceable material as directed, within a lead of 50 m : - Hard rock (requiring blasting) - Pipes, cables etc. exceeding 300 mm dia but not exceeding 600 mm dia",
 "unit":"metre","base_qty":30,"dsr_rate":1306.60,
 "ref_pattern":True,
 "sections":[
  ("REFERENCE ITEMS (A)",[
   ("REF#2.9.2","Excavation hard rock blasting (40.399 cum)","cum",40.399,729.00),
   ("REF#2.25","Filling in trenches (40.399 cum)","cum",40.399,219.65),
  ]),
  ("LABOUR",[
   ("0114","Beldar (rock ramming)","day",1.17,558.00),
  ]),
 ],
},
{
 "id":"2.13.3.1","desc":"Excavating trenches of required width for pipes, cables, etc, including excavation for sockets, depth upto 1.5 m, including getting out the excavated materials, returning the soil as required in layers not exceeding 20 cm in depth, including consolidating each deposited layers by ramming, watering etc., stacking serviceable material for measurements and disposal of unserviceable material as directed, within a lead of 50 m : - Hard rock (blasting prohibited) - Pipes, cables etc. not exceeding 80 mm dia",
 "unit":"metre","base_qty":180,"dsr_rate":726.65,
 "ref_pattern":True,
 "sections":[
  ("REFERENCE ITEMS (A)",[
   ("REF#2.9.3","Excavation hard rock no blasting (85.05 cum)","cum",85.05,1080.55),
   ("REF#2.25","Filling in trenches (85.05 cum)","cum",85.05,219.65),
  ]),
  ("LABOUR",[
   ("0114","Beldar (rock ramming)","day",2.50,558.00),
  ]),
 ],
},
{
 "id":"2.13.3.2","desc":"Excavating trenches of required width for pipes, cables, etc, including excavation for sockets, depth upto 1.5 m, including getting out the excavated materials, returning the soil as required in layers not exceeding 20 cm in depth, including consolidating each deposited layers by ramming, watering etc., stacking serviceable material for measurements and disposal of unserviceable material as directed, within a lead of 50 m : - Hard rock (blasting prohibited) - Pipes, cables etc. exceeding 80 mm dia but not exceeding 300 mm dia",
 "unit":"metre","base_qty":80,"dsr_rate":1546.90,
 "ref_pattern":True,
 "sections":[
  ("REFERENCE ITEMS (A)",[
   ("REF#2.9.3","Excavation hard rock no blasting (93.60 cum)","cum",93.60,1080.55),
   ("REF#2.25","Filling in trenches (93.60 cum)","cum",93.60,219.65),
  ]),
  ("LABOUR",[
   ("0114","Beldar (rock ramming)","day",2.50,558.00),
  ]),
 ],
},
{
 "id":"2.13.3.3","desc":"Excavating trenches of required width for pipes, cables, etc, including excavation for sockets, depth upto 1.5 m, including getting out the excavated materials, returning the soil as required in layers not exceeding 20 cm in depth, including consolidating each deposited layers by ramming, watering etc., stacking serviceable material for measurements and disposal of unserviceable material as directed, within a lead of 50 m : - Hard rock (blasting prohibited) - Pipes, cables etc. exceeding 300 mm dia but not exceeding 600 mm dia",
 "unit":"metre","base_qty":30,"dsr_rate":1780.00,
 "ref_pattern":True,
 "sections":[
  ("REFERENCE ITEMS (A)",[
   ("REF#2.9.3","Excavation hard rock no blasting (40.399 cum)","cum",40.399,1080.55),
   ("REF#2.25","Filling in trenches (40.399 cum)","cum",40.399,219.65),
  ]),
  ("LABOUR",[
   ("0114","Beldar (rock ramming)","day",2.50,558.00),
  ]),
 ],
},

# ─── 2.14, 2.15 EXTRA DEPTH ROCK (%) ─────────────────────────────────────────
{
 "id":"2.14","desc":"Extra for excavating trenches for pipes, cables, etc. in ordinary/hard rock exceeding 1.5 m in depth but not exceeding 3 m. (Rate is over corresponding basic item for depth upto 1.5 metre)",
 "unit":"%","base_qty":1,"dsr_rate":103.75,
 "pct_item":True,"sections":[],
 "notes":"DAR computed 103.60% (DSR 2021: 103.75%).",
},
{
 "id":"2.15","desc":"Extra for excavating trenches for pipes, cables, etc. in ordinary/hard rock exceeding 3m in depth but not exceeding 4.5 m. (Rate is over correspondingbasic item for depth upto 1.5 metre)",
 "unit":"%","base_qty":1,"dsr_rate":256.15,
 "pct_item":True,"sections":[],
 "notes":"DAR computed 255.60% (DSR 2021: 256.15%).",
},

# ─── 2.16 CLOSE TIMBERING IN TRENCHES ────────────────────────────────────────
{
 "id":"2.16.1","desc":"Close timbering in trenches including strutting, shoring and packing cavities (wherever required) complete. (Measurements to be taken of the face area timbered). - Depth not exceeding 1.5 m",
 "unit":"sqm","base_qty":90,"dsr_rate":129.95,
 "sections":[
  ("MATERIAL",[
   ("1198","Second class kail wood in planks","10 cudm",21.375,260.00),
   ("1197","Second class kail wood in scantling","10 cudm",7.50,260.00),
   ("0302","Safeda ballies 125 mm diameter and 1.5m long","metre",3.1875,40.00),
  ]),
  ("CARRIAGE",[
   ("2204","Carriage of Timber","cum",1.3125,118.59),
  ]),
  ("LABOUR",[
   ("0112","Carpenter 2nd class","day",0.50,679.00),
   ("0114","Beldar","day",1.00,558.00),
   ("9999","Sundries","LS",26.91,2.00),
  ]),
 ],
},
{
 "id":"2.16.2","desc":"Close timbering in trenches including strutting, shoring and packing cavities (wherever required) complete. (Measurements to be taken of the face area timbered). - Depth exceeding 1.5 m but not exceeding 3 m",
 "unit":"sqm","base_qty":90,"dsr_rate":141.20,
 "sections":[
  ("MATERIAL",[
   ("1198","Second class kail wood in planks","10 cudm",21.375,260.00),
   ("1197","Second class kail wood in scantling","10 cudm",7.50,260.00),
   ("0302","Safeda ballies 125 mm diameter and 1.5m long","metre",3.1875,40.00),
  ]),
  ("CARRIAGE",[
   ("2204","Carriage of Timber","cum",1.3125,118.59),
  ]),
  ("LABOUR",[
   ("0112","Carpenter 2nd class","day",0.75,679.00),
   ("0114","Beldar","day",2.00,558.00),
   ("9999","Sundries","LS",40.43,2.00),
  ]),
 ],
},
{
 "id":"2.16.3","desc":"Close timbering in trenches including strutting, shoring and packing cavities (wherever required) complete. (Measurements to be taken of the face area timbered). - Depth exceeding 3 m but not exceeding 4.5 m",
 "unit":"sqm","base_qty":90,"dsr_rate":166.55,
 "sections":[
  ("MATERIAL",[
   ("1198","Second class kail wood in planks","10 cudm",21.375,260.00),
   ("1197","Second class kail wood in scantling","10 cudm",7.50,260.00),
   ("0302","Safeda ballies 125 mm diameter and 1.5m long","metre",3.1875,40.00),
  ]),
  ("CARRIAGE",[
   ("2204","Carriage of Timber","cum",1.3125,118.59),
  ]),
  ("LABOUR",[
   ("0112","Carpenter 2nd class","day",1.50,679.00),
   ("0114","Beldar","day",4.00,558.00),
   ("9999","Sundries","LS",80.73,2.00),
  ]),
 ],
},

# ─── 2.17 CLOSE TIMBERING IN SHAFTS ──────────────────────────────────────────
{
 "id":"2.17.1","desc":"Close timbering in case of shafts, wells, cesspits, manholes and the like including strutting, shoring and packing cavities (wherever required) etc. complete. (Measurements to be taken of the face area timbered). - Depth not exceeding 1.5 m",
 "unit":"sqm","base_qty":6.6,"dsr_rate":138.45,
 "sections":[
  ("MATERIAL",[
   ("1198","Kail/Deodar planks 38 mm thick","cum",0.00990,12000.00),
   ("1197","Kail/Deodar scantlings","cum",0.00310,12000.00),
   ("0302","Safeda poles/Ballies","each",0.0825,40.00),
  ]),
  ("CARRIAGE",[
   ("2204","Carriage of timber","cum",0.01300,118.59),
  ]),
  ("LABOUR",[
   ("0112","Carpenter 2nd class","day",0.145,679.00),
   ("0114","Beldar","day",0.073,558.00),
   ("9999","Sundries","LS",1.56,2.00),
  ]),
 ],
},
{
 "id":"2.17.2","desc":"Close timbering in case of shafts, wells, cesspits, manholes and the like including strutting, shoring and packing cavities (wherever required) etc. complete. (Measurements to be taken of the face area timbered). - Depth exceeding 1.5 m but not exceeding 3 m",
 "unit":"sqm","base_qty":6.6,"dsr_rate":162.55,
 "sections":[
  ("MATERIAL",[
   ("1198","Kail/Deodar planks 38 mm thick","cum",0.01300,12000.00),
   ("1197","Kail/Deodar scantlings","cum",0.00420,12000.00),
   ("0302","Safeda poles/Ballies","each",0.0825,40.00),
  ]),
  ("CARRIAGE",[
   ("2204","Carriage of timber","cum",0.01720,118.59),
  ]),
  ("LABOUR",[
   ("0112","Carpenter 2nd class","day",0.175,679.00),
   ("0114","Beldar","day",0.088,558.00),
   ("9999","Sundries","LS",1.82,2.00),
  ]),
 ],
},
{
 "id":"2.17.3","desc":"Close timbering in case of shafts, wells, cesspits, manholes and the like including strutting, shoring and packing cavities (wherever required) etc. complete. (Measurements to be taken of the face area timbered). - Depth exceeding 3 m but not exceeding 4.5 m",
 "unit":"sqm","base_qty":6.6,"dsr_rate":187.95,
 "sections":[
  ("MATERIAL",[
   ("1198","Kail/Deodar planks 38 mm thick","cum",0.01560,12000.00),
   ("1197","Kail/Deodar scantlings","cum",0.00560,12000.00),
   ("0302","Safeda poles/Ballies","each",0.0825,40.00),
  ]),
  ("CARRIAGE",[
   ("2204","Carriage of timber","cum",0.02120,118.59),
  ]),
  ("LABOUR",[
   ("0112","Carpenter 2nd class","day",0.210,679.00),
   ("0114","Beldar","day",0.105,558.00),
   ("9999","Sundries","LS",2.34,2.00),
  ]),
 ],
},

# ─── 2.18 CLOSE TIMBERING OVER AREAS ─────────────────────────────────────────
{
 "id":"2.18.1","desc":"Close timbering over areas including strutting, shoring and packing cavities (wherever required) etc. complete. (Measurements to be taken of the face area timbered): Depth not exceeding 1.5 m",
 "unit":"sqm","base_qty":45,"dsr_rate":116.25,
 "sections":[
  ("MATERIAL",[
   ("1198","Kail/Deodar planks 38 mm thick","cum",0.0560,12000.00),
   ("1197","Kail/Deodar scantlings","cum",0.0182,12000.00),
   ("0302","Safeda poles/Ballies","each",0.4500,40.00),
  ]),
  ("CARRIAGE",[
   ("2204","Carriage of timber","cum",0.0742,118.59),
  ]),
  ("LABOUR",[
   ("0112","Carpenter 2nd class","day",0.750,679.00),
   ("0114","Beldar","day",0.375,558.00),
   ("9999","Sundries","LS",8.06,2.00),
  ]),
 ],
},
{
 "id":"2.18.2","desc":"Close timbering over areas including strutting, shoring and packing cavities (wherever required) etc. complete. (Measurements to be taken of the face area timbered): Depth exceeding 1.5 m but not exceeding 3 m",
 "unit":"sqm","base_qty":45,"dsr_rate":129.75,
 "sections":[
  ("MATERIAL",[
   ("1198","Kail/Deodar planks 38 mm thick","cum",0.0680,12000.00),
   ("1197","Kail/Deodar scantlings","cum",0.0230,12000.00),
   ("0302","Safeda poles/Ballies","each",0.4500,40.00),
  ]),
  ("CARRIAGE",[
   ("2204","Carriage of timber","cum",0.0910,118.59),
  ]),
  ("LABOUR",[
   ("0112","Carpenter 2nd class","day",0.900,679.00),
   ("0114","Beldar","day",0.450,558.00),
   ("9999","Sundries","LS",10.90,2.00),
  ]),
 ],
},
{
 "id":"2.18.3","desc":"Close timbering over areas including strutting, shoring and packing cavities (wherever required) etc. complete. (Measurements to be taken of the face area timbered): Depth exceeding 3 m but not exceeding 4.5 m",
 "unit":"sqm","base_qty":45,"dsr_rate":143.90,
 "sections":[
  ("MATERIAL",[
   ("1198","Kail/Deodar planks 38 mm thick","cum",0.0830,12000.00),
   ("1197","Kail/Deodar scantlings","cum",0.0300,12000.00),
   ("0302","Safeda poles/Ballies","each",0.4500,40.00),
  ]),
  ("CARRIAGE",[
   ("2204","Carriage of timber","cum",0.1130,118.59),
  ]),
  ("LABOUR",[
   ("0112","Carpenter 2nd class","day",1.100,679.00),
   ("0114","Beldar","day",0.550,558.00),
   ("9999","Sundries","LS",12.72,2.00),
  ]),
 ],
},

# ─── 2.19 EXTRA PERMANENT CLOSE TIMBERING ────────────────────────────────────
{
 "id":"2.19","desc":"Extra for planking, strutting and packing materials for cavities (in close timbering) if required to be left permanently in position. (Face area of timber permanently left to be measured).",
 "unit":"sqm","base_qty":90,"dsr_rate":1596.50,
 "sections":[
  ("MATERIAL",[
   ("1198","Kail/Deodar planks 38 mm (permanent)","cum",3.3250,12000.00),
   ("1197","Kail/Deodar scantlings (permanent)","cum",1.1667,12000.00),
   ("0302","Safeda poles/Ballies (permanent)","each",44.625,40.00),
  ]),
  ("CARRIAGE",[
   ("2204","Carriage of timber","cum",4.4917,118.59),
  ]),
 ],
},

# ─── 2.20 OPEN TIMBERING IN TRENCHES ─────────────────────────────────────────
{
 "id":"2.20.1","desc":"Open timbering in trenches including strutting and shoring complete (measurements to be taken of the face area timbered): Depth not exceeding 1.5 m",
 "unit":"sqm","base_qty":90,"dsr_rate":67.00,
 "sections":[
  ("MATERIAL",[
   ("1198","Kail/Deodar planks 38 mm","cum",0.0540,12000.00),
   ("1197","Kail/Deodar scantlings","cum",0.0184,12000.00),
   ("0302","Safeda poles/Ballies","each",0.4500,40.00),
  ]),
  ("CARRIAGE",[
   ("2204","Carriage of timber","cum",0.0724,118.59),
  ]),
  ("LABOUR",[
   ("0112","Carpenter 2nd class","day",0.750,679.00),
   ("0114","Beldar","day",0.375,558.00),
   ("9999","Sundries","LS",8.06,2.00),
  ]),
 ],
},
{
 "id":"2.20.2","desc":"Open timbering in trenches including strutting and shoring complete (measurements to be taken of the face area timbered): Depth exceeding 1.5 m but not exceeding 3 m",
 "unit":"sqm","base_qty":90,"dsr_rate":74.05,
 "sections":[
  ("MATERIAL",[
   ("1198","Kail/Deodar planks 38 mm","cum",0.0640,12000.00),
   ("1197","Kail/Deodar scantlings","cum",0.0218,12000.00),
   ("0302","Safeda poles/Ballies","each",0.4500,40.00),
  ]),
  ("CARRIAGE",[
   ("2204","Carriage of timber","cum",0.0858,118.59),
  ]),
  ("LABOUR",[
   ("0112","Carpenter 2nd class","day",0.900,679.00),
   ("0114","Beldar","day",0.450,558.00),
   ("9999","Sundries","LS",9.88,2.00),
  ]),
 ],
},
{
 "id":"2.20.3","desc":"Open timbering in trenches including strutting and shoring complete (measurements to be taken of the face area timbered): Depth exceeding 3 m but not exceeding 4.5 m",
 "unit":"sqm","base_qty":90,"dsr_rate":85.70,
 "sections":[
  ("MATERIAL",[
   ("1198","Kail/Deodar planks 38 mm","cum",0.0840,12000.00),
   ("1197","Kail/Deodar scantlings","cum",0.0288,12000.00),
   ("0302","Safeda poles/Ballies","each",0.4500,40.00),
  ]),
  ("CARRIAGE",[
   ("2204","Carriage of timber","cum",0.1128,118.59),
  ]),
  ("LABOUR",[
   ("0112","Carpenter 2nd class","day",1.100,679.00),
   ("0114","Beldar","day",0.550,558.00),
   ("9999","Sundries","LS",12.72,2.00),
  ]),
 ],
},

# ─── 2.21 OPEN TIMBERING IN SHAFTS ───────────────────────────────────────────
{
 "id":"2.21.1","desc":"Open timbering in case of shafts, wells, cesspits, manholes and the like including strutting and shoring complete (Measurements to be taken of the face area timbered): Depth not exceeding 1.5 m",
 "unit":"sqm","base_qty":6.6,"dsr_rate":60.30,
 "sections":[
  ("MATERIAL",[
   ("1198","Kail/Deodar planks","cum",0.00495,12000.00),
   ("1197","Kail/Deodar scantlings","cum",0.00168,12000.00),
   ("0302","Safeda poles/Ballies","each",0.04125,40.00),
  ]),
  ("CARRIAGE",[
   ("2204","Carriage of timber","cum",0.00663,118.59),
  ]),
  ("LABOUR",[
   ("0112","Carpenter 2nd class","day",0.068,679.00),
   ("0114","Beldar","day",0.034,558.00),
   ("9999","Sundries","LS",0.735,2.00),
  ]),
 ],
},
{
 "id":"2.21.2","desc":"Open timbering in case of shafts, wells, cesspits, manholes and the like including strutting and shoring complete (Measurements to be taken of the face area timbered): Depth exceeding 1.5 m but not exceeding 3 m",
 "unit":"sqm","base_qty":6.6,"dsr_rate":71.60,
 "sections":[
  ("MATERIAL",[
   ("1198","Kail/Deodar planks","cum",0.00594,12000.00),
   ("1197","Kail/Deodar scantlings","cum",0.00200,12000.00),
   ("0302","Safeda poles/Ballies","each",0.04125,40.00),
  ]),
  ("CARRIAGE",[
   ("2204","Carriage of timber","cum",0.00794,118.59),
  ]),
  ("LABOUR",[
   ("0112","Carpenter 2nd class","day",0.083,679.00),
   ("0114","Beldar","day",0.042,558.00),
   ("9999","Sundries","LS",0.910,2.00),
  ]),
 ],
},
{
 "id":"2.21.3","desc":"Open timbering in case of shafts, wells, cesspits, manholes and the like including strutting and shoring complete (Measurements to be taken of the face area timbered): Depth exceeding 3 m but not exceeding 4.5 m",
 "unit":"sqm","base_qty":6.6,"dsr_rate":86.50,
 "sections":[
  ("MATERIAL",[
   ("1198","Kail/Deodar planks","cum",0.00770,12000.00),
   ("1197","Kail/Deodar scantlings","cum",0.00263,12000.00),
   ("0302","Safeda poles/Ballies","each",0.04125,40.00),
  ]),
  ("CARRIAGE",[
   ("2204","Carriage of timber","cum",0.01033,118.59),
  ]),
  ("LABOUR",[
   ("0112","Carpenter 2nd class","day",0.100,679.00),
   ("0114","Beldar","day",0.050,558.00),
   ("9999","Sundries","LS",1.105,2.00),
  ]),
 ],
},

# ─── 2.22 OPEN TIMBERING OVER AREAS ──────────────────────────────────────────
{
 "id":"2.22.1","desc":"Open timbering over areas including strutting, shoring etc. complete. (Measurements to be taken of the face area timbered): Depth not exceeding 1.5 m",
 "unit":"sqm","base_qty":45,"dsr_rate":40.90,
 "sections":[
  ("MATERIAL",[
   ("1198","Kail/Deodar planks","cum",0.0270,12000.00),
   ("1197","Kail/Deodar scantlings","cum",0.0091,12000.00),
   ("0302","Safeda poles/Ballies","each",0.2250,40.00),
  ]),
  ("CARRIAGE",[
   ("2204","Carriage of timber","cum",0.0361,118.59),
  ]),
  ("LABOUR",[
   ("0112","Carpenter 2nd class","day",0.375,679.00),
   ("0114","Beldar","day",0.188,558.00),
   ("9999","Sundries","LS",4.03,2.00),
  ]),
 ],
},
{
 "id":"2.22.2","desc":"Open timbering over areas including strutting, shoring etc. complete. (Measurements to be taken of the face area timbered): Depth exceeding 1.5 m but not exceeding 3 m",
 "unit":"sqm","base_qty":45,"dsr_rate":48.45,
 "sections":[
  ("MATERIAL",[
   ("1198","Kail/Deodar planks","cum",0.0330,12000.00),
   ("1197","Kail/Deodar scantlings","cum",0.0113,12000.00),
   ("0302","Safeda poles/Ballies","each",0.2250,40.00),
  ]),
  ("CARRIAGE",[
   ("2204","Carriage of timber","cum",0.0443,118.59),
  ]),
  ("LABOUR",[
   ("0112","Carpenter 2nd class","day",0.450,679.00),
   ("0114","Beldar","day",0.225,558.00),
   ("9999","Sundries","LS",5.04,2.00),
  ]),
 ],
},
{
 "id":"2.22.3","desc":"Open timbering over areas including strutting, shoring etc. complete. (Measurements to be taken of the face area timbered): Depth exceeding 3 m but not exceeding 4.5 m",
 "unit":"sqm","base_qty":45,"dsr_rate":60.70,
 "sections":[
  ("MATERIAL",[
   ("1198","Kail/Deodar planks","cum",0.0450,12000.00),
   ("1197","Kail/Deodar scantlings","cum",0.0154,12000.00),
   ("0302","Safeda poles/Ballies","each",0.2250,40.00),
  ]),
  ("CARRIAGE",[
   ("2204","Carriage of timber","cum",0.0604,118.59),
  ]),
  ("LABOUR",[
   ("0112","Carpenter 2nd class","day",0.563,679.00),
   ("0114","Beldar","day",0.281,558.00),
   ("9999","Sundries","LS",6.36,2.00),
  ]),
 ],
},

# ─── 2.23 EXTRA PERMANENT OPEN TIMBERING ─────────────────────────────────────
{
 "id":"2.23","desc":"Extra for planking and strutting in open timbering if required to be left permanently in position. (Face area of the timber permanently left to be measured).",
 "unit":"sqm","base_qty":90,"dsr_rate":820.55,
 "sections":[
  ("MATERIAL",[
   ("1198","Kail/Deodar planks (permanent)","cum",1.7100,12000.00),
   ("1197","Kail/Deodar scantlings (permanent)","cum",0.5850,12000.00),
   ("0302","Safeda poles/Ballies (permanent)","each",22.500,40.00),
  ]),
  ("CARRIAGE",[
   ("2204","Carriage of timber","cum",2.2950,118.59),
  ]),
 ],
},

# ─── 2.24 PERCENTAGE EXTRAS ───────────────────────────────────────────────────
{
 "id":"2.24.1","desc":"Extra rates for quantities of works, executed: In or under water and/or liquid mud, including pumping out water as required",
 "unit":"%","base_qty":1,"dsr_rate":None,
 "pct_item":True,"conditional_extra":True,"percent":20,
 "conditional_rule":"20% of qualifying selected base rate",
 "sections":[],
 "notes":"Conditional 20% extra over each applicable earthwork item, limited to qualifying work. Measure depth from sub-soil water level to centre of gravity. No resource rows.",
},
{
 "id":"2.24.2","desc":"Extra rates for quantities of works, executed: In or under foul position, including pumping out water as required",
 "unit":"%","base_qty":1,"dsr_rate":None,
 "pct_item":True,"conditional_extra":True,"percent":25,
 "conditional_rule":"25% of qualifying selected base rate",
 "sections":[],
 "notes":"Conditional 25% extra over each applicable earthwork item, limited to qualifying work. Measure depth from sub-soil water level to centre of gravity. No resource rows.",
},

# ─── 2.25 FILLING IN TRENCHES ────────────────────────────────────────────────
{
 "id":"2.25","desc":"Filling available excavated earth (excluding rock) in trenches, plinth, sides of foundations etc. in layers not exceeding 20cm in depth, consolidating each deposited layer by ramming and watering, lead up to 50 m and lift upto 1.5 m.",
 "unit":"cum","base_qty":10,"dsr_rate":219.65,
 "sections":[
  ("LABOUR",[
   ("0128","Mate","day",0.20,617.00),
   ("0115","Coolie","day",2.50,558.00),
   ("0101","Bhisti","day",0.20,617.00),
  ]),
 ],
},

# ─── 2.25(a) FILLING WITH EXCAVATED EARTH (MACHINE) ─────────────────────────
{
 "id":"2.25(a)","desc":"Excavating, supplying and filling of local earth (including royalty) by mechanical transport upto a lead of 5km also including ramming and watering of the earth in layers not exceeding 20 cm in trenches, plinth, sides of foundation etc. complete.",
 "unit":"cum","base_qty":10,"dsr_rate":323.90,
 "ref_pattern":True,
 "sections":[
  ("REFERENCE ITEMS (A - excluded from multipliers)",[
   ("REF 2.6.1","Hydraulic excavation ordinary soil (10 cum)","cum",10.00,181.85),
   ("REF 1.2.2","DEDUCT: dozing/spreading (10 cum)","cum",-10.00,175.78),
   ("REF 1.1.2","Carriage 10 cum","cum",10.00,201.83),
   ("0979","Royalty on earth","cum",10.00,40.00),
  ]),
  ("LABOUR (subject to multiplier chain)",[
   ("0114","Beldar","day",0.45,558.00),
   ("0101","Bhisti","day",0.35,617.00),
  ]),
 ],
},

# ─── 2.26 EXTRA LIFT ──────────────────────────────────────────────────────────
{
 "id":"2.26.1","desc":"Extra for every additional lift of 1.5 m or part thereof in excavation / banking excavated or stacked materials. - All kinds of soil",
 "unit":"cum","base_qty":10,"dsr_rate":90.40,
 "sections":[
  ("LABOUR",[
   ("0128","Mate","day",0.10,617.00),
   ("0114","Beldar","day",1.10,558.00),
  ]),
 ],
},
{
 "id":"2.26.2","desc":"Extra for every additional lift of 1.5 m or part thereof in excavation / banking excavated or stacked materials. - Ordinary or hard rock",
 "unit":"cum","base_qty":10,"dsr_rate":162.10,
 "sections":[
  ("LABOUR",[
   ("0128","Mate","day",0.20,617.00),
   ("0114","Beldar","day",1.95,558.00),
  ]),
 ],
},

# ─── 2.27 FILLING WITH SAND ───────────────────────────────────────────────────
{
 "id":"2.27","desc":"Supplying and filling in plinth with sand under floors, including watering, ramming, consolidating and dressing complete.",
 "unit":"cum","base_qty":10,"dsr_rate":1953.05,
 "sections":[
  ("MATERIAL",[
   ("6501","Sand Zone V","cum",10.00,1225.00),
  ]),
  ("CARRIAGE",[
   ("2335","Carriage of sand by road","cum",10.00,103.77),
  ]),
  ("LABOUR",[
   ("0114","Beldar","day",0.89,558.00),
   ("0115","Coolie","day",1.07,558.00),
   ("0101","Bhisti","day",0.35,617.00),
  ]),
 ],
},

# ─── 2.28 DRESSING OF SLOPES ──────────────────────────────────────────────────
{
 "id":"2.28.1","desc":"Surface dressing of the ground including removing vegetation and in-equalities not exceeding 15 cm deep and disposal of rubbish, lead up to 50 m and lift up to 1.5 m. - All kinds of soil",
 "unit":"sqm","base_qty":100,"dsr_rate":24.35,
 "sections":[
  ("LABOUR",[
   ("0114","Beldar","day",1.97,558.00),
   ("0115","Coolie","day",1.29,558.00),
  ]),
 ],
},
{
 "id":"2.28.2","desc":"Surface dressing of the ground including removing vegetation and in-equalities not exceeding 15 cm deep and disposal of rubbish, lead up to 50 m and lift up to 1.5 m. - Hard soil",
 "unit":"sqm","base_qty":100,"dsr_rate":35.15,
 "sections":[
  ("LABOUR",[
   ("0114","Beldar","day",2.50,558.00),
   ("0115","Coolie","day",1.60,558.00),
  ]),
 ],
},

# ─── 2.29 RAMMING & WATERING ──────────────────────────────────────────────────
{
 "id":"2.29.1","desc":"Ploughing the existing ground to a depth of 15 cm to 25 cm and watering the same. - All kinds of soil",
 "unit":"sqm","base_qty":100,"dsr_rate":24.65,
 "sections":[
  ("LABOUR",[
   ("0114","Beldar","day",2.75,558.00),
   ("0101","Bhisti","day",0.50,617.00),
  ]),
 ],
},
{
 "id":"2.29.2","desc":"Ploughing the existing ground to a depth of 15 cm to 25 cm and watering the same. - Hard soil",
 "unit":"sqm","base_qty":100,"dsr_rate":34.70,
 "sections":[
  ("LABOUR",[
   ("0114","Beldar","day",3.30,558.00),
   ("0101","Bhisti","day",0.60,617.00),
  ]),
 ],
},

# ─── 2.30 BORING HOLES IN ROCK ───────────────────────────────────────────────
{
 "id":"2.30.1","desc":"Excavating holes more than 0.10 cum & upto 0.5 cum including getting out the excavated soil, then returning the soil as required in layers not exceeding 20cm in depth, including consolidating each deposited layer by ramming, watering etc, disposing of surplus excavated soil, as directed within a lead of 50 m and lift upto 1.5 m. - All kinds of soil",
 "unit":"each","base_qty":10,"dsr_rate":79.30,
 "ref_pattern":True,
 "sections":[
  ("REFERENCE ITEMS (A)",[
   ("REF 2.8.1","Excavation pipe trench ordinary soil (3.00 cum)","cum",3.00,252.30),
  ]),
  ("MATERIAL (subject to chain)",[
   ("9999","Sundries","LS",13.52,2.00),
  ]),
 ],
},
{
 "id":"2.30.2","desc":"Excavating holes more than 0.10 cum & upto 0.5 cum including getting out the excavated soil, then returning the soil as required in layers not exceeding 20cm in depth, including consolidating each deposited layer by ramming, watering etc, disposing of surplus excavated soil, as directed within a lead of 50 m and lift upto 1.5 m. - Ordinary rock",
 "unit":"each","base_qty":10,"dsr_rate":138.05,
 "ref_pattern":True,
 "sections":[
  ("REFERENCE ITEMS (A)",[
   ("REF 2.9.1","Excavation ordinary rock (3.00 cum)","cum",3.00,448.15),
  ]),
  ("MATERIAL",[
   ("9999","Sundries","LS",27.04,2.00),
  ]),
 ],
},
{
 "id":"2.30.3","desc":"Excavating holes more than 0.10 cum & upto 0.5 cum including getting out the excavated soil, then returning the soil as required in layers not exceeding 20cm in depth, including consolidating each deposited layer by ramming, watering etc, disposing of surplus excavated soil, as directed within a lead of 50 m and lift upto 1.5 m. - Hard rock (requiring blasting)",
 "unit":"each","base_qty":10,"dsr_rate":222.30,
 "ref_pattern":True,
 "sections":[
  ("REFERENCE ITEMS (A)",[
   ("REF 2.9.2","Excavation hard rock blasting (3.00 cum)","cum",3.00,729.00),
  ]),
  ("MATERIAL",[
   ("9999","Sundries","LS",27.04,2.00),
  ]),
 ],
},
{
 "id":"2.30.4","desc":"Excavating holes more than 0.10 cum & upto 0.5 cum including getting out the excavated soil, then returning the soil as required in layers not exceeding 20cm in depth, including consolidating each deposited layer by ramming, watering etc, disposing of surplus excavated soil, as directed within a lead of 50 m and lift upto 1.5 m. - Hard rock (blasting prohibited)",
 "unit":"each","base_qty":10,"dsr_rate":327.80,
 "ref_pattern":True,
 "sections":[
  ("REFERENCE ITEMS (A)",[
   ("REF 2.9.3","Excavation hard rock no blasting (3.00 cum)","cum",3.00,1080.55),
  ]),
  ("MATERIAL",[
   ("9999","Sundries","LS",27.04,2.00),
  ]),
 ],
},

# ─── 2.31 SCARIFYING ──────────────────────────────────────────────────────────
{
 "id":"2.31","desc":"Clearing jungle including uprooting of rank vegetation, grass, brush wood, trees and saplings of girth up to 30 cm measured at a height of 1 m above ground level and removal of rubbish up to a distance of 50 m outside the periphery of the area cleared.",
 "unit":"sqm","base_qty":100,"dsr_rate":12.55,
 "sections":[
  ("LABOUR",[
   ("0114","Beldar","day",1.08,558.00),
   ("0115","Coolie","day",0.60,558.00),
  ]),
 ],
},

# ─── 2.32 GRADING BORROW PITS ────────────────────────────────────────────────
{
 "id":"2.32","desc":"Clearing grass and removal of the rubbish up to a distance of 50 m outside the periphery of the area cleared.",
 "unit":"sqm","base_qty":100,"dsr_rate":6.40,
 "sections":[
  ("LABOUR",[
   ("0114","Beldar","day",0.60,558.00),
   ("0115","Coolie","day",0.25,558.00),
   ("9999","Sundries","LS",1.82,2.00),
  ]),
 ],
},

# ─── 2.33 TREE FELLING ────────────────────────────────────────────────────────
{
 "id":"2.33.1","desc":"Felling trees of the girth (measured at a height of 1 m above ground level), including cutting of trunks and branches, removing the roots and stacking of serviceable material and disposal of unserviceable material. - Beyond 30 cm girth upto and including 60 cm girth",
 "unit":"each","base_qty":1,"dsr_rate":380.60,
 "sections":[
  ("LABOUR",[
   ("0114","Beldar","day",0.33,558.00),
   ("0115","Coolie","day",0.17,558.00),
   ("9999","Sundries","LS",2.73,2.00),
  ]),
 ],
},
{
 "id":"2.33.2","desc":"Felling trees of the girth (measured at a height of 1 m above ground level), including cutting of trunks and branches, removing the roots and stacking of serviceable material and disposal of unserviceable material. - Beyond 60 cm girth upto and including 120 cm girth",
 "unit":"each","base_qty":1,"dsr_rate":1694.40,
 "sections":[
  ("LABOUR",[
   ("0114","Beldar","day",1.50,558.00),
   ("0115","Coolie","day",0.75,558.00),
   ("9999","Sundries","LS",5.46,2.00),
  ]),
 ],
},
{
 "id":"2.33.3","desc":"Felling trees of the girth (measured at a height of 1 m above ground level), including cutting of trunks and branches, removing the roots and stacking of serviceable material and disposal of unserviceable material. - Beyond 120 cm girth upto and including 240 cm girth",
 "unit":"each","base_qty":1,"dsr_rate":7860.55,
 "sections":[
  ("LABOUR",[
   ("0114","Beldar","day",7.00,558.00),
   ("0115","Coolie","day",3.50,558.00),
   ("9999","Sundries","LS",8.06,2.00),
  ]),
 ],
},
{
 "id":"2.33.4","desc":"Felling trees of the girth (measured at a height of 1 m above ground level), including cutting of trunks and branches, removing the roots and stacking of serviceable material and disposal of unserviceable material. - Above 240 cm girth",
 "unit":"each","base_qty":1,"dsr_rate":15749.95,
 "sections":[
  ("LABOUR",[
   ("0114","Beldar","day",14.00,558.00),
   ("0115","Coolie","day",7.00,558.00),
   ("9999","Sundries","LS",26.91,2.00),
  ]),
 ],
},

# ─── 2.34 ANTI-TERMITE CHEMICAL ──────────────────────────────────────────────
{
 "id":"2.34.1","desc":"Supplying chemical emulsion in sealed containers including delivery as specified. - Chlorpyriphos/ Lindane emulsifiable concentrate of 20%",
 "unit":"litre","base_qty":100,"dsr_rate":200.85,
 "sections":[
  ("MATERIAL",[
   ("7022","Chlorpyriphos 20% EC concentrate","litre",100.00,150.00),
  ]),
  ("CARRIAGE",[
   ("2342","Carriage of chemical","cum",1.00,10.38),
  ]),
 ],
},

# ─── 2.35 POST-CONSTRUCTIONAL ANTI-TERMITE ───────────────────────────────────
{
 "id":"2.35.1.1","desc":"Diluting and injecting chemical emulsion for POST-CONSTRUCTIONAL anti- termite treatment (excluding the cost of chemical emulsion) : - Along external wall where the apron is not provided using chemical emulsion @ 7.5 litres / sqm of the vertical surface of the substructure to a depth of 300mm including excavation channel along the wall & rodding etc. complete: - With Chlorpyriphos/ Lindane E.C. 20% with 1% concentration",
 "unit":"metre","base_qty":1,"dsr_rate":28.25,
 "sections":[
  ("LABOUR",[
   ("0114","Beldar","day",0.33,558.00),
   ("9999","Sundries","LS",13.52,2.00),
  ]),
 ],
},
{
 "id":"2.35.2.1","desc":"Diluting and injecting chemical emulsion for POST-CONSTRUCTIONAL anti- termite treatment (excluding the cost of chemical emulsion) : - Along the external wall below concrete or masonry apron using chemical emulsion @ 2.25 litres per linear metre including drilling and plugging holes etc.: - With Chlorpyriphos/ Lindane E.C. 20% with 1% concentration",
 "unit":"metre","base_qty":1,"dsr_rate":39.45,
 "sections":[
  ("LABOUR",[
   ("0114","Beldar","day",0.40,558.00),
   ("9999","Sundries","LS",35.88,2.00),
  ]),
 ],
},
{
 "id":"2.35.3.1","desc":"Diluting and injecting chemical emulsion for POST-CONSTRUCTIONAL anti- termite treatment (excluding the cost of chemical emulsion) : - Treatment of soil under existing floors using chemical emulsion @ one litre per hole, 300 mm apart including drilling 12 mm diameter holes and plugging with cement mortar 1 :2 (1 cement : 2 Coarse sand) to match the existing floor: - With Chlorpyriphos/Lindane E.C. 20% with 1% concentration",
 "unit":"sqm","base_qty":1,"dsr_rate":227.05,
 "sections":[
  ("LABOUR",[
   ("0114","Beldar","day",2.00,558.00),
   ("0124","Mason 2nd class","day",0.50,679.00),
   ("9999","Sundries","LS",35.88,2.00),
  ]),
 ],
},
{
 "id":"2.35.4.1","desc":"Diluting and injecting chemical emulsion for POST-CONSTRUCTIONAL anti- termite treatment (excluding the cost of chemical emulsion) : - Treatment of existing masonry using chemical emulsion @ one litre per hole at 300 mm interval including drilling holes at 45 degree and plugging them with cement mortar 1:2 (1 cement : 2 coarse sand) to the full depth of the hole : - With Chlorpyriphos/Lindane E.C. 20% with 1% concentration",
 "unit":"metre","base_qty":1,"dsr_rate":31.75,
 "sections":[
  ("LABOUR",[
   ("0114","Beldar","day",0.30,558.00),
   ("0124","Mason 2nd class","day",0.05,679.00),
   ("9999","Sundries","LS",17.94,2.00),
  ]),
 ],
},
{
 "id":"2.35.5","desc":"Diluting and injecting chemical emulsion for POST-CONSTRUCTIONAL anti- termite treatment (excluding the cost of chemical emulsion) : Treatment at points of contact of wood work by chemical emulsion Chlorpyriphos/ Lindane (in oil or kerosene based solution) @ 0.5 litres per hole by drilling 6 mm dia holes at downward angle of 45 degree at 150 mm centre to centre and sealing the same.",
 "unit":"metre","base_qty":1,"dsr_rate":254.00,
 "sections":[
  ("MATERIAL",[
   ("","Kerosene oil","litre",32.30,50.00),
  ]),
  ("LABOUR",[
   ("0112","Carpenter 2nd class","day",0.20,679.00),
   ("0114","Beldar","day",0.20,558.00),
   ("9999","Sundries","LS",17.94,2.00),
  ]),
 ],
},

# ─── 2.36 COMPACTION BY SHEEP-FOOT ROLLER ────────────────────────────────────
{
 "id":"2.36","desc":"Extra for levelling & neatly dressing of disposed soil completely as directed by Engineer-in-charge.",
 "unit":"cum","base_qty":10,"dsr_rate":66.35,
 "sections":[
  ("LABOUR",[
   ("0128","Mate","day",0.08,617.00),
   ("0114","Beldar/Coolie","day",0.80,558.00),
  ]),
 ],
},

# ─── 2.37 FILLING WITH FLY ASH ───────────────────────────────────────────────
{
 "id":"2.37","desc":"Supply and stacking of Fly ash conforming to IRC- 58 at site, including carriage, loading , unloading & stacking up to any lead (measured stacks will be reduced by 20% for payment).",
 "unit":"cum","base_qty":1,"dsr_rate":152.20,
 "sections":[
  ("MATERIAL",[
   ("1980","Fly ash","cum",1.00,10.00),
  ]),
  ("CARRIAGE",[
   ("2262","Carriage of fly ash","cum",1.00,103.77),
  ]),
 ],
},

# ─── 2.38 FILLING – IMPORTED EARTH ───────────────────────────────────────────
{
 "id":"2.38","desc":"Filling with available fly ash and earth (excluding rock) in trenches or embankment in layers (each layer should not exceed 15 cm), with intermediate layer of compacted earth (Soil density of 98%) after every four layers of compacted depth of fly ash, sides & top layer of filling shall be done with earth having total minimum compacted thickness 30 cm or as decided by Engineer - in-charge, including compacting each layer by rolling/ ramming and watering, all complete as per drawing and direction of Engineer -in - charge.",
 "unit":"cum","base_qty":10,"dsr_rate":219.65,
 "sections":[
  ("LABOUR",[
   ("0128","Mate","day",0.20,617.00),
   ("0115","Coolie","day",2.50,558.00),
   ("0101","Bhisti","day",0.20,617.00),
  ]),
 ],
},

]  # end ITEMS


# ── Sheet builder ─────────────────────────────────────────────────────────────

COL_CODE   = 1   # A
COL_DESC   = 2   # B
COL_UNIT   = 3   # C
COL_QTY    = 4   # D
COL_RATE   = 5   # E
COL_AMT    = 6   # F
COL_PROJ_Q = 7   # G  (user's project quantity)
COL_PROJ_A = 8   # H  (user's project amount = G × per-unit-rate)

HEADERS = ["Code","Description","Unit","Quantity","Rate","Amount","Project Qty","Project Amount"]

NUM_FMT = '#,##0.00'
PCT_FMT = '#,##0.00'


def set_cell(ws, row, col, value, font=None, fill_obj=None, align=None, num_format=None):
    c = ws.cell(row=row, column=col, value=value)
    if font:      c.font      = font
    if fill_obj:  c.fill      = fill_obj
    if align:     c.alignment = align
    if num_format: c.number_format = num_format
    c.border = thin_border()
    return c


from pathlib import Path
import json

_SPECS_PATH = Path(__file__).resolve().parents[1] / "data" / "reference_json" / "earthwork_pdf_specs.json"
if _SPECS_PATH.exists():
    with open(_SPECS_PATH, encoding="utf-8") as _f:
        PDF_SPECS = json.load(_f)
else:
    PDF_SPECS = {}



def write_item(ws, item, start_row):
    """Write one complete first-principles item table with 7 columns:
    Item Code | Labour / Machine / Material | Work done | Condition / When used | Category | Productivity | Quantity
    """
    r = start_row
    item_code = item["id"]
    base_qty = float(item.get("base_qty", 10.0))
    unit = item.get("unit", "cum")

    # 1. Header rows (Parent description & Sub-item description spanning A:G)
    hdr_lines = get_item_headers(item_code)
    for idx, line in enumerate(hdr_lines):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
        cell = ws.cell(row=r, column=1, value=line)
        cell.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        cell.fill = fill("1F4E79" if idx == 0 else "2E75B6")
        cell.alignment = AL_L
        ws.row_dimensions[r].height = max(24, min(100, (len(line) // 80 + 1) * 16))
        for col in range(2, 8):
            ws.cell(row=r, column=col).border = thin_border()
        cell.border = thin_border()
        r += 1

    # 2. Blank spacing row
    ws.row_dimensions[r].height = 6
    r += 1

    # 3. Table Column Header (7 Columns):
    headers = [
        "Item Code",
        "Labour / Machine / Material",
        "Work done",
        "Condition / When used",
        "Category",
        "Productivity",
        "Quantity",
    ]
    ws.row_dimensions[r].height = 22
    for col_idx, h_text in enumerate(headers, start=1):
        cell = ws.cell(row=r, column=col_idx, value=h_text)
        cell.font = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
        cell.fill = fill("2E75B6")
        cell.alignment = AL_C if col_idx in (1, 5, 7) else (AL_R if col_idx == 6 else AL_L)
        cell.border = thin_border()
    r += 1

    # 4. Extract resources
    labour_rows = []
    machine_rows = []
    material_rows = []
    reference_rows = []

    has_power_roller = False
    for sec_name, rows in item.get("sections", []):
        if "REFERENCE" in sec_name:
            # Process referenced base items
            for row in rows:
                ref_id = str(row[0]).strip()
                ref_desc = str(row[1]).strip()
                ref_unit = str(row[2]).strip()
                ref_qty = float(row[3])
                unit_factor = ref_qty / base_qty if base_qty > 0 else 0
                if ref_id.isdigit():
                    ref_label = ref_desc
                    spec_desc = f"Direct direct cost/material scope executed as per CPWD {ref_id} specification"
                    cat_name = "Material"
                else:
                    ref_label = f"{ref_id} ({ref_desc})"
                    spec_desc = f"Sub-item scope executed as per CPWD {ref_id} specification"
                    cat_name = "Reference"
                ref_data = (
                    ref_label,
                    f"Referenced base execution: {ref_desc}",
                    spec_desc,
                    cat_name,
                    f"{ref_qty:g} {ref_unit} ({unit_factor:.3f} {ref_unit}/{unit})",
                    f"{int(base_qty) if base_qty == int(base_qty) else base_qty} {unit}",
                )
                if cat_name == "Material":
                    material_rows.append(ref_data)
                else:
                    reference_rows.append(ref_data)
            continue

        for row in rows:
            res_code = str(row[0]).strip()
            res_name = str(row[1]).strip()
            res_unit = str(row[2]).strip()
            coeff = float(row[3])

            if res_code == "9999" or "sundries" in res_name.lower():
                continue

            sec_upper = sec_name.upper()
            sec_type = "LABOUR" if "LABOUR" in sec_upper else ("MACHINERY" if "MACHINERY" in sec_upper else ("MATERIAL" if "MATERIAL" in sec_upper else "CARRIAGE"))

            data = earthwork_teaching_knowledge.get_teaching_row_data(
                item_code=item_code,
                res_code=res_code,
                raw_name=res_name,
                section_name=sec_type,
                coeff=coeff,
                unit=res_unit,
                batch_qty=base_qty,
                batch_unit=unit,
            )
            if res_code == "0003":
                has_power_roller = True

            cat = data[3]
            if cat == "Labour":
                labour_rows.append(data)
            elif cat in ("Machine", "Equipment"):
                machine_rows.append(data)
            elif cat == "Material":
                material_rows.append(data)
            else:
                reference_rows.append(data)

    # Add 1/2-t roller for rolling items 2.2.1, 2.2.2, 2.3.1, 2.3.2
    if has_power_roller and item_code in ("2.2.1", "2.2.2", "2.3.1", "2.3.2"):
        roller_05_data = (
            "½-ton roller / wooden or steel rammers",
            "Initial compaction of each earth layer (<=20 cm depth)",
            "Compacting loose layers 1 and 2 of each 3-layer cycle; manual/light compaction; output = 1.14 cum/hr",
            "Equipment",
            "8.8 machine-hrs (1.14 cum/hr)",
            f"{int(base_qty) if base_qty == int(base_qty) else base_qty} {unit}",
        )
        machine_rows.insert(0, roller_05_data)

    # Add durmats for trench/plinth filling 2.25
    if item_code == "2.25":
        durmat_data = (
            "Wooden or steel rammers (durmats)",
            "Compacting backfilled earth layers in confined foundation trenches and plinth floors",
            "Each layer <= 20 cm; confined areas inaccessible to power rollers; output = 1.56 cum/hr",
            "Equipment",
            "6.4 machine-hrs (1.56 cum/hr)",
            f"{int(base_qty) if base_qty == int(base_qty) else base_qty} {unit}",
        )
        machine_rows.insert(0, durmat_data)

    # Composite Pipe Trench Items (2.10.1.1, 2.10.1.2, 2.10.1.3)
    if item_code in ("2.10.1.1", "2.10.1.2", "2.10.1.3"):
        trench_notes = {
            "2.10.1.1": "180 m pipe run (180m × 0.45m × 1.05m = 85.05 cum, 0.4725 cum/m)",
            "2.10.1.2": "110 m pipe run (110m × 0.60m × 1.225m + 5% = 84.89 cum, 0.7717 cum/m)",
            "2.10.1.3": "60 m pipe run (60m × 0.75m × 1.45m + 10% = 72.29 cum, 1.2048 cum/m)",
        }[item_code]

        # Standard uniform batch of 10 cum:
        # REF#2.8.1 (per 10 cum): Mate = 0.40 day (3.20 h), Coolie = 2.05 day (16.40 h), Excavator = 0.04125 day (0.33 h), Loader = 0.04125 day (0.33 h)
        # REF#2.25 (per 10 cum): Mate = 0.20 day (1.60 h), Coolie = 2.50 day (20.00 h), Bhisti = 0.20 day (1.60 h)
        # Combined: Mate = 0.60 day (4.80 h), Coolie = 4.55 day (36.40 h), Bhisti = 0.20 day (1.60 h)
        labour_rows = [
            (
                "Mate",
                "Supervising trench excavation alignment, socket pits, and backfill layer compaction (REF#2.8.1 + REF#2.25)",
                f"Combined REF#2.8.1 (0.40 day) + REF#2.25 (0.20 day) = 0.60 day/10 cum. Derived from {trench_notes}. Standardized to uniform 10 cum batch.",
                "Labour",
                "4.80 man-hrs (2.08 cum/man-hr)",
                "10 cum",
            ),
            (
                "Coolie",
                "Shifting/carrying excavated soil, disposing within 50 m lead, and returning backfill soil (REF#2.8.1 + REF#2.25)",
                f"Combined REF#2.8.1 (2.05 day) + REF#2.25 (2.50 day) = 4.55 day/10 cum. Spoil disposal & 20 cm layer backfill. Derived from {trench_notes}.",
                "Labour",
                "36.40 man-hrs (0.27 cum/man-hr)",
                "10 cum",
            ),
            (
                "Bhisti",
                "Watering backfilled soil layers to optimum moisture content for consolidation (REF#2.25)",
                f"0.20 day/10 cum (1.60 man-hrs/10 cum). Applied in layers not exceeding 20 cm depth around pipe. Derived from {trench_notes}.",
                "Labour",
                "1.60 man-hrs (6.25 cum/man-hr)",
                "10 cum",
            ),
        ]
        machine_rows = [
            (
                "Hydraulic Excavator 0.9 cum",
                "Mechanical trench cutting to required depth and profile (REF#2.8.1)",
                f"0.04125 day/10 cum (0.33 machine-hr/10 cum); bucket capacity 0.9 cum; output = 30.30 cum/hr. Derived from {trench_notes}.",
                "Machine",
                "0.33 machine-hrs (30.30 cum/hr)",
                "10 cum",
            ),
            (
                "Front End Loader 1.0 cum",
                "Loading surplus excavated trench soil into tippers for disposal (REF#2.8.1)",
                f"0.04125 day/10 cum (0.33 machine-hr/10 cum); bucket capacity 1.0 cum; output = 30.30 cum/hr. Derived from {trench_notes}.",
                "Machine",
                "0.33 machine-hrs (30.30 cum/hr)",
                "10 cum",
            ),
        ]
        reference_rows = [
            (
                "REF#2.8.1 (Pipe trench excavation in ordinary soil)",
                "Referenced base execution: Excavation pipe trench ordinary soil",
                f"Trench cut volume: 10 cum base scope executed as per CPWD REF#2.8.1 specification. Derived from {trench_notes}.",
                "Reference",
                "10.00 cum (1.000 cum/cum)",
                "10 cum",
            ),
            (
                "REF#2.25 (Filling plinth/trenches in 20cm layers, watering & ramming)",
                "Referenced base execution: Filling in trenches",
                f"Backfill volume: 10 cum base scope executed as per CPWD REF#2.25 specification. Derived from {trench_notes}.",
                "Reference",
                "10.00 cum (1.000 cum/cum)",
                "10 cum",
            ),
        ]

    # Extra Depth Pipe Trench Items in Soil (2.11 & 2.12)
    if item_code in ("2.11", "2.12"):
        trench_info = (
            "300 m pipe trench depth >1.5m to 3.0m in soil (+127.00% extra)"
            if item_code == "2.11"
            else "100 m pipe trench depth >3.0m to 4.5m in soil (+315.05% extra)"
        )
        labour_rows = [
            (
                "Mate",
                "Supervising deep trench excavation alignment, safety staging, and backfill compaction",
                f"Combined REF#2.8.1 + REF#2.6.1 + REF#2.25 + REF#2.26.1. Derived from {trench_info}. Standardized to uniform 10 cum batch.",
                "Labour",
                "5.20 man-hrs (1.92 cum/man-hr)" if item_code == "2.11" else "5.60 man-hrs (1.79 cum/man-hr)",
                "10 cum",
            ),
            (
                "Coolie",
                "Staging spoil removal from deep trench, disposal within 50m lead, and shifting backfill soil",
                f"Multi-stage vertical handling & 20 cm layer backfill. Derived from {trench_info}.",
                "Labour",
                "40.40 man-hrs (0.25 cum/man-hr)" if item_code == "2.11" else "48.40 man-hrs (0.21 cum/man-hr)",
                "10 cum",
            ),
            (
                "Bhisti",
                "Watering backfilled soil layers in deep trench to optimum moisture content (REF#2.25)",
                f"0.20 day/10 cum (1.60 man-hrs/10 cum). Applied in layers not exceeding 20 cm depth. Derived from {trench_info}.",
                "Labour",
                "1.60 man-hrs (6.25 cum/man-hr)",
                "10 cum",
            ),
        ]
        machine_rows = [
            (
                "Hydraulic Excavator 0.9 cum",
                "Mechanical deep trench cutting and bulk excavation (REF#2.8.1 + REF#2.6.1)",
                f"0.04125 day/10 cum (0.33 machine-hr/10 cum); bucket capacity 0.9 cum; output = 30.30 cum/hr. Derived from {trench_info}.",
                "Machine",
                "0.33 machine-hrs (30.30 cum/hr)",
                "10 cum",
            ),
            (
                "Front End Loader 1.0 cum",
                "Loading surplus excavated deep trench soil into tippers (REF#2.8.1)",
                f"0.04125 day/10 cum (0.33 machine-hr/10 cum); bucket capacity 1.0 cum; output = 30.30 cum/hr. Derived from {trench_info}.",
                "Machine",
                "0.33 machine-hrs (30.30 cum/hr)",
                "10 cum",
            ),
        ]
        if item_code == "2.11":
            reference_rows = [
                (
                    "REF#2.8.1 (Trench excavation in all kinds of soil)",
                    "Referenced base execution: Excavation volume 362.25 cum for 300 m pipe trench",
                    "Base trench excavation cut scope executed as per CPWD REF#2.8.1 specification",
                    "Reference",
                    "10.00 cum (1.000 cum/cum)",
                    "10 cum",
                ),
                (
                    "REF#2.6.1 (Bulk excavation over area/depth)",
                    "Referenced base execution: Bulk extra depth excavation 160.00 cum",
                    "Deep trench extra width/depth cut scope executed as per CPWD REF#2.6.1 specification",
                    "Reference",
                    "4.417 cum (0.442 cum/cum)",
                    "10 cum",
                ),
                (
                    "REF#2.25 (Trench backfilling, watering and ramming)",
                    "Referenced base execution: Trench backfill 522.25 cum in 20 cm layers",
                    "Backfill and consolidation scope executed as per CPWD REF#2.25 specification",
                    "Reference",
                    "14.417 cum (1.442 cum/cum)",
                    "10 cum",
                ),
                (
                    "REF#2.26.1 (Extra vertical lift >1.5 m to 3.0 m)",
                    "Referenced base execution: Additional lift 141.75 cum for spoil >1.5 m depth",
                    "Vertical stage lift scope executed as per CPWD REF#2.26.1 specification",
                    "Reference",
                    "3.913 cum (0.391 cum/cum)",
                    "10 cum",
                ),
                (
                    "DEDUCT: REF#2.10.1.2 (Basic pipe trench depth <= 1.5 m)",
                    "Deduction of basic 1.5 m trench cost already paid: 300 m @ basic rate",
                    "Basic 1.5 m depth rate deduction to arrive at net extra rate",
                    "Reference",
                    "-8.281 m (-0.828 m/cum)",
                    "10 cum",
                ),
                (
                    "Depth Extra Rule (+127.00% over Item 2.10.1.2)",
                    "Net extra cost calculated as +127.00% over basic pipe trench rate 2.10.1.2",
                    "Applies to qualifying pipe trench length exceeding 1.5 m but <= 3.0 m depth in all soils",
                    "Reference",
                    "+127.00% over basic",
                    "10 cum",
                ),
            ]
        else:
            reference_rows = [
                (
                    "REF#2.8.1 (Trench excavation in all kinds of soil)",
                    "Referenced base execution: Excavation volume 126.00 cum for 100 m pipe trench",
                    "Base trench cut scope executed as per CPWD REF#2.8.1 specification",
                    "Reference",
                    "10.00 cum (1.000 cum/cum)",
                    "10 cum",
                ),
                (
                    "REF#2.6.1 (Bulk excavation over area/depth)",
                    "Referenced base execution: Bulk deep cut excavation 200.00 cum",
                    "Deep trench extra cut scope executed as per CPWD REF#2.6.1 specification",
                    "Reference",
                    "15.873 cum (1.587 cum/cum)",
                    "10 cum",
                ),
                (
                    "REF#2.25 (Trench backfilling, watering and ramming)",
                    "Referenced base execution: Trench backfilling 326.00 cum in 20 cm layers",
                    "Backfill and consolidation scope executed as per CPWD REF#2.25 specification",
                    "Reference",
                    "25.873 cum (2.587 cum/cum)",
                    "10 cum",
                ),
                (
                    "REF#2.26.1 (Extra vertical lift >1.5 m to 3.0 m)",
                    "Referenced base execution: Additional lift 126.00 cum for spoil from depth > 1.5 m",
                    "Vertical stage lift scope executed as per CPWD REF#2.26.1 specification",
                    "Reference",
                    "10.000 cum (1.000 cum/cum)",
                    "10 cum",
                ),
                (
                    "DEDUCT: REF#2.10.1.2 (Basic pipe trench depth <= 1.5 m)",
                    "Deduction of basic 1.5 m trench cost already paid: 100 m @ basic rate",
                    "Basic 1.5 m depth rate deduction to arrive at net extra rate",
                    "Reference",
                    "-7.937 m (-0.794 m/cum)",
                    "10 cum",
                ),
                (
                    "Depth Extra Rule (+315.05% over Item 2.10.1.2)",
                    "Net extra cost calculated as +315.05% over basic pipe trench rate 2.10.1.2",
                    "Applies to qualifying pipe trench length exceeding 3.0 m but <= 4.5 m depth in all soils",
                    "Reference",
                    "+315.05% over basic",
                    "10 cum",
                ),
            ]

    # Ordinary Rock Pipe Trenches (2.13.1.1, 2.13.1.2, 2.13.1.3)
    if item_code in ("2.13.1.1", "2.13.1.2", "2.13.1.3"):
        trench_notes = {
            "2.13.1.1": "180 m pipe run <=80 mm dia (180m × 0.45m × 1.05m = 85.05 cum, 0.4725 cum/m)",
            "2.13.1.2": "80 m pipe run 80-300 mm dia (80m × 0.90m × 1.30m = 93.60 cum, 1.1700 cum/m)",
            "2.13.1.3": "30 m pipe run 300-600 mm dia (30m × 0.90m × 1.425m + 5% = 40.399 cum, 1.3466 cum/m)",
        }[item_code]
        labour_rows = [
            (
                "Mate",
                "Supervising ordinary rock trench excavation, socket cutting, and backfill consolidation (REF#2.25)",
                f"0.20 day/10 cum (1.60 man-hrs/10 cum). Gang supervision & safety monitoring. Derived from {trench_notes}. Standardized to uniform 10 cum batch.",
                "Labour",
                "1.60 man-hrs (6.25 cum/man-hr)",
                "10 cum",
            ),
            (
                "Rock Excavator",
                "Manual extraction and pickaxe excavation of ordinary rock strata in trench bed (REF#2.9.1)",
                f"0.885 day/10 cum (7.08 man-hrs/10 cum). Pickaxe breaking and extraction in rock trenches. Derived from {trench_notes}.",
                "Labour",
                "7.08 man-hrs (1.41 cum/man-hr)",
                "10 cum",
            ),
            (
                "Rock Breaker",
                "Sledgehammer splitting and sizing excavated rock masses into manageable rubble (REF#2.9.1)",
                f"1.765 day/10 cum (14.12 man-hrs/10 cum). Heavy hammer breaking in trench line. Derived from {trench_notes}.",
                "Labour",
                "14.12 man-hrs (0.71 cum/man-hr)",
                "10 cum",
            ),
            (
                "Rock Hole Driller",
                "Crowbar jumper drilling and splitting of ordinary rock bedding planes (REF#2.9.1)",
                f"0.53 day/10 cum (4.24 man-hrs/10 cum). Manual jumper drilling and wedging. Derived from {trench_notes}.",
                "Labour",
                "4.24 man-hrs (2.36 cum/man-hr)",
                "10 cum",
            ),
            (
                "Coolie",
                "Handling excavated rock rubble, disposal within 50m lead, and shifting backfill soil (REF#2.9.1 + REF#2.25)",
                f"Combined REF#2.9.1 (1.30 day) + REF#2.25 (2.50 day) = 3.80 day/10 cum. Derived from {trench_notes}.",
                "Labour",
                "30.40 man-hrs (0.33 cum/man-hr)",
                "10 cum",
            ),
            (
                "Beldar",
                "Trimming trench sides, squaring socket pits, and spreading backfill in 20cm layers (REF#2.9.1 + Extra dressing)",
                f"Combined REF#2.9.1 (0.50 day) + extra dressing (0.21 day) = 0.71 day/10 cum. Derived from {trench_notes}.",
                "Labour",
                "5.68 man-hrs (1.76 cum/man-hr)",
                "10 cum",
            ),
            (
                "Bhisti",
                "Watering backfilled earth layers around pipe to optimum moisture content for compaction (REF#2.25)",
                f"0.20 day/10 cum (1.60 man-hrs/10 cum). Applied in layers not exceeding 20 cm depth. Derived from {trench_notes}.",
                "Labour",
                "1.60 man-hrs (6.25 cum/man-hr)",
                "10 cum",
            ),
        ]
        machine_rows = [
            (
                "Hydraulic Excavator 0.9 cum",
                "Mechanical trench cutting in ordinary rock strata (REF#2.9.1)",
                f"0.0625 day/10 cum (0.50 machine-hr/10 cum); output = 20.00 cum/hr. Derived from {trench_notes}.",
                "Machine",
                "0.50 machine-hrs (20.00 cum/hr)",
                "10 cum",
            ),
            (
                "Tipper 10 tonne",
                "Hauling and disposing surplus unserviceable rock spoil within 50m lead (REF#2.9.1)",
                f"0.0625 day/10 cum (0.50 machine-hr/10 cum); output = 20.00 cum/hr. Derived from {trench_notes}.",
                "Machine",
                "0.50 machine-hrs (20.00 cum/hr)",
                "10 cum",
            ),
        ]
        reference_rows = [
            (
                "REF#2.9.1 (Excavation ordinary rock in trenches)",
                "Referenced base execution: Excavation ordinary rock in foundation trenches",
                f"Trench cut volume: 10 cum base scope executed as per CPWD REF#2.9.1 specification. Derived from {trench_notes}.",
                "Reference",
                "10.00 cum (1.000 cum/cum)",
                "10 cum",
            ),
            (
                "REF#2.25 (Filling plinth/trenches in 20cm layers, watering & ramming)",
                "Referenced base execution: Filling in trenches",
                f"Backfill volume: 10 cum base scope executed as per CPWD REF#2.25 specification. Derived from {trench_notes}.",
                "Reference",
                "10.00 cum (1.000 cum/cum)",
                "10 cum",
            ),
        ]

    # Hard Rock Blasting Pipe Trenches (2.13.2.1, 2.13.2.2, 2.13.2.3)
    if item_code in ("2.13.2.1", "2.13.2.2", "2.13.2.3"):
        trench_notes = {
            "2.13.2.1": "180 m pipe run <=80 mm dia (180m × 0.45m × 1.05m = 85.05 cum, 0.4725 cum/m)",
            "2.13.2.2": "80 m pipe run 80-300 mm dia (80m × 0.90m × 1.30m = 93.60 cum, 1.1700 cum/m)",
            "2.13.2.3": "30 m pipe run 300-600 mm dia (30m × 0.90m × 1.425m + 5% = 40.399 cum, 1.3466 cum/m)",
        }[item_code]
        labour_rows = [
            (
                "Mate",
                "Supervising hard rock blasting trench safety, socket profiling, and backfill consolidation (REF#2.25)",
                f"0.20 day/10 cum (1.60 man-hrs/10 cum). Safety supervision. Derived from {trench_notes}. Standardized to uniform 10 cum batch.",
                "Labour",
                "1.60 man-hrs (6.25 cum/man-hr)",
                "10 cum",
            ),
            (
                "Rock Excavator",
                "Mucking and clearing blasted hard rock fragments from trench invert (REF#2.9.2)",
                f"1.24 day/10 cum (9.92 man-hrs/10 cum). Shovel clearing of blasted rock. Derived from {trench_notes}.",
                "Labour",
                "9.92 man-hrs (1.01 cum/man-hr)",
                "10 cum",
            ),
            (
                "Rock Breaker",
                "Secondary breaking of hard rock boulders by heavy sledgehammers in trench (REF#2.9.2)",
                f"3.00 day/10 cum (24.00 man-hrs/10 cum). Sledgehammer secondary fracturing. Derived from {trench_notes}.",
                "Labour",
                "24.00 man-hrs (0.42 cum/man-hr)",
                "10 cum",
            ),
            (
                "Rock Hole Driller",
                "Operating jackhammer/pneumatic rock drills for blast hole pattern in trench line (REF#2.9.2)",
                f"0.90 day/10 cum (7.20 man-hrs/10 cum). Blast hole pattern drilling. Derived from {trench_notes}.",
                "Labour",
                "7.20 man-hrs (1.39 cum/man-hr)",
                "10 cum",
            ),
            (
                "Blaster",
                "Charging blast holes with explosives, stemming, connecting detonators, and firing shots (REF#2.9.2)",
                f"0.15 day/10 cum (1.20 man-hrs/10 cum). Licensed shot-firing operations. Derived from {trench_notes}.",
                "Labour",
                "1.20 man-hrs (8.33 cum/man-hr)",
                "10 cum",
            ),
            (
                "Coolie",
                "Carrying rock fragments, loading disposal skips, and shifting backfill soil (REF#2.9.2 + REF#2.25)",
                f"Combined REF#2.9.2 (1.50 day) + REF#2.25 (2.50 day) = 4.00 day/10 cum. Derived from {trench_notes}.",
                "Labour",
                "32.00 man-hrs (0.31 cum/man-hr)",
                "10 cum",
            ),
            (
                "Beldar",
                "Trench side dressing, socket trimming, and extra ramming of rock backfill (REF#2.9.2 + Extra ramming)",
                f"Combined REF#2.9.2 (0.60 day) + extra ramming (0.29 day) = 0.89 day/10 cum. Derived from {trench_notes}.",
                "Labour",
                "7.12 man-hrs (1.40 cum/man-hr)",
                "10 cum",
            ),
            (
                "Bhisti",
                "Watering backfilled earth/soil layers around pipe to optimum moisture content for compaction (REF#2.25)",
                f"0.20 day/10 cum (1.60 man-hrs/10 cum). Applied in layers not exceeding 20 cm depth. Derived from {trench_notes}.",
                "Labour",
                "1.60 man-hrs (6.25 cum/man-hr)",
                "10 cum",
            ),
        ]
        machine_rows = [
            (
                "Hydraulic Excavator 0.9 cum",
                "Mucking and loading blasted hard rock trench debris (REF#2.9.2)",
                f"0.125 day/10 cum (1.00 machine-hr/10 cum); output = 10.00 cum/hr. Derived from {trench_notes}.",
                "Machine",
                "1.00 machine-hrs (10.00 cum/hr)",
                "10 cum",
            ),
            (
                "Tipper 10 tonne",
                "Hauling surplus blasted rock spoil to designated dumping ground within 50m lead (REF#2.9.2)",
                f"0.125 day/10 cum (1.00 machine-hr/10 cum); output = 10.00 cum/hr. Derived from {trench_notes}.",
                "Machine",
                "1.00 machine-hrs (10.00 cum/hr)",
                "10 cum",
            ),
        ]
        reference_rows = [
            (
                "REF#2.9.2 (Excavation hard rock requiring blasting in trenches)",
                "Referenced base execution: Excavation hard rock requiring blasting in foundation trenches",
                f"Trench cut volume: 10 cum base scope executed as per CPWD REF#2.9.2 specification. Derived from {trench_notes}.",
                "Reference",
                "10.00 cum (1.000 cum/cum)",
                "10 cum",
            ),
            (
                "REF#2.25 (Filling plinth/trenches in 20cm layers, watering & ramming)",
                "Referenced base execution: Filling in trenches",
                f"Backfill volume: 10 cum base scope executed as per CPWD REF#2.25 specification. Derived from {trench_notes}.",
                "Reference",
                "10.00 cum (1.000 cum/cum)",
                "10 cum",
            ),
        ]

    # Hard Rock No Blasting Pipe Trenches (2.13.3.1, 2.13.3.2, 2.13.3.3)
    if item_code in ("2.13.3.1", "2.13.3.2", "2.13.3.3"):
        trench_notes = {
            "2.13.3.1": "180 m pipe run <=80 mm dia (180m × 0.45m × 1.05m = 85.05 cum, 0.4725 cum/m)",
            "2.13.3.2": "80 m pipe run 80-300 mm dia (80m × 0.90m × 1.30m = 93.60 cum, 1.1700 cum/m)",
            "2.13.3.3": "30 m pipe run 300-600 mm dia (30m × 0.90m × 1.425m + 5% = 40.399 cum, 1.3466 cum/m)",
        }[item_code]
        labour_rows = [
            (
                "Mate",
                "Supervising hard rock chiselling/splitting in built-up area, socket profiling, and backfill consolidation (REF#2.25)",
                f"0.20 day/10 cum (1.60 man-hrs/10 cum). Alignment and safety supervision. Derived from {trench_notes}. Standardized to uniform 10 cum batch.",
                "Labour",
                "1.60 man-hrs (6.25 cum/man-hr)",
                "10 cum",
            ),
            (
                "Rock Excavator",
                "Non-blast excavation, wedging and prying laminated hard rock in trench bed (REF#2.9.3)",
                f"2.65 day/10 cum (21.20 man-hrs/10 cum). Heavy mechanical wedging and prying. Derived from {trench_notes}.",
                "Labour",
                "21.20 man-hrs (0.47 cum/man-hr)",
                "10 cum",
            ),
            (
                "Rock Breaker",
                "Manual fracturing of hard rock using heavy moils, points and sledgehammers (REF#2.9.3)",
                f"6.175 day/10 cum (49.40 man-hrs/10 cum). Intensive manual splitting without explosives. Derived from {trench_notes}.",
                "Labour",
                "49.40 man-hrs (0.20 cum/man-hr)",
                "10 cum",
            ),
            (
                "Stone Chiseller",
                "Chiselling rock trench sides vertically and dressing bottom true to gradient (REF#2.9.3)",
                f"1.06 day/10 cum (8.48 man-hrs/10 cum). Precision chiselling of rock trench walls. Derived from {trench_notes}.",
                "Labour",
                "8.48 man-hrs (1.18 cum/man-hr)",
                "10 cum",
            ),
            (
                "Blacksmith 2nd class",
                "On-site sharpening, tempering and re-forging of chisels, moil points and crowbars (REF#2.9.3)",
                f"0.175 day/10 cum (1.40 man-hrs/10 cum). Tool maintenance and bit re-sharpening. Derived from {trench_notes}.",
                "Labour",
                "1.40 man-hrs (7.14 cum/man-hr)",
                "10 cum",
            ),
            (
                "Coolie",
                "Carrying rock chips, loading disposal tippers, and shifting backfill soil (REF#2.9.3 + REF#2.25)",
                f"Combined REF#2.9.3 (1.50 day) + REF#2.25 (2.50 day) = 4.00 day/10 cum. Derived from {trench_notes}.",
                "Labour",
                "32.00 man-hrs (0.31 cum/man-hr)",
                "10 cum",
            ),
            (
                "Beldar",
                "Trench side clearance, socket dressing, and extra ramming of rock backfill (REF#2.9.3 + Extra ramming)",
                f"Combined REF#2.9.3 (0.75 day) + extra ramming (0.29 day) = 1.04 day/10 cum. Derived from {trench_notes}.",
                "Labour",
                "8.35 man-hrs (1.20 cum/man-hr)",
                "10 cum",
            ),
            (
                "Bhisti",
                "Watering backfilled earth/soil layers around pipe to optimum moisture content for compaction (REF#2.25)",
                f"0.20 day/10 cum (1.60 man-hrs/10 cum). Applied in layers not exceeding 20 cm depth. Derived from {trench_notes}.",
                "Labour",
                "1.60 man-hrs (6.25 cum/man-hr)",
                "10 cum",
            ),
        ]
        machine_rows = [
            (
                "Hydraulic Excavator 0.9 cum",
                "Mechanical excavation assist and removal of fractured rock boulders (REF#2.9.3)",
                f"0.125 day/10 cum (1.00 machine-hr/10 cum); output = 10.00 cum/hr. Derived from {trench_notes}.",
                "Machine",
                "1.00 machine-hrs (10.00 cum/hr)",
                "10 cum",
            ),
            (
                "Tipper 10 tonne",
                "Hauling surplus rock rubble to designated disposal heaps within 50m lead (REF#2.9.3)",
                f"0.125 day/10 cum (1.00 machine-hr/10 cum); output = 10.00 cum/hr. Derived from {trench_notes}.",
                "Machine",
                "1.00 machine-hrs (10.00 cum/hr)",
                "10 cum",
            ),
        ]
        reference_rows = [
            (
                "REF#2.9.3 (Excavation hard rock no blasting in trenches)",
                "Referenced base execution: Excavation hard rock without blasting in foundation trenches",
                f"Trench cut volume: 10 cum base scope executed as per CPWD REF#2.9.3 specification. Derived from {trench_notes}.",
                "Reference",
                "10.00 cum (1.000 cum/cum)",
                "10 cum",
            ),
            (
                "REF#2.25 (Filling plinth/trenches in 20cm layers, watering & ramming)",
                "Referenced base execution: Filling in trenches",
                f"Backfill volume: 10 cum base scope executed as per CPWD REF#2.25 specification. Derived from {trench_notes}.",
                "Reference",
                "10.00 cum (1.000 cum/cum)",
                "10 cum",
            ),
        ]

    # Rock Pipe Trench Depth Extras (2.14 & 2.15)
    if item_code in ("2.14", "2.15"):
        trench_info = (
            "300 m pipe trench depth >1.5m to 3.0m in ordinary/hard rock (+103.60% extra)"
            if item_code == "2.14"
            else "100 m pipe trench depth >3.0m to 4.5m in ordinary/hard rock (+255.60% extra)"
        )
        labour_rows = [
            (
                "Mate",
                "Supervising deep rock trench blasting, stage scaffolding safety, and backfill consolidation",
                f"Combined REF#2.9.2 + REF#2.7.2 + REF#2.25 + REF#2.26.2. Derived from {trench_info}. Standardized to uniform 10 cum batch.",
                "Labour",
                "5.20 man-hrs (1.92 cum/man-hr)",
                "10 cum",
            ),
            (
                "Rock Excavator",
                "Deep trench rock extraction and clearing blasted rock from lower bench (REF#2.9.2 + REF#2.7.2)",
                f"Mucking and extraction in deep rock trench. Derived from {trench_info}.",
                "Labour",
                "9.92 man-hrs (1.01 cum/man-hr)",
                "10 cum",
            ),
            (
                "Rock Breaker",
                "Secondary breaking of hard rock masses in deep trench confines (REF#2.9.2 + REF#2.7.2)",
                f"Heavy sledgehammer fragmentation. Derived from {trench_info}.",
                "Labour",
                "24.00 man-hrs (0.42 cum/man-hr)",
                "10 cum",
            ),
            (
                "Rock Hole Driller",
                "Drilling deep blast hole pattern in trench rock bed using crawler/pneumatic drills (REF#2.9.2)",
                f"Blast hole drilling at depth > 1.5 m. Derived from {trench_info}.",
                "Labour",
                "7.20 man-hrs (1.39 cum/man-hr)",
                "10 cum",
            ),
            (
                "Coolie",
                "Multi-stage vertical lifting of rock debris, disposal within 50m lead, and shifting backfill soil",
                f"Staging rock debris from deep trench and layer backfill. Derived from {trench_info}.",
                "Labour",
                "44.00 man-hrs (0.23 cum/man-hr)",
                "10 cum",
            ),
            (
                "Beldar",
                "Trench bottom leveling, socket pit excavation in rock, and ramming backfilled layers",
                f"Deep rock bottom preparation and layer ramming. Derived from {trench_info}.",
                "Labour",
                "7.60 man-hrs (1.32 cum/man-hr)",
                "10 cum",
            ),
            (
                "Bhisti",
                "Watering backfilled earth/soil layers in deep trench around pipe (REF#2.25)",
                f"0.20 day/10 cum (1.60 man-hrs/10 cum). Applied in layers not exceeding 20 cm depth. Derived from {trench_info}.",
                "Labour",
                "1.60 man-hrs (6.25 cum/man-hr)",
                "10 cum",
            ),
        ]
        machine_rows = [
            (
                "Hydraulic Excavator 0.9 cum",
                "Mechanical rock excavation and deep mucking (REF#2.9.2 + REF#2.7.2)",
                f"0.125 day/10 cum (1.00 machine-hr/10 cum); output = 10.00 cum/hr. Derived from {trench_info}.",
                "Machine",
                "1.00 machine-hrs (10.00 cum/hr)",
                "10 cum",
            ),
            (
                "Tipper 10 tonne",
                "Hauling surplus rock rubble from deep trench to dumping heaps (REF#2.9.2)",
                f"0.125 day/10 cum (1.00 machine-hr/10 cum); output = 10.00 cum/hr. Derived from {trench_info}.",
                "Machine",
                "1.00 machine-hrs (10.00 cum/hr)",
                "10 cum",
            ),
        ]
        if item_code == "2.14":
            reference_rows = [
                (
                    "REF#2.9.2 (Excavation hard rock requiring blasting in trenches)",
                    "Referenced base execution: Excavation volume 517.50 cum for 300 m rock trench",
                    "Base rock trench excavation scope executed as per CPWD REF#2.9.2 specification",
                    "Reference",
                    "10.00 cum (1.000 cum/cum)",
                    "10 cum",
                ),
                (
                    "REF#2.7.2 (Deep rock excavation over area)",
                    "Referenced base execution: Extra width/depth rock excavation 190.00 cum",
                    "Bulk deep rock cut scope executed as per CPWD REF#2.7.2 specification",
                    "Reference",
                    "3.671 cum (0.367 cum/cum)",
                    "10 cum",
                ),
                (
                    "REF#2.25 (Trench backfilling, watering and ramming)",
                    "Referenced base execution: Trench backfill 707.50 cum in 20 cm layers",
                    "Backfill and consolidation scope executed as per CPWD REF#2.25 specification",
                    "Reference",
                    "13.671 cum (1.367 cum/cum)",
                    "10 cum",
                ),
                (
                    "REF#2.26.2 (Extra vertical lift >1.5 m to 3.0 m in rock)",
                    "Referenced base execution: Additional lift 202.50 cum for rock spoil >1.5 m depth",
                    "Vertical stage lift scope executed as per CPWD REF#2.26.2 specification",
                    "Reference",
                    "3.913 cum (0.391 cum/cum)",
                    "10 cum",
                ),
                (
                    "DEDUCT: REF#2.13.2.2 (Basic rock pipe trench depth <= 1.5 m)",
                    "Deduction of basic 1.5 m rock trench cost already paid: 300 m @ basic rate",
                    "Basic 1.5 m depth rock rate deduction to arrive at net extra rate",
                    "Reference",
                    "-5.797 m (-0.580 m/cum)",
                    "10 cum",
                ),
                (
                    "Depth Extra Rule (+103.60% over Item 2.13 basic)",
                    "Net extra cost calculated as +103.60% over basic rock pipe trench rate",
                    "Applies to qualifying pipe trench length exceeding 1.5 m but <= 3.0 m depth in rock",
                    "Reference",
                    "+103.60% over basic",
                    "10 cum",
                ),
            ]
        else:
            reference_rows = [
                (
                    "REF#2.9.2 (Excavation hard rock requiring blasting in trenches)",
                    "Referenced base execution: Excavation volume 180.00 cum for 100 m rock trench",
                    "Base rock trench excavation scope executed as per CPWD REF#2.9.2 specification",
                    "Reference",
                    "10.00 cum (1.000 cum/cum)",
                    "10 cum",
                ),
                (
                    "REF#2.7.2 (Deep rock excavation over area)",
                    "Referenced base execution: Extra width/depth rock excavation 237.50 cum",
                    "Bulk deep rock cut scope executed as per CPWD REF#2.7.2 specification",
                    "Reference",
                    "13.194 cum (1.319 cum/cum)",
                    "10 cum",
                ),
                (
                    "REF#2.25 (Trench backfilling, watering and ramming)",
                    "Referenced base execution: Trench backfilling 417.50 cum in 20 cm layers",
                    "Backfill and consolidation scope executed as per CPWD REF#2.25 specification",
                    "Reference",
                    "23.194 cum (2.319 cum/cum)",
                    "10 cum",
                ),
                (
                    "REF#2.26.2 (Extra vertical lift >3.0 m to 4.5 m in rock)",
                    "Referenced base execution: Additional lift 180.00 cum for rock spoil >3.0 m depth",
                    "Vertical stage lift scope executed as per CPWD REF#2.26.2 specification",
                    "Reference",
                    "10.000 cum (1.000 cum/cum)",
                    "10 cum",
                ),
                (
                    "DEDUCT: REF#2.13.2.2 (Basic rock pipe trench depth <= 1.5 m)",
                    "Deduction of basic rock trench cost up to 1.5 m depth already accounted for: 100 m @ basic rate",
                    "Basic 1.5 m depth rock rate deduction to arrive at net extra rate",
                    "Reference",
                    "-5.556 m (-0.556 m/cum)",
                    "10 cum",
                ),
                (
                    "Depth Extra Rule (+255.60% over Item 2.13 basic)",
                    "Net extra cost calculated as +255.60% over basic rock pipe trench rate",
                    "Applies to qualifying pipe trench length exceeding 3.0 m but <= 4.5 m depth in rock",
                    "Reference",
                    "+255.60% over basic",
                    "10 cum",
                ),
            ]

    # Mechanical Earth Filling (2.25(a))
    if item_code == "2.25(a)":
        labour_rows = [
            (
                "Beldar",
                "Leveling mechanically dumped local earth in plinth/trenches into uniform layers not exceeding 20 cm depth",
                "Spreading and dressing dumped earth into 20 cm layers in plinth, foundation sides and trenches",
                "Labour",
                "3.60 man-hrs (2.78 cum/man-hr)",
                "10 cum",
            ),
            (
                "Bhisti",
                "Watering each layer of mechanically transported earth to optimum moisture content before ramming",
                "Spraying water uniformly over 20 cm loose layers for compaction conditioning",
                "Labour",
                "2.80 man-hrs (3.57 cum/man-hr)",
                "10 cum",
            ),
        ]
        machine_rows = [
            (
                "Hydraulic Excavator 0.9 cum",
                "Mechanical excavation of borrow earth at source pit (REF#2.6.1)",
                "0.04125 day/10 cum (0.33 machine-hr/10 cum); output = 30.30 cum/hr",
                "Machine",
                "0.33 machine-hrs (30.30 cum/hr)",
                "10 cum",
            ),
            (
                "Front End Loader 1.0 cum",
                "Loading excavated borrow earth into tippers for haulage (REF#2.6.1)",
                "0.04125 day/10 cum (0.33 machine-hr/10 cum); output = 30.30 cum/hr",
                "Machine",
                "0.33 machine-hrs (30.30 cum/hr)",
                "10 cum",
            ),
            (
                "Tipper 10 tonne",
                "Mechanical transport and carriage of local earth from borrow pit to site up to 5 km lead (REF#1.1.2)",
                "Tipper haulage cycle over 5 km lead: 10 cum batch payload",
                "Machine",
                "10.00 cum (1.000 cum/cum)",
                "10 cum",
            ),
            (
                "Wooden or steel rammers (durmats)",
                "Compacting spread earth layers in trenches, plinth and foundation sides",
                "Consolidating 20 cm layers in confined plinth/trenches where power rollers cannot operate; output = 1.56 cum/hr",
                "Equipment",
                "6.40 machine-hrs (1.56 cum/hr)",
                "10 cum",
            ),
        ]
        material_rows = [
            (
                "Royalty on earth",
                "Statutory royalty paid on excavated local borrow earth",
                "Royalty rate per cum of borrow earth as per local state mining department norms",
                "Material",
                "10.00 cum (1.000 cum/cum)",
                "10 cum",
            ),
        ]
        reference_rows = [
            (
                "REF#2.6.1 (Hydraulic excavation ordinary soil)",
                "Referenced base execution: Borrow pit mechanical excavation 10 cum",
                "Excavation at borrow source executed as per CPWD REF#2.6.1 specification",
                "Reference",
                "10.00 cum (1.000 cum/cum)",
                "10 cum",
            ),
            (
                "REF#1.1.2 (Carriage of earth up to 5 km lead)",
                "Referenced base execution: Mechanical carriage by tipper/truck over 5 km lead",
                "Haulage scope executed as per CPWD Carriage Item 1.1.2 specification",
                "Reference",
                "10.00 cum (1.000 cum/cum)",
                "10 cum",
            ),
            (
                "DEDUCT: REF#1.2.2 (Dozing and spreading omitted)",
                "Omission of mechanical dozer spreading at borrow pit since spreading is done manually on site",
                "Rate deduction as per CPWD Sub-Head 1 specification",
                "Reference",
                "-10.00 cum (-1.000 cum/cum)",
                "10 cum",
            ),
        ]

    # Handle percentage extra items (like 2.24.1, 2.24.2)
    if item.get("conditional_extra") or item_code in ("2.24.1", "2.24.2"):
        pct_rule = item.get("conditional_rule", "+20% over base item rate")
        cond_text = (
            "Quantities of work executed in or under water and/or liquid mud; depth to C.G. (conditional % extra, not a flat Rs rate)"
            if item_code == "2.24.1"
            else "Quantities of work executed in or under foul position; depth to C.G. (conditional % extra, not a flat Rs rate)"
        )
        work_text = (
            "Dewatering, extra handling, and difficulty allowance in wet/mud conditions"
            if item_code == "2.24.1"
            else "Extra handling and foul conditions execution allowance"
        )
        extra_row = (
            "Pumping equipment & labour crew",
            work_text,
            cond_text,
            "Equipment",
            pct_rule,
            f"Qualifying {unit}",
        )
        machine_rows.append(extra_row)

    # Sequence: Labour first, then Machine/Equipment, then Material, then Reference
    all_data_rows = labour_rows + machine_rows + material_rows + reference_rows

    for row_idx, data_row in enumerate(all_data_rows):
        disp_name, work_done, condition, category, prod_str, qty_str = data_row
        row_fill = fill("F9FBFD" if row_idx % 2 == 1 else "FFFFFF")

        max_len = max(len(work_done), len(condition))
        ws.row_dimensions[r].height = max(20, min(85, (max_len // 45 + 1) * 14))

        # Col A: Item Code
        c1 = ws.cell(row=r, column=1, value=item_code)
        c1.font = Font(name="Calibri", size=9, bold=True)
        c1.alignment = AL_C
        c1.fill = row_fill
        c1.border = thin_border()

        # Col B: Labour / Machine / Material
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
        c7 = ws.cell(row=r, column=7, value=qty_str)
        c7.font = Font(name="Calibri", size=9)
        c7.alignment = AL_C
        c7.fill = row_fill
        c7.border = thin_border()

        r += 1

    # Blank row separating items
    ws.row_dimensions[r].height = 12
    r += 1
    return r


def build_sheet(ws):
    # Sheet title spanning A1:G1
    ws.merge_cells("A1:G1")
    t1 = ws.cell(row=1, column=1, value="Sub-Head 2.0 — EARTH WORK  |  First-Principles Resource, Work & Gang Analysis")
    t1.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    t1.fill = fill("1A1A2E")
    t1.alignment = AL_C
    ws.row_dimensions[1].height = 28

    # Subtitle spanning A2:G2
    ws.merge_cells("A2:G2")
    t2 = ws.cell(row=2, column=1, value="Evidence: CPWD DAR 2019 Vol 1  |  Format: Item Code • Labour / Machine / Material • Work done • Condition / When used • Category • Productivity • Standard Batch Quantity")
    t2.font = Font(name="Calibri", size=9, italic=True, color="333333")
    t2.fill = fill("E8F4FD")
    t2.alignment = AL_L
    ws.row_dimensions[2].height = 18

    current_row = 4
    for item in ITEMS:
        current_row = write_item(ws, item, current_row)

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


def main():
    WB_PATH = WB_VOL1_LATEST_FILE
    TMP_PATH = WB_PATH.replace(".xlsx", "_REBUILDING.xlsx")

    print("Loading workbook …")
    wb = openpyxl.load_workbook(WB_PATH)

    # Remove old sheet if exists
    if "02_support_earth_work" in wb.sheetnames:
        del wb["02_support_earth_work"]

    # Insert after 02_Earth_Work
    target_pos = 2
    if "02_Earth_Work" in wb.sheetnames:
        target_pos = wb.sheetnames.index("02_Earth_Work") + 1

    ws = wb.create_sheet("02_support_earth_work", target_pos)
    ws.sheet_properties.tabColor = "2E75B6"

    print("Building rate analysis sheet …")
    build_sheet(ws)

    print(f"Saving to temp: {TMP_PATH}")
    wb.save(TMP_PATH)
    print("Saved temp. Replacing original …")
    os.replace(TMP_PATH, WB_PATH)
    print(f"Done -> {WB_PATH}")
    print(f"Items written: {len(ITEMS)}")


if __name__ == "__main__":
    main()
