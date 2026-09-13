"""
support_builder_earth_v2.py
Rebuilds 02_support_earth_work sheet with FULL CPWD rate analysis format.
Each item mirrors CivilDAR_2019_Vol_1.pdf exactly:
  Code | Description | Unit | Quantity | Rate | Amount(=D*E)
  Sections: MACHINERY / LABOUR / MATERIAL / CARRIAGE
  Chain: W → Water(1%) → X → GST(14.05%) → Y → CP&OH(15%) → Z → Cess(1%) → Total → Per unit

Formula constraint: only SUM, +, -, *, / (no VLOOKUP/INDEX/MATCH)
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os, sys, shutil
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir)))
from scripts.paths import WB_VOL1_LATEST_FILE

# ── Colours ──────────────────────────────────────────────────────────────────
C_ITEM_HDR   = "1F4E79"   # dark blue  – item heading
C_SEC_HDR    = "2E75B6"   # mid blue   – MACHINERY/LABOUR/MATERIAL/CARRIAGE
C_CHAIN_LBL  = "FFF2CC"   # pale yellow – W / X / Y / Z labels
C_CHAIN_VAL  = "FFFACD"   # lighter yellow
C_REF_ROW    = "E2EFDA"   # pale green  – referenced sub-item rows
C_FINAL      = "F4B942"   # amber       – per-unit rate row
C_DSR        = "D9D9D9"   # grey        – DSR cross-ref
C_WHITE      = "FFFFFF"
C_BLANK      = None

FT_HDR  = Font(bold=True, color="FFFFFF", size=10)
FT_SEC  = Font(bold=True, color="FFFFFF", size=9)
FT_BODY = Font(size=9)
FT_BOLD = Font(bold=True, size=9)
FT_CHAIN= Font(bold=True, size=9, italic=True)

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

# ── Item data ─────────────────────────────────────────────────────────────────
# Structure per item:
#   "id"       : "2.1.1"
#   "desc"     : full description
#   "unit"     : "sqm" / "cum" etc.
#   "base_qty" : numeric base for the analysis (e.g. 100, 10, 1)
#   "dsr_rate" : DSR 2021 final rate (for cross-reference)
#   "sections" : list of (section_name, rows)
#              : each row = (code, description, unit, qty, rate)  — amount = qty*rate
#              : code="" for direct resource rows; code="REF" for cross-ref rows (marked A)
#   "notes"    : optional string appended after the chain

ITEMS = [

# ─── 2.1 SURFACE DRESSING ────────────────────────────────────────────────────
{
 "id":"2.1.1","desc":"Surface dressing of ground including removing of vegetation, grass, bushes, shrubs, rank weeds, etc., and disposal of the same upto 50 metres lead - in ordinary soil",
 "unit":"sqm","base_qty":100,"dsr_rate":107.00,
 "sections":[
  ("LABOUR",[
   ("0114","Beldar","day",6.80,558.00),
   ("0115","Coolie","day",5.60,558.00),
  ]),
 ],
},
{
 "id":"2.1.2","desc":"Surface dressing of ground including removing of vegetation, grass, bushes, shrubs, rank weeds, etc., and disposal of the same upto 50 metres lead - in hard soil",
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
 "id":"2.2.1","desc":"Earthwork in excavation by mechanical means (Hydraulic excavator)/manual means in foundation trenches or drains (not exceeding 1.5 m in width or 10 sqm on plan) including dressing of sides and ramming of bottoms, lift upto 1.5 m, including getting out the excavated soil and disposal of surplus excavated soil as directed, within a lead of 50 m - in ordinary soil",
 "unit":"cum","base_qty":10,"dsr_rate":862.70,
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
{
 "id":"2.2.2","desc":"Earthwork in excavation in hard soil",
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
 "id":"2.3.1","desc":"Earthwork in excavation by mechanical means/manual means over areas (exceeding 30 cm in depth, 1.5 m in width and 10 sqm on plan) including dressing of sides and ramming of bottoms, lift upto 1.5 m, disposal of surplus within 50 m lead - in ordinary soil",
 "unit":"cum","base_qty":10,"dsr_rate":543.40,
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
{
 "id":"2.3.2","desc":"Earthwork in excavation over areas - in hard soil",
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
 "id":"2.4","desc":"Deduct for not using road roller and not providing Chowkidar when excavation is done manually",
 "unit":"cum","base_qty":10,"dsr_rate":4.40,
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
 "id":"2.5","desc":"Deduct for not providing Bhisti in dry weather",
 "unit":"cum","base_qty":10,"dsr_rate":38.20,
 "sections":[
  ("LABOUR",[
   ("0101","Bhisti","day",0.40,617.00),
  ]),
 ],
 "notes":"DEDUCT item",
},

# ─── 2.6 HYDRAULIC EXCAVATION ─────────────────────────────────────────────────
{
 "id":"2.6.1","desc":"Excavation for foundation in trenches/drains by hydraulic excavator, lead up to 50 m, lift up to 1.5 m - in ordinary soil",
 "unit":"cum","base_qty":10,"dsr_rate":205.45,
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
{
 "id":"2.6.2","desc":"Excavation for foundation in trenches/drains by hydraulic excavator - in hard soil",
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
 "id":"2.7.1","desc":"Excavation for foundation in trenches/drains in ordinary rock by blasting - all lifts",
 "unit":"cum","base_qty":10,"dsr_rate":412.95,
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
 "id":"2.7.2","desc":"Excavation for foundation in hard rock by blasting",
 "unit":"cum","base_qty":10,"dsr_rate":711.35,
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
 "id":"2.7.3","desc":"Excavation for foundation in hard rock where blasting is prohibited - by chiselling",
 "unit":"cum","base_qty":10,"dsr_rate":1184.30,
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
 "id":"2.8.1","desc":"Earthwork in excavation by mechanical means/manual means in pipe trenches including dressing of sides and ramming of bottom, lift upto 1.5 m, disposal of surplus within 50 m - in ordinary soil",
 "unit":"cum","base_qty":10,"dsr_rate":286.85,
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
{
 "id":"2.8.2","desc":"Earthwork in excavation in pipe trenches - in hard soil",
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
 "id":"2.9.1","desc":"Earthwork in excavation in pipe trenches in ordinary rock by blasting",
 "unit":"cum","base_qty":10,"dsr_rate":523.50,
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
 "id":"2.9.2","desc":"Earthwork in excavation in pipe trenches in hard rock by blasting",
 "unit":"cum","base_qty":10,"dsr_rate":846.25,
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
 "id":"2.9.3","desc":"Earthwork in excavation in pipe trenches in hard rock where blasting is prohibited",
 "unit":"cum","base_qty":10,"dsr_rate":1258.60,
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
 "id":"2.10.1.1","desc":"Earthwork in excavation in pipe trenches, dia 150 mm, depth up to 1.5 m - ordinary soil (composite: excav 2.8.1 + filling 2.25)",
 "unit":"metre","base_qty":180,"dsr_rate":255.55,
 "ref_pattern":True,
 "sections":[
  ("REFERENCE ITEMS (excluded from multiplier chain)",[
   ("REF 2.8.1","Excavation pipe trench ordinary soil (85.05 cum)","cum",85.05,252.30),
   ("REF 2.25","Filling in trenches (85.05 cum)","cum",85.05,219.65),
  ]),
 ],
 "notes":"W-A pattern: entire cost = sum of referenced items × base qty. No separate Water/GST/CPOH applied on top.",
},
{
 "id":"2.10.1.2","desc":"Pipe trench dia 150 mm, depth 1.5–3 m - ordinary soil",
 "unit":"metre","base_qty":110,"dsr_rate":417.35,
 "ref_pattern":True,
 "sections":[
  ("REFERENCE ITEMS",[
   ("REF 2.8.1","Excavation pipe trench ordinary soil (84.89 cum)","cum",84.89,252.30),
   ("REF 2.25","Filling in trenches (84.89 cum)","cum",84.89,219.65),
  ]),
 ],
},
{
 "id":"2.10.1.3","desc":"Pipe trench dia 150 mm, depth 3–4.5 m - ordinary soil",
 "unit":"metre","base_qty":60,"dsr_rate":651.55,
 "ref_pattern":True,
 "sections":[
  ("REFERENCE ITEMS",[
   ("REF 2.8.1","Excavation pipe trench ordinary soil (72.29 cum)","cum",72.29,252.30),
   ("REF 2.25","Filling in trenches (72.29 cum)","cum",72.29,219.65),
  ]),
 ],
},

# ─── 2.11 EXTRA DEPTH (%) ─────────────────────────────────────────────────────
{
 "id":"2.11","desc":"Extra over item 2.10 for excavation and filling in pipe trenches for additional depth beyond 1.5 m upto 3 m - % of item 2.10",
 "unit":"%","base_qty":1,"dsr_rate":127.00,
 "pct_item":True,
 "sections":[],
 "notes":"Percentage-based extra: 127% over base rate of Item 2.10. No resource rows. DAR computed value = 127.00%.",
},
{
 "id":"2.12","desc":"Extra over item 2.10 for additional depth beyond 3 m - % of item 2.10",
 "unit":"%","base_qty":1,"dsr_rate":314.95,
 "pct_item":True,
 "sections":[],
 "notes":"Percentage-based extra: 315.05% over base rate of Item 2.10. DAR computed value = 315.05% (DSR 2021: 314.95%).",
},

# ─── 2.13 ROCK PIPE TRENCHES (COMPOSITE) ─────────────────────────────────────
{
 "id":"2.13.1.1","desc":"Pipe trench in ordinary rock, dia 150 mm, depth up to 1.5 m",
 "unit":"metre","base_qty":180,"dsr_rate":376.95,
 "ref_pattern":True,
 "sections":[
  ("REFERENCE ITEMS (A - excluded from multipliers)",[
   ("REF 2.9.1","Excavation ordinary rock (85.05 cum)","cum",85.05,448.15),
   ("REF 2.25","Filling (85.05 cum)","cum",85.05,219.65),
  ]),
  ("LABOUR (subject to Water/GST/CPOH/Cess)",[
   ("0114","Beldar (extra for rock dressing)","day",1.80,558.00),
  ]),
 ],
},
{
 "id":"2.13.1.2","desc":"Pipe trench in hard rock by blasting, dia 150 mm, depth up to 1.5 m",
 "unit":"metre","base_qty":180,"dsr_rate":933.35,
 "ref_pattern":True,
 "sections":[
  ("REFERENCE ITEMS (A)",[
   ("REF 2.9.2","Excavation hard rock blasting (85.05 cum)","cum",85.05,729.00),
   ("REF 2.25","Filling (85.05 cum)","cum",85.05,219.65),
  ]),
  ("LABOUR",[
   ("0114","Beldar (extra dressing)","day",1.80,558.00),
  ]),
 ],
},
{
 "id":"2.13.1.3","desc":"Pipe trench in hard rock blasting prohibited, dia 150 mm, depth up to 1.5 m",
 "unit":"metre","base_qty":180,"dsr_rate":1074.00,
 "ref_pattern":True,
 "sections":[
  ("REFERENCE ITEMS (A)",[
   ("REF 2.9.3","Excavation hard rock no blasting (85.05 cum)","cum",85.05,1080.55),
   ("REF 2.25","Filling (85.05 cum)","cum",85.05,219.65),
  ]),
  ("LABOUR",[
   ("0114","Beldar (extra dressing)","day",1.80,558.00),
  ]),
 ],
},
{
 "id":"2.13.2.1","desc":"Pipe trench ordinary rock, dia 150 mm, depth 1.5–3 m",
 "unit":"metre","base_qty":110,"dsr_rate":531.85,
 "ref_pattern":True,
 "sections":[
  ("REFERENCE ITEMS (A)",[
   ("REF 2.9.1","Excavation ordinary rock (84.89 cum)","cum",84.89,448.15),
   ("REF 2.25","Filling (84.89 cum)","cum",84.89,219.65),
  ]),
  ("LABOUR",[
   ("0114","Beldar","day",2.20,558.00),
  ]),
 ],
},
{
 "id":"2.13.2.2","desc":"Pipe trench hard rock blasting, dia 150 mm, depth 1.5–3 m",
 "unit":"metre","base_qty":110,"dsr_rate":1316.90,
 "ref_pattern":True,
 "sections":[
  ("REFERENCE ITEMS (A)",[
   ("REF 2.9.2","Excavation hard rock (84.89 cum)","cum",84.89,729.00),
   ("REF 2.25","Filling (84.89 cum)","cum",84.89,219.65),
  ]),
  ("LABOUR",[
   ("0114","Beldar","day",2.20,558.00),
  ]),
 ],
},
{
 "id":"2.13.2.3","desc":"Pipe trench hard rock no blasting, dia 150 mm, depth 1.5–3 m",
 "unit":"metre","base_qty":110,"dsr_rate":1515.20,
 "ref_pattern":True,
 "sections":[
  ("REFERENCE ITEMS (A)",[
   ("REF 2.9.3","Excavation hard rock no blasting (84.89 cum)","cum",84.89,1080.55),
   ("REF 2.25","Filling (84.89 cum)","cum",84.89,219.65),
  ]),
  ("LABOUR",[
   ("0114","Beldar","day",2.20,558.00),
  ]),
 ],
},
{
 "id":"2.13.3.1","desc":"Pipe trench ordinary rock, dia 150 mm, depth 3–4.5 m",
 "unit":"metre","base_qty":60,"dsr_rate":726.65,
 "ref_pattern":True,
 "sections":[
  ("REFERENCE ITEMS (A)",[
   ("REF 2.9.1","Excavation ordinary rock (72.29 cum)","cum",72.29,448.15),
   ("REF 2.25","Filling (72.29 cum)","cum",72.29,219.65),
  ]),
  ("LABOUR",[
   ("0114","Beldar","day",1.80,558.00),
  ]),
 ],
},
{
 "id":"2.13.3.2","desc":"Pipe trench hard rock blasting, dia 150 mm, depth 3–4.5 m",
 "unit":"metre","base_qty":60,"dsr_rate":1799.35,
 "ref_pattern":True,
 "sections":[
  ("REFERENCE ITEMS (A)",[
   ("REF 2.9.2","Excavation hard rock (72.29 cum)","cum",72.29,729.00),
   ("REF 2.25","Filling (72.29 cum)","cum",72.29,219.65),
  ]),
  ("LABOUR",[
   ("0114","Beldar","day",1.80,558.00),
  ]),
 ],
},
{
 "id":"2.13.3.3","desc":"Pipe trench hard rock no blasting, dia 150 mm, depth 3–4.5 m",
 "unit":"metre","base_qty":60,"dsr_rate":2070.50,
 "ref_pattern":True,
 "sections":[
  ("REFERENCE ITEMS (A)",[
   ("REF 2.9.3","Excavation hard rock no blasting (72.29 cum)","cum",72.29,1080.55),
   ("REF 2.25","Filling (72.29 cum)","cum",72.29,219.65),
  ]),
  ("LABOUR",[
   ("0114","Beldar","day",1.80,558.00),
  ]),
 ],
},

# ─── 2.14, 2.15 EXTRA DEPTH ROCK (%) ─────────────────────────────────────────
{
 "id":"2.14","desc":"Extra over item 2.13 for depth beyond 1.5 m upto 3 m in rock (percentage extra)",
 "unit":"%","base_qty":1,"dsr_rate":103.75,
 "pct_item":True,"sections":[],
 "notes":"DAR computed 103.60% (DSR 2021: 103.75%).",
},
{
 "id":"2.15","desc":"Extra over item 2.13 for depth beyond 3 m in rock (percentage extra)",
 "unit":"%","base_qty":1,"dsr_rate":256.15,
 "pct_item":True,"sections":[],
 "notes":"DAR computed 255.60% (DSR 2021: 256.15%).",
},

# ─── 2.16 CLOSE TIMBERING IN TRENCHES ────────────────────────────────────────
{
 "id":"2.16.1","desc":"Close timbering in trenches including strutting and bracing etc. and removing the same after the work is done - depth up to 2 m",
 "unit":"sqm","base_qty":90,"dsr_rate":132.90,
 "sections":[
  ("MATERIAL",[
   ("1198","Kail/Deodar planks 38 mm thick","cum",0.1050,12000.00),
   ("1197","Kail/Deodar scantlings","cum",0.0358,12000.00),
   ("0302","Safeda poles/Ballies 3–4 m","each",0.9000,40.00),
  ]),
  ("CARRIAGE",[
   ("2204","Carriage of timber by road","cum",0.1408,118.59),
  ]),
  ("LABOUR",[
   ("0112","Carpenter 2nd class","day",1.500,679.00),
   ("0114","Beldar","day",0.750,558.00),
   ("9999","Sundries","LS",13.52,2.00),
  ]),
 ],
},
{
 "id":"2.16.2","desc":"Close timbering in trenches - depth 2–4 m",
 "unit":"sqm","base_qty":90,"dsr_rate":145.55,
 "sections":[
  ("MATERIAL",[
   ("1198","Kail/Deodar planks 38 mm thick","cum",0.1250,12000.00),
   ("1197","Kail/Deodar scantlings","cum",0.0420,12000.00),
   ("0302","Safeda poles/Ballies 3–4 m","each",0.9000,40.00),
  ]),
  ("CARRIAGE",[
   ("2204","Carriage of timber by road","cum",0.1670,118.59),
  ]),
  ("LABOUR",[
   ("0112","Carpenter 2nd class","day",1.750,679.00),
   ("0114","Beldar","day",0.875,558.00),
   ("9999","Sundries","LS",17.94,2.00),
  ]),
 ],
},
{
 "id":"2.16.3","desc":"Close timbering in trenches - depth 4–6 m",
 "unit":"sqm","base_qty":90,"dsr_rate":174.00,
 "sections":[
  ("MATERIAL",[
   ("1198","Kail/Deodar planks 38 mm thick","cum",0.1600,12000.00),
   ("1197","Kail/Deodar scantlings","cum",0.0580,12000.00),
   ("0302","Safeda poles/Ballies 3–4 m","each",0.9000,40.00),
  ]),
  ("CARRIAGE",[
   ("2204","Carriage of timber by road","cum",0.2180,118.59),
  ]),
  ("LABOUR",[
   ("0112","Carpenter 2nd class","day",2.250,679.00),
   ("0114","Beldar","day",1.125,558.00),
   ("9999","Sundries","LS",26.91,2.00),
  ]),
 ],
},

# ─── 2.17 CLOSE TIMBERING IN SHAFTS ──────────────────────────────────────────
{
 "id":"2.17.1","desc":"Close timbering in shafts - depth up to 2 m",
 "unit":"sqm","base_qty":6.6,"dsr_rate":142.50,
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
 "id":"2.17.2","desc":"Close timbering in shafts - depth 2–4 m",
 "unit":"sqm","base_qty":6.6,"dsr_rate":169.35,
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
 "id":"2.17.3","desc":"Close timbering in shafts - depth 4–6 m",
 "unit":"sqm","base_qty":6.6,"dsr_rate":197.60,
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
 "id":"2.18.1","desc":"Close timbering over areas - depth up to 2 m",
 "unit":"sqm","base_qty":45,"dsr_rate":119.10,
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
 "id":"2.18.2","desc":"Close timbering over areas - depth 2–4 m",
 "unit":"sqm","base_qty":45,"dsr_rate":134.20,
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
 "id":"2.18.3","desc":"Close timbering over areas - depth 4–6 m",
 "unit":"sqm","base_qty":45,"dsr_rate":149.95,
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
 "id":"2.19","desc":"Extra for permanent close timbering (timber left in) over item 2.16",
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
 "id":"2.20.1","desc":"Open timbering in trenches - depth up to 2 m",
 "unit":"sqm","base_qty":90,"dsr_rate":68.55,
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
 "id":"2.20.2","desc":"Open timbering in trenches - depth 2–4 m",
 "unit":"sqm","base_qty":90,"dsr_rate":76.40,
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
 "id":"2.20.3","desc":"Open timbering in trenches - depth 4–6 m",
 "unit":"sqm","base_qty":90,"dsr_rate":89.35,
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
 "id":"2.21.1","desc":"Open timbering in shafts - depth up to 2 m",
 "unit":"sqm","base_qty":6.6,"dsr_rate":62.35,
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
 "id":"2.21.2","desc":"Open timbering in shafts - depth 2–4 m",
 "unit":"sqm","base_qty":6.6,"dsr_rate":74.95,
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
 "id":"2.21.3","desc":"Open timbering in shafts - depth 4–6 m",
 "unit":"sqm","base_qty":6.6,"dsr_rate":91.60,
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
 "id":"2.22.1","desc":"Open timbering over areas - depth up to 2 m",
 "unit":"sqm","base_qty":45,"dsr_rate":42.40,
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
 "id":"2.22.2","desc":"Open timbering over areas - depth 2–4 m",
 "unit":"sqm","base_qty":45,"dsr_rate":50.80,
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
 "id":"2.22.3","desc":"Open timbering over areas - depth 4–6 m",
 "unit":"sqm","base_qty":45,"dsr_rate":64.30,
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
 "id":"2.23","desc":"Extra for permanent open timbering (timber left in) over item 2.20",
 "unit":"sqm","base_qty":90,"dsr_rate":822.05,
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
 "id":"2.24.1","desc":"Extra rates for quantities of work executed in or under water and/or liquid mud, including pumping out water as required - 20% over each applicable item; apply only to qualifying quantity measured by metre depth from sub-soil water level to centre of gravity",
 "unit":"%","base_qty":1,"dsr_rate":None,
 "pct_item":True,"conditional_extra":True,"percent":20,
 "conditional_rule":"20% of qualifying selected base rate",
 "sections":[],
 "notes":"Conditional 20% extra over each applicable earthwork item, limited to qualifying work. Measure depth from sub-soil water level to centre of gravity. No resource rows.",
},
{
 "id":"2.24.2","desc":"Extra rates for quantities of work executed in or under foul position, including pumping out water as required - 25% over each applicable item; apply only to qualifying quantity measured by metre depth from sub-soil water level to centre of gravity",
 "unit":"%","base_qty":1,"dsr_rate":None,
 "pct_item":True,"conditional_extra":True,"percent":25,
 "conditional_rule":"25% of qualifying selected base rate",
 "sections":[],
 "notes":"Conditional 25% extra over each applicable earthwork item, limited to qualifying work. Measure depth from sub-soil water level to centre of gravity. No resource rows.",
},

# ─── 2.25 FILLING IN TRENCHES ────────────────────────────────────────────────
{
 "id":"2.25","desc":"Filling in trenches, plinth, sides of foundations, etc., with available excavated earth, in layers not exceeding 20 cm in depth, consolidated each layer by ramming and watering",
 "unit":"cum","base_qty":10,"dsr_rate":253.95,
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
 "id":"2.25(a)","desc":"Filling in trenches/plinth with sand under plinth by mechanical means using excavated earth from borrow pits with all leads and lifts",
 "unit":"cum","base_qty":10,"dsr_rate":368.65,
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
 "id":"2.26.1","desc":"Extra over item 2.3 for lift above 1.5 m upto 3 m",
 "unit":"cum","base_qty":10,"dsr_rate":104.50,
 "sections":[
  ("LABOUR",[
   ("0128","Mate","day",0.10,617.00),
   ("0114","Beldar","day",1.10,558.00),
  ]),
 ],
},
{
 "id":"2.26.2","desc":"Extra over item 2.3 for lift above 3 m upto 4.5 m",
 "unit":"cum","base_qty":10,"dsr_rate":187.40,
 "sections":[
  ("LABOUR",[
   ("0128","Mate","day",0.20,617.00),
   ("0114","Beldar","day",1.95,558.00),
  ]),
 ],
},

# ─── 2.27 FILLING WITH SAND ───────────────────────────────────────────────────
{
 "id":"2.27","desc":"Filling in trenches/plinth with sand under floors including watering and compaction",
 "unit":"cum","base_qty":10,"dsr_rate":2161.20,
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
 "id":"2.28.1","desc":"Dressing of slopes (cut or fill) in ordinary soil including trimming to required slope",
 "unit":"sqm","base_qty":100,"dsr_rate":28.15,
 "sections":[
  ("LABOUR",[
   ("0114","Beldar","day",1.97,558.00),
   ("0115","Coolie","day",1.29,558.00),
  ]),
 ],
},
{
 "id":"2.28.2","desc":"Dressing of slopes in hard soil",
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
 "id":"2.29.1","desc":"Ramming and watering the bottoms of excavations or tops of fillings including Bhisti charges",
 "unit":"sqm","base_qty":100,"dsr_rate":28.50,
 "sections":[
  ("LABOUR",[
   ("0114","Beldar","day",2.75,558.00),
   ("0101","Bhisti","day",0.50,617.00),
  ]),
 ],
},
{
 "id":"2.29.2","desc":"Ramming and watering (for hard soil)",
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
 "id":"2.30.1","desc":"Boring holes for fence posts in ordinary soil (avg 0.30 cum each), base 10 holes",
 "unit":"each","base_qty":10,"dsr_rate":89.90,
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
 "id":"2.30.2","desc":"Boring holes in ordinary rock",
 "unit":"each","base_qty":10,"dsr_rate":160.90,
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
 "id":"2.30.3","desc":"Boring holes in hard rock by blasting",
 "unit":"each","base_qty":10,"dsr_rate":257.70,
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
 "id":"2.30.4","desc":"Boring holes in hard rock blasting prohibited",
 "unit":"each","base_qty":10,"dsr_rate":381.40,
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
 "id":"2.31","desc":"Scarifying existing road surface 50 mm deep by means of a grader with ripper attachment",
 "unit":"sqm","base_qty":100,"dsr_rate":14.50,
 "sections":[
  ("LABOUR",[
   ("0114","Beldar","day",1.08,558.00),
   ("0115","Coolie","day",0.60,558.00),
  ]),
 ],
},

# ─── 2.32 GRADING BORROW PITS ────────────────────────────────────────────────
{
 "id":"2.32","desc":"Grading around borrow pits after earthwork is done",
 "unit":"sqm","base_qty":100,"dsr_rate":7.40,
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
 "id":"2.33.1","desc":"Felling and removal of trees including cutting of branches and roots - girth up to 0.30 m",
 "unit":"each","base_qty":1,"dsr_rate":439.25,
 "sections":[
  ("LABOUR",[
   ("0114","Beldar","day",0.33,558.00),
   ("0115","Coolie","day",0.17,558.00),
   ("9999","Sundries","LS",2.73,2.00),
  ]),
 ],
},
{
 "id":"2.33.2","desc":"Felling and removal of trees - girth 0.30–0.60 m",
 "unit":"each","base_qty":1,"dsr_rate":1957.15,
 "sections":[
  ("LABOUR",[
   ("0114","Beldar","day",1.50,558.00),
   ("0115","Coolie","day",0.75,558.00),
   ("9999","Sundries","LS",5.46,2.00),
  ]),
 ],
},
{
 "id":"2.33.3","desc":"Felling and removal of trees - girth 0.60–0.90 m",
 "unit":"each","base_qty":1,"dsr_rate":9084.05,
 "sections":[
  ("LABOUR",[
   ("0114","Beldar","day",7.00,558.00),
   ("0115","Coolie","day",3.50,558.00),
   ("9999","Sundries","LS",8.06,2.00),
  ]),
 ],
},
{
 "id":"2.33.4","desc":"Felling and removal of trees - girth 0.90–1.20 m",
 "unit":"each","base_qty":1,"dsr_rate":18198.70,
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
 "id":"2.34.1","desc":"Pre-constructional anti-termite treatment: applying Chlorpyriphos 1% emulsion in water @ 5 litres per sqm",
 "unit":"litre","base_qty":100,"dsr_rate":200.90,
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
 "id":"2.35.1.1","desc":"Post-constructional anti-termite: drilling holes 12 mm dia at 300 mm c/c in brick/RCC columns at ground level and injecting chemical - per metre depth of drilling",
 "unit":"metre","base_qty":1,"dsr_rate":32.30,
 "sections":[
  ("LABOUR",[
   ("0114","Beldar","day",0.33,558.00),
   ("9999","Sundries","LS",13.52,2.00),
  ]),
 ],
},
{
 "id":"2.35.2.1","desc":"Post-constructional anti-termite: drilling and injecting chemical in wall at 450 mm c/c both faces - per metre run of wall",
 "unit":"metre","base_qty":1,"dsr_rate":44.70,
 "sections":[
  ("LABOUR",[
   ("0114","Beldar","day",0.40,558.00),
   ("9999","Sundries","LS",35.88,2.00),
  ]),
 ],
},
{
 "id":"2.35.3.1","desc":"Post-constructional anti-termite: rodding in floor junction with wall and injecting chemical - per sqm",
 "unit":"sqm","base_qty":1,"dsr_rate":256.15,
 "sections":[
  ("LABOUR",[
   ("0114","Beldar","day",2.00,558.00),
   ("0124","Mason 2nd class","day",0.50,679.00),
   ("9999","Sundries","LS",35.88,2.00),
  ]),
 ],
},
{
 "id":"2.35.4.1","desc":"Post-constructional anti-termite: treatment on external perimeter of building - per metre run",
 "unit":"metre","base_qty":1,"dsr_rate":35.75,
 "sections":[
  ("LABOUR",[
   ("0114","Beldar","day",0.30,558.00),
   ("0124","Mason 2nd class","day",0.05,679.00),
   ("9999","Sundries","LS",17.94,2.00),
  ]),
 ],
},
{
 "id":"2.35.5","desc":"Treatment of timber against white ants using kerosene solution - per metre of sill/frame/chowkat",
 "unit":"metre","base_qty":1,"dsr_rate":257.55,
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
 "id":"2.36","desc":"Compaction of earth fill by mechanical sheep-foot roller, lift not exceeding 300 mm",
 "unit":"cum","base_qty":10,"dsr_rate":76.70,
 "sections":[
  ("LABOUR",[
   ("0128","Mate","day",0.08,617.00),
   ("0114","Beldar/Coolie","day",0.80,558.00),
  ]),
 ],
},

# ─── 2.37 FILLING WITH FLY ASH ───────────────────────────────────────────────
{
 "id":"2.37","desc":"Filling in plinth/trenches with fly ash obtained from thermal power plant",
 "unit":"cum","base_qty":1,"dsr_rate":234.05,
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
 "id":"2.38","desc":"Filling in plinth/trenches with earth obtained from approved borrow pits including all leads/lifts, watering and compaction",
 "unit":"cum","base_qty":10,"dsr_rate":253.95,
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


def write_item(ws, item, start_row):
    """Write one complete item analysis block. Returns next free row."""
    r = start_row
    base = item["base_qty"]
    unit = item["unit"]
    is_pct   = item.get("pct_item", False)
    is_conditional_extra = item.get("conditional_extra", False)
    is_ref   = item.get("ref_pattern", False)
    notes    = item.get("notes", "")

    # ── Item header ──
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    hdr_txt = f"Item {item['id']}  |  {item['desc']}"
    c = ws.cell(row=r, column=1, value=hdr_txt)
    c.font  = FT_HDR
    c.fill  = fill(C_ITEM_HDR)
    c.alignment = AL_L
    r += 1

    # Base quantity header
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    if is_conditional_extra:
        base_txt = (
            f"Conditional percentage extra: {item['conditional_rule']} (not a flat Rs rate). "
            "Apply only to the qualifying quantity/depth."
        )
    else:
        base_txt = f"Details of cost for {base} {unit}" + (f"  [DSR 2021 Rate: {item['dsr_rate']:.2f}/{unit}]")
    c = ws.cell(row=r, column=1, value=base_txt)
    c.font  = Font(bold=True, italic=True, size=9)
    c.fill  = fill("D6E4F0")
    c.alignment = AL_L
    r += 1

    # Column headers
    for ci, h in enumerate(HEADERS, 1):
        c = ws.cell(row=r, column=ci, value=h)
        c.font  = Font(bold=True, size=9, color="FFFFFF")
        c.fill  = fill("4472C4")
        c.alignment = AL_C
        c.border = thin_border()
    r += 1

    # ── Percentage-only items ──
    if is_pct:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        txt = notes if notes else f"Percentage extra: {item.get('percent', item['dsr_rate']):.2f}%"
        c = ws.cell(row=r, column=1, value=txt)
        c.font  = FT_BODY
        c.fill  = fill(C_CHAIN_LBL)
        c.alignment = AL_L
        r += 1
        # per-unit row
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        if is_conditional_extra:
            ws.cell(row=r, column=1, value="Conditional composer rule (not a flat Rs rate)").font = FT_BOLD
            set_cell(ws, r, 6, item["conditional_rule"], font=FT_BOLD, fill_obj=fill(C_FINAL), align=AL_L)
        else:
            ws.cell(row=r, column=1, value="DAR 2019 Rate (CivilDAR computed)").font = FT_BOLD
            set_cell(ws, r, 6, item["dsr_rate"], font=FT_BOLD, fill_obj=fill(C_FINAL), num_format=NUM_FMT)
        ws.cell(row=r, column=6).alignment = AL_R
        r += 2
        return r

    # ── Resource sections ──
    # Collect rows for W calculation (exclude REF rows if ref_pattern)
    resource_rows = []  # list of (sheet_row, is_ref_A)
    ref_total = 0.0     # sum of ref amounts (not in W chain)
    w_rows    = []      # (qty_cell_ref, rate_cell_ref, amount_cell_ref)

    for sec_name, rows in item["sections"]:
        if not rows:
            continue
        # Section header
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
        c = ws.cell(row=r, column=1, value=sec_name)
        c.font  = FT_SEC
        c.fill  = fill(C_SEC_HDR)
        c.alignment = AL_L
        r += 1

        is_ref_section = ("(A" in sec_name or "REFERENCE" in sec_name.upper())

        for code, desc, u, qty, rate in rows:
            is_ref_row = (code == "REF" or is_ref_section)
            f_bg = C_REF_ROW if is_ref_row else C_WHITE

            set_cell(ws, r, COL_CODE, code,  font=FT_BODY, fill_obj=fill(f_bg), align=AL_C)
            set_cell(ws, r, COL_DESC, desc,  font=FT_BODY, fill_obj=fill(f_bg), align=AL_L)
            set_cell(ws, r, COL_UNIT, u,     font=FT_BODY, fill_obj=fill(f_bg), align=AL_C)
            qc = set_cell(ws, r, COL_QTY,  qty,  font=FT_BODY, fill_obj=fill(f_bg), num_format=NUM_FMT)
            rc = set_cell(ws, r, COL_RATE, rate, font=FT_BODY, fill_obj=fill(f_bg), num_format=NUM_FMT)

            # Amount formula: =D_r * E_r
            qty_ref  = f"D{r}"
            rate_ref = f"E{r}"
            amt_ref  = f"F{r}"
            ac = ws.cell(row=r, column=COL_AMT)
            ac.value  = f"={qty_ref}*{rate_ref}"
            ac.font   = FT_BODY
            ac.fill   = fill(f_bg)
            ac.number_format = NUM_FMT
            ac.border = thin_border()

            if is_ref_row:
                ref_total += qty * rate
            else:
                w_rows.append(amt_ref)

            r += 1

    # ── Calculation chain ──
    chain_fill = fill(C_CHAIN_LBL)
    chain_font = FT_CHAIN

    # TOTAL W
    if w_rows:
        w_formula = "=SUM(" + ",".join(w_rows) + ")"
    else:
        w_formula = 0

    # For ref_pattern items: show ref total separately, then add W for chain
    if is_ref and not w_rows:
        # Pure ref item (e.g. 2.10.1.x) — chain on the ref total
        set_cell(ws, r, 1, "Sub-total of Referenced Items (A)", font=chain_font, fill_obj=fill("C6EFCE"), align=AL_L)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        ref_sum_cell = f"F{r}"
        # sum all F rows in the section block — just use literal
        all_f = [f"F{rr}" for rr in range(start_row+3, r) if True]
        # simpler: just put the computed ref_total as value
        set_cell(ws, r, 6, ref_total, font=chain_font, fill_obj=fill("C6EFCE"), num_format=NUM_FMT)
        ref_sum_row = r
        r += 1

        # per-unit = ref_total / base
        per_unit_dar = ref_total / base
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        ws.cell(row=r, column=1, value=f"DAR 2019 Rate per {unit} = Total / {base} {unit}").font = chain_font
        set_cell(ws, r, 6, per_unit_dar, font=Font(bold=True,size=10), fill_obj=fill(C_FINAL), num_format=NUM_FMT)
        r += 1

        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        ws.cell(row=r, column=1, value=f"DSR 2021 Rate (cross-reference)").font = Font(italic=True,size=9)
        set_cell(ws, r, 6, item["dsr_rate"], font=Font(italic=True,size=9), fill_obj=fill(C_DSR), num_format=NUM_FMT)
        r += 2
        return r

    # Standard chain (or W-A where W rows exist alongside refs)
    w_row_num = r
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    ws.cell(row=r, column=1, value="TOTAL W  (Direct cost of resources)").font = chain_font
    ws.cell(row=r, column=1).fill = chain_fill
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
    w_cell = f"F{r}"
    ac = ws.cell(row=r, column=6)
    ac.value = w_formula
    ac.font  = chain_font
    ac.fill  = chain_fill
    ac.number_format = NUM_FMT
    ac.border = thin_border()
    r += 1

    # Water 1%
    water_row = r
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    ws.cell(row=r, column=1, value=f"Add: Water charges @ 1% of W").font = chain_font
    ws.cell(row=r, column=1).fill = chain_fill
    set_cell(ws, r, 6, f"={w_cell}*0.01", font=chain_font, fill_obj=chain_fill, num_format=NUM_FMT)
    water_cell = f"F{r}"
    r += 1

    # TOTAL X
    x_row = r
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    ws.cell(row=r, column=1, value="TOTAL X  (W + Water)").font = chain_font
    ws.cell(row=r, column=1).fill = chain_fill
    x_cell = f"F{r}"
    set_cell(ws, r, 6, f"={w_cell}+{water_cell}", font=chain_font, fill_obj=chain_fill, num_format=NUM_FMT)
    r += 1

    # GST 14.05%
    gst_row = r
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    ws.cell(row=r, column=1, value="Add: GST @ 14.05% of X").font = chain_font
    ws.cell(row=r, column=1).fill = chain_fill
    gst_cell = f"F{r}"
    set_cell(ws, r, 6, f"={x_cell}*0.1405", font=chain_font, fill_obj=chain_fill, num_format=NUM_FMT)
    r += 1

    # TOTAL Y
    y_row = r
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    ws.cell(row=r, column=1, value="TOTAL Y  (X + GST)").font = chain_font
    ws.cell(row=r, column=1).fill = chain_fill
    y_cell = f"F{r}"
    set_cell(ws, r, 6, f"={x_cell}+{gst_cell}", font=chain_font, fill_obj=chain_fill, num_format=NUM_FMT)
    r += 1

    # CP&OH 15%
    cpoh_row = r
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    ws.cell(row=r, column=1, value="Add: Contractor's Profit & Overheads (CP&OH) @ 15% of Y").font = chain_font
    ws.cell(row=r, column=1).fill = chain_fill
    cpoh_cell = f"F{r}"
    set_cell(ws, r, 6, f"={y_cell}*0.15", font=chain_font, fill_obj=chain_fill, num_format=NUM_FMT)
    r += 1

    # TOTAL Z
    z_row = r
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    ws.cell(row=r, column=1, value="TOTAL Z  (Y + CP&OH)").font = chain_font
    ws.cell(row=r, column=1).fill = chain_fill
    z_cell = f"F{r}"
    set_cell(ws, r, 6, f"={y_cell}+{cpoh_cell}", font=chain_font, fill_obj=chain_fill, num_format=NUM_FMT)
    r += 1

    # Cess 1%
    cess_row = r
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    ws.cell(row=r, column=1, value="Add: Cess @ 1% of Z").font = chain_font
    ws.cell(row=r, column=1).fill = chain_fill
    cess_cell = f"F{r}"
    set_cell(ws, r, 6, f"={z_cell}*0.01", font=chain_font, fill_obj=chain_fill, num_format=NUM_FMT)
    r += 1

    # Total for base qty
    total_row = r
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    if is_ref:
        ws.cell(row=r, column=1, value=f"Cost for {base} {unit} [Z + Cess + Referenced items {ref_total:.2f}]").font = chain_font
        total_cell = f"F{r}"
        set_cell(ws, r, 6, f"={z_cell}+{cess_cell}+{ref_total}", font=chain_font, fill_obj=fill("BDD7EE"), num_format=NUM_FMT)
    else:
        ws.cell(row=r, column=1, value=f"Total cost for {base} {unit}  (Z + Cess)").font = chain_font
        total_cell = f"F{r}"
        set_cell(ws, r, 6, f"={z_cell}+{cess_cell}", font=chain_font, fill_obj=fill("BDD7EE"), num_format=NUM_FMT)
    ws.cell(row=r, column=1).fill = fill("BDD7EE")
    r += 1

    # Per-unit rate
    rate_row = r
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    ws.cell(row=r, column=1, value=f"DAR 2019 Rate per {unit}  =  Total / {base}").font = Font(bold=True,size=10)
    ws.cell(row=r, column=1).fill = fill(C_FINAL)
    per_unit_cell = f"F{r}"
    set_cell(ws, r, 6, f"={total_cell}/{base}", font=Font(bold=True,size=10), fill_obj=fill(C_FINAL), num_format=NUM_FMT)
    r += 1

    # Say rounded (informational)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    ws.cell(row=r, column=1, value=f"Say (rounded)").font = Font(italic=True,size=9)
    ws.cell(row=r, column=1).fill = fill(C_FINAL)
    # ROUND formula
    set_cell(ws, r, 6, f"=ROUND({per_unit_cell},2)", font=Font(italic=True,bold=True,size=9), fill_obj=fill(C_FINAL), num_format=NUM_FMT)
    r += 1

    # DSR 2021 cross-reference
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    ws.cell(row=r, column=1, value="DSR 2021 Rate (cross-reference only)").font = Font(italic=True,size=9)
    ws.cell(row=r, column=1).fill = fill(C_DSR)
    set_cell(ws, r, 6, item["dsr_rate"], font=Font(italic=True,size=9), fill_obj=fill(C_DSR), num_format=NUM_FMT)
    r += 1

    # Project columns — user fills G; H = G * per-unit
    proj_hdr_row = rate_row  # we'll add project cols at rate row
    # Set formulas at per-unit row
    pu_row = rate_row
    ws.cell(row=pu_row, column=COL_PROJ_Q).value = None   # user fills
    ws.cell(row=pu_row, column=COL_PROJ_Q).fill  = fill("FFF2CC")
    ws.cell(row=pu_row, column=COL_PROJ_Q).border = thin_border()
    ws.cell(row=pu_row, column=COL_PROJ_A).value  = f"=G{pu_row}*F{pu_row}"
    ws.cell(row=pu_row, column=COL_PROJ_A).number_format = NUM_FMT
    ws.cell(row=pu_row, column=COL_PROJ_A).fill   = fill("FFF2CC")
    ws.cell(row=pu_row, column=COL_PROJ_A).border = thin_border()

    if notes:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
        ws.cell(row=r, column=1, value=f"Note: {notes}").font = Font(italic=True,size=8,color="666666")
        r += 1

    r += 1  # blank separator
    return r


def build_sheet(ws):
    # Sheet title
    ws.merge_cells("A1:H1")
    ws.cell(row=1, column=1, value="Sub-Head 2.0 — EARTH WORK  |  CPWD DAR 2019 Complete Rate Analysis")
    ws.cell(row=1, column=1).font  = Font(bold=True, size=13, color="FFFFFF")
    ws.cell(row=1, column=1).fill  = fill("1A1A2E")
    ws.cell(row=1, column=1).alignment = AL_C

    ws.merge_cells("A2:H2")
    ws.cell(row=2, column=1, value="Source: CivilDAR_2019_Vol_1.pdf  |  Basic rates: 2018-based  |  Chain: W → Water(1%) → X → GST(14.05%) → Y → CP&OH(15%) → Z → Cess(1%) → Total → Per Unit Rate")
    ws.cell(row=2, column=1).font  = Font(italic=True, size=8, color="333333")
    ws.cell(row=2, column=1).fill  = fill("E8F4FD")
    ws.cell(row=2, column=1).alignment = AL_L

    current_row = 3
    for item in ITEMS:
        current_row = write_item(ws, item, current_row)

    # Column widths
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 55
    ws.column_dimensions["C"].width = 8
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 14
    ws.column_dimensions["F"].width = 14
    ws.column_dimensions["G"].width = 14
    ws.column_dimensions["H"].width = 16

    # Row heights: basic
    ws.sheet_view.showGridLines = True
    ws.freeze_panes = "A3"


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
    print(f"Done → {WB_PATH}")
    print(f"Items written: {len(ITEMS)}")


if __name__ == "__main__":
    main()
