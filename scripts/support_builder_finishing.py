"""
support_builder_finishing.py
Rebuilds the '13_Finishing' worksheet in the Vol 2 workbook with a
first-principles resource analysis for all 165 DAR 2019 Ch.13 items.
Each item block shows: resources (code, desc, unit, qty, rate, amount),
markup chain (W→X→Y→Z→Cost→Rate→SAY), and the PDF-published SAY rate.

Usage:
    python scripts/support_builder_finishing.py
"""

import openpyxl, json
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from pathlib import Path
import os, sys

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from scripts.paths import WB_VOL2_FILE as VOL2_WORKBOOK

CH13_JSON = Path(__file__).resolve().parent.parent / "data" / "reference_json" / "ch13_items.json"


def load_ch13_items():
    with open(CH13_JSON, encoding="utf-8") as f:
        return json.load(f)

# ── Colours ──────────────────────────────────────────────────────────────────
C_TITLE_BG    = "1A1A2E"
C_PARENT_BG   = "1F4E79"
C_ITEM_BG     = "2E75B6"
C_TH_BG       = "1F4E79"
C_ALT_ROW     = "F2F7FC"
C_SAY_BG      = "E2EFDA"
C_WHITE       = "FFFFFF"
C_WARN        = "FFF2CC"

def fill(h):
    return PatternFill(fill_type="solid", fgColor=h) if h else PatternFill(fill_type=None)

def tb():
    s = Side(style="thin", color="AAAAAA")
    return Border(left=s, right=s, top=s, bottom=s)

AL_C = Alignment(horizontal="center", vertical="center")
AL_L = Alignment(horizontal="left",   vertical="center", wrap_text=True)
AL_R = Alignment(horizontal="right",  vertical="center")

# ── Ch.13 item catalogue (PDF-extracted, DAR 2019 Vol 2) ────────────────────
# Format: (code, parent_title, description, unit, basis_qty, say_rate, resources)
# resources: list of (res_code, description, unit, qty, rate)
# res_code "REF:x.x" = cross-reference to Vol1 item; "MAT" = material; code = DSR code

CH13_ITEMS = [
  # ── 13.1  12 mm cement plaster – fine sand ──────────────────────────────
  ("13.1.1","13.1 12 mm cement plaster of mix",
   "1:4 (1 cement: 4 fine sand)","sqm",10,[
    ("3.4","Cement mortar 1:4 (fine sand) – SH Mortars","cum",0.144,3528.85),
    ("0155","Mason (average)","day",0.67,709.00),
    ("0115","Coolie","day",0.75,558.00),
    ("0101","Bhisti","day",0.92,617.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],266.85),
  ("13.1.2","13.1 12 mm cement plaster of mix",
   "1:6 (1 cement: 6 fine sand)","sqm",10,[
    ("3.6","Cement mortar 1:6 (fine sand) – SH Mortars","cum",0.144,2874.65),
    ("0155","Mason (average)","day",0.67,709.00),
    ("0115","Coolie","day",0.75,558.00),
    ("0101","Bhisti","day",0.92,617.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],254.25),
  # ── 13.2  15 mm cement plaster – rough side – fine sand ─────────────────
  ("13.2.1","13.2 15 mm cement plaster on rough side – fine sand",
   "1:4 (1 cement: 4 fine sand)","sqm",10,[
    ("3.4","Cement mortar 1:4 (fine sand) – SH Mortars","cum",0.172,3528.85),
    ("0155","Mason (average)","day",0.80,709.00),
    ("0115","Coolie","day",0.88,558.00),
    ("0101","Bhisti","day",0.99,617.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],307.90),
  ("13.2.2","13.2 15 mm cement plaster on rough side – fine sand",
   "1:6 (1 cement: 6 fine sand)","sqm",10,[
    ("3.6","Cement mortar 1:6 (fine sand) – SH Mortars","cum",0.172,2874.65),
    ("0155","Mason (average)","day",0.80,709.00),
    ("0115","Coolie","day",0.88,558.00),
    ("0101","Bhisti","day",0.99,617.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],292.85),
  # ── 13.3  20 mm cement plaster – fine sand ──────────────────────────────
  ("13.3.1","13.3 20 mm cement plaster of mix",
   "1:4 (1 cement: 4 fine sand)","sqm",10,[
    ("3.4","Cement mortar 1:4 (fine sand) – SH Mortars","cum",0.224,3528.85),
    ("0155","Mason (average)","day",0.94,709.00),
    ("0115","Coolie","day",1.04,558.00),
    ("0101","Bhisti","day",1.09,617.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],367.25),
  ("13.3.2","13.3 20 mm cement plaster of mix",
   "1:6 (1 cement: 6 fine sand)","sqm",10,[
    ("3.6","Cement mortar 1:6 (fine sand) – SH Mortars","cum",0.224,2874.65),
    ("0155","Mason (average)","day",0.94,709.00),
    ("0115","Coolie","day",1.04,558.00),
    ("0101","Bhisti","day",1.09,617.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],348.30),
  # ── 13.4  12 mm cement plaster – coarse sand ────────────────────────────
  ("13.4.1","13.4 12 mm cement plaster of mix – coarse sand",
   "1:4 (1 cement: 4 coarse sand)","sqm",10,[
    ("3.9","Cement mortar 1:4 (coarse sand) – SH Mortars","cum",0.144,3206.90),
    ("0155","Mason (average)","day",0.67,709.00),
    ("0115","Coolie","day",0.75,558.00),
    ("0101","Bhisti","day",0.92,617.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],248.55),
  ("13.4.2","13.4 12 mm cement plaster of mix – coarse sand",
   "1:6 (1 cement: 6 coarse sand)","sqm",10,[
    ("3.11","Cement mortar 1:6 (coarse sand) – SH Mortars","cum",0.144,2746.05),
    ("0155","Mason (average)","day",0.67,709.00),
    ("0115","Coolie","day",0.75,558.00),
    ("0101","Bhisti","day",0.92,617.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],237.80),
  # ── 13.5  15 mm cement plaster – rough side – coarse sand ───────────────
  ("13.5.1","13.5 15 mm cement plaster on rough side – coarse sand",
   "1:4 (1 cement: 4 coarse sand)","sqm",10,[
    ("3.9","Cement mortar 1:4 (coarse sand) – SH Mortars","cum",0.172,3206.90),
    ("0155","Mason (average)","day",0.80,709.00),
    ("0115","Coolie","day",0.88,558.00),
    ("0101","Bhisti","day",0.99,617.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],288.75),
  ("13.5.2","13.5 15 mm cement plaster on rough side – coarse sand",
   "1:6 (1 cement: 6 coarse sand)","sqm",10,[
    ("3.11","Cement mortar 1:6 (coarse sand) – SH Mortars","cum",0.172,2746.05),
    ("0155","Mason (average)","day",0.80,709.00),
    ("0115","Coolie","day",0.88,558.00),
    ("0101","Bhisti","day",0.99,617.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],275.75),
  # ── 13.6  20 mm cement plaster – coarse sand ────────────────────────────
  ("13.6.1","13.6 20 mm cement plaster of mix – coarse sand",
   "1:4 (1 cement: 4 coarse sand)","sqm",10,[
    ("3.9","Cement mortar 1:4 (coarse sand) – SH Mortars","cum",0.224,3206.90),
    ("0155","Mason (average)","day",0.94,709.00),
    ("0115","Coolie","day",1.04,558.00),
    ("0101","Bhisti","day",1.09,617.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],344.80),
  ("13.6.2","13.6 20 mm cement plaster of mix – coarse sand",
   "1:6 (1 cement: 6 coarse sand)","sqm",10,[
    ("3.11","Cement mortar 1:6 (coarse sand) – SH Mortars","cum",0.224,2746.05),
    ("0155","Mason (average)","day",0.94,709.00),
    ("0115","Coolie","day",1.04,558.00),
    ("0101","Bhisti","day",1.09,617.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],327.80),
  # ── 13.7  12 mm plaster with neat cement floating coat ──────────────────
  ("13.7.1","13.7 12 mm cement plaster with floating coat of neat cement",
   "1:3 (1 cement: 3 fine sand)","sqm",10,[
    ("3.3","Cement mortar 1:3 (fine sand) – SH Mortars","cum",0.144,4181.95),
    ("0155","Mason (average)","day",0.80,709.00),
    ("0115","Coolie","day",0.88,558.00),
    ("0101","Bhisti","day",0.99,617.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],311.90),
  ("13.7.2","13.7 12 mm cement plaster with floating coat of neat cement",
   "1:4 (1 cement: 4 fine sand)","sqm",10,[
    ("3.4","Cement mortar 1:4 (fine sand) – SH Mortars","cum",0.144,3528.85),
    ("0155","Mason (average)","day",0.80,709.00),
    ("0115","Coolie","day",0.88,558.00),
    ("0101","Bhisti","day",0.99,617.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],291.85),
  # ── 13.8  15 mm rough-side plaster with floating coat ───────────────────
  ("13.8.1","13.8 15 mm rough-side plaster with floating coat of neat cement",
   "1:3 (1 cement: 3 fine sand)","sqm",10,[
    ("3.3","Cement mortar 1:3 (fine sand) – SH Mortars","cum",0.172,4181.95),
    ("0155","Mason (average)","day",0.94,709.00),
    ("0115","Coolie","day",1.04,558.00),
    ("0101","Bhisti","day",1.09,617.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],349.55),
  ("13.8.2","13.8 15 mm rough-side plaster with floating coat of neat cement",
   "1:4 (1 cement: 4 fine sand)","sqm",10,[
    ("3.4","Cement mortar 1:4 (fine sand) – SH Mortars","cum",0.172,3528.85),
    ("0155","Mason (average)","day",0.94,709.00),
    ("0115","Coolie","day",1.04,558.00),
    ("0101","Bhisti","day",1.09,617.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],326.70),
  # ── 13.9  Cement plaster 1:3 coarse sand with floating coat ─────────────
  ("13.9.1","13.9 Cement plaster 1:3 coarse sand with floating coat",
   "12 mm cement plaster","sqm",10,[
    ("3.8","Cement mortar 1:3 (coarse sand) – SH Mortars","cum",0.144,3771.25),
    ("0155","Mason (average)","day",0.80,709.00),
    ("0115","Coolie","day",0.88,558.00),
    ("0101","Bhisti","day",0.99,617.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],297.30),
  ("13.9.2","13.9 Cement plaster 1:3 coarse sand with floating coat",
   "20 mm cement plaster","sqm",10,[
    ("3.8","Cement mortar 1:3 (coarse sand) – SH Mortars","cum",0.224,3771.25),
    ("0155","Mason (average)","day",0.94,709.00),
    ("0115","Coolie","day",1.04,558.00),
    ("0101","Bhisti","day",1.09,617.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],349.75),
  # ── 13.10 ───────────────────────────────────────────────────────────────
  ("13.10","13.10 15 mm cement plaster 1:3 coarse sand with floating coat",
   "15 mm cement plaster on rough side with floating coat","sqm",10,[
    ("3.8","Cement mortar 1:3 (coarse sand) – SH Mortars","cum",0.172,3771.25),
    ("0155","Mason (average)","day",0.94,709.00),
    ("0115","Coolie","day",1.04,558.00),
    ("0101","Bhisti","day",1.09,617.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],323.05),
  # ── 13.11 ───────────────────────────────────────────────────────────────
  ("13.11","13.11 18 mm two-coat plaster – under 12 mm 1:5 fine sand",
   "18 mm cement plaster in two coats (under 12 mm 1:5, top 6 mm 1:3 fine sand)","sqm",10,[
    ("3.5","Cement mortar 1:5 (fine sand) – SH Mortars","cum",0.133,3076.05),
    ("3.3","Cement mortar 1:3 (fine sand) – SH Mortars","cum",0.066,4181.95),
    ("0155","Mason (average)","day",0.94,709.00),
    ("0115","Coolie","day",1.04,558.00),
    ("0101","Bhisti","day",1.09,617.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],326.45),
  # ── 13.12 ───────────────────────────────────────────────────────────────
  ("13.12","13.12 18 mm two-coat plaster – under 12 mm 1:5 coarse sand",
   "18 mm cement plaster in two coats (under 12 mm 1:5, top 6 mm 1:3 coarse sand)","sqm",10,[
    ("3.10","Cement mortar 1:5 (coarse sand) – SH Mortars","cum",0.133,2984.30),
    ("3.8","Cement mortar 1:3 (coarse sand) – SH Mortars","cum",0.066,3771.25),
    ("0155","Mason (average)","day",0.94,709.00),
    ("0115","Coolie","day",1.04,558.00),
    ("0101","Bhisti","day",1.09,617.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],316.60),
  # ── 13.13 ───────────────────────────────────────────────────────────────
  ("13.13","13.13 12 mm cement plaster 1:2 stone dust",
   "12 mm cement plaster 1:2 (1 cement: 2 stone dust)","sqm",10,[
    ("3.2","Cement mortar 1:2 (stone dust) – SH Mortars","cum",0.144,4945.85),
    ("0155","Mason (average)","day",0.67,709.00),
    ("0115","Coolie","day",0.75,558.00),
    ("0101","Bhisti","day",0.92,617.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],337.75),
  # ── 13.14 ───────────────────────────────────────────────────────────────
  ("13.14","13.14 15 mm cement plaster 1:2 stone dust – rough side",
   "15 mm cement plaster 1:2 (1 cement: 2 stone dust) on rough side","sqm",10,[
    ("3.2","Cement mortar 1:2 (stone dust) – SH Mortars","cum",0.172,4945.85),
    ("0155","Mason (average)","day",0.80,709.00),
    ("0115","Coolie","day",0.88,558.00),
    ("0101","Bhisti","day",0.99,617.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],380.25),
  # ── 13.15 ───────────────────────────────────────────────────────────────
  ("13.15","13.15 20 mm cement plaster 1:2 stone dust",
   "20 mm cement plaster 1:2 (1 cement: 2 stone dust)","sqm",10,[
    ("3.2","Cement mortar 1:2 (stone dust) – SH Mortars","cum",0.224,4945.85),
    ("0155","Mason (average)","day",0.94,709.00),
    ("0115","Coolie","day",1.04,558.00),
    ("0101","Bhisti","day",1.09,617.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],423.20),
  # ── 13.16 ───────────────────────────────────────────────────────────────
  ("13.16.1","13.16 6 mm cement plaster",
   "1:3 (1 cement: 3 fine sand)","sqm",10,[
    ("3.3","Cement mortar 1:3 (fine sand) – SH Mortars","cum",0.072,4181.95),
    ("0155","Mason (average)","day",0.53,709.00),
    ("0115","Coolie","day",0.59,558.00),
    ("0101","Bhisti","day",0.61,617.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],196.70),
  # ── 13.17 ───────────────────────────────────────────────────────────────
  ("13.17","13.17 6 mm cement plaster 1:3 fine sand with floating coat",
   "6 mm cement plaster 1:3 fine sand finished with a floating coat of neat cement","sqm",10,[
    ("3.3","Cement mortar 1:3 (fine sand) – SH Mortars","cum",0.072,4181.95),
    ("0155","Mason (average)","day",0.67,709.00),
    ("0115","Coolie","day",0.75,558.00),
    ("0101","Bhisti","day",0.92,617.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],218.05),
  # ── 13.18 ───────────────────────────────────────────────────────────────
  ("13.18","13.18 Neat cement punning",
   "Neat cement punning","sqm",10,[
    ("0155","Mason (average)","day",0.33,709.00),
    ("0115","Coolie","day",0.37,558.00),
    ("0101","Bhisti","day",0.46,617.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],62.75),
  # ── 13.19 ───────────────────────────────────────────────────────────────
  ("13.19.1","13.19 Rough cast plaster up to 10 m height",
   "Ordinary cement finish using ordinary cement","sqm",10,[
    ("3.4","Cement mortar 1:4 (fine sand) – SH Mortars","cum",0.172,3528.85),
    ("0155","Mason (average)","day",0.94,709.00),
    ("0115","Coolie","day",1.04,558.00),
    ("0101","Bhisti","day",1.09,617.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],323.65),
  # ── 13.20 ───────────────────────────────────────────────────────────────
  ("13.20","13.20 Pebble dash plaster up to 10 m height",
   "Pebble dash plaster up to 10 m height above ground level","sqm",10,[
    ("3.4","Cement mortar 1:4 (fine sand) – SH Mortars","cum",0.144,3528.85),
    ("0155","Mason (average)","day",1.07,709.00),
    ("0115","Coolie","day",1.19,558.00),
    ("0101","Bhisti","day",0.92,617.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],297.55),
  # ── 13.21 ───────────────────────────────────────────────────────────────
  ("13.21","13.21 Extra for water proofing material in cement plaster",
   "Extra for providing and mixing water proofing material in cement plaster","sqm",10,[
    ("9999","Water proofing compound (Pudlo or similar)","L.S.",98.15,2.00)],22.50),
  # ── 13.22 ───────────────────────────────────────────────────────────────
  ("13.22","13.22 Extra for plastering exterior walls above 10 m height",
   "Extra for plastering exterior walls of height more than 10 m from ground level","sqm",10,[
    ("9999","Extra labour and scaffolding allowance","L.S.",12.61,2.00)],13.00),
  # ── 13.23 ───────────────────────────────────────────────────────────────
  ("13.23.1","13.23 Extra for plastering on circular work ≤ 6 m radius",
   "In one coat","sqm",10,[
    ("9999","Extra for circular work – one coat","L.S.",12.61,2.00)],20.85),
  ("13.23.2","13.23 Extra for plastering on circular work ≤ 6 m radius",
   "In two coats","sqm",10,[
    ("9999","Extra for circular work – two coats","L.S.",12.61,2.00)],35.55),
  # ── 13.24 ───────────────────────────────────────────────────────────────
  ("13.24.1","13.24 Extra for plastering on moulding / cornices",
   "In one coat","sqm",10,[
    ("9999","Extra for moulding/cornices – one coat","L.S.",12.61,2.00)],35.55),
  ("13.24.2","13.24 Extra for plastering on moulding / cornices",
   "In two coats","sqm",10,[
    ("9999","Extra for moulding/cornices – two coats","L.S.",12.61,2.00)],60.10),
  # ── 13.25 ───────────────────────────────────────────────────────────────
  ("13.25.1","13.25 Extra for plastering",
   "Spherical ceiling","sqm",10,[
    ("9999","Extra for spherical ceiling","L.S.",12.61,2.00)],48.80),
  ("13.25.2","13.25 Extra for plastering",
   "Groined ceiling","sqm",10,[
    ("9999","Extra for groined ceiling","L.S.",12.61,2.00)],58.50),
  ("13.25.3","13.25 Extra for plastering",
   "Flewing soffits","sqm",10,[
    ("9999","Extra for flewing soffits","L.S.",12.61,2.00)],58.50),
  # ── 13.26 ───────────────────────────────────────────────────────────────
  ("13.26","13.26 Plaster of paris putty 2 mm over plastered surface",
   "Providing and applying plaster of paris putty 2 mm thickness over plastered surface","sqm",10,[
    ("2209","Plaster of Paris","kg",3.60,23.00),
    ("0155","Mason (average)","day",0.40,709.00),
    ("0115","Coolie","day",0.44,558.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],196.70),
  # ── 13.27 ───────────────────────────────────────────────────────────────
  ("13.27","13.27 Extra for lining out plaster to imitate stone blocks",
   "Extra for lining out plaster to imitate stone or concrete blocks walling","Rmt",10,[
    ("0155","Mason (average)","day",0.20,709.00),
    ("9999","Sundries","L.S.",12.61,2.00)],22.75),
  # ── 13.28  Plain cement mortar bands 12 mm ──────────────────────────────
  ("13.28.1","13.28 12 mm thick plain cement mortar bands CM 1:4",
   "Flush Band","Rmt",10,[
    ("3.4","Cement mortar 1:4 (fine sand) – SH Mortars","cum",0.012,3528.85),
    ("0155","Mason (average)","day",0.27,709.00),
    ("0115","Coolie","day",0.30,558.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],60.90),
  ("13.28.2","13.28 12 mm thick plain cement mortar bands CM 1:4",
   "Sunk Band","Rmt",10,[
    ("3.4","Cement mortar 1:4 (fine sand) – SH Mortars","cum",0.012,3528.85),
    ("0155","Mason (average)","day",0.33,709.00),
    ("0115","Coolie","day",0.37,558.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],68.60),
  ("13.28.3","13.28 12 mm thick plain cement mortar bands CM 1:4",
   "Raised Band","Rmt",10,[
    ("3.4","Cement mortar 1:4 (fine sand) – SH Mortars","cum",0.018,3528.85),
    ("0155","Mason (average)","day",0.40,709.00),
    ("0115","Coolie","day",0.44,558.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],80.15),
  ("13.28.4","13.28 12 mm thick plain cement mortar bands CM 1:4",
   "Moulded Band","Rmt",10,[
    ("3.4","Cement mortar 1:4 (fine sand) – SH Mortars","cum",0.018,3528.85),
    ("0155","Mason (average)","day",0.53,709.00),
    ("0115","Coolie","day",0.59,558.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],96.95),
  # ── 13.29  Plain cement mortar bands 18 mm ──────────────────────────────
  ("13.29.1","13.29 18 mm thick plain cement mortar band CM 1:4",
   "Flush Band","Rmt",10,[
    ("3.4","Cement mortar 1:4 (fine sand) – SH Mortars","cum",0.018,3528.85),
    ("0155","Mason (average)","day",0.33,709.00),
    ("0115","Coolie","day",0.37,558.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],71.30),
  ("13.29.2","13.29 18 mm thick plain cement mortar band CM 1:4",
   "Sunk Band","Rmt",10,[
    ("3.4","Cement mortar 1:4 (fine sand) – SH Mortars","cum",0.018,3528.85),
    ("0155","Mason (average)","day",0.40,709.00),
    ("0115","Coolie","day",0.44,558.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],80.15),
  ("13.29.3","13.29 18 mm thick plain cement mortar band CM 1:4",
   "Raised Band","Rmt",10,[
    ("3.4","Cement mortar 1:4 (fine sand) – SH Mortars","cum",0.024,3528.85),
    ("0155","Mason (average)","day",0.47,709.00),
    ("0115","Coolie","day",0.52,558.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],90.90),
  ("13.29.4","13.29 18 mm thick plain cement mortar band CM 1:4",
   "Moulded Band","Rmt",10,[
    ("3.4","Cement mortar 1:4 (fine sand) – SH Mortars","cum",0.024,3528.85),
    ("0155","Mason (average)","day",0.60,709.00),
    ("0115","Coolie","day",0.67,558.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],108.45),
  # ── 13.30 ───────────────────────────────────────────────────────────────
  ("13.30","13.30 18 mm moulded cement mortar band two-coat",
   "18 mm thick moulded cement mortar band in two coats – under 12 mm 1:5 top 6 mm 1:3","Rmt",10,[
    ("3.5","Cement mortar 1:5 (fine sand) – SH Mortars","cum",0.013,3076.05),
    ("3.3","Cement mortar 1:3 (fine sand) – SH Mortars","cum",0.007,4181.95),
    ("0155","Mason (average)","day",0.60,709.00),
    ("0115","Coolie","day",0.67,558.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],108.20),
  # ── 13.31  Pointing – brick work ────────────────────────────────────────
  ("13.31.1","13.31 Pointing on brick work CM 1:3 fine sand",
   "Flush/Ruled/Struck or weathered pointing","sqm",10,[
    ("3.3","Cement mortar 1:3 (fine sand) – SH Mortars","cum",0.018,4181.95),
    ("0155","Mason (average)","day",0.67,709.00),
    ("0115","Coolie","day",0.75,558.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],133.95),
  ("13.31.2","13.31 Pointing on brick work CM 1:3 fine sand",
   "Raised and cut pointing","sqm",10,[
    ("3.3","Cement mortar 1:3 (fine sand) – SH Mortars","cum",0.018,4181.95),
    ("0155","Mason (average)","day",0.94,709.00),
    ("0115","Coolie","day",1.04,558.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],162.50),
  # ── 13.32 ───────────────────────────────────────────────────────────────
  ("13.32.1","13.32 Pointing on tile brick work CM 1:3 fine sand",
   "Flush/Ruled/Struck or weathered pointing","sqm",10,[
    ("3.3","Cement mortar 1:3 (fine sand) – SH Mortars","cum",0.009,4181.95),
    ("0155","Mason (average)","day",0.94,709.00),
    ("0115","Coolie","day",1.04,558.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],159.60),
  # ── 13.33 ───────────────────────────────────────────────────────────────
  ("13.33.1","13.33 Pointing on stone work CM 1:3 fine sand",
   "Flush/Ruled pointing","sqm",10,[
    ("3.3","Cement mortar 1:3 (fine sand) – SH Mortars","cum",0.018,4181.95),
    ("0155","Mason (average)","day",0.94,709.00),
    ("0115","Coolie","day",1.04,558.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],162.50),
  ("13.33.2","13.33 Pointing on stone work CM 1:3 fine sand",
   "Raised and cut pointing","sqm",10,[
    ("3.3","Cement mortar 1:3 (fine sand) – SH Mortars","cum",0.018,4181.95),
    ("0155","Mason (average)","day",1.07,709.00),
    ("0115","Coolie","day",1.19,558.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],179.30),
  # ── 13.34 ───────────────────────────────────────────────────────────────
  ("13.34","13.34 Raised and cut pointing on stone work in white cement mortar 1:3",
   "Raised and cut pointing on stone work in white cement mortar 1:3","sqm",10,[
    ("3.3W","White cement mortar 1:3 – SH Mortars","cum",0.018,8000.00),
    ("0155","Mason (average)","day",1.07,709.00),
    ("0115","Coolie","day",1.19,558.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],248.20),
  # ── 13.35 ───────────────────────────────────────────────────────────────
  ("13.35.1","13.35 Pointing on stone slab ceiling CM 1:2 fine sand",
   "Flush/Ruled pointing","sqm",10,[
    ("3.2","Cement mortar 1:2 (fine sand) – SH Mortars","cum",0.009,4945.85),
    ("0155","Mason (average)","day",0.94,709.00),
    ("0115","Coolie","day",1.04,558.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],161.50),
  # ── 13.36 ───────────────────────────────────────────────────────────────
  ("13.36","13.36 Extra for pointing on outside walls above 10 m height",
   "Extra for pointing on walls on the outside at height more than 10 m","sqm",10,[
    ("9999","Extra allowance for height","L.S.",12.61,2.00)],12.20),
  # ── 13.37  White washing ─────────────────────────────────────────────────
  ("13.37.1","13.37 White washing with lime – even shade",
   "New work (three or more coats)","sqm",10,[
    ("9977","Carriage of lime","L.S.",0.91,2.00),
    ("0141","White Washer","day",0.20,617.00),
    ("0115","Coolie","day",0.10,558.00),
    ("9999","Indigo gum etc.","L.S.",4.42,2.00),
    ("9999","Sundries, ladders etc.","L.S.",2.73,2.00)],28.55),
  ("13.37.2","13.37 White washing with lime – even shade",
   "Old work (two or more coats)","sqm",10,[
    ("9977","Carriage of lime","L.S.",0.65,2.00),
    ("0141","White Washer","day",0.16,617.00),
    ("0115","Coolie","day",0.08,558.00),
    ("9999","Sundries, ladders etc.","L.S.",2.73,2.00)],21.45),
  # ── 13.38 ───────────────────────────────────────────────────────────────
  ("13.38","13.38 Satna lime wash on walls with one coat",
   "Satna lime wash on walls with one coat","sqm",10,[
    ("9977","Carriage of Satna lime","L.S.",1.30,2.00),
    ("0141","White Washer","day",0.12,617.00),
    ("0115","Coolie","day",0.06,558.00),
    ("9999","Sundries","L.S.",2.73,2.00)],16.50),
  # ── 13.39 ───────────────────────────────────────────────────────────────
  ("13.39.1","13.39 Colour washing – new work",
   "New work (two or more coats) with base coat of white washing with lime","sqm",10,[
    ("9977","Carriage of lime and colour","L.S.",1.56,2.00),
    ("0141","White Washer","day",0.24,617.00),
    ("0115","Coolie","day",0.12,558.00),
    ("9999","Sundries, colour pigment etc.","L.S.",7.15,2.00)],38.30),
  ("13.39.2","13.39 Colour washing – new work with base coat of whiting",
   "New work (two or more coats) with base coat of whiting","sqm",10,[
    ("9977","Carriage of material","L.S.",1.56,2.00),
    ("0141","White Washer","day",0.24,617.00),
    ("0115","Coolie","day",0.12,558.00),
    ("9999","Sundries, whiting, colour etc.","L.S.",9.75,2.00)],41.70),
  # ── 13.40  Dry distemper ─────────────────────────────────────────────────
  ("13.40","13.40 Distempering with dry distemper (two or more coats) new work",
   "Two or more coats on new work – dry distemper approved brand","sqm",10,[
    ("9999","Putty, glue etc.","L.S.",2.73,2.00),
    ("9977","Carriage of material","L.S.",1.56,2.00),
    ("9999","Brushes, sand paper etc.","L.S.",7.15,2.00),
    ("0131","Painter","day",0.80,679.00),
    ("0115","Coolie","day",0.40,558.00),
    ("9999","Sundries","L.S.",5.33,2.00)],117.75),
  # ── 13.41  Oil-bound washable distemper ──────────────────────────────────
  ("13.41.1","13.41 Distempering with oil bound washable distemper – new work",
   "New work (two or more coats) over and including water thinnable priming coat","sqm",10,[
    ("9999","Brushes, putty etc.","L.S.",7.15,2.00),
    ("9988","Sundries incl. carriage","L.S.",8.06,2.00),
    ("9977","Carriage of material","L.S.",4.42,2.00),
    ("0131","Painter","day",1.00,679.00),
    ("0115","Coolie","day",0.50,558.00),
    ("9999","Sundries","L.S.",8.06,2.00)],153.45),
  # ── 13.42  1st quality acrylic distemper – new work ──────────────────────
  ("13.42.1","13.42 Distempering with 1st quality acrylic distemper (ready mixed) VOC – new",
   "Two or more coats on new work","sqm",10,[
    ("0835","Acrylic distemper (ready mixed)","litre",1.43,90.00),
    ("9977","Carriage of material","L.S.",1.43,2.00),
    ("0131","Painter","day",0.80,679.00),
    ("0115","Coolie","day",0.40,558.00),
    ("9999","Brushes, sand paper etc.","L.S.",8.06,2.00)],178.00),
  # ── 13.43  Water thinnable cement primer ──────────────────────────────────
  ("13.43.1","13.43 Applying one coat water thinnable cement primer",
   "Water thinnable cement primer","sqm",10,[
    ("9977","Carriage of primer","L.S.",1.43,2.00),
    ("0131","Painter","day",0.54,679.00),
    ("0115","Coolie","day",0.27,558.00),
    ("9999","Brushes etc.","L.S.",4.42,2.00)],72.75),
  # ── 13.44  Water proofing cement paint ────────────────────────────────────
  ("13.44.1","13.44 Finishing walls with water proofing cement paint – new",
   "New work (two or more coats applied @ 3.84 kg/10 sqm)","sqm",10,[
    ("9977","Carriage of material","L.S.",2.86,2.00),
    ("0131","Painter","day",0.80,679.00),
    ("0115","Coolie","day",0.40,558.00),
    ("9999","Brushes, sand paper etc.","L.S.",8.06,2.00)],117.75),
  # ── 13.45  Textured exterior paint ────────────────────────────────────────
  ("13.45.1","13.45 Finishing walls with textured exterior paint – new",
   "New work (two or more coats @ 3.28 ltr/10 sqm) over and including priming","sqm",10,[
    ("9977","Carriage of material","L.S.",4.42,2.00),
    ("0131","Painter","day",1.00,679.00),
    ("0115","Coolie","day",0.50,558.00),
    ("9999","Brushes, sand paper etc.","L.S.",8.06,2.00)],153.45),
  # ── 13.46  Acrylic smooth exterior paint ──────────────────────────────────
  ("13.46.1","13.46 Finishing walls with Acrylic Smooth exterior paint – new",
   "New work (two or more coats @ 1.67 ltr/10 sqm) over and including priming","sqm",10,[
    ("9977","Carriage of material","L.S.",4.42,2.00),
    ("0131","Painter","day",1.00,679.00),
    ("0115","Coolie","day",0.50,558.00),
    ("9999","Brushes, sand paper etc.","L.S.",8.06,2.00)],153.45),
  # ── 13.47  Premium acrylic smooth exterior paint with silicone ────────────
  ("13.47.1","13.47 Finishing walls with Premium Acrylic Smooth Exterior Paint (Silicone) – new",
   "New work (two or more coats @ 1.43 ltr/10 sqm) over and including priming","sqm",10,[
    ("9977","Carriage of material","L.S.",4.42,2.00),
    ("0131","Painter","day",1.00,679.00),
    ("0115","Coolie","day",0.50,558.00),
    ("9999","Brushes, sand paper etc.","L.S.",8.06,2.00)],153.45),
  # ── 13.48  Deluxe multi-surface paint ────────────────────────────────────
  ("13.48.1","13.48 Deluxe Multi Surface Paint – walls",
   "Two or more coats on walls @ 1.25 ltr/10 sqm over and including one priming coat","sqm",10,[
    ("9977","Carriage of material","L.S.",4.42,2.00),
    ("0131","Painter","day",1.00,679.00),
    ("0115","Coolie","day",0.50,558.00),
    ("9999","Brushes, sand paper etc.","L.S.",8.06,2.00)],153.45),
  ("13.48.2","13.48 Deluxe Multi Surface Paint – wood work",
   "Painting wood work with Deluxe Multi Surface Paint, two or more coats","sqm",10,[
    ("9977","Carriage of material","L.S.",4.42,2.00),
    ("0131","Painter","day",1.20,679.00),
    ("0115","Coolie","day",0.60,558.00),
    ("9999","Brushes, sand paper etc.","L.S.",8.06,2.00)],175.80),
  ("13.48.3","13.48 Deluxe Multi Surface Paint – steel work",
   "Painting steel work with Deluxe Multi Surface Paint, two or more coats","sqm",10,[
    ("9977","Carriage of material","L.S.",4.42,2.00),
    ("0131","Painter","day",1.40,679.00),
    ("0115","Coolie","day",0.70,558.00),
    ("9999","Brushes, sand paper etc.","L.S.",8.06,2.00)],198.95),
  # ── 13.50  Priming coats ─────────────────────────────────────────────────
  ("13.50.1","13.50 Applying priming coat",
   "Ready mixed pink or grey primer on wood work (hard and soft wood)","sqm",10,[
    ("9977","Carriage of primer","L.S.",1.30,2.00),
    ("0131","Painter","day",0.54,679.00),
    ("0115","Coolie","day",0.27,558.00),
    ("9999","Brushes etc.","L.S.",4.42,2.00)],72.75),
  ("13.50.2","13.50 Applying priming coat",
   "Ready mixed aluminium primer on wood work","sqm",10,[
    ("9977","Carriage of primer","L.S.",1.30,2.00),
    ("0131","Painter","day",0.54,679.00),
    ("0115","Coolie","day",0.27,558.00),
    ("9999","Brushes etc.","L.S.",4.42,2.00)],72.75),
  ("13.50.3","13.50 Applying priming coat",
   "Ready mixed red oxide zinc chromate primer on steel and iron work (one coat)","sqm",10,[
    ("9977","Carriage of primer","L.S.",1.30,2.00),
    ("0131","Painter","day",0.54,679.00),
    ("0115","Coolie","day",0.27,558.00),
    ("9999","Brushes etc.","L.S.",4.42,2.00)],72.75),
  ("13.50.4","13.50 Applying priming coat",
   "Ready mixed red oxide zinc chromate primer on steel and iron work (two coats)","sqm",10,[
    ("9977","Carriage of primer","L.S.",2.60,2.00),
    ("0131","Painter","day",1.07,679.00),
    ("0115","Coolie","day",0.54,558.00),
    ("9999","Brushes etc.","L.S.",8.06,2.00)],138.45),
  # ── 13.51  Silicon & acrylic emulsion sealer ──────────────────────────────
  ("13.51.1","13.51 Painting with silicon & acrylic emulsion sealer",
   "One coat","sqm",10,[
    ("9977","Carriage of material","L.S.",2.86,2.00),
    ("0131","Painter","day",0.54,679.00),
    ("0115","Coolie","day",0.27,558.00),
    ("9999","Brushes etc.","L.S.",4.42,2.00)],80.80),
  ("13.51.2","13.51 Painting with silicon & acrylic emulsion sealer",
   "Two coats","sqm",10,[
    ("9977","Carriage of material","L.S.",4.42,2.00),
    ("0131","Painter","day",1.00,679.00),
    ("0115","Coolie","day",0.50,558.00),
    ("9999","Brushes etc.","L.S.",7.15,2.00)],144.45),
  # ── 13.52  Epoxy paint ────────────────────────────────────────────────────
  ("13.52.1","13.52 Finishing with Epoxy paint (two or more coats)",
   "On steel work","sqm",10,[
    ("9977","Carriage of epoxy paint","L.S.",4.42,2.00),
    ("0131","Painter","day",1.40,679.00),
    ("0115","Coolie","day",0.70,558.00),
    ("9999","Brushes, thinner etc.","L.S.",8.06,2.00)],199.70),
  ("13.52.2","13.52 Finishing with Epoxy paint (two or more coats)",
   "On concrete work","sqm",10,[
    ("9977","Carriage of epoxy paint","L.S.",4.42,2.00),
    ("0131","Painter","day",1.20,679.00),
    ("0115","Coolie","day",0.60,558.00),
    ("9999","Brushes, thinner etc.","L.S.",8.06,2.00)],176.55),
  # ── 13.53  Synthetic enamel on G.S. sheet ────────────────────────────────
  ("13.53.1","13.53 Painting on G.S. sheet with synthetic enamel – new",
   "New work (two or more coats) including approved steel primer but excl. mordant","sqm",10,[
    ("9977","Carriage of paint","L.S.",4.42,2.00),
    ("0131","Painter","day",1.40,679.00),
    ("0115","Coolie","day",0.70,558.00),
    ("9999","Brushes, sand paper etc.","L.S.",8.06,2.00)],199.70),
  # ── 13.54  Mordant solution on G.S. sheet ────────────────────────────────
  ("13.54.1","13.54 Applying mordant solution on G.S. sheet",
   "Copper acetate solution (38 gm in 1 litre soft water)","sqm",10,[
    ("9977","Carriage of chemical","L.S.",0.91,2.00),
    ("0131","Painter","day",0.27,679.00),
    ("0115","Coolie","day",0.14,558.00),
    ("9999","Sundries","L.S.",2.73,2.00)],37.85),
  ("13.54.2","13.54 Applying mordant solution on G.S. sheet",
   "HCl/copper sulphate solution (13 gm HCl + 13 gm copper sulphate in 1 litre water)","sqm",10,[
    ("9977","Carriage of chemical","L.S.",0.91,2.00),
    ("0131","Painter","day",0.27,679.00),
    ("0115","Coolie","day",0.14,558.00),
    ("9999","Sundries","L.S.",2.73,2.00)],37.85),
  # ── 13.55  Painting on pipes – new 2 coats ────────────────────────────────
  ("13.55.1","13.55 Painting (two or more coats) on rain water/soil pipes – new",
   "100 mm diameter pipes","Rmt",10,[
    ("9977","Carriage of paint","L.S.",2.60,2.00),
    ("0131","Painter","day",0.80,679.00),
    ("0115","Coolie","day",0.40,558.00),
    ("9999","Brushes etc.","L.S.",4.42,2.00)],117.75),
  ("13.55.2","13.55 Painting (two or more coats) on rain water/soil pipes – new",
   "150 mm diameter pipes","Rmt",10,[
    ("9977","Carriage of paint","L.S.",4.42,2.00),
    ("0131","Painter","day",1.07,679.00),
    ("0115","Coolie","day",0.54,558.00),
    ("9999","Brushes etc.","L.S.",5.33,2.00)],150.15),
  # ── 13.56  Painting on pipes with primer – new ────────────────────────────
  ("13.56.1","13.56 Painting (two or more coats) on pipes with primer – new",
   "100 mm diameter pipes","Rmt",10,[
    ("9977","Carriage of paint","L.S.",2.60,2.00),
    ("0131","Painter","day",1.00,679.00),
    ("0115","Coolie","day",0.50,558.00),
    ("9999","Brushes etc.","L.S.",4.42,2.00)],141.65),
  ("13.56.2","13.56 Painting (two or more coats) on pipes with primer – new",
   "150 mm diameter pipes","Rmt",10,[
    ("9977","Carriage of paint","L.S.",4.42,2.00),
    ("0131","Painter","day",1.34,679.00),
    ("0115","Coolie","day",0.67,558.00),
    ("9999","Brushes etc.","L.S.",5.33,2.00)],183.60),
  # ── 13.57  Oil type wood preservative – new ───────────────────────────────
  ("13.57.1","13.57 Painting with oil type wood preservative – new",
   "New work (two or more coats)","sqm",10,[
    ("9977","Carriage of preservative","L.S.",2.86,2.00),
    ("0131","Painter","day",0.80,679.00),
    ("0115","Coolie","day",0.40,558.00),
    ("9999","Brushes etc.","L.S.",7.15,2.00)],120.55),
  # ── 13.58 ───────────────────────────────────────────────────────────────
  ("13.58","13.58 Two coats of fire retardant paint on cleaned wood/ply",
   "Providing and applying two coats of fire retardant paint on cleaned wood/ply","sqm",10,[
    ("9977","Carriage of fire retardant paint","L.S.",4.42,2.00),
    ("0131","Painter","day",1.20,679.00),
    ("0115","Coolie","day",0.60,558.00),
    ("9999","Brushes etc.","L.S.",8.06,2.00)],175.80),
  # ── 13.59 ───────────────────────────────────────────────────────────────
  ("13.59","13.59 Coal tarring two coats – new work",
   "Coal tarring two coats using 0.16 litre and 0.12 litre coal tar per sqm","sqm",10,[
    ("9977","Carriage of coal tar","L.S.",1.30,2.00),
    ("0131","Painter","day",0.40,679.00),
    ("0115","Coolie","day",0.20,558.00),
    ("9999","Brushes etc.","L.S.",4.42,2.00)],59.30),
  # ── 13.60  Acrylic emulsion paint ────────────────────────────────────────
  ("13.60.1","13.60 Wall painting with acrylic emulsion paint – new",
   "Two or more coats on new work","sqm",10,[
    ("0835","Plastic emulsion paint","litre",1.21,200.00),
    ("9977","Carriage of material","L.S.",1.43,2.00),
    ("0131","Painter","day",0.54,679.00),
    ("0115","Coolie","day",0.54,558.00),
    ("9999","Brushes, sand paper etc.","L.S.",10.79,2.00)],128.65),
  # ── 13.61  Synthetic enamel – general – new ───────────────────────────────
  ("13.61.1","13.61 Painting with synthetic enamel paint – general – new",
   "Two or more coats on new work","sqm",10,[
    ("9977","Carriage of paint","L.S.",4.42,2.00),
    ("0131","Painter","day",1.20,679.00),
    ("0115","Coolie","day",0.60,558.00),
    ("9999","Brushes, sand paper etc.","L.S.",8.06,2.00)],175.80),
  # ── 13.62  Synthetic enamel – with under coat ─────────────────────────────
  ("13.62.1","13.62 Painting with synthetic enamel – with under coat – new",
   "Two or more coats on new work with under coat of suitable shade","sqm",10,[
    ("9977","Carriage of paint","L.S.",4.42,2.00),
    ("0131","Painter","day",1.40,679.00),
    ("0115","Coolie","day",0.70,558.00),
    ("9999","Brushes, sand paper etc.","L.S.",8.06,2.00)],199.70),
  # ── 13.63  Aluminium paint ────────────────────────────────────────────────
  ("13.63.1","13.63 Painting with aluminium paint – new",
   "Two or more coats on new work","sqm",10,[
    ("9977","Carriage of paint","L.S.",4.42,2.00),
    ("0131","Painter","day",1.20,679.00),
    ("0115","Coolie","day",0.60,558.00),
    ("9999","Brushes etc.","L.S.",8.06,2.00)],175.80),
  # ── 13.64  Acid proof paint ───────────────────────────────────────────────
  ("13.64.1","13.64 Painting with acid proof paint – new",
   "Two or more coats on new work","sqm",10,[
    ("9977","Carriage of paint","L.S.",4.42,2.00),
    ("0131","Painter","day",1.40,679.00),
    ("0115","Coolie","day",0.70,558.00),
    ("9999","Brushes, thinner etc.","L.S.",8.06,2.00)],199.70),
  # ── 13.65  Black anti-corrosive bitumastic paint ──────────────────────────
  ("13.65.1","13.65 Painting with black anti-corrosive bitumastic paint – new",
   "Two or more coats on new work","sqm",10,[
    ("9977","Carriage of paint","L.S.",2.86,2.00),
    ("0131","Painter","day",0.80,679.00),
    ("0115","Coolie","day",0.40,558.00),
    ("9999","Brushes etc.","L.S.",7.15,2.00)],117.75),
  # ── 13.66  Floor enamel paint ─────────────────────────────────────────────
  ("13.66.1","13.66 Floor painting with floor enamel paint – new",
   "Two or more coats on new work","sqm",10,[
    ("9977","Carriage of paint","L.S.",4.42,2.00),
    ("0131","Painter","day",1.20,679.00),
    ("0115","Coolie","day",0.60,558.00),
    ("9999","Brushes etc.","L.S.",8.06,2.00)],175.80),
  # ── 13.67  Varnishing ────────────────────────────────────────────────────
  ("13.67.1","13.67 Varnishing",
   "Two or more coats glue sizing with copal varnish over under coat of flatting","sqm",10,[
    ("9977","Carriage of varnish","L.S.",4.42,2.00),
    ("0131","Painter","day",1.60,679.00),
    ("0115","Coolie","day",0.80,558.00),
    ("9999","Brushes, sand paper, pumice stone etc.","L.S.",8.06,2.00)],221.70),
  ("13.67.2","13.67 Varnishing",
   "Two or more coats glue sizing with spar varnish over under coat of flatting","sqm",10,[
    ("9977","Carriage of varnish","L.S.",4.42,2.00),
    ("0131","Painter","day",1.60,679.00),
    ("0115","Coolie","day",0.80,558.00),
    ("9999","Brushes, sand paper, pumice stone etc.","L.S.",8.06,2.00)],221.70),
  # ── 13.68  French spirit polishing ────────────────────────────────────────
  ("13.68.1","13.68 French spirit polishing",
   "Two or more coats on new works including a coat of wood filler","sqm",10,[
    ("9977","Carriage of shellac, methylated spirit","L.S.",4.42,2.00),
    ("0131","Painter","day",2.00,679.00),
    ("0115","Coolie","day",1.00,558.00),
    ("9999","Brushes, pumice stone, cotton wool etc.","L.S.",10.79,2.00)],268.55),
  # ── 13.69  Wax polish – wood work ────────────────────────────────────────
  ("13.69.1","13.69 Polishing on wood work with wax polish – new",
   "New work","sqm",10,[
    ("9977","Carriage of wax polish","L.S.",2.86,2.00),
    ("0131","Painter","day",1.40,679.00),
    ("0115","Coolie","day",0.70,558.00),
    ("9999","Brushes, sand paper etc.","L.S.",8.06,2.00)],199.70),
  # ── 13.70  Floor polishing – masonry/concrete ────────────────────────────
  ("13.70","13.70 Floor polishing on masonry or concrete floors with wax polish",
   "Floor polishing on masonry or concrete floors with wax polish","sqm",10,[
    ("9977","Carriage of wax polish","L.S.",2.86,2.00),
    ("0131","Painter","day",1.40,679.00),
    ("0115","Coolie","day",0.70,558.00),
    ("9999","Brushes etc.","L.S.",8.06,2.00)],199.70),
  # ── 13.71  Lettering ─────────────────────────────────────────────────────
  ("13.71","13.71 Lettering with black Japan paint",
   "Lettering with black Japan paint (25 mm height letters)","letters",100,[
    ("9977","Carriage of Japan paint","L.S.",0.52,2.00),
    ("0131","Painter","day",1.00,679.00),
    ("0115","Coolie","day",0.50,558.00),
    ("9999","Brushes etc.","L.S.",5.33,2.00)],4.70),
  # ── 13.72  Washed stone grit plaster ─────────────────────────────────────
  ("13.72","13.72 Washed stone grit plaster on exterior walls up to 10 m",
   "Washed stone grit plaster on exterior walls height up to 10 m above ground level","sqm",10,[
    ("3.4","Cement mortar 1:4 (fine sand) – SH Mortars","cum",0.172,3528.85),
    ("0155","Mason (average)","day",1.34,709.00),
    ("0115","Coolie","day",1.49,558.00),
    ("0101","Bhisti","day",0.99,617.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],385.80),
  # ── 13.73  Forming groove in washed stone grit plaster ───────────────────
  ("13.73.1","13.73 Forming groove in top layer of washed stone grit plaster",
   "15 mm wide and 15 mm deep groove","Rmt",10,[
    ("0155","Mason (average)","day",0.27,709.00),
    ("9999","Sundries","L.S.",2.73,2.00)],27.65),
  ("13.73.2","13.73 Forming groove in top layer of washed stone grit plaster",
   "20 mm wide and 15 mm deep groove","Rmt",10,[
    ("0155","Mason (average)","day",0.33,709.00),
    ("9999","Sundries","L.S.",2.73,2.00)],32.35),
  # ── 13.74 ───────────────────────────────────────────────────────────────
  ("13.74","13.74 Extra for washed grit plaster on exterior walls above 10 m",
   "Extra for washed stone grit plaster on exterior walls of height more than 10 m","sqm",10,[
    ("9999","Extra allowance for height","L.S.",12.61,2.00)],13.00),
  # ── 13.75 ───────────────────────────────────────────────────────────────
  ("13.75","13.75 Extra for washed stone grit plaster on circular work",
   "Extra for washed stone grit plaster on circular work not exceeding 6 m radius","sqm",10,[
    ("9999","Extra for circular work","L.S.",12.61,2.00)],15.55),
  # ── 13.76 ───────────────────────────────────────────────────────────────
  ("13.76","13.76 Forming groove from 12x12 mm up to 25x15 mm in top layer",
   "Forming groove of uniform size from 12x12 mm and up to 25x15 mm in top layer","Rmt",10,[
    ("0155","Mason (average)","day",0.53,709.00),
    ("9999","Sundries","L.S.",2.73,2.00)],50.00),
  # ── 13.77 ───────────────────────────────────────────────────────────────
  ("13.77","13.77 Extra for white cement in top layer of washed stone grit plaster",
   "Extra for using white cement in place of ordinary cement in the top layer","sqm",10,[
    ("9999","Extra cost of white cement over OPC","L.S.",12.61,2.00)],43.00),
  # ── 13.78 ───────────────────────────────────────────────────────────────
  ("13.78","13.78 12 mm thick premixed formulated one coat plaster",
   "Providing and applying 12 mm thick (average) premixed formulated one coat plaster","sqm",10,[
    ("9977","Premixed plaster material and carriage","L.S.",12.61,2.00),
    ("0155","Mason (average)","day",0.67,709.00),
    ("0115","Coolie","day",0.75,558.00),
    ("0101","Bhisti","day",0.92,617.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],278.40),
  # ── 13.79 ───────────────────────────────────────────────────────────────
  ("13.79","13.79 Extra for addition of synthetic polyester fibre 6 mm in plaster",
   "Extra for addition of synthetic polyester triangular fibre 6 mm length in plaster","sqm",10,[
    ("9999","Synthetic polyester fibre @ 50 gm/sqm","L.S.",12.61,2.00)],16.55),
  # ── 13.80 ───────────────────────────────────────────────────────────────
  ("13.80","13.80 White cement based putty 1 mm average thickness",
   "Providing and applying white cement based putty of average thickness 1 mm","sqm",10,[
    ("2209","White cement based putty","kg",1.50,40.00),
    ("9977","Carriage of putty","L.S.",1.43,2.00),
    ("0155","Mason (average)","day",0.20,709.00),
    ("0115","Coolie","day",0.22,558.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],14.58),
  # ── 13.81  1st quality acrylic distemper VOC – new ───────────────────────
  ("13.81.1","13.81 Distempering with 1st quality acrylic distemper VOC – new",
   "One coat","sqm",10,[
    ("0835","Acrylic distemper (ready mixed)","litre",0.75,90.00),
    ("9977","Carriage of material","L.S.",0.91,2.00),
    ("0131","Painter","day",0.40,679.00),
    ("0115","Coolie","day",0.20,558.00),
    ("9999","Brushes etc.","L.S.",4.42,2.00)],95.75),
  ("13.81.2","13.81 Distempering with 1st quality acrylic distemper VOC – new",
   "Two coats","sqm",10,[
    ("0835","Acrylic distemper (ready mixed)","litre",1.43,90.00),
    ("9977","Carriage of material","L.S.",1.43,2.00),
    ("0131","Painter","day",0.80,679.00),
    ("0115","Coolie","day",0.40,558.00),
    ("9999","Brushes etc.","L.S.",7.15,2.00)],178.00),
  # ── 13.82  Acrylic emulsion paint VOC – new ───────────────────────────────
  ("13.82.1","13.82 Wall painting with acrylic emulsion paint VOC – new",
   "One coat","sqm",10,[
    ("0835","Acrylic emulsion paint","litre",0.63,200.00),
    ("9977","Carriage of material","L.S.",0.91,2.00),
    ("0131","Painter","day",0.27,679.00),
    ("0115","Coolie","day",0.27,558.00),
    ("9999","Brushes etc.","L.S.",5.33,2.00)],69.05),
  ("13.82.2","13.82 Wall painting with acrylic emulsion paint VOC – new",
   "Two coats","sqm",10,[
    ("0835","Acrylic emulsion paint","litre",1.21,200.00),
    ("9977","Carriage of material","L.S.",1.43,2.00),
    ("0131","Painter","day",0.54,679.00),
    ("0115","Coolie","day",0.54,558.00),
    ("9999","Brushes etc.","L.S.",10.79,2.00)],128.65),
  # ── 13.83  Premium acrylic emulsion interior paint VOC ────────────────────
  ("13.83.1","13.83 Wall painting with premium acrylic emulsion interior paint VOC – new",
   "One coat","sqm",10,[
    ("9977","Carriage of premium paint","L.S.",2.86,2.00),
    ("0131","Painter","day",0.27,679.00),
    ("0115","Coolie","day",0.27,558.00),
    ("9999","Brushes etc.","L.S.",5.33,2.00)],73.45),
  ("13.83.2","13.83 Wall painting with premium acrylic emulsion interior paint VOC – new",
   "Two coats","sqm",10,[
    ("9977","Carriage of premium paint","L.S.",4.42,2.00),
    ("0131","Painter","day",0.54,679.00),
    ("0115","Coolie","day",0.54,558.00),
    ("9999","Brushes etc.","L.S.",10.79,2.00)],139.15),
  # ── 13.85  Primer – VOC – new ─────────────────────────────────────────────
  ("13.85.1","13.85 Applying priming coat with primer VOC",
   "Ready mixed pink or grey primer on wood work having VOC < 150 g/litre","sqm",10,[
    ("9977","Carriage of primer","L.S.",1.30,2.00),
    ("0131","Painter","day",0.54,679.00),
    ("0115","Coolie","day",0.27,558.00),
    ("9999","Brushes etc.","L.S.",4.42,2.00)],72.75),
  ("13.85.3","13.85 Applying priming coat with primer VOC",
   "Water thinnable cement primer on wall surface having VOC < 150 g/litre","sqm",10,[
    ("9977","Carriage of primer","L.S.",1.30,2.00),
    ("0131","Painter","day",0.54,679.00),
    ("0115","Coolie","day",0.27,558.00),
    ("9999","Brushes etc.","L.S.",4.42,2.00)],72.75),
  # ── 13.86  6 mm plaster on RCC work ──────────────────────────────────────
  ("13.86","13.86 6 mm plaster on cement concrete or RCC work",
   "6 mm plaster on cement concrete or RCC work with cement mortar 1:3","sqm",10,[
    ("3.3","Cement mortar 1:3 (fine sand) – SH Mortars","cum",0.072,4181.95),
    ("0155","Mason (average)","day",0.53,709.00),
    ("0115","Coolie","day",0.59,558.00),
    ("0101","Bhisti","day",0.61,617.00),
    ("9999","Scaffolding and sundries","L.S.",12.61,2.00)],196.70),
  # ── 13.87  White washing – old work ───────────────────────────────────────
  ("13.87.1","13.87 White washing with lime – old work",
   "Old work (two or more coats)","sqm",10,[
    ("9977","Carriage of lime","L.S.",0.65,2.00),
    ("0141","White Washer","day",0.16,617.00),
    ("0115","Coolie","day",0.08,558.00),
    ("9999","Sundries, ladders etc.","L.S.",2.73,2.00)],21.45),
  ("13.87.2","13.87 White washing with lime – old work",
   "Old work (one or more coats)","sqm",10,[
    ("9977","Carriage of lime","L.S.",0.39,2.00),
    ("0141","White Washer","day",0.10,617.00),
    ("0115","Coolie","day",0.05,558.00),
    ("9999","Sundries, ladders etc.","L.S.",2.73,2.00)],13.55),
  # ── 13.88 ───────────────────────────────────────────────────────────────
  ("13.88","13.88 Removing white/colour wash by scraping and preparing surface",
   "Removing white or colour wash by scraping and sand papering and preparing surface","sqm",10,[
    ("0115","Coolie","day",0.40,558.00),
    ("9999","Sand paper, scrapers etc.","L.S.",2.73,2.00)],31.65),
  # ── 13.89  Dry distemper – old ────────────────────────────────────────────
  ("13.89","13.89 Distempering with dry distemper (one or more coats) old work",
   "One or more coats on old work","sqm",10,[
    ("9999","Putty, glue etc.","L.S.",1.56,2.00),
    ("9977","Carriage of material","L.S.",0.78,2.00),
    ("0131","Painter","day",0.40,679.00),
    ("0115","Coolie","day",0.20,558.00),
    ("9999","Brushes, sand paper etc.","L.S.",4.42,2.00)],63.90),
  # ── 13.90  Acrylic distemper old ─────────────────────────────────────────
  ("13.90.1","13.90 Distempering with 1st quality acrylic distemper – old",
   "Old work (one or more coats)","sqm",10,[
    ("0835","Acrylic distemper","litre",0.75,90.00),
    ("9977","Carriage of material","L.S.",0.91,2.00),
    ("0131","Painter","day",0.40,679.00),
    ("0115","Coolie","day",0.20,558.00),
    ("9999","Brushes etc.","L.S.",4.42,2.00)],95.75),
  # ── 13.91 ───────────────────────────────────────────────────────────────
  ("13.91","13.91 Removing dry/oil bound distemper, WP cement paint and preparing",
   "Removing dry or oil bound distemper, water proofing cement paint and preparing","sqm",10,[
    ("0115","Coolie","day",0.67,558.00),
    ("9999","Sand paper, scrapers etc.","L.S.",2.73,2.00)],46.85),
  # ── 13.92  Synthetic enamel on G.S. – old ────────────────────────────────
  ("13.92.1","13.92 Painting on G.S. sheet with synthetic enamel – old",
   "Old work (one or more coats)","sqm",10,[
    ("9977","Carriage of paint","L.S.",2.86,2.00),
    ("0131","Painter","day",0.80,679.00),
    ("0115","Coolie","day",0.40,558.00),
    ("9999","Brushes, sand paper etc.","L.S.",5.33,2.00)],117.75),
  # ── 13.93  Painting on 75 mm pipes – new two coats ────────────────────────
  ("13.93.1","13.93 Painting (two or more coats) on 75 mm dia pipes – new",
   "75 mm diameter pipes","Rmt",10,[
    ("9977","Carriage of paint","L.S.",1.30,2.00),
    ("0131","Painter","day",0.54,679.00),
    ("0115","Coolie","day",0.27,558.00),
    ("9999","Brushes etc.","L.S.",4.42,2.00)],80.80),
  # ── 13.94  Painting on pipes – old one coat ───────────────────────────────
  ("13.94.1","13.94 Painting (one or more coats) on pipes – old",
   "75 mm diameter pipes","Rmt",10,[
    ("9977","Carriage of paint","L.S.",0.65,2.00),
    ("0131","Painter","day",0.27,679.00),
    ("0115","Coolie","day",0.14,558.00),
    ("9999","Brushes etc.","L.S.",2.73,2.00)],41.65),
  ("13.94.2","13.94 Painting (one or more coats) on pipes – old",
   "100 mm diameter pipes","Rmt",10,[
    ("9977","Carriage of paint","L.S.",1.30,2.00),
    ("0131","Painter","day",0.40,679.00),
    ("0115","Coolie","day",0.20,558.00),
    ("9999","Brushes etc.","L.S.",4.42,2.00)],63.90),
  ("13.94.3","13.94 Painting (one or more coats) on pipes – old",
   "150 mm diameter pipes","Rmt",10,[
    ("9977","Carriage of paint","L.S.",2.60,2.00),
    ("0131","Painter","day",0.67,679.00),
    ("0115","Coolie","day",0.34,558.00),
    ("9999","Brushes etc.","L.S.",5.33,2.00)],97.25),
  # ── 13.95  Painting on pipes – new two coats with primer ──────────────────
  ("13.95.1","13.95 Painting (two or more coats) on 75 mm dia pipes – new with primer",
   "75 mm diameter pipes","Rmt",10,[
    ("9977","Carriage of paint","L.S.",1.30,2.00),
    ("0131","Painter","day",0.67,679.00),
    ("0115","Coolie","day",0.34,558.00),
    ("9999","Brushes etc.","L.S.",4.42,2.00)],98.05),
  ("13.95.2","13.95 Painting (two or more coats) on 100 mm dia pipes – new with primer",
   "100 mm diameter pipes","Rmt",10,[
    ("9977","Carriage of paint","L.S.",2.60,2.00),
    ("0131","Painter","day",1.00,679.00),
    ("0115","Coolie","day",0.50,558.00),
    ("9999","Brushes etc.","L.S.",5.33,2.00)],141.65),
  ("13.95.3","13.95 Painting (two or more coats) on 150 mm dia pipes – new with primer",
   "150 mm diameter pipes","Rmt",10,[
    ("9977","Carriage of paint","L.S.",4.42,2.00),
    ("0131","Painter","day",1.34,679.00),
    ("0115","Coolie","day",0.67,558.00),
    ("9999","Brushes etc.","L.S.",5.33,2.00)],183.60),
  # ── 13.96  Painting on pipes – old one coat with primer ───────────────────
  ("13.96.1","13.96 Painting (one or more coats) on 75 mm dia pipes – old with primer",
   "75 mm diameter pipes","Rmt",10,[
    ("9977","Carriage of paint","L.S.",0.65,2.00),
    ("0131","Painter","day",0.34,679.00),
    ("0115","Coolie","day",0.17,558.00),
    ("9999","Brushes etc.","L.S.",2.73,2.00)],51.55),
  ("13.96.2","13.96 Painting (one or more coats) on 100 mm dia pipes – old with primer",
   "100 mm diameter pipes","Rmt",10,[
    ("9977","Carriage of paint","L.S.",1.30,2.00),
    ("0131","Painter","day",0.54,679.00),
    ("0115","Coolie","day",0.27,558.00),
    ("9999","Brushes etc.","L.S.",4.42,2.00)],80.80),
  ("13.96.3","13.96 Painting (one or more coats) on 150 mm dia pipes – old with primer",
   "150 mm diameter pipes","Rmt",10,[
    ("9977","Carriage of paint","L.S.",2.60,2.00),
    ("0131","Painter","day",0.80,679.00),
    ("0115","Coolie","day",0.40,558.00),
    ("9999","Brushes etc.","L.S.",5.33,2.00)],117.75),
  # ── 13.97  Oil type wood preservative – old ───────────────────────────────
  ("13.97.1","13.97 Painting with oil type wood preservative – old",
   "Old work (one or more coats)","sqm",10,[
    ("9977","Carriage of preservative","L.S.",1.56,2.00),
    ("0131","Painter","day",0.40,679.00),
    ("0115","Coolie","day",0.20,558.00),
    ("9999","Brushes etc.","L.S.",4.42,2.00)],64.65),
  # ── 13.98  Plastic emulsion – old ────────────────────────────────────────
  ("13.98.1","13.98 Wall painting with plastic emulsion paint – old",
   "One or more coats on old work","sqm",10,[
    ("0835","Plastic emulsion paint","litre",0.63,200.00),
    ("9977","Carriage of material","L.S.",0.91,2.00),
    ("0131","Painter","day",0.27,679.00),
    ("0115","Coolie","day",0.27,558.00),
    ("9999","Brushes etc.","L.S.",5.33,2.00)],69.05),
  # ── 13.99  Synthetic enamel – old ────────────────────────────────────────
  ("13.99.1","13.99 Painting with synthetic enamel paint – old",
   "One or more coats on old work","sqm",10,[
    ("9977","Carriage of paint","L.S.",2.86,2.00),
    ("0131","Painter","day",0.80,679.00),
    ("0115","Coolie","day",0.40,558.00),
    ("9999","Brushes, sand paper etc.","L.S.",5.33,2.00)],117.75),
  # ── 13.100  Aluminium paint – old ────────────────────────────────────────
  ("13.100.1","13.100 Painting with aluminium paint – old",
   "One or more coats on old work","sqm",10,[
    ("9977","Carriage of paint","L.S.",2.86,2.00),
    ("0131","Painter","day",0.80,679.00),
    ("0115","Coolie","day",0.40,558.00),
    ("9999","Brushes etc.","L.S.",5.33,2.00)],117.75),
  # ── 13.101  Acid proof – old ──────────────────────────────────────────────
  ("13.101.1","13.101 Painting with acid proof paint – old",
   "One or more coats on old work","sqm",10,[
    ("9977","Carriage of paint","L.S.",2.86,2.00),
    ("0131","Painter","day",0.80,679.00),
    ("0115","Coolie","day",0.40,558.00),
    ("9999","Brushes, thinner etc.","L.S.",5.33,2.00)],117.75),
  # ── 13.102  Bitumastic paint – old ────────────────────────────────────────
  ("13.102.1","13.102 Painting with black anti-corrosive bitumastic paint – old",
   "One or more coats on old work","sqm",10,[
    ("9977","Carriage of paint","L.S.",1.56,2.00),
    ("0131","Painter","day",0.40,679.00),
    ("0115","Coolie","day",0.20,558.00),
    ("9999","Brushes etc.","L.S.",4.42,2.00)],64.65),
  # ── 13.103  French spirit polishing – old ────────────────────────────────
  ("13.103.1","13.103 French spirit polishing – old",
   "One or more coats on old work","sqm",10,[
    ("9977","Carriage of shellac, spirit","L.S.",2.86,2.00),
    ("0131","Painter","day",1.20,679.00),
    ("0115","Coolie","day",0.60,558.00),
    ("9999","Cotton wool, pumice stone etc.","L.S.",7.15,2.00)],175.80),
  # ── 13.104  Wax polish – wood – old ──────────────────────────────────────
  ("13.104.1","13.104 Polishing on wood work with wax polish – old",
   "Old work","sqm",10,[
    ("9977","Carriage of wax polish","L.S.",1.56,2.00),
    ("0131","Painter","day",0.80,679.00),
    ("0115","Coolie","day",0.40,558.00),
    ("9999","Brushes, sand paper etc.","L.S.",5.33,2.00)],117.75),
  # ── 13.105 ───────────────────────────────────────────────────────────────
  ("13.105","13.105 Re-lettering with black Japan paint",
   "Re-lettering with black Japan paint (25 mm height letters)","letters",100,[
    ("9977","Carriage of Japan paint","L.S.",0.26,2.00),
    ("0131","Painter","day",0.50,679.00),
    ("0115","Coolie","day",0.25,558.00),
    ("9999","Brushes etc.","L.S.",2.73,2.00)],2.65),
  # ── 13.106 ───────────────────────────────────────────────────────────────
  ("13.106","13.106 Painting (one or more coats) with black Japan paint",
   "Painting (one or more coats) with black Japan paint on new work","sqm",10,[
    ("9977","Carriage of Japan paint","L.S.",2.86,2.00),
    ("0131","Painter","day",0.80,679.00),
    ("0115","Coolie","day",0.40,558.00),
    ("9999","Brushes etc.","L.S.",5.33,2.00)],117.75),
  # ── 13.107  C.P. brass chain and rubber plug ──────────────────────────────
  ("13.107.1","13.107 Providing and fixing C.P. brass chain and rubber plug",
   "32 mm dia","no",1,[
    ("8509","CP brass chain with rubber plug 32 mm","no",1.0,200.00),
    ("0367","Plumber","day",0.07,750.00),
    ("0115","Coolie","day",0.04,558.00)],75.10),
  ("13.107.2","13.107 Providing and fixing C.P. brass chain and rubber plug",
   "40 mm dia","no",1,[
    ("8510","CP brass chain with rubber plug 40 mm","no",1.0,250.00),
    ("0367","Plumber","day",0.07,750.00),
    ("0115","Coolie","day",0.04,558.00)],89.45),
  # ── 13.108  Acrylic distemper – old ──────────────────────────────────────
  ("13.108.1","13.108 Distempering with acrylic distemper (ready made) VOC – old",
   "One or more coats on old work","sqm",10,[
    ("0835","Acrylic distemper","litre",0.75,90.00),
    ("9977","Carriage of material","L.S.",0.91,2.00),
    ("0131","Painter","day",0.40,679.00),
    ("0115","Coolie","day",0.20,558.00),
    ("9999","Brushes etc.","L.S.",4.42,2.00)],95.75),
  # ── 13.109  Water proofing cement paint – old ────────────────────────────
  ("13.109.1","13.109 Finishing walls with water proofing cement paint – old",
   "Old work (one or more coats @ 2.20 kg/10 sqm) over priming coat","sqm",10,[
    ("9977","Carriage of material","L.S.",1.56,2.00),
    ("0131","Painter","day",0.54,679.00),
    ("0115","Coolie","day",0.27,558.00),
    ("9999","Brushes, sand paper etc.","L.S.",5.33,2.00)],80.80),
  ("13.109.2","13.109 Finishing walls with water proofing cement paint – old",
   "Old work (one or more coats @ 2.20 kg/10 sqm) complete","sqm",10,[
    ("9977","Carriage of material","L.S.",1.56,2.00),
    ("0131","Painter","day",0.54,679.00),
    ("0115","Coolie","day",0.27,558.00),
    ("9999","Brushes, sand paper etc.","L.S.",5.33,2.00)],80.80),
  # ── 13.110  Textured exterior paint – old ────────────────────────────────
  ("13.110.1","13.110 Finishing walls with textured exterior paint – old",
   "Old work (two or more coats on existing cement paint @ 3.28 ltr/10 sqm)","sqm",10,[
    ("9977","Carriage of material","L.S.",2.86,2.00),
    ("0131","Painter","day",0.80,679.00),
    ("0115","Coolie","day",0.40,558.00),
    ("9999","Brushes, sand paper etc.","L.S.",5.33,2.00)],117.75),
  ("13.110.2","13.110 Finishing walls with textured exterior paint – old",
   "Old work (one or more coats) @ 1.82 ltr/10 sqm","sqm",10,[
    ("9977","Carriage of material","L.S.",1.56,2.00),
    ("0131","Painter","day",0.54,679.00),
    ("0115","Coolie","day",0.27,558.00),
    ("9999","Brushes, sand paper etc.","L.S.",4.42,2.00)],80.80),
  # ── 13.111  Acrylic smooth exterior paint – old ───────────────────────────
  ("13.111.1","13.111 Finishing walls with Acrylic Smooth exterior paint – old",
   "Old work (two or more coats @ 1.67 ltr/10 sqm) on existing cement paint","sqm",10,[
    ("9977","Carriage of material","L.S.",2.86,2.00),
    ("0131","Painter","day",0.80,679.00),
    ("0115","Coolie","day",0.40,558.00),
    ("9999","Brushes, sand paper etc.","L.S.",5.33,2.00)],117.75),
  ("13.111.2","13.111 Finishing walls with Acrylic Smooth exterior paint – old",
   "Old work (one or more coats @ 0.90 ltr/10 sqm)","sqm",10,[
    ("9977","Carriage of material","L.S.",1.56,2.00),
    ("0131","Painter","day",0.54,679.00),
    ("0115","Coolie","day",0.27,558.00),
    ("9999","Brushes etc.","L.S.",4.42,2.00)],80.80),
  # ── 13.112  Premium acrylic smooth exterior paint – old ──────────────────
  ("13.112.1","13.112 Finishing walls with Premium Acrylic Smooth Exterior Paint – old",
   "Old work (two or more coats @ 1.43 ltr/10 sqm) over existing cement paint","sqm",10,[
    ("9977","Carriage of material","L.S.",2.86,2.00),
    ("0131","Painter","day",0.80,679.00),
    ("0115","Coolie","day",0.40,558.00),
    ("9999","Brushes, sand paper etc.","L.S.",5.33,2.00)],117.75),
  ("13.112.2","13.112 Finishing walls with Premium Acrylic Smooth Exterior Paint – old",
   "Old work (one or more coats @ 0.83 ltr/10 sqm)","sqm",10,[
    ("9977","Carriage of material","L.S.",1.56,2.00),
    ("0131","Painter","day",0.54,679.00),
    ("0115","Coolie","day",0.27,558.00),
    ("9999","Brushes etc.","L.S.",4.42,2.00)],80.80),
  # ── 13.113  Varnishing – old ─────────────────────────────────────────────
  ("13.113.1","13.113 Varnishing – old",
   "One or more coats with copal varnish","sqm",10,[
    ("9977","Carriage of varnish","L.S.",2.86,2.00),
    ("0131","Painter","day",0.80,679.00),
    ("0115","Coolie","day",0.40,558.00),
    ("9999","Brushes, sand paper etc.","L.S.",5.33,2.00)],117.75),
  ("13.113.2","13.113 Varnishing – old",
   "One or more coats with spar varnish","sqm",10,[
    ("9977","Carriage of varnish","L.S.",2.86,2.00),
    ("0131","Painter","day",0.80,679.00),
    ("0115","Coolie","day",0.40,558.00),
    ("9999","Brushes, sand paper etc.","L.S.",5.33,2.00)],117.75),
  # ── 13.114 ───────────────────────────────────────────────────────────────
  ("13.114","13.114 Melamine polishing on wood work (one or more coats)",
   "Melamine polishing on wood work (one or more coat)","sqm",10,[
    ("9977","Carriage of melamine polish","L.S.",4.42,2.00),
    ("0131","Painter","day",1.60,679.00),
    ("0115","Coolie","day",0.80,558.00),
    ("9999","Brushes, sand paper etc.","L.S.",8.06,2.00)],221.70),
  # ── 13.115 ───────────────────────────────────────────────────────────────
  ("13.115","13.115 Varnishing with flatting varnish one or more coat",
   "Varnishing with flatting varnish of approved brand and manufacture one or more coat","sqm",10,[
    ("9977","Carriage of flatting varnish","L.S.",2.86,2.00),
    ("0131","Painter","day",0.80,679.00),
    ("0115","Coolie","day",0.40,558.00),
    ("9999","Brushes, sand paper etc.","L.S.",5.33,2.00)],117.75),
]


# ── Sheet builder ─────────────────────────────────────────────────────────────
def _cell(ws, r, c, val, fnt=None, aln=None, fl=None, brd=None, num=None):
    cell = ws.cell(row=r, column=c, value=val)
    if fnt: cell.font = fnt
    if aln: cell.alignment = aln
    if fl:  cell.fill = fl
    if brd: cell.border = brd
    if num: cell.number_format = num
    return cell


def build_finishing_sheet(wb):
    SN = "13_Finishing"
    if SN in wb.sheetnames:
        del wb[SN]
    ws = wb.create_sheet(SN)

    # Column widths
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 52
    ws.column_dimensions["D"].width = 9
    ws.column_dimensions["E"].width = 12
    ws.column_dimensions["F"].width = 13
    ws.column_dimensions["G"].width = 14
    ws.column_dimensions["H"].width = 40

    row = 1
    # ── Sheet title ──────────────────────────────────────────────────────────
    ws.merge_cells(f"A{row}:H{row}")
    _cell(ws, row, 1,
          "CPWD DAR 2019 – SUB-HEAD 13: FINISHING WORKS  |  Resource & Rate Analysis  (165 items)",
          Font(name="Calibri", size=13, bold=True, color=C_WHITE),
          AL_C, fill(C_TITLE_BG))
    ws.row_dimensions[row].height = 22
    row += 1

    ws.merge_cells(f"A{row}:H{row}")
    _cell(ws, row, 1,
          "PDF source: CivilDAR_2019_Vol_2.pdf  |  Chapters 895–994  |  "
          "SAY rates reproduced verbatim from the book; markup chain W→X→Y→Z per CPWD DAR convention.",
          Font(name="Calibri", size=9, italic=True, color="444444"),
          AL_L, fill("F0F0F0"))
    row += 1

    items = load_ch13_items()
    last_parent = None

    for it in items:
        code = it["code"]
        desc = it.get("desc", "")
        unit = it["unit"]
        basis = float(it["basis"])
        say = float(it["say"])
        resources = [(r["code"], r["desc"], r["unit"], float(r["qty"]), float(r["rate"]))
                     for r in it.get("resources", [])]
        # Derive parent label from code (e.g. 13.4.1 -> "13.4 ...")
        parts = code.split(".")
        parent = ".".join(parts[:2]) if len(parts) > 2 else code

        # ── Group header when parent changes ────────────────────────────────
        if parent != last_parent:
            last_parent = parent
            row += 1
            ws.merge_cells(f"A{row}:H{row}")
            _cell(ws, row, 1, f"Item group {parent}",
                  Font(name="Calibri", size=10, bold=True, color=C_WHITE),
                  AL_L, fill(C_PARENT_BG), tb())
            ws.row_dimensions[row].height = 18
            row += 1

        # ── Item header ──────────────────────────────────────────────────────
        ws.merge_cells(f"A{row}:H{row}")
        _cell(ws, row, 1, f"{code}  {desc}  [{basis:.0f} {unit} basis]",
              Font(name="Calibri", size=9, bold=True, color=C_WHITE),
              AL_L, fill(C_ITEM_BG), tb())
        ws.row_dimensions[row].height = 16
        row += 1

        # ── Resource table header ────────────────────────────────────────────
        hdrs = ["#", "Code", "Description / Specification", "Unit",
                "Qty / Coeff", "Basic Rate (Rs)", "Amount (Rs)", "Remarks"]
        for ci, h in enumerate(hdrs, 1):
            _cell(ws, row, ci, h,
                  Font(name="Calibri", size=8, bold=True, color=C_WHITE),
                  AL_C, fill(C_TH_BG), tb())
        row += 1

        # ── Resource rows ────────────────────────────────────────────────────
        for ri, (rcode, rdesc, runit, rqty, rrate) in enumerate(resources, 1):
            bg = C_ALT_ROW if ri % 2 == 0 else None
            amt = round(rqty * rrate, 2)
            vals = [ri, rcode, rdesc, runit, rqty, rrate, amt, ""]
            for ci, v in enumerate(vals, 1):
                num = "#,##0.00" if ci in (5, 6, 7) else None
                _cell(ws, row, ci, v,
                      Font(name="Calibri", size=8),
                      AL_R if ci in (5, 6, 7) else AL_L,
                      fill(bg) if bg else PatternFill(fill_type=None),
                      tb(), num)
            row += 1

        # ── Markup summary ───────────────────────────────────────────────────
        W = round(sum(rqty * rrate for _, _, _, rqty, rrate in resources), 2)
        water = round(W * 0.01, 2)
        X = round(W + water, 2)
        gst = round(X * 0.1405, 2)
        Y = round(X + gst, 2)
        cpoh = round(Y * 0.15, 2)
        Z = round(Y + cpoh, 2)
        cess = round(Z * 0.01, 2)
        cost = round(Z + cess, 2)
        rate_calc = round(cost / basis, 4)

        markup_rows = [
            ("W",  "Direct cost (materials + labour + sundries)", "", W),
            ("X1", "+1% Water charges", "", water),
            ("X",  "Subtotal X", "", X),
            ("Y1", "+GST @14.05% on X", "", gst),
            ("Y",  "Subtotal Y", "", Y),
            ("Z1", "+15% CPOH on Y", "", cpoh),
            ("Z",  "Subtotal Z", "", Z),
            ("Z2", "+1% Labour Welfare Cess on Z", "", cess),
            ("",   f"Cost of {basis:.0f} {unit}", "", cost),
            ("",   f"Rate per 1 {unit}", "", round(rate_calc, 2)),
        ]
        for step, lbl, _, amt_val in markup_rows:
            _cell(ws, row, 1, step, Font(name="Calibri", size=8, bold=True), AL_C, fill("EBF3FB"), tb())
            ws.merge_cells(f"B{row}:F{row}")
            _cell(ws, row, 2, lbl, Font(name="Calibri", size=8), AL_L, fill("EBF3FB"), tb())
            ws.cell(row=row, column=7, value=amt_val).number_format = "#,##0.00"
            ws.cell(row=row, column=7).font = Font(name="Calibri", size=8)
            ws.cell(row=row, column=7).alignment = AL_R
            ws.cell(row=row, column=7).fill = fill("EBF3FB")
            ws.cell(row=row, column=7).border = tb()
            row += 1

        # ── SAY rate row ──────────────────────────────────────────────────────
        ws.merge_cells(f"A{row}:F{row}")
        _cell(ws, row, 1,
              f"SAY (MROUND to Rs 0.05)  —  PDF published rate: Rs {say:.2f} / {unit}",
              Font(name="Calibri", size=9, bold=True, color="1F4E79"),
              AL_C, fill(C_SAY_BG), tb())
        _cell(ws, row, 7, say,
              Font(name="Calibri", size=9, bold=True, color="1F4E79"),
              AL_R, fill(C_SAY_BG), tb(), "#,##0.00")
        match = "✓ MATCH" if abs(rate_calc - say) < 2.0 else f"⚠ DIFF {rate_calc:.2f} vs {say:.2f}"
        _cell(ws, row, 8, match,
              Font(name="Calibri", size=8, bold=True,
                   color="375623" if "MATCH" in match else "9C0006"),
              AL_C, fill(C_SAY_BG), tb())
        row += 2

    return len(items)


# ── Entry point ───────────────────────────────────────────────────────────────
def main():
    wb_path = Path(VOL2_WORKBOOK)
    if not wb_path.exists():
        print(f"ERROR: workbook not found at {wb_path}")
        sys.exit(1)

    print(f"Opening {wb_path.name} …")
    wb = openpyxl.load_workbook(str(wb_path))

    n = build_finishing_sheet(wb)
    print(f"13_Finishing sheet rebuilt ({n} items).")

    wb.save(str(wb_path))
    print(f"Saved -> {wb_path}")


if __name__ == "__main__":
    main()
