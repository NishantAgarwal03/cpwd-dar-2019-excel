"""
support_builder_steel.py
Rebuilds the existing '10_Steel_Work' worksheet with First-Principles Resource,
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

STEEL_ITEMS = [
    # ─── 10.1 STRUCTURAL STEEL WORK IN SINGLE SECTION ───────────────────────────
    {
        "id": "10.1",
        "parent_title": "10.1 STRUCTURAL STEEL WORK IN SINGLE SECTION",
        "title": "10.1 Structural steel work in single section, fixed with or without connecting plate, including cutting, hoisting & primer (1 quintal basis)",
        "unit": "quintal", "base_qty": 1.0,
        "rows": [
            ("Structural steel sections (RSJ, channels, angles, tees)", "Mild steel rolled sections conforming to IS:2062 Grade E250 (Fe 410 W) cut to length with mechanical saw / gas torch (1.05 quintal incl. 5% cutting wastage)", "Standard hot-rolled structural sections; beams, columns, lintels and purlins", "Material", "1.050 quintal (1.050 q/q)", "1 quintal"),
            ("Carriage of Steel", "Transport, loading, unloading and staging of steel sections at site", "Carriage allowance", "Material", "0.105 tonne", "1 quintal"),
            ("Fitter (grade 1)", "Laying out cutting marks, straightening, squaring ends, drilling bolt/rivet holes, assembling and aligning true to plumb/line", "Skilled structural fitter; 0.50 day = 4.00 man-hrs per quintal", "Labour", "4.00 man-hrs (0.250 q/man-hr)", "1 quintal"),
            ("Blacksmith 2nd class", "Forging cleat plates, bending, flame cutting, bevelling edges, and cold chisel dressing", "Skilled metal craftsman; 0.75 day = 6.00 man-hrs per quintal", "Labour", "6.00 man-hrs (0.167 q/man-hr)", "1 quintal"),
            ("Beldar", "Rigging, manual hoisting with chain pulley blocks / derricks, holding sections during bolting/welding, and de-rigging", "Handling & lifting crew; 1.00 day = 8.00 man-hrs per quintal", "Labour", "8.00 man-hrs (0.125 q/man-hr)", "1 quintal"),
            ("Sundries", "Gas cutting oxygen/acetylene gases, hacksaw blades, drill bits, temporary bolts, and staging ropes", "Sundries allowance (L.S. 20.67 x cost index)", "Equipment", "L.S. allowance", "1 quintal"),
            ("REF#13.50.3 (Steel primer)", "Shop coat of approved red oxide zinc chromate steel primer (IS:2074) (3.00 sqm per quintal steel)", "Finishing scope executed per CPWD REF#13.50.3 specification", "Reference", "3.00 sqm (3.00 sqm/q)", "1 quintal"),
        ]
    },

    # ─── 10.2 STRUCTURAL STEEL IN BUILT-UP TRUSSES & FRAMED WORK ────────────────
    {
        "id": "10.2",
        "parent_title": "10.2 STRUCTURAL STEEL IN BUILT-UP TRUSSES & FRAMED WORK",
        "title": "10.2 Structural steel work riveted, bolted or welded in built up sections, trusses and framed work (1 quintal basis)",
        "unit": "quintal", "base_qty": 1.0,
        "rows": [
            ("Structural steel sections & plates (tees, angles, flats)", "Fabrication of principal rafters, struts, ties, gusset plates (6-12mm) and purlin cleats (1.05 quintal net incl. 5% cutting/trimming wastage)", "Conforming to IS:2062 Grade E250; complex built-up lattice trusses, girders, and portal frames", "Material", "1.050 quintal (1.050 q/q)", "1 quintal"),
            ("Carriage of Steel", "Haulage and mechanical crane/winch staging of fabricated trusses to roof level", "Carriage allowance", "Material", "0.105 tonne", "1 quintal"),
            ("Connecting bolts, nuts & welding electrodes", "High-strength friction grip (HSFG) or Grade 4.6 bolts with hardened washers, and E6013 welding electrodes", "Connecting hardware for truss node joints", "Material", "3.50 kg (3.50 kg/q)", "1 quintal"),
            ("Fitter (grade 1)", "Setting out truss geometry on template floor, jig assembly, matching gusset holes, plumbing and alignment of trusses on bearing plates", "Master structural fabricator; 0.60 day = 4.80 man-hrs per quintal", "Labour", "4.80 man-hrs (0.208 q/man-hr)", "1 quintal"),
            ("Blacksmith 2nd class", "Cutting gussets, edge planing, pre-heating, tack welding, full-fillet welding of truss joints, and chipping slag", "Welder/blacksmith crew; 0.90 day = 7.20 man-hrs per quintal", "Labour", "7.20 man-hrs (0.139 q/man-hr)", "1 quintal"),
            ("Beldar", "Heavy rigging, gin-pole/crane hoisting of assembled trusses across spans up to 15m, fixing temporary guy ropes and bracings", "Rigging crew; 1.20 day = 9.60 man-hrs per quintal", "Labour", "9.60 man-hrs (0.104 q/man-hr)", "1 quintal"),
            ("Sundries", "Welding flux, grinding wheels, crane slings, guy wires, turnbuckles, and scaffolding timbers", "Sundries allowance", "Equipment", "L.S. allowance", "1 quintal"),
            ("REF#13.50.3 (Steel primer)", "Shop coat of approved red oxide zinc chromate steel primer (3.00 sqm per quintal steel)", "Finishing scope executed per CPWD REF#13.50.3 specification", "Reference", "3.00 sqm (3.00 sqm/q)", "1 quintal"),
        ]
    },

    # ─── 10.3 COLLAPSIBLE STEEL SHUTTERS ────────────────────────────────────────
    {
        "id": "10.3",
        "parent_title": "10.3 COLLAPSIBLE STEEL SHUTTERS",
        "title": "10.3 Providing & fixing collapsible steel shutters with vertical channels 20x10x2 mm & 20x5 mm diagonals (10 sqm basis)",
        "unit": "10 sqm", "base_qty": 10.0,
        "rows": [
            ("M.S. channels 20x10x2mm", "Double vertical channels (20x10x2mm) spaced at 100-120mm centres forming shutter gate leaves (1.50 quintal incl. 10% wastage)", "Cold-formed mild steel channels conforming to IS:2062", "Material", "1.500 quintal (0.150 q/sqm)", "10 sqm"),
            ("M.S. Tees 40x40x6mm (top & bottom tracks)", "Top guide rail and bottom inverted roller track tee sections (0.35 quintal)", "Structural tees for smooth sliding guidance", "Material", "0.350 quintal (0.035 q/sqm)", "10 sqm"),
            ("M.S. flat iron diagonals 20x5mm", "Lattice brace diagonals (20x5mm) connecting vertical channels with 6mm steel pins (0.80 quintal)", "Pivoting scissor mechanism ensuring uniform collapsible movement", "Material", "0.800 quintal (0.080 q/sqm)", "10 sqm"),
            ("Steel pulleys & ball bearings 40mm", "40mm diameter hardened steel roller pulleys fitted with sealed ball bearings (20 nos.)", "Roller mechanism fixed at alternate vertical channel nodes", "Material", "20.00 nos (2.00 nos/sqm)", "10 sqm"),
            ("Locking arrangement, handles & stoppers", "M.S. locking brackets, padlock eyes, cast steel handles, rubber stops and holddown anchor pins", "Hardware & locking fittings", "Material", "L.S. set", "10 sqm"),
            ("Fitter (grade 1)", "Aligning top and bottom tracks true to level, inserting channels, assembling scissor diagonals, adjusting roller clearance and swing", "Skilled fitter; 2.00 days = 16.00 man-hrs per 10 sqm", "Labour", "16.00 man-hrs (0.625 sqm/man-hr)", "10 sqm"),
            ("Blacksmith 2nd class", "Drilling pivot holes, riveting/pinning diagonals, welding stop plates and track anchors", "Metal craftsman; 2.50 days = 20.00 man-hrs per 10 sqm", "Labour", "20.00 man-hrs (0.500 sqm/man-hr)", "10 sqm"),
            ("Beldar", "Lifting, holding shutter assemblies in gate openings, grouting anchor holdfasts, and cleaning tracks", "Labour crew; 3.00 days = 24.00 man-hrs per 10 sqm", "Labour", "24.00 man-hrs (0.417 sqm/man-hr)", "10 sqm"),
            ("Sundries", "Grease, rivets, anchor bolts, drill bits, and touch-up red oxide", "Sundries allowance", "Equipment", "L.S. allowance", "10 sqm"),
            ("REF#13.50.3 (Steel primer)", "Shop coat of red oxide primer applied to all sides of channels and diagonals before assembly", "Finishing scope executed per CPWD REF#13.50.3 specification", "Reference", "25.00 sqm (2.50 sqm/sqm)", "10 sqm"),
        ]
    },

    # ─── 10.5 M.S. SHEET DOORS ──────────────────────────────────────────────────
    {
        "id": "10.5",
        "parent_title": "10.5 M.S. SHEET DOORS",
        "title": "10.5 Providing and fixing 1mm thick M.S. sheet door with frame of 40x40x6 mm angle iron & 3mm gussets (10 sqm basis)",
        "unit": "10 sqm", "base_qty": 10.0,
        "rows": [
            ("M.S. sheet 1mm thick", "Cold-rolled close annealed (CRCA) mild steel sheet 1.0mm thick (7.85 kg/sqm + 5% cutting wastage = 8.24 kg/sqm)", "Conforming to IS:513; welded or riveted to angle iron frame", "Material", "0.824 quintal (0.082 q/sqm)", "10 sqm"),
            ("M.S. angle iron 40x40x6mm", "Perimeter frame and horizontal/vertical cross stiffeners of 40x40x6mm angle iron (3.5 kg/m) (1.20 quintal)", "Structural frame conforming to IS:2062", "Material", "1.200 quintal (0.120 q/sqm)", "10 sqm"),
            ("M.S. gusset plates 3mm & holdfasts", "Corner gussets (3mm thick) at all junctions, diagonal braces, and 40x5mm flat iron hold fasts", "Bracing & wall anchoring hardware", "Material", "0.250 quintal (0.025 q/sqm)", "10 sqm"),
            ("Heavy duty M.S. butt hinges & fittings", "Forged 100mm heavy-duty hinges welded to frame and leaf, sliding door bolt, tower bolts and handles", "Door hardware", "Material", "L.S. set", "10 sqm"),
            ("Carriage of Steel & sheets", "Haulage of sheet panels and angle iron to welding shop and site", "Carriage allowance", "Material", "0.230 tonne", "10 sqm"),
            ("Blacksmith 1st class", "Cutting angle iron, mitring corners, stitching 1mm sheet to angles with continuous/tack welds, dressing welds smooth", "Skilled blacksmith; 3.50 days = 28.00 man-hrs per 10 sqm", "Labour", "28.00 man-hrs (0.357 sqm/man-hr)", "10 sqm"),
            ("Beldar", "Handling sheets, holding during welding, lifting finished doors into reveals, temporary propping during masonry embedding", "Labour crew; 3.50 days = 28.00 man-hrs per 10 sqm", "Labour", "28.00 man-hrs (0.357 sqm/man-hr)", "10 sqm"),
            ("Sundries", "Welding electrodes, grinding discs, cutting gas, and wire brushes", "Sundries allowance", "Equipment", "L.S. allowance", "10 sqm"),
            ("REF#13.50.3 (Steel primer)", "Complete priming coat of approved zinc chromate primer on both sides of leaf and frame (22 sqm)", "Finishing scope executed per CPWD REF#13.50.3 specification", "Reference", "22.00 sqm (2.20 sqm/sqm)", "10 sqm"),
        ]
    },

    # ─── 10.6 ROLLING SHUTTERS ──────────────────────────────────────────────────
    {
        "id": "10.6",
        "parent_title": "10.6 ROLLING SHUTTERS",
        "title": "10.6 Supplying and fixing rolling shutters with interlocking M.S. laths, pipe shaft, wire springs & top cover (10 sqm basis)",
        "unit": "10 sqm", "base_qty": 10.0,
        "rows": [
            ("Factory made rolling shutter assembly", "Interlocking 18/20 gauge cold-rolled M.S. laths with end locks, heavy-duty side guides, bottom lock rail with sliding shoot bolts (10 sqm)", "Complete shutter curtain conforming to IS:6248; balanced operation", "Material", "10.00 sqm (1.000 sqm/sqm)", "10 sqm"),
            ("M.S. pipe shaft, brackets & ball bearings", "Seamless steel pipe shaft of adequate dia with heavy steel end bearing brackets and cast iron hubs", "Suspension shaft assembly", "Material", "L.S. set", "10 sqm"),
            ("High tensile steel wire helical springs", "27.5 cm long tempered helical coiled wire springs conforming to IS:4454 (Part 1) counter-balancing shutter weight", "Spring counterbalance mechanism", "Material", "L.S. set", "10 sqm"),
            ("M.S. top hood cover", "0.90mm thick CRCA sheet hood cover enclosing coiled curtain and barrel shaft against weather", "Protective enclosure conforming to CPWD specification", "Material", "1.00 set", "10 sqm"),
            ("Carriage of rolling shutters", "Transport of heavy barrel shaft, curtain laths and side channels to site", "Carriage allowance", "Material", "10.00 sqm", "10 sqm"),
            ("Fitter (grade 1)", "Anchoring side guide channels to jambs with dash fasteners, hoisting barrel shaft, winding counterbalance springs, feeding laths, testing tension", "Master shutter technician; 1.50 days = 12.00 man-hrs per 10 sqm", "Labour", "12.00 man-hrs (0.833 sqm/man-hr)", "10 sqm"),
            ("Beldar", "Rigging barrel shaft, hoisting heavy curtain roll, assisting during spring tensioning, and grouting holdfasts", "Handling crew; 1.50 days = 12.00 man-hrs per 10 sqm", "Labour", "12.00 man-hrs (0.833 sqm/man-hr)", "10 sqm"),
            ("Sundries", "Dash fasteners 12x100mm, machine bolts, grease, lubricating oil, and alignment levels", "Sundries allowance", "Equipment", "L.S. allowance", "10 sqm"),
            ("REF#13.50.3 (Steel primer)", "Priming coat of approved steel primer on hood cover, side guides and lath faces", "Finishing scope executed per CPWD REF#13.50.3 specification", "Reference", "24.00 sqm (2.40 sqm/sqm)", "10 sqm"),
        ]
    },

    # ─── 10.14 PRESSED STEEL DOOR FRAMES ────────────────────────────────────────
    {
        "id": "10.14.1",
        "parent_title": "10.14 PRESSED STEEL DOOR FRAMES",
        "title": "10.14.1 Providing & fixing pressed steel door frames conforming to IS:4351 (Profile B) with concrete filling (10 m run basis)",
        "unit": "10 m", "base_qty": 10.0,
        "rows": [
            ("Pressed steel door frame profile B", "Factory roll-formed 1.25mm / 1.6mm commercial mild steel sheet profile with rebates, hinge jamb mortises, and strike plate cutouts (10 m run)", "Conforming to IS:4351 Profile B; factory phosphated and dip-primed", "Material", "10.00 metre (1.000 m/m)", "10 m"),
            ("Adjustable base ties & hold fast lugs", "Base tie angle (25x25x3mm) to preserve width at floor level, plus 6 nos. adjustable pressed steel corrugated holdfast lugs", "Fixing hardware conforming to IS:4351", "Material", "6.00 nos (0.60 no/m)", "10 m"),
            ("Portland Cement", "Hydraulic binder decomposed from 0.035 cum cavity filling concrete 1:3:6 (REF#4.1.8) (0.035 x 0.220 t = 0.008 t)", "OPC-43 Grade cement; solid core backing prevent hollow sound & denting", "Material", "0.008 tonne (0.0008 t/m)", "10 m"),
            ("Coarse sand", "Zone III coarse sand decomposed from 0.035 cum cavity concrete 1:3:6 (REF#4.1.8) (0.035 x 0.46 = 0.016 cum)", "Clean natural river sand", "Material", "0.016 cum (0.0016 cum/m)", "10 m"),
            ("Graded stone aggregate 10/20mm", "Graded coarse aggregate decomposed from 0.035 cum concrete 1:3:6 (REF#4.1.8) (0.035 x 0.92 = 0.032 cum)", "Stone ballast", "Material", "0.032 cum (0.0032 cum/m)", "10 m"),
            ("Carriage of frames & concrete materials", "Haulage to installation location", "Carriage allowance", "Material", "0.10 tonne equiv", "10 m"),
            ("Fitter (grade 1)", "Plumbing frame, squaring diagonals, checking rebate alignment, anchoring base ties, and securing holdfasts in masonry joints", "Skilled fitter; 0.40 day = 3.20 man-hrs per 10 m", "Labour", "3.20 man-hrs (3.125 m/man-hr)", "10 m"),
            ("Beldar", "Filling frame hollow rear cavity solidly with 1:3:6 concrete in 30cm lifts, compacting with rod, and cleaning steel face", "Labour crew; 0.40 day = 3.20 man-hrs per 10 m", "Labour", "3.20 man-hrs (3.125 m/man-hr)", "10 m"),
            ("Bhisti", "Water curing filled concrete core inside frames", "Watering crew; 0.10 day = 0.80 man-hrs per 10 m", "Labour", "0.80 man-hrs (12.50 m/man-hr)", "10 m"),
            ("REF#4.1.8 (Cement concrete 1:3:6)", "Referenced core filling: 0.035 cum CC 1:3:6 per 10 m pressed steel frame", "Concrete scope executed per CPWD REF#4.1.8 specification", "Reference", "0.035 cum (0.0035 cum/m)", "10 m"),
        ]
    },

    # ─── 10.16 TUBULAR STEEL WORK ───────────────────────────────────────────────
    {
        "id": "10.16.1",
        "parent_title": "10.16 STEEL WORK IN BUILT-UP TUBULAR SECTIONS",
        "title": "10.16.1 Steel work in built up tubular trusses & purlins using hot finished welded type tubes IS:1161 (1 quintal basis)",
        "unit": "quintal", "base_qty": 1.0,
        "rows": [
            ("Mild steel hollow circular / rectangular tubes", "Hot finished welded steel tubes conforming to IS:1161 / IS:4923 Grade YSt 210/240 (1.05 quintal net incl. 5% cutting wastage)", "Structural hollow sections offering high torsional and buckling resistance per unit weight", "Material", "1.050 quintal (1.050 q/q)", "1 quintal"),
            ("Carriage of Steel tubes", "Transport and staging of tubular members to fabrication area", "Carriage allowance", "Material", "0.105 tonne", "1 quintal"),
            ("M.S. plates, end caps & gussets", "6mm to 10mm thick mild steel gusset plates, flattened tube ends, and welded capping discs to seal hollow voids", "Conforming to IS:2062 Grade E250", "Material", "0.120 quintal (0.120 q/q)", "1 quintal"),
            ("Fitter (grade 1)", "Profiling tube ends with fish-mouth / saddle cuts, setting up geometry on jig, squaring, drilling hole connections, and plumbing trusses", "Skilled structural fitter; 0.55 day = 4.40 man-hrs per quintal", "Labour", "4.40 man-hrs (0.227 q/man-hr)", "1 quintal"),
            ("Blacksmith / Welder", "Continuous all-round fillet welding of tubular nodes, sealing ends against internal atmospheric corrosion, and weld dressing", "Certified welder; 0.80 day = 6.40 man-hrs per quintal", "Labour", "6.40 man-hrs (0.156 q/man-hr)", "1 quintal"),
            ("Beldar", "Handling tubular trusses, rigging, hoisting onto RCC / steel column caps, temporary propping with purlins", "Rigging crew; 1.10 day = 8.80 man-hrs per quintal", "Labour", "8.80 man-hrs (0.114 q/man-hr)", "1 quintal"),
            ("Sundries", "Welding rods, gas cutting oxygen/DA, grinding wheels, and hoisting slings", "Sundries allowance", "Equipment", "L.S. allowance", "1 quintal"),
            ("REF#13.50.3 (Steel primer)", "Shop coat of approved epoxy/zinc chromate steel primer (3.20 sqm per quintal tubular steel)", "Finishing scope executed per CPWD REF#13.50.3 specification", "Reference", "3.20 sqm (3.20 sqm/q)", "1 quintal"),
        ]
    },

    # ─── 10.22 ELECTRIC ARC / GAS WELDING ───────────────────────────────────────
    {
        "id": "10.22",
        "parent_title": "10.22 WELDING BY GAS OR ELECTRIC PLANT",
        "title": "10.22 Welding by gas or electric plant including transportation of plant at site (100 cm run basis)",
        "unit": "100 cm", "base_qty": 100.0,
        "rows": [
            ("Electric arc / gas welding operation", "Welding 100 cm run of structural joints using heavy-duty portable transformer/generator plant", "Conforming to IS:816 code of practice for use of metal arc welding", "Machinery", "100.00 cm run (1.000 cm/cm)", "100 cm"),
            ("Welding electrodes (E6013 / heavy coated)", "Heavy-coated shielded metal arc welding electrodes 3.15mm / 4.0mm dia conforming to IS:814", "Consumable electrode", "Material", "1.20 kg (0.012 kg/cm)", "100 cm"),
            ("Welder (grade 1)", "Pre-cleaning joint, striking arc, laying root and cover weld beads, maintaining uniform penetration, chipping slag, and dye check", "Certified structural welder; 0.15 day = 1.20 man-hrs per 100 cm", "Labour", "1.20 man-hrs (83.3 cm/man-hr)", "100 cm"),
            ("Beldar", "Shifting welding cables, earth clamps, chipping hammer, wire brush, holding components with clamps", "Assisting crew; 0.15 day = 1.20 man-hrs per 100 cm", "Labour", "1.20 man-hrs (83.3 cm/man-hr)", "100 cm"),
            ("Sundries", "Plant shifting allowance, electrical energy / fuel charges, face shield lenses, and chipping chisels", "Sundries allowance (L.S. allowance)", "Equipment", "L.S. allowance", "100 cm"),
        ]
    },

    # ─── 10.25 STEEL WORK WELDED IN BUILT-UP FRAMED WORK (GRILLS, GATES) ────────
    {
        "id": "10.25",
        "parent_title": "10.25 STEEL WORK WELDED IN BUILT-UP SECTIONS / FRAMED WORK",
        "title": "10.25 Steel work welded in built up sections / framed work for window grills, railings & gates (1 quintal basis)",
        "unit": "quintal", "base_qty": 1.0,
        "rows": [
            ("Mild steel flats, square bars & angles", "M.S. flats (20x5mm, 25x6mm), square bars (10mm, 12mm), angles (35x35x5mm) cut to pattern, forged and punched (1.05 quintal incl. 5% wastage)", "Conforming to IS:2062 Grade E250; architectural window grills, balustrades, staircase railings and compound gates", "Material", "1.050 quintal (1.050 q/q)", "1 quintal"),
            ("Carriage of Steel", "Transport and handling of steel bars and fabricated panels", "Carriage allowance", "Material", "0.105 tonne", "1 quintal"),
            ("Blacksmith 1st class", "Hot/cold bending scrolls, forging spear heads, punching bar holes, assembling grid panels, welding intersection points, and grinding smooth", "Master blacksmith; 1.00 day = 8.00 man-hrs per quintal", "Labour", "8.00 man-hrs (0.125 q/man-hr)", "1 quintal"),
            ("Fitter (grade 1)", "Squaring grill panels, drilling perimeter holes, countersinking, positioning holdfasts, plumbing in window reveals, and fixing", "Skilled fitter; 0.70 day = 5.60 man-hrs per quintal", "Labour", "5.60 man-hrs (0.179 q/man-hr)", "1 quintal"),
            ("Beldar", "Handling materials, holding panels during welding, assisting in fixing grills to brickwork/concrete with cement mortar", "Labour crew; 1.30 day = 10.40 man-hrs per quintal", "Labour", "10.40 man-hrs (0.096 q/man-hr)", "1 quintal"),
            ("Sundries", "Welding rods, cutting gas, emery cloth, grinding stones, and staging scaffolding", "Sundries allowance", "Equipment", "L.S. allowance", "1 quintal"),
            ("REF#13.50.3 (Steel primer)", "Complete priming coat of red oxide zinc chromate primer on intricate grill surfaces (3.50 sqm per quintal)", "Finishing scope executed per CPWD REF#13.50.3 specification", "Reference", "3.50 sqm (3.50 sqm/q)", "1 quintal"),
        ]
    },

    # ─── 10.28 STAINLESS STEEL RAILING (GRADE 304) ──────────────────────────────
    {
        "id": "10.28",
        "parent_title": "10.28 STAINLESS STEEL RAILING (GRADE 304)",
        "title": "10.28 Stainless steel (Grade 304) railing with hollow tubes, balusters, dash fasteners & mirror buffing (100 kg basis)",
        "unit": "100 kg", "base_qty": 100.0,
        "rows": [
            ("Stainless steel (Grade 304) hollow tubes & sections", "SS-304 hollow circular tubes (50mm dia top handrail, 38mm dia balusters, 19mm intermediate runners) with 1.5mm wall thickness (105 kg incl. 5% cutting wastage)", "Premium austenitic stainless steel conforming to ASTM A554 Grade 304; corrosion resistant, mirror / hairline satin finish", "Material", "105.00 kg (1.050 kg/kg)", "100 kg"),
            ("SS-304 base plates, cover flanges & bolts", "SS solid base plates (100x100x6mm), decorative cover rosettes, and SS-304 dash fasteners / expansion anchor bolts (12x100mm)", "Non-corrosive structural anchor hardware", "Material", "50.00 kg (0.500 kg/kg)", "100 kg"),
            ("Carriage of SS materials", "Careful transport and handling of polished SS tubes with protective plastic wrapping", "Carriage allowance", "Material", "0.155 tonne", "100 kg"),
            ("Blacksmith / Welder (SS specialist)", "Bending elbows/returns, TIG (Argon arc) welding joints with SS-308 filler wire, grinding weld beads flush with coarse to fine flap discs", "Certified TIG welder & finisher; 3.00 days = 24.00 man-hrs per 100 kg", "Labour", "24.00 man-hrs (4.17 kg/man-hr)", "100 kg"),
            ("Bandhani (Polishing specialist)", "Buffing welded seams and tubes with sisal wheels, cotton mops and green rouge polishing compound to uniform 600-grit mirror finish", "Master polisher; 1.00 day = 8.00 man-hrs per 100 kg", "Labour", "8.00 man-hrs (12.50 kg/man-hr)", "100 kg"),
            ("Beldar", "Drilling anchor holes in staircase marble/granite treads, vacuuming dust, driving anchor dash fasteners, holding balusters plumb", "Labour crew; 10.00 days = 80.00 man-hrs per 100 kg", "Labour", "80.00 man-hrs (1.25 kg/man-hr)", "100 kg"),
            ("Sundries", "Argon shielding gas, tungsten electrodes, SS-308 filler wire, 3M Scotch-Brite buffing mops, polishing rouge, and masking tape", "Sundries allowance", "Equipment", "L.S. allowance", "100 kg"),
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

def build_steel_sheet(ws):
    if ws.views.sheetView:
        ws.views.sheetView[0].showGridLines = True

    # Sheet title spanning A1:G1
    ws.merge_cells("A1:G1")
    t1 = ws.cell(row=1, column=1, value="Sub-Head 10.0 — STEEL WORK  |  First-Principles Resource, Work & Gang Analysis")
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
    for item in STEEL_ITEMS:
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

def rebuild_steel_in_workbook(file_path):
    print(f"Opening {file_path} ...")
    wb = openpyxl.load_workbook(file_path)
    if "10_Steel_Work" not in wb.sheetnames:
        raise ValueError(f"'10_Steel_Work' sheet not found in {file_path}")

    pos = wb.sheetnames.index("10_Steel_Work")
    del wb["10_Steel_Work"]
    ws = wb.create_sheet("10_Steel_Work", pos)
    ws.sheet_properties.tabColor = "2E75B6"
    build_steel_sheet(ws)

    temp_path = file_path.replace(".xlsx", "_TMP_STEEL.xlsx")
    wb.save(temp_path)
    wb.close()
    os.replace(temp_path, file_path)
    print(f"Successfully updated {file_path} -> sheet '10_Steel_Work' ({len(STEEL_ITEMS)} items).")

def main():
    repo_root = Path(__file__).resolve().parents[1]
    main_wb = repo_root / "CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_1_Latest.xlsx"
    output_wb = repo_root / "outputs" / "earthwork-custom-rate-composer" / "CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_1_Latest.xlsx"

    if main_wb.exists():
        rebuild_steel_in_workbook(str(main_wb))
    if output_wb.exists():
        rebuild_steel_in_workbook(str(output_wb))

if __name__ == "__main__":
    main()
