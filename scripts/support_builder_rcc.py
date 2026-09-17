"""
support_builder_rcc.py
Rebuilds the existing '05_RCC_Work' worksheet with First-Principles Resource,
Productivity, Gang and Material Analysis per CPWD DAR 2019 Vol 1 Ch.5:
  Item Code | Labour / Machine / Material | Work done | Condition / When used | Category | Productivity | Quantity
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from pathlib import Path
import os

C_TITLE_BG   = "1A1A2E"
C_SUBTITLE_BG= "E8F4FD"
C_HDR_PARENT = "1F4E79"
C_HDR_SUB    = "2E75B6"
C_TH_BG      = "1F4E79"
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

# ── Chapter 5 RCC Items ────────────────────────────────────────────────────
RCC_ITEMS = [
    # ── 5.1.2 ─────────────────────────────────────────────────────────────
    {
        "id": "5.1.2",
        "parent_title": "5.0 REINFORCED CEMENT CONCRETE — IN FOUNDATIONS & FOOTINGS (UP TO PLINTH LEVEL)",
        "title": "5.1.2 RCC 1:1.5:3 in foundations, footings, bases of columns etc. up to plinth level",
        "unit": "cum",
        "base_qty": 1.0,
        "rows": [
            ("Stone Aggregate 20mm",
             "Single-size 20 mm crushed stone forming the primary coarse aggregate skeleton; provides compressive strength and limits water demand",
             "Main structural aggregate for 1:1.5:3 mix; 0.57 cum per cum of finished RCC; sourced from approved quarry conforming to IS:383",
             "Material", "0.57 cum (0.570 cum/cum)", "1 cum"),
            ("Stone Aggregate 10mm",
             "Single-size 10 mm crushed stone filling the voids between 20 mm particles to improve packing density and workability",
             "Secondary aggregate for grading gap; 0.28 cum per cum; combined total with 20 mm = 0.85 cum",
             "Material", "0.28 cum (0.280 cum/cum)", "1 cum"),
            ("Carriage of Stone Aggregate (below 40 mm)",
             "Mechanical transport, unloading and stacking of combined 20 mm + 10 mm stone aggregate at site batching platform",
             "Combined carriage for both sizes; 0.57 + 0.28 = 0.85 cum per cum of concrete",
             "Material", "0.85 cum (0.850 cum/cum)", "1 cum"),
            ("Coarse Sand (Zone III)",
             "Zone III graded coarse sand providing fine-aggregate mortar matrix between aggregate particles; reduces permeability",
             "CPWD-specified Zone III coarse sand free from clay; 0.425 cum per cum of concrete for 1:1.5:3 proportion",
             "Material", "0.425 cum (0.425 cum/cum)", "1 cum"),
            ("Carriage of Coarse Sand",
             "Mechanical transport, unloading and stacking of coarse sand at site mixing area",
             "Standard site carriage of coarse sand; 0.425 cum per cum",
             "Material", "0.425 cum (0.425 cum/cum)", "1 cum"),
            ("Portland Cement",
             "OPC hydraulic binder forming the paste matrix; 0.2833 cum loose volume equals 0.40 tonne for 1:1.5:3 mix",
             "OPC 43/53 grade; 0.40 tonne per cum of finished RCC; controls mix ratio 1 cement : 1.5 sand : 3 aggregate by volume",
             "Material", "0.40 tonne (0.400 t/cum)", "1 cum"),
            ("Carriage of Cement",
             "Transport of cement bags from site store to batching point",
             "Standard allowance for cement carriage within site; 0.40 tonne",
             "Material", "0.40 tonne (0.400 t/cum)", "1 cum"),
            ("Mason (average)",
             "Skilled concrete mason directing mix consistency, supervising placing and compacting concrete in foundation forms",
             "Controlling batching, placing, surface finishing; 0.17 day per cum for straightforward foundation work",
             "Labour", "0.17 day (5.9 cum/day)", "1 cum"),
            ("Beldar",
             "Semi-skilled labourer measuring aggregates by farma, loading mixer hopper, and distributing wet concrete into forms",
             "Primary manual crew for aggregate handling and concrete distribution; 2.00 day per cum",
             "Labour", "2.00 day (0.5 cum/day)", "1 cum"),
            ("Bhisti",
             "Water carrier supplying measured mixing water and maintaining early-age curing of foundation concrete",
             "Mix water supply, curing water; 0.90 day per cum",
             "Labour", "0.90 day (1.1 cum/day)", "1 cum"),
            ("Concrete Mixer (0.25–0.40 cum with hopper)",
             "Mechanical drum mixer with loading hopper batching aggregates, cement and water into homogeneous concrete",
             "Hired mechanical mixing; 0.07 day per cum; approx 14 cum/day throughput; mandatory for RCC",
             "Machine", "0.07 day (14 cum/day)", "1 cum"),
            ("Vibrator (Needle type 40 mm)",
             "Immersion poker vibrator for internal compaction eliminating voids and honeycombing around reinforcement",
             "Needle vibrator compaction; 0.07 day per cum; essential for dense, durable RCC around bars",
             "Machine", "0.07 day (14 cum/day)", "1 cum"),
            ("Sundries",
             "Measuring boxes, mortar pans, wheelbarrows, water buckets, and site cleanup allowance",
             "L.S. 14.30 per cum as minor tools and consumables",
             "Equipment", "L.S. 14.30/cum", "1 cum"),
        ]
    },
    # ── 5.1.3 ─────────────────────────────────────────────────────────────
    {
        "id": "5.1.3",
        "parent_title": "5.0 REINFORCED CEMENT CONCRETE — IN FOUNDATIONS & FOOTINGS (UP TO PLINTH LEVEL)",
        "title": "5.1.3 RCC 1:2:4 in foundations, footings, bases of columns etc. up to plinth level",
        "unit": "cum",
        "base_qty": 1.0,
        "rows": [
            ("Stone Aggregate 20mm",
             "Single-size 20 mm crushed stone as main coarse aggregate for leaner 1:2:4 mix",
             "Higher aggregate content than 1:1.5:3; 0.67 cum per cum of finished RCC",
             "Material", "0.67 cum (0.670 cum/cum)", "1 cum"),
            ("Stone Aggregate 10mm",
             "Single-size 10 mm crushed stone filling voids between 20 mm particles",
             "Reduced proportion vs 1:1.5:3; 0.22 cum per cum; combined stone = 0.89 cum",
             "Material", "0.22 cum (0.220 cum/cum)", "1 cum"),
            ("Carriage of Stone Aggregate (below 40 mm)",
             "Mechanical transport of combined 20 mm + 10 mm stone aggregate to site batching platform",
             "Total stone carriage; 0.67 + 0.22 = 0.89 cum per cum of concrete",
             "Material", "0.89 cum (0.890 cum/cum)", "1 cum"),
            ("Coarse Sand (Zone III)",
             "Zone III coarse sand providing fine-aggregate matrix; higher proportion in leaner 1:2:4 mix",
             "0.445 cum per cum vs 0.425 in 1:1.5:3; reflects increased sand-cement ratio",
             "Material", "0.445 cum (0.445 cum/cum)", "1 cum"),
            ("Carriage of Coarse Sand",
             "Mechanical transport and stacking of coarse sand at batching platform",
             "Standard site carriage; 0.445 cum per cum",
             "Material", "0.445 cum (0.445 cum/cum)", "1 cum"),
            ("Portland Cement",
             "OPC hydraulic binder; 0.2225 cum loose volume = 0.32 tonne for leaner 1:2:4 mix",
             "OPC 43/53 grade; 0.32 tonne per cum; 20% less cement than 1:1.5:3 (0.40 t), hence lower cost",
             "Material", "0.32 tonne (0.320 t/cum)", "1 cum"),
            ("Carriage of Cement",
             "Transport of cement bags from store to batching point",
             "Standard site carriage; 0.32 tonne per cum",
             "Material", "0.32 tonne (0.320 t/cum)", "1 cum"),
            ("Mason (average)",
             "Skilled concrete mason supervising mix, placement and finishing of foundation concrete",
             "0.17 day per cum — same as 5.1.2; foundation work complexity unchanged",
             "Labour", "0.17 day (5.9 cum/day)", "1 cum"),
            ("Beldar",
             "Semi-skilled labourer measuring and loading aggregates, handling and distributing concrete",
             "2.00 day per cum — unchanged; mix ratio change does not affect placing labour",
             "Labour", "2.00 day (0.5 cum/day)", "1 cum"),
            ("Bhisti",
             "Water carrier for mixing water and curing of fresh concrete",
             "0.90 day per cum — same as 5.1.2",
             "Labour", "0.90 day (1.1 cum/day)", "1 cum"),
            ("Concrete Mixer (0.25–0.40 cum with hopper)",
             "Drum mixer batching 1:2:4 concrete ingredients into homogeneous mix",
             "0.07 day per cum; same machine productivity as 1:1.5:3",
             "Machine", "0.07 day (14 cum/day)", "1 cum"),
            ("Vibrator (Needle type 40 mm)",
             "Poker vibrator for compaction of concrete around reinforcement bars",
             "0.07 day per cum; essential to eliminate honeycombing",
             "Machine", "0.07 day (14 cum/day)", "1 cum"),
            ("Sundries",
             "Measuring boxes, pans, wheelbarrows and minor site tools",
             "L.S. 14.30 per cum",
             "Equipment", "L.S. 14.30/cum", "1 cum"),
        ]
    },
    # ── 5.2.2 ─────────────────────────────────────────────────────────────
    {
        "id": "5.2.2",
        "parent_title": "5.0 REINFORCED CEMENT CONCRETE — IN WALLS (ABOVE PLINTH TO FLOOR V LEVEL)",
        "title": "5.2.2 RCC 1:1.5:3 in walls (any thickness) above plinth to floor five level",
        "unit": "cum",
        "base_qty": 1.0,
        "rows": [
            ("Stone Aggregate 20mm",
             "Single-size 20 mm crushed stone; same aggregate content as 5.1.2 — walls require identical mix design",
             "0.57 cum per cum; mix proportions unchanged from foundation RCC",
             "Material", "0.57 cum (0.570 cum/cum)", "1 cum"),
            ("Stone Aggregate 10mm",
             "Single-size 10 mm crushed stone filling voids; same as 5.1.2",
             "0.28 cum per cum",
             "Material", "0.28 cum (0.280 cum/cum)", "1 cum"),
            ("Carriage of Stone Aggregate (below 40 mm)",
             "Mechanical transport of stone aggregate to batching platform",
             "0.85 cum per cum combined",
             "Material", "0.85 cum (0.850 cum/cum)", "1 cum"),
            ("Coarse Sand (Zone III)",
             "Zone III coarse sand; same as 5.1.2 mix design",
             "0.425 cum per cum",
             "Material", "0.425 cum (0.425 cum/cum)", "1 cum"),
            ("Carriage of Coarse Sand",
             "Transport of coarse sand to batching platform",
             "0.425 cum per cum",
             "Material", "0.425 cum (0.425 cum/cum)", "1 cum"),
            ("Portland Cement",
             "OPC binder; 0.2833 cum = 0.40 tonne for 1:1.5:3 — same mix design as 5.1.2",
             "0.40 tonne per cum; mix ratio drives cement content independent of height",
             "Material", "0.40 tonne (0.400 t/cum)", "1 cum"),
            ("Carriage of Cement",
             "Transport of cement bags to batching point",
             "0.40 tonne per cum",
             "Material", "0.40 tonne (0.400 t/cum)", "1 cum"),
            ("Mason 1st Class",
             "Skilled first-class mason for wall forming, bar-cover maintenance, and careful concrete placement in wall forms",
             "Wall concrete demands more skilled supervision than foundations; 0.10 day per cum (from 9.18 cum batch = 0.92 day total)",
             "Labour", "0.10 day (10 cum/day)", "1 cum"),
            ("Mason 2nd Class",
             "Second-class mason assisting 1st class in placing, compacting and levelling wall concrete",
             "0.10 day per cum — equal to 1st class; wall work requires paired mason team",
             "Labour", "0.10 day (10 cum/day)", "1 cum"),
            ("Beldar",
             "Semi-skilled labourer loading mixer, distributing concrete and compacting in wall forms",
             "1.23 day per cum — higher than foundation (2.00 day) because wall shuttering requires more careful placement; from 11.29 day/9.18 cum",
             "Labour", "1.23 day (0.81 cum/day)", "1 cum"),
            ("Coolie",
             "Unskilled labourer assisting in aggregate handling, concrete carrying in pans/buckets to wall forms",
             "0.82 day per cum — extra handling for elevated wall work; from 7.53 day/9.18 cum",
             "Labour", "0.82 day (1.22 cum/day)", "1 cum"),
            ("Bhisti",
             "Water carrier for mixing water and curing of wall concrete faces",
             "0.90 day per cum — same as foundation; wall area requires equivalent curing water",
             "Labour", "0.90 day (1.1 cum/day)", "1 cum"),
            ("Concrete Mixer (0.25–0.40 cum with hopper)",
             "Drum mixer for batching wall concrete mix",
             "0.07 day per cum — unchanged from 5.1.2",
             "Machine", "0.07 day (14 cum/day)", "1 cum"),
            ("Vibrator (Needle type 40 mm)",
             "Poker vibrator for compaction inside wall forms around reinforcement bars",
             "0.07 day per cum — critical for wall sections where honeycombing cannot be seen",
             "Machine", "0.07 day (14 cum/day)", "1 cum"),
            ("Scaffolding",
             "Temporary working scaffold providing safe access platform for placing and compacting concrete in upper wall sections",
             "L.S. 14.36 per cum — from 131.82/9.18 cum batch; adjustable for wall height",
             "Equipment", "L.S. 14.36/cum", "1 cum"),
            ("Sundries",
             "Mortar pans, buckets, measuring boxes, vibrator hoses and minor consumables",
             "L.S. 14.36 per cum — same rate as scaffolding from 9.18 cum analysis",
             "Equipment", "L.S. 14.36/cum", "1 cum"),
            ("Coolie (extra — lifting to floor V)",
             "Additional unskilled labour for hoisting aggregate, cement and wet concrete from ground to upper floor levels by hoist or manual carry",
             "0.75 × 2.5 floors × 1 cum = 1.88 day per cum; extra over ground-level batching labour",
             "Labour", "1.88 day extra lifting", "1 cum"),
        ]
    },
    # ── 5.3 ───────────────────────────────────────────────────────────────
    {
        "id": "5.3",
        "parent_title": "5.0 REINFORCED CEMENT CONCRETE — BEAMS, SUSPENDED FLOORS, ROOFS, BALCONIES, LINTELS (PLINTH TO FLOOR V)",
        "title": "5.3 RCC 1:1.5:3 in beams, suspended floors/roofs ≤15° slope, balconies, lintels, above plinth to floor five",
        "unit": "cum",
        "base_qty": 1.0,
        "rows": [
            ("Stone Aggregate 20mm",
             "Single-size 20 mm crushed stone; same 1:1.5:3 mix design — beam/slab/roof concrete requires identical proportions",
             "0.57 cum per cum of RCC",
             "Material", "0.57 cum (0.570 cum/cum)", "1 cum"),
            ("Stone Aggregate 10mm",
             "Single-size 10 mm stone filling voids; same as 5.1.2/5.2.2",
             "0.28 cum per cum",
             "Material", "0.28 cum (0.280 cum/cum)", "1 cum"),
            ("Carriage of Stone Aggregate (below 40 mm)",
             "Mechanical transport of stone aggregate to site batching platform",
             "0.85 cum per cum combined",
             "Material", "0.85 cum (0.850 cum/cum)", "1 cum"),
            ("Coarse Sand (Zone III)",
             "Zone III coarse sand; same mix as 5.1.2",
             "0.425 cum per cum",
             "Material", "0.425 cum (0.425 cum/cum)", "1 cum"),
            ("Carriage of Coarse Sand",
             "Transport of coarse sand to batching area",
             "0.425 cum per cum",
             "Material", "0.425 cum (0.425 cum/cum)", "1 cum"),
            ("Portland Cement",
             "OPC binder; 0.40 tonne per cum for 1:1.5:3 mix",
             "Same cement content as foundations; mix design unchanged",
             "Material", "0.40 tonne (0.400 t/cum)", "1 cum"),
            ("Carriage of Cement",
             "Transport of cement bags from store to batching point",
             "0.40 tonne per cum",
             "Material", "0.40 tonne (0.400 t/cum)", "1 cum"),
            ("Mason (average)",
             "Skilled mason for overhead concrete placement in beam/slab/roof forms; more difficult than foundation work requiring precise levelling",
             "0.24 day per cum — significantly higher than 5.1.2 (0.17 day); overhead slab work requires greater mason skill and care",
             "Labour", "0.24 day (4.2 cum/day)", "1 cum"),
            ("Beldar",
             "Semi-skilled labourer loading mixer and distributing concrete into beam/slab forms against gravity",
             "2.75 day per cum — higher than foundation (2.00 day); elevated overhead work and shuttering complexity increases labour",
             "Labour", "2.75 day (0.36 cum/day)", "1 cum"),
            ("Bhisti",
             "Water carrier for mixing water and overhead curing of suspended slabs and beams",
             "0.90 day per cum — same as 5.1.2",
             "Labour", "0.90 day (1.1 cum/day)", "1 cum"),
            ("Concrete Mixer (0.25–0.40 cum with hopper)",
             "Drum mixer for batching beam/slab concrete",
             "0.08 day per cum — slightly higher than 5.1.2 (0.07 day) due to more frequent small pours",
             "Machine", "0.08 day (12.5 cum/day)", "1 cum"),
            ("Vibrator (Needle type 40 mm)",
             "Poker vibrator for thorough compaction in beams and slabs; critical for overhead sections where gravity cannot assist compaction",
             "0.08 day per cum — matches mixer at 0.08 day",
             "Machine", "0.08 day (12.5 cum/day)", "1 cum"),
            ("Sundries",
             "Measuring boxes, pans, buckets, vibrator hoses and consumables",
             "L.S. 14.30 per cum",
             "Equipment", "L.S. 14.30/cum", "1 cum"),
            ("Coolie (extra — lifting to floor V)",
             "Additional unskilled labour for hoisting materials and wet concrete from ground level to upper floors",
             "2.5 floors × 0.75 = 1.88 day per cum extra over ground-level rates; standard CPWD allowance for up to floor V",
             "Labour", "1.88 day extra lifting", "1 cum"),
        ]
    },
    # ── 5.4 ───────────────────────────────────────────────────────────────
    {
        "id": "5.4",
        "parent_title": "5.0 REINFORCED CEMENT CONCRETE — IN KERBS, STEPS (UP TO FLOOR V)",
        "title": "5.4 RCC 1:1.5:3 in kerbs, steps, up to floor five level",
        "unit": "cum",
        "base_qty": 1.0,
        "rows": [
            ("Cement Concrete 1:1.5:3 (ref Item 4.1.2 — Concrete Work)",
             "Complete concrete production: all stone aggregate, coarse sand, cement, batching labour, mixer and vibrator as per Ch.4 Concrete Work item 4.1.2",
             "Reference rate ₹7210.55/cum (marked A); full concrete resources derived from Ch.4; kerbs/steps use same mix design but with extra handling",
             "Material", "1.00 cum @ ₹7210.55 (A)", "1 cum"),
            ("Mason 1st Class",
             "Skilled mason for accurate kerb profile and step nosing formation; requires precise screeding and edging",
             "0.04 day per cum; small-section kerb and step forms demand skilled finishing",
             "Labour", "0.04 day", "1 cum"),
            ("Mason 2nd Class",
             "Second-class mason assisting in kerb shuttering, placing and finishing",
             "0.04 day per cum",
             "Labour", "0.04 day", "1 cum"),
            ("Beldar",
             "Semi-skilled labourer handling concrete distribution into small kerb/step sections",
             "0.10 day per cum extra over base concrete rate (ref 4.1.2 already includes main placing labour)",
             "Labour", "0.10 day extra", "1 cum"),
            ("Bhisti",
             "Water carrier for curing of kerb and step concrete surfaces",
             "0.20 day per cum extra curing water for exposed small sections",
             "Labour", "0.20 day extra", "1 cum"),
            ("Mate",
             "Supervisor/gang leader coordinating kerb-setting and step-forming crew",
             "0.04 day per cum; needed for setting out kerb lines and stair geometry",
             "Labour", "0.04 day", "1 cum"),
            ("Coolie (extra — lifting to floor V)",
             "Additional labour for hoisting materials to upper floors for stair/landing step work",
             "2.5 floors × 0.75 = 1.88 day per cum extra; applied to W-A chain (extra beyond ref A rate)",
             "Labour", "1.88 day extra lifting", "1 cum"),
        ]
    },
    # ── 5.5 ───────────────────────────────────────────────────────────────
    {
        "id": "5.5",
        "parent_title": "5.0 REINFORCED CEMENT CONCRETE — ARCHES, DOMES, VAULTS, SHELLS, FOLDED PLATES (PLINTH TO FLOOR V)",
        "title": "5.5 RCC 1:1.5:3 in arches, archribs, domes, vaults, shells, folded plates, roofs >15° slope, up to floor five",
        "unit": "cum",
        "base_qty": 1.0,
        "rows": [
            ("Stone Aggregate 20mm",
             "Single-size 20 mm crushed stone; same mix design as 5.1.2 — curved surfaces do not change concrete proportions",
             "0.57 cum per cum; from 26.73 cum arch example: 15.236/26.73",
             "Material", "0.57 cum (0.570 cum/cum)", "1 cum"),
            ("Stone Aggregate 10mm",
             "Single-size 10 mm stone; same as other 1:1.5:3 items",
             "0.28 cum per cum; from 7.484/26.73",
             "Material", "0.28 cum (0.280 cum/cum)", "1 cum"),
            ("Carriage of Stone Aggregate (below 40 mm)",
             "Mechanical transport of stone aggregate to batching platform",
             "0.85 cum per cum combined",
             "Material", "0.85 cum (0.850 cum/cum)", "1 cum"),
            ("Coarse Sand (Zone III)",
             "Zone III coarse sand; same proportion as other 1:1.5:3 items",
             "0.425 cum per cum",
             "Material", "0.425 cum (0.425 cum/cum)", "1 cum"),
            ("Carriage of Coarse Sand",
             "Transport of coarse sand to batching area",
             "0.425 cum per cum",
             "Material", "0.425 cum (0.425 cum/cum)", "1 cum"),
            ("Portland Cement",
             "OPC binder; 0.40 tonne per cum for 1:1.5:3 — unchanged from other items",
             "0.40 tonne per cum; from 10.692/26.73",
             "Material", "0.40 tonne (0.400 t/cum)", "1 cum"),
            ("Carriage of Cement",
             "Transport of cement bags to batching point",
             "0.40 tonne per cum",
             "Material", "0.40 tonne (0.400 t/cum)", "1 cum"),
            ("Mason (average)",
             "Skilled mason for placing concrete in curved/sloped forms; requires careful control of workability for arch/dome shapes",
             "0.24 day per cum — same as beam/slab (5.3); overhead curved work equals slab complexity",
             "Labour", "0.24 day (4.2 cum/day)", "1 cum"),
            ("Beldar",
             "Semi-skilled labourer loading mixer and distributing concrete into complex curved arch/dome forms",
             "2.75 day per cum — same as 5.3; arch forms require equivalent labour intensity to beams",
             "Labour", "2.75 day (0.36 cum/day)", "1 cum"),
            ("Bhisti",
             "Water carrier for mixing water and curing of curved dome/arch concrete surfaces",
             "0.90 day per cum",
             "Labour", "0.90 day (1.1 cum/day)", "1 cum"),
            ("Concrete Mixer (0.25–0.40 cum with hopper)",
             "Drum mixer for batching arch/dome concrete",
             "0.08 day per cum",
             "Machine", "0.08 day (12.5 cum/day)", "1 cum"),
            ("Vibrator (Needle type 40 mm)",
             "Poker vibrator for compacting concrete in curved forms; critical to prevent voids under arch soffits",
             "0.08 day per cum",
             "Machine", "0.08 day (12.5 cum/day)", "1 cum"),
            ("Sundries",
             "Measuring boxes, pans, buckets and minor consumables; slightly higher for arch geometry",
             "L.S. 15.03 per cum — marginally higher than flat work (14.30) due to arch geometry",
             "Equipment", "L.S. 15.03/cum", "1 cum"),
            ("Extra — Mason 1st Class (curved surface)",
             "Additional skilled mason time for placing concrete over curved formwork maintaining correct thickness and profile of arch/dome",
             "5.00 day total / 26.73 cum = 0.19 day/cum extra over standard placing labour",
             "Labour", "0.19 day extra (curved)", "1 cum"),
            ("Extra — Mason 2nd Class (curved surface)",
             "Additional second-class mason for curved surface placement assistance",
             "5.00 day / 26.73 cum = 0.19 day/cum extra",
             "Labour", "0.19 day extra (curved)", "1 cum"),
            ("Extra — Coolie (curved surface)",
             "Additional unskilled help for awkward concrete handling and compacting on curved surfaces",
             "4.50 day / 26.73 cum = 0.17 day/cum extra; restricted access on arch/dome forms",
             "Labour", "0.17 day extra (curved)", "1 cum"),
            ("Coolie (extra — lifting to floor V)",
             "Additional labour for hoisting materials and concrete to upper floors",
             "0.75 × 2.5 × 1 cum = 1.88 day per cum extra; same standard as other floor V items",
             "Labour", "1.88 day extra lifting", "1 cum"),
        ]
    },
    # ── 5.6 ───────────────────────────────────────────────────────────────
    {
        "id": "5.6",
        "parent_title": "5.0 REINFORCED CEMENT CONCRETE — IN CHIMNEYS AND SHAFTS (UP TO FLOOR V)",
        "title": "5.6 RCC 1:1.5:3 in chimneys, shafts, up to floor five level",
        "unit": "cum",
        "base_qty": 1.0,
        "rows": [
            ("RCC 1:1.5:3 in Walls (ref Item 5.2.2 — RCC Work)",
             "All concrete materials, shuttering, scaffold and standard placing labour as per Item 5.2.2 — RCC in walls above plinth to floor V",
             "Reference rate ₹9306.00/cum (marked A); chimney/shaft sections are treated as vertical walls with same resources",
             "Material", "1.00 cum @ ₹9306.00 (A)", "1 cum"),
            ("Coolie (extra — lifting materials in chimney shaft)",
             "Additional unskilled labour for hoisting all materials through restricted chimney/shaft openings; more confined than open wall access",
             "0.26 day per cum extra over ref 5.2.2; chimney throat restricts hoist access requiring more manual carry",
             "Labour", "0.26 day extra (confined lift)", "1 cum"),
        ]
    },
    # ── 5.7 ───────────────────────────────────────────────────────────────
    {
        "id": "5.7",
        "parent_title": "5.0 REINFORCED CEMENT CONCRETE — IN WELL STEINING",
        "title": "5.7 RCC 1:1.5:3 in well steining",
        "unit": "cum",
        "base_qty": 1.0,
        "rows": [
            ("Cement Concrete 1:1.5:3 (ref Item 4.1.2 — Concrete Work)",
             "Complete plain concrete at 1:1.5:3 as per Ch.4 Item 4.1.2; all aggregate, sand, cement, mixing and placing labour included in ref rate",
             "Reference rate ₹7210.55/cum (marked A); well-steining uses same plain concrete base with extra labour for underwater/confined conditions",
             "Material", "1.00 cum @ ₹7210.55 (A)", "1 cum"),
            ("Coolie (extra — well steining conditions)",
             "Additional unskilled labour for confined well-shaft conditions: lowering materials by rope, working in restricted circular space at depth",
             "0.08 day per cum extra over ref 4.1.2; lower than chimney extra (0.26 day) as well steining is below ground and materials are lowered not hoisted",
             "Labour", "0.08 day extra (well shaft)", "1 cum"),
        ]
    },
    # ── 5.8 ───────────────────────────────────────────────────────────────
    {
        "id": "5.8",
        "parent_title": "5.0 REINFORCED CEMENT CONCRETE — VERTICAL/HORIZONTAL FINS, BOX LOUVERS, FACIAS, EAVES BOARDS (PLINTH TO FLOOR V)",
        "title": "5.8 RCC 1:1.5:3 in vertical/horizontal fins, box louvers, facias and eaves boards, above plinth to floor five",
        "unit": "cum",
        "base_qty": 1.0,
        "rows": [
            ("Stone Aggregate 20mm",
             "Single-size 20 mm crushed stone; same 1:1.5:3 mix — thin fin sections require identical mix for strength",
             "0.57 cum per cum; calculated from 0.66 cum batch: 0.3762/0.66",
             "Material", "0.57 cum (0.570 cum/cum)", "1 cum"),
            ("Stone Aggregate 10mm",
             "Single-size 10 mm stone; same as other 1:1.5:3 items",
             "0.28 cum per cum; from 0.1848/0.66",
             "Material", "0.28 cum (0.280 cum/cum)", "1 cum"),
            ("Carriage of Stone Aggregate (below 40 mm)",
             "Transport of stone aggregate to batching platform",
             "0.85 cum per cum; from 0.561/0.66",
             "Material", "0.85 cum (0.850 cum/cum)", "1 cum"),
            ("Coarse Sand (Zone III)",
             "Zone III coarse sand; same proportion",
             "0.425 cum per cum; from 0.2805/0.66",
             "Material", "0.425 cum (0.425 cum/cum)", "1 cum"),
            ("Carriage of Coarse Sand",
             "Transport of coarse sand",
             "0.425 cum per cum",
             "Material", "0.425 cum (0.425 cum/cum)", "1 cum"),
            ("Portland Cement",
             "OPC binder; 0.40 tonne per cum; from 0.264/0.66",
             "Same cement content as other 1:1.5:3 items",
             "Material", "0.40 tonne (0.400 t/cum)", "1 cum"),
            ("Carriage of Cement",
             "Transport of cement bags to batching point",
             "0.40 tonne per cum",
             "Material", "0.40 tonne (0.400 t/cum)", "1 cum"),
            ("Mason 1st Class",
             "Skilled first-class mason for precision fin and louver concrete placement; thin sections need expert control of workability",
             "0.09 day per cum; from 0.06/0.66 — thin fins need more care per unit volume than mass concrete",
             "Labour", "0.09 day (11 cum/day)", "1 cum"),
            ("Mason 2nd Class",
             "Second-class mason assisting 1st class in fin placement and surface finishing",
             "0.09 day per cum; from 0.06/0.66",
             "Labour", "0.09 day (11 cum/day)", "1 cum"),
            ("Beldar",
             "Semi-skilled labourer for aggregate handling and concrete pouring into small fin forms",
             "1.20 day per cum; from 0.79/0.66 — lower than wall work (1.23/cum) since fins are smaller discrete pours",
             "Labour", "1.20 day (0.83 cum/day)", "1 cum"),
            ("Coolie",
             "Unskilled helper for concrete handling and material supply to fin locations",
             "0.85 day per cum; from 0.56/0.66",
             "Labour", "0.85 day (1.18 cum/day)", "1 cum"),
            ("Bhisti",
             "Water carrier for mixing water and curing of exposed fin surfaces",
             "0.91 day per cum; from 0.60/0.66 — marginally higher than mass concrete due to high surface-area-to-volume ratio",
             "Labour", "0.91 day (1.1 cum/day)", "1 cum"),
            ("Concrete Mixer (0.25–0.40 cum with hopper)",
             "Drum mixer for batching small fin concrete pours",
             "0.076 day per cum; from 0.05/0.66",
             "Machine", "0.076 day", "1 cum"),
            ("Vibrator (Needle type 40 mm)",
             "Poker vibrator for compaction in narrow fin sections; essential to fill around close-spaced bars in thin sections",
             "0.076 day per cum; from 0.05/0.66",
             "Machine", "0.076 day", "1 cum"),
            ("Scaffolding",
             "Working scaffold for access to vertical and horizontal fin/louver locations; covers full wall face",
             "L.S. 45.70 per cum — much higher than wall work (14.36/cum) due to large scaffold area relative to thin fin volume",
             "Equipment", "L.S. 45.70/cum", "1 cum"),
            ("Sundries",
             "Measuring boxes, small pans, buckets and minor consumables",
             "L.S. 14.38 per cum; from 9.49/0.66",
             "Equipment", "L.S. 14.38/cum", "1 cum"),
            ("Coolie (extra — restricted fin working)",
             "Additional unskilled labour for working inside restricted fin and louver box cavities; compaction and clean-up in confined spaces",
             "0.27 day per cum extra; from 0.18/0.66 — access restriction within fin box is more severe than open wall sections",
             "Labour", "0.27 day extra (confined fins)", "1 cum"),
            ("Extra — Mason 1st Class (fin geometry)",
             "Additional first-class mason time for precision alignment and striking off at top of fin sections",
             "0.076 day per cum extra; from 0.05/0.66",
             "Labour", "0.076 day extra", "1 cum"),
            ("Extra — Mason 2nd Class (fin geometry)",
             "Additional second-class mason for fin geometry work",
             "0.076 day per cum extra",
             "Labour", "0.076 day extra", "1 cum"),
            ("Extra — Beldar (fin geometry)",
             "Additional beldar for concrete distribution in confined fin cavities",
             "0.15 day per cum extra; from 0.10/0.66",
             "Labour", "0.15 day extra", "1 cum"),
            ("Extra — Bhisti (fin geometry)",
             "Additional water carrier for curing of high surface-area fin and louver sections",
             "0.23 day per cum extra; from 0.15/0.66",
             "Labour", "0.23 day extra", "1 cum"),
        ]
    },
    # ── 5.9.1 ─────────────────────────────────────────────────────────────
    {
        "id": "5.9.1",
        "parent_title": "5.9 CENTERING & SHUTTERING (incl. strutting, propping and removal of form) — per sqm of contact area",
        "title": "5.9.1 Centering & shuttering — Foundations, footings, bases of columns etc. for mass concrete",
        "unit": "sqm",
        "base_qty": 1.0,
        "rows": [
            ("Wall Form Panel 1250×500 mm",
             "Steel wall-form panel providing smooth rigid shuttering face for foundation/footing sides; reused up to 40 times",
             "0.34 each per sqm; qty = 16×0.85/40 = 0.34 (using once cost); @ ₹860 per panel",
             "Material", "0.34 each (per sqm)", "1 sqm"),
            ("Corner Angle 45×45×5 mm, 1.50 m long",
             "Steel angle providing rigid corner junction between adjacent form panels to prevent blowout at panel joints",
             "0.085 each per sqm; @ ₹240 per piece",
             "Material", "0.085 each (per sqm)", "1 sqm"),
            ("100 mm Channel Shoulder, 2.50 m long",
             "Rolled steel channel waler providing horizontal stiffness to panel face against wet concrete pressure",
             "0.17 each per sqm; @ ₹910 per piece",
             "Material", "0.17 each (per sqm)", "1 sqm"),
            ("Double Clip (bridge clip)",
             "Steel spring clip locking adjacent panel edges together; prevents panel splaying under concrete pressure",
             "0.34 each per sqm; @ ₹76 per clip",
             "Material", "0.34 each (per sqm)", "1 sqm"),
            ("Single Clip",
             "Steel spring clip for single-panel-edge locking where panels abut at footing ends",
             "0.17 each per sqm; @ ₹59 per clip",
             "Material", "0.17 each (per sqm)", "1 sqm"),
            ("M.S. Tube 40 mm dia",
             "Mild steel tube used as push-pull prop or tie rod across narrow footing form faces",
             "0.2295 m per sqm; @ ₹215 per metre",
             "Material", "0.2295 m (per sqm)", "1 sqm"),
            ("Assembly Nuts & Bolts",
             "Fasteners for joining form panels, walers and clips during erection and striking",
             "L.S. 22.10 per sqm; from 1040×0.85/40 = 22.10 allowance",
             "Material", "L.S. 22.10/sqm", "1 sqm"),
            ("Carriage",
             "Transport, loading and unloading of all shuttering hardware from yard to site and between pours",
             "L.S. 78.00 per sqm",
             "Material", "L.S. 78.00/sqm", "1 sqm"),
            ("Fitter (Grade 1)",
             "Skilled fitter erecting, aligning, bracing and striking steel formwork panels to specification",
             "0.75 day per sqm; @ ₹738/day; approx 1.3 sqm/day for footing shuttering",
             "Labour", "0.75 day (1.3 sqm/day)", "1 sqm"),
            ("Beldar",
             "Semi-skilled labourer assisting fitter in panel handling, propping and dismantling",
             "1.50 day per sqm; 2:1 ratio Beldar:Fitter for erection and striking",
             "Labour", "1.50 day (0.67 sqm/day)", "1 sqm"),
            ("Shuttering Oil",
             "Release agent (mineral/vegetable oil) applied to form face before concrete pour to prevent adhesion and ease striking",
             "L.S. 52.00 per sqm",
             "Equipment", "L.S. 52.00/sqm", "1 sqm"),
            ("Sundries",
             "Wire ties, nails, wedges, spacers, and minor accessories for shuttering assembly",
             "L.S. 26.00 per sqm",
             "Equipment", "L.S. 26.00/sqm", "1 sqm"),
        ]
    },
    # ── 5.9.2 ─────────────────────────────────────────────────────────────
    {
        "id": "5.9.2",
        "parent_title": "5.9 CENTERING & SHUTTERING (incl. strutting, propping and removal of form) — per sqm of contact area",
        "title": "5.9.2 Centering & shuttering — Walls (any thickness) incl. pilasters, buttresses, plinth and string courses",
        "unit": "sqm",
        "base_qty": 1.0,
        "rows": [
            ("Wall Form Panel 1250×500 mm",
             "Steel wall-form panel for both faces of RC wall; two-sided formwork doubles panel count vs footings",
             "0.51 each per sqm; 2×3×2×0.85/40 = 0.51; @ ₹860 per panel",
             "Material", "0.51 each (per sqm)", "1 sqm"),
            ("100 mm Channel Shoulder, 2.50 m long",
             "Horizontal waler stiffening wall form panels against hydrostatic pressure of tall concrete walls",
             "0.17 each per sqm; @ ₹910 per piece",
             "Material", "0.17 each (per sqm)", "1 sqm"),
            ("Double Clip (bridge clip)",
             "Steel spring clip locking panel edges; double count for two-face wall shuttering",
             "0.51 each per sqm; @ ₹76 per clip",
             "Material", "0.51 each (per sqm)", "1 sqm"),
            ("Single Clip",
             "Steel spring clip for panel edge joints at ends of wall runs",
             "0.255 each per sqm; @ ₹59 per clip",
             "Material", "0.255 each (per sqm)", "1 sqm"),
            ("M.S. Tube 40 mm dia",
             "Tie-rod tube passing through wall thickness holding two form faces together against concrete pressure",
             "0.68 m per sqm; @ ₹215 per metre; through-ties critical for wall sections >200 mm",
             "Material", "0.68 m (per sqm)", "1 sqm"),
            ("Nut & Bolts",
             "Fasteners securing walers to panels and tie-rods to walers",
             "L.S. 27.62 per sqm",
             "Material", "L.S. 27.62/sqm", "1 sqm"),
            ("Carriage",
             "Transport of heavy wall form panels and walers between pours",
             "L.S. 78.00 per sqm",
             "Material", "L.S. 78.00/sqm", "1 sqm"),
            ("Fitter (Grade 1)",
             "Skilled fitter erecting two-sided wall formwork, threading tie-rods and aligning panels plumb",
             "3.50 day per sqm — significantly higher than footings (0.75 day); two-face walls with tie-rods and plumb checking need more skill and time",
             "Labour", "3.50 day (0.29 sqm/day)", "1 sqm"),
            ("Beldar",
             "Semi-skilled labourer handling heavy wall panels, propping and striking both faces",
             "6.00 day per sqm; 1.7:1 ratio Beldar:Fitter; wall panel handling requires substantial manual effort",
             "Labour", "6.00 day (0.17 sqm/day)", "1 sqm"),
            ("Shuttering Oil",
             "Release agent applied to both wall form faces before each pour",
             "L.S. 78.00 per sqm — higher than footing (52.00) since both faces of wall are coated",
             "Equipment", "L.S. 78.00/sqm", "1 sqm"),
            ("Sundries",
             "Wire ties, wedges, cone/spacers for tie-rod plastic cones, and accessories",
             "L.S. 52.00 per sqm",
             "Equipment", "L.S. 52.00/sqm", "1 sqm"),
        ]
    },
    # ── 5.9.3 ─────────────────────────────────────────────────────────────
    {
        "id": "5.9.3",
        "parent_title": "5.9 CENTERING & SHUTTERING (incl. strutting, propping and removal of form) — per sqm of contact area",
        "title": "5.9.3 Centering & shuttering — Suspended floors, roofs, landings, balconies and access platforms",
        "unit": "sqm",
        "base_qty": 1.0,
        "rows": [
            ("S.H. Steel Work (ref Item 10.1 — Steel Work)",
             "Pressed steel plates (0.75×0.60 m, 1.6 mm thick) used as flat decking spanning between props; ref SH Steel Work rate for fabricated pressed plates",
             "9.2055 kg per sqm; plate weight 14.44 kg × 5×6 plates = 433.2 kg; qty once = 433.2×0.85/40 = 9.2055 kg; @ ₹86.05/kg (A)",
             "Material", "9.2055 kg @ ref 10.1 (A)", "1 sqm"),
            ("Adjustable Span Prop ESO+SI (2.35–3.40 m)",
             "Proprietary adjustable steel prop (inner and outer tubes) supporting deck plates at standard floor heights 2.35–3.40 m",
             "0.1063 each per sqm; 5×0.85/40 = 0.1063; @ ₹1480 per prop",
             "Material", "0.1063 each (per sqm)", "1 sqm"),
            ("Adjustable Telescopic Prop 3 m (2.02–3.75 m)",
             "Secondary telescopic prop providing additional support between ESO props; covers intermediate span points",
             "0.1275 each per sqm; 6×0.85/40 = 0.1275; @ ₹955 per prop",
             "Material", "0.1275 each (per sqm)", "1 sqm"),
            ("Assembly Nuts & Bolts",
             "Fasteners for prop head-plate connections and plate-to-plate joints",
             "L.S. 22.10 per sqm; 1040×0.85/40 = 22.10",
             "Material", "L.S. 22.10/sqm", "1 sqm"),
            ("Carriage",
             "Transport of heavy steel deck plates and props from yard to site and between floors",
             "L.S. 130.00 per sqm — higher than wall shuttering (78.00) due to weight of pressed steel decking system",
             "Material", "L.S. 130.00/sqm", "1 sqm"),
            ("Fitter (Grade 1)",
             "Skilled fitter erecting prop grid, adjusting prop heights, laying deck plates level and plumb-checking soffit",
             "3.00 day per sqm; @ ₹738/day; soffit forming requires accurate level-setting of all props",
             "Labour", "3.00 day (0.33 sqm/day)", "1 sqm"),
            ("Beldar",
             "Semi-skilled labourer handling deck plates and props, adjusting base plates and striking after concrete gains strength",
             "6.00 day per sqm; 2:1 Beldar:Fitter; heavy plate work requires additional manual labour",
             "Labour", "6.00 day (0.17 sqm/day)", "1 sqm"),
            ("Shuttering Oil",
             "Release agent applied to top face of steel deck plates before each pour",
             "L.S. 78.00 per sqm",
             "Equipment", "L.S. 78.00/sqm", "1 sqm"),
            ("Sundries / Paper Tape",
             "Paper tape sealing plate joints to prevent concrete slurry loss; sundry fasteners and packing",
             "L.S. 49.70 per sqm — paper tape is critical for pressed plate systems to prevent bleeding through joints",
             "Equipment", "L.S. 49.70/sqm", "1 sqm"),
        ]
    },
    # ── 5.9.5 ─────────────────────────────────────────────────────────────
    {
        "id": "5.9.5",
        "parent_title": "5.9 CENTERING & SHUTTERING (incl. strutting, propping and removal of form) — per sqm of contact area",
        "title": "5.9.5 Centering & shuttering — Lintels, beams, plinth beams, girders, bressumers and cantilevers",
        "unit": "sqm",
        "base_qty": 1.0,
        "rows": [
            ("S.H. Steel Work — Pressed Steel Plates (ref Item 10.1)",
             "Pressed steel plates (1.20×0.50 m, 1.6 mm thick) for beam side and soffit; angle stiffeners at edges; ref SH Steel Work",
             "6.4356 kg per sqm; 3×5×20.19 kg = 302.85 kg; qty once = 302.85×0.85/40 = 6.4356 kg; @ ₹86.05/kg (A)",
             "Material", "6.4356 kg @ ref 10.1 (A)", "1 sqm"),
            ("Adjustable Telescopic Prop 3 m (2.02–3.75 m)",
             "Prop supporting beam bottom plate; fewer props than slabs as beam loads concentrate on soffit not distributed across floor area",
             "0.1275 each per sqm; 6×0.85/40 = 0.1275; @ ₹955 per prop",
             "Material", "0.1275 each (per sqm)", "1 sqm"),
            ("Beam Clamp 300–380 mm (450–1070 mm range)",
             "Proprietary U-clamp spanning beam width to hold side plates vertical against concrete pressure; replaces through-ties for beams",
             "0.1063 each per sqm; 5×0.85/40 = 0.1063 sets; @ ₹355 per set",
             "Material", "0.1063 sets (per sqm)", "1 sqm"),
            ("Assembly Nuts & Bolts",
             "Fasteners for plate joints, prop head-plates and clamp bolts",
             "L.S. 22.10 per sqm",
             "Material", "L.S. 22.10/sqm", "1 sqm"),
            ("Carriage",
             "Transport of beam form plates, clamps and props",
             "L.S. 78.00 per sqm — same as wall/footing shuttering",
             "Material", "L.S. 78.00/sqm", "1 sqm"),
            ("Fitter (Grade 1)",
             "Skilled fitter assembling U-channel beam bottom, side plates, and beam clamps; setting soffit level with adjustable props",
             "1.25 day per sqm; @ ₹738/day; less complex than slab decking (3.00 day) as beam contact area is smaller per component",
             "Labour", "1.25 day (0.80 sqm/day)", "1 sqm"),
            ("Beldar",
             "Semi-skilled labourer handling beam plates and props; 2:1 ratio with Fitter",
             "2.50 day per sqm",
             "Labour", "2.50 day (0.40 sqm/day)", "1 sqm"),
            ("Shuttering Oil",
             "Release agent on beam soffit and side plates",
             "L.S. 39.00 per sqm — lower than slab (78.00) as beam contact area per sqm is predominantly soffit not top surface",
             "Equipment", "L.S. 39.00/sqm", "1 sqm"),
            ("Sundries / Paper Tape",
             "Paper tape sealing plate joints at beam corners; sundry fasteners and packing shims",
             "L.S. 24.61 per sqm",
             "Equipment", "L.S. 24.61/sqm", "1 sqm"),
        ]
    },
    # ── 5.9.6 ─────────────────────────────────────────────────────────────
    {
        "id": "5.9.6",
        "parent_title": "5.9 CENTERING & SHUTTERING (incl. strutting, propping and removal of form) — per sqm of contact area",
        "title": "5.9.6 Centering & shuttering — Columns, Pillars, Piers, Abutments, Posts and Struts",
        "unit": "sqm",
        "base_qty": 1.0,
        "rows": [
            ("Wall Form Panel 1250×450 mm",
             "Narrower 450 mm wide steel form panel fitting the 450 mm column face; four panels wrap around column perimeter",
             "0.17 each per sqm; 8×0.85/40 = 0.17; @ ₹860 per panel",
             "Material", "0.17 each (per sqm)", "1 sqm"),
            ("Corner Angle 45×45×5 mm, 2.50 m long",
             "Longer corner angle (2.5 m vs 1.5 m in footings) spanning full column height; provides rigid corner junction between face panels",
             "0.085 each per sqm; 4×0.85/40 = 0.085; @ ₹255 per piece",
             "Material", "0.085 each (per sqm)", "1 sqm"),
            ("Column Clamp 450×1070 mm",
             "Proprietary yoke clamp encircling column form at regular vertical intervals to resist hoop bursting pressure from wet concrete",
             "0.1063 each per sqm; 5×0.85/40 = 0.1063; @ ₹965 per clamp",
             "Material", "0.1063 each (per sqm)", "1 sqm"),
            ("Prop 2 m (adjustable 3–3.5 m)",
             "Inclined prop providing lateral stability to column form against accidental impact and wind during concrete pour",
             "0.085 each per sqm; 4×0.85/40 = 0.085; @ ₹635 per prop",
             "Material", "0.085 each (per sqm)", "1 sqm"),
            ("Assembly Nut & Bolt",
             "Fasteners for column clamp bolts and panel-corner angle connections",
             "L.S. 27.62 per sqm; 1300×0.85/40 = 27.63",
             "Material", "L.S. 27.62/sqm", "1 sqm"),
            ("Carriage",
             "Transport of column form panels, clamps and props to site and between pours",
             "L.S. 52.00 per sqm — lower than wall (78.00) as column form kits are more compact and lighter per sqm",
             "Material", "L.S. 52.00/sqm", "1 sqm"),
            ("Fitter (Grade 1)",
             "Skilled fitter assembling four panel faces around column reinforcement, fitting corner angles and column clamps at each tier",
             "1.00 day per sqm; @ ₹738/day; column forms are simpler than two-face walls but require careful plumb alignment",
             "Labour", "1.00 day (1.0 sqm/day)", "1 sqm"),
            ("Beldar",
             "Semi-skilled labourer handling column panels, tightening clamp bolts and supporting fitter",
             "2.00 day per sqm; 2:1 ratio with Fitter",
             "Labour", "2.00 day (0.5 sqm/day)", "1 sqm"),
            ("Shuttering Oil",
             "Release agent applied to inner face of column panels before each pour",
             "L.S. 39.00 per sqm — same as beam (5.9.5); only inner column face area is coated",
             "Equipment", "L.S. 39.00/sqm", "1 sqm"),
            ("Sundries",
             "Wire ties, wedges, plastic spacer cones for tie-rods, and minor accessories",
             "L.S. 26.00 per sqm — includes carriage and minor consumables for column form kit",
             "Equipment", "L.S. 26.00/sqm", "1 sqm"),
        ]
    },
]

# ── Sheet builder ─────────────────────────────────────────────────────────

def write_item_block(ws, item, start_row):
    r = start_row
    item_code = item["id"]
    base_qty = item["base_qty"]
    unit = item["unit"]

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
    ws.row_dimensions[r].height = 30
    for col in range(1, 8):
        ws.cell(row=r, column=col).border = thin_border()
    r += 1

    # 3. Spacing row
    ws.row_dimensions[r].height = 6
    r += 1

    # 4. Column Header
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

        c1 = ws.cell(row=r, column=1, value=item_code)
        c1.font = Font(name="Calibri", size=9, bold=True)
        c1.alignment = AL_C
        c1.fill = row_fill
        c1.border = thin_border()

        c2 = ws.cell(row=r, column=2, value=disp_name)
        c2.font = Font(name="Calibri", size=9, bold=(category in ("Machine", "Equipment", "Material")))
        c2.alignment = AL_L
        c2.fill = row_fill
        c2.border = thin_border()

        c3 = ws.cell(row=r, column=3, value=work_done)
        c3.font = Font(name="Calibri", size=9)
        c3.alignment = AL_L
        c3.fill = row_fill
        c3.border = thin_border()

        c4 = ws.cell(row=r, column=4, value=condition)
        c4.font = Font(name="Calibri", size=9)
        c4.alignment = AL_L
        c4.fill = row_fill
        c4.border = thin_border()

        c5 = ws.cell(row=r, column=5, value=category)
        c5.font = Font(name="Calibri", size=9)
        c5.alignment = AL_C
        c5.fill = row_fill
        c5.border = thin_border()

        c6 = ws.cell(row=r, column=6, value=prod_str)
        c6.font = Font(name="Calibri", size=9, bold=True)
        c6.alignment = AL_R
        c6.fill = row_fill
        c6.border = thin_border()

        c7 = ws.cell(row=r, column=7, value=row_qty)
        c7.font = Font(name="Calibri", size=9)
        c7.alignment = AL_C
        c7.fill = row_fill
        c7.border = thin_border()

        r += 1

    ws.row_dimensions[r].height = 12
    r += 1
    return r


def build_rcc_sheet(ws):
    if ws.views.sheetView:
        ws.views.sheetView[0].showGridLines = True

    ws.merge_cells("A1:G1")
    t1 = ws.cell(row=1, column=1, value="Sub-Head 5.0 — REINFORCED CEMENT CONCRETE  |  First-Principles Resource, Work & Gang Analysis")
    t1.font = Font(name="Calibri", size=13, bold=True, color=C_WHITE)
    t1.fill = fill(C_TITLE_BG)
    t1.alignment = AL_C
    ws.row_dimensions[1].height = 28
    for col in range(1, 8):
        ws.cell(row=1, column=col).border = thin_border()

    ws.merge_cells("A2:G2")
    t2 = ws.cell(row=2, column=1,
                 value="Evidence: CPWD DAR 2019 Vol 1 Ch.5 (pp.181–204)  |  Format: Item Code • Labour/Machine/Material • Work done • Condition/When used • Category • Productivity • Quantity  |  Rates: W+X+Y+Z chain (W=+1% water, X=+14.05% GST, Y=+15% CPOH, Z=+1% Cess)")
    t2.font = Font(name="Calibri", size=9, italic=True, color="333333")
    t2.fill = fill(C_SUBTITLE_BG)
    t2.alignment = AL_L
    ws.row_dimensions[2].height = 18
    for col in range(1, 8):
        ws.cell(row=2, column=col).border = thin_border()

    ws.row_dimensions[3].height = 8

    current_row = 4
    for item in RCC_ITEMS:
        current_row = write_item_block(ws, item, current_row)

    # DAR rates reference row at the end
    ws.row_dimensions[current_row].height = 8
    current_row += 1
    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=7)
    ref_cell = ws.cell(row=current_row, column=1,
                       value="DAR 2019 Consolidated Rates (W+X+Y+Z applied): 5.1.2=₹7718/cum | 5.1.3=₹7296/cum | 5.2.2=₹9306/cum | 5.3=₹9764/cum | 5.4=₹8963/cum | 5.5=₹10289/cum | 5.6=₹9500/cum | 5.7=₹7270/cum | 5.8=₹8550/cum || Shuttering: 5.9.1=₹285/sqm | 5.9.2=₹609/sqm | 5.9.3=5.9.4=₹693/sqm | 5.9.5=₹552/sqm | 5.9.6=₹734/sqm | 5.9.7=₹622/sqm | 5.9.8=₹600/sqm | 5.9.9=₹1713/sqm | 5.9.11=₹609/sqm | 5.9.12=₹246/sqm | 5.9.13=₹1024/sqm | 5.9.15=₹285/sqm | 5.9.16.1=₹173/m | 5.9.16.2=₹737/sqm")
    ref_cell.font = Font(name="Calibri", size=8, italic=True, color="444444")
    ref_cell.fill = fill(C_SUBTITLE_BG)
    ref_cell.alignment = AL_L
    for col in range(1, 8):
        ws.cell(row=current_row, column=col).border = thin_border()
    ws.row_dimensions[current_row].height = 30

    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 34
    ws.column_dimensions["C"].width = 54
    ws.column_dimensions["D"].width = 52
    ws.column_dimensions["E"].width = 16
    ws.column_dimensions["F"].width = 28
    ws.column_dimensions["G"].width = 16

    ws.sheet_view.showGridLines = True
    ws.freeze_panes = "A4"


def rebuild_rcc_in_workbook(file_path):
    print(f"Opening {file_path} ...")
    wb = openpyxl.load_workbook(file_path)
    if "05_RCC_Work" not in wb.sheetnames:
        raise ValueError(f"'05_RCC_Work' sheet not found in {file_path}")

    pos = wb.sheetnames.index("05_RCC_Work")
    del wb["05_RCC_Work"]
    ws = wb.create_sheet("05_RCC_Work", pos)
    ws.sheet_properties.tabColor = "2E75B6"
    build_rcc_sheet(ws)

    temp_path = file_path.replace(".xlsx", "_TMP_RCC.xlsx")
    wb.save(temp_path)
    wb.close()
    os.replace(temp_path, file_path)
    print(f"Successfully updated {file_path} -> sheet '05_RCC_Work' ({len(RCC_ITEMS)} items).")


def main():
    repo_root = Path(__file__).resolve().parents[1]
    main_wb = repo_root / "CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_1_Latest.xlsx"
    output_wb = repo_root / "outputs" / "earthwork-custom-rate-composer" / "CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_1_Latest.xlsx"

    if main_wb.exists():
        rebuild_rcc_in_workbook(str(main_wb))
    if output_wb.exists():
        rebuild_rcc_in_workbook(str(output_wb))


if __name__ == "__main__":
    main()
