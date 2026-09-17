"""First-Principles Knowledge Base for CPWD Earth Work (Sub-Head 02).

Provides contextual engineering explanations:
- Work done: Specific physical operation performed by each labour, machine, material, or reference
- Condition / When used: Site circumstances, layer depths, leads, lifts, soil types, moisture rules, and mathematically correlated output rates
- Category: Labour, Machine, Equipment, Material, Reference
- Productivity: Formatted effort (man-hrs or machine-hrs) with exact mathematical output rate correlation (e.g. 0.50 machine-hrs (20.00 cum/hr))
"""

from __future__ import annotations
import re

SHIFT_HOURS = 8.0

# General resource role mappings by trade/craft/material
_DEFAULT_ROLES: dict[str, tuple[str, str, str]] = {
    # res_code -> (default_work_done, default_condition, default_category)
    # --- LABOUR ---
    "0114": (
        "Manual excavation, breaking soil clods, trimming trench sides, ramming bottoms, and spreading earth in layers",
        "Primary manual labour; used in all manual cutting, dressing, filling, and formation preparation",
        "Labour",
    ),
    "0115": (
        "Carrying and shifting excavated earth/debris by head-load baskets, loading into carts/skips, and disposing within 50 m lead",
        "Manual handling and disposal; required where motorized transport is not deployed for lead up to 50 m and lift up to 1.5 m",
        "Labour",
    ),
    "0101": (
        "Sprinkling and watering earth layers to bring dry soil to Optimum Moisture Content (OMC) prior to rolling/ramming",
        "Used during compaction and filling operations; water applied uniformly in layers not exceeding 20 cm",
        "Labour",
    ),
    "0113": (
        "General watch, site safety guarding, boundary demarcation, and safeguarding machinery/tools overnight",
        "Mandatory supporting watchman where motorized road rollers or specialized plant are stationed at site",
        "Labour",
    ),
    "0128": (
        "Supervision of excavation gang, alignment checking, depth measurement, and safety control",
        "Used to direct labour gangs in mechanized excavation, trench dressing, and post-construction treatments",
        "Labour",
    ),
    "0132": (
        "Excavating rock, levering fractured rock blocks, removing blasted fragments, and dressing rock faces",
        "Used in rock excavation works in trenches, pits, and open areas",
        "Labour",
    ),
    "0133": (
        "Breaking excavated boulders into transportable fragments using heavy sledge hammers and steel wedges",
        "Used in ordinary rock and hard rock where boulders require manual size reduction before disposal",
        "Labour",
    ),
    "0134": (
        "Drilling jumper holes in rock face for charging explosives or inserting splitting wedges",
        "Used in hard rock excavation requiring blasting or manual wedge splitting",
        "Labour",
    ),
    "0135": (
        "Precision chiselling of rock surface to required trench levels, vertical planes, and founding bedrock",
        "Used in hard rock where blasting is prohibited by safety regulations or adjacent structures",
        "Labour",
    ),
    "0103": (
        "Sharpening chisels, repairing drill jumpers, maintaining crowbars, and dressing steel wedges on site forge",
        "Supporting craftsman for continuous dressing of rock-cutting tools during chiselling operations",
        "Labour",
    ),
    "0112": (
        "Fabricating, cutting, assembling, and placing timber shoring, strutting, and walings; drilling woodwork for termite treatment",
        "Skilled carpentry required in trench/shaft timbering and timber anti-termite treatment",
        "Labour",
    ),
    "0124": (
        "Plugging drilled holes with cement mortar 1:2 to match original floor/wall surface finish after chemical injection",
        "Skilled masonry required in anti-termite post-construction remedial drilling",
        "Labour",
    ),

    # --- MACHINERY & EQUIPMENT ---
    "0003": (
        "Heavy mechanical consolidation and compaction of embankment and subgrade layers",
        "Every 3rd layer and topmost layer; roller capacity >= 8 tonne; loose layers <= 20 cm",
        "Machine",
    ),
    "0020": (
        "Bulk mechanical digging, trench cutting, and loading spoil directly into tippers or spoil heaps",
        "Mechanized excavation in trenches and open areas; operating rate matches soil/rock hardness",
        "Machine",
    ),
    "0018": (
        "Picking up bulk excavated spoil, stockpiling, and charging hauling transport or backfilling trenches",
        "Complementary wheel/tracked loader operating in tandem with hydraulic excavators",
        "Machine",
    ),
    "0017": (
        "Hauling fragmented rock and excavated spoil from trench face to dumping grounds or stacking yards",
        "Mechanized earth and rock haulage within site boundaries; operating rate matches loading cycle",
        "Machine",
    ),
    "EQUIP-0.5T": (
        "Initial compaction of each earth layer (<=20 cm depth)",
        "Compacting loose layers 1 and 2 of each 3-layer cycle; manual/light compaction where specified",
        "Equipment",
    ),
    "EQUIP-DURMAT": (
        "Layer-by-layer compaction of backfilled earth in confined foundation trenches and plinth areas",
        "Confined trench and plinth backfill; layers <= 20 cm where power rollers cannot manoeuvre",
        "Equipment",
    ),

    # --- MATERIALS ---
    "1198": (
        "Protective timber walling/sheeting preventing trench side collapse",
        "Close or open timbering in foundation trenches/shafts; 38 mm thick Kail/Deodar planks",
        "Material",
    ),
    "1197": (
        "Horizontal walings and cross struts bracing timber sheeting against lateral earth pressure",
        "Timbering framework; Kail/Deodar scantlings as per CPWD specifications",
        "Material",
    ),
    "0302": (
        "Vertical and cross strutting poles/props supporting timbering framework",
        "Intermediate trench strutting and propping; Safeda poles/ballies 3 to 4 m long",
        "Material",
    ),
    "7022": (
        "Chemical anti-termite concentrate forming lethal chemical barrier in soil",
        "Pre/post construction termite treatment; diluted to 1% aqueous emulsion as per IS 6313 (Part 2)",
        "Material",
    ),
    "6501": (
        "Plinth filling and under-floor cushioning material",
        "Bedding and plinth filling; clean coarse/medium sand compacted in layers not exceeding 15 cm",
        "Material",
    ),
    "1980": (
        "Bulk filling material for plinth and ground depression reclamation",
        "Pond or mound fly ash compacted in uniform layers for plinth filling under floors",
        "Material",
    ),
    "EXPLOSIVE-POWDER": (
        "High explosive charge fracturing intact bedrock and hard rock mass",
        "Controlled rock excavation requiring blasting; handled by licensed shot-firer",
        "Material",
    ),
    "EXPLOSIVE-FUSE": (
        "Safety fuse providing timed delay ignition for blasting cartridges",
        "Rock blasting operations; regulated burn rate approximately 120 seconds per metre",
        "Material",
    ),
    "KEROSENE": (
        "Solvent vehicle and penetrating oil carrier for wood preservative treatment",
        "Solvent applied with chemical insecticide for remedial anti-termite woodwork treatment",
        "Material",
    ),

    # --- CARRIAGE ---
    "CARRIAGE": (
        "Transportation of materials from source/store to site of work",
        "Standard CPWD carriage lead and handling specification",
        "Carriage",
    ),
}

# Contextual overrides for specific item activities
_ITEM_SPECIFIC_OVERRIDES: dict[tuple[str, str], tuple[str, str, str, str]] = {
    # (item_code, res_code) -> (display_name, work_done, condition, category)
    # --- 2.2 / 2.2.1 Embankment Banking ---
    ("2.2.1", "0114"): (
        "Beldar",
        "Rough excavation, breaking clods, spreading earth in layers, dressing embankment, assisting in ramming/rolling",
        "Main manual earthwork; used throughout excavation, filling and layer preparation",
        "Labour",
    ),
    ("2.2.1", "0115"): (
        "Coolie",
        "Carrying and shifting excavated earth/material, distributing earth over the filling area and assisting Beldars",
        "Used where manual handling/carriage is required; lead up to 50 m",
        "Labour",
    ),
    ("2.2.1", "0101"): (
        "Bhisti",
        "Watering each earth layer to achieve suitable moisture before compaction",
        "Used during layer-wise filling/compaction; water applied as required",
        "Labour",
    ),
    ("2.2.1", "0113"): (
        "Chowkidar",
        "General watch and assistance at the work site; safeguarding materials/equipment",
        "Supporting labour; not directly productive in earthwork",
        "Labour",
    ),
    ("2.2.1", "EQUIP-0.5T"): (
        "½-ton roller / wooden or steel rammers",
        "Initial compaction of each earth layer (<=20 cm depth)",
        "Compacting loose layers 1 and 2 of each 3-layer cycle; manual/light compaction; output = 1.14 cum/hr",
        "Equipment",
    ),
    ("2.2.1", "0003"): (
        "8–10 tonne power roller Diesel road roller",
        "Heavy consolidation/compaction of embankment",
        "Every 3rd layer and the top-most layer; roller capacity >= 8 tonne; output = 156.25 cum/hr",
        "Machine",
    ),

    # --- 2.2.2 Hard Soil Embankment ---
    ("2.2.2", "0114"): (
        "Beldar",
        "Rough excavation in hard soil, breaking clods, spreading earth in layers, dressing embankment, assisting in ramming/rolling",
        "Main manual earthwork in hard soil; used throughout excavation, filling and layer preparation",
        "Labour",
    ),
    ("2.2.2", "0115"): (
        "Coolie",
        "Carrying and shifting hard excavated earth/material, distributing earth over filling area",
        "Used where manual handling/carriage is required; lead up to 50 m",
        "Labour",
    ),
    ("2.2.2", "0101"): (
        "Bhisti",
        "Watering hard soil layers to achieve optimum moisture before rolling",
        "Used during layer-wise filling/compaction; water applied as required",
        "Labour",
    ),
    ("2.2.2", "0113"): (
        "Chowkidar",
        "General watch and assistance at the work site; safeguarding materials/equipment",
        "Supporting labour; not directly productive in earthwork",
        "Labour",
    ),
    ("2.2.2", "EQUIP-0.5T"): (
        "½-ton roller / wooden or steel rammers",
        "Initial compaction of each hard soil layer (<=20 cm depth)",
        "Compacting loose layers 1 and 2 of each 3-layer cycle; output = 1.14 cum/hr",
        "Equipment",
    ),
    ("2.2.2", "0003"): (
        "8–10 tonne power roller Diesel road roller",
        "Heavy consolidation/compaction of hard soil embankment",
        "Every 3rd layer and top-most layer; roller capacity >= 8 tonne; output = 156.25 cum/hr",
        "Machine",
    ),

    # --- 2.3.1 Banking Excavated Earth ---
    ("2.3.1", "0114"): (
        "Beldar",
        "Breaking clods, spreading received excavated earth in uniform 20 cm layers, trimming side slopes, and ramming",
        "Embankment formation from pre-excavated earth; layers <= 20 cm depth",
        "Labour",
    ),
    ("2.3.1", "0115"): (
        "Coolie",
        "Transporting excavated earth by headload from dump heaps to fill location within 50 m lead",
        "Carriage of banking earth; lead <= 50 m and lift <= 1.5 m",
        "Labour",
    ),
    ("2.3.1", "0101"): (
        "Bhisti",
        "Watering deposited earth layers to bring fill to optimum moisture content for compaction",
        "Required before rolling; water added to OMC",
        "Labour",
    ),
    ("2.3.1", "0113"): (
        "Chowkidar",
        "Safeguarding road roller and night watch at embankment site",
        "Continuous site presence during roller deployment",
        "Labour",
    ),
    ("2.3.1", "EQUIP-0.5T"): (
        "½-ton roller / wooden or steel rammers",
        "Initial compaction of intermediate loose earth layers",
        "Compacting loose layers 1 and 2 of each 3-layer cycle; layers <= 20 cm; output = 1.14 cum/hr",
        "Equipment",
    ),
    ("2.3.1", "0003"): (
        "8–10 tonne power roller Diesel road roller",
        "Heavy compaction and final rolling of embankment",
        "Every 3rd layer and topmost layer; minimum 8-tonne capacity; output = 156.25 cum/hr",
        "Machine",
    ),

    # --- 2.6.1 Mechanized Excavation (All kinds of soil) ---
    ("2.6.1", "0020"): (
        "Hydraulic Excavator 0.9 cum",
        "Bulk mechanical digging and loading spoil into tippers",
        "Bulk area excavation in all kinds of soil; rated output = 30.49 cum/hr (10 cum in 0.33 machine-hr)",
        "Machine",
    ),
    ("2.6.1", "0017"): (
        "Tipper 6 cum",
        "Hauling excavated earth within 50 m lead to dump site",
        "Mechanized earth haulage within site; rated output = 34.72 cum/hr (10 cum in 0.29 machine-hr)",
        "Machine",
    ),

    # --- 2.7.1 Mechanized Rock Excavation (Ordinary rock) ---
    ("2.7.1", "0020"): (
        "Hydraulic Excavator 0.9 cum",
        "Mechanical tearing and loading fragmented ordinary rock into tippers",
        "Bulk excavation in ordinary rock; rated output = 20.00 cum/hr (10 cum in 0.50 machine-hr)",
        "Machine",
    ),
    ("2.7.1", "0017"): (
        "Tipper 6 cum",
        "Hauling fragmented ordinary rock to dumping yards",
        "Mechanized rock haulage within site; rated output = 19.84 cum/hr (10 cum in 0.50 machine-hr)",
        "Machine",
    ),

    # --- 2.7.2 Mechanized Rock Excavation (Hard rock, no blasting) ---
    ("2.7.2", "0020"): (
        "Hydraulic Excavator 0.9 cum",
        "Mechanical loading and handling chiselled/split hard rock fragments",
        "Hard rock excavation (blasting prohibited); rated output = 10.00 cum/hr (10 cum in 1.0 machine-hr)",
        "Machine",
    ),
    ("2.7.2", "0017"): (
        "Tipper 6 cum",
        "Hauling heavy hard rock boulders and fragments to disposal dumps",
        "Mechanized hard rock haulage; rated output = 10.00 cum/hr (10 cum in 1.0 machine-hr)",
        "Machine",
    ),

    # --- 2.9.1 Trench Rock Excavation (Ordinary rock) ---
    ("2.9.1", "0020"): (
        "Hydraulic Excavator 0.9 cum",
        "Trench digging and mechanical spoil loading in ordinary rock",
        "Foundation trench excavation in ordinary rock; rated output = 20.00 cum/hr (10 cum in 0.50 machine-hr)",
        "Machine",
    ),
    ("2.9.1", "0017"): (
        "Tipper 6 cum",
        "Hauling ordinary rock trench spoil to disposal heaps",
        "Mechanized rock haulage; rated output = 20.00 cum/hr (10 cum in 0.50 machine-hr)",
        "Machine",
    ),

    # --- 2.25 Trench/Plinth Filling ---
    ("2.25", "0114"): (
        "Beldar",
        "Spreading available excavated earth in trenches/plinth in layers not exceeding 20 cm and operating durmat rammers",
        "Backfilling foundations, plinth, and trenches; layers <= 20 cm",
        "Labour",
    ),
    ("2.25", "0115"): (
        "Coolie",
        "Shifting excavated earth from perimeter spoil heaps to trenches/plinth within 50 m lead",
        "Manual basket haulage; lead <= 50 m and lift <= 1.5 m",
        "Labour",
    ),
    ("2.25", "0101"): (
        "Bhisti",
        "Watering backfilled earth layers to optimum moisture before durmat ramming",
        "Applied during trench/plinth filling for each 20 cm layer",
        "Labour",
    ),
    ("2.25", "EQUIP-DURMAT"): (
        "Wooden or steel rammers (durmats)",
        "Compacting backfilled earth layers in confined foundation trenches and plinth floors",
        "Each layer <= 20 cm; confined areas inaccessible to power rollers; output = 1.56 cum/hr",
        "Equipment",
    ),

    # --- 2.1.1 Surface Excavation ---
    ("2.1.1", "0114"): (
        "Beldar",
        "Surface excavation not exceeding 30 cm depth, clearing vegetation, roots, and trimming ground surface",
        "Surface cutting over wide plan area (>10 sqm, >1.5m width); depth <= 30 cm in ordinary soil",
        "Labour",
    ),
    ("2.1.1", "0115"): (
        "Coolie",
        "Collecting excavated spoil and root debris into baskets and carrying to disposal heap within 50 m lead",
        "Disposal of excavated surface soil up to 50 m lead and 1.5 m lift",
        "Labour",
    ),

    # --- 2.31 Jungle Clearance ---
    ("2.31", "0114"): (
        "Beldar",
        "Clearing thick jungle, cutting thorny brushwood, uprooting rank vegetation, and grubbing trees <=30 cm girth",
        "Site preparation; clearing rank growth and saplings measured at 1 m height",
        "Labour",
    ),
    ("2.31", "0115"): (
        "Coolie",
        "Collecting cleared brushwood, thorny scrub, and cut saplings into disposal stacks within 50 m lead",
        "Stacking serviceable timber and disposing unserviceable scrub",
        "Labour",
    ),

    # --- 2.32 Grass Clearing ---
    ("2.32", "0114"): (
        "Beldar",
        "Scraping and cutting rank grass, weeds, and light surface vegetation with spade/khurpa",
        "Clearing grass and light vegetation from ground prior to earthwork or building construction",
        "Labour",
    ),
    ("2.32", "0115"): (
        "Coolie",
        "Raking, collecting, and carrying cleared grass and organic debris to disposal heaps within 50 m lead",
        "Disposal of cleared organic spoil to prevent contamination of fill",
        "Labour",
    ),

    # --- 2.27 Sand Filling ---
    ("2.27", "6501"): (
        "Sand (Zone IV/V)",
        "Coarse/medium sand for bedding, plinth filling, and floor level cushion",
        "Clean river/quarry sand compacted in layers not exceeding 15 cm in plinth and under floors",
        "Material",
    ),
    ("2.27", "0114"): (
        "Beldar",
        "Spreading and levelling sand in uniform layers not exceeding 15 cm and watering for consolidation",
        "Plinth filling under floors; light watering and wooden float leveling",
        "Labour",
    ),

    # --- 2.34.1 Anti-termite Emulsion Supply ---
    ("2.34.1", "7022"): (
        "Chlorpyriphos 20% EC chemical concentrate",
        "Supplying chemical emulsion in sealed containers for anti-termite treatment",
        "Conforming to IS 8944; diluted to 1% aqueous emulsion for soil/woodwork treatment",
        "Material",
    ),

    # --- 2.37 Fly Ash Filling ---
    ("2.37", "1980"): (
        "Fly ash (pond / mound)",
        "Bulk filling material for plinth, foundation trenches, and low ground reclamation",
        "Compacted in layers not exceeding 20 cm in plinth/embankments as per CPWD specifications",
        "Material",
    ),
    ("2.37", "0114"): (
        "Beldar",
        "Spreading supplied fly ash in layers not exceeding 20 cm, watering, and consolidating",
        "Plinth and depression filling under floors; watering to optimum moisture content",
        "Labour",
    ),
}


def get_teaching_row_data(
    item_code: str,
    res_code: str,
    raw_name: str,
    section_name: str,
    coeff: float,
    unit: str,
    batch_qty: float,
    batch_unit: str,
    hours: float | None = None,
) -> tuple[str, str, str, str, str, str]:
    """Return (disp_name, work_done, condition, category, productivity, quantity)
    with 100% mathematical correlation between operating hours, output rates, and batch quantities.
    """
    # 1. Determine Identity & Descriptions
    clean_code = res_code.strip()
    override = _ITEM_SPECIFIC_OVERRIDES.get((item_code, clean_code))
    if override:
        disp_name, work_done, cond, cat = override
    else:
        # Check default roles
        def_info = _DEFAULT_ROLES.get(clean_code)
        if not def_info:
            # Check by raw name keywords
            name_lower = raw_name.lower()
            if "blasting powder" in name_lower:
                def_info = _DEFAULT_ROLES.get("EXPLOSIVE-POWDER")
            elif "fuse" in name_lower:
                def_info = _DEFAULT_ROLES.get("EXPLOSIVE-FUSE")
            elif "kerosene" in name_lower:
                def_info = _DEFAULT_ROLES.get("KEROSENE")
            elif "fly ash" in name_lower:
                def_info = _DEFAULT_ROLES.get("1980")
            elif "sand" in name_lower:
                def_info = _DEFAULT_ROLES.get("6501")
            elif "chlorpyriphos" in name_lower or "chlorpyrifos" in name_lower:
                def_info = _DEFAULT_ROLES.get("7022")
            elif "plank" in name_lower:
                def_info = _DEFAULT_ROLES.get("1198")
            elif "scantling" in name_lower:
                def_info = _DEFAULT_ROLES.get("1197")
            elif "balli" in name_lower or "pole" in name_lower:
                def_info = _DEFAULT_ROLES.get("0302")

        if def_info:
            w_done, c_used, default_cat = def_info
            disp_name = raw_name
            work_done = w_done
            cond = c_used
            if section_name == "MATERIAL":
                cat = "Material"
            elif section_name == "MACHINERY":
                cat = "Machine"
            elif section_name == "CARRIAGE":
                cat = "Carriage"
            else:
                cat = default_cat
        else:
            disp_name = raw_name
            if section_name == "MATERIAL":
                cat = "Material"
                work_done = f"Direct material supply for {item_code}"
                cond = f"CPWD specification for {raw_name.lower()}"
            elif section_name == "MACHINERY":
                cat = "Machine"
                work_done = f"Mechanical operation of {raw_name.lower()} for {item_code}"
                cond = f"Mechanized execution under {section_name.lower()}"
            elif section_name == "CARRIAGE":
                cat = "Carriage"
                work_done = f"Carriage of materials for {item_code}"
                cond = f"CPWD carriage lead specification"
            elif section_name == "LABOUR":
                cat = "Labour"
                work_done = f"Manual execution of {raw_name.lower()} tasks for {item_code}"
                cond = f"Standard manual labour under {section_name.lower()}"
            else:
                cat = "Equipment"
                work_done = f"Execution of {raw_name.lower()} tasks for {item_code}"
                cond = f"Standard CPWD specification"

    # 2. Calculate Hours & Mathematically Correlated Rates
    if hours is not None:
        tot_hrs = abs(hours)
    elif coeff != 0 and cat in ("Labour", "Machine", "Equipment"):
        tot_hrs = round(abs(coeff) * SHIFT_HOURS, 3)
    else:
        tot_hrs = 0.0

    # Format Quantity
    if batch_qty == int(batch_qty):
        qty_str = f"{int(batch_qty)} {batch_unit}"
    else:
        qty_str = f"{batch_qty:.2f} {batch_unit}"

    # Format Productivity with Mathematical Correlation
    if cat == "Labour":
        if tot_hrs > 0 and batch_qty > 0:
            rate_val = batch_qty / tot_hrs
            prod_str = f"{tot_hrs:.2f} man-hrs ({rate_val:.2f} {batch_unit}/man-hr)"
        elif tot_hrs >= 1.0:
            prod_str = f"{tot_hrs:.2f} man-hrs"
        else:
            prod_str = f"{tot_hrs:.3f} man-hrs"

    elif cat in ("Machine", "Equipment"):
        if tot_hrs > 0 and batch_qty > 0:
            rate_val = batch_qty / tot_hrs
            rate_fmt = f"{rate_val:.2f}" if rate_val != int(rate_val) else f"{int(rate_val)}"
            hrs_fmt = f"{tot_hrs:g}" if tot_hrs >= 0.1 else f"{tot_hrs:.3f}"
            prod_str = f"{hrs_fmt} machine-hrs ({rate_fmt} {batch_unit}/hr)"
            # Synchronize condition if generic output was mentioned
            if "output ~" in cond or "rated output" in cond:
                cond = re.sub(r"(?:rated output|output)\s*~?\s*[\d.]+\s*cum/hour", f"rated output: {rate_fmt} {batch_unit}/hr", cond)
        else:
            prod_str = f"{tot_hrs:g} machine-hrs"

    elif cat == "Material":
        if batch_qty > 0 and coeff > 0:
            unit_rate = coeff / batch_qty
            if unit_rate == int(unit_rate):
                rate_part = f"{int(unit_rate)} {unit}/{batch_unit}"
            else:
                rate_part = f"{unit_rate:.4g} {unit}/{batch_unit}"
            prod_str = f"{coeff:g} {unit} ({rate_part})"
        else:
            prod_str = f"{coeff:g} {unit}"

    elif cat == "Reference":
        if batch_qty > 0 and coeff > 0:
            unit_rate = coeff / batch_qty
            prod_str = f"{coeff:g} {unit} ({unit_rate:.3f} {unit}/{batch_unit})"
        else:
            prod_str = f"{coeff:g} {unit}"

    else:
        prod_str = f"{coeff:g} {unit}"

    return disp_name, work_done, cond, cat, prod_str, qty_str
