# -*- coding: utf-8 -*-
"""
Per-trade input analysis: which clause of an item's nomenclature actually
moves cost, and whether the DAR 2019 data supports it.

Each entry answers, for one clause, the question "is there a cost impact that
I have data for?" with one of three verdicts:

  DRIVES COST         the DAR prints different figures for the listed choices
  NOMENCLATURE ONLY   the DAR prices one rate for every variant of this clause
  USER-SUPPLIED COST  real cost, no DAR figure - the user enters an amount

`evidence` cites the DAR items that justify the verdict, so the classification
can be checked rather than taken on trust.

This module also carries the cross-volume MATERIAL lines that sheets 08, 09
and 10 import from `Resolved_Cross_Volume_Items`, each tagged "A" so it is
excluded from the markup base per the CPWD (W-A) convention.
"""

import re

from scripts.trade_layout import R_SAY, R_LIB_FIRST

# The 03_Mortars builder computes ONE mix at a time. Consuming trades read its
# live Say rate; a second mortar (e.g. Cladding's white-cement pointing mix) is
# built there too and logged in the item library, and is read from that row.
# Both are derived from the layout constants so a row shift cannot silently
# repoint them at a blank or, worse, at a zeroed markup cell.
MORTAR_LIVE_SAY = "='03_Mortars'!G%d" % R_SAY
MORTAR_LIBRARY_2 = "='03_Mortars'!H%d" % (R_LIB_FIRST + 1)

DRIVES = 'DRIVES COST'
NOMEN = 'NOMENCLATURE ONLY'
USER = 'USER-SUPPLIED COST'


def _d(label, default, impact, drives, evidence, options=None):
    return {'label': label, 'default': default, 'impact': impact,
            'drives': drives, 'evidence': evidence, 'options': options}


def _k(label, default, unit, applies_to, note, number_format=None, options=None):
    return {'label': label, 'default': default, 'unit': unit, 'applies_to': applies_to,
            'note': note, 'number_format': number_format, 'options': options}


# ---------------------------------------------------------------------------
# 02 - Earth Work
# ---------------------------------------------------------------------------
EARTH = {
    'scope_inputs': [
        _d('Strata / soil classification', 'All kinds of soil', DRIVES, 'Labour rows',
           'DAR 2.7.1 (ordinary rock), 2.7.2 (hard rock, blasting), 2.7.3 (hard rock, blasting prohibited) '
           'are priced separately from 2.1.1 (all kinds of soil) with different crews and machinery - so the '
           'strata is a real cost driver with data behind it.',
           ['All kinds of soil', 'Ordinary rock', 'Hard rock (requiring blasting)',
            'Hard rock (blasting prohibited)']),
        _d('Operation', 'Surface excavation', DRIVES, 'Labour rows',
           'DAR 2.1 (surface excavation), 2.2 (rough excavation and banking), 2.3 (trench excavation), '
           '2.6 (excavation by hydraulic excavator) each carry different Beldar/Coolie coefficients.',
           ['Surface excavation', 'Rough excavation and banking', 'Trench / foundation excavation',
            'Filling in layers with watering and ramming', 'Excavation by hydraulic excavator']),
        _d('Depth / thickness band', 'not exceeding 30 cm in depth', DRIVES, 'Output basis',
           'The depth band changes the output basis the crew is quoted against (100 sqm for surface work '
           'not exceeding 30 cm, 10 cum for deeper excavation) - see DAR 2.1.1 vs 2.2.1.',
           ['not exceeding 30 cm in depth', 'exceeding 30 cm but not exceeding 1.5 m depth',
            'exceeding 1.5 m depth']),
        _d('Lead for disposal', 'upto 50 m lead', NOMEN, '-',
           'Within an Earth Work item the DAR quotes one rate for disposal upto 50 m lead; beyond 50 m the '
           'carriage is measured separately under Sub-Head 01, not by varying this item. No differential '
           'figure exists inside Sub-Head 02.',
           ['upto 50 m lead', 'lead beyond 50 m (price carriage separately under Sub-Head 01)']),
        _d('Lift', 'and 1.5 m lift', NOMEN, '-',
           'The DAR prices one rate for the stated lift; "all lifts" and "1.5 m lift" wordings do not carry '
           'separate figures in Sub-Head 02.',
           ['and 1.5 m lift', 'for all lifts']),
        _d('Mechanical plant used', 'Manual (no plant)', DRIVES, 'Labour rows',
           'DAR 2.6.1 prices excavation by hydraulic excavator with a machinery line and a much smaller '
           'labour gang than the manual item 2.1.1 - selecting plant should replace crew rows with a '
           'plant hire row from Rates_Master.',
           ['Manual (no plant)', 'Hydraulic excavator', 'Power roller 8 tonne (compaction)',
            'Excavator + tipper combination']),
    ],
    'driver_inputs': [
        _k('Output basis quantity', 100.0, 'sqm or cum', 'Batch basis',
           'DAR Earth Work items are quoted per 100 sqm (surface work) or per 10 cum (volume work). Keep '
           'this consistent with the coefficients entered below.', '0.00'),
        _k('Number of watering passes', 1, 'passes', 'Bhisti coeff',
           'DAR 2.5 is an explicit deduct item for NOT watering the excavated earth for banking - so watering '
           'is a costed operation with a printed figure.', '0'),
        _k('Rolling with power roller', 'YES', 'YES / NO', 'Plant row',
           'DAR 2.4 is an explicit deduct item for not rolling with a minimum 8 tonne power roller, which '
           'gives the differential directly.', None, ['YES', 'NO']),
        _k('Site clearance / dressing included', 'NO', 'YES / NO', 'Labour rows',
           'Dressing is priced in its own DAR items rather than as an uplift on excavation; if included, add '
           'the dressing crew as extra labour rows.', None, ['YES', 'NO']),
    ],
}

# ---------------------------------------------------------------------------
# 03 - Mortars
# ---------------------------------------------------------------------------
MORTARS = {
    'scope_inputs': [
        _d('Binder type', 'Ordinary Portland cement', DRIVES, 'Material row 1',
           'DAR 3.1-3.14 use ordinary cement (code 0367); 3.15-3.17 use white cement at a different rate. '
           'Changing the binder changes the material code and rate.',
           ['Ordinary Portland cement', 'White cement', 'Lime', 'Mud mortar']),
        _d('Mix ratio (binder : aggregate)', '1:4', DRIVES, 'Material coeffs',
           'DAR 3.1 (1:1) through 3.11 (1:6) print a different cement volume for every ratio - 0.7175 cum '
           'of cement at 1:1 down to 0.178 cum at 1:6. This is the strongest cost driver on the sheet.',
           ['1:1', '1:2', '1:3', '1:4', '1:5', '1:6', '1:8']),
        _d('Fine aggregate type', 'coarse sand', DRIVES, 'Material row 3',
           'DAR splits the same ratios by aggregate: 3.1-3.6 fine sand, 3.7-3.11 coarse sand, 3.12 stone '
           'dust, 3.13-3.14 marble dust - each a different Rates_Master code and rate.',
           ['fine sand', 'coarse sand', 'stone dust', 'marble dust']),
        _d('Mixing method', 'Manual (hand mixing)', DRIVES, 'Labour rows',
           'Hand mixing is priced with Beldar days; mechanical mixing adds a plant hire line (code 0002 '
           'onwards) whose day rate already includes operator and fuel per CPWD Note 1.',
           ['Manual (hand mixing)', 'Mechanical mixer']),
        _d('Location / lift of use', 'at ground level', NOMEN, '-',
           'The Mortars sub-head prices the mix itself, not its placement. Lifting to upper floors is priced '
           'in the consuming trade (see the "extra labour for lifting material" lines inside DAR 4.2.5), not '
           'here.',
           ['at ground level', 'for all floors / lifts']),
        _d('Purpose of the mix', 'for masonry bedding', NOMEN, '-',
           'The DAR does not vary a mortar rate by what it will be used for; the same item 3.9 is imported '
           'by Masonry, Stone, Cladding and Flooring alike.',
           ['for masonry bedding', 'for plaster', 'for pointing', 'for flooring bedding']),
    ],
    'driver_inputs': [
        _k('Dry volume factor', 1.30, 'ratio', 'Material coeffs',
           'The DAR derives wet-to-dry mortar volume inside each item header (e.g. 3.9 "0.268 cum of cement"). '
           'Adjust only if you are deriving a ratio the book does not print.', '0.000'),
        _k('Cement bulk density', 1.44, 'tonne / cum', 'Cement coeff',
           'Used to convert the printed cement volume into the tonne basis that code 0367 is rated in - the '
           'conversion the DAR performs inline (0.269 cum = 0.38 tonne).', '0.000'),
        _k('Batch output basis', 1.00, 'cum', 'Batch basis',
           'Every DAR mortar item is analysed for 1.00 cum. Keep 1.00 unless you deliberately want a batch '
           'rate.', '0.00'),
        _k('Wastage allowance', 0.00, '%', 'Material coeffs',
           'The DAR mortar items carry no wastage percentage - wastage on mortar is absorbed in the consuming '
           'trade. Leave at 0 to stay faithful to the book.', '0.00'),
    ],
}

# ---------------------------------------------------------------------------
# 04 / 05 - Concrete and RCC
# ---------------------------------------------------------------------------
def _concrete_scope(is_rcc):
    element = (['in plinth beams', 'in columns', 'in suspended slabs', 'in walls', 'in footings']
               if is_rcc else
               ['in foundation and plinth', 'in superstructure', 'in retaining walls', 'as levelling course'])
    return [
        _d('Mix / grade', '1:2:4', DRIVES, 'Material coeffs',
           'DAR 4.1.2 (1:1.5:3) through 4.1.8 (1:4:8) print different cement, sand and aggregate volumes for '
           'each grade, giving Say rates from about Rs 8,025 to well over Rs 10,000 per cum.',
           ['1:1.5:3', '1:2:4', '1:3:6', '1:4:8', '1:5:10', 'M20 design mix', 'M25 design mix']),
        _d('Coarse aggregate nominal size', '20 mm nominal size', DRIVES, 'Material rows 1-2',
           'DAR 4.1.3 (20 mm) and 4.1.4 (40 mm) are the same 1:2:4 grade but use different aggregate codes '
           'and a different voids deduction, so the two print different rates.',
           ['20 mm nominal size', '40 mm nominal size', '10 mm nominal size', 'graded 20 mm and 10 mm']),
        _d('Placement location in the structure',
           'in plinth beams' if is_rcc else 'in foundation and plinth', NOMEN, '-',
           'The DAR does not print a separate concrete rate per structural element; what changes with height '
           'is the lifting labour, which the book adds as an explicit extra Coolie line (see 4.2.5, '
           '"extra labour for lifting material upto floor V level"). Add that as a labour row rather than '
           'expecting a different base rate.',
           element),
        _d('Centering and shuttering', 'excluding centering and shuttering', NOMEN, '-',
           'Every DAR concrete and RCC item is analysed excluding the cost of centering and shuttering, which '
           'is measured under its own sub-head. There is no in-item differential.',
           ['excluding centering and shuttering', 'including centering and shuttering (price separately)']),
        _d('Reinforcement', 'excluding reinforcement' if is_rcc else 'not applicable (plain concrete)',
           NOMEN, '-',
           'RCC items in the DAR exclude reinforcement, which is measured and priced separately under Steel '
           'Work. No in-item figure exists.',
           ['excluding reinforcement', 'including reinforcement (price separately)']),
        _d('Finish to exposed surface', 'no special finish', USER, 'Sundries row',
           'Float finish, hardener topping and similar finishes are priced in Flooring items (11.3-11.5), not '
           'inside the concrete item. If required here, enter the allowance as a Sundries line.',
           ['no special finish', 'float finish', 'trowel finish', 'concrete hardener topping']),
    ]


def _concrete_drivers(is_rcc):
    return [
        _k('Voids deduction in aggregate', 0.075 if not is_rcc else 0.075, 'fraction', 'Aggregate coeff',
           'DAR 4.1.4 deducts 7.5% for voids inline ("0.56 - 7.5% for voids = 0.52 cum"). Adjust only with a '
           'documented reason.', '0.000'),
        _k('Extra lift labour (floor level)', 0.00, 'Coolie days', 'Labour rows',
           'The DAR adds this as its own Coolie line - 4.2.5 uses 0.75 x 2.5 = 1.88 Coolie days for lifting '
           'material up to floor V level. Enter the days here and mirror them on a Coolie row.', '0.00'),
        _k('Concrete mixer productivity', 0.07, 'days / cum', 'Plant row',
           'DAR 4.2.5 uses 0.07 mixer-days per cum. Plant codes 0001-0083 already include operator, fuel and '
           'lubricants for an 8-hour shift per CPWD Note 1.', '0.000'),
        _k('Vibrator productivity', 0.07, 'days / cum', 'Plant row',
           'DAR 4.2.5 uses 0.07 needle-vibrator days per cum on the same basis.', '0.000'),
    ]


CONCRETE = {'scope_inputs': _concrete_scope(False), 'driver_inputs': _concrete_drivers(False)}
RCC = {'scope_inputs': _concrete_scope(True), 'driver_inputs': _concrete_drivers(True)}

# ---------------------------------------------------------------------------
# 06 - Masonry
# ---------------------------------------------------------------------------
MASONRY = {
    'scope_inputs': [
        _d('Masonry unit type and class', 'common burnt clay F.P.S. (non modular) bricks class 7.5',
           DRIVES, 'Material row 1',
           'DAR 6.1 (F.P.S. non-modular 7.5), 6.2 (modular 7.5) and 6.3 (machine moulded perforated 12.5) use '
           'different brick codes and different bricks-per-cum coefficients - 6.1.1 and 6.2.1 differ by about '
           'Rs 460 per cum at the same mortar.',
           ['common burnt clay F.P.S. (non modular) bricks class 7.5',
            'common burnt clay modular bricks class 7.5',
            'machine moulded perforated bricks class 12.5',
            'fly ash lime bricks', 'concrete solid blocks', 'AAC blocks']),
        _d('Mortar ratio', 'cement mortar 1:4 (1 cement : 4 coarse sand)', DRIVES,
           'Mortar import row',
           'DAR 6.1.1 imports Mortar item 3.9 (1:4) and 6.1.2 imports 3.11 (1:6); the printed Say rates are '
           'Rs 6,376.25 and Rs 6,153.05 per cum. Change this and the imported mortar rate must change with it.',
           ['cement mortar 1:3 (1 cement : 3 coarse sand)',
            'cement mortar 1:4 (1 cement : 4 coarse sand)',
            'cement mortar 1:5 (1 cement : 5 coarse sand)',
            'cement mortar 1:6 (1 cement : 6 coarse sand)',
            'lime mortar', 'mud mortar']),
        _d('Location in structure', 'in foundation and plinth', NOMEN, '-',
           'DAR 6.1 is titled "in foundation and plinth" and prices one rate for it. Superstructure masonry is '
           'a separate printed item whose difference is scaffolding labour - add that as a labour or sundries '
           'line rather than expecting this clause to move the rate on its own.',
           ['in foundation and plinth', 'in superstructure upto floor V level',
            'in superstructure above floor V level']),
        _d('Wall thickness', 'one brick thick and above', NOMEN, '-',
           'The DAR prices brickwork by volume (per cum) for one-brick-and-above walls, so thickness does not '
           'change the cum rate. Half-brick walls are a separate per-sqm item.',
           ['one brick thick and above', 'half brick thick (separate per-sqm item)']),
        _d('Curing', 'including curing', NOMEN, '-',
           'Curing is inside the Bhisti coefficient and the 1% water charge; the DAR gives no separate figure '
           'for including or excluding it.',
           ['including curing', 'excluding curing']),
        _d('Raking of joints for plaster', 'not required', USER, 'Sundries row',
           'The DAR carries no separate raking allowance in Sub-Head 06. If the specification demands it, put '
           'the cost in the Sundries line.',
           ['not required', 'joints raked for plaster']),
    ],
    'driver_inputs': [
        _k('Bricks per cum of masonry', 0.494, 'thousand / cum', 'Material row 1',
           'DAR 6.1.1 uses 0.494 thousand bricks per cum, which already includes the mortar joint volume.',
           '0.000'),
        _k('Mortar consumption', 0.25, 'cum / cum', 'Mortar import row',
           'DAR 6.1.1 uses 0.25 cum of mortar per cum of brickwork.', '0.000'),
        _k('Brick breakage allowance', 0.00, '%', 'Material row 1',
           'Sub-Head 06 items carry no separate breakage percentage - it is absorbed in the 0.494 coefficient. '
           'Leave at 0 to stay faithful to the book.', '0.00'),
        _k('Scaffolding allowance', 0.00, 'Rs L.S.', 'Sundries row',
           'The DAR prices scaffolding as an explicit L.S. line where it applies (see 4.2.5, "Scaffolding '
           '17.60 x 2.5"). Enter an amount here and mirror it on the Sundries row.', '0.00'),
    ],
}

# ---------------------------------------------------------------------------
# 07 - Stone Work
# ---------------------------------------------------------------------------
STONE = {
    'scope_inputs': [
        _d('Stone masonry type', 'random rubble masonry', DRIVES, 'Labour rows',
           'The DAR prices random rubble, coursed rubble and ashlar as separate items with progressively '
           'higher mason coefficients for the extra dressing and coursing work.',
           ['random rubble masonry', 'coursed rubble masonry (first sort)',
            'coursed rubble masonry (second sort)', 'ashlar masonry']),
        _d('Stone type', 'stone for masonry (rubble)', DRIVES, 'Material row 1',
           'Different quarry stones carry different Rates_Master codes and rates; the stone line is the '
           'largest single component of a stone masonry item.',
           ['stone for masonry (rubble)', 'sand stone', 'granite', 'quartzite', 'lime stone']),
        _d('Mortar ratio', 'cement mortar 1:6 (1 cement : 6 coarse sand)', DRIVES, 'Mortar import row',
           'As in Masonry, the mortar is imported from Sub-Head 03 at the printed ratio; changing the ratio '
           'changes the imported rate (item 3.9 vs 3.11 and so on).',
           ['cement mortar 1:3 (1 cement : 3 coarse sand)',
            'cement mortar 1:4 (1 cement : 4 coarse sand)',
            'cement mortar 1:6 (1 cement : 6 coarse sand)', 'lime mortar']),
        _d('Face dressing', 'hammer dressed', DRIVES, 'Labour rows',
           'Dressing grade is delivered through the stone mason coefficient - a finer dressed face takes more '
           'mason days for the same volume, which is how the DAR differentiates its stone items.',
           ['undressed', 'hammer dressed', 'chisel dressed', 'fine dressed / polished']),
        _d('Through / bond stones', 'including through stones', DRIVES, 'Material row 3',
           'The DAR carries through or bond stone (code 0300) as its own material line inside stone masonry '
           'items, so including or omitting it is a costed choice.',
           ['including through stones', 'without through stones']),
        _d('Backing concrete', 'not required', USER, 'Sundries row',
           'Some DAR stone items carry a "Cement concrete 1:6:12" style L.S. backing line. There is no single '
           'printed figure that applies generally, so enter the allowance yourself.',
           ['not required', 'cement concrete 1:6:12 backing', 'cement concrete 1:5:10 backing']),
    ],
    'driver_inputs': [
        _k('Stone consumption', 1.05, 'cum / cum', 'Material row 1',
           'Quarry stone is measured loose against finished masonry volume; the DAR builds the difference '
           'into the printed coefficient.', '0.000'),
        _k('Mortar consumption', 0.34, 'cum / cum', 'Mortar import row',
           'Rubble masonry takes a larger mortar fraction than brickwork because of the irregular voids.',
           '0.000'),
        _k('Through stone spacing', 1.50, 'm', 'Material row 3',
           'Through stones are specified at a spacing, which converts to the bond stone quantity on the '
           'material row.', '0.00'),
        _k('Scaffolding allowance', 0.00, 'Rs L.S.', 'Sundries row',
           'Enter an amount if the work is above scaffolding height, and mirror it on the Sundries row.',
           '0.00'),
    ],
}

# ---------------------------------------------------------------------------
# 08 - Cladding Work
# ---------------------------------------------------------------------------
CLADDING = {
    'scope_inputs': [
        _d('Slab area band', 'area of slab over 0.50 sqm', DRIVES, 'Batch basis',
           'This is the dominant cost dimension in Sub-Head 08. DAR 8.1.1.1 (slab area upto 0.50 sqm) is '
           'analysed for 0.50 sqm and 8.1.1.2 (over 0.50 sqm) for 1.00 sqm, and the two print materially '
           'different rates for the same stone.',
           ['area of slab upto 0.50 sqm', 'area of slab over 0.50 sqm']),
        _d('Stone / marble type', 'Raj Nagar plain white marble', DRIVES, 'Material row 1',
           'DAR 8.1.1 and 8.2.1 list the marble variety explicitly (Raj Nagar plain white / Udaipur green / '
           'Zebra black), each a different Rates_Master code and rate.',
           ['Raj Nagar plain white marble', 'Udaipur green marble', 'Zebra black marble',
            'red sand stone', 'Kota stone', 'granite']),
        _d('Bedding mortar ratio', 'cement mortar 1:3 (1 cement : 3 coarse sand)', DRIVES,
           'Bedding mortar row',
           'DAR 8.1.1.1 imports Mortar item 3.8 (1:3 coarse sand) as the bedding mortar line.',
           ['cement mortar 1:3 (1 cement : 3 coarse sand)',
            'cement mortar 1:4 (1 cement : 4 coarse sand)',
            'cement mortar 1:6 (1 cement : 6 coarse sand)']),
        _d('Pointing mortar', 'white cement mortar 1:2 (1 white cement : 2 marble dust)', DRIVES,
           'Pointing mortar row',
           'DAR 8.1.1.1 imports a SECOND mortar - item 3.15, white cement 1:2 with marble dust - for pointing, '
           'alongside the bedding mortar. Both lines are priced in the same item.',
           ['white cement mortar 1:2 (1 white cement : 2 marble dust)',
            'white cement mortar 1:3 (1 white cement : 3 marble dust)',
            'cement mortar 1:3 (1 cement : 3 coarse sand)', 'no pointing']),
        _d('Slab thickness and finish', '18 mm gang saw cut, mirror polished', DRIVES, 'Material row 1',
           'DAR 8.1 (gang saw cut, polished and machine cut) and 8.2 (premoulded and prepolished) are separate '
           'items at different rates for the same 18 mm thickness.',
           ['18 mm gang saw cut, mirror polished', '18 mm premoulded and prepolished',
            '20 mm thick', '25 mm thick', '30 mm thick']),
        _d('Chase cutting for concealed fixings', 'not required', DRIVES, 'Resolved 18.78 row',
           'Where the specification requires chases in the wall, the rate is DAR item 18.78 (Water Supply), '
           'resolved on the Resolved_Cross_Volume_Items sheet at Rs 154.15 per metre and imported here as an '
           'A-tagged line.',
           ['not required', 'chases upto 7.5 x 7.5 cm in walls (DAR 18.78)']),
    ],
    'driver_inputs': [
        _k('Cutting and wastage allowance', 5.00, '%', 'Material row 1',
           'DAR 8.1.1.1 states "Add for wastage" inline on the stone line. 5% is the common printed figure; '
           'change it if your slab schedule justifies more.', '0.00'),
        _k('Bedding mortar thickness', 12.00, 'mm', 'Bedding mortar row',
           'Converts to the bedding mortar volume per sqm of cladding.', '0.00'),
        _k('Chase length per sqm', 0.00, 'metre / sqm', 'Resolved 18.78 row',
           'Multiply by the resolved 18.78 rate. Leave at 0 when no chase cutting is required.', '0.000'),
        _k('Scaffolding allowance', 0.00, 'Rs L.S.', 'Sundries row',
           'Cladding above scaffolding height needs an access allowance; the DAR carries it as an L.S. line '
           'rather than a percentage.', '0.00'),
    ],
}

# ---------------------------------------------------------------------------
# 09 - Wood and PVC Work
# ---------------------------------------------------------------------------
WOOD = {
    'scope_inputs': [
        _d('Timber species', 'second class teak wood', DRIVES, 'Material row 1',
           'The single biggest driver in Sub-Head 09. Same chowkhat, three species, three printed rates: '
           'DAR 9.1.1 second class teak, 9.1.2 sal wood, 9.1.3 kiln seasoned hollock - the GST lines alone '
           'show Rs 497.07 / 402.17 / 273.95, so the finished rates differ by nearly half.',
           ['second class teak wood', 'sal wood', 'kiln seasoned and chemically treated hollock wood',
            'laminated veneer lumber (LVL)', 'PVC / plastic section']),
        _d('Member type', 'chowkhat / frame of door', DRIVES, 'Material coeffs',
           'DAR 9.1 (door and window frames), 9.2 (LVL factory made frames) and 9.3 (false ceiling and '
           'partition frames) are analysed against different member sizes and therefore different timber '
           'volumes.',
           ['chowkhat / frame of door', 'window frame', 'clerestory window frame',
            'false ceiling / partition frame', 'shutter', 'beading / architrave']),
        _d('Wrought and framed work', 'wrought framed and fixed in position', DRIVES, 'Labour rows',
           'DAR 9.1 is priced "wrought framed and fixed in position"; 9.3 is "sawn and fixed in position" '
           'with a lower carpenter coefficient, so the workmanship clause carries a real labour differential.',
           ['wrought framed and fixed in position', 'sawn and fixed in position']),
        _d('Circular / shaped work', 'straight work only', DRIVES, 'Labour rows',
           'DAR 9.4 is an explicit EXTRA item for additional labour on circular work such as fan light frames '
           '- it gives the differential directly rather than leaving it to judgement.',
           ['straight work only', 'circular work (DAR 9.4 extra labour applies)']),
        _d('Priming coat on wood', 'priming coat applied (DAR 13.50.1)', DRIVES, 'Resolved 13.50.1 row',
           'Priming on wood work is DAR item 13.50.1, resolved on Resolved_Cross_Volume_Items at Rs 57.05 per '
           'sqm and imported here as an A-tagged line so it is not marked up twice.',
           ['priming coat applied (DAR 13.50.1)', 'no priming coat']),
        _d('Wood preservative treatment', 'not required', DRIVES, 'Resolved 13.57.1 row',
           'Oil type wood preservative, two or more coats on new work, is DAR item 13.57.1, resolved at '
           'Rs 44.50 per sqm and available as an A-tagged import.',
           ['not required', 'oil type wood preservative, two or more coats (DAR 13.57.1)']),
    ],
    'driver_inputs': [
        _k('Timber wastage allowance', 5.00, '%', 'Material row 1',
           'DAR 9.1.1 works in cudm and adds a wastage allowance inline when converting the net section volume '
           'to the ordered volume (36 cudm net becomes 38 cudm).', '0.00'),
        _k('Surface area for priming', 3.00, 'sqm / unit', 'Resolved 13.50.1 row',
           'The priming import is priced per sqm of finished timber surface; enter the girth-derived area for '
           'the member being built.', '0.000'),
        _k('Output unit conversion', 0.0380, 'cum / unit', 'Rate row',
           'DAR 9 items are analysed per door / per frame and then converted to a standard cum rate at the end '
           'of the item. Enter the timber volume of one unit to convert between the two.', '0.0000'),
        _k('Holdfasts per frame', 6, 'Nos', 'Material rows',
           'Holdfasts and fixing hardware are separate material lines in the DAR frame items.', '0'),
    ],
}

# ---------------------------------------------------------------------------
# 10 - Steel Work
# ---------------------------------------------------------------------------
STEEL = {
    'scope_inputs': [
        _d('Fabrication type', 'single section', DRIVES, 'Material coeff',
           'DAR 10.1 (single section, fixed with or without connecting plate) uses 1.05 quintal of structurals '
           'per quintal of finished work; 10.2 (riveted, bolted or welded built-up sections and trusses) uses '
           '1.60 - a 52% difference on the dominant material line.',
           ['single section', 'built-up section, riveted / bolted / welded', 'truss / lattice girder',
            'gusset plated connection']),
        _d('Section profile', 'tees, angles, channels and R.S. joists', DRIVES, 'Material row 1',
           'DAR 10.1 uses code 1007 (structurals: tees, angles, channels, joists); 10.2 additionally uses code '
           '1009 (flats exceeding 10 mm) at a different rate for ties and braces.',
           ['tees, angles, channels and R.S. joists', 'flats exceeding 10 mm in thickness',
            'flats upto 10 mm in thickness', 'plates / gussets', 'hollow sections']),
        _d('Priming coat', 'red oxide zinc chromate primer (DAR 13.50.3)', DRIVES, 'Resolved 13.50.3 row',
           'DAR 10.1 imports Finishing item 13.50.3 at Rs 50.70 per sqm and tags it "A", excluding it from '
           'every markup base. It is resolved here on Resolved_Cross_Volume_Items rather than held as a '
           'literal.',
           ['red oxide zinc chromate primer (DAR 13.50.3)', 'ready mixed pink / grey primer (DAR 13.50.1)',
            'no priming coat']),
        _d('Connecting plate', 'with or without connecting plate', NOMEN, '-',
           'DAR 10.1 is titled "fixed with or without connecting plate" - the book deliberately prices one '
           'rate for both, so this clause carries no differential.',
           ['with or without connecting plate', 'with connecting plate', 'without connecting plate']),
        _d('Erection height / lift', 'for all heights', NOMEN, '-',
           'Sub-Head 10 items do not vary by erection height; where extra lifting is genuinely incurred the '
           'DAR adds an explicit labour line, as it does in Concrete Work.',
           ['for all heights', 'upto floor V level', 'above floor V level']),
        _d('Galvanising / protective coating', 'not required', USER, 'Sundries row',
           'Hot dip galvanising is not analysed inside Sub-Head 10 in Vol.1. If specified, enter the amount as '
           'a Sundries line or add a material row for the coating.',
           ['not required', 'hot dip galvanised', 'epoxy coated']),
    ],
    'driver_inputs': [
        _k('Steel wastage allowance', 5.00, '%', 'Material row 1',
           'DAR 10.2 shows the arithmetic openly: "Total = 50.65 kg + Add wastage @ 5% = 2.53 kg". The 1.05 '
           'coefficient on 10.1 is the same 5% expressed as a multiplier.', '0.00'),
        _k('Painted surface area', 3.00, 'sqm / quintal', 'Resolved 13.50.3 row',
           'DAR 10.1 prices 3.00 sqm of priming per quintal of steel. Change only if your section girth '
           'differs materially.', '0.000'),
        _k('Bolts / rivets allowance', 0.00, '% of steel', 'Material rows',
           'Built-up items carry their fasteners as separate material lines rather than as a percentage; use '
           'this only if you are approximating.', '0.00'),
        _k('Output unit conversion', 100.0, 'kg / quintal', 'Rate row',
           'DAR 10.1 is analysed per quintal and then reported per kg ("Cost of 1 quintal 8604.22, Cost of '
           '1 Kg 86.04"). This is the divisor for that conversion.', '0.00'),
    ],
}

# ---------------------------------------------------------------------------
# 11 - Flooring
# ---------------------------------------------------------------------------
FLOORING = {
    'scope_inputs': [
        _d('Flooring material', 'precast terrazzo tiles', DRIVES, 'Material row 1',
           'DAR 11.1 (brick on edge), 11.2 (dry brick on edge on mud mortar), 11.3 (cement concrete flooring) '
           'and the tile items are separate analyses with entirely different material lines.',
           ['brick on edge (class 7.5)', 'precast terrazzo tiles', 'cement concrete flooring',
            'Kota stone slab', 'marble slab', 'ceramic / vitrified tile']),
        _d('Bedding mortar ratio', 'cement mortar 1:4 (1 cement : 4 coarse sand)', DRIVES,
           'Mortar import row',
           'DAR 11.1.1 imports Mortar item 3.9 (1:4) and 11.1.2 imports 3.11 (1:6) for the identical brick '
           'flooring; 11.2 imports 3.18 (mud mortar) instead. The bed is a live cost lever.',
           ['cement mortar 1:3 (1 cement : 3 coarse sand)',
            'cement mortar 1:4 (1 cement : 4 coarse sand)',
            'cement mortar 1:6 (1 cement : 6 coarse sand)', 'mud mortar (DAR 3.18)']),
        _d('Flooring thickness', '40 mm thick', DRIVES, 'Material coeffs',
           'DAR 11.3.1 (40 mm), 11.4 (52 mm with hardener topping) and 11.5 (62 mm with hardener topping) are '
           'separate printed items with rising material quantities.',
           ['20 mm thick', '40 mm thick', '52 mm thick', '62 mm thick']),
        _d('Surface finish', 'finished with a floating coat of neat cement', DRIVES, 'Material / labour rows',
           'DAR 11.3 is finished "with a floating coat of neat cement" - a costed cement line - while 11.4 and '
           '11.5 carry a concrete hardener topping instead, at a different rate.',
           ['finished with a floating coat of neat cement', 'concrete hardener topping',
            'machine rubbed and polished', 'no special finish']),
        _d('Bedding thickness', 'on a bed of 12 mm cement mortar', DRIVES, 'Mortar import row',
           'DAR 11.1 states the bed thickness in the item title ("on a bed of 12 mm cement mortar"), which '
           'sets the mortar volume per sqm.',
           ['on a bed of 12 mm cement mortar', 'on a bed of 20 mm cement mortar',
            'on a bed of 25 mm cement mortar']),
        _d('Laying pattern', 'in required pattern', NOMEN, '-',
           'DAR 11.2 says "in required pattern" and prices one rate; the book gives no pattern-by-pattern '
           'differential.',
           ['in required pattern', 'herringbone', 'straight joint', 'diagonal']),
    ],
    'driver_inputs': [
        _k('Neat cement slurry', 4.40, 'kg / sqm', 'Cement row',
           'The floating coat of neat cement is priced as a separate cement quantity per sqm in the DAR '
           'flooring items.', '0.00'),
        _k('Bedding mortar consumption', 0.012, 'cum / sqm', 'Mortar import row',
           'A 12 mm bed over 1 sqm is 0.012 cum, which is how the DAR derives the mortar line.', '0.0000'),
        _k('Tile wastage allowance', 5.00, '%', 'Material row 1',
           'Cutting waste at edges and around obstructions; the DAR absorbs it in the printed coefficient, so '
           'set 0 to reproduce the book exactly.', '0.00'),
        _k('Polishing machine productivity', 0.10, 'days / 10 sqm', 'Plant row',
           'Floor rubbing and polishing machine (code 0013) is a day-rated plant line including operator and '
           'fuel per CPWD Note 1.', '0.000'),
    ],
}

# ---------------------------------------------------------------------------
# 12 - Roofing
# ---------------------------------------------------------------------------
ROOFING = {
    'scope_inputs': [
        _d('Sheet thickness', '1.00 mm thick', DRIVES, 'Material row 1',
           'The dominant driver. DAR 12.1.1 (1.00 mm), 12.1.2 (0.80 mm) and 12.1.3 (0.63 mm) derive the sheet '
           'weight explicitly - 17.72 kg per sheet at 0.80 mm against 13.38 kg at 0.63 mm - so the material '
           'line moves with thickness.',
           ['1.00 mm thick', '0.80 mm thick', '0.63 mm thick', '0.50 mm thick']),
        _d('Sheet type', 'corrugated G.S. sheet', DRIVES, 'Material row 1',
           'DAR 12.1 (corrugated G.S. sheet roofing) and 12.4 (plain G.S. sheet ridges and hips) use different '
           'material codes and different fixing labour.',
           ['corrugated G.S. sheet', 'plain G.S. sheet', 'colour coated profile sheet',
            'polycarbonate sheet', 'A.C. sheet']),
        _d('Zinc coating', 'zinc coating not less than 275 gm/sqm', NOMEN, '-',
           'DAR 12.1.1-12.1.3 all state "zinc coating not less than 275 gm/m2" - it is a specification '
           'threshold the book applies to every thickness, not a rate differential.',
           ['zinc coating not less than 275 gm/sqm', 'zinc coating not less than 120 gm/sqm']),
        _d('Fixing hardware', 'polymer coated J or L hooks, bolts and nuts', DRIVES, 'Material rows 3-6',
           'The DAR prices the hooks (code 1023), bolts and nuts (1022), G.I. limpet washers (1207) and '
           'bitumen washers (1208) as four separate costed material lines.',
           ['polymer coated J or L hooks, bolts and nuts', 'G.I. J or L hooks',
            'self drilling screws', 'seam bolts']),
        _d('Surface type', 'vertical / curved surface included', NOMEN, '-',
           'DAR 12.1 is titled "including vertical / curved surface" and prices one rate for both, so the '
           'clause does not carry its own figure.',
           ['vertical / curved surface included', 'flat sloping surface only']),
        _d('Openings and cutting', 'no openings', DRIVES, 'Labour rows',
           'DAR 12.2 (straight cutting) and 12.3 (circular cutting) are explicit EXTRA items for openings '
           'exceeding 40 sq decimetre, each analysed by thickness - so cutting is separately costed data.',
           ['no openings', 'straight cutting for openings (DAR 12.2)',
            'circular cutting for openings (DAR 12.3)']),
    ],
    'driver_inputs': [
        _k('Lap and wastage allowance', 5.00, '%', 'Material row 1',
           'Side and end laps mean more sheet is bought than roof is covered; the DAR builds the lap into its '
           'sheet-count derivation for each roof geometry.', '0.00'),
        _k('Sheet unit weight', 17.72, 'kg / sheet', 'Material row 1',
           'DAR 12.1.2 derives this openly: "2.8 x 0.9 @ 17.72 kg each = 956.8 kg". Change it when you change '
           'the thickness above.', '0.00'),
        _k('Hooks per sheet', 8, 'Nos', 'Material rows 3-4',
           'Fixing hardware is counted per sheet in the DAR derivation rather than per sqm.', '0'),
        _k('Priming and painting of laps', 'YES', 'YES / NO', 'Painter row',
           'DAR 12.1 items carry a Painter line for priming and painting the overlaps; toggle drives whether '
           'that labour row applies.', None, ['YES', 'NO']),
    ],
}

SCOPE_BY_SHEET = {
    '02_Earth_Work': EARTH,
    '03_Mortars': MORTARS,
    '04_Concrete_Work': CONCRETE,
    '05_RCC_Work': RCC,
    '06_Masonry_Work': MASONRY,
    '07_Stone_Work': STONE,
    '08_Cladding_Work': CLADDING,
    '09_Wood_and_PVC_Work': WOOD,
    '10_Steel_Work': STEEL,
    '11_Flooring': FLOORING,
    '12_Roofing': ROOFING,
}

# ---------------------------------------------------------------------------
# Cross-volume MATERIAL lines, imported from Resolved_Cross_Volume_Items.
# Every one is A-tagged: its rate already carries Water, GST, CPOH and Cess.
# ---------------------------------------------------------------------------
RESOLVED_LINES = {
    '08_Cladding_Work': [
        {'code': '3.15', 'coeff': 0.004,
         'custom_desc': 'White cement mortar 1:2 (1 white cement : 2 marble dust) - pointing mortar, '
                        'rate from 03_Mortars or DAR item 3.15',
         'custom_unit': 'cum',
         'custom_rate_formula': MORTAR_LIBRARY_2,
         'note': 'Second mortar line. DAR 8.1.1.1 imports BOTH item 3.8 (bedding) and item 3.15 (white cement '
                 'pointing) in the same cladding item. The 03_Mortars builder computes one mix at a time, so '
                 'the bedding mortar reads its live Say rate while this pointing mortar reads row 2 of '
                 'the 03_Mortars item library - build the pointing mix there and log it. Not A-tagged: '
                 'Mortars carry no markups, so this line takes the full chain.'},
        {'code': '18.78', 'coeff': 0.0,
         'custom_desc': 'Making chases upto 7.5 x 7.5 cm in walls including making good (DAR 18.78, Water Supply)',
         'custom_unit': 'metre',
         'custom_rate_formula': '=XV_18_78',
         'a_tag': True,
         'note': 'Resolved on Resolved_Cross_Volume_Items from WB1 data only (Rs 154.15 per metre). A-tagged: '
                 'the rate already includes Water, GST, CPOH and Cess, so it is excluded from every markup '
                 'base. Set the coefficient from the "Chase length per sqm" driver in Panel 2.'},
    ],
    '09_Wood_and_PVC_Work': [
        {'code': '13.50.1', 'coeff': 3.0,
         'custom_desc': 'Priming coat with ready mixed pink or grey primer on wood work (DAR 13.50.1, Finishing)',
         'custom_unit': 'sqm',
         'custom_rate_formula': '=XV_13_50_1',
         'a_tag': True,
         'note': 'Resolved on Resolved_Cross_Volume_Items from WB1 data only (Rs 57.05 per sqm). A-tagged per '
                 'the CPWD (W-A) convention.'},
        {'code': '13.57.1', 'coeff': 0.0,
         'custom_desc': 'Oil type wood preservative, new work, two or more coats (DAR 13.57.1, Finishing)',
         'custom_unit': 'sqm',
         'custom_rate_formula': '=XV_13_57_1',
         'a_tag': True,
         'note': 'Resolved on Resolved_Cross_Volume_Items from WB1 data only (Rs 44.50 per sqm). A-tagged. '
                 'Set the coefficient above 0 only when preservative treatment is specified.'},
    ],
    '10_Steel_Work': [],  # handled by replacing the existing literal line below
}


def _retarget_mortar_refs(configs):
    """Point every '03_Mortars'!G<n> reference at the sheet's actual Say row.

    trade_configs.py carries these as literal strings; rather than trusting the
    row number written there, rewrite it from the live layout constants.
    """
    for cfg in configs.values():
        for m in cfg.get('default_materials', []) or []:
            f = m.get('custom_rate_formula')
            if isinstance(f, str) and "'03_Mortars'!G" in f:
                m['custom_rate_formula'] = re.sub(r"'03_Mortars'!G\d+",
                                                  "'03_Mortars'!G%d" % R_SAY, f)
            elif isinstance(f, str) and "'03_Mortars'!H" in f:
                m['custom_rate_formula'] = re.sub(r"'03_Mortars'!H\d+",
                                                  "'03_Mortars'!H%d" % (R_LIB_FIRST + 1), f)
    return configs


def merge_scope_metadata(configs):
    """Attach the input-analysis panels and cross-volume lines to each config."""
    for sheet, meta in SCOPE_BY_SHEET.items():
        cfg = configs.get(sheet)
        if not cfg:
            continue
        cfg['scope_inputs'] = meta['scope_inputs']
        cfg['driver_inputs'] = meta['driver_inputs']
        cfg['nomenclature_formula'] = _nomenclature_formula(sheet, cfg)

    # 10 Steel Work: replace the hard-coded 50.70 literal with the resolved item.
    steel = configs.get('10_Steel_Work')
    if steel:
        for m in steel.get('default_materials', []):
            if str(m.get('code')) == '13.50.3':
                m['custom_rate_formula'] = '=XV_13_50_3'
                m.pop('custom_rate', None)
                m['a_tag'] = True
                m['custom_desc'] = ('Priming coat with red oxide zinc chromate primer on steel work '
                                    '(DAR 13.50.3, Finishing)')
                m['custom_unit'] = 'sqm'
                m['note'] = ('Resolved on Resolved_Cross_Volume_Items from WB1 data only (Rs 50.70 per sqm). '
                             'A-tagged: DAR item 10.1 tags this line "A" and computes every markup on '
                             '(W-A), (X-A), (Y-A), (Z-A). Verified: 1% of (6469.38-152.10) = 63.17, exactly '
                             'the figure the book prints.')

    # 08 and 09: append the resolved cross-volume lines.
    for sheet, lines in RESOLVED_LINES.items():
        cfg = configs.get(sheet)
        if not cfg or not lines:
            continue
        mats = cfg.setdefault('default_materials', [])
        existing = {str(m.get('code')) for m in mats}
        for line in lines:
            if str(line['code']) not in existing:
                mats.append(line)

    _retarget_mortar_refs(configs)
    return configs


def _nomenclature_formula(sheet, cfg):
    """Assemble the item name from the Panel 1 selections, with a manual override."""
    from scripts.trade_layout import R_P1_FIRST, R_NOMEN
    dims = cfg.get('scope_inputs') or []
    n = min(len(dims), 6)
    if n == 0:
        return cfg.get('default_item_desc', '')
    stem = {
        '02_Earth_Work': '"Earth work in "',
        '03_Mortars': '"Providing and mixing "',
        '04_Concrete_Work': '"Providing and laying in position cement concrete of grade "',
        '05_RCC_Work': '"Providing and laying in position reinforced cement concrete of grade "',
        '06_Masonry_Work': '"Brick work with "',
        '07_Stone_Work': '"Stone work - "',
        '08_Cladding_Work': '"Stone / marble cladding - "',
        '09_Wood_and_PVC_Work': '"Providing wood work in "',
        '10_Steel_Work': '"Structural steel work in "',
        '11_Flooring': '"Flooring with "',
        '12_Roofing': '"Providing roofing with "',
    }.get(sheet, '""')
    parts = [f'B{R_P1_FIRST + i}' for i in range(n)]
    joined = ' & ", " & '.join(parts)
    return (f'=IF(I{R_NOMEN}<>"", I{R_NOMEN}, {stem} & {joined} & '
            f'", complete as per CPWD specifications and directions of Engineer-in-charge.")')
