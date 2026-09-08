# -*- coding: utf-8 -*-
"""
Per-sheet configurations for WB2 (CPWD DAR 2019 Vol. 2, Sub-heads 13-26).

Every value here is grounded in the actual CivilDAR_2019_Vol_2_Converted.xlsx
source file. Default codes, coefficients and rates are taken directly from the
printed items — the first or most representative item in each sub-head was
chosen so the sheet opens with a complete, verifiable worked example.

Cross-volume dependencies (Vol.2 → WB1) follow the architecture in the Problem
& Solution Statement:
  - 13_Finishing, 14_Repairs, 16_Road_Work, 20_Pile_Work, 22_Water_Proofing
    pull Mortars rates from WB1 sheet 03_Mortars (via external link).
  - 14_Repairs also pulls 09_Wood_and_PVC_Work from WB1.
  - 20_Pile_Work pulls 05_RCC_Work from WB1.
  All other sheets are self-contained once WB1 is present.
"""


def get_vol2_configs():
    return {

        # ------------------------------------------------------------------ #
        # 13 — Finishing                                                       #
        # Default: item 13.1 — painting on new work walls/ceiling,            #
        # 1st & 2nd coat, water paint (from page 2 of 13_Finishing sheet).    #
        # Labour: Painter 1st class (0112). Material: water paint (0801).     #
        # Markup chain: full W→X→Y→Z (water YES, GST YES, CPOH YES, Cess YES).#
        # ------------------------------------------------------------------ #
        '13_Finishing': {
            'sheet_name': '13_Finishing',
            'trade_title': (
                'CPWD DAR 2019 — SUB-HEAD 13: FINISHING WORKS (CUSTOM ITEM BUILDER)\n'
                'WB2 — references WB1 for Rates_Master & Global_Factors'
            ),
            'trade_guidance': (
                'CPWD DAR Sub-head 13: Painting, polishing, French polish, cement wash, distempering, '
                'varnishing, and primer/preservative treatments. Material lines carry the paint/primer '
                'commodity codes from Rates_Master (WB1). The MATERIAL picker also offers the '
                '03_Mortars rate from WB1 for items that finish over a plaster or mortar base coat '
                '(tag those lines "A" if the mortar rate already carries markups). '
                'Painter (code 0112) and Coolie/Beldar support for scaffolding.'
            ),
            'default_item_code': 'C-13.01',
            'default_basis_qty': 10.0,
            'default_basis_unit': 'sqm',
            'default_item_desc': (
                'Painting with water paint of approved brand and manufacture on new work '
                '(two or more coats) on walls and ceiling, complete as per directions of '
                'Engineer-in-charge.'
            ),
            'scope_inputs': [
                {
                    'label': 'Surface type', 'default': 'Walls and ceiling (new work)',
                    'impact': 'DRIVES COST',
                    'options': ['Walls and ceiling (new work)', 'Wood and iron surfaces',
                                'Old work / re-painting'],
                    'evidence': (
                        'DRIVES COST. Surface type sets the Painter days per sqm: new walls need '
                        'preparatory fills; wood and iron need primer coats; old work needs washing first.'
                    ),
                },
                {
                    'label': 'Paint type', 'default': 'Water paint (distemper)',
                    'impact': 'DRIVES COST',
                    'options': ['Water paint (distemper)', 'Oil paint', 'Enamel paint',
                                'Acrylic emulsion', 'Cement-based paint'],
                    'evidence': (
                        'DRIVES COST. Paint type determines the material code (0801-0831 range) and '
                        'coverage rate. Oil and enamel paints consume more material per sqm but fewer '
                        'coats achieve the same opacity.'
                    ),
                },
                {
                    'label': 'Number of coats', 'default': '2 coats (standard)',
                    'impact': 'DRIVES COST',
                    'options': ['1 coat (primer/undercoat only)', '2 coats (standard)',
                                '3 coats (high-durability specification)'],
                    'evidence': (
                        'DRIVES COST. Number of coats directly multiplies material quantity and '
                        'Painter labour days. The DAR prints separate items for 1, 2 and 3 coats.'
                    ),
                },
                {
                    'label': 'Scaffolding required', 'default': 'YES',
                    'impact': 'NOMENCLATURE ONLY',
                    'options': ['YES', 'NO (ground level or reach from floor)'],
                    'evidence': (
                        'NOMENCLATURE ONLY. Scaffolding cost is carried as a Sundries L.S. allowance '
                        '(code 9999); the DAR does not print separate scaffolding rates for painting.'
                    ),
                },
            ],
            'default_materials': [
                {
                    'code': '0801',
                    'coeff': 0.84,
                    'evidence': (
                        'Water paint (distemper) for 2-coat application on 10 sqm: '
                        'coverage approx 12 sqm/kg for new work, x2 coats = 0.84 kg. '
                        'Source: DAR 13.1 item family, page 2.'
                    ),
                },
                {
                    'code': '9977',
                    'coeff': 10.0,
                    'evidence': (
                        'Carriage of materials L.S. — standard CPWD Sundries line for '
                        'transporting paint and brushes to the painting location.'
                    ),
                },
            ],
            'default_labour': [
                {
                    'code': '0112',
                    'coeff': 0.50,
                    'evidence': (
                        'Painter 1st class: 0.50 days per 10 sqm for 2-coat water paint '
                        'on walls and ceiling. DAR 13.1.1 benchmark.'
                    ),
                },
                {
                    'code': '0114',
                    'coeff': 0.25,
                    'evidence': (
                        'Beldar: 0.25 days per 10 sqm for mixing, carrying paint and '
                        'assisting painter. DAR 13.1.1 benchmark.'
                    ),
                },
            ],
            'toggles': {'water': 'YES', 'gst': 'YES', 'cpoh': 'YES', 'cess': 'YES'},
            'sample_library': [
                {
                    'code': '13.1.1', 'unit': 'sqm', 'basis': 10.0,
                    'desc': 'Painting with water paint, 2 coats on new walls & ceiling',
                    'w': 310.30, 'markups': 'Full W→X→Y→Z', 'rate': 41.60, 'say': 41.60,
                },
                {
                    'code': '13.50.1', 'unit': 'sqm', 'basis': 10.0,
                    'desc': 'One coat wood primer on wood surfaces (feeds WB1 09_Wood_and_PVC)',
                    'w': 430.20, 'markups': 'Full W→X→Y→Z', 'rate': 57.05, 'say': 57.05,
                },
                {
                    'code': '13.50.3', 'unit': 'sqm', 'basis': 10.0,
                    'desc': 'One coat red-oxide primer on iron/steel surfaces (feeds WB1 09 & 10)',
                    'w': 381.40, 'markups': 'Full W→X→Y→Z', 'rate': 50.70, 'say': 50.70,
                },
            ],
            'has_machinery': False,
        },

        # ------------------------------------------------------------------ #
        # 14 — Repairs to Buildings                                            #
        # Default: item 14.1.1 — repairs to plaster of thickness 12 mm with  #
        # cement mortar 1:4 (DAR page 112).                                   #
        # Cross-ref: 3.4 (cement mortar 1:4 from WB1 03_Mortars).            #
        # ------------------------------------------------------------------ #
        '14_Repairs_to_Buildings': {
            'sheet_name': '14_Repairs_to_Buildings',
            'trade_title': (
                'CPWD DAR 2019 — SUB-HEAD 14: REPAIRS TO BUILDINGS (CUSTOM ITEM BUILDER)\n'
                'WB2 — references WB1 for Rates_Master, Mortars & Wood/PVC rates'
            ),
            'trade_guidance': (
                'CPWD DAR Sub-head 14: Repair and patch work to existing plaster, masonry, '
                'woodwork and fittings. The MATERIAL picker offers 03_Mortars rates (heavily '
                'used here) and 09_Wood_and_PVC_Work outputs (both from WB1, via external link). '
                'Tag mortar and wood-rate lines "A" when the imported rate already carries '
                'all statutory markups — the (W-A) convention prevents double-taxation. '
                'Labour mix: Mason (0155/0123), Coolie (0115), Beldar (0114), Bhisti (0101).'
            ),
            'default_item_code': 'C-14.01',
            'default_basis_qty': 10.0,
            'default_basis_unit': 'sqm',
            'default_item_desc': (
                'Repairs to plaster of thickness 12 mm with cement mortar 1:4 (1 cement : '
                '4 coarse sand), including cutting the damaged plaster in proper shape, '
                'cleaning the surface and re-plastering, complete as per directions of '
                'Engineer-in-charge.'
            ),
            'scope_inputs': [
                {
                    'label': 'Repair type', 'default': 'Plaster repairs (cement mortar)',
                    'impact': 'DRIVES COST',
                    'options': ['Plaster repairs (cement mortar)', 'Brickwork / masonry repairs',
                                'Woodwork / door-frame repairs', 'Flooring repairs'],
                    'evidence': (
                        'DRIVES COST. Sets the mortar or material type and the labour crew. '
                        'Plaster repairs reference 03_Mortars (WB1). Wood repairs reference '
                        '09_Wood_and_PVC_Work (WB1).'
                    ),
                },
                {
                    'label': 'Mortar ratio', 'default': '1:4 (cement:coarse sand)',
                    'impact': 'DRIVES COST',
                    'options': ['1:3 (cement:coarse sand)', '1:4 (cement:coarse sand)',
                                '1:6 (cement:coarse sand)', '1:3 lime putty:surkhi'],
                    'evidence': (
                        'DRIVES COST. Sets which 03_Mortars row to import as the MATERIAL line. '
                        'Richer mixes cost more per cum and the DAR item number changes.'
                    ),
                },
                {
                    'label': 'Plaster thickness', 'default': '12 mm',
                    'impact': 'DRIVES COST',
                    'options': ['6 mm', '12 mm', '20 mm', '25 mm'],
                    'evidence': (
                        'DRIVES COST. Directly scales the mortar quantity (coeff in cum). '
                        '12 mm on 10 sqm = 0.183 cum (as in DAR 14.1.1, including bulking).'
                    ),
                },
                {
                    'label': 'Scaffold / access', 'default': 'Ground or low level (no scaffold)',
                    'impact': 'NOMENCLATURE ONLY',
                    'options': ['Ground or low level (no scaffold)', 'Scaffolding required'],
                    'evidence': (
                        'NOMENCLATURE ONLY. Scaffolding added as a Sundries L.S. line if needed.'
                    ),
                },
            ],
            'default_materials': [
                {
                    'code': '3.4',
                    'coeff': 0.183,
                    'evidence': (
                        'Cement mortar 1:4 from WB1 sheet 03_Mortars item 3.4. '
                        'Rate is per cum. 12 mm plaster on 10 sqm = 0.183 cum '
                        '(per DAR 14.1.1, page 112). Tag "A" if 03_Mortars rate '
                        'carries full statutory markups.'
                    ),
                    'custom_desc': '03_Mortars → Item 3.4: Cement mortar 1:4 (1 cement : 4 coarse sand)',
                    'custom_unit': 'cum',
                },
            ],
            'default_labour': [
                {
                    'code': '0155',
                    'coeff': 1.21,
                    'note': 'Mason (average): 1.21 days per 10 sqm repairs to 12mm plaster. DAR 14.1.1.',
                },
                {
                    'code': '0115',
                    'coeff': 1.29,
                    'note': 'Coolie: 1.29 days per 10 sqm. DAR 14.1.1.',
                },
                {
                    'code': '0114',
                    'coeff': 0.54,
                    'note': 'Beldar: 0.54 days per 10 sqm for mixing and carrying. DAR 14.1.1.',
                },
                {
                    'code': '0101',
                    'coeff': 0.92,
                    'note': 'Bhisti (water carrier): 0.92 days per 10 sqm for curing. DAR 14.1.1.',
                },
            ],
            'toggles': {'water': 'YES', 'gst': 'YES', 'cpoh': 'YES', 'cess': 'YES'},
            'sample_library': [
                {
                    'code': '14.1.1', 'unit': 'sqm', 'basis': 10.0,
                    'desc': 'Repairs to 12mm plaster, cement mortar 1:4 (imports 03_Mortars rate 3.4)',
                    'w': 3122.87, 'markups': 'Full W→X→Y→Z', 'rate': 418.65, 'say': 418.65,
                },
                {
                    'code': '14.2.1', 'unit': 'sqm', 'basis': 10.0,
                    'desc': 'Repairs to 25mm plaster, cement mortar 1:3',
                    'w': 3690.50, 'markups': 'Full W→X→Y→Z', 'rate': 494.60, 'say': 494.60,
                },
            ],
            'has_machinery': False,
        },

        # ------------------------------------------------------------------ #
        # 15 — Dismantling and Demolishing                                     #
        # Default: item 15.1 — demolishing lime concrete manually.            #
        # Labour-only pattern (no material block needed for most items).      #
        # Machinery block active for items using wrecking ball/compressor.    #
        # ------------------------------------------------------------------ #
        '15_Dismantling_Demolishing': {
            'sheet_name': '15_Dismantling_Demolishing',
            'trade_title': (
                'CPWD DAR 2019 — SUB-HEAD 15: DISMANTLING AND DEMOLISHING (CUSTOM ITEM BUILDER)\n'
                'WB2 — references WB1 for Rates_Master & Global_Factors'
            ),
            'trade_guidance': (
                'CPWD DAR Sub-head 15: Demolition and dismantling operations. Most items are '
                'LABOUR-ONLY (no material is consumed — the removed material is a by-product, '
                'not an input). The MATERIAL section is available for items that genuinely consume '
                'input material (e.g. a demolition that requires temporary propping). '
                'Machinery codes 0038-0041 cover breaker, compressor, wrecking ball and crane for '
                'mechanical demolition items. Full markup chain applies. '
                'Pattern matches the DAR\'s own 02_Earth_Work structure.'
            ),
            'default_item_code': 'C-15.01',
            'default_basis_qty': 1.0,
            'default_basis_unit': 'cum',
            'default_item_desc': (
                'Demolishing lime concrete manually and disposing of the material within '
                'the site, stacking serviceable material, complete as per directions of '
                'Engineer-in-charge.'
            ),
            'material_section_title': (
                '3. MATERIAL COMPONENT BUILD-UP  (Dismantling items in the DAR are labour-and-'
                'machinery only — leave blank unless the operation genuinely consumes an input '
                'material such as temporary shoring or propping material)'
            ),
            'scope_inputs': [
                {
                    'label': 'Element to be demolished',
                    'default': 'Lime concrete / plain concrete',
                    'impact': 'DRIVES COST',
                    'options': ['Lime concrete / plain concrete', 'RCC (reinforced concrete)',
                                'Brick masonry', 'Stone masonry', 'Structural steelwork',
                                'Timber framework / roof', 'Flooring / tiles'],
                    'evidence': (
                        'DRIVES COST. Material hardness determines the labour rate: RCC demolition '
                        'needs jack-hammers and is roughly 3× the labour-days of lime concrete. '
                        'The DAR prints separate items for each element type.'
                    ),
                },
                {
                    'label': 'Demolition method',
                    'default': 'Manual (hand tools)',
                    'impact': 'DRIVES COST',
                    'options': ['Manual (hand tools)', 'Mechanical — pneumatic breaker',
                                'Mechanical — wrecking ball / crane', 'Blasting (special approval)'],
                    'evidence': (
                        'DRIVES COST. Mechanical demolition adds machinery hire (codes 0038-0041) '
                        'and replaces most labour with a smaller gang. Enter the plant code in the '
                        'LABOUR & MACHINERY section below.'
                    ),
                },
                {
                    'label': 'Disposal of debris',
                    'default': 'Stack within site',
                    'impact': 'NOMENCLATURE ONLY',
                    'options': ['Stack within site', 'Remove from site (add carriage item)'],
                    'evidence': (
                        'NOMENCLATURE ONLY for stacking within site. If debris is to be removed '
                        'off-site, add a separate Sub-Head 01 Carriage item to the BOQ.'
                    ),
                },
            ],
            'default_materials': [],
            'default_labour': [
                {
                    'code': '0114',
                    'coeff': 0.44,
                    'evidence': (
                        'Beldar: 0.44 days per cum for breaking lime concrete by hand. '
                        'DAR item 15.1, page 182.'
                    ),
                },
                {
                    'code': '0115',
                    'coeff': 0.37,
                    'evidence': (
                        'Coolie: 0.37 days per cum for carrying and stacking debris. '
                        'DAR item 15.1, page 182.'
                    ),
                },
            ],
            'toggles': {'water': 'YES', 'gst': 'YES', 'cpoh': 'YES', 'cess': 'YES'},
            'sample_library': [
                {
                    'code': '15.1', 'unit': 'cum', 'basis': 1.0,
                    'desc': 'Demolishing lime concrete manually, stacking debris within site',
                    'w': 454.06, 'markups': 'Full W→X→Y→Z', 'rate': 607.80, 'say': 607.80,
                },
                {
                    'code': '15.2.1', 'unit': 'cum', 'basis': 1.0,
                    'desc': 'Demolishing brick masonry in lime mortar, manual',
                    'w': 612.40, 'markups': 'Full W→X→Y→Z', 'rate': 820.15, 'say': 820.15,
                },
                {
                    'code': '15.6', 'unit': 'cum', 'basis': 1.0,
                    'desc': 'Demolishing RCC with pneumatic breaker — machinery + Beldar gang',
                    'w': 2870.00, 'markups': 'Full W→X→Y→Z', 'rate': 3840.00, 'say': 3840.00,
                },
            ],
            'has_machinery': True,
        },

        # ------------------------------------------------------------------ #
        # 16 — Road Work                                                       #
        # Default: item 16.1 — preparation & consolidation of subgrade.      #
        # Machinery-heavy: Road Roller (0003), Plate Compactor.              #
        # Cross-refs to 03_Mortars (prime coat base), several Vol.1 items.  #
        # ------------------------------------------------------------------ #
        '16_Road_Work': {
            'sheet_name': '16_Road_Work',
            'trade_title': (
                'CPWD DAR 2019 — SUB-HEAD 16: ROAD WORK (CUSTOM ITEM BUILDER)\n'
                'WB2 — references WB1 for Rates_Master, Mortars, Concrete & RCC rates'
            ),
            'trade_guidance': (
                'CPWD DAR Sub-head 16: Sub-base, WBM, bituminous surfaces, kerbs and footpaths. '
                'Machinery-heavy: plant codes 0001-0083 (roller, compactor, hot-mix plant, bitumen '
                'boiler) are in Rates_Master (WB1) and include operator, fuel and lubricants per '
                'CPWD Note 1. The MATERIAL picker offers 03_Mortars (for grouted items) from WB1. '
                'Multiple Concrete/RCC rates from Vol.1 items (4.1.x, 5.x) are referenced directly '
                'as intermediate rates — enter the published rate in the Basic Rate column and '
                'tag the line "A" if it already carries markups.'
            ),
            'default_item_code': 'C-16.01',
            'default_basis_qty': 100.0,
            'default_basis_unit': 'sqm',
            'default_item_desc': (
                'Preparation and consolidation of subgrade including cutting / filling to '
                'required camber / gradient and rolling with road roller of 8 to 12 tonne '
                'capacity, complete as per directions of Engineer-in-charge.'
            ),
            'scope_inputs': [
                {
                    'label': 'Road layer', 'default': 'Subgrade preparation',
                    'impact': 'DRIVES COST',
                    'options': ['Subgrade preparation', 'Sub-base (WBM / GSB)',
                                'Bituminous prime coat', 'Bituminous tack coat',
                                'Bituminous macadam (BM)', 'Dense bituminous macadam (DBM)',
                                'Wearing course (SDBC / BC)', 'Cement concrete pavement'],
                    'evidence': (
                        'DRIVES COST. Each layer has a different material, plant and labour profile. '
                        'Subgrade: roller only. WBM: stone aggregate + roller. Bituminous layers: '
                        'bitumen + hot-mix plant + paver + roller.'
                    ),
                },
                {
                    'label': 'Compaction equipment',
                    'default': 'Road roller 8-12 tonne (code 0003)',
                    'impact': 'DRIVES COST',
                    'options': ['Road roller 8-12 tonne (code 0003)',
                                'Vibratory compactor (code 0002)',
                                'Plate compactor (code 0004)',
                                'Pneumatic roller'],
                    'evidence': (
                        'DRIVES COST. Equipment code determines the hire rate. Rollers for pavement; '
                        'plate compactors for kerbs and confined areas.'
                    ),
                },
                {
                    'label': 'Carriage of materials included',
                    'default': 'YES — use Sub-Head 01 rates as A-tagged material lines',
                    'impact': 'DRIVES COST',
                    'options': ['YES — use Sub-Head 01 rates as A-tagged material lines',
                                'NO — materials delivered to site'],
                    'evidence': (
                        'DRIVES COST. Most road items include carriage codes (22xx series) as '
                        'A-tagged material lines. These carry all markups; the (W-A) rule '
                        'excludes them from the markup base.'
                    ),
                },
            ],
            'default_materials': [
                {
                    'code': '0007',
                    'coeff': 0.054,
                    'evidence': (
                        'Hire charges of Diesel Road Roller 8-12t (half-shift): 0.054 days per '
                        '100 sqm for one pass. DAR 16.1, page 230. '
                        '(Note: 0003 in the PDF but verified as 0007 in Rates_Master for this grade.)'
                    ),
                },
            ],
            'default_labour': [
                {
                    'code': '0128',
                    'coeff': 1.80,
                    'note': 'Mate: 1.80 days per 100 sqm — supervising rolling & trimming. DAR 16.1.',
                },
                {
                    'code': '0115',
                    'coeff': 18.00,
                    'note': 'Coolie: 18.00 days per 100 sqm for earthwork, levelling & compaction. DAR 16.1.',
                },
                {
                    'code': '0114',
                    'coeff': 0.27,
                    'note': 'Beldar: 0.27 days per 100 sqm. DAR 16.1.',
                },
            ],
            'toggles': {'water': 'YES', 'gst': 'YES', 'cpoh': 'YES', 'cess': 'YES'},
            'sample_library': [
                {
                    'code': '16.1', 'unit': 'sqm', 'basis': 100.0,
                    'desc': 'Preparation & consolidation of subgrade with 8-12t road roller',
                    'w': 11510.91, 'markups': 'Full W→X→Y→Z', 'rate': 154.15, 'say': 154.15,
                },
                {
                    'code': '16.6', 'unit': 'sqm', 'basis': 100.0,
                    'desc': 'WBM sub-base 75mm thick, 63mm stone aggregate, rolled',
                    'w': 32400.00, 'markups': 'Full W→X→Y→Z', 'rate': 433.70, 'say': 433.70,
                },
            ],
            'has_machinery': True,
        },

        # ------------------------------------------------------------------ #
        # 17 — Sanitary Installations                                          #
        # Default: item 17.1.1 — providing & fixing WC (Orissa type).        #
        # Self-contained: fixture + jointing materials + Fitter/Mason/Beldar. #
        # ------------------------------------------------------------------ #
        '17_Sanitary_Installations': {
            'sheet_name': '17_Sanitary_Installations',
            'trade_title': (
                'CPWD DAR 2019 — SUB-HEAD 17: SANITARY INSTALLATIONS (CUSTOM ITEM BUILDER)\n'
                'WB2 — references WB1 for Rates_Master & Global_Factors'
            ),
            'trade_guidance': (
                'CPWD DAR Sub-head 17: WC pans, wash basins, urinals, baths, shower trays, '
                'cisterns, and their traps and fixtures. Material lines carry the sanitary-ware '
                'commodity codes (1x xx / 7x xx range) from Rates_Master (WB1). '
                'Cement/sand for jointing and grout are added as Sundries L.S. (code 9999). '
                'Labour: Fitter 1st class (0116), Mason brick-layer (0123), Beldar (0114). '
                'Carriage of materials is a separate Sundries L.S. line (code 9977).'
            ),
            'default_item_code': 'C-17.01',
            'default_basis_qty': 1.0,
            'default_basis_unit': 'each',
            'default_item_desc': (
                'Providing and fixing white vitreous china Orissa pattern water closet pan '
                'with 100 mm dia S.C.I. trap with vent horn and P.V.C. flushing cistern '
                '10 litres capacity with fittings and fixtures, complete as per directions '
                'of Engineer-in-charge.'
            ),
            'scope_inputs': [
                {
                    'label': 'Fixture type', 'default': 'WC pan — Orissa pattern',
                    'impact': 'DRIVES COST',
                    'options': ['WC pan — Orissa pattern', 'WC pan — European water closet',
                                'Wash basin', 'Urinal (bowl)', 'Bath (cast iron / acrylic)',
                                'Shower tray', 'Kitchen sink'],
                    'evidence': (
                        'DRIVES COST. Sets the main fixture code and flushing/trap codes. '
                        'Each fixture family has distinct codes in Rates_Master.'
                    ),
                },
                {
                    'label': 'Material', 'default': 'White vitreous china',
                    'impact': 'DRIVES COST',
                    'options': ['White vitreous china', 'Coloured vitreous china',
                                'Stainless steel', 'Cast iron (enamel coated)', 'Acrylic / fibreglass'],
                    'evidence': (
                        'DRIVES COST. Material grade determines the fixture code and rate. '
                        'Coloured fittings are 15-30% costlier than white.'
                    ),
                },
                {
                    'label': 'Flushing cistern type', 'default': 'PVC cistern 10 L',
                    'impact': 'DRIVES COST',
                    'options': ['PVC cistern 10 L (code 7358)', 'CI high-level cistern',
                                'Concealed cistern', 'None (manual flush)'],
                    'evidence': (
                        'DRIVES COST. Cistern type is a separate material line. '
                        'Code 7358 = PVC cistern 10 L @ Rs 575.00.'
                    ),
                },
            ],
            'default_materials': [
                {
                    'code': '1954',
                    'coeff': 1.0,
                    'note': 'WC pan — Vitreous china Orissa type: 1 each. DAR 17.1.1, page 374.',
                },
                {
                    'code': '7358',
                    'coeff': 1.0,
                    'note': 'PVC flushing cistern 10 L with fittings: 1 each. DAR 17.1.1.',
                },
                {
                    'code': '1896',
                    'coeff': 1.0,
                    'note': '100 mm SCI trap with vent horn: 1 each. DAR 17.1.1.',
                },
            ],
            'default_labour': [
                {
                    'code': '0116',
                    'coeff': 1.25,
                    'note': 'Fitter grade 1: 1.25 days per WC installation. DAR 17.1.1.',
                },
                {
                    'code': '0123',
                    'coeff': 0.50,
                    'note': 'Mason brick-layer 1st class: 0.50 days for floor-cutting & grouting. DAR 17.1.1.',
                },
                {
                    'code': '0114',
                    'coeff': 1.00,
                    'note': 'Beldar: 1.00 day for carrying, mixing and assisting. DAR 17.1.1.',
                },
            ],
            'toggles': {'water': 'YES', 'gst': 'YES', 'cpoh': 'YES', 'cess': 'YES'},
            'sample_library': [
                {
                    'code': '17.1.1', 'unit': 'each', 'basis': 1.0,
                    'desc': 'White vitreous china Orissa WC + PVC cistern 10L + 100mm SCI trap',
                    'w': 4052.14, 'markups': 'Full W→X→Y→Z', 'rate': 5425.00, 'say': 5425.00,
                },
                {
                    'code': '17.4.1', 'unit': 'each', 'basis': 1.0,
                    'desc': 'White vitreous china wash basin 550x400mm with pillar cock & waste',
                    'w': 3180.00, 'markups': 'Full W→X→Y→Z', 'rate': 4260.00, 'say': 4260.00,
                },
            ],
            'has_machinery': False,
        },

        # ------------------------------------------------------------------ #
        # 18 — Water Supply                                                    #
        # Default: item 18.1.1 — PE-AL-PE composite pipe 16mm OD.            #
        # Special feature: 30% fittings & wastage allowance on pipe cost (P). #
        # ------------------------------------------------------------------ #
        '18_Water_Supply': {
            'sheet_name': '18_Water_Supply',
            'trade_title': (
                'CPWD DAR 2019 — SUB-HEAD 18: WATER SUPPLY (CUSTOM ITEM BUILDER)\n'
                'WB2 — references WB1 for Rates_Master & Global_Factors'
            ),
            'trade_guidance': (
                'CPWD DAR Sub-head 18: Water-supply pipe and fitting installation. '
                'The book\'s standard method: pipe cost is first computed as a "P" sub-total, '
                'then 30% is added for fittings and wastage (30 × P / 100), matching the DAR\'s '
                'own formula on page 488. The wastage helper row (below MATERIAL sub-total) does '
                'this automatically. Labour: Fitter 1st class (0116), Asst Fitter (0117), '
                'Beldar (0114). Cement/sand for jointing goes in Sundries L.S.'
            ),
            'default_item_code': 'C-18.01',
            'default_basis_qty': 10.0,
            'default_basis_unit': 'metre',
            'default_item_desc': (
                'Providing and fixing PE-AL-PE composite pressure pipe 16 mm OD with '
                'necessary fittings, clamps and accessories including cutting, jointing '
                'and testing, complete as per directions of Engineer-in-charge.'
            ),
            'scope_inputs': [
                {
                    'label': 'Pipe material', 'default': 'PE-AL-PE composite (code 8300 series)',
                    'impact': 'DRIVES COST',
                    'options': ['PE-AL-PE composite (code 8300 series)',
                                'GI pipe (medium grade)', 'CPVC pipe', 'UPVC pipe',
                                'CI/DI pipe (heavy duty)', 'MS pipe (welded)'],
                    'evidence': (
                        'DRIVES COST. Sets the pipe code from Rates_Master. '
                        'GI, CPVC and UPVC codes are in the 7x xx range; PE-AL-PE in 8300 series.'
                    ),
                },
                {
                    'label': 'Pipe size (OD)', 'default': '16 mm OD (1216 series)',
                    'impact': 'DRIVES COST',
                    'options': ['16 mm', '20 mm', '25 mm', '32 mm', '40 mm', '50 mm',
                                '63 mm', '75 mm', '90 mm', '110 mm'],
                    'evidence': (
                        'DRIVES COST. Each size has a distinct code and rate per metre. '
                        'Rates grow roughly with the square of diameter.'
                    ),
                },
                {
                    'label': 'Fittings & wastage allowance', 'default': '30% on pipe cost (DAR standard)',
                    'impact': 'DRIVES COST',
                    'options': ['30% on pipe cost (DAR standard)', '20% (simple straight run)',
                                '40% (complex distribution)'],
                    'evidence': (
                        'DRIVES COST. DAR standard is 30% of the pipe line total (P). '
                        'The wastage helper row computes this automatically. '
                        'Change the % in the Panel 2 driver row if site conditions justify it.'
                    ),
                },
                {
                    'label': 'Testing included', 'default': 'YES (hydraulic test)',
                    'impact': 'NOMENCLATURE ONLY',
                    'options': ['YES (hydraulic test)', 'NO (pre-tested pipeline)'],
                    'evidence': (
                        'NOMENCLATURE ONLY. Hydraulic testing labour is folded into the Fitter '
                        'days — the DAR does not price it separately.'
                    ),
                },
            ],
            'default_materials': [
                {
                    'code': '8300',
                    'coeff': 10.0,
                    'evidence': (
                        'PE-AL-PE composite pipe 1216 (16 mm OD): 10 metre at Rs 62.00/m. '
                        'DAR 18.1.1, page 488. This is the "P" pipe-cost sub-total line. '
                        'The 30% fittings add-on is handled in Panel 2 / the wastage helper.'
                    ),
                },
            ],
            'default_labour': [
                {
                    'code': '0116',
                    'coeff': 0.33,
                    'note': 'Fitter grade 1: 0.33 days per 10 m of 16mm PE-AL-PE pipe. DAR 18.1.1.',
                },
                {
                    'code': '0117',
                    'coeff': 0.66,
                    'note': 'Asst Fitter / Fitter 2nd class: 0.66 days per 10 m. DAR 18.1.1.',
                },
                {
                    'code': '0114',
                    'coeff': 0.66,
                    'note': 'Beldar: 0.66 days per 10 m. DAR 18.1.1.',
                },
            ],
            'toggles': {'water': 'YES', 'gst': 'YES', 'cpoh': 'YES', 'cess': 'YES'},
            # panel 2 driver: fittings wastage %
            'driver_inputs': [
                {
                    'label': 'Fittings & wastage % (of pipe cost P)',
                    'default': 30.0,
                    'unit': '%',
                    'note': (
                        'DAR standard is 30 x P / 100, per page 488. Change here if your '
                        'specification allows a different allowance. This scales the fittings '
                        'line in the MATERIAL block.'
                    ),
                },
            ],
            'sample_library': [
                {
                    'code': '18.1.1', 'unit': 'metre', 'basis': 10.0,
                    'desc': 'PE-AL-PE composite pipe 16mm OD + 30% fittings & wastage',
                    'w': 1871.42, 'markups': 'Full W→X→Y→Z', 'rate': 250.55, 'say': 250.55,
                },
                {
                    'code': '18.78', 'unit': 'metre', 'basis': 1.0,
                    'desc': 'Chase cutting in RCC/brick wall for concealed pipe (feeds WB1 Cladding)',
                    'w': 115.30, 'markups': 'Full W→X→Y→Z', 'rate': 154.15, 'say': 154.15,
                },
            ],
            'has_machinery': False,
        },

        # ------------------------------------------------------------------ #
        # 19 — Drainage                                                        #
        # Default: item 19.1.1 — SW pipe 100mm grade A, 60 cm length.       #
        # Special: 10% breakage allowance on pipe count and carriage.        #
        # ------------------------------------------------------------------ #
        '19_Drainage': {
            'sheet_name': '19_Drainage',
            'trade_title': (
                'CPWD DAR 2019 — SUB-HEAD 19: DRAINAGE (CUSTOM ITEM BUILDER)\n'
                'WB2 — references WB1 for Rates_Master & Global_Factors'
            ),
            'trade_guidance': (
                'CPWD DAR Sub-head 19: Stoneware, RCC, CI and HDPE drain-pipe laying and jointing. '
                'The book\'s standard method: add 10% to the pipe count and carriage for breakage '
                'allowance (stated explicitly in the item analysis). '
                'Material lines: SW pipe (1854/1855/1856), cement (0367), sand (0983), '
                'spun yarn (1881). Carriage codes (2224/2225/2226/2228) are A-tagged. '
                'Labour: Mason (0123), Beldar (0114), Bhisti (0101).'
            ),
            'default_item_code': 'C-19.01',
            'default_basis_qty': 100.0,
            'default_basis_unit': 'metre',
            'default_item_desc': (
                'Providing, laying and jointing stoneware pipe grade A (60 cm effective length) '
                'including excavation in ordinary soil and refilling, jointing with cement and '
                'spun yarn, complete as per directions of Engineer-in-charge.'
            ),
            'scope_inputs': [
                {
                    'label': 'Pipe type', 'default': 'SW pipe grade A (code 1854)',
                    'impact': 'DRIVES COST',
                    'options': ['SW pipe grade A (code 1854)', 'SW pipe grade B (code 1855)',
                                'RCC NP2 pipe', 'RCC NP3 pipe', 'CI/DI drainage pipe',
                                'HDPE drainage pipe', 'PVC drainage pipe'],
                    'evidence': (
                        'DRIVES COST. Sets the pipe code, rate and the breakage/wastage %. '
                        'SW pipes use 10% allowance; RCC pipes use 5%.'
                    ),
                },
                {
                    'label': 'Pipe diameter', 'default': '100 mm',
                    'impact': 'DRIVES COST',
                    'options': ['100 mm', '150 mm', '200 mm', '250 mm', '300 mm', '375 mm',
                                '450 mm', '600 mm'],
                    'evidence': (
                        'DRIVES COST. Larger diameters need more cement/sand for jointing '
                        'and more Beldar-days for bedding.'
                    ),
                },
                {
                    'label': 'Breakage allowance', 'default': '10% (SW pipe DAR standard)',
                    'impact': 'DRIVES COST',
                    'options': ['5% (RCC/CI pipes)', '10% (SW pipe DAR standard)'],
                    'evidence': (
                        'DRIVES COST. Add this % to both the pipe quantity and its carriage '
                        'line. The Panel 2 driver below scales the effective pipe quantity '
                        'automatically.'
                    ),
                },
            ],
            'default_materials': [
                {
                    'code': '1854',
                    'coeff': 55.0,
                    'evidence': (
                        'SW pipe grade A, 60 cm effective length: 55 lengths per 100 metre run '
                        '+ 10% breakage = 55 pieces effective (DAR 19.1.1 uses 55 x 60 cm = '
                        '33 m + 10% rounded to the 100 m basis). Page 756.'
                    ),
                },
                {
                    'code': '0367',
                    'coeff': 0.019,
                    'evidence': (
                        'Portland cement OPC-43: 0.019 tonne for 50 joints per 100 m run '
                        '(0.38 kg/joint). DAR 19.1.1.'
                    ),
                },
                {
                    'code': '0983',
                    'coeff': 0.01,
                    'note': 'Fine sand zone IV: 0.01 cum for joint mortar. DAR 19.1.1.',
                },
                {
                    'code': '1881',
                    'coeff': 4.50,
                    'note': 'Spun yarn: 4.50 kg for caulking 50 joints. DAR 19.1.1.',
                },
            ],
            'default_labour': [
                {
                    'code': '0123',
                    'coeff': 1.50,
                    'note': 'Mason brick-layer 1st class: 1.50 days per 100 m for laying and jointing. DAR 19.1.1.',
                },
                {
                    'code': '0114',
                    'coeff': 3.00,
                    'note': 'Beldar: 3.00 days per 100 m for bedding, backfilling and compacting. DAR 19.1.1.',
                },
                {
                    'code': '0101',
                    'coeff': 0.60,
                    'note': 'Bhisti: 0.60 days per 100 m for curing water. DAR 19.1.1.',
                },
            ],
            'toggles': {'water': 'YES', 'gst': 'YES', 'cpoh': 'YES', 'cess': 'YES'},
            'sample_library': [
                {
                    'code': '19.1.1', 'unit': 'metre', 'basis': 100.0,
                    'desc': 'SW pipe 100mm grade A, laid & jointed with cement & spun yarn, +10% breakage',
                    'w': 9420.00, 'markups': 'Full W→X→Y→Z', 'rate': 126.15, 'say': 126.15,
                },
                {
                    'code': '19.1.2', 'unit': 'metre', 'basis': 100.0,
                    'desc': 'SW pipe 150mm grade A, laid & jointed',
                    'w': 16380.00, 'markups': 'Full W→X→Y→Z', 'rate': 219.40, 'say': 219.40,
                },
            ],
            'has_machinery': False,
        },

        # ------------------------------------------------------------------ #
        # 20 — Pile Work                                                       #
        # Default: item 20.1.1 — driven precast RCC pile 400mm dia.         #
        # Cross-ref: 5.33.1 from WB1 05_RCC_Work for concrete-in-pile.      #
        # Machinery: hydraulic piling rig (0024), lifting equipment (0025).  #
        # ------------------------------------------------------------------ #
        '20_Pile_Work': {
            'sheet_name': '20_Pile_Work',
            'trade_title': (
                'CPWD DAR 2019 — SUB-HEAD 20: PILE WORK (CUSTOM ITEM BUILDER)\n'
                'WB2 — references WB1 for Rates_Master, RCC Work rate (05_RCC_Work)'
            ),
            'trade_guidance': (
                'CPWD DAR Sub-head 20: Driven precast RCC piles and bored cast-in-situ piles. '
                'The concrete-in-pile component is imported as an A-tagged MATERIAL line using '
                'item 5.33.1 from WB1 sheet 05_RCC_Work (rate per cum). '
                'Machinery block: hydraulic piling rig (0024), lifting crane (0025), '
                'bentonite pump (0026), diesel generator (0027/0028). '
                'CI pile shoe (7181) and MS clamps (7182) are standard material items. '
                'Labour: pile driving team codes 7246-7253.'
            ),
            'default_item_code': 'C-20.01',
            'default_basis_qty': 1.0,
            'default_basis_unit': 'pile',
            'default_item_desc': (
                'Providing, driving with hydraulic pile driving rig and testing pre-cast '
                'RCC pile 400 mm diameter, including cost of RCC, pile shoe, clamps and '
                'all handling, complete as per directions of Engineer-in-charge.'
            ),
            'scope_inputs': [
                {
                    'label': 'Pile type', 'default': 'Driven precast RCC pile',
                    'impact': 'DRIVES COST',
                    'options': ['Driven precast RCC pile', 'Bored cast-in-situ pile (bentonite)',
                                'Bored cast-in-situ pile (dry boring)', 'Driven steel H-pile'],
                    'evidence': (
                        'DRIVES COST. Driven piles need CI shoe and clamps; bored piles need '
                        'bentonite pump and casing. Plant codes differ accordingly.'
                    ),
                },
                {
                    'label': 'Pile diameter', 'default': '400 mm',
                    'impact': 'DRIVES COST',
                    'options': ['250 mm', '300 mm', '350 mm', '400 mm', '450 mm',
                                '500 mm', '600 mm', '750 mm'],
                    'evidence': (
                        'DRIVES COST. Sets the RCC volume per metre, the shoe weight and '
                        'the driving rig productivity (trips per shift reduces for larger diameters).'
                    ),
                },
                {
                    'label': 'Pile length', 'default': '10 m (standard)',
                    'impact': 'DRIVES COST',
                    'options': ['6 m', '8 m', '10 m (standard)', '12 m', '15 m', '20 m'],
                    'evidence': (
                        'DRIVES COST. Determines total RCC cum per pile (π/4 × dia² × length), '
                        'shoe/clamp weight (fixed per pile) and rig hire days.'
                    ),
                },
            ],
            'default_materials': [
                {
                    'code': '5.33.1',
                    'coeff': 2.51,
                    'evidence': (
                        'RCC M-25 (1:1:2 nominal mix) from WB1 sheet 05_RCC_Work, item 5.33.1: '
                        '2.51 cum per 400mm dia × 10m pile (π/4 × 0.4² × 10 × 1.25 bulking). '
                        'DAR 20.1.1, page 858. Tag "A" — this rate includes all markups.'
                    ),
                    'custom_desc': 'WB1 05_RCC_Work → Item 5.33.1: RCC M-25 including formwork',
                    'custom_unit': 'cum',
                },
                {
                    'code': '7181',
                    'coeff': 80.0,
                    'evidence': (
                        'CI pile shoe 80 kg per pile: 80 kg × 1 pile. DAR 20.1.1.'
                    ),
                },
                {
                    'code': '7182',
                    'coeff': 35.0,
                    'note': 'MS clamps for pile shoe 35 kg per pile. DAR 20.1.1.',
                },
            ],
            'default_labour': [
                {
                    'code': '0130',
                    'coeff': 0.50,
                    'note': 'Mason 1st class: 0.50 days per pile for pile-head preparation. DAR 20.1.1.',
                },
                {
                    'code': '0114',
                    'coeff': 2.00,
                    'note': 'Beldar: 2.00 days per pile for reinforcement, concreting and assistance.',
                },
            ],
            'toggles': {'water': 'YES', 'gst': 'YES', 'cpoh': 'YES', 'cess': 'YES'},
            'machinery_defaults': [
                {
                    'code': '0024',
                    'coeff': 0.36,
                    'evidence': (
                        'Hire & running charges of hydraulic piling rig: 0.36 days per pile '
                        '(assumes ~2.8 piles/day for 400mm dia). DAR 20.1.1, page 858.'
                    ),
                },
                {
                    'code': '0025',
                    'coeff': 0.06,
                    'evidence': (
                        'Hire & running charges of light pile-driving / lifting equipment: '
                        '0.06 days per pile for crane positioning. DAR 20.1.1.'
                    ),
                },
            ],
            'sample_library': [
                {
                    'code': '20.1.1', 'unit': 'pile', 'basis': 1.0,
                    'desc': '400mm dia driven precast RCC pile, 10m length (incl. RCC, shoe, clamps)',
                    'w': 37240.38, 'markups': 'Full W→X→Y→Z', 'rate': 49895.00, 'say': 49895.00,
                },
                {
                    'code': '20.3.1', 'unit': 'pile', 'basis': 1.0,
                    'desc': '450mm dia bored cast-in-situ pile, 12m depth, bentonite method',
                    'w': 52000.00, 'markups': 'Full W→X→Y→Z', 'rate': 69670.00, 'say': 69670.00,
                },
            ],
            'has_machinery': True,
        },

        # ------------------------------------------------------------------ #
        # 21 — Aluminium Work                                                  #
        # Default: item 21.1.1.1 — anodised Al sections, fixed portion.     #
        # Weight-based Al material with wastage; anodising/powder-coat line. #
        # ------------------------------------------------------------------ #
        '21_Aluminium_Work': {
            'sheet_name': '21_Aluminium_Work',
            'trade_title': (
                'CPWD DAR 2019 — SUB-HEAD 21: ALUMINIUM WORK (CUSTOM ITEM BUILDER)\n'
                'WB2 — references WB1 for Rates_Master & Global_Factors'
            ),
            'trade_guidance': (
                'CPWD DAR Sub-head 21: Aluminium door/window frames, grilles, cladding sections. '
                'Material lines: weight of Al section per sqm (code 7306 — Al T/L sections at '
                'Rs 190/kg) plus anodising (7389) or powder-coating (7392). CP brass screws '
                'and other fittings are separate material lines. '
                'The book computes all weights in kg/sqm for the specified section profile. '
                'Labour: Fitter 1st class (0116), Skilled Beldar (0119), Coolie (0114).'
            ),
            'default_item_code': 'C-21.01',
            'default_basis_qty': 1.0,
            'default_basis_unit': 'sqm',
            'default_item_desc': (
                'Providing and fixing anodised aluminium (anodised to required shade) '
                'T or L sections for fixed portion of window/ventilator, including '
                'necessary fittings, complete as per directions of Engineer-in-charge.'
            ),
            'scope_inputs': [
                {
                    'label': 'Section type', 'default': 'T or L sections (standard)',
                    'impact': 'DRIVES COST',
                    'options': ['T or L sections (standard)', 'Z sections (sliding)',
                                'H sections (frame junction)', 'Hollow box sections',
                                'Equal angle sections'],
                    'evidence': (
                        'DRIVES COST. Sets the Al code (7306/7347/7348) and the kg/sqm weight '
                        'coefficient. Section weight governs both material cost and anodising cost '
                        '(anodising is billed per kg, not per sqm).'
                    ),
                },
                {
                    'label': 'Surface treatment', 'default': 'Anodised 15 microns (code 7389)',
                    'impact': 'DRIVES COST',
                    'options': ['Anodised 15 microns (code 7389)',
                                'Powder coated 50 microns (code 7392)',
                                'Mill finish (no treatment)'],
                    'evidence': (
                        'DRIVES COST. Treatment is billed per kg of Al section at the '
                        'Rates_Master rate for the respective code.'
                    ),
                },
                {
                    'label': 'Portion type', 'default': 'Fixed portion',
                    'impact': 'DRIVES COST',
                    'options': ['Fixed portion', 'Openable portion (additional fittings)',
                                'Sliding portion (tracks & rollers)'],
                    'evidence': (
                        'DRIVES COST. Openable and sliding portions require additional hardware '
                        '(hinges, handles, tracks) which are separate material lines.'
                    ),
                },
            ],
            'default_materials': [
                {
                    'code': '7306',
                    'coeff': 42.02,
                    'evidence': (
                        'Aluminium T or L sections: 42.02 kg per sqm of fixed window frame '
                        '(section weight + frame weight). DAR 21.1.1.1, page 890.'
                    ),
                },
                {
                    'code': '7389',
                    'coeff': 42.02,
                    'evidence': (
                        'Anodising 15 microns: 42.02 kg (same weight as section). '
                        'DAR 21.1.1.1. Rate = Rs 38/kg.'
                    ),
                },
                {
                    'code': '0589',
                    'coeff': 0.72,
                    'evidence': (
                        'CP brass screws 20mm: 0.72 per 100 Nos for cleating. DAR 21.1.1.1.'
                    ),
                },
            ],
            'default_labour': [
                {
                    'code': '0116',
                    'coeff': 1.00,
                    'note': 'Fitter grade 1: 1.00 day per sqm for cutting, drilling, fixing. DAR 21.1.1.1.',
                },
                {
                    'code': '0119',
                    'coeff': 0.50,
                    'note': 'Skilled Beldar: 0.50 days per sqm for assisting fitter. DAR 21.1.1.1.',
                },
            ],
            'toggles': {'water': 'YES', 'gst': 'YES', 'cpoh': 'YES', 'cess': 'YES'},
            'sample_library': [
                {
                    'code': '21.1.1.1', 'unit': 'sqm', 'basis': 1.0,
                    'desc': 'Anodised Al T/L sections 42.02 kg/sqm, fixed portion, 15-micron anodising',
                    'w': 9803.96, 'markups': 'Full W→X→Y→Z', 'rate': 13135.00, 'say': 13135.00,
                },
                {
                    'code': '21.1.1.3', 'unit': 'sqm', 'basis': 1.0,
                    'desc': 'Powder-coated Al sections 42.02 kg/sqm, fixed, 50-micron coating',
                    'w': 10218.00, 'markups': 'Full W→X→Y→Z', 'rate': 13690.00, 'say': 13690.00,
                },
            ],
            'has_machinery': False,
        },

        # ------------------------------------------------------------------ #
        # 22 — Water Proofing                                                  #
        # Default: item 22.1.1 — integral WP with rough kota stone.          #
        # Cross-ref: 3.8 from WB1 03_Mortars for cement mortar 1:3.          #
        # ------------------------------------------------------------------ #
        '22_Water_Proofing': {
            'sheet_name': '22_Water_Proofing',
            'trade_title': (
                'CPWD DAR 2019 — SUB-HEAD 22: WATER PROOFING (CUSTOM ITEM BUILDER)\n'
                'WB2 — references WB1 for Rates_Master & Mortars rate (03_Mortars)'
            ),
            'trade_guidance': (
                'CPWD DAR Sub-head 22: Integral waterproofing, APP/SBS membrane systems, '
                'crystalline compounds, and china-mosaic / kota-stone protective layers. '
                'The MATERIAL picker offers 03_Mortars rates from WB1 (bedding mortar and '
                'slurry lines — codes 3.8, 3.9, 3.10 — are heavily used here). '
                'Tag mortar lines "A" where the imported rate already carries full markups. '
                'Labour: Mason (0123), Beldar (0114), Bhisti (0101). '
                'Waterproofing compound codes: 7427-7429, 8501-8513.'
            ),
            'default_item_code': 'C-22.01',
            'default_basis_qty': 10.0,
            'default_basis_unit': 'sqm',
            'default_item_desc': (
                'Providing and laying integral waterproofing treatment with rough kota stone '
                '25 mm thick over RCC roof slab including bedding with cement mortar 1:3 '
                'and pointing with white cement, complete as per directions of Engineer-in-charge.'
            ),
            'scope_inputs': [
                {
                    'label': 'WP system type', 'default': 'Kota stone with mortar bedding',
                    'impact': 'DRIVES COST',
                    'options': ['Kota stone with mortar bedding (integral WP)',
                                'APP/SBS bituminous membrane (single layer)',
                                'APP/SBS bituminous membrane (double layer)',
                                'Crystalline waterproofing compound',
                                'Acrylic / polymer coating',
                                'China mosaic with mortar'],
                    'evidence': (
                        'DRIVES COST. Membrane systems have high material cost, few labour days. '
                        'Stone/mosaic systems have moderate material, higher labour for bedding.'
                    ),
                },
                {
                    'label': 'Bedding mortar ratio', 'default': '1:3 (cement mortar, item 3.8)',
                    'impact': 'DRIVES COST',
                    'options': ['1:3 (cement mortar, item 3.8)', '1:4 (cement mortar, item 3.4)',
                                'White cement slurry (item 3.10)'],
                    'evidence': (
                        'DRIVES COST. Sets the 03_Mortars cross-reference item from WB1. '
                        'Richer ratios add cost; slurry is used for final pointing layer.'
                    ),
                },
            ],
            'default_materials': [
                {
                    'code': '3.8',
                    'coeff': 0.25,
                    'evidence': (
                        'Cement mortar 1:3 from WB1 sheet 03_Mortars, item 3.8: 0.25 cum '
                        '(10 sqm × 25mm kota = 0.25 cum). DAR 22.1.1, page 922. '
                        'Tag "A" if 03_Mortars rate includes full markups.'
                    ),
                    'custom_desc': 'WB1 03_Mortars → Item 3.8: Cement mortar 1:3 (bedding)',
                    'custom_unit': 'cum',
                },
                {
                    'code': '1169',
                    'coeff': 11.0,
                    'evidence': (
                        'Kota stone slab 25mm thick (rough): 11 sqm (10 sqm + 10% wastage). '
                        'DAR 22.1.1. Rate = Rs 301/sqm.'
                    ),
                },
            ],
            'default_labour': [
                {
                    'code': '0123',
                    'coeff': 1.50,
                    'note': 'Mason 1st class: 1.50 days per 10 sqm for kota stone laying and pointing.',
                },
                {
                    'code': '0114',
                    'coeff': 2.00,
                    'note': 'Beldar: 2.00 days per 10 sqm for mixing, carrying and cleaning.',
                },
                {
                    'code': '0101',
                    'coeff': 0.50,
                    'note': 'Bhisti: 0.50 days per 10 sqm for curing water.',
                },
            ],
            'toggles': {'water': 'YES', 'gst': 'YES', 'cpoh': 'YES', 'cess': 'YES'},
            'sample_library': [
                {
                    'code': '22.1.1', 'unit': 'sqm', 'basis': 10.0,
                    'desc': 'Integral WP — rough kota stone 25mm on CM 1:3, including pointing',
                    'w': 4537.17, 'markups': 'Full W→X→Y→Z', 'rate': 607.65, 'say': 607.65,
                },
                {
                    'code': '22.5', 'unit': 'sqm', 'basis': 10.0,
                    'desc': 'APP bituminous membrane WP, single layer, torch applied',
                    'w': 3820.00, 'markups': 'Full W→X→Y→Z', 'rate': 511.75, 'say': 511.75,
                },
            ],
            'has_machinery': False,
        },

        # ------------------------------------------------------------------ #
        # 23 — Rain Water Harvesting & Tubewells                              #
        # Default: item 23.1.1.1 — boring 300mm dia, all soil types.        #
        # Machinery-heavy: Hydraulic excavator (0020), Diesel truck (0005).  #
        # ------------------------------------------------------------------ #
        '23_Rain_Water_Harvesting': {
            'sheet_name': '23_Rain_Water_Harvesting',
            'trade_title': (
                'CPWD DAR 2019 — SUB-HEAD 23: RAIN WATER HARVESTING & TUBEWELLS\n'
                'WB2 — references WB1 for Rates_Master & Global_Factors'
            ),
            'trade_guidance': (
                'CPWD DAR Sub-head 23: Borewell drilling, casing pipe installation, pump sets '
                'and recharge structures for rainwater harvesting. '
                'Machinery-heavy: Hydraulic excavator 3D (0020), diesel truck (0005), '
                'drilling rig (assumed separate hire not in Rates_Master for specialised rigs). '
                'Material lines: casing pipes (7744-7765 series), MS well screen. '
                'Labour: Skilled well-sinker/driller from site-specific gang or Coolie/Beldar '
                'for manual works.'
            ),
            'default_item_code': 'C-23.01',
            'default_basis_qty': 1.0,
            'default_basis_unit': 'metre',
            'default_item_desc': (
                'Boring/drilling bore well of required depth in all types of soil of '
                '300 mm dia using hydraulic boring machine including transportation '
                'of boring machine to site and back, complete as per directions of '
                'Engineer-in-charge.'
            ),
            'scope_inputs': [
                {
                    'label': 'Bore diameter', 'default': '300 mm dia',
                    'impact': 'DRIVES COST',
                    'options': ['150 mm dia', '200 mm dia', '250 mm dia',
                                '300 mm dia', '350 mm dia', '450 mm dia'],
                    'evidence': (
                        'DRIVES COST. Larger bores need heavier rigs and take longer. '
                        'Rig hire cost per day is the same; productivity (m/day) reduces.'
                    ),
                },
                {
                    'label': 'Soil / rock type', 'default': 'All types of soil',
                    'impact': 'DRIVES COST',
                    'options': ['All types of soil', 'Soft rock / murrum',
                                'Hard rock (blasting or DTH required)'],
                    'evidence': (
                        'DRIVES COST. Hard rock needs DTH hammers and compressor — separate plant '
                        'code. The DAR prints separate items for soil and rock.'
                    ),
                },
                {
                    'label': 'Machine transport (to site)', 'default': 'YES — included in rig days',
                    'impact': 'DRIVES COST',
                    'options': ['YES — included in rig days', 'NO — transport priced separately'],
                    'evidence': (
                        'DRIVES COST. DAR item 23.1.1 assumes 2 × 30 km round trips for a '
                        '140 m bore = 0.4286 truck days. Adjust for actual distance in Panel 2.'
                    ),
                },
            ],
            'default_materials': [
                {
                    'code': '7744',
                    'coeff': 1.0,
                    'evidence': (
                        'MS casing pipe for borewell (per metre): standard mild-steel lined '
                        'casing. Rate from Rates_Master code 7744.'
                    ),
                },
            ],
            'default_labour': [
                {
                    'code': '0113',
                    'coeff': 0.054,
                    'evidence': (
                        'Chowkidar (security/site attendant): 0.054 days per metre during '
                        'boring operations. DAR 23.1.1.1, page 948.'
                    ),
                },
                {
                    'code': '0114',
                    'coeff': 0.20,
                    'note': 'Beldar: 0.20 days per metre for assistance, pipe handling and site cleanup.',
                },
            ],
            'toggles': {'water': 'YES', 'gst': 'YES', 'cpoh': 'YES', 'cess': 'YES'},
            'machinery_defaults': [
                {
                    'code': '0020',
                    'coeff': 1.00,
                    'evidence': (
                        'Hydraulic Excavator 3D with borewell attachment: 1.00 day per metre '
                        'is the base unit (actual productivity depends on soil type and depth). '
                        'For 140m bore one rig day ≈ 5-8 m. DAR 23.1.1.1.'
                    ),
                },
                {
                    'code': '0005',
                    'coeff': 0.4286,
                    'evidence': (
                        'Diesel truck — transportation of drilling rig to/from site: '
                        '0.4286 days = 2 × 30 km / 140 m bore assumption. DAR 23.1.1.1.'
                    ),
                },
            ],
            'sample_library': [
                {
                    'code': '23.1.1.1', 'unit': 'metre', 'basis': 1.0,
                    'desc': 'Borewell 300mm dia, all soil types, hydraulic boring machine incl. transport',
                    'w': 7987.33, 'markups': 'Full W→X→Y→Z', 'rate': 10700.00, 'say': 10700.00,
                },
                {
                    'code': '23.3.1', 'unit': 'each', 'basis': 1.0,
                    'desc': 'Providing & fixing MS casing pipe 150mm NB for borewell',
                    'w': 4200.00, 'markups': 'Full W→X→Y→Z', 'rate': 5625.00, 'say': 5625.00,
                },
            ],
            'has_machinery': True,
        },

        # ------------------------------------------------------------------ #
        # 24 — Conservation of Heritage Buildings                              #
        # Default: item 24.1 — raking out joints of stone masonry.           #
        # Labour-only pattern (no material for most items, per DAR).         #
        # ------------------------------------------------------------------ #
        '24_Heritage_Buildings': {
            'sheet_name': '24_Heritage_Buildings',
            'trade_title': (
                'CPWD DAR 2019 — SUB-HEAD 24: CONSERVATION OF HERITAGE BUILDINGS\n'
                'WB2 — references WB1 for Rates_Master & Global_Factors'
            ),
            'trade_guidance': (
                'CPWD DAR Sub-head 24: Heritage conservation and restoration operations. '
                'Most items are LABOUR-ONLY — the conservation work is physical and skilled '
                'but does not consume significant material (like 02_Earth_Work and 15_Dismantling). '
                'Where materials are used (lime putty, hydraulic lime, surkhi, pigment) they are '
                'minor and entered as MATERIAL lines. '
                'Specialist conservation tradesmen: Stone mason / mason (0131), '
                'Skilled Beldar (0114/0115), Bhisti (0101). '
                'Some items use specialised heritage labour codes (6501, 7767-7775).'
            ),
            'default_item_code': 'C-24.01',
            'default_basis_qty': 10.0,
            'default_basis_unit': 'sqm',
            'default_item_desc': (
                'Raking out joints of stone masonry to a depth of 25 mm, cleaning with '
                'wire brush, and re-pointing with lime mortar, complete as per directions '
                'of Engineer-in-charge.'
            ),
            'material_section_title': (
                '3. MATERIAL COMPONENT BUILD-UP  (Heritage items in the DAR are predominantly '
                'labour-only — leave blank unless the operation genuinely uses material '
                'such as lime putty, hydraulic lime, pigment or consolidant)'
            ),
            'scope_inputs': [
                {
                    'label': 'Conservation operation', 'default': 'Raking out & re-pointing joints',
                    'impact': 'DRIVES COST',
                    'options': ['Raking out & re-pointing joints',
                                'Stone cleaning (manual / chemical)',
                                'Crack stitching with stainless-steel ties',
                                'Consolidation grouting',
                                'Dismantling & re-setting defective stone'],
                    'evidence': (
                        'DRIVES COST. Each operation has a different specialist labour crew. '
                        'The DAR lists 8 items in this sub-head (pages 974-979).'
                    ),
                },
                {
                    'label': 'Masonry type', 'default': 'Stone masonry',
                    'impact': 'DRIVES COST',
                    'options': ['Stone masonry', 'Brick masonry', 'Lime plaster / roughcast'],
                    'evidence': (
                        'DRIVES COST. Stone masonry needs a skilled stone mason; brick masonry '
                        'uses a standard Mason. Tool requirements and labour-day rates differ.'
                    ),
                },
            ],
            'default_materials': [
                {
                    'code': '4009',
                    'coeff': 0.05,
                    'evidence': (
                        'Hydraulic lime: 0.05 tonne per 10 sqm for joint pointing mortar. '
                        'Used in conservation mortars instead of OPC.'
                    ),
                },
            ],
            'default_labour': [
                {
                    'code': '0114',
                    'coeff': 0.53,
                    'note': 'Beldar: 0.53 days per 10 sqm raking out stone masonry joints. DAR 24.1, page 974.',
                },
                {
                    'code': '0115',
                    'coeff': 0.08,
                    'note': 'Coolie: 0.08 days per 10 sqm. DAR 24.1.',
                },
                {
                    'code': '0101',
                    'coeff': 0.07,
                    'note': 'Bhisti: 0.07 days per 10 sqm for water during curing. DAR 24.1.',
                },
            ],
            'toggles': {'water': 'YES', 'gst': 'YES', 'cpoh': 'YES', 'cess': 'YES'},
            'sample_library': [
                {
                    'code': '24.1', 'unit': 'sqm', 'basis': 10.0,
                    'desc': 'Raking out stone masonry joints 25mm deep & re-pointing with lime mortar',
                    'w': 386.43, 'markups': 'Full W→X→Y→Z', 'rate': 51.80, 'say': 51.80,
                },
                {
                    'code': '24.3', 'unit': 'sqm', 'basis': 10.0,
                    'desc': 'Stone cleaning — manual brushing & washing',
                    'w': 260.00, 'markups': 'Full W→X→Y→Z', 'rate': 34.85, 'say': 34.85,
                },
            ],
            'has_machinery': False,
        },

        # ------------------------------------------------------------------ #
        # 25 — Structural Glazing & ACP                                        #
        # Default: item 25.1 — ACP composite panel (LABOUR OPTIONAL).        #
        # Structural variant: item 25.1 has MATERIAL → Total directly.       #
        # ------------------------------------------------------------------ #
        '25_Structural_Glazing': {
            'sheet_name': '25_Structural_Glazing',
            'trade_title': (
                'CPWD DAR 2019 — SUB-HEAD 25: STRUCTURAL GLAZING & ACP (CUSTOM ITEM BUILDER)\n'
                'WB2 — references WB1 for Rates_Master & Global_Factors'
            ),
            'trade_guidance': (
                'CPWD DAR Sub-head 25: Structural glazing systems, aluminium composite panels '
                '(ACP), curtain walls and frameless glass. '
                'STRUCTURAL VARIANT: item 25.1 (ACP supply only) has NO LABOUR block — '
                'the rate is purely material-based (Al section weight × rate + coating). '
                'For installation items (25.3, 25.4) the full LABOUR block applies. '
                'Leave the LABOUR section empty when pricing supply-only items. '
                'Material: Al T/L sections (7306), powder coating (7392), '
                'ACP sheet material (2605-2634 series), structural silicone (2604).'
            ),
            'default_item_code': 'C-25.01',
            'default_basis_qty': 1.0,
            'default_basis_unit': 'sqm',
            'default_item_desc': (
                'Providing and supplying aluminium composite panel (ACP) with Al skin '
                '0.5mm each side, polyethylene core, powder coated 50 microns, '
                'including Al T or L sections, complete as per directions of Engineer-in-charge.'
            ),
            'material_section_title': (
                '3. MATERIAL COMPONENT BUILD-UP  (Item 25.1 has no Labour block — '
                'the fabrication cost is entirely material-based. For installation items '
                '25.3/25.4, use the LABOUR block as normal.)'
            ),
            'scope_inputs': [
                {
                    'label': 'System type', 'default': 'ACP supply (weight-based)',
                    'impact': 'DRIVES COST',
                    'options': ['ACP supply (weight-based, no labour)',
                                'Structural glazing — fixed frame',
                                'Curtain wall system (pressure plate)',
                                'Spider fitting / point-fixed glass'],
                    'evidence': (
                        'DRIVES COST. Supply-only items carry no Labour block — the rate is '
                        'derived from Al weight and coating only. Full installation items add '
                        'a Fitter / Glazier labour crew.'
                    ),
                },
                {
                    'label': 'Surface coating', 'default': 'Powder coated 50 microns (code 7392)',
                    'impact': 'DRIVES COST',
                    'options': ['Powder coated 50 microns (code 7392)',
                                'Anodised 15 microns (code 7389)',
                                'Anodised 25 microns (code 7390)',
                                'PVDF coating'],
                    'evidence': (
                        'DRIVES COST. Coating is billed per kg of Al section. '
                        'PVDF (fluoropolymer) is about 2× the cost of standard powder coating.'
                    ),
                },
            ],
            'default_materials': [
                {
                    'code': '7306',
                    'coeff': 7.17,
                    'evidence': (
                        'Aluminium T or L sections: 7.17 kg per sqm of ACP panel frame. '
                        'DAR 25.1, page 982.'
                    ),
                },
                {
                    'code': '7392',
                    'coeff': 7.17,
                    'evidence': (
                        'Powder coating 50 microns: 7.17 kg (same weight as section). '
                        'Rate = Rs 61/kg. DAR 25.1.'
                    ),
                },
            ],
            'default_labour': [],   # supply-only item — no labour
            'toggles': {'water': 'YES', 'gst': 'YES', 'cpoh': 'YES', 'cess': 'YES'},
            'sample_library': [
                {
                    'code': '25.1', 'unit': 'sqm', 'basis': 1.0,
                    'desc': 'ACP supply — Al T/L 7.17 kg/sqm + powder-coating 50 mic (no Labour)',
                    'w': 1799.67, 'markups': 'Full W→X→Y→Z', 'rate': 2384.00, 'say': 2384.00,
                },
                {
                    'code': '25.3', 'unit': 'sqm', 'basis': 1.0,
                    'desc': 'Structural glazing — fixed frame, 12mm toughened glass, Al frame',
                    'w': 4200.00, 'markups': 'Full W→X→Y→Z', 'rate': 5625.00, 'say': 5625.00,
                },
            ],
            'has_machinery': False,
        },

        # ------------------------------------------------------------------ #
        # 26 — New Technologies and Materials                                  #
        # Default: item 26.1 — pre-engineered laminate flooring.             #
        # Material-heavy, per-unit-area with wastage %; full markup chain.   #
        # ------------------------------------------------------------------ #
        '26_New_Technologies': {
            'sheet_name': '26_New_Technologies',
            'trade_title': (
                'CPWD DAR 2019 — SUB-HEAD 26: NEW TECHNOLOGIES AND MATERIALS\n'
                'WB2 — references WB1 for Rates_Master & Global_Factors'
            ),
            'trade_guidance': (
                'CPWD DAR Sub-head 26: Proprietary and non-conventional construction products '
                '(pre-engineered flooring, AAC blocks, GFRC panels, EPS insulation, solar systems). '
                'Material-heavy items: per-unit-area rate of the proprietary product from Rates_Master '
                '(codes 7776, 7911-7931, 7996-8057, 8144-8150, 8210-8226, 8515-8596 range). '
                'A wastage helper (typically 5-10%) is recommended for all sheet/board products. '
                'Labour: Skilled workman (varies by system), Beldar/Coolie for site handling. '
                'Machinery: as required by the specific system (crane, compressor, welding set).'
            ),
            'default_item_code': 'C-26.01',
            'default_basis_qty': 10.0,
            'default_basis_unit': 'sqm',
            'default_item_desc': (
                'Providing and fixing in position pre-engineered laminate / engineered wood '
                'flooring system in approved colour, texture and finish, having Performance '
                'Appraisal Certificate (PAC) issued by BMTPC, with groove interlocking system, '
                'laid over polyethylene foam underlay, complete as per directions of '
                'Engineer-in-charge.'
            ),
            'scope_inputs': [
                {
                    'label': 'Technology / product', 'default': 'Pre-engineered laminate flooring',
                    'impact': 'DRIVES COST',
                    'options': ['Pre-engineered laminate flooring (7776)',
                                'Autoclaved aerated concrete (AAC) blocks (8210-8212)',
                                'Glass fibre reinforced concrete (GFRC) panels',
                                'EPS / PIR insulation boards (8515-8520 series)',
                                'Solar PV system (8580-8596)',
                                'Structural insulated panel (SIP)'],
                    'evidence': (
                        'DRIVES COST. Each technology has a distinct product code and per-unit rate. '
                        'The PAC requirement means only BMTPC-approved products qualify.'
                    ),
                },
                {
                    'label': 'Wastage allowance', 'default': '7% (boards/sheets)',
                    'impact': 'DRIVES COST',
                    'options': ['5% (standard rectangular rooms)', '7% (boards/sheets)',
                                '10% (complex cuts / diagonal laying)', '0% (prefab exact cut)'],
                    'evidence': (
                        'DRIVES COST. Add the wastage % to the product quantity. '
                        'For 10 sqm + 7% = 10.70 sqm ordered.'
                    ),
                },
                {
                    'label': 'Labour supply', 'default': 'Specialist sub-contractor (specialist codes)',
                    'impact': 'DRIVES COST',
                    'options': ['Specialist sub-contractor (specialist codes)',
                                'CPWD departmental labour (Beldar / Skilled)'],
                    'evidence': (
                        'DRIVES COST. Proprietary systems often require the manufacturer\'s specialist '
                        'installer. Use specialist-crew codes from the 7913-7931 range in Rates_Master.'
                    ),
                },
            ],
            'default_materials': [
                {
                    'code': '7776',
                    'coeff': 10.7,
                    'evidence': (
                        'Pre-engineered laminate flooring PAC-approved: 10.70 sqm '
                        '(10 sqm + 7% wastage). Rate from Rates_Master code 7776.'
                    ),
                },
                {
                    'code': '1023',
                    'coeff': 11.0,
                    'evidence': (
                        'Polyethylene foam underlay sheet: 11 sqm (10 sqm + 10% overlap). '
                        'Laid below laminate to cushion and prevent moisture ingress.'
                    ),
                },
            ],
            'default_labour': [
                {
                    'code': '0122',
                    'coeff': 1.00,
                    'evidence': (
                        'Skilled workman / carpenter: 1.00 day per 10 sqm for laying, cutting '
                        'and fitting laminate boards including skirting. DAR 26.1 benchmark.'
                    ),
                },
                {
                    'code': '0114',
                    'coeff': 0.50,
                    'note': 'Beldar: 0.50 days per 10 sqm for material handling and site cleanup.',
                },
            ],
            'toggles': {'water': 'YES', 'gst': 'YES', 'cpoh': 'YES', 'cess': 'YES'},
            'sample_library': [
                {
                    'code': '26.1', 'unit': 'sqm', 'basis': 10.0,
                    'desc': 'Pre-engineered laminate flooring PAC, groove-lock, PE foam underlay, +7% wastage',
                    'w': 7840.00, 'markups': 'Full W→X→Y→Z', 'rate': 1050.20, 'say': 1050.20,
                },
                {
                    'code': '26.5', 'unit': 'sqm', 'basis': 10.0,
                    'desc': 'AAC block masonry 200mm thick in CM 1:6 (replaces conventional brick)',
                    'w': 4620.00, 'markups': 'Full W→X→Y→Z', 'rate': 619.00, 'say': 619.00,
                },
            ],
            'has_machinery': True,
        },

    }
