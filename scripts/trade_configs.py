# -*- coding: utf-8 -*-

def get_all_trade_configs():
    return {
        '01_Carriage_of_Materials': {
            'sheet_name': '01_Carriage_of_Materials',
            'trade_title': 'CPWD DAR 2019 — SUB-HEAD 01: CARRIAGE OF MATERIALS (ANALYTICAL SIMULATOR & BENCHMARKS)',
            'trade_guidance': 'CPWD DAR Sub-head 01: Dynamic Mechanical Carriage Simulator based on Data Sheet No. 1 (Notes 1 to 5). Supports custom lead (L), truck speed (S), turnaround time (T), and payload capacities (C). Evaluates trips N, km done, dynamic fuel consumption, and net unit rates with 15% CPOH.',
            'default_item_code': '1.1.CUSTOM',
            'default_basis_qty': 10.98,
            'default_basis_unit': 'metre',
            'default_material': 'R.C.C./C.I./Steel pipes 1000, 1100 & 1200 mm dia',
            'default_scope': 'including loading, transporting, unloading and stacking',
            'default_lift': 'for all lifts',
            'default_item_desc': 'Transport of 1000, 1100 & 1200 mm dia R.C.C./C.I./Steel cylinder pipes by mechanical transport including loading, transporting, unloading to approved municipal dumping ground/site for lead upto 26 km for all lifts complete as per directions of Engineer-in-charge.',
            'default_lead': 26.0,
            'default_speed': 29.0,
            'default_turnaround': 1.0,
            'default_capacity': 10.98,
            'sample_library': [
                {'code': '1.1.18', 'desc': 'Disposal of building malba by mechanical transport - lead 10 km (Restricted urban 3 trips, 8 cum/trip)', 'unit': 'cum', 'basis': 24.0, 'w': 5939.58, 'markups': '15% CPOH only', 'rate': 284.60, 'say': 284.60},
                {'code': '1.1.17.12-VAR', 'desc': 'Transport of 1000, 1100 & 1200 mm dia pipes - lead 26 km (N=2.86 trips, Speed 29 km/h, Payload 10.98 m/trip)', 'unit': 'metre', 'basis': 31.40, 'w': 7470.17, 'markups': '15% CPOH only', 'rate': 273.56, 'say': 273.60},
                {'code': '1.1.1', 'desc': 'Carriage of Lime, moorum, building rubbish by mechanical transport - lead 5 km (N=5.19 trips, 8 cum/trip)', 'unit': 'cum', 'basis': 41.52, 'w': 5829.54, 'markups': '15% CPOH only', 'rate': 161.46, 'say': 161.50}
            ]
        },
        '02_Earth_Work': {
            'sheet_name': '02_Earth_Work',
            'default_materials': [],
            'material_section_title': ('3. MATERIAL COMPONENT BUILD-UP  (Earth Work items in the DAR are '
                                      'labour-and-plant only - leave blank unless the operation genuinely '
                                      'consumes material, e.g. imported earth or moorum)'),
            'trade_title': 'CPWD DAR 2019 — SUB-HEAD 02: EARTH WORK (CUSTOM ITEM BUILDER)',
            'trade_guidance': 'CPWD DAR Sub-head 02: Earthwork operations (surface excavation, trench excavation, filling, compaction). Operates without raw material inputs, applying labour crews and earthmoving machinery hire with standard 5-step compounding statutory markups.',
            'default_item_code': 'C-02.01',
            'default_basis_qty': 100.0,
            'default_basis_unit': 'sqm',
            'default_item_desc': 'Earth work in surface excavation not exceeding 30 cm in depth, exceeding 1.5m width and 10 sqm plan, getting out and disposal upto 50m lead and 1.5m lift in all kinds of soil',
            'default_labour': [
                {'code': '0114', 'coeff': 6.80, 'note': 'Beldar (excavation & throwing)'},
                {'code': '0115', 'coeff': 5.60, 'note': 'Coolie (carrying & disposal upto 50m)'}
            ],
            'default_sundries_base': 0.0,
            'toggles': {'water': 'YES', 'gst': 'YES', 'cpoh': 'YES', 'cess': 'YES'},
            'toggle_notes': {
                'water': 'CPWD Standard: 1% water charges on labour cost for site compaction water and dust suppression.',
                'gst': 'CPWD DAR multiplying factor 0.1405 (works contract tax).',
                'cpoh': 'Standard 15% Contractor Profit & Overheads.',
                'cess': 'Statutory 1% BOCW Welfare Cess.'
            },
            'sample_library': [
                {'code': '2.1.1', 'desc': 'Earth work in surface excavation in all kinds of soil upto 50m lead and 1.5m lift', 'unit': 'sqm', 'basis': 100.0, 'w': 6919.20, 'markups': 'Full W->X->Y->Z', 'rate': 92.57, 'say': 92.55},
                {'code': '2.2.1', 'desc': 'Earth work in rough excavation, banking excavated earth in layers with watering & rolling', 'unit': 'cum', 'basis': 10.0, 'w': 5581.72, 'markups': 'Full W->X->Y->Z', 'rate': 746.80, 'say': 746.80}
            ]
        },
        '03_Mortars': {
            'sheet_name': '03_Mortars',
            'trade_title': 'CPWD DAR 2019 — SUB-HEAD 03: MORTARS (INTERMEDIATE MIX BUILDER)',
            'trade_guidance': 'CPWD DAR Sub-head 03: Mortars serve as intermediate building blocks for Masonry, Stone, Cladding, and Flooring. Direct Cost (W) equals Final Cost. Markup chain is turned OFF by default to avoid double-taxation and double-profit in consuming trades.',
            'default_item_code': 'C-03.01',
            'default_basis_qty': 1.0,
            'default_basis_unit': 'cum',
            'default_item_desc': 'Cement mortar 1:4 (1 cement : 4 coarse sand) — Intermediate mix for downstream trade consumption',
            'default_materials': [
                {'code': '0367', 'coeff': 0.38, 'note': 'Portland Cement (0.269 cum = 0.38 tonne)'},
                {'code': '2209', 'coeff': 0.38, 'note': 'Carriage of Cement'},
                {'code': '0982', 'coeff': 1.07, 'note': 'Coarse sand (zone III)'},
                {'code': '2203', 'coeff': 1.07, 'note': 'Carriage of Coarse sand'}
            ],
            'default_labour': [
                {'code': '0114', 'coeff': 0.75, 'note': 'Beldar (measuring, carrying & mixing)'},
                {'code': '0101', 'coeff': 0.07, 'note': 'Bhisti (watering)'}
            ],
            'default_sundries_base': 13.52,
            'toggles': {'water': 'NO', 'gst': 'NO', 'cpoh': 'NO', 'cess': 'NO'},
            'toggle_notes': {
                'water': 'CPWD RULE: 0% markup default. Water is not added here as it is marked up once in Masonry/Flooring.',
                'gst': 'CPWD RULE: GST is not added here to prevent double taxation when imported into Masonry.',
                'cpoh': 'CPWD RULE: Contractor profit is not added to intermediate mortar. Consuming trades add 15% CPOH over combined total.',
                'cess': 'CPWD RULE: 0% Cess here; applied once at the end of the consuming trade item.'
            },
            'sample_library': [
                {'code': '3.1', 'desc': 'Cement mortar 1:1 (1 cement : 1 fine sand)', 'unit': 'cum', 'basis': 1.0, 'w': 6390.60, 'markups': 'No markups (base cost)', 'rate': 6390.60, 'say': 6390.60},
                {'code': '3.9', 'desc': 'Cement mortar 1:4 (1 cement : 4 coarse sand)', 'unit': 'cum', 'basis': 1.0, 'w': 4010.35, 'markups': 'No markups (base cost)', 'rate': 4010.35, 'say': 4010.35}
            ]
        },
        '04_Concrete_Work': {
            'sheet_name': '04_Concrete_Work',
            'trade_title': 'CPWD DAR 2019 — SUB-HEAD 04: CONCRETE WORK (CUSTOM ITEM BUILDER)',
            'trade_guidance': 'CPWD DAR Sub-head 04: Plain cement concrete. Built from coarse aggregate, fine aggregate, cement, carriage lines, mixing & placing labour, mixer and vibrator hire. Full 5-step compounding markup chain applies.',
            'default_item_code': 'C-04.01',
            'default_basis_qty': 1.0,
            'default_basis_unit': 'cum',
            'default_item_desc': 'Cement concrete 1:2:4 (1 cement : 2 coarse sand : 4 graded stone aggregate 20 mm nominal size) up to plinth level',
            'default_materials': [
                {'code': '0295', 'coeff': 0.57, 'note': 'Stone Aggregate 20 mm nominal size'},
                {'code': '0297', 'coeff': 0.28, 'note': 'Stone Aggregate 10 mm nominal size'},
                {'code': '2202', 'coeff': 0.85, 'note': 'Carriage of Stone aggregate below 40 mm'},
                {'code': '0982', 'coeff': 0.425, 'note': 'Coarse sand (zone III)'},
                {'code': '2203', 'coeff': 0.425, 'note': 'Carriage of Coarse sand'},
                {'code': '0367', 'coeff': 0.32, 'note': 'Portland Cement'},
                {'code': '2209', 'coeff': 0.32, 'note': 'Carriage of Cement'}
            ],
            'default_labour': [
                {'code': '0155', 'coeff': 0.10, 'note': 'Mason (average)'},
                {'code': '0114', 'coeff': 1.63, 'note': 'Beldar (carrying & placing)'},
                {'code': '0101', 'coeff': 0.70, 'note': 'Bhisti (curing)'},
                {'code': '0002', 'coeff': 0.07, 'note': 'Hire charges of Concrete Mixer with hopper'},
                {'code': '0012', 'coeff': 0.07, 'note': 'Vibrator (Needle type 40mm)'}
            ],
            'default_sundries_base': 14.30,
            'toggles': {'water': 'YES', 'gst': 'YES', 'cpoh': 'YES', 'cess': 'YES'},
            'toggle_notes': {
                'water': 'CPWD Standard: 1% for curing & concrete water.',
                'gst': 'CPWD DAR factor 0.1405 (works contract tax).',
                'cpoh': 'Standard 15% Contractor Profit & Overheads.',
                'cess': 'Statutory 1% BOCW Welfare Cess.'
            },
            'sample_library': [
                {'code': '4.1.2', 'desc': 'PCC 1:1.5:3 with 20mm graded stone aggregate up to plinth level', 'unit': 'cum', 'basis': 1.0, 'w': 5389.89, 'markups': 'Full W->X->Y->Z', 'rate': 7210.58, 'say': 7210.55},
                {'code': '4.1.3', 'desc': 'PCC 1:2:4 with 20mm graded stone aggregate up to plinth level', 'unit': 'cum', 'basis': 1.0, 'w': 5074.69, 'markups': 'Full W->X->Y->Z', 'rate': 6788.62, 'say': 6788.60}
            ]
        },
        '05_RCC_Work': {
            'sheet_name': '05_RCC_Work',
            'trade_title': 'CPWD DAR 2019 — SUB-HEAD 05: REINFORCED CEMENT CONCRETE (RCC) BUILDER',
            'trade_guidance': 'CPWD DAR Sub-head 05: Reinforced cement concrete in plinth beams, columns, and slabs. Excludes centering, shuttering, and reinforcement which are measured and priced separately per CPWD specifications.',
            'default_item_code': 'C-05.01',
            'default_basis_qty': 1.0,
            'default_basis_unit': 'cum',
            'default_item_desc': 'Reinforced cement concrete 1:1.5:3 (1 cement : 1.5 coarse sand : 3 graded stone aggregate 20mm nominal size) up to plinth level',
            'default_materials': [
                {'code': '0295', 'coeff': 0.57, 'note': 'Stone Aggregate 20 mm nominal size'},
                {'code': '0297', 'coeff': 0.28, 'note': 'Stone Aggregate 10 mm nominal size'},
                {'code': '2202', 'coeff': 0.85, 'note': 'Carriage of Stone aggregate'},
                {'code': '0982', 'coeff': 0.425, 'note': 'Coarse sand (zone III)'},
                {'code': '2203', 'coeff': 0.425, 'note': 'Carriage of Coarse sand'},
                {'code': '0367', 'coeff': 0.40, 'note': 'Portland Cement'},
                {'code': '2209', 'coeff': 0.40, 'note': 'Carriage of Cement'}
            ],
            'default_labour': [
                {'code': '0155', 'coeff': 0.17, 'note': 'Mason (average)'},
                {'code': '0114', 'coeff': 2.00, 'note': 'Beldar (mixing & placing)'},
                {'code': '0101', 'coeff': 0.90, 'note': 'Bhisti (curing)'},
                {'code': '0002', 'coeff': 0.07, 'note': 'Hire charges of Concrete Mixer'},
                {'code': '0012', 'coeff': 0.07, 'note': 'Needle Vibrator 40mm'}
            ],
            'default_sundries_base': 14.30,
            'toggles': {'water': 'YES', 'gst': 'YES', 'cpoh': 'YES', 'cess': 'YES'},
            'toggle_notes': {
                'water': 'CPWD Standard: 1% for curing & site water.',
                'gst': 'CPWD DAR factor 0.1405 (works contract tax).',
                'cpoh': 'Standard 15% Contractor Profit & Overheads.',
                'cess': 'Statutory 1% BOCW Welfare Cess.'
            },
            'sample_library': [
                {'code': '5.1.2', 'desc': 'RCC 1:1.5:3 up to plinth level (excl. shuttering & reinforcement)', 'unit': 'cum', 'basis': 1.0, 'w': 5768.78, 'markups': 'Full W->X->Y->Z', 'rate': 7717.15, 'say': 7717.15},
                {'code': '5.1.3', 'desc': 'RCC 1:2:4 up to plinth level (excl. shuttering & reinforcement)', 'unit': 'cum', 'basis': 1.0, 'w': 5373.98, 'markups': 'Full W->X->Y->Z', 'rate': 7189.02, 'say': 7189.00}
            ]
        },
        '06_Masonry_Work': {
            'sheet_name': '06_Masonry_Work',
            'trade_title': 'CPWD DAR 2019 — SUB-HEAD 06: MASONRY WORK (CUSTOM ITEM BUILDER)',
            'trade_guidance': 'CPWD DAR Sub-head 06: Brick and block masonry. Imports un-marked Mortar rates from 03_Mortars as a direct material input, then applies the full statutory markup chain once over the combined total.',
            'default_item_code': 'C-06.01',
            'default_basis_qty': 1.0,
            'default_basis_unit': 'cum',
            'default_item_desc': 'Brick work with common burnt clay F.P.S. bricks class 7.5 in foundation and plinth in cement mortar 1:4 (1 cement : 4 coarse sand)',
            'default_materials': [
                {'code': '2602', 'coeff': 0.494, 'note': 'Common burnt clay F.P.S. bricks class 7.5 (thousand)'},
                {'code': '2201', 'coeff': 0.494, 'note': 'Carriage of Bricks (thousand)'},
                {
                    'code': '03_Mortars',
                    'custom_desc': 'Cement mortar 1:4 (Rate pulled from 03_Mortars builder or DAR Item 3.9)',
                    'custom_unit': 'cum',
                    'coeff': 0.25,
                    'custom_rate_formula': "='03_Mortars'!G66",
                    'note': 'Live link to 03_Mortars output rate'
                }
            ],
            'default_labour': [
                {'code': '0123', 'coeff': 0.36, 'note': 'Mason (brick layer) 1st class'},
                {'code': '0124', 'coeff': 0.36, 'note': 'Mason (brick layer) 2nd class'},
                {'code': '0115', 'coeff': 1.37, 'note': 'Coolie'},
                {'code': '0101', 'coeff': 0.20, 'note': 'Bhisti'}
            ],
            'default_sundries_base': 2.73,
            'toggles': {'water': 'YES', 'gst': 'YES', 'cpoh': 'YES', 'cess': 'YES'},
            'toggle_notes': {
                'water': 'CPWD Standard: 1% water charges on combined material and labour total.',
                'gst': 'CPWD DAR factor 0.1405 (works contract tax).',
                'cpoh': 'Standard 15% Contractor Profit & Overheads.',
                'cess': 'Statutory 1% BOCW Welfare Cess.'
            },
            'sample_library': [
                {'code': '6.1.1', 'desc': 'Brick work class 7.5 in foundation and plinth in CM 1:4', 'unit': 'cum', 'basis': 1.0, 'w': 4765.73, 'markups': 'Full W->X->Y->Z', 'rate': 6375.31, 'say': 6375.30},
                {'code': '6.1.2', 'desc': 'Brick work class 7.5 in foundation and plinth in CM 1:6', 'unit': 'cum', 'basis': 1.0, 'w': 4539.06, 'markups': 'Full W->X->Y->Z', 'rate': 6072.07, 'say': 6072.05}
            ]
        },
        '07_Stone_Work': {
            'sheet_name': '07_Stone_Work',
            'trade_title': 'CPWD DAR 2019 — SUB-HEAD 07: STONE WORK (CUSTOM ITEM BUILDER)',
            'trade_guidance': 'CPWD DAR Sub-head 07: Random rubble and ashlar stone masonry. Consumes un-marked Mortar from 03_Mortars, quarry stone, bond stones, dressing labour, and masons.',
            'default_item_code': 'C-07.01',
            'default_basis_qty': 1.0,
            'default_basis_unit': 'cum',
            'default_item_desc': 'Random rubble masonry with hard stone in foundation and plinth in cement mortar 1:6 (1 cement : 6 coarse sand)',
            'default_materials': [
                {'code': '0299', 'coeff': 1.00, 'note': 'Stone for masonry (rubble)'},
                {'code': '2202', 'coeff': 1.00, 'note': 'Carriage of Stone'},
                {'code': '0300', 'coeff': 0.10, 'note': 'Through or bond stone'},
                {
                    'code': '03_Mortars',
                    'custom_desc': 'Cement mortar 1:6 (Rate from 03_Mortars or DAR 3.11)',
                    'custom_unit': 'cum',
                    'coeff': 0.30,
                    'custom_rate_formula': "='03_Mortars'!G66",
                    'note': 'Mortar bedding rate'
                }
            ],
            'default_labour': [
                {'code': '0155', 'coeff': 0.70, 'note': 'Stone Mason (average)'},
                {'code': '0115', 'coeff': 1.40, 'note': 'Coolie'},
                {'code': '0101', 'coeff': 0.20, 'note': 'Bhisti'}
            ],
            'default_sundries_base': 3.50,
            'toggles': {'water': 'YES', 'gst': 'YES', 'cpoh': 'YES', 'cess': 'YES'},
            'toggle_notes': {
                'water': 'Standard 1% Water charges.',
                'gst': 'CPWD factor 0.1405.',
                'cpoh': '15% Contractor Profit & Overheads.',
                'cess': '1% Labour Welfare Cess.'
            },
            'sample_library': [
                {'code': '7.1.1', 'desc': 'Random rubble masonry with hard stone in CM 1:6 in foundation and plinth', 'unit': 'cum', 'basis': 1.0, 'w': 4150.20, 'markups': 'Full W->X->Y->Z', 'rate': 5552.00, 'say': 5552.00}
            ]
        },
        '08_Cladding_Work': {
            'sheet_name': '08_Cladding_Work',
            'trade_title': 'CPWD DAR 2019 — SUB-HEAD 08: CLADDING WORK (CUSTOM ITEM BUILDER)',
            'trade_guidance': 'CPWD DAR Sub-head 08: Stone and marble veneer cladding. Incorporates bedding mortar and pointing mortar lines, slab cutting/wastage, scaffolding, and stone masons.',
            'default_item_code': 'C-08.01',
            'default_basis_qty': 10.0,
            'default_basis_unit': 'sqm',
            'default_item_desc': '30 mm thick gang saw cut red sand stone / white sand stone sun-shade / cladding in wall with CM 1:3',
            'default_materials': [
                {'code': '0308', 'coeff': 10.50, 'note': 'Red sand stone slab 30mm thick (with 5% cutting wastage)'},
                {'code': '2202', 'coeff': 0.315, 'note': 'Carriage of Stone slabs'},
                {
                    'code': '03_Mortars',
                    'custom_desc': 'Cement mortar 1:3 (Bedding mortar from 03_Mortars or DAR 3.3)',
                    'custom_unit': 'cum',
                    'coeff': 0.20,
                    'custom_rate_formula': "='03_Mortars'!G66",
                    'note': 'Bedding mortar'
                }
            ],
            'default_labour': [
                {'code': '0155', 'coeff': 1.50, 'note': 'Stone Mason (fixing & pointing)'},
                {'code': '0115', 'coeff': 1.80, 'note': 'Coolie'},
                {'code': '0101', 'coeff': 0.30, 'note': 'Bhisti'}
            ],
            'default_sundries_base': 12.00,
            'toggles': {'water': 'YES', 'gst': 'YES', 'cpoh': 'YES', 'cess': 'YES'},
            'toggle_notes': {
                'water': 'Standard 1% Water charges.',
                'gst': 'CPWD factor 0.1405.',
                'cpoh': '15% Contractor Profit & Overheads.',
                'cess': '1% Labour Welfare Cess.'
            },
            'sample_library': [
                {'code': '8.1.1', 'desc': 'Red sand stone cladding 30mm thick with backing mortar', 'unit': 'sqm', 'basis': 10.0, 'w': 11250.00, 'markups': 'Full W->X->Y->Z', 'rate': 1505.20, 'say': 1505.20}
            ]
        },
        '09_Wood_and_PVC_Work': {
            'sheet_name': '09_Wood_and_PVC_Work',
            'trade_title': 'CPWD DAR 2019 — SUB-HEAD 09: WOOD & PVC WORK (CUSTOM ITEM BUILDER)',
            'trade_guidance': 'CPWD DAR Sub-head 09: Timber frames, shutters, and fittings. Built from timber scantling (with standard 5% wastage allowance), carriage, holdfasts, and carpenter crews.',
            'default_item_code': 'C-09.01',
            'default_basis_qty': 36.0,
            'default_basis_unit': 'cudm',
            'default_item_desc': 'Second class teak wood in frames of doors, windows and other frames wrought, framed and fixed in position',
            'default_materials': [
                {'code': '1189', 'coeff': 3.80, 'note': 'Second class teak wood in scantling (36 cudm + 5% wastage = 38 cudm)'},
                {'code': '2204', 'coeff': 0.038, 'note': 'Carriage of Timber'}
            ],
            'default_labour': [
                {'code': '0156', 'coeff': 0.72, 'note': 'Carpenter (average)'},
                {'code': '0114', 'coeff': 0.07, 'note': 'Beldar'}
            ],
            'default_sundries_base': 0.00,
            'toggles': {'water': 'YES', 'gst': 'YES', 'cpoh': 'YES', 'cess': 'YES'},
            'toggle_notes': {
                'water': 'Standard 1% Water charges.',
                'gst': 'CPWD factor 0.1405.',
                'cpoh': '15% Contractor Profit & Overheads.',
                'cess': '1% Labour Welfare Cess.'
            },
            'sample_library': [
                {'code': '9.1.1', 'desc': 'Second class teak wood in frames of doors and windows', 'unit': 'cum', 'basis': 36.0, 'w': 3502.85, 'markups': 'Full W->X->Y->Z', 'rate': 130183.06, 'say': 130183.05}
            ]
        },
        '10_Steel_Work': {
            'sheet_name': '10_Steel_Work',
            'trade_title': 'CPWD DAR 2019 — SUB-HEAD 10: STEEL WORK (CUSTOM ITEM BUILDER)',
            'trade_guidance': 'CPWD DAR Sub-head 10: Structural steel fabrication. Incorporates rolling sections (with 5% wastage), carriage, fitter/blacksmith crew, and priming coat (CPWD DAR 13.50.3 @ ₹50.70/sqm).',
            'default_item_code': 'C-10.01',
            'default_basis_qty': 1.0,
            'default_basis_unit': 'quintal',
            'default_item_desc': 'Structural steel work in single section, fixed with or without connecting plate, including cutting, hoisting, fixing and applying priming coat',
            'default_materials': [
                {'code': '1007', 'coeff': 1.05, 'note': 'Structurals (tees, angles, channels, joists) + 5% wastage = 1.05 q'},
                {'code': '2205', 'coeff': 0.105, 'note': 'Carriage of Steel (0.105 tonne)'},
                {
                    'code': '13.50.3',
                    'custom_desc': 'Red oxide zinc chromate primer on steel (Rate as per Item 13.50.3 of SH: Finishing)',
                    'custom_unit': 'sqm',
                    'coeff': 3.00,
                    'custom_rate': 50.70,
                    'note': 'Priming coat @ ₹50.70/sqm (Deducted from W before markup per CPWD W-A rule)'
                }
            ],
            'default_labour': [
                {'code': '0116', 'coeff': 0.50, 'note': 'Fitter (grade 1)'},
                {'code': '0103', 'coeff': 0.75, 'note': 'Blacksmith 2nd class'},
                {'code': '0114', 'coeff': 1.00, 'note': 'Beldar'}
            ],
            'default_sundries_base': 20.67,
            'toggles': {'water': 'YES', 'gst': 'YES', 'cpoh': 'YES', 'cess': 'YES'},
            'toggle_notes': {
                'water': 'Standard 1% Water charges applied on (W - Priming coat).',
                'gst': 'CPWD factor 0.1405 applied on (X - Priming coat).',
                'cpoh': '15% CPOH applied on (Y - Priming coat).',
                'cess': '1% Cess applied on (Z - Priming coat).'
            },
            'sample_library': [
                {'code': '10.1', 'desc': 'Structural steel work in single section fixed complete with priming coat', 'unit': 'quintal', 'basis': 1.0, 'w': 6469.38, 'markups': 'Full W->X->Y->Z (excl. primer)', 'rate': 8604.22, 'say': 8604.20}
            ]
        },
        '11_Flooring': {
            'sheet_name': '11_Flooring',
            'trade_title': 'CPWD DAR 2019 — SUB-HEAD 11: FLOORING (CUSTOM ITEM BUILDER)',
            'trade_guidance': 'CPWD DAR Sub-head 11: Tile, stone, marble, and concrete flooring. Consumes mortar bedding (from 03_Mortars), flooring tiles, neat cement slurry, and floor polishing masons.',
            'default_item_code': 'C-11.01',
            'default_basis_qty': 10.0,
            'default_basis_unit': 'sqm',
            'default_item_desc': 'Precast terrazzo tiles 20 mm thick with marble chips of size upto 6 mm laid in floors in CM 1:4 with neat cement slurry',
            'default_materials': [
                {'code': '0291', 'coeff': 10.00, 'note': 'Precast terrazzo tiles 20mm thick (sqm)'},
                {'code': '2201', 'coeff': 0.05, 'note': 'Carriage of tiles'},
                {
                    'code': '03_Mortars',
                    'custom_desc': 'Cement mortar 1:4 (Bedding mortar from 03_Mortars or DAR 3.9)',
                    'custom_unit': 'cum',
                    'coeff': 0.20,
                    'custom_rate_formula': "='03_Mortars'!G66",
                    'note': 'Bedding mortar'
                },
                {'code': '0367', 'coeff': 0.044, 'note': 'Cement for neat cement slurry (4.4 kg/sqm)'}
            ],
            'default_labour': [
                {'code': '0155', 'coeff': 1.20, 'note': 'Mason (average)'},
                {'code': '0115', 'coeff': 1.20, 'note': 'Coolie'},
                {'code': '0101', 'coeff': 0.20, 'note': 'Bhisti'},
                {'code': '0013', 'coeff': 0.30, 'note': 'Machine for rubbing & polishing of floors'}
            ],
            'default_sundries_base': 10.00,
            'toggles': {'water': 'YES', 'gst': 'YES', 'cpoh': 'YES', 'cess': 'YES'},
            'toggle_notes': {
                'water': 'Standard 1% Water charges.',
                'gst': 'CPWD factor 0.1405.',
                'cpoh': '15% Contractor Profit & Overheads.',
                'cess': '1% Labour Welfare Cess.'
            },
            'sample_library': [
                {'code': '11.3.1', 'desc': 'Precast terrazzo tiles 20 mm thick laid in floors in CM 1:4', 'unit': 'sqm', 'basis': 10.0, 'w': 7850.00, 'markups': 'Full W->X->Y->Z', 'rate': 1049.80, 'say': 1049.80}
            ]
        },
        '12_Roofing': {
            'sheet_name': '12_Roofing',
            'trade_title': 'CPWD DAR 2019 — SUB-HEAD 12: ROOFING (CUSTOM ITEM BUILDER)',
            'trade_guidance': 'CPWD DAR Sub-head 12: Sheet roofing and roof waterproofing. Includes corrugated G.S. sheets (with 5% lap wastage), J/L hooks, bolts, bitumen/limpet washers, and painter/carpenter labour.',
            'default_item_code': 'C-12.01',
            'default_basis_qty': 184.52,
            'default_basis_unit': 'sqm',
            'default_item_desc': 'Providing corrugated G.S. sheet roofing 1.00 mm thick with zinc coating 275 gm/sqm fixed with J/L hooks, bolts and limpet washers complete',
            'default_materials': [
                {'code': '3050', 'coeff': 23.26, 'note': 'Galvanised steel corrugated sheets (quintal)'},
                {'code': '2302', 'coeff': 2.326, 'note': 'Carriage of G.I. sheet and accessories (tonne)'},
                {'code': '1022', 'coeff': 88.40, 'note': 'Galvanised steel bolts & nuts 6mm dia, 25mm long (10 Nos)'},
                {'code': '1023', 'coeff': 81.00, 'note': 'Galvanised steel J or L hooks 8 mm dia (10 Nos)'},
                {'code': '1207', 'coeff': 16.94, 'note': 'G.I. Limpet washer (100 Nos)'},
                {'code': '1208', 'coeff': 16.94, 'note': 'Bitumen washer (100 Nos)'}
            ],
            'default_labour': [
                {'code': '0156', 'coeff': 4.00, 'note': 'Carpenter / Erector'},
                {'code': '0114', 'coeff': 6.00, 'note': 'Beldar'},
                {'code': '0131', 'coeff': 2.00, 'note': 'Painter (for overlap priming & painting)'}
            ],
            'default_sundries_base': 15.00,
            'toggles': {'water': 'YES', 'gst': 'YES', 'cpoh': 'YES', 'cess': 'YES'},
            'toggle_notes': {
                'water': 'Standard 1% Water charges.',
                'gst': 'CPWD factor 0.1405.',
                'cpoh': '15% Contractor Profit & Overheads.',
                'cess': '1% Labour Welfare Cess.'
            },
            'sample_library': [
                {'code': '12.1.1', 'desc': 'Corrugated G.S. sheet roofing 1.00 mm thick with accessories', 'unit': 'sqm', 'basis': 184.52, 'w': 152340.00, 'markups': 'Full W->X->Y->Z', 'rate': 1104.50, 'say': 1104.50}
            ]
        }
    }
