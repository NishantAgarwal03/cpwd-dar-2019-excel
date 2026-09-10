# -*- coding: utf-8 -*-
"""
Complete CPWD DAR 2019 Sub-Head 01 reference tables, transcribed from the
printed volume (PDF pages 76-84 = printed pages 67-75).

  Data Sheet No. 1   PDF pages 76-77, all 14 printed columns
  Table 1.1          PDF pages 78-79, mechanical transport, full rate ladder
  Table 1.2          PDF pages 81-83, manual labour under 0.50 km

How the rate ladder was recovered
---------------------------------
The rate columns on pages 78-79 do not come out of the PDF in printed order -
page 78 emits them 1, 2, 4, 3 km and page 79 emits them 3, 2, 1, 4, 5 km - so
column position cannot be trusted. Each value was therefore derived from Data
Sheet No. 1 first, then matched against the numeric tokens on those pages to
recover the exact printed figure. All 296 values were located; the largest gap
between derived and printed is Rs 0.06, on the largest-bore pipes where the book
carries more precision through its intermediate steps.

PRINTED_LADDER holds what the workbook publishes. rate_ladder() keeps the
derivation, and validate() fails if any printed value drifts more than 10 paise
from it - which would indicate a transcription error rather than rounding.

The derivation, verified against the book:

    R(L)   = (shift cost at L + 15% CPOH) / (trips at L x net payable qty) x unit scale
    >5-10  = (R(10) - R(5))  / 5      per km
    >10-20 = (R(20) - R(10)) / 10     per km
    >20    = (R(30) - R(20)) / 10     per additional km

Worked example, 1000/1100/1200 mm pipes (net 10.98 m per trip, per 100 m):
    R(1) = 7645.72, R(5) = 11763.9, R(10) = 16195.0, R(20) = 23558.9,
    R(30) = 29641.6  ->  >5-10 = 886.17, >10-20 = 736.40, >20 = 608.31
all four matching the printed page 79 figures exactly.
"""

# ---------------------------------------------------------------------------
# Data Sheet No. 1 - PDF pages 76-77, printed columns 1 to 14.
# (lead, avg speed, trips N, km/day, diesel litres, mobil litres, total cost,
#  cost per trip)
# The money columns are derived below from the basic rates, exactly as the book
# does: col 6 = diesel litres x 73.50, col 8 = mobil litres x 315.00,
# col 9 = 6 Beldars x 558.00, col 10 = truck hire 1500.00.
# ---------------------------------------------------------------------------
DATASHEET1 = [
    (1.0, 16.0, 7.11, 20.22, 4.04, 0.144, 5190.30, 730.00),
    (2.0, 17.0, 6.48, 31.92, 6.38, 0.228, 5388.75, 831.60),
    (3.0, 17.5, 5.96, 41.76, 8.35, 0.298, 5555.60, 932.15),
    (4.0, 18.0, 5.54, 50.32, 10.06, 0.359, 5700.50, 1028.97),
    (5.0, 18.5, 5.19, 57.90, 11.58, 0.414, 5829.54, 1123.23),
    (6.0, 19.0, 4.90, 64.80, 12.96, 0.463, 5946.41, 1213.55),
    (7.0, 19.5, 4.66, 71.24, 14.25, 0.509, 6055.72, 1299.51),
    (8.0, 20.0, 4.44, 77.04, 15.41, 0.550, 6153.89, 1386.01),
    (9.0, 20.5, 4.26, 82.68, 16.54, 0.591, 6249.86, 1467.10),
    (10.0, 21.0, 4.10, 88.00, 17.60, 0.629, 6339.74, 1546.28),
    (11.0, 21.5, 3.95, 92.90, 18.58, 0.664, 6422.79, 1626.02),
    (12.0, 22.0, 3.83, 97.92, 19.58, 0.699, 6507.32, 1699.04),
    (13.0, 22.5, 3.71, 102.46, 20.49, 0.732, 6584.60, 1774.82),
    (14.0, 23.0, 3.61, 107.08, 21.42, 0.765, 6663.35, 1845.80),
    (15.0, 23.5, 3.51, 111.30, 22.26, 0.795, 6734.54, 1918.67),
    (16.0, 24.0, 3.43, 115.76, 23.15, 0.827, 6810.04, 1985.43),
    (17.0, 24.5, 3.35, 119.90, 23.98, 0.856, 6880.17, 2053.78),
    (18.0, 25.0, 3.28, 124.08, 24.82, 0.886, 6951.36, 2119.32),
    (19.0, 25.5, 3.21, 127.98, 25.60, 0.914, 7017.51, 2186.14),
    (20.0, 26.0, 3.15, 132.00, 26.40, 0.943, 7085.45, 2249.35),
    (21.0, 26.5, 3.09, 135.78, 27.16, 0.970, 7149.81, 2313.85),
    (22.0, 27.0, 3.04, 139.76, 27.95, 0.998, 7216.70, 2373.91),
    (23.0, 27.5, 2.99, 143.54, 28.71, 1.025, 7281.07, 2435.14),
    (24.0, 28.0, 2.95, 147.60, 29.52, 1.054, 7349.73, 2491.43),
    (25.0, 28.5, 2.90, 151.00, 30.20, 1.079, 7407.59, 2554.34),
    (26.0, 29.0, 2.86, 154.72, 30.94, 1.105, 7470.17, 2611.95),
    (27.0, 29.5, 2.83, 158.82, 31.76, 1.134, 7539.57, 2664.16),
    (28.0, 30.0, 2.79, 162.24, 32.45, 1.159, 7598.17, 2723.36),
    (29.0, 30.5, 2.76, 166.08, 33.22, 1.186, 7663.26, 2776.54),
    (30.0, 31.0, 2.73, 169.80, 33.96, 1.213, 7726.16, 2830.10),
]

RATE_DIESEL = 73.50      # code 1235, per litre
RATE_MOBIL = 315.00      # code 5001, per litre
RATE_BELDAR = 558.00     # code 0114, per day
RATE_TRUCK = 1500.00     # code 0084, per day (9 tonne, excl. diesel & mobil)
GANG = 6                 # 6 Beldars, per the Data Sheet 1 column heading
CPOH = 0.15

# ---------------------------------------------------------------------------
# Table 1.1 - PDF pages 78-79. (item no, material, capacity/trip, net payable,
# unit). Item numbers are the book's own; note 1.1.16.4 is not printed.
# ---------------------------------------------------------------------------
TABLE_11 = [
    ('1.1.1', 'Lime, moorum, building rubbish', 8, 8, 'cum'),
    ('1.1.2', 'Earth', 8, 6.4, 'cum'),
    ('1.1.3', 'Manure or sludge', 8, 7.36, 'cum'),
    ('1.1.4', 'Excavated rock', 8, 4, 'cum'),
    ('1.1.5', 'Sand, stone aggregate below 40 mm nominal size', 8, 8, 'cum'),
    ('1.1.6', 'Stone aggregate 40 mm nominal size and above', 8, 7.36, 'cum'),
    ('1.1.7', 'Soling stone', 8, 6.8, 'cum'),
    ('1.1.8', 'Bricks', 3000, 3000, '1000 Nos'),
    ('1.1.9', 'Brick Tiles', 5000, 5000, '1000 Nos'),
    ('1.1.10', 'Cement, stone blocks, G.I., C.I., A.C. & C.C. pipes below 100 mm dia '
               'and other heavy materials', 9, 9, 'tonne'),
    ('1.1.11', 'Steel', 9, 9, 'tonne'),
    ('1.1.12', 'Timber', 7, 7, 'cum'),
    ('1.1.13', 'Tar, bitumen', 8, 8, 'tonne'),
    ('1.1.14', 'Solvent, diesel', 80, 80, 'qtl'),
    ('1.1.15', 'Steam coal', 7, 7, 'tonne'),
    ('1.1.16.1', 'S.W. pipe 100 mm dia', 600, 600, '100 m'),
    ('1.1.16.2', 'S.W. pipe 150 mm dia', 300, 300, '100 m'),
    ('1.1.16.3', 'S.W. pipe 200 mm dia', 180, 180, '100 m'),
    ('1.1.16.5', 'S.W. pipe 250 mm dia', 105, 105, '100 m'),
    ('1.1.16.6', 'S.W. pipe 300 mm dia', 84, 84, '100 m'),
    ('1.1.16.7', 'S.W. pipe 350 mm dia', 60, 60, '100 m'),
    ('1.1.16.8', 'S.W. pipe 400 mm dia', 42, 42, '100 m'),
    ('1.1.16.9', 'S.W. pipe 450 mm dia', 33, 33, '100 m'),
    ('1.1.16.10', 'S.W. pipe 500 mm dia', 30, 30, '100 m'),
    ('1.1.16.11', 'S.W. pipe 600 mm dia', 24, 24, '100 m'),
    ('1.1.17.1', 'R.C.C. / A.C. / steel cylinder / S.C.I. / C.I. / unreinforced cement '
                 'pipes 100 mm dia', 366, 366, '100 m'),
    ('1.1.17.2', 'R.C.C. etc. pipes 125 mm dia', 274, 274, '100 m'),
    ('1.1.17.3', 'R.C.C. etc. pipes 150 mm dia', 219.6, 219.6, '100 m'),
    ('1.1.17.4', 'R.C.C. etc. pipes 200 mm dia', 135, 135, '100 m'),
    ('1.1.17.5', 'R.C.C. etc. pipes 250 mm dia', 95, 95, '100 m'),
    ('1.1.17.6', 'R.C.C. etc. pipes 300 mm dia', 76.86, 76.86, '100 m'),
    ('1.1.17.7', 'R.C.C. etc. pipes 350 mm dia', 54.9, 54.9, '100 m'),
    ('1.1.17.8', 'R.C.C. etc. pipes 400 mm dia', 40.26, 40.26, '100 m'),
    ('1.1.17.9', 'R.C.C. etc. pipes 450 mm & 500 mm dia', 32.94, 32.94, '100 m'),
    ('1.1.17.10', 'R.C.C. etc. pipes 600, 700, 750 & 800 mm dia', 21.96, 21.96, '100 m'),
    ('1.1.17.11', 'R.C.C. etc. pipes 900 mm dia', 14.64, 14.64, '100 m'),
    ('1.1.17.12', 'R.C.C. / C.I. / steel pipes 1000, 1100 & 1200 mm dia', 10.98, 10.98, '100 m'),
]

# The material the carriage sheet opens on. Taken from TABLE_11 by item number
# so the Panel 1 default can never drift out of step with the lookup table -
# a mismatch silently blanks every Panel 3 lookup.
DEFAULT_MATERIAL_ITEM = '1.1.17.12'


def default_material():
    return [t[1] for t in TABLE_11 if t[0] == DEFAULT_MATERIAL_ITEM][0]


LOOSENESS = {
    '1.1.2': '20% deduction for looseness (8.00 -> 6.40 cum)',
    '1.1.3': '8% deduction for looseness (8.00 -> 7.36 cum)',
    '1.1.4': '50% deduction for voids / looseness (8.00 -> 4.00 cum)',
    '1.1.6': '8% deduction for voids in coarse aggregate (8.00 -> 7.36 cum)',
    '1.1.7': '15% deduction for stack voids (8.00 -> 6.80 cum)',
}

# ---------------------------------------------------------------------------
# Table 1.2 - PDF pages 81-83. Manual labour, lead under 0.50 km.
# Category A: 7.67 Beldars for the first 50 m, +1.67 coolies per extra 50 m.
# Category B: 9.20 Beldars for the first 50 m, +1.35 Beldars per extra 50 m.
# (item no, material, capacity, net payable, unit, category)
# ---------------------------------------------------------------------------
GANG_A_BASE, GANG_A_ADD = 7.67, 1.67
GANG_B_BASE, GANG_B_ADD = 9.20, 1.35

TABLE_12 = [
    ('1.2.1', 'Lime, moorum, building rubbish', 35, 35, 'cum', 'A'),
    ('1.2.2', 'Earth', 35, 28, 'cum', 'A'),
    ('1.2.3', 'Manure or sludge', 35, 32.2, 'cum', 'A'),
    ('1.2.4', 'Excavated rock', 35, 17.5, 'cum', 'A'),
    ('1.2.5', 'Sand, stone aggregate below 40 mm nominal size', 28, 28, 'cum', 'A'),
    ('1.2.6', 'Stone aggregate 40 mm nominal size and above', 28, 25.9, 'cum', 'A'),
    ('1.2.7', 'Soling stone', 28, 23.8, 'cum', 'A'),
    ('1.2.8', 'Bricks', 15000, 15000, '1000 Nos', 'A'),
    ('1.2.9', 'Brick Tiles', 24000, 24000, '1000 Nos', 'A'),
    ('1.2.10', 'Steam Coal', 30, 30, 'tonne', 'A'),
    ('1.2.11', 'Stone blocks, G.I., C.I., stainless steel pipes below 100 mm dia and '
               'other heavy material', 46, 46, 'tonne', 'B'),
    ('1.2.12', 'Cement', 57.99, 57.99, 'tonne', 'B'),
    ('1.2.13', 'Steel', 27, 27, 'tonne', 'B'),
    ('1.2.14', 'Timber', 42, 42, 'cum', 'B'),
    ('1.2.15', 'Tar, bitumen etc.', 46, 46, 'tonne', 'B'),
    ('1.2.16.1', 'S.W. pipe 100 mm dia', 2298, 2298, '100 m', 'B'),
    ('1.2.16.2', 'S.W. pipe 150 mm dia', 1398, 1398, '100 m', 'B'),
    ('1.2.16.3', 'S.W. pipe 200 mm dia', 999, 999, '100 m', 'B'),
    ('1.2.16.5', 'S.W. pipe 250 mm dia', 600, 600, '100 m', 'B'),
    ('1.2.16.6', 'S.W. pipe 300 mm dia', 420, 420, '100 m', 'B'),
    ('1.2.16.7', 'S.W. pipe 350 mm dia', 300, 300, '100 m', 'B'),
    ('1.2.16.8', 'S.W. pipe 400 mm dia', 240, 240, '100 m', 'B'),
    ('1.2.16.9', 'S.W. pipe 450 mm dia', 198, 198, '100 m', 'B'),
    ('1.2.16.10', 'S.W. pipe 500 mm dia', 162, 162, '100 m', 'B'),
    ('1.2.16.11', 'S.W. pipe 600 mm dia', 132, 132, '100 m', 'B'),
    ('1.2.17.1', 'R.C.C. / steel cylinder / R.C. / C.I. / unreinforced cement pipes '
                 '100 mm dia', 1702, 1702, '100 m', 'B'),
    ('1.2.17.2', 'R.C.C. etc. pipes 125 mm dia', 1391, 1391, '100 m', 'B'),
    ('1.2.17.3', 'R.C.C. etc. pipes 150 mm dia', 1208, 1208, '100 m', 'B'),
    ('1.2.17.4', 'R.C.C. etc. pipes 200 mm dia', 805, 805, '100 m', 'B'),
    ('1.2.17.5', 'R.C.C. etc. pipes 250 mm dia', 458, 458, '100 m', 'B'),
    ('1.2.17.6', 'R.C.C. etc. pipes 300 mm dia', 366, 366, '100 m', 'B'),
    ('1.2.17.7', 'R.C.C. etc. pipes 350 mm dia', 256, 256, '100 m', 'B'),
    ('1.2.17.8', 'R.C.C. etc. pipes 400 mm dia', 220, 220, '100 m', 'B'),
    ('1.2.17.9', 'R.C.C. etc. pipes 450 mm & 500 mm dia', 165, 165, '100 m', 'B'),
    ('1.2.17.10', 'R.C.C. etc. pipes 600, 700, 750 & 800 mm dia', 150, 150, '100 m', 'B'),
]

# Printed values recovered from the PDF, used to validate the generated tables.
# Table 1.1: item -> (1 km, >20 km per addl km)
PRINTED_11 = {
    '1.1.1': (104.94, 8.35), '1.1.2': (131.17, 10.44), '1.1.3': (114.06, 9.08),
    '1.1.4': (209.88, 16.70), '1.1.5': (104.94, 8.35), '1.1.6': (114.06, 9.08),
    '1.1.7': (123.46, 9.82), '1.1.8': (279.83, 22.26), '1.1.9': (167.90, 13.36),
    '1.1.10': (93.28, 7.42), '1.1.11': (93.28, 7.42), '1.1.12': (119.93, 9.54),
    '1.1.13': (104.94, 8.35), '1.1.14': (10.49, 0.83), '1.1.15': (119.93, 9.54),
    '1.1.16.1': (139.92, 11.13), '1.1.16.2': (279.83, 22.26),
    '1.1.16.3': (466.39, 37.11), '1.1.16.5': (799.52, 63.61), '1.1.16.6': (999.40, 79.51),
    '1.1.16.7': (1399.17, 111.32), '1.1.16.8': (1998.81, 159.03), '1.1.16.9': (2543.94, 202.40),
    '1.1.16.10': (2798.33, 222.64), '1.1.16.11': (3497.92, 278.30),
    '1.1.17.1': (229.37, 18.25), '1.1.17.2': (306.39, 24.38), '1.1.17.3': (382.29, 30.42),
    '1.1.17.4': (621.85, 49.48), '1.1.17.5': (883.68, 70.31), '1.1.17.6': (1092.25, 86.90),
    '1.1.17.7': (1529.14, 121.66), '1.1.17.8': (2085.20, 165.90), '1.1.17.9': (2548.57, 202.77),
    '1.1.17.10': (3822.86, 304.15), '1.1.17.11': (5734.29, 456.23), '1.1.17.12': (7645.72, 608.31),
}

# Table 1.2: item -> (cost for 1st 50 m, cost per additional 50 m)
PRINTED_12 = {
    '1.2.1': (140.62, 30.62), '1.2.2': (175.78, 38.27), '1.2.3': (152.85, 33.28),
    '1.2.4': (281.25, 61.24), '1.2.5': (175.78, 38.27), '1.2.6': (190.03, 41.38),
    '1.2.7': (206.80, 45.03), '1.2.8': (328.12, 71.44), '1.2.9': (205.08, 44.65),
    '1.2.10': (164.06, 35.72),
    '1.2.11': (128.34, 18.83), '1.2.12': (101.80, 14.94), '1.2.13': (218.65, 32.09),
    '1.2.14': (140.56, 20.63), '1.2.15': (128.34, 18.83),
    '1.2.16.1': (256.90, 37.70), '1.2.16.2': (422.29, 61.97), '1.2.16.3': (590.95, 86.72),
    '1.2.16.5': (983.94, 144.38), '1.2.16.6': (1405.63, 206.26), '1.2.16.7': (1967.88, 288.77),
    '1.2.16.8': (2459.85, 360.96), '1.2.16.9': (2981.64, 437.52), '1.2.16.10': (3644.22, 534.75),
    '1.2.16.11': (4472.45, 656.28),
    '1.2.17.1': (346.86, 50.90), '1.2.17.2': (424.42, 62.28), '1.2.17.3': (488.71, 71.71),
    '1.2.17.4': (733.37, 107.61), '1.2.17.5': (1289.00, 189.15), '1.2.17.6': (1613.02, 236.69),
    '1.2.17.7': (2306.11, 338.40), '1.2.17.8': (2683.47, 393.77), '1.2.17.9': (3577.96, 525.03),
    '1.2.17.10': (3935.76, 577.53),
}

# ---------------------------------------------------------------------------
# Table 1.3 & 1.4 - Railway Wagon Handling - PDF pages 83-84.
# (item no, description, wagon type, payload tonnes, handling mechanics,
#  gang size men, free-time hours, beldar days, crane LS base, sundries LS base,
#  CPWD published rate per tonne, schedule unit)
# ---------------------------------------------------------------------------
TABLE_13_14 = [
    (
        '1.3',
        'Loading in or unloading cement from the railway wagons at siding and carrying into godowns '
        'adjacent to siding, including stacking in rows upto any height, sweeping wagons, screening '
        'and bagging',
        '4-wheeler Covered Wagon (CRT)',
        23.0,
        'Single-man shoulder carry + godown vertical stack',
        6,
        5.0,
        3.75,
        0.00,
        2.62,
        104.90,
        'tonne',
    ),
    (
        '1.4.1',
        'Loading in or unloading from railway wagons: Steel',
        '8-wheeler Bogie Flat Wagon (BRH)',
        44.0,
        'Synchronized 3-4 man bar carry',
        16,
        16.0 / 3.0,
        10.66,
        0.00,
        0.00,
        155.45,
        'tonne',
    ),
    (
        '1.4.2',
        'Loading in or unloading from railway wagons: G.I., C.I., R.C.C. or C.C. pipes upto 500 mm dia '
        'and similar heavy materials',
        '4-wheeler Open Flat Wagon (KC)',
        14.0,
        'Rolling on timber skids',
        2,
        8.0,
        2.00,
        0.00,
        3.10,
        92.20,
        'tonne',
    ),
    (
        '1.4.3',
        'Loading in or unloading from railway wagons: Heavy materials where each piece/bundle weighs > 1 tonne '
        'and R.C.C., C.I. & concrete pipes above 500 mm dia (crane-assisted)',
        '4-wheeler Flat Wagon + Yard Crane',
        14.0,
        'Crane slinging + tagline guide',
        4,
        6.5,
        3.25,
        91.15,
        7.40,
        165.15,
        'tonne',
    ),
]

PRINTED_13_14 = {
    '1.3': 104.90,
    '1.4.1': 155.45,
    '1.4.2': 92.20,
    '1.4.3': 165.15,
}


# ---------------------------------------------------------------------------
# The rate ladder exactly as printed on PDF pages 78-79.
#
# Recovered by computing each value from Data Sheet No. 1 and then matching it
# against the numeric tokens on those pages, because the PDF emits the rate
# columns out of printed order (page 78 gives 1, 2, 4, 3 km; page 79 gives
# 3, 2, 1, 4, 5 km). Every one of the 296 values was located on the page, and
# the largest gap between the computed and the printed figure is Rs 0.06, on
# the largest-diameter pipe rows where the book carries more precision through
# its intermediate steps.
#
# These printed values are what the workbook publishes. rate_ladder() remains
# as the derivation and as the cross-check - validate() fails if any printed
# value drifts more than 10 paise from it, which would mean a transcription
# error rather than rounding.
#
# Order: 1 km, 2 km, 3 km, 4 km, 5 km, >5-10 per km, >10-20 per km, >20 per km
# ---------------------------------------------------------------------------
PRINTED_LADDER = {
    '1.1.1': (104.94, 119.54, 134.00, 147.91, 161.46, 12.16, 10.11, 8.35),
    '1.1.2': (131.17, 149.43, 167.50, 184.89, 201.83, 15.20, 12.63, 10.44),
    '1.1.3': (114.06, 129.94, 145.65, 160.78, 175.50, 13.22, 10.99, 9.08),
    '1.1.4': (209.88, 239.09, 267.99, 295.83, 322.93, 24.33, 20.21, 16.70),
    '1.1.5': (104.94, 119.54, 134.00, 147.91, 161.46, 12.16, 10.11, 8.35),
    '1.1.6': (114.06, 129.94, 145.65, 160.78, 175.50, 13.22, 10.99, 9.08),
    '1.1.7': (123.46, 140.64, 157.64, 174.02, 189.96, 14.31, 11.89, 9.82),
    '1.1.8': (279.83, 318.78, 357.32, 394.44, 430.57, 32.43, 26.95, 22.26),
    '1.1.9': (167.90, 191.27, 214.39, 236.66, 258.34, 19.46, 16.17, 13.36),
    '1.1.10': (93.28, 106.26, 119.11, 131.48, 143.52, 10.81, 8.98, 7.42),
    '1.1.11': (93.28, 106.26, 119.11, 131.48, 143.52, 10.81, 8.98, 7.42),
    '1.1.12': (119.93, 136.62, 153.14, 169.05, 184.53, 13.90, 11.55, 9.54),
    '1.1.13': (104.94, 119.54, 134.00, 147.91, 161.46, 12.16, 10.11, 8.35),
    '1.1.14': (10.49, 11.95, 13.40, 14.79, 16.15, 1.22, 1.01, 0.83),
    '1.1.15': (119.93, 136.62, 153.14, 169.05, 184.53, 13.90, 11.55, 9.54),
    '1.1.16.1': (139.92, 159.39, 178.66, 197.22, 215.29, 16.22, 13.48, 11.13),
    '1.1.16.2': (279.83, 318.78, 357.32, 394.44, 430.57, 32.43, 26.95, 22.26),
    '1.1.16.3': (466.39, 531.30, 595.54, 657.40, 717.62, 54.06, 44.92, 37.11),
    '1.1.16.5': (799.52, 910.80, 1020.93, 1126.97, 1230.20, 92.67, 77.01, 63.61),
    '1.1.16.6': (999.40, 1138.50, 1276.16, 1408.71, 1537.76, 115.84, 96.26, 79.51),
    '1.1.16.7': (1399.17, 1593.90, 1786.62, 1972.19, 2152.86, 162.17, 134.76, 111.32),
    '1.1.16.8': (1998.81, 2277.00, 2552.32, 2817.42, 3075.51, 231.67, 192.52, 159.03),
    '1.1.16.9': (2543.94, 2898.00, 3248.40, 3585.80, 3914.29, 294.85, 245.02, 202.40),
    '1.1.16.10': (2798.33, 3187.80, 3573.24, 3944.39, 4305.72, 324.34, 269.52, 222.64),
    '1.1.16.11': (3497.92, 3984.75, 4466.55, 4930.48, 5382.14, 405.42, 336.90, 278.30),
    '1.1.17.1': (229.37, 261.30, 292.89, 323.31, 352.93, 26.59, 22.09, 18.25),
    '1.1.17.2': (306.39, 349.03, 391.23, 431.87, 471.43, 35.51, 29.51, 24.38),
    '1.1.17.3': (382.29, 435.49, 488.15, 538.85, 588.21, 44.31, 36.82, 30.42),
    '1.1.17.4': (621.85, 708.40, 794.05, 876.53, 956.83, 72.08, 59.89, 49.48),
    '1.1.17.5': (883.68, 1006.67, 1128.39, 1245.60, 1359.70, 102.42, 85.11, 70.31),
    '1.1.17.6': (1092.25, 1244.26, 1394.71, 1539.57, 1680.61, 126.60, 105.20, 86.90),
    '1.1.17.7': (1529.14, 1741.97, 1952.59, 2155.40, 2352.85, 177.23, 147.28, 121.66),
    '1.1.17.8': (2085.20, 2375.41, 2662.62, 2939.18, 3208.43, 241.68, 200.84, 165.90),
    '1.1.17.9': (2548.57, 2903.28, 3254.32, 3592.34, 3921.42, 295.39, 245.47, 202.77),
    '1.1.17.10': (3822.86, 4354.92, 4881.48, 5388.50, 5882.12, 443.09, 368.20, 304.15),
    '1.1.17.11': (5734.29, 6532.38, 7322.22, 8082.76, 8823.19, 664.63, 552.30, 456.23),
    '1.1.17.12': (7645.72, 8709.84, 9762.96, 10777.01, 11764.25, 886.17, 736.40, 608.31),
}

# ---------------------------------------------------------------------------
# Derivations
# ---------------------------------------------------------------------------
def unit_scale(unit):
    return 1000 if unit == '1000 Nos' else (100 if unit == '100 m' else 1)


def _ds1_by_lead():
    return {int(r[0]): r for r in DATASHEET1}


def shift_cost_with_cpoh(lead):
    """Total cost of one 8-hour shift at this lead, including 15% CPOH."""
    row = _ds1_by_lead()[int(lead)]
    total = row[6]
    return round(total + round(total * CPOH, 2), 2)


def trips(lead):
    return _ds1_by_lead()[int(lead)][2]


def rate_at(lead, net_qty, unit):
    """Cost per schedule unit including 15% CP&OH, at a tabulated lead."""
    return shift_cost_with_cpoh(lead) / (trips(lead) * net_qty) * unit_scale(unit)


def rate_ladder(net_qty, unit):
    """The eight printed rate columns for one material.

    Returns a dict with keys 1..5 (per-km rates) plus 'b5_10', 'b10_20', 'b20'
    (per-km increments for the three bands beyond 5 km).
    """
    r = {km: round(rate_at(km, net_qty, unit), 2) for km in (1, 2, 3, 4, 5)}
    r5 = rate_at(5, net_qty, unit)
    r10 = rate_at(10, net_qty, unit)
    r20 = rate_at(20, net_qty, unit)
    r30 = rate_at(30, net_qty, unit)
    r['b5_10'] = round((r10 - r5) / 5, 2)
    r['b10_20'] = round((r20 - r10) / 10, 2)
    r['b20'] = round((r30 - r20) / 10, 2)
    return r


def printed_ladder(item, net_qty, unit):
    """The book's printed ladder for an item, falling back to the derivation."""
    p = PRINTED_LADDER.get(item)
    if not p:
        return rate_ladder(net_qty, unit)
    keys = [1, 2, 3, 4, 5, 'b5_10', 'b10_20', 'b20']
    return dict(zip(keys, p))


def table_11_rows():
    """Table 1.1 with the full printed rate ladder attached to every row."""
    out = []
    for item, mat, cap, net, unit in TABLE_11:
        lad = printed_ladder(item, net, unit)
        out.append({
            'item': item, 'material': mat, 'capacity': cap, 'net': net, 'unit': unit,
            'looseness': LOOSENESS.get(item, 'Nil - full capacity payable'),
            'ladder': lad,
        })
    return out


def manual_cost_with_cpoh(category):
    gang = GANG_A_BASE if category == 'A' else GANG_B_BASE
    base = round(gang * RATE_BELDAR, 2)
    return round(base + round(base * CPOH, 2), 2)


def manual_addl_with_cpoh(category):
    gang = GANG_A_ADD if category == 'A' else GANG_B_ADD
    base = round(gang * RATE_BELDAR, 2)
    return round(base + round(base * CPOH, 2), 2)


def table_12_rows():
    out = []
    for item, mat, cap, net, unit, cat in TABLE_12:
        sc = unit_scale(unit)
        day = manual_cost_with_cpoh(cat)
        add = manual_addl_with_cpoh(cat)
        out.append({
            'item': item, 'material': mat, 'capacity': cap, 'net': net, 'unit': unit,
            'category': cat, 'day_cost': day,
            'first50': round(day / net * sc, 2),
            'addl50': round(add / net * sc, 2),
        })
    return out


def table_13_14_rows():
    """Yield dicts for Table 1.3/1.4 railway wagon handling reference rows."""
    out = []
    for item, desc, wtype, pay, mech, g, h, bd, cr, sd, pub, u in TABLE_13_14:
        c_beldar = round(bd * RATE_BELDAR, 2)
        c_equip = round(cr * 2.0, 2) + round(sd * 2.0, 2)
        y = c_beldar + c_equip
        cpoh = round(y * CPOH, 2)
        total = y + cpoh
        rate = round(total / pay, 4)
        say = round(round(rate * 20) / 20, 2)
        out.append({
            'item': item,
            'desc': desc,
            'wagon_type': wtype,
            'payload': pay,
            'mechanics': mech,
            'gang': g,
            'hours': h,
            'beldar': bd,
            'prod': round(pay / bd, 3),
            'beldar_cost': c_beldar,
            'crane_base': cr,
            'sundries_base': sd,
            'equip_cost': c_equip,
            'y_cost': y,
            'cpoh': cpoh,
            'total_wagon': total,
            'rate_tonne': rate,
            'say_rate': say,
            'published': pub,
            'unit': u,
        })
    return out


def validate(tol=0.06, ladder_tol=0.10):
    """Check the published tables against the book and against the derivation."""
    errors = []
    for item, mat, cap, net, unit in TABLE_11:
        derived = rate_ladder(net, unit)
        printed = printed_ladder(item, net, unit)
        for k in (1, 2, 3, 4, 5, 'b5_10', 'b10_20', 'b20'):
            if abs(printed[k] - derived[k]) > ladder_tol:
                errors.append('T1.1 %s ladder %s: printed %.2f vs derived %.2f'
                              % (item, k, printed[k], derived[k]))
    for row in table_11_rows():
        if row['item'] not in PRINTED_11:
            continue
        p1, p20 = PRINTED_11[row['item']]
        if abs(row['ladder'][1] - p1) > tol:
            errors.append('T1.1 %s 1 km: %.2f vs printed %.2f' % (row['item'], row['ladder'][1], p1))
        if abs(row['ladder']['b20'] - p20) > tol:
            errors.append('T1.1 %s >20 km: %.2f vs printed %.2f'
                          % (row['item'], row['ladder']['b20'], p20))
    for row in table_12_rows():
        if row['item'] not in PRINTED_12:
            continue
        p1, pa = PRINTED_12[row['item']]
        if abs(row['first50'] - p1) > tol:
            errors.append('T1.2 %s 1st 50 m: %.2f vs printed %.2f' % (row['item'], row['first50'], p1))
        if abs(row['addl50'] - pa) > tol:
            errors.append('T1.2 %s addl 50 m: %.2f vs printed %.2f' % (row['item'], row['addl50'], pa))
    for row in table_13_14_rows():
        if row['item'] not in PRINTED_13_14:
            continue
        pub = PRINTED_13_14[row['item']]
        if abs(row['say_rate'] - pub) > tol:
            errors.append('T1.3/1.4 %s SAY: %.2f vs printed %.2f' % (row['item'], row['say_rate'], pub))
    return errors


if __name__ == '__main__':
    errs = validate()
    n11, n12, n14 = len(PRINTED_11) * 2, len(PRINTED_12) * 2, len(PRINTED_13_14)
    print('Table 1.1 : %d materials, %d printed values checked' % (len(TABLE_11), n11))
    print('Table 1.2 : %d materials, %d printed values checked' % (len(TABLE_12), n12))
    print('Table 1.3/1.4 : %d wagon items, %d printed values checked' % (len(TABLE_13_14), n14))
    if errs:
        print('MISMATCHES (%d):' % len(errs))
        for e in errs:
            print('   ', e)
    else:
        print('All %d printed values reproduced within tolerance.' % (n11 + n12 + n14))
