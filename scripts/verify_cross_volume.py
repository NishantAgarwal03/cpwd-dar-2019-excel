# -*- coding: utf-8 -*-
"""
CHECK 9 - Cross-volume resolution and the CPWD (W-A) markup exclusion.

Two independent verifications:

  9A NUMERIC   Recompute each resolved item from Rates_Master values using the
               same (W-A) chain the sheet implements, and compare the result to
               the rate the DAR 2019 actually prints. If the arithmetic model is
               wrong, the resolved rate will not land on the book's figure.

  9B STRUCTURAL Read the generated workbook and assert the wiring: every markup
               base subtracts the A-total cell, A-tagged lines carry "A" in the
               tag column, the A-total SUMIF covers exactly the material rows,
               and the consuming sheets (08, 09, 10) import the resolved items
               by defined name rather than by literal.
"""

import sys, os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir)))
import json
import openpyxl
from scripts.paths import RATES_MASTER_JSON, WB_VOL1_FILE

from scripts.cross_volume import RESOLVED_ITEMS
from scripts.trade_layout import (
    R_MAT_FIRST, R_MAT_LAST, R_MAT_SUB, R_A_TOTAL, R_LAB_SUB, R_SUN,
    R_W, R_X1, R_X, R_Y1, R_Y, R_Z1, R_Z, R_Z2,
)

WATER, GST, CPOH, CESS, SUND = 0.01, 0.1405, 0.15, 0.01, 2.00

CONSUMERS = {
    '08_Cladding_Work': ['XV_18_78'],
    '09_Wood_and_PVC_Work': ['XV_13_50_1', 'XV_13_57_1'],
    '10_Steel_Work': ['XV_13_50_3'],
}


def _rate_index():
    with open(RATES_MASTER_JSON, 'r', encoding='utf-8') as f:
        return {str(r['code']): float(r['rate']) for r in json.load(f)}


def recompute_resolved():
    """9A - reproduce each resolved item's Say rate from first principles."""
    rates = _rate_index()
    results = []
    prior = {}

    for item in RESOLVED_ITEMS:
        total = 0.0
        a_total = 0.0

        for code, coeff, _note, a_tag in item['materials']:
            if a_tag:
                rate = prior.get(code, 0.0)
            else:
                rate = rates.get(str(code), 0.0)
            amt = round(coeff * rate, 2)
            total += amt
            if a_tag:
                a_total += amt

        for code, coeff, _note in item['labour']:
            total += round(coeff * rates.get(str(code), 0.0), 2)

        for base, _note in item['sundries']:
            total += round(base * SUND, 2)

        w = round(total, 2)
        tg = item['toggles']
        water = round((w - a_total) * WATER, 2) if tg['water'] == 'YES' else 0.0
        x = w + water
        gst = round((x - a_total) * GST, 2) if tg['gst'] == 'YES' else 0.0
        y = x + gst
        cpoh = round((y - a_total) * CPOH, 2) if tg['cpoh'] == 'YES' else 0.0
        z = y + cpoh
        cess = round((z - a_total) * CESS, 2) if tg['cess'] == 'YES' else 0.0
        cost = round(z + cess, 2)
        # CPWD quotes Say rates to the nearest 5 paise - MROUND(x, 0.05).
        say = round(round((cost / item['basis_qty']) / 0.05) * 0.05, 2)

        prior[item['item_no']] = say
        results.append({
            'item_no': item['item_no'], 'w': w, 'a': a_total, 'say': say,
            'book_w': item['book_w'], 'book_say': item['book_say'],
            'unit': item['basis_unit'],
        })
    return results


def check_cross_volume(wb, verbose=True):
    """Runs 9A and 9B. Returns (ok, messages)."""
    msgs = []
    ok = True

    # --- 9A numeric ------------------------------------------------------
    for r in recompute_resolved():
        dw = abs(r['w'] - r['book_w'])
        ds = abs(r['say'] - r['book_say'])
        # With MROUND(x, 0.05) applied the Say rate must land exactly on the
        # book's printed figure; W still tolerates a paisa of line rounding.
        if dw <= 0.05 and ds < 0.001:
            msgs.append(f"  [PASS] Item {r['item_no']:<9} W = Rs {r['w']:>9,.2f} (book {r['book_w']:>9,.2f}), "
                        f"Say = Rs {r['say']:>8,.2f} / {r['unit']} (book {r['book_say']:>8,.2f})"
                        + (f", of which A = Rs {r['a']:,.2f}" if r['a'] else ""))
        else:
            ok = False
            msgs.append(f"  [FAIL] Item {r['item_no']} does not reproduce the book: "
                        f"W {r['w']:,.2f} vs {r['book_w']:,.2f}, Say {r['say']:,.2f} vs {r['book_say']:,.2f}")

    # --- 9B structural ---------------------------------------------------
    if 'Resolved_Cross_Volume_Items' not in wb.sheetnames:
        return False, msgs + ['  [FAIL] Resolved_Cross_Volume_Items sheet is missing.']

    names = {n for n in wb.defined_names}
    for item in RESOLVED_ITEMS:
        if item['key'] not in names:
            ok = False
            msgs.append(f"  [FAIL] Defined name {item['key']} not published.")
    if all(i['key'] in names for i in RESOLVED_ITEMS):
        msgs.append(f"  [PASS] All {len(RESOLVED_ITEMS)} resolved items published as defined names "
                    f"({', '.join(i['key'] for i in RESOLVED_ITEMS)}).")

    builders = [s for s in wb.sheetnames
                if s[:2].isdigit() and s != '01_Carriage_of_Materials']
    wired = 0
    for name in builders:
        ws = wb[name]

        expected_sumif = (f'=ROUND(SUMIF($I${R_MAT_FIRST}:$I${R_MAT_LAST}, "A", '
                          f'$G${R_MAT_FIRST}:$G${R_MAT_LAST}), 2)')
        if str(ws.cell(row=R_A_TOTAL, column=7).value) != expected_sumif:
            ok = False
            msgs.append(f"  [FAIL] {name}: A-total cell G{R_A_TOTAL} does not sum the material tag column.")
            continue

        if str(ws.cell(row=R_W, column=7).value) != f'=G{R_MAT_SUB} + G{R_LAB_SUB} + G{R_SUN}':
            ok = False
            msgs.append(f"  [FAIL] {name}: W does not total materials + labour + sundries.")
            continue

        bases = {
            R_X1: f'=G{R_W} - G{R_A_TOTAL}',
            R_Y1: f'=G{R_X} - G{R_A_TOTAL}',
            R_Z1: f'=G{R_Y} - G{R_A_TOTAL}',
            R_Z2: f'=G{R_Z} - G{R_A_TOTAL}',
        }
        bad = [r for r, f in bases.items() if str(ws.cell(row=r, column=5).value) != f]
        if bad:
            ok = False
            msgs.append(f"  [FAIL] {name}: markup base rows {bad} do not subtract the A-total.")
            continue
        wired += 1

    if wired == len(builders):
        msgs.append(f"  [PASS] All {wired} standard builders apply the CPWD (W-A) exclusion at every "
                    f"markup step (Water, GST, CPOH, Cess).")

    # Consumers import by defined name, and tag the line "A".
    for sheet, keys in CONSUMERS.items():
        if sheet not in wb.sheetnames:
            ok = False
            msgs.append(f"  [FAIL] {sheet} missing.")
            continue
        ws = wb[sheet]
        found = {}
        for r in range(R_MAT_FIRST, R_MAT_LAST + 1):
            f = str(ws.cell(row=r, column=6).value or '')
            for k in keys:
                if k in f:
                    found[k] = r
        missing = [k for k in keys if k not in found]
        if missing:
            ok = False
            msgs.append(f"  [FAIL] {sheet} does not import {missing} from Resolved_Cross_Volume_Items.")
            continue
        untagged = [k for k, r in found.items() if str(ws.cell(row=r, column=9).value or '').strip() != 'A']
        if untagged:
            ok = False
            msgs.append(f"  [FAIL] {sheet}: imported rates {untagged} are not tagged 'A' and would be "
                        f"marked up a second time.")
            continue
        msgs.append(f"  [PASS] {sheet} imports {', '.join(keys)} by defined name, each tagged 'A'.")

    # The old hard-coded literal must be gone from Steel Work.
    steel = wb['10_Steel_Work']
    literals = [r for r in range(R_MAT_FIRST, R_MAT_LAST + 1)
                if isinstance(steel.cell(row=r, column=6).value, (int, float))
                and abs(float(steel.cell(row=r, column=6).value) - 50.70) < 0.001]
    if literals:
        ok = False
        msgs.append(f"  [FAIL] 10_Steel_Work still holds the hard-coded 50.70 literal at row(s) {literals}.")
    else:
        msgs.append("  [PASS] 10_Steel_Work no longer holds the interim 50.70 literal; the priming coat "
                    "is a live reference to the resolved item.")

    if verbose:
        for m in msgs:
            print(m)
    return ok, msgs


if __name__ == '__main__':
    wb = openpyxl.load_workbook(WB_VOL1_FILE)
    good, _ = check_cross_volume(wb)
    print('\nCROSS-VOLUME CHECK:', 'PASS' if good else 'FAIL')
