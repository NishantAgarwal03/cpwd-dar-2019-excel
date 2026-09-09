# -*- coding: utf-8 -*-
"""
run_audit.py
CPWD DAR 2019 — Rate Audit
Compares every published item analysis in both converted workbooks against
the rates in rates_master_clean.json and the shadow markup chain.

Usage:
    python scripts/run_audit.py                   # full audit, summary table
    python scripts/run_audit.py --section 02      # only sheets starting with "02"
    python scripts/run_audit.py --item 2.2.1      # one specific item
    python scripts/run_audit.py --verbose         # show resource detail on failures
    python scripts/run_audit.py --export audit_report.txt
"""
import sys, os, json, argparse
sys.path.insert(0, os.path.abspath('.'))

from scripts.audit_parser import parse_all
from scripts.shadow_calc  import shadow_calculate

VOL1_CONVERTED = 'CivilDAR_2019_Vol_1_Converted.xlsx'
VOL2_CONVERTED = 'CivilDAR_2019_Vol_2_Converted.xlsx'
RATES_MASTER   = 'rates_master_clean.json'

# Width constants for report columns
W_SHEET = 32
W_ITEM  = 10
W_DESC  = 50
W_RATE  = 10

# ------------------------------------------------------------------
def load_rates_master(path):
    with open(path, encoding='utf-8') as f:
        data = json.load(f)
    return {item['code']: item for item in data}


def _trunc(s, n):
    s = str(s or '')
    return s[:n-1] + '…' if len(s) > n else s.ljust(n)


def _fmt_rate(v):
    if v is None:
        return '   —    '
    return f'{v:>10.2f}'


def print_section_header(title, out):
    line = '=' * 120
    out.append(line)
    out.append(f'  {title}')
    out.append(line)


def print_table_header(out):
    out.append(
        f"{'SHEET':<{W_SHEET}} {'ITEM':<{W_ITEM}} {'DESCRIPTION':<{W_DESC}} "
        f"{'PDF SAY':>{W_RATE}} {'CALC SAY':>{W_RATE}}  STATUS"
    )
    out.append('-' * 120)


def run_audit(section_filter=None, item_filter=None, verbose=False):
    # Load data
    rates = load_rates_master(RATES_MASTER)
    items = parse_all(VOL1_CONVERTED, VOL2_CONVERTED)

    # Filter
    if section_filter:
        items = [i for i in items if i['sheet'].lower().startswith(section_filter.lower())]
    if item_filter:
        items = [i for i in items if i['item_code'] == item_filter]

    # Run shadow calculations
    results = [shadow_calculate(i, rates) for i in items]

    # Tally
    total   = len(results)
    skipped = sum(1 for r in results if r.status == 'SKIP')
    audited = total - skipped
    passed  = sum(1 for r in results if r.status == 'PASS')
    failed  = sum(1 for r in results if r.status == 'FAIL')
    missing = sum(1 for r in results if r.status == 'MISSING_CODE')
    rate_mm = sum(1 for r in results if r.status == 'RATE_MISMATCH')
    errors  = sum(1 for r in results if r.status == 'ERROR')

    out = []

    # ---- Summary header
    print_section_header('CPWD DAR 2019  |  RATE AUDIT REPORT', out)
    out.append(f'  Total items found  : {total}')
    out.append(f'  SKIP (composite)   : {skipped}  (reference other items as inputs; cannot shadow-calc)')
    out.append(f'  Independently audited : {audited}')
    out.append(f'  PASS          : {passed}  ({100*passed/audited:.1f}%)' if audited else '  (no items)')
    out.append(f'  FAIL          : {failed}')
    out.append(f'  MISSING CODE  : {missing}  (code not in Rates_Master)')
    out.append(f'  RATE MISMATCH : {rate_mm}  (code in Rates_Master but rate differs from PDF)')
    out.append(f'  ERROR         : {errors}')
    out.append('')

    # ---- Detail table
    print_section_header('DETAIL BY ITEM', out)
    print_table_header(out)

    for r in results:
        say_pdf  = _fmt_rate(r.pdf_say)
        say_calc = _fmt_rate(r.shadow_say)
        status_sym = {
            'PASS':         'OK  PASS',
            'FAIL':         'XX  FAIL',
            'SKIP':         '--  SKIP',
            'MISSING_CODE': '??  MISSING_CODE',
            'RATE_MISMATCH':'~~  RATE_MISMATCH',
            'ERROR':        '!!  ERROR',
        }.get(r.status, r.status)

        out.append(
            f"{_trunc(r.sheet, W_SHEET)} {_trunc(r.item_code, W_ITEM)} "
            f"{_trunc(r.description, W_DESC)} "
            f"{say_pdf} {say_calc}  {status_sym}"
        )

        # Verbose: show resource detail for non-passing items
        if verbose and r.status != 'PASS':
            if r.missing_codes:
                out.append(f"         MISSING CODES: {', '.join(r.missing_codes)}")
            if r.rate_mismatches:
                for mm in r.rate_mismatches:
                    out.append(f"         RATE DIFF:    {mm}")
            if not r.say_match and r.shadow_say is not None and r.pdf_say is not None:
                diff = round(r.shadow_say - r.pdf_say, 2)
                out.append(f"         SAY DIFF:     shadow={r.shadow_say}  pdf={r.pdf_say}  delta={diff:+.2f}")
            if r.W_match is False and r.shadow_W and r.pdf_W:
                out.append(f"         W DIFF:       shadow_W={r.shadow_W}  pdf_W={r.pdf_W}")
            out.append('')

    out.append('')

    # ---- Rate mismatch detail
    mm_items = [r for r in results if r.rate_mismatches]
    if mm_items:
        print_section_header('RATE MISMATCHES — Rates_Master vs PDF Analysis Rates', out)
        out.append(f"{'CODE':<8} {'MASTER RATE':>12} {'PDF RATE':>12}  DESCRIPTION")
        out.append('-' * 80)
        seen = set()
        for r in mm_items:
            for mm in r.rate_mismatches:
                if mm not in seen:
                    seen.add(mm)
                    parts = mm.split(':')
                    code = parts[0].strip()
                    rest = parts[1].strip() if len(parts) > 1 else ''
                    entry = rates.get(code, {})
                    out.append(f"{code:<8} {entry.get('rate',0):>12.2f} {rest:>12}  {_trunc(entry.get('desc',''),60)}")
        out.append('')

    # ---- Missing codes detail
    mc_items = [r for r in results if r.missing_codes]
    if mc_items:
        print_section_header('MISSING CODES — Codes in PDF Analysis but not in Rates_Master', out)
        seen = set()
        for r in mc_items:
            for code in r.missing_codes:
                if code not in seen:
                    seen.add(code)
                    out.append(f"  {code}  (first seen in {r.sheet} item {r.item_code})")
        out.append('')

    return out, results


def main():
    parser = argparse.ArgumentParser(description='CPWD DAR 2019 Rate Audit')
    parser.add_argument('--section', help='Filter by sheet prefix, e.g. "02"')
    parser.add_argument('--item',    help='Filter by item code, e.g. "2.2.1"')
    parser.add_argument('--verbose', action='store_true', help='Show resource detail on failures')
    parser.add_argument('--export',  help='Write report to this file path')
    args = parser.parse_args()

    lines, results = run_audit(
        section_filter = args.section,
        item_filter    = args.item,
        verbose        = args.verbose,
    )

    report = '\n'.join(lines)
    print(report)

    if args.export:
        with open(args.export, 'w', encoding='utf-8') as f:
            f.write(report)
        print(f'\nReport written to {args.export}')

    # Exit code: 0 if all pass, 1 if any failures
    any_fail = any(r.status not in ('PASS',) for r in results)
    sys.exit(1 if any_fail else 0)


if __name__ == '__main__':
    main()
