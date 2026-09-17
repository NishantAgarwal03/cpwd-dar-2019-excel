"""extract_norms.py
Builds labour_productivity.json and sundries_reference.json directly from the
PDF-verified Ch01/03-12 item JSONs (scripts/vol1_chapter_extractor_pdf.py) and
Ch02's PDF-verified support_builder_earth_v2.py ITEMS list, instead of the
unverifiable data/converted_xlsx/CivilDAR_2019_Vol_1_Converted.xlsx (see
reports/audit_architecture_process_2026-09-17.md, finding D2, and
PIPELINE.md's "known gap" note this replaces).

Every LABOUR (code 01xx)/MACHINERY (code 00xx) resource row across every item
becomes a productivity record; every code-9999 (or "sundries"-described)
resource row becomes a sundries record. This is a pure re-shaping of resource
rows that are already independently PDF-verified per chapter -- nothing new
is invented here.
"""
import sys, os, json, re
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir)))
from scripts.paths import REFERENCE_JSON_DIR, LABOUR_PRODUCTIVITY_JSON, SUNDRIES_REFERENCE_JSON

SUBHEAD_NAMES = {
    1: '01_Carriage_of_Materials',
    2: '02_Earth_Work',
    3: '03_Mortars',
    4: '04_Concrete_Work',
    5: '05_RCC_Work',
    6: '06_Masonry_Work',
    7: '07_Stone_Work',
    8: '08_Cladding_Work',
    9: '09_Wood_and_PVC_Work',
    10: '10_Steel_Work',
    11: '11_Flooring',
    12: '12_Roofing',
}

# PDF-direct chapter item JSONs (Ch02 is handled separately, below, since its
# canonical source is the hand-verified support_builder_earth_v2.py ITEMS).
_JSON_CHAPTERS = [1, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]

_LABOUR_MACHINERY_PAT = re.compile(r'^(01\d{2}|00\d{2})$')


def _res_type(code: str) -> str:
    return 'Labour' if code.startswith('01') else 'Machinery'


def _emit(subhead, item_no, item_desc, basis, code, desc, unit, qty, rate,
          productivity_records, sundries_records):
    code = str(code or '')
    desc = desc or ''
    if _LABOUR_MACHINERY_PAT.match(code):
        productivity_records.append({
            'subhead': subhead,
            'item_no': item_no,
            'item_desc': item_desc,
            'basis': basis,
            'code': code,
            'description': desc,
            'unit': unit,
            'coefficient': qty,
            'rate': rate,
            'type': _res_type(code),
        })
    if code == '9999' or 'sundries' in desc.lower():
        sundries_records.append({
            'subhead': subhead,
            'item_no': item_no,
            'item_desc': item_desc,
            'basis': basis,
            'description': desc if desc else 'Sundries',
            'base_ls': qty,
            'multiplier': 2.0,
            'amount': round((qty or 0.0) * (rate or 0.0), 2),
        })


def main():
    productivity_records = []
    sundries_records = []

    for ch in _JSON_CHAPTERS:
        path = os.path.join(REFERENCE_JSON_DIR, f'ch{ch:02d}_items.json')
        with open(path, encoding='utf-8') as f:
            items = json.load(f)
        subhead = SUBHEAD_NAMES[ch]
        for item in items:
            item_no = item.get('code', '')
            item_desc = (item.get('desc') or '')[:100]
            basis = f"{item.get('basis', 1.0):g} {item.get('unit', 'nos')}"
            for res in item.get('resources', []):
                _emit(subhead, item_no, item_desc, basis,
                      res.get('code'), res.get('desc'), res.get('unit'),
                      res.get('qty'), res.get('rate'),
                      productivity_records, sundries_records)

    # Ch02: from support_builder_earth_v2.py's ITEMS (dsr_rate literals and
    # the 2.16.x family verified against the PDF 2026-09-18). Composite
    # "REFERENCE ITEMS" rows use synthetic codes like "REF#2.8.1" and simply
    # never match the 4-digit Labour/Machinery/Sundries patterns above.
    sys.path.insert(0, os.path.join(os.path.dirname(__file__)))
    import support_builder_earth_v2 as _earth

    subhead = SUBHEAD_NAMES[2]
    for item in _earth.ITEMS:
        item_no = item.get('id', '')
        item_desc = (item.get('desc') or '')[:100]
        unit = item.get('unit', 'nos')
        base_qty = item.get('base_qty', 1.0) or 1.0
        basis = f"{base_qty:g} {unit}"
        for _section_name, rows in item.get('sections', []):
            for row in rows:
                rcode, rdesc, runit, rqty, rrate = row
                _emit(subhead, item_no, item_desc, basis,
                      rcode, rdesc, runit, rqty, rrate,
                      productivity_records, sundries_records)

    print(f'Mined {len(productivity_records)} productivity records.')
    print(f'Mined {len(sundries_records)} sundries records.')

    with open(LABOUR_PRODUCTIVITY_JSON, 'w', encoding='utf-8') as f:
        json.dump(productivity_records, f, indent=2)

    with open(SUNDRIES_REFERENCE_JSON, 'w', encoding='utf-8') as f:
        json.dump(sundries_records, f, indent=2)

    print('Saved reference data successfully.')


if __name__ == '__main__':
    main()
