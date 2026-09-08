# -*- coding: utf-8 -*-
"""
Work-type classification for the `Labour_Machinery_Productivity` sheet.

The Problem & Solution Statement (section 2.2) requires this reference to be
organised by WORK TYPE -- "what does brickwork need" -- rather than by the
sub-head sheet a record happened to be printed on. This module derives that
dimension from each record's sub-head plus keywords in its item nomenclature
and resource description.

Nothing is invented: the classification only re-groups records that already
exist in the base volumes.
"""

# Ordered list of (work_type, trade_group). Order fixes the display sequence.
WORK_TYPE_ORDER = [
    ('Carriage - Mechanical Transport',      'Carriage & Handling'),
    ('Carriage - Manual Handling',           'Carriage & Handling'),
    ('Loading / Unloading at Siding',        'Carriage & Handling'),
    ('Earthwork - Excavation',               'Earthwork'),
    ('Earthwork - Filling, Banking & Compaction', 'Earthwork'),
    ('Earthwork - Rock Cutting',             'Earthwork'),
    ('Mortar Mixing',                        'Mortars'),
    ('Concreting - Plain Cement Concrete',   'Concrete'),
    ('Concreting - Reinforced Cement Concrete', 'Concrete'),
    ('Brickwork & Block Masonry',            'Masonry'),
    ('Stone Masonry',                        'Stone'),
    ('Stone / Marble Cladding & Veneer Fixing', 'Cladding'),
    ('Carpentry & Joinery',                  'Wood & PVC'),
    ('PVC / Plastic Fabrication & Fixing',   'Wood & PVC'),
    ('Steel Fabrication & Fixing',           'Steel'),
    ('Roofing & Sheeting',                   'Roofing'),
    ('Flooring - Laying & Bedding',          'Flooring'),
    ('Flooring - Rubbing, Polishing & Finishing', 'Flooring'),
    ('Painting & Priming',                   'Finishing'),
    ('General / Unclassified',               'General'),
]

WORK_TYPE_GROUP = dict(WORK_TYPE_ORDER)
WORK_TYPE_INDEX = {wt: i for i, (wt, _) in enumerate(WORK_TYPE_ORDER)}


def _has(text, *words):
    return any(w in text for w in words)


def classify(subhead, item_desc, resource_desc):
    """Return the work type for one productivity record."""
    d = (item_desc or '').lower()
    res = (resource_desc or '').lower()
    sh = subhead or ''

    # Painting / priming labour shows up inside several trades (roofing,
    # steel, wood). Resource-led, so it is checked before the sub-head rules.
    if _has(res, 'painter') or _has(d, 'priming coat', 'painting with', 'primer'):
        return 'Painting & Priming'

    if sh.startswith('01'):
        if _has(d, 'manual labour', 'by manual') or _has(d, 'beldars (male/female)'):
            return 'Carriage - Manual Handling'
        if _has(d, 'loading in or unloading', 'railway wagon', 'siding'):
            return 'Loading / Unloading at Siding'
        return 'Carriage - Mechanical Transport'

    if sh.startswith('02'):
        if _has(d, 'rock'):
            return 'Earthwork - Rock Cutting'
        if _has(d, 'filling', 'banking', 'rolling', 'compact', 'watering', 'dressing', 'consolidat'):
            return 'Earthwork - Filling, Banking & Compaction'
        return 'Earthwork - Excavation'

    if sh.startswith('03'):
        return 'Mortar Mixing'

    if sh.startswith('04'):
        return 'Concreting - Plain Cement Concrete'

    if sh.startswith('05'):
        return 'Concreting - Reinforced Cement Concrete'

    if sh.startswith('06'):
        return 'Brickwork & Block Masonry'

    if sh.startswith('07'):
        return 'Stone Masonry'

    if sh.startswith('08'):
        return 'Stone / Marble Cladding & Veneer Fixing'

    if sh.startswith('09'):
        if _has(d, 'pvc', 'plastic', 'polyvinyl'):
            return 'PVC / Plastic Fabrication & Fixing'
        return 'Carpentry & Joinery'

    if sh.startswith('10'):
        return 'Steel Fabrication & Fixing'

    if sh.startswith('11'):
        if _has(d, 'polish', 'rubbing', 'grinding', 'hardener', 'float finish'):
            return 'Flooring - Rubbing, Polishing & Finishing'
        if _has(res, 'rubbing', 'polishing'):
            return 'Flooring - Rubbing, Polishing & Finishing'
        return 'Flooring - Laying & Bedding'

    if sh.startswith('12'):
        return 'Roofing & Sheeting'

    return 'General / Unclassified'


def median(values):
    vs = sorted(values)
    n = len(vs)
    if n == 0:
        return 0.0
    mid = n // 2
    if n % 2:
        return vs[mid]
    return (vs[mid - 1] + vs[mid]) / 2.0


def summarise(records):
    """Roll records up to one row per (work type, resource type, code).

    Returns a list of dicts sorted by work type order then descending
    observation count, i.e. the most representative crew members first.
    """
    buckets = {}
    for rec in records:
        wt = rec['work_type']
        key = (wt, rec['type'], rec['code'])
        b = buckets.setdefault(key, {
            'work_type': wt,
            'trade_group': WORK_TYPE_GROUP[wt],
            'res_type': rec['type'],
            'code': rec['code'],
            'description': rec['description'],
            'unit': rec['unit'],
            'rate': rec['rate'],
            'coeffs': [],
            'bases': {},
            'subheads': set(),
        })
        # Zero coefficients are extraction artefacts of blank continuation
        # rows; they would drag every median to zero.
        if rec['coefficient'] and rec['coefficient'] > 0:
            b['coeffs'].append(rec['coefficient'])
        basis = rec.get('basis') or ''
        b['bases'][basis] = b['bases'].get(basis, 0) + 1
        b['subheads'].add(rec['subhead'])

    out = []
    for b in buckets.values():
        if not b['coeffs']:
            continue
        modal_basis = max(b['bases'].items(), key=lambda kv: kv[1])[0] if b['bases'] else ''
        out.append({
            'work_type': b['work_type'],
            'trade_group': b['trade_group'],
            'res_type': b['res_type'],
            'code': b['code'],
            'description': b['description'],
            'unit': b['unit'],
            'rate': b['rate'],
            'typical': round(median(b['coeffs']), 3),
            'low': round(min(b['coeffs']), 3),
            'high': round(max(b['coeffs']), 3),
            'observations': len(b['coeffs']),
            'basis': modal_basis,
            'subheads': ', '.join(sorted(b['subheads'])),
        })

    out.sort(key=lambda x: (WORK_TYPE_INDEX[x['work_type']],
                            0 if x['res_type'] == 'Labour' else 1,
                            -x['observations'],
                            x['code']))
    return out
