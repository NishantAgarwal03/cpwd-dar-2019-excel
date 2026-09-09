# -*- coding: utf-8 -*-
"""
audit_parser.py
Parse CivilDAR_2019_Vol_*_Converted.xlsx to extract every item analysis:
  - item code, description, basis qty+unit
  - resource rows (code, desc, qty, pdf_rate, pdf_amount)
  - whether the item has a W→X→Y→Z markup chain
  - published W total and Say rate

Returns a list of ItemAnalysis dicts consumed by run_audit.py.
"""
import re
import openpyxl

# ------------------------------------------------------------------
# Helpers
# ------------------------------------------------------------------

_ITEM_CODE_RE = re.compile(r'^\d+\.\d+')          # e.g. "2.2.1", "3.1", "4.1.2a"
_RESOURCE_CODE_RE = re.compile(r'^\d{4}$')        # exactly 4 digits

# Matches "Details of cost for/of/per X unit" and bare "for X unit"
_BASIS_RE = re.compile(
    r'(?:for|per|of)\s+([\d,.]+)\s*(sqm|cum|nos|kg|tonne|m|rmt|quintal|litre|day|trip|set|pair)',
    re.IGNORECASE
)
# "Cost of 23 tonne" rows that carry the aggregate-to-unit conversion
_COST_OF_RE = re.compile(
    r'cost\s+of\s+([\d,.]+)\s*(sqm|cum|nos|kg|tonne|m|rmt|quintal|litre|day|trip|set|pair)',
    re.IGNORECASE
)

def _to_float(v):
    if v is None:
        return None
    try:
        return float(str(v).replace(',', '').strip())
    except (ValueError, TypeError):
        return None


def _row_text(row):
    """Return all non-empty cell strings joined."""
    return ' '.join(str(c).strip() for c in row if str(c).strip())


def _is_item_header(col_b):
    """True if col B looks like a dotted item code."""
    return bool(_ITEM_CODE_RE.match(str(col_b or '').strip()))


def _is_resource(col_b, col_e, col_f, col_g):
    """True if row looks like a resource line (4-digit code or 9999, numeric amounts)."""
    code = str(col_b or '').strip()
    if not (_RESOURCE_CODE_RE.match(code) or code == '9999'):
        return False
    return _to_float(col_e) is not None or _to_float(col_g) is not None


def _extract_basis(text):
    """Return (basis_qty, basis_unit) from a 'Details of cost for X cum' string."""
    m = _BASIS_RE.search(text)
    if m:
        qty = _to_float(m.group(1))
        unit = m.group(2).lower()
        return qty, unit
    return 1.0, 'unit'


def _extract_say(col_b, row):
    """If row is a Say row, return the rate value; else None."""
    s = str(col_b or '').strip().lower()
    if s == 'say':
        for v in row[2:]:
            f = _to_float(v)
            if f is not None and f > 0:
                return f
    return None


def _extract_W(col_b, row):
    """If row is the W-total row, return W value."""
    s = str(col_b or '').strip()
    # e.g. "TOTAL"  with the W value embedded in the text of another cell
    if s.upper().startswith('TOTAL'):
        for v in row[2:]:
            vs = str(v or '').strip()
            # "5581.72 W" or "5581.72" in the amount column
            m = re.search(r'([\d,]+\.?\d*)\s*W?', vs)
            if m:
                f = _to_float(m.group(1))
                if f and f > 1:
                    return f
        # Also check col G (index 6) directly
        f = _to_float(row[6] if len(row) > 6 else None)
        if f and f > 1:
            return f
    # "TOTAL  5389.29 W" packed into col_b itself
    m = re.search(r'([\d,]+\.?\d*)\s*W', s)
    if m:
        return _to_float(m.group(1))
    return None


# ------------------------------------------------------------------
# Main parser
# ------------------------------------------------------------------

def parse_sheet(ws, sheet_name):
    """
    Parse one worksheet and return a list of item dicts.
    Each dict:
      sheet, item_code, description, basis_qty, basis_unit,
      resources: [{code, desc, qty, pdf_rate, pdf_amount}],
      markup_type: 'full' | 'cpoh_only' | 'none'
      has_markup (bool, backward-compat = markup_type != 'none'),
      W, say_rate
    """
    items = []
    current      = None
    parent_desc  = ''    # description of the nearest parent item code (e.g. "1.4")
    header_text_buf = []

    def _finalise(item):
        if item and item.get('say_rate') is not None:
            # Resolve markup_type from what we saw
            if item.get('_saw_water') and item.get('_saw_gst'):
                item['markup_type'] = 'full'
            elif item.get('_saw_cpoh') and not item.get('_saw_water'):
                item['markup_type'] = 'cpoh_only'
            else:
                item['markup_type'] = 'none'
            item['has_markup'] = item['markup_type'] != 'none'
            # Clean internal flags
            for k in ('_saw_water', '_saw_gst', '_saw_cpoh', '_saw_cess'):
                item.pop(k, None)
            items.append(item)

    def _new_item(code, desc, parent):
        """Create a fresh item dict, prepending parent description if present."""
        full_desc = f'{parent}: {desc}' if parent and desc and parent not in desc else desc
        return {
            'sheet':          sheet_name,
            'item_code':      code,
            'description':    full_desc.strip(),
            'basis_qty':      1.0,
            'basis_unit':     'unit',
            'resources':      [],
            'has_cross_ref':  False,
            'markup_type':    'none',
            'has_markup':     False,
            '_saw_water':     False,
            '_saw_gst':       False,
            '_saw_cpoh':      False,
            '_saw_cess':      False,
            'W':              None,
            'say_rate':       None,
        }

    rows = list(ws.iter_rows(values_only=True))

    for row in rows:
        if len(row) < 8:
            row = list(row) + [None] * (8 - len(row))

        col_b = str(row[1] or '').strip()
        col_c = str(row[2] or '').strip()
        col_e = row[4]
        col_f = row[5]
        col_g = row[6]
        rtext = _row_text(row)

        # ---- Say row → closes an item --------------------------------
        say = _extract_say(col_b, row)
        if say is not None and current is not None:
            current['say_rate'] = say
            _finalise(current)
            current = None
            header_text_buf = []
            continue

        # ---- Item header row -----------------------------------------
        if _is_item_header(col_b):
            # Cross-reference lines look like item codes but description says
            # "Rate as per item no. X of SH: …" — not a new item.
            if re.match(r'rate\s+as\s+per', col_c, re.IGNORECASE):
                if current is not None:
                    current['has_cross_ref'] = True
                continue

            # Determine depth: parent item has fewer dot-segments
            # e.g. "1.4" is parent of "1.4.1"
            depth = col_b.count('.')
            if depth == 1:
                # This is a parent-level item (e.g. "1.4", "2.10")
                # Save its description for child items; only open a new current
                # if it has resources of its own (usually it doesn't).
                _finalise(current)
                current = None
                parent_desc = col_c
                # Start a new current anyway; if it gets no resources it won't finalise
                current = _new_item(col_b, col_c, '')
                basis_q, basis_u = _extract_basis(rtext)
                current['basis_qty'] = basis_q
                current['basis_unit'] = basis_u
            else:
                # Child item (e.g. "1.4.1", "2.10.1.1")
                _finalise(current)
                current = _new_item(col_b, col_c, parent_desc)
                basis_q, basis_u = _extract_basis(rtext)
                current['basis_qty'] = basis_q
                current['basis_unit'] = basis_u
            continue

        if current is None:
            continue

        # ---- Accumulate basis qty from "Details of cost for/of X unit" rows ---
        rlow = rtext.lower()
        if 'details of cost' in rlow or 'details of rate' in rlow:
            bq, bu = _extract_basis(rtext)
            if bq != 1.0:
                current['basis_qty'] = bq
                current['basis_unit'] = bu
            # Also accumulate into description if it adds context
            if col_c and col_c not in current['description']:
                current['description'] = (current['description'] + ' — ' + col_c).strip(' — ')
            continue

        # ---- "Cost of X tonne" rows → update basis qty ---------------
        m_cost = _COST_OF_RE.search(rtext)
        if m_cost and col_b.lower().startswith('cost of'):
            qty_v = _to_float(m_cost.group(1))
            unit_v = m_cost.group(2).lower()
            if qty_v and qty_v > 1:
                current['basis_qty']  = qty_v
                current['basis_unit'] = unit_v
            continue

        # ---- Resource row --------------------------------------------
        if _is_resource(col_b, col_e, col_f, col_g):
            qty    = _to_float(col_e)
            rate   = _to_float(col_f)
            amount = _to_float(col_g)
            if amount is None and qty is not None and rate is not None:
                amount = round(qty * rate, 2)
            current['resources'].append({
                'code':       col_b,
                'desc':       col_c,
                'qty':        qty,
                'pdf_rate':   rate,
                'pdf_amount': amount,
            })
            continue

        # ---- TOTAL / W row ------------------------------------------
        w_val = _extract_W(col_b, row)
        if w_val is not None and current['W'] is None:
            current['W'] = w_val
            continue

        # ---- Markup type detection ----------------------------------
        if re.search(r'water\s+charges', rtext, re.IGNORECASE):
            current['_saw_water'] = True
        if re.search(r'\bGST\b', rtext, re.IGNORECASE):
            current['_saw_gst'] = True
        if re.search(r'\bCPOH\b', rtext, re.IGNORECASE):
            current['_saw_cpoh'] = True
        if re.search(r'\bCess\b', rtext, re.IGNORECASE):
            current['_saw_cess'] = True

    _finalise(current)
    return items


def parse_all(vol1_path, vol2_path):
    """Return all parsed items from both converted workbooks."""
    all_items = []
    for path, label in [(vol1_path, 'V1'), (vol2_path, 'V2')]:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        for sname in wb.sheetnames:
            if sname == '00_Basic_Rates':
                continue
            ws = wb[sname]
            items = parse_sheet(ws, sname)
            all_items.extend(items)
        wb.close()
    return all_items
