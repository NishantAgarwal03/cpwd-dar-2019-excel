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
_BASIS_RE = re.compile(
    r'(?:for|per)\s+([\d,.]+)\s*(sqm|cum|nos|kg|tonne|m|rmt|quintal|litre|day|trip|set|pair)',
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
      has_markup, W, say_rate
    """
    items = []
    current = None
    header_text_buf = []   # multi-row description accumulation

    def _finalise(item):
        if item and item.get('say_rate') is not None:
            items.append(item)

    rows = list(ws.iter_rows(values_only=True))

    for row in rows:
        if len(row) < 8:
            row = list(row) + [None] * (8 - len(row))

        col_b = str(row[1] or '').strip()
        col_c = str(row[2] or '').strip()
        col_e = row[4]
        col_f = row[5]
        col_g = row[6]

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
            # Cross-reference lines look like item codes but their description
            # says "Rate as per item no. X of SH: …" — they are NOT new items;
            # they are resource references inside the current item.
            if re.match(r'rate\s+as\s+per', col_c, re.IGNORECASE):
                # Treat as a cross-reference resource (mark with prefix 'XREF')
                if current is not None:
                    current['has_cross_ref'] = True
                continue

            _finalise(current)
            current = {
                'sheet': sheet_name,
                'item_code': col_b,
                'description': col_c,
                'basis_qty': 1.0,
                'basis_unit': 'unit',
                'resources': [],
                'has_markup': False,
                'has_cross_ref': False,
                'W': None,
                'say_rate': None,
            }
            header_text_buf = [col_c]
            # Try to extract basis from this row's text
            basis_q, basis_u = _extract_basis(_row_text(row))
            current['basis_qty'] = basis_q
            current['basis_unit'] = basis_u
            continue

        if current is None:
            continue

        # ---- Accumulate basis qty from continuation rows -------------
        rtext = _row_text(row)
        if 'details of cost' in rtext.lower() or 'details of rate' in rtext.lower():
            bq, bu = _extract_basis(rtext)
            if bq != 1.0:
                current['basis_qty'] = bq
                current['basis_unit'] = bu
            continue

        # ---- Resource row --------------------------------------------
        if _is_resource(col_b, col_e, col_f, col_g):
            qty    = _to_float(col_e)
            rate   = _to_float(col_f)
            amount = _to_float(col_g)
            # If amount missing but qty+rate present, compute it
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

        # ---- Markup presence ----------------------------------------
        if re.search(r'water charges|GST|CPOH|Cess', rtext, re.IGNORECASE):
            current['has_markup'] = True

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
