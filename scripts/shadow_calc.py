# -*- coding: utf-8 -*-
"""
shadow_calc.py
Python replica of the CPWD DAR 2019 markup chain.
Takes an ItemAnalysis dict (from audit_parser) + rates_master dict,
and returns a ShadowResult with computed W, X, Y, Z, Say and
any rate discrepancies found.

Markup chain (standard items):
    W = sum(resource amounts)
    X = W + Water(1% of W)
    Y = X + GST(X × 0.1405)
    Z = Y + CPOH(Y × 0.15)
    Final = Z + Cess(Z × 0.01)
    Say = round(Final / basis_qty, 2)   → then MROUND to 0.05

Component items (no markup flag, e.g. Mortars):
    Cost = sum(resource amounts) / basis_qty
    Say = Cost  (no chain)
"""
import math
from dataclasses import dataclass, field
from typing import List, Optional

# CPWD standard factors (per Global_Factors sheet)
FACTOR_WATER   = 0.01
FACTOR_GST     = 0.1405
FACTOR_CPOH    = 0.15
FACTOR_CESS    = 0.01
FACTOR_SUNDRIES = 2.00  # Rs/unit — rate for code 9999

MROUND_STEP = 0.05   # Excel MROUND step used for say rate


def _mround(val, step=MROUND_STEP):
    """Replicate Excel MROUND(val, step)."""
    if step == 0:
        return val
    return round(round(val / step) * step, 10)


def _r2(v):
    return round(v, 2)


@dataclass
class ResourceResult:
    code: str
    desc: str
    qty: Optional[float]
    pdf_rate: Optional[float]
    master_rate: Optional[float]        # rate in rates_master_clean.json
    pdf_amount: Optional[float]
    shadow_amount: Optional[float]
    rate_match: bool = True             # pdf_rate == master_rate (within 0.01)
    note: str = ''


@dataclass
class ShadowResult:
    sheet: str
    item_code: str
    description: str
    basis_qty: float
    basis_unit: str
    has_markup: bool
    resources: List[ResourceResult] = field(default_factory=list)
    shadow_W: Optional[float] = None
    shadow_say: Optional[float] = None
    pdf_W: Optional[float] = None
    pdf_say: Optional[float] = None
    say_match: bool = False
    W_match: bool = False
    missing_codes: List[str] = field(default_factory=list)
    rate_mismatches: List[str] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)

    skip_reason: str = ''   # set when item cannot be independently audited

    @property
    def status(self):
        if self.skip_reason:
            return 'SKIP'
        if self.errors:
            return 'ERROR'
        if self.missing_codes:
            return 'MISSING_CODE'
        if self.rate_mismatches:
            return 'RATE_MISMATCH'
        if self.say_match:
            return 'PASS'
        return 'FAIL'


def shadow_calculate(item: dict, rates_master: dict) -> ShadowResult:
    """
    item: one dict from audit_parser.parse_all()
    rates_master: {code: {desc, unit, rate}} from rates_master_clean.json
    """
    res = ShadowResult(
        sheet       = item['sheet'],
        item_code   = item['item_code'],
        description = item['description'],
        basis_qty   = item.get('basis_qty', 1.0) or 1.0,
        basis_unit  = item.get('basis_unit', 'unit') or 'unit',
        has_markup  = item.get('has_markup', False),
        pdf_W       = item.get('W'),
        pdf_say     = item.get('say_rate'),
    )

    # Items that reference other items (cross-refs) as inputs cannot be
    # independently shadow-calculated — skip them with a note.
    if item.get('has_cross_ref'):
        res.skip_reason = 'composite: references other items as inputs'
        res.pdf_say = item.get('say_rate')
        return res

    # Items with no resource rows cannot be audited
    if not item.get('resources'):
        res.skip_reason = 'no resource rows parsed'
        return res

    shadow_sum = 0.0
    missing = []
    rate_mm  = []

    for r in item.get('resources', []):
        code       = r['code']
        qty        = r['qty']
        pdf_rate   = r['pdf_rate']
        pdf_amount = r['pdf_amount']

        # Special: 9999 = Sundries at Rs 2.00/unit (variable qty)
        if code == '9999':
            master_rate = FACTOR_SUNDRIES
            shadow_amt  = _r2((qty or 0) * master_rate) if qty is not None else 0.0
            rr = ResourceResult(
                code=code, desc=r.get('desc','Sundries'),
                qty=qty, pdf_rate=pdf_rate, master_rate=master_rate,
                pdf_amount=pdf_amount, shadow_amount=shadow_amt,
                rate_match=(abs((pdf_rate or 2.0) - 2.0) < 0.01),
            )
            shadow_sum += shadow_amt
            res.resources.append(rr)
            continue

        # Look up in rates_master
        entry = rates_master.get(code)
        if entry is None:
            missing.append(code)
            rr = ResourceResult(
                code=code, desc=r.get('desc',''),
                qty=qty, pdf_rate=pdf_rate, master_rate=None,
                pdf_amount=pdf_amount, shadow_amount=None,
                rate_match=False, note='NOT IN RATES_MASTER',
            )
            # Fall back to pdf_amount so the rest of the chain still runs
            shadow_sum += pdf_amount or 0.0
            res.resources.append(rr)
            continue

        master_rate  = entry['rate']
        rate_ok      = (pdf_rate is None) or (abs(pdf_rate - master_rate) < 0.05)
        if not rate_ok:
            rate_mm.append(f"{code}: pdf={pdf_rate} master={master_rate}")

        if qty is not None and master_rate is not None:
            shadow_amt = _r2(qty * master_rate)
        elif pdf_amount is not None:
            shadow_amt = pdf_amount   # use PDF amount as fallback
        else:
            shadow_amt = 0.0

        rr = ResourceResult(
            code=code, desc=r.get('desc', entry.get('desc','')),
            qty=qty, pdf_rate=pdf_rate, master_rate=master_rate,
            pdf_amount=pdf_amount, shadow_amount=shadow_amt,
            rate_match=rate_ok,
        )
        shadow_sum += shadow_amt
        res.resources.append(rr)

    res.missing_codes  = missing
    res.rate_mismatches = rate_mm

    W = _r2(shadow_sum)
    res.shadow_W = W

    if res.pdf_W is not None:
        res.W_match = abs(W - res.pdf_W) < 0.10

    markup_type = item.get('markup_type', 'full' if res.has_markup else 'none')

    if markup_type == 'full':
        # Standard CPWD chain: W → +1% Water → +14.05% GST → +15% CPOH → +1% Cess
        X     = _r2(W + _r2(W * FACTOR_WATER))
        Y     = _r2(X + _r2(X * FACTOR_GST))
        Z     = _r2(Y + _r2(Y * FACTOR_CPOH))
        final = _r2(Z + _r2(Z * FACTOR_CESS))
    elif markup_type == 'cpoh_only':
        # Carriage / labour-only items: W → +15% CPOH only (no Water, GST, Cess)
        final = _r2(W + _r2(W * FACTOR_CPOH))
    else:
        # Component item (mortars, etc.) — raw sum only
        final = W

    # Divide by basis quantity
    unit_rate = _r2(final / res.basis_qty) if res.basis_qty else final

    # Apply MROUND to match "Say" presentation
    say = _mround(unit_rate)
    res.shadow_say = say

    if res.pdf_say is not None:
        # For small rates: allow ±0.10 absolute (MROUND to 0.05 → worst-case 0.05 error)
        # For large rates: allow 0.1% relative (rounding accumulates through chain)
        tol = max(0.10, res.pdf_say * 0.001)
        res.say_match = abs(say - res.pdf_say) <= tol

    return res
