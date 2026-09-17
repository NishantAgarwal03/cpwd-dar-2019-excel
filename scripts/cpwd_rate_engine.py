"""CPWD Custom Rate Analysis Calculation Engine.

Links resolved composition -> gang registry -> productivity registry -> custom unit rate.
Calculates direct cost W, 5-step statutory markups, unit rate, conditional extras,
and benchmark comparison with engineering teaching rationale.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from cpwd_gang_registry import GANG_REGISTRY, ComposedGang, GangResource
from cpwd_productivity_registry import PRODUCTIVITY_REGISTRY, ProductivityEntry


@dataclass
class StatutoryMarkups:
    water_rate: float = 0.01
    gst_rate: float = 0.1405
    cpoh_rate: float = 0.15
    cess_rate: float = 0.01

    def calculate(self, direct_cost_w: float) -> dict[str, float]:
        w = round(direct_cost_w, 2)
        water = round(w * self.water_rate, 2)
        x = round(w + water, 2)
        gst = round(x * self.gst_rate, 2)
        y = round(x + gst, 2)
        cpoh = round(y * self.cpoh_rate, 2)
        z = round(y + cpoh, 2)
        cess = round(z * self.cess_rate, 2)
        total = round(z + cess, 2)
        return {
            "w_direct": w,
            "water": water,
            "x_subtotal": x,
            "gst": gst,
            "y_subtotal": y,
            "cpoh": cpoh,
            "z_subtotal": z,
            "cess": cess,
            "total_cost": total,
        }


@dataclass
class ResourceAnalysisLine:
    code: str
    name: str
    category: str
    coefficient: float
    unit: str
    rate: float
    amount: float
    productivity_summary: str
    teaching_derivation: str


@dataclass
class CustomRateResult:
    composed_scope: str
    base_item_code: str
    base_description: str
    batch_quantity: float
    batch_unit: str
    direct_cost_w: float
    markups: dict[str, float]
    base_unit_rate: float
    conditional_extras: list[dict[str, Any]]
    final_unit_rate: float
    benchmark_item_code: str
    benchmark_unit_rate: float
    variance_amount: float
    variance_percent: float
    variance_explanation: str
    resource_lines: list[ResourceAnalysisLine]
    deducted_lines: list[dict[str, Any]]
    teaching_notes: list[str]

    def summary_text(self) -> str:
        var_sign = "+" if self.variance_amount >= 0 else ""
        return (
            f"Custom Unit Rate: Rs {self.final_unit_rate:.2f} per {self.batch_unit} "
            f"(Benchmark {self.benchmark_item_code}: Rs {self.benchmark_unit_rate:.2f} | "
            f"Variance: {var_sign}Rs {self.variance_amount:.2f} / {var_sign}{self.variance_percent:.1f}%). "
            f"Direct Cost: Rs {self.direct_cost_w:.2f}, Markups: Rs {self.markups['total_cost'] - self.direct_cost_w:.2f} per batch."
        )


def calculate_custom_rate(
    resolved_composition: dict[str, Any],
    markups: StatutoryMarkups | None = None,
) -> CustomRateResult:
    """Calculate the custom unit rate from resolved keywords composition."""
    if markups is None:
        markups = StatutoryMarkups()

    base = resolved_composition["base"]
    base_code = base["item_code"]
    deduction_codes = [d["item_code"] for d in resolved_composition.get("deductions", [])]
    addition_codes = [a["item_code"] for a in resolved_composition.get("additions", [])]
    conditional_extras = resolved_composition.get("conditional_extras", [])

    # 1. Compose Gang
    composed: ComposedGang = GANG_REGISTRY.compose_gang(base_code, deduction_codes, addition_codes)
    batch_qty = composed.batch_quantity
    batch_unit = composed.batch_unit

    # 2. Build Resource Lines with Productivity derivations
    resource_lines: list[ResourceAnalysisLine] = []
    for res in composed.all_active_resources:
        prod_entry = PRODUCTIVITY_REGISTRY.lookup("02_Earth_Work", base_code, res.code)
        if prod_entry and prod_entry.daily_productivity:
            prod_summary = f"{prod_entry.daily_productivity} {prod_entry.metric_unit}"
            derivation = prod_entry.derivation
        elif prod_entry and prod_entry.hourly_productivity:
            prod_summary = f"{prod_entry.hourly_productivity} {prod_entry.metric_unit}"
            derivation = prod_entry.derivation
        else:
            prod_summary = f"{res.coefficient} {res.unit} / batch"
            derivation = f"Direct allowance: {res.coefficient} {res.unit} per {batch_qty} {batch_unit}"

        resource_lines.append(
            ResourceAnalysisLine(
                code=res.code,
                name=res.name,
                category=res.category,
                coefficient=res.coefficient,
                unit=res.unit,
                rate=res.rate,
                amount=res.amount,
                productivity_summary=prod_summary,
                teaching_derivation=derivation,
            )
        )

    # 3. Direct Cost W & Markups
    direct_w = composed.direct_cost_w
    markup_dict = markups.calculate(direct_w)
    base_batch_total = markup_dict["total_cost"]
    base_unit_rate = round(base_batch_total / batch_qty, 2)

    # 4. Conditional Extras (e.g. 2.24.1 +20%, 2.24.2 +25% on qualifying base rate)
    calculated_extras: list[dict[str, Any]] = []
    extra_rate_sum = 0.0
    for extra in conditional_extras:
        pct = float(extra.get("percent", 0.0))
        extra_unit_val = round(base_unit_rate * pct / 100.0, 2)
        extra_rate_sum += extra_unit_val
        calculated_extras.append({
            **extra,
            "calculated_extra_rate": extra_unit_val,
            "unit": base_unit_rate,
        })

    final_unit_rate = round(base_unit_rate + extra_rate_sum, 2)

    # 5. Benchmark Comparison
    # Calculate unamended base DAR rate under the same markup chain
    base_direct_w = composed.base_item.direct_cost_w
    base_markups = markups.calculate(base_direct_w)
    base_dar_rate = round(base_markups["total_cost"] / batch_qty, 2)

    benchmark_unit_rate = base_dar_rate
    dsr_benchmark = composed.base_item.dsr_rate or benchmark_unit_rate

    variance_amount = round(final_unit_rate - benchmark_unit_rate, 2)
    variance_pct = round((variance_amount / benchmark_unit_rate) * 100.0, 1) if benchmark_unit_rate else 0.0

    # 6. Explanations & Teaching Notes
    teaching_notes: list[str] = []
    variance_parts: list[str] = []

    if deduction_codes:
        deduct_names = []
        for d in resolved_composition.get("deductions", []):
            code = d["item_code"]
            if code == "2.4":
                deduct_names.append("omitted 8-t power roller & chowkidar")
                teaching_notes.append(
                    "Omission of 8-tonne road roller (Item 2.4): removes 0.008 roller-day, "
                    "0.008 chowkidar-day and 1.82 sundries. Remaining compaction is by manual wooden/steel rammers."
                )
            elif code == "2.5":
                deduct_names.append("omitted watering to OMC")
                teaching_notes.append(
                    "Omission of watering (Item 2.5): removes 0.40 Bhishti-day (3.2 hours of watering). "
                    "Applicable only where contractor did not moisten earth to optimum moisture content."
                )
            else:
                deduct_names.append(f"deduction {code}")
        variance_parts.append(f"Scope deductions: {', '.join(deduct_names)}")

    if addition_codes:
        add_names = [a["item_code"] for a in resolved_composition.get("additions", [])]
        variance_parts.append(f"Scope additions: {', '.join(add_names)}")
        teaching_notes.append(f"Operations added to base scope: {', '.join(add_names)}.")

    if calculated_extras:
        for ex in calculated_extras:
            variance_parts.append(f"+{ex['percent']}% extra for {ex['condition'][:30]}")
            teaching_notes.append(
                f"Conditional extra {ex['item_code']}: +{ex['percent']}% applied only to qualifying quantity "
                f"executed {ex['condition']}. Depth measured from sub-soil water level to centre of gravity."
            )

    variance_explanation = " | ".join(variance_parts) if variance_parts else "Standard base scope without modifications."

    deducted_lines = [
        {"code": r.code, "name": r.name, "category": r.category, "coefficient": r.coefficient, "unit": r.unit, "rate": r.rate}
        for r in composed.deducted_resources
    ]

    return CustomRateResult(
        composed_scope=resolved_composition.get("composed_scope", ""),
        base_item_code=base_code,
        base_description=composed.base_item.description,
        batch_quantity=batch_qty,
        batch_unit=batch_unit,
        direct_cost_w=direct_w,
        markups=markup_dict,
        base_unit_rate=base_unit_rate,
        conditional_extras=calculated_extras,
        final_unit_rate=final_unit_rate,
        benchmark_item_code=base_code,
        benchmark_unit_rate=benchmark_unit_rate,
        variance_amount=variance_amount,
        variance_percent=variance_pct,
        variance_explanation=variance_explanation,
        resource_lines=resource_lines,
        deducted_lines=deducted_lines,
        teaching_notes=teaching_notes,
    )
