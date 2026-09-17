"""
audit_ch13_finishing.py
CPWD DAR 2019 – Ch.13 Finishing: Builder vs PDF Audit

Reads every SAY rate written by support_builder_finishing.py into the
13_Finishing worksheet and verifies it against the corresponding rate
extracted directly from CivilDAR_2019_Vol_2.pdf.

Tolerance: ±Rs 0.05 (MROUND rounding band).

Usage:
    python scripts/audit_ch13_finishing.py
    python scripts/audit_ch13_finishing.py --export reports/audit_ch13.md
"""

import sys, os, re, argparse
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import openpyxl, pdfplumber
from scripts.paths import WB_VOL2_FILE, PDF_CIVIL_DAR_VOL2

TOLERANCE = 0.05


# ── 1. Extract SAY rates from the built worksheet ────────────────────────────
def read_sheet_rates(wb_path):
    """Return dict {item_code: sheet_say_rate} from 13_Finishing."""
    wb = openpyxl.load_workbook(str(wb_path), read_only=True, data_only=True)
    if "13_Finishing" not in wb.sheetnames:
        raise ValueError("13_Finishing sheet not found in workbook")
    ws = wb["13_Finishing"]
    rates = {}
    code_pat = re.compile(r'^(13\.\d+(?:\.\d+)?)\s+')
    current_code = None
    for row in ws.iter_rows(values_only=True):
        if not row or row[0] is None:
            continue
        cell_a = str(row[0]).strip()
        m = code_pat.match(cell_a)
        if m:
            current_code = m.group(1)
        # SAY row: column A contains "SAY" text, column G has the rate
        if current_code and "SAY" in cell_a and "PDF published rate" in cell_a:
            try:
                rates[current_code] = float(row[6])
            except (TypeError, ValueError):
                pass
    wb.close()
    return rates


# ── 2. Extract SAY rates from the PDF (pages 10-109, Ch.13) ──────────────────
def read_pdf_rates(pdf_path):
    """Return dict {item_code: pdf_say_rate} from Vol 2 PDF, Ch.13."""
    with pdfplumber.open(str(pdf_path)) as pdf:
        all_text = []
        for i in range(9, 110):
            all_text.append(pdf.pages[i].extract_text() or "")
        full = "\n".join(all_text)

    pattern = re.compile(r"(?m)^(13\.\d+(?:\.\d+)?)\s+(.+?)$")
    segments = list(pattern.finditer(full))
    rates = {}
    for idx, m in enumerate(segments):
        code = m.group(1)
        start = m.end()
        end = segments[idx + 1].start() if idx + 1 < len(segments) else len(full)
        block = full[start:end]
        say_m = re.search(r"\bSay\s+([\d,]+\.?\d*)", block)
        if say_m:
            rates[code] = float(say_m.group(1).replace(",", ""))
    return rates


# ── 3. Compare and format report ─────────────────────────────────────────────
def run_audit(wb_path, pdf_path):
    print("Reading worksheet rates …")
    sheet_rates = read_sheet_rates(wb_path)
    print(f"  {len(sheet_rates)} items read from sheet.")

    print("Reading PDF rates (Ch.13, pages 895-994) …")
    pdf_rates = read_pdf_rates(pdf_path)
    print(f"  {len(pdf_rates)} items read from PDF.")

    all_codes = sorted(set(list(sheet_rates.keys()) + list(pdf_rates.keys())),
                       key=lambda c: [int(x) for x in c.split(".")])

    passed = []
    failed = []
    missing_in_sheet = []
    missing_in_pdf = []

    for code in all_codes:
        in_sheet = code in sheet_rates
        in_pdf = code in pdf_rates
        if not in_sheet:
            missing_in_sheet.append(code)
        elif not in_pdf:
            missing_in_pdf.append(code)
        else:
            diff = abs(sheet_rates[code] - pdf_rates[code])
            if diff <= TOLERANCE:
                passed.append((code, sheet_rates[code], pdf_rates[code], diff))
            else:
                failed.append((code, sheet_rates[code], pdf_rates[code], diff))

    return {
        "passed": passed,
        "failed": failed,
        "missing_in_sheet": missing_in_sheet,
        "missing_in_pdf": missing_in_pdf,
        "total_pdf": len(pdf_rates),
        "total_sheet": len(sheet_rates),
    }


def format_report(result):
    lines = []
    W = 100
    lines.append("=" * W)
    lines.append("  CPWD DAR 2019 – CH.13 FINISHING  |  Builder vs PDF Rate Audit")
    lines.append(f"  Tolerance: Rs {TOLERANCE:.2f}  |  Date: 2026-09-17")
    lines.append("=" * W)
    lines.append(f"  PDF items:    {result['total_pdf']:>4}")
    lines.append(f"  Sheet items:  {result['total_sheet']:>4}")
    lines.append(f"  PASS:         {len(result['passed']):>4}")
    lines.append(f"  FAIL:         {len(result['failed']):>4}")
    lines.append(f"  Missing in sheet: {len(result['missing_in_sheet']):>3}")
    lines.append(f"  Missing in PDF:   {len(result['missing_in_pdf']):>3}")
    lines.append("=" * W)

    if result["failed"]:
        lines.append("")
        lines.append("FAILURES  (sheet rate differs from PDF rate by more than Rs 0.05)")
        lines.append("-" * W)
        lines.append(f"  {'CODE':<14} {'SHEET':>10} {'PDF':>10} {'DIFF':>8}")
        lines.append("-" * W)
        for code, sr, pr, diff in result["failed"]:
            lines.append(f"  {code:<14} {sr:>10.2f} {pr:>10.2f} {diff:>8.2f}")
    else:
        lines.append("")
        lines.append("  No failures – all sheet rates match PDF within tolerance.")

    if result["missing_in_sheet"]:
        lines.append("")
        lines.append("MISSING IN SHEET (PDF has rate but sheet does not)")
        lines.append("-" * W)
        for code in result["missing_in_sheet"]:
            lines.append(f"  {code}")

    if result["missing_in_pdf"]:
        lines.append("")
        lines.append("MISSING IN PDF (sheet has rate but PDF does not – possible extra items)")
        lines.append("-" * W)
        for code in result["missing_in_pdf"]:
            lines.append(f"  {code}")

    lines.append("")
    lines.append("PASS DETAIL")
    lines.append("-" * W)
    lines.append(f"  {'CODE':<14} {'SHEET':>10} {'PDF':>10} {'DIFF':>8}")
    lines.append("-" * W)
    for code, sr, pr, diff in result["passed"]:
        lines.append(f"  {code:<14} {sr:>10.2f} {pr:>10.2f} {diff:>8.2f}")

    lines.append("")
    lines.append("=" * W)
    total = len(result["passed"]) + len(result["failed"])
    pct = 100 * len(result["passed"]) / total if total else 0
    lines.append(f"  RESULT: {len(result['passed'])}/{total} items PASS  ({pct:.1f}%)")
    lines.append("=" * W)
    return "\n".join(lines)


# ── Entry point ───────────────────────────────────────────────────────────────
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--export", metavar="FILE",
                    help="Save report to this path (default: print to stdout)")
    args = ap.parse_args()

    result = run_audit(Path(WB_VOL2_FILE), Path(PDF_CIVIL_DAR_VOL2))
    report = format_report(result)

    if args.export:
        out = Path(args.export)
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(report, encoding="utf-8")
        print(f"Report saved -> {out}")
    else:
        print(report)


if __name__ == "__main__":
    main()
