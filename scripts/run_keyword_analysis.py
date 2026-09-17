#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
scripts/run_keyword_analysis.py
===============================
CLI runner for CPWD DSR Volume 1 Progressive Keyword Identification & Diagnostics.

Features:
  - Discovers minimum unique keyword identifiers (1-5 words) across all DSR items.
  - Computes both Global and Chapter-Local identifiers for dynamic trade selectors.
  - Classifies specification ambiguities and subset overlaps (STRUCTURALLY_IMPOSSIBLE).
  - Resolves borderline items via deep-pass combinatorial search (up to 8 words).
  - Generates audit-ready Excel workbook and reference JSON for estimation pipelines.

Usage:
  py -3.11 scripts/run_keyword_analysis.py
  py -3.11 scripts/run_keyword_analysis.py --scope dual --max-words 5
"""
from __future__ import annotations

import os
import sys
import time
import argparse
from pathlib import Path

# Ensure UTF-8 output on Windows consoles
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# Ensure repository root is in sys.path
PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from scripts.paths import (
    DSR_VOL1_SCHEDULE_XLSX,
    KEYWORD_LOOKUP_JSON,
    KEYWORD_ANALYSIS_XLSX,
)
from scripts.keywords import (
    run_keyword_pipeline,
    build_inverted_index,
    classify_and_resolve,
    write_excel,
    write_json,
    print_summary,
    print_sample,
)

DEFAULT_SHEETS = [
    "Carriage of Materials",
    "Earth Work",
    "Mortars",
    "Concrete Work",
    "RCC Work",
    "Masonry Work",
    "Stone Work",
    "Cladding Work",
    "Wood and PVC Work",
    "Steel Work",
    "Flooring",
    "Roofing",
]


def main() -> None:
    parser = argparse.ArgumentParser(
        description="CPWD DSR Volume 1 Progressive Keyword Identification & Diagnostics"
    )
    parser.add_argument(
        "--input", "-i",
        default=DSR_VOL1_SCHEDULE_XLSX,
        help=f"Path to DSR Schedule XLSX (default: {DSR_VOL1_SCHEDULE_XLSX})"
    )
    parser.add_argument(
        "--output-excel", "-oe",
        default=KEYWORD_ANALYSIS_XLSX,
        help=f"Path to output Excel report (default: {KEYWORD_ANALYSIS_XLSX})"
    )
    parser.add_argument(
        "--output-json", "-oj",
        default=KEYWORD_LOOKUP_JSON,
        help=f"Path to output JSON lookup (default: {KEYWORD_LOOKUP_JSON})"
    )
    parser.add_argument(
        "--max-words", "-w",
        type=int,
        default=5,
        help="Maximum keyword search depth for primary pass (default: 5)"
    )
    parser.add_argument(
        "--max-words-deep", "-wd",
        type=int,
        default=8,
        help="Maximum keyword search depth for secondary deep pass (default: 8)"
    )
    parser.add_argument(
        "--scope", "-s",
        choices=["dual", "global", "chapter"],
        default="dual",
        help="Identification scope: dual (both global & chapter), global, or chapter (default: dual)"
    )
    parser.add_argument(
        "--sheets",
        nargs="*",
        default=None,
        help="Specific sheet names to analyze (default: all 12 DSR Vol 1 chapters)"
    )

    args = parser.parse_args()

    print("=" * 70)
    print("  CPWD DSR VOL 1: PROGRESSIVE KEYWORD IDENTIFICATION SYSTEM")
    print("  Integrated Rate Estimation & Lexical Ambiguity Audit Engine")
    print("=" * 70)

    input_path = Path(args.input)
    if not input_path.exists():
        sys.exit(f"[ERROR] DSR Schedule file not found: {input_path}")

    target_sheets = args.sheets if args.sheets else DEFAULT_SHEETS
    t0 = time.perf_counter()

    # Pass 1: Primary Search & Token Extraction
    items = run_keyword_pipeline(
        xlsx_path=str(input_path),
        target_sheets=target_sheets,
        max_words=args.max_words,
        scope=args.scope,
    )

    # Pass 2: Diagnostic Failure Classification & Deep Pass Search
    if args.scope in ("global", "dual"):
        print("\n  Running post-processing failure diagnostics & deep search ...")
        global_index = build_inverted_index(items)
        classify_and_resolve(items, global_index, max_words_deep=args.max_words_deep)

    # Summary and Sample Display
    print_summary(items)
    print_sample(items, n=25)

    # Write Output Artifacts
    print("  Writing output artifacts ...")
    write_excel(items, args.output_excel)
    write_json(items, args.output_json)

    elapsed = time.perf_counter() - t0
    print(f"\n  ✓ Completed analysis in {elapsed:.1f} s")
    print("=" * 70)


if __name__ == "__main__":
    main()
