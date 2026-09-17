"""run_pipeline.py — CPWD DAR 2019 master pipeline runner.

Runs every step in dependency order. Each step is idempotent — re-running
it overwrites the previous output with a fresh one.

Usage:
    python run_pipeline.py              # full pipeline
    python run_pipeline.py --step 3     # single step by number
    python run_pipeline.py --from 4     # from step 4 onward
    python run_pipeline.py --list       # print step list and exit

Steps
-----
 1  extract_rates       PDF → rates_master_clean.json
 2  vol1_extract        Vol 1 PDF → ch01,03-12_items.json     (Vol 1 chapter JSONs)
 3  extract_norms       ch01,03-12_items.json + Ch02 ITEMS → labour_productivity.json, sundries_reference.json
 4  vol2_extract        Vol 2 PDF → ch13-26_items.json        (Vol 2 chapter JSONs)
 5  registry_vol1       ch02-12 sources → gang/productivity_registry_vol1.txt/.ini
 6  registry_vol2       ch13-26_items.json → gang/productivity_registry_vol2.txt/.ini
 7  workbook_vol1       all sources → CPWD_DAR_2019_Vol1_Workbook.xlsx
 8  workbook_vol2       ch13-26_items.json → CPWD_DAR_2019_Vol2_Workbook.xlsx

Note: extract_norms (step 3) now reads vol1_extract's (step 2) output JSONs
directly, so vol1_extract must run first -- this is the opposite order from
before 2026-09-18, when both independently read the same converted XLSX.
"""

from __future__ import annotations

import argparse
import importlib
import sys
import time
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

BASE = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE))


def _step_extract_rates():
    import scripts.extract_rates as m
    m.main()


def _step_extract_norms():
    import scripts.extract_norms as m
    m.main()


def _step_vol1_extract():
    import scripts.vol1_chapter_extractor_pdf as m
    m.main(["--all"])


def _step_vol2_extract():
    import scripts.vol2_chapter_extractor as m
    m.main(["--all"])


def _step_registry_vol1():
    from scripts.cpwd_item_parser import build_vol1_registries
    build_vol1_registries()


def _step_registry_vol2():
    from scripts.cpwd_item_parser_vol2 import build_vol2_registries
    build_vol2_registries()


def _step_workbook_vol1():
    import scripts.main as m
    m.generate_full_workbook()


def _step_workbook_vol2():
    import scripts.generate_vol2_workbook as m
    m.generate_vol2_workbook()


STEPS: list[tuple[str, str, callable]] = [
    ("extract_rates",  "PDF → rates_master_clean.json",                     _step_extract_rates),
    ("vol1_extract",   "Vol 1 PDF → ch01,03-12_items.json",                 _step_vol1_extract),
    ("extract_norms",  "ch01,03-12_items.json + Ch02 ITEMS → labour/sundries JSON", _step_extract_norms),
    ("vol2_extract",   "Vol 2 PDF → ch13-26_items.json",                    _step_vol2_extract),
    ("registry_vol1",  "ch02-12 → gang/productivity_registry_vol1.*",       _step_registry_vol1),
    ("registry_vol2",  "ch13-26 → gang/productivity_registry_vol2.*",       _step_registry_vol2),
    ("workbook_vol1",  "All sources → Vol 1 workbook",                      _step_workbook_vol1),
    ("workbook_vol2",  "ch13-26 → Vol 2 workbook",                          _step_workbook_vol2),
]


def run_step(i: int) -> None:
    name, desc, fn = STEPS[i]
    print(f"\n[{i+1}/{len(STEPS)}] {name} — {desc}", flush=True)
    t0 = time.time()
    fn()
    print(f"  ✓ done in {time.time()-t0:.1f}s", flush=True)


def main() -> None:
    ap = argparse.ArgumentParser(description="CPWD DAR 2019 pipeline runner")
    grp = ap.add_mutually_exclusive_group()
    grp.add_argument("--step",   type=int, help="Run single step number (1-based)")
    grp.add_argument("--from",   dest="from_step", type=int, help="Run from step N onward")
    grp.add_argument("--list",   action="store_true", help="List steps and exit")
    args = ap.parse_args()

    if args.list:
        for i, (name, desc, _) in enumerate(STEPS, 1):
            print(f"  {i:2d}  {name:<20s}  {desc}")
        return

    if args.step:
        run_step(args.step - 1)
    elif args.from_step:
        for i in range(args.from_step - 1, len(STEPS)):
            run_step(i)
    else:
        for i in range(len(STEPS)):
            run_step(i)


if __name__ == "__main__":
    main()
