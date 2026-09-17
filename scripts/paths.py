# -*- coding: utf-8 -*-
"""
scripts/paths.py
================
Centralized path manager and resolver for CPWD DAR 2019 codebase.
Anchors all relative paths to the project repository root and transparently
supports migrated subdirectories (data/raw_pdf, data/reference_json, data/converted_xlsx, output)
with zero-breakage fallbacks.
"""
import os
import sys

PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir))

# Standard Directories
DATA_DIR = os.path.join(PROJECT_ROOT, "data")
RAW_PDF_DIR = os.path.join(DATA_DIR, "raw_pdf")
REFERENCE_JSON_DIR = os.path.join(DATA_DIR, "reference_json")
CONVERTED_XLSX_DIR = os.path.join(DATA_DIR, "converted_xlsx")
OUTPUT_DIR = os.path.join(PROJECT_ROOT, "output")
SCRATCH_DIR = os.path.join(PROJECT_ROOT, "scratch")
DOCS_DIR = os.path.join(PROJECT_ROOT, "docs")
LOGS_DIR = os.path.join(PROJECT_ROOT, "logs")
BACKUPS_DIR = os.path.join(PROJECT_ROOT, "backups")

def resolve_path(filename: str, subfolder: str = "") -> str:
    """
    Resolve path to a data or project file.
    Checks subfolder under data/ first, then root, then subfolder under root.
    """
    candidates = []
    if subfolder:
        candidates.append(os.path.join(DATA_DIR, subfolder, filename))
        candidates.append(os.path.join(PROJECT_ROOT, subfolder, filename))
    candidates.append(os.path.join(PROJECT_ROOT, filename))
    candidates.append(os.path.join(DATA_DIR, filename))

    for i, c in enumerate(candidates):
        if os.path.exists(c):
            if i > 0:
                print(f"[paths.py] WARNING: '{filename}' resolved to non-canonical "
                      f"location '{c}' (canonical would be '{candidates[0]}'). "
                      f"A stale duplicate may be shadowing the current file.",
                      file=sys.stderr)
            return c

    # Default to preferred organized location
    if subfolder:
        return os.path.join(DATA_DIR, subfolder, filename)
    return os.path.join(PROJECT_ROOT, filename)

# Reference JSONs
RATES_MASTER_JSON = resolve_path("rates_master_clean.json", "reference_json")
LABOUR_PRODUCTIVITY_JSON = resolve_path("labour_productivity.json", "reference_json")
SUNDRIES_REFERENCE_JSON = resolve_path("sundries_reference.json", "reference_json")

# Converted XLSX
VOL1_CONVERTED_XLSX = resolve_path("CivilDAR_2019_Vol_1_Converted.xlsx", "converted_xlsx")
VOL2_CONVERTED_XLSX = resolve_path("CivilDAR_2019_Vol_2_Converted.xlsx", "converted_xlsx")
AUDIT_TEST_CASES_XLSX = resolve_path("CPWD_DAR_2019_Audit_Test_Cases.xlsx", "converted_xlsx")
DSR_VOL1_SCHEDULE_XLSX = resolve_path("DSR_Vol1_Schedule.xlsx", "converted_xlsx")

# Raw PDFs
PDF_CIVIL_DAR_VOL1 = resolve_path("CivilDAR_2019_Vol_1.pdf", "raw_pdf")
PDF_CIVIL_DAR_VOL2 = resolve_path("CivilDAR_2019_Vol_2.pdf", "raw_pdf")
PDF_DSR_VOL1 = resolve_path("DSR_Vol1_UPDATED_DEC_2021.pdf", "raw_pdf")
PDF_DSR_VOL2 = resolve_path("DSR_Vol2 _UPDATED_DEC_2021.pdf", "raw_pdf")

# Production Output Workbooks
WB_VOL1_FILE = resolve_path("CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_1.xlsx")
WB_VOL1_LATEST_FILE = resolve_path("CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_1_Latest.xlsx")
WB_VOL1_TEMPLATE_FILE = resolve_path("CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_1.xltx")
WB_VOL2_FILE = resolve_path("CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_2.xlsx")
WB_VOL2_TEMPLATE_FILE = resolve_path("CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_2.xltx")

# Keyword Engine & Analysis Artifacts
KEYWORD_LOOKUP_JSON = resolve_path("dsr_keyword_lookup.json", "reference_json")
KEYWORD_ANALYSIS_XLSX = os.path.join(OUTPUT_DIR, "CPWD_DSR_Vol1_Keyword_Analysis.xlsx")

