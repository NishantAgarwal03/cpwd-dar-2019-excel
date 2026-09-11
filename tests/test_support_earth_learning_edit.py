"""Tests for the source-backed earthwork productivity derivation model."""

from __future__ import annotations

import sys
import tempfile
import unittest
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

from openpyxl import Workbook

ROOT = Path(__file__).resolve().parents[1]
BASELINE_WORKBOOK = ROOT / "CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_1_Latest.xlsx"
sys.path.insert(0, str(ROOT / "scripts"))

from support_earth_learning_edit import (  # noqa: E402
    build_derivation_catalog,
    derive_resource_norm,
    export_ascii_formula_workbook,
)


def sheet8_formulas(workbook_path: Path) -> list[str]:
    """Read formula nodes directly so Excel's XML compatibility is tested."""
    with zipfile.ZipFile(workbook_path) as archive:
        root = ET.fromstring(archive.read("xl/worksheets/sheet8.xml"))
    return [node.text or "" for node in root.findall(".//{*}f")]


def source_sheet_with_verified_norms():
    """Build the same source-table shape used in 02_Earth_Work (rows 109+)."""
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "02_Earth_Work"

    # A:D identify the source work item; E/F identify its resource; G is output.
    # H/I are the source activity and verbatim source calculation; K is the
    # published coefficient and L is the support-resource lookup key.
    sheet.append([None] * 17)
    sheet.append([
        "Surface Excavation (≤30 cm)", "All kinds of soil",
        "Manual labor (depth ≤1.5m)", "General surface cut", 1, "0114",
        14.70588235294118,
        "Manual pickaxe and phowrah loosening.",
        "Gang: 4 Beldars × 13.60 h = 54.40 man-hrs ÷ 8.00 h = 6.800 Beldar-days per 100 sqm.",
        "source-reference", 6.8,
        "2. Surface Preparation & Earthwork|General surface cut|Manual labor (depth ≤1.5m)|1",
        "Beldar", "day", 558, "W",
        "Surface Excavation (≤30 cm)|All kinds of soil|Manual labor (depth ≤1.5m)|General surface cut|1",
    ])
    sheet.append([
        "Embankment Banking & Dressing", "All kinds of soil",
        "Manual labor (depth ≤1.5m)", "Full cycle: rough excavation + banking + roll", 5, "0003",
        1250,
        "8-10 tonne diesel roller conducting compaction passes.",
        "Plant: 1 Roller × 0.064 h ÷ 8.00 h = 0.008 roller-day per 10 cum.",
        "source-reference", 0.008,
        "2. Surface Preparation & Earthwork|Full cycle: rough excavation + banking + roll|Manual labor (depth ≤1.5m)|5",
        "Hire – Diesel Road Roller 8-10t", "day", 3000, "W",
        "Embankment Banking & Dressing|All kinds of soil|Manual labor (depth ≤1.5m)|Full cycle: rough excavation + banking + roll|5",
    ])
    return sheet


class DerivationCatalogTests(unittest.TestCase):
    def setUp(self):
        sheet = source_sheet_with_verified_norms()
        # Excel commonly stores CPWD resource code 0114 as the number 114.
        sheet["F2"] = 114
        self.catalog = build_derivation_catalog(sheet)

    def test_derives_labour_norm_from_unique_source_record(self):
        result = derive_resource_norm(
            {"code": "0114", "description": "Beldar", "coefficient": 6.8, "unit": "day"},
            {"source_key": "2. Surface Preparation & Earthwork|General surface cut|Manual labor (depth ≤1.5m)|1", "batch_quantity": 100, "batch_unit": "sqm"},
            self.catalog,
        )

        self.assertEqual(result["role"], "Beldar")
        self.assertEqual(result["gang_or_machine"], 4)
        self.assertAlmostEqual(result["task_hours"], 13.6)
        self.assertEqual(result["shift_hours"], 8)
        self.assertAlmostEqual(result["coefficient"], 6.8)
        self.assertEqual(result["unit"], "day")
        self.assertEqual(result["batch_quantity"], 100)
        self.assertEqual(result["batch_unit"], "sqm")
        self.assertEqual(result["interpretation_label"], "CPWD fixed norm / source evidence")
        self.assertIn("Gang: 4 Beldars", result["source_norm"])

    def test_derives_machine_norm_from_unique_source_record(self):
        result = derive_resource_norm(
            {"code": "0003", "description": "Hire – Diesel Road Roller 8-10t", "coefficient": 0.008, "unit": "day"},
            {"source_key": "2. Surface Preparation & Earthwork|Full cycle: rough excavation + banking + roll|Manual labor (depth ≤1.5m)|5", "batch_quantity": 10, "batch_unit": "cum"},
            self.catalog,
        )

        self.assertEqual(result["role"], "Hire – Diesel Road Roller 8-10t")
        self.assertEqual(result["gang_or_machine"], 1)
        self.assertAlmostEqual(result["task_hours"], 0.064)
        self.assertAlmostEqual(result["machine_hours"], 0.064)
        self.assertAlmostEqual(result["coefficient"], 0.008)
        self.assertEqual(result["interpretation_label"], "CPWD fixed norm / source evidence")

    def test_returns_none_when_source_mapping_is_not_unique_or_not_found(self):
        result = derive_resource_norm(
            {"code": "9999", "description": "Unmatched resource", "coefficient": 1, "unit": "day"},
            {"source_key": "missing", "batch_quantity": 1, "batch_unit": "cum"},
            self.catalog,
        )

        self.assertIsNone(result)

    def test_keeps_final_coefficient_without_inventing_task_hours(self):
        sheet = source_sheet_with_verified_norms()
        sheet.append([
            "Other", "Soil", "Manual", "Final coefficient only", 1, "0114", None,
            "Published final norm only.", "Final coefficient: 1.25 Beldar-day per 10 cum.", "source-reference",
            1.25, "final-only", "Beldar", "day", 558, "W", "Other|Soil|Manual|Final coefficient only|1",
        ])
        result = derive_resource_norm(
            {"code": "0114", "description": "Beldar", "coefficient": 1.25, "unit": "day"},
            {"source_key": "final-only", "batch_quantity": 10, "batch_unit": "cum"},
            build_derivation_catalog(sheet),
        )

        self.assertIsNone(result["task_hours"])
        self.assertEqual(result["interpretation_label"], "Teaching interpretation, not a published CPWD rule")

    def test_matches_excel_numeric_resource_code_to_zero_padded_support_code(self):
        result = derive_resource_norm(
            {"code": "0114", "description": "Beldar", "coefficient": 6.8, "unit": "day"},
            {"source_key": "2. Surface Preparation & Earthwork|General surface cut|Manual labor (depth ≤1.5m)|1", "batch_quantity": 100, "batch_unit": "sqm"},
            self.catalog,
        )

        self.assertIsNotNone(result)


class FormulaCompatibilityTests(unittest.TestCase):
    def test_exported_sheet8_formula_nodes_are_ascii_only(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            output = Path(temporary_directory) / "sanitized.xlsx"
            export_ascii_formula_workbook(BASELINE_WORKBOOK, output)
            formulas = sheet8_formulas(output)

        self.assertEqual(len(formulas), 473)
        invalid = [formula for formula in formulas if any(ord(char) > 127 for char in formula)]
        self.assertEqual(invalid, [])

    def test_targeted_export_rewrites_only_formula_literals_in_sheet8(self):
        source_formulas = sheet8_formulas(BASELINE_WORKBOOK)
        with tempfile.TemporaryDirectory() as temporary_directory:
            output = Path(temporary_directory) / "sanitized.xlsx"
            export_ascii_formula_workbook(BASELINE_WORKBOOK, output)
            output_formulas = sheet8_formulas(output)

        self.assertEqual(len(output_formulas), 473)
        self.assertTrue(all(all(ord(char) <= 127 for char in formula) for formula in output_formulas))
        self.assertEqual(
            output_formulas,
            [formula.replace("—", "-").replace("→", "->").replace("›", ">") for formula in source_formulas],
        )


if __name__ == "__main__":
    unittest.main()
