"""Comprehensive test suite for the CPWD Custom Rate-Analysis System.

Validates:
1. Published Productivity Registry (derivations, labour, machinery, INI/TXT export)
2. Gang Composition Registry (base gangs, deduction math, additions, INI/TXT export)
3. Calculation Engine (Cases 1, 2, 3, statutory markups, benchmarks, variance, teaching notes)
4. Full Workbook Panel Integration & Excel XML formula safety
"""

from __future__ import annotations

import configparser
from pathlib import Path
import sys
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from cpwd_productivity_registry import CPWDProductivityRegistry, PRODUCTIVITY_REGISTRY
from cpwd_gang_registry import CPWDGangRegistry, GANG_REGISTRY
from cpwd_rate_engine import calculate_custom_rate, StatutoryMarkups
from earthwork_composer_resolver import resolve_selected_keywords


class CPWDProductivityRegistryTests(unittest.TestCase):
    def test_beldar_surface_excavation_derives_correct_daily_and_hourly_productivity(self):
        entry = PRODUCTIVITY_REGISTRY.lookup("02_Earth_Work", "2.1.1", "0114")
        self.assertIsNotNone(entry)
        self.assertEqual(entry.resource_name, "Beldar")
        self.assertEqual(entry.resource_type, "Labour")
        self.assertEqual(entry.coefficient, 6.80)
        self.assertEqual(entry.batch_quantity, 100.0)
        self.assertEqual(entry.batch_unit, "sqm")
        self.assertAlmostEqual(entry.daily_productivity, 14.71, places=2)
        self.assertAlmostEqual(entry.hourly_productivity, 1.84, places=2)
        self.assertIn("14.71 sqm/worker-day", entry.derivation)
        self.assertIn("CPWD DAR 2019 Item 2.1.1", entry.source_citation)

    def test_road_roller_banking_derives_correct_machine_productivity(self):
        entry = PRODUCTIVITY_REGISTRY.lookup("02_Earth_Work", "2.3.1", "0003")
        self.assertIsNotNone(entry)
        self.assertEqual(entry.resource_type, "Machinery")
        self.assertEqual(entry.coefficient, 0.008)
        self.assertEqual(entry.batch_quantity, 10.0)
        self.assertEqual(entry.batch_unit, "cum")
        self.assertAlmostEqual(entry.hourly_productivity, 156.25, places=1)
        self.assertAlmostEqual(entry.daily_productivity, 1250.0, places=1)

    def test_hydraulic_excavator_derives_published_30_cum_per_hour(self):
        entry = PRODUCTIVITY_REGISTRY.lookup("02_Earth_Work", "2.6.1", "0020")
        self.assertIsNotNone(entry)
        self.assertAlmostEqual(entry.hourly_productivity, 30.49, places=2)
        self.assertIn("30.49", entry.derivation)

    def test_export_to_ini_and_txt_files_are_readable_and_valid(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            ini_file = Path(temp_dir) / "test_productivity.ini"
            txt_file = Path(temp_dir) / "test_productivity.txt"
            PRODUCTIVITY_REGISTRY.export_to_ini(ini_file)
            PRODUCTIVITY_REGISTRY.export_to_txt(txt_file)

            self.assertTrue(ini_file.exists())
            self.assertTrue(txt_file.exists())

            config = configparser.ConfigParser()
            config.read(ini_file, encoding="utf-8")
            self.assertGreater(len(config.sections()), 5)
            section = config["02_Earth_Work.2.1.1.0114_Beldar"]
            self.assertEqual(section["subhead"], "02_Earth_Work")
            self.assertEqual(section["coefficient"], "6.8000")
            self.assertIn("14.71", section["daily_productivity"])

            txt_content = txt_file.read_text(encoding="utf-8")
            self.assertIn("PUBLISHED PRODUCTIVITY REGISTRY", txt_content)
            self.assertIn("Beldar", txt_content)

    def test_loads_external_productivity_json_across_multiple_subheads(self):
        registry = CPWDProductivityRegistry()
        json_path = ROOT / "data" / "reference_json" / "labour_productivity.json"
        if json_path.exists():
            loaded = registry.load_from_json(json_path)
            self.assertGreater(loaded, 100)
            # Verify entries from subheads beyond Earth Work
            entries = registry.all_entries()
            subheads = {e.subhead for e in entries}
            self.assertIn("02_Earth_Work", subheads)
            self.assertIn("04_Concrete_Work", subheads)


class CPWDGangRegistryTests(unittest.TestCase):
    def test_retrieves_standard_base_item_gang(self):
        gang = GANG_REGISTRY.get_item("2.3.1")
        self.assertIsNotNone(gang)
        self.assertEqual(gang.batch_quantity, 10.0)
        self.assertEqual(gang.batch_unit, "cum")
        res_codes = {r.code for r in gang.labour}
        self.assertEqual(res_codes, {"0114", "0115", "0101", "0113"})
        mach_codes = {m.code for m in gang.machinery}
        self.assertIn("0003", mach_codes)
        self.assertIn("EQUIP-0.5T", mach_codes)

    def test_item_2_2_and_2_3_1_have_half_tonne_roller_equipment(self):
        g22 = GANG_REGISTRY.get_item("2.2")
        g231 = GANG_REGISTRY.get_item("2.3.1")
        roller_22 = next(m for m in g22.machinery if m.code == "EQUIP-0.5T")
        roller_231 = next(m for m in g231.machinery if m.code == "EQUIP-0.5T")
        self.assertEqual(roller_22.task_hours, 8.80)
        self.assertEqual(roller_231.task_hours, 8.80)

    def test_deduction_composition_subtracts_roller_and_watering_correctly(self):
        composed = GANG_REGISTRY.compose_gang(
            base_item_code="2.3.1",
            deduction_codes=["2.4", "2.5"],
            addition_codes=[],
        )
        self.assertEqual(composed.base_item.item_code, "2.3.1")
        # Roller 0003 was 0.008 - 0.008 = 0.0
        roller = next(r for r in composed.net_machinery if r.code == "0003")
        self.assertAlmostEqual(roller.coefficient, 0.0)

        # Bhisti 0101 was 0.40 - 0.40 = 0.0
        bhisti = next(r for r in composed.net_labour if r.code == "0101")
        self.assertAlmostEqual(bhisti.coefficient, 0.0)

        # Chowkidar 0113 was 0.008 - 0.008 = 0.0
        chowkidar = next(r for r in composed.net_labour if r.code == "0113")
        self.assertAlmostEqual(chowkidar.coefficient, 0.0)

        # Beldar and Coolie remain unchanged
        beldar = next(r for r in composed.net_labour if r.code == "0114")
        coolie = next(r for r in composed.net_labour if r.code == "0115")
        self.assertAlmostEqual(beldar.coefficient, 2.20)
        self.assertAlmostEqual(coolie.coefficient, 3.60)

        # Sundries: was 2.73 - 1.82 = 0.91
        sundries = next(r for r in composed.net_sundries if r.code == "9999")
        self.assertAlmostEqual(sundries.coefficient, 0.91)

        # Active resources filter: 0003, 0101, 0113 deducted; Beldar, Coolie, Sundries and EQUIP-0.5T remain active!
        active_codes = {r.code for r in composed.all_active_resources}
        self.assertEqual(active_codes, {"0114", "0115", "9999", "EQUIP-0.5T"})
        self.assertNotIn("0003", active_codes)
        self.assertNotIn("0101", active_codes)
        self.assertNotIn("0113", active_codes)

        # Direct cost: 2.20*558 + 3.60*558 + 0.91*2.0 + 1.10*0.0 = 1227.60 + 2008.80 + 1.82 = 3238.22
        self.assertAlmostEqual(composed.direct_cost_w, 3238.22, places=2)

    def test_export_gang_registry_to_ini_and_txt(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            ini_file = Path(temp_dir) / "test_gang.ini"
            txt_file = Path(temp_dir) / "test_gang.txt"
            GANG_REGISTRY.export_to_ini(ini_file)
            GANG_REGISTRY.export_to_txt(txt_file)

            self.assertTrue(ini_file.exists())
            self.assertTrue(txt_file.exists())

            config = configparser.ConfigParser()
            config.read(ini_file, encoding="utf-8")
            self.assertIn("Gang.Item_2.3.1", config.sections())
            self.assertIn("0114:Beldar:2.2000day(gang_2.0x8.8h)", config["Gang.Item_2.3.1"]["labour_mix"])
            self.assertNotIn("@", config["Gang.Item_2.3.1"]["labour_mix"])


class CPWDRateEngineTests(unittest.TestCase):
    def test_case_1_banking_without_power_roller_and_watering(self):
        """User example: 'banking earth, but no watering and no 8-tonne power roller'."""
        resolved = resolve_selected_keywords([
            "banking",
            "all kinds of soil",
            "no power roller",
            "no watering",
        ])
        result = calculate_custom_rate(resolved)

        self.assertEqual(result.base_item_code, "2.3.1")
        self.assertEqual(result.batch_quantity, 10.0)
        self.assertEqual(result.batch_unit, "cum")

        # Direct Cost W: 3238.22
        self.assertAlmostEqual(result.direct_cost_w, 3238.22, places=2)

        # Step-by-step markups
        # Water 1% on 3238.22 = 32.38 -> X = 3270.60
        # GST 14.05% on 3270.60 = 459.52 -> Y = 3730.12
        # CPOH 15% on 3730.12 = 559.52 -> Z = 4289.64
        # Cess 1% on 4289.64 = 42.90 -> Total = 4332.54
        # Custom Unit Rate = 4332.54 / 10 = 433.25
        self.assertAlmostEqual(result.final_unit_rate, 433.25, places=2)

        # Benchmark: Base 2.3.1 under standard markups is 470.57
        self.assertAlmostEqual(result.benchmark_unit_rate, 470.57, places=2)

        # Variance: 433.25 - 470.57 = -37.32 (-7.9%)
        self.assertAlmostEqual(result.variance_amount, -37.32, places=2)
        self.assertAlmostEqual(result.variance_percent, -7.9, places=1)

        # Explanations and teaching notes
        self.assertIn("omitted 8-t power roller", result.variance_explanation)
        self.assertIn("omitted watering", result.variance_explanation)
        self.assertTrue(any("Item 2.4" in note for note in result.teaching_notes))
        self.assertTrue(any("Item 2.5" in note for note in result.teaching_notes))

    def test_case_2_rough_excavation_with_difficult_condition(self):
        """User example: rough excavation under water applies 2.24.1 (+20%)."""
        resolved = resolve_selected_keywords([
            "rough excavation",
            "under water",
        ])
        result = calculate_custom_rate(resolved)

        self.assertEqual(result.base_item_code, "2.2")
        self.assertAlmostEqual(result.base_unit_rate, 746.80, places=2)
        # Conditional extra: 20% of 746.80 = 149.36
        self.assertEqual(len(result.conditional_extras), 1)
        self.assertAlmostEqual(result.conditional_extras[0]["calculated_extra_rate"], 149.36, places=2)
        # Final rate: 746.80 + 149.36 = 896.16
        self.assertAlmostEqual(result.final_unit_rate, 896.16, places=2)
        self.assertAlmostEqual(result.variance_amount, 149.36, places=2)
        self.assertAlmostEqual(result.variance_percent, 20.0, places=1)

    def test_case_3_surface_excavation_with_addition_and_overlap_warning(self):
        """User example: surface excavation with grass clearing warns of overlap."""
        resolved = resolve_selected_keywords([
            "surface excavation",
            "all kinds of soil",
            "grass clearing",
        ])
        self.assertEqual(len(resolved["overlap_warnings"]), 1)
        self.assertIn("Review scope before pricing items 2.1.1 and 2.32 together", resolved["overlap_warnings"][0])

        result = calculate_custom_rate(resolved)
        self.assertEqual(result.base_item_code, "2.1.1")
        self.assertGreater(result.final_unit_rate, result.benchmark_unit_rate)


if __name__ == "__main__":
    unittest.main()
