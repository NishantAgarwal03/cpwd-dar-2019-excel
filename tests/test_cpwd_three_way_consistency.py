"""Tri-Party Mathematical Consistency and Non-Contradiction Test Suite.

Ensures strict mathematical alignment and zero contradiction across:
1. CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_1_Latest.xlsx (Rates_Master, 02_Earth_Work, Trade Sheets)
2. scripts/cpwd_gang_registry.py (GANG_REGISTRY)
3. scripts/cpwd_productivity_registry.py (PRODUCTIVITY_REGISTRY)
"""

from __future__ import annotations

import re
import sys
import unittest
from pathlib import Path
import openpyxl

ROOT = Path(__file__).resolve().parents[1]
WORKBOOK_PATH = ROOT / "CPWD_DAR_2019_Custom_Rate_Analysis_Workbook_Vol_1_Latest.xlsx"
sys.path.insert(0, str(ROOT / "scripts"))

from cpwd_gang_registry import GANG_REGISTRY
from cpwd_productivity_registry import PRODUCTIVITY_REGISTRY


class TestCPWDThreeWayConsistency(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.wb = openpyxl.load_workbook(WORKBOOK_PATH, data_only=True)
        cls.ws_rm = cls.wb["Rates_Master"]
        cls.ws_ew = cls.wb["02_Earth_Work"]

        # Parse master rates from Rates_Master sheet
        cls.master_rates: dict[str, float] = {}
        for r in range(2, cls.ws_rm.max_row + 1):
            code = str(cls.ws_rm.cell(r, 1).value or "").strip()
            rate = cls.ws_rm.cell(r, 5).value
            if code and rate is not None:
                try:
                    cls.master_rates[code] = float(rate)
                    cls.master_rates[str(int(code))] = float(rate)
                except ValueError:
                    pass

        # Parse source derivation norms from 02_Earth_Work (rows 109+)
        cls.wb_earth_norms: dict[str, list[dict]] = {}
        for r in range(109, cls.ws_ew.max_row + 1):
            c_code = str(cls.ws_ew.cell(r, 6).value or "").strip()
            if not c_code:
                continue
            c_coef = cls.ws_ew.cell(r, 11).value
            c_prod = cls.ws_ew.cell(r, 7).value
            c_rate = cls.ws_ew.cell(r, 15).value
            c_calc = cls.ws_ew.cell(r, 9).value
            cls.wb_earth_norms.setdefault(c_code, []).append({
                "row": r,
                "coef": float(c_coef) if c_coef is not None else None,
                "prod": float(c_prod) if c_prod is not None else None,
                "rate": float(c_rate) if c_rate is not None else None,
                "calc": str(c_calc) if c_calc is not None else "",
            })

    @classmethod
    def tearDownClass(cls):
        cls.wb.close()

    def test_gang_registry_matches_productivity_registry_coefficients(self):
        """Verify that every resource in Gang Registry matches Productivity Registry in coefficient magnitude."""
        for item_code, gang in GANG_REGISTRY._items.items():
            all_res = [*gang.machinery, *gang.labour, *gang.materials, *gang.sundries]
            for res in all_res:
                p_entry = PRODUCTIVITY_REGISTRY.lookup("02_Earth_Work", item_code, res.code)
                if p_entry:
                    self.assertAlmostEqual(
                        abs(res.coefficient),
                        abs(p_entry.coefficient),
                        places=4,
                        msg=f"Coefficient mismatch for {item_code} {res.code} ({res.name})",
                    )

    def test_gang_mathematical_derivation_equation(self):
        """Verify the invariant: coefficient == (gang_count * task_hours) / 8.0."""
        for item_code, gang in GANG_REGISTRY._items.items():
            for res in gang.labour:
                if res.gang_count is not None and res.task_hours is not None:
                    derived_coef = (res.gang_count * res.task_hours) / 8.0
                    self.assertAlmostEqual(
                        derived_coef,
                        abs(res.coefficient),
                        places=3,
                        msg=f"Gang math violated for {item_code} {res.code}: {res.gang_count} * {res.task_hours} / 8.0 != {res.coefficient}",
                    )

    def test_productivity_mathematical_derivation_equations(self):
        """Verify invariants:
        1. For Labour: daily_productivity == batch_quantity / abs(coefficient)
        2. For Machinery: daily_productivity == round(hourly_productivity * 8.0, 2)
        3. hourly_productivity == batch_quantity / (abs(coefficient) * 8.0)
        """
        for entry in PRODUCTIVITY_REGISTRY.all_entries():
            if entry.daily_productivity is not None and entry.coefficient != 0:
                if entry.resource_type == "Labour":
                    expected_daily = entry.batch_quantity / abs(entry.coefficient)
                    self.assertAlmostEqual(
                        expected_daily,
                        abs(entry.daily_productivity),
                        places=1,
                        msg=f"Labour daily productivity math violated for {entry.item_code} {entry.resource_code}",
                    )
                elif entry.resource_type == "Machinery":
                    # In Machinery, registry computes hourly first (rounded to 2 places) then daily = hourly * 8.0
                    self.assertAlmostEqual(
                        abs(entry.daily_productivity),
                        round(abs(entry.hourly_productivity) * 8.0, 2),
                        places=2,
                        msg=f"Machinery daily productivity math violated for {entry.item_code} {entry.resource_code}",
                    )

    def test_gang_registry_rates_match_workbook_rates_master(self):
        """Verify that basic rates in Gang Registry match Rates_Master in the master workbook."""
        for item_code, gang in GANG_REGISTRY._items.items():
            all_res = [*gang.machinery, *gang.labour, *gang.materials]
            for res in all_res:
                if res.rate > 0 and res.code in self.master_rates:
                    master_rate = self.master_rates[res.code]
                    self.assertAlmostEqual(
                        res.rate,
                        master_rate,
                        places=2,
                        msg=f"Rate mismatch for {item_code} {res.code}: gang rate {res.rate} != Rates_Master {master_rate}",
                    )

    def test_workbook_earthwork_norms_align_with_registries(self):
        """Verify that 02_Earth_Work source table in the workbook aligns with the Python registries."""
        # For item 2.1.1 Beldar (0114): coef = 6.8, prod = 100/6.8 = 14.70588, rate = 558
        beldar_norms = self.wb_earth_norms.get("0114", [])
        self.assertTrue(len(beldar_norms) > 0, "Missing Beldar 0114 in 02_Earth_Work norms")
        beldar_211 = next((n for n in beldar_norms if abs(n["coef"] - 6.8) < 1e-3), None)
        self.assertIsNotNone(beldar_211, "02_Earth_Work missing 6.80 Beldar norm")
        self.assertAlmostEqual(beldar_211["prod"], 100.0 / 6.8, places=3)
        self.assertEqual(beldar_211["rate"], 558.0)

        # For item 2.2 / 2.3.1 Road Roller (0003): coef = 0.008, prod = 1250, rate = 3000
        roller_norms = self.wb_earth_norms.get("0003", [])
        self.assertTrue(len(roller_norms) > 0, "Missing Roller 0003 in 02_Earth_Work norms")
        roller_norm = next((n for n in roller_norms if abs(n["coef"] - 0.008) < 1e-4), None)
        self.assertIsNotNone(roller_norm, "02_Earth_Work missing 0.008 Roller norm")
        self.assertEqual(roller_norm["prod"], 1250.0)
        self.assertEqual(roller_norm["rate"], 3000.0)

    def test_trade_sheets_batch_quantity_and_productivity_consistency(self):
        """Verify that newly built trade sheets have consistent quantity and productivity columns."""
        trade_sheets = [
            "03_Mortars", "04_Concrete_Work", "06_Masonry_Work",
            "07_Stone_Work", "08_Cladding_Work", "09_Wood_and_PVC_Work",
            "10_Steel_Work", "11_Flooring", "12_Roofing"
        ]
        for sheet_name in trade_sheets:
            ws = self.wb[sheet_name]
            for r in range(7, ws.max_row + 1):
                cat = ws.cell(r, 5).value
                prod = ws.cell(r, 6).value
                qty = ws.cell(r, 7).value
                if cat in ["Labour", "Machinery", "Material"]:
                    self.assertIsNotNone(prod, f"Missing productivity at {sheet_name}:Row {r}")
                    self.assertIsNotNone(qty, f"Missing quantity at {sheet_name}:Row {r}")
                    # Ensure qty has recognized batch unit
                    self.assertTrue(
                        any(u in str(qty).lower() for u in ["cum", "sqm", "quintal", "kg", "m", "nos", "tonne", "each", "bag"]),
                        f"Unrecognized batch quantity format at {sheet_name}:Row {r}: '{qty}'"
                    )

    def test_all_trades_registries_cover_subheads_02_to_12(self):
        """Verify that all-trades exports exist and cover all Sub-heads 02 to 12."""
        import configparser
        g_ini = ROOT / "data" / "gang_registry_all_trades.ini"
        p_ini = ROOT / "data" / "productivity_registry_all_trades.ini"
        g_txt = ROOT / "data" / "gang_registry_all_trades.txt"
        p_txt = ROOT / "data" / "productivity_registry_all_trades.txt"

        self.assertTrue(g_ini.exists(), "Missing gang_registry_all_trades.ini")
        self.assertTrue(p_ini.exists(), "Missing productivity_registry_all_trades.ini")
        self.assertTrue(g_txt.exists(), "Missing gang_registry_all_trades.txt")
        self.assertTrue(p_txt.exists(), "Missing productivity_registry_all_trades.txt")

        cp_g = configparser.ConfigParser()
        cp_g.read(g_ini, encoding="utf-8")
        self.assertGreater(len(cp_g.sections()), 500)

        cp_p = configparser.ConfigParser()
        cp_p.read(p_ini, encoding="utf-8")
        self.assertGreater(len(cp_p.sections()), 1500)

        # Check subheads present in productivity registry sections
        expected_prefixes = {
            "02_Earth_Work", "03_Mortars", "04_Concrete_Work", "05_RCC_Work",
            "06_Masonry_Work", "07_Stone_Work", "08_Cladding_Work",
            "09_Wood_and_PVC_Work", "10_Steel_Work", "11_Flooring", "12_Roofing"
        }
        found_subheads = {s.split(".")[0] for s in cp_p.sections() if "." in s}
        for exp in expected_prefixes:
            self.assertIn(exp, found_subheads, f"Sub-head {exp} missing from all-trades productivity registry")


if __name__ == "__main__":
    unittest.main()

