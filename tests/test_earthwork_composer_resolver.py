"""Behaviour tests for controlled-keyword earthwork composition."""

from __future__ import annotations

import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from earthwork_composer_resolver import resolve_selected_keywords  # noqa: E402


class EarthworkComposerResolverTests(unittest.TestCase):
    def test_banking_without_power_roller_or_watering_selects_base_and_deductions(self):
        result = resolve_selected_keywords([
            "banking excavated earth",
            "all kinds of soil",
            "no power roller",
            "no watering",
        ])

        self.assertEqual(result["base"]["item_code"], "2.3.1")
        self.assertEqual(
            [component["item_code"] for component in result["deductions"]],
            ["2.4", "2.5"],
        )
        self.assertEqual(result["additions"], [])
        self.assertEqual(result["conditional_extras"], [])
        self.assertEqual(result["overlap_warnings"], [])

    def test_rough_excavation_banking_with_water_adds_only_a_qualified_percentage_extra(self):
        result = resolve_selected_keywords([
            "rough excavation and banking",
            "water or liquid mud",
        ])

        self.assertEqual(result["base"]["item_code"], "2.2")
        self.assertEqual(result["deductions"], [])
        self.assertEqual(result["conditional_extras"][0]["item_code"], "2.24.1")
        self.assertEqual(result["conditional_extras"][0]["percent"], 20)
        self.assertEqual(result["conditional_extras"][0]["qualifying_base_item"], "2.2")

    def test_rough_excavation_banking_with_foul_position_resolves_25_percent_extra(self):
        result = resolve_selected_keywords([
            "rough excavation and banking",
            "foul position",
        ])

        self.assertEqual(result["base"]["item_code"], "2.2")
        self.assertEqual(result["conditional_extras"][0]["item_code"], "2.24.2")
        self.assertEqual(result["conditional_extras"][0]["percent"], 25)
        self.assertEqual(result["conditional_extras"][0]["qualifying_base_item"], "2.2")

    def test_surface_excavation_with_grass_clearing_warns_of_possible_scope_overlap(self):
        result = resolve_selected_keywords([
            "surface excavation",
            "all kinds of soil",
            "clearing grass",
        ])

        self.assertEqual(result["base"]["item_code"], "2.1.1")
        self.assertEqual(result["additions"][0]["item_code"], "2.32")
        self.assertEqual(len(result["overlap_warnings"]), 1)
        self.assertIn("2.1.1", result["overlap_warnings"][0])
        self.assertIn("2.32", result["overlap_warnings"][0])

    def test_selected_components_keep_the_user_keyword_order_within_each_category(self):
        result = resolve_selected_keywords([
            "rough excavation and banking",
            "foul position",
            "water or liquid mud",
        ])

        self.assertEqual(
            [component["item_code"] for component in result["conditional_extras"]],
            ["2.24.2", "2.24.1"],
        )

    def test_unknown_controlled_keyword_is_rejected_instead_of_being_silently_ignored(self):
        with self.assertRaisesRegex(ValueError, "Unknown controlled keyword.*borrow pit"):
            resolve_selected_keywords(["rough excavation and banking", "borrow pit"])

    def test_result_exposes_plain_language_composed_scope(self):
        result = resolve_selected_keywords(["banking excavated earth", "all kinds of soil", "no watering"])

        self.assertIn("Base 2.3.1", result["composed_scope"])
        self.assertIn("Deduct 2.5", result["composed_scope"])
        self.assertEqual(result["incompatibilities"], [])

    def test_no_watering_rejects_non_banking_base_before_any_deduction_is_executable(self):
        with self.assertRaisesRegex(ValueError, "2.5.*2.1.1"):
            resolve_selected_keywords(["surface excavation", "all kinds of soil", "no watering"])

    def test_duplicate_difficult_condition_selection_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "Duplicate controlled keyword.*foul position"):
            resolve_selected_keywords([
                "rough excavation and banking",
                "foul position",
                "foul position",
            ])

    def test_water_and_foul_conditions_require_separate_qualifying_quantities(self):
        result = resolve_selected_keywords([
            "rough excavation and banking",
            "water or liquid mud",
            "foul position",
        ])

        self.assertEqual(
            [component["item_code"] for component in result["conditional_extras"]],
            ["2.24.1", "2.24.2"],
        )
        self.assertEqual(len(result["incompatibilities"]), 1)
        self.assertIn("separate qualifying quantities", result["incompatibilities"][0])

    def test_approved_literal_keyword_aliases_resolve_without_renaming_by_the_student(self):
        banking = resolve_selected_keywords(["banking", "all kinds of soil", "no power roller", "no watering"])
        rough_water = resolve_selected_keywords(["rough excavation", "under water"])
        surface_grass = resolve_selected_keywords(["surface excavation", "all kinds of soil", "grass clearing"])

        self.assertEqual(banking["base"]["item_code"], "2.3.1")
        self.assertEqual([item["item_code"] for item in banking["deductions"]], ["2.4", "2.5"])
        self.assertEqual(rough_water["base"]["item_code"], "2.2")
        self.assertEqual(rough_water["conditional_extras"][0]["item_code"], "2.24.1")
        self.assertEqual(surface_grass["additions"][0]["item_code"], "2.32")


if __name__ == "__main__":
    unittest.main()
