# -*- coding: utf-8 -*-
"""
tests/test_keyword_engine.py
============================
Unit test suite for scripts.keywords engine, diagnostic classifier, and reports.
Run with:
  py -3.11 -m unittest tests/test_keyword_engine.py
"""
import unittest
import os
import tempfile
import json
from pathlib import Path

from scripts.keywords.engine import (
    ScheduleItem,
    extract_keywords,
    keyword_priority,
    combo_score,
    build_inverted_index,
    find_minimum_identifier,
)
from scripts.keywords.diagnose import classify_and_resolve
from scripts.keywords.report import write_json, write_excel


class TestKeywordExtraction(unittest.TestCase):
    def test_phrase_slugs_and_ratios(self):
        desc = "Providing and laying in position cement concrete 1:2:4 (1 cement : 2 coarse sand : 4 graded stone aggregate 20mm)"
        tokens = extract_keywords(desc)
        
        # Verify mix ratio is preserved
        self.assertIn("1:2:4", tokens)
        self.assertIn("mix_1-2-4", tokens)
        self.assertIn("coarse-sand", tokens)
        
        # Verify stop words are excluded
        self.assertNotIn("work", tokens)
        self.assertNotIn("and", tokens)
        self.assertNotIn("in", tokens)

    def test_timbering_and_depth_slugs(self):
        desc = "Close timbering in trenches including strutting, shoring and packing depth not exceeding 1.5 m"
        tokens = extract_keywords(desc)
        
        self.assertIn("close-timbering-in-trenches", tokens)
        self.assertIn("depth_upto_1-5", tokens)
        self.assertIn("strutting", tokens)
        self.assertIn("shoring", tokens)

    def test_priority_hierarchy(self):
        # Slug / P1 terms should have priority 1
        self.assertEqual(keyword_priority("anti-termite"), 1)
        self.assertEqual(keyword_priority("depth_upto_1-5"), 1)
        self.assertEqual(keyword_priority("chlorpyriphos"), 1)

        # Qualifiers should have priority 2
        self.assertEqual(keyword_priority("foundation"), 2)
        self.assertEqual(keyword_priority("external"), 2)

        # Ratios should have priority 3
        self.assertEqual(keyword_priority("1:2"), 3)
        self.assertEqual(keyword_priority("1:1.5:3"), 3)

        # Action verbs should have priority 4
        self.assertEqual(keyword_priority("excavation"), 4)
        self.assertEqual(keyword_priority("compacting"), 4)

        # Numbers should have priority 5
        self.assertEqual(keyword_priority("20"), 5)

        # Stop words should have priority 6
        self.assertEqual(keyword_priority("work"), 6)
        self.assertEqual(keyword_priority("the"), 6)


class TestInvertedIndexAndSearch(unittest.TestCase):
    def setUp(self):
        self.item1 = ScheduleItem(
            sheet="Earth Work",
            code="2.1.1",
            parent_code="2.1",
            parent_desc="Earth work in surface excavation",
            child_desc="in all kinds of soil",
            complete_desc="Earth work in surface excavation in all kinds of soil",
            unit="sqm",
            rate=120.0,
            keywords=["surface", "excavation", "soil"],
        )
        self.item2 = ScheduleItem(
            sheet="Earth Work",
            code="2.1.2",
            parent_code="2.1",
            parent_desc="Earth work in surface excavation",
            child_desc="in hard rock requiring blasting",
            complete_desc="Earth work in surface excavation in hard rock requiring blasting",
            unit="sqm",
            rate=450.0,
            keywords=["surface", "excavation", "hard-rock-requiring", "blasting"],
        )
        self.item3 = ScheduleItem(
            sheet="Mortars",
            code="3.1.1",
            parent_code="3.1",
            parent_desc="Cement mortar 1:2",
            child_desc="1 cement : 2 fine sand",
            complete_desc="Cement mortar 1:2 1 cement : 2 fine sand",
            unit="cum",
            rate=3800.0,
            keywords=["1:2", "fine-sand", "mortar"],
        )

    def test_uniqueness_resolution(self):
        items = [self.item1, self.item2, self.item3]
        index = build_inverted_index(items)

        combo1, status1 = find_minimum_identifier(self.item1, index, max_words=3)
        self.assertEqual(status1, "UNIQUE")
        self.assertIn("soil", combo1)

        combo2, status2 = find_minimum_identifier(self.item2, index, max_words=3)
        self.assertEqual(status2, "UNIQUE")
        self.assertTrue("hard-rock-requiring" in combo2 or "blasting" in combo2)

        combo3, status3 = find_minimum_identifier(self.item3, index, max_words=3)
        self.assertEqual(status3, "UNIQUE")
        self.assertTrue("fine-sand" in combo3 or "1:2" in combo3 or "mortar" in combo3)

    def test_structurally_impossible_classification(self):
        # item_base has keywords [excavation, trench]
        # item_superset has keywords [excavation, trench, deep]
        # item_base keywords are a strict subset of item_superset keywords,
        # so any combo matching item_base ALSO matches item_superset.
        item_base = ScheduleItem(
            sheet="Earth Work",
            code="2.9.1",
            parent_code="2.9",
            parent_desc="Trench work",
            child_desc="excavation in trench",
            complete_desc="Trench work excavation in trench",
            unit="cum",
            rate=200.0,
            keywords=["excavation", "trench"],
        )
        item_superset = ScheduleItem(
            sheet="Earth Work",
            code="2.9.2",
            parent_code="2.9",
            parent_desc="Trench work",
            child_desc="excavation in deep trench",
            complete_desc="Trench work excavation in deep trench",
            unit="cum",
            rate=300.0,
            keywords=["excavation", "trench", "deep"],
        )
        items = [item_base, item_superset]
        index = build_inverted_index(items)

        combo_b, status_b = find_minimum_identifier(item_base, index, max_words=2)
        item_base.identifier = combo_b
        item_base.identifier_status = status_b
        self.assertEqual(status_b, "NOT_UNIQUELY_IDENTIFIABLE")

        classify_and_resolve(items, index, max_words_deep=4)
        self.assertEqual(item_base.identifier_status, "STRUCTURALLY_IMPOSSIBLE")
        self.assertIn("2.9.2", item_base.ambiguous_with)


class TestExportIntegrity(unittest.TestCase):
    def test_json_and_excel_export(self):
        item = ScheduleItem(
            sheet="Concrete Work",
            code="4.1.1",
            parent_code="4.1",
            parent_desc="Providing cement concrete",
            child_desc="1:2:4",
            complete_desc="Providing cement concrete 1:2:4",
            unit="cum",
            rate=5200.0,
            keywords=["1:2:4", "mix_1-2-4", "concrete"],
            identifier=("mix_1-2-4",),
            identifier_status="UNIQUE",
            chapter_identifier=("1:2:4",),
            chapter_identifier_status="UNIQUE",
        )
        items = [item]

        with tempfile.TemporaryDirectory() as tmpdir:
            json_path = os.path.join(tmpdir, "test_lookup.json")
            xlsx_path = os.path.join(tmpdir, "test_analysis.xlsx")

            write_json(items, json_path)
            self.assertTrue(os.path.exists(json_path))

            with open(json_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                self.assertEqual(data["metadata"]["total_items"], 1)
                self.assertIn("4.1.1", data["quick_lookup"])

            write_excel(items, xlsx_path)
            self.assertTrue(os.path.exists(xlsx_path))
            self.assertGreater(os.path.getsize(xlsx_path), 1000)


if __name__ == "__main__":
    unittest.main()
