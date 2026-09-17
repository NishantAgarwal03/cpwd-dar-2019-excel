# -*- coding: utf-8 -*-
"""
scripts.keywords.diagnose
=========================
Diagnostic module for analyzing ambiguous and non-unique DSR items.

Classifies failure cases into:
  1. STRUCTURALLY_IMPOSSIBLE:
     The item's keyword set is a strict subset of another item's keywords.
     No purely inclusion-based keyword combination can ever isolate it.
     Computes the best partial identifier (minimal combo yielding the fewest peer matches).
  2. NEEDS_MORE_KEYWORDS:
     A unique combination likely exists, but exceeds the primary max_words threshold.
     Executes a deep-pass search (e.g. max_words_deep=8) to resolve it.
"""
from __future__ import annotations

import itertools
from typing import List, Dict, Optional, Tuple

from scripts.keywords.engine import (
    ScheduleItem,
    InvertedIndex,
    find_minimum_identifier,
)


def classify_and_resolve(
    items: List[ScheduleItem],
    index: InvertedIndex,
    max_words_deep: int = 8,
) -> None:
    """
    Mutates items in place:
      - Re-classifies failures
      - Attempts deeper search for NEEDS_MORE_KEYWORDS
      - Computes best_partial + ambiguous_with for STRUCTURALLY_IMPOSSIBLE
    """
    all_kw_sets = {it.code: set(it.keywords) for it in items}
    item_map = {it.code: it for it in items}

    failures = [
        it for it in items
        if it.identifier_status in ("NOT_UNIQUELY_IDENTIFIABLE", "NEEDS_MORE_KEYWORDS")
    ]
    if not failures:
        print("  No failures to classify.")
        return

    print(f"  Classifying {len(failures)} non-unique items ...")
    structurally_impossible = 0
    needs_more = 0
    resolved = 0

    for item in failures:
        my_kw = all_kw_sets[item.code]
        # Structurally impossible <-> exists another item whose keyword set is a superset
        dominated_by = [
            c for c, kws in all_kw_sets.items()
            if c != item.code and my_kw.issubset(kws)
        ]

        if dominated_by:
            item.identifier_status = "STRUCTURALLY_IMPOSSIBLE"
            item.ambiguous_with = dominated_by
            _attach_best_partial(item, index, item_map)
            structurally_impossible += 1
        else:
            item.identifier_status = "NEEDS_MORE_KEYWORDS"
            needs_more += 1
            combo, status = find_minimum_identifier(item, index, max_words=max_words_deep)
            if status == "UNIQUE":
                item.identifier = combo
                item.identifier_status = "UNIQUE"
                resolved += 1

    print(f"    STRUCTURALLY_IMPOSSIBLE  : {structurally_impossible}")
    print(f"    NEEDS_MORE_KEYWORDS      : {needs_more - resolved}")
    print(f"    Resolved by deeper search : {resolved}")


def _attach_best_partial(
    target: ScheduleItem,
    index: InvertedIndex,
    item_map: Dict[str, ScheduleItem],
) -> None:
    """
    Find the keyword combo that narrows matches to the fewest items
    (even if > 1). Store as target.best_partial and target.partial_matches.
    """
    cand = [(kw, index[kw]) for kw in target.keywords if kw in index]
    if not cand:
        target.best_partial = None
        target.partial_matches = []
        return

    best_combo: Optional[Tuple[str, ...]] = None
    best_matches: List[str] = list(item_map.keys())

    for n in range(1, min(4, len(cand) + 1)):
        for combo in itertools.combinations(cand, n):
            keys, sets = zip(*combo)
            matching = sets[0]
            for s in sets[1:]:
                matching = matching & s
                if len(matching) <= len(best_matches):
                    break
            if 1 < len(matching) <= len(best_matches):
                best_matches = list(matching)
                best_combo = keys
                if len(best_matches) == 2:
                    break
        if best_combo and len(best_matches) == 2:
            break

    target.best_partial = best_combo
    target.partial_matches = [c for c in best_matches if c != target.code]
