"""Regression tests for src/discovery.py against test-fixtures/.

These were proven by hand during initial design (2026-07-19) before this
project existed; codified here as the starting backpressure for the loop.
"""

import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from discovery import discover_from_pptx  # noqa: E402

FIXTURES = os.path.join(os.path.dirname(__file__), "..", "test-fixtures")


def test_shp_groupshape_finds_all_four_leaf_shapes():
    path = os.path.join(FIXTURES, "shp-groupshape.pptx")
    candidates = discover_from_pptx(path)
    assert len(candidates) == 4


def test_shp_groupshape_recurses_into_the_group_not_opaque():
    path = os.path.join(FIXTURES, "shp-groupshape.pptx")
    candidates = discover_from_pptx(path)
    grouped = [c for c in candidates if c.group_path]
    top_level = [c for c in candidates if not c.group_path]
    assert len(grouped) == 3
    assert len(top_level) == 1
    assert grouped[0].group_path == ("Group 4",)


def test_shp_groupshape_finds_zero_candidate_fields():
    # All four shapes are empty decoration (no text, not pictures) -- correct
    # behavior is to find zero fields, not force a match onto empty shapes.
    path = os.path.join(FIXTURES, "shp-groupshape.pptx")
    candidates = discover_from_pptx(path)
    assert sum(1 for c in candidates if c.is_candidate_field) == 0
