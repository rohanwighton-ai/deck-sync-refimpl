"""Regression tests for src/discovery.py against test-fixtures/.

These were proven by hand during initial design (2026-07-19) before this
project existed; codified here as the starting backpressure for the loop.
"""

import os
import sys
import xml.etree.ElementTree as ET
from zipfile import ZipFile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from discovery import discover, discover_from_pptx  # noqa: E402

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


def test_shp_groupshape_shapes_have_no_placeholder():
    path = os.path.join(FIXTURES, "shp-groupshape.pptx")
    candidates = discover_from_pptx(path)
    assert all(not c.has_placeholder for c in candidates)
    assert all(c.placeholder_type is None and c.placeholder_idx is None for c in candidates)


def test_mst_slide_layouts_captures_placeholder_type_and_idx():
    # mst-slide-layouts.pptx has no ppt/slides/* entries at all -- only
    # slideLayouts/slideMasters (see IMPLEMENTATION_PLAN.md). discover() is
    # root-agnostic (just looks for p:spTree), so it works unchanged on a
    # slideLayout XML root even before a dedicated loader exists.
    path = os.path.join(FIXTURES, "mst-slide-layouts.pptx")
    with ZipFile(path) as z:
        with z.open("ppt/slideLayouts/slideLayout1.xml") as f:
            root = ET.parse(f).getroot()
    candidates = discover(root)

    title = next(c for c in candidates if c.name == "Title 1")
    assert title.placeholder_type == "title"
    assert title.placeholder_idx == 0

    body = next(c for c in candidates if c.name == "Text Placeholder 3")
    assert body.placeholder_type == "body"
    assert body.placeholder_idx == 10
