"""Tests for src/identity_tags.py against specs/identity-tags.md.

Real disk round-trips throughout (temp copies of shp-groupshape.pptx /
mst-slide-layouts.pptx) -- no fixture on disk carries any tags to begin
with, so every test starts from "no tags yet" and proves the write path,
same limitation specs/identity-tags.md notes for why this was verified
against ECMA-376 directly rather than an existing example file.
"""

import os
import shutil
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from discovery import NS, discover_from_pptx, discover_from_pptx_layout  # noqa: E402
from identity_tags import (  # noqa: E402
    read_shape_tags,
    read_slide_tags,
    upsert_shape_tags,
    upsert_slide_tags,
)

FIXTURES = os.path.join(os.path.dirname(__file__), "..", "test-fixtures")
SLIDE_PART = "ppt/slides/slide1.xml"
LAYOUT_PART = "ppt/slideLayouts/slideLayout1.xml"


def _copy_fixture(name):
    fd, dst = tempfile.mkstemp(suffix=".pptx")
    os.close(fd)
    shutil.copyfile(os.path.join(FIXTURES, name), dst)
    return dst


def _tags_parts(path):
    with zipfile.ZipFile(path) as z:
        return sorted(n for n in z.namelist() if n.startswith("ppt/tags/tag") and n.endswith(".xml"))


def _read_part(path, part_name):
    with zipfile.ZipFile(path) as z:
        return z.read(part_name)


def test_read_slide_tags_returns_empty_when_no_relationship_exists():
    path = _copy_fixture("shp-groupshape.pptx")
    try:
        assert read_slide_tags(path, SLIDE_PART) == {}
    finally:
        os.remove(path)


def test_upsert_slide_tags_creates_the_full_relationship_graph_from_scratch():
    path = _copy_fixture("shp-groupshape.pptx")
    try:
        upsert_slide_tags(path, SLIDE_PART, {"instance_key": "rec-1", "slide_type": "quarterly-update"})

        assert read_slide_tags(path, SLIDE_PART) == {"instance_key": "rec-1", "slide_type": "quarterly-update"}

        tags_parts = _tags_parts(path)
        assert len(tags_parts) == 1

        rels_bytes = _read_part(path, "ppt/slides/_rels/slide1.xml.rels")
        rels_root = ET.fromstring(rels_bytes)
        rel_types = {
            el.get("Type") for el in rels_root.findall("{http://schemas.openxmlformats.org/package/2006/relationships}Relationship")
        }
        assert "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags" in rel_types
        # The pre-existing slideLayout relationship must survive untouched.
        assert "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" in rel_types

        ct_bytes = _read_part(path, "[Content_Types].xml")
        ct_root = ET.fromstring(ct_bytes)
        overrides = {
            el.get("PartName")
            for el in ct_root.findall("{http://schemas.openxmlformats.org/package/2006/content-types}Override")
        }
        assert f"/{tags_parts[0]}" in overrides
    finally:
        os.remove(path)


def test_upsert_slide_tags_merges_without_dropping_existing_tags():
    path = _copy_fixture("shp-groupshape.pptx")
    try:
        upsert_slide_tags(path, SLIDE_PART, {"a": "1"})
        upsert_slide_tags(path, SLIDE_PART, {"b": "2"})

        assert read_slide_tags(path, SLIDE_PART) == {"a": "1", "b": "2"}
        assert len(_tags_parts(path)) == 1  # reused, not duplicated
    finally:
        os.remove(path)


def test_upsert_slide_tags_overwrites_only_the_given_key():
    path = _copy_fixture("shp-groupshape.pptx")
    try:
        upsert_slide_tags(path, SLIDE_PART, {"a": "1", "b": "2"})
        upsert_slide_tags(path, SLIDE_PART, {"a": "revised"})

        assert read_slide_tags(path, SLIDE_PART) == {"a": "revised", "b": "2"}
    finally:
        os.remove(path)


def test_read_shape_tags_returns_empty_when_no_custdatalst():
    path = _copy_fixture("shp-groupshape.pptx")
    try:
        shape = discover_from_pptx(path)[0]
        assert read_shape_tags(path, SLIDE_PART, shape) == {}
    finally:
        os.remove(path)


def test_upsert_shape_tags_creates_and_persists():
    path = _copy_fixture("shp-groupshape.pptx")
    try:
        shape = discover_from_pptx(path)[0]
        upsert_shape_tags(path, SLIDE_PART, shape, {"role": "ph_status"})

        # Re-discover to get a fresh Candidate (z_order-stable, but proves
        # the write didn't disturb shape discovery itself).
        reread_shape = discover_from_pptx(path)[0]
        assert read_shape_tags(path, SLIDE_PART, reread_shape) == {"role": "ph_status"}
    finally:
        os.remove(path)


def test_upsert_shape_tags_merges_without_dropping_existing_tags():
    path = _copy_fixture("shp-groupshape.pptx")
    try:
        shape = discover_from_pptx(path)[0]
        upsert_shape_tags(path, SLIDE_PART, shape, {"role": "ph_status"})
        upsert_shape_tags(path, SLIDE_PART, shape, {"format": "percentage"})

        assert read_shape_tags(path, SLIDE_PART, shape) == {"role": "ph_status", "format": "percentage"}
        assert len(_tags_parts(path)) == 1  # reused, not duplicated
    finally:
        os.remove(path)


def test_multiple_shapes_on_one_slide_get_independent_tags():
    path = _copy_fixture("shp-groupshape.pptx")
    try:
        shapes = discover_from_pptx(path)
        assert len(shapes) >= 2
        upsert_shape_tags(path, SLIDE_PART, shapes[0], {"role": "title"})
        upsert_shape_tags(path, SLIDE_PART, shapes[1], {"role": "value"})

        reread = discover_from_pptx(path)
        assert read_shape_tags(path, SLIDE_PART, reread[0]) == {"role": "title"}
        assert read_shape_tags(path, SLIDE_PART, reread[1]) == {"role": "value"}
        assert len(_tags_parts(path)) == 2  # independent parts, no cross-contamination
    finally:
        os.remove(path)


def test_upsert_shape_tags_preserves_the_ph_element_and_schema_order_on_a_placeholder():
    path = _copy_fixture("mst-slide-layouts.pptx")
    try:
        title = next(c for c in discover_from_pptx_layout(path, 1) if c.placeholder_type == "title")
        upsert_shape_tags(path, LAYOUT_PART, title, {"role": "ph_title"})

        reread = discover_from_pptx_layout(path, 1)
        reread_title = next(c for c in reread if c.placeholder_type == "title")
        # The ph element (and its type/idx) must still be intact -- proves
        # inserting custDataLst didn't clobber a sibling element.
        assert reread_title.placeholder_type == "title"
        assert read_shape_tags(path, LAYOUT_PART, reread_title) == {"role": "ph_title"}

        # Schema order: ph must still come before custDataLst in nvPr. Locate
        # the title's specific nvPr -- slideLayout1.xml has multiple nvPr
        # elements (one per shape), and the first one in document order isn't
        # necessarily the title's.
        with zipfile.ZipFile(path) as z:
            root = ET.parse(z.open(LAYOUT_PART)).getroot()
        title_nvpr = next(
            nvpr
            for nvpr in root.findall(f".//{{{NS['p']}}}nvPr")
            if nvpr.find(f"{{{NS['p']}}}ph[@type='title']") is not None
        )
        children_tags = [c.tag.split("}")[-1] for c in title_nvpr]
        assert children_tags.index("ph") < children_tags.index("custDataLst")
    finally:
        os.remove(path)


def test_write_never_touches_unrelated_zip_entries():
    path = _copy_fixture("shp-groupshape.pptx")
    try:
        with zipfile.ZipFile(path) as z:
            before = {name: z.read(name) for name in z.namelist() if name != "ppt/slides/_rels/slide1.xml.rels"}

        upsert_slide_tags(path, SLIDE_PART, {"instance_key": "rec-1"})

        with zipfile.ZipFile(path) as z:
            untouched_names = set(before) - {SLIDE_PART, "[Content_Types].xml"}
            after = {name: z.read(name) for name in untouched_names}
        assert after == {name: before[name] for name in untouched_names}
    finally:
        os.remove(path)
