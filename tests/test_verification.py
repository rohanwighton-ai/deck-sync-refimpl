"""Tests for src/verification.py's inject_primitive against specs/verification.md.

Uses a temp copy of mst-slide-layouts.pptx's title placeholder (real text:
"Click to edit Master title style") so writes exercise a real zip write-back,
never the checked-in fixture itself.
"""

import os
import shutil
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from discovery import NS, discover_from_pptx, discover_from_pptx_layout  # noqa: E402
from verification import inject_primitive  # noqa: E402

FIXTURES = os.path.join(os.path.dirname(__file__), "..", "test-fixtures")

LAYOUT_PART = "ppt/slideLayouts/slideLayout1.xml"
TITLE_SEED_TEXT = "Click to edit Master title style"


def _copy_fixture(name):
    fd, dst = tempfile.mkstemp(suffix=".pptx")
    os.close(fd)
    shutil.copyfile(os.path.join(FIXTURES, name), dst)
    return dst


def _read_text(path, part_name, z_order):
    with zipfile.ZipFile(path) as z, z.open(part_name) as f:
        root = ET.parse(f).getroot()
    spTree = root.find(".//p:spTree", NS)
    shapes = [c for c in spTree.iter() if c.tag.split("}")[-1] in ("sp", "pic")]
    sp = shapes[z_order - 1]
    return "".join(t.text or "" for t in sp.findall(".//a:t", NS))


def test_inject_primitive_is_a_noop_when_current_value_already_matches_source():
    path = _copy_fixture("mst-slide-layouts.pptx")
    try:
        title = next(c for c in discover_from_pptx_layout(path, 1) if c.placeholder_type == "title")
        result = inject_primitive(path, LAYOUT_PART, title, TITLE_SEED_TEXT)

        assert result.written is False
        assert result.verified is True
        assert result.initial_hash == result.source_hash == result.final_hash
        # No-op really means no-op: bytes on disk are untouched.
        with open(path, "rb") as f:
            after = f.read()
        with open(os.path.join(FIXTURES, "mst-slide-layouts.pptx"), "rb") as f:
            original = f.read()
        assert after == original
    finally:
        os.remove(path)


def test_inject_primitive_writes_and_confirms_when_value_differs():
    path = _copy_fixture("mst-slide-layouts.pptx")
    try:
        title = next(c for c in discover_from_pptx_layout(path, 1) if c.placeholder_type == "title")
        result = inject_primitive(path, LAYOUT_PART, title, "Q3 Revenue")

        assert result.written is True
        assert result.verified is True
        assert result.initial_hash != result.source_hash
        assert result.final_hash == result.source_hash

        assert _read_text(path, LAYOUT_PART, title.z_order) == "Q3 Revenue"
    finally:
        os.remove(path)


def test_inject_primitive_leaves_every_other_zip_entry_byte_for_byte_untouched():
    path = _copy_fixture("mst-slide-layouts.pptx")
    try:
        with zipfile.ZipFile(path) as z:
            before = {name: z.read(name) for name in z.namelist() if name != LAYOUT_PART}

        title = next(c for c in discover_from_pptx_layout(path, 1) if c.placeholder_type == "title")
        inject_primitive(path, LAYOUT_PART, title, "Q3 Revenue")

        with zipfile.ZipFile(path) as z:
            assert set(z.namelist()) == set(before) | {LAYOUT_PART}
            after = {name: z.read(name) for name in z.namelist() if name != LAYOUT_PART}
        assert after == before
    finally:
        os.remove(path)


def test_inject_primitive_raises_rather_than_silently_dropping_a_write_it_cant_make():
    # shp-groupshape.pptx's shapes are empty decoration: a <p:txBody> with no
    # <a:r>/<a:t> run at all, only an empty endParaRPr. There's nowhere to
    # write a value, so this must fail loudly, not silently no-op or corrupt
    # the shape's XML by inventing a run.
    path = _copy_fixture("shp-groupshape.pptx")
    try:
        shape = discover_from_pptx(path)[0]
        try:
            inject_primitive(path, "ppt/slides/slide1.xml", shape, "anything")
            assert False, "expected ValueError: shape has no text runs"
        except ValueError:
            pass
    finally:
        os.remove(path)
