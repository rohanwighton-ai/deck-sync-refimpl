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

from discovery import NS, Candidate, discover_from_pptx, discover_from_pptx_layout  # noqa: E402
from verification import inject_primitive, verify_structure, verify_structure_from_pptx  # noqa: E402

for _prefix, _uri in NS.items():
    ET.register_namespace(_prefix, _uri)

FIXTURES = os.path.join(os.path.dirname(__file__), "..", "test-fixtures")

LAYOUT_PART = "ppt/slideLayouts/slideLayout1.xml"
TITLE_SEED_TEXT = "Click to edit Master title style"
SOURCE_SLIDE_PART = "ppt/slides/slide1.xml"
DUPLICATE_SLIDE_PART = "ppt/slides/slide2.xml"


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


def _duplicate_slide_fixture(mutate=None):
    """Build a temp copy of shp-groupshape.pptx with a second slide part
    (ppt/slides/slide2.xml) appended -- a duplicate of slide1.xml, optionally
    passed through `mutate` first. No checked-in fixture has more than one
    slide, so this synthesizes the minimal thing verify_structure needs: two
    shape-tree-bearing parts in one zip. discover_from_pptx_part() only ever
    opens a named zip member directly, so the archive doesn't need a valid
    [Content_Types].xml/relationships/presentation.xml entry for slide2 to be
    readable here -- only a real pptx viewer would care about those.
    """
    path = _copy_fixture("shp-groupshape.pptx")
    with zipfile.ZipFile(path) as z:
        slide1_bytes = z.read(SOURCE_SLIDE_PART)
    slide2_bytes = mutate(slide1_bytes) if mutate else slide1_bytes
    with zipfile.ZipFile(path, "a") as z:
        z.writestr(DUPLICATE_SLIDE_PART, slide2_bytes)
    return path


def _drop_last_shape(xml_bytes):
    root = ET.fromstring(xml_bytes)
    spTree = root.find(".//p:spTree", NS)
    direct_shapes = [c for c in spTree if c.tag.split("}")[-1] in ("sp", "pic")]
    spTree.remove(direct_shapes[-1])
    return ET.tostring(root, encoding="unicode").encode("utf-8")


def _retag_last_shape_as_picture(xml_bytes):
    root = ET.fromstring(xml_bytes)
    spTree = root.find(".//p:spTree", NS)
    direct_shapes = [c for c in spTree if c.tag.split("}")[-1] == "sp"]
    direct_shapes[-1].tag = direct_shapes[-1].tag.replace("}sp", "}pic")
    return ET.tostring(root, encoding="unicode").encode("utf-8")


def test_verify_structure_ok_for_an_identical_duplicate():
    path = _duplicate_slide_fixture()
    try:
        result = verify_structure_from_pptx(path, SOURCE_SLIDE_PART, DUPLICATE_SLIDE_PART)
        assert result.ok
        assert result.source_count == result.duplicate_count == 4
        assert result.mismatches == []
    finally:
        os.remove(path)


def test_verify_structure_flags_a_missing_shape_rather_than_assuming_duplication_succeeded():
    path = _duplicate_slide_fixture(mutate=_drop_last_shape)
    try:
        result = verify_structure_from_pptx(path, SOURCE_SLIDE_PART, DUPLICATE_SLIDE_PART)
        assert not result.ok
        assert result.source_count == 4
        assert result.duplicate_count == 3
        kinds = {m.kind for m in result.mismatches}
        assert "shape_count" in kinds
        assert "missing_in_duplicate" in kinds
    finally:
        os.remove(path)


def test_verify_structure_flags_a_shape_type_change():
    path = _duplicate_slide_fixture(mutate=_retag_last_shape_as_picture)
    try:
        result = verify_structure_from_pptx(path, SOURCE_SLIDE_PART, DUPLICATE_SLIDE_PART)
        assert not result.ok
        assert result.source_count == result.duplicate_count == 4
        type_mismatches = [m for m in result.mismatches if m.kind == "type"]
        assert len(type_mismatches) == 1
        assert type_mismatches[0].index == 3  # "Rectangle 5", the 4th (last) shape
    finally:
        os.remove(path)


def test_verify_structure_flags_an_identity_tag_mismatch():
    # identity_tag is always None straight out of discover() (no physical
    # storage format decided yet -- see IMPLEMENTATION_PLAN.md's notes), so
    # this constructs Candidates directly, same as test_matching.py's tier-1
    # tests, rather than round-tripping through a pptx that can't yet carry
    # a tag on disk.
    def _c(z_order, tag):
        return Candidate(
            name=f"shape{z_order}",
            group_path=(),
            z_order=z_order,
            shape_type="autoshape_or_textbox",
            placeholder_type=None,
            placeholder_idx=None,
            has_text=True,
            identity_tag=tag,
        )

    source = [_c(1, "title_field")]
    duplicate = [_c(1, "subtitle_field")]

    result = verify_structure(source, duplicate)
    assert not result.ok
    tag_mismatches = [m for m in result.mismatches if m.kind == "identity_tag"]
    assert len(tag_mismatches) == 1
    assert tag_mismatches[0].index == 0


def test_verify_structure_reports_extra_shapes_in_duplicate_too():
    duplicate_candidates = discover_from_pptx(os.path.join(FIXTURES, "shp-groupshape.pptx"))

    result = verify_structure([], duplicate_candidates)

    assert not result.ok
    assert result.source_count == 0
    assert result.duplicate_count == 4
    kinds = {m.kind for m in result.mismatches}
    assert "shape_count" in kinds
    assert "extra_in_duplicate" in kinds
