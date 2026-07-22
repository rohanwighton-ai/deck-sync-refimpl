"""End-to-end composition test: real deck -> identity tags on disk ->
resolve_slide_instance() -> sync_operations dispatch -> real Excel row.

This is the "primitives actually compose" proof flagged as missing after
Priority 6: every module was tested in isolation, but nothing before this
exercised discovery + identity_tags + sync_operations + excel_output
together against a real fixture.
"""

import os
import shutil
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from discovery import NS, discover_from_pptx_layout  # noqa: E402
from excel_output import Sheet  # noqa: E402
from identity_tags import upsert_shape_tags, upsert_slide_tags  # noqa: E402
from resolve import resolve_slide_instance  # noqa: E402
from sync_operations import Flagged, InPlaceCorrection, NewRecord, NoChange, plan_routine_sync  # noqa: E402

FIXTURES = os.path.join(os.path.dirname(__file__), "..", "test-fixtures")
LAYOUT_PART = "ppt/slideLayouts/slideLayout1.xml"
TITLE_SEED_TEXT = "Click to edit Master title style"


def _copy_fixture(name):
    fd, dst = tempfile.mkstemp(suffix=".pptx")
    os.close(fd)
    shutil.copyfile(os.path.join(FIXTURES, name), dst)
    return dst


def _onboard(path):
    """Simulate onboarding having already happened: tag the slide and its
    title field, exactly as a real onboarding workflow would leave a deck."""
    title = next(c for c in discover_from_pptx_layout(path, 1) if c.placeholder_type == "title")
    upsert_slide_tags(path, LAYOUT_PART, {"slide_type": "quarterly-update", "instance_key": "rec-1"})
    upsert_shape_tags(path, LAYOUT_PART, title, {"role": "Title"})


def _current_title_text(path):
    with zipfile.ZipFile(path) as z:
        root = ET.parse(z.open(LAYOUT_PART)).getroot()
    for sp in root.iter(f"{{{NS['p']}}}sp"):
        if sp.find(f".//{{{NS['p']}}}ph[@type='title']") is not None:
            return "".join(t.text or "" for t in sp.findall(f".//{{{NS['a']}}}t"))
    raise AssertionError("title shape not found")


def test_resolve_reads_real_tags_off_disk_into_a_slide_instance():
    path = _copy_fixture("mst-slide-layouts.pptx")
    try:
        _onboard(path)

        instance = resolve_slide_instance(path, LAYOUT_PART)

        assert instance.instance_key == "rec-1"
        assert instance.type_tag == "quarterly-update"
        assert set(instance.field_shapes) == {"Title"}
    finally:
        os.remove(path)


def test_untagged_slide_resolves_to_none_key_and_type():
    path = _copy_fixture("mst-slide-layouts.pptx")
    try:
        instance = resolve_slide_instance(path, LAYOUT_PART)

        assert instance.instance_key is None
        assert instance.type_tag is None
        assert instance.field_shapes == {}
    finally:
        os.remove(path)


def test_end_to_end_no_change_then_in_place_correction():
    path = _copy_fixture("mst-slide-layouts.pptx")
    try:
        _onboard(path)
        instance = resolve_slide_instance(path, LAYOUT_PART)

        no_change_sheet = Sheet(
            deck_reference="deck-v1", fields=["Title"], instance_order=["rec-1"], rows={"rec-1": {"Title": TITLE_SEED_TEXT}}
        )
        actions = plan_routine_sync(path, [instance], no_change_sheet)
        assert actions == [NoChange(instance_key="rec-1")]
        assert _current_title_text(path) == TITLE_SEED_TEXT  # untouched

        # Re-resolve (fresh Candidates -- z_order-stable but proves nothing broke).
        instance = resolve_slide_instance(path, LAYOUT_PART)
        correction_sheet = Sheet(
            deck_reference="deck-v1", fields=["Title"], instance_order=["rec-1"], rows={"rec-1": {"Title": "Q3 Revenue"}}
        )
        actions = plan_routine_sync(path, [instance], correction_sheet)

        assert len(actions) == 1
        assert isinstance(actions[0], InPlaceCorrection)
        assert actions[0].changed_fields["Title"].verified is True
        assert _current_title_text(path) == "Q3 Revenue"  # the real write actually landed
    finally:
        os.remove(path)


def test_end_to_end_new_record_when_data_sheet_row_has_no_onboarded_instance():
    path = _copy_fixture("mst-slide-layouts.pptx")
    try:
        _onboard(path)
        instance = resolve_slide_instance(path, LAYOUT_PART)

        sheet = Sheet(
            deck_reference="deck-v1",
            fields=["Title"],
            instance_order=["rec-1", "rec-new"],
            rows={"rec-1": {"Title": TITLE_SEED_TEXT}, "rec-new": {"Title": "Brand New"}},
        )
        actions = plan_routine_sync(path, [instance], sheet)

        no_change = [a for a in actions if isinstance(a, NoChange)]
        new_records = [a for a in actions if isinstance(a, NewRecord)]
        assert no_change == [NoChange(instance_key="rec-1")]
        assert new_records == [
            NewRecord(row_instance_key="rec-new", values={"Title": "Brand New"}, reason="no known slide instance carries this row's instance key")
        ]
    finally:
        os.remove(path)


def test_end_to_end_never_onboarded_slide_is_flagged_not_silently_skipped():
    path = _copy_fixture("mst-slide-layouts.pptx")
    try:
        instance = resolve_slide_instance(path, LAYOUT_PART)  # no _onboard() call

        sheet = Sheet(deck_reference="deck-v1", fields=["Title"], instance_order=[], rows={})
        actions = plan_routine_sync(path, [instance], sheet)

        assert len(actions) == 1
        assert isinstance(actions[0], Flagged)
        assert actions[0].kind == "unclassified_slide"
    finally:
        os.remove(path)
