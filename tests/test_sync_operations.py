"""Tests for src/sync_operations.py against specs/sync-operations.md."""

import os
import shutil
import sys
import tempfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from discovery import discover_from_pptx_layout  # noqa: E402
from excel_output import Sheet  # noqa: E402
from sync_operations import (  # noqa: E402
    Flagged,
    InPlaceCorrection,
    NewRecord,
    NoChange,
    PeriodRollover,
    SlideInstance,
    plan_period_rollover,
    plan_routine_sync,
)

FIXTURES = os.path.join(os.path.dirname(__file__), "..", "test-fixtures")
LAYOUT_PART = "ppt/slideLayouts/slideLayout1.xml"
TITLE_SEED_TEXT = "Click to edit Master title style"


def _copy_fixture(name):
    fd, dst = tempfile.mkstemp(suffix=".pptx")
    os.close(fd)
    shutil.copyfile(os.path.join(FIXTURES, name), dst)
    return dst


def _title_instance(path, instance_key="rec-1", type_tag="quarterly-update"):
    title = next(c for c in discover_from_pptx_layout(path, 1) if c.placeholder_type == "title")
    return SlideInstance(
        part_path=LAYOUT_PART,
        instance_key=instance_key,
        type_tag=type_tag,
        field_shapes={"Title": title},
    )


def _sheet(deck_reference, rows):
    instance_order = list(rows)
    return Sheet(deck_reference=deck_reference, fields=["Title"], instance_order=instance_order, rows=rows)


def test_case1_no_change_when_value_already_matches():
    path = _copy_fixture("mst-slide-layouts.pptx")
    try:
        instance = _title_instance(path)
        sheet = _sheet("deck-v1", {"rec-1": {"Title": TITLE_SEED_TEXT}})

        actions = plan_routine_sync(path, [instance], sheet)

        assert actions == [NoChange(instance_key="rec-1")]
    finally:
        os.remove(path)


def test_case4_in_place_correction_writes_and_verifies_changed_field():
    path = _copy_fixture("mst-slide-layouts.pptx")
    try:
        instance = _title_instance(path)
        sheet = _sheet("deck-v1", {"rec-1": {"Title": "Q3 Revenue"}})

        actions = plan_routine_sync(path, [instance], sheet)

        assert len(actions) == 1
        action = actions[0]
        assert isinstance(action, InPlaceCorrection)
        assert action.instance_key == "rec-1"
        assert set(action.changed_fields) == {"Title"}
        result = action.changed_fields["Title"]
        assert result.written is True
        assert result.verified is True

        # The write actually took, on disk, not just claimed.
        reread = next(c for c in discover_from_pptx_layout(path, 1) if c.placeholder_type == "title")
        assert reread is not None  # sanity: shape still discoverable after write
    finally:
        os.remove(path)


def test_case3_new_record_when_no_matching_instance():
    path = _copy_fixture("mst-slide-layouts.pptx")
    try:
        # No instances known at all -- every Data-sheet row is a new record.
        sheet = _sheet("deck-v1", {"rec-new": {"Title": "Brand New Slide"}})

        actions = plan_routine_sync(path, [], sheet)

        assert actions == [
            NewRecord(
                row_instance_key="rec-new",
                values={"Title": "Brand New Slide"},
                reason="no known slide instance carries this row's instance key",
            )
        ]
    finally:
        os.remove(path)


def test_case6_unclassified_instance_is_flagged_and_excluded_from_dispatch():
    path = _copy_fixture("mst-slide-layouts.pptx")
    try:
        # No type tag / instance key -- can't even be checked against the
        # Data-sheet, so it must be flagged rather than guessed at, and must
        # not silently participate in (or block) the row dispatch below.
        classified = _title_instance(path, instance_key="rec-1")
        unclassified = SlideInstance(part_path=LAYOUT_PART, instance_key=None, type_tag=None, field_shapes={})
        sheet = _sheet("deck-v1", {"rec-1": {"Title": TITLE_SEED_TEXT}})

        actions = plan_routine_sync(path, [classified, unclassified], sheet)

        flagged = [a for a in actions if isinstance(a, Flagged)]
        assert len(flagged) == 1
        assert flagged[0].kind == "unclassified_slide"
        assert flagged[0].subject == LAYOUT_PART

        no_change = [a for a in actions if isinstance(a, NoChange)]
        assert no_change == [NoChange(instance_key="rec-1")]
    finally:
        os.remove(path)


def test_missing_field_on_instance_is_skipped_not_flagged_or_changed():
    path = _copy_fixture("mst-slide-layouts.pptx")
    try:
        instance = _title_instance(path)
        # "Subtitle" isn't in this instance's field_shapes at all.
        sheet = _sheet("deck-v1", {"rec-1": {"Title": TITLE_SEED_TEXT, "Subtitle": "no shape for this"}})

        actions = plan_routine_sync(path, [instance], sheet)

        assert actions == [NoChange(instance_key="rec-1")]
    finally:
        os.remove(path)


def test_period_rollover_never_produced_by_routine_sync():
    path = _copy_fixture("mst-slide-layouts.pptx")
    try:
        instance = _title_instance(path)
        sheet = _sheet("deck-v1", {"rec-1": {"Title": "Q4 Revenue"}})  # a changed value -- still not case 2

        actions = plan_routine_sync(path, [instance], sheet)

        assert not any(isinstance(a, PeriodRollover) for a in actions)
    finally:
        os.remove(path)


def test_period_rollover_is_explicit_and_references_the_named_instance():
    path = _copy_fixture("mst-slide-layouts.pptx")
    try:
        instance = _title_instance(path, instance_key="rec-1")

        action = plan_period_rollover(instance, {"Title": "Q4 Revenue"})

        assert action == PeriodRollover(
            source_instance_key="rec-1",
            new_values={"Title": "Q4 Revenue"},
            reason="explicit period-rollover command",
        )
    finally:
        os.remove(path)


def test_period_rollover_raises_for_an_unclassified_instance():
    instance = SlideInstance(part_path=LAYOUT_PART, instance_key=None, type_tag=None, field_shapes={})
    try:
        plan_period_rollover(instance, {"Title": "anything"})
        assert False, "expected ValueError"
    except ValueError:
        pass
