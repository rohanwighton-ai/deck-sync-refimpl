"""Tests for src/onboarding.py against specs/onboarding.md.

Uses mst-slide-layouts.pptx's two layouts as template vs. new-instance:
layout1 (tagged, the template) and layout2 (untagged, the "new slide" being
matched against it). Real fixture geometry drift between the two layouts is
what already makes their title placeholder match high-confidence and their
body placeholder match medium-confidence (see test_matching.py's own
cross-layout test) -- this exercises both branches for real, not by
constructing synthetic scores.
"""

import os
import shutil
import sys
import tempfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from discovery import discover_from_pptx_layout  # noqa: E402
from identity_tags import upsert_shape_tags, upsert_slide_tags  # noqa: E402
from matching import Confidence  # noqa: E402
from onboarding import confirm_field_match, match_slide_against_template, onboard_new_instance  # noqa: E402
from resolve import resolve_slide_instance  # noqa: E402

FIXTURES = os.path.join(os.path.dirname(__file__), "..", "test-fixtures")
LAYOUT1_PART = "ppt/slideLayouts/slideLayout1.xml"
LAYOUT2_PART = "ppt/slideLayouts/slideLayout2.xml"


def _copy_fixture(name):
    fd, dst = tempfile.mkstemp(suffix=".pptx")
    os.close(fd)
    shutil.copyfile(os.path.join(FIXTURES, name), dst)
    return dst


def _build_template(path):
    """Tag layout1 (title + body) and resolve it -- simulates first-time
    onboarding having already happened for this type, exactly as
    specs/onboarding.md says needs no new code (direct tagging, no
    scoring)."""
    title = next(c for c in discover_from_pptx_layout(path, 1) if c.placeholder_type == "title")
    body = next(c for c in discover_from_pptx_layout(path, 1) if c.placeholder_type == "body")
    upsert_slide_tags(path, LAYOUT1_PART, {"slide_type": "quarterly-update", "instance_key": "rec-1"})
    upsert_shape_tags(path, LAYOUT1_PART, title, {"role": "Title"})
    upsert_shape_tags(path, LAYOUT1_PART, body, {"role": "Body"})
    return resolve_slide_instance(path, LAYOUT1_PART)


def test_match_slide_against_template_scores_both_high_and_medium_confidence():
    path = _copy_fixture("mst-slide-layouts.pptx")
    try:
        template = _build_template(path)

        matches = match_slide_against_template(path, LAYOUT2_PART, template)

        by_role = {m.role: m.result for m in matches}
        assert set(by_role) == {"Title", "Body"}
        assert by_role["Title"].confidence is Confidence.HIGH
        assert by_role["Title"].candidate is not None
        assert by_role["Body"].confidence is Confidence.MEDIUM
        assert by_role["Body"].candidate is None  # medium is never auto-accepted
    finally:
        os.remove(path)


def test_onboard_new_instance_auto_accepts_high_confidence_and_leaves_medium_unresolved():
    path = _copy_fixture("mst-slide-layouts.pptx")
    try:
        template = _build_template(path)

        onboard_new_instance(path, LAYOUT2_PART, template, slide_type="quarterly-update", instance_key="rec-2")

        instance = resolve_slide_instance(path, LAYOUT2_PART)
        assert instance.instance_key == "rec-2"
        assert instance.type_tag == "quarterly-update"
        # Title (high confidence) got auto-tagged; Body (medium) did not.
        assert set(instance.field_shapes) == {"Title"}
    finally:
        os.remove(path)


def test_confirm_field_match_resolves_a_flagged_medium_confidence_match():
    # Simulates the human-decides step for a medium-confidence flag --
    # whether that decision came from re-running the matcher or (per
    # specs/onboarding.md's Non-goals) an eventual direct shape-selection UI,
    # this is the primitive either path calls.
    path = _copy_fixture("mst-slide-layouts.pptx")
    try:
        template = _build_template(path)
        onboard_new_instance(path, LAYOUT2_PART, template, slide_type="quarterly-update", instance_key="rec-2")

        body_candidate = next(c for c in discover_from_pptx_layout(path, 2) if c.placeholder_type == "body")
        confirm_field_match(path, LAYOUT2_PART, "Body", body_candidate)

        instance = resolve_slide_instance(path, LAYOUT2_PART)
        assert set(instance.field_shapes) == {"Title", "Body"}
    finally:
        os.remove(path)


def test_onboard_new_instance_tags_slide_identity_unconditionally():
    # Slide-level identity is supplied by whatever created the instance, not
    # matched -- it must be written even if every field match were to fail.
    path = _copy_fixture("mst-slide-layouts.pptx")
    try:
        template = _build_template(path)

        onboard_new_instance(path, LAYOUT2_PART, template, slide_type="quarterly-update", instance_key="rec-2")

        instance = resolve_slide_instance(path, LAYOUT2_PART)
        assert instance.instance_key == "rec-2"
        assert instance.type_tag == "quarterly-update"
    finally:
        os.remove(path)
