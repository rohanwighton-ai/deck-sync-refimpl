"""Tests for src/matching.py against specs/matching.md.

Sibling-ambiguity / z-order cases use test-fixtures/shp-groupshape.pptx (its
SOURCE.md-stated purpose); everything else is exercised with directly
constructed Candidate objects since matching operates purely on Candidate
data, independent of how a Candidate was discovered.
"""

import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from discovery import Candidate, discover_from_pptx, discover_from_pptx_layout  # noqa: E402
from matching import Confidence, match, score_candidate  # noqa: E402

FIXTURES = os.path.join(os.path.dirname(__file__), "..", "test-fixtures")


def _candidate(name, z_order=1, **overrides):
    defaults = dict(
        name=name,
        group_path=(),
        z_order=z_order,
        shape_type="autoshape_or_textbox",
        placeholder_type=None,
        placeholder_idx=None,
        has_text=True,
    )
    defaults.update(overrides)
    return Candidate(**defaults)


def test_tier1_trusts_a_single_valid_existing_tag_without_scoring():
    reference = _candidate("Reference")
    tagged = _candidate("Tagged", identity_tag="status_field")
    other = _candidate("Other", shape_type="picture")  # would score terribly if scored

    result = match([tagged, other], reference, valid_tags={"status_field"})

    assert result.candidate is tagged
    assert result.confidence == Confidence.HIGH
    assert result.score is None
    assert "identity tag" in result.reason


def test_tier1_ignores_a_tag_not_in_valid_tags_and_falls_back_to_scoring():
    reference = _candidate("Reference", position=(0, 0), size=(914400, 914400))
    stale_tagged = _candidate(
        "Stale", identity_tag="wrong_field", shape_type="picture", position=(9000000, 9000000), size=(1, 1)
    )
    good_match = _candidate("Good", position=(0, 0), size=(914400, 914400))

    result = match([stale_tagged, good_match], reference, valid_tags={"status_field"})

    assert result.candidate is good_match
    assert result.confidence == Confidence.HIGH
    assert result.score is not None


def test_tier1_flags_a_same_tag_collision():
    reference = _candidate("Reference")
    a = _candidate("A", identity_tag="status_field")
    b = _candidate("B", identity_tag="status_field")

    result = match([a, b], reference, valid_tags={"status_field"})

    assert result.candidate is None
    assert result.confidence == Confidence.MEDIUM
    assert "collision" in result.reason


def test_placeholder_index_match_scores_high_confidence():
    reference = _candidate("Reference", placeholder_type="body", placeholder_idx=10)
    same_idx = _candidate("Same idx", placeholder_type="body", placeholder_idx=10)
    different_idx = _candidate("Different idx", placeholder_type="body", placeholder_idx=2)

    result = match([different_idx, same_idx], reference)

    assert result.candidate is same_idx
    assert result.confidence == Confidence.HIGH


def test_no_matching_signal_is_low_confidence_and_unmatched():
    reference = _candidate(
        "Reference", shape_type="picture", has_text=False, position=(0, 0), size=(914400, 914400)
    )
    unrelated = _candidate(
        "Unrelated",
        shape_type="autoshape_or_textbox",
        has_text=True,
        position=(50_000_000, 50_000_000),
        size=(1, 1),
    )

    result = match([unrelated], reference)

    assert result.candidate is None
    assert result.confidence == Confidence.LOW


def test_score_candidate_ignores_inapplicable_signals_rather_than_penalizing():
    # Reference carries no placeholder and no geometry -- score_candidate must
    # renormalize across only the signals that apply (shape type, content),
    # not silently treat the missing ones as zero.
    reference = _candidate("Reference", has_text=True)
    perfect_on_applicable_signals = _candidate("Match", has_text=True)

    assert score_candidate(perfect_on_applicable_signals, reference) == 1.0


def test_shp_groupshape_sibling_ambiguity_resolved_by_zorder():
    # All four leaf shapes share shape_type and has_text=False. A reference
    # with no placeholder/geometry data leaves only those two signals
    # applicable, so all four score identically -- a real 4-way tie. z-order
    # (Oval 2 is z=2, matching the reference's z_order=2 uniquely) is the
    # only thing that can break it.
    path = os.path.join(FIXTURES, "shp-groupshape.pptx")
    candidates = discover_from_pptx(path)
    reference = _candidate("Reference", z_order=2, has_text=False)

    result = match(candidates, reference)

    assert result.confidence == Confidence.HIGH
    assert result.candidate is not None
    assert result.candidate.name == "Oval 2"


def test_shp_groupshape_sibling_ambiguity_flagged_when_zorder_also_ties():
    # "Rounded Rectangle 1" (z=1) and "Isosceles Triangle 3" (z=3) are
    # equidistant in z-order from a reference at z=2 -- z-order can't
    # disambiguate here, so this must be flagged, not arbitrarily resolved.
    path = os.path.join(FIXTURES, "shp-groupshape.pptx")
    candidates = discover_from_pptx(path)
    rr1 = next(c for c in candidates if c.name == "Rounded Rectangle 1")
    tri3 = next(c for c in candidates if c.name == "Isosceles Triangle 3")
    reference = _candidate("Reference", z_order=2, has_text=False)

    result = match([rr1, tri3], reference)

    assert result.candidate is None
    assert result.confidence == Confidence.MEDIUM
    assert "sibling ambiguity" in result.reason


def test_match_with_no_candidates_is_low_confidence_unmatched():
    reference = _candidate("Reference")
    result = match([], reference)
    assert result.candidate is None
    assert result.confidence == Confidence.LOW


def test_mst_slide_layouts_placeholder_index_alone_does_not_force_high_confidence():
    # The body placeholder (idx=10) is present on both layouts but has drifted
    # to very different geometry between them. Placeholder-index match is the
    # strongest signal, but specs/matching.md is explicit that no single
    # signal should be taken blindly -- combined with geometry this far off,
    # the correct outcome is medium confidence (flagged), not an auto-accept.
    path = os.path.join(FIXTURES, "mst-slide-layouts.pptx")
    layout1 = discover_from_pptx_layout(path, layout_index=1)
    layout2 = discover_from_pptx_layout(path, layout_index=2)
    reference = next(c for c in layout1 if c.name == "Text Placeholder 3")

    result = match(layout2, reference)

    assert result.candidate is None
    assert result.confidence == Confidence.MEDIUM
