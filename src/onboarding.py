"""Onboarding: match a new slide of an already-established type against its
template, per specs/onboarding.md.

First-time onboarding of a type needs no code here at all -- per the
underlying skill's own design, the first example slide *becomes* the
reference by direct tagging (identity_tags) once a human confirms it, and
its verification is exactly inject_primitive hitting the no-op path (already
proven in tests/test_resolve.py). This module is specifically the gap that
was missing: matching a *subsequent* slide's candidates against an
established template.
"""

from __future__ import annotations

from dataclasses import dataclass

from discovery import Candidate, discover_from_pptx_part
from identity_tags import read_shape_tags, upsert_shape_tags, upsert_slide_tags
from matching import Confidence, MatchResult, match
from sync_operations import SlideInstance


@dataclass(frozen=True)
class FieldMatch:
    role: str
    result: MatchResult


def match_slide_against_template(path: str, slide_part: str, template: SlideInstance) -> list[FieldMatch]:
    """For each field role `template` defines, score every untagged
    candidate on `slide_part` against the template's reference shape for
    that role, per specs/matching.md's tier-2 path (already-tagged
    candidates are excluded from the pool -- they're either this same field
    from a prior pass or a different field entirely, neither is a fresh
    match target).
    """
    candidates = discover_from_pptx_part(path, slide_part)
    untagged = [c for c in candidates if read_shape_tags(path, slide_part, c).get("role") is None]

    return [FieldMatch(role=role, result=match(untagged, reference_shape)) for role, reference_shape in template.field_shapes.items()]


def confirm_field_match(path: str, slide_part: str, role: str, shape: Candidate) -> None:
    """Write `role` onto `shape` -- the confirmation primitive for a match a
    human has decided on, whether that's accepting a medium-confidence
    result matching.py already scored, or a direct selection an eventual
    onboarding UI would make (see specs/onboarding.md's Non-goals: this is
    the primitive that UI would call, not the UI itself).
    """
    upsert_shape_tags(path, slide_part, shape, {"role": role})


def onboard_new_instance(
    path: str, slide_part: str, template: SlideInstance, slide_type: str, instance_key: str
) -> list[FieldMatch]:
    """Tag a new instance's slide-level identity (supplied by whatever
    created it -- e.g. a sync-operations duplication -- not matched) and
    auto-accept any high-confidence field matches against `template`,
    writing their tags immediately so they're a tier-1 fast match next time
    (specs/matching.md's confidence_thresholds: "the system gets more
    self-healing over time, not less"). Medium/low-confidence matches are
    returned but never auto-tagged -- a human decides via
    confirm_field_match(), same as any other unresolved match.
    """
    upsert_slide_tags(path, slide_part, {"slide_type": slide_type, "instance_key": instance_key})

    matches = match_slide_against_template(path, slide_part, template)
    for field_match in matches:
        if field_match.result.confidence is Confidence.HIGH and field_match.result.candidate is not None:
            confirm_field_match(path, slide_part, field_match.role, field_match.result.candidate)
    return matches
