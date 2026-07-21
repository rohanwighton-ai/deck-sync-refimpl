"""Field matching: given a pool of discovered candidate shapes and a reference
candidate (the example a field type was configured from), determine which
candidate, if any, represents the same field.

See specs/matching.md for the requirements this implements.
"""

from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from typing import Container, Sequence

from discovery import Candidate

# Signal weights, in the reliability order specs/matching.md prescribes:
# placeholder index > geometric similarity > shape type > content pattern.
_PLACEHOLDER_WEIGHT = 0.5
_GEOMETRY_WEIGHT = 0.3
_SHAPE_TYPE_WEIGHT = 0.15
_CONTENT_WEIGHT = 0.05

# EMU tolerances for "close enough" geometry (914400 EMU = 1 inch).
_POSITION_TOLERANCE_EMU = 914400
_SIZE_TOLERANCE_EMU = 914400

_HIGH_THRESHOLD = 0.75
_MEDIUM_THRESHOLD = 0.4

# Two scores within this much of each other are "similarly close" per
# specs/matching.md's sibling_ambiguity rule.
_SIBLING_GAP_THRESHOLD = 0.1


class Confidence(str, Enum):
    HIGH = "high"
    MEDIUM = "medium"
    LOW = "low"


@dataclass(frozen=True)
class MatchResult:
    candidate: Candidate | None  # None unless a match was accepted
    confidence: Confidence
    score: float | None  # None for a tier-1 tag-trust match (no scoring performed)
    reason: str


def _placeholder_score(candidate: Candidate, reference: Candidate) -> float | None:
    if reference.placeholder_idx is None:
        return None  # signal not applicable: reference isn't a placeholder
    if candidate.placeholder_idx is None:
        return 0.0
    return (
        1.0
        if (candidate.placeholder_type, candidate.placeholder_idx)
        == (reference.placeholder_type, reference.placeholder_idx)
        else 0.0
    )


def _geometry_score(candidate: Candidate, reference: Candidate) -> float | None:
    if reference.position is None or reference.size is None:
        return None  # signal not applicable: reference carries no geometry
    if candidate.position is None or candidate.size is None:
        return 0.0
    dx = candidate.position[0] - reference.position[0]
    dy = candidate.position[1] - reference.position[1]
    pos_score = max(0.0, 1.0 - ((dx**2 + dy**2) ** 0.5) / _POSITION_TOLERANCE_EMU)
    dw = candidate.size[0] - reference.size[0]
    dh = candidate.size[1] - reference.size[1]
    size_score = max(0.0, 1.0 - ((dw**2 + dh**2) ** 0.5) / _SIZE_TOLERANCE_EMU)
    return (pos_score + size_score) / 2


def _shape_type_score(candidate: Candidate, reference: Candidate) -> float:
    return 1.0 if candidate.shape_type == reference.shape_type else 0.0


def _content_pattern_score(candidate: Candidate, reference: Candidate) -> float:
    # Weakest signal, last resort per spec. discover() only captures whether a
    # shape has text, not the text itself, so this degrades to a has-text
    # match rather than a real pattern comparison (e.g. "looks like a date").
    return 1.0 if candidate.has_text == reference.has_text else 0.0


def score_candidate(candidate: Candidate, reference: Candidate) -> float:
    """Combine every applicable scored signal into one confidence score in
    [0, 1], in specs/matching.md's reliability order. A signal that isn't
    applicable to this reference (e.g. the reference isn't a placeholder, or
    carries no geometry) is excluded and the remaining weights are
    renormalized -- never padded with a fabricated value."""
    signals = [
        (_PLACEHOLDER_WEIGHT, _placeholder_score(candidate, reference)),
        (_GEOMETRY_WEIGHT, _geometry_score(candidate, reference)),
        (_SHAPE_TYPE_WEIGHT, _shape_type_score(candidate, reference)),
        (_CONTENT_WEIGHT, _content_pattern_score(candidate, reference)),
    ]
    applicable = [(weight, value) for weight, value in signals if value is not None]
    total_weight = sum(weight for weight, _ in applicable)
    if total_weight == 0:
        return 0.0
    return sum(weight * value for weight, value in applicable) / total_weight


def _confidence_for(score: float) -> Confidence:
    if score >= _HIGH_THRESHOLD:
        return Confidence.HIGH
    if score >= _MEDIUM_THRESHOLD:
        return Confidence.MEDIUM
    return Confidence.LOW


def match(
    candidates: Sequence[Candidate],
    reference: Candidate,
    valid_tags: Container[str] | None = None,
) -> MatchResult:
    """Match `reference`'s field against the best candidate in `candidates`.

    Two-tier per spec:
    - Tier 1 (trust, no scoring): any candidate whose identity_tag is already
      set is trusted directly, provided it's "valid" -- contained in
      `valid_tags` if that's given, or simply non-None otherwise. Exactly one
      trusted candidate is an immediate high-confidence match; more than one
      is a same-tag collision, which can't be silently resolved.
    - Tier 2 (scored fallback): every untagged candidate is scored against
      `reference` via score_candidate(). The top scorer is accepted only if
      it clears the high-confidence threshold AND isn't ambiguously close to
      other similarly-scored siblings (sibling_ambiguity: closeness itself is
      signal, not noise to break arbitrarily -- z-order is tried as a
      supplementary disambiguator before giving up and flagging). Medium
      confidence is always flagged for human confirmation; low confidence is
      always unmatched.
    """
    tagged = [
        c
        for c in candidates
        if c.identity_tag is not None and (valid_tags is None or c.identity_tag in valid_tags)
    ]
    if len(tagged) == 1:
        return MatchResult(
            candidate=tagged[0], confidence=Confidence.HIGH, score=None, reason="existing identity tag"
        )
    if len(tagged) > 1:
        return MatchResult(
            candidate=None,
            confidence=Confidence.MEDIUM,
            score=None,
            reason=f"{len(tagged)} candidates already carry this identity tag -- collision",
        )

    if not candidates:
        return MatchResult(candidate=None, confidence=Confidence.LOW, score=None, reason="no candidates to match against")

    scored = sorted(
        ((c, score_candidate(c, reference)) for c in candidates),
        key=lambda pair: pair[1],
        reverse=True,
    )
    best_candidate, best_score = scored[0]
    confidence = _confidence_for(best_score)

    if confidence is Confidence.HIGH:
        tied = [(c, s) for c, s in scored if (best_score - s) < _SIBLING_GAP_THRESHOLD]
        if len(tied) > 1:
            z_distances = [(c, abs(c.z_order - reference.z_order)) for c, _ in tied]
            min_z = min(dist for _, dist in z_distances)
            winners = [c for c, dist in z_distances if dist == min_z]
            if len(winners) == 1:
                best_candidate = winners[0]
                best_score = next(s for c, s in tied if c is best_candidate)
                confidence = _confidence_for(best_score)
            else:
                return MatchResult(
                    candidate=None,
                    confidence=Confidence.MEDIUM,
                    score=best_score,
                    reason=(
                        f"sibling ambiguity: {len(tied)} candidates score within "
                        f"{_SIBLING_GAP_THRESHOLD} of each other and z-order doesn't disambiguate"
                    ),
                )

    if confidence is Confidence.HIGH:
        return MatchResult(candidate=best_candidate, confidence=confidence, score=best_score, reason="scored match")
    if confidence is Confidence.MEDIUM:
        return MatchResult(
            candidate=None, confidence=confidence, score=best_score, reason="medium confidence -- flagged for human confirmation"
        )
    return MatchResult(candidate=None, confidence=Confidence.LOW, score=best_score, reason="low confidence -- unmatched")
