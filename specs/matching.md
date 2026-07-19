# Matching

Given a discovered candidate shape and a reference (the example this type was configured
from), determine whether they represent the same field — robust to a deck that has
drifted organically (resized, reformatted, slightly edited) since the reference was
captured.

## Requirements

- Two-tier: if a candidate already carries a valid identity tag from a previous pass,
  trust it directly — no scoring needed. Only fall back to scoring when untagged.
- Scored fallback, in order of reliability: layout placeholder index (if applicable, most
  stable), geometric similarity (position/size within tolerance), shape type, content
  pattern (weakest signal, last resort). Combine into a confidence score — never take the
  single best-scoring candidate blindly.
- Confidence thresholds: high confidence auto-accepts; medium confidence must be flagged
  for human confirmation, never silently guessed; low confidence is unmatched, not forced
  onto the nearest candidate.
- Sibling ambiguity: when multiple untagged candidates score similarly close to each
  other (a real pattern — multiple overlapping/stacked shapes on the same underlying
  element), that closeness IS the signal, not noise to break arbitrarily. Add z-order as
  a supplementary signal when geometry alone can't distinguish siblings, and treat a
  close score gap the same as medium confidence — flag, don't pick.

## Non-goals

- Writing the identity tag once a match is accepted — that's the caller's responsibility
  (part of onboarding/verification, not matching itself).
- Resolving what to do about a low-confidence or flagged match — matching reports
  confidence, it does not decide policy.

## Reference

Language-agnostic reproduction of the underlying skill's `shape-identity-and-matching.md`
`matching_tiers`, `confidence_thresholds`, and `sibling_ambiguity` sections.
