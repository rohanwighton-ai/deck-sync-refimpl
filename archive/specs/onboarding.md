# Onboarding

Given an established type's template (an already-tagged reference slide) and a new,
untagged slide of that same type, match its candidate shapes against the template's field
shapes and either accept the match or flag it for a human — never guess a mapping onto a
drifted deck.

## Requirements

- **First-time onboarding of a type needs no scoring at all** — per the underlying skill's
  own design, "the working copy IS becoming the reference." A first example slide is
  discovered, confirmed with a human, and tagged directly (`identity_tags.upsert_slide_tags`/
  `upsert_shape_tags`); its verification is exactly `inject_primitive` hitting the no-op path
  on every field, since the seed value came from the shape's own harvested content. This is
  already fully covered by existing primitives (discovery, identity_tags, verification) —
  nothing new to build for this case.
- **Boilerplate-vs-varying pre-filter, when 2+ example slides are supplied** (the
  underlying skill's `onboard-slide-type.md` already allows 1-2 examples, and
  `deck-adoption.md`'s bulk case routinely has many more). A richly-designed real slide
  can discover 60-90 raw candidates (confirmed against a real 46-slide deck — see
  `test-fixtures/SOURCE.md`'s `crc-real-deck-redacted.pptx` entry), the overwhelming
  majority of which are static design chrome (section headers, icon labels, legend
  text) that happen to contain text, not real dynamic fields — a human cannot review
  that many shapes at the Step 4 confirm-with-user phase gate. Before presenting that
  review: cluster corresponding shapes across the supplied instances using the same
  geometry/placeholder-type signals `matching.md`'s tier-2 scoring already computes (no
  new geometry logic), and partition candidates into **varies across instances**
  (shown first, expanded — the likely real fields) vs. **identical across every
  instance** (collapsed by default, never discarded — still fully visible and
  includable with one action). With only one example slide supplied, there is no
  second instance to diff against, so this filter cannot apply and the full unfiltered
  candidate list is shown exactly as before — this is a pre-filter for the review
  step, not a change to what `discover()` itself returns.
- **Matching a subsequent slide against an established template** is the actual gap:
  discover the new slide's candidates, and for each field role the template defines, score
  every untagged candidate against the template's reference shape per specs/matching.md's
  tier-2 path.
- **Confidence thresholds dispatch, per specs/matching.md**: high confidence auto-accepts
  and writes the tag immediately (self-healing — becomes a tier-1 fast match next time,
  same as the underlying skill's `confidence_thresholds` describes); medium confidence
  produces a match result but is never auto-tagged — a human decides; low confidence is
  unmatched and never forced.
- **Confirming an unresolved match is a distinct, explicit action**, not a side effect of
  scoring — given a specific shape and role a human has decided on (however they decided —
  re-running the matcher, or picking the shape directly), write the tag. This is the
  primitive an eventual object-selection UI (a human selects the shape in the deck, tells
  the tool which field it is) would call once built — that selection mechanism itself is
  out of scope here (see Non-goals), but the primitive it needs already exists as this
  spec's confirmation step.
- The slide-level identity (`slide_type`, `instance_key`) for a new instance is supplied by
  whatever created it (e.g. sync-operations' case 3/2 duplication), not matched — tagging
  it is unconditional, separate from the per-field scoring loop.

## Non-goals

- **The selection UI/mechanism itself** (a human clicking a shape in PowerPoint,
  `Application.ActiveWindow.Selection.ShapeRange` in VBA) — out of scope for this Python
  reference implementation, which has no real deck-editing UI to select from. This spec
  builds the primitive the UI would call (confirm a role onto a specific shape), not the
  UI.
- **Deciding where a type's template itself is physically stored** (a dedicated part? a
  designated slide within the deck?) — this spec takes a template as an already-resolved
  `SlideInstance` (e.g. via `resolve.py`'s tier-1 composition), the same boundary
  `sync_operations.py` already draws for its own inputs.
- Reconciling a template that has itself drifted (its own tags lost/corrupted) — that's
  onboarding *of the template*, not matching *against* one; out of scope here.

## Reference

Language-agnostic reproduction of the underlying skill's `onboard-slide-type.md` (Step 2's
untagged_fallback note: "if there's already a partial reference to compare against") and
`shape-identity-and-matching.md`'s `matching_tiers`/`confidence_thresholds` sections,
restricted to the template-matching case — first-time onboarding needs no new logic here,
per the Requirements section above.

The boilerplate-vs-varying pre-filter reuses `matching.md`'s tier-2 geometry/placeholder
scoring for shape correspondence across instances; it's a review-time filter only, not a
change to `discovery.md`'s candidate rule. Motivating evidence:
`test-fixtures/SOURCE.md`'s `crc-real-deck-redacted.pptx` entry (2026-07-25).
