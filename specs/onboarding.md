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
