# Deck Adoption (Bulk Retroactive Linking)

Every spec so far assumes the Data sheet is the source of truth and slides get created
or corrected to match it (`sync-operations.md`'s cases 1/3/4). This spec covers the
opposite, unaddressed direction: **a real deck that already has N populated slides
representing N real records, with no tags yet, that need to become the ground truth**
— captured verbatim into the Data sheet and tagged — before ongoing sync starts
governing them. `onboarding.md` only harvests values for *one* example slide (the
template); `matching.md`'s "subsequent slide" case only handles one slide at a time,
tested at that granularity. Neither covers "here are the 11 other real slides this
deck already has of this type." `DECISIONS.md`'s 2026-07-19 entry names this
explicitly as an accepted, unvalidated gap ("no real deck has been run through this
yet"); this spec closes it.

This is a **one-time bootstrap operation per type**, not a new sync-operations case —
it runs once (or occasionally, as more historical slides surface) *before* routine
sync takes over, not on every pass.

## Requirements

- **Explicit selection, not whole-deck scanning.** Entry point takes the user's
  current slide multi-selection (`Application.ActiveWindow.Selection.SlideRange`) as
  the scope — same mechanism convention as every other selection-driven step in this
  project. The 2026-07-19 decision's reasoning (never require real deck content
  wholesale by default) still governs the *default* posture; this spec is the
  deliberate, explicit opt-in a user reaches for specifically because they want to
  adopt real content — not a reversal of that default.
- **Two starting conditions, same downstream loop:**
  - **Greenfield** (no template yet for this type): the user picks one slide from the
    selection as the template example. That one slide goes through
    `onboard-slide-type.md`'s existing flow completely unchanged — this spec adds
    nothing to establishing the template itself. Every *other* slide in the selection
    then goes through the per-slide loop below.
  - **Established template**: every slide in the selection goes through the per-slide
    loop below directly.
- **Per-slide loop, run over every slide in scope except the template:**
  - **Idempotent skip.** A slide that already carries `instance_key` + type tag is
    already linked — report it as "already linked," touch nothing. Reruns (adopting
    slides missed on a first pass) must be safe.
  - **Match against the template** using `matching.md`'s existing tier-2 scoring,
    unchanged — no new matching logic, just applied per-slide across the batch instead
    of to a single incoming slide.
  - **Confidence dispatch** (same thresholds as `onboarding.md`): high confidence →
    included in the batch as ready; medium confidence → flagged, needs the existing
    `confirm_field_match` resolution before it can proceed; low/no match → excluded
    from this batch and reported as unclassified, never forced in.
  - **Harvest current values verbatim** from every matched field on the slide — the
    same harvest primitive `onboarding.md` already uses for its one template slide,
    generalized here to run once per slide in the batch.
  - **Instance-key resolution, never silently invented:** if the type's Data sheet
    already has one or more rows with no `instance_key` yet (e.g. a user hand-typed
    rows ahead of running this), compare the slide's harvested values against each
    such row's non-key fields; a slide matching exactly one such row verbatim links to
    it directly. Zero or more-than-one match falls back to creating a fresh row rather
    than guessing which to merge into. For a genuinely new row, the instance_key is
    never auto-generated silently — the human supplies or confirms it at the phase
    gate below (auto-suggested from the type's designated key field's harvested value,
    if one is defined).
- **One phase gate for the whole batch, before any write.** A single review listing
  every slide in scope with its proposed disposition (ready/high-confidence, needs
  confirmation/medium, excluded/unclassified) and the instance_key it would receive.
  Nothing is written until this is confirmed. The user can exclude any individual
  slide even if it scored ready — this is a review, not just a rubber stamp.
  **This is exactly where `onboarding.md`'s boilerplate-vs-varying pre-filter matters
  most** — a bulk-adoption batch routinely has many real instances of the same type to
  diff against (unlike ordinary onboarding's 1-2 examples), so the filter has maximum
  signal here: confirmed against a real 46-slide deck, raw discovery finds 60-90
  candidates per slide, and the phase gate is unusable without collapsing the
  identical-across-every-instance shapes first (see `test-fixtures/SOURCE.md`'s
  `crc-real-deck-redacted.pptx` entry).
- **Row order bootstraps from deck order**, not the other way around: newly created
  rows are appended to the Data sheet in the same order their source slides currently
  appear in the deck. This is the one-time exception to
  `slide-duplication-trigger.md`'s standing invariant (deck order governed by
  Data-sheet row order) — that invariant takes over immediately once this bootstrap
  finishes, reconciling any future reordering as normal.
- **Verify the link, not just the write** — for every slide just linked, run the
  existing `inject_primitive` no-op round-trip (comparing the slide's current values
  against its brand-new row), exactly as `onboarding.md`'s Step 6 already mandates for
  the template slide, generalized across the whole batch. Any field that doesn't hit
  the no-op path is a bug in this pass's harvest, not something a later sync
  "corrects" — stop and flag it the same way onboarding.md treats that condition for
  the template.
- **Reporting** uses the same shape as `ribbon-ui.md`'s shared result form: counts of
  linked / already-linked-skipped / excluded-unclassified / failed-verification, each
  listed by slide name/index.

## Non-goals

- **Auto-classifying a mixed selection spanning multiple untagged types at once.**
  Run once per type (established or being established in the same pass via the
  greenfield path); slides of a different type in the same selection simply score
  low/unmatched and land in the excluded/unclassified bucket, not auto-sorted into
  additional templates.
- **Reconstructing multi-period history between existing slides** (recognizing that
  two of the provided slides are actually the same underlying record at two different
  periods, one archival). Every slide in scope becomes an independent new instance;
  this spec does not infer rollover lineage. A human wanting that relationship
  expressed would need to supply matching key context manually — not asked for here.
- **The UI itself.** `ribbon-ui.md` gets an "Adopt Existing Slides" entry point in a
  future pass, the same relationship it already has to `onboarding.md`/`matching.md`.
  This spec is the engine layer beneath that button, not the button.
- **Automatic whole-deck scanning with no explicit selection** — deliberately not
  built; see the Requirements note on why the 2026-07-19 default posture still stands.
- **Merging a slide's harvested values into more than one plausibly-matching existing
  empty row** — ambiguous cases always fall back to a fresh row rather than guessing.

## Reference

- Extends `onboarding.md` (harvest + match primitives, generalized from "one template
  slide" / "one subsequent slide" to a batch) and reuses `matching.md`'s tier-2
  scoring and `identity-tags.md`'s tag-write primitives unchanged.
- Instance-key semantics match the row-drives-slide direction already established in
  `vba/SyncOperations.bas` (`RowInstanceKey`) — same key meaning, opposite direction
  of derivation.
- Row-order bootstrap rule hands off directly to `slide-duplication-trigger.md`'s
  standing deck-order invariant once adoption finishes.
- Motivating gap: `DECISIONS.md`'s 2026-07-19 onboarding-design entry names this exact
  scenario ("validated against an actual organically-drifted CRC deck") as an accepted,
  still-open uncertainty.
