# Sync Operations

Given a type's known slide instances and its linked Data-sheet rows, decide which of the
sync event cases applies to each row/instance and what should happen — never guessing a
fit into a case the input doesn't actually satisfy.

## Requirements

- Two distinct entry points, never conflated: **routine sync** dispatches cases 1
  (no_change), 3 (new_record), and 4 (in_place_correction) automatically; **period
  rollover** (case 2) only ever runs from an explicit, separately-invoked command, never
  inferred from routine sync just because a value looks different.
- Routine sync, per Data-sheet row:
  - No known instance carries this row's persistent instance key → **case 3
    (new_record)**: a new slide instance is needed. This module decides that fact and
    what values to inject once the instance exists; it does not perform the physical
    slide duplication itself (see Non-goals).
  - A known instance carries this row's instance key → run the value-verification
    primitive (`inject_primitive`) per field and use its own result as the classifier:
    a no-op result (hash already matches) is **case 1 (no_change)**; a write result (hash
    differed, value written and verified) is **case 4 (in_place_correction)**. No
    separate diff step — the write-verification primitive already had to compute this.
- Case 2 (period_rollover) is a distinct operation, not a variant dispatch of routine
  sync: given an explicit command naming a specific instance key, locate its current
  slide, decide that a duplicate is needed as the new period's slide (leaving the
  original untouched as history), and decide what to inject onto the duplicate from the
  new period's Data-sheet row. Never triggered by a value merely differing — only by the
  explicit command.
- Case 6 (unclassified_slide) detection: a slide's type is an explicit declaration (a
  tag), never inferred from shape-set similarity. A slide with no recognized type tag is
  flagged for reclassification, not silently skipped or guessed into the nearest type.
  This makes detection trivial — a missing/unrecognized tag, not a confidence score.
- Every dispatch decision must be traceable to which rule produced it (why this row is
  case 1 vs 3 vs 4, or why a slide was flagged) — a plan a human can audit, not just a
  final action with no rationale attached.

## Non-goals

- **Physical slide duplication.** Deciding a duplicate is needed (cases 2 and 3) is in
  scope; performing it — inserting a new `ppt/slides/slideN.xml` part and wiring
  `presentation.xml`/relationships/content-types to match — is not. That is a distinct,
  materially harder OOXML-surgery problem; `verification.py`'s `verify_structure`/
  `verify_z_order` already assume a duplicate exists and only check it, and this module
  keeps that same boundary rather than quietly absorbing the harder problem.
- **Case 5 (record_retired).** The source design has no agreed trigger convention for
  "this record is done" vs. "temporarily unchanged" — building detection now would mean
  inventing a convention the design never settled, not implementing something already
  specified. Left as an open design question.
- **Case 7 (deck_side_conflict).** Detecting this needs a 3-way comparison (current
  on-slide value vs. last-synced value vs. new target value); nothing in this project
  persists a "last-synced value" today (`inject_primitive` only ever compares current vs.
  target, a 2-way check). Needs its own storage decision before it can be built — left as
  an open design question, not guessed at.
- Resolution/remediation UI for anything flagged (cases 5-7, or a low-confidence match) —
  this spec defines dispatch and detection, not what a human does once flagged, same
  boundary `verification.md` already draws for its own checks.
- The "add new quarterly record" command's actual trigger/UI mechanism — out of scope
  here; this spec assumes the command has already been issued for a specific instance key
  and starts from there.

## Reference

Language-agnostic reproduction of the underlying skill's `sync-cases.md` (case taxonomy)
and `run-sync.md` (dispatch process, entry points, success criteria) sections, restricted
to cases 1/3/4/6 per those files' own stated v1 scope.
