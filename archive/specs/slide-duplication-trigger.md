# Slide Duplication & Trigger Semantics

`sync-operations.md` decides *that* a duplicate is needed for cases 2 (period_rollover)
and 3 (new_record); this spec governs what happens the instant that decision becomes a
real `Slide.Duplicate` call — where the result lands in the deck, what gets checked before
it's trusted, and how it behaves when many duplications happen in one pass. Closes the gap
both `vba-port.md` ("trigger semantics... aren't decided") and `sync-operations.md`
("Physical slide duplication... Separate spec, pending decision") name explicitly as
out of scope for themselves.

## Requirements

- **One primitive, two callers.** Both cases terminate in the same call —
  `DuplicateAndTag(sourceSlide, newInstanceKey)` — but only `sync_operations` decides
  *whether* to invoke it, per its own case 2/3 dispatch rules. This spec does not
  re-decide when duplication is warranted, only what the call itself does once triggered.
- **Deck order is governed by Data-sheet row order, not by `Slide.Duplicate`'s own
  default placement** (which inserts immediately after the source). A type's slides in
  the deck must reflect the current row order of that type's rows in the Data sheet.
- **Order is an enforced invariant, not a one-time stamp at creation.** Every routine
  sync run reconciles deck order against current Data-sheet row order for the type being
  synced, moving already-existing slides (`Slide.MoveTo`) if a human has reordered rows
  since the last sync — not just placing newly-created slides correctly and leaving prior
  ones wherever they already sit. Chosen over insert-at-creation-only because a resorted
  sheet silently drifting from deck order with no flag would contradict the
  never-silently-diverge posture `sync-operations.md` already holds for case 7 — a
  resequence step makes the drift resolve automatically instead of accumulating unnoticed.
- **Bulk case-3 ordering.** When routine sync finds multiple unmatched rows in one pass
  (e.g. generating a full deck from a freshly-populated Data sheet), process them in
  Data-sheet row order, one full duplicate-tag-inject-verify cycle per row — never
  duplicate all of them first and inject afterward. Keeps `sync-operations.md`'s
  per-decision traceability requirement intact per-slide, not just per-batch.
- **Post-duplication structural verification is mandatory before tagging.** Immediately
  after every `Slide.Duplicate` call, run `verification.md`'s `verify_structure` and
  `verify_z_order` against the duplicate before writing any tag or injecting any value.
  A malformed duplicate must never receive an `instance_key`.
- **Instance-key collision guard.** Before duplicating, check that no existing slide
  already carries the `instance_key` about to be assigned. Two Data-sheet rows sharing a
  key by mistake (bad paste, copy-drag error) must refuse and flag the row — same
  never-silently-guess posture as case 6 — rather than silently producing two slides for
  one key.
- **Partial-row handling.** A row missing a value for one of the type's fields still gets
  a slide created — creation is never withheld pending row "completeness," since no
  completeness rule exists anywhere in the spec chain. Injection is skipped only for the
  missing field, and the row is flagged for visibility (same treatment as case 6's
  unclassified-slide flag) rather than silently leaving the field blank with no signal.

## Non-goals

- Case 5 (`record_retired`)'s archive-move mechanics — separately deferred, no agreed
  convention yet (`sync-operations.md`'s own Non-goals).
- The "add new quarterly record" command's actual invocation/UI mechanism — already
  deferred in `sync-operations.md`.
- Picture-field injection semantics (what a Data-sheet cell means for an image field) —
  an `inject_primitive`/verification question, not a duplication-placement one; separate
  spec if it becomes real.
- Multi-deck concerns — exporting slides to a separate file (tagged or untagged), a
  workbook having more than one connected deck, cross-deck conflict resolution, and
  propagating a resolved value to other connected decks. These were scoped in design
  discussion alongside this spec but depend on capabilities this project hasn't built or
  proven yet (deck-workbook pairing is still 1:1; no deck has been synced against real
  Office once). Deliberately not specified here — revisit once onboarding (port-order
  step 5) and Excel-side I/O (step 6) are ported and this has run against a real deck.

## Reference

- Governs the "how" for `sync-operations.md`'s cases 2/3 duplication decisions; does not
  re-derive when duplication is warranted.
- Reuses `verification.md`'s `verify_structure`/`verify_z_order` unchanged — no new
  verification logic, just a mandatory call site for existing checks.
- Closes the gap named in `vba-port.md`'s Non-goals ("trigger semantics... aren't
  decided") and `sync-operations.md`'s Non-goals ("Physical slide duplication...
  Separate spec, pending decision").
