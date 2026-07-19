# Excel Output

Given a set of discovered, confirmed fields for a slide type, produce the structure of
a correct linked data sheet.

## Requirements

- One column per confirmed field, named consistently with the field's identity (not
  regenerated per slide instance — the same logical field must always map to the same
  column, across every instance of a type).
- One row per slide instance, keyed by that instance's persistent identity — never by
  position/order, since row order and slide order can each change independently.
- Seed rows come from harvested values (the content already on the source shape at
  onboarding time), not blank cells the user has to backfill.
- No data loss on write: a column addition or row addition must never silently drop or
  overwrite an existing column/row it wasn't asked to touch.
- The sheet must carry a stable reference back to which specific deck it's paired with —
  never inferred solely from column-name matching against multiple candidate decks.

## Non-goals

- Formatting/styling of the sheet (fonts, colors, conditional formatting) — this spec is
  about structural and data correctness, not presentation.
- The archival-copy behavior for period rollover (that's a slide-side/deck-side concern,
  not an Excel-structure concern) — out of scope until a spec for sync operations exists.

## Reference

Derived from the underlying skill's `input-contract.md` (`unique_named_shapes`,
`no_duplicate_field_names`, `persistent_instance_identity`, `deck_workbook_pairing`
rules) and `onboard-slide-type.md`'s harvest/commit steps, translated into what a correct
output artifact must satisfy rather than the workflow steps that produce it.
