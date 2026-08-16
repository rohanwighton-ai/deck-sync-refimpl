# Verification

Given a tagged field shape and its linked data-source value, prove the link resolves
correctly — never assume a tag-and-seed pairing that merely looks consistent is actually
wired correctly.

## Requirements

- The core operation ("inject_primitive"): hash the shape's current value and the linked
  source value; if they match, no-op (write nothing); if they differ, write the source
  value into the shape, then re-hash to confirm the write actually took. Never assume a
  write succeeded without checking.
- A no-op result immediately after tagging is evidence the link is correct — the seed
  value was harvested from the shape itself, so a genuine link resolves with zero visible
  change. A mismatch at that point means the tag resolved to the wrong source, not that
  the sync "corrected" something.
- Structural verification after any duplication (not just value verification): shape
  count, type, and identity-tag correspondence between a duplicate and its source must
  all be checked, not assumed from the duplication API succeeding.
- Z-order (stacking order) must be checked separately from value/tag correspondence after
  duplication. A duplicate can have the right shapes, right tags, and right values, while
  a stacking-order change makes an overlaid field invisible (e.g. a transparent text box
  ending up behind its background shape). Value-correctness and visual-correctness are
  different claims — verifying one does not verify the other.

## Non-goals

- Deciding what to do about a failed verification (this spec defines detection, not
  remediation policy).

## Reference

Language-agnostic reproduction of the underlying skill's `shape-identity-and-matching.md`
`inject_primitive` and `verification_checkpoints` sections.
