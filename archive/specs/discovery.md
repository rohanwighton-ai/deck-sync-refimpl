# Discovery

Given a slide's shape tree, find every shape that is a candidate dynamic field.

## Requirements

- Type-agnostic: a candidate can be a dedicated text box, text typed directly into an
  autoshape or placeholder, or a picture. Never filter by shape type — check for the
  actual signal (non-empty text content, or being an image).
- Recurse into groups: a group is not a leaf shape and must never be treated as one
  opaque candidate. Walk into every group's members, recursively (groups can nest), and
  evaluate each leaf shape independently.
- A shape with no text content and that is not a picture is not a candidate, however
  deeply nested or however plausible-looking its shape type. Pure decoration must be
  correctly excluded, not force-matched.
- Preserve enough structural metadata per discovered candidate to support later matching
  and verification: shape name (for humans, not identity), group path (the chain of
  group names it's nested inside, empty if top-level), z-order (document position within
  its containing tree), shape type, and whether it carries a placeholder type/index.

## Non-goals (out of scope for this spec)

- Assigning meaning to a discovered field (which one is "status", which is "title") —
  that's a human confirmation step, not something discovery infers.
- Writing any identity tag — discovery only finds and describes candidates, it does not
  mutate the source file.

## Reference

Existing (VBA-target) design: this project's understanding of the underlying skill's
`shape-identity-and-matching.md` `discovery_scope` section — reproduced here in
language-agnostic form for a Python reference implementation. Test fixtures live in
`test-fixtures/` at the project root.
