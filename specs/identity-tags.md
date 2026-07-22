# Identity Tags

Given a shape or slide that needs a persistent, hidden identity (role, type, instance key),
read and write it via PowerPoint's own `Shape.Tags`/`Slide.Tags` mechanism at the OOXML
level — not the visible/user-editable shape name, per input-contract.md's
`unique_named_shapes` rule and shape-identity-and-matching.md's `identity_layer`.

This was an open question through Priority 2-3 (`matching.py`/`verification.py`'s
`identity_tag` field has always been populated by the caller directly, never read from or
written to disk) — python-pptx has no built-in support for it, and no fixture on disk
carried any tags to reverse-engineer from. Resolved by consulting ECMA-376's own "User
Defined Tags Part" definition and a real-world example, not by guessing (see Reference).

## Requirements

- **Slide-level tags** (`slide_type`, `instance_key`, `period_key`): stored in a dedicated
  Tags Part (`ppt/tags/tagN.xml`, root `<p:tagLst>` containing `<p:tag name="" val=""/>`
  children), related to the slide part directly via a relationship of type
  `.../relationships/tags` in the slide's own `.rels` file.
- **Shape-level tags** (`role`): a shape is not its own OOXML part (no `.rels` of its own),
  so it references a Tags Part indirectly — inside the shape's `<p:nvPr>` (the same
  element `discovery.py` already reads for `<p:ph>`), a
  `<p:custDataLst><p:tags r:id="rIdN"/></p:custDataLst>` element whose `r:id` resolves
  through the *owning slide's* `.rels`, landing on the same kind of Tags Part.
- **Schema ordering matters.** `<p:nvPr>`'s children are schema-ordered
  (`ph?, media?, custDataLst?, extLst?`) — inserting `custDataLst` must go after any
  existing `<p:ph>` and before any existing `<p:extLst>`, not just appended blindly.
- **Read-merge-write, no data loss**, same pattern as `excel_output.upsert_row`: setting
  one tag must never drop or overwrite another already-present tag on the same shape or
  slide that wasn't part of this call.
- **Idempotent relationship-graph creation**: writing a tag when no Tags Part/relationship
  exists yet must create exactly one of each (new part number, new relationship ID, new
  `[Content_Types].xml` override) — writing again must reuse the existing ones, never
  create a second Tags Part or a duplicate relationship for the same shape/slide.
- Two shapes on the same slide, tagged independently, must not interfere with each other —
  each gets its own Tags Part and relationship.

## Non-goals

- Deciding *what* role/type/instance-key values to assign — that's matching.py's
  (tier-1 trust) and the onboarding workflow's job. This is read/write plumbing only.
- Verifying that a tag survives a real PowerPoint ungroup/regroup round-trip — flagged as
  an open question in the source design (`shape-identity-and-matching.md`'s
  `discovery_scope`), not something this module can test without a real PowerPoint
  instance.
- General-purpose custom XML data (`<p:custData r:id>`, a sibling mechanism in the same
  `custDataLst` container for arbitrary XML unrelated to tags) — out of scope; this module
  only ever writes/reads the `<p:tags>` variant.

## Reference

Verified against ECMA-376's "User Defined Tags Part" definition (content type
`application/vnd.openxmlformats-officedocument.presentationml.tags+xml`, relationship
type `http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags`,
`CT_TagList`/`p:tag` schema) and a real-world shape-level example
(`<p:custDataLst><p:tags r:id="..."/></p:custDataLst>` resolving to a `<p:tag name=".."
val=".."/>` list) found via `singerla/pptx-automizer` issue #103. `nvPr`'s child ordering
(`ph, media, custDataLst, extLst`) verified against `CT_ApplicationNonVisualDrawingProps`'s
schema definition. Language-agnostic reproduction of `shape-identity-and-matching.md`'s
`identity_layer` and `vba_implementation_notes` sections' storage mechanism, translated
from the VBA object model (`Shape.Tags`/`Slide.Tags`) to the underlying file format.
