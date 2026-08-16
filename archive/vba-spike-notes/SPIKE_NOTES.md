# VBA spike: `inject_primitive` port

De-risks this project's core assumption -- "Python proves the logic, VBA
just inherits it" -- by porting the narrowest real slice: writing one
primitive value into one tagged shape. Not a port of discovery, matching,
sync dispatch, or duplication. Agreed as outstanding work in the
2026-07-23 roadmap-correction note (project_active_ventures memory).

**Not executed or verified in this environment** -- there is no
Windows/Office install here. `InjectPrimitive.bas` has not been run. The
manual verification recipe below is how to actually prove it on a real
machine.

## What was ported

`src/verification.py`'s `inject_primitive(path, part_name, shape,
source_value)`:
1. Locate the shape.
2. If its current value already equals the source value, no-op.
3. Otherwise write the value, then re-read it back and confirm the
   write actually took (never assumed from the write call alone).
4. Error if the shape has no text to write into.

`InjectPrimitive` in `InjectPrimitive.bas` preserves all four steps, plus
the tag-based lookup discipline from `src/identity_tags.py` (find the shape
by its hidden identity tag, not its visible name -- the project's own
real-deck validation found 5 shapes on one real slide all literally named
"Rounded Rectangle 1").

## Where the VBA port genuinely simplifies things (not a corner cut)

The Python side (`identity_tags.py`) had to hand-roll the OOXML mechanism
behind PowerPoint's `Shape.Tags` -- reverse-engineering the
`custDataLst`/Tags-Part/relationship chain from ECMA-376, because
python-pptx has no support for it and no fixture on disk carried tags to
reverse-engineer from. VBA doesn't need any of that: `Shape.Tags` is a
native COM object PowerPoint already implements on top of that exact same
XML mechanism. So `FindShapeByRoleTag` is a plain `For Each shp In
sld.Shapes` loop checking `shp.Tags("role")` -- there is no XML surgery
in this file at all. This is the clearest confirmation so far that "Python
proves the logic, VBA inherits it" holds: the logic (find-by-tag,
compare-before-write, verify-after-write) is identical; only the mechanism
for reading/writing tags collapses from ~250 lines of zip/XML handling to
one object-model property.

## Deliberate divergences from the Python semantics

1. **No SHA-256 hashing.** `inject_primitive` hashes values before
   comparing (`_hash`). VBA holds the actual strings live in memory via
   the object model already -- there's no large-file or streaming reason
   to hash, so this port does direct string equality. Same rigor (exact
   match), no crypto library needed or available in stock VBA.

2. **"Verified" is a weaker guarantee here than in Python.** The Python
   version re-reads the written value by **re-opening the .pptx zip from
   disk** (`_read_part` again, after `write_zip_parts`) -- so `verified`
   proves the write persisted to the actual OOXML part. This VBA port
   re-reads the value from the **same live Shape object** it just wrote
   to, which only proves the in-memory object-model state changed, not
   that a `.Save` would actually persist it correctly. This is the one
   real gap in the port and the reason the manual verification recipe
   below includes closing and reopening the file, not just running the
   macro once.

3. **`Shape.Tags("role")` can't distinguish "no such tag" from "tag
   present but empty string."** Real VBA API behavior: the accessor
   returns `""` in both cases. The Python side's tag storage is a real
   `dict`, so `identity_tags.py`'s `read_shape_tags` returns a genuine
   `{}` when absent. Since an empty-string role is not a meaningful
   identity tag in this project's scheme, `ShapeHasRoleTag` treats both
   cases as "not tagged," which is safe here -- just noting the API
   limitation so it isn't mistaken for an oversight later if this spike
   grows into a full port (a full port would need `Tags.Count`/`Tags.Name(i)`
   iteration to tell the two cases apart, same technique
   `FindShapeByRoleTag` already uses when counting matches).

4. **Multiple-match handling.** If more than one shape on the slide
   carries the same `role` tag value, `FindShapeByRoleTag` returns
   `Nothing` (treated as not-found) rather than picking the first match.
   The Python side doesn't hit this case directly (it resolves a specific
   `Candidate` upstream via `matching.py`), but given the project's stated
   philosophy throughout ("prove the link, don't assume a tag-and-seed
   pairing that merely looks consistent is actually wired up right" --
   verification.py's own module docstring), refusing to guess on an
   ambiguous tag is the same posture applied consistently, not a new one
   invented for this port.

## What was deliberately left out of scope

- No discovery, matching, or sync-dispatch port -- this file only writes
  a value into an already-identified shape.
- No physical slide duplication -- already cut from this project's scope
  entirely (2026-07-23 roadmap correction: VBA's native `Slide.Duplicate`
  solves that mechanic directly, nothing to prove there).
- No structural/z-order verification port (`verify_structure`,
  `verify_z_order`) -- those operate across two shape collections
  (source vs. duplicate); this spike only touches one shape on one slide.
- No real VBA unit-test harness. `ManualSmokeTest` is a smoke test you run
  from the VBA IDE, not an automated suite.

## Manual verification recipe (run on a real Windows/PowerPoint machine)

1. Open (or create) a `.pptx` with at least one slide containing a text
   shape (a plain text box is fine).
2. Tag that shape: in the VBA Immediate Window,
   `ActivePresentation.Slides(1).Shapes(1).Tags.Add "role", "demo_field"`.
   (Pick the actual shape index/name for your text box if it isn't
   `Shapes(1)`.)
3. Import `InjectPrimitive.bas` into the presentation's VBA project
   (VBA IDE -> File -> Import File...).
4. Run `ManualSmokeTest` (F5, or Alt+F8 -> `ManualSmokeTest` -> Run).
5. Confirm the message box reports `Found=True Written=True Verified=True`
   and that the text box's visible text actually changed to
   "hello from VBA <timestamp>".
6. Run `ManualSmokeTest` again immediately. Confirm it now reports
   `Written=False` (since the value already matches) -- proves the no-op
   path, not just the write path.
7. **Close persistence check** (covers divergence #2 above, which
   in-session re-reads can't catch): `Ctrl+S` to save, fully close
   PowerPoint, reopen the file, and manually confirm the text box still
   shows the last-written value. This is the step that actually proves
   the write reached the underlying OOXML part, not just the live COM
   object.
8. Confirm no other shape on the slide moved, changed formatting, or lost
   text -- `InjectPrimitive` should touch only the tagged shape.
9. Edge case: add a second shape with the same `role="demo_field"` tag
   and re-run. Confirm it now reports `Found=False` with the ambiguous-tag
   error message, rather than silently picking one of the two shapes.
