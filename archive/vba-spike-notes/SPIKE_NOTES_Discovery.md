# VBA port: `discovery` module

Module 1 of `specs/vba-port.md`'s port order. Ports `src/discovery.py`'s
`discover()` (see `specs/discovery.md`) to native VBA, walking a real open
presentation's `Shapes`/`GroupShapes` collections instead of parsing raw
`<p:spTree>` OOXML.

**Not executed or verified in this environment** -- there is no
Windows/Office install here, same constraint as `InjectPrimitive.bas`.
`Discovery.bas` has not been run. The manual verification recipe below is
how to actually prove it against a real Office install, cross-checked
against this project's own `test-fixtures/*.pptx` and the exact values
`tests/test_discovery.py` already asserts for the Python side.

## What was ported

`discover(slide_xml_root)`'s walk, mirrored field-for-field onto a new
`Candidate` VBA `Type`:

1. Recurse into `msoGroup` shapes via `GroupItems`, never treating a group
   as one opaque candidate (matches `discovery.md`'s explicit recursion
   requirement).
2. For every non-group shape PowerPoint's object model can carry a text
   frame or image on (see "Scope note" below), capture: name, group path,
   z-order, shape type, placeholder type, has-text, position, size.
3. Leave `IdentityTag` always `""`, matching `discovery.py`'s own non-goal
   ("discovery only finds and describes candidates, it does not read or
   write identity tags") -- reading tags is `identity_tags`' module
   boundary (already ported), not this one's.

## Where the VBA port genuinely simplifies things (not a corner cut)

No XML parsing at all. `Shape.Type`, `Shape.GroupItems`, `Shape.Left/Top/
Width/Height`, and `Shape.TextFrame.TextRange.Text` are all native COM
properties PowerPoint already computes; the Python side has to hand-parse
`<p:grpSp>`/`<p:sp>`/`<p:pic>` and `<a:xfrm>` because it has no host
application to ask.

## Scope note: this reproduces discovery.py's real behavior, not just discovery.md's language

`discovery.md` says "type-agnostic... never filter by shape type." Taken
literally that would mean *every* shape type is a candidate if it has text
or is a picture. But `discovery.py`'s actual `walk()` only ever recognizes
three OOXML tag names: `<p:grpSp>` (recurse), `<p:sp>`, `<p:pic>` (leaves).
A `<p:graphicFrame>` (table, chart, embedded object) or `<p:cxnSp>`
(connector) never matches either leaf tag, so the Python implementation is
*already* silently blind to those shapes regardless of their content --
that's a real, existing scope boundary of the code being ported, not
something invented here.

`IsCandidateLeafType` in `Discovery.bas` reproduces that same boundary:
`msoAutoShape`, `msoTextBox`, `msoPlaceholder`, `msoFreeform` map to
`<p:sp>`; `msoPicture`/`msoLinkedPicture` map to `<p:pic>`; anything else
(`msoChart`, `msoTable`/`msoEmbeddedOLEObject`/`msoLinkedOLEObject`,
`msoLine`/`msoConnector`, `msoComment`, `msoDiagram`, ...) is skipped
entirely, exactly as the Python walk never reaches them. If a future
fixture needs one of those shape kinds treated as a candidate, that is a
`discovery.py` scope change first (this port would follow it, not lead it).

## Deliberate divergences from the Python semantics

1. **`PlaceholderIdx` is always `-1` -- not ported, not invented.** OOXML's
   `<p:ph idx="...">` is the *most reliable* signal in `matching.md`'s
   scoring order (weight 0.5), and `discover_from_pptx_layout` on
   `mst-slide-layouts.pptx` proves Python reads it correctly (e.g. "Text
   Placeholder 3" on layout 1 has `idx=10`, per
   `tests/test_discovery.py::test_mst_slide_layouts_captures_placeholder_type_and_idx`).
   But PowerPoint's object model has no `Shape.PlaceholderFormat.Idx` (or
   equivalent) -- `PlaceholderFormat.Type` only gives the *category*
   (title/body/...), not the numeric slot that disambiguates two
   same-category placeholders on one layout. Unlike `Shape.Tags`
   (`InjectPrimitive.bas`'s spike already confirmed COM implements the
   exact OOXML mechanism natively), there is no native path here at all --
   getting `idx` genuinely requires falling back to raw OOXML per
   `specs/vba-port.md`'s explicit-fallback instruction (e.g.
   `Shell`-invoking Python, or reading the slide/layout part's XML via a
   zip library). That fallback is *not* built in this module -- flagging it
   here as the real, unresolved gap it is, left for whichever module
   actually needs `idx` first (almost certainly `matching`, port-order
   step 3, since it is `matching.md`'s top-weighted signal).

2. **`PlaceholderType` is a best-effort label, not authoritative.**
   `PlaceholderTypeLabel()` maps `ppPlaceholderType` enum values to
   OOXML-style strings for the common cases (title/ctrTitle/subTitle/body/
   obj/chart/tbl/media/clipArt/dt/sldNum/hdr/ftr) and falls through to
   `"unknown"` for enum values with no clean OOXML string this module
   confirmed (e.g. vertical-text variants). This is a many-to-one
   approximation of a one-to-one OOXML attribute; do not treat a matched
   string here as proof the underlying `<p:ph type="...">` attribute is
   literally that value the way the Python read is.

3. **Position/size are points converted to EMU, not EMU read directly --
   and the round-trip is not always bit-exact.** `Shape.Left/Top/Width/
   Height` are `Single` (single-precision float) in points; `PointsToEmu`
   multiplies by 12700 (1 pt = 12700 EMU) to match `discovery.py`'s EMU
   convention. Sizes that are whole multiples of an inch round-trip exactly
   (e.g. `mst-slide-layouts.pptx` layout 1's "Title 1" is `cx=8229600`,
   `cy=1143000` EMU -- both exact multiples of 12700). Offsets frequently
   are **not** clean multiples of 12700 in the source XML (e.g. that same
   shape's `x=467544` EMU is `36.81...` pt) -- expect the VBA-read value to
   be within a handful of EMU of Python's, not necessarily identical, once
   it passes through a `Single` and back.

4. **Nested-shape position may already be slide-absolute in VBA, unlike
   Python's raw local offset.** `discovery.py`'s `_geometry()` reads a
   shape's own local `<p:spPr>/<a:xfrm>` verbatim without walking up
   through parent group transforms (documented in `discovery.py` and
   `AGENTS.md`'s Notes as exact only when a group's `chOff`/`chExt` equals
   its own `off`/`ext`). PowerPoint's object model is reported to resolve
   `Shape.Left`/`Top` for a shape inside a group to slide-absolute
   coordinates in common cases -- if true, that would make this port
   *more* correct than the Python original for nested shapes, not less.
   **Unconfirmed** -- there is no Office install here to verify it, and
   real-world behavior for rotated/flipped groups is known to be
   inconsistent. Treat position values captured for grouped shapes (e.g.
   `shp-groupshape.pptx`'s three shapes inside "Group 4") as unverified
   until the manual recipe below is actually run.

5. **`GroupPath` is a single `"/"`-joined string, not a tuple.** VBA has no
   convenient tuple type for a UDT field; `Candidate.GroupPath` joins the
   same group-name chain `discovery.py`'s `group_path` tuple carries.
   `shp-groupshape.pptx`'s three nested shapes should read `"Group 4"` here
   (single level of nesting in that fixture), matching Python's
   `("Group 4",)`.

## What was deliberately left out of scope

- `identity_tag` reading -- always `""`, per discovery's own module
  boundary (see "What was ported" above). Reading `Shape.Tags` is
  `InjectPrimitive.bas`'s territory (already ported).
- No matching, scoring, or sync-dispatch -- this module only finds and
  describes candidates, exactly like its Python counterpart's non-goals.
- No automated test harness, for the same reason `InjectPrimitive.bas` has
  none: no VBA unit-test framework wired up in this project, and no
  Office/COM available in this sandbox to run one against.

## Manual verification recipe

Run from the VBA IDE (Alt+F11) with the Immediate window open (Ctrl+G).

### 1. `shp-groupshape.pptx` -- group recursion, zero candidate fields

1. Open `test-fixtures/shp-groupshape.pptx` directly in PowerPoint.
2. Import `Discovery.bas` (File > Import File... in the VBA IDE).
3. Run `ManualSmokeTest`.
4. Expected, per `tests/test_discovery.py`'s already-proven Python values:
   - 4 candidates total (`test_shp_groupshape_finds_all_four_leaf_shapes`).
   - 1 with `GroupPath = ""`, 3 with `GroupPath = "Group 4"`
     (`test_shp_groupshape_recurses_into_the_group_not_opaque`).
   - Every one has `HasText = False` and `ShapeType = "autoshape_or_textbox"`
     (none are pictures) -- i.e. zero candidate *fields* even though 4 shapes
     are discovered, matching `test_shp_groupshape_finds_zero_candidate_fields`'s
     point that pure decoration must be found but not force-matched as a field.
   - Every one has `HasPlaceholder = False`
     (`test_shp_groupshape_shapes_have_no_placeholder`).
   - The shape named "Oval 2": Python reads `position=(5940152, 2708920)`,
     `size=(914400, 914400)` EMU exactly
     (`test_shp_groupshape_captures_position_and_size`) -- confirm the VBA
     values are close (within a few EMU/points, per divergence #3 above;
     `size` here is a clean 1in x 1in square so it should round-trip exact).

### 2. `mst-slide-layouts.pptx` -- placeholder type/idx, two layouts

This fixture has no `Slides` at all (only slide layouts + a slide master),
per `test_mst_slide_layouts_has_no_ppt_slides_entries`. Open it in
PowerPoint and use `ActivePresentation.Designs(1).SlideMaster.CustomLayouts`
to reach its layouts (or, from any presentation, temporarily apply one of
its layouts to a slide and read `Slide.CustomLayout`).

1. In the Immediate window, run:
   `? Discovery.DiscoverCustomLayout(Application.ActivePresentation.Designs(1).SlideMaster.CustomLayouts(1)).UBound`
   (or step through `DiscoverCustomLayout` with a watch on `results`).
2. Expected, per Python's already-proven values:
   - A shape named "Title 1" with `HasPlaceholder = True`,
     `PlaceholderType = "title"` (Python: `placeholder_type == "title"`,
     `placeholder_idx == 0`). **`PlaceholderIdx` will read `-1` here per
     divergence #1** -- confirm this is the expected gap, not treat it as a
     bug.
   - A shape named "Text Placeholder 3" with `PlaceholderType = "body"`
     (Python: `placeholder_type == "body"`, `placeholder_idx == 10`). Same
     `PlaceholderIdx = -1` caveat.
   - "Title 1"'s position/size should read close to `(467544, 2060848)` /
     `(8229600, 1143000)` EMU (the size should be exact; the offset may
     drift slightly per divergence #3).
3. Repeat against `CustomLayouts(2)` and confirm it returns a non-empty
   candidate list too (Python:
   `test_mst_slide_layouts_layout2_also_discoverable`), proving the loader
   isn't hardcoded to the first layout.

### 3. Close the loop on divergence #4 (grouped-shape position)

While `shp-groupshape.pptx` is open, manually compare one of "Group 4"'s
child shapes' VBA-reported `Shape.Left`/`Top` (converted to EMU) against
its raw `<a:off>` in `slide1.xml` (unzip the `.pptx` and inspect the XML
directly, the same way this project's own `tests/test_discovery.py` header
comment says its expected values were originally hand-proven). If they
differ by more than the rounding noise in divergence #3, VBA is resolving
group-relative coordinates to slide-absolute ones and this note should be
upgraded from "unconfirmed" to a confirmed, permanent divergence (in VBA's
favor) in a follow-up edit to this file.
