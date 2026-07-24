# VBA port: `matching` module

Module 3 of `specs/vba-port.md`'s port order. Ports `src/matching.py`'s
`score_candidate()`/`match()` (see `specs/matching.md`) to native VBA,
operating on `Discovery.Candidate` arrays.

**Not executed or verified in this environment** -- there is no
Windows/Office install here, same constraint as `InjectPrimitive.bas` and
`Discovery.bas`. `Matching.bas` has not been run. The manual verification
recipe below is how to actually prove it against a real Office install,
cross-checked against `tests/test_matching.py`'s already-proven Python
values.

## What was ported

`score_candidate(candidate, reference)` and `match(candidates, reference,
valid_tags)`, field-for-field:

1. Tier 1 (trust, no scoring): a single non-empty, valid `IdentityTag` wins
   immediately at high confidence; more than one is a flagged collision.
2. Tier 2 (scored fallback): every candidate scored via `ScoreCandidate`,
   combining placeholder-index / geometry / shape-type / content-has-text
   signals at the same weights as Python (0.5/0.3/0.15/0.05), renormalized
   over only the applicable signals.
3. Confidence thresholds (0.75 high / 0.4 medium) identical to Python.
4. Sibling ambiguity: candidates scoring within 0.1 of the top score are
   tied; z-order (`abs(ZOrder - reference.ZOrder)`) is tried as a
   supplementary disambiguator; if it doesn't uniquely resolve the tie,
   the result is flagged (medium), never arbitrarily picked.

## Where this module closes a real, previously-flagged gap

`SPIKE_NOTES_Discovery.md`'s divergence #1 flagged `PlaceholderIdx` as
always `-1` in `Discovery.bas` -- PowerPoint's object model has no
`Shape.PlaceholderFormat.Idx` equivalent, unlike `Shape.Tags` (which does
have a native equivalent, per `InjectPrimitive.bas`'s spike). That note
predicted this module would need to resolve it via a raw-OOXML fallback,
per `specs/vba-port.md`'s explicit allowance ("only fall back to raw OOXML
where VBA's object model genuinely has no path, and flag that fallback
explicitly"). This module does exactly that:

- `LoadPartXml` extracts a `.pptx`'s given part (e.g.
  `ppt/slideLayouts/slideLayout1.xml`) to a temp folder using
  `Shell.Application`'s native zip-folder support (`Namespace(...).CopyHere`)
  -- there is no zip library in stock VBA, and this is the standard
  escape hatch for exactly this situation.
- `PlaceholderIdxFromDom` finds the shape's `<p:cNvPr name="...">` element,
  walks to its parent's `<p:nvPr>/<p:ph>` sibling, and reads `idx`,
  applying OOXML's own default (`0`) when `<p:ph>` is present but `idx` is
  omitted -- the same default `discovery.py`'s `_placeholder_info` applies.
- `EnrichPlaceholderIdx` is the public entry point: given a `Candidate()`
  array and the source file's path, it fills in `PlaceholderIdx` for every
  `HasPlaceholder` candidate. Callers **must** run this before scoring a
  reference/candidate pool that includes placeholders, or the placeholder
  signal silently degrades to "not applicable" (see below -- which is safe,
  not a crash, but loses the strongest signal `matching.md` specifies).

This means `Matching.bas` is a full port of `matching.md`, not a partial
one deferring the placeholder-index signal -- the fallback exists
specifically so this module doesn't have to leave that gap open further.

## Deliberate divergences from the Python semantics

1. **Unresolved placeholder idx (`-1`) is treated as "signal not
   applicable," not "idx zero."** If `EnrichPlaceholderIdx` is never called
   (or fails), every placeholder candidate still carries Discovery's `-1`
   sentinel. `PlaceholderScore` explicitly checks for `-1` and reports the
   whole signal as inapplicable in that case -- **not** as two unresolved
   placeholders coincidentally matching each other, which would otherwise
   be a real false-positive risk (both sides read `-1`, "match"). This is
   a safety guard the Python original doesn't need (its `placeholder_idx`
   is a real `int | None`, never a fake sentinel value), added here purely
   because VBA's `Long` has no `None`.

2. **Geometry, shape-type, and content signals are always "applicable" in
   this port.** Python's `_geometry_score`/etc. can report "not applicable"
   when the reference itself carries no geometry (raw OOXML with no
   `<a:xfrm>`). PowerPoint's object model always reports
   `Left`/`Top`/`Width`/`Height`/`Type`/text for any live shape -- there is
   no "shape with no geometry" case to detect. `ScoreCandidate` therefore
   never renormalizes away these three weights, only the placeholder one.

3. **Raw-OOXML fallback reads last-saved disk state, not the live,
   possibly-unsaved in-memory shapes.** `EnrichPlaceholderIdx` takes a file
   path and extracts from it directly -- if the presentation has unsaved
   edits (a renamed shape, a shape added/removed since the last save), the
   fallback's view of shape names/placeholders can disagree with the live
   `Shapes` collection `DiscoverSlide`/`DiscoverCustomLayout` just walked.
   Save the file before calling this in the same session it was edited.

4. **Name-based lookup inherits the same duplicate-name risk
   `InjectPrimitive.bas` already documented.** `PlaceholderIdxFromDom`
   matches shapes by `<p:cNvPr name="...">`, since that's the only handle
   the raw XML and the live `Candidate.Name` share. If a part has more than
   one shape with the same name (a real, observed case on a live deck per
   `SPIKE_NOTES.md`), the lookup can't disambiguate and reports `-1`
   (unresolved) rather than guessing -- consistent with this project's
   posture throughout, but a real precision loss versus a hypothetical
   z-order-aware XML lookup this module doesn't build.

5. **`Shell.Application.Namespace(...).CopyHere` is an asynchronous shell
   operation.** `LoadPartXml` polls (up to 10s, 200ms intervals) for the
   target file to actually appear on disk after `CopyHere` returns, rather
   than assuming completion -- a documented real gotcha with this
   technique, not defensive paranoia. If extraction is unusually slow
   (very large deck, slow disk), 10s may not be enough; there's no
   automatic retry beyond that window.

6. **`MatchResult.CandidateIndex` is an array index, not a `Candidate`
   value.** VBA has no clean "optional value" the way Python's
   `Candidate | None` return works for a UDT. Callers check `HasCandidate`
   first, then index the same `candidates()` array they passed to `Match`.

## What was deliberately left out of scope

- No sync-dispatch, resolve, or onboarding logic -- this module only
  scores/matches, exactly like its Python counterpart's non-goals (writing
  the identity tag, or deciding what to do about a flagged match, are the
  caller's job).
- No automated test harness, for the same reason every other module here
  has none: no VBA unit-test framework wired up, and no Office/COM
  available in this sandbox to run one against.

## Manual verification recipe

Run from the VBA IDE (Alt+F11) with the Immediate window open (Ctrl+G).
Import both `Discovery.bas` and `Matching.bas` into the same VBA project
first (`Matching.bas` calls `Discovery.DiscoverSlide`/`DiscoverCustomLayout`
directly).

### 1. `shp-groupshape.pptx` -- sibling ambiguity resolved by z-order

1. Open `test-fixtures/shp-groupshape.pptx` in PowerPoint.
2. Run `Matching.ManualSmokeTest_SiblingAmbiguity`.
3. Expected, per
   `tests/test_matching.py::test_shp_groupshape_sibling_ambiguity_resolved_by_zorder`:
   `Confidence=high`, `HasCandidate=True`, `Name=Oval 2` -- all four leaf
   shapes tie on shape-type/has-text, but z-order (`Oval 2` is z=2, exactly
   matching the reference's `ZOrder=2`) uniquely breaks the tie.

### 2. `mst-slide-layouts.pptx` -- placeholder index alone doesn't force a match

1. Open `test-fixtures/mst-slide-layouts.pptx`. This fixture has no
   `Slides` (only slide layouts + a master, per
   `test_mst_slide_layouts_has_no_ppt_slides_entries`) -- reach its two
   layouts via `Application.ActivePresentation.Designs(1).SlideMaster.CustomLayouts(1)`
   and `(2)`.
2. In the Immediate window:
   `Matching.ManualSmokeTest_PlaceholderIndex "<full path to mst-slide-layouts.pptx>", Application.ActivePresentation.Designs(1).SlideMaster.CustomLayouts(1), Application.ActivePresentation.Designs(1).SlideMaster.CustomLayouts(2)`
3. Expected, per
   `tests/test_matching.py::test_mst_slide_layouts_placeholder_index_alone_does_not_force_high_confidence`:
   `Confidence=medium`, `HasCandidate=False` -- the body placeholder
   (`idx=10` on both layouts) matches on placeholder index, the strongest
   signal, but geometry has drifted too far between the two layouts for
   that alone to auto-accept.
4. Confirm the fallback actually resolved real idx values (not silently
   falling back to "inapplicable"): before running step 2, call
   `Discovery.DiscoverCustomLayout` and `Matching.EnrichPlaceholderIdx`
   directly and inspect `candidates1(i).PlaceholderIdx` in the Immediate
   window for "Text Placeholder 3" -- expect `10`, matching Python's
   `placeholder_idx == 10`.

### 3. Confirm the unresolved-idx safety guard (divergence #1)

1. With `mst-slide-layouts.pptx`'s layout 1 candidates freshly discovered
   (via `Discovery.DiscoverCustomLayout`, **before** calling
   `EnrichPlaceholderIdx`), manually call
   `Matching.ScoreCandidate(candidates(i), candidates(j))` for two
   different placeholder shapes (both still carrying `PlaceholderIdx = -1`).
2. Confirm the placeholder signal does NOT contribute -- the returned
   score should equal what geometry/shape-type/content alone would produce
   (renormalized over those three weights only), not include a false
   "idx -1 = idx -1" match. This proves the guard in divergence #1 actually
   fires rather than being dead code.
