# VBA port: `onboarding` module

Module 5 of `specs/vba-port.md`'s port order (after discovery, identity_tags,
matching, resolve+sync_operations). Ports `match_slide_against_template()`,
`confirm_field_match()`, and `onboard_new_instance()` (`Onboarding.bas`) --
matching a *subsequent* untagged slide's candidates against an
already-established template. First-time onboarding of a type needs no code
here at all, per `specs/onboarding.md`: "the working copy IS becoming the
reference" -- direct `Slide.Tags.Add`/`Shape.Tags.Add` calls after a human
confirms the discovered fields, exactly as `SyncOperations.bas`'s own
`ManualSmokeTest_NoChangeThenInPlaceCorrection` recipe already demonstrates.

**Not executed or verified in this environment** -- there is no Windows/Office
install here, same constraint as every prior module. `Onboarding.bas` has not
been run. The manual verification recipe below is how to actually prove it
against a real Office install, cross-checked against
`tests/test_onboarding.py`'s already-proven Python values.

## A real gap this module closes: Candidate -> live Shape

Neither `Discovery.DiscoverSlide` nor `Discovery.DiscoverCustomLayout` ever
exposed the live `Shape` object behind a returned `Candidate` -- no prior
caller needed one. `InjectPrimitive.bas` re-locates a shape by its role
**tag** (`FindShapeByRoleTag`), never by `Candidate` identity. Onboarding is
the first caller needing the opposite direction: given a `Candidate` a human
or the matcher accepted, write a tag onto its actual shape; given a
template's already-tagged shapes, read which `Candidate` each role belongs
to.

Closed in `Discovery.bas` (not duplicated here) by threading a parallel
`ByRef shapes() As Object` array through `Walk` and adding two new public
entry points, `DiscoverSlideWithShapes`/`DiscoverCustomLayoutWithShapes`,
that expose it -- `DiscoverSlide`/`DiscoverCustomLayout`'s own signatures and
output are unchanged (they pass a local throwaway array). Safe because
`Candidate.ZOrder` is exactly "1-based position among discovered leaf
shapes" and `Walk` assigns it in one deterministic pass over the same tree
both old and new entry points traverse -- pairing candidate index to shape
index is an exact match, not a heuristic re-match by name (shape names are
not unique, per `InjectPrimitive.bas`'s own real-deck finding).

## A real VBA restriction found while building this (not previously documented anywhere in this project)

**A UDT (`Type ... End Type`) cannot be assigned to a `Variant`** in VBA --
this is a compile-time restriction ("Invalid use of type"), not a style
preference. `Scripting.Dictionary`'s `Item` property is `Variant`-typed, so
`dict(key) = someTypedVariable` fails to compile whenever `someTypedVariable`
is a `Candidate`, `MatchResult`, `InjectResult`, etc. This matters here
because the natural design for `BuildTemplateFieldShapes`'s role->shape
mapping would have been a `Scripting.Dictionary`, matching this project's
own Python `dict[str, Candidate]` -- **not viable in VBA.** Used parallel
arrays instead (`roles() As String` alongside `Candidate()`, same indices).

**This same pattern already appears in the previously-shipped
`SyncOperations.bas`**: `PlanRoutineSync`'s `changed(fieldName) = r` (line
~135) assigns an `InjectResult` UDT into a `Scripting.Dictionary`. Per the
restriction above, this looks like it would fail to compile the moment
`SyncOperations.bas` is actually imported into a real VBA project -- but
this is a finding, not a fix made here. `SyncOperations.bas` is a different,
already-committed module with its own `SPIKE_NOTES_Resolve.md`; fixing it is
outside this module's scope and is flagged for a separate decision rather
than silently patched in passing.

## What was ported

- `Onboarding.BuildTemplateFieldShapes(templateSld, ByRef roles())`: for
  every discovered candidate on an already-tagged slide whose live shape
  carries a non-blank `role` tag, records the role and its `Candidate`.
  Mirrors `resolve.py`'s `field_shapes` construction (the half `Resolve.bas`
  deliberately didn't build -- see its own header comment on why -- but
  that `MatchSlideAgainstTemplate` genuinely needs, since scoring requires
  the reference shape's geometry/placeholder data, not just "a shape exists
  for this role").
- `Onboarding.MatchSlideAgainstTemplate(sld, templateRoles(),
  templateFieldShapes(), ByRef untaggedShapes())`: per role, scores every
  untagged, non-decoration candidate on `sld` against the template's
  reference via `Matching.Match`, field-for-field against
  `match_slide_against_template()`.
- `Onboarding.ConfirmFieldMatch(shp, role)`: writes `role` directly onto a
  live shape -- the primitive an eventual selection UI would call, per
  `specs/onboarding.md`'s Non-goals.
- `Onboarding.OnboardNewInstance(sld, templateRoles(), templateFieldShapes(),
  slideType, instanceKey)`: tags slide-level identity unconditionally, then
  auto-accepts and tags every high-confidence field match, leaving
  medium/low for a human -- field-for-field against `onboard_new_instance()`.

## Deliberate divergences from the Python semantics

1. **Parallel arrays, not a role->Candidate dictionary.** See the UDT/Variant
   restriction above -- not a style choice, a compile-time constraint.
   `templateRoles()`/`templateFieldShapes()` must be passed together and
   stay index-aligned; there is no `.Keys`/`.Item(role)` convenience the
   Python `dict` gave.
2. **`BuildTemplateFieldShapes` is new, not a reuse of `Resolve.SlideInstance`.**
   `Resolve.bas`'s `SlideInstance` UDT carries only `HasTypeTag`/`TypeTag`/
   `HasInstanceKey`/`InstanceKey` -- no `field_shapes`, by that module's own
   documented choice (`SyncOperations` never needed one; `InjectPrimitive`
   does its own per-field tag lookup instead). Onboarding's template
   argument genuinely needs reference *geometry*, not just "does a shape
   exist for this role," so it builds its own separate structure rather
   than retrofitting `Resolve.SlideInstance` -- keeps `Resolve.bas`
   untouched, consistent with this project's practice of extending in the
   new module rather than reopening an already-shipped one.
3. **`IsCandidateField` lives in `Onboarding.bas`, not on `Discovery.Candidate`.**
   Matches `specs/discovery.md`'s own framing (quoted directly in
   `src/onboarding.py`'s docstring): the filter belongs at the point
   candidates are handed to the matcher, not inside discovery itself --
   other callers (`verify_structure`-equivalent work) still need the full,
   unfiltered shape list.
4. **`Shape.Tags.Add` is a native upsert; no read-merge-write.** Both
   `ConfirmFieldMatch` and `OnboardNewInstance`'s slide-tag writes rely on
   the documented behavior of `Tags.Add`: if the name already exists, its
   value is replaced, not duplicated. The Python original's read-merge-write
   discipline (`identity_tags.md`'s requirement) existed to protect against
   hand-rolled OOXML surgery dropping *other* tags -- irrelevant here, since
   the object model's `Tags` collection is a real key-value store that
   VBA/PowerPoint itself maintains.
5. **No SHA-256 hashing, no re-open-from-disk verification** -- same
   already-documented divergence as every module whose writes ultimately go
   through the object model (`InjectPrimitive.bas`'s tag writes here, not
   text writes, but the same reasoning applies: verification is against the
   live in-memory state, not a reopened file).

## What was deliberately left out of scope

- The shape-selection UI a human would use for medium/low-confidence matches
  or first-time onboarding -- `specs/onboarding.md`'s own Non-goals. This
  module builds `ConfirmFieldMatch`, the primitive such a UI would call, not
  the UI itself.
- Deciding where a type's template is physically stored (a dedicated slide?
  a separate file?) -- `BuildTemplateFieldShapes` takes an already-resolved
  template `Slide` object as given, same boundary `specs/onboarding.md`
  draws for itself.
- Reconciling a template whose own tags have drifted/corrupted -- that's
  onboarding *of the template*, not matching *against* one; out of scope
  here, same as the Python original.
- Port-order step 6 (Excel-side reads/writes) -- still not built; unrelated
  to this module.
- Fixing the UDT-in-Dictionary issue found in `SyncOperations.bas` -- flagged
  above, not fixed here.

## Manual verification recipe

Run from the VBA IDE (Alt+F11) with the Immediate window open (Ctrl+G).
Import `Discovery.bas` (updated), `InjectPrimitive.bas`, `Matching.bas`,
`Resolve.bas`, `SyncOperations.bas`, and `Onboarding.bas` into the same VBA
project (`Onboarding.bas` calls `Discovery.DiscoverSlideWithShapes` and
`Matching.Match` directly).

### 1. Set up the template and new slide, matching `tests/test_onboarding.py`'s `_build_template()`

Using `test-fixtures/mst-slide-layouts.pptx` (or, if `CustomLayouts` proves
awkward to drive live, two ordinary slides built to the same shape): apply
layout 1 to slide 1 (the template) and layout 2 to slide 2 (the new,
untagged slide) -- real fixture geometry drift between the two layouts is
what makes Title score high and Body score medium, not a constructed
scenario.

In the Immediate window, tag slide 1 as the template:
```
Application.ActivePresentation.Slides(1).Tags.Add "slide_type", "quarterly-update"
Application.ActivePresentation.Slides(1).Tags.Add "instance_key", "rec-1"
Application.ActivePresentation.Slides(1).Shapes(1).Tags.Add "role", "Title"
Application.ActivePresentation.Slides(1).Shapes(2).Tags.Add "role", "Body"
```
(adjust shape indices to whichever are actually the title/body placeholders
on your slide 1). Leave slide 2 completely untagged.

### 2. `MatchSlideAgainstTemplate` -- confirm both confidence branches

1. In the Immediate window:
   ```
   Dim roles() As String, shapes() As Candidate, m() As FieldMatch
   shapes = Onboarding.BuildTemplateFieldShapes(Application.ActivePresentation.Slides(1), roles)
   Dim untagged() As Object
   m = Onboarding.MatchSlideAgainstTemplate(Application.ActivePresentation.Slides(2), roles, shapes, untagged)
   ```
2. Expected, per
   `tests/test_onboarding.py::test_match_slide_against_template_scores_both_high_and_medium_confidence`:
   one role (Title) with `Confidence="high"`, `HasCandidate=True`; one role
   (Body) with `Confidence="medium"`, `HasCandidate=False`.

### 3. `Onboarding.ManualSmokeTest_OnboardNewInstance` -- auto-accept + slide identity

1. Run `Onboarding.ManualSmokeTest_OnboardNewInstance`.
2. Expected, per
   `tests/test_onboarding.py::test_onboard_new_instance_auto_accepts_high_confidence_and_leaves_medium_unresolved`:
   Title shows `confidence=high hasCandidate=True`; Body shows
   `confidence=medium hasCandidate=False`.
3. Confirm in the Immediate window that slide 2's Title-role shape actually
   got tagged (`Application.ActivePresentation.Slides(2).Shapes(1).Tags("role")`
   should now read `"Title"` -- adjust the shape reference to whichever one
   matched), and that Body's shape was **not** tagged.
4. Confirm slide-level identity landed unconditionally, per
   `tests/test_onboarding.py::test_onboard_new_instance_tags_slide_identity_unconditionally`:
   `Application.ActivePresentation.Slides(2).Tags("slide_type")` =
   `"quarterly-update"`, `Tags("instance_key")` = `"rec-2"`.

### 4. `ConfirmFieldMatch` -- resolving a flagged medium-confidence match

1. After step 3, manually identify slide 2's Body-role shape (whichever
   candidate `MatchSlideAgainstTemplate`'s medium-confidence result pointed
   at, from step 2's `untagged`/`m` output) and confirm it:
   ```
   Onboarding.ConfirmFieldMatch untagged(<index from m's Body result>), "Body"
   ```
2. Expected, per
   `tests/test_onboarding.py::test_confirm_field_match_resolves_a_flagged_medium_confidence_match`:
   that shape now reads `Tags("role") = "Body"`.

### 5. Pure decoration is never matched or auto-tagged

Using `test-fixtures/shp-groupshape.pptx` (its "Oval 2" and three siblings
are pure decoration -- no text, not pictures, per
`tests/test_discovery.py::test_shp_groupshape_finds_zero_candidate_fields`):

1. Build a template whose only field role points at one of those decoration
   shapes (this is a synthetic setup purely to exercise the filter -- a real
   template would never do this in practice).
2. Run `MatchSlideAgainstTemplate`/`OnboardNewInstance` against the same
   slide. Expected, per
   `tests/test_onboarding.py::test_match_slide_against_template_never_matches_pure_decoration`
   and `test_onboard_new_instance_never_auto_tags_pure_decoration`: the
   result is `Confidence="low"`, `HasCandidate=False`, and no tag gets
   written anywhere.
