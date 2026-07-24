# VBA port: `resolve` + `sync_operations` modules

Module 4 of `specs/vba-port.md`'s port order (grouped together there, same
as the Python source: `src/resolve.py` composes `src/sync_operations.py`'s
`SlideInstance` input). Ports `resolve_slide_instance()` (`Resolve.bas`) and
`plan_routine_sync()`/`plan_period_rollover()` (`SyncOperations.bas`) to
native VBA, operating on live `Slide` objects and `InjectPrimitive.bas`'s
existing tag-based write primitive instead of file-surgery.

**Not executed or verified in this environment** -- there is no
Windows/Office install here, same constraint as every prior module.
`Resolve.bas`/`SyncOperations.bas` have not been run. The manual
verification recipe below is how to actually prove them against a real
Office install, cross-checked against `tests/test_resolve.py`'s and
`tests/test_sync_operations.py`'s already-proven Python values.

## A real, previously-mis-stated gap this module closes

`specs/vba-port.md`'s port order describes module 2 (`identity_tags`) as
"already done (`InjectPrimitive.bas`'s `Shape.Tags`/`Slide.Tags` reads)".
That is only half true: grepping `InjectPrimitive.bas` confirms it reads
`Shape.Tags("role")` (via `FindShapeByRoleTag`) but never once reads
`Slide.Tags` for anything. No VBA file in this project read a slide-level
tag before this one. This module is the first consumer that actually needs
`slide_type`/`instance_key` (to resolve a `SlideInstance`), so
`Resolve.ResolveSlideInstance()` adds the missing `Slide.Tags` reads here --
exactly the "extend only if a gap is found" instruction module 2's own
port-order entry gives, rather than carrying the mis-statement forward
silently. Scoped to **reads only**: writing slide-level tags is
onboarding's job (port-order step 5, not yet built) and stays an open item
for that module, not invented speculatively here.

## What was ported

- `Resolve.ResolveSlideInstance(sld)`: reads `sld.Tags("slide_type")` and
  `sld.Tags("instance_key")` into a `SlideInstance` UDT
  (`HasTypeTag`/`TypeTag`, `HasInstanceKey`/`InstanceKey`), mirroring
  `resolve_slide_instance()`'s slide-level half exactly.
- `SyncOperations.PlanRoutineSync(instances, instanceOrder, dataRows)`:
  cases 1 (no_change), 3 (new_record), 4 (in_place_correction), and 6
  (unclassified_slide), field-for-field against
  `plan_routine_sync()`: instances missing a type/instance tag are flagged
  independently of the Data-sheet walk; a Data-sheet row with no known
  instance is `new_record`; a known instance's row is dispatched per field
  via `InjectPrimitive.InjectPrimitive()`, whose own `Written` result is
  the case-1-vs-4 classifier, exactly as `inject_primitive`'s result is in
  Python -- no separate diff step.
- `SyncOperations.PlanPeriodRollover(instance, newValues)`: case 2, a
  distinct function never called from `PlanRoutineSync`, matching
  `plan_period_rollover()`'s "never inferred from a value merely looking
  different" rule. Raises if `instance` has no `InstanceKey`, same guard
  as the Python `ValueError`.

## Deliberate divergences from the Python semantics

1. **No separate `field_shapes` pre-resolution step.** `resolve.py` builds
   a `role -> Candidate` dict once per slide because Python's
   `inject_primitive` needs a `Candidate`'s `z_order` to re-locate the
   shape during raw zip/XML surgery. VBA's `InjectPrimitive` already does
   its own tag-based lookup internally (`FindShapeByRoleTag`) every time
   it's called -- there is no z-order or file-offset to pre-resolve. So
   `SyncOperations.PlanRoutineSync` calls
   `InjectPrimitive.InjectPrimitive(instanceSlide, fieldName, value)`
   directly per field, using the field name itself as the `role` tag
   value to look up (matching the Python fixture convention: a shape
   tagged `role="Title"` corresponds to a Data-sheet field named
   `"Title"`). Pre-building a map first would be pure indirection with no
   benefit in this port -- a genuine simplification, not a corner cut,
   same category as `Shape.Tags` collapsing `identity_tags.py`'s ~250
   lines of XML handling in the original spike.

2. **`InjectPrimitive`'s `Found = False` conflates two Python-distinguishable
   situations into one skip outcome.** In Python, `field_shapes.get(field)`
   returning `None` (no shape has this role) causes a clean skip; a
   duplicate role tag on two shapes would instead silently let
   `resolve.py`'s dict comprehension overwrite one `Candidate` with
   another (last write, by z-order, wins) with no error raised anywhere.
   VBA's `FindShapeByRoleTag` refuses to guess in that same duplicate-tag
   case and reports `Found = False` -- which `PlanRoutineSync` here treats
   identically to "field not present" (skip, not written this round).
   This is arguably **safer** than the Python original (it never silently
   picks a wrong shape), but it does mean a real duplicate-role bug on a
   live deck reads as "field simply didn't sync" here rather than as any
   kind of visible error. Disambiguating that is case-7
   (`deck_side_conflict`) adjacent territory, a non-goal per
   `specs/sync-operations.md` -- not solved here, just not silently
   guessed at either.

3. **`instances()` is caller-supplied, not discovered.** Gathering "every
   live slide belonging to this type" (e.g. walking
   `Application.ActivePresentation.Slides` and checking tags) is the
   caller's job, matching `sync_operations.py`'s own non-goal --
   `SlideInstance` there is likewise "already resolved," not something the
   module reads off a deck itself.

4. **The Data-sheet input is a placeholder shape for a module that doesn't
   exist yet.** Port-order step 6 (Excel-side reads/writes) is not built.
   `PlanRoutineSync` accepts `dataRows`/`instanceOrder` already read into
   memory (a `Scripting.Dictionary` of `Scripting.Dictionary`s, plus a row-
   order `Collection`) rather than reading a worksheet itself -- mirroring
   `plan_routine_sync()`'s own separation (it takes an already-read
   `Sheet`, never touches a file). Whichever module implements step 6
   **must** produce exactly this shape (documented in
   `SyncOperations.bas`'s `PlanRoutineSync` header) or adapt it; this is a
   real interface decision made here, not deferred.

5. **`Flagged.Subject` uses `Slide.SlideID`, not `part_path`.** There is no
   OOXML part-path concept for a live in-memory `Slide` the way Python's
   file-surgery approach has one; `SlideID` is the closest stable native
   handle (survives reordering, unlike `SlideIndex`; unlike `.Name`, which
   is frequently blank on an untouched slide).

6. **`CustomLayout.Tags` support is unconfirmed.** `ResolveSlideInstance`
   takes `sld As Object` generically so it also works against a
   `CustomLayout` (needed for `mst-slide-layouts.pptx`-style fixtures,
   which have no `Slides` at all -- the exact fixture
   `tests/test_resolve.py` uses). Whether `CustomLayout` actually exposes
   a native `.Tags` property the same way `Slide` does is **not
   confirmed** -- no Office install here to check. If it turns out
   `CustomLayout` has no `.Tags`, the manual recipe below falls back to a
   normal slide-based fixture instead; this note should be upgraded from
   "unconfirmed" to a confirmed finding in a follow-up edit once checked.
7. **No SHA-256 hashing, no re-open-from-disk verification.** Both
   already-documented divergences of `InjectPrimitive.bas` (direct string
   equality; verification against the live object model, not a re-opened
   file) apply unchanged here since this module's writes go through
   `InjectPrimitive` -- not reintroduced or re-explained per call site.

## What was deliberately left out of scope

- Cases 5 (`record_retired`) and 7 (`deck_side_conflict`) -- non-goals per
  `specs/sync-operations.md`, not produced anywhere in this file.
- Physical slide duplication -- deciding a duplicate is needed (cases 2/3)
  is in scope; performing it is not (`Slide.Duplicate` is the real target's
  native mechanism, per `specs/sync-operations.md`'s Non-goals).
- Excel-side reads/writes (port-order step 6) -- `dataRows`/`instanceOrder`
  are accepted pre-built, per divergence 4 above.
- Onboarding (port-order step 5) -- writing slide-level tags
  (`slide_type`/`instance_key`) and shape-level `role` tags is not this
  module's job; it only reads what onboarding is assumed to have already
  written.
- No automated test harness, same reason as every other module here: no
  VBA unit-test framework wired up, no Office/COM available in this
  sandbox.

## Manual verification recipe

Run from the VBA IDE (Alt+F11) with the Immediate window open (Ctrl+G).
Import `Discovery.bas`, `InjectPrimitive.bas`, `Resolve.bas`, and
`SyncOperations.bas` into the same VBA project first (`SyncOperations.bas`
calls both `Resolve.ResolveSlideInstance` and
`InjectPrimitive.InjectPrimitive` directly).

### 1. Tag a fixture, matching `tests/test_resolve.py`'s `_onboard()`

`test-fixtures/mst-slide-layouts.pptx` has no `Slides` collection -- try
`CustomLayouts(1)` first per divergence 6 above; if `.Tags` isn't available
on a `CustomLayout` object in your Office version, use any ordinary
single-slide `.pptx` with a text/title shape instead and adjust the object
references below accordingly (the dispatch logic under test doesn't care
which).

1. Open the fixture in PowerPoint.
2. In the Immediate window (substituting `CustomLayouts(1)` for `Slides(1)`
   if using the layouts fixture):
   ```
   Application.ActivePresentation.Slides(1).Tags.Add "slide_type", "quarterly-update"
   Application.ActivePresentation.Slides(1).Tags.Add "instance_key", "rec-1"
   Application.ActivePresentation.Slides(1).Shapes(1).Tags.Add "role", "Title"
   ```
   (pick the actual title shape's index/name if it isn't `Shapes(1)`).

### 2. `Resolve.ManualSmokeTest` -- confirm the slide-level tag gap is closed

1. Run `Resolve.ManualSmokeTest`.
2. Expected, per `tests/test_resolve.py::test_resolve_reads_real_tags_off_disk_into_a_slide_instance`:
   `HasTypeTag=True TypeTag=quarterly-update HasInstanceKey=True InstanceKey=rec-1`.
3. Remove the two `Slide.Tags.Add` calls from step 1 (or test against a
   fresh untagged slide) and re-run. Expected, per
   `tests/test_resolve.py::test_untagged_slide_resolves_to_none_key_and_type`:
   `HasTypeTag=False HasInstanceKey=False`.

### 3. `SyncOperations.ManualSmokeTest_NoChangeThenInPlaceCorrection` -- cases 1 and 4

With the slide/shape tagged per step 1:

1. Run `SyncOperations.ManualSmokeTest_NoChangeThenInPlaceCorrection` once
   with its `TITLE_TEXT` constant left blank. Expected, per
   `tests/test_resolve.py::test_end_to_end_no_change_then_in_place_correction`'s
   first half: `Kind=no_change`.
2. Edit `TITLE_TEXT` to a different string (e.g. `"Q3 Revenue"`) and
   re-run. Expected, per that same test's second half: `Kind=in_place_correction`,
   and the title shape's visible text actually changes to the new value.

### 4. Case 3 (`new_record`) and case 6 (`unclassified_slide`)

Not wired into the smoke-test subs above (they need a second, deliberately
untagged slide) -- verify these two directly from the Immediate window:

1. **`new_record`**: build `instances`/`order`/`rowsDict` as in
   `ManualSmokeTest_NoChangeThenInPlaceCorrection`, but add a second row to
   `order`/`rowsDict` under an instance key no tagged slide carries (e.g.
   `"rec-new"`). Call `PlanRoutineSync` directly. Expected, per
   `tests/test_resolve.py::test_end_to_end_new_record_when_data_sheet_row_has_no_onboarded_instance`:
   one `no_change` action for `rec-1` and one `new_record` action for
   `rec-new` with `Values("Title") = "Brand New"` (or whatever value you
   used).
2. **`unclassified_slide`**: use a slide with no `slide_type`/`instance_key`
   tags at all as the sole entry in `instances`, with an empty `order`
   Collection. Expected, per
   `tests/test_resolve.py::test_end_to_end_never_onboarded_slide_is_flagged_not_silently_skipped`:
   exactly one action, `Kind=flagged`, `FlagKind=unclassified_slide`.

### 5. `PlanPeriodRollover` -- case 2 is never auto-dispatched

1. Confirm by inspection (or by stepping through with the debugger) that
   no code path in `PlanRoutineSync` calls `PlanPeriodRollover` -- matches
   `tests/test_sync_operations.py::test_period_rollover_never_produced_by_routine_sync`.
2. Call `PlanPeriodRollover` directly against a resolved `SlideInstance`
   with `HasInstanceKey=True`; confirm it returns a `PeriodRollover` with
   `Reason="explicit period-rollover command"`.
3. Call it against a `SlideInstance` with `HasInstanceKey=False`; confirm
   it raises (`Err.Raise`) rather than returning a rollover with a blank
   key.
