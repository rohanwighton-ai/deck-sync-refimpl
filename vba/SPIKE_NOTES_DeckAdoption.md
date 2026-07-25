# VBA implementation: `DeckAdoption.bas`

Implements `specs/deck-adoption.md` -- bulk retroactive linking of a deck's
already-populated, untagged slides. No Python equivalent exists (confirmed via
`find`/`grep`: no `src/deck_adoption.py`, no `adopt`/`batch` symbol anywhere in
`vba/` before this pass) -- this spec landed VBA-only, so this module is a
direct port of the spec's prose, not a translation of an existing
implementation the way every other `vba/*.bas` module (except
`InjectPrimitive.bas`, the original spike) has been so far.

**Not executed or verified in this environment** -- same constraint as every
module built this pass (2026-07-25): this container has no `powershell.exe`
(confirmed via `which powershell.exe`), so there is no Windows/Office install
reachable to run `run_vba_tests.ps1` against. The manual verification recipe
below is how to actually prove it against a real Office install; the 5 tests
added to `vba/tests/TestRunner.bas` are ready to run the next time this
project is picked up on the WSL/Windows host.

## Scope boundary: engine layer only, per the spec's own Non-goals

`specs/deck-adoption.md`'s Non-goals section is explicit: "The UI itself...
`ribbon-ui.md` gets an 'Adopt Existing Slides' entry point in a *future*
pass... This spec is the engine layer beneath that button, not the button."
`DeckAdoption.bas` has no ribbon/form code and takes plain arrays rather than
`Application.ActiveWindow.Selection.SlideRange` directly -- turning a live UI
selection into `slidesToAdopt()` (and excluding whichever slide was picked as
the template) is left to a future caller, keeping this module headlessly
testable the same way every other module in this port stays testable without
a live user selection.

Two other explicit Non-goals also drove real scoping decisions here:

- **Greenfield template establishment is not this module's job.** "The
  greenfield path... hands the user's chosen template slide to the existing,
  unchanged `onboard-slide-type.md` flow." `PlanAdoption`/`CommitAdoption`
  both take an already-onboarded `templateSld` (tagged fields) as a given --
  they never create one. A caller doing a greenfield adoption runs that
  slide through the existing onboarding flow first (or the underlying
  primitives directly), then calls into this module for every *other* slide
  in the selection.
- **A `needs_confirmation` slide is never partially committed.** The
  spec's per-slide dispatch says medium confidence "needs the existing
  `confirm_field_match` resolution before it can proceed" -- read literally
  as "resolve, then re-run," not "commit whatever did match now and defer
  the rest." `CommitAdoption` reports `needs_confirmation` slides in the
  excluded bucket, untouched, every time; a human must call the existing
  `Onboarding.ConfirmFieldMatch` on the ambiguous shape(s) and re-run
  `PlanAdoption`/`CommitAdoption` for that slide once every role scores high.

## Design decision: plan/commit split, mirroring `SyncOperations.Plan*` / `RunSync.Run*`

The spec's "one phase gate for the whole batch, before any write" requirement
needed a concrete shape without inventing a UI. Followed this project's own
existing precedent instead of designing something new:
`SyncOperations.PlanRoutineSync` (decide, nothing written for the
duplication case) feeding `RunSync.RunRoutineSync` (execute the decision).
`DeckAdoption.PlanAdoption` is the decide half (writes nothing -- confirmed
by the fact its own body never calls `.Tags.Add`, `UpsertRow`, or any other
mutating primitive); `DeckAdoption.CommitAdoption` is the execute half, and
is the *only* function in this module that writes anything. A caller (a
future review form, or a human at the Immediate window per the manual
recipe below) inspects the `AdoptionSlidePlan()` array between the two calls
-- that inspection *is* the phase gate, the same way a human reading
`RunRoutineSync`'s report string today is this project's only "review" for
routine sync.

One consequence worth flagging: **`CommitAdoption` requires a
`confirmedInstanceKeys()` array, parallel to `plans()`**, supplying the
instance_key for every `"ready"` slide the caller wants actually committed.
The spec is explicit that "the instance_key is never auto-generated
silently -- the human supplies or confirms it," and no per-type "designated
key field" concept exists anywhere in this project to auto-suggest one from
(confirmed via `grep -rn "designated\|period_key" vba/*.bas`: `period_key`
is named once, in `specs/identity-tags.md`, as a slide-level tag concept --
but no VBA module reads or writes it anywhere, including this one). Leaving
`confirmedInstanceKeys(i) = ""` for a `"ready"` slide is a legal, safe
no-op (that slide is reported excluded, not guessed) -- this is the hook a
future review form would fill in, not a gap that blocks headless testing.

## Design decision: per-slide confidence dispatch (a judgment call the spec leaves open)

`specs/deck-adoption.md` says "same thresholds as `onboarding.md`" but
`onboarding.md`/`matching.md`'s thresholds are inherently **per-field**
(`Matching.Match` scores one candidate against one reference role), while
this spec's phase-gate review is explicitly **per-slide** ("every slide in
scope with its proposed disposition"). Aggregating N field-level confidences
into one slide-level disposition is not spelled out by any spec in this
project. Chose:

- **`ready`**: every template role the matcher considered scored **high**
  confidence on this slide. (`OnboardNewInstance` -- reused unchanged for
  the actual tag-write -- only ever auto-accepts high-confidence matches,
  so this is also the exact condition under which reusing it commits every
  field, not just some.)
- **`needs_confirmation`**: at least one role scored high or medium, but not
  every role scored high -- the slide is recognizably this type but
  incomplete/ambiguous.
- **`unclassified`**: every role scored low (or there were zero template
  roles at all) -- nothing on this slide plausibly belongs to this type.

Documented here rather than silently baked in, same posture as
`RunSync.bas`'s resequencing-anchor choice or `SlideDuplication.bas`'s
delete-malformed-duplicate choice -- a different, equally defensible
aggregation (e.g. "ready if >=50% of roles are high") was not chosen.

## The one genuinely new piece: keyless Data-sheet row resolution

`ExcelOutput.ReadSheet` deliberately excludes any row with a blank
Instance ID cell from `Sheet.Rows`/`Sheet.InstanceOrder` -- its own read
loop only includes `If instanceId <> ""`. That is correct behavior for
every existing caller (`RunSync.RunRoutineSync` et al. never care about a
row with no key), but it means `ExcelOutput.bas` has **no way to see** the
exact rows this spec needs to resolve against ("a user hand-typed rows
ahead of running this"). Not a bug in `ExcelOutput.bas` -- a real gap this
task is the first to actually need closed, so it's closed locally in
`DeckAdoption.bas` (`ReadKeylessRows`) rather than by reopening
`ExcelOutput.bas`'s already-shipped, already-tested read contract.

`ReadKeylessRows` finds its own last-used-row bound via
`ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1` rather than
`ExcelOutput.bas`'s own `Cells(Rows.Count, 1).End(xlUp)` idiom --
deliberately, since that idiom walks *column A specifically*, which is
exactly the column a keyless row has nothing in. `UsedRange` bounds the
whole sheet's real extent regardless of which column holds data, and needs
no `xlToLeft`/`xlUp` numeric-constant workaround (`.Row`/`.Rows.Count` are
plain properties, not `Range.End` calls) -- so this sidesteps the
cross-app-named-constant gotcha `ExcelOutput.bas` already hit entirely,
rather than needing to reproduce its `XL_UP`-as-numeric-literal fix.

`FindMatchingKeylessRow`/`RowValuesMatchHarvested` implement the spec's
"exact non-key-field match required... zero-or-multiple matches always
falls back to a fresh row" rule literally: a row matches only if *every*
field the row itself has a value for equals the slide's harvested value for
that field exactly. A field the row lacks is not evidence either way (the
row simply never had it); a field the slide harvested but the row lacks is
likewise not evidence either way. `usedRowIds` (threaded through
`PlanAdoption`'s per-slide loop) prevents two different slides in the same
batch from both claiming the same keyless row -- the spec's ambiguity rule
applies within a batch, not just against the sheet's starting state.

## A real bug found and fixed during review (2026-07-25): base-index mismatch between `plans()` and the other three parallel arrays

An adversarial review pass (before this module was ever committed) traced
through `PlanAdoption`'s indexing carefully and found a real, if
not-yet-triggered, correctness bug. The first version of `PlanAdoption`
tracked plan entries with a separate 1-based counter (`n`, incremented once
per slide, `ReDim Preserve plans(1 To n)`), while `harvestedValues` was
allocated over `slidesToAdopt`'s own `LBound`/`UBound` (`ReDim
harvestedValues(lo To hi)`) -- deliberately base-agnostic, since `PlanAdoption`
already reads `slidesToAdopt`'s bounds generically rather than assuming
1-based. `CommitAdoption` then derived its own loop range from
`LBound(plans)`/`UBound(plans)` (always `1`/`n` under the old scheme) and
indexed `slidesToAdopt(i)`/`harvestedValues(i)`/`confirmedInstanceKeys(i)`
with that same `i`. This is only correct if the caller's `slidesToAdopt`
happens to be 1-based -- for any other base (e.g. a 0-based array, which is
completely ordinary VBA and which `PlanAdoption`'s own generic
`LBound`/`UBound` handling implicitly anticipated as a real possibility),
`CommitAdoption` would either raise "Subscript out of range" (once its
`plans`-derived `hi` exceeded `slidesToAdopt`'s real `UBound`) or, worse,
silently misattribute which slide's harvested values/confirmed instance_key
got written to which slide's tags.

Every one of this module's first 5 tests happened to build a 1-based
`Dim slidesToAdopt(1 To 1) As Object`, so none of them could have caught
this -- a single-slide, 1-based test array can never distinguish "indices
happen to line up" from "indices are the same *array*."

**Fix**: `PlanAdoption` now allocates `plans` over the exact same range as
`harvestedValues` (`ReDim plans(lo To hi)`, both sized directly rather than
grown via `ReDim Preserve` from a separate counter -- every slide produces
exactly one plan entry, so the final size is known up front) and indexes
`plans(i)` directly using the same loop variable `i` that indexes
`slidesToAdopt`/`harvestedValues`. `CommitAdoption`'s own docstring now
states explicitly that `confirmedInstanceKeys` (the one array `PlanAdoption`
doesn't produce) must be built by the caller over that identical range.
Added `Test_DeckAdoption_MultiSlideZeroBasedBatchKeepsIndicesAligned` to
`vba/tests/TestRunner.bas` specifically to catch a regression of this: a
deliberately 0-based, 3-slide `slidesToAdopt(0 To 2)` mixing all three live
dispositions (`already_linked`/`ready`/`unclassified`), asserting each
slide receives its *own* tag/instance_key/harvested-value, not another
slide's.

## What was deliberately left out of scope (beyond the Non-goals already covered above)

- **No rollback on a verify-the-link failure.** The spec says a
  non-no-op verification result "is a bug in this pass's harvest... stop
  and flag it." Read narrowly: the tags and Data-sheet row for that one
  slide are already written by the time verification runs (verification is
  a *check*, not a pre-write gate -- the write has to happen before there's
  anything to round-trip). No undo/rollback primitive exists anywhere in
  this project (every write primitive here -- `Shape.Tags.Add`,
  `ExcelOutput.UpsertRow`, the object-model tag writes `OnboardNewInstance`
  makes -- is a direct, immediate mutation with no transaction concept), so
  "stop and flag" is implemented as "report this one slide in
  `FailedVerificationLabels`, continue processing the rest of the batch" --
  not a full batch abort, and not an automatic undo of this slide's own
  writes. A human seeing a non-zero `FailedVerificationCount` in the report
  needs to manually inspect and correct that slide's tags/row. Flagged here
  as a real, load-bearing limitation, not glossed over.
- **Row-order-bootstrap is a documented caller precondition, not
  self-enforced.** "Newly created rows are appended... in the same order
  their source slides currently appear in the deck" falls out of
  `ExcelOutput.UpsertRow`'s own append-in-call-order behavior *only if*
  `CommitAdoption` iterates `slidesToAdopt()` in deck order -- which in turn
  requires the caller to have supplied it in deck order to begin with. This
  function does not sort by `SlideIndex` itself (no sort primitive exists
  anywhere in this project's VBA to reach for, and `Application.
  ActiveWindow.Selection.SlideRange` iterates in the order the underlying
  `Slides` collection already exposes, which is deck order, for a normal
  contiguous or ctrl-click selection) -- documented as a precondition on
  `PlanAdoption`'s own header comment rather than defended against
  defensively.
- **`AdoptionSlidePlan` is deliberately scalar-only.** No member is a
  dynamic array of another UDT (e.g. `Matches() As FieldMatch`) -- this
  project has never exercised "array-of-UDT where each element itself
  contains a dynamic array of a different UDT" in a real Office run (every
  existing array-of-UDT type -- `SyncAction`, `FieldMatch`, `MatchResult`
  -- is scalars plus at most one `Object` member), and given how many real,
  Office-specific VBA gotchas this project has already found the hard way
  (see `AGENTS.md`'s Known Patterns), introducing an untested construct
  here rather than an already-proven one was judged not worth the risk for
  what plan/commit's own design doesn't actually require. `CommitAdoption`
  re-derives field-shape matches via `Onboarding.OnboardNewInstance`'s own
  internal call to `MatchSlideAgainstTemplate` instead of threading
  `PlanAdoption`'s own match results through -- redundant scoring work
  (bounded by one slide's shape count, not the whole batch), traded
  deliberately for staying inside proven patterns only.

## Manual verification recipe

Import `Discovery.bas`, `InjectPrimitive.bas`, `Matching.bas`, `Resolve.bas`,
`SyncOperations.bas`, `Onboarding.bas`, `Verification.bas`,
`SlideDuplication.bas`, `ExcelOutput.bas`, and `DeckAdoption.bas` into the
same VBA project (`DeckAdoption.bas` calls into `Onboarding`, `Resolve`,
`ExcelOutput`, and `InjectPrimitive` directly). Automated coverage lives in
`vba/tests/TestRunner.bas`'s 5 new `Test_DeckAdoption_*` functions -- run via
`vba/tests/run_vba_tests.ps1` on a real Windows/Office host, same as every
other module's tests, per `AGENTS.md`'s Testing section. They cover:
idempotent already-linked skip, a high-confidence slide committed and
verified (fresh Data-sheet row), a medium-confidence slide correctly left
untouched (`needs_confirmation`, no partial tag write), a fully unrelated
slide excluded (`unclassified`), and linking to a pre-existing keyless
Data-sheet row (Instance ID cell filled in, no duplicate row created)
rather than always appending fresh.

For an ad hoc, interactive check beyond what the harness covers (e.g. the
whole-batch phase-gate flavor, since the automated tests exercise
`PlanAdoption`/`CommitAdoption` per-slide rather than as one multi-slide
`Immediate`-window batch): build a template slide (one tagged field is
enough), 2-3 further slides at varying drift from it, and a worksheet with
one keyless row matching one of those slides' expected harvested value
verbatim. In the Immediate window:

```
Dim harvestedValues() As Object
Dim plans() As AdoptionSlidePlan
Dim slides(1 To 3) As Object
Set slides(1) = ActivePresentation.Slides(2) ' etc -- whichever slides are in scope, excluding the template
plans = DeckAdoption.PlanAdoption(slides, ActivePresentation.Slides(1), Worksheets("Data"), harvestedValues)
```

Then inspect `plans(1).Disposition`/`.Reason`/`.MatchedKeylessRowId` per
slide before deciding `confirmedInstanceKeys()` and calling
`CommitAdoption` -- this Immediate-window inspection step *is* the manual
stand-in for the phase-gate review form a future `ribbon-ui.md` pass would
build.
