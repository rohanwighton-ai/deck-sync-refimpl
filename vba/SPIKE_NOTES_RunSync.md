# VBA implementation: `RunSync.bas` (the orchestration layer)

No Python equivalent -- every module in this port explicitly draws a
boundary around orchestration as "someone else's job": gathering live
instances is `sync_operations.py`'s own stated non-goal; executing a
decided duplication is `SlideDuplication.bas`'s caller's job; reconciling
deck order against Data-sheet row order every sync is a whole-type,
driver-level concern per `specs/slide-duplication-trigger.md`. This module
is that someone else -- the first place all 9 other modules
(discovery/identity_tags/matching/resolve/sync_operations/onboarding/
excel_output/verification/slide_duplication) get composed into one real,
runnable sync pass. Written and executed against real Office the same day
(2026-07-25).

## What was built

- **`RunSync.GatherInstances(slideType)`**: every slide in the active
  presentation whose `slide_type` tag matches, in current deck order.
  Possibly unallocated (a genuinely empty type). The gathering step every
  other module explicitly deferred.
- **`RunSync.RunRoutineSync(ws, slideType, templateSld)`**: the main entry
  point. Reads `ws` via `ExcelOutput.ReadSheet`, gathers instances, calls
  `SyncOperations.PlanRoutineSync`, and dispatches its decisions:
  `no_change`/`in_place_correction` are already executed as a side effect
  of planning itself (`PlanRoutineSync` calls `InjectPrimitive` directly
  per field -- nothing left to do here but report); `new_record` is
  executed via `SlideDuplication.DuplicateAndTag`; `flagged` is reported,
  never forced. Returns a human-readable report string (counts + per-row
  detail), not just a silent side effect.
- **`RunSync.ResequenceByRowOrder(slideType, instanceOrder)`**: applied
  after every routine sync pass, reconciling deck order for the *whole*
  type (both newly-created and pre-existing slides) against current
  Data-sheet row order -- the standing invariant `specs/slide-duplication-
  trigger.md` requires, not a one-time stamp at creation.

## The real, previously-hidden bug this module's first real run exposed

`ExcelOutput.bas`'s `xlToLeft`/`xlUp` named constants don't resolve when
driven cross-app from a PowerPoint-hosted VBA project -- see
`SPIKE_NOTES_ExcelOutput.md`'s own entry for the full finding. This had
been invisible through every prior `ExcelOutput.bas` test because those all
ran *inside* Excel's own VBA project, where the names happen to resolve
natively. `RunSync.RunRoutineSync` was the first code in this entire
project to actually call `ExcelOutput.ReadSheet` from a foreign host
application, which is exactly what surfaced it -- a real, concrete
demonstration of why "runs inside Excel" and "drives Excel from PowerPoint"
(both named as valid targets in `vba-port.md`) are not interchangeable
assumptions, and why a module tested only in one context can silently carry
a bug the other context alone reveals.

Diagnosed by screenshotting the actual PowerPoint/VBE window mid-hang
(window-title/`Err.Number` probing alone wasn't resolving it -- a compile
error surfaced as an indefinite hang rather than a clean COM exception,
consistent with this project's earlier finding that uncaught compile-time
issues can pop a blocking modal rather than propagate cleanly; see
`AGENTS.md`'s Testing section) -- the compile-error dialog's own text
("Variable not defined", `xlToLeft` highlighted) named the exact cause
directly.

## A real design decision the spec left open: how resequencing is anchored

`specs/slide-duplication-trigger.md` says deck order must match Data-sheet
row order but doesn't say *where* in the deck that ordered block should
sit. `ResequenceByRowOrder` anchors it at the **current lowest `SlideIndex`**
among the type's existing slides (found before any moves), not the front or
back of the deck -- preserving wherever a human originally chose to place
this type's slides within a deck that mixes multiple types, and fixing only
their *relative* order among themselves. A different, equally defensible
choice (always push to front/back) was not made. Implemented via repeated
`Slide.MoveTo` calls walking `instanceOrder` in order, each one placed
immediately after the previous one's already-corrected position -- which
does mean this type's slides end up **contiguous** even if they started
interspersed with other types' slides. Not fully specified by
`slide-duplication-trigger.md` either way; flagged here as a real judgment
call, same as `SlideDuplication.bas`'s delete-malformed-duplicate choice.

## What was deliberately left out of scope

- Cases 5/7 -- non-goals throughout this entire project.
- Deciding where a type's template lives -- `templateSld` is supplied by
  the caller, per `onboarding.md`'s own non-goal.

## Case 2 (period rollover): `RunSync.RunPeriodRollover` (2026-07-25 pass)

Added after `specs/ribbon-ui.md` landed and flagged its "New Period" button
as needing an execution primitive that didn't exist yet -- `SyncOperations.
PlanPeriodRollover` only ever *decided* a rollover (returned a
`PeriodRollover` struct); nothing called `SlideDuplication.DuplicateAndTag`
with that decision the way `RunRoutineSync` already does for case 3.

`RunPeriodRollover(sourceSld, slideType, newInstanceKey, newValues)` closes
that gap, mirroring `RunRoutineSync`'s own case-3 handling: resolves
`sourceSld` (the instance's own current slide -- confirmed by
`Resolve.ResolveSlideInstance`, no separate "look up the slide for this
instance_key" step exists or is needed, since the caller already has the
slide in hand) into a `SlideInstance`, calls `SyncOperations.
PlanPeriodRollover` to get the rollover decision (this is what actually
validates the instance carries an `instance_key` -- raises otherwise, same
guard `PlanPeriodRollover` already had), gathers `slideType`'s current
instances itself (same non-goal-gathering-is-someone-else's-job posture
every other module in this port takes, and the same thing
`RunRoutineSync` does before dispatching), then calls `SlideDuplication.
DuplicateAndTag(sourceSld, slideType, newInstanceKey, rollover.NewValues,
existingInstances)` -- `sourceSld` itself is never mutated by
`DuplicateAndTag` (it only writes to the new duplicate), which is exactly
`sync-operations.md`'s case-2 requirement that the original stays
untouched as history, so no separate "leave the source alone" logic was
needed here -- it falls out of reusing the existing primitive as-is.

Deliberately narrow, matching the task's own scope: takes the *slide*
directly (`sourceSld`), not an instance_key to look up -- resolving
"which slide does this instance_key currently live on" across a whole
deck is a search a caller (e.g. the New Period picker UI, Priority 21)
is better positioned to have already done via its own selection/lookup,
not something this primitive should reinvent. Returns `DuplicateResult`
directly (the same structured type `RunRoutineSync`'s own case-3 branch
already consumes) rather than a formatted `String` report -- deliberately
not mirroring `RunRoutineSync`'s return type here, since `ribbon-ui.md`'s
shared result form is unbuilt UI-layer work that hasn't decided yet
whether `RunRoutineSync` itself should move off `String` reports too (see
that task's own notes); a fresh function returning the already-structured
type avoids adding a second `String`-report format to later un-parse.

## Verification

Automated: `Test_RunSync_GatherInstancesFiltersByType` and
`Test_RunSync_EndToEndCreatesSlidesFromFreshSheet` (a real, full pass: a
template, a live cross-app Excel worksheet with one stale existing row and
two brand-new rows, confirming case 4 correction, case 3 creation ×2,
correct field injection, and post-sync resequencing all in one real run),
both passing via `run_vba_tests.ps1`. No separate manual recipe written.

`Test_RunSync_RunPeriodRolloverDuplicatesLeavingSourceUntouched` (added
2026-07-25 alongside `RunPeriodRollover` itself) confirms: the new period's
slide receives the injected new value; it's tagged with the new
instance_key; the source slide's own value is unchanged after the call
(the actual case-2-specific claim, not just "a duplicate was made"); and a
second rollover onto the same already-used instance_key is refused, not
silently double-created. **Not executed in this environment** --
`powershell.exe` is unreachable from this pass's plain Linux container
(confirmed: `which powershell.exe` fails), the same constraint Priority 20
of `IMPLEMENTATION_PLAN.md` already documented for this whole pass; needs a
real run via `run_vba_tests.ps1` on the WSL/Windows host next time this
project is picked up there.
