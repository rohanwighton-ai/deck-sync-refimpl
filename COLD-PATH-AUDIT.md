# Cold-path audit — the other six buttons, the shared utilities, and the
modules not yet read tonight

Companion to `HOT-PATH-AUDIT.md` (items AF-AJ) and `FIX-LIST.md` items
AL-AQ. Written 2026-08-17 evening, commissioned to check whether the same
defect class found in the two busiest buttons also lives in the rest of the
codebase. Produced by a fresh research pass (model: fable, read-only, no
code changes), reviewed here before acting on it.

**Gist:** the biggest thing found is that the Lobby's "pin on tick" watcher
(added yesterday) charges a toll on every single cell the tool itself
writes into the big spreadsheet — a toll of roughly a hundred slow calls
per cell — and most of the tool's own writing paths never switch that toll
off, even though one of them already knows to. Beyond that: the "add/retire
slides" buttons run the tool's slowest full-comparison machinery just to
answer a yes/no membership question, the review-sheet writer has no speed
wrapper at all, and the publish path re-searches the spreadsheet from the
top for every single row it writes.

**Scope and method.** Read every toolbar button not covered by
`HOT-PATH-AUDIT.md` (`CommandBarUI.bas`'s `AddButton` list is the
authoritative roster of what's reachable): "Review changes (writes
nothing)", "Add missing slides", "Retire slides with no row", "Tag fields
on this slide", "Change which workbook this deck uses", "Create template
slide" -- plus the shared utility modules (`WorkbookBridge.bas`,
`DeckRegistry.bas`, `ExcelOutput.bas`, `FieldSpec.bas`) and the modules not
yet read tonight (`BatchOnboardFlow.bas`, `AdoptFlow.bas`, `Harvest.bas`,
`MilestoneDevice.bas`, `InjectPrimitive.bas`, `TemplateAudit.bas`,
`DiscoverUI.bas`, `AppEvents.cls`). Cost figures reuse tonight's
calibration only: a full register read is 9.3s for 43 rows (~3-7ms per
cross-app COM call, cross-checked against item W's own numbers), unwrapped-
vs-wrapped fast mode is ~24-63x (items AE/AB), a verified save is 8.8s.

## Ranked findings

| # | Finding | Shape | Where | Weight |
|---|---|---|---|---|
| AL | Lobby pin watcher taxes every cell write; two of three hot write paths don't disable it | eager work behind a falsely-cheap guard | `AppEvents.cls:55`, `DraftingLobby.bas:83`, `DraftingUI.bas:595` | Minutes-class on the two most-pressed buttons |
| AM | `WriteQueueSheet` (review-sheet writer) has no fast-mode wrapper at all | AE's own omission, unfixed here | `ReviewQueue.bas:651-738`, `RibbonUI.bas:944-1063` | Every "Review changes" press AND every fresh cycle of button 2 |
| AN | `PublishDrafts` -> `UpsertRow` rescans the whole register per published row | AB's shape, second-hottest chain | `Drafting.bas:1386`, `ExcelOutput.bas:706-760` | Tens of seconds per field, x13 in a full publish, stacked on AF |
| AO | "Add/Retire slides" run the full register-vs-deck diff to answer a membership question, then re-walk the deck again | eager full work for a small answer | `RibbonUI.bas:681-870` | Minutes-class per press; a few times a quarter |
| AP | `KnownFieldNames` does a full register read per type on every "mark field" press during onboarding | repeated per-call cost, one-time flow | `BatchOnboardFlow.bas:1312-1346` | ~9s x marks-per-session; not worth fixing unless onboarding recurs |
| AQ | `MilestoneDevice` re-walks its shape group 3-5x per write | repeated computation | `MilestoneDevice.bas:142-436` | Sub-second total; found the shape, not worth the risk of touching |

AL-AO are real, cheap, low-risk fixes. AP-AQ are "found the shape, honest
answer is leave it."

## AL — the pin watcher's guard is a full spec-sheet read per write event

**The claim vs. the code.** `AppEvents.cls:49-54`: *"FieldIdForSheet's own
check is what keeps this cheap and safe: anything that is not a real
drafting sheet name exits in one comparison."* `DraftingLobby.bas:78-82`
repeats it: *"exits in one comparison per field, cheap enough to run on
every keystroke."* But `FieldIdForSheet` (`DraftingLobby.bas:83`) BEGINS
with `DraftingUI.ProseFields(wb)` -- which calls `WorkbookBridge.
WorksheetExists` (a `For Each` over all ~54 worksheets, cross-process),
then `GetOrAddWorksheet`, then a full scan of the Field Spec sheet at 2
cell reads per row. The name comparison the comments describe happens
AFTER all of that. Roughly 100-170 cross-app COM calls per event -- at the
calibrated 3-7ms each, **~0.3-1.2s per cell write**.

**Why this hits the hot buttons.** The watcher is wired in `WorkbookBridge.
OpenOrGetWorkbook`, live on every path that touches the register. Excel's
`SheetChange` fires for programmatic writes too, unless `EnableEvents =
False`. `RefreshDraftingSheets` already protects itself (`EnableEvents =
False`, restored after the Lobby rebuild). `ApplyApproved`'s fast-mode
wrapper sets `ScreenUpdating`/`Calculation` but **not `EnableEvents`** --
so every log line and review-sheet write during an apply pays the toll. At
Phase-3 scale (221 items) that's ~1,000 events, minutes-class.
`WriteQueueSheet` (AM) and `PublishAllDraftedFields`'s register writes
(AN) are unprotected the same way.

**Calibration caveat:** item W's own proof numbers were measured in a
fresh throwaway workbook with no spec sheet -- the handler exits almost
free there. That isolated number structurally excluded this exact cost;
the real register workbook pays it in full. Same shape as AE: the isolated
proof wasn't wrong, it was never exposed to this.

**Possible bearing on item X** (the still-unexplained multi-minute stall,
`Ctrl+Break` inert, from the live demo). A long chain of synchronous
cross-process handler invocations during a big write (`WriteQueueSheet`'s
~1,800 writes at Phase-3 scale) fits that signature. Not claimed as the
cause -- cheaply falsifiable: one Timing line around `WriteQueueSheet`,
once with events on, once with `EnableEvents=False`. If the wrong
hypothesis is true the two times match; if right, the gap is unmissable.

**Fix direction, three independent layers, cheapest first:**
1. Make the guard match its own comments -- check `sheetName` against the
   drafting-sheet shape BEFORE touching the workbook. One in-memory string
   comparison, zero COM.
2. Cache the sheetName->fieldId map per workbook (the `ShapeAddressBook`
   precedent, same hook point).
3. Add `EnableEvents=False` to `ApplyApproved`'s existing wrapper and to
   AF's restructure when it happens -- matching `RefreshDraftingSheets`'
   already-shipped discipline.

## AM — the review-sheet writer has no fast-mode wrapper at all

`WriteQueueSheet` does 8 `.Cells` writes per queue item plus banner,
headers, and nine formatting operations. Its caller `ReviewChangesCore`
manages nothing -- no `ScreenUpdating`, `Calculation`, or `EnableEvents`.
Runs on every "Review changes" press AND on every "2. Put it on the
slides" press that starts a fresh cycle (`pending = 0` delegates straight
to `ReviewChangesCore`) -- the first press of every publish round, on the
most-pressed button. Item AE measured exactly this omission on this
workbook at ~24-63x. Fix: the AE-pattern wrapper around `ReviewChangesCore`'s
type loop, `EnableEvents` included, same error-path restore discipline
`ApplyApproved` already has. This exact fix has shipped three times
tonight; risk is low, the pattern is proven.

## AN — `UpsertRow` inside the publish loop is item AB's shape, in the
second-hottest chain

`PublishDrafts` calls `ExcelOutput.UpsertRow` once per ticked row. Every
call re-runs from scratch: `LocateStructuralColumns` (a header scan),
`FindOrAppendInstanceRow` (compares every register row until the instance
matches), `FindOrAppendFieldColumn` (scans every header cell) -- **for a
fieldId that is constant across the entire loop**, since publish writes
one field per call. For a 43-row field: ~3,000+ cross-app reads per field,
~10-20s at calibrated rates, x13 fields in a full chain publish -- stacked
on AF's own ceremony, and each write also pays AL's toll. This is
`BuildLobbyFromScratch`'s pre-AB structure: a find-my-row scan inside a
loop whose caller already read every instance key moments earlier.

**Fix direction:** hoist the two constant lookups out of the row loop,
build a key->rowNum dictionary once per field from the read the caller
already paid for. `UpsertRow` is shared with `Harvest.bas` and
`DeckAdoption.bas`, each with different correctness assumptions -- don't
reshape the shared function blind; do it in the caller or via an optional
precomputed index. Natural to fold into AF's restructure, same function
family.

## AO — deck membership answered by running the tool's most expensive diff,
twice over

"Add missing slides"/"Retire slides with no row" call `ReviewQueue.
BuildQueue` per type but consume only four counters (orphan/no-row counts
and keys). `BuildQueue` earns those counts by doing the whole job: a full
deck walk, then `PlanRoutineSync`, which runs an `InjectField` DRY PROBE
per field per row against the live slides -- the multi-minute-class
operation item X implicated on the real deck, all of it discarded except
four counters. Then a second function walks `pres.Slides` again,
re-resolving every slide -- the same membership facts `BuildQueue`'s own
parity loop just computed, this time as objects.

Frequency is honest-per-quarter, not per-session -- but pressing "Add
missing slides" at quarter start on the real deck plausibly costs minutes
to print "3 rows have no slide."

**Fix direction:** extract the parity loop already inside `BuildQueue`
(instance keys resolved from slides, dictionary-tested against the
register) into its own membership-only scan that stops before the
per-field dry-probing. Keep the object-returning half for the delete path
-- its guards are pinned by an existing test and untouched by this fix.
Care level: moderate -- read-only until the existing consent gates, so the
blast radius of a bug here is a wrong COUNT, not a wrong delete.

## Found the shape, pressed rarely — honest to leave

- **AP, `KnownFieldNames`** -- full register read per type on every "mark
  field" press during onboarding. ~9.3s-class per mark. One-time-per-type
  flow; only worth fixing if more onboarding is actually coming.
- **AQ, `MilestoneDevice` group re-walks** -- ~5 walks of a ~37-shape group
  per device write, in-process PowerPoint COM, sub-second total at one
  device per project slide. Not worth the risk of touching a writer family
  with this project's incident history.
- **`SyncPreview`/`PreviewRoutineSync`** -- a second, parallel full-diff
  engine beside `BuildQueue`, reachable only from the standalone publish
  button's Yes/No offer. Not a live cost, barely reachable -- flagged as an
  AJ-adjacent consolidation candidate if that offer is ever rethought, not
  urgent.
- **`DiscoverFieldsOnSlide`** -- `BuildDiscoverySheet` writes ~7 cells per
  text shape with no wrapper, but gets fixed for free by AL's guard change
  (the Discovery sheet fails the name check only after the full guard cost
  today). No standalone work justified.

## Checked and clean — no action

`ChangePairedWorkbook`/`RepointWorkbookUI` (legitimate write-verification,
rare press). `CreateTemplateSlide` (bounded, rare). `TemplateAudit.
BuildAudit` (collects each instance's texts exactly once, in-memory
lookups after -- the module that looked likely to hide an O(n^2) doesn't).
`BatchOnboardFlow.BuildBatchPlanFromCandidates` (already fixed 2026-07-25,
documented in place -- the O(n^2) it once had is gone). `AdoptFlow`/
`DeckAdoption` (one-time, small scale). `WorkbookBridge.
FormatRegisterSheet` (looked like item W's shape, checked directly, is
O(n) not O(n^2), now inside AE's wrapper). `ArrangeTabs`/
`WriteWorkbookIndex` (bounded, inside AE's wrapper). `ReviewQueue.
ReadQueueSheet` itself (single pass, the load-bearing read -- see the note
below on how it's CALLED, which is a separate, already-tracked problem).
`ShapeAddressBook` (one scan per workbook attach, correct).

## Cross-cutting observations

1. **Two comments assert a cheapness the code never had** (`AppEvents.
   cls`, `DraftingLobby.bas` -- "one comparison"). Not stale, aspirational
   from birth -- written to describe the guard as designed, not as built.
   Exactly what stopped this being noticed: a reader who trusts the
   comment doesn't open `ProseFields` to check.
2. **The fast-mode wrapper is being rediscovered one call site at a time**
   (W -> Z -> AE -> AL/AM/AN here), and `EnableEvents` has now joined
   `ScreenUpdating`/`Calculation` as a third thing each site must remember
   separately. The shape-not-site fix: one `Timing.BeginFastMode(wb)`/
   `EndFastMode(wb)` pair owning all three plus the error-path restore,
   replacing three hand-rolled copies -- then a missing wrapper is one
   absent call, greppable, instead of three absent lines.
3. **Measure before believing AL's magnitude** -- the 0.3-1.2s/event range
   is derived from calibration, not measured directly. The check that can
   actually fail: time `WriteQueueSheet` with events on vs.
   `EnableEvents=False` on the real register. If the wrong hypothesis is
   true, the two times match.

## The larger question, again

Same closing caveat as `HOT-PATH-AUDIT.md`: these are real, mostly cheap,
low-risk fixes, and doing them is still speed work -- which the project's
own needs-vs-build comparison found isn't the dominant cost of a quarter.
AL and AM have the strongest case for an exception anyway: they tax the
exact chains Rohan presses while doing the real work, and AL is also a
correctness-adjacent cleanup (a live event handler with a falsely-
documented cost profile is exactly the kind of thing the next debugging
session pays to rediscover from scratch).

**UPDATE, same evening, real measurement, not projection:** Rohan pressed
"2. Put it on the slides" for real on `addin125` (all of AB/AC/AE/AK
already live). `PublishAllDraftedFields (total)`: **362.2 seconds for 4
fields — 90.6 sec/field**, read straight off the Timing sheet. That is
worse than AF's own estimate in `HOT-PATH-AUDIT.md`, and it is exactly the
path AL (`EnableEvents`) and AN (`UpsertRow`'s per-row rescan) both stack
onto, on top of AF's own per-field ceremony. This clears the standing PM
condition for fixing AF: the real number came back large, on the real
button, pressed by Rohan. AF, AL, and AN are now the justified next fix —
not speculation.
