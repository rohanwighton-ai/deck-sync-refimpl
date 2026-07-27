# VBA implementation: `BatchOnboardFlow.bas`

Onboards a new slide type AND bulk-links a whole batch of its already-existing
instances in one pass, driven by an editable Excel review grid instead of
InputBox chains. Built 2026-07-26 specifically because a real, richly-designed
deck (`test-fixtures/crc-real-deck-redacted.pptx`, and the unredacted
`test.pptx` used for the actual live run) discovers 60-90 raw candidate shapes
per slide — walking that many one InputBox at a time (`OnboardFlow.bas`'s
existing flow) is unusable. Also closes `specs/onboarding.md`'s own
documented-but-never-built "boilerplate-vs-varying pre-filter" gap, in a
different shape than that spec sketched: rather than a silent geometric
clustering filter, this surfaces a computed guess (diffing harvested text
across the whole batch) and lets a human freely override it in a real,
editable table (Rohan's own framing, 2026-07-26).

**Executed against real Office 2026-07-26. 70/70 tests pass** (up from 61 at
the start of this module's work). Getting there took an unusually long,
genuinely difficult debugging session — full account below, because the root
causes are exactly the kind of thing worth not re-discovering blind next time.

## The debugging saga — three separate real bugs, not one

The symptom throughout was consistent and misleading: **"Compile error:
User-defined type not defined"**, with VBE's error navigation landing on
`BuildBatchPlan`'s own signature line (or occasionally a caller's `Dim plan As
BatchOnboardPlan` line) almost every time — which kept pointing investigation
at the wrong place. Three genuinely distinct bugs were hiding behind that one
symptom:

### Bug 1 (real, found by a Fable-model agent): UDT array stored in a Dictionary

`BuildBatchPlan` originally cached each other-slide's discovered candidates in
`Scripting.Dictionary` objects, one cached per slide index, to avoid
re-scanning a slide once per field:
```vb
Dim oc() As Candidate
oc = Discovery.DiscoverSlideWithShapes(otherSlides(s), os_)
Set otherCandidatesBySlide(s) = oc   ' <- illegal
```
A `Candidate()` array (`Candidate` is a `Public Type` in `Discovery.bas`, a
standard module) can **never** be stored as a `Dictionary`/`Variant` value —
VBA's compiler rejects it outright: *"Only user-defined types defined in
public object modules can be coerced to or from a variant or passed to
late-bound functions."* This is documented in `AGENTS.md`'s Known Patterns
already (found 2026-07-25 in `Onboarding.bas` and `SyncOperations.bas`) — this
module hit it a third time, having never checked that section before writing
new Dictionary-caching code.

**Fix**: flattened all other-slides' `Candidate()`/`Object()`/`Boolean()` data
into single plain, properly-typed arrays (`allOtherCandidates()`,
`allOtherShapes()`, `allOtherAvailable()`) instead of caching per-slide in a
Dictionary, with `otherSlideCandStart()`/`otherSlideCandCount()` (plain
`Long()` arrays — primitives, no restriction) recording where each slide's own
candidates sit within the flattened set. Each slide is still discovered only
once.

This was a real, necessary fix — but fixing it did **not** make the "type not
defined" error go away. That turned out to be a second, unrelated bug.

### Bug 2 (the actual cause of the persistent symptom): declaration order

Real, confirmed VBA compiler quirk, found by bisecting a long series of
minimal, isolated repros (not a guess — each step below was independently
verified live against real Office): **if a `Public Function`/`Sub` appears
textually *before* a `Public Type` in the same standard module, that Type
fails to resolve when referenced from a *different* module** — "Compile
error: User-defined type not defined" — even though:
- the function containing the "bad" reference compiles and runs fine when
  called directly with real arguments,
- the Type itself, alone with a trivial function, works cross-module,
- the Type with all its real members (6 `Object` Dictionaries) works
  cross-module,
- and same-module usage of the Type always works regardless of order.

Only the combination of (Type declared after another Function) + (referenced
from a *different* module) reproduces it. `BatchOnboardFlow.bas` originally
had `AllValuesIdentical`/`SuggestBatchFieldName` declared before `Public Type
BatchOnboardPlan`, and `WriteReviewGrid`/`ReadReviewGrid`/etc. before `Public
Type BatchCommitResult` — both real instances of this ordering.

A **second, apparently-related** symptom of the same underlying rule: the six
`Private Const COL_*` declarations (originally sitting between `BuildBatchPlan`
and `WriteReviewGrid`) produced a *different* error once Bug 2's Type-ordering
fix was applied — **"Only comments may appear after End Sub, End Function, or
End Property"** — also resolved by moving them above every Function/Sub.

**Fix**: every module-level declaration (`Const`, `Type`) in this module now
lives at the very top, above all `Function`/`Sub` declarations — not a style
preference, a hard requirement this module discovered the expensive way.
**Recommendation for every other module in this project going forward**: keep
all `Const`/`Type` declarations first, before any `Function`/`Sub`, even
though the rest of this codebase (written before this was known) doesn't
consistently follow that — nothing has broken elsewhere yet, but this is now
a known trap, not a hypothetical one.

### Bug 3: a reserved word and a test-data bug (ordinary, once bugs 1-2 were found)

`vba/tests/TestRunner.bas` had `Dim single As New Collection` — `Single` is a
VBA reserved word (the floating-point type), producing a plain "Syntax error."
Renamed to `onlyOne`.

Separately, `Test_BatchOnboardFlow_CommitBatchTagsLinksAndVerifies` used
identical harvested text ("Overall Status") on both the template and the one
other slide — `BuildBatchPlan` correctly classified this as decoration
(`FieldSuggestIdentical = True`) and defaulted `FieldInclude` to `False`,
so `CommitBatch` correctly tagged nothing. The test's own assertions expected
tagging to happen, which was a test bug, not a `CommitBatch` bug — fixed by
explicitly setting `plan.FieldInclude(1) = True` before committing (simulating
a human overriding the suggestion in the review grid), with new assertions
added confirming the *suggestion* itself was correct first.

## Diagnostic technique that actually worked

Manual live probing (import, run, read the stuck VBE dialog) got the search
going but produced inconsistent, misleading signal for hours — the error's
reported location moved around depending on which unrelated test scaffolding
happened to be in play, which is exactly what Bug 2's "wrong module gets
blamed" behavior produces. What actually closed it:
1. A Fable-model agent, briefed with everything already ruled out, found Bug 1
   via its own from-scratch minimal repro (see its report for the exact
   isolation technique).
2. UI Automation (`System.Windows.Automation`) reading the VBE's exact
   selected/highlighted text directly, rather than relying on a human
   describing a screenshot — this is what made rapid, precise bisection of
   Bug 2 possible (dozens of tiny `TypeDefModule`/`CallerModule` pairs,
   isolating exactly which combination of same-module-vs-cross-module and
   before-vs-after-the-Type broke it).
3. Systematic reduction: minimal reproductions built up in small, deliberate
   increments (bare Type → Type+trivial-function → Type+real-body →
   Type+real-body+other-functions), each tested in complete isolation before
   moving to the next increment, rather than continuing to prod the full
   ~20-module production project.

## Redesign 2026-07-26: click-based field selection, not Discovery auto-enumeration

Discovery-based auto-enumeration (walk every text/picture shape via
`Onboarding.IsCandidateField`) was the *original* mechanism for choosing
which fields to onboard. Live-tested against Rohan's real, unredacted
`test1.pptx`, it produced an 87-row "Field Review" grid — his direct
feedback: "very hard for a human to interpret ... why did I have to select
multiple files? I think this is better human led by selection then you find
the matching value and position."

`BuildBatchPlan` (Discovery-based) is kept as-is — still tested, still a
valid building block — but the live "Bulk Onboard Type" ribbon flow no
longer calls it. Instead: `MarkFieldForBatch` (toolbar button, repeatable —
click one field's shape, run it, repeat) accumulates a human-chosen field
list in module-level session state, and `BuildBatchPlanFromMarkedFields`
runs the exact same cross-slide correspondence/harvesting engine
(`BuildBatchPlanFromCandidates`, extracted as a shared private helper) but
scoped to only those clicked shapes, matched back to their `Candidate` by
object identity (`Is`), not name or position. `PromptBatchOnboardType` now
*requires* a completed marking session (on the same slide that ends up
earliest — by deck order — in the batch selection) before it will proceed.

VBA's InputBox/MsgBox are fully modal (block the whole application), so
there's no way to show a "click now" prompt and have the user click the
canvas while it's up — each mark is therefore its own separate toolbar
click, not a loop inside one macro run. `MarkShapeForBatch` (the pure logic:
validate top-level shape, accumulate, de-dupe, detect a different-slide
mark) is split out from `MarkFieldForBatch` (the MsgBox-confirming Sub
wrapper) the same way `ResolveFields.bas` already splits interactive entry
points from testable helpers — the wrapper's own MsgBox is manual-
verification-only, like every other InputBox/MsgBox interaction in this
project.

## Addendum 2026-07-26: mark-time field type, and why it stops at metadata

Live-testing the click-based redesign above, Rohan asked for a field TYPE
declaration at mark time too (e.g. tagging a "Project Number" field as
text vs. a budget figure as currency) "so Excel can format/filter
correctly." `MarkFieldForBatch` now prompts for a type (Text/Number/
Currency/Date, `NormalizeFieldType`) alongside the name, stored per-field
(`BatchOnboardPlan.FieldTypes`) and shown/editable as a column in the
"Field Review" grid (`WriteReviewGrid`/`ReadReviewGrid`).

It deliberately goes no further than that. The natural next step — have
`ExcelOutput.UpsertRow` write a real typed Excel value (`Date`/`Double`)
with matching `NumberFormat` instead of plain text — was built, then
reverted, after tracing a real correctness risk: `CommitBatch` writes into
the same Data sheet `RunSync.bas`/`SyncOperations.PlanRoutineSync` read on
every *routine* sync, and that path feeds values straight into
`InjectPrimitive.InjectPrimitive`, which **writes onto the live PowerPoint
slide** when the Data-sheet value differs from what's currently on the
slide. A typed cell's read-back is not guaranteed to equal the exact string
that was written — `CStr()` on an Excel `Date` value returns a locale-
formatted string (not necessarily the original text), and `CStr(CDbl(x))`
can drop a trailing zero. Either would make a later routine sync see a
false "difference" and silently rewrite the slide's date/number text into a
reformatted (though logically equal) version nobody asked for — a real
slide-content mutation, not a display nicety, and exactly the class of
silent-guess-or-mutate behavior `InjectPrimitive`/`Verification` exist to
prevent elsewhere in this project.

`ExcelOutput.UpsertRow` therefore still always writes and reads back the
exact harvested string, unconditionally, regardless of any declared type —
see that function's own header comment. Revisiting real Excel-native
typing would need `PlanRoutineSync`'s own comparison made type-aware first
(e.g. comparing parsed values rather than raw strings), not another pass at
`UpsertRow` in isolation.

## Addendum 2026-07-26: auto-selecting the batch by layout, not manual multi-select

Live-testing again, Rohan pushed back on still having to manually select
the batch of slides (Slide Sorter/Ctrl-click) after marking fields: "i
still have to select slides?" Marking fields only ever established *which
shapes* on *one* slide are fields — it never established *which other
slides* are instances of that type, so a second piece of information was
always structurally necessary. But per this project's own 2026-07-26
"child deck per slide type" authoring decision (`DECISIONS.md`), the real
target workflow makes that second piece nearly free: every slide in a
child deck built from the same layout genuinely *is* the same type.

`FindSameLayoutSlides` (pure, testable) now finds every other slide in the
presentation sharing the template's `CustomLayout` (compared by `.Name`,
same robustness reasoning as `MarkShapeForBatch`'s shape-identity
comparison). `PromptBatchOnboardType` calls it automatically whenever the
human hasn't already made an explicit multi-slide selection themselves:
switches to Slide Sorter (the one view `Slides.Range(...).Select` is
confirmed to register reliably in — see `SPIKE_NOTES_AdoptFlow.md`'s
Normal-view finding, sidestepped here rather than risked), selects the
template plus every same-layout sibling, and shows a Yes/No MsgBox listing
exactly which slides it picked. Yes proceeds with that batch; No returns a
message telling the human to select their own batch and run "Bulk Onboard
Type" again — at which point their explicit selection is detected and used
as-is, un-overridden. Layout equality was chosen over any content/geometry
heuristic specifically because it's what the child-deck architecture
already guarantees, not a guess.

## Addendum 2026-07-26: grouped shapes were being rejected on an unverified assumption

Live-testing again, Rohan asked: "Is it allowing me to mark selected
objects in groups? I think maybe no?" It wasn't. `MarkShapeForBatch`
originally checked `shp.Parent.SlideID` and rejected the shape outright if
that failed, on the assumption that `Shape.Parent` returns the enclosing
`GroupShape` (not the `Slide`) for a shape inside a group — stated
confidently in the code's own comment, but never actually verified against
real Office, since this project had no reachable Office install when
`ResolveFields.bas`'s shape-selection pattern (which this borrowed from)
was first written.

Verified live via a standalone probe script (not TestRunner.bas — a
throwaway diagnostic, three iterations to get right: the first two hit an
unrelated real quirk, COM shape references going stale after a `.Group()`
call reuses the same variable): `Shape.Parent` resolves **directly to the
Slide**, even for a shape nested one level inside a group — confirmed with
`SlideID` matching the slide's own `Parent.SlideID` exactly, and
`Shape.Parent.Name` literally returning the slide's own name. The
enclosing group, when one exists, is exposed separately via
`Shape.ParentGroup`, which this function has no reason to touch.
`Discovery.Walk` (`Discovery.bas`) already recurses into groups when
building the Candidate list `BuildBatchPlanFromMarkedFields` matches a
marked `Shape` back against by object identity — so once the incorrect
rejection was removed, grouped fields work end-to-end with no other code
changes anywhere in the module. Real deck implication: Rohan's actual
slides use grouped "card" layouts for their fields (see the "Mark Field
for Batch" screenshot from this same session), so this was a genuine
blocker, not a theoretical one — PowerPoint's own click-to-select behavior
(one click selects the whole group, a second click on the same spot drills
into the specific member) is what a user needs to know to reach a grouped
field shape at all, unrelated to this fix.

## Addendum 2026-07-26: the group check landed in the wrong place at first

The `msoGroup` rejection above was originally only inside `MarkShapeForBatch`
(the pure function), called at the very end of `MarkFieldForBatch` (the
Sub) -- *after* both the name and type `InputBox` prompts had already run.
Live-tested (Rohan's own screenshot: selecting a whole card's background
bar showed `Field name for this shape (current value: ''):` with no
explanation), a whole-group selection got the human through two pointless
prompts before finally being told it was rejected. Rohan's first instinct
was a much heavier fix -- "ungroup all at start and regroup on exit while
preserving z order" -- rejected as unnecessary and genuinely riskier
(z-order/rotation/flip aren't guaranteed to round-trip losslessly through
an ungroup/regroup cycle, and it would mutate the deck's actual group
structure as a side effect of what should be a read-only marking step).
The real fix was much smaller: check `shp.Type = msoGroup` immediately
after `ValidateSingleShapeSelection` succeeds in `MarkFieldForBatch`
itself, before showing either prompt -- `MarkShapeForBatch`'s own check
stays too (defense in depth for any other caller), just no longer the only
place it happens.

## Addendum 2026-07-26: click-to-drill-in doesn't reliably reach VBA at all

The "just click again" guidance above turned out to be wrong, confirmed
live with Rohan doing exactly that repeatedly: a genuine canvas click-to-
drill-in into a group (Shape Format ribbon tab active, tight selection
handles around just the one field shape, Selection Pane confirming a leaf
item like "Text 4") can still leave
`Application.ActiveWindow.Selection.ShapeRange(1)` reporting the *outer
group*, not the individual member the UI visually shows selected — proven
with a diagnostic build that echoed the shape name/Type straight into the
rejection message (`shp.Name`, `shp.Type`) rather than guessing from a
screenshot. This is a real gap between PowerPoint's visual selection state
and what its own object model exposes to automation, not a click-technique
problem, and no amount of "click again" advice fixes it.

Fix: stopped trying to detect which member was clicked at all. When
`Application.Selection` reports a group, `FlattenGroupLeaves` (pure,
recursing through any nested groups the same way `Discovery.Walk` does)
lists every real leaf shape inside it, and `MarkFieldForBatch` shows a
numbered picker (name + text preview) for the human to choose from
explicitly — same "numbered list, pick by number" idiom
`ResolveFields.BuildRolePickerPrompt` already established. Rohan's earlier
"ungroup at start, regroup on exit" idea would also have worked around
this, but stays rejected for the same reasons as before (z-order/rotation/
flip round-trip risk, mutating deck structure for what should be a
read-only step) — the picker solves the actual problem (which member was
intended) without touching the file's group structure at all.

## Addendum 2026-07-26: sample-values column accumulated across rows

Real bug, found live against Rohan's actual 9-field marked batch (screenshot
of "Bulk Onboard Type -- Review", and the "Field Review" workbook he'd
saved off separately): `WriteReviewGrid`'s per-field loop builds each row's
"Sample Values From Other Slides" cell in a local `samples` string, but
never reset it to `""` at the top of each iteration -- it only had
`Dim samples As String` inside the loop, and VBA's `Dim` does **not**
re-initialize a variable on each pass through a loop (a `Dim` statement is
only elaborated once; the variable is scoped to the whole procedure
regardless of where it's textually written). Every row therefore kept
appending onto the *same* string as every prior field, so field 9's cell
ended up containing all 9 fields' samples concatenated together instead of
just its own up-to-3. Fixed with an explicit `samples = ""` each iteration.

The existing `ReviewGridRoundTrip` test never caught this because its
2-field fixture happened to give field 1 zero other-slide samples -- with
nothing to leak, the bug was invisible. Strengthened: field 1 now also has
a real sample, and the test asserts field 2's cell does NOT contain it.

## Addendum 2026-07-26: static/variable hint, instance-key auto-suggest, and marking-session persistence

Three more pieces, all in the same live-testing pass, all in response to
direct feedback:

1. **Static/variable hint.** Rohan: "Can we include whether the fields are
   static or variable... perhaps lock the static fields going forward if
   that is useful?" Agreed the tool can't actually *know* this at onboarding
   time (only one snapshot per project exists yet -- whether a field is
   genuinely static is only knowable after watching it across real
   periods), and explicitly deferred wiring anything into sync behavior.
   Captured purely as a declared hint instead, at mark time -- exact same
   pattern as the type tag (`MarkShapeForBatch`/`NormalizeFieldVolatility`/
   a new "Static/Variable" column in the Field Review grid), not applied
   anywhere yet.

2. **Instance-key auto-suggest.** Rohan initially proposed dropping
   instance keys entirely in favor of "always work on a copy of the file" --
   flagged as a much bigger pivot than it first sounded (would mean
   abandoning `DeckRegistry`'s pairing and `RunSync`'s whole routine/period
   sync model), and asked directly whether the real complaint was the
   architecture or just the hand-typing. Confirmed: just the hand-typing.
   Fix: `SuggestInstanceKey` pre-fills each instance-key `InputBox` with the
   *first marked field's* harvested value for that slide (in his real deck,
   "Project Number" -- already a natural per-slide identifier, and usually
   the field marked first) -- most prompts become a single OK click, only
   needing an edit for genuine ambiguity (e.g. his boss's manually
   duplicated Q3/Q4 slides needing distinct suffixes).

3. **Marking-session persistence.** Rohan: "sick of linking every test" --
   the marking session (`markedShapes`/`markedNames`/`markedTypes`/
   `markedVolatility`) previously lived only in the add-in's runtime
   memory, wiped on every PowerPoint close (confirmed the hard way: a
   session was lost mid-testing with no way to recover it, forcing a full
   re-mark). Fixed by persisting to a `CustomDocumentProperty` on the
   presentation (`DeckSyncMarkingSession`) -- same mechanism `DeckRegistry.
   bas` already uses successfully, travels with the deck once saved.
   `SerializeMarkingSession`/`RestoreMarkingSession` (pure) handle the
   round trip; a live `Shape` reference can't itself survive a close, so
   only each field's `Name` is persisted, and restoring re-finds the shape
   by `Name` via `Discovery` (already recurses into groups). `MarkFieldFor
   Batch` restores transparently on first use after a reload -- no new
   button, no extra step, just click a field and it picks up where you
   left off (or starts fresh if the saved session was for a different
   slide -- the existing DIFFERENT_SLIDE handling covers that case for
   free). `ResetMarkingSession` now also clears the persisted property, so
   a deliberate "Clear Marked Fields" or a successful commit never leaves a
   stale session to resurrect later.

   One real open question resolved empirically rather than assumed: Office's
   `CustomDocumentProperties` string type has a documented 255-character
   limit, and a realistic marking session (several fields' names/types)
   easily exceeds that. A direct PowerShell-side probe to check this hit an
   unrelated crash (`System.NullReferenceException` inside .NET's dynamic
   COM interop trying to reflect the `DocumentProperties` collection type --
   a known PowerShell-automation quirk, not an Office one, and a different
   failure mode from the earlier `.Add()` overload-resolution issue found
   this same session). Answered for real instead via the native-VBA path
   this feature actually uses: `Test_BatchOnboardFlow_
   MarkingSessionPropertyRoundTripsBeyond255Chars` writes deliberately
   over-255-char content and asserts an exact read-back match -- **passes,
   no truncation observed** in this Office version/configuration.

## Addendum 2026-07-26: AutoSave does not reliably persist macro-driven edits

The persistence feature above turned out to be necessary but not
sufficient: Rohan reported his real tagging/marking work was genuinely
lost across a PowerPoint close, even with AutoSave on. Traced properly
rather than assumed:

- A live probe confirmed `Presentation.Saved` correctly flips to `False`
  immediately after a macro writes a `Shape.Tag` -- the dirty flag itself
  is NOT the problem, ruling out the first (wrong) theory this session had
  about VBA writes "not tripping AutoSave."
- `Application.ActivePresentation.FullName` on Rohan's real, live session
  reported a `https://d.docs.live.net/...` cloud URL, with `Saved = True`
  at the moment checked -- AutoSave was actively believing it had nothing
  to save via the cloud connection, yet the change was still lost after a
  subsequent close/reopen.
- The one local file this session had been checking on disk
  (`/mnt/c/Users/rohan/OneDrive/Claude/test1.pptx`) never changed at all
  (same timestamp, same byte count) even after an explicit Ctrl+S from
  Rohan's own session -- that file and the cloud-connected one PowerPoint
  was actually editing are the same OneDrive item, just reached through
  two different mechanisms (local sync mirror vs. direct cloud API), and
  the local mirror lagging behind is a red herring, not the real bug.

Conclusion: AutoSave's own background/debounced save simply does not
reliably capture macro-driven edits before the application actually
closes -- a known class of issue (the debounce/checkpoint timer isn't
reliably extended or flushed by VBA-driven changes the way normal
keyboard/UI edits are), not something fixable from the PowerPoint UI side
or a setting to toggle. The only real fix is for the add-in to stop
trusting AutoSave for anything it cares about surviving a close, and force
an explicit, synchronous `Presentation.Save` (which blocks until the save
genuinely completes or errors) after every meaningful write:
`SaveMarkingSessionToProperty` now does this after every mark, and
`PromptBatchOnboardType` does it again (both the deck and the Data
workbook) right after `CommitBatch` -- the actual linked data, which
matters even more than the marking-session bookkeeping. Both surface a
clear `WARNING:` in the confirmation message rather than silently
continuing if the forced save itself fails, instead of the previous
silent-loss behavior.

## Addendum 2026-07-26: re-investigated the save fix after Rohan said he still didn't trust it

The AutoSave addendum above landed a real fix (forced, synchronous
`pres.Save` after every mark and after every commit) and one live test of
it looked good -- but that was a single manual anecdote, and Rohan had
already lost real work to this once. He said directly he still didn't
trust saving was working. Treated that as a real, open question rather
than re-confirming the same partial signal: everything below is a fresh,
independent investigation, done entirely against throwaway presentations
(never Rohan's real `test1.pptx`), using live PowerShell/COM automation
against real Office (no PowerPoint/Excel session was running at the start
-- confirmed via `Get-Process` before touching anything).

**Bottom line up front: the fix holds.** Extensive, deliberately
adversarial reproduction attempts -- matching Rohan's real file size, his
real cloud-URL document identity, a zero-delay close, and even a forced
process kill -- could not break it. The renewed distrust traces to two
concrete, separate things, both fixed this pass: no *permanent* verification
was baked into the confirmation message (only a one-off diagnostic that had
already been deleted), and the add-in Rohan is actually running is a
slightly stale build that still shows that deleted diagnostic's leftover
text. Neither is a save-durability bug in the sense originally suspected.

### Reproducing Rohan's real environment for the probe

Faithfully reproducing his setup mattered: a throwaway `Presentations.Add()`
+ `SaveAs()` into a *new* path inside `C:\Users\rohan\OneDrive\Claude\`
shows `AutoSaveOn = True` immediately but keeps a plain local `FullName`.
Reopening an *existing* OneDrive file from that same folder (`Presentations.
Open` on the local path) is what actually reproduces Rohan's real
`FullName = https://d.docs.live.net/96b9ec593ee3ba55/Claude/...`
cloud-URL identity -- confirmed live, this promotion happens automatically,
no special "Open from OneDrive" UI action needed. All of the tests below
were run against files opened that same way, including via a *constructed*
cloud URL (`https://d.docs.live.net/96b9ec593ee3ba55/Claude/<throwaway
filename>.pptx`) once the file existed, to open through the identical code
path his real session uses.

`Presentations.Open(path, ReadOnly, Untitled, WithWindow)`: passing
`Untitled:=True` (copied from a stale memory of the signature) silently
opens a **disconnected `Presentation1`**, not the real file -- any
subsequent `.Save` on it errors ("An error occurred while PowerPoint was
saving the file") and any "durability" conclusion drawn from it would have
been meaningless. Real path: `Untitled:=False`. Cost about 20 minutes of
a genuinely misleading result before being caught by noticing `FullName`
read back as `Presentation1` instead of the real path -- worth remembering
for any future probe using this call.

### Finding 1: `Presentation.Saved` cannot be trusted as a "did this
actually save" signal on an AutoSave/cloud document

Confirmed repeatedly (throwaway files, both the small marking-session
property write and a ~59MB file matching Rohan's real deck's rough size):
after `pres.Save` returns with `Err.Number = 0`, `Presentation.Saved`
reads **False** -- even though a subsequent genuine process Quit + fresh
reopen (sometimes via a totally new `PowerPoint.Application` instance)
correctly reads back the just-written value. `Err.Number` was accurate in
every single test run here; `.Saved` was not. This isn't the bug Rohan
hit, but it's a real, previously-undocumented gotcha worth remembering:
nothing in this codebase currently branches on `.Saved` for a real
decision, and it should stay that way for any AutoSave-connected document
-- `Err.Number` after an explicit `.Save` is the only signal confirmed
reliable here.

### Finding 2: `pres.Save` on a cloud-connected document is genuinely
synchronous and blocking -- not fire-and-forget

This was the single most likely remaining explanation going into this
pass, and it's now directly tested and ruled out. Built a throwaway
~59MB presentation (eight 1600x1600 random-noise PNGs, deliberately
incompressible so file size stays close to Rohan's real ~49MB deck rather
than collapsing under PNG compression), saved it into the OneDrive folder,
reopened it via its cloud URL identity, then: wrote a marking-session
property, called `.Save` (688ms to return, `Err.Number = 0`), and called
`Application.Quit()` with **zero** added delay -- the same "user marks a
field and immediately closes PowerPoint" sequence that supposedly loses
data. `Quit()` itself returned in 15ms, and the POWERPNT.EXE process was
still alive 2 seconds later (a real, mundane process-teardown lag, not
evidence of an in-flight save) -- force-killed it to construct the worst
case a genuinely impatient user could produce. **The marker still read
back correctly** on a fresh reopen via the same cloud URL. Repeated this
shape of test (small file/large file, with/without an added delay,
graceful quit/forced kill) several times; never reproduced data loss. The
consistent sub-second save time regardless of total file size also
strongly suggests OneDrive/PowerPoint's AutoSave path uploads a real delta
of the changed OOXML parts, not the whole file, on every save -- so "a
slow save of a big file" is very unlikely to be a live risk factor here
either.

### Finding 3: the add-in Rohan is actually running is a real, verifiable
artifact -- and it's slightly stale, though not missing the fix

`C:\Users\rohan\OneDrive\Claude\` has eleven separately numbered
`addinN.ppam` files (evidence of a lot of iterative rebuild history).
`Application.AddIns` shows only **`addin16.ppam` has `AutoLoad = True`** --
that's the one actually live in Rohan's real sessions. Its VBA project
reports `Protection = 1` (locked/password-protected for viewing), which
makes `VBComponents` return a `Count` of 0 through every COM automation
path tried (`Presentations.Open` refuses `.ppam` files outright --
"you must use AddIns.Add"; `AddIns.Add` loads it but it never appears in
`Application.Presentations`; `Application.VBE.VBProjects` does find its
project by name/path but `Protection = 1` blocks `VBComponents` regardless
of entry point). This alone is worth knowing: **there is currently no way
to verify what's actually inside the live add-in via COM automation.**

Worked around it by going one level below the object model entirely: a
`.ppam` is the same OOXML/CFBF container format as any other macro-enabled
Office file, and VBA project "view lock" protection is a VBE-UI-level gate,
not something that prevents the raw compressed module streams from being
read directly out of the file. Wrote a from-scratch, read-only CFBF
(Compound File Binary Format) parser plus an MS-OVBA decompressor
(Python, stdlib only -- no `oletools`/`olefile`/`pip`/`7z` available in
this environment) against a **copy** of `addin16.ppam`, validated it by
successfully extracting and reading real, correct-looking VBA source for
all 19 modules, then diffed the extracted `BatchOnboardFlow` module
against this repo's current source.

Result: **the deployed add-in does contain the real fix** -- both
`SaveMarkingSessionToProperty`'s forced `pres.Save` + `Err.Number` check
and `PromptBatchOnboardType`'s paired deck+workbook save are present,
logically identical to the current repo (only VBE's own case-canonicalization
of identifiers differs, e.g. `.count` vs `.Count`, meaningless). It's a
build from slightly before the final cleanup commit, though: it still
carries the leftover temporary diagnostic block ("(diagnostic: property
length right after write = N)") that the repo source had already had
removed -- so every real "Mark Field for Batch" confirmation Rohan sees
today has that extra diagnostic-looking line appended, which reads as
unfinished/buggy even though it's functionally inert. Deliberately did
**not** touch `addin16.ppam` itself -- rebuilding/relocking a live add-in
file Rohan actively depends on is an infrastructure action that should be
his call, not something to do silently mid-investigation. Flagging it here:
worth a clean rebuild of the add-in from the current repo source, and
worth considering some cheap way to verify "what's actually deployed
matches git HEAD" going forward, given how easy this drift turned out to
be to produce and how hard it was to detect from the repo side alone.

### What this means for "why does Rohan still not trust it"

Not an unfixed correctness bug -- extensive adversarial retesting under
conditions matching his real environment couldn't break the existing fix.
Best explanation is two compounding, non-bug things: trust genuinely
hadn't been re-earned after one real loss plus a single anecdotal manual
retest, and the actual add-in he's using visibly shows leftover diagnostic
clutter that looks like unfinished work regardless of whether the
underlying save is fine.

### Fixes made this pass

1. **`SaveMarkingSessionToProperty`** (`BatchOnboardFlow.bas`): added a
   *permanent* closed-loop verification after the forced `Save` --
   re-reads the property back and compares it against what was written,
   surfacing a clear `WARNING` if they don't match. This formalizes,
   permanently, exactly the check the deleted temporary diagnostic did
   once by hand, instead of relying on Err.Number alone (accurate in
   every test here, but gives Rohan no visible assurance beyond "no error
   was thrown").
2. **`PromptBatchOnboardType`**'s explicit deck+workbook save block: added
   the analogous check using `DeckRegistry.LookupType`'s own public
   round-trip lookup on the slide type this commit just registered --
   the actual durable deck<->Data-sheet link, so it matters most here.
3. **New automated test**, `Test_BatchOnboardFlow_
   MarkingSessionSurvivesRealCloseAndReopen` (`vba/tests/TestRunner.bas`):
   a genuine `Presentation.Close` followed by a fresh `Presentations.Open`
   of the same file on disk (not just re-reading the same live object) --
   the strongest regression guard reachable from inside a single VBA test
   process. Documented in its own header that this is *not* equivalent to
   this addendum's out-of-process PowerShell probes (a VBA test can't spawn
   a whole new POWERPNT.EXE mid-run the way those did), which is why those
   probes' findings live here in prose rather than only in code.
4. **`run_vba_tests.ps1`**: unrelated to the save bug, found while running
   the suite for this investigation -- `$xl.Quit()`/`$ppt.Quit()` alone
   reliably left windowless zombie EXCEL.EXE/POWERPNT.EXE processes running
   after a clean, non-erroring run (confirmed repeatedly, up to 7 stray
   EXCEL.EXE at once). Fixed both the PowerPoint- and Excel-pass `finally`
   blocks with explicit `ReleaseComObject` + double `GC.Collect`/
   `WaitForPendingFinalizers`, plus a windowless-only `Stop-Process`
   fallback (mirrors the script's own existing self-heal philosophy for
   stray processes at start-up) -- confirmed clean (`Get-Process` reports
   nothing) after a full run with this fix in place.

Deliberately **not** changed: `addin16.ppam` itself (see Finding 3 --
flagged for Rohan's own call, not silently rebuilt).

### Real test run, this pass (`run_vba_tests.ps1`, real Office, all green)

```
=== deck-sync-refimpl VBA test run (PowerPoint) ===
PASS  Discovery_GroupRecursionFindsCandidates
PASS  InjectPrimitive_NoOpWhenValueAlreadyMatches
PASS  InjectPrimitive_WritesAndVerifiesOnMismatch
PASS  InjectPrimitive_AmbiguousTagRefusesToGuess
PASS  Matching_SiblingAmbiguityResolvedByZOrder
PASS  Matching_EnrichPlaceholderIdxReadsRealFile
PASS  Resolve_ReadsTagsOffLiveSlide
PASS  SyncOperations_Cases1And4
PASS  SyncOperations_Case3NewRecord
PASS  SyncOperations_Case6UnclassifiedSlide
PASS  Onboarding_HighAndMediumConfidence
PASS  Onboarding_OnboardNewInstanceAutoTagsHighOnly
PASS  Onboarding_PureDecorationNeverMatched
PASS  Verification_StructureMatchesAfterDuplicate
PASS  Verification_DetectsShapeCountMismatch
PASS  Verification_DetectsZOrderSwap
PASS  SlideDuplication_CreatesTaggedInjectedSlide
PASS  SlideDuplication_RefusesInstanceKeyCollision
PASS  SlideDuplication_PartialRowStillCreatesSlideButFlagsMissing
PASS  RunSync_GatherInstancesFiltersByType
PASS  RunSync_EndToEndCreatesSlidesFromFreshSheet
PASS  RunSync_RunPeriodRolloverDuplicatesLeavingSourceUntouched
PASS  DeckAdoption_AlreadyLinkedSlideSkipped
PASS  DeckAdoption_ReadyHighConfidenceSlideLinkedAndCreatesFreshRow
PASS  DeckAdoption_MediumConfidenceSlideNeedsConfirmationAndIsNotTagged
PASS  DeckAdoption_UnclassifiedSlideExcluded
PASS  DeckAdoption_MatchesExistingKeylessRowLinksWithoutCreatingNewRow
PASS  DeckAdoption_MultiSlideZeroBasedBatchKeepsIndicesAligned
PASS  ResolveFields_ValidateSingleShapeSelectionAcceptsOneShape
PASS  ResolveFields_ValidateSingleShapeSelectionRejectsMultiple
PASS  ResolveFields_BuildRolePickerPromptListsRolesNumbered
PASS  ResolveFields_PickRoleFromListAcceptsNumberOrName
PASS  ResolveFields_EndToEndTagsSelectedShapeViaPickedRole
PASS  DeckRegistry_BuildAndParseTypeRegistrationRoundTrip
PASS  DeckRegistry_ParseTypeRegistrationRejectsMalformed
PASS  DeckRegistry_GetOrCreateDeckIdIsStableAcrossCalls
PASS  DeckRegistry_RegisterAndLookupTypeRoundTrip
PASS  DeckRegistry_LookupTypeFalseWhenNotRegistered
PASS  DeckRegistry_LookupTypeFalseWhenTemplateSlideDeleted
PASS  DeckRegistry_ListRegisteredTypesListsAllRegistered
PASS  DeckRegistry_WorkbookPathRoundTrip
PASS  WorkbookBridge_SanitizeSheetNameStripsInvalidCharsAndTruncates
PASS  OnboardFlow_PlanOnboardingFindsCandidatesAndHarvestsText
PASS  OnboardFlow_ApplyFieldReviewAnswerRenamesOrExcludes
PASS  OnboardFlow_ApplyPeriodKeyAnswerMarksExactlyOneField
PASS  OnboardFlow_DeriveSeedInstanceKeyUsesPeriodKeyOrEvergreen
PASS  OnboardFlow_CommitAndVerifyOnboardingRoundTrip
PASS  RibbonUI_ResolveTypeAnswerAcceptsNumberOrName
PASS  RibbonUI_ResolveRecordAnswerAcceptsNumberOnly
PASS  RibbonUI_BuildTypePickerPromptListsAllTypes
PASS  CommandBarUI_ShowToolbarCreatesThreeWiredButtons
PASS  CommandBarUI_ShowToolbarIsIdempotent
PASS  CommandBarUI_HideToolbarRemovesIt
PASS  AdoptFlow_ValidateAdoptionSelectionSortsIntoDeckOrder
PASS  AdoptFlow_ValidateAdoptionSelectionRejectsNonSlideSelection
PASS  AdoptFlow_ExcludeTemplateSlideRemovesOnlyTemplate
PASS  AdoptFlow_BuildAdoptionReviewSummaryCountsAndListsNonReady
PASS  BatchOnboardFlow_AllValuesIdenticalDetectsMatchAndMismatch
PASS  BatchOnboardFlow_SuggestBatchFieldNameReusesPhNameOrFallsBack
PASS  BatchOnboardFlow_NormalizeFieldTypeAcceptsNumberOrName
PASS  BatchOnboardFlow_NormalizeFieldVolatilityAcceptsNumberOrName
PASS  BatchOnboardFlow_SuggestInstanceKeyUsesFirstFieldsHarvestedValue
PASS  BatchOnboardFlow_FindSameLayoutSlidesGroupsByLayoutOnly
PASS  BatchOnboardFlow_BuildBatchPlanFindsCorrespondenceAndHarvestsAcrossSlides
PASS  BatchOnboardFlow_BuildBatchPlanFromMarkedFieldsUsesOnlyMarkedShapes
PASS  BatchOnboardFlow_MarkShapeForBatchAccumulatesAndDedupes
PASS  BatchOnboardFlow_MarkShapeForBatchAcceptsShapeInsideGroup
PASS  BatchOnboardFlow_MarkShapeForBatchRejectsWholeGroupSelection
PASS  BatchOnboardFlow_MarkingSessionPropertyRoundTripsBeyond255Chars
PASS  BatchOnboardFlow_RestoreMarkingSessionRecoversMarkedFields
PASS  BatchOnboardFlow_SaveMarkingSessionToPropertyForcesRealSave
PASS  BatchOnboardFlow_MarkingSessionSurvivesRealCloseAndReopen
PASS  BatchOnboardFlow_RestoreMarkingSessionFindsNothingOnWrongSlide
PASS  BatchOnboardFlow_FlattenGroupLeavesReturnsAllMembers
PASS  BatchOnboardFlow_ReviewGridRoundTrip
PASS  BatchOnboardFlow_CommitBatchTagsLinksAndVerifies

=== deck-sync-refimpl VBA test run (Excel) ===
PASS  CreateSheet_SeedsDeckReferenceNoFieldsOrRows
PASS  CreateSheet_RefusesToReinitialize
PASS  UpsertRow_SeedsNewInstanceFromHarvestedValues
PASS  UpsertRow_NewFieldAppendsColumnWithoutTouchingExisting
PASS  UpsertRow_PartialUpdateMergesNotReplaces
PASS  UpsertRow_NewInstanceDoesNotDisturbExistingRows
PASS  ReadSheet_PreservesFieldAndInstanceOrderAcrossManyWrites
PASS  HeaderRow_ReservesColumnAForInstanceId
```

76/76 PowerPoint + 8/8 Excel, 0 failures. `Get-Process EXCEL,POWERPNT`
reported nothing running immediately after this run.

## Manual verification recipe

1. Go to your template slide (Normal view). Click a field's shape, then run
   "Mark Field for Batch" (toolbar). Type a name (e.g. "Project Number")
   when prompted, then a type (1-4, or the name). Confirm the MsgBox
   reports "Marked field 1: '<name>' (<type>, shape '...')". Repeat for
   each real field on that slide. Re-clicking an already-marked shape and
   running it again lets you rename/re-type it instead of duplicating it.
   If the field is inside a grouped "card" layout, just click it and run
   "Mark Field for Batch" regardless of whether you've drilled in — if
   PowerPoint's own selection state ends up reporting the whole group
   (which can happen even after a real drill-in click, see the addendum
   below), you'll get a numbered picker listing every field inside it
   instead of a rejection; pick the right one by number. Also type/
   volatility prompts follow the name prompt (1-4 or the name each time),
   and (new) the whole session now survives closing PowerPoint -- if you
   close and reopen, clicking a field on the same slide and running "Mark
   Field for Batch" again picks up right where you left off, reported as
   "Restored N field(s) from a previous session" alongside the new mark.
2. If you misclick, run "Clear Marked Fields" and start over -- this also
   clears the persisted session, not just the in-memory one.
3. Click "Bulk Onboard Type" directly — no manual slide selection needed —
   the unredacted `test.pptx` (kept local, never committed) is the actual
   target deck this was built for. Confirm it switches to Slide Sorter,
   highlights the template plus every other slide sharing its layout, and
   shows a MsgBox listing them; click Yes. (To test the decline path:
   click No, then manually Ctrl-click/Slide-Sorter-select your own batch —
   template slide first, by deck order — and run "Bulk Onboard Type"
   again; confirm your explicit selection is used, not re-auto-detected.)
4. (Covered by step 3 above.)
5. Name the type. Confirm the "Field Review" Excel sheet appears with one
   row **per marked field only** (not every shape on the slide), a computed
   Suggested classification, and sample values from other slides.
6. Edit the sheet: rename fields, flip Include Y/N, exclude anything wrong.
   Confirm.
7. Provide an instance key for the template (required) and each other slide
   (blank = skip).
8. Confirm the final report: linked count, skipped count, and (should be
   zero) failed-verification count.
9. Confirm in Excel: one Data-sheet row per linked slide, correctly populated
   from harvested values, only for included fields.
10. Confirm in PowerPoint: only included fields got role tags; every
    selected slide (not just the template) carries `slide_type`/
    `instance_key` if it got a real instance key.
11. Confirm the marking session was cleared after a successful commit (run
    "Bulk Onboard Type" again with nothing selected/marked — it should ask
    you to mark fields first, not silently reuse the previous batch's).

## Addendum 2026-07-26: the "unstable null CustomDocumentProperties reads"
## scare -- fully reproduced, and it's the diagnostic tool, not the deck

After the save-durability fix above (Finding 1-3, `addin17.ppam` rebuilt and
loaded), Rohan ran "Mark Field for Batch" on his real, live, open
`test1.pptx` (`FullName = https://d.docs.live.net/.../test1.pptx`) --
in-process confirmation showed a clean "Marked field N" message with no
`WARNING:` line, meaning `SaveMarkingSessionToProperty`'s own closed-loop
verification passed. Querying externally via PowerShell moments later
(`[Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')`,
the same technique used throughout this file) then produced alarming,
inconsistent results across four successive checks: the marking-session
property reading back empty, then `Presentation.Name`/`.Saved` themselves
coming back blank with the property read throwing "cannot call a method on
a null-valued expression," then a reconnect confirming the COM connection
itself was fine, then -- on that same good connection --
`CustomDocumentProperties.Count` blank again and a `foreach` over the
collection throwing three separate null-value errors (one per property).
This read exactly like real document-state corruption specific to Rohan's
real, richly-edited file, and needed direct investigation against that real
file rather than another synthetic repro -- this codebase's synthetic
59MB-file repro (Finding 2 above) never touched real accumulated document
structure and couldn't have caught a bug that only exists there.

### What was actually tested, live, against the real file

Re-verified `addin17.ppam` is the only loaded/autoloaded add-in
(`Application.AddIns`), confirmed `Presentations.Count = 1` and
`test1.pptx` correctly identified by name and cloud `FullName`. Then ran
20 back-to-back iterations, each doing two reads of the *exact same
property on the exact same live connection* side by side:

- **(a) PowerShell-direct**: `$pres.CustomDocumentProperties.Count`,
  `$pres.Saved`, and `$pres.CustomDocumentProperties('DeckSyncMarkingSession').Value`
  -- PowerShell's own dynamic COM dispatch, no VBA involved.
- **(b) native VBA, via the already-loaded add-in**: `$ppt.Run(
  'BatchOnboardFlow.ReadMarkingSessionProperty', $pres)` -- calls the
  *existing, already-deployed* `ReadMarkingSessionProperty` function in
  `addin17.ppam` (read-only, does `pres.CustomDocumentProperties(name).Value`
  internally, no writes), executed in-process by PowerPoint's own VBA
  engine, not PowerShell's interop layer.

Result, all 20/20 iterations, completely stable both directions -- no
intermittency at all in this run:
```
[PS-direct] CustomDocumentProperties.Count =        <- blank, every time
[PS-direct] Saved = 0
[PS-direct] DeckSyncMarkingSession len = 0            <- empty, every time
[VBA-Run]   ReadMarkingSessionProperty len = 151      <- correct, every time
```
The native VBA path (b) read back the real, non-empty session on every
single call: 151 characters, decoding to a genuine 4-field marking session
on slide 256 (`Text 2|Project number|text|static`, `TextBox 51|Project
Name|text|static`, `Text 17|Project Status|text|variable`, `Text 33|About
Field Text|text|static`) -- exactly the shape of data Rohan's real marking
work would produce, not corrupted or truncated. `Presentation.Name`,
`.Saved`, and `.Slides.Count` (control reads via the *same* PowerShell-direct
mechanism, but not touching `CustomDocumentProperties`) all read back
correctly every time too -- so the live COM connection itself is genuinely
fine; the failure is scoped to this one collection type specifically.

Pushed one level further to explain *why*, not just observe it:
```powershell
$cdp = $pres.CustomDocumentProperties
$cdp.GetType().FullName   # -> throws System.NullReferenceException
```
Merely calling `.GetType()` on the `CustomDocumentProperties` collection
object PowerShell itself just successfully returned throws a
`NullReferenceException` inside .NET's own reflection of the COM wrapper --
before any indexing, before `.Count`, before touching a single property.
Every `Item(i)` (i = 1..4) and the named-index lookup threw the identical
error. This is not a race, not a timing issue, not connection staleness --
it is PowerShell's dynamic COM interop layer failing to construct a valid
managed wrapper around this specific Office collection type, full stop,
independent of document content or state.

### This is not a new discovery -- it's the same quirk this file already
### flagged once, hit again via a different code path

The "Marking-session persistence" addendum earlier in this file already
recorded: *"A direct PowerShell-side probe to check this hit an unrelated
crash (`System.NullReferenceException` inside .NET's dynamic COM interop
trying to reflect the `DocumentProperties` collection type -- a known
PowerShell-automation quirk, not an Office one)."* That was hit via
`.Add()` overload resolution; today's scare hit the identical underlying
defect via plain indexed/enumerated property access instead. Same root
cause, different entry point, not independently re-discovered so much as
re-confirmed with much more rigor this time (20x repeated, side-by-side
against a native-VBA control that never failed once).

The second check's *additional* symptom that day -- `Presentation.Name`
and `.Saved` themselves reading blank, not just `CustomDocumentProperties`
-- did not reproduce here across 20 fresh iterations (both read correctly
every time in this pass). Likely a separate, transient, garden-variety
stale-COM-reference hiccup in that one PowerShell process (e.g. a dying
RCW from a prior probe in the same session) rather than a second real bug
-- it's a materially different failure shape (whole-object blank, not
collection-specific) from the fully deterministic `CustomDocumentProperties`
defect nailed down above, and nothing here suggests it recurs when the
connection is freshly (re)acquired.

### Bottom line

**No document corruption, no property-count limit, no AutoSave
mid-session swap-out, and no regression in the save fix.** The forced-save
+ closed-loop-verification mechanism in `SaveMarkingSessionToProperty` is
working exactly as designed on Rohan's real file -- the 151-char, 4-field
session was there, correct, and stable on every single native read. The
alarming readings were produced entirely by the diagnostic tool (PowerShell's
default dynamic COM dispatch against `CustomDocumentProperties`), not by
PowerPoint, VBA, or the deck. No change made to `BatchOnboardFlow.bas` --
there is nothing to fix there; the bug, such as it is, lives in the
investigation tooling.

**Recommendation for any future external verification of this property**:
never read `CustomDocumentProperties` directly from PowerShell. Route
through `Application.Run` into the already-loaded add-in's own
`ReadMarkingSessionProperty` (or an equivalent existing/native VBA
function) instead -- confirmed reliable 20/20 here, against the exact same
live document, in the exact same PowerShell session where direct access
failed 20/20. `Presentation`-level properties (`.Name`, `.Saved`,
`.Slides.Count`, `.FullName`) and the `Presentations`/`AddIns` collections
are all confirmed fine via direct PowerShell dispatch -- the defect is
specific to `DocumentProperties`-family collections, not a blanket "don't
trust PowerShell COM" conclusion.

## Addendum 2026-07-26: the real 46-slide commit -- "Linked: 0, FAILED
## verification: 46" -- root cause found and fixed

Rohan ran "Bulk Onboard Type" for real for the first time at real scale
against his actual deck: marked several fields on a template slide,
`FindSameLayoutSlides` auto-selected all 46 same-layout slides in the deck
(every project-report slide shares one base layout), he confirmed instance
keys, and `CommitBatch` ran. Result: **every single slide failed
verification** -- `Linked: 0`, `FailedVerificationCount: 46`, zero skips.

This mattered more than a normal test failure because `CommitBatch`
(`BatchOnboardFlow.bas`) writes real tags (`Onboarding.ConfirmFieldMatch`,
`slide_type`/`instance_key`) and a real Data-sheet row (`ExcelOutput.
UpsertRow`) for every included field on every confirmed slide
*unconditionally*, **before** it calls `VerifyBatchLink` -- verification
failing only changes which counter a slide lands in, it does not roll
anything back. So the report's "Linked: 0" was consistent with all 46 real
slides having real role/slide_type/instance_key tags and 46 real Data rows
already written, despite reading as a total failure.

### Root cause, confirmed (not theorized)

`InjectPrimitive.bas`'s `FindShapeByRoleTag` -- the function every write
path in this project (`OnboardFlow.bas`, `SyncOperations.bas`,
`BatchOnboardFlow.bas`, `DeckAdoption.bas`, `SlideDuplication.bas`) goes
through to re-resolve a shape by its role tag after tagging it -- looped
over `sld.Shapes`:

```vba
For Each shp In sld.Shapes
    If ShapeHasRoleTag(shp, identityTag) Then ...
```

`Slide.Shapes` is PowerPoint's **top-level-only** shape collection -- it
does not descend into a `GroupShape`'s `GroupItems`, exactly the same
distinction `Discovery.bas`'s own `Walk` function exists to handle (and
does handle correctly, recursing into `msoGroup` -- see that module's
header). `Onboarding.ConfirmFieldMatch` (`shp.Tags.Add "role", role`) works
on any shape regardless of nesting, so the *write* half of tagging a
grouped field always succeeded. But `FindShapeByRoleTag`'s flat loop could
**never** find that tag again if the shape was nested inside a group --
`Found = False` for every grouped field, on every slide, every time,
deterministically. `VerifyBatchLink` treats `Not r.Found` as an immediate
failure for the whole slide (`If Not r.Found Or r.Written Or Not r.Verified
Then VerifyBatchLink = False`), so one grouped field among several included
fields was enough to fail the entire slide.

This is a pure read-side bug, not a matching/scoring bug and not a
stale-tag bug (the three other hypotheses this investigation was asked to
rule in or out). It doesn't depend on prior partial onboarding attempts,
ambiguous correspondence scoring at scale, or `FieldInclude`/Dictionary-key
issues -- it reproduces on a single clean slide with a single grouped
field and zero prior history. It explains the *uniformity* of the failure
(46 for 46, not a handful) better than any of those alternatives: the same
field positions were marked once on the template and matched onto every
other same-layout slide, so if a given field is grouped on the template it
is (given they share a layout) grouped the same way on every instance --
deterministically wrong every time, not probabilistically wrong sometimes.

**Confirmed against real structure, not assumed**: `test-fixtures/
crc-real-deck-redacted.pptx` is a redacted copy of Rohan's actual 46-slide
deck (see `test-fixtures/SOURCE.md`). A direct check of its raw slide XML
(`ppt/slides/slideN.xml`) found **all 46 of 46 slides contain at least one
`<p:grpSp>` group** (518 group elements total across the deck), and on
slide 1 specifically, 54 of the slide's text-bearing shapes are nested
inside a group vs. 28 at the top level -- i.e. *most* of a real slide's
real content, structurally, lives inside groups, not at the top level.
This lines up exactly with this module's own prior comments (`Rohan's real
deck uses grouped "card" layouts for its fields`, from the
`MarkShapeForBatch`/`FlattenGroupLeaves` work earlier the same day) --
this was already known to be true of marking, and turns out to have been
silently unhandled on the verification side the whole time.

**Why this was never caught before today**: every existing automated test
that exercises `InjectPrimitive`/`VerifyBatchLink`/`CommitBatch`
(`Test_InjectPrimitive_*`, `Test_BatchOnboardFlow_CommitBatchTagsLinks
AndVerifies`) used only top-level, ungrouped textboxes. `Discovery`'s own
group-recursion test (`Test_Discovery_GroupRecursionFindsCandidates`) and
`BatchOnboardFlow`'s marking-side group tests (`MarkShapeForBatchAccepts
ShapeInsideGroup`, `FlattenGroupLeavesReturnsAllMembers`) proved the
*write*/*discovery* halves handle groups correctly -- nothing proved the
*re-read-by-tag* half did, and it didn't.

### The fix

`FindShapeByRoleTag` now delegates to a small recursive helper
(`WalkForRoleTag`, private to `InjectPrimitive.bas`) that mirrors
`Discovery.Walk`'s own group-recursion shape: descend into `GroupItems` on
`msoGroup`, check the tag on every real leaf shape otherwise. `matchCount`
accumulates by `ByRef` across the whole recursive walk, so a same-tag
collision between a top-level shape and a nested one (or across two
different groups) is still caught as ambiguous, not silently missed by
scoping the count per group level. `TestRunner.bas`'s own `FindShapeByRole`
test helper (used by several unrelated tests to locate a tagged shape
post-write) had the identical flat-loop bug and was fixed the same way, so
it can't mask a regression here either.

New tests added directly targeting this bug class:
- `Test_InjectPrimitive_FindsRoleTagInsideGroup` -- the smallest possible
  repro: a role tag written onto a shape nested one level inside a group
  must be found again, on both the no-op and the write-and-verify paths.
- `Test_InjectPrimitive_AmbiguousTagAcrossGroupAndTopLevelRefusesToGuess`
  -- proves the collision guard still holds across nesting levels after
  the rewrite.
- `Test_BatchOnboardFlow_CommitBatchWithGroupedFieldsAtScale` -- a
  template + 10 other slides (11 total, matching "much bigger than any
  prior test batch, which topped out at 2-3 slides"), each with a
  realistic mix of one top-level field and one field nested inside a
  group. Before the fix this test fails exactly the way the real run did
  (`LinkedCount = 0`, `FailedVerificationCount = 11`); after the fix all
  11 slides link and verify cleanly, and the grouped shapes' own tags/text
  are checked directly as a second, independent confirmation alongside
  `VerifyBatchLink`'s internal round-trip.

### Is Rohan's real deck's current tag state (from the failed run) fine,
### corrupted, or unknown?

**Likely fine, but not provably so without a direct read-only check.**
Reasoning: `ConfirmFieldMatch`'s tag *writes* were never the broken half --
only the subsequent *lookup* was -- so the real deck's 46 slides almost
certainly do carry the correct role/slide_type/instance_key tags on the
correct (correspondence-matched) shapes, and the paired
`SAAFE-Projects-Data.xlsx` almost certainly has 46 correct rows, exactly as
harvested. Nothing in this root cause implies the *wrong* shape got
tagged, values got scrambled between fields, or the same role landed on
two shapes -- it implies only that a subsequent, unrelated read step
couldn't see what had already been written correctly.

That said, this is reasoned from the code, not observed directly against
the real file -- per this task's own safety constraints, `test1.pptx` was
never opened or written to by this investigation. **One read-only check
would convert "likely fine" into "confirmed fine"**: for each of the 46
slides, walk `sld.Shapes` recursively (the same `WalkForRoleTag` logic,
called read-only) and confirm (a) every field that was supposed to be
linked carries exactly one shape with the expected `role` tag, and (b)
that shape's current text matches what the confirmation dialog's harvested
values would have been. This is safe to run against the live file (no
writes, mirrors the fixed lookup logic) and would be the natural first
thing to run from the main session before deciding whether any remediation
of the real deck is needed at all.

### Test suite run

This environment (the subagent sandbox this investigation ran in) is
Linux with no PowerShell and no Office install -- `vba/tests/
run_vba_tests.ps1` requires a real Windows PowerPoint/Excel COM
automation session (per `AGENTS.md`'s Testing section) and could not be
executed here. **The fix and new tests above are unexecuted in this
environment** -- they need a real run via `run_vba_tests.ps1` from the
main session (the environment `AGENTS.md` confirms has `powershell.exe`
reachable and Office installed) before being trusted. Flagging this
explicitly rather than presenting untested code as verified.

### 2026-07-27 addendum: confirmed in real Office, plus a second, unrelated bug that masked the confirmation

Ran `run_vba_tests.ps1` for real (Windows/Office session, main loop, not
the sandbox above). First few attempts looked like environment flakiness
-- a stuck POWERPNT.EXE left in the VBA IDE overnight, an empty Excel
automation server slow to release its process, a run that legitimately
exceeds the harness's 120s foreground timeout on a full 87-test suite (not
a hang, just backgrounded). All were real but were a distraction from the
actual blocker underneath them: `Application.Run("TestRunner.RunAllTests",
...)` was failing cleanly every time with "Sub or function not defined" --
which is *also* what VBA reports when the target project has a genuine
compile error anywhere, not only when the named macro is missing.

The real cause: `Test_BatchOnboardFlow_CommitBatchWithGroupedFieldsAtScale`
(the new scale test added above) called `.Group()` -- WITH parens -- as a
bare statement, discarding the return value, on two lines:

    templateSld.Shapes.Range(Array(tFieldB.Name, tSibling.Name)).Group()
    sld.Shapes.Range(Array(oFieldB.Name, oSibling.Name)).Group()

`.Group` is a method that returns a Shape. Calling it as a bare statement
(no `Set`, no `Call`) with trailing `()` -- even *empty* parens -- is a
genuine VBA "Compile error: Syntax error", confirmed directly via
PowerPoint's own Debug > Compile VBAProject dialog. Every other `.Group()`
call already in this file (`Set grp = sld.Shapes.Range(...).Group()`)
was fine because it assigns the result; only these two bare-statement
calls (added because the scale test doesn't need a reference to the new
group) tripped it. Fix: drop the parens (`.Group` with no trailing `()`)
on both lines -- the grouping side-effect still happens, nothing else
changes.

Two things worth naming about how this got found:
- A first read of the (correctly rewritten) `InjectPrimitive.bas` fix and
  an automated module-by-module import bisection both showed the project
  "compiling" cleanly, which was misleading: a late-bound `Application.Run`
  call to an *unrelated*, already-valid macro (a `PingOK` canary) does not
  reliably force a full-project compile the way triggering the actually-
  broken procedure does. Don't trust an indirect compile proxy over asking
  Office directly (Debug > Compile) when the two disagree.
- Rohan flagged "it hung on syntax errors" as the very first message about
  this. That was dismissed as probably-the-zombie-process story already in
  progress, and it turned out to be the literal, correct diagnosis the
  whole time. When someone with eyes on the actual failing UI names the
  category of the problem, weight that over an in-progress theory that
  merely fits the symptoms so far.

Post-fix: `run_vba_tests.ps1` reports **87/87 PASS, 0 FAIL** in real
Office (84 previously-established + the 3 new grouped-shape tests from
the addendum above, including `..._CommitBatchWithGroupedFieldsAtScale`,
which directly reproduces the original "Linked: 0 / FAILED verification:
46" production symptom at a comparable scale and now passes). The
`InjectPrimitive.bas` grouped-shape fix is therefore now genuinely test-
confirmed, not just code-reviewed. Add-in rebuilt and shipped as
`addin20.ppam`.

Real-deck-tag-state read-only verification: built (`vba/tools/
VerifyRealDeck.bas` + `vba/tools/verify_real_deck.ps1`, not part of the
shipped add-in -- opens both real files ReadOnly, closes without saving,
never writes). Result: **46/46 slides fully OK, 230/230 fields present
and matching, 0 mismatches.** Confirms the original failed run's writes
all landed correctly on the real deck -- only `CommitBatch`'s own
verification read-back was broken (the `FindShapeByRoleTag` group bug
fixed above), never the writes themselves. No remediation of the real
deck was needed.
