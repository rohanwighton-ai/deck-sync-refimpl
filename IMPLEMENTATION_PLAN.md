# Implementation Plan

## Priority 26 (2026-07-26 pass): BatchOnboardFlow.bas — Excel-grid-driven bulk onboarding, closing the real-deck usability gap

Built specifically for the real, richly-designed deck (`test.pptx`, kept
local/unredacted) that discovers 60-90 raw candidate shapes per slide —
`OnboardFlow.bas`'s InputBox-chain review is unusable at that density. New
module `vba/BatchOnboardFlow.bas`: select 2+ slides of one type, it finds
per-field correspondence across the whole batch via `Matching.Match` (reused
unchanged), computes a "decoration vs. real field" suggestion by diffing
harvested text across every selected slide, and presents it as an **editable
Excel grid** instead of a sequential prompt chain — closes
`specs/onboarding.md`'s long-documented-but-unbuilt "boilerplate-vs-varying
pre-filter" gap, in a different shape than that spec sketched (a human-edited
table beats an algorithm guessing alone, and Excel is a far more capable
editing surface than anything hand-built here would be). Onboarding a new
type and bulk-linking its existing instances happen in one pass. Sixth
`CommandBarUI.bas` toolbar button: "Bulk Onboard Type."

**Cost an unusually long debugging session** — full account in
`SPIKE_NOTES_BatchOnboardFlow.md`, worth reading before touching VBA
declaration order or Dictionary-of-array patterns in this project again.
Three separate real bugs hid behind one consistent, misleading symptom
("Compile error: User-defined type not defined" with VBE's navigation
pointing at the wrong line almost every time): (1) a `Candidate()` UDT array
stored in a `Scripting.Dictionary` — illegal, already logged in `AGENTS.md`
but not checked before writing this module; (2) a newly-confirmed VBA
compiler quirk — a `Public Type`/`Private Const` declared *after* a
Function/Sub in the same module fails to resolve cross-module, fixed by
moving every declaration to the top of the file; (3) an ordinary reserved-word
typo (`single` as a variable name) and one test-data bug. **70/70 real-Office
tests pass.**

Closes the one deferral Priority 22 (`deck-adoption.md`) explicitly left open:
"`ribbon-ui.md` gets an 'Adopt Existing Slides' entry point in a *future*
pass." Built `vba/AdoptFlow.bas` (mirrors `OnboardFlow.bas`'s relationship to
`Onboarding.bas` — pure interactive glue, `DeckAdoption.bas`'s `PlanAdoption`/
`CommitAdoption` still does every real decision) and wired a fifth
`CommandBarUI.bas` toolbar button to it.

Two real bugs found and fixed this pass, full account in
`SPIKE_NOTES_AdoptFlow.md`:
1. A whole-array `ByRef` assignment (`outSlides = unsorted`) left the
   caller's array genuinely unallocated — first time this project tried that
   shape for a `ByRef` array out-param; fixed to match `PlanAdoption`'s own
   `ReDim`-then-populate-in-place convention.
2. `Slides.Range(...).Select` under COM automation only registers as a real
   slide-type selection in Slide Sorter view — in Normal view (where
   `NewBlankSlide()` leaves the test window), `Selection.Type` silently
   stayed `ppSelectionNone`. Confirmed automation-only, not a real-user
   limitation; fixed in the test, not the production code.

65/65 real-Office tests pass (up from 61 at the end of Priority 24).
Requires a full rebuild of `dist/deck-sync-refimpl.ppam` (import all 18
modules, Save As again) to actually reach a real installed add-in — same
manual release-step constraint Priority 24 already established, not a new
one.

## Priority 23 (2026-07-26 pass, hand-driven, WSL host with real Office): real-Office verification + starting Priority 21

Run directly on the WSL host (not the Docker container) with `powershell.exe`
and real PowerPoint/Excel reachable, so this pass gets the real-Office
verification every prior VBA pass since Priority 20a had to defer.

First real run of the full suite found two real bugs, both in test/harness
code, not in `DeckAdoption.bas`/`RunSync.RunPeriodRollover` themselves:
1. `vba/tests/run_vba_tests.ps1`'s module import lists were never updated
   when `DeckAdoption.bas`/`ResolveFields.bas` landed (2026-07-25) — caused a
   real `TestRunner.bas` compile error ("user-defined type not defined" on
   `AdoptionSlidePlan`) on the first attempt. Fixed.
2. `ResolveFields`'s 3 shape-selection tests failed with `Shape.Select ::
   Invalid request. To select a shape, its view must be active` — the shared
   `NewBlankSlide()` test helper never navigated the window's view to a newly
   added slide, so `.Select` on its shapes failed. Fixed by adding
   `Application.ActiveWindow.View.GotoSlide` to `NewBlankSlide()`. Full
   account in `SPIKE_NOTES_ResolveFields.md`/`SPIKE_NOTES_DeckAdoption.md`.

After both fixes: **41/41 real-Office tests pass**, PowerPoint and Excel
sides both. `RunSync.RunPeriodRollover` (Priority 20a) and `DeckAdoption.bas`
(Priority 22) are now genuinely proven against real Office, not just
line-traced — both flagged "not executed this pass" caveats below are
resolved as of this run.

Moving on to Priority 21 (`ribbon-ui.md`) in this same pass, starting with
the prerequisite gap Priority 21's own text didn't fully resolve: no module
anywhere stores which template slide / worksheet / workbook backs a given
slide type on a given deck (`ExcelOutput.bas`'s `WriteDeckReference`/
`ReadDeckReference` only cover the workbook→deck direction, and are never
actually called by anything — confirmed via `grep -rn WriteDeckReference
vba/*.bas` matching only `ExcelOutput.bas` itself). Every existing entry
point takes `ws`/`templateSld` as caller-supplied parameters specifically
because "where a template lives" was deferred everywhere (see the Priority 19
carry-forward note below) — a one-click ribbon button has no caller to
supply them. See Priority 24 below for the registry this pass builds to
close that gap before wiring any button to it.

## Priority 24 (2026-07-26 pass): closed Priority 21 entirely — deck registry, all four ribbon actions, and a real shipped UI

Everything in Priority 21's checklist is now `[x]`, but the *shape* of what
shipped differs from the spec in one real way: no ribbon tab exists, because
none can for a `.ppam` (see Priority 21's own updated bullet and
`SPIKE_NOTES_RibbonUI.md` for the full, costly proof). What actually shipped:

- `specs/deck-registry.md` + `vba/DeckRegistry.bas` — the prerequisite this
  priority's own header flagged: deck GUID, paired workbook path, and
  per-type (template SlideID, worksheet name) lookups, stored in
  `Presentation.CustomDocumentProperties` mirroring `ExcelOutput.bas`'s
  existing (never-called-until-now) `WriteDeckReference` pattern. 8 tests,
  real-Office proven.
- `vba/WorkbookBridge.bas` — small shared plumbing (get-or-open workbook,
  get-or-add worksheet, sheet-name sanitizing) both `RibbonUI.bas` and
  `OnboardFlow.bas` needed identically.
- `vba/OnboardFlow.bas` — the Onboard New Slide Type flow, phase-gated
  InputBox review, real design call on seeding (example slide becomes
  instance #1 too, avoiding a second "keyless row" mechanism). 6 tests.
- `vba/RibbonUI.bas` — all four action Subs (`SyncNow`/`NewPeriod`/
  `OnboardNewType`/`ResolveUnmatchedFields`), refactored to be
  callback-shape-agnostic (plain parameterless Subs) once the ribbon path
  died, so `CommandBarUI.bas` could wire them directly. 5 tests plus the
  full live add-in proof below.
- `vba/CommandBarUI.bas` — the actual shipped UI: a `CommandBars` toolbar
  (`Auto_Open`-driven), because a real ribbon is provably impossible for a
  `.ppam`. 3 tests, plus the real end-to-end proof: 17 modules imported,
  Rohan hand-saved a real `.ppam` (`Addin2.ppam`), loaded via `AddIns.Add` +
  `.Loaded = True` — `Auto_Open` fired, the "Deck Sync" toolbar appeared with
  all 4 buttons correctly wired, confirmed both via COM and visually by
  Rohan. Copied into the repo at `dist/deck-sync-refimpl.ppam`.
- Also fixed along the way: `run_vba_tests.ps1`'s stale import lists (missing
  `DeckAdoption.bas`/`ResolveFields.bas`, caused a real compile error on first
  attempt) and `TestRunner.bas`'s `NewBlankSlide()` (missing a
  `View.GotoSlide` call, caused 3 real `Shape.Select` failures) — see
  Priority 23 above.

**Real, load-bearing finding for whoever packages a build next**: there is no
way to automate producing or updating a `.ppam` end-to-end. `SaveAs(30)` fails
via automation for reasons that aren't discoverable from outside Office, and a
loaded add-in's `VBComponents` is genuinely `null` to automation (not just
empty) — both confirmed, not assumed. `vba/tests/build_ppam.ps1` now documents
and does the one half that *is* fully automatable (importing every production
module into a fresh presentation) and stops there; the File > Save As click is
real, permanent, human-required. This is a release-time cost, not a dev-loop
one — `run_vba_tests.ps1` stays fully automated regardless.

Full test count this pass: 61/61 real-Office tests pass (PowerPoint + Excel
combined), up from 41 at the start of Priority 23. `python3 -m pytest`/`mypy`
unaffected throughout (VBA-only work).

## Notes carried forward, still open after this pass
- The 1-vs-2-example-slide onboarding gap (Priority 21's `OnboardFlow.bas`
  bullet) is real and deferred, not decided against.
- The shared result reporting is `MsgBox`-backed, not a structured form —
  deferred, `ShowSyncResult` is the single point a future upgrade would touch.
- A real COM/VSTO add-in (genuinely different technology from `.ppam`) is the
  only path to an actual ribbon tab, if that polish is ever wanted — not
  started, `customUI14.xml` is kept as its starting point.
- Cases 5/7 and the multi-deck design remain entirely unspecified, unchanged
  from every prior pass's carry-forward note.

## Priority 20 (2026-07-25 pass): two new specs landed with zero implementation — this is the real gap

Re-verified this pass by actually running `python3 -m pytest tests/ -v` (**70 passed**)
and `python3 -m mypy src/` (**no issues, 10 source files**) — both green, no
regressions since Priority 19. Note: this pass runs in a plain Linux container
(`Dockerfile`/`loop-docker.sh`), not the WSL host with a real Windows/Office install
Priorities 17-19 used — `powershell.exe` is not reachable here (confirmed:
`which powershell.exe` fails), so nothing VBA-side can be executed or re-verified
against real Office this pass. All VBA gap analysis below is via direct code reading
only, same constraint the original 6-module port order operated under before
Priority 17.

Confirmed via `git log` that two commits landed since Priority 19 (`ab86ea0`) closed:
`be6ddee` (`specs/ribbon-ui.md`) and `8b1bb32` (`specs/deck-adoption.md`). Both read in
full this pass. Confirmed via `find`/`grep` across `vba/` and `src/` that **neither has
any implementation at all** — no `RibbonUI.bas`, no `customUI14.xml` anywhere in the
repo, no `DeckAdoption.bas`, no `Ribbon`/`Adopt` symbol of any kind in `vba/` or `src/`.
This is the first planning pass where the gap isn't "one module in an established
port order" but two full new specs with nothing built yet — treat both as the
headline priority below, ahead of the stale lower-numbered priorities kept for
history.

While reading both specs against the existing engine, found one concrete,
already-confirmed **prerequisite gap that blocks part of `ribbon-ui.md`**, not just a
UI-wiring task: `ribbon-ui.md`'s "New Period" button says to call into "an existing
Sub" for case 2, but no such Sub exists. `SyncOperations.PlanPeriodRollover`
(`vba/SyncOperations.bas:183`) only *decides* a rollover (returns a `PeriodRollover`
struct); nothing calls `SlideDuplication.DuplicateAndTag` with that decision the way
`RunSync.RunRoutineSync` already does for case 3. `RunSync.bas`'s own header comment
and `SPIKE_NOTES_RunSync.md:79` already flag this as intentionally out of scope for
Priority 19 ("case 2... not driven from here") — confirmed still true by reading the
full file, not just grepping for a TODO (there are none; this project doesn't leave
TODO markers, per `grep -rn TODO` returning nothing repo-wide). See Priority 20a below.

## Priority 20a: `RunSync.RunPeriodRollover` — the missing case-2 execution primitive

- [x] Add an execution path for `SyncOperations.PlanPeriodRollover`'s decision,
      mirroring `RunRoutineSync`'s case-3 handling (why: this is the one real engine
      gap blocking `ribbon-ui.md`'s "New Period" button — everything else that spec
      needs, per the Reference-code check below, already exists and is already
      tested). Given an existing `SlideInstance` (the record rolling to a new
      period) and `newValues`, call `SyncOperations.PlanPeriodRollover` to get the
      `PeriodRollover` decision, then call `SlideDuplication.DuplicateAndTag` with
      the *current* slide as source and a caller-supplied new `instance_key` for the
      new period (mirrors `slide-duplication-trigger.md`'s "one primitive, two
      callers" framing for cases 2/3 — `DuplicateAndTag`'s signature is already
      generic enough for this, confirmed by reading `vba/SlideDuplication.bas:34`
      directly; it takes `sourceSld`/`slideType`/`newInstanceKey`/`values`/
      `existingInstances` with no case-3-specific assumption baked in). Should live
      in `RunSync.bas` next to `RunRoutineSync`/`ResequenceByRowOrder` (same module
      that already owns "decision struct in, real deck mutation out" for case 3) —
      not a new module, to avoid duplicating `GatherInstances`/collision-guard
      wiring. Needs its own `SPIKE_NOTES_RunSync.md` update (new section, not a new
      file — the existing one already governs `RunSync.bas`) documenting the design
      choice of where the source slide comes from (the instance's own current slide,
      confirmed by `Resolve.ResolveSlideInstance`) and a manual verification recipe,
      per `vba-port.md`'s per-module-recipe requirement that still applies to every
      VBA addition in this repo. This function is a prerequisite for Priority 21's
      "New Period" button — do this first so that task is pure UI wiring, not engine
      design under a UI task's scope.
      Added `RunSync.RunPeriodRollover(sourceSld, slideType, newInstanceKey,
      newValues)` in `vba/RunSync.bas`, next to `RunRoutineSync`/`ResequenceByRowOrder`
      as planned. Resolves `sourceSld` into a `SlideInstance`, calls `SyncOperations.
      PlanPeriodRollover` to get the decision, gathers `slideType`'s current instances
      itself (same non-goal-gathering posture every module in this port takes), then
      calls `SlideDuplication.DuplicateAndTag(sourceSld, slideType, newInstanceKey,
      rollover.NewValues, existingInstances)` — `sourceSld` is never mutated by
      `DuplicateAndTag`, so "leave the original untouched as history" falls out of
      reusing the existing primitive rather than needing new logic. Takes the slide
      directly (not an instance_key to look up) — resolving "which slide is this
      instance_key on" is left to the caller (e.g. Priority 21's picker), not
      reinvented here. Returns `DuplicateResult` directly (already-structured, same
      type `RunRoutineSync`'s case-3 branch consumes) rather than a new `String`
      report format, since Priority 21's shared-result-form task hasn't yet decided
      whether `RunRoutineSync` itself should move off `String` reports. Added
      `Test_RunSync_RunPeriodRolloverDuplicatesLeavingSourceUntouched` to
      `vba/tests/TestRunner.bas`: confirms the new slide gets the injected value and
      the new instance_key tag, confirms the *source* slide's value is unchanged
      after the call (the actual case-2-specific claim), and confirms a second
      rollover onto an already-used instance_key is refused via the same collision
      guard `DuplicateAndTag` already enforces for case 3. Documented in a new
      `SPIKE_NOTES_RunSync.md` section. **Not executed against real Office this
      pass** — `powershell.exe` unreachable from this plain-Linux container
      (confirmed via `which powershell.exe`), same constraint the rest of this pass
      operates under; needs a real `run_vba_tests.ps1` run on the WSL/Windows host
      next time this project is picked up there. `python3 -m pytest tests/ -v` (70
      passed) and `python3 -m mypy src/` (no issues, 10 source files) both confirmed
      unaffected (VBA-only change).

## Priority 21: `specs/ribbon-ui.md` — the ribbon/forms layer

Confirmed by reading `ribbon-ui.md` requirement-by-requirement against the current
engine (`vba/RunSync.bas`, `Onboarding.bas`, `SlideDuplication.bas`, `ExcelOutput.bas`,
`Resolve.bas`) that every button besides "New Period" (Priority 20a) already has a
tested Sub/Function to call — this task is genuinely "thin ribbon calling existing
code," as the spec's own opening paragraph claims, once 20a lands. Split into
sub-tasks below since the spec itself describes several independent pieces (a ribbon
XML, several forms, one shared result form) — each is completable in one iteration;
don't attempt all of `ribbon-ui.md` in a single pass.

- [x] `customUI14.xml` + ribbon callback module — **superseded, provably
      impossible for a `.ppam`.** Built `customUI14.xml` and `RibbonUI.bas`'s
      original `(control As IRibbonControl)` callbacks exactly as planned, then
      spent most of a 2026-07-26 pass discovering a `.ppam`'s add-in loader
      rejects the package outright if it contains anything beyond its exact
      expected 5-part structure — proven definitively by adding a harmless,
      totally unrelated, unreferenced dummy part and watching it fail to load
      identically to the real ribbon injection attempt. Full blow-by-blow
      (SaveAs(30)'s automation-only COMException, the reverse-engineered
      5-part real structure, the failed vbaProject.bin splice, the
      null-VBComponents-on-a-loaded-addin dead end) is in
      `SPIKE_NOTES_RibbonUI.md` — read that before attempting this again.
      **Shipped instead**: `CommandBarUI.bas`, the pre-Ribbon `CommandBars`
      mechanism — pure runtime VBA, zero package changes needed, so it works
      inside the exact `.ppam` structure already proven to load. `Auto_Open`
      confirmed firing automatically on add-in load; toolbar with all 4
      buttons visually confirmed live by Rohan. `RibbonUI.bas`'s action logic
      was refactored into plain parameterless Subs (`SyncNow`/`NewPeriod`/
      `OnboardNewType`/`ResolveUnmatchedFields`) so `CommandBarButton.OnAction`
      can call them directly — `customUI14.xml` is kept in the repo,
      unwired, as a starting point if a real COM/VSTO add-in (a genuinely
      different technology) ever gets built. The one-click "Sync Now" open
      sub-decision this bullet originally flagged is resolved by
      `DeckRegistry.bas` (Priority 24 below): `GetWorkbookPath`/
      `ListRegisteredTypes`/`LookupType` supply exactly what was missing.
- [x] "New Period" picker (type, then record) → `RunSync.RunPeriodRollover`.
      Built as `RibbonUI.NewPeriod`: type picker sourced from
      `DeckRegistry.ListRegisteredTypes`, record picker from
      `RunSync.GatherInstances`, new instance key entered directly (must
      already have a row in the Data sheet — `run-sync.md` Step 3's "create
      the row first" is a spreadsheet-first action the user takes before
      clicking, not something this button invents values for). Real-Office
      tested (`RibbonUI_ResolveTypeAnswerAcceptsNumberOrName`,
      `RibbonUI_ResolveRecordAnswerAcceptsNumberOnly`, and the picker-prompt
      builders); the interactive InputBox chain itself is manual-only, same
      constraint as every other InputBox flow.
- [x] "Onboard New Slide Type" flow. Built as `vba/OnboardFlow.bas`: duplicates
      the selection into a working copy immediately (Step 1's invariant),
      `Discovery` + a `PlanOnboarding` pass builds the reviewable field list,
      an InputBox-chain Review phase-gates every write (rename/exclude/mark
      period-key, final Yes/No confirm before anything is written — matches
      the spec's phase-gate requirement, just not via a `UserForm` grid), then
      `CommitOnboarding` tags fields/establishes the template+first-instance/
      registers in `DeckRegistry`/seeds the Data-sheet row, and
      `VerifyOnboarding` runs the Step-6 no-op-path check per field. Real bug
      caught and fixed by this same design work: the seed row needs a real
      `instance_key` or it becomes an invisible keyless row (`ExcelOutput.
      ReadSheet` excludes blank-Instance-ID rows) — resolved by having the
      example slide double as instance #1, not inventing new "keyless seed
      row" handling. Real-Office tested end-to-end, including a full commit
      against a real throwaway Excel workbook created via COM
      (`OnboardFlow_CommitAndVerifyOnboardingRoundTrip`). Full design
      rationale in `SPIKE_NOTES_OnboardFlow.md`. **Scope reduction from the
      spec**: supports exactly 1 example slide, not "1-2" — flagged as
      deferred, not silently dropped.
- [x] "Resolve Unmatched Fields" flow: user clicks a shape
      (`Application.ActiveWindow.Selection.ShapeRange`), picks a role from the
      template's defined roles, calls `Onboarding.ConfirmFieldMatch` (why: spec
      names this exact mechanism and primitive — the only new code is the shape-
      click capture and role picker, both pure UI; `ConfirmFieldMatch` already
      exists and is tested per `SPIKE_NOTES_Onboarding.md`).
      Built `vba/ResolveFields.bas`: `PromptResolveUnmatchedField(templateSld)`
      is the ribbon-button entry point (reads
      `Application.ActiveWindow.Selection`, drives an `InputBox`, calls
      `Onboarding.ConfirmFieldMatch` unchanged — no new matching logic).
      Split into that thin interactive entry point plus three `Public`
      pure-logic helpers (`ValidateSingleShapeSelection`,
      `BuildRolePickerPrompt`, `PickRoleFromList`), mirroring
      `DeckAdoption.bas`'s posture of keeping decision logic testable by
      taking objects as parameters. **Deliberately used `InputBox` instead of
      a `UserForm` ListBox**: this project has never authored a VBA UserForm
      before, and whether a `.frm` can carry its full control layout as plain
      text (vs. requiring an opaque binary `.frx`) can't be confirmed without
      a real VBE to export/import against — this container has none
      (`powershell.exe` unreachable, same constraint as every VBA task this
      pass). Rather than hand-author an unverifiable binary-adjacent file,
      picked the lower-risk, already-proven mechanism; flagged as an open
      question for `ribbon-ui.md`'s shared-result-form bullet too, which will
      hit this identical question for its own dialog. Added 5 tests to
      `vba/tests/TestRunner.bas`: two `ValidateSingleShapeSelection` cases
      that turned out to be genuinely testable against a real selection (not
      a mock) since `run_vba_tests.ps1` runs PowerPoint visible, so
      `shp.Select`/`Shapes.Range(Array(...)).Select` really do change
      `Application.ActiveWindow.Selection`; `BuildRolePickerPrompt` and
      `PickRoleFromList` unit tests; and one end-to-end test wiring
      selection→role-lookup→`ConfirmFieldMatch` together (skipping only the
      `InputBox` call itself) to prove the pieces actually compose. The
      `InputBox` interaction itself has no automated coverage (no headless
      harness can click through a live modal) — covered by a manual
      verification recipe in the new `SPIKE_NOTES_ResolveFields.md` instead.
      `python3 -m pytest tests/ -v` (70 passed) and `python3 -m mypy src/`
      (no issues, 10 source files) both confirmed unaffected (VBA-only
      change). **Not executed against real Office this pass** —
      `powershell.exe` unreachable from this container; needs a real
      `run_vba_tests.ps1` run on the WSL/Windows host next time this project
      is picked up there.
- [x] Shared result reporting: `RibbonUI.ShowSyncResult(title, report)`, called
      from every action (Sync Now, New Period, Onboard New Slide Type) rather
      than divergent per-button `MsgBox` calls — meets the spec's "not a
      bespoke dialog per action" intent, just backed by `MsgBox` instead of a
      `UserForm` (same no-proven-`.frm`-precedent reasoning as everywhere
      else this pass; upgrading the display mechanism later only touches this
      one function, by design). Structured counts/flagged-item display (vs.
      the current formatted `String` report) is a real, deliberate deferral,
      not an oversight — `RunRoutineSync`'s return type would need to change
      to a structured result for a form to consume directly, and that's a
      bigger, separate refactor.
- [x] Manual verification recipes added: `SPIKE_NOTES_OnboardFlow.md`,
      `SPIKE_NOTES_RibbonUI.md`, `SPIKE_NOTES_CommandBarUI.md` (the last of
      these is the one that actually matters for click-through verification,
      since the ribbon itself was superseded).

## Priority 22: `specs/deck-adoption.md` — bulk retroactive linking

Confirmed by reading `deck-adoption.md` against `Onboarding.bas`/`Matching.bas`/
`Verification.bas`/`InjectPrimitive.bas` that the *scoring* and *tag-write*
primitives it needs already exist and are reusable as-is (`Onboarding.
BuildTemplateFieldShapes`, `MatchSlideAgainstTemplate`, `ConfirmFieldMatch`) — but the
**batch orchestration itself is new logic**, not just wiring, unlike most of
`ribbon-ui.md` above. In particular: idempotent-skip-if-already-tagged, harvesting
verbatim values per matched slide, resolving against pre-existing keyless Data-sheet
rows (exact non-key-field match required, ambiguous → fresh row, never guessed), the
whole-batch phase gate before any write, deck-order-bootstraps-row-order (the one-time
exception to `slide-duplication-trigger.md`'s normal invariant), and the
verify-the-link pass generalized across a batch — none of these exist anywhere in
`vba/` today (confirmed via `grep -rli "adopt\|batch" vba/` returning nothing beyond
this task's own future filename).

- [x] Build `vba/DeckAdoption.bas`: entry point takes
      `Application.ActiveWindow.Selection.SlideRange` as scope (why: spec's explicit
      "explicit selection, not whole-deck scanning" requirement, deliberately not a
      reversal of the 2026-07-19 default-posture decision the spec itself cites).
      Greenfield path (no template yet) hands the user's chosen template slide to
      the *existing*, unchanged `onboard-slide-type.md` flow (i.e. Priority 21's
      Onboard New Slide Type flow, once built — or the underlying primitives
      directly if this lands first) — this spec adds nothing to template
      establishment itself, confirm no new code duplicates that path.
      Per-slide loop over every other slide in scope:
        - Skip (report "already linked") if the slide already carries
          `instance_key` + type tag — read via `Resolve.ResolveSlideInstance`,
          already handles this exact check (`HasInstanceKey`/`HasTypeTag`).
        - Score via `Onboarding.MatchSlideAgainstTemplate` unchanged (no new
          matching logic, spec is explicit about this).
        - Dispatch on confidence using the same thresholds `onboarding.md`
          already defines (high → ready, medium → needs `ConfirmFieldMatch`,
          low/none → excluded/unclassified) — reuse `MatchResult.Confidence`
          directly, don't reimplement threshold comparison.
        - Harvest verbatim values from matched fields (same technique
          `InjectPrimitive.bas` already uses for reading current text —
          `shp.TextFrame.TextRange.Text` — confirmed no dedicated "harvest"
          function exists yet anywhere in `vba/`; this task is where one first
          gets written, generalized enough that Priority 21's onboarding verify
          step could reuse it too if convenient, though that's not required).
        - Instance-key resolution against existing keyless Data-sheet rows: this
          is the one piece of genuinely new decision logic in the whole spec (not
          present in `onboarding.md` or `sync_operations.md` at all — those never
          handle "a Data-sheet row that predates any linked slide"). Needs
          `ExcelOutput.ReadSheet`'s row data plus an exact-match comparison
          against each keyless row's non-key fields; zero-or-multiple matches
          always falls back to a fresh row per the spec's explicit
          never-guess rule.
      One phase-gate review (all slides in scope, proposed disposition, proposed
      instance_key) before any write — mirrors the Review-form pattern Priority 21
      needs anyway for onboarding; consider sharing form-building code between the
      two rather than writing two divergent review UIs, but the underlying data
      each needs to display differs (per-field for onboarding vs. per-slide for
      this), so don't force a shared abstraction if it doesn't fit cleanly.
      Row-order bootstrap: newly created rows appended to the Data sheet in
      current deck order of their source slides — a one-time exception, confirm
      this doesn't fight `RunSync.ResequenceByRowOrder`'s normal invariant (which
      takes over immediately after, per the spec).
      Verify-the-link: reuse `InjectPrimitive`'s no-op round-trip per linked slide
      (same check `onboarding.md` Step 6 already does for the template), generalized
      across the batch; any non-no-op result is a harvest bug this pass must flag
      and stop on, not something a later sync silently "corrects."
      Needs `SPIKE_NOTES_DeckAdoption.md` (new file, same format as every other VBA
      module) with a manual verification recipe — no automated harness exists for
      this either, same constraint as Priority 21.
      **Deliberately excluded per the spec's own Non-goals at the time**: no UI
      for this yet (`ribbon-ui.md` gets an "Adopt Existing Slides" entry point
      in a *future* pass, per that spec's own Non-goals section) — build the
      engine layer only, callable from the VBE, same as every other module before
      its own ribbon wiring landed. **That future pass is Priority 25** (same
      day, 2026-07-26) — `vba/AdoptFlow.bas` + a fifth toolbar button now wire
      this in. No mixed-type auto-classification, no multi-period-history
      reconstruction, no whole-deck implicit scanning remain true non-goals.

      Built `vba/DeckAdoption.bas` as a plan/commit pair, mirroring this project's
      own existing `SyncOperations.Plan*` / `RunSync.Run*` split rather than
      inventing a new shape: `PlanAdoption(slidesToAdopt(), templateSld, ws,
      ByRef harvestedValues())` writes nothing and returns one `AdoptionSlidePlan`
      per slide (`Disposition`: `already_linked` / `ready` / `needs_confirmation` /
      `unclassified`, plus `MatchedKeylessRowId` and a human-readable `Reason`) —
      inspecting this array *is* the phase-gate review, standing in for the
      not-yet-built form the same way `RunRoutineSync`'s report string already
      does for routine sync. `CommitAdoption(plans(), slidesToAdopt(),
      harvestedValues(), confirmedInstanceKeys(), slideType, templateSld, ws)` is
      the only function that writes: reuses `Onboarding.OnboardNewInstance`
      unchanged for tagging (it already does exactly "tag identity + auto-accept
      every high-confidence field," which is precisely what a `ready` slide
      needs — no new tagging logic was written), `ExcelOutput.UpsertRow` for the
      Data-sheet write (linking into a matched keyless row by writing its
      Instance ID cell directly first, so `UpsertRow` finds and merges into it
      rather than appending a duplicate), and `InjectPrimitive.InjectPrimitive`'s
      no-op round trip for the verify-the-link check. `ready`/`needs_confirmation`/
      `unclassified` dispatch is a per-slide aggregation of `Matching.Match`'s
      existing per-field confidence (spec says "same thresholds as onboarding.md"
      but those are inherently per-field, not per-slide) — documented as a real,
      un-pinned-down judgment call in `SPIKE_NOTES_DeckAdoption.md`, same posture
      as `RunSync.bas`'s resequencing-anchor choice. `ReadKeylessRows`/
      `FindMatchingKeylessRow` close the one genuinely new gap the spec calls
      out: `ExcelOutput.ReadSheet` deliberately excludes blank-Instance-ID rows
      (correct for every existing caller, but exactly the rows this task needs to
      see) — read locally in `DeckAdoption.bas` via `ws.UsedRange` rather than
      reopening `ExcelOutput.bas`'s already-shipped read contract, and rather
      than reusing its `Cells(Rows.Count,1).End(xlUp)` idiom (which walks column
      A specifically, exactly the column a keyless row has nothing in).
      `AdoptionSlidePlan` is deliberately scalar-only (no member is a dynamic
      array of another UDT) — this project has never exercised that construct in
      a real Office run, and given how many genuine Office-specific VBA gotchas
      this project has already hit the hard way (`AGENTS.md`'s Known Patterns),
      untested territory was avoided where the design didn't strictly require it;
      `CommitAdoption` re-derives field-shape matches via `OnboardNewInstance`'s
      own internal `MatchSlideAgainstTemplate` call instead of threading
      `PlanAdoption`'s match results through. Added 6 tests to
      `vba/tests/TestRunner.bas`: idempotent already-linked skip, a
      high-confidence slide committed end-to-end (tags + fresh Data-sheet row +
      verified), a medium-confidence slide correctly left fully untouched (no
      partial tag write), a fully unrelated slide excluded, linking into a
      pre-existing keyless Data-sheet row without creating a duplicate row, and
      (added after an adversarial review pass, see below) a 0-based 3-slide
      batch mixing all three live dispositions.

      A dedicated review pass (this container has no Office to compile against,
      so a careful line-by-line manual trace stood in for a compiler) caught one
      real, un-triggered bug before commit: `PlanAdoption` originally tracked
      its returned `AdoptionSlidePlan()` with a separate 1-based counter while
      `harvestedValues()` kept whatever index base the caller's
      `slidesToAdopt()` used, and `CommitAdoption` indexed both (plus
      `confirmedInstanceKeys()`) with one shared loop variable — correct only
      for a 1-based `slidesToAdopt`, which every original test happened to use,
      so nothing caught it. Fixed by allocating `plans` over `slidesToAdopt`'s
      own `LBound`/`UBound` directly rather than a separate counter, and added
      the 0-based regression test above specifically to catch a recurrence.
      Full account in `SPIKE_NOTES_DeckAdoption.md`'s new "A real bug found and
      fixed during review" section.
      **Not executed against real Office this pass** — `powershell.exe`
      unreachable from this plain-Linux container (confirmed via `which
      powershell.exe`), same constraint as Priority 20a; needs a real
      `run_vba_tests.ps1` run on the WSL/Windows host next time this project is
      picked up there. `python3 -m pytest tests/ -v` (70 passed) and `python3 -m
      mypy src/` (no issues, 10 source files) both confirmed unaffected
      (VBA-only change). Full design rationale, the plan/commit phase-gate
      design, and the manual verification recipe are in the new
      `SPIKE_NOTES_DeckAdoption.md`.

## Notes carried forward from Priority 19 (still open, unaffected by this pass)
- Cases 5 (`record_retired`) and 7 (`deck_side_conflict`) remain non-goals
  throughout — no spec has been written for either.
- The parked multi-deck design (export, conflict-resolution/propagation,
  stale-queue) remains entirely unspecified.
- Where a type's template actually lives/is looked up is still caller-supplied
  everywhere (`RunRoutineSync`'s `templateSld` parameter, `onboarding.md`'s own
  non-goal) — Priority 21's ribbon buttons will need *some* answer for this to be
  clickable at all (a picker, a naming convention, a stored reference — not
  decided by any spec yet); flagged again here since Priority 21 is the first
  place this actually has to be resolved concretely rather than left to a caller
  that doesn't exist yet.
- This pass found no regressions and no other undocumented gaps in `src/*.py` or
  the 6 already-ported core VBA modules — confirmed via full `pytest`/`mypy` run
  (green) and re-reading `AGENTS.md`'s Known Patterns/Constraints for anything
  contradicted by current code (none found).

---

# Implementation Plan (history below this line)

Generated by gap analysis of specs/* against src/* (2026-07-21). Confirmed by
reading src/discovery.py, tests/test_discovery.py, both test-fixtures/*.pptx
files (unzipped and inspected directly), and by actually running
`python3 -m pytest tests/ -v` and `python3 -m mypy src/`.

Current state: only `src/discovery.py` exists. `matching.py`, `verification.py`,
and an excel-output module do not exist yet (confirmed via `find`/`grep` — no
stubs, TODOs, or partial files anywhere in the repo). Tests pass (3/3) but
`mypy src/` currently **fails** — this is a real, pre-existing gap, not a
hypothetical one.

## Priority 1: Fix and harden discovery (foundational — everything else reads its output)

- [x] Fix the `mypy src/` failure in `src/discovery.py:76` (why: `spTree = slide_xml_root.find(...)`
      returns `Element | None`, but `walk()` requires `Element`; `discover()` calls
      `walk(spTree, ())` unchecked. AGENTS.md's mandated validation command is
      `python3 -m mypy src/` and PROMPT_build.md step 4 requires validation to pass
      before every commit — this blocks every future iteration's ability to commit
      cleanly until fixed. Confirmed by running mypy directly, not assumed.)
      Fixed by raising `ValueError` when `spTree` is `None` instead of walking it
      unchecked — narrows the type for mypy and matches discovery's "never silently
      guess" ethos (a slide XML with no `p:spTree` is malformed input, not something
      to paper over). Confirmed `python3 -m pytest tests/ -v` (3 passed) and
      `python3 -m mypy src/` (no issues) both pass.

- [x] Capture actual placeholder `type`/`idx` attributes on `Candidate`, not just the
      current `has_placeholder: bool` (why: specs/matching.md's most reliable scoring
      signal is "layout placeholder index (if applicable, most stable)" — a boolean
      can't express an index. Confirmed by reading `mst-slide-layouts.pptx`'s
      `slideLayout1.xml`/`slideLayout2.xml`: both contain `<p:ph type="title"/>` and
      `<p:ph type="body" sz="quarter" idx="10"/>` with real type/idx values that a
      future matching module will need. `_has_placeholder()` in discovery.py currently
      discards this data.)
      Replaced `has_placeholder: bool` with `placeholder_type: str | None` and
      `placeholder_idx: int | None` on `Candidate`; `has_placeholder` is now a derived
      property (`placeholder_type is not None`) kept for convenience. New
      `_placeholder_info()` reads `<p:ph>` and applies OOXML's own attribute defaults
      (`type` defaults to `"obj"`, `idx` defaults to `0`) when the element is present
      but the attribute is omitted — the element's mere presence, not its attributes,
      is what marks a shape as a placeholder. Tested against
      `mst-slide-layouts.pptx`'s `slideLayout1.xml` by calling `discover()` directly on
      the parsed layout root: `discover()` is root-agnostic (only looks for
      `.//p:spTree`), so this works even before a dedicated slideLayout loader exists
      (that loader is the next task below). Confirmed `python3 -m pytest tests/ -v`
      (5 passed) and `python3 -m mypy src/` (no issues) both pass.

- [x] Add discovery support and tests for `test-fixtures/mst-slide-layouts.pptx` (why:
      this fixture is currently unused by any test — confirmed via grep across
      `tests/`. It is also structurally different from what `discover_from_pptx`
      assumes: unzipping it shows it has **no `ppt/slides/*` entries at all**, only
      `ppt/slideLayouts/slideLayout{1,2}.xml` and `ppt/slideMasters/slideMaster1.xml`.
      `discover_from_pptx()` hardcodes `ppt/slides/slide{slide_index}.xml` and will
      raise `KeyError` on this fixture as-is. `discover()` itself is root-agnostic
      (just looks for `.//p:spTree`) so it likely works unchanged on a slideLayout
      XML root — the gap is the loader, not the walker. Needs a loader that can target
      `ppt/slideLayouts/slideLayoutN.xml`, plus tests asserting placeholder type/idx
      are correctly discovered per spec's "whether it carries a placeholder type/index"
      metadata requirement, per test-fixtures/SOURCE.md's stated purpose for this file.)
      Refactored `discover_from_pptx` to delegate to a new generic
      `discover_from_pptx_part(path, part_name)` that opens any shape-tree-bearing
      zip member, and added `discover_from_pptx_layout(path, layout_index=1)` on top
      of it targeting `ppt/slideLayouts/slideLayoutN.xml`. Replaced the earlier
      hand-rolled `ZipFile`/`ET.parse` test (which bypassed the loader gap entirely)
      with one using the new loader, plus a `KeyError`-on-`ppt/slides/*` regression
      test confirming the fixture really does lack slide entries, and a second-layout
      test confirming the loader isn't hardcoded to index 1. Confirmed
      `python3 -m pytest tests/ -v` (7 passed) and `python3 -m mypy src/` (no issues)
      both pass.

## Priority 2: Matching module (specs/matching.md)

- [x] Implement `src/matching.py` per specs/matching.md, depends on Priority 1's
      placeholder type/idx data being available on `Candidate` (why: nothing in
      src/ implements matching yet — confirmed via `find`/`grep`, no stub exists).
      Required behavior per spec:
  - Two-tier: trust an existing valid identity tag directly, no scoring; only score
    when untagged.
  - Scored fallback in reliability order: placeholder index > geometric similarity
    (position/size within tolerance) > shape type > content pattern (weakest,
    last resort) — combined into a single confidence score, never take the single
    best-scoring candidate blindly.
  - Confidence thresholds: high = auto-accept, medium = flag for human confirmation
    (never silently guess), low = unmatched (never force onto nearest candidate).
  - Sibling ambiguity: when untagged candidates score closely to each other, treat
    the closeness itself as signal (flag, don't arbitrarily break the tie); add
    z-order as a supplementary disambiguating signal before falling back to "flag."
      Extended `Candidate` (discovery.py) with `position`/`size` (EMU, read from each
      shape's own `<p:spPr>/<a:xfrm>` — deliberately does not walk up through parent
      group transforms; documented as exact only when a group's chOff/chExt equals its
      own off/ext, true for every current fixture) and `identity_tag: str | None = None`
      (always None out of `discover()` today, since discovery still doesn't read/write
      tags per its non-goals — the field exists purely so `matching.py`'s tier-1 logic
      has somewhere to look). `matching.py`'s `match(candidates, reference, valid_tags)`
      implements: tier-1 tag trust (single valid tag wins immediately; >1 same-tag
      candidates is a flagged collision); tier-2 `score_candidate()` combining
      placeholder-index/geometry/shape-type/content-has-text signals by weight
      (0.5/0.3/0.15/0.05), renormalized over only the signals applicable to the
      reference (e.g. geometry skipped, not zeroed, if the reference carries none);
      high/medium/low thresholds at 0.75/0.4; sibling-ambiguity handling that flags any
      set of candidates scoring within 0.1 of the top score unless z-order uniquely
      picks one winner among them. Confirmed `python3 -m pytest tests/ -v` (20 passed)
      and `python3 -m mypy src/` (no issues) both pass.

- [x] Add matching tests using `shp-groupshape.pptx` (sibling ambiguity / z-order
      disambiguation — this fixture's whole SOURCE.md purpose per its own
      description) and `mst-slide-layouts.pptx` (placeholder-index tier) once
      Priority 1's loader work lands.
      Added `tests/test_matching.py` (11 tests): tier-1 trust/collision using synthetic
      `Candidate`s; `shp-groupshape.pptx`'s 4 tied leaf shapes (identical shape_type/
      has_text) to exercise a genuine 4-way score tie resolved by z-order, plus a 2-way
      case where z-order also ties and must be flagged, not guessed; `mst-slide-
      layouts.pptx`'s two layouts' body placeholder (same idx=10, very different
      geometry) to show placeholder-index match alone correctly lands at medium
      confidence (flagged) rather than forcing an auto-accept. Also added 3 tests to
      `tests/test_discovery.py` covering the new `position`/`size`/`identity_tag`
      fields (position/size checked against both fixtures' raw XML `a:off`/`a:ext`
      values directly). Confirmed `python3 -m pytest tests/ -v` (20 passed) and
      `python3 -m mypy src/` (no issues) both pass.

## Priority 3: Verification module (specs/verification.md)

- [x] Implement `src/verification.py`'s core `inject_primitive` operation: hash the
      shape's current value and the linked source value; no-op if equal; if
      different, write the source value then re-hash to confirm the write actually
      took (why: nothing in src/ implements this yet — confirmed via `find`/`grep`).
      This needs a shape-value read/write primitive that doesn't exist yet in
      discovery.py (which is read-only) — will likely need `zipfile` write-back of
      modified slide XML, following the same stdlib-only constraint AGENTS.md
      documents for discovery.py.
      `inject_primitive(path, part_name, shape: Candidate, source_value)` locates the
      shape by re-walking to `shape.z_order` (same numbering discover() already
      assigns), reads its concatenated `<a:t>` text, sha256-hashes it against
      `source_value`'s hash; no-op (zero bytes written) if equal. If different, writes
      `source_value` into the shape's first text run (clearing any extra runs so the
      concatenation stays exact), rewrites the single changed zip entry via a
      temp-file-then-`os.replace` swap (every other entry copied byte-for-byte,
      preserving `p:`/`a:` namespace prefixes via `ET.register_namespace`), then
      re-opens the file from disk and re-hashes the written-back value — never assumes
      the write succeeded from the write call alone. Raises `ValueError` (does not
      silently no-op or invent a run) when a shape has no `<a:t>` run to write into,
      per `shp-groupshape.pptx`'s empty-decoration shapes. Added
      `tests/test_verification.py` (4 tests) using a temp copy of
      `mst-slide-layouts.pptx`'s title placeholder (real seed text "Click to edit
      Master title style"): no-op leaves the file byte-identical to the original
      fixture; a real write is confirmed both via a fresh read and via
      `result.verified`; every other zip entry is confirmed byte-for-byte untouched
      after a write (no data loss from the rewrite); and the no-text-runs case on
      `shp-groupshape.pptx` raises rather than corrupting. Confirmed
      `python3 -m pytest tests/ -v` (24 passed) and `python3 -m mypy src/` (no issues)
      both pass.

- [x] Implement structural verification after duplication: shape count, type, and
      identity-tag correspondence between a duplicate and its source, checked
      explicitly rather than assumed from the duplication API succeeding (why: spec
      explicitly calls out that duplication succeeding is not sufficient evidence).
      No duplication mechanism or fixture currently exists in the repo — will likely
      need a small synthetic fixture (a two-slide pptx where one slide is a
      duplicate of the other) since none of the pulled python-pptx fixtures cover
      this; note this as a new-fixture task rather than assuming an existing one
      applies.
      Added `verify_structure(source, duplicate)` (plus `verify_structure_from_pptx`
      convenience wrapper) to `src/verification.py`, operating on two `discover()`
      output lists paired positionally by z_order — the same canonical ordering
      `_find_shape_by_z_order` already relies on. Reports every count/type/tag
      mismatch as a `StructuralMismatch` rather than stopping at the first one or
      silently truncating to the shorter list when counts differ (`missing_in_duplicate`
      / `extra_in_duplicate` kinds cover the leftover positions explicitly). No
      multi-slide fixture existed in test-fixtures/, so tests synthesize one: a temp
      copy of `shp-groupshape.pptx` with a second `ppt/slides/slide2.xml` zip entry
      appended (byte-identical, a dropped-shape mutation, and a retagged-shape-type
      mutation) — valid since `discover_from_pptx_part()` only ever opens a named zip
      member directly and never validates `[Content_Types].xml`/relationships. The
      identity-tag correspondence case uses directly-constructed `Candidate`s (like
      `test_matching.py`'s tier-1 tests) since `identity_tag` is always `None` out of
      `discover()` — no physical tag storage format is decided yet. Confirmed
      `python3 -m pytest tests/ -v` (29 passed) and `python3 -m mypy src/` (no issues)
      both pass.

- [x] Implement the z-order check as a check distinct from value/tag correspondence
      (why: spec explicitly warns these are different claims — a duplicate can have
      correct shapes/tags/values while a stacking-order regression still makes an
      overlaid field invisible, e.g. transparent text box behind its background).
      `verify_z_order(source, duplicate)` (plus `verify_z_order_from_pptx`) pairs
      shapes by `identity_tag` (untagged shapes excluded — no reliable
      correspondence signal for them) and compares every pair of commonly-tagged
      shapes' relative stacking order (not just adjacent ones), so a swap deep in
      the stack is caught regardless of how many shapes sit between the two that
      moved. The code and its tests already existed on disk from the prior
      iteration but were uncommitted-plan-wise; running the suite surfaced a real
      bug in the already-checked-off `verify_structure` (Priority 3, first
      structural-verification task above): it paired shapes purely by list
      position, so a pure reorder of tagged shapes (same shapes/tags/values, just
      restacked) *also* tripped its tag-correspondence check, since a shape's tag
      moves with it when the list is reordered — directly contradicting the
      spec's claim that structural and stacking correctness are separate claims a
      duplicate can satisfy independently. Fixed by having `verify_structure` pair
      tagged shapes by `identity_tag` (order-independent) and only fall back to
      positional pairing for untagged shapes, which have no other correspondence
      signal. This changes one existing test's semantics:
      `test_verify_structure_flags_an_identity_tag_mismatch`'s single relabeled
      shape now correctly reads as "expected tagged shape gone" +
      "unexpected tagged shape appeared" (`missing_in_duplicate`/
      `extra_in_duplicate`) rather than a same-position `identity_tag` mismatch,
      since tag-based pairing has no positional signal to distinguish "renamed"
      from "swapped for something else" — updated its assertions accordingly, and
      removed the now-impossible `"identity_tag"` mismatch kind. Confirmed
      `python3 -m pytest tests/ -v` (33 passed, was 32 passed/1 failed before this
      fix) and `python3 -m mypy src/` (no issues) both pass.

## Priority 4: Excel output module (specs/excel-output.md)

- [x] Implement an excel-output module (e.g. `src/excel_output.py`) producing the
      structure specs/excel-output.md defines (why: nothing in src/ implements this
      yet — confirmed via `find`/`grep`). Required behavior:
  - One column per confirmed field, named by stable field identity (not
    regenerated per slide instance).
  - One row per slide instance keyed by persistent instance identity, never by
    position/order.
  - Seed rows populated from harvested (onboarding-time) values, not left blank.
  - No data loss on write: adding a column/row must never silently drop or
    overwrite an existing column/row it wasn't asked to touch — this implies a
    read-merge-write pattern, not a blind overwrite.
  - A stable reference back to the specific paired deck (not inferred solely from
    column-name matching against multiple candidate decks).
  - Note: AGENTS.md says no pip packages beyond pytest/mypy are installed and this
    project deliberately stays dependency-light (no openpyxl). `.xlsx` is itself a
    zipped OOXML package, so this can likely follow the same `zipfile` +
    `xml.etree.ElementTree` approach discovery.py already uses — confirm this
    holds before reaching for a dependency, per AGENTS.md's stated preference.
  - No existing `.xlsx` test fixtures exist in `test-fixtures/` — will need to
    create minimal synthetic ones (or generate the minimal valid OOXML by hand
    to test against, since none were pulled from upstream for this spec.

      Implemented `src/excel_output.py` fully stdlib (zipfile + xml.etree), hand-
      writing the minimal valid `.xlsx` OOXML package: `[Content_Types].xml`,
      `_rels/.rels`, `xl/workbook.xml`(+`.rels`), `xl/styles.xml` as static
      boilerplate (never varies by data), plus `xl/worksheets/sheet1.xml` and
      `docProps/custom.xml` built dynamically via ElementTree. Uses inline
      strings (`t="inlineStr"`) rather than a shared-strings table, so a write
      never has to merge/dedupe a separate string part — one less moving piece
      for a read-merge-write design to get wrong. Column A is a reserved,
      position-addressed "Instance ID" column (a structural convention of this
      module, not a field subject to the spec's name-based-identity rule);
      confirmed fields occupy columns B.. in first-seen order, looked up by
      header text on every read, never assumed from position. The deck
      reference lives in `docProps/custom.xml` as a real OOXML custom document
      property — a dedicated metadata slot, not a row/column that could collide
      with real data, satisfying "never inferred solely from column-name
      matching." `create_sheet(path, deck_reference)` makes a fresh empty sheet
      and refuses to overwrite an existing file (a second "create" would
      silently discard whatever's already there). `upsert_row(path, instance_id,
      values)` is the read-merge-write primitive: unknown field names in
      `values` are appended as new columns (existing columns never reordered or
      touched); a new `instance_id` is appended as a new row seeded entirely
      from `values`; an existing `instance_id` is merged field-by-field, so a
      partial re-sync (e.g. one changed field) can never blank out that
      instance's other already-populated fields. A field absent from a given
      instance's dict renders as an omitted cell, not a forced blank. Whole-file
      writes go through a temp-file-then-`os.replace` swap, same safety pattern
      `verification.py`'s `_write_part` uses. No `.xlsx` fixtures existed to test
      against (none pulled from upstream for this spec, per the note above still
      accurate at the time this task started) — `tests/test_excel_output.py`
      round-trips writer against reader instead (8 tests): deck-reference
      round-trip, create-refuses-overwrite, seed-from-harvested-values,
      column-append-without-touching-existing-data, partial-update-merges,
      new-instance-doesn't-disturb-existing-rows, field/instance order
      preserved across multiple writes, and column-A-reservation checked
      directly against the written XML. Confirmed `python3 -m pytest tests/ -v`
      (41 passed) and `python3 -m mypy src/` (no issues, 4 source files) both
      pass.

## All tasks complete (Priority 1-4)

Every Priority 1-4 item above is checked off. No further gaps identified
against specs/discovery.md, specs/matching.md, specs/verification.md, or
specs/excel-output.md as of this pass — see "Notes for next planning pass"
above for open design questions (identity-tag physical storage format,
`src/lib/` promotion) that weren't blocking but are worth a decision if this
project continues.

## Priority 5: Sync operations decision tree (specs/sync-operations.md)

- [x] Implement dispatch logic for cases 1 (no_change), 3 (new_record), and 4
      (in_place_correction) per the underlying `crc-vba-deck-sync` skill's
      `sync-cases.md`/`run-sync.md`, translated into specs/sync-operations.md.
      `plan_routine_sync(path, instances, data_sheet)` reuses existing
      primitives directly rather than reimplementing anything: `inject_primitive`'s
      own no-op/write result is the case-1-vs-4 classifier; a Data-sheet row
      whose instance key matches no known `SlideInstance` is case 3 (the
      decision only — physically duplicating a slide is a non-goal, see below).
      Case 6 (unclassified_slide) is also implemented, cheaply: type is an
      explicit declaration per input-contract.md, so a missing/unrecognized
      type tag or instance key is a trivial flag, not a confidence score.
      Case 2 (period_rollover) is `plan_period_rollover(instance, new_values)`,
      a separate, explicitly-invoked function never reachable from
      `plan_routine_sync` — matches the source design's "never inferred from a
      value looking different" rule (tested directly:
      `test_period_rollover_never_produced_by_routine_sync`).
      **Deliberately not built** (see specs/sync-operations.md's Non-goals for
      the reasoning): physical slide duplication (a distinct, harder
      OOXML-surgery problem), case 5 record_retired (no agreed
      retirement-status convention exists in the source design), case 7
      deck_side_conflict (needs a last-synced-value store nothing in this
      project persists today — `inject_primitive` only ever does a 2-way
      current-vs-target comparison, not the 3-way one this needs). All three
      are open design questions to raise before building, not oversights.
      `SlideInstance` (the input this module operates on) is deliberately an
      already-resolved dataclass, not something this module reads off a real
      pptx itself — instance_key/type_tag/field-to-shape correspondence still
      depends on the identity-tag physical storage format that's remained an
      open question since Priority 2-3 (see "Notes for next planning pass"
      below). Confirmed `python3 -m pytest tests/ -v` (49 passed) and
      `python3 -m mypy src/` (no issues, 5 source files) both pass.

## Priority 6: Identity tag physical storage (specs/identity-tags.md)

- [x] Resolve and implement the physical storage format for `identity_tag`,
      open since Priority 2-3 (see below). Verified against ECMA-376's User
      Defined Tags Part definition and a real-world example rather than
      guessed, since python-pptx has no built-in support for this and no
      fixture on disk carried any tags to reverse-engineer from (citations
      in specs/identity-tags.md). `src/identity_tags.py`: slide-level tags
      (`slide_type`/`instance_key`/`period_key`) via a direct relationship
      from the slide part's own `.rels` to a Tags Part (`ppt/tags/tagN.xml`,
      `<p:tagLst>`/`<p:tag>`); shape-level tags (`role`) via the shape's own
      `<p:nvPr>/<p:custDataLst>/<p:tags r:id>`, resolved through the owning
      slide's `.rels` to the same kind of Tags Part (a shape has no `.rels`
      of its own). Respects `CT_ApplicationNonVisualDrawingProps`' schema
      child order (`ph, media, custDataLst, extLst`). `upsert_slide_tags`/
      `upsert_shape_tags` are read-merge-write, same pattern as
      `excel_output.upsert_row`: first write creates the Tags Part +
      relationship + `[Content_Types].xml` override; later writes reuse them
      idempotently. Confirmed `python3 -m pytest tests/ -v` (59 passed) and
      `python3 -m mypy src/` (no issues, 6 source files) both pass.
      **Not yet done**: wiring this into an actual end-to-end "resolve a real
      slide into a `sync_operations.SlideInstance`" function (compose
      `discover()` + `identity_tags.read_*` + `matching.match()`) — every
      module so far is individually tested but nothing composes them against
      a real deck yet. Flagged as the next real gap, not an oversight.

## Priority 7: src/lib promotion + real end-to-end composition proof

- [x] Promote the two real duplications flagged after Priority 6 into
      `src/lib/ooxml.py`: `find_shape_element_by_z_order()` (was
      `verification.py`'s `_find_shape_by_z_order` and the tree-walk half of
      `identity_tags.py`'s `_find_nvpr`) and `write_zip_parts()` (was
      `verification.py`'s single-part `_write_part` and `identity_tags.py`'s
      multi-part `_write_parts` — the same shape once generalized to a dict
      of updates). Both call sites refactored to import from `lib.ooxml`
      rather than keeping private copies. `excel_output.py`'s `_write_xlsx`
      deliberately kept separate — it always rebuilds every part from
      scratch (it fully owns its file, no "preserve untouched existing
      parts" concern the way a `.pptx` we don't fully own has), so forcing
      it onto the shared helper would be the wrong generalization.
- [x] Close the integration gap flagged after Priority 6: `src/resolve.py`'s
      `resolve_slide_instance(path, slide_part)` composes
      `discover_from_pptx_part()` + `identity_tags.read_slide_tags()` +
      `identity_tags.read_shape_tags()` into a real `sync_operations.
      SlideInstance` — tier-1 trust only (a shape reaching this function is
      assumed already tagged from onboarding; scoring untagged candidates
      against a reference is matching.py's tier-2 path and needs a
      per-type reference-shape configuration that doesn't exist anywhere in
      this project yet — deliberately not invented here, flagged instead).
      `tests/test_resolve.py` is the actual end-to-end proof this project was
      missing: tag a real fixture on disk via `identity_tags`, resolve it,
      run it through `sync_operations.plan_routine_sync()`, and confirm
      real dispatch against a real Data-sheet row — no_change when the seed
      value matches, in_place_correction (with the write actually landing on
      disk, re-read and confirmed) when it doesn't, new_record for an
      unmatched row, and unclassified_slide flagging for a never-onboarded
      slide. Confirmed `python3 -m pytest tests/ -v` (64 passed) and
      `python3 -m mypy src/` (no issues, 9 source files) both pass.

## Priority 8: Onboarding — matching a new slide against an established template

- [x] Close the onboarding gap flagged after Priority 7. Reading the underlying skill's
      own `onboard-slide-type.md` first resolved the "reference config storage" question
      it seemed to imply: **first-time onboarding of a type needs no scoring or new code
      at all** — "the working copy IS becoming the reference," tagged directly after a
      human confirms it, verified by `inject_primitive` hitting the no-op path (already
      fully covered by `tests/test_resolve.py`'s no-change case). The actual gap was
      narrower: matching a *subsequent* slide of an already-established type against its
      template. `src/onboarding.py`: `match_slide_against_template()` scores every
      untagged candidate on a new slide against each of the template's field roles via
      `matching.match()`'s tier-2 path; `onboard_new_instance()` tags the new instance's
      slide-level identity unconditionally (supplied by whatever created it, e.g. a
      duplication — not matched) and auto-accepts + tags any high-confidence field match
      immediately (self-healing into a tier-1 fast match next time, per
      specs/matching.md's confidence_thresholds), leaving medium/low-confidence matches
      unresolved; `confirm_field_match()` is the explicit primitive a human's decision
      (or an eventual shape-selection UI — see Non-goals) calls to resolve one. Tested
      against real fixture drift, not synthetic scores: `mst-slide-layouts.pptx`'s two
      layouts' title placeholder cross-matches at high confidence (auto-accepted), its
      body placeholder at medium (correctly left flagged, per the geometry drift
      `test_matching.py` already established), then resolved via `confirm_field_match()`.
      **Deliberately not built**: the selection UI/mechanism itself (a human clicking a
      shape in PowerPoint) — this Python reference implementation has no real deck-editing
      UI to select from; this module builds the primitive that UI would call, not the UI.
      Confirmed `python3 -m pytest tests/ -v` (68 passed) and `python3 -m mypy src/` (no
      issues, 10 source files) both pass.

## Priority 9 (2026-07-24 pass): VBA spike follow-up + one stale doc found

Re-verified this pass by actually running `python3 -m pytest tests/ -v` (68 passed) and
`python3 -m mypy src/` (no issues, 10 source files) — both still green, no regressions
since Priority 8. Confirmed via `git log`/`git show` that one thing changed since this
plan was last written: commit `333c53a` added `vba/InjectPrimitive.bas` +
`vba/SPIKE_NOTES.md`, a hand-port of `inject_primitive` to VBA. That spike is real,
already self-documents its own scope/divergences/manual-verification recipe in
`vba/SPIKE_NOTES.md`, and needs no Python-side work — noted here for context, not as an
open task. Gap analysis below is what's actually new/actionable.

- [x] Fix `src/resolve.py`'s module docstring (lines 9-15): it says scoring untagged
      candidates against a reference "needs a per-type reference configuration that
      doesn't exist anywhere in this project yet" and calls this "a real, separate gap."
      That gap is already closed — confirmed by reading `tests/test_onboarding.py`
      (`_onboard()` helper at line 47 and every test after it): `onboarding.py`'s
      `match_slide_against_template(path, slide_part, template: SlideInstance)` takes
      exactly the `SlideInstance` shape `resolve_slide_instance()` already produces, and
      the tests literally call `resolve_slide_instance(path, LAYOUT1_PART)` to build the
      `template` argument passed into it. `resolve.py`'s docstring was written in commit
      `637a130`, before `onboarding.py` existed (`b7f07b4`, later) — it's stale
      documentation actively misleading a future reader into re-solving an already-solved
      problem, not a real gap. Fixed: replaced the "real, separate gap" paragraph with a
      pointer to `onboarding.py`'s actual resolution (an onboarded template slide, run
      through `resolve_slide_instance()` itself, *is* the per-type reference — no
      separate config format was ever needed). Pure doc fix, no behavior change. Confirmed
      `python3 -m pytest tests/ -v` (68 passed) and `python3 -m mypy src/` (no issues, 10
      source files) both still pass after editing.

- [x] Decide and document whether `vba/` follows the specs-driven process the rest of
      this repo uses (why: every file under `src/` traces to a `specs/*.md` file that
      defines its scope, requirements, and non-goals before code is written — confirmed
      by checking all 7 specs against all 10 `src/*.py`/`src/lib/*.py` files, every one
      has a governing spec. `vba/InjectPrimitive.bas` has no `specs/*.md` counterpart; its
      scope lives only in its own `SPIKE_NOTES.md`, self-authored alongside the code
      rather than preceding it.)
      Decision: no further VBA porting (discovery, matching, or sync-dispatch logic) is
      currently planned or wanted beyond the existing `InjectPrimitive.bas` spike — there
      is no task anywhere in this plan calling for more VBA work, so writing a
      `specs/vba-port.md` now would be speculative scope for hypothetical future work.
      Documented this explicitly in `AGENTS.md`'s Constraints section instead (a new
      bullet after the existing "reference/test implementation... real target is VBA"
      one), including the fallback instruction: if more VBA porting *is* wanted later,
      write `specs/vba-port.md` first, mirroring how every other module here started from
      a spec. This closes the open question so a future planning pass doesn't have to
      re-derive it from git history again. Pure doc change, no code touched. Confirmed
      `python3 -m pytest tests/ -v` (68 passed) and `python3 -m mypy src/` (no issues, 10
      source files) both still pass.

## Priority 12 (2026-07-24 pass): VBA port order, module 1 (discovery)

- [x] Port `discovery` per `specs/vba-port.md`'s numbered port order (module 1 of 6;
      module 2, `identity_tags`, was already done out-of-order as the original spike —
      confirmed nothing else in that order has been started: `ls vba/` before this task
      showed only `InjectPrimitive.bas`/`SPIKE_NOTES.md`). Added `vba/Discovery.bas`:
      `DiscoverSlide(sld)`/`DiscoverCustomLayout(layout)` walk the live `Shapes`/
      `GroupShapes` collections recursively (no XML at all, unlike the Python original,
      which has no host application to ask), reproducing `discovery.py`'s actual walked
      scope — including the non-obvious detail that it only ever recognizes
      `<p:grpSp>`/`<p:sp>`/`<p:pic>`, so tables/charts/connectors/OLE objects are already
      invisible to the Python version regardless of content, not merely "type-agnostic"
      per discovery.md's requirement language taken alone. Added
      `vba/SPIKE_NOTES_Discovery.md` (mirroring `SPIKE_NOTES.md`'s format) documenting
      real, unresolved divergences found while porting rather than glossed over:
      **`PlaceholderIdx` cannot be ported at all** — PowerPoint's object model has no
      `Shape.PlaceholderFormat.Idx`-equivalent (unlike `Shape.Tags`, which the earlier
      spike found *does* have a native equivalent), so this is flagged as real
      unfinished work for whichever module needs it first (almost certainly `matching`,
      port-order step 3, since idx is `matching.md`'s top-weighted signal) rather than
      invented or silently dropped; placeholder *type* is a best-effort enum-to-string
      mapping, not authoritative; position/size convert points→EMU (×12700) and are not
      always bit-exact against Python's raw integer EMU reads; whether VBA already
      resolves nested-shape position to slide-absolute coordinates (which would make
      this port *more* correct than Python's documented local-offset-only limit) is
      flagged explicitly as unconfirmed, not assumed either way. Manual verification
      recipe cross-checks against both existing fixtures' exact already-proven Python
      values (`tests/test_discovery.py`'s shp-groupshape.pptx 4-shape/zero-field/
      Group-4-nesting case and mst-slide-layouts.pptx's title/body placeholder type
      case) — not executed here (no Windows/Office in this environment, same
      constraint `InjectPrimitive.bas` already documented). No `src`/`tests` changes;
      confirmed `python3 -m pytest tests/ -v` (70 passed) and `python3 -m mypy src/`
      (no issues, 10 source files) both still pass (unaffected, as expected for a
      vba/-only change).
      **Still open for a future pass**: port-order steps 3 (`matching`), 4
      (`resolve`+`sync_operations`), 5 (`onboarding`), 6 (Excel-side) remain unported.
      Step 3 in particular inherits this task's flagged `PlaceholderIdx` gap and will
      need to resolve it (native fallback to raw OOXML, per `specs/vba-port.md`'s
      explicit-fallback allowance) before it can port `matching.md`'s
      placeholder-index signal faithfully.

## Priority 13 (2026-07-24 pass): VBA port order, module 3 (matching)

- [x] Port `matching` per `specs/vba-port.md`'s numbered port order (module 3 of 6,
      after discovery and identity_tags; module 2/identity_tags was already done
      out-of-order as the original spike). Confirmed nothing else in the order had
      started beyond `Discovery.bas` (`ls vba/` before this task showed only
      `InjectPrimitive.bas`, `SPIKE_NOTES.md`, `Discovery.bas`,
      `SPIKE_NOTES_Discovery.md`). Added `vba/Matching.bas`: `ScoreCandidate()`/
      `Match()` port `matching.py`'s tier-1 tag-trust + tier-2 weighted scoring
      (placeholder-index/geometry/shape-type/content at the same 0.5/0.3/0.15/0.05
      weights and 0.75/0.4 thresholds) and sibling-ambiguity z-order tie-break,
      field-for-field.
      Also resolved `SPIKE_NOTES_Discovery.md`'s flagged `PlaceholderIdx` gap (always
      `-1` in `Discovery.bas` since the object model has no
      `Shape.PlaceholderFormat.Idx` equivalent) via the raw-OOXML fallback
      `specs/vba-port.md` explicitly allows when no native path exists:
      `EnrichPlaceholderIdx`/`LoadPartXml`/`PlaceholderIdxFromDom` extract a part's XML
      via `Shell.Application`'s native zip-folder support (no zip library in stock
      VBA) and read `<p:cNvPr name>`'s sibling `<p:nvPr>/<p:ph idx>`, applying OOXML's
      own idx-omitted-defaults-to-0 rule, matching `discovery.py`'s
      `_placeholder_info`. Added a safety guard `score_candidate.py` doesn't need:
      `PlaceholderScore` treats an unresolved `-1` sentinel as "signal not applicable"
      rather than risking two unresolved placeholders falsely comparing `-1=-1` as a
      match. Added `vba/SPIKE_NOTES_Matching.md` (same format as prior modules)
      documenting this and 5 other real divergences (geometry/shape-type/content
      always-applicable in VBA vs. Python's optionality; fallback reads last-saved
      disk state, not live unsaved edits; duplicate-shape-name lookup ambiguity
      inherited from `InjectPrimitive.bas`'s already-documented risk; `CopyHere`'s
      async-completion gotcha; no `Candidate | None` equivalent so `MatchResult` uses
      an array-index convention instead), plus a manual verification recipe
      cross-checked against `tests/test_matching.py`'s already-proven values
      (`shp-groupshape.pptx`'s 4-way z-order tie-break landing on "Oval 2";
      `mst-slide-layouts.pptx`'s body placeholder idx=10 match alone correctly
      landing at medium, not auto-accepted). No `src`/`tests` changes; confirmed
      `python3 -m pytest tests/ -v` (70 passed) and `python3 -m mypy src/` (no issues,
      10 source files) both still pass (unaffected, as expected for a vba/-only
      change).
      **Still open for a future pass**: port-order steps 4 (`resolve`+
      `sync_operations`), 5 (`onboarding`), 6 (Excel-side) remain unported. Step 4 can
      now assume `EnrichPlaceholderIdx` is available when it needs placeholder-index
      scoring, rather than inheriting a fresh gap.

## Priority 14 (2026-07-24 pass): VBA port order, module 4 (resolve + sync_operations)

- [x] Port `resolve` + `sync_operations` per `specs/vba-port.md`'s numbered port order
      (module 4 of 6, after discovery/identity_tags/matching). Confirmed nothing else in
      the order had started beyond `Discovery.bas`/`Matching.bas` (`ls vba/` before this
      task showed only `InjectPrimitive.bas`, `SPIKE_NOTES.md`, `Discovery.bas`,
      `SPIKE_NOTES_Discovery.md`, `Matching.bas`, `SPIKE_NOTES_Matching.md`).
      Found and closed a real, previously-mis-stated gap before porting further:
      `specs/vba-port.md`'s port order describes module 2 (`identity_tags`) as "already
      done (`InjectPrimitive.bas`'s `Shape.Tags`/`Slide.Tags` reads)" -- grepping
      `InjectPrimitive.bas` confirmed it only ever reads `Shape.Tags("role")`; it never
      once reads `Slide.Tags`. Module 4 is the first module that actually needs
      slide-level tags (`slide_type`/`instance_key`), so `vba/Resolve.bas`'s
      `ResolveSlideInstance(sld)` adds the missing native `Slide.Tags` reads itself
      (scoped to reads only -- writing slide-level tags is onboarding's job, port-order
      step 5, still open) rather than carrying the mis-statement forward.
      Added `vba/Resolve.bas` (`ResolveSlideInstance`) and `vba/SyncOperations.bas`
      (`PlanRoutineSync` for cases 1/3/4/6, `PlanPeriodRollover` for case 2, never
      dispatched from routine sync). Deliberately does **not** port `resolve.py`'s
      `field_shapes` pre-resolution step: Python needs it because file-surgery
      `inject_primitive` requires a `Candidate`'s `z_order`, but VBA's `InjectPrimitive`
      already does its own tag-based shape lookup on every call, so
      `PlanRoutineSync` calls `InjectPrimitive.InjectPrimitive(slide, fieldName, value)`
      directly per field -- a genuine simplification, not a corner cut, same category as
      `Shape.Tags` collapsing `identity_tags.py`'s hand-rolled XML in the original spike.
      Since port-order step 6 (Excel-side reads) isn't built yet, `PlanRoutineSync`
      accepts an already-read Data-sheet shape (`Scripting.Dictionary` of
      `Scripting.Dictionary`s plus an order `Collection`) as a documented interface
      contract for that future module to satisfy, rather than reading a worksheet
      itself -- mirrors `plan_routine_sync()`'s own separation from `excel_output.py`.
      Added `vba/SPIKE_NOTES_Resolve.md` (covering both files, since `specs/vba-port.md`
      groups them as one port-order item) documenting the closed `Slide.Tags` gap, 7
      deliberate divergences (including one arguably *safer* than the Python original:
      `InjectPrimitive`'s ambiguous-tag refusal vs. `resolve.py`'s silent
      last-write-wins on a duplicate role tag), and a manual verification recipe
      cross-checked against `tests/test_resolve.py`'s and `tests/test_sync_operations.py`'s
      already-proven values for all of cases 1/2/3/4/6. No `src`/`tests` changes;
      confirmed `python3 -m pytest tests/ -v` (70 passed) and `python3 -m mypy src/` (no
      issues, 10 source files) both still pass (unaffected, as expected for a vba/-only
      change).
      **Still open for a future pass**: port-order steps 5 (`onboarding`) and 6
      (Excel-side) remain unported. Step 5 will need to close the write-side of the
      `Slide.Tags` gap flagged above (writing `slide_type`/`instance_key`, not just
      reading them). Step 6's Data-sheet Dictionary/Collection shape is now a fixed
      interface contract this module already depends on, not an open design question.
      Also unconfirmed and worth resolving whenever a real Office install is available:
      whether `CustomLayout` objects expose a native `.Tags` property the same way
      `Slide` does (needed to resolve `mst-slide-layouts.pptx`-style fixtures directly,
      per `SPIKE_NOTES_Resolve.md`'s divergence 6).

## Priority 15 (2026-07-25 pass): VBA port order, module 5 (onboarding)

- [x] Port `onboarding` per `specs/vba-port.md`'s numbered port order (module 5 of 6,
      after discovery/identity_tags/matching/resolve+sync_operations). Confirmed via `ls
      vba/` that modules 1-4 were the only ones present before this task.
      Found and closed a real gap before porting: neither `Discovery.DiscoverSlide` nor
      `DiscoverCustomLayout` ever exposed the live `Shape` object behind a returned
      `Candidate` -- no prior caller needed the Candidate-to-shape direction
      (`InjectPrimitive.bas` goes tag-to-shape, never candidate-to-shape). Onboarding is
      the first caller needing it (to write a tag onto a matched candidate, or read a
      template's existing role tags), so `Discovery.bas` gained
      `DiscoverSlideWithShapes`/`DiscoverCustomLayoutWithShapes` (parallel `shapes()`
      array, same index as the returned `Candidate()`, since `Candidate.ZOrder` is exactly
      the deterministic position `Walk` already assigns) -- purely additive,
      `DiscoverSlide`/`DiscoverCustomLayout`'s own signatures and output are unchanged.
      Also found a real VBA restriction while designing this: a UDT cannot be assigned to
      a `Variant`, so a `Scripting.Dictionary` (the natural port of Python's
      `dict[str, Candidate]`) cannot hold `Candidate`/`MatchResult`/`InjectResult` values
      -- confirmed this is why `Onboarding.bas` uses parallel arrays
      (`roles() As String` / `Candidate()`) instead of a dictionary, and flagged (not
      fixed) that `SyncOperations.bas`'s already-shipped `PlanRoutineSync`
      (`changed(fieldName) = r` where `r As InjectResult`) appears to hit this exact
      restriction -- a different, already-committed module, left for a separate decision
      rather than silently patched here.
      Added `vba/Onboarding.bas`: `BuildTemplateFieldShapes` (role -> reference `Candidate`,
      the field-shapes half `Resolve.SlideInstance` deliberately doesn't carry),
      `MatchSlideAgainstTemplate` (per-role tier-2 scoring via `Matching.Match`, excluding
      already-tagged and pure-decoration candidates), `ConfirmFieldMatch` (the write
      primitive a selection UI would call), and `OnboardNewInstance` (unconditional
      slide-identity tagging + auto-accept of high-confidence matches only), field-for-field
      against `src/onboarding.py`. `Shape.Tags.Add`/`Slide.Tags.Add` upsert natively, so no
      read-merge-write logic was needed the way `identity_tags.py`'s hand-rolled XML
      required.
      Added `vba/SPIKE_NOTES_Onboarding.md` documenting both real findings above, 5
      deliberate divergences, and a manual verification recipe cross-checked against
      `tests/test_onboarding.py`'s already-proven values (high/medium confidence split on
      `mst-slide-layouts.pptx`'s two layouts, pure-decoration exclusion on
      `shp-groupshape.pptx`). No `src`/`tests` changes -- `python3 -m pytest tests/ -v` /
      `python3 -m mypy src/` were not re-run this pass (no pytest/mypy on this bare host;
      those only ever ran inside the Ralph Docker container per `AGENTS.md`, which wasn't
      available this pass either), but since nothing under `src/`/`tests/` was touched,
      there is nothing on the Python side that could have regressed.
      **Still open for a future pass**: port-order step 6 (Excel-side reads/writes) is the
      only module left. Also still open: whether `CustomLayout` exposes a native `.Tags`
      property (flagged since `SPIKE_NOTES_Resolve.md`, still unconfirmed -- no Office
      install here), and the `SyncOperations.bas` UDT/Dictionary finding above, which needs
      its own fix decision.

## Priority 16 (2026-07-25 pass): VBA port order, module 6 (excel_output) -- port order complete

- [x] Port `excel_output` per `specs/vba-port.md`'s numbered port order (module 6 of 6, the
      last one). Confirmed via `ls vba/` that modules 1-5 were the only ones present before
      this task.
      Unlike every prior module, `vba-port.md` itself flags this one as strictly simpler
      than its Python source, not just a mechanism swap: `excel_output.py` hand-rebuilds
      the entire `.xlsx` zip package (six OOXML parts) on every single write, purely
      because a headless script has no host application to lean on. A live `Worksheet`
      doesn't need that -- `ExcelOutput.UpsertRow` finds-or-appends the target column/row
      directly and writes only those specific cells, per `vba-port.md`'s explicit
      instruction not to port the zip-rebuilding approach.
      Added `vba/ExcelOutput.bas`: `CreateSheet` (writes the `"Instance ID"` header, stores
      a deck reference via native `Workbook.CustomDocumentProperties` rather than a
      hand-rolled custom-properties XML part, refusing to re-initialize an already-set-up
      sheet), `ReadSheet` (recovers fields/instance rows/deck reference into a `Sheet` UDT,
      using `IsEmpty()` rather than `= ""` to distinguish "never harvested" from "harvested
      as an empty string," matching `read_sheet`'s structural cell-presence check), and
      `UpsertRow` (appends a new field column or instance row as needed; only ever writes
      the specific cells a call's `values` actually mentions, never touching an unrelated
      field/instance).
      Confirmed the port-order-4 interface contract is satisfied with **no adapter
      needed**: `SPIKE_NOTES_Resolve.md`'s divergence 4 documented `PlanRoutineSync`'s
      `dataRows`/`instanceOrder` parameters as "an interface contract for that future
      module to satisfy" -- `ReadSheet(ws).Rows`/`.InstanceOrder` are exactly that shape
      (`Scripting.Dictionary` of `Scripting.Dictionary`s / `Collection` of instance-ID
      strings in row order) and plug straight in, confirmed in
      `SPIKE_NOTES_ExcelOutput.md`'s own manual-recipe step 5.
      Added `vba/SPIKE_NOTES_ExcelOutput.md`: 6 deliberate divergences and a manual
      verification recipe cross-checked against `tests/test_excel_output.py`'s
      already-proven round-trip values (no `.xlsx` fixture exists for this spec on either
      language's side -- both are necessarily round-trip/self-consistency tests, not
      checks against an externally-produced file). No `src`/`tests` changes; `python3 -m
      pytest tests/`/`mypy src/` were not re-run this pass (no pytest/mypy on this bare
      host, same as the prior pass), but nothing under `src/`/`tests/` was touched, so
      there is nothing on the Python side that could have regressed.
      **All 6 port-order modules now exist** (`Discovery.bas`, `InjectPrimitive.bas`,
      `Matching.bas`, `Resolve.bas`+`SyncOperations.bas`, `Onboarding.bas`,
      `ExcelOutput.bas`), each with its own `SPIKE_NOTES_*.md`. This is *not* the same as
      "the VBA port is done": **nothing has ever been executed against a real Office
      install** -- every module's correctness rests entirely on field-for-field comparison
      against the Python side's 70 passing tests plus manual reasoning, not a single real
      run. Also still genuinely open, independent of Office access: no orchestration/driver
      exists tying the 6 modules together into an actual usable sync flow (each module's
      own spec explicitly draws that boundary as someone else's job); `specs/slide-
      duplication-trigger.md`'s row-order/resequence decisions aren't ported to VBA at all
      yet (that spec was written after this port-order work started and targets a still-
      unbuilt duplication primitive); and the parked multi-deck design (export,
      conflict-resolution/propagation, stale-queue -- see `claude-brain`'s
      `project_active_ventures.md`, 2026-07-25 entries) remains entirely unspecified.

## Priority 17 (2026-07-25 pass): first real-Office execution + hardening loop

- [x] Every VBA module's `SPIKE_NOTES_*.md` said "not executed or verified in this
      environment" since there was no Windows/Office install available. That changed this
      pass: `powershell.exe` is reachable from this WSL environment and can drive a real
      Windows-side PowerPoint 16.0 + Excel 16.0 install directly via COM automation
      (confirmed, not assumed). Built `vba/tests/TestRunner.bas` (PowerPoint: Discovery,
      InjectPrimitive, Matching, Resolve, SyncOperations, Onboarding -- 13 assertion-based
      test functions) and `vba/tests/TestRunnerExcel.bas` (Excel: ExcelOutput -- 8 test
      functions), deliberately NOT reusing the existing `ManualSmokeTest*` subs (those use
      `MsgBox`, which hangs a headless run waiting on a click that never comes). Built
      `vba/tests/run_vba_tests.ps1`: stages `.bas` files + `test-fixtures/shp-groupshape.pptx`
      into a Windows-native temp dir (COM automation across the `\\wsl.localhost\...` UNC
      boundary is flaky; plain file copies across it aren't), imports everything, runs both
      test suites, prints one machine-readable report per host app.
      Real environment blockers hit and resolved, not hypothetical: (1) "Trust access to the
      VBA project object model" was actually off (a first probe looked like it was already
      enabled but was a false positive from PowerShell silently returning null on a null COM
      *property* access rather than throwing -- only method *calls* on null reliably throw;
      confirmed via registry, `AccessVBOM` unset under both `HKCU:\...\PowerPoint\Security`
      and `\Excel\Security`) -- enabled with Rohan's explicit sign-off (real security
      setting, not flipped silently, see claude-brain's DECISIONS.md 2026-07-25 entry). (2) A zombie
      POWERPNT.EXE from an interrupted probe made `New-Object -ComObject` attach to the
      stale instance instead of spawning fresh, breaking `VBProject` access in a way that
      looked identical to the trust-setting problem -- script now always kills stray
      processes first. (3) `Application.Run` with inline arguments fails PowerShell's own
      COM method-overload resolution; fixed via `InvokeMember` reflection with explicit
      `[string]` casts on every argument.
      **The real finding, not incidental to the tooling**: `ReDim arr(1 To 0)` throws
      "Subscript out of range" at runtime in real VBA -- confirmed via multiple clean
      isolated repros, contradicting the "(1 To 0) = empty array" convention every module
      in this port used since Priority 12. Root-caused and fixed across `Discovery.bas`,
      `Matching.bas`, `SyncOperations.bas`, `Onboarding.bas`, and `TestRunner.bas`: never
      pre-ReDim to an empty range; leave arrays genuinely unallocated until the first
      `ReDim Preserve` (confirmed safe as a first allocation, including for UDT arrays);
      callers must error-guard `LBound`/`UBound` rather than assume a returned array is
      allocated. Documented as a new AGENTS.md Known Pattern.
      **Result after fixing**: 12 of 13 PowerPoint tests pass for real, all 8 Excel tests
      pass for real. The one remaining failure is itself a genuine, confirmed finding, not
      a bug in this pass's fix: `Matching.EnrichPlaceholderIdx`'s `Shell.Application.
      Namespace(zipPath)` reliably returns `Nothing` when driven via COM automation, even
      against an independently-verified-valid `.pptx` (opened cleanly by .NET's
      `ZipFile.OpenRead`, 39 real entries) -- reproduced 3 times including with an explicit
      delay (rules out a shell-notification race). Not fixed this pass -- needs a real
      redesign of the zip-extraction technique, not a patch; documented in
      `SPIKE_NOTES_Matching.md` and left open.
      No `src`/`tests` changes. `python3 -m pytest`/`mypy` not re-run (bare host has
      neither; unaffected regardless since nothing under `src/`/`tests/` changed).

## Priority 18 (2026-07-25 pass): fix EnrichPlaceholderIdx's zip-extraction, all 21 real tests now pass

- [x] Fix the `Matching.EnrichPlaceholderIdx` zip-extraction failure Priority 17 found and left
      open. Replaced `LoadPartXml`'s `Shell.Application.Namespace()`/`CopyHere` technique
      entirely with shelling out to `tar.exe` (bsdtar, bundled with Windows 10 1803+/Windows 11
      by default) via `WScript.Shell.Run("cmd.exe /c tar -xf ... -C ... <entry>", 0, True)` --
      confirmed working standalone against a real saved `.pptx` (correct real XML content
      extracted) before being wired in, avoiding the Windows shell namespace layer entirely
      (the layer that was actually failing, per Priority 17's finding).
      That fix immediately surfaced a **second, previously-masked bug**: `CreateObject
      ("MSXML2.DOMDocument60")` (no dots) raises Err 429 "ActiveX component can't create
      object" on this machine -- execution had never actually reached that line before,
      since the `Shell.Application` failure always short-circuited first. Root-caused by
      probing every plausible MSXML ProgID variant directly: `"MSXML2.DOMDocument.6.0"`
      (dotted) works and returns the identical real `DOMDocument60` object (confirmed via
      `TypeName`). Fixed by switching the ProgID string only -- the object type/API used
      afterward (`.async`, `SelectionLanguage`/XPath, `.Load`) is unchanged.
      Diagnosed via a temporary debug-log instrumentation added directly to `LoadPartXml`
      (write cmd/exitCode/file-existence/dom-load-result to a log file on every call),
      re-run through the real `run_vba_tests.ps1` pipeline rather than isolated probes --
      isolated ad hoc probes of the exact same file/part succeeded even while the real test
      still failed, which is what actually pointed at "something after the point my probes
      covered" rather than the extraction step itself. Instrumentation removed before commit.
      **Result: all 21 real tests now pass** (13/13 PowerPoint, 8/8 Excel) -- the full suite
      is clean for the first time. Updated `SPIKE_NOTES_Matching.md`'s finding from "confirmed
      broken, not fixed this pass" to "found and fixed," with both root causes and the fix
      technique documented.
      No `src`/`tests` changes. `python3 -m pytest`/`mypy` not re-run (bare host has neither;
      unaffected regardless since nothing under `src/`/`tests/` changed).
      **New assumption now load-bearing, worth flagging**: the raw-OOXML fallback path
      depends on `tar.exe` being present (Windows 10 1803+/Windows 11 default -- true of the
      real target machine, confirmed, but not universally true of every conceivable Windows
      install `EnrichPlaceholderIdx` might someday run on). Not treated as a blocker; noted
      for awareness if this ever needs to run somewhere older.

## Priority 19 (2026-07-25 pass): the orchestration layer -- deck-sync-refimpl is now a real, runnable sync tool

- [x] Build the piece every module in this port explicitly deferred: gathering live
      instances, executing a decided duplication, and reconciling deck order every sync.
      Three new modules, all written and executed against real Office the same day:
      **`Verification.bas`** (new port, not in the original 6-module order --
      `verify_structure`/`verify_z_order` from `src/verification.py`, needed because
      `specs/slide-duplication-trigger.md` makes them mandatory before tagging any
      duplicate; had to do its own `DiscoverSlideWithShapes` + role-tag read internally
      since `Candidate.IdentityTag` is always blank, same reason `Onboarding.
      BuildTemplateFieldShapes` already does this). **`SlideDuplication.bas`**
      (`DuplicateAndTag`, the first real implementation of `specs/slide-duplication-
      trigger.md`: instance-key collision guard, `Slide.Duplicate`, mandatory
      structural/z-order verification before any tag write -- a failed check deletes the
      malformed duplicate rather than leaving it as untagged debris, a judgment call the
      spec itself doesn't dictate -- then tags and injects with partial-row handling).
      **`RunSync.bas`** (`GatherInstances`, `RunRoutineSync` dispatching `SyncOperations.
      PlanRoutineSync`'s decisions to `SlideDuplication.DuplicateAndTag`, and
      `ResequenceByRowOrder` applying the row-order standing invariant to the whole type
      every pass -- anchored at the type's current lowest `SlideIndex` rather than the
      front/back of the deck, another judgment call the spec leaves open).
      Extended `vba/tests/TestRunner.bas` with 8 new real tests (3 Verification, 3
      SlideDuplication, 2 RunSync -- including a full end-to-end pass: a template, a live
      cross-app Excel worksheet with one stale row and two brand-new rows, confirming
      case 4 correction, case 3 creation ×2, correct injection, and post-sync
      resequencing all in one real run).
      **Two real bugs found and fixed via the hardening loop, not code review**: (1) the
      driver script's own PowerPoint-side import list never included `ExcelOutput.bas`
      (a tooling gap in `run_vba_tests.ps1`, not production code) -- `RunSync.bas`
      declares `Dim sheet As Sheet` and calls `ExcelOutput.ReadSheet`, so without that
      import the whole project failed to compile ("Sub or function not defined"). (2) A
      genuine production bug found the same way: `ExcelOutput.bas`'s `xlToLeft`/`xlUp`
      named constants (Excel's `XlDirection` enum) only resolve inside Excel's own VBA
      project -- every prior `ExcelOutput.bas` test had run there and looked completely
      clean, but `RunSync.bas` (PowerPoint-hosted) was the first code in this whole
      project to actually call `ExcelOutput.ReadSheet` cross-app, and hit "Variable not
      defined" immediately. Fixed by switching both to numeric literals (`XL_TO_LEFT =
      -4159`, `XL_UP = -4162`).
      Both bugs manifested as an indefinite hang rather than a clean COM exception
      (compile errors aren't catchable by `On Error`, consistent with this project's
      earlier modal-dialog finding) -- diagnosed by literally screenshotting the
      PowerPoint/VBE window mid-hang (bringing it to the foreground first via
      `SetForegroundWindow`, since a hung Office process sits behind whatever terminal
      triggered it) rather than continuing to guess blindly. Both new techniques
      documented in `AGENTS.md`'s Known Patterns for future passes.
      **Result: all 29 real tests now pass** (21/21 PowerPoint, 8/8 Excel) -- the fullest
      real end-to-end proof this project has produced: a fresh Excel sheet genuinely
      drives real slide creation, correction, and reordering in a real PowerPoint deck.
      No `src`/`tests` changes (this priority has no Python equivalent at all --
      orchestration was explicitly out of scope for every Python module by design).
      **Still open**: case 2 (period rollover) is not driven from `RunRoutineSync` (by
      design, per `sync-operations.md`'s own "never inferred from routine sync" rule) --
      no orchestration exists yet for the explicit rollover command itself. Cases 5/7
      remain non-goals throughout. The parked multi-deck design (export,
      conflict-resolution/propagation, stale-queue -- see `claude-brain`'s
      `project_active_ventures.md`) remains entirely unspecified, as does deciding where
      a type's template is actually stored/looked up (currently caller-supplied to
      `RunRoutineSync`, per `onboarding.md`'s own non-goal).

## Notes for next planning pass
- No `pyproject.toml`/`mypy.ini`/`setup.cfg` exists — mypy is running with default
  settings. Worth confirming this stays intentional as more modules are added. Still true
  as of this pass (confirmed: no such files at repo root).
- `matching.py`'s content-pattern signal degrades to a has-text boolean match because
  `Candidate` doesn't store actual text content, only `has_text`. If a fixture surfaces
  a real need for pattern-based matching (e.g. distinguishing a date field from a
  status field by format), `Candidate` will need a text-content field added.
- `matching.py`'s geometry reading (`_geometry()` in discovery.py) uses each shape's
  raw local `a:off`/`a:ext` without walking up through parent group transforms — exact
  only when a group's `chOff`/`chExt` equals its own `off`/`ext` (true for every
  current fixture, not guaranteed in general OOXML).
- `vba/InjectPrimitive.bas` is unexecuted and unverified (no Windows/Office in this
  environment, per `SPIKE_NOTES.md`) — if a Windows/Office environment ever becomes
  available, running the manual verification recipe in `SPIKE_NOTES.md` is real
  outstanding work, just not automatable from here.

## Priority 11 (2026-07-24 pass): AGENTS.md's VBA-port cross-reference went stale mid-iteration

- [x] Context: commit `67f4df8` had a human revert the loop's own prior "no further VBA
      porting is wanted" call (`ba05df2`), correcting it to "full VBA production porting is
      planned — write `specs/vba-port.md` before that work starts." That gap closed during
      this same pass, but via a separate concurrent commit (`06767da`, "Add
      specs/vba-port.md: scope the production VBA port") that landed while this iteration
      was independently drafting the identical file — confirmed by comparing timestamps and
      `git log`; not this iteration's own work, and not re-done here to avoid clobbering it.
      Reviewed `06767da`'s `specs/vba-port.md` directly: it satisfies AGENTS.md's ask
      (port order mirroring the real dependency chain, each existing `specs/*.md` staying
      the governing spec for its own VBA translation, "translate the mechanism not the
      Python workaround" called out per module via `vba/SPIKE_NOTES.md`'s already-demonstrated
      pattern, and a per-module manual-verification-recipe requirement since this
      environment has no Windows/Office install). What was still stale after that commit:
      `AGENTS.md`'s own Constraints bullet still said "write `specs/vba-port.md` before that
      work starts" as an open ask, now factually wrong since the spec exists. Fixed by
      updating that bullet to point at the spec instead of asking for it. Pure doc fix, no
      `src`/`tests`/`specs` change. Confirmed `python3 -m pytest tests/ -v` (70 passed) and
      `python3 -m mypy src/` (no issues, 10 source files) both still pass (unaffected, as
      expected for a docs-only change).

## Priority 10 (2026-07-24 pass): pure decoration could be auto-tagged as a field

- [x] Fix `src/onboarding.py`'s `match_slide_against_template()`: it built its untagged
      candidate pool as "discovered minus already-tagged" only, never excluding pure
      decoration (no text, not a picture). specs/discovery.md is explicit: "A shape with
      no text content and that is not a picture is not a candidate... Pure decoration
      must be correctly excluded, not force-matched." `Candidate.is_candidate_field`
      already computes this (discovery.py:42-44) but was, until this fix, checked in
      exactly one place in the whole repo (`tests/test_discovery.py`, a post-hoc
      assertion) — never as an actual filter anywhere real matching happened. Confirmed
      concretely, not hypothetically: `tests/test_matching.py`'s existing
      `test_shp_groupshape_sibling_ambiguity_resolved_by_zorder` already showed `match()`
      itself will happily return a HIGH-confidence match onto `shp-groupshape.pptx`'s
      "Oval 2" — a shape `test_discovery.py`'s own
      `test_shp_groupshape_finds_zero_candidate_fields` proves is pure decoration. Since
      `onboard_new_instance()` auto-tags any HIGH-confidence match immediately with no
      human in the loop, this was a real path to silently linking a business field to an
      empty decoration shape.
      `discover()` itself was deliberately left unchanged — it must keep returning every
      shape (decoration included), since other callers (e.g. `verification.
      verify_structure`) need the full shape list for structural correspondence, not just
      fields. The filter belongs at the point candidates are handed to the matcher, so it
      was added to `match_slide_against_template()`'s `untagged` list comprehension
      (`c.is_candidate_field and ...`), not to `discover()` or to `matching.match()`
      itself (matching.md's spec is silent on decoration — that's discovery.md's
      requirement, and `match()`'s own tests deliberately use decoration shapes as
      synthetic tie-breaking fixtures, unrelated to this gap). Added two regression tests
      to `tests/test_onboarding.py` using `shp-groupshape.pptx`'s real decoration shapes
      and the same reference `test_matching.py` uses to previously force a HIGH match:
      confirms `match_slide_against_template` now reports LOW/unmatched instead, and that
      `onboard_new_instance` never auto-tags them. Confirmed `python3 -m pytest tests/ -v`
      (70 passed, was 68) and `python3 -m mypy src/` (no issues, 10 source files) both
      pass.
