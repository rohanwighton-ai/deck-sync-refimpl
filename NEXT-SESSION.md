# NEXT SESSION — start here

> ## 17 AUG, LATE NIGHT (~22:10) — the Device Registry (P3) turned out to be
> one small fix, not a new module — three of the four suspected consumers
> were already fine. Full suite 243/243. NOT yet built into an addin.
> **STATUS: CURRENT, supersedes the block below.**
>
> Re-checked the "A DEVICE REGISTRY" design (below in this same file) against
> the actual current code before writing anything, per this project's own
> "read the file, don't produce a theory" rule — the picture had moved since
> that doc was written. **Discovery is already fixed** (`SlotCount`-based
> recognition). **Marking is already handled reasonably** (asks before
> opening a device up). **Template Audit is already clean, confirmed by
> tracing**: it shares Discovery's candidate list, and separately its
> `ShapeText` helper silently returns `""` for a group shape (no
> `TextFrame`), so a device candidate never becomes an audit row either way.
> **Only `FieldWiring` still had zero device awareness** — full detail in
> `FIX-LIST.md`, item P3.
>
> Fix: `MilestoneDevice.IsColumnForThisDevice`, plus a new `DeviceOwnedCount`
> on `FieldWiringResult` that `ScanFieldWiring`'s per-field loop routes
> device columns into instead of the ordinary carrier/unmarked machinery.
> Matches the pattern the codebase already chose twice (direct
> `MilestoneDevice` calls, no registry indirection) rather than building a
> plugin table for a population of one device.
>
> New test made to fail first, harder than usual: with the fix stashed, it
> doesn't even COMPILE (references a result field that doesn't exist without
> the fix). Full suite 243/243 with it restored. **Not yet built into an
> addin or measured live** — next session, build the next addin (whatever
> follows `addin133`) and confirm the run report/START HERE sheet no longer
> lists the 21 `MS*` columns.

> ## 17 AUG, LATE NIGHT (~21:45) — item AT fixed in full (3 rounds), `addin133`
> deployed and confirmed live. P1 also confirmed built and deployed
> (`addin131`/`132`/`133` all carry it) — the earlier "not yet built" note below
> is stale. `purpose-hound` agent commissioned and run once.
> **STATUS: CURRENT, supersedes every block below.**
>
> **What happened, in order:** Rohan asked to check item X's real freeze
> ("nothing anywhere, looks frozen" — a genuinely different symptom from P1's
> "dialog behind the window," confirmed via a direct question rather than
> assumed to be the same bug). A fable-run `waste-hound` traced it to
> `ReviewChangesCore` -> `BuildQueue` -> `PlanRoutineSync` -> the injector
> family, landing on `ShapeAddressBook.Lookup` — the cache built earlier that
> same night to fix item AR turned out to contain items W and AB's own defect
> shape, twice over, plus a borrowed justification that didn't transfer. Full
> detail, all three rounds, and the measurement-methodology trap found and
> corrected mid-investigation (cross-project `Application.Run` adds real
> overhead beyond trivial dispatch — confirmed via an isolated floor probe):
> **FIX-LIST.md, item AT.**
>
> **Numbers: 508.5ms/call (unfixed, cross-project baseline) -> 346.3ms
> in-process (true apples-to-apples baseline) -> 222.1ms after rounds 1-2
> (resolve-once caching) -> proven FLAT regardless of miss volume after round
> 3** (in-memory negative cache — a direct probe recorded 300 distinct misses
> then re-read the sheet's row extent fresh from the file: 17 rows before, 17
> after). At the project's own estimated ~5,000+ Lookup calls per "Put it on
> the slides" press, this was ~42 minutes worst-case; it is now structurally
> bounded by real hit count alone, not by how many absent fields exist.
>
> **Round 3 came from Rohan's own question, not code review**: "but why is it
> caching misses?" — the negative cache had copied the positive cache's
> "worth surviving a reopen" justification without re-deriving whether it
> actually transferred to the majority (miss) case. It didn't. This is now
> the canonical example `purpose-hound` (see below) is calibrated against.
>
> **A real hang, found and cleared, unrelated to the fix:** the full 242-test
> suite hung 5+ minutes once during verification, leaving 4 orphaned Excel
> processes. Bisected by calling each of the 3 `ShapeAddressBook` tests
> individually via the real `TestRunner.RunAllTests` entry point (not TestRunner
> as a black box) — all three pass clean in isolation. A clean re-run afterward
> completed normally, 242/242 twice. Root cause judged environmental
> (accumulated Office automation state from a very long session), not a defect
> in the fix — worth re-checking if it recurs, not chased further tonight.
>
> **Deploy hygiene finding, worth remembering:** `File > Save As` twice this
> session defaulted somewhere OTHER than the AddIns folder without warning —
> once to `OneDrive` root, once to `OneDrive\Claude\`. Neither failed loudly;
> both looked like a normal successful save. **Always verify the actual saved
> path after a Save-As**, don't assume the dialog remembered the AddIns
> folder. Separately: `addin131` and `addin132` were BOTH still set
> `AutoLoad=True` when `addin133` was registered — three versions of the same
> add-in auto-loading at once, a real duplicate-module-collision risk that
> nothing would have surfaced until something broke mysteriously. Now cleaned
> up (only `addin133` auto-loads); **check for this after every deploy from
> now on**, not just after ones where something looked wrong.
>
> **New agent commissioned: `purpose-hound`** (`~/claude-brain/agents/
> purpose-hound.md`, symlinked into `~/.claude/agents`, `model: sonnet`).
> Cold, read-only auditor for borrowed/copied reasoning — for each function
> whose comment argues for its own design ("same reason as X," "worth
> persisting," etc.), asks four questions: what's the role, what greater good
> is claimed, does that greater good actually apply HERE (checked against
> BOTH the sibling function it was borrowed from AND the project's own
> strategic intent — personal tool, not org-wide; "a boundary earns its place
> only where a person decides"), and if it applies, has it actually been
> tested per this repo's own bar (a named test or a "proven live" note), not
> just reasoned into confidence. Calibrated against the AT/RecordAbsent defect
> above. **First real run, same night, on this repo:** mostly clean — the
> codebase's habit of narrating and self-correcting its own past transplant
> errors in comments is real and held up. One live finding: `DeckRegistry.
> SetWorkbookPathVerified` borrowed its retry/escalation design from
> `SetDeckPeriodVerified` (whose fix has a named "8/8 proven live" note) but
> has never had its OWN live proof — plausibly inherits the same underlying
> fix, never independently confirmed. Cheap to close (a live retest, and
> Scenario 8 exercises this path anyway). **DONE, same night**: 2 rounds against
> a real 49MB test deck (a scratch copy, never the real one), each a repoint +
> close + reopen + read-back from the SAVED FILE'S BYTES via `WorkbookPathOnDisk`
> — both matched. `SetWorkbookPathVerified` now has its own named live proof,
> same bar as its siblings. (Test's own first attempt used non-existent fake
> paths and was correctly REFUSED by the function's own existence guard — not a
> bug, a good find: re-ran against real dummy `.xlsx` targets.)
>
> **Not yet done:** the actual live retest against item X's real stall this
> was diagnosed from — press "1. Set up my quarter" then "2. Put it on the
> slides" for real, read the Timing sheet, confirm the freeze is actually
> gone (or much shorter) on the real deck, not just proven in isolation.
> **This is next session's first action.** Everything else below this block
> (Scenario 1's unaided close, the P1 dialog-visibility fix, the earlier
> OneDrive write-reliability work) is unchanged and still the larger context —
> read down from here for it, but the immediate next step is the item X
> retest.

> ## 17 AUG, LATE — P1's code fix written (background session), NOT yet
> built/deployed/live-tested.
> **STATUS: CURRENT, supersedes the block below. The plan in that block was
> followed exactly; what's left is steps 5 onward — build the next addin,
> deploy, and retry Scenario 1 for real.**
>
> `DraftingUI.BringPowerPointToFront` written (same marker-caption/
> `AppActivate` technique as `BringExcelToFront`, right beside it,
> `DraftingUI.bas`), wired into `RibbonUI.SyncNowChainCore` at both points
> the plan below specified: the top of the sub (before `Set pres =
> Application.ActivePresentation`) and immediately after `RollForwardUI`
> returns (before `RefreshDraftingSheets`). `check_vba_static.py` clean
> across 38 modules. **The Python `pytest` suite could NOT be run** — this
> was done in a Linux sandbox with no `pip`/network and no Office/COM
> access, so neither the Python suite nor any live VBA behavior could be
> exercised; the change touches VBA only, no Python files. Committed on
> branch `worktree-p1-bring-ppt-front`, not merged to `main`.
>
> **Next session's first action:** review this diff, merge to main if it
> looks right, build the next addin (whatever `addin13x` follows `addin130`),
> deploy, confirm loaded live, THEN retry the real Scenario 1 attempt from
> scratch per the original plan's step 5 — "1. Set up my quarter" is
> idempotent and safe to re-run.

> ## 17 AUG, EVENING — First-ever real Scenario 1 ATTEMPT, killed twice,
> AS fixed (a real bug found live), P1 diagnosed but NOT YET FIXED.
> **STATUS: CURRENT, supersedes the block below. READ THIS FIRST — there is
> an unfinished fix (P1) with a precise plan already written out below;
> pick that up before anything else.**
>
> **What actually happened, in order:** with `addin126` deployed (AF/AL/AJ/AR
> all live), Rohan attempted the real thing for the first time this
> project's whole history — open the real deck, "1. Set up my quarter",
> review, tick, "2. Put it on the slides", unaided. Two real defects were
> found live doing this, both fixed and pushed; a third (P1) was found,
> diagnosed precisely, and NOT yet fixed when the session ended.
>
> **AS, FIXED, PUSHED (`36f9947`).** `Drafting.PruneParked`'s `.Delete`
> call (line ~434) had no `DisplayAlerts` guard — a raw, unsuppressed
> Excel "permanently delete this sheet?" alert fired once per field during
> the drafting-sheet rebuild, confirmed live by screenshot. My first theory
> (the sheet-scan itself getting slower) was WRONG — a second opinion
> (waste-hound) caught it before I fixed the wrong thing: this exact code
> path measured 1.8s/field a few hours earlier the same session, so a scan
> over a modestly-larger sheet count cannot explain the observed 82s/103s/
> 110s per field. The real driver: most fields had reached `PruneParked`'s
> `keepNewest=2` threshold, so nearly every field's park this run triggered
> a genuine delete+alert that blocked on a human click — invisible to
> `Timing`'s own instrumentation, silently counted as computation. Fix
> copies `DraftingLobby.bas:304-306`'s own already-proven pattern verbatim
> (`wb.Application.DisplayAlerts`, not bare `Application`). Swept the whole
> codebase for a third sheet-delete site — none exists. Full suite 242/242.
> **Deployed as `addin130`, confirmed loaded live.**
>
> **P1, DIAGNOSED PRECISELY, FIX NOT YET WRITTEN.** Retesting with
> `addin130` (AS fix in place) hit a THIRD recurrence of an old, known,
> never-actually-fixed defect: FIX-LIST's own P1 ("A dialog opens BEHIND
> the PowerPoint window, and reads as nothing happened"), documented
> 2026-08-13 with a proposed fix that was **never implemented for this
> chain**. Confirmed live via window-title enumeration: hidden windows
> titled "Start a Quarter", "Roll Forward", "1. Set up my quarter", and
> "PopupHost" appeared in sequence across two separate presses, each one
> genuinely a legitimate dialog the tool is supposed to show, just buried
> behind whatever window had focus.
>
> **Root cause, read from the actual source, not assumed:** exactly ONE
> call site in the whole codebase does the right thing —
> `DraftingUI.RollForwardUI` calls `BringExcelToFront wb` (`DraftingUI.
> bas:1648`) before its own Excel range-picker `InputBox`, because that
> specific dialog is genuinely Excel-owned (`wb.Application.InputBox`,
> `Type:=8`) and needs Excel visible to click a cell in. Every OTHER prompt
> in the "1. Set up my quarter" chain — the period-confirm `MsgBox` at
> `RibbonUI.bas:1199` included — is a bare, unguarded `MsgBox`/`InputBox`
> call, owned by POWERPOINT's own VBA process (since that's where this
> code runs), with **nothing anywhere bringing PowerPoint's window
> forward** before showing it. Grepped every `.Activate`/`AppActivate`
> call in production `vba/*.bas` to confirm: `BringExcelToFront` is the
> ONLY foreground-bringing helper that exists, and it exists only for
> Excel. There is no PowerPoint-side equivalent anywhere.
>
> **Worse than a missing call — an actively self-defeating one.**
> `RollForwardUI` correctly and deliberately brings EXCEL to the front for
> its own picker. But nothing brings PowerPoint back afterward, so the
> VERY NEXT prompt in the chain (`RefreshDraftingSheets`, called right
> after `RollForwardUI` in `SyncNowChainCore`, `RibbonUI.bas:1356-1358`)
> is now hidden behind the Excel window Roll Forward JUST correctly
> raised. This precisely explains the sequence of different hidden window
> titles observed across two live presses tonight.
>
> **THE PLAN, not yet built — do this first:**
> 1. Write `BringPowerPointToFront()` in `DraftingUI.bas`, right beside the
>    existing `BringExcelToFront` (`DraftingUI.bas:747-760`) — same proven
>    technique (set `Application.Caption` to a marker string, `AppActivate`
>    on that marker since it matches on title PREFIX and PowerPoint's
>    default caption doesn't start with a fixed string either, restore the
>    caption after), but for bare `Application` (correctly PowerPoint here,
>    since this code runs IN PowerPoint's VBA project — this is NOT the
>    bare-`Application` trap `DraftingLobby.bas` hit, that trap was
>    specifically about Excel-hosted code needing `wb.Application`; here
>    bare `Application` is already correct).
> 2. Call it at the very top of `SyncNowChainCore` (`RibbonUI.bas:1176`),
>    before `Set pres = Application.ActivePresentation` — covers the
>    period-confirm `MsgBox` and everything else early in the chain.
> 3. Call it again immediately after `DraftingUI.RollForwardUI` returns,
>    before `DraftingUI.RefreshDraftingSheets` runs (`RibbonUI.bas:1357-
>    1358`) — covers everything from the drafting-sheet rebuild onward,
>    undoing Roll Forward's deliberate (and correct, for ITS purpose)
>    Excel-activation.
> 4. Best-effort only, matching `BringExcelToFront`'s own documented
>    stance — `On Error Resume Next` around it, a failed activation still
>    leaves every dialog fully functional, just possibly behind something,
>    same as today. Do not add waits/sleeps to "help" this — same
>    "Do NOT fix by adding waits" rule this project already wrote down once
>    for the identical class of defect (FIX-LIST P1's own text).
> 5. Test, full suite, build the next addin, redeploy, THEN retry the real
>    Scenario 1 attempt from scratch — "1. Set up my quarter" is idempotent
>    and safe to re-run; nothing was lost by killing the process either
>    time tonight (nothing had reached the final Save both times).
>
> **Current live state at handover:** PowerPoint and Excel both killed
> (force-closed, safe both times — nothing had been saved to disk past
> what `addin130`'s partial run + AutoSave already preserved). `addin130`
> is the currently-registered/loaded addin (has AS, does NOT have the P1
> fix). A safety copy of the mid-incident Excel state was saved earlier to
> `OneDrive\Claude\register-wide.INCIDENT-BACKUP-20260817-183751.xlsx` —
> not needed for recovery (the live register-wide.xlsx itself is fine,
> nothing was lost), kept only as a paranoia artifact, safe to delete
> whenever.
>
> **Also still queued, not started, not forgotten:** the "workaround
> hound" agent idea (Rohan's, rhymes with waste-hound — hunts band-aids
> left in instead of real fixes) — explicitly judged NOT necessary
> tonight, stays queued. `CopyAiDraftsToSubmit`/`PublishDraftsForField`
> (old pre-AF dead Subs) still confirmed dead, still deliberately left in
> place. AN, AD Phase A/B, AG-AJ, AM-AQ (minus AR/AS, now fixed) all still
> open per FIX-LIST.md. The drafting-sheet "simple, color-block" styling
> note is captured and still just a discussion note, not scoped.
>
> **The actual finish line has still not moved.** Tonight got closer than
> any prior session — a real attempt was made, twice, on the real deck,
> unaided except for live debugging — but both attempts were killed before
> reaching a tick/approve/publish cycle. Next session's very first action
> should be finishing the P1 fix above, then trying again. Don't start
> anything else first.

> ## 17/18 AUG, LATE NIGHT — AR fixed (the real dominant cost), the AF live
> retest turned out invalid, `SyncNow`/`SyncNowCore` (AJ) deleted.
> **STATUS: SUPERSEDED by the block above.**
>
> **The AF retest (block below) was invalid — read this before trusting
> anything it says about AF being slow.** `addin135` was built and pressed
> for real; it ran 14.5+ minutes and got killed. Live diagnosis in the
> moment drew two wrong inferences from correct observations. A cold
> second opinion (fable, full transcript this session) read the saved
> file's actual bytes and found the truth: the register under test was a
> reverted baseline missing the `Drafting Lobby` sheet entirely, so
> `PublishAllDraftedFields` exited in SECONDS ("Nothing is pinned"). The
> 14.5 minutes were the button chain's NEXT stage —
> `ReviewChangesCore` → `BuildQueue` → `PlanRoutineSync` — a pre-existing
> path with no relationship to AF/AL, never fast-mode-wrapped or
> Timing-instrumented, hence invisible until now. **AF has still never
> been genuinely retested live.** Next real action: press "1. Set up my
> quarter" FIRST (rebuilds the Lobby), THEN "2. Put it on the slides" —
> only that sequence tests AF for real.
>
> **AR fixed same night — the actual dominant cost fable found.**
> `InjectorFor` calls `FindShapeByRoleTag` up to 4x per field (base, ".1",
> ".track", ".rest"); a genuine miss (no shape for a field on this slide
> type — the majority case now that the register carries more populated
> columns than most slide types have fields) was never cached, only a hit
> was. Fixed via `ShapeAddressBook.RecordAbsent`/`NO_SHAPE_MARKER`, the
> negative twin of the existing positive cache, same invariant. Explicit,
> honest limit stated in the code and in FIX-LIST.md: no verify-on-read for
> negatives (would cost the walk it avoids), so this is trusted, not
> verified — safe only as long as nothing hand-adds a shape to an instance
> outside this tool, which has never happened but isn't prevented either.
> Proven via deliberate-break-first. Full suite 242/242. Committed
> (`e416dd2`), pushed. **Not yet built into an addin or measured live.**
>
> **AJ deleted same night** — `RibbonUI.SyncNow`/`SyncNowCore`, confirmed
> dead via grep (nothing calls either except each other), 329 lines gone.
> Committed (`2fa16b6`), pushed.
>
> **Next session, in order:** build the next addin (staging via
> `build_ppam.ps1`, your manual Save-As, then the trusted-folder-copy +
> registry swap + verify cycle) → press "1. Set up my quarter" → press
> "2. Put it on the slides" for real → read the Timing sheet. That single
> retest is now the actual proof for BOTH AF and AR at once, on a register
> in the state the button expects. Do not trust any number from a run that
> skips the "Set up my quarter" step first.
>
> **Also queued, not started:** "workaround hound" agent idea (Rohan's,
> rhymes with waste-hound — hunts band-aids left in instead of real fixes,
> same read-only diagnose-only pattern). `CopyAiDraftsToSubmit`/
> `PublishDraftsForField` (the old pre-AF per-field Subs) are confirmed
> dead-as-buttons but deliberately left in place — safe next-session
> deletion once AF is proven live. AN, AD Phase A/B, AM/AG-AJ/AO-AQ all
> still open per FIX-LIST.md.

> ## 17 AUG, NIGHT — AF + AL fixed properly (not band-aided), a real bug found
> and fixed inside the proof, `waste-hound` agent built. **STATUS: SUPERSEDED
> by the block above, except where noted.**
> Continuation of the same very long session (16 Aug evening through past
> midnight 17/18 Aug).
>
> **Commissioned `waste-hound`** (`~/claude-brain/agents/waste-hound.md`,
> committed and pushed to `claude-brain`, symlinked into `~/.claude/agents`):
> read-only, cold performance auditor — `model: sonnet` (fable wrote the spec,
> corrected mid-drafting to run on Sonnet, not fable, per Rohan). Encodes six
> defect shapes found across tonight's two audits (eager full work where
> incremental would do; rescan-from-scratch inside a loop; machinery outliving
> the dialog/gate it fed; re-verifying a fast known-good copy against a slow
> authoritative source with nothing invalidating it; no batching on a loop of
> small expensive ops; the same computation repeated twice for the same data
> in one pass) plus a "checked and clean reported with equal confidence as a
> finding" rigor discipline. Never edits code — diagnoses only.
> **Open idea, Rohan's, not yet built:** a companion "workaround hound" —
> same read-only pattern, hunting band-aids left in instead of the real fix
> (the exact thing "proper fundamental fix, don't bandaid unless you're
> bleeding" is policing). Backlogged, not speced.
>
> **Fixed AF and AL properly — Rohan's explicit instruction after the real
> 362.2s/4-field number landed**: "lets go with the proper fundamental fix,
> every time. dont bandaid unless you bleeding," then "apply it across the
> class in line with best practice." Full technical detail in `FIX-LIST.md`'s
> "AF + AL, the real fix" section — summary:
> - **AF**: `PublishAllDraftedFields`'s 13x-per-field redundant resolve/
>   register-read/save/Run-Log-write collapsed to once-per-press, via a new
>   `Public Function DraftingUI.PublishOneFieldForChain` that does the actual
>   per-field work only (wet-only publish, no separate dry preview pass) and
>   is genuinely unit-testable for the first time (explicit params, no
>   `Application.ActivePresentation` dependency).
> - **AL**: `DraftingLobby.FieldIdForSheet` (the Lobby pin-watcher's guard,
>   fired by `AppEvents.cls` on every cell-change event in every open
>   workbook) fixed at its root cause — a real zero-COM prefix check now
>   runs first, instead of the full Field Spec scan its own header comment
>   falsely claimed didn't happen. `ApplyApproved`'s fast-mode wrapper also
>   now disables `EnableEvents` as belt-and-suspenders.
> - **AF-NTP, found DURING the fix**: proving AF's wet-only optimization safe
>   surfaced a real, previously-latent bug in `Drafting.NothingToPublish` —
>   it only ever matched dry-run phrasing, so it could never detect "nothing
>   published" against a wet-mode result. Caught by the "make it fail once"
>   discipline (a new test deliberately broken and reverted first; a full
>   unfiltered suite run then caught this second, real, unrelated failure on
>   its own). Fixed at the source, not worked around at the call site.
>   241/241 tests pass with the fix in.
>
> **Not yet re-measured live.** `addin125` (currently loaded) predates all of
> this. Next session: build `addin126` (Rohan's manual Save-As step, same
> pattern as every prior build), deploy, then press "2. Put it on the slides"
> for real and get the actual post-fix number against the 362.2s/4-field
> pre-fix baseline. Commit made this session covers all of AF/AL/AF-NTP plus
> the `waste-hound` agent (separate repo, already pushed).
>
> **Still explicitly deferred, correctly, not forgotten**: AN
> (`ExcelOutput.UpsertRow`'s per-row rescan inside the publish loop, same
> shape one level deeper) — shared-function surgery across multiple callers
> with different assumptions, needs its own session; `DRAFTING-SPEED-
> STRATEGY.md`'s Phase A/B (bulk-array `WriteDraftingSheet`); remaining
> COLD-PATH-AUDIT.md items AM, AG-AJ, AO-AQ.
>
> **DISCUSSION NOTE, not a task — Rohan explicitly flagged this as "note to
> discuss," not implement**: "I'd like to make the drafting sheets very
> simple and color block stylistic." No design done yet; do not start
> implementing from this note alone. Rohan's answer when asked what "simple/
> color-block" means (2026-08-17/18): **status-color-coding, not just visual
> decluttering.** A cell's state — AI-generated vs. draft vs. final, etc. —
> should be "immediately recognisable by their existence" — the colour block
> itself is the signal, not a label someone has to read. Style cue given
> explicitly: solid colour block (e.g. hot pink), bold white lowercase text
> in the block (his example: bottom-left of the square). "Colour important…
> simple blocking allows for clearer UI indicators too" — i.e. this isn't
> pure aesthetics, it's meant to make cell state legible at a glance where
> it currently isn't. Still needs: the actual state→colour mapping (how many
> states, which colours), and where in the codebase cell formatting is set
> (`Drafting.bas`'s `WriteDraftingSheet` is the likely site, ~600 COM calls
> per field already, so this should probably ride along with `DRAFTING-
> SPEED-STRATEGY.md`'s Phase A bulk-write work rather than be a separate
> pass over the same cells). Scope this properly before touching code.

> ## 17 AUG, AFTERNOON — Timing instrumentation + running-long checkpoint,
> `addin121` deployed. **STATUS: SUPERSEDED by the block above.** Continuation
> of the same very long session (16 Aug evening through 17 Aug afternoon).
>
> **Built `Timing.bas`** (new module), Rohan mid-demo-debrief: "can you please set
> timing machinery in the code... within an order of magnitude regarding what you
> expect." Per-stage duration AND a real per-unit rate (Sec/Unit as its own column,
> not buried in text — "make sure the variables are quantifiable... per unit of
> whatever") written to a `Timing` sheet in the register, documented targets stated
> in the same units. Wired into the two real hot paths (`DraftingUI.
> RefreshDraftingSheets`/`PublishAllDraftedFields`, `ReviewQueue.ApplyApproved`) with
> sub-timing (dry probe vs. real write, accumulated not logged per-item — same
> "don't add the cost you're trying to measure" reasoning as item W's own fix), a
> click-log (`Timing.LogClick` — "have a register of what I clicked when so you can
> diagnose it") so a hang can be read from its own log instead of guessed at, and a
> human-wait exclusion (`LogWait`/`excludeSeconds`) so a MsgBox someone is sitting on
> doesn't inflate the processing rate logged next to it.
>
> **Classic Excel VBA speed tricks researched and applied**, not asserted from
> memory: `ScreenUpdating`/`Calculation = xlCalculationManual` — confirmed used
> NOWHERE before tonight, now wrapped around both hot loops (restored in
> `DraftingUI`'s `Failed:` handler too, guarded by `screenSettingsCaptured` so an
> error before either was ever touched can't wrongly "restore" them to their
> uninitialized defaults). `.Copy`/`PasteSpecial` and `.Select`/`.Activate` checked
> and ruled out. Bulk array read/write (the single most-cited win, one cited case
> 886s→<1s) NOT done — `Drafting.bas` alone has 97 individual `.Cells(` calls, a real
> future opportunity, correctly too large to fold into tonight.
>
> **Real bug caught before shipping: `ApplyApproved` has no top-level error handler**,
> so a re-raised item crash (Error 50290, or the deliberate test fault) skipped the
> new ScreenUpdating/Calculation restore entirely and would have left Excel's screen
> updating off and calculation on manual for the rest of the session. Fixed by
> restoring immediately before each of the two `Err.Raise itemErrNum` sites, not just
> once after the loop.
>
> **Added a running-long checkpoint** — Rohan, right after the demo's ~2 minute
> stall: "include some cancel lines if target exceeded at a reasonable point."
> `Timing.CheckBudgetAndMaybeCancel`, checked every 10 items in `ApplyApproved`'s
> loop, silent whenever elapsed stays under budget (`items * 2 sec/item`, floored at
> 15s so the ratio isn't noise on a handful of items) — only pops a Yes/No once a run
> has genuinely blown past it. Does **NOT** rescue a genuine single-call hang like the
> demo's own (still open, FIX-LIST item X) — a between-items check cannot interrupt
> something already blocked inside one COM call. Cancelling stops the loop cleanly:
> remaining items are logged as "cancelled: user stopped the run early," and — caught
> before shipping — `MarkConsumed` is now **skipped** on a cancelled run, because it
> stamps the whole review sheet consumed and `PendingApprovals` treats a consumed
> sheet as nothing-left-to-do; consuming it on a partial run would have silently
> discarded the still-ticked remaining items.
>
> **Also fixed while touching this: `PublishAllDraftedFields`'s final Timing call was
> passing a String (`list`, the comma-joined field names) into a `Long` parameter
> position** — a real compile-breaking type mismatch, not cosmetic, that would have
> broken every "2. Put it on the slides" press. Caught by re-reading the call site,
> not by the test suite (nothing exercised that exact line).
>
> **Two new tests, one deliberately proven to fail first**
> (`Test_Timing_LogDurationWritesQuantifiablePerUnitRate`,
> `Test_Timing_CheckBudgetAndMaybeCancelSilentWhenUnderBudgetOrFloor`) — the first
> broken on purpose (`elapsed * unitCount` instead of `/`), confirmed it failed with
> the exact wrong number (40 instead of 2.5), reverted, confirmed green. The
> cancel-dialog's ACTUAL cancel branch (ratio exceeded past the floor) is **not**
> covered — it necessarily pops a real MsgBox, which would hang a headless run; a
> named gap, same shape as Phase 3's own untested no-modal apply path.
>
> **Full suite: 240/240 passed** (238 pre-existing + 2 new). Static checks,
> module-list consistency, and doc-control checks all clean.
>
> **`addin121` built, moved to trusted
> (`...\AppData\Roaming\Microsoft\AddIns\`), registered `AutoLoad=1`, `addin120` set
> to `0`, confirmed loaded live** — two ways: `Application.AddIns("addin121").
> Loaded = True` from a fresh PowerPoint launch, AND `Application.Run
> ("addin121.ppam!Timing.StartClock")` actually executed and returned a real Timer
> value — stronger evidence than the AddIns flag alone, since it proves the specific
> new code is genuinely compiled and callable, not just that a correctly-named file
> is sitting in the folder. (Minor snag along the way: PowerPoint's Save As landed
> the file in `OneDrive\Claude\`, not the trusted `AddIns` folder addin120 lives in —
> copied across before registering, same as every prior build.)
>
> **UPDATE, same afternoon: the retest happened, and the Timing sheet answered its
> own question.** Deck deliberately reverted to an older `.bak` snapshot first (per
> Rohan's choice) so "1. Set up my quarter" would have real work to do. Real numbers,
> read straight off the `Timing` sheet: 13× `WriteDraftingSheet` ~40s total (2.5-4.6s
> each), **`BuildLobbyFromScratch` 503.4s (559 rows scanned, 115 pinned) — 75% of the
> entire 665.4s run.** First-ever measurement of this stage (NEXT-SESSION and
> FIX-LIST had both only inferred it before), and it dwarfs everything items W/Y/Z
> touched combined — now FIX-LIST item AB, still open, not yet diagnosed against the
> actual source. Two more real findings from watching this run live: (1) a genuine
> long, currently-unmeasured delay happens INSIDE `Resolve()`, before its own
> period-rollover confirmation dialog even appears — looks exactly like a hang from
> outside (flat CPU, `Responding=True`, no visible dialog checked mid-delay) but
> isn't; FIX-LIST item AA. (2) `RefreshDraftingSheets`'s "(total)" Timing row fires
> BEFORE `WriteRunLog`/`SaveWorkbookVerified` actually run — caught because Rohan
> reported the real completion dialog arrived ~30s after that row appeared; folded
> into item AB. **Also found and worth remembering: `WriteRunLog` doesn't clear
> trailing rows from a longer previous run** — this run's real content ended at row
> 30, rows 31-86 were untouched leftovers from the prior (86-row) run, a minor but
> real bug, not yet its own FIX-LIST item.
>
> **Both AA and AB have a live confound not yet ruled out**: this run followed a
> deliberate deck-swap for the retest, so the register had more than usual to
> reconcile. One clean baseline run against an already-synced deck, untouched, is
> needed before treating 503s as typical rather than worst-case.
>
> **UPDATE, same afternoon: AB diagnosed and FIXED, not yet deployed.**
> `BuildLobbyFromScratch` went through `PinToLobby`/`FindLobbyRow` for every pin,
> which did THREE full rescans-from-row-1 per call (item W's exact shape, stacked
> three deep) — and since this function had just deleted and recreated the Lobby
> sheet, every one of the 115 pins that run was provably new, so the whole lookup
> was 100% wasted work (~33,000 wasted COM calls). Fixed two ways: `LastLobbyRow`
> rewritten to `End(XL_UP)` (one COM call, not n — fixes every caller including
> `AppEvents`' real-time path too), and `BuildLobbyFromScratch` no longer calls
> `PinToLobby` at all — it tracks its own `nextRow` and writes directly, since it's
> the only writer of a sheet it just wiped. `PinToLobby` itself untouched. Full
> suite green (240/240), including both Lobby correctness tests — proves the
> rewrite gives identical results, not just faster ones.
>
> **UPDATE, same evening: `addin122` built, deployed, confirmed loaded live, AB
> re-measured in isolation — 503.4s -> 2.67s, ~188x, identical output.** Measured
> by calling `DraftingLobby.BuildLobbyFromScratch` directly via `Application.Run`
> against the same register workbook, bypassing `Resolve()` and the whole
> `RefreshDraftingSheets` chain entirely — deliberate, not a shortcut: item AA's
> own delay was still live in the same chain during this retest attempt (it hit
> again, ~5+ minutes of real CPU burn this time, no dialog needed since the
> register was already on the rolled-over period — so AA is a real, repeatable
> cost, not a one-off), and isolating AB from a known-slow unrelated neighbour is
> what makes 2.67s trustworthy. The deck-swap confound noted above does NOT apply
> to this number — `BuildLobbyFromScratch` never reads the deck.
>
> **Also found live, mid-wait on AA: the register workbook itself is a bad test
> fixture.** 54 sheets, of which 26 (48%) are `SAVED ...` archive tabs accumulated
> over a week of repeated test cycles, only 15 are genuine working content. Tonight's
> numbers (including AA's delay) were measured against a workbook roughly 2x the
> size a real single-quarter register would be. Not yet acted on — a trimmed or
> fresh test fixture is a real follow-up, separate from anything below.
>
> **UPDATE, same evening: FIX-LIST item AD opened — `Drafting.WriteDraftingSheet`
> diagnosed, strategy written, not yet built.** Full diagnosis and plan in
> `DRAFTING-SPEED-STRATEGY.md` (fresh research pass, fable model, read-only). Real
> shape found in the actual source: ~600 COM calls per field, of which ~240 are
> CONSTANT overhead (cosmetics/formatting) paid on every rebuild regardless of
> whether anything changed — meaning bulk-array reads/writes alone only buy ~2-3x;
> hitting the requested 10x needs a cosmetic-skip-when-unchanged stamp as well.
> Expected: 2.5-4.6s/field -> ~0.15-0.35s/field. Evaluated (and recommended AGAINST
> bundling) merging the 13 per-field drafting sheets into one-per-type — real
> option, but not on the speed critical path, and it means operating again on the
> function with 5 prior data-loss incidents before Lobby Phase 3's own "does this
> still need fixing" gate has even been tested. Also found and fixed in passing: a
> stale comment in `Drafting.bas` claiming parking is unconditional when it isn't —
> the ordinary same-layout-same-period path still has no pre-write backup, same gap
> as the 2026-08-14 incident; comment corrected to state reality, behavior NOT
> changed (that's a separate decision).
>
> **UPDATE, same evening: step 0 (measurement gap) and item AC both DONE.** Every
> previously-unattributed stage inside `RefreshDraftingSheets` now has its own
> `Timing.LogTiming` line (spec+sources write, missing-columns check, period
> validation, register read, tab/index/format, the three FieldSpec validations,
> WriteRunLog+Save), and the "(total)" row moved to after the actual save -- the
> true end of the function. Full suite green (240/240). **NOT YET DEPLOYED** — no
> addin build since. The next real "1. Set up my quarter" press on a build that
> includes this is what actually answers where the missing ~120s lives; not
> guessed at further before that.
>
> **UPDATE, same evening: `addin123` built, deployed, confirmed loaded live, retest
> run for real.** Full run 665.4s -> 284.84s, ~2.3x, every stage now individually
> visible: spec+sources write 5.2s, register read 9.3s (43 rows), 13x
> `WriteDraftingSheet` 23.7s total (1.6-2.1s/field, down from 2.5-4.6s but Phase A/B
> not started so this is incidental, not the real fix), `BuildLobbyFromScratch`
> 168.6s, `ArrangeTabs+WriteWorkbookIndex+FormatRegisterSheet` 53.7s, validations
> 1.8s, `WriteRunLog+Save` 8.8s. Residual unattributed gap ~12.7s (down from ~120s)
> — attribution problem solved.
>
> **UPDATE, same evening: found and FIXED item AE the moment the data landed.**
> `BuildLobbyFromScratch` — already fixed (AB), proven 2.67s in isolation — cost
> 168.6s in this real run: 63x slower, same code. Cause: the fast-mode wrapper
> (`ScreenUpdating`/`Calculation`) added earlier tonight only covers the
> `WriteDraftingSheet` field loop, restored to normal immediately after — so
> `BuildLobbyFromScratch` and the tabs/index/format cluster (also unexpectedly
> slow at 53.7s) both paid full screen-redraw and full automatic recalculation on
> every write, against a real 45-slide deck and 54-sheet register. The isolated
> AB proof wasn't wrong — it just never faced this cost. Fixed by widening the
> wrapper to cover both (checked first for the AutoFit+ScreenUpdating interaction
> that can silently miscompute column widths — neither function uses AutoFit, no
> risk). Full suite green (240/240). **NOT YET DEPLOYED** — `addin124` and one
> more live retest are the real proof.
>
> **UPDATE, same evening: item AA's own theory was WRONG, corrected.** Retesting
> `addin124` (the AE fix), Rohan pushed back on why the period dialog appeared in
> ~17s once and ~5 minutes the next run with no code difference — a fair
> challenge that AA's original diagnosis ("lives inside `Resolve()`") couldn't
> answer, because it doesn't live there. Traced the REAL chain: "1. Set up my
> quarter" runs `SyncNowChainCore`, which calls `WhereAmICore` (a full deck/
> register status scan) BEFORE `StartQuarter`'s dialog, and `Resolve()` doesn't
> run until `RefreshDraftingSheets`, the LAST thing in the chain. `WhereAmICore`
> -> `Readiness.Build` -> two full copies of the ~49MB deck file plus slow
> `Shell.Application` ZIP extraction (re-verifying period/workbook-path from disk)
> plus a full `ReviewQueue.BuildQueue` diff per registered type — on EVERY press,
> to produce one line of status text.
>
> **UPDATE, same evening: `Readiness.bas`/`WhereAmI` DELETED ENTIRELY (item AK,
> FIXED).** Traced through and found almost nothing it checked was novel — the
> real operations (`RefreshDraftingSheets`, `ApplyApprovedCore`,
> `RollForwardPeriod`, `RunSync`) already independently catch and explain every
> failure mode it pre-checked, the moment they actually run. The one distinct
> check (period reported-but-not-saved) is already verified at write time by
> `SetDeckPeriodVerified` inside `StartQuarter`, earlier in the same chain.
> Rohan: "Please get rid of stupid stuff" → "delete the whole thing, keep
> anything useful but otherwise get rid of it." Nothing was worth keeping — full
> reasoning and every redundancy traced in `FIX-LIST.md` item AK. Full suite
> green (240/240). **NOT YET DEPLOYED.**
>
> **UPDATE, same evening: hot-path audit (fable) found the same shape 5 more
> times** — `HOT-PATH-AUDIT.md`, FIX-LIST items AF-AJ. Headline: `PublishAllDraftedFields`
> ("2. Put it on the slides", the most-pressed button) redoes press-level work
> 13x, once per field — ~4 min of redundant register re-reads, ~2-3.5 min of
> redundant saves, no fast-mode wrapper on the loop at all (item AF). Four more
> in the same "fed a since-deleted dialog" shape (AG-AJ). None fixed yet.
>
> **UPDATE, same evening: needs-vs-build comparison (separate fork) raised a real
> "are we on track" question, and Rohan gave the real answer.** The comparison
> found tonight's whole session went to sync-speed work while the project's own
> manual-baseline memory says speed isn't the dominant cost of a quarter, and the
> stated finish line (a real quarter reviewed/approved/published UNAIDED) hasn't
> moved. **Rohan's correction, worth recording verbatim: "the performance was
> killing the testing, I wasn't able to use it."** The speed work wasn't a
> distraction from the real goal — multi-minute stalls and unclear hang-vs-working
> states were making the tool's OWN testing/verification loop unusable, which is
> a genuine precondition for reaching the finish line, not orthogonal to it. Both
> things are true: the finish line still hasn't moved, AND tonight's fixes were
> plausibly necessary before it safely could. Worth an honest PM-style check next
> session with BOTH facts on the table, not just one.
>
> **UPDATE, same evening: `addin125` built, deployed, confirmed loaded live —
> Rohan pressed "1. Set up my quarter" for real, unaided.** 665.4s -> 72.8s
> total (~9.1x), `BuildLobbyFromScratch` 168.6s -> 6.9s, tabs/index/format
> cluster 53.7s -> 6.3s, the pre-dialog wait (item AA/AK) ~20s (was 5+ minutes).
> Direct, live, real-button-press evidence for AE and AK both landing correctly.
>
> **UPDATE, same evening: PM re-check on Rohan's "performance was killing the
> testing" claim, verified against the live Timing sheet, not the transcription
> above.** Verdict: **on track** — W/Y were hit live during the actual manager
> demo, AC came from Rohan's own confusion about whether a run had finished, AK
> was a tax paid on every real press regardless of who was watching. This was
> defect remediation of things that hit him directly, not speed-for-its-own-
> sake — the earlier needs-vs-build framing was the wrong lens. But the finish
> line (Scenario 1, unaided) still hasn't moved, and the PM's exact words: "the
> actual test — does Scenario 1 go unaided now — has not happened... this is
> Rohan's five minutes, not more code." Explicit stopping condition: don't open
> new lettered FIX-LIST items speculatively before that's tried.
>
> **UPDATE, same evening: second fable audit (`COLD-PATH-AUDIT.md`, items
> AL-AQ) found the same defect class 6 more times across the rest of the
> codebase**, most significantly AL — the Lobby pin watcher (`AppEvents.cls`)
> taxes every single cell write anywhere in the tool (~100-170 COM calls/event),
> despite its own comments claiming "one comparison" — and AN, `UpsertRow`
> rescanning the whole register per row inside the publish loop, item AB's
> shape again.
>
> **UPDATE, same evening: Rohan pressed "2. Put it on the slides" for real —
> the actual other half of Scenario 1's chain, and the PM's own stopping
> condition is now satisfied.** `PublishAllDraftedFields (total)`: **362.2s for
> 4 fields, 90.6 sec/field** — worse than AF's own estimate, and exactly where
> AF + AL + AN all stack on the same real path. This is real, Rohan-witnessed
> evidence, not a projection — the standing PM condition for fixing AF next
> ("if the number comes back large... that's sufficient real-user evidence")
> is now met.
>
> **UPDATE, same evening: Rohan's real "2. Put it on the slides" press is still
> in progress at time of writing** — `PublishAllDraftedFields` completed and
> saved (362.2s, 4 fields, logged above), now inside `ApplyApproved`'s own
> uninstrumented review-sheet read (finding AI/AL: `ScanPendingApprovals` +
> `ApplyApproved` both call `ReadQueueSheet` on both review sheets --
> `project-status-2D3D` (38 stale rows since 2026-08-10) and
> `project-progress-A32C` (223 rows, CONSUMED but still walked in full) --
> per-cell reads, no bulk read, no Timing instrumentation on this specific
> stage). CPU climbing steadily (Excel ~9.5min, PowerPoint ~8min accumulated
> at last check), workbook dirty again (new write cycle active). Whatever the
> real `ApplyApproved (total)` number turns out to be is the actual close of
> Scenario 1's second half -- read it off the Timing sheet next session if it
> wasn't captured live.
>
> **UPDATE, same evening: a standing "efficiency auditor" subagent is being
> built**, at Rohan's request after tonight's two fable audits both paid off
> live. Not deck-sync-specific -- goes in `~/claude-brain/agents/` so it's
> reusable across any of Rohan's repos. Authored by fable (drawing on what
> just worked twice tonight: the six defect-shape taxonomy, file:line
> grounding, frequency x cost ranking, explicit "checked and clean" reporting,
> reusing a repo's own calibration numbers instead of inventing estimates),
> but runs on Sonnet once built -- fable's role was one-time authorship, not
> the agent's ongoing identity. Read-only (Read/Glob/Grep/Bash, no Edit/
> Write) -- diagnoses, does not implement, matching the diagnose-cold/
> implement-with-context separation this whole session was built on. Rohan's
> explicit brief for its persona: "offended by even the slightest
> inefficiency, but volcanic when significant" -- with the hard constraint
> that the tone never substitutes for evidence; every strong reaction has to
> survive being checked against the actual file line before it's written
> down. Not yet saved to disk as of this update -- check
> `~/claude-brain/agents/` for the actual filename next session.
>
> **Next session's actual priority, in order: (1) fix AF + AL + AN together** —
> they're the same real path, all three stack on the same measured 362.2s
> number, and fixing them separately would mean re-measuring the same press
> three times. Restructure `PublishAllDraftedFields` to do press-level work
> once (one Resolve, one register read, one save, one log), add
> `EnableEvents=False` to the fast-mode wrapper, and hoist `UpsertRow`'s
> constant lookups out of its row loop using the register read the caller
> already has. Build alongside/behind nothing risky needed here — unlike
> `WriteDraftingSheetBulk`, none of this touches typed-content preservation
> logic. **(2)** Once fixed and deployed, complete the actual close of
> Scenario 1 (review, approve, publish, unaided) — the real outstanding test.
> **(3)** AM (`WriteQueueSheet`'s missing wrapper), AG-AJ, AO (membership-check
> redundancy) — same shape, lower individual weight, fold in opportunistically.
> **(4)** `WriteDraftingSheetBulk` per `DRAFTING-SPEED-STRATEGY.md`'s Phase A —
> still deliberately not started, real surgery on a function with 5 prior
> data-loss incidents deserves a fresh run, not more of an already marathon
> one. **(5)** Build a trimmed or fresh test fixture — this session's register
> is ~2x a real deck's sheet count from accumulated test archives.
>
> ## 17 AUG, MIDDAY — Lobby fixes W/Y deployed as `addin120`. **STATUS: SUPERSEDED by
> the block above**, kept for the detail. Continuation of the
> same very long session (16 Aug evening through 17 Aug midday, including a live demo
> to Rohan's manager on this same machine, using `addin119` as-deployed).
>
> **Two real fixes are BUILT, TESTED, and PUSHED but NOT YET DEPLOYED — `addin119`
> still lacks both.** Discovered this precisely because the live demo used `addin119`
> unmodified and hit exactly the slowness both fixes address:
>
> 1. **FIX-LIST item W** — `ReviewQueue.AppendLogLine`'s O(n²) Sync Log rescan.
> 2. **FIX-LIST item Y** — `InjectPrimitive.FindShapeByRoleTag` walking every shape on
>    the slide, twice per item, every time. New module `ShapeAddressBook.bas`: a
>    persistent, self-healing cache of "which shape answers to this field on this
>    slide type," wired the same way `DraftingLobby.bas` is (a module-level workbook
>    reference set from `WorkbookBridge.OpenOrGetWorkbook`, no new parameters threaded
>    through the injector family). Found a real, load-bearing PowerPoint quirk
>    building it — auto-generated shape names resolve via type+ordinal even after a
>    rename, confirmed to survive a real save/close/reopen — now in `AGENTS.md`.
>
> **Live evidence during the demo, not just theory:** pressed "2. Put it on the
> slides" on `addin119`, watched it grind for several minutes with the same content
> already mostly synced from earlier. A background monitor (polling Sync Log row
> count + CPU every 20s) confirmed a genuine ~2 minute stall — CPU nearly flat, no
> dialog (Rohan checked the screen directly), `Ctrl+Break` sent to PowerPoint did NOT
> interrupt it. That last point is itself real evidence: consistent with the stall
> being inside a blocked synchronous COM call into Excel (a big register scan), not
> a PowerPoint-side hang — logged as an update to FIX-LIST item X. Closed by killing
> both processes (test deck, zero real risk), not chased further live.
>
> **`addin120` built, deployed, confirmed loaded live** — both fixes (W: AppendLogLine,
> Y: ShapeAddressBook) are now actually running, not just committed. Next real action:
> re-run the same "2. Put it on the slides" scenario that stalled during the demo and
> see whether it's gone or just smaller — that comparison is the actual proof, not a
> synthetic test.
>
> ## 17 AUG, MORNING — Lobby phases 0-3 deployed, demo prep. **STATUS: SUPERSEDED by
> the block above**, kept for the detail.
>
> **Lobby Phases 0-3 all built, tested, and DEPLOYED (`addin119`).** Phase 3 (pre-ticked
> queue + no-modal apply) was built and pushed overnight but deliberately held back from
> deployment — the most consequential change in the design, changing what counts as a
> human decision on content reaching a real slide. Rohan reviewed it this morning: asked
> real questions about the layered no-overwrite mechanism (build-time diff filter,
> apply-time hash revalidation, the injector's own no-op check) and about a proposed
> "register as memory" shortcut to skip live shape reads (correctly rejected — it would
> trust a remembered value over the live slide, the exact anti-pattern this project has
> been burned by before), then approved: *"let's get it built properly."* `addin119`
> built, moved to trusted, hash-verified, registered `AutoLoad=1`, `addin118` set to
> `0`, confirmed loaded live via a fresh PowerPoint launch.
>
> **Phase 4 (sheet-merging) intentionally not started.** Explicitly gated on whether
> Phases 0-3 already solve the tab-clutter complaint once used for real (they haven't
> been yet), and structurally riskier than anything else this session — real surgery on
> `Drafting.WriteDraftingSheet`'s fixed-row-position assumptions, the same function with
> five prior real incidents including one that wiped 129 paragraphs, 43 ticks and 75
> notes. Discussed and agreed with Rohan not to bundle it in.
>
> **`PutItOnTheSlidesCore`'s new no-modal path still has no automated test of its own**
> (same gap `PublishAllDraftedFields` has — needs `Application.ActivePresentation`,
> nothing in this codebase exercises that chain end-to-end). The next real "2. Put it
> on the slides" press, live, is the actual first proof of Phase 3 in anger.
>
> **Lobby Phase 2 built, proven live, committed (`af74908`).** `DraftingUI.
> PublishAllDraftedFields` now reads `DraftingLobby.ReadLobby`/`DistinctPinnedFields`
> instead of crawling all 13 declared Prose fields — proven on
> `PRESERVED-known-good-20260815-1050`: with 39 rows pinned across two fields
> (`ABOUT_BODY` ×1, `PROGRESS_BODY` ×38), "2. Put it on the slides" ran Copy+Publish
> for exactly those two, confirmed from the saved workbook's own `Drafting Lobby`
> sheet. Safety valve: `RefreshDraftingSheets` ("1. Set up my quarter") now silently
> repairs the Lobby every run, at no extra cost. `LOBBY-DESIGN.md` and `CHECKLIST.md`
> updated in the same commit. **Phases 3-4 not started.**
>
> **Error 50290 (FIX-LIST item V) interrupted the live retest** — third occurrence
> across three sessions, three different call sites, still no root cause. Rather than
> chase it live at 1am, closed the actual diagnostic gap instead (`211f7c8`):
> `ReviewQueue.ApplyApproved`'s per-item write now traps `Err.Number`/`Description`/
> `Source` locally, logs `"CRASHED in dry probe/real write: ..."` to the Sync Log
> BEFORE re-raising, and re-raises with the specific `EntityKey`/`FieldID` folded into
> `Err.Source`. **Next occurrence will finally name which item was mid-write instead
> of just "VBAProject".** Proven correct via a gated test-only hook
> (`ReviewQueue.mTestForceInjectCrash`) since Office cannot be made to raise 50290 on
> demand — the test genuinely failed against the unwrapped code first, then passed.
> Root cause itself is STILL OPEN.
>
> **Byproduct: fixed a real gap in `check_vba_static.py`.** Its declaration-order
> check only ever matched `Type`/`Const`/`Enum` — not a bare module-level variable —
> and that exact blind spot let TWO real compile errors through "clean" in one night
> (`DraftingLobby.mAppEvents` earlier, `ReviewQueue.mTestForceInjectCrash` while
> building the 50290 fix itself). Added `VAR_DECL_RE`, proved it by deliberately
> reintroducing the real defect and confirming the checker now catches it.
>
> **A real mistake this session, owned, not hidden:** closing Excel "without saving"
> mid-session (checked the register file's timestamp, didn't check whether the
> review sheet's tick marks had been persisted) discarded the in-progress review
> ticks that would have let the 50290 retest resume exactly where it crashed.
> Re-ticking the resulting fresh 183-item "project-progress" review by hand wasn't
> attempted — not worth it at 1am for a diagnostic retest. Next real occurrence
> (whenever Rohan is doing genuine review work) is now the actual test.
>
> **MACHINE STATE AT HANDOVER (final, ~02:20):** `addin118.ppam` saved, copied to
> the trusted `AddIns` folder, hash-verified against the `OneDrive\Claude\` copy,
> registered `HKCU\...\PowerPoint\AddIns\addin118` with `AutoLoad=1`, `addin117`
> set to `AutoLoad=0`. **Confirmed loaded live** via a fresh PowerPoint launch
> (`AddIns` collection, `Loaded = -1`) — this is the current build, includes
> tonight's `ReviewQueue.bas`/`TestRunner.bas`/`DraftingUI.bas`/`AGENTS.md`/
> `FIX-LIST.md` changes. PowerPoint and Excel both closed. Full suite 235/0,
> static/module-list/doc checks all clean, everything pushed to `main` at
> `3f18caa`.

> ## 17 AUG, LATE NIGHT (session started 16 Aug evening). **SESSION-END HANDOVER.
> STATUS: CURRENT.** Very long single session — document control catch-up, the elapsed
> bar built as a real new field, then a full architecture pivot (the Lobby) designed and
> two of its four phases built and PROVEN LIVE. Supersedes everything below on current
> build/design state. **Read `LOBBY-DESIGN.md` in full before touching anything it
> names** — it is now the primary design doc, `CHECKLIST.md`'s Lobby section just points
> to it.
>
> ### THE BIG THING: THE LOBBY. Architecture pivot, not a bug fix.
> Rohan, after living through the two-press review/apply cycle and the 13-sheet crawl
> one too many times: *"it now needs to start speeding up and be far less annoying to
> use... I need to start showing it to people."* Full plan written and reviewed BEFORE
> any of it was built, per his explicit request (*"do a full architectural plan for
> this before we start, large changes"*) — see `LOBBY-DESIGN.md` for the complete
> reasoning, not summarised again here. Headline decisions, all made live with him, all
> recorded with the actual reasoning in the doc: a "pin, not scan" model (drafting
> sheets pin to one shared Lobby on the APPROVE tick, nothing crawls anymore); the
> register-to-slide queue defaults to PRE-TICKED/opt-out (not the exception-list model
> first proposed — Rohan pushed back twice, correctly); onboarding stays out of scope
> for pre-ticking (its own existing deliberate flow); a numeric threshold idea was
> raised and explicitly rejected as worse than the categorical signal already available.
>
> **Phase 0 (cold-start crawl + core mechanics) and Phase 1 (the live pin-on-tick event)
> are BUILT AND COMMITTED. Phase 1 is PROVEN LIVE**, not just green in the test suite:
> with the real add-in loaded and a real deck/register open, a raw COM write to a
> drafting sheet's APPROVE column — zero direct calls to `PinToLobby` — caused the
> `Drafting Lobby` sheet to be created and correctly pinned automatically, and
> un-ticking correctly cleared it. Both directions verified from the saved workbook.
> **Phases 2-4 (wire Publish to read the Lobby instead of crawling; pre-tick the queue
> and drop the Yes/No/Cancel gate; revisit sheet-merging only if still needed) are NOT
> STARTED.** The Lobby populates itself correctly now, but nothing downstream reads it
> yet — the 13-sheet crawl is still what actually runs on every "Put it on the slides"
> press today.
>
> **New file: `vba/AppEvents.cls`** — the first class module this codebase has ever
> had. Getting it working surfaced two genuinely new, well-documented gotchas (both
> logged as classes in `AGENTS.md`, worth reading before touching `WithEvents` or any
> future `.cls` file again):
> - `WithEvents` needs an early-bound type, so this project now carries a reference to
>   Excel's own object library, added *programmatically* in all four scripts that build
>   a presentation from scratch (`build_ppam.ps1`, `run_vba_tests.ps1`,
>   `compile_check.ps1`, `field_e2e.ps1`) — each one starts fresh every run, so the
>   reference has to be re-added every time, it cannot be a one-off VBE setting.
> - **The real one:** a `.cls` file with this repo's normal LF-only line endings
>   imports SILENTLY as a Standard Module instead of a Class Module —
>   `VBComponents.Import()` needs CRLF to recognise the `VERSION 1.0 CLASS` header at
>   all. The only symptom was a generic `WithEvents` compile error that read exactly
>   like a `WithEvents`-specific problem; two wrong theories (Public-vs-Private,
>   qualified-vs-bare type name) were tried and discarded before checking the imported
>   component's actual `.Type` property directly settled it. Fixed by CRLF-normalising
>   `.cls` files specifically during staging, in all four scripts.
>
> **New file: `vba/DraftingLobby.bas`** — `PinToLobby`, `ReadLobby`, `ClearLobbyEntry`,
> `LobbyCount`, `BuildLobbyFromScratch` (the cold-start/repair crawl, usable standalone
> right now), `EnsureWatching` (wires the event sink, called from
> `WorkbookBridge.OpenOrGetWorkbook`), `FieldIdForSheet`. Also fixed along the way: a
> `Collection` cannot hold a VBA `Type` (compile error, hit a fourth time — see
> `AGENTS.md`).
>
> **Three commits tonight, all pushed to `main`:** `eee22d1` (the milestone/derived-field
> apply-path fix, item V, plus the elapsed-time bar's first cut and the `build_ppam.ps1`
> save-before-quit fix, item O), `5eb4c28` (`LOBBY-DESIGN.md` + Lobby phase 0),
> `b67941c` (Lobby phase 1, proven live).
>
> ### THE ELAPSED-TIME BAR — built, tagged live on `3_P001`, ONE BUG STILL OPEN
> New `Kind = Derived` field, `TIMELINE_ELAPSED`, computed fresh every sync from
> `START_DATE`/`END_DATE` (`SyncOperations.ElapsedFraction`), never a register column.
> Discovery proven live (correct fraction computed, ~0.9578, matches the real dates).
> Applying it is NOT reliably proven — `ReviewQueue.ApplyApproved` was extended for
> `Kind = Derived` fields (same shape as the earlier milestone-device fix, item V) and
> the actual bar-write succeeded once, but a SEPARATE, still-unexplained "changed since
> you approved it" hash-mismatch recurred on retry even after rounding the width to 2dp
> (the first theory, tried and it didn't fully fix it). **FIX-LIST has the full
> diagnostic trail — read it before re-diagnosing from scratch.** Session ended this
> thread deliberately, at Rohan's call, to go think about the workflow architecture
> instead of one more live-debug round; that's what led directly to the Lobby.
>
> ### SCENARIO 1'S OFFICIAL UNAIDED CLOSE IS STILL NOT ACHIEVED
> Unchanged from before tonight — every mechanism proof this session (milestone device,
> elapsed bar, the Lobby) was heavily assisted, same as the earlier ABOUT_BODY/milestone
> proofs. This remains open and is not what tonight's work was aimed at closing.
>
> ### MACHINE STATE AT HANDOVER — CHECK FRESH, DO NOT ASSUME
> Checked directly at session end: PowerPoint has **0 presentations open** (the deck
> closed at some point, cause not chased). Excel still has
> `PRESERVED-known-good-20260815-1050\register-wide.xlsx` open, **Saved=False**. That
> workbook now contains, live but unsaved: the `Drafting Lobby` sheet (header row only,
> both live test pins were cleaned up in the same test), the `TIMELINE_ELAPSED`/
> `TIMELINE_ELAPSED.rest` tags on `3_P001`'s slide (in the closed deck, separately
> saved earlier and confirmed surviving a rebuild), and a rebuilt
> `Review project-progress-A32C` queue (183 changes, untouched, "nothing has been
> written" per its own report). **`PRESERVED-known-good-20260815-1050` is further from
> pristine than ever** — treat the name as historical, not descriptive, same note as
> last time. If Excel is closed without saving, the Lobby sheet and this session's
> register-side test state will not persist; the CODE is safe regardless (committed and
> pushed), only this one workbook's live content is at risk.
>
> ### addin115 IS THE CURRENT BUILD, confirmed loaded and working tonight.
>
> ### FULL PIPELINE PROVEN END TO END, ONE FIELD, VERIFIED AT EVERY STEP FROM FILE BYTES.
> `TPL_ABOUT_BODY` row for `3_P001`: typed into `SUBMIT`/`APPROVE` -> `Publish Drafts`
> wrote it into the register (verified in the saved `.xlsx`'s `sharedStrings.xml`,
> independent of Excel) -> `2. Put it on the slides` built a review queue
> (`Review project-progress-A32C`) with the diff -> ticked `Y` on that one row only ->
> `Apply Approved` wrote it to the real slide (verified in the saved `.pptx`'s
> `slide1.xml`, independent of PowerPoint): `ABOUTTEST TEXT AS INSTRUCTED TO
> TYPECI / PI...`. **Not counted as the official unaided Scenario 1 close** — heavily
> guided throughout — but the mechanism itself is now genuinely proven, at the file
> level, not just trusted from dialog text.
>
> ### THE REVIEW QUEUE (R13) IS A DELIBERATE SECOND GATE, NOT REDUNDANT WITH DRAFTING
> APPROVAL — worth recording since it came up as a real question tonight. Drafting's
> `APPROVE` answers "is this text right?". The review queue answers "is this change,
> to this exact slide, still right at the moment it's about to become irreversible?" —
> catches drift between draft-approval-time and apply-time (slide hand-edited
> meanwhile, etc). Confirmed from `ReviewQueue.bas`'s own header: built after a real
> incident (`PROJECT_STATUS`, a non-prose field with no prior approval at all, changed
> 19 slides with nobody seeing a before/after). Landed on: keep it for prose fields too
> — it's usually a no-op confirmation, but it's the last check before the write is
> irreversible, same as a final look at an envelope before it's posted.
>
> ### MACHINE STATE AT HANDOVER — DECK AND REGISTER ARE IN TWO DIFFERENT FOLDERS.
> PowerPoint has `AppData\Local\deck-sync-backups\WORKING-20260816-180335\3. Project
> Progress.pptx` open (Saved=True). Excel has
> `PRESERVED-known-good-20260815-1050\register-wide.xlsx` open — the deck's stored
> `DeckSyncWorkbookPath` still resolves there (FIX-LIST L, hit live twice tonight).
> This mismatched pair is WORKING correctly via the fallback, but don't assume it's
> tidy — check state fresh next session rather than assume either folder is what you
> left it as. **`PRESERVED-known-good-20260815-1050` is no longer pristine** (period
> is `Q1F27`, register has real Q1F27 rows and one real published field) — the name is
> now aspirational, not accurate.
>
> ### THE REVIEW QUEUE STILL HOLDS 181 UNREVIEWED CHANGES.
> `Review project-progress-A32C`, rebuilt fresh tonight (run stamp `2026-08-16
> 18:52:50`), has 181 changes still sitting unapproved after tonight's one deliberate
> tick was applied. Real diffs from the Q1F27 roll-forward against the still-Q4F26
> slides — includes the known bad ones (`PROJECT_PROGRESS` format bug, item M; invisible
> date-character diffs, item N — both re-confirmed live in this exact queue tonight, not
> re-derived). **Do not bulk-approve this queue** — same standing rule as before.
>
> ### FOUR FIX-LIST ITEMS TOUCHED TONIGHT, ONE NEW.
> **L** (workbook path fallback) and **M**/**N** (format/invisible-diff bugs in the
> review queue) were all hit LIVE, corroborating existing entries — not re-derived from
> scratch. **T is new**: `Sources.ApplyPeriodValidation` swallowed a real `Err.Number`
> during tonight's successful roll-forward, reporting only a generic "Excel refused"
> message. Did not block or corrupt anything else. See `FIX-LIST.md` for detail.
>
> ### THE ONEDRIVE\CLAUDE SIGHTING FROM EARLIER TONIGHT IS EXPLAINED, CONFIRMED HARMLESS.
> Reproduced and SEEN this time (screenshot): a normal File Explorer window, tab titled
> "Claude", browsing straight into the live `OneDrive\Claude` folder. Matches last
> session's theory exactly — `RegistryValueOnDisk`'s `Shell.Application`/`CopyHere`
> verification technique popping the window open as an unwanted but harmless visible
> side effect. Read-only browsing, confirmed nothing written by it. Not chased further
> tonight (not asked to); if it needs closing for real, the fix is removing
> `Shell.Application`/`CopyHere` from the verification path, per last session's note.
>
> ### SCENARIO 1'S MECHANISM SUCCEEDED FOR REAL THIS TIME — ROLL FORWARD ACTUALLY RAN.
> On `AppData\Local\deck-sync-backups\PRESERVED-known-good-20260815-1050\` (the deck's
> stored `DeckSyncWorkbookPath` keeps resolving here regardless of which copy the deck
> itself is opened from — see FIX-LIST item L, hit live tonight, not re-derived).
> `1. Set up my quarter` ran its full chain: deck period confirmed `Q1F27`; Roll Forward
> copied all **43 rows Q4F26 -> Q1F27**; `Q4F26` archived as its own file beside the
> register; workbook saved; Refresh Drafting Sheets rebuilt all `TPL_*` sheets and
> reported ready. Verified independently via COM at each step, not just trusted from the
> dialog text. **Still NOT counted as the official unaided close** — heavy diagnostic
> assistance and guidance throughout, same reasoning as the 15 Aug attempt (`769a280`).
> What this proves: the mechanism itself is now sound on a real quarter turn end to end.
> A clean solo rerun (deck already knows the steps) is the only thing still needed to
> flip the count.
>
> **New, real, and separate from the above:** `Sources validation: NOT APPLIED -- Excel
> refused` appeared in the Run Log during this same successful run — a genuine live
> defect, not routine noise. Logged as FIX-LIST item **T**: the real `Err.Number`/
> `Err.Description` is swallowed by `Sources.bas`'s `On Error Resume Next` block, so the
> actual cause was never learned. Did not block or corrupt anything else in the run.
>
> **Testing-methodology lesson, not a code defect:** copying a deck+register pair to a
> new folder (`WORKING-20260816-180335` was built for this, unused in the end) does NOT
> repoint the copy at its own sibling register, because `GetWorkbookPath` only falls back
> to the sibling when the ORIGINAL stored path is missing — and it wasn't, because the
> original folder was left in place alongside the copy. To make a truly independent test
> copy: either don't leave the original at the same absolute path, or use
> `RepointWorkbookUI` (`CAP_REPOINT_WORKBOOK`, written and statically proven per FIX-LIST
> L, not yet confirmed in a built add-in) to explicitly point the copy at its own
> register.
>
> **PRESERVED-known-good-20260815-1050 is no longer pristine** — tonight's real run
> changed its deck's period to `Q1F27` and rebuilt its register for real. Don't assume
> "known-good" still means untouched; check state fresh next time this folder is used,
> same discipline as everywhere else.
>
> **Next action:** you're mid content-review — the 43 rolled-forward `Q1F27` rows are
> sitting in the `TPL_*` drafting sheets (`ORIGINAL`/`REPORTED LAST TIME` columns
> populated, `SUBMIT`/`APPROVE` columns blank, ready for column G wording + `Y` in column
> H), waiting to be approved and published via `2. Put it on the slides`. That publish
> step has not been attempted yet this session.

> ## READ `SESSION-PROTOCOL.md` FIRST, EVERY SESSION. THIS IS NOW MANDATED.
>
> Not this file. `SESSION-PROTOCOL.md`. It says what order to read things in
> (`DOCUMENT-MAP.md` before this one), the documentation-upkeep discipline, where
> test results actually live, and how Rohan works on this project specifically.
> Written 2026-08-16 after a real architecture decision sat undiscovered for four
> days because no session had done what it now mandates.

> ## 16 AUG, EVENING. **SESSION-END HANDOVER. STATUS: CURRENT.** Long session,
> multiple compactions. Supersedes everything below on build state, OneDrive risk,
> and Scenario 1 status. `CHECKLIST.md` is still the primary tickable surface —
> this block is the narrative *why* and the one open thread that needs eyes.
>
> ### BUILD: `addin110` IS CURRENT, TICKED, LOADED. Includes BOTH fixes below.
> Copies in `OneDrive\Claude\` and the trusted `AppData\Roaming\Microsoft\AddIns\`.
> `addin109` should be UNTICKED if it still is — having both loaded risks duplicate
> toolbar buttons (`CommandBarUI`'s Auto_Open runs for each loaded add-in).
>
> ### THE ONEDRIVE WRITE-RELIABILITY RISK IS FULLY CLOSED, BOTH SIDES, PROVEN LIVE.
> `FIX-LIST.md` items **P** (deck) and **S** (register) — read those two entries for
> the full story, this is the short version:
> - **P, the deck:** `DeckSyncPeriod`/`DeckSyncWorkbookPath`/`DeckSyncType`/
>   `DeckSyncTemplate`/`DeckSyncId` moved off `Presentation.CustomDocumentProperties`
>   (which only ever lands a session's FIRST write on a cloud deck, permanently stuck
>   after that — confirmed, three rescue attempts failed) onto a dedicated hidden
>   slide (`DeckSyncRegistry`), keyed by shape name. Proven: **8/8 repeated writes
>   landed on one reused open cloud file** through the real `addin109`/`addin110`
>   `SetDeckPeriodVerified`, the exact scenario that failed 0/8 before.
> - **S, the register:** same class of defect, `ExcelOutput.WriteDeckReference`'s
>   `Workbook.CustomDocumentProperties`, found by asking "check the register too"
>   after P closed — not re-derived, generalised on purpose. Narrower shape (new
>   properties land fine, RE-writing an existing one never does) but the same real
>   consequence: `StampPairing` re-stamps on every repoint, so a workbook re-paired
>   to a different deck would silently keep reporting the OLD one forever. Moved onto
>   a cell on a very-hidden `DeckSyncMeta` sheet. Proven across two genuinely separate
>   sessions (stamp, close, reopen, confirm; re-stamp with a DIFFERENT value, close,
>   reopen, confirm the new value landed).
> - Both fixes read-fall-back to the old `CustomDocumentProperties` location, so
>   decks/workbooks stamped before 2026-08-16 keep reading correctly until the next
>   real write moves them onto the new mechanism. No separate migration step.
> - **Untested: OneDrive for Business / SharePoint** (the likely home of the REAL
>   work deck). Everything above was proven against personal OneDrive
>   (`d.docs.live.net`). Rohan's explicit call, 2026-08-16: get everything working
>   locally/personally FIRST, verify SharePoint later. Don't re-open that question
>   unprompted — it's deliberately deferred, not forgotten.
> - Two real bugs found and fixed building this, both worth remembering as a CLASS
>   not an instance: `sh.Namespace()` needs a `Variant` not a bare `String` (same
>   defect `PropertyOnDisk` already documented once — check every `Namespace()` call
>   in the repo if this class resurfaces); and a genuine VBA `""` success return
>   marshals as PowerShell `$null` through `Application.Run`/`InvokeMember` — an
>   early "8/8 failed" probe report was a false negative from the TEST SCRIPT, not
>   the fix, caught only by checking the saved file's actual bytes independently of
>   the return value.
>
> ### SCENARIO 1 ATTEMPT: STARTED, OUTCOME UNCONFIRMED, NOTHING UNSAFE HAPPENED.
> Opened `AppData\Local\deck-sync-backups\PRESERVED-known-good-20260815-1050\
> 3. Project Progress.pptx` (the sanctioned local test copy) for Rohan via COM,
> visible, 44 slides. Rohan pressed `1. Set up my quarter` (the REAL button caption
> is `CAP_SET_UP_QUARTER` = "1. Set up my quarter" -- `TOOLBAR.md`'s own prose table
> says "Start the quarter", which is STALE; trust the `CAP_*` constant in
> `CommandBarUI.bas`, not the table). Rohan then reported seeing something involving
> "onedrive/claude" on screen and was asked to stop rather than press further.
>
> **Verified directly afterward, not assumed:** PowerPoint's process was still alive
> but reported ZERO presentations open (the deck appears to have closed along the
> way — cause unknown, not confirmed a crash vs. a normal close). Checked the actual
> files: the LOCAL test deck and register **both show their original 15 Aug
> timestamps, completely untouched** — no write happened. The LIVE OneDrive register
> (`OneDrive\Claude\register-wide.xlsx`) **also shows its untouched timestamp from
> before this entire session began** — confirmed by direct comparison against the
> very first listing taken at the start of today's OneDrive investigation. Whatever
> the "onedrive/claude" sighting was, it did not write to anything, safe or live.
>
> **Theory, PARTIALLY corroborated by Rohan directly:** `RegistryValueOnDisk` (P's
> new verification function) uses the same `Shell.Application` `CopyHere`
> zip-extraction technique that was already caught red-handed once today causing an
> unrelated blank Word window to pop up mid-probe. Pressing "Set up my quarter" calls
> `SetDeckPeriodVerified`, which calls this same mechanism repeatedly. Rohan confirmed
> what he saw WAS `OneDrive\Claude` specifically — the exact folder this whole
> mechanism operates in (the live deck, register, and every scratch probe folder all
> live there) — which matches the theory's location. **Still not fully explained**:
> the precise OS/Explorer-level mechanism that turns real `CopyHere` shell activity
> into a visible window is genuinely outside this code's (or any code's) visibility —
> that gap was named directly to Rohan, not glossed over. If it recurs and needs
> closing for real, the fix is removing `Shell.Application`/`CopyHere` from the
> verification path entirely, not a better explanation of the current one.
>
> **MACHINE STATE AT HANDOVER, checked directly, not assumed:** PowerPoint process
> alive, ZERO presentations open (the deck really is gone). Excel process ALSO alive
> (not seen in the first process check — appeared between checks) with the LOCAL test
> `register-wide.xlsx` genuinely open (`PRESERVED-known-good-20260815-1050`), UNSAVED
> (file on disk still shows the old 15 Aug timestamp). This is real signal: "Set up my
> quarter" got far enough to open the paired workbook before something stopped it —
> consistent with the deck's own window closing along the way, whether from the
> OneDrive\Claude sighting or something else. Left exactly as found, both processes,
> for next session to inspect fresh rather than guess from a discarded state.
>
> **Next action:** don't assume the state above is still true by the time you read
> this — check both processes fresh first. If the Excel workbook is still open exactly
> as described, that is useful forensic evidence, not something to discard casually.
> Once state is understood, retry `1. Set up my quarter` on the SAME local test copy —
> watch closely for what exactly triggers any OneDrive/Explorer-looking window, and
> report it in detail rather than stopping again if it recurs (now that both file
> checks above confirm it's not writing anywhere unsafe). This is still the FIRST
> unaided close of Scenario 1 -- not yet achieved. `2. Draft and publish` and
> `3. Put it on the slides` haven't been reached yet either.
>
> **RE-CONFIRMED at final handover, unchanged from the check above:** PowerPoint 0
> presentations, Excel still has the same `register-wide.xlsx` open, `Saved=False`.
> Left alone deliberately, both times.
>
> **ONE THREAD RAISED, NEVER RESOLVED THIS SESSION — SAY SO PLAINLY, DON'T GUESS AT
> IT:** Rohan asked, referencing a screenshot, "why blocked on those fields? fix it
> forever more" — some report or dialog showed one or more FIELDS with a "blocked"
> status he wanted explained and permanently fixed. **No image ever actually reached
> this session** — every screenshot attempt this entire session failed to attach
> (confirmed repeatedly, not a one-off), so the specific fields and exact wording were
> never learned. Do not assume this is the same "OneDrive\Claude" sighting above —
> Rohan's replies suggest it might be, but he never confirmed that reading, and
> "blocked" on FIELDS (plural) sounds more like a Readiness/sync report line (this
> project's own established vocabulary — e.g. "Period: BLOCKED -- not set in the saved
> file") than a Windows folder window. **First thing next session: ask what specific
> fields and get the exact text, in words, before doing anything about it.**
>
> ## 16 AUG. **`CHECKLIST.md` IS NOW THE PRIMARY HANDOVER SURFACE. READ IT FIRST.**
>
> Compiled 2026-08-16 from a full pass of every CURRENT-status document
> (`DOCUMENT-MAP.md`'s own list), not just the usual three. Flat, tickable, every item
> linked to its source. This block and everything below it is still true and still worth
> reading for the *why* — but "what's actually left to do" now lives in `CHECKLIST.md`,
> and it gets ticked there, together, not re-derived from this prose each time.

> ## 15 AUG, ~21:20. **SESSION-END HANDOVER. STATUS: CURRENT.** Rohan is calling it for
> tonight. Supersedes the ~20:15 block below on build state and the milestone finding.
> Everything else there stands.
>
> ### FIRST ACTION: BUILD `addin104`. FOUR REAL FIXES ARE SOURCE-ONLY RIGHT NOW.
>
> `addin102` is currently loaded. `addin103` is built and sitting unticked. **Neither has
> the milestone fixes (Q, R) or the drafting-report label fix.** Build fresh, call it
> `addin104`, tick it, restart. Until that happens nothing below about milestones or the
> drafting report is visible on any real slide — this is not optional, it's the very first
> thing to do next session.
>
> ### TONIGHT'S SIX COMMITS, WHAT EACH ONE ACTUALLY DID
>
> 1. **`b0eab50` — Fix P.** Cloud-hosted decks no longer get bricked read-only.
>    Root cause: `PropertyOnDisk` took its path ByRef and reassigned it during URL
>    translation, so *reading* the file rewrote the caller's `path` variable from the
>    `https://` URL to a local one — silently disabling the cloud branch and pointing
>    `SaveAs` at the wrong location, which is what caused the read-only lock. Now `ByVal`.
>    Cloud persistence is still intermittent (uncharacterised, eight hypotheses eliminated,
>    documented in `FIX-LIST` P) — but the tool can no longer damage a deck trying.
> 2. **`769a280` — Scenario 1, mechanism proven, honestly not counted closed.** A real
>    quarter turn ran end to end for the first time — `Q4F26`→`Q1F27`, 43/43/43 rows,
>    verified from the saved files. **Not counted as closed**, because the pass condition
>    is unaided and Claude drove much of the run. Count stayed 5 of 9.
> 3. **`c509e3b` + `f110af6` — File-per-quarter, safe half, tested.** `ArchiveWorkbookForPeriod`
>    freezes the outgoing quarter to its own file before roll-forward — `SaveCopyAs`, never
>    `SaveAs` (which would silently re-point the live workbook at the archive). Three tests,
>    proven by a deliberate break. **Not a gate yet** — a missing archive is reported and the
>    run continues, since roll-forward only appends. The prune half (drop old rows, retire
>    `ParkSheetCopy`, sweep up `Sync Log` too — see `SCENARIOS.md`) is still not built.
> 4. **`d31869c` — Readiness now catches a partial quarter before it refuses you.** The
>    exact five-stub-row problem that cost an hour tonight now shows as `Rows for <period>:
>    PARTIAL — N row(s) where a full quarter is M` in the pre-flight, with the remedy
>    spelled out, before you ever press the button.
> 5. **`72bd4c1` — The 13-field chained report now labels each block.** `2. Put it on the
>    slides` runs Copy-AI-to-Submit and Publish across all 13 drafting fields in one press
>    (deliberate, not a bug) but every block used to land under the same two fixed headers
>    with the field name buried in prose. Rohan: *"this msg makes zero sense."* Now each
>    block reads `-- Publish Drafts (KEY_EVENTS_BODY) --`.
> 6. **`d621b1f` + `08f1da1` + `4b099be` + `cc4bd04` — the milestone device, fixed in two
>    real layers, chased at Rohan's direct insistence rather than demoed around.**
>    - **Q (`d621b1f`):** the device's two writers (`SetVisible`/`WriteText`) used to
>      suppress their only write and return nothing — no postcondition, no report. Now they
>      confirm by reading the property back, same as everywhere else this project has
>      learned to distrust a write that didn't raise. Proven by a deliberate break on a
>      real slide with real shapes.
>    - **R (`cc4bd04`), the bigger one:** Q's fixed writers had never once been exercised,
>      because **the device was structurally unreachable.** `PlanRoutineSync` walks the
>      register's own column headers as field identities; the device's identity tag
>      (`MILESTONE_TIMELINE`) is not a register column — the real data lives across 21
>      separate `MS1_LABEL`..`MS7_DONE` columns instead. So `InjectField` was never once
>      called with that tag, on any slide, ever. Reproduced live: seeded real test data,
>      ran the actual button macro against the real deck, confirmed nothing moved.
>      **Fixed properly**, not routed around: `InjectPrimitive.DeviceRoleTagsOnSlide` walks
>      a slide's shapes the same way `InjectorFor` already decides a tag routes to a
>      device, and `PlanRoutineSync` now asks about those tags too. Proven end-to-end with
>      a test whose row dictionary contains only column-style keys, never the device's own
>      tag name — matching the real register exactly — going from `no_change` to
>      `in_place_correction`, with real shapes confirmed changed afterward.
>    - **`08f1da1`, a side effect worth knowing:** `Application.Run` cannot marshal a VBA
>      `Type` (`MilestoneDrawResult`) back to an external caller — confirmed live, not
>      theorised. `DrawFromRowReport` is a permanent String-returning wrapper that exists
>      solely so this device can be inspected from outside VBA at all, same shape every
>      other verified-write function in this project already uses.
>
> ### WHAT WAS NOT DONE, AND WHY IT'S FINE
>
> None of tonight's milestone/report fixes have been **seen** working on a real slide —
> `addin104` was never built and ticked before the session ended. That's the very first
> action above. The **evidence they work is the tests**, each proven by a deliberate break
> tonight — stronger evidence than a one-off visual demo would have been, just not a
> picture. Say so plainly if picking this up cold: don't re-diagnose Q or R, build and go
> straight to trying it on the rig.
>
> ### THE RIG, AS LEFT
>
> `AppData\Local\deck-sync-quarter-20260815-1623\` — deck at `Q1F27`, register clean (no
> stub rows, no stale ticks), **209 real sync changes already applied to slides** (verified
> from the deck's mtime, which moved for the first time all session). `3_P001`'s Q1F27 row
> also carries seeded test milestone data (slots 1-2) from tonight's chase — harmless, real
> content, safe to leave or clear. **Do not use `OneDrive\deck-sync-known-good\`** — still
> points at the live register.
>
> ### YOUR LIVE REGISTER (OneDrive\Claude\register-wide.xlsx) IS ALSO CLEAN
>
> Both cleaned tonight, separately, each backed up first and verified by diff before/after:
> the 5 `Q1F27` stub rows and the 38 stale `Y` ticks on `Review project-status-2D3D`. Your
> real quarter turn will not hit either obstacle now. Backup:
> `AppData\Local\deck-sync-backups\register-wide.PRE-CLEAN-20260815-1927.xlsx`.
>
> ### THE CRITICAL PATH, UNCHANGED FROM THE ~20:15 BLOCK
>
> 1. Close scenario 1 for real — review the drafted content, it's a content decision now.
> 2. Re-run scenario 1 **unaided.** This is the only step that moves the count, and by
>    definition happens without Claude.
> 3. Land the file-per-quarter prune half (`Sync Log` included in scope now).
> 4. Scenario 3 (per-letter templates) — plan written, step 1 done.
> 5. ~~Fix the milestone device writers~~ — **done tonight, Q and R both.**
> 6. Prove scenario 8 — a fresh deck and employer from nothing.
> 7. Scenario 9 (provenance) — designed, not built.
>
> ---
>
> ### THE CRITICAL PATH TO A GENUINELY FINISHED TOOL, IN ORDER
>
> Not a task list — the actual dependency order, asked for directly by Rohan the same
> evening:
>
> 1. **Close scenario 1 for real.** Not a code blocker — 43 rolled-forward rows need
>    approval ticks before publish touches a slide. A content decision, not a coding one.
> 2. **Re-run scenario 1 unaided.** The pass condition is Rohan alone, no Claude. This is
>    the only step that actually moves the count, and by definition Claude cannot do it.
> 3. **Land the file-per-quarter prune half** (drop old rows, retire `ParkSheetCopy`,
>    **now also sweep up `Sync Log`** — see `SCENARIOS.md`'s file-per-quarter section,
>    updated 2026-08-15 evening). Archive half is built and tested; prune needs its own
>    tests and a keyboard run before touching anything real.
> 4. **Unstick scenario 3** (per-letter templates) — plan written, step 1 done.
> 5. **Fix the milestone device writers** (FIX-LIST item Q) before the milestone
>    read-back feature is built on writers that can't report failure — 21 of 29 fields
>    depend on this being solid first.
> 6. **Prove scenario 8** — a fresh deck and employer from nothing. The actual product
>    test; everything so far is validated against one real deck only.
> 7. **Scenario 9** (provenance) — designed, not built, and the easiest to defer forever.
>
> Steps 1–2 need no live-loop diagnosis and no Claude in the room — button presses and
> content review. Start there.
>
> ### THE DRAFTING-CHAIN REPORT WAS UNREADABLE, NOW FIXED — `72bd4c1`
>
> `2. Put it on the slides` runs 13 fields through Copy-AI-to-Submit and Publish in one
> press (deliberate — "a person is not asked thirteen times"). But every field's block
> landed under the same two fixed headers with the field name buried in prose or absent.
> Rohan, reading a real run: *"this msg makes zero sense"* — correctly, since two blocks
> reporting 0 rows and 38 rows with no visible label read as self-contradictory. Fixed:
> `DraftingUI.ChainBlockHeader` labels each block with the field it belongs to, proven by
> a deliberate break. **Not re-exercised against a live deck since the fix** — next
> `2. Put it on the slides` run should confirm the dialog now reads as distinct blocks.
>
> ### FIRST ACTION: PRESS PUBLISH ON THE RIG. IT IS ONE BUTTON FROM CLOSING SCENARIO 1.
>
> The rig is `AppData\Local\deck-sync-quarter-20260815-1623\` — a local copy of the REAL
> deck and register, already re-pointed at its own register (verified in the deck's bytes).
> `Q1F27` is rolled forward and sitting in it. **Nothing has been pushed to a slide:** the
> `Sync Log` has no entry for the 16:23–16:52 window. That publish is the gap between
> "proven in the register" and scenario 1 closed.
>
> ### WHAT ACTUALLY MOVED, AND WHAT DID NOT
>
> **Moved:** the quarter turn ran end to end on real files for the first time — period
> `Q4F26`→`Q1F27` confirmed in the saved `.pptx`, **43 rows / 43 distinct keys** in the
> register, old quarters intact at 43 each, 13 drafting sheets rebuilt and 13 parked.
> Content is real, not stubs.
>
> **Did not move: the scenario count is still 5 of 9.** Scenario 1's pass condition is
> *unaided*, and Claude found the blocker and the row number. **The mechanism is proven;
> the scenario is owed a run where Claude says nothing.** Do not let this session's chat
> transcript convince anyone it was closed — `SCENARIOS.md` is the count.
>
> ### THE THING MOST LIKELY TO STALL A REAL QUARTER AT WORK
>
> **Stale state, discovered one collision at a time.** Five leftover `Q1F27` stub rows
> silently refused the roll-forward tonight; **the same five are in the live register**,
> along with **38 `Y` ticks from 10 Aug**. There is no pre-flight. At work, that is a
> refusal you cannot interpret with no Claude to ask. Full table and two more defects from
> the run (`Roll Forward` asking a question it can answer itself; a bare Excel delete
> prompt) are in `SCENARIOS.md`.
>
> ### RANKING, ARGUED FROM THE CODE
>
> **File-per-quarter belongs ahead of scenario 3 and the template thread.** Verified not
> built (`ParkSheetCopy` copies within one workbook; the only `SaveCopyAs` is for deck
> backups). It CLOSES four open problems rather than adding one — see `SCENARIOS.md`.
>
> ### DO NOT
>
> - **Do not resume probing the intermittent OneDrive write.** Eight hypotheses are dead
>   and named in `FIX-LIST` P. A ninth without a new idea is the sunk-cost loop.
> - **"P is fixed" does NOT mean "OneDrive works."** Four setup properties, ~50% land rate.
> - **Do not approve anything** in `Review project-status-2D3D` — 38 stale ticks.
>
> ---

> ## 15 AUG, ~16:20. **HANDOVER. STATUS: CURRENT.** Supersedes the 12:45 block below on
> P and on the build. Everything else there stands.
>
> ### `addin101` IS LIVE. P's DESTRUCTIVE HALF IS FIXED AND PROVEN.
>
> Suite `203/0`, `COMPILE OK (34 modules)`, static clean. `addin101` registered
> `AutoLoad=1` and verified loaded; md5 `8897F5ED…`, distinct from `addin100`'s
> `DD1F5B24…`.
>
> **The root cause was NOT SaveAs — it was a reader mutating its caller.**
> `PropertyOnDisk` took `deckPath` ByRef and reassigned it during URL translation, so
> reading the file rewrote `path` from the URL to the local path. That made `IsUrl(path)`
> always False on a cloud deck, and handed `pres.SaveAs` a location the document was not
> open from — which is what left decks read-only. Now `ByVal`. Full evidence: `FIX-LIST`
> item **P**.
>
> **Old build: one failed attempt bricked the deck for the session. New build: healthy
> after every failure, five runs.** That is the win; take it as the win.
>
> ### WHAT IS STILL OPEN, AND DO NOT RE-DERIVE IT
>
> Cloud persistence is **intermittent** — identical writes to fresh cloud decks land about
> half the time. **Eight hypotheses tested and dead:** file size, AutoSave, sync latency,
> URL translation, fixture poisoning, aggressive polling, wrapper-vs-direct, dirty flag.
> Settle window raised to 30s in 5s steps. Anyone picking this up: read P before running a
> single probe, and do not spend an evening re-killing those eight.
>
> **This is NOT a non-cloud tool.** The affected surface is four setup-only document
> properties. Slide content wrote fine to a cloud deck in the same session, and the
> register is Excel on a different path entirely.
>
> ### FIRST ACTION: SCENARIO 1 — and confirm where the live deck actually lives
>
> Start a Quarter writes the period, which is the exposed property, so scenario 1 is the
> one most likely to meet this. **Before starting, confirm whether the live deck is
> OneDrive-hosted or local** — that decides whether any of P matters for the quarter turn,
> and it was never established.
>
> Scratch folders safe to delete: `OneDrive\Claude\onedrive-write-probe\` (several
> `fresh-*`, `poll-*`, `dirty-*` decks and `target-*.xlsx`) and
> `AppData\Local\deck-sync-probe\`.
>
> ---

> ## 15 AUG, ~12:45. **HANDOVER. STATUS: CURRENT.** Supersedes the ~12:00 block below on
> the build state and, completely, on the OneDrive finding. Everything else there stands.
>
> ### THE ADD-IN IS BUILT AND LIVE. `addin99`.
>
> Built from a clean tree at `95ea7dd`, stamp `2026-08-15 12:10`, all 33 modules imported.
> Saved to `OneDrive\Claude\addin99.ppam`, copied to the trusted `AppData\Roaming\
> Microsoft\AddIns\` and hash-verified across both (md5 `4D5D78F783137BEB2BE89D372A1D6141`,
> distinct from `addin98`'s `0A1C17FC…`, so it is not a re-save). Registered `AutoLoad=1`
> and the ONLY entry — `addin98` was removed from the list, not just unticked; its file is
> still on disk if it is ever needed back.
>
> **Verified loaded, not assumed:** `CodeLetterOf("3_P001")` returns `P`, and that function
> exists only in this build. A nonsense macro name fails on the same channel, so the
> success discriminates. Module list on disk and the build script's import list were
> diffed and are identical — no module silently missing, the defect that once shipped a
> `.ppam` that could not run Sync Now.
>
> ### THE ONEDRIVE RISK IS SOLVED AS A DIAGNOSIS. THE FIX IS NOT WRITTEN.
>
> **The recorded remedy was wrong and the risk was misattributed. OneDrive is not broken.**
> Plain `pres.Save` works on a cloud-hosted deck. What fails is `pres.SaveAs path, 24` —
> it raises `0x80CD1001` and leaves the open presentation READ-ONLY, after which every
> save fails with *"must be saved with a different name"* for the life of that document.
>
> `DeckRegistry` escalates to exactly that call at **three sites** — `SaveDeckVerified:807`,
> `SetDeckPeriodVerified:869`, `SetWorkbookPathVerified:940` — the instant a read-back does
> not confirm. A cloud save lands a beat after the call returns, so the read is too early,
> the escalation bricks the document, and the retry loop burns its remaining attempts
> against a presentation it has already broken. **The rescue is the failure.**
>
> Measured on a scratch deck, every phase read back from the saved file: plain Save **3 of
> 4 landed**; one SaveAs **raised**; plain Save after it **0 of 4**. Four theories died
> against measurement first — file size (a 32KB deck fails identically), AutoSave
> (settable, makes no difference either way), sync latency (two minutes plus a close,
> never arrived), and URL translation (`LocalPathForUrl` maps the `d.docs.live.net` URL to
> the local file correctly). **Full evidence: `FIX-LIST.md` item P.**
>
> **FIRST ACTION: WRITE P's FIX.** Branch on the existing `IsUrl(path)` at all three sites;
> on a cloud deck never escalate — settle and re-read instead. Two private helpers and two
> constants, and **the constants go at the TOP of the module** (a `Const` after a procedure
> is a VBA compile error). Then suite → build → `addin100`.
>
> **Operational, until it is fixed:** a run that hits this leaves the deck read-only, so
> anything done afterwards in the same session silently fails to save too. Close and
> reopen the deck to clear it.
>
> ### WHAT DID NOT CHANGE
>
> Delivery count is still **5 of 9**. Suite `203/0` and `COMPILE OK (34 modules)` still
> stand — no source was touched today, so the build matches the tested tree. Scenario 1 is
> still the last of "the quarter". `SCENARIOS.md` is still the frame; its OneDrive section
> has been corrected to match the above.
>
> Two scratch folders exist and are safe to delete: `OneDrive\Claude\onedrive-write-probe\`
> and `AppData\Local\deck-sync-probe\`.
>
> ---

> ## 15 AUG, ~12:00. **SUPERSEDED** on build state and on the OneDrive finding by the block
> above — its "biggest unlit risk / nothing is proven on OneDrive" reading is now known to
> be a defect in our own escalation, not a platform problem. Everything else here stands.
>
> ### FIRST ACTION: BUILD. THE ADD-IN IS BEHIND THE SOURCE.
>
> **Loaded add-in is `addin98`** (md5 `0a1c17fc4ab97fff4bb6586d9b7dd128`, in
> `OneDrive\Claude\` and the trusted `AppData\Roaming\Microsoft\AddIns\`). It has the
> Add/Retire split. **It does NOT have** the tooltip fix, the `Tag fields on this slide`
> button, or `CodeLetterOf`. Build → Save As `addin99` → tick → untick 98 → restart.
> Note `build_ppam.ps1` **QUITS POWERPOINT** first (fix-list O) and the run is not
> automatable past the Save As click — that is a real, permanent human step.
>
> ### BUILD STATE
>
> - **Suite `203 passed, 0 failed`. `COMPILE OK: whole project compiled clean (34 modules).`**
>   Static checks clean across 35 modules; all three module lists satisfied.
> - **Everything is committed and pushed** — last commit `e517a93`.
> - The suite REFUSES to run if PowerPoint or Excel has an open window, and **exits `2`
>   either way — the same code as a compile-gate failure.** Read the text, not the code.
>
> ### DELIVERY COUNT IS 5 OF 9 SCENARIOS. `SCENARIOS.md` IS THE FRAME — READ IT SECOND.
>
> Closed: **2, 4, 5, 6, 7.** Open: **1** (generate a new quarter — the last of "the
> quarter", never run end to end), **3** (plan below, step 1 done), **8** (untested, though
> the pairing button is a real piece of it), **9** (designed, not built).
>
> ### WHAT SHIPPED TODAY, AND THE ONE PATTERN BEHIND IT
>
> Three new buttons, all fixing the SAME defect shape — a tested function a person could
> not reach. That is now **six** instances in this project and it is the dominant failure
> mode, ahead of anything algorithmic:
> 1. **`Change which workbook this deck uses`** — `RepointWorkbookUI` was offered only when
>    the pairing was EMPTY, never when it was WRONG.
> 2. **`Add missing slides` / `Retire slides with no row`** — split from one button by
>    declared intent, per Rohan. Each reports the other's count without prompting.
> 3. **`Tag fields on this slide`** — `DiscoverFields` was reachable only inside the
>    `Not hasTypes` setup gate, so a configured deck could never tag another field. **This
>    is the one to use next: 32 fields are untagged on the real deck and this is what tags
>    them.**
>
> ### THE THREE OPEN QUESTIONS, IN THE ORDER THEY MATTER
>
> 1. **Does any of this work on OneDrive?** Everything today was proven on a LOCAL copy
>    because the OneDrive-hosted write failed outright — AutoSave ON, 4 verified attempts,
>    file mtime never moved. The work machine is OneDrive-hosted. **This is the biggest
>    unlit risk on the board and it is not a scenario.**
> 2. **Scenario 1** — the last of "the quarter".
> 3. **Scenario 3** — plan below; step 3 lands in the path that closed scenario 2, so
>    re-run S2 after touching it.
>
> ### DO NOT REPEAT THESE
>
> - **Do not test on `OneDrive\deck-sync-known-good\`** — its deck still points at the LIVE
>   register (verified, unchanged). Use `AppData\Local\deck-sync-backups\
>   PRESERVED-known-good-20260815-1050\` (deck + register together, local, re-pointed and
>   proven) or re-point a fresh copy with the new button.
> - **Do not try to exercise `DiscoverUI.BuildDiscoverySheet`'s "keep existing marks" fix on
>   a configured deck** — unreachable there by design. It belongs to scenario 8.
> - **Do not approve anything in the live review queue** — `PROJECT_PROGRESS` shows `80%`
>   against a proposed `0.8` (fix-list M) and would write the string `0.8` onto a slide.
>
> ---
>
> ### SCENARIOS 2 AND 7 ARE CLOSED ON REAL FILES. DELIVERY COUNT IS 5.
>
> Both by button, unaided, **read from the saved files**. S2: one register row added →
> `Add or retire slides` → `1 created, 0 failed`, and the saved `.pptx` went to **45 slide
> parts** with `ppt/tags/tag666.xml` carrying `SLIDE_TYPE=project-progress` +
> `INSTANCE_KEY=S999`. S7: same row removed → **slide 44 deleted**, back to 44 parts, **0
> tags carrying S999**. The retire warning named the slide by index and key before asking.
>
> ### THE THING THAT BLOCKED IT, AND THE BUTTON THAT FIXED IT
>
> **A COPIED DECK KEEPS POINTING AT THE ORIGINAL'S REGISTER.** `GetWorkbookPath:170`
> returns the stored path unchanged whenever that path EXISTS; the sibling lookup is a
> fallback for a MISSING file only. The known-good snapshot's deck stored
> `…\OneDrive\Claude\register-wide.xlsx` in its own `docProps/custom.xml` — so "test it on
> the snapshot" would have read and WRITTEN THE LIVE REGISTER, including the `Field
> Discovery` fixture. Caught by reading the file, before anything ran.
>
> `RepointWorkbookUI` already fixed this and was **unreachable**: offered only when the
> pairing was EMPTY, never when it was WRONG. Now a 5th button, `CAP_REPOINT_WORKBOOK`,
> and **it worked by button, verified from the deck's bytes** (mtime moved, stored path
> changed; the same check on the untouched copy still shows the old path, so it
> discriminates). Fifth reachability defect of the run; first caught before it cost
> anything.
>
> ### ENVIRONMENT — THE RECORDED ONEDRIVE REMEDY DID NOT HOLD
>
> `NEXT-SESSION` records "turning AutoSave ON made the write land". **AutoSave was ON and
> the write still did not land** on the OneDrive-hosted snapshot: `SetWorkbookPathVerified`
> refused after 4 attempts, and the file's mtime was **still 06:27** — never written at
> all, while PowerPoint reported `Saved=False` with the change sitting in its cache.
> **The whole sitting was then done on a LOCAL copy outside any sync folder
> (`AppData\Local\deck-sync-backups\PRESERVED-known-good-20260815-1050\`) and everything
> worked first time.** The work machine will be OneDrive-hosted, so this is not academic —
> but it is a separate problem from the code.
>
> ### NOT DONE, AND WHY
>
> - **The `DiscoverUI` change CANNOT RUN ON A CONFIGURED DECK — sixth reachability finding.**
>   `1. Set up my quarter` went Start a Quarter → Roll Forward → Refresh Drafting Sheets and
>   stopped; no discovery offer appeared. **Not a bug — a gate.** `RibbonUI.bas:1549`:
>   *"SETUP IS A PRECONDITION, NOT AN ACTIVITY — once ever per slide type, and only on a
>   deck that has none. A configured deck never sees this."* `DiscoverFields` is reached
>   only from `If setupAnswer = vbYes` (`:1599`), and `setupAnswer` exists only inside
>   `If Not hasTypes`. This deck has `project-progress` registered, so the question is never
>   asked. `DiscoverFields` is `BuildDiscoverySheet`'s only production caller — the split-out
>   was for tests.
>   **Corollary: the 9 marks the change protects cannot be wiped either**, because the code
>   that would wipe them cannot run here. **It belongs to SCENARIO 8** (bring up a fresh
>   deck and register from nothing) — that is the only path that reaches it. Do not try to
>   exercise it on a configured deck again.
>   Fixture recorded for whenever it is reachable: `Field Discovery`, header row 6, data
>   rows 7–65 = **59 rows**, exactly **9 marks** (`PROJECT_CODE`, `PROJECT_NAME`,
>   `PROJECT_PROGRESS`, `PROJECT_STATUS`, `STRATEGIC_ALIGNMENT_BODY`, `ABOUT_BODY`,
>   `PROBLEM_BODY`, `PROGRESS_BODY`, `KEY_EVENTS_BODY`). Pass condition: *"9 existing
>   mark(s) kept."*
> - **The two-button split and the caption fix are IN SOURCE ONLY.** Not built, not
>   pressed, and **the VBA suite has not been re-run** (Office was in use). Static checks
>   clean across 35 modules, all three module lists satisfied. **First action next session:
>   build → Save As `addin98` → tick → restart → run the suite.**
>
> ### WHY THE TIMELINES LOOK FUNKY — ANSWERED, AND IT IS NOT WIRING
>
> Rohan asked. Publish is wired: `InjectorFor` routes a tagged group with
> `MilestoneDevice.SlotCount > 0` to `INJECTOR_DEVICE` before anything else. **The data is
> not there and never has been.** Counted from the register, columns found by header name:
> **21 milestone columns × 43 `Q4F26` rows = 0 non-empty cells.** And **`MILESTONE_TIMELINE`
> has no register column at all** — same shape as `STRATEGIC_LINKAGES`. So every timeline
> on every slide is still hand-drawn, and the tool has never had one value to write.
>
> **CORRECTED SAME SESSION, after Rohan asked "I thought we harvested from the shapes?" —
> and he was right.** This was first written as "the harvest CANNOT fill them, milestone
> state lives in shape visibility (GAP 3)", taken from the handover without opening
> `MilestoneDevice.bas`. **It is not a capability gap, it is a missing function plus a
> policy refusal:**
> - `Harvest.ShapeIsNotHarvestableText:410` refuses anything whose injector is not
>   `INJECTOR_TEXT`. Devices are refused BY POLICY at the routing layer.
> - `MilestoneDevice`'s entire public surface is `ColumnFor`, `IsDoneWord`, `SlotCount`,
>   `DeviceIntegrity`, `DrawFromRow`, `DrawMilestones` — structural or write-direction.
>   `SetVisible`/`WriteText` are private writers. **Nothing returns values.**
> - **But the addressing is already built and in use:** `PartsOf`/`CollectNamed` enumerate
>   the group's parts by name, and `ColumnFor(i, part)` already maps slot + part → register
>   column. Labels and dates are plain text in named parts; `DONE` is `shp.Visible`.
>
> **So the work is a `RowFromDevice` mirroring `DrawFromRow` and sharing `ColumnFor`** (so
> read and write cannot disagree about columns, exactly as `InjectorFor` stopped harvest
> and publish disagreeing about injectors), plus letting the harvest route a device to it.
> Plus a `MILESTONE_TIMELINE` column. That is a feature, not a research problem — and it is
> **21 of the 29 `Given` fields**, so it is the largest remaining stopped data flow.
>
> ### SCENARIO 3 — TRACED TO THE BOTTOM. STEP 1 DONE, STEPS 2–5 PLANNED.
>
> **`FindTemplateFor` is NOT the blocker.** Slide creation never calls it. `SlideMembership`
> resolves the template via `DeckRegistry.LookupType`, which reads a custom document
> property `DeckSyncType:<slideType>` holding `slideId|worksheetName`. **One template per
> type, by construction** — the property can hold one slide ID. `MakeTemplateFrom:78` says
> so itself: *"a type with two templates has no defined behaviour — LookupType returns
> whichever SlideID was registered last, so which one gets cloned depends on click order."*
>
> **The deck, from its own tags:** exactly ONE slide tagged `is_template` (slide 44, the
> green/`P` one), all 44 slides `slide_type=project-progress`. Rohan's ranges: **1–11 green
> `P`, 12–26 orange `K`, 27–43 purple `S`** — matching the tag counts P=11, K=15, S=17.
> `K` and `S` have no template.
>
> **DONE (step 1):** `TemplateSlide.CodeLetterOf` + 10 assertions, proven by a deliberate
> break. Handles both key shapes; returns `""` for no-letter rather than guessing, and `""`
> means *"use the type's unlettered template"* — which is what keeps single-template decks
> working unchanged.
>
> **STEP 2 — per-letter registration.** A NEW property namespace
> `DeckSyncTemplate:<type>:<letter>` → slide ID, **alongside** `DeckSyncType:` rather than
> replacing it. `LookupType` gains an optional letter: lettered property first, existing
> unlettered registration as fallback. Backwards compatible by construction — a deck with
> no lettered properties behaves exactly as today.
>
> **STEP 3 — choose PER ROW, not per type.** The letter varies row by row *inside* one type,
> so the choice belongs where rows are iterated: **inside `RunSync.CreateMissingSlides`**,
> not in `SlideMembershipCore` (which resolves one template per type and passes it in).
> **This is the code path that closed S2 — re-run scenario 2 after touching it.**
>
> **STEP 4 — the guard at `RibbonUI.bas:2375`** refuses to create a second template for a
> type (*"a type must have exactly one"*), and `MakeTemplateFrom:82` has its own. **Both
> must become one-per-type-per-LETTER or the feature cannot be built at all.** Do this
> before any deck surgery, or making the orange template will simply be refused.
>
> **STEP 5 — the deck surgery.** Make the `K` template from a slide in 12–26 and the `S`
> template from one in 27–43, via the existing tested `MakeTemplateFrom`, then tag each with
> its `code_letter` and register it. Name/tag the shapes. **On a copy — never the live deck.**
>
> **Still undecided:** whether a template's letter is a slide tag (`code_letter`, matching
> the existing lowercase `slide_type`/`instance_key`/`is_template` vocabulary) or lives only
> in the registration property. A tag is repairable by eye in the Selection Pane; the
> property is invisible. Rohan's standing preference has been repairability.
>
> ### ALSO SEEN, NOT YET ACTED ON
>
> - **The `START HERE` panel is a stale snapshot presented as fact.** It showed
>   `Paired workbook … OneDrive\Claude\register-wide.xlsx` with state **`ok`**, "read from
>   saved .pptx", four hours after that stopped being true. Row 2 says it is not live; the
>   row says `ok`. Anyone re-pointing a deck and then reading this panel concludes the
>   re-point failed.
> - **"17 value(s) are not in their allowed list"** from the chain — the Controlled-field
>   vocabulary problem, still unenforced.
> - `Readiness.bas:55` and `WorkbookBridge.bas:17` are **two separate constants both
>   hardcoding `"START HERE"`**. Change either and they diverge silently.
> - The tab is named `START HERE` while its own A1 heading says `DECK SYNC -- WHERE YOU
>   ARE`. Cost a minute tonight looking for a tab that does not exist.
> - Toolbar is now **"Deck Sync 40"** — forty accumulated bars — beside 25 stale `.ppam`
>   files. Both still worth a purge.
>
> Four new fix-list entries were added tonight: **L** (pairing, detection half still open),
> **M** (`PROJECT_PROGRESS` `80%` vs `0.8` false diff, live in the 47-row queue), **N**
> (invisible-character diffs with the explaining column cut off), **O** (`build_ppam.ps1`
> quits PowerPoint with the user's deck open; its safety is a timeout side effect).
>
> ---

> ## 15 AUG, ~10:45. **SUPERSEDED** by the block above. The 08:30 block below remains CURRENT for state; this
> only records a scope ruling and three costs. Nothing was built.
>
> ### ROHAN RULED: FINISH THE QUARTER BEFORE THE TEMPLATE LAYER
>
> chat proposed a shape-name join key (`=FIELD[#INDEX][.PART]`) replacing role tags, and
> flagged possible drift toward the more interesting problem. **The ruling is
> quarter-first**, made explicitly rather than by which document arrived next. Template
> thread parked, not rejected. Full trace: `claude-brain/DECISIONS.md` 2026-08-15;
> proposal at `OneDrive\Claude\chat-to-code-2026-08-15-template-names.md`, reply at
> `code-to-chat-2026-08-15-scope-ruled-quarter-first.md`.
>
> **So the FIRST ACTION is unchanged — the snapshot sitting in the block below.**
>
> ### THREE COSTS PRICED, ALL UNPAID, ALL ON THE NAMING THREAD
>
> 1. **It reverses a July decision.** `TemplateAudit.bas:350` cites `specs/identity-tags.md`:
>    the shape name was rejected as an identity key for being **visible and user-editable**,
>    per `input-contract.md`'s `unique_named_shapes` rule — which PowerPoint does not
>    enforce. The duplicate check must ship in the SAME commit as the first name read.
> 2. **`=` collides with Excel.** `WriteAuditGrid` already writes a leading apostrophe on the
>    TEXT column because slide text starting with `=` renders `#NAME?` — and writes
>    `COL_SHAPE` with no guard. The audit grid breaks on day one under this convention.
> 3. **`TemplateAudit.bas:377` stays open.** Verified: `ws.Cells(r, COL_DECISION).Value = ""`,
>    unconditional, so removing the `Cells.Clear` above it fixes nothing. The fix needs a
>    stable per-row key — the same key the naming thread would supply, and the one Discovery
>    already solves with `shp.Id`. Design call, deliberately not done now.
>
> ### ONE TEST STILL OWED, AND IT GATES THE WHOLE THREAD
>
> **Does `=` survive what PowerPoint does to shape names** — copy/paste duplication,
> group/ungroup, Reset Slide? Ten minutes on a SCRATCH deck, not the real one. If PowerPoint
> mangles or auto-renames on any of them, the proposal needs a different sigil or a different
> mechanism, and everything above is moot.
>
> **A correction worth keeping, because it is the third of its kind:** chat's argument rested
> on S5 having never run. **S5 closed 14 Aug 20:45** (`3_P001`/`KEY_EVENTS_BODY`, by button).
> It was reading a stale handover — the same stale line sat in Claude Code's session reminder
> until this morning. Its instinct was right and its evidence was wrong.
>
> ---

> ## 15 AUG, ~08:30. **STATUS: CURRENT.** Everything below is historical.
>
> ### SCENARIO 6 IS CLOSED ON REAL SLIDES. DELIVERY COUNT IS 3.
>
> **Read `SCENARIOS.md` first** — status is re-derived from the code there, and scenario 6
> now carries its evidence.
>
> ### WHAT WAS DELIVERED
>
> **`PROJECT_STATUS` corrected on 8 real slides, by button, unaided.** Review -> one `Y`
> in `F38` -> `2. Put it on the slides` -> Yes. One tick covered eight changes via
> `PropagateBatchApprovals`. **Verified from the saved `.pptx` against the pre-write
> backup**, never from a dialog: before 8 slides `'Not started'` + 1 already `'Not
> Started'`; after 0 + 9. The untouched ninth is why "8 written" was honest. The check was
> demonstrated to discriminate (different answers on the two files). Backup confirmed on
> disk at `AppData\Local\deck-sync-backups\...r13-20260815-074533.bak.pptx`, written
> outside the synced folder.
>
> This also settled, by button rather than by decision, the `PROJECT_STATUS` casing
> question `COLUMNS.md` lists as open — the register's vocabulary won. Reversible from
> that backup if the deck should have won instead.
>
> ### BUILD STATE
>
> - **`addin96`, stamp `2026-08-15 08:14`, is the loaded add-in.** `addin95` (07:04) was
>   ticked off. Copies in `OneDrive\Claude\` and the trusted `AppData\Roaming\Microsoft\
>   AddIns\`, md5 `15e044d096cd05b7957b1ff5667dda46`, hash-verified across both.
> - **Suite 202/0. `COMPILE OK: whole project compiled clean (34 modules).`** Static
>   checks clean across 35 modules; all three module lists satisfied.
> - **24 stale `.ppam` files** now clutter the AddIns folder. Still worth a purge.
>
> ### THE DEFECT THAT ATE THE EVENING, AND ITS FIX (PROVEN)
>
> **A review grid rebuilt under a live Excel AutoFilter did not clear.** On the real
> register: 108 rows where 57 were written, 21 rows carrying a change id and nothing else,
> 26 change ids duplicated, and **13 rows pre-ticked `Y` that no human typed** — including
> a second copy of the batch (`B1` unapproved, `B2` approved). Because approval applies by
> CHANGE ID and both copies share one, a single stale tick approves its invisible twin.
> `BuildQueue` and `WriteQueueSheet` were both innocent; reading them found nothing.
>
> **Fixed in `ReviewQueue.WriteQueueSheet`:** drop `AutoFilterMode` and unhide rows before
> the clear, then assert `COL_HASH` is empty on the row below the grid and raise if not.
> **Proven by before/after on the same input condition** — filter on, rebuild: was 108
> rows / 21 orphans / 26 dupes / 13 phantom ticks / filter kept / 60 hidden; now **49 rows
> / 0 / 0 / 0 / filter gone / 0 hidden**, banner `08:20:29`, reconciling independently with
> the dialog's "47 changes".
>
> **The `Err.Raise` backstop has never fired and is NOT tested.** It guards a condition the
> filter drop now prevents, so it cannot be exercised from the UI without breaking
> `Cells.Clear` itself. The filter drop is what does the work.
>
> ### SHIPPED BUT COMPLETELY UNTESTED — DO THIS BEFORE TRUSTING IT
>
> **`DiscoverUI.BuildDiscoverySheet` no longer clears the sheet.** It wiped a grid whose
> own row 3 tells a person to type in columns F and G, contradicting the tool's own "the
> grid is still there if you want to come back to it". Rows are now matched by shape id and
> updated in place, new shapes append, and only rows for departed shapes are cleared. Two
> new private helpers, `ExistingRowsById` and `CountMarks`; the result message reports marks
> kept and rows removed.
>
> **It has never run.** And it is NOT reachable by a button: `DiscoverFields` is called only
> from inside `1. Set up my quarter` after answering **Yes**, immediately followed by
> `BatchOnboardFlow.BatchOnboardType`, on a path that also runs `StartQuarter`. So testing
> it means running setup — do it on the snapshot, not the live pair.
>
> **The fixture is already perfect and is sitting in the live register:** the `Field
> Discovery` sheet holds **59 rows, all 59 shape ids matching slide 44, and 9 already
> marked**. Old code destroys those 9; new code must report *"9 existing mark(s) kept."*
> Register backed up first at `OneDrive\Claude\backups\PREDISCOVER-20260815-0826.xlsx`.
>
> ### ROHAN'S RULING, CAPTURED — READ BEFORE TOUCHING ANY `Clear`
>
> **The archive is last quarter's FILE. `REPORTED LAST TIME` is not storage.** His words:
> *"I can go back to previous quarters non destroyed drafting sheets to see field progeny,
> it is only in the new quarter for some of the ai tools and human to use its structure and
> narrative consistency."* Recorded as `DOCUMENT-MAP.md` decision **6** and
> `claude-brain/DECISIONS.md` 2026-08-15, plus a memory file. **Kills as a category:**
> ferrying more columns "so nothing is lost", widening `REPORTED LAST TIME` into a history,
> or adding a park/refusal around the rollover clear.
>
> **The live gap it exposes:** file-per-quarter is DECIDED, NOT BUILT — the register still
> stacks `Q3F26`/`Q4F26`/`Q1F27` in one workbook, so the rollover clear fires in the only
> copy. **Until it is built, `ParkSheetCopy` is load-bearing and must not be deleted.** A
> proposal to delete it as vestigial was made this session and was wrong.
>
> ### EVERY `Clear` IN THE CODEBASE, AUDITED — "why is clear needed there?"
>
> | site | verdict |
> |---|---|
> | `ReviewQueue.bas` (grid) | **needed** — the tail must not survive. Was failing under a filter; fixed. |
> | `DiscoverUI.bas` (discovery grid) | **NOT needed** — tail problem wiping human columns. Removed. |
> | `TemplateAudit.bas:355` | needed, but **it is not what discards your decisions** — line 377 blanks `COL_DECISION` unconditionally. The comment misattributes its own loss. Its "no stable per-row key" claim is true of the grid as built; Discovery solves the same problem with a shape-id column from `shp.Id`. **Open, and a design call, not a fix.** |
> | `Readiness.bas:426`, `WorkbookBridge.bas:416`/`466` | needed — `Cells.Clear` also drops stale FORMATTING, and these sheets change shape between builds. Nothing human on them. |
> | `Drafting.bas:751` | needed — fires only on an UNKNOWN layout, the one case where old cells cannot be located. |
> | `Drafting.bas:349`, `:791` | needed — a move, and the tool's own instruction rows. |
> | `Drafting.bas:970–975` (ferry) | needed — justified by the ruling above. |
>
> 24 `Err.Clear` calls not audited — they reset VBA's error object, not data. They can mask
> errors inside `On Error Resume Next`; a separate question.
>
> ### STILL OPEN
>
> - **`Review project-status-2D3D` is `OPEN` with 38 rows ticked `Y` from `2026-08-10
>   10:22:28`.** Five-day-old approvals sitting live. The change-hash design should stop
>   them applying to values that have since moved — **untested claim.**
> - **`TemplateAudit` line 377**, above.
> - **The apply dialog is titled "1. Set up my quarter -- slide changes"** and is reached
>   from "2. Put it on the slides" — a caption hardcoded where it should derive from `CAP_*`.
> - Scenario 1 (generate a new quarter) — last of "the quarter" neither closed nor built.
> - Scenario 3 blocked by `TemplateSlide.FindTemplateFor` returning the first type match.
> - Scenarios 2 and 7 built and unrun.
>
> ### FIRST ACTION NEXT SESSION
>
> **One sitting on the known-good snapshot (`OneDrive\deck-sync-known-good\
> 2026-08-15-0625\`), doing three things that all need it open anyway:**
> 1. **Scenario 2** — add one register row with a new instance key at `Q4F26`, press `Add
>    or retire slides`, read the result from the SAVED file.
> 2. **Scenario 7** — same button, the retire half.
> 3. **The `DiscoverUI` test** — `1. Set up my quarter`, answer Yes, confirm *"N existing
>    mark(s) kept."*
>
> Read every result from the saved file, never a dialog. Excel's AutoSave writes DURING a
> macro run, so a read taken too early returns a half-written sheet — that cost two wrong
> conclusions this session. Wait for the mtime to settle, then read twice and compare.
>
> ---

> ## 15 AUG, ~06:45. **SUPERSEDED** by the block above.
>
> ### THE HARVEST IS FINISHED. DECK MEMBERSHIP IS BUILT AND HAS NEVER RUN.
>
> **Read `SCENARIOS.md` first** — it is the frame, with status re-derived from the
> code, and it now says which scenarios are closed, unblocked, built-and-unrun, and
> genuinely blocked.
>
> ### FIRST ACTION: BUILD THE ADD-IN, THEN RUN SCENARIO 2 ON ONE PROJECT
>
> The loaded build is `addin93` (05:54), which predates the `Add or retire slides`
> button entirely. Build -> `addin94` -> tick -> restart.
>
> Then, against the known-good snapshot: **add ONE row to the register** with a new
> instance key at `Q4F26`, press `Add or retire slides`, and read the result from the
> saved file. That is scenario 2, and it has never run.
>
> ### WHAT SHIPPED — suite state recorded below, NOT yet re-run at handover
>
> - **`Add or retire slides`** (4th button). Creates a slide for every register row
>   that has none; DELETES every slide whose key the register has no row for. **Asked
>   separately**, because one "make the deck match" confirmation would buy consent for
>   the destructive half using the safe half's reasoning. The delete warning names
>   every slide by index and key.
> - **Retirement deletes.** Rohan's call: the register is the source of truth and last
>   quarter's saved deck is the archive, so hiding would grow the deck forever to avoid
>   a loss already covered.
> - **`ResolveDeckContext`** — the four guards that must pass before anything reads the
>   register, extracted from `ReviewChangesCore` unchanged: workbook path, open, the
>   UNSAVED-BUFFER refusal, and R9 duplicate keys. Creation and retirement now cannot
>   skip them.
> - **`SlidesWithNoRow`** returns slide OBJECTS, not keys, and guards two states the
>   count does not imply: never the template (it carries a type and no instance key by
>   design, so any rule keying off "has a type" deletes it), never an unclassified
>   slide.
>
> ### THE FINDING, AND IT IS THE FOURTH OF ITS KIND
>
> `RunSync.CreateMissingSlides` has existed and been tested all along. It was reachable
> only through `SyncNowCore`, called only from a `Private SyncNow` that **nothing
> called**. Its comment read *"NO LONGER A BUTTON TARGET. The chain is the entry
> point"* — the chain calls `StartQuarter`, `RollForwardUI`, `RefreshDraftingSheets`,
> marking and discovery, and never this. And it had been made `Private` **specifically
> so the reachability check would not report it**.
>
> So the checker built to find orphans was silenced on a genuine orphan, by a comment
> asserting a reachability fact that had stopped being true. `FIX-LIST` item D predicted
> exactly this class the night before. **Scenario 2 was blocked for weeks by a stale
> comment, not by missing code.**
>
> ### WHAT IS LEFT
>
> - **Scenario 1 (generate a new quarter)** is the last of "the quarter" neither closed
>   nor built. The ferry has still never touched the real workbook.
> - **Scenario 3 is blocked by a real defect**: `TemplateSlide.FindTemplateFor` returns
>   the first slide matching a type, and three colour templates share `project-progress`.
> - The milestone device (21 of the 29 `Given` fields) cannot be read back — slot state
>   lives in shape visibility.
> - `SECTOR`/`TRL` have no shapes; `STRATEGIC_LINKAGES` has no template tag; 12 slides
>   need dates by hand.
>
> ---

> ## 15 AUG, ~00:35. **SUPERSEDED** by the block above.
>
> ### THE HARVEST WORKS BY BUTTON. IT IS NOT SAFE TO RUN AT SCALE YET.
>
> ### ~~FIRST ACTION: TEACH THE HARVEST ABOUT FIELD KIND~~ — **DONE 2026-08-15, suite 200/0**
>
> `Harvest.HarvestSlide` now asks `InjectPrimitive.InjectorFor(sld, field)` — the
> SAME decision `InjectField` routes on, **extracted from it rather than copied**,
> so harvest and publish cannot disagree about a field — and refuses anything that
> is not `INJECTOR_TEXT`. A group is refused separately: the router calls a group a
> device only when it carries milestone slots, so a tagged group without them would
> otherwise route to the text writer and be reported as "blank on the slide", which
> is false. **Refusal, not conversion**, because a wrong guess cannot be corrected.
>
> It needed a FOURTH axis word. `Kind` (Controlled/Prose/Static/Derived),
> `FieldType` (text/number/currency/date) and `Behaviour` (fill/fit/as-is) are all
> taken and all mean something else; overloading one would be the defect that has
> already cost this project a wiped feature.
>
> **NOT IN A BUILT ADD-IN.** The loaded build is `addin91` (`00:14`), which predates
> this, the dialog fix and the doc rename. **First action now: build, tick, restart.**
>
> ### THE ORIGINAL FIRST ACTION, KEPT FOR THE REASONING:
>
> `PROJECT_PROGRESS` reads `33%` off the slide. `InjectPrimitive.bas:340` refuses
> any non-numeric progress value and names **`'90%'` as wrong in those exact
> words** — so harvesting what is displayed writes a value the tool itself cannot
> publish. **And the empty-cell rule means a corrected harvest can never overwrite
> it**: the cell is no longer empty, so it must be cleared by hand, exactly like
> the four coerced cells cleared at 23:40.
>
> Caught at the dialog on the last press of the night. **Nothing was written.**
>
> The harvest currently assumes "what is displayed" is what the register wants.
> True for prose, names and dates; FALSE wherever the slide shows a formatted
> VIEW of a stored value. It already refuses devices by name; it needs the same
> treatment for numeric-contract fields — convert (`33%` -> `0.33`) or refuse,
> never write the string.
>
> ### THE HARVEST IS DONE FOR EVERY FIELD THAT CAN BE READ — 15 AUG 06:25
>
> | field | filled |
> |---|---|
> | `STRATEGIC_ALIGNMENT_BODY`, `PROBLEM_BODY`, `SUBTITLE_A` | **43 of 43** |
> | `INDUSTRY_CASH`, `TOTAL_VALUE`, `PROJECT_LEAD` | **43 of 43** |
> | `START_DATE`, `END_DATE` | **31 of 43** — 12 slides display no dates |
>
> This morning the first column read **1 of 43**. Deck **440 -> 682 tag parts**.
> Roughly **320 values** now in the register that existed only as slide text.
>
> **Five batches, by button. Every one: offer matched result exactly** (24/34,
> 22/34, 60/90, 56/85, 66/102) and the register's own arithmetic agreed
> independently of the dialog. Every refusal traceable to a named slide.
>
> **KNOWN-GOOD COPY** of the deck and register at this state:
> `OneDrive\deck-sync-known-good\2026-08-15-0625\`, md5-verified. Restore point
> if anything after this breaks.
>
> **WHAT IS LEFT NEEDS CODE OR A PERSON, NOT MORE PRESSING:**
> the milestone device (21 of the 29 `Given` fields, unreadable — slot state
> lives in shape visibility, GAP 3); `SECTOR`/`TRL` (no shapes);
> `STRATEGIC_LINKAGES` (no template tag); 12 slides of dates by hand.
>
> ### HARVEST PROGRESS — 15 AUG 06:02
>
> **9 of 43 register rows filled** for the six scalar fields (was 1 at 00:35). Deck at
> **500 tag parts**, the seven roles on 10 slides. Two batches run by button:
> slides 2-5 (24 labelled, 34 values) and slides 6-9 (22 labelled, 34 values).
>
> The remaining ~34 slides are the SAME PRESS with a bigger selection. The preview now
> tells the truth, so the offer can be read and trusted before approving.
>
> **Known, from the runs rather than from theory:** slide 8's `START_DATE`/`END_DATE`
> collide and are refused — it has no date-shaped text, so both roles land on some other
> shape and the guard stops them. It needs its dates by hand. **One collision prints one
> line per role rather than one line per collision**, so a colliding slide looks twice as
> bad as it is. Cosmetic, unfixed.
>
> ### PROVEN ON REAL FILES, BY BUTTON, READ BACK FROM SAVED BYTES
>
> | | |
> |---|---|
> | harvest direction exists at all | `Harvest.bas` — new tonight |
> | verbatim storage | `30 Oct 2023`, `$275,598` in the register as text, not `45229`/`275598` |
> | template tagged | 7 roles on slide 44; tag parts 440 -> 447 |
> | propagation | 7 roles carried to slide 1; 447 -> 454 |
> | values harvested | 8 into `3_P001` Q4F26; 4 repaired after the coercion fix |
> | dates disambiguated | 55 shapes renamed; **32 of 44** slides now carry both `Text 212a`/`Text 216a` |
> | batch dry run, slides 2-5 | **24 labelled, 0 collisions**, dates resolving separately |
>
> ### BUILD STATE
>
> - Suite **199 passed / 0 failed**, compile clean, 34 modules. Static + module
>   lists + docs all clean.
> - **`addin91` is the only registered add-in**, `AutoLoad=1`, build stamp
>   `2026-08-15 00:14`. Verified loaded by calling `CountShapesWithRoleTag`,
>   which exists only in tonight's builds.
> - Deck `00:29`, 454 tag parts. Register `23:46`.
> - Backups, each md5-verified: `PRE-TAG-20260814-222900`,
>   `PRE-HARVEST-20260814-232900`, `PRE-BLANK-20260814-234000`,
>   `PRE-RENAME-20260815-002300`. `backups/` is now gitignored (48MB, was one
>   `git add -A` from being welded into history).
>
> ### FOUR DEFECTS FIXED, AND WHAT FOUND EACH
>
> 1. **`AdoptFlow.AdoptExistingSlides` was orphaned** — no button, no caller, for
>    the whole life of the three-button toolbar, its header still reading
>    "Toolbar entry point". `check_vba_static.py` could not see it because
>    `AdoptFlow.bas` was missing from `UI_MODULES`. **The checker asks whether a
>    NAME appears in another module, not whether anything reachable calls it** —
>    a chain of private orphans is still invisible to it.
> 2. **Excel silently coerced harvested values.** `.Value = CStr(...)` is not a
>    string write. `UpsertRow` gained opt-in `asText` setting `NumberFormat = "@"`
>    BEFORE assignment. Found by reading the register file after the first real
>    harvest.
> 3. **`SaveDeckVerified` reported failure when there was nothing to save** — its
>    only proof was "did the mtime advance", which cannot separate "save failed"
>    from "nothing pending", and it forced a full `SaveAs` of a 49MB deck to find
>    out. Now consults `pres.Saved` first (to ask whether anything was PENDING —
>    the file still has the only word on success).
> 4. **The matcher was name-blind.** Name now breaks a sibling tie, after
>    geometry, never before it.
>
> ### THREE OPEN, EACH CAUGHT BY A GUARD RATHER THAN A TEST
>
> 1. **The progress-format mismatch above.** The blocker.
> 2. **`OfferHarvestForSelectedSlides`'s dialog mislabels and truncates.** All
>    propagation detail is accumulated into the `collisions` string, so successful
>    stamps print under a "Refused -- two fields matched one shape" header; and it
>    hits `CapReport`'s 900-char cap mid-word, so collisions can be invisible.
> 3. **Slide 27 has a shape already named `Text 216a` that is not the date.** The
>    rename refused it. That slide's `END_DATE` will keep colliding.
>
> ### TWO THINGS INTRODUCED TONIGHT THAT ARE NOT DEFECTS BUT WILL BITE LATER
>
> - **`START_DATE`/`END_DATE` are now stored as TEXT** (`30 Oct 2023`), deliberately,
>   so they round-trip to slides verbatim. **Nothing in the VBA reads those columns
>   today** — checked, the only references are comments. But `COLUMNS.md:46` says
>   *"Time elapsed (the horizontal bar and its percentage) is computed from
>   `START_DATE` and `END_DATE`"*, so that derivation is planned and not built. When
>   it is built it must PARSE the text rather than assume a serial. This is the same
>   tension as open item 1: the register stores what the slide shows, and a computed
>   field wants the underlying value.
> - **`SOURCE-HARVEST.md` and `Harvest.bas` are unrelated and both say "harvest".**
>   The doc is a paper provenance form from 2026-08-08; the module reads values off
>   slides into the register. Both declare themselves current. A word doing two jobs
>   has cost this project before — rename one before it costs again.
>
> ### WHAT I GOT WRONG TWICE — READ THIS BEFORE THEORISING ABOUT THE MATCHER
>
> - **Theory 1: `POSITION_TOLERANCE_EMU` (1 inch) could not separate shapes 0.14"
>   apart.** Wrong — arithmetic done properly gives a geometry gap of 0.075, and
>   the two DO tie, but that was not why the fix failed.
> - **Theory 2: a name tier above scoring.** Wrong, and the codebase said so:
>   `Discovery.bas:22` reads *"shape name ... never used as an identity key"*. It
>   turned a deliberately-drifted shape into a `high` auto-accept and broke
>   `Onboarding_HighAndMediumConfidence` and
>   `DeckAdoption_MediumConfidenceSlideNeedsConfirmation`. **Two tests were
>   standing exactly where that rule needed guarding.** Name as a TIE-BREAK
>   passes both.
> - **The real cause was the deck, not the code:** only **4 of 44** slides carried
>   the date shape names. The tie-break worked on those and nowhere else. 20 =
>   6+6+4+4 was the tell, and it was in the counts the whole time.
> - **COORDINATE SPACES.** A shape inside a group is stored group-relative in the
>   slide XML; COM's `.Left`/`.Top` are absolute. The first rename dry run matched
>   **0 of 88** because of this. The dates are at `y=3.16"` in the XML and
>   `top=120.8pt` (1.68") via COM.
>
> ### DELIVERY COUNT IS 2 — AND THE COUNT HAS STOPPED MEASURING THE BLOCKED HALF
>
> Rohan, tonight: *"I've pressed that button several times in testing what more
> can I do"*. He is right, and it is the second time he has said it (14 Aug: *"we
> have proved we can push to slides over and over"*). **Publishing is proven.
> Holding the count at 2 scores the half that works while the blocked half goes
> unscored.** The blocked half is INPUT — the register has little worth
> publishing:
>
> | | |
> |---|---|
> | 6 scalar fields | harvested on **1 of 43** slides |
> | 21 milestone fields | cannot be read back at all — slot state lives in shape visibility (GAP 3) |
> | `SECTOR`, `TRL` | no shape; Rohan chose to split them out, not built |
> | `STRATEGIC_LINKAGES` | register column, no template tag, no values |
> | `PROJECT_STATUS` | 17 values out of vocabulary, nothing enforcing it |
>
> Score that checklist, not the delivery count.
>
> ---

> ## 14 AUG, ~23:10. **SUPERSEDED** by the block above.
>
> ### THE HARVEST DID NOT EXIST. IT DOES NOW, AND IT HAS NEVER RUN ON THE DECK.
>
> The block below says the harvest is a press away. **It is not, and it never
> was.** `DeckAdoption.PlanAdoption` skips any slide carrying `slide_type` +
> `instance_key` at `DeckAdoption.bas:149`, **before any matching runs**. All 43
> project slides carry both. So adoption on a live deck reports 43 already-linked
> and reads nothing — by construction, not by accident. Adoption is for slides
> the tool has never seen; every slide here was adopted months ago, which is how
> the nine prose roles reached all 44.
>
> Searched for any other route: `PlanAdoption`/`CommitAdoption` are the only two
> functions that read a value off a slide into the register. The `CurrentValue`
> reads in `InjectPrimitive` are what-was-there-before records inside the
> INJECTION path, not a harvest.
>
> ### FIRST ACTION: REBUILD THE ADD-IN, THEN DRY-RUN PROPAGATION ON **ONE** SLIDE
>
> Nothing below is reachable until the modules are re-imported and
> `File > Save As > PowerPoint Add-in` is clicked. The loaded add-in is still the
> pre-harvest build.
>
> Then: select ONE project slide, press `1. Set up my quarter`, and answer **No**
> at the prompt. The dry run writes nothing. **Read what it offers, not what it
> claims** — and note the suite cannot substitute for this: every test fixture is
> a two-shape slide, and slide 1 has **136 shapes, four of them named
> `Shape 46`**. The matcher has never been asked anything that hard.
>
> ### THE TAGGING WAS NOT 29 FIELDS. IT WAS 7, AND IT IS DONE.
>
> Read from the deck's XML, not from the field list:
>
> | | |
> |---|---|
> | 21 `MS*_LABEL/_DATE/_DONE` | ONE tag. `Discovery.bas:165` recognises the timeline by counting slots, not by name, so the device is one candidate |
> | 6 scalars | real shapes, tagged |
> | `SECTOR`, `TRL` | **no shape at all** — substrings of one composed subtitle line |
>
> **Written to slide 44 and verified from the saved bytes**: tag parts 440 → 447,
> each of the seven ROLE values appearing exactly once, file moved 19:20 → 22:33.
> `START_DATE`, `END_DATE`, `PROJECT_LEAD`, `SUBTITLE_A`, `INDUSTRY_CASH`,
> `TOTAL_VALUE`, `MILESTONE_TIMELINE`. Backup: `backups/PRE-TAG-20260814-222900/`,
> md5-verified.
>
> ### WHAT THE DECK ACTUALLY LOOKS LIKE — settled from the file, do not re-derive
>
> - **Slide 44 is the template** (the one tag file with `SLIDE_TYPE` and no
>   `INSTANCE_KEY`), and it is a **clone of slide 1 with only the nine prose
>   fields placeholder-ised**. Every Given value on it is still project
>   `3_P001`'s real data, including a shape named `3_P001 Timeline`.
> - **The milestone device is named on 2 of 44 slides** — 44 and 1. Nowhere else.
> - **There is a stray `MS2_ON` OUTSIDE the timeline group**, at slot 3's
>   position, on both slides. `PartsOf` only walks the group, so the device can
>   never control it. Uncontrolled dark circle over slot 3.
> - **The slide shows four money figures and the register names two.** SAAFE Cash
>   and In-Kind are labelled on the slide and have no field.
>
> ### WHAT SHIPPED — suite **197 passed / 0 failed**, compile clean, 34 modules
>
> - **`vba/Harvest.bas` (new).** `HarvestSlide` reads role-tagged shapes into the
>   register row for the slide and period. **A cell is written ONLY when the
>   register holds nothing**, read structurally as `Not rowValues.Exists(field)`
>   because `ReadSheetForPeriod` only dictionaries non-empty cells. Refuses on
>   duplicate instance rows, never creates a column, never invents a row, and
>   **skips a tagged group by name** rather than reading a device's empty text as
>   a blank field.
> - **`PropagateTemplateTags`** — the missing middle link. Carries the template's
>   roles onto an already-linked slide, reusing `MatchSlideAgainstTemplate`
>   unchanged. **The caller must filter the roles first**: that function loops
>   EVERY template role while filtering the target to UNTAGGED shapes, so asking
>   about a role the slide already carries scores it against some unrelated
>   leftover. Stamps only on `high`; refuses when two roles claim one shape.
> - **`AdoptFlow.AdoptExistingSlides` was orphaned** — no button, no caller, for
>   the whole life of the three-button toolbar. Now called from
>   `RibbonUI.OfferAdoptionForSelectedSlides`. Its header said "Toolbar entry
>   point" the entire time.
> - **`check_vba_static.py` could not see it**: `AdoptFlow.bas` was missing from
>   `UI_MODULES`. Added. **And the checker is weaker than its name** — it asks
>   whether a NAME appears in another module, not whether anything reachable
>   calls it, so a chain of private orphans is still invisible to it.
> - Both new doors are gated on there being something to do. In Normal view a
>   slide is ALWAYS selected, so an ungated "you have slides selected" prompt
>   would fire on nearly every press.
>
> ### PROVEN BY A DELIBERATE BREAK — AND ONE THAT WAS WEAKER THAN CLAIMED
>
> - **Harvest's empty-cell rule: real damage.** Guard removed → all three
>   assertions failed, `FIELD_A was NOT overwritten, got 'FROM THE SLIDE A'`,
>   `Written` 2 not 1. The device test stayed green, isolating it.
> - **Propagation's role filter: bookkeeping only.** Filter removed → **one**
>   assertion failed (`already on the slide, got 0`). `FIELD_B` still landed
>   correctly and `FIELD_A` was not corrupted. **So the test proves the guard is
>   reached, NOT that removing it puts a role on the wrong shape.** The fixture
>   has two easily-distinguished shapes; the real slide has 136. This is the
>   weaker of the two guards and it is the one that writes to the deck.
> - **The first break attempt did not break anything** — commenting the call to
>   the wrapper left the callee's name still written inside it, and the checker
>   only greps for the name. That is what exposed the checker's real contract.
>
> ### ALSO: A SKIPPED SUITE RUN EXITS 0
>
> `run_vba_tests.ps1` refuses to close an open PowerPoint (correctly, since
> 14 Aug) and **exits 0 having done nothing**. Its own header still describes the
> old "ask it to Quit()" behaviour, which is what produced a wrong prediction to
> Rohan. Read the output file, never the exit code.
>
> ### OPEN, NAMED
>
> 1. **Propagation has never touched a real slide.** Dry run one first.
> 2. **The device covers 21 of the 29 fields and cannot be read back.** Slot
>    state lives in shape visibility — GAP 3, unbuilt. Harvest skips it by name.
> 3. **`SECTOR`/`TRL`: Rohan chose to split them into their own shapes.** Not
>    built. Note the split does NOT reach the other 42 slides, and `Text 4`'s
>    other two facts already exist as their own shapes (`Text 110`, `Text 112`) —
>    so `SUBTITLE_A` may be a composed output, not a field.
> 4. The stray `MS2_ON` on slides 1 and 44.
> 5. `STRATEGIC_LINKAGES` has a register column, no values, no template tag.
> 6. `PROJECT_STATUS` still has nothing enforcing its vocabulary.
>
> **Delivery count is 2.** Nothing reached a slide, and none of tonight's code
> has run against the real deck. Everything is committed to the working tree
> only — `Harvest.bas` and `backups/` are untracked.
>
> ---

> ## 14 AUG, LATE NIGHT (~22:15). **SUPERSEDED** by the block above.
>
> ### THE THING THAT CHANGED TONIGHT: DRAFTING → REGISTER WORKS.
>
> `3_P001`, `KEY_EVENTS_BODY`, 272 chars, published **by button**, 20:45.
> That link had never once run in this project's life. S5 is closed.
>
> ### FIRST ACTION TOMORROW: **THE HARVEST.** Nothing else comes first.
>
> There are roughly **1,247 values** — 29 `Given` fields × 43 projects — sitting on the
> slides and **nowhere else**. Read from the register tonight: `START_DATE`, `END_DATE`,
> `INDUSTRY_CASH`, `TOTAL_VALUE`, `PROJECT_LEAD`, `SECTOR`, `TRL`, `SUBTITLE_A`, every
> `MS*_LABEL/_DATE/_DONE` — **0 of 43 rows populated**. `PROJECT_STATUS` (43) and
> `PROJECT_PROGRESS` (1) are the only non-prose fields with anything in them.
>
> Tagging is not a chore, it is what makes harvesting possible: an untagged shape is
> anonymous, so the tool can neither read from it nor write to it. **Tag once on the
> template, harvest, and the register fills itself from the deck.** Nobody types 1,247
> values.
>
> **Do not do the quarter turn before the harvest** — it would roll forward emptiness.
>
> ### STATE, READ FROM FILES
>
> - **`addin88` is the only registered add-in**, `AutoLoad=1`, confirmed loaded.
> - Suite **194 passed / 0 failed**. Compile clean, 33 modules. Static clean.
> - Commits tonight: `a9f52b7`, `6c74912`, `3e7db62`, `db364b3`, `d1278b4`, `368d873`.
>   All pushed.
> - Register **394,082 bytes**, zip integrity OK, 39 sheets. Drafting sheets at
>   **layout 5**, 43 rows / 43 SUBMIT each; `PROGRESS_BODY` holds its 43 approve ticks.
> - Review queue **open, nothing applied**: 32 rows ticked = the 19 INVISIBLE changes only
>   (punctuation and casing). The 36 visible ones are deliberately unticked.
> - Backups tonight: `PRE-ADDIN86-20260814-185040`, `GOOD-STATE-20260814-202633`,
>   `register-wide.PRE-RESTORE-MIGRATION-*`, `register-wide.PRE-TICKSET-*`.
>
> ### THE TOOLBAR IS THREE BUTTONS, SPLIT BY ARTIFACT
>
> `1. Set up my quarter` (workbook) · `2. Put it on the slides` (deck) ·
> `Review changes (writes nothing)`. **Neither side can trigger the other** — that coupling
> is what wiped 43 approve ticks in the morning. Seven dialogs became **one approval (the
> review tick) plus two selections**. The field picker is **deleted**, not improved:
> publish walks every Prose field. `Rebuild my sheets` is deleted — "press this when a
> sheet looks wrong" is a defect with instructions attached.
>
> ### WHAT COST THE EVENING, AND WHAT IT TAUGHT
>
> **`MigrateSheetLayout` wiped every drafting sheet on the real workbook at 19:11** — 129
> drafted paragraphs, 43 approve ticks, 75 notes. **All recovered**, 412 values verified
> through Excel against the backup, 0 overwritten.
>
> The defect was a **bootstrap error**: `sheetLayout` was read at `COL_D_LAYOUT` and
> `sheetPeriod` at `COL_D_PERIOD` — the CURRENT layout's stamp columns. *You need the
> layout to know where the layout stamp is.* On a layout-4 sheet those cells hold the
> prompt and nothing, so the version came back 0, `layoutMatches` went False, the
> whole-sheet clear fired, and the migration — guarded on the same flag — never ran.
> **Fixed:** `DetectLayoutFromRow` searches the intro row right-to-left; the period is then
> read at *that* layout's position.
>
> **The test that should have caught it encoded the bug.** Its fixture wrote layout-3 DATA
> under layout-5 STAMPS — a sheet that has never existed — and passed only because the
> reading code made the same mistake. That is why 192 green tests certified a defect that
> destroys real work.
>
> ### ALSO SHIPPED, EACH PROVEN BY A DELIBERATE BREAK
>
> - **The pairing cross-wiring hole.** `specs/deck-registry.md` claimed the deck↔workbook
>   pairing "closes the cross-wiring risk". It did not: the workbook's `DeckReference` GUID
>   was written once at onboarding and **never read for its purpose**. Now stamped on both
>   ends at repoint and checked before every register write. A blank stamp is NOT a
>   mismatch — every pre-existing register has one.
> - **Discovery is device-aware.** A group with slots is ONE candidate. The timeline used to
>   present as 21 fields to hand-tag.
> - **The test runner stopped closing your Office windows** and now clears more than one
>   husk (`Count -ne 1` meant one husk self-heals and two are permanent).
>
> ### WHAT ROHAN'S QUESTIONS DELETED — the pattern, again
>
> - *"It shouldn't have to ask"* → deleted the field picker, its wording, and the
>   Excel-focus bug just built for it.
> - *"What happens when I tag them?"* → reordered the plan. Tagging enables the HARVEST.
> - *"I don't have 26 duplicate slides"* → stopped a deck-integrity hunt. Read the deck:
>   43 `INSTANCE_KEY`, 44 `SLIDE_TYPE`, clean. The duplicates are 30 **queue rows** with
>   identical ChangeIDs.
> - *"I thought the only generative text was coming from drafting sheets"* → exposed that
>   the review queue shows every register↔slide difference regardless of origin. Of 55
>   queued changes, **one** came from tonight's drafting.
>
> ### OPEN, NAMED
>
> 1. **The harvest** — the 1,247 values. Everything queues behind it.
> 2. **`GAP 3` is hours, not a build.** `MilestoneDevice.bas:630` already writes
>    `shp.Visible`; the template already carries `MS1_ON/_NOW/_OFF` … `MS7`. Expose it as a
>    field behaviour. `EXPECTED-TRACE`'s "no `.Visible` write anywhere" is **wrong** and now
>    annotated as such.
> 3. **The review queue emits duplicate rows** — 85 rows for 55 changes, identical
>    ChangeIDs. Cosmetic; Rohan ranked it below connectivity.
> 4. **`PROJECT_STATUS` has no vocabulary enforcement.** 17 values out of list tonight,
>    reported and not written.
> 5. **`STRATEGIC_LINKAGES` now HAS a register column** (added tonight with 16 others) but
>    no values and no template tag.
> 6. The unsaved-workbook prompt before review always answers Yes — a save-and-report, not
>    a question. And something after publish dirties the workbook without saving.
>
> ### DOCUMENT CONTROL, DONE THIS SESSION
>
> `TOOLBAR.md`, `WORKFLOW.md`, `FIX-LIST.md` and `EXPECTED-TRACE-2026-08-14.md` were
> scanned and corrected where tonight made them false — button names, layout 4 → 5 column
> letters, and the wrong `.Visible` claim. **Historical records were annotated, not
> rewritten.** `NEXT-SESSION.md`'s older blocks quote past suite counts; those are dated
> observations and stay as they are.
>
> **Delivery count is 2.** Tonight's 8 `PROJECT_STATUS` writes applied ticks made at 16:51,
> for a field already delivered on 8 August. The machinery moved a long way; the count did
> not.
>
> ---

> ## 14 AUG, NIGHT (~19:00). **SUPERSEDED** by the block above.
>
> ### FIRST ACTION: SAVE `addin86`, THEN PRESS `1. Sync Now` AND WALK THE LOOP.
>
> `addin86` is imported and waiting on the one permanent manual step
> (`File > Save As > PowerPoint Add-in`). `addin85` is registered and loaded but
> **predates everything below it.** Nothing in this block is reachable until 86 is saved,
> ticked, and 85 is unticked.
>
> Deck and register are **backed up and md5-verified** at
> `backups/PRE-ADDIN86-20260814-185040/`.
>
> ### THE SIZING IN THE BLOCK BELOW WAS WRONG. ROHAN CAUGHT IT. DO NOT REPEAT IT.
>
> Earlier tonight this file's plan cut Rohan's five imagined steps ("A") from the weekend
> as **weeks** of work. He challenged it and he was right. **The estimate was made without
> opening the modules** — the failure this project has logged nine times.
>
> | the claim he made | what the source says |
> |---|---|
> | marking/confirmation runs against the template, manual **or** matrix review | `BatchOnboardFlow.BuildBatchPlan(templateSld, otherSlides())` and `BuildBatchPlanFromMarkedFields(...)` — **both exist and are tested** |
> | onboarding runs off slides matching the template | that IS `BuildBatchPlan` |
> | templates are provided at the start | `TemplateSlide.MakeTemplateFrom(...)` exists |
>
> **`EXPECTED-TRACE-2026-08-14.md` is WRONG on GAP 3.** It says there is "no `.Visible`
> write anywhere". `MilestoneDevice.bas:630` writes `shp.Visible` and it works — it is
> private to that device. The job is EXPOSING it as a field behaviour, not inventing it.
>
> **Revised: GAP 1 hours (spec + wiring), GAP 3 hours, GAP 2 already plumbing, GAP 4 days
> and deferrable.** The scoring doc itself said "3 of 5 substantially built" and this file
> still said weeks. **A handover document is a summary of the code and goes stale in hours.
> Read the module.**
>
> ### THE "PUBLISH ONE FIELD BY BUTTON" MILESTONE IS DISSOLVED, NOT SKIPPED
>
> Rohan: *"no we have proved we can push to slides over and over. finish this against
> target"*. S6 is proven repeatedly. The target loop IS tweak → sync → tweak → sync, so
> **the first real sync exercises S5 as a by-product.** Building a separate ceremony for it
> was making a step out of something that happens anyway.
>
> ### WHAT SHIPPED TONIGHT — suite 192/0, static clean, compile gate passed
>
> **1. The pairing defect: the GUID was written and never read.**
> `specs/deck-registry.md` claims the pairing is "mutually verifiable" and closes the
> cross-wiring risk. It did not. Deck stores the workbook PATH; workbook stores the deck's
> `DeckSyncId` GUID (path one way, identity the other — deliberate, since OneDrive breaks
> paths and not GUIDs). But the GUID was written once at onboarding and **nothing ever
> compared the two.** Only readers were a struct assignment and a `MsgBox` in a demo
> expecting the literal `"deck-v1"`.
> - `ExcelOutput.WriteDeckReference`/`ReadDeckReference` made **Public**
> - `DeckRegistry.StampPairing` — a repoint now writes BOTH ends, save-verified
> - `DeckRegistry.PairingProblem` + pure `PairingVerdict` — wired into
>   `PublishDraftsForField`, the write path into the register
> - **Blank is not a mismatch.** Every pre-existing register has a blank `DeckReference`;
>   refusing those would strand him at work. A DIFFERENT GUID refuses.
> - **Proven by two deliberate breaks**, each failing a different assertion by name.
>
> **2. Seven dialogs → one approval + two selections.** Deleted: "This will, in order…",
> "Write these into the register?", "Ready to build the list" (3-way), the Roll Forward
> confirm. **Bulk-approve is not lost** — it keeps its own toolbar entry, so it stays chosen
> by name rather than reachable by answering "No" to a question about something else.
> KEPT: quarter selection, field selection, review tick.
>
> **3. Two typed boxes became clicks, and one was a DATA HAZARD.** The field picker needed
> an exact match against a 30-item list that pushed its own text box off screen. The
> roll-forward source period was free text — and periods are matched EXACTLY, so `"q3f26"`
> or a quarter not present produces a clean run that copies nothing and reports success. The
> period is now read from a row the person clicks, so a typo is impossible **by construction
> rather than by validation**. Range-picker chosen over a UserForm because `build_ppam.ps1`
> imports `.bas` only.
>
> **4. `check_vba_static.py` caught a real defect in my own code** (a retry that recursed
> with identical arguments) hours after I called it "nearly worthless". Blind in one
> direction is not blind.
>
> ### THE PREDICTION FOR THE FIRST PRESS — recorded before it happens
>
> 1 quarter prompt · 2 ~~"Go ahead?"~~ gone · 3 roll-forward reports non-modally ·
> **4 drafting rebuild — the layout 4→5 migration AND the ferry, first time ever on a real
> workbook** · 5 click a row to pick the field · 6 ~~register confirm~~ gone ·
> 7 ~~"ready to build"~~ gone · 8 **review tick** · 9 apply.
>
> Step 4 is the only one that has never touched real files. If what he sees differs, **the
> diff is the finding** — take what he saw, not what was expected.
>
> ### STILL OPEN
>
> 1. **`STRATEGIC_LINKAGES` has no register column** — the whole output slide type depends
>    on it.
> 2. **`PROJECT_STATUS` drifted clean → dirty in three days.** A Controlled field with a
>    declared vocabulary has **nothing enforcing it**. Normalising fixes today; only
>    enforcement fixes next time.
> 3. **The template specification is now the critical path** — GAP 1 is hours of wiring
>    *given a spec*, so the spec is the bottleneck. Handed to chat.
>
> **Delivery count is still 2.** It moves on the next real sync, not on a ceremony.
>
> ---

> ## 14 AUG, EVENING. **SUPERSEDED** by the block above where they disagree.
>
> ### THE BUILD IS GREEN. THE QUARTER-TURN FERRY IS IN.
>
> Compile gate **passed** (proven: the gate blocks the suite entirely, and 191 tests ran).
> Suite **191 passed / 0 failed**. 191 = the previous 193 minus the two tests deleted below.
>
> ### FIRST ACTION: PRESS `Sync Now` AND PUBLISH ONE FIELD, END TO END, BY BUTTON
>
> Chat proposed this ordering and it is right. **The publish path has still never run
> through the tool** — the 43 `PROGRESS_BODY` values were written into the register by hand
> over COM. *Register -> slides* is proven; ***drafting -> register* is not.** Those are
> different claims and only one is true.
>
> Do it BEFORE restructuring the chain (rename, split, delete the three invariant prompts).
> If publish then fails, a restructuring bug is indistinguishable from the original defect.
>
> **Green does not mean reachable.** No test asks "can a person cause this to run". This
> project has shipped a tested picture injector, tested progress bars and a tested publish
> path all behind locked doors. Only the press answers it, and the answer is read from the
> SAVED FILE, never from a dialog.
>
> ### ROHAN'S RULE, 14 AUG: **ONE APPROVAL STEP, PLUS SELECTION.** NOT YET BUILT.
>
> *"there should just be one approval step (besides any selection)"* — after walking the
> chain himself and hitting seven dialogs.
>
> **The one approval is the REVIEW TICK.** That is where consent belongs: it is the only
> gate whose answer varies, and the only one standing between the register and his slides.
> Everything else becomes non-blocking and goes to the Run Log.
>
> **Dialogs to delete, in order met when pressing `1. Sync Now`:**
>
> | dialog | why it dies |
> |---|---|
> | `MS*` "fields with nothing to write into" | **DONE this session** — answer never varied; the 21 names are `MILESTONE_TIMELINE` internals, and "Yes" would have destroyed the device the warning pointed at. Now `WriteRunLog`. |
> | "This will, in order... Go ahead?" | Invariant. Orientation belongs on START HERE, not in the way. |
> | Roll Forward offer | Invariant when the destination already holds rows — it says so itself and then asks anyway. |
> | "Write these into the register for Q4F26?" | The register is not the deck. Nothing reaches a slide from it without the review tick. Backed up, reversible. |
> | "Ready to build the list of slide changes" | Pure ceremony in front of the real gate. |
> | "Nothing was changed" confirmations | Report, never a modal. |
>
> **KEEP:** the field selection (it is a selection, not an approval) and the review tick.
>
> **The field picker is a defect in its own right.** It is a typed `InputBox` requiring an
> exact match against a 30-item list, and the list is so long it pushes the text field OFF
> THE BOTTOM OF THE SCREEN. Rohan: *"I don't want to do any secret hidden typing, I need to
> select the field by clicking on it."* Make it a click.
>
> **COMPILED AND GREEN.** The `MS*` deletion is verified: compile clean across 33 modules,
> suite 191 passed / 0 failed.
>
> ---
>
> ### WHAT SHIPPED THIS SESSION
>
> - **Layout migration restored** (`MigrateSheetLayout`). The five `kept*` dictionaries WERE
>   the migration, as a side effect of clear-and-rebuild; deleting them deleted it, and a
>   layout-3 sheet was being read with current column numbers -- SUBMIT text becoming a
>   source ID, the AI draft landing in the column that publishes.
> - **The quarter-turn ferry.** `DRAFT_LAYOUT_VERSION` 4 -> 5, new `COL_D_PREV` at column 4
>   ("REPORTED LAST TIME"). On a period change SUBMIT moves sideways, the working columns are
>   handed to the new quarter empty, and the sheet is parked first because sources and notes
>   ARE cleared.
> - **The rollover refusal is DELETED** -- it was a deadlock. It told people to publish
>   first; publish never cleared what it objected to; the only `ClearContents` sat past the
>   exit. The tool could not start a new quarter. The per-row cadence machinery and the
>   `cadence` parameter went with it.
> - **Every column letter in headers and instructions is now DERIVED.** They were literals
>   and the one-column shift made all six name the wrong column while still reading correctly.
> - Tests: two deleted (`RolloverCadenceGovernsUntypedRows`,
>   `RefusesRatherThanDiscardOnPeriodChange` -- they asserted the deadlock), one replaced by
>   `Drafting_QuarterTurnFerriesSubmitIntoReportedLastTime`, which also asserts idempotence
>   on a second rebuild.
> - **Your register was repaired**: the two 11:40-damaged tabs had no layout/period stamp,
>   which would have triggered `Cells.Clear` on both. Re-stamped and verified from the file.
> - **The 27 lost paragraphs are RECOVERED** -- both sheets at 43 rows / 43 SUBMIT, verified
>   from the saved bytes. Restored by writing ONLY into cells empty in the live file, which
>   made the backup predating the 08:21 publish irrelevant.
>
> ### THE COMPILE ERROR, AND WHY IT MATTERS BEYOND ITSELF
>
> `ByRef argument type mismatch`. `DraftingUI:521` passed a bare positional `Nothing` into
> the removed `cadence` slot, so `Nothing` bound to `srcWs` and `srcWs` to a Long. The
> comment three lines above said *"No cadence argument"* and was READ INSTEAD OF THE CODE.
> **It described the argument BY NAME for a call that passes POSITIONALLY** -- true about
> intent, false about the call.
>
> ### KNOWN GAPS, NAMED
>
> 1. **The ferry has never touched a real workbook.** The test builds its own two-project
>    fixture. Your live sheets are layout 4 and will migrate to 5 on the next refresh -- that
>    is the first real exercise of BOTH new things at once. Back up first.
> 2. **No test for corrected text.** Chat's point: if column F is corrected mid-quarter the
>    ferry must carry the CORRECTED value. The code reads SUBMIT at the moment of the turn so
>    it should be right; nothing proves it.
> 3. **`STRATEGIC_LINKAGES` has no register column.** 32 columns and it is not one. The
>    output slide's chips are computable in principle and have **no input at sync time**.
>    Fold into the sixteen-column pass; it is the one that whole slide type depends on.
> 4. **`PROJECT_STATUS` drifted clean -> dirty between 11 and 14 Aug.** 17 rows, all Q3F26,
>    16 pure casing. The finding underneath: **a Controlled field with a declared vocabulary
>    has nothing enforcing it.** A defect in the control, not the data.
>
> **Delivery count is 2.** It moves when a drafted field reaches a slide through the tool.
>
> ---


> ## 14 AUG, LATEST. **FIRST ACTION: COMPILE `Drafting.bas` IN THE VBE AND READ THE DIALOG.**
>
> The reported-last-time build is committed and pushed and **THE COMPILE GATE IS RED.
> NO TESTS RAN. NOTHING IN IT IS VERIFIED.** `vba/Drafting.bas` carries a warning banner
> on line 1; delete it once green. Committed only because losing it was the bigger risk.
>
> The gate prints no line number — the VBE dialog does, and PowerPoint is left on screen
> holding it. Open it and read it. Suspects, in order: the deleted `droppedQuarterly` /
> `keptStatic` counters (report block ~line 1050), the removed `cadence` parameter from
> `WriteDraftingSheet` (a caller may still pass it), and the new ferry block in the row loop.
> `check_vba_static.py` says clean and CANNOT see this — it does not check block balance.
>
> ### FROM CHAT, 14 AUG 14:49 — TWO FINDINGS TO CARRY
>
> `OneDrive\Claude\chat-status-for-claude-code-2026-08-14-night.md`.
>
> 1. **`STRATEGIC_LINKAGES` HAS NO REGISTER COLUMN.** 32 columns and it is not one of them.
>    This is the reopener named in the chips reply and it is REAL: the output slide's chips
>    are computable in principle but have **no input at sync time**. It does not stop Rohan
>    drawing; it stops the slide being populated. Fold it into the sixteen-column pass — it
>    is not just another missing column, it is the one the whole output slide type depends on.
>
> 2. **`PROJECT_STATUS` drifted from clean to dirty between 11 and 14 Aug.** 17 non-conformant
>    rows, ALL in Q3F26 (Q4F26 and Q1F27 clean), 16 pure casing. Legacy, so normalising cannot
>    break anything live. **The finding underneath it matters more: a Controlled field with a
>    declared vocabulary has NOTHING ENFORCING IT.** That will recur, and it is a defect in the
>    control, not in the data.
>
> Also settled by Rohan: progress bars stay GREEN (retires the May traffic-light spec), and
> `Project Closed` is TERMINAL. Chat's operational claims are hypotheses to check against the
> source, not instructions — it has twice advised something impossible in the tool.
>
> ---
>
> ### WHAT THE BUILD DOES, once it compiles
>
> Rohan, 14 Aug: *"whenever a 1/4 changes at the top, the ferries belonging to that system
> deal with information for the new quarter. The last previous set that runs move info into
> 'reported last time' column."*
>
> **This REVERSED the earlier "derived from the register" decision, deliberately.** The
> ferry runs at the moment the update notices the period changed, so it needs no predecessor
> tracking, no second register read, and no period ordering — and it preserves text that was
> never published, which the derived version could not show.
>
> - `DRAFT_LAYOUT_VERSION` **4 → 5**, new `COL_D_PREV` at column 4; everything from SOURCES
>   rightwards shifts one right. `ColumnInLayout` gains an explicit `Case 4`.
> - **The rollover refusal is DELETED** — it was the deadlock. So is the per-row cadence
>   machinery and the `cadence` parameter.
> - On a quarter turn: SUBMIT moves to `COL_D_PREV`, the working columns are cleared, and
>   the sheet is parked first because sources and notes are genuinely destroyed.
> - Every column letter in the headers and instructions is now DERIVED. They were literals
>   and the shift made all six name the wrong column.
> - Intro rows 1–2 are wiped before rewrite so a layout bump cannot strand an old stamp.
> - Tests: `Drafting_RolloverCadenceGovernsUntypedRows` and
>   `Drafting_RefusesRatherThanDiscardOnPeriodChange` **deleted** (they asserted the
>   deadlock). `RolloverRebuildsOnlyWhenNothingIsAtRisk` **replaced** by
>   `Drafting_QuarterTurnFerriesSubmitIntoReportedLastTime`, which also asserts idempotence
>   on a second rebuild — none of it has ever run.
>
> **The invariant was restated, and this matters.** It said "if the carry dictionaries come
> back, so has the clear". Rohan: *"can't they just sate the need then run on update?"* They
> can. Holding values in a variable was never the defect; opening a gap was. The real
> invariant is **no `ws.Cells.Clear` on the normal path**.
>
> ### ALSO DONE 14 AUG, AND VERIFIED
>
> - Layout migration restored (`MigrateSheetLayout`), suite 192/1 at commit `a5156f8`.
> - **The 27 lost paragraphs are RECOVERED** — both sheets at 43 rows / 43 SUBMIT, read
>   back from the saved file. `PROGRESS_BODY` also carries 43 restored approve ticks.
> - The live register's two damaged tabs were re-stamped to layout 4 / `Q4F26`.
> - **Chat-side template handover written:**
>   `OneDrive\Claude\handover-to-chat-templates-2026-08-14.md`.
>
> **Delivery count is still 2.**
>
> ---


> ## 14 AUG, LATE. **SUPERSEDED** by the blocks above — it was current on 14 Aug
> and its banner was never demoted, so this file carried TWO blocks claiming to
> be current until the 14 Aug 23:10 block above. Kept for reasoning only.
>
> ### THE COMPILE IS DONE AND IT WAS GREEN
>
> `aff84d6`'s restructure **compiles clean, whole project, 33 modules**. That was the
> previous block's blocking first action and it is closed. Suite came back **191 passed,
> 2 failed** — both failures real, both diagnosed below.
>
> ### YOUR REGISTER WAS REPAIRED. VERIFIED FROM THE SAVED FILE.
>
> `TPL_KEY_EVENTS_BODY` and `TPL_PROGRESS_BODY` had **no layout stamp and no period
> stamp** — the 11:40 abort died before writing them. A blank K1 reads as layout 0,
> `ColumnInLayout(0,"SUBMIT")` returns 0, `layoutMatches` is False, and **`ws.Cells.Clear`
> would have fired on both sheets** at the next rebuild. Stamped back to `4` / `Q4F26`
> through Excel. Confirmed from the saved XML: both sheets `K1=4`, `M1=Q4F26`, nothing
> beyond column M, `maxrow=52` matching every healthy sheet, counts intact at 23 rows /
> 22 SUBMIT and 37 / 37. Backup: `backups/register-wide.PRE-STAMPREPAIR-20260814-132207.xlsx`.
>
> **A COM diagnostic written as read-only WROTE to the live register** — `Close($false)`
> cannot discard while AutoSave is on. The AutoSave hazard was already documented in this
> file for PowerPoint and was not carried across to Excel. It left a `Z999` probe cell on
> both sheets; removed and verified gone. Logged to FRICTION.md.
>
> ### DEFECT: THE LAYOUT MIGRATION WAS DELETED WITH THE CARRY DICTIONARIES. FIXED.
>
> The five `kept*` dictionaries were the layout migration, as a side effect of clearing
> and rebuilding. Deleting them deleted it. `layoutMatches` is
> `(ColumnInLayout(sheetLayout,"SUBMIT") > 0)` — **True for a layout-3 sheet** — so the
> clear never fires, the row loop reads the sheet with layout-4 numbers, and SUBMIT text
> becomes a source ID while the AI draft lands in the column that publishes. The sheet is
> then stamped forward, so it can never self-correct.
>
> **Fixed** by a new `MigrateSheetLayout` — in place, per row, read-all-then-write
> (3 -> 4 is a permutation of the same six columns), old positions emptied between, parked
> first. **No clear, no carry dictionaries: the structural proof that the clear is gone
> still holds.**
>
> **This was NOT urgent until it was.** Nothing on the live register is layout 3. But the
> "reported last quarter" column below is a layout bump to 5, which runs this exact path
> against sheets stamped 4. Fix had to land first.
>
> ### DEFECT: A ROLLOVER CANNOT HAPPEN AT ALL. NOT YET FIXED.
>
> Traced end to end, not inferred:
> - `Start a Quarter` sets the deck period; does not touch drafting sheets.
> - `Roll Forward` copies register rows; does not touch drafting sheets.
> - `Refresh Drafting Sheets` with `periodChanged` → at-risk scan finds typed work →
>   **REFUSED, nothing changed.**
> - The refusal says *"publish this sheet's work before rebuilding"*. `PublishDrafts` only
>   READS `COL_D_SUBMIT` and `COL_D_APPROVED` (`Drafting.bas:1123,1125`). It never clears
>   them. The next attempt refuses identically.
> - The only `ClearContents` for those columns is **inside the branch the refusal already
>   exited past**. No other route clears or deletes a drafting sheet.
>
> **Once a sheet holds typed work — the wanted state — it can never be rolled forward.**
> Predates tonight: the refusal landed 13 Aug in `ddf867b`. Tonight's FIX-LIST 1c only
> widened the trigger to SOURCES and NOTES. Never hit because no rollover has been
> attempted since. **It bites at Q1F27.**
>
> ### ROHAN'S ANSWER DISSOLVES IT — AND DELETES CODE
>
> *"an owner should be able to edit their quarter's slides by editing their quarter's
> spreadsheet and hitting sync. When the deck is set to a new quarter, items that were
> current last quarter but now need replacing will move to a 'reported last quarter'
> column in the drafting sheet, adding fodder for the AI prompt re style and narrative
> consistency."*
>
> If last quarter's text moves aside instead of being destroyed, **the refusal has nothing
> left to protect** — republishing is prevented structurally. That kills the refusal, the
> cadence machinery and the deadlock together, and adds prompt material. Same move as the
> in-place fix and the per-quarter split. `Drafting.bas:228` already states the principle:
> *"nothing is destroyed, only superseded."*
>
> **SETTLED: the column is DERIVED from the register, not moved from the sheet.** The
> register already holds it — 43 rows at `Q3F26`, 43 at `Q4F26`. Derived shows what
> actually reached a slide, which is the honest prompt fodder; moved would mix in drafts
> that were never published and make the heading false.
>
> **The one cost, named:** periods are free text with **no ordering** (`Q3F26` is just a
> string). `RollForwardPeriod` is told both periods by the UI and does not record the
> link. So deriving needs the predecessor stored at roll-forward time. That is the
> prerequisite, and it is small.
>
> ### THE SECOND SUITE FAILURE IS A DESIGN COLLISION, NOT A BUG
>
> `Drafting_RolloverCadenceGovernsUntypedRows` asserts the old rebuild-and-clear
> behaviour; FIX-LIST 1c's widened scan now refuses first, so its assertions never run.
> Rohan called it: the scan wins, delete the cadence machinery. **NOT DONE, deliberately**
> — that machinery holds the only `ClearContents` in the rollover path, so deleting it
> before the "reported last quarter" column exists would weld the deadlock shut. Delete it
> as part of that build, not before.
>
> Also stale prose: the refusal says *"N row(s) have drafted or submitted text"* when it
> now also fires on sources or notes alone. Machine fact written into prose.
>
> ### NEXT, IN ORDER
>
> 1. ~~**Recover the 27 texts.**~~ **DONE 14 Aug, verified from the saved file.** Both
>    sheets are back to **43 rows / 43 SUBMIT**, nothing missing, every text byte-identical
>    to the backup. The 27 = 20 appended rows on `KEY_EVENTS_BODY` + `4_K017`, whose row
>    survived with an emptied SUBMIT (309 chars), + 6 appended rows on `PROGRESS_BODY`.
>
>    **The rule that made a wholesale restore unnecessary: a cell was written ONLY when
>    empty in the live file.** Nothing existing could be overwritten by construction rather
>    than by care, so the backup predating the 08:21 publish stopped mattering. Backup:
>    `backups/register-wide.PRE-RESTORE27-20260814-134147.xlsx`.
>
>    **`PROGRESS_BODY`'s 43 approve ticks were restored too** — they were wiped at 11:40 and
>    every SUBMIT on that sheet is byte-identical to what they approved, so this restores
>    consent rather than inventing it. `KEY_EVENTS_BODY` has 0; its backup had none. Flagged
>    to Rohan as a consent gate touched without asking.
> 2. **Build the "reported last quarter" column**: store the predecessor period at roll
>    forward, derive the column, bump `DRAFT_LAYOUT_VERSION` to 5, delete the refusal and
>    the cadence machinery, feed the prior text into the L2 prompt.
> 3. **Scenario 5**, once the sheets are whole.
>
> **Delivery count is still 2.** Nothing reached a slide on 14 Aug evening.
>
> ---
>
> ## 14 AUG, EVENING — SUPERSEDED by the block above. Kept for reasoning.
>
> ### FIRST ACTION: COMPILE AND RUN THE SUITE IN POWERPOINT. NOTHING ELSE.
>
> `vba/Drafting.bas` carries a substantial restructure that is **committed and pushed
> (`aff84d6`) but has NEVER BEEN COMPILED**. Everything else waits on that. It is
> committed because losing it was the bigger risk, not because it is trusted.
>
> **Do not trust `check_vba_static.py` as compile evidence.** It was handed a
> `Drafting.bas` containing a deliberate `If True Then` with no `End If` and printed
> `static checks clean across 34 module(s)`. It has a function named
> `check_structural_sanity` that does not check block balance. Proven by breaking it
> on purpose, 14 Aug.
>
> If the compile is red it is almost certainly one of five edits, each isolated and
> easy to bisect (listed below). A known-good copy of the file as it stood before the
> harvest deletion is at
> `scratchpad/Drafting.bas.pre-harvest-delete` — likely cleaned up; the real fallback
> is `git checkout vba/Drafting.bas`.
>
> ---
>
> ### WHAT HAPPENED: THE 11:40 RUN DESTROYED DRAFTED TEXT, NOT JUST TICKS
>
> The morning block below says *"No text was lost."* **That was wrong.** It counted the
> 22 and 37 SUBMIT cells that survived and never compared them to the 43/43 baseline
> sitting in that morning's own backup.
>
> | sheet | 14 Aug 07:16 | after 11:40 | lost |
> |---|---|---|---|
> | `TPL_KEY_EVENTS_BODY` | 43 rows / 43 drafted | 23 / 22 | **20 rows, 21 texts** |
> | `TPL_PROGRESS_BODY` | 43 rows / 43 drafted | 37 / 37 | **6 rows, 6 texts** |
>
> **Named, because a count is not evidence.** `KEY_EVENTS_BODY` lost every S-coded
> project plus `4_K021`, `1_K022`, `3_K023`. `PROGRESS_BODY` lost a strict subset:
> `1_S018, 3_S019, 3_S020, S021, S022, S023`.
>
> **Established from the files:**
> - Both sheets stopped **mid-function**. Rows present; prompt, layout stamp and period
>   stamp all absent — and all three are written *after* the row loop.
> - Truncation is contiguous at the tail. KEY_EVENTS stopped after register row 72,
>   PROGRESS after row 86 (the Q4F26 block spans physical rows 50–92).
> - **The register never changed.** 92 rows in all four backups and live.
> - **Not a cell-length problem.** Longest value 609 chars against Excel's 32,767. This
>   hypothesis was tested and is dead.
>
> **NOT established: why the write stopped where it did — AND DO NOT GO LOOKING.**
>
> Rohan, 14 Aug: *"please don't blow my budget on mysteries that are not part of the
> existing architecture and current decision making."* He is right. The in-place fix
> bounds this failure mode — a mid-write abort now leaves untouched rows untouched — so
> the cause is a curiosity, not a blocker. It is a one-off on a specific machine state
> and may never recur.
>
> **Reopen it only if it happens again**, and then with the evidence in hand rather than
> by hunting. Do not write a speculative cause into this file.
>
> **RECOVERY STILL OUTSTANDING.** The 27 texts exist in
> `backups/register-wide.PRE-ABOUTFIX-20260814-071658.xlsx`. **A wholesale file restore
> is wrong** — that backup predates the 08:21 PROGRESS_BODY publish. The rows must go
> back through Excel once the rebuild is proven safe. A fresh full backup of both files
> is at `backups/PRESCENARIO5-20260814/`, md5-verified against the originals.
>
> ---
>
> ### THE FIX: THE DRAFTING SHEET IS UPDATED IN PLACE, NOT REBUILT
>
> Rohan, after being shown that the park-before-clear change made the destruction
> survivable rather than removing it: *"why is clear still happening? For a fresh
> project or new quarter?"* He was right — it was patch six in a family of five.
>
> **`ws.Cells.Clear` now runs in exactly one branch: `If Not layoutMatches`** — a column
> renumbering, which has happened three times in the tool's life
> (`DRAFT_LAYOUT_VERSION` is 4). Everything else updates in place.
>
> **The proof condition, and it is structural rather than a claim:** the five `kept*`
> carry dictionaries are **deleted**. They existed only to ferry a person's work across
> the gap `Cells.Clear` opened. *If they ever come back, so has the clear.*
>
> The five edits, in the order they were made:
> 1. Park before every clear, not only on a layout mismatch (the old guard was gated on
>    `strandedRows > 0`, counted only inside `If Not layoutMatches` — so the ordinary
>    case took **no copy at all**, which is why 11:40 had no archive).
> 2. Layout + period stamps written **immediately after** the clear, not at the end. A
>    mid-write failure used to leave rows with no stamp, which reads as "unknown layout"
>    next time — and unknown layout carries nothing. A crash armed the *next* run to
>    discard everything that survived.
> 3. At-risk scan now counts **SOURCES and NOTES** (FIX-LIST 1c).
> 4. FIX-LIST 1d's late `ParkSheetCopy` call **removed** — it copied the already-cleared
>    sheet and reported "nothing was lost".
> 5. Rows addressed **by project code, not by position** (`rowOf` index, `appendAt`
>    cursor); human columns never written; harvest block and dictionaries deleted;
>    `r = appendAt` restored after the loop for the chrome's benefit.
>
> **Behaviour now:** an existing project keeps its row and its typed columns are not
> touched at all. A new project appends one row at the bottom. On a **period change**
> only, and per row, the work columns are cleared — the one case that must still
> destroy, because last quarter's text must not be republishable as this quarter's.
>
> **Known loose end:** `lostWithContent` (line ~530) is now declared and never
> incremented — dead, harmless, mine to clean.
>
> ---
>
> ### THE COLUMN CONTRACT, SETTLED 14 AUG — do not re-derive it
>
> - **Nothing in the add-in ever writes column E (AI DRAFT).** The only two references
>   are the header text and the rollover clear. Copilot writes it, externally, from the
>   prompt the tool puts in `L2`. There is no AI call inside the tool.
> - **Column F (SUBMIT) is written in exactly one place** — `CopyAiToSubmit`,
>   `Drafting.bas:1272` — and it skips any row where F already holds text.
> - **Publish reads F and G. It never reads E.** That is what makes "NEVER published"
>   a mechanism rather than a label.
>
> So AI text reaches the register only through the person's own column, via an explicit
> press that cannot overwrite them, plus a tick.
>
> ---
>
> ### ROHAN'S IMAGINED USE OF THE TOOL — captured before scenario 5 ran
>
> **`EXPECTED-TRACE-2026-08-14.md`** (untracked, in the repo root). His five steps in
> his own words, plus the diff against what the code does. Written down first
> deliberately: without a recorded prediction, whatever the tool does will look like
> what it was supposed to do.
>
> Scoring: **3 of 5 substantially built, 1 half, 1 absent.** The four gaps:
> 1. The template does not produce the workbook — direction is reversed (this is the
>    13 Aug template-first pivot, decided and never built).
> 2. **There are TWO approval gates and he imagines one.** Publish is per-field
>    (`FieldForRun` asks which); the compiled view he describes is the review queue
>    behind it. **He never once mentions choosing a field.** Very likely the root of the
>    tick confusion.
> 3. **Shape visibility is entirely unbuilt.** No `.Visible` write anywhere; `Behaviour`
>    is fill/fit/as-is for pictures only. The only one of the five with zero machinery.
> 4. **He expects a file per quarter.** **SETTLED 14 Aug — see `DECISIONS.md`.**
>
> ### GAP 4 IS SETTLED: ONE FILE PAIR PER QUARTER, PLUS A DERIVED CENTRAL REGISTER
>
> Deck **and** register split per quarter. A central register for multi-period analysis
> is **derived and never authored** — rebuilt by reading the quarter files, read-only,
> nothing ever flows back out of it into a quarter. Rohan's addition, and it closes the
> only real objection to splitting.
>
> **The cost was checked, not assumed, and it is nearly zero.** All twelve register
> reads are `ReadSheetForDeckPeriod(ws, <the deck's CURRENT period>)`. Exactly one
> function crosses periods: `RollForwardPeriod`. Column C reads the current period, not
> the previous one — the feared cross-quarter reads **do not exist**. An earlier version
> of this file named that as the unpriced cost; that was wrong.
>
> **The actual argument** is that the drafting sheets live *inside* the register
> workbook. If the register does not split, they do not split, and clearing last
> quarter's work out of a live sheet stays a normal-path operation — the operation that
> lost 27 paragraphs on 14 Aug. Per-quarter, the old quarter survives *by construction*.
> Same move as the in-place fix, one level up.
>
> **Not designed yet:** the aggregate's rebuild trigger and location; and the naming
> convention, which stops being cosmetic — `GetWorkbookPath` is already a bug source
> ("could not open the paired workbook", three causes in one morning) and more files
> means more pairing, on a machine with no Claude, Python or WSL.
>
> ### GAP 2 IS SMALLER THAN IT LOOKED — plumbing, not architecture
>
> Checked 14 Aug: **the review queue is already matrix-shaped.** `ChangeHash` is keyed
> per entity *and* field, so the compiled projects-x-fields view Rohan describes already
> exists downstream. The per-field bottleneck is only in **publish** (`FieldForRun`
> asking "which field?"), sitting in front of a machine that already handles the matrix.
> So this is closer to "publish across all fields" than to building a new surface.
>
> ### ON SHAPES vs TEXT BOXES — asked 14 Aug, answered from the code
>
> Injection already handles four kinds: text (`InjectPrimitive`), pictures
> (`InjectPictureVia`), progress bars (`InjectProgressVia`) and **devices**
> (`InjectDeviceVia` — a group consuming several register columns as one addressable
> thing). Visibility on/off is the one with no machinery at all (GAP 3).
>
> **Writing is solved; RECOGNISING is not.** Injection treats a device as one thing;
> discovery, wiring, marking and the audit all see its parts — which is why the timeline
> appeared as 21 candidate fields. That is Rohan's "load shape modules as prenamed per
> slide entities", and it is the device registry decided 13 Aug.
>
> ---
>
> ### DEFERRED, DELIBERATELY
>
> - **Scenario 5 has not run.** It should not, until the sheets are recovered and the
>   rebuild is proven — the chain rebuilds drafting sheets as a step, which is the
>   operation that caused this, and the baseline is currently 27 rows short.
> - **The rename** (`RefreshDraftingSheets` → `BuildDraftingSheets`) and the two-button
>   split. Reasoning is sound and recorded below; blocked on GAP 4.
> - The two-button shape agreed with Rohan: the boundary is **where he stops typing**,
>   not workbook-vs-deck. `[1. Start the quarter]` / `[2. Put it on the slides]`, one
>   add-in, count unchanged. Chat side had no quarrel with it.
>
> ### PROCESS NOTE
>
> Rohan's plain questions did the work again tonight, twice: *"aren't they permanent
> now?"* exposed the clear-and-restore shape, and *"why is clear still happening?"*
> caught the sixth local patch being applied to the same defect. Two FRICTION entries
> logged (the wrong "no text was lost" claim, and the static checker's blind spot).

> ## 14 AUG, ~12:00 — READ THIS BLOCK FIRST. It supersedes the 10:15 block below.
>
> ### STATE
>
> - **`addin84` IS LOADED.** Verified from the registry, not from a dialog:
>   `HKCU\...\Office\16.0\PowerPoint\AddIns\addin84`, `AutoLoad=1`, and **`addin83` is
>   gone**. Build stamp `2026-08-14 10:09`. File is byte-identical in `OneDrive\Claude\`
>   and the trusted `AppData\Roaming\Microsoft\AddIns\` (md5 `2b5514d6…`).
> - **The tick fix is committed and pushed** — `644ed16` on `main`.
> - **Delivery count is still 2.** No new field reached a slide today.
>
> ### WHAT HAPPENED AT 11:40, AND WHAT IT COST
>
> Rohan pressed `1. Sync Now` **while `addin83` was still the loaded add-in**. The chain
> rebuilt the drafting sheets and **wiped every approve tick on the two sheets it
> rebuilt** — defect 1 demonstrating itself on the real workbook, an hour after being
> fixed in source. Read from the saved file:
>
> | sheet | SUBMIT cells | APPROVE cells (data rows) |
> |---|---|---|
> | `TPL_KEY_EVENTS_BODY` | 22 | **0** |
> | `TPL_PROGRESS_BODY` | 37 | **0** |
> | the other eleven `TPL_*` | 43 | 43 present |
>
> **No text was lost. The deck was never touched** (still 08:21). **Nothing was
> published** — the Sync Log's newest entry is still `2026-08-14 08:21`, and that check
> was calibrated: the same search finds `07:48` and `08:21`, so its silence about 11:xx
> is evidence rather than absence.
>
> **NO BACKUP WAS TAKEN before that rebuild.** That is FIX-LIST 1d — the park runs
> *after* `ws.Cells.Clear`, not before. Still unfixed.
>
> **TO RE-TICK: `TPL_KEY_EVENTS_BODY` and `TPL_PROGRESS_BODY`.** Those two are certain.
> The other eleven have approve cells PRESENT, which is **not** the same as holding `Y` —
> a cell containing `0` counts identically, and this project has already recorded one
> wrong conclusion from exactly that (FIX-LIST P6). Unverified; check column G by eye.
> Under `addin84` those ticks now survive a same-quarter rebuild.
>
> ### ARCHITECTURE DECIDED THIS SESSION — Rohan's calls, all four
>
> 1. **The tick default INVERTS.** Everything auto-ticked at drafting; you untick what
>    you do not want shipped. Writing the text is the declaration; the exception should
>    carry the marking cost, not the majority case. **Two conditions attached:** the
>    review grid keeps FIX-LIST 3's conditional pre-tick (pre-tick only where the slide's
>    current text still exactly matches what the register last wrote, so a hand-edited
>    slide still demands a read); and the period reset becomes load-bearing rather than
>    incidental, because auto-tick plus a rollover would otherwise re-approve last
>    quarter's prose. Today's fix already drops the tick with the text on a period
>    change — that property now needs a test that says so out loud.
>
> 2. **THE CHAIN SPLITS BY ARTIFACT, NOT BY STEP.** One set of actions touches the
>    workbook, one touches the deck, and **neither can trigger the other**. This reverses
>    the 2026-08-09 two-button chain decision. The chain was right about the problem
>    (orientation — "which step am I up to?") and wrong about the remedy (removing the
>    choice), and the coupling it created is what wiped the ticks above: a person pressed
>    a button to PUBLISH and it REBUILT first.
>
> 3. **The deck side runs fully independently** of whether the workbook side has just
>    run. It reads the register as it stands on disk. This is what makes each half
>    testable alone — scenario 5 currently cannot be exercised without walking scenarios
>    1 and 2's machinery first.
>
> 4. **Review is its own action, not a step inside something else** — with its purpose
>    AND its exclusions stated ("for reading current vs proposed and ticking; not for
>    writing to slides"). Rohan: *"clear what it is and isn't for, part of a sequence, or
>    not."*
>
> ### THE ONLY THING TO BUILD NEXT — PHASE 1, AND NOTHING ELSE
>
> **Split the chain into independent buttons, and delete the three invariant prompts.**
>
> This is mostly DELETION. `RefreshDraftingSheets`, `PublishDraftsForField`,
> `ReviewChangesCore` and `ApplyApprovedCore` already exist as independent Subs;
> `RibbonUI.SyncNowChainCore` is a wrapper that calls them in order, and
> `CommandBarUI.AddButton` is called exactly twice. Removing the wrapper and adding
> buttons is small and subtractive.
>
> **DEFERRED, DELIBERATELY — do not build these as part of phase 1.** Scope grew in
> consecutive rounds this session and was cut back at Rohan's prompt (*"are we
> complicating this or making it simpler"*): self-describing action objects, promoting
> the `Readiness` surface as an orientation map, Excel-hosting the workbook half as an
> `.xlam` (FIX-LIST 8 — needs the deck's path and period written into the workbook's
> custom properties first), template-first, the device registry. All still wanted. None
> before scenario 5 has been walked once.
>
> ### PROCESS NOTE, WORTH KEEPING
>
> **Four rounds of chain design happened and the delivery count did not move.** Rohan
> asked "are we complicating this?" — the question this session should have asked itself
> at round three. The tell was scope growing in consecutive rounds while nothing new
> could be demonstrated. Consider running `deck-sync-pm` before the next design round.
>
> ---
>
> ## 14 AUG, 10:15 — THE APPROVE-TICK DEFECT IS FIXED AND PROVEN.
>
> **Working tree carries UNCOMMITTED changes to `vba/Drafting.bas` and
> `vba/tests/TestRunner.bas`.** Commit them. Nothing else is modified.
>
> **The fix.** `WriteDraftingSheet` no longer clears `COL_D_APPROVED`. The tick is
> harvested into a fifth `keptApproved` dictionary under the SAME `carryThisRow` test
> as SUBMIT, so it travels with the text it approves: same quarter keeps both, rollover
> drops both. Defect 1 below is closed — **publish is reachable**.
>
> **Rohan's framing is what produced it**, and it beat both proposals on the table.
> Chat side argued approval cannot live on the register row (the register is downstream
> of approval); Claude Code proposed reversing the chain order. Rohan asked *"why are
> drafting sheets getting rebuilt??? I thought they were staying as part of the record?"*
> — and the answer is that the FREQUENCY was the defect. The code already computed
> `periodChanged` and then ignored its own answer for one column.
>
> **Evidence — five suite runs, every check watched failing:**
>
> | run | state | verdict |
> |---|---|---|
> | 1 | tick fix in | 192/1 · tick test PASS |
> | break 1 | unconditional clear restored | 191/2 · *"SAME-PERIOD REBUILD KEEPS THE TICK ... got `''`"* |
> | break 2 | harvest outside `carryThisRow` | 191/2 · *"ROLLOVER DROPS THE TICK ... got `'Y'`"* |
> | 5 | + index test derived | **193 passed, 0 failed** |
> | break 3 | `DescribeSheet` points at draft column | 192/1 · both index assertions red, naming column E |
>
> Each break named a DIFFERENT assertion, so both halves are independently pinned.
>
> **THE SUITE HAD BEEN RED SINCE `a6e57af`, AND THIS FILE SAID 192/0 THE WHOLE TIME.**
> `WorkbookBridge_IndexExplainsEachSheet` hardcoded layout-3 letters (read C, type D,
> tick E) while `DescribeSheet` was correctly fixed to derive them from the `COL_D_*`
> constants. The test had already been written once to catch exactly this drift, and
> held the next generation of it in place for the same reason: **it typed a
> machine-knowable fact instead of deriving it.** Now derived on both sides. This is
> the write-it-twice class landing in a test, where no document checker can see it.
>
> **SUPERSEDED — `addin84` has since been saved, installed and confirmed loaded; see the
> block above.** As written: build stamp `2026-08-14 10:09`, all 32
> production modules imported. The `File > Save As > PowerPoint Add-in (*.ppam)` click
> is a permanent manual step (see `build_ppam.ps1` header — proven impossible to
> automate, twice, for independent reasons). **Verify the live build by the stamp
> reading `10:09`** — 15 `.ppam` files are on that machine.
>
> **Still open, unchanged:** defect 2 (chain dies with `Error 50290` when the register
> is open in Excel, then names the wrong file) and defect 3 (three invariant prompts).
> Both sit in scenario 5's path, so it is walkable-with-traps, not walkable.
>
> **Roll-forward spec received** — `OneDrive\Claude\rollforward-spec-2026-08-14.md`,
> scenario 1, all 48 fields with five treatments. Not yet filed into this repo. Findings
> returned to Rohan, highest first:
> 1. **"Placeholder" is doing two jobs** — treatment D means "drafting cell arrives
>    empty awaiting a human", the MECE rule means "standard wording printed on a
>    funder-facing slide". Collapsing them would placeholder active projects that should
>    be reported as GAPS. Pin before building anything on the table.
> 2. Chat's §0 approval argument is overtaken by the fix above.
> 3. §2.3 is a real catch — one milestone device holds three roll-forward cadences;
>    write that test before the device registry lands.
> 4. Totals line does not close (states 51 across treatments, table holds 47 rows,
>    `STRATEGIC_LINKAGES` is argued in §2.1 but has no row).
> 5. Open for Rohan: does the per-project-state rule LAYER with the per-field table or
>    override it? Chat assumed layering and asked to be told.
>
> **Rohan, 14 Aug: he owns the workbook and Claude Code may work that side directly**
> when it is the right tool — holding `feedback_write_office_files_through_office`
> (never hand-write the `.xlsx`) and backing up before any write to the live register.

> ## TWO FIELDS ARE ON REAL SLIDES. The delivery count is 2, not 0.
>
> **14 Aug 2026, 08:21 — `PROGRESS_BODY` written to 43 slides.** Verified from the saved
> `.pptx`, not from a dialog: 43 match the register, **0 differ**, 27 carry real paragraph
> breaks, no literal `||` anywhere. Build `addin83`, stamp `2026-08-14 08:09`.
>
> The workbook's **`Sync Log` is the durable record** (it appends; `Run Log` is replaced
> every run and cannot answer "what has this ever delivered"):
>
> ```
> 2026-08-13 16:55   KEY_EVENTS_BODY   written  21
> 2026-08-14 08:21   PROGRESS_BODY     written  43
> ```
>
> **Closed with it:** whether `||` reaches a slide as a real paragraph break. It does.
> FIX-LIST called that "the last unverified link in the chain".
>
> ### But read this before claiming the loop works
>
> **The 43 values were written into the register BY HAND** (Excel COM), because the
> publish path cannot be reached — see 1 below. So what is proven is *register → slides*.
> *Drafting → register* has still never happened through the tool. They are different
> claims and only one of them is true.
>
> Those 43 values also carry no sources and no recipe hash, because publish is where
> provenance would be written.
>
> ---
>
> ## STATE
>
> - **Deck** `OneDrive\Claude\3. Project Progress.pptx` — 44 slides, period `Q4F26`,
>   nine `ROLE` tags per slide. Tags live in `ppt/tags/tagN.xml`, **not** in the slide XML;
>   grepping `slideN.xml` for them returns nothing and looks exactly like an untagged deck.
> - **Register** `OneDrive\Claude\register-wide.xlsx` — 25 sheets, 43 rows each at
>   `Q3F26` / `Q4F26`, 5 at `Q1F27`. `PROGRESS_BODY` and `ABOUT_BODY` both 43/43 at Q4F26.
> - **Build `addin83` is the ONLY registered add-in** (`AutoLoad=1`; 82 was unticked and
>   removed from the registry). Thirteen stale `.ppam` files still clutter
>   `AppData\Roaming\Microsoft\AddIns` — delete them before one gets loaded by accident.
> - Backups from 14 Aug in `OneDrive\Claude\backups\`: `PREPUBLISH-20260813-204031` is
>   the last Excel-written baseline and is guaranteed openable.
>
> ---
>
> ## THE THREE DEFECTS THAT COST THE MORNING
>
> ### 1. The publish path cannot be reached. Still unfixed.
>
> `Drafting.bas:694` clears `COL_D_APPROVED` on every drafting-sheet rebuild. Draft,
> submit, sources and notes are carried; the tick is not. `SyncNowChainCore` runs
> **rebuild (step 3) immediately before publish (step 4)**, and nothing but a person ever
> writes that column. So a tick can never survive to be read.
>
> Demonstrated, not inferred: 43 rows ticked `Y`, saved to disk, `Sync Now` pressed,
> publish reported `0 would be published, 43 drafted but not ticked`.
>
> **Do not "fix" it by never clearing the column** — a tick approves a *specific* text.
> The real question is whether approval belongs on the **register row** instead of on a
> surface that is rebuilt by design. That is a design call, not a patch.
>
> ### 2. The chain fails when the register is open in Excel, and does not say so
>
> `Refresh Drafting Sheets` raises **`Error 50290`** when Excel already holds the
> workbook; every later stage then reports *"Could not open the paired workbook at
> C:\...\register-wide.xlsx"* — naming the file that is fine and discarding Excel's real
> reason. Self-defeating too: step 3 leaves the workbook dirty, and the apply's own guard
> then correctly refuses to read values that exist only in Excel's memory.
>
> **Workaround until fixed: close Excel before pressing `1. Sync Now`.** Cost four failed
> runs on 14 Aug before the pattern was visible.
>
> ### 3. Three prompts have an invariant answer
>
> Two presses produced ~20 dialogs; four were real decisions. The `MS*` warning, the
> 17-column question and Roll Forward can only ever be answered the same way. Rohan:
> *"way too many msgboxes and having to answer no is confusing."* This is not polish —
> it trains click-through past the one dialog where `No` destroys 43 ticks.
>
> ---
>
> ## FIXED IN `addin83`
>
> `DeckRegistry.SaveDeckVerified` and `WorkbookBridge.SaveWorkbookVerified` now resolve a
> OneDrive URL through `LocalPathForUrl` before touching the filesystem. Previously
> `fso.FileExists(url)` returned **False** for an `https://` path, so both functions
> returned "this workbook has never been saved to a file" and **exited before calling
> Save** — the one function written to guarantee the save was the only thing skipping it.
> Nine call sites for the workbook, two for the deck. Where the path cannot be resolved
> the save is now *attempted* and reported as unverified, rather than silently refused.
>
> ---
>
> ## NEXT, IN ORDER
>
> 1. **Approve-tick placement** (defect 1). Unblocks drafting → register, which is the
>    whole point of the recipes.
> 2. **Delete the invariant prompts** (defect 3), and make defect 2 say "close Excel first".
> 3. **Tag the sixteen standing `Given` fields on the template.** Propagation to
>    already-linked slides **is supported and tested** — `ExistingInstanceKey` protects
>    their keys, hardened after the 2026-07-26 incident that orphaned 46 slides. Only the
>    entry point is wrong: it asks *"Name for this new slide type"*. Verify what it does
>    on re-registering `project-progress`, on a carved copy, first.
> 4. **Then the template pivot** — but not making the template authoritative for the field
>    set until a quarter has been produced from it.
>
> **ANSWERED 2026-08-14, and do not re-ask it.** The load-bearing question of the whole
> project -- did the recipe tell him what to write, or did he still work it out himself --
> got its answer minutes after he was shown the `PROGRESS_BODY` Field Spec row:
> **"It worked and we will continue to refine over time."** The central bet holds: the
> recipes remove the re-deciding. A first affirmative read across two fields, not a full
> pass over thirteen; refinement is expected and he said so. `TRACKER.md` item 9 is ticked
> on the strength of it -- 9 of 10.
>
> **NOT blocked, and both halves are settled: Rohan owns the declared-linkage list, and
> its home is a PERMANENT sheet in the register workbook beside `Sources`.** It was put
> back to him as an open question on 14 Aug after both parts had already been answered --
> twice for ownership, and once by reading `WorkbookBridge.LifespanOf`. What remains is a
> task, not a decision: create the sheet, give it an `as at` date, and put the four codes
> already visible on the real slide (`1.1.1`, `2.1.1`, `3.3.1`, `3.4.1`) into it.
>
> Original note follows. FIX-LIST
> records this twice (*"Rohan owns it, colleagues later"*) and it was still written up as
> a question for him on 14 Aug, from a file that had been read in full the same morning.
> The real gap is that the list does not exist as an artifact yet: it needs a home
> **outside** the register workbook, which the tool rebuilds and clears, plus an `as at`
> date. Four declared codes are already on the real slide (`1.1.1`, `2.1.1`, `3.3.1`,
> `3.4.1`) with nowhere to live.
>
> ---
>
> ## RULES THIS MORNING EARNED
>
> **Never hand-write an `.xlsx`.** Re-serialising a worksheet with ElementTree produced a
> file that passed zip integrity, XML parsing, relationship, content-type, cell-order and
> inline-string checks — and Excel refused to open it, because only two namespaces were
> registered and `mc`/`x14ac`/`xr` got invented prefixes. **The only valid test of "can
> Excel open this" is Excel opening it**, A/B'd against a known-good file in the same
> session. Drive Excel over COM to write; read by hand is still fine.
>
> **A defect is a class, and so is a diagnostic.** "Could not open the paired workbook"
> appeared three times this morning with three different underlying causes, all hidden by
> the same discarded error. FIX-LIST item 1 is now the highest-value cheap fix in the repo.
>
> ---

**Written 13 August 2026, ~16:15.** Previous version archived as `NEXT-SESSION-2026-08-12.md`.


> ## SUPERSEDED — written mid-session, kept for the reasoning below it
>
> This banner read "THE DELIVERY COUNT IS STILL ZERO". **It is no longer true.** 21 drafted
> `KEY_EVENTS_BODY` values reached real slides at 17:23 on 13 Aug, verified three ways: the
> apply dialog, the slide XML in the saved file, and a screenshot of the rendered slide.
>
> 43 slides tagged and linked, 0 failed verification, register saved, deck period `Q4F26`
> on disk. The section below explains the defect that had blocked it, which is now fixed
> and proven — keep it for the reasoning, not as a description of current state.

---

## THE BLOCKER — FIXED AND PROVEN. Kept for the reasoning and the rejected workaround.

`RibbonUI.SyncNowChainCore` step 4 is `DraftingUI.PublishDraftsForField`, which begins:

```vba
fieldId = ActiveDraftField(wb)          ' whatever TPL_ sheet is ACTIVE in Excel
If fieldId = "" Then fieldId = AskForField(CAP, wb)
```

It only asks **if** the active sheet is not a drafting sheet. But the step immediately
before it — `RefreshDraftingSheets` — ends with `ShowSheet wb, firstSheet`, and
`firstSheet` is the **first `Kind = Prose` row on the Field Spec sheet**
(`DraftingUI.ProseFields`, row order).

That row is `ABOUT_BODY`, which has **0 submitted, 0 approved**. So every run publishes
an empty sheet, reports "0 would be published", and finishes quietly. **The chain cannot
reach any other field.** There is no field picker on the toolbar (two buttons only), so
this is the only publish route.

This is very likely a large part of why this project has never got a field onto a slide.

**A ROW-REORDER WORKAROUND WAS PROPOSED AND REJECTED — DO NOT USE IT.** Moving
`KEY_EVENTS_BODY` above `ABOUT_BODY` on the Field Spec would work, because
`FieldSpec.WriteSpecSheet` only seeds *missing* FieldIDs and never reorders existing rows.
It is still the wrong move, and Rohan stopped it with one question: *"Why are you having
to move register rows manually? Worries me that the code won't work when it needs to."*

He is right, for three reasons:

1. **It makes which field reaches a slide depend on spreadsheet row order** — invisible,
   unstated coupling of exactly the kind that has bitten this project repeatedly.
2. **It is not available at work.** No Claude, no WSL, no Python there — a quarter has to
   be runnable from toolbar buttons. "Reorder rows in a spec sheet so the right field
   publishes" is not a procedure; it is a defect with instructions attached.
3. **It would have hidden the defect behind a successful-looking run**, which is the
   failure mode this project keeps rediscovering.

**FIXED in `addin81`** (build stamp `2026-08-13 16:24`). New `DraftingUI.FieldForRun`:
inside a collected chain it ASKS; standalone it still reads the active sheet, because
there the answer really is on screen. Asked ONCE per run and reused, so the two stages
that need it do not ask twice.

**It had TWO call sites.** `CopyAiDraftsToSubmit` carried the identical line and the
identical consequence — fixed together rather than only where it was noticed.

**PROVEN 2026-08-13 17:23 on the real deck.** 21 KEY_EVENTS_BODY values written and confirmed by reading slide XML from the saved file -- slide 1 now carries the drafted "The industry partner's withdrawal..." wording. The delivery count is no longer zero. Previously read: 192 tests pass and the project compiles, but no test exercises the
chain's field selection — which is precisely the gap that allowed this defect. Green here
means "nothing broke", not "the fix works". Prove it by pressing the button: `Sync Now`
must now ASK which field, and `KEY_EVENTS_BODY` must be selectable.

**Note what the test suite did NOT do here.** 192 tests pass. Not one of them asks "can a
person cause `KEY_EVENTS_BODY` to be published?" — they test that publishing works when
called, not that the chain can reach it. Same "tested unit behind a locked door" shape as
the picture injection and the progress bars, found the same way: by pressing the button.

---

## STATE, VERIFIED FROM FILES (not from dialogs)

- **Deck** `OneDrive\Claude\3. Project Progress.pptx` — 44 slides, 49,247,250 bytes.
  `DeckSyncPeriod = Q4F26` confirmed by property name in `docProps/custom.xml`.
  Slide 44 is the hidden master template, 9 fields set to `<<placeholders>>`.
- **Register** `OneDrive\Claude\register-wide.xlsx` — 308,072 bytes, `Register` sheet has
  92 rows (1 header + 91: 43 Q3F26 + 43 Q4F26 + 5 Q1F27). All 43 Q4F26 instance keys
  match the Q3F26 keys exactly — **the handover's "stale/foreign Q4F26 rows" warning was
  wrong**, they describe the same slides.
- **Backups** `OneDrive\Claude\backups\2026-08-13-1520-post-onboard-Q3F26 - *` — deck and
  register, both verified byte-identical by md5 at the time of copy.
- **Build `addin82`**, stamp `2026-08-13 19:22` — CURRENT, confirmed loaded and the only
  add-in registered. **Source is one commit ahead of it** (`20dc2a9`, six sheet-name
  literals replaced by their constants). The values are identical, so behaviour is
  unchanged and a rebuild is not required — but `addin82` is not byte-equivalent to `HEAD`. In `OneDrive\Claude\` and the trusted location
  `AppData\Roaming\Microsoft\AddIns\`. (Superseded: `addin80` stamp 14:37, `addin81`
  stamp 16:24.)

### Drafting sheets — real counts (header row EXCLUDED)

| sheet | submitted | approved |
|---|---|---|
| `TPL_KEY_EVENTS_BODY` | 43 | 39 |
| `TPL_PROGRESS_BODY` | 34 | **42** |
| `TPL_HIGHLIGHTS_BODY` | 43 | 42 |
| `TPL_ABOUT_BODY` | 0 | 0 |
| `TPL_STRATEGIC_ALIGNMENT_BODY`, `TPL_PROBLEM_BODY`, `TPL_STRATEGIC_LINKAGES` | 0 | 0 |

`PROGRESS_BODY` has **more approvals than submitted text** (42 vs 34). Those 8 rows
publish nothing — both text and tick are required — but the count will look wrong.

`HIGHLIGHTS_BODY` has the most work in it and **cannot publish**: it is not one of the
nine tagged fields and needs slot columns, not one column. See FIX-LIST.

---

## ENVIRONMENT FINDING — SAVES AND ONEDRIVE

Both files are open via **OneDrive URLs**, not local paths:
`https://d.docs.live.net/96b9ec593ee3ba55/Claude/…`

With AutoSave **off**, the deck period write failed **4 verified attempts**, and a manual
`Ctrl+S` did not change the file's mtime either. `SetDeckPeriodVerified` correctly
detected this and refused to continue:

> `THE PERIOD DID NOT REACH THE FILE after 4 attempt(s). Asked for: Q4F26  On disk: Q3F26`

**Turning AutoSave ON made the write land.** This is an environment condition, not a code
defect — and it is the configuration the work machine will be in. The tool's behaviour
here was correct and is what a week ago was missing: it checked the file, not its own
cache, and refused rather than reporting success.

---

## WHAT SHIPPED TONIGHT

- **Suite green: 192 passed, 0 failed** (was 190/2), behind the compile gate.
- The two tests asserting the deleted defect were rewritten **and renamed**, because the
  old names stated the defect as the requirement:
  - `Drafting_PeriodRolloverDropsStaleSubmit` → `Drafting_RolloverRebuildsOnlyWhenNothingIsAtRisk`
  - `Drafting_RolloverKeepsEntityStaticRows` → `Drafting_RolloverCadenceGovernsUntypedRows`
- **`RefreshDraftingSheets` no longer reports success over a refusal.** It collected
  refusals into the Run Log and then said *"drafting sheets are ready. Workbook saved."*
  with an information icon. Now: refusal first (so MsgBox truncation eats the guidance,
  not the warning), refused field names listed, warning icon.

### NOT YET TRUSTED

**Neither rewritten test has been made to fail on purpose.** Green alone is not evidence.
Break each before relying on it — for `...RolloverCadenceGovernsUntypedRows`, put SUBMIT
text back on the fixture and it should stop testing anything, because the refusal
pre-empts the whole path.

---

## FILES CHANGED THIS SESSION — ALL COMMITTED AND PUSHED

`deck-sync-refimpl` (main, in sync with origin):
- `vba/DraftingUI.bas` — refusal surfaced in the dialog; `FieldForRun`; roll-forward skip
- `vba/RibbonUI.bas` — `PickType`; sync-log constant at two creators
- `vba/AdoptFlow.bas` — third `PickType` call site
- `vba/ExcelOutput.bas` — `PeriodRowCount`
- `vba/WorkbookBridge.bas` — `SYNC_LOG_SHEET_NAME`; lifecycle tab ordering
- `vba/tests/TestRunner.bas` — two tests rewritten + renamed
- `FIX-LIST.md`, `NEXT-SESSION.md`

`claude-brain`: `DECISIONS.md` — the template-first decision.
`Zettelkasten`: `20260813-the-device-is-the-unit-of-addressing-not-its-parts.md`.
`OneDrive\Claude`: `reply-from-claude-code-2026-08-13-evening.md` — answers chat side's
three questions; not a repo, so not versioned.

Builds: `addin80` (14:37) and `addin81` (16:24), both superseded. **`addin82` (19:22) is loaded and
the only add-in)**. Suite green at 192 passed / 0 failed with the compile gate clean across
33 modules.

---

## OPEN, IN PRIORITY ORDER

1. **Publish one field.** Field Spec row move → `Sync Now` → `KEY_EVENTS_BODY` → review →
   apply. Then verify by reading the slide XML out of the saved deck, not the dialog.
2. **Fix the publish-target defect properly** (above), then revert the row move.
3. **FIX-LIST 1c/1d** — at-risk scan misses SOURCES/NOTES; the park that reports "nothing
   was lost" runs *after* `ws.Cells.Clear`. Fixing 1c makes 1d unreachable.
4. **The cadence machinery is probably dead code.** The refusal pre-empts it; it now
   governs only SOURCES/NOTES on untyped rows. If 1c is fixed, delete it rather than
   maintain it. Rohan's call.
5. **`MILESTONE_TIMELINE` group tagging is UNVERIFIED.** It was not among the nine tagged
   fields. If the timeline renders blank after a sync, check this first.
6. **Field Spec `Kind` values look wrong for the milestones**: `MS1_LABEL`/`MS7_LABEL` are
   `Given` while `MS2`–`MS6_LABEL` are `Prose`; MS1/MS7 DATE+DONE are `Derived` while
   MS2–MS6 are `Given`. 13 fields are `Prose` but only 7 have drafting sheets, so the next
   refresh will create six more tabs.
7. Slide 44 still carries P001's unmanaged content (figures, photo, team). The audit found
   **50 unmanaged text items on slide 1, 21 of which look like project data** — that is the
   next tagging backlog, and the same set the Field Spec wants columns for.

---

## THREE THINGS WORTH KEEPING

**A warning that only reaches the log is not a warning.** The refusal guard was correct
and invisible; the dialog said "ready" over seven refused sheets. Fixed, but the shape
recurs — check where a message *lands*, not just that it exists.

**The check that found the save failure was the one that read the file.** Four
in-process attempts all "succeeded". Only comparing against `docProps/custom.xml` on disk
told the truth. Evidence must come from the far side of the boundary.

**"Nothing happened" meant a dialog behind the window — twice.** A VBA modal can open
behind PowerPoint. Before diagnosing a dead button, Alt+Tab. A calibrated test:
PowerPoint stops answering COM (`ActivePresentation.Name` comes back empty) while a modal
is open, and answers normally when idle.

---

## ARCHITECTURE — DECIDED IN PRINCIPLE 2026-08-13, NOT YET BUILT

Two calls made at the end of the first successful publish. Both are Rohan's, both are
right, and both should be settled properly before more feature work.

### 1. TEMPLATE-FIRST, NOT DISCOVERY-FIRST

Rohan: *"Are we better off gearing it to be an expert template builder and pushing a
pattern we know? ... maybe that's best rather than a sensory beast that doesn't quite
know what it is trying to be."*

**Yes.** The strongest argument is the WORK MACHINE. A sensory tool needs an operator with
judgement at every step — tonight it needed Claude to decide which of 59 shapes were
fields, whether the `MS*` warning was real, which register to pair, whether `TESTFILL` was
junk, and whether 88 proposed changes were safe. At work there is no Claude. A
template-driven tool needs no run-time judgement, because the judgement was made once, in
advance, and frozen into an artifact.

**The cost asymmetry says the same.** Discovery runs ONCE per slide type, ever. Publishing
runs 43 slides x 9 fields x 4 quarters, forever. Nearly all the code, nearly all the
defects and nearly all of 13 Aug went into the once-ever path.

**And look where the defects actually were:** the grid that loses marks, the blank grid
that unmarks, the 21 `MS*` false positive, 50 "unmanaged" items it can only guess at,
"which register?", the pairing. Every one is a PERCEPTION defect. The publish path — the
thing that runs every quarter — had one defect, four lines long.

**The mechanical diagnosis under "doesn't know what it's trying to be":** the tool holds
THREE sources of truth about what a field is — a tagged shape, a register column, and a
Field Spec row — and reconciles them at run time. Every reconciliation is a place to be
wrong. A template collapses all three into one, decided at design time.

**What that means concretely**
- The template is the authority; register schema derives from it.
- `FieldWiring`'s orphan-column question dissolves: nothing can be orphaned if the
  template defines the set.
- Discovery is DEMOTED to a one-off migration tool. Run once per deck, then never
  developed again. It has already been run — 13 Aug — so for this deck it is done.
- New projects clone slide 44 and are conformant by construction.

**What NOT to throw away:** the verification discipline (the file is the evidence; prove a
check can fail; a defect is a class), the Office/COM/OOXML knowledge, the consent-gate
design, and the register-deck contract. None of it is discovery-specific; all of it
transfers to the next project.

**The honest risk:** real slides are messier than any taxonomy — the finding that made
discovery seem necessary in the first place. But that now cuts the other way: owning the
template BOUNDS the mess instead of trying to perceive it, and the messy migration has
already happened.

### 2. A DEVICE REGISTRY — PROTECT COMPOUND SHAPES FROM THE GENERIC MACHINERY

Rohan: *"I can already see the need for a specialist module spot to protect complex shape
mechanisms like the timeline being pulled into marking etc."*

Three pieces of evidence from one evening:

- `MS1_DATE` … `MS7_LABEL` appeared as 21 rows in the Discover Fields grid, and the only
  thing that stopped them being tagged was **Claude telling Rohan not to**.
- `FieldWiring.ScanFieldWiring` reported those same 21 columns as orphaned on EVERY run —
  the recurring "21 field(s) on the register that no slide carries" warning.
- The Template Audit counted device internals among its "50 unmanaged text items, 21 of
  which look like project data".

Three generic mechanisms seeing the device's PARTS; none of them seeing the device.

**The model already exists, in exactly one place.** Injection has it right:

```vba
InjectPrimitive.InjectField(sld, "MILESTONE_TIMELINE", "", False, Nothing, row)
```

One addressable thing, consuming its own columns off the register row. That understanding
never propagated to discovery, wiring, marking or audit — so a device is a first-class
citizen at write time and a pile of loose shapes everywhere else.

**Principle: the device is the unit of addressing, not its parts.**

**Shape of the fix.** One declaration per device — its role tag, the register columns it
consumes, its internal shape-name pattern — read by four consumers:

| consumer | today | with the registry |
|---|---|---|
| Discovery | lists 21 internals as candidate fields | skips anything inside a declared device |
| `FieldWiring` | reports 21 orphan columns every run | counts them as OWNED by the device |
| Marking | will happily tag a device internal | refuses |
| Template Audit | counts internals as unmanaged project data | classifies as device internals |

One declaration, four consumers — versus four independent special cases, which is what
would get written if this is approached site by site.

**Why it is urgent rather than tidy:** "leave rows 41-65 alone" was ADVICE GIVEN TO A
HUMAN. Zettel `20260719-telling-an-agent-not-to-do-something-isnt-a-control` — and it is
not a control when you tell a person either. At work, with nobody to say it, the next run
of Discover Fields tags 21 timeline shapes as individual fields and quietly destroys the
device.

**It folds into the template pivot.** In a template-owned world the device is PART of the
template, so it is declared by construction rather than looked up — the registry stops
being a side-table and becomes a property of the thing already controlled. The expensive
half, knowing what a device is, is already written.

### 3. ALSO REQUESTED, NOT STARTED

**Logical tab numbering and Excel best practice across all workbooks.** `register-wide.xlsx`
has 18 sheets in arrival order with no scheme (`START HERE`, `Sources`, `SRC_EXTRACTS`,
`Field Spec`, seven `TPL_*`, `Register`, `Run Log`, `Sync Log`, `Field Discovery`,
`Template Audit`, a stale `Review project-status-2D3D`, plus the live review sheet).
`WorkbookBridge.ArrangeTabs` already orders them, so the scheme belongs there rather than
in a manual pass. Do this FIRST next session — it is small, bounded, and was explicitly
asked for.

---

## EXCEL TAB ORDER — DONE BY POSITION, DEFERRED BY NAME (13 Aug, late)

Rohan asked for logical tab numbering and Excel best practice across the workbooks, then
added: **"anything that threatens the data chain fix it"** and **"if you are going to
renumber use logic"**. Both shaped what was and was not done.

### Done: a data-chain fix found while auditing the sheet names

**`"Sync Log"` was a bare literal in SEVEN places across two modules** — alone among the
tool-owned sheets, every one of which otherwise has a constant. Two of the seven are
`GetOrAddWorksheet` calls, which **create** the sheet when the name does not match. So one
divergent literal would not fail loudly: it would quietly open a second log sheet while
`IsToolOwnedSheet` and `ArrangeTabs` went on guarding the first, splitting the audit trail
while looking healthy. Now `WorkbookBridge.SYNC_LOG_SHEET_NAME`, all seven replaced.

### Done: tab order follows the lifecycle of a quarter

`ArrangeTabs` now orders by position — no renaming, so nothing can break a lookup:

```
 1  START HERE     where a person begins; the readiness checklist
 2  Field Spec     what fields exist at all -- configuration before data
 3  Sources        the evidence values may cite
 4  SRC_*          harvested source data
 5  Register       THE DATA. The reason the workbook exists.
 6  TPL_*          where a person works, in Field Spec order
 7  Review *       the approval gate, between work and the deck
 8  Field Discovery, Template Audit    diagnostics, off the normal path
 9  Run Log, Sync Log                  the audit trail
10  SAVED *        parked archives, absolutely last
```

**`Register` was previously unplaced** — it fell into "everything else in its current
order" beside the diagnostics, so the most important sheet in the workbook sat wherever it
happened to land. That, rather than any cosmetic gain, is what this fixes. `SAVED *`
archives now sort last so they cannot be mistaken for the live sheet they were copied from;
typing in one is silent, because publish reads the live sheet only.

### Compile-verified 13 Aug, after Rohan closed Office

```
COMPILE OK: whole project compiled clean (33 modules).
RESULT: OK
=== 192 passed, 0 failed ===
```

Static checks also clean across 34 modules. The new cross-module references
(`DiscoverUI.DISCOVERY_SHEET_NAME` and `TemplateAudit.AUDIT_SHEET_NAME` read from
`WorkbookBridge`) compile.

**NOT YET IN AN ADD-IN.** The ordering and the `SYNC_LOG_SHEET_NAME` constant are in source
only at the time of writing — `addin81` predated them. **`addin82` has since been built,
installed and confirmed loaded**, so the ordering is live in code; it takes visible effect
on the next completed drafting rebuild, since `ArrangeTabs` runs during one.

### Deferred deliberately: numbering the NAMES

`01_FIELD_SPEC`-style names would be better still, and are a **migration, not a tidy-up**:
sheet names are this tool's addressing mechanism — nine constants, dozens of literals, plus
the `TPL_`, `Review ` and `SAVED ` prefix matches. It needs the constants changed, every
literal found, and a rename pass over a live workbook holding drafted work, with a fallback
for a workbook that has not been migrated yet.

**If it is done, the scheme should be the ordering above**, so position and name agree and
neither can drift from the other. Do it as its own session with the suite green before and
after — not alongside anything else.

---

## REQUESTED 13 Aug: a standard placeholder for projects that have stopped reporting

Rohan: *"if it's missing because they have stopped reporting please include a standard
placeholder for now."*

**The need is real.** `3_P001` is `Project Closed` and its `PROGRESS_BODY` drafting row is
`SUBMIT` empty / `APPROVED = '0'`. A closed project should not silently render last
quarter's prose or an empty box.

**Proposed standard wording:**

> `No update this quarter — project closed Q3F26. Last reported Q1F26.`

States that the absence is intentional, why, and where the last real content is.

**THE TRAP, and why this was not just typed in.** A placeholder typed into `SUBMIT`
publishes as ordinary content and is then indistinguishable from drafted prose — next
quarter nobody can tell which rows are real. That is exactly the `TESTFILL-1256` string
found in the Q4F26 register tonight, which was one bundled "Yes" from reaching a slide.

**So it should be DERIVED, not typed.** The register already knows `PROJECT_STATUS` and the
period. A placeholder generated at publish time for rows where status is closed/not-started
AND the field is empty is:
- consistent by construction, so the wording cannot drift between projects;
- distinguishable from human prose, so a later review can find every one;
- self-correcting — the moment real text is drafted it takes precedence.

Typing it into 43 drafting rows gets the same pixels this quarter and a mess next quarter.

**Open decisions for Rohan:**
1. Which statuses trigger it — `Project Closed` only, or `Not Started` too?
2. Does it apply to every prose field, or only the quarterly ones (`PROGRESS_BODY`,
   `KEY_EVENTS_BODY`) and not entity-static ones like `ABOUT_BODY`?
3. Exact wording, including whether to name the closing quarter dynamically.

**Do not bulk-type placeholders into the drafting sheets before deciding 1–3.** Undoing 43
rows of typed placeholder is harder than generating it once.

### ANSWERED 13 Aug — the placeholder rule

Rohan on the three open decisions: *"1) whatever makes MECE logical sense 2) only the
quarterly ones, good instinct 3) not sure it should be one of a set selection yeah? like
validated data?"*

**2 is settled: quarterly fields only** (`PROGRESS_BODY`, `KEY_EVENTS_BODY`). Entity-static
fields are excluded — a closed project still *is* what it always was, so a placeholder on
`ABOUT_BODY` would be false.

**3 is settled: a validated set, not free text.** The wording is SELECTED, never typed —
same mechanism as the Controlled fields and `ApplyControlledValidation`. Selecting it is
the consent, and it cannot drift across 43 projects.

**1, the MECE partition. It is over `status x has-text`, not status alone** — that is what
makes it collectively exhaustive:

| project state | field empty | treatment |
|---|---|---|
| Closed | yes | placeholder — no further update will ever come |
| Not yet commenced | yes | placeholder — none due yet |
| Active, explicitly marked nothing-to-report | yes | placeholder — a deliberate statement |
| **Active, no explicit mark** | yes | **NOT placeholdered. Reported as a GAP.** |
| any | no | real text always wins |

**Row four is the whole point.** An active project with an empty quarterly field is a
MISSING UPDATE, not a non-update. Auto-placeholdering it would silently convert "chase this
up" into a tidy sentence on a funder-facing slide — the same shape as every
reports-success-without-confirming-the-effect defect in this file, pointed at content
instead of code.

So the vocabulary needs a fourth member that a person CHOOSES, e.g. `Nothing to report`.
Its presence is what authorises a placeholder on an active project; its absence is a gap
the run should name.

**Prerequisite: normalise `PROJECT_STATUS` first.** The vocabulary is currently inconsistent
(`In progress`/`In Progress`, `Not started`/`Not Started`, plus `Not yet commenced`) — see
"Five status values differ by one capital letter". A MECE rule keyed on status cannot be
built on a status set that has near-duplicate members. Fix that before implementing this.
