# Fix list

> **CURRENT — the live list of what is known-broken and not yet fixed.** Re-audited
> against the code 2026-08-14; six entries added 2026-08-15; A and E since fixed; S
> added and fixed 2026-08-16; T added 2026-08-16 (late evening).
> Entries say whether they are still live; anything marked fixed names the build it was
> fixed in. U added 2026-08-16 (night); V added and FIXED same night, addin111.
> W added and FIXED 2026-08-17 morning. X added 2026-08-17 morning, still open.
> Y added and FIXED 2026-08-17 midday (`ShapeAddressBook.bas`), deployed `addin120`.
> Z added and FIXED 2026-08-17 afternoon (`Timing.bas`, running-long checkpoint in
> `ApplyApproved`), deployed `addin121`.
> AA added 2026-08-17 afternoon, still open -- long unmeasured delay before Resolve()'s
> period-rollover dialog even appears. AB added and FIXED same afternoon --
> BuildLobbyFromScratch's O(n^2) pin-scanning was 75% of a real run's cost. DEPLOYED
> `addin122`, confirmed loaded live, and independently re-measured isolated (no
> confounds): 503.4s -> 2.67s, identical output, ~188x. AC added and FIXED
> 2026-08-17 evening -- RefreshDraftingSheets' "(total)" Timing row now fires after
> WriteRunLog+SaveWorkbookVerified, the true end of the function, and every
> previously-unattributed stage (spec/sources write, missing-columns check,
> validation, register read, tab/index/format, the three FieldSpec validations,
> WriteRunLog+Save) now has its own Timing line -- DEPLOYED `addin123`, confirmed
> loaded live, and confirmed by real data: full run 665.4s -> 284.8s, ~2.3x, with
> every stage now individually visible for the first time. AD added 2026-08-17
> evening, still open -- `WriteDraftingSheet` is ~600 COM calls per field; strategy
> in `DRAFTING-SPEED-STRATEGY.md`, Phase A (bulk read/write) not yet started. AE
> added and FIXED same evening -- the fast-mode wrapper (ScreenUpdating=False,
> Calculation=Manual) only covered the drafting-field loop, so BuildLobbyFromScratch
> (already fixed, item AB) still cost 168.6s live against an isolated 2.67s, and
> the tab/index/format cluster cost 53.7s for simple sheet operations. Widened to
> cover both. DEPLOYED `addin125`, confirmed loaded live, confirmed by real data:
> `RefreshDraftingSheets` 665.4s -> 72.8s overall, `BuildLobbyFromScratch` 168.6s
> -> 6.9s, tabs/index/format 53.7s -> 6.3s. AF added 2026-08-17 evening; FIXED
> PROPERLY (not band-aided) later the same night on Rohan's explicit instruction
> ("proper fundamental fix, every time, don't bandaid unless you're bleeding") --
> `PublishAllDraftedFields` restructured so the per-field loop calls one new
> function, `DraftingUI.PublishOneFieldForChain`, that does exactly one wet-only
> `Drafting.PublishDrafts` call (no separate dry preview) per field, with resolve/
> register-read/save/Run-Log-write hoisted to run ONCE per press instead of once
> per field, and the fast-mode wrapper (now including EnableEvents, see AL) around
> the whole loop. Caught and fixed a real second bug while proving this correct
> (see AF-NTP below). Not yet re-measured live -- addin125 predates this build;
> addin126 pending. AG added 2026-08-17 evening, still open -- `OfferMarkingFor
> UnwiredFields` costs a full register read + full deck shape-walk per press,
> output destroyed before anyone sees it. AH added 2026-08-17 evening, still open
> -- harvest dry-run reads the register up to 3x in one press of button 1. AI
> added 2026-08-17 evening, still open -- `ScanPendingApprovals` computes dead
> detail for a gate deleted this morning, double-reads the review sheet. AJ added
> 2026-08-17 evening, still open -- `SyncNow`/`SyncNowCore` is fully dead code
> containing the worst call pattern found tonight (4x `BuildQueue` per type);
> recommend deletion, matching the bulk-approve precedent. AF-AJ full detail in
> `HOT-PATH-AUDIT.md`. AK added and FIXED 2026-08-17 evening --
> `Readiness.bas`/`WhereAmI`/`WhereAmICore` deleted entirely (Rohan: "delete the
> whole thing"); every check it made was independently redundant with what the
> real operations already catch and explain when actually run, at a cost of two
> full deck-file copies plus a full BuildQueue diff per type, on every single
> press of the tool's most-used button. AL added 2026-08-17 evening; FIXED
> PROPERLY same night, at its root cause rather than only in `ApplyApproved`'s
> wrapper -- `DraftingLobby.FieldIdForSheet` (the pin-watcher's guard, called on
> EVERY cell-change event in EVERY open workbook) had a header comment claiming
> "exits in one comparison" while its actual first line ran a full Field Spec
> scan (`DraftingUI.ProseFields`, ~54-worksheet `WorksheetExists` walk) before any
> comparison happened. Added a zero-COM `StrComp(Left$(sheetName,4),"TPL_",
> vbTextCompare)<>0` pre-check as the true first line, since every drafting
> sheet name is guaranteed to start with that literal prefix
> (`Drafting.DraftSheetNameFor`). `ApplyApproved`'s fast-mode wrapper also now
> disables `EnableEvents` (belt-and-suspenders on top of the root-cause fix).
> AM added 2026-08-17 evening, still open --
> `WriteQueueSheet` (the review-sheet writer) has no fast-mode wrapper at all,
> item AE's own omission unfixed here. AN added 2026-08-17 evening, still open
> -- `UpsertRow` inside the publish loop rescans the whole register per row, per
> field, item AB's shape in the second-hottest chain. AO added 2026-08-17
> evening, still open -- "Add/Retire slides" run the full register-vs-deck diff
> to answer a yes/no membership question. AP, AQ added 2026-08-17 evening,
> found-not-fixed -- real but low-frequency, not worth the risk of touching.
> AL-AQ full detail in `COLD-PATH-AUDIT.md`. **Real measurement, same evening:
> Rohan pressed "2. Put it on the slides" for real on `addin125` --
> `PublishAllDraftedFields (total)`: 362.2s for 4 fields, 90.6 sec/field, worse
> than AF's own estimate. This is AF + AL + AN all stacking on the same real
> path. Clears the standing condition for fixing AF next -- the number came
> back large, on the real button.** AS added and FIXED 2026-08-17/18 night --
`Drafting.PruneParked`'s `.Delete` had no `DisplayAlerts` guard, same defect as
`DraftingLobby.bas`'s (fixed 2026-08-16), never propagated to this second call
site. AT added and FIXED 2026-08-17/18 late night -- `ShapeAddressBook`'s own
cache (built to fix AR) contained items W and AB's defect shape TWICE over
(resolve-from-scratch every call, then a redundant row-scan underneath the
fix), plus a borrowed justification that didn't transfer (persisting misses
to a sheet that never needed to survive a reopen). 508.5ms/call -> flat
regardless of miss volume, proven live (sheet stayed at 17 rows after 300
distinct misses). DEPLOYED `addin133`, confirmed loaded live; two stale
auto-loading predecessors (`addin131`/`addin132`) found and disabled during
deploy.

## Added 2026-08-17 afternoon — AA, STILL OPEN — long delay BEFORE the period dialog
even appears, in code the new Timing instrumentation does not cover

**AA. PRESSING "1. SET UP MY QUARTER" TOOK A VERY LONG TIME TO EVEN SHOW THE
"is this the new period, Q1F27?" CONFIRMATION DIALOG.** Found live during the
first real retest of `addin121` (immediately after the deck-swap prep for that
retest -- see below for why that matters). Rohan, watching it: "it took a
LOOOOOOOOOOOOONG time for that dialogue to arrive" -- explicitly the delay
BEFORE the dialog appeared, not how long he took to answer it once it showed.

**Live evidence, not assumption:** ~2+ minutes of polling (`Get-Process`
`TotalProcessorTime`, matching item X's own method) showed CPU essentially
flat throughout, `Responding=True` the whole time, no dialog visible on
screen when checked -- looked exactly like a hang. It was not: the dialog
was genuinely still coming, just very slowly. **This is a real gap in
tonight's own instrumentation**: `Timing.LogClick` is the very first line
inside `RefreshDraftingSheets`'s body, called right after `Resolve()`
succeeds -- but the period-rollover check and its confirmation prompt live
INSIDE `Resolve()` (or whatever precedes it), before that line is ever
reached. The `Timing` sheet showed zero rows for this run the entire time
it was stuck, which is exactly correct given where the delay actually is,
but looks indistinguishable from a genuine hang without knowing that.

**Confound not yet ruled out:** this happened right after swapping the live
deck for an older `.bak` snapshot (see this session's retest prep) so the
register and the reverted deck had more to reconcile than usual -- possible
the slow part is specifically period/pairing detection working harder
against a genuinely bigger mismatch, not a general cost that happens on
every normal "Set up my quarter" press. Not yet isolated from a normal run
against an already-synced deck.

**Not fixed, not measured.** Next session: instrument `Resolve()`'s own
period-detection path the same way tonight's `RefreshDraftingSheets`/
`ApplyApproved` work was, including a `Timing.LogClick`-equivalent BEFORE
whatever does the slow part, so a real stall there is distinguishable from
this same kind of legitimate-but-slow wait next time.

## Added 2026-08-17 afternoon — FIXED (AB) — `BuildLobbyFromScratch` was O(n^2),
75% of a real run's cost

**AB. FIRST REAL MEASUREMENT OF `RefreshDraftingSheets` EVER TAKEN: 665
SECONDS TOTAL, 503 OF THEM (75%) IN `BuildLobbyFromScratch` ALONE. FIXED.**
From the `addin121` retest, read straight off the `Timing` sheet:

```
13x WriteDraftingSheet     ~40s total  (2.5-4.6s each)
BuildLobbyFromScratch      503.4s      (559 rows scanned, 115 pinned, across 13 fields)
RefreshDraftingSheets (total) 665.4s   (51.2 sec/field average, dominated by the above)
```

**Diagnosed against the actual source**, not just measured: every
`PinToLobby` call goes through `FindLobbyRow`, which calls `LastLobbyRow`
for its bound (a VBA loop reading one cell per row from `LOBBY_FIRST_ROW` --
item W's exact shape) and then its own comparison loop checking 3 cells per
row up to that bound. `PinToLobby` then calls `LastLobbyRow` a SECOND time,
redundantly, if nothing was found. Three full rescans-from-row-1 per pin,
not one. Worse than item W specifically inside `BuildLobbyFromScratch`,
because that function deletes and recreates the Lobby sheet at its own
start -- every one of the 115 pins that run is *provably* new, so
`FindLobbyRow`'s entire comparison loop was 100% guaranteed-wasted work,
summing to roughly 33,000 wasted COM calls across the run.

**Fix, two parts:**
1. `LastLobbyRow` rewritten to `Cells(Rows.Count, COL_L_FIELDID).End(XL_UP).Row`
   -- one native COM call instead of up to n. Fixes the general case for
   every caller, including `AppEvents`' real-time single-pin path.
2. `BuildLobbyFromScratch` no longer calls `PinToLobby`/`FindLobbyRow` at
   all -- it tracks its own `nextRow` counter (valid specifically because it
   just wiped the sheet in this same call) and writes the 5 Lobby columns
   directly. Turns the whole pin-writing part from O(n^2) to strict O(n).
   `PinToLobby` itself is untouched; `AppEvents`' real-time path still gets
   the genuine existing-pin lookup it actually needs.

Full suite green (240/240), including
`DraftingLobby_BuildFromScratchFindsOnlyApprovedRows` and
`DraftingLobby_PinTwiceUpdatesInPlaceNotDuplicate` -- the rewrite produces
identical results to the old code, not just faster ones.

**DEPLOYED (`addin122`), confirmed loaded live, and re-measured in
isolation -- 503.4s -> 2.67s, ~188x, IDENTICAL output** ("115 approved
row(s) pinned, 559 row(s) scanned across 13 field(s)", byte-for-byte the
same string). Measured by calling `DraftingLobby.BuildLobbyFromScratch`
directly via `Application.Run` against the same register workbook, bypassing
`Resolve()` and the whole `RefreshDraftingSheets` chain entirely -- this
was deliberate, not a shortcut: item AA's own unrelated delay was still
live in the same chain, and isolating the ONE thing under test from a
known-slow, unrelated neighbor is what makes 2.67s a real number rather
than a number contaminated by something else already known to be broken.

**Confound now ruled out** (was open above): the deck-swap concern doesn't
apply to this measurement -- BuildLobbyFromScratch reads only the
drafting sheets and their APPROVE columns, never the deck, so the retest
deck's state is irrelevant to this specific number.

## Added 2026-08-17 afternoon, FIXED same evening — AC — the "(total)" Timing row
fired before `RefreshDraftingSheets` actually ended, and ~120s of the same run
was entirely unmeasured

**AC. `DraftingUI.RefreshDraftingSheets`'s `Timing.LogTiming ...,
"RefreshDraftingSheets (total)", ...` call fired BEFORE
`WorkbookBridge.WriteRunLog` and `WorkbookBridge.SaveWorkbookVerified` (the
actual save-to-disk, verified) -- only after both of those did the
completion `MsgBox` appear. FIXED.** Found live: Rohan reported the
completion dialog arrived "about 30 seconds" after the Timing sheet's own
"(total)" row appeared, exactly consistent with watching the wrong signal.

**Fix, folded together with closing item AD's own measurement gap** (its
step 0 -- 665.4s total measured, 503.4s Lobby (now fixed), ~40s drafting
sheets, ~120s entirely unattributed to anything named): every previously
untimed stage inside `RefreshDraftingSheets` now has its own
`Timing.LogTiming` line -- `WriteSpecSheet+WriteSourcesSheet`,
`MissingRegisterColumns check+add` (excluding the dialog wait, same
discipline as the rest of this module), `ApplyPeriodValidation`,
`ReadSheetForDeckPeriod` (with a real row count -- guarded so it does not
crash on the one early-failure path where `Sheet.Rows` is never
initialized, found reading the source before trusting `.Count` on it),
`ArrangeTabs+WriteWorkbookIndex+FormatRegisterSheet`,
`ControlledValidation+BehaviourValidation+RendersValidation`, and
`WriteRunLog+SaveWorkbookVerified`. The "(total)" log itself moved to
after the save -- the true end of the function, matching what a person
watching the screen actually experiences as "done".

Full suite green (240/240). **DEPLOYED `addin123`, confirmed loaded live.**
Real retest with the completed instrumentation: full run 665.4s -> 284.84s
(~2.3x), with every stage individually visible for the first time --
`WriteSpecSheet+WriteSourcesSheet` 5.2s, `ReadSheetForDeckPeriod` 9.3s (43
rows), 13x `WriteDraftingSheet` 23.7s total, `BuildLobbyFromScratch` 168.6s,
`ArrangeTabs+WriteWorkbookIndex+FormatRegisterSheet` 53.7s, validations
1.8s, `WriteRunLog+SaveWorkbookVerified` 8.8s. Residual unattributed gap:
~12.7s (down from ~120s before this fix) -- close enough to call the
attribution problem solved.

## Added 2026-08-17 evening — FIXED (AE) — the fast-mode wrapper only covered
the drafting loop, so the ALREADY-FIXED item AB still cost 168.6s live

**AE. `BuildLobbyFromScratch` (item AB, already fixed to be O(n), proven
2.67s in isolation) STILL COST 168.6s IN THE REAL RUN ABOVE -- 63x SLOWER
THAN THE ISOLATED PROOF, WITH THE SAME CODE.** Found immediately from
reading item AC's own newly-complete Timing sheet: `ArrangeTabs+
WriteWorkbookIndex+FormatRegisterSheet` was ALSO surprisingly slow (53.7s
for simple tab/formatting operations), the same shape of surprise in a
second place.

**Cause, found reading the source, not guessed:** `DraftingUI.
RefreshDraftingSheets`'s `ScreenUpdating=False`/`Calculation=Manual`
wrapper (added earlier tonight, items W/Z's own speed work) is scoped ONLY
to the `WriteDraftingSheet` field loop -- it gets explicitly restored to
normal (`ScreenUpdating=True`, `Calculation=Automatic`) immediately after
`Next i`, BEFORE `BuildLobbyFromScratch` and the tabs/index/format cluster
ever run. Both of those write to the same workbook the loop was just
protecting, for the identical reason -- every one of their now-efficient
writes was still paying full screen-redraw and full automatic dependent-
formula recalculation, on a live 45-slide deck and a 54-sheet register.
This is why the isolated proof of AB (a blank PowerPoint, no deck, no
recalculation pressure) measured 2.67s while the real run measured 168.6s
-- the isolated test was never wrong, it just wasn't exposed to the cost
this fix addresses.

**Fix:** the restore moved from immediately after `Next i` to after the
`ArrangeTabs+WriteWorkbookIndex+FormatRegisterSheet` cluster -- widening
fast mode to cover both. Checked first for the known ScreenUpdating+AutoFit
interaction that can silently miscompute column widths: neither
`DraftingLobby.bas` nor the three `WorkbookBridge` functions in that
cluster use `AutoFit` anywhere, so this widening carries no correctness
risk from that specific quirk.

Full suite green (240/240). **NOT YET DEPLOYED** -- next addin build plus a
live re-run is the actual proof, not the passing suite alone.

## Added and FIXED 2026-08-17 evening — AK — `Readiness.bas`/`WhereAmI` deleted
entirely

**AK. THE "WHERE AM I" STATUS CHECK, RUN QUIETLY ON EVERY SINGLE PRESS OF
THE TOOL'S MOST-USED BUTTON, DID TWO FULL COPIES OF THE ENTIRE ~49MB DECK
FILE PLUS SLOW SHELL.APPLICATION ZIP EXTRACTION, PLUS A FULL
`ReviewQueue.BuildQueue` DIFF PER REGISTERED SLIDE TYPE -- TO PRODUCE ONE
LINE OF THROWAWAY STATUS TEXT. DELETED.**

Found live while chasing item AA: the "Start a Quarter" period dialog took
~17s to appear once, then ~5+ minutes the very next run with no code
difference between the two -- meaning whatever's slow ISN'T the dialog
logic itself. Traced the real chain (`SyncNowChainCore` -> `WhereAmICore`
-> `Readiness.Build`, called BEFORE `StartQuarter`'s dialog, not inside
`Resolve()` as originally assumed -- that earlier theory, documented in
item AA below, was wrong). `Readiness.Build` calls `DeckRegistry.
PropertyOnDisk` (period) and `WorkbookPathOnDisk` -> `RegistryValueOnDisk`
(workbook path) -- EACH does its own `fso.CopyFile` of the entire deck to a
temp folder, then `Shell.Application` Namespace/CopyHere ZIP extraction
with an async polling wait, to re-verify state from the SAVED FILE BYTES.
Plus, per registered type, a full `ReviewQueue.BuildQueue` diff just to
print a row/slide count.

**Traced through and found almost none of it was novel information.**
`RefreshDraftingSheets`/`ApplyApprovedCore` already independently check for
a missing period or workbook pairing and refuse with a clear message the
moment the real button is pressed. `RollForwardPeriod` already refuses on
its own when a partial/leftover period would collide.
`ReviewQueue.BuildQueue`'s row/slide counts get recomputed by the real sync
moments later anyway. `RunSync.bas` already refuses clearly on a missing
template ("REFUSED: this slide type has no template slide registered").
The one thing that WAS genuinely distinct -- catching a period reported-
as-set but never actually saved, the 2026-08-08 defect class -- is ALSO
already verified at the moment of WRITING, via `DeckRegistry.
SetDeckPeriodVerified` inside `DraftingUI.StartQuarter`, called earlier in
the same chain. Re-verifying it again from disk, moments later in the same
session, was checking something already proven true.

Rohan, on hearing the diagnosis: **"Please get rid of stupid stuff,"** then,
after confirming nothing else independently relies on the mechanism:
**"delete the whole thing, keep anything useful but otherwise get rid of
it."** Nothing was worth relocating -- every check traced back to something
already caught elsewhere with its own message, confirmed by grep before
deleting (e.g. `RunSync.bas`'s own template-slide refusal).

**Fix:** `Readiness.bas` deleted entirely. `RibbonUI.WhereAmI`/
`WhereAmICore` deleted (the quiet chain call, and the dialog version, which
was itself already "NO LONGER A BUTTON TARGET" per its own prior comment --
its only remaining caller was the chain call just removed).
`WorkbookBridge.ArrangeTabs`'s tab-order list no longer references the now-
nonexistent "READY" sheet. `DeckRegistry.bas`'s underlying disk-read
functions (`PropertyOnDisk`, `RegistryValueOnDisk`, `WorkbookPathOnDisk`,
`PeriodOnDisk`) are UNTOUCHED -- confirmed by checking every caller first --
they're genuinely load-bearing elsewhere (`SetWorkbookPathVerified`,
`BatchOnboardFlow`'s onboarding verification), this deletion only removes
`Readiness.Build`'s redundant calls INTO them.

A dead `seenPreview` test variable (declared, set, never asserted) was
found and removed from `TestRunner.bas` in the same pass -- caught a real
self-inflicted compile error from removing its declaration but missing one
usage line; fixed and re-verified before trusting the suite.

Full suite green (240/240). **NOT YET DEPLOYED** -- next addin build is the
real proof.

## Added 2026-08-17 evening — AF-AJ, STILL OPEN — hot-path audit, same shape
as AK found five more times

Commissioned after AK: does the same "machinery that outlived the dialog it
fed" defect exist elsewhere in the two most-pressed chains? Full diagnosis,
evidence, and proposed fix direction for each in `HOT-PATH-AUDIT.md` --
summarized here, not duplicated in full:

- **AF.** `PublishAllDraftedFields` ("2. Put it on the slides", the most-
  pressed button) redoes press-level work 13x, once per field: 2 resolves,
  2 full register reads, 2 verified saves, 2 Run Log writes (each erasing
  the last -- all 26 unobservable) per field, and NO fast-mode wrapper on
  the loop at all. Rough bound: several minutes off every full publish.
  **FIXED PROPERLY, later the same night** -- see "AF + AL, the real fix"
  below, not a summary duplicate of the band-aid this bullet originally
  described.
- **AG.** `OfferMarkingForUnwiredFields` (runs on every "1. Set up my
  quarter" press) costs a full register read + full deck shape-walk --
  its own output gets destroyed by `RefreshDraftingSheets`'s `WriteRunLog`
  later in the SAME press. All cost, zero observable output.
- **AH.** The harvest dry-run gate reads the register up to 3x in one press
  of button 1 (~28s where 9s would do) -- legitimate work, needs sharing
  across sub-steps, not deleting.
- **AI.** `ScanPendingApprovals` computes `sheetNames`/`stamp` detail for
  the Yes/No/Cancel gate deleted this morning (Lobby Phase 3) -- never
  used by its only caller -- and double-reads the review sheet right
  before `ApplyApproved` reads it again anyway.
- **AJ.** `RibbonUI.SyncNow`/`SyncNowCore` is fully dead, unreachable code
  (the toolbar targets `SyncNowChain` only) containing the single worst
  call pattern found tonight -- `ReviewQueue.BuildQueue` FOUR times per
  type in one run. Not a live cost today since nothing can reach it, but
  recommend deletion matching this project's own precedent (the
  bulk-approve removal), not a fix -- nothing needs its pattern repaired
  if it doesn't exist.

**The cross-cutting finding, worth more than any single item:** four of
these six (counting AK) share one shape -- a mechanism built to feed a
human-facing surface survived the deletion of that surface, across TWO
separate dialog-deletion campaigns (2026-08-14's consent-dialog removal,
this morning's Lobby Phase 3 gate removal). Cheap sweep for the class: for
each `MsgBox`/prompt deleted since 2026-08-14, grep for what computed its
inputs and check whether anything else still consumes them.

**Not yet fixed, and not obviously the right thing to fix next.** A
separate needs-vs-build comparison run the same evening found that tonight's
entire session went to sync-speed work while the project's own manual-
baseline memory says that's not the dominant cost of a quarter, and the
stated finish line (a real quarter reviewed, approved and published
UNAIDED) hasn't moved. AF-AJ are real, cheap, low-risk fixes -- but doing
them is still choosing to extend tonight's pattern, not correct it. Full
reasoning in `HOT-PATH-AUDIT.md`'s closing section.

## Added and FIXED 2026-08-17 night — AF + AL, the real fix, not a band-aid

Rohan, once the real 362.2s/4-field number landed: "lets go with the proper
fundamental fix, every time. dont bandaid unless you bleeding" -- then "apply
it across the class in line with best practice." This is that fix, for AF and
AL together (they land on the same loop). AN (`UpsertRow`'s per-row rescan
inside the publish loop, same shape one level deeper) was explicitly left
alone tonight -- shared-function surgery, needs its own session and caution,
not a bandaid-vs-fix question.

**AF, root cause and fix.** `PublishAllDraftedFields`'s loop called two
full-ceremony per-field Subs (`CopyAiDraftsToSubmit`, `PublishDraftsForField`)
13 times, each independently resolving the presentation, reading the whole
register, saving, and rewriting the Run Log (which REPLACES the sheet, so 12
of every 13 writes were destroyed by the next one before ever being seen --
this session's own "machinery that outlived its output" shape, found again).
Fix: a new `Public Function DraftingUI.PublishOneFieldForChain(wb, regWs,
srcWs, fieldId, period) As String` does only the per-field work -- copy AI
drafts to submit, refresh counts, one `Drafting.PublishDrafts(..., dryRun:=
False, ...)` call -- and returns a report string. `PublishAllDraftedFields`
now resolves once, reads the register once, loops calling this new function,
then saves once, writes the Run Log once, and shows once. Made `Public`
(matching the existing `DistinctPinnedFields` precedent) specifically so
`TestRunner.bas` can drive it directly with explicit parameters, instead of
requiring a live `Application.ActivePresentation` the way the old Subs did --
genuinely unit-testable for the first time. The old two Subs are untouched
and still exist, complete and correct, confirmed via grep to be unreachable
from any toolbar button (`CommandBarUI.bas`) or any other caller -- left in
place because deleting known-good dead code is a separate decision from
fixing the loop that used to call it.

**AF-NTP -- a real bug found and fixed while proving AF correct.** The old
per-field path always ran `Drafting.PublishDrafts` TWICE per field: once
`dryRun:=True` to build preview text and answer "was there anything to
publish?", then again `dryRun:=False` to actually write. The new function
runs wet-only -- one call, not two -- and checks `Drafting.NothingToPublish`
against the WET result instead. That exposed a genuine, previously-latent
bug: `NothingToPublish` hardcoded ONLY the dry-run summary phrasing ("0
would be published") as its match string, so it could never match a wet
result's phrasing ("0 published", no "would be"). Every caller before
tonight always passed it dry-mode text, so the gap was never exercised.
Found by the "make it fail once before trusting it" discipline: a new test,
`Test_DraftingUI_PublishOneFieldForChainRunsWetOnly`, deliberately broke the
dry/wet flag first (confirmed the write-detection assertion could actually
fail), reverted it, then a full unfiltered suite run caught this SECOND,
real, unrelated failure on its own. Fixed at the actual source --
`Drafting.NothingToPublish` now checks both phrasings (`Or`) -- verified
safe via grep (only two real callers: the new function, wet-mode; the old
`PublishDraftsForField`, still dry-mode, unaffected) and the existing
correctness test (`Test_Drafting_NothingToPublishReadsTheSameCounts`) only
exercises the dry branch, so widening cannot break it. Suite re-run clean,
241/241.

**AL, root cause and fix.** `AppEvents.cls`'s `mApp_SheetChange` fires on
every sheet-change event in every open workbook (Application-level
`WithEvents`) and gates on `DraftingLobby.FieldIdForSheet(wb, sheetName)`.
That function's own header comment claimed "exits in one comparison," but
its real first line called `DraftingUI.ProseFields(wb)` -- a full Field Spec
sheet read plus a `WorksheetExists` scan across all ~54 worksheets -- before
any name comparison ran at all. Measured at ~100-170 COM calls per event, on
every write path that didn't explicitly disable events. Fix, at the root:
`FieldIdForSheet` now exits in one real comparison first --
`If StrComp(Left$(sheetName, 4), "TPL_", vbTextCompare) <> 0 Then Exit
Function` -- since every drafting sheet name is guaranteed to start with
that literal prefix (`Drafting.DraftSheetNameFor`). `ReviewQueue.
ApplyApproved`'s fast-mode wrapper also now captures/restores
`Application.EnableEvents` (belt-and-suspenders alongside the root-cause
fix, restored on both the normal-completion and error paths, matching the
existing ScreenUpdating/Calculation restore-before-re-raise pattern).

**Not yet re-measured live.** `addin125` (currently loaded) predates all of
this; `addin126` needs to be built and deployed, then "2. Put it on the
slides" pressed for real, before the post-fix number replaces the 362.2s/
4-field pre-fix measurement above.

**UPDATE, same night — the retest attempt was invalid, and found something
bigger instead.** `addin135` was built and deployed with this exact fix and
pressed for real. It ran 14.5+ minutes and was killed. Live diagnosis
(mine, in the moment) drew two wrong inferences from correct observations
("settings back to normal ⇒ the loop finished", "file mtime changed ⇒ the
loop's own save ran") — a cold second opinion (fable, full transcript in
this session) read the saved file's actual bytes and found the real story:
**the publish exited in seconds**, because the register being tested was a
reverted baseline with no `Drafting Lobby` sheet at all — "Nothing is
pinned" fired almost immediately. The 14.5 minutes were spent entirely in
the NEXT stage the button chain falls through to when nothing is pending
approval -- `RibbonUI.ReviewChangesCore` -> `ReviewQueue.BuildQueue` ->
`SyncOperations.PlanRoutineSync` -- a completely different, pre-existing
path with no relationship to AF/AL, never fast-mode-wrapped, never
Timing-instrumented, and therefore invisible until tonight. **AF itself is
not proven regressed — it has still never been genuinely retested live**
(needs "1. Set up my quarter" pressed first, to rebuild the Lobby, before
"2. Put it on the slides" means anything). Fable's math on `WriteRunLog`
also cleared AF of the hypothesis I raised live: the new single-body write
(~150-200 lines) is *cheaper* than the old path's 26 destructive
replace-writes (~1300 cell writes total) -- AF reduced Run Log cost, it
did not add to it.

## Added and FIXED 2026-08-17/18 night — AR, the real dominant cost

**AR. `InjectorFor` CALLS `FindShapeByRoleTag` UP TO FOUR TIMES PER FIELD
(base tag, ".1", ".track", ".rest"), AND A GENUINE MISS WAS NEVER CACHED --
ONLY A HIT WAS.** `ShapeAddressBook.bas` (built earlier tonight, see its own
header) already fixed the common case of a shape that EXISTS: cache its
name, verify-on-read, fall back to a full walk only on drift. But a field
with NO matching shape on a given slide type -- the majority case once a
register carries more populated columns than any one slide type has fields
for (roll-forward/publish work over the last two days inflated exactly this)
-- got no caching at all. Every one of `InjectorFor`'s up-to-four calls
re-walked the whole slide from scratch, every time, for every such field.
Diagnosed cold by fable (full evidence trail: the saved register's own
bytes, `InjectPrimitive.bas:408-465`'s call pattern, `ShapeAddressBook.bas`'s
existing `Record` contract) after my own live diagnosis chased the wrong
function entirely (see AF+AL's update above) -- this is the actual dominant
unmeasured cost in the tool tonight, unrelated to AF/AL, and it gets worse
every time more register columns get filled in.

**Fix:** `ShapeAddressBook.NO_SHAPE_MARKER` + `RecordAbsent(slideType,
fieldId)` -- the negative twin of the existing `Record`/`Lookup`. Same
invariant the positive cache already relies on: `Slide.Duplicate` copies
the template's shapes unchanged, nothing in this codebase adds or renames
one afterward, so "confirmed absent" is exactly as stable as "confirmed
present" -- IF that invariant holds. `InjectPrimitive.FindShapeByRoleTag`
checks the marker first and skips the walk entirely on a confirmed miss;
records one on a genuine zero-match walk, but deliberately NEVER on an
ambiguous 2+ match (caching ambiguity as "absent" would hide a real
problem instead of surfacing it on every call the way it does today).
Fixing the one choke point (`FindShapeByRoleTag`) collapses all four of
`InjectorFor`'s calls automatically -- no change needed to `InjectorFor`
itself.

**Honest limit, stated plainly, not glossed over** (Rohan asked directly
whether this is infallible under strict rule-following -- it is not, and
not for the same reason the positive cache isn't): the positive cache has
a cheap verify-on-read (the tag check) as a real safety net against
PowerPoint's own ordinal-based name resolution silently pointing a cached
name at a different shape. The negative cache has **no verify-on-read at
all** -- checking "is it really still gone?" costs the same full walk this
fix exists to avoid. It is trusted, not verified. If a template shape is
ever added to an individual slide instance by hand outside this tool (never
observed, not prevented by any code), this cache would be silently wrong
with no self-heal, only a manual clear of the `Shape Address Book` sheet.
Not fixed, because no automatic invalidation exists for either cache
direction today, and none was added here on the strength of a case that
has not happened.

**Proven with the "make it fail once" discipline**: new test
`Test_InjectPrimitive_NegativeCacheSkipsTheWalk` deliberately disabled
(`If False And ...`), confirmed it fails with the exact expected message
("still Nothing despite a real matching shape now existing"), reverted.
Also proves ambiguity is never cached as absent. Full suite 242/242.

**Not yet built into an addin or measured live.** `addin135` predates this;
next session (or later tonight): build, deploy, then press "1. Set up my
quarter" FIRST (rebuilds the Lobby), then "2. Put it on the slides" -- that
retest now finally proves both AF and AR for real, on a register in the
state the button actually expects.

## Added and FIXED 2026-08-17/18 night — AS, live during the real Scenario 1 attempt

**AS. `Drafting.PruneParked`'S `.Delete` CALL HAD NO `DisplayAlerts` GUARD --
SAME DEFECT AS `DraftingLobby.bas`'s LOBBY-SHEET DELETE, FIXED 2026-08-16,
NEVER PROPAGATED TO THIS SECOND CALL SITE.** Found live, mid-attempt, during
this project's first-ever real Scenario 1 close: pressing "1. Set up my
quarter" after a period rollover, `WriteDraftingSheet`'s per-field cost
measured 82.0s, 103.4s, 110.5s across the first three fields -- increasing,
not flat, and roughly 12x the entire 13-field total (23.7s) this same run
shape measured earlier the SAME evening (`addin123` retest). A raw,
unsuppressed Excel "permanently delete this sheet?" alert also fired and
had to be clicked through by hand -- confirmed by screenshot, and matching
a previously-unconfirmed defect noted in `SCENARIOS.md`'s 15 Aug run.

**Root cause, confirmed by a second opinion (waste-hound) before fixing --
my own first theory (the sheet-scan itself getting slower) was wrong.**
`WriteDraftingSheet` calls `ParkSheetCopy` on every period-change rollover
(every field, this run), which calls `PruneParked` to cap old backups at
`keepNewest = 2` per field. `PruneParked` does `For Each sh In wb.Sheets`
-- a real, but bounded (well under a second even at real scale) scan --
then `wb.Sheets(oldest).Delete` with only `On Error Resume Next` around it,
which suppresses runtime errors, not Excel's native delete-confirmation
(a wholly separate mechanism gated by `Application.DisplayAlerts`). The
decisive evidence against the scan-cost theory: this exact code path,
same workbook, same session, measured 1.8s/field a few hours earlier --
a scan over a sheet count that grew only modestly since cannot explain a
150x per-field jump. The real driver: most fields had already reached
`keepNewest`'s steady state from repeated past rollovers, so THIS run was
the one where nearly every field's park crossed the threshold and
triggered a genuine delete+alert, blocking on a human click, once per
field -- and that wait time, invisible to `Timing`'s own instrumentation
(no `LogWait` wraps a native OS dialog), got silently counted as
computation.

**Fix, narrow and precedent-matched, not a redesign:** `wb.Application.
DisplayAlerts = False` / `= True` wrapped around the one `.Delete` call,
copying `DraftingLobby.bas:304-306`'s own already-proven pattern verbatim
-- including `wb.Application`, not bare `Application` (a real trap that
codebase's own comment documents having hit once: bare `Application` in
this PowerPoint-hosted VBA project resolves to PowerPoint, not the Excel
instance that actually owns `wb`, so it would silently suppress the wrong
app's alerts while Excel's own dialog still fired).

**Explicitly NOT done, on the second opinion's advice:** no run-scoped
cache for `PruneParked`'s scan. The whole park/prune mechanism is already
slated for deletion once file-per-quarter lands (`SCENARIOS.md`'s "GAP 4
-- THE BIG ONE") -- hardening a scan that isn't the real bottleneck, inside
code that's going away, would be effort spent on the wrong asset.

**Swept for a third instance**, per this project's own "a defect found is
a class, not an instance" rule: grepped every `.Delete` call across
production `vba/*.bas`/`.cls`. Only two sheet/worksheet deletes exist in
the whole codebase -- `DraftingLobby.bas`'s (already fixed) and this one.
Every other `.Delete` call found (data validation, PowerPoint slides,
shape tags, command bars, a custom document property) is a different
object type with no relationship to this specific native alert. Clean.

**Full suite 242/242.** No new automated test written for this fix
specifically -- the thing it prevents is a native OS/Excel modal dialog,
which this project's VBA test harness has no way to observe appearing or
not appearing; the fix is a direct, verbatim copy of an already-tested
pattern from a sibling call site, not new logic. **The real verification
is the live retest**, per the second opinion's own proposed discriminating
check: if `WriteDraftingSheet` costs collapse back toward the ~1.8s/field
baseline with the dialog suppressed, the modal was the driver and this is
done. If costs stay high, the scan/Copy/Delete's own Excel-internal cost
is real after all, and the caching direction becomes the actual priority.
Not yet retested live as of this writing.

## Added and FIXED 2026-08-17 late night — AT, the cache built to fix item AR
## turned out to contain items W and AB's own defect shape

**AT. `ShapeAddressBook`'s LOOKUP/RECORD/RECORDABSENT RE-RESOLVED THE SHEET
FROM SCRATCH ON EVERY CALL, THEN A SECOND REDUNDANT SCAN INSIDE
`FindBookRow`, THEN PERSISTED THE MAJORITY-CASE ANSWER (MISSES) TO A SHEET
THAT NEVER NEEDED TO SURVIVE A REOPEN.** Found via a fable-run `waste-hound`
diagnosis of item X's live freeze (Rohan: "nothing anywhere, looks frozen" --
which turned out to be a different symptom from P1's "dialog behind the
window," not the same bug). The audit traced the freeze to `ReviewChangesCore`
-> `BuildQueue` -> `SyncOperations.PlanRoutineSync` -> `InjectPrimitive`'s
injector family, landing on `ShapeAddressBook.Lookup` itself -- the cache
BUILT to fix item AR contained the exact defect shape items W and AB already
fixed elsewhere in this codebase, inside the module written hours earlier
that same night to fix a different instance of it.

**Round 1 -- resolve-once caching (`mSheet`/`mNextRow`).** Every single
`Lookup`/`Record`/`RecordAbsent` call re-resolved the "Shape Address Book"
sheet via `WorkbookBridge.WorksheetExists` and/or `GetOrAddWorksheet`, each a
`For Each ws In wb.Worksheets` scan of the WHOLE WORKBOOK (~54 sheets on the
real register) -- independent of how big the address book itself is.
**Measured live against the real register (17-row book), cross-project via
`Application.Run`: 508.5ms/call.** Fixed by caching the resolved worksheet
object once per workbook (invalidated only when a genuinely different
workbook is wired in via `SetActiveWorkbook`).

**A measurement-methodology trap found and corrected mid-investigation:**
cross-project `Application.Run` (needed to drive an already-loaded add-in
from outside) adds real overhead on top of substantive COM work, not just
trivial dispatch -- confirmed by an isolated "floor" probe (a guaranteed
early-exit call, ~0ms) that ruled out dispatch cost as the explanation for a
smaller-than-expected improvement. Re-measured the RIGHT way: same-project,
in-process, one cross-project hop total instead of one per call, with a
clean git-stash-based before/after on the SAME real register.
**True apples-to-apples: 346.3ms -> 222.1ms/call after round 1 alone --
real, but far short of what eliminating a ~54-sheet scan should produce.**

**Round 2 -- `FindBookRow` still called `LastBookRow` (its own full
row-by-row scan) on EVERY invocation**, even with the sheet itself cached --
the identical "second redundant scan underneath the fix" shape item AB found
in `PinToLobby`. Fixed by having `FindBookRow` use the already-tracked
`mNextRow - 1` instead of recomputing it.

**Round 3 -- Rohan's own question, not a code review, found the real
remaining cost: "but why is it caching misses?"** The negative cache
(`RecordAbsent`, added for item AR) borrowed the POSITIVE cache's
justification wholesale -- "the template fixes this answer forever, worth
surviving a reopen" -- without re-deriving whether it actually transferred.
It doesn't: a HIT is rare and worth persisting; a MISS is the MAJORITY case
(most fields aren't on most slide types, and each gets recorded per suffix
variant -- base/`.1`/`.track`/`.rest`) and only needs to survive the REST OF
THE CURRENT BUTTON PRESS, not a reopen -- every other row of the same slide
type hits it within the same session regardless of whether it ever touches
disk. Persisting misses anyway is exactly what grew the sheet `FindBookRow`
has to scan. Fixed by moving `RecordAbsent` to an in-memory
`Scripting.Dictionary` (`mNegativeCache`, keyed `slideType|fieldId`,
`vbTextCompare`) -- zero sheet growth, zero Excel COM calls for a cached
miss. The positive cache (`Record`) is unchanged; for hits the original
"worth surviving a reopen" reasoning genuinely holds.

**Proven, not assumed -- the architectural claim directly demonstrated:**
a probe recorded 300 DISTINCT misses via `RecordAbsent`, then re-read the
sheet's actual row extent FRESH FROM THE FILE (not an in-memory counter):
**17 rows before, 17 rows after 300 misses.** The sheet cannot grow from
misses anymore, structurally, not just currently-small-so-not-yet-visible.
100 lookups timed immediately after: 81.6ms/call, same range as round 1+2's
number, confirming no degradation under real-shaped load.

**Combined result: 508.5ms/call (unfixed) -> ~120-220ms/call after rounds
1-2 -> flat regardless of miss volume after round 3**, at the project's own
estimated ~5,000+ Lookup calls per "Put it on the slides" press
(item AF/AR) -- was ~42 minutes worst-case, now structurally bounded by hit
count alone (a dozen-ish rows), not by how many absent fields the register
carries.

**A real hang found and cleared during verification, unrelated to the fix
itself:** the full 242-test suite hung for 5+ minutes once, leaving four
orphaned Excel processes with zero workbooks open. Bisected by calling each
of the three `ShapeAddressBook`-related tests individually via the real
`TestRunner.RunAllTests` entry point (not a guess) -- all three pass cleanly
in isolation, meaning the hang was test-sequencing/environmental (likely
accumulated Office automation state from a very long session), not a defect
in this fix. A clean re-run afterward completed normally, 242/242.

**Swept for the same class**, per this project's own "a defect found is a
class, not an instance" rule: `Record` (the positive-cache write path) still
calls `EnsureSheet`-equivalent (`ResolveSheetForWrite`) which is now cached
the same way as the read path -- no third instance found.

**Full suite 242/242 (twice, clean both times).** `check_vba_static.py`
clean across 38 modules. **DEPLOYED `addin133`, confirmed loaded live.**
Cleanup found and fixed during deployment: `addin131` and `addin132` were
BOTH still set `AutoLoad=True` alongside the new build (from earlier same-
night deploys never having their predecessor unregistered) -- three
versions of the same add-in auto-loading simultaneously, a real risk of
duplicate-module collisions. `addin131`/`addin132` explicitly unloaded and
`AutoLoad` cleared; only `addin133` remains active.

**Not yet re-tested against the real Scenario 1 stall this was diagnosed
from** (item X) -- the actual live "Put it on the slides" retest, with
Timing sheet numbers, is still the next real proof and hasn't happened yet.

## Added 2026-08-17 afternoon — FIXED (Z) — no way to interrupt a long apply run

**Z. A LONG-RUNNING `ApplyApproved` HAD NO WAY TO STOP IT SHORT OF FORCE-CLOSING
BOTH OFFICE APPS.** Directly follows item X's own live demo incident -- the ~2 minute
stall was eventually resolved by killing both processes, losing whatever partial
progress existed. Rohan, right after: "include some cancel lines if target exceeded
at a reasonable point."

**Fix:** `Timing.CheckBudgetAndMaybeCancel`, checked every 10 items inside
`ApplyApproved`'s loop. Silent whenever elapsed time is still inside budget (`items
done * 2 sec/item`, floored at 15s so the ratio isn't noise on a handful of items) --
a normal-speed run never sees a dialog, same "don't reintroduce a modal on the fast
path" discipline as Phase 3's own pre-ticked queue. Only once a run has genuinely
blown past budget does it offer a Yes/No to stop. On stop: the loop exits cleanly,
remaining items are logged to the Sync Log as `"cancelled: user stopped the run
early (running long)"` so nothing has to be guessed at afterward, and -- caught
before shipping, not after -- `MarkConsumed` is skipped on a cancelled run, since it
stamps the WHOLE review sheet consumed and `PendingApprovals` treats a consumed
sheet as nothing-left-to-do; consuming it on a partial run would have silently
discarded the still-approved remaining items on the very run meant to protect them.

**Does NOT fix item X itself.** A between-items check can only ever run between
items -- it cannot interrupt something already blocked inside a single synchronous
COM call, which is exactly what item X's own `Ctrl+Break` test showed was happening.
Item X's root cause remains open.

**While building this, also caught and fixed:**
- `ApplyApproved` had no top-level error handler, so a re-raised item crash (Error
  50290, or the deliberate test fault) skipped the `Timing`-motivated
  `ScreenUpdating`/`Calculation` restore entirely and would have left Excel's screen
  updating off and calculation on manual for the rest of the session. Restore now
  happens immediately before each of the two `Err.Raise itemErrNum` sites.
- `PublishAllDraftedFields`'s final `Timing.LogTiming` call was passing a String
  (the comma-joined field list) into a `Long` parameter position -- a real
  compile-breaking type mismatch that would have broken every "2. Put it on the
  slides" press, not a cosmetic issue. Found re-reading the call site.

Two new tests, one deliberately broken and confirmed to fail with the exact wrong
number before being fixed and confirmed green. Full suite: 240/240. **DEPLOYED,
`addin121`, confirmed loaded live** via `Application.AddIns("addin121").Loaded` AND
by actually calling `Timing.StartClock()` inside the running add-in via
`Application.Run` and getting a real value back.

## Added 2026-08-17 morning — one FIXED (W), one LIVE and unexplained (X)

**W. `ReviewQueue.AppendLogLine` RESCANNED THE ENTIRE SYNC LOG FROM ROW 2 ON EVERY
SINGLE APPEND -- O(n^2) COM CALLS ACROSS A BIG APPLY RUN. FIXED.** Found live, at the
worst possible moment to notice it: Rohan watching the first real-scale Phase 3 apply
(221 items, deliberately built to stress-test the new pre-tick/no-modal path) grind for
several minutes, asking "why is it soooo slow." `AppendLogLine`'s row-finding was
`r = 2 : Do While Cells(r,1) <> "" : r = r + 1 : Loop` -- item 1 scans 1 row, item 221
scans 221 rows, ~24,500 wasted reads total for one run, each one a cross-application COM
call (PowerPoint calling into Excel). This wasn't new tonight -- it's always been there
-- but nothing had ever pushed 221 real diffs through one apply before Phase 3 removed
the friction (manually ticking 221 checkboxes) that was accidentally capping batch size.
Small batches never felt it; this one did.

Fixed by reusing `ExcelOutput.bas`'s own already-proven idiom: `Cells(Rows.Count,
1).End(XL_UP).Row + 1`, a single native "find last used row" COM call instead of a VBA
loop of up to n calls. `XL_UP` as a numeric literal (`-4162`), not the named constant
`xlUp` -- same reason `ExcelOutput.bas`'s own `XL_UP` exists: the name only resolves
inside Excel's own VBA project, and this module is PowerPoint-hosted.

**Measured, not assumed**: isolated timing test (fresh throwaway workbook, no other
Office state involved) -- old approach, 100 appends: 15.4s; 300 appends didn't finish in
2 minutes. New approach, 300 appends: 1.67s. Extrapolated to tonight's real 221-item run,
the old code was spending roughly 75 seconds just finding where to write in the log, on
top of everything else. Full suite green (236/0) before and after.

**X. STILL OPEN, NOT YET INVESTIGATED: Rohan reports the apply run appeared to STALL
while PowerPoint was not the focused/active window, and progressed again once he
clicked back onto it.** A real, first-hand, repeated observation ("I know it does
this"), not a one-off guess -- and mechanically plausible: VBA is single-threaded and
this add-in's macros run inside PowerPoint's own message pump, which Windows can
deprioritize for a background window. If real, this would mean a long-running sync
genuinely cannot be left to run unattended in the background while working in another
window, which matters for real usage. **Not yet verified against real Office** (this
project's own rule) -- the one attempt at a controlled test that night was
inconclusive, because the run also needed the CPU time regardless, and a blocking
MsgBox halts execution whether the window is focused or not (unrelated to this claim,
but easy to mistake for confirming it). Next session: a clean test is a long-running
loop with NO modal in it (so nothing but window focus could explain a stall), left
alone with the window genuinely unfocused, watching whether CPU usage goes flat.
Possible mitigation if confirmed: periodic `DoEvents` calls inside the long loops
(`ApplyApproved`, `BuildQueue`) to yield to the message pump -- not attempted yet,
because the phenomenon itself isn't confirmed.

**UPDATE, same day, live during a real demo:** a genuine ~2+ minute stall was observed
and measured properly this time (a background monitor polling Sync Log row count and
CPU every 20s, not manual glances) -- PowerPoint's own CPU went essentially flat while
Excel's crept up slightly, with `Responding=True` throughout and no visible dialog
(confirmed by Rohan looking directly at the screen). **`Ctrl+Break` sent to PowerPoint
did not interrupt it** -- no VBE break, no change in behaviour. That is itself a real
data point, not a dead end: `Ctrl+Break` interrupts actively-executing VBA statements,
not a VBA thread genuinely blocked inside a synchronous cross-process COM call waiting
for Excel to return. This is consistent with the stall being on the EXCEL side of a
big register scan (likely `ReviewChangesCore`/`BuildQueue` comparing the whole
register against the deck for the `project-progress` type), not a PowerPoint-side
hang and not the window-focus theory this entry originally chased. Closed by killing
both processes (test deck, zero real risk) rather than waiting further. Root cause of
*why* that scan is so slow is still open -- worth a real look with `ShapeAddressBook`'s
speed fix (item Y) actually deployed first, since that may explain a meaningful share
of it on its own.

One place for what is known-broken and not yet fixed, so each new review stops
re-deriving the same findings. Three reviews have now paid to rediscover items that
were already known — that cost is what this file exists to stop.

Ranked by how much real work is destroyed, or wasted, before anyone notices.

## Added 2026-08-17 midday — one FIXED (Y), the real cause behind item X's slowness

**Y. `InjectPrimitive.FindShapeByRoleTag` WALKED EVERY SHAPE ON THE SLIDE (RECURSING
INTO EVERY GROUP) ON EVERY CALL, TWICE PER QUEUED ITEM. FIXED.** Measured live during
tonight's Phase 3 stress test: roughly 4-5 seconds per item on a real 221-item apply
run. Ruled out `Application.ScreenUpdating` first (does not exist on
`PowerPoint.Application` -- confirmed by actually trying it, not assumed from Excel's
API) and window focus second (measured with a controlled isolated test, no stall). The
same operation on a one-shape test slide took under 1ms -- roughly a 5000x gap,
pointing at real slide complexity (many shapes, nested groups) as the actual cost, not
COM overhead in general.

**Fix:** `ShapeAddressBook.bas` (new module) -- a persistent, self-healing cache of
"which shape, by name, answers to this role tag on this slide type." `Slide.Duplicate`
(how every instance is created) copies shape names from the template, and nothing in
this codebase ever renames a shape afterwards, so the same name answers the same
question on every instance of a type, indefinitely. `FindShapeByRoleTag` now tries the
cached name first (`Shapes.Item(name)` -- confirmed against Microsoft's own guidance to
be a genuine fast path, not a marginal one), verifies the candidate's role tag before
trusting it, and only falls back to the full walk on a miss or a mismatch. Every full
walk records what it found, so a cache miss self-heals for next time -- no separate
"discovery" pass anywhere (Rohan, 2026-08-17: "to avoid discovery").

Wired via the same pattern as `DraftingLobby.bas`'s `EnsureWatching`:
`ShapeAddressBook.SetActiveWorkbook` is called from `WorkbookBridge.OpenOrGetWorkbook`,
the one place every path through this add-in already reaches the register workbook --
avoids threading a new `wb` parameter through the whole injector family
(`InjectField`/`InjectProgressVia`/`InjectPictureVia`/`InjectDeviceVia` and every one of
their callers).

**A real, load-bearing PowerPoint quirk found building this, now in `AGENTS.md`:**
auto-generated shape names ("TextBox 1") keep resolving via a type+ordinal fallback
even after the shape is renamed -- confirmed live, and confirmed it survives a real
SaveAs/Close/Reopen, not just a live session. Only an explicit custom name invalidates
cleanly on rename. The first version of this fix's self-healing test used rename as
the drift trigger and could not have failed against genuinely broken code, because the
stale name kept resolving to the right shape anyway, for reasons that had nothing to
do with the cache. Corrected to use a realistic drift (shape deleted and replaced) once
this was understood, then proven to fail against deliberately-broken code before being
trusted.

Full suite green (238/0), static/module-list/doc checks clean. **DEPLOYED same day,
`addin120`, confirmed loaded live** -- built and tested during a live demo alongside
the AppendLogLine fix (item W), then both shipped together in a deliberate
build+deploy pass right after, rather than rushed mid-meeting.

## Added 2026-08-16 (night) — one, FIXED same night in addin111 — R's discovery fix had no matching write fix

**V. THE MILESTONE DEVICE WAS DISCOVERABLE (R) BUT STILL COULD NOT BE WRITTEN THROUGH THE
REVIEW QUEUE — THE ONLY PATH THIS TOOL EVER WRITES A SLIDE THROUGH. FIXED 2026-08-16.**
Found chasing "are we there with shapes yet" live: with real `MS1_LABEL`..`MS3_DONE` data
seeded and ticked `Y` in `Review project-progress-A32C`, pressing "2. Put it on the
slides" reported `DROPPED 3_P001/MILESTONE_TIMELINE -- the register no longer has a
value for this field` — every single time, regardless of the register actually holding
real data (checked directly, it did).

Root cause: `ReviewQueue.ApplyApproved`'s proposed-value lookup (`ReviewQueue.bas:~1328`)
assumes every `FieldID` is a literal register column and requires
`rowValues.Exists(FieldID)` before it will even attempt a write. That is true for
ordinary fields and never true for a device tag — `MILESTONE_TIMELINE`'s data lives
across 21 columns (`MS1_LABEL`..`MS7_DONE`), never one cell named after the device,
which is the exact fact R's own comment already states
(`InjectPrimitive.bas:100-109`). R fixed `SyncOperations.PlanRoutineSync` (the
*discovery* path, used by `BuildQueue`) to ask about device tags via
`InjectPrimitive.DeviceRoleTagsOnSlide` — but `ApplyApproved` is a *different* consumer
of the same `FieldID` list, downstream, and was never updated to match. So R made the
device reachable for preview and completely unreachable for the write that preview was
supposed to lead to — the exact "built for one consumer, unreachable from the next"
shape this project has hit repeatedly (milestone writers themselves, picture injection,
progress bars).

Fixed by making `ApplyApproved` ask the same question `BuildQueue` already does: when
`rowValues.Exists(FieldID)` is false, check `InjectPrimitive.DeviceRoleTagsOnSlide(sld)`
for the tag before giving up, and if found, use the same literal
`"(redrawn from its register columns)"` `BuildQueue` already stores as `ProposedValue`
(so the drift-protection hash still agrees with what was approved).
`InjectPrimitive.InjectField`'s `INJECTOR_DEVICE` case ignores the `sourceValue`
argument entirely regardless, so this string is never actually written anywhere — it
only has to match itself.

**Proven properly, not just patched:** every downstream function
(`MilestoneDevice.DrawFromRow`, `.DeviceIntegrity`, `InjectPrimitive.InjectorFor`,
`.FindShapeByRoleTag`) was probed directly against the real slide via `Application.Run`
first and each came back clean — isolating the defect specifically to
`ApplyApproved`'s own lookup before touching any code. One transient `Error 50290`
occurred mid-diagnosis with no code difference before/after (backup created, no Sync Log
line written, nothing changed on disk) and did not reproduce on immediate retry — logged
here as a loose end, not chased further, since the deterministic defect it was found
alongside is confirmed fixed and independently verified: `written: 3_P001/
MILESTONE_TIMELINE` in the real Sync Log, and `Kickoff`/`Design complete`/`Trial
underway` confirmed present in the saved `.pptx`'s own `slide1.xml`, read directly, not
through PowerPoint. Built in `addin111` (2026-08-16 19:44).

**CORRECTION, same night, ~21:33: "one-off" does not hold up.** `Error 50290` recurred —
this time during `Drafting.PublishDrafts` re-publishing `ABOUT_BODY` ("Could not
publish."), a completely different code path from the milestone-apply crash it first
appeared in. Two occurrences in two unrelated call sites in one session is a real,
recurring intermittent failure, not a fluke to stop tracking. **Not yet root-caused** —
still no `Err.Description` captured either time (the generic `UnexpectedErrorText`
wrapper reports only the number), and it has not been possible to reproduce on demand.
Next session: before dismissing it again, add `Err.Description`/`Err.Source` capture at
the point it's actually raised (not just the top-level chain handler), and watch for
whether it clusters around long sessions / many consecutive Office automation calls
rather than any specific code path.

**THIRD OCCURRENCE, next session, 2026-08-17 ~00:56, a different call site again.**
During the live Phase 2 proof (LOBBY-DESIGN.md) on `PRESERVED-known-good-20260815-1050`:
"1. Set up my quarter" and "2. Put it on the slides"' Copy/Publish half both completed
cleanly (Lobby correctly held only `ABOUT_BODY`/`PROGRESS_BODY`, register saved) — the
fault fired one stage later, in `PutItOnTheSlidesCore` (the review-queue apply step,
`RibbonUI.bas`), reported as `"Put it on the slides stopped early... Error 50290...
Reported by: VBAProject"`. Three occurrences across three sessions, three different call
sites (milestone-apply, `PublishDrafts`, now `ApplyApproved`'s chain) rules out a
single-function cause — this reads as a genuine intermittent COM/automation fault, not a
logic bug in any one of them.

**FIXED (the diagnostic gap, not the underlying fault) same night, ~01:30.**
`ReviewQueue.ApplyApproved`'s per-item write block (both the dry-probe and the real
`InjectPrimitive.InjectField` calls) now traps locally: captures `Err.Number`/
`Description`/`Source` immediately, writes a `"CRASHED in dry probe/real write: ..."`
line to the Sync Log via `AppendLogLine` BEFORE re-raising (so it survives even if
PowerPoint dies entirely, same reasoning as `AppendLogLine`'s own header), then
re-raises with the specific `EntityKey`/`FieldID` folded into `Err.Source` so the
top-level dialog names the item instead of just "VBAProject". **Proven, not just
compiled**: Office cannot be made to raise 50290 on demand — that unreliability is the
whole reason this bug has gone three sessions unfixed — so a test-only hook
(`ReviewQueue.mTestForceInjectCrash`, gated, never reachable from a button) simulates a
deterministic fault; `Test_ReviewQueue_ApplyApprovedNamesTheItemWhenInjectFieldCrashes`
genuinely failed against the unwrapped code first (nothing propagated, everything
empty — the fixture's own presentation was never saved, so `BackupBeforeWrite` was
exiting the whole function before the write loop ever ran; fixed by saving the test
presentation to a real temp path), then passed once the wrapper was added. Full suite
235/0.

Next session, when this recurs for real: the Sync Log line and the dialog's `Reported
by:` field will finally name which item was mid-write. Root cause is still open --
this only makes the next occurrence diagnosable, it does not explain why Office raises
the fault at all.

**Byproduct: a real gap found in `check_vba_static.py`'s own declaration-order check.**
Building this fix hit the exact bug `AGENTS.md` already documents from 2026-07-30 --
TWICE in one night (`DraftingLobby.mAppEvents` earlier, then
`ReviewQueue.mTestForceInjectCrash` here) — a module-level variable declaration placed
after other procedures, which compiles quietly wrong and reports its error in a
different module. `check_vba_static.py` already had a check for this shape, but its
regex only matched `Type`/`Const`/`Enum` keywords, never a bare `Public foo As Bar` or
`Private WithEvents ...` variable declaration — so it reported "static checks clean"
immediately before both live compile failures tonight, the exact "a check that cannot
fail looks exactly like one that passed" shape. Fixed by adding `VAR_DECL_RE`; proven
by deliberately reintroducing the real defect (moved `mTestForceInjectCrash` back to
its broken position) and confirming the checker now names it, then restoring.

**SEPARATE, STILL OPEN: `TIMELINE_ELAPSED` (the elapsed bar) still fails "changed since
you approved it" even after rounding `CurrentValue` to 2dp (`InjectPrimitive.bas`,
`InjectProgressField`).** Verified the rounding fix is genuinely active (`Current` in the
queue now reads `127.28`, not `127.2757`) and hand-checked the hash algorithm against the
stored `Current`/`Proposed` twice — both times the stored hash IS correctly derivable
from the stored values, so the hash function itself is not the bug. Still failed on
apply regardless. **The precision theory was incomplete or wrong** — paused mid-diagnosis
at Rohan's call, to think about the workflow architecture instead of one more live
debug round. Next session: don't re-assume precision: instead compare the build-time
`Current` against a fresh apply-time dry-run read side by side, in the same run, to see
concretely which one differs and by how much, rather than reasoning about it from two
separate reads taken minutes apart.

**Rohan asked directly, correctly: "what if it's the first thing every time?"** — a fair
challenge, since the one successful button-driven write happened right after a
diagnostic probe had already written to that same slide once, which could have masked a
"fails on a truly untouched group" defect. Tested properly rather than assumed: found
`PRESERVED-known-good-20260815-1050`'s own copy of the deck, confirmed via raw XML it
had never been touched by anything this session (no `Kickoff`, still the original
authored `ABOUT_BODY` text), confirmed its `MILESTONE_TIMELINE` group was genuinely
pristine (37/37 shapes visible, nothing ever toggled), then ran the real chain against
it with zero diagnostic probes in between — a true first-ever write, through the actual
button. **Succeeded cleanly, no crash, no error, no probe priming it first**, verified
again from the saved file's own bytes. The 50290 crash was a one-off, not a
first-write-every-time defect.

## Added 2026-08-16 (night) — one, LIVE, cosmetic — looks like the OLD destructive bug but isn't

**U. "2. PUT IT ON THE SLIDES" VISIBLY FLICKERS THROUGH ALL 13 DRAFTING SHEETS, AND IT
READS AS THE REBUILD-DESTROYS-WORK DEFECT EVEN THOUGH IT ISN'T.** Rohan, live tonight,
after watching it run: *"why are drafting sheets still getting rebuilt like that I can
see clear at work? I thought we'd been over and over this?"* — the exact question this
project has already answered structurally (decision 1, `DOCUMENT-MAP.md`). Checked
directly, not assumed: `ABOUT_BODY`'s `SUBMIT`/`APPROVE` cells (`G10`/`H10`) survived
completely intact across the run. `Drafting.bas:751`'s `ws.Cells.Clear` is correctly
gated to `If Not layoutMatches` only (a genuine layout migration) — it did not fire.
The real cause: `PublishAllDraftedFields`/`RibbonUI.PutItOnTheSlides` steps through all
13 `TPL_*` sheets, one at a time, checking each field's 43 rows for an AI draft to copy
into SUBMIT, without suppressing `Application.ScreenUpdating`. Watching Excel flip
through 13 tabs touching cells in real time LOOKS like the destructive rebuild even
though nothing is being lost — confirmed by the "38 left alone" count staying identical
across repeated runs tonight. Fix: wrap the loop in `ScreenUpdating = False` /
`= True`. Cosmetic, not data loss — but worth fixing precisely because it keeps
re-triggering the same alarm as the real, already-fixed defect.

## Added 2026-08-16 (late evening) — one, LIVE, found during a real successful quarter turn

**T. `Sources.ApplyPeriodValidation` SWALLOWS ITS OWN ERROR NUMBER, THE EXACT PATTERN
ITEM 1 ALREADY NAMES.** Hit live during tonight's Scenario 1 mechanism run (the one that
actually succeeded — 43 rows Q4F26 -> Q1F27, drafting sheets refreshed): the Run Log
recorded `Sources validation: NOT APPLIED -- Excel refused (the workbook may be read-only
or the sheet protected)`, the exact generic text `Sources.bas:255` always emits on this
path. The real cause is thrown away — `rng.Validation.Add` fails inside an
`On Error Resume Next` block (`Sources.bas:244-252`) and only a Boolean survives; `Err.Number`
and `Err.Description` are read, checked, then discarded (`Err.Clear`) without ever being
put in the message. Nothing else in the run was affected — the roll-forward, the drafting
sheet rebuild, and the workbook save all completed and saved correctly regardless. Not yet
reproduced on demand or root-caused (candidates: the `Sources` sheet mid-write from the
same run, or a genuine protection/read-only state) — next session, reproduce it once with
`Err.Number`/`Err.Description` temporarily appended to the message before deciding a fix.

## Added 2026-08-16 (evening) — one, FIXED same day, the register-side twin of P

**S. THE REGISTER HAD THE SAME CLASS OF DEFECT AS P, ON A DIFFERENT MECHANISM AND
APPLICATION. FIXED 2026-08-16.** Prompted directly by "check the register too" after P
closed — not found by re-deriving from scratch, found by asking whether a defect found
on one application generalises to the sibling one doing the analogous job.
`ExcelOutput.WriteDeckReference`/`ReadDeckReference` used `Workbook.CustomDocumentProperties`
— same storage class as the deck's `Presentation.CustomDocumentProperties`, different
application (Excel, not PowerPoint), so tested independently rather than assumed safe by
analogy.

Probed directly (PowerShell's own COM interop cannot introspect
`CustomDocumentProperties`'s type info at all — a `NullReferenceException` in
`ComRuntimeHelpers.GetTypeAttrForTypeInfo`, confirmed twice with different call shapes,
a limitation of that interop layer, not a fact about Excel — so driven from VBA
instead, `vba/tools/ExcelPropertyProbe.bas`): a brand-new property lands, and a
SECOND, different new property in the SAME session also lands — narrower than the
deck's version, which locked up entirely after one write. But RE-WRITING an EXISTING
property never persists, via `.Value =` (`WriteDeckReference`'s actual pattern) or via
Delete+Add (the "safe" pattern the deck side used to rely on). `StampPairing` calls
`WriteDeckReference` on every repoint, so a workbook re-paired to a different deck
after its first stamp would have silently kept reporting the OLD deck's identity
forever — exactly the cross-wiring risk this mechanism exists to close
(`DeckRegistry.bas`'s own header comment).

**Fixed**: moved off `CustomDocumentProperties` onto a cell (A1) on a dedicated,
very-hidden `DeckSyncMeta` sheet, appended at the end — same shape as the deck's fix,
using the plain `Cells`/`Range` writes this whole module already relies on for every
other piece of register data, proven reliable on OneDrive by this project's entire
operating history rather than freshly probed. `ReadDeckReference` falls back to the old
`CustomDocumentProperties` location so workbooks stamped before this change keep
reading correctly until the next repoint moves them over. Static checks clean, suite
230/0. **Proven against the real re-pairing scenario**, not just a same-session repeat:
stamped a real cloud-hosted scratch workbook, closed, reopened (genuinely separate
session), read back correct; re-stamped with a DIFFERENT value, closed, reopened again,
read back the NEW value, not stuck on the old one.

## Added 2026-08-15 (late evening) — one, LIVE, and it is bigger than Q

**R. THE MILESTONE DEVICE IS UNREACHABLE THROUGH THE NORMAL SYNC LOOP — Q FIXED THE
WRITERS, THIS IS WHY NOTHING EVER CALLS THEM.** Confirmed from the shape's own tag, not
assumed: slide 1's `MILESTONE_TIMELINE` group carries exactly one tag, `ROLE =
MILESTONE_TIMELINE`. The register has **no column named `MILESTONE_TIMELINE`** — only the
21 individual `MS1_LABEL`/`MS1_DATE`/`MS1_DONE` … `MS7_*` columns.

`ReviewQueue.BuildQueue` → `SyncOperations.PlanRoutineSync` (`SyncOperations.bas:152`)
walks `For Each fieldName In rowValues.Keys` — the register's OWN column headers — and
calls `InjectPrimitive.InjectField(sld, fieldName, ...)` once per column. `InjectorFor`
(`InjectPrimitive.bas:286`) only routes to `INJECTOR_DEVICE` when `FindShapeByRoleTag(sld,
identityTag)` finds a group tagged with that exact identity. Since `"MILESTONE_TIMELINE"`
never appears in `rowValues.Keys` — it isn't a register column — `InjectField` is never
once called with that identity tag, on any slide, ever. Not routed wrong: **never called.**

**Reproduced live, 2026-08-15:** seeded real test data into `3_P001`'s `MS1-3_LABEL/DATE/
DONE` columns in the rig register, saved, ran `RibbonUI.PutItOnTheSlides` (the actual
button macro) against the real deck. It completed cleanly, reported nothing to change, and
slide 1's circles were unchanged — confirmed both from the deck's unmoved mtime and by
reading the shapes' `.Visible`/text directly off the slide afterward.

**This is why FIX-LIST Q, real as it is, hasn't been observed to matter in practice**: the
writers Q fixed have never been exercised by an ordinary sync, because the ordinary sync
never reaches them. Same shape as the picture injector, the progress bars, and
`CreateMissingSlides` — a tested function nobody could reach — just found on the WRITE
side this time instead of the read/harvest side.

**Not a quick patch — a real design call**, deliberately not made tonight: either (a) give
the register a literal `MILESTONE_TIMELINE` column so the ordinary loop has something to
iterate that routes to the device (cheapest, but a fake column with no real per-project
value, existing only to be a routing trigger), or (b) have `PlanRoutineSync`/`BuildQueue`
separately scan each slide's ROLE tags for anything routing to `INJECTOR_DEVICE` and
handle those outside the column-driven loop (more correct, touches more code). Whoever
picks this up: read `SyncOperations.bas:140-186` and `InjectPrimitive.bas:286-342` first —
both are short and the whole mechanism is visible in them.

## Added 2026-08-15 (evening) — one, LIVE

**Q. FIXED 2026-08-15 (evening).** THE MILESTONE DEVICE'S TWO WRITERS CANNOT FAIL VISIBLY.
`MilestoneDevice.SetVisible:627` and `MilestoneDevice.WriteText:634` each suppress the
error around their only write, restore handling, and return nothing:

```vba
Private Sub SetVisible(shp As Object, show As Boolean)
    If shp Is Nothing Then Exit Sub
    On Error Resume Next
    shp.Visible = IIf(show, msoTrue, msoFalse)
    On Error GoTo 0
End Sub
```

No return value, no postcondition, no report. **The caller cannot know whether the write
happened**, so `DrawFromRow` can report a milestone drawn against a shape that never
changed. This is the project's signature defect — *reports success without confirming the
effect* — sitting in the writers for **21 of the 29 `Given` fields**, i.e. the largest
stopped data flow in the tool. It has not bitten yet only because the device has never
had data to draw.

**Why it deserves the postcondition treatment specifically:** `shp.Visible` is where
milestone DONE state lives, and PowerPoint has already been measured silently overriding a
property write on this project (`LockAspectRatio`, 2026-08-10). A suppressed `.Visible`
write that PowerPoint declines is invisible by construction — the exact case
`SlideDuplication.bas:115` and `TemplateSlide.bas:122` already guard with an explicit
postcondition and a comment saying the guard can genuinely fail. Those two are the model
to copy; these two are the same shape without the guard.

**Fix, done:** both are now `Function`s returning `Boolean`, confirmed by reading the
property back rather than trusting the assignment did not raise. `DrawMilestones` tracks
a per-slot outcome across all 8 writes for an achieved slot (5 for a hidden one) and
reports through the existing `NoteOnce`/`.Detail` mechanism — no new plumbing, same shape
every other note on this device already uses. A failed write no longer changes the slot
count: it still reports Drawn/Hidden, but now also names which slot had a write that did
not take.

**Proven by a deliberate break**, real slide, real shapes: swapped a real slot's label
shape for a `Line` (genuinely no `TextFrame` — checked via `HasTextFrame`, not assumed;
`AddShape` ovals actually do carry a text frame in modern PowerPoint, which would have
been a false assumption here) and called `DrawMilestones` through its real public API.
Before the fix: silent, `Drawn = 3`, nothing in the report. After: `Drawn = 3` still, but
`.Detail` names the slot and says a write did not take. Reverting the fix (stubbing
`WriteText` back to unconditional success) failed exactly that one assertion and no
others — the other two slots' writes were confirmed unaffected. Suite 209/0, compile
clean, static clean.

**How it was found, because the method generalises.** Applying an "Invisible Failure"
audit — an error suppressed and then never tested for by ANY means. Three iterations were
needed and the first two were wrong: 68 sites when the criterion was only `Err.Number`
(no criterion at all — plenty of code suppresses the call and tests the RESULT), then 40
when the window was six lines (sampled two, BOTH were handled by a postcondition further
down), then **24 when the scope became the rest of the procedure**, which is this
codebase's actual idiom. Sampled two of those 24 and both were real. **The other 22 are
unaudited candidates, not findings** — do not act on them without opening each one. Script
kept at `scratchpad/find_invisible_failure.py`; it is NOT in `check_vba_static.py`
deliberately, because a check that cries wolf gets switched off.

## Added 2026-08-15 (midday) — one, LIVE, and it is the top of the list

**P. `SaveAs`-TO-SELF POISONS A CLOUD-HOSTED DECK, AND IT IS THE ESCALATION WE WROTE TO
RESCUE A FAILED WRITE.** Measured on a scratch deck in `OneDrive\Claude\
onedrive-write-probe\`, 2026-08-15 midday, all three phases read back from the saved file:

| phase | result |
|---|---|
| plain `pres.Save` only, cloud-hosted | **3 of 4 landed** |
| one `pres.SaveAs path, 24` | **raised `0x80CD1001`** |
| plain `pres.Save` only, after that SaveAs | **0 of 4** — `"This presentation is read-only and must be saved with a different name."` |

So on a OneDrive-hosted deck the SaveAs leaves the OPEN PRESENTATION FLAGGED READ-ONLY,
and every later save fails for the life of that document. `DeckRegistry` escalates to
exactly that call at three sites — `SaveDeckVerified:807`, `SetDeckPeriodVerified:869`,
`SetWorkbookPathVerified:940` — each the moment its first read-back does not confirm. A
cloud save lands a beat AFTER the call returns, so the read is simply too early; the
escalation then converts a working save into a permanently broken document, and the
retry loop spends its remaining attempts against a presentation it has already bricked.

**This is the whole of "nothing works on OneDrive".** Plain `Save` works there. Four
theories died against measurement first: file size (a 32KB deck fails identically),
AutoSave (`AutoSaveOn` is settable and makes no difference either way), sync latency (two
minutes plus a close, never arrived), and URL translation (`LocalPathForUrl` maps
`https://d.docs.live.net/...` to the local file correctly — both reads agree).

**FIXED 2026-08-15 (afternoon), the destructive half — `addin101`.** All three sites now
branch on `IsUrl(path)` and never escalate on a cloud deck; they settle and re-read
instead. **Proven against the demonstrated red:** the old build bricked the document into
read-only on the first failed attempt; the new one is healthy after every failure, in five
separate runs.

**AND THE REAL ROOT CAUSE, WHICH WAS NOT SaveAs.** `PropertyOnDisk` took `deckPath`
**ByRef** and reassigned it (`deckPath = mapped`) during URL translation — so *reading the
file rewrote the caller's variable* from the `https://` URL to the local path. Two
consequences, both severe: `If IsUrl(path)` a line later was always False on a cloud deck,
so the new branch could never fire; and the pre-existing `pres.SaveAs path, 24` was handed
the LOCAL path for a document PowerPoint had open from the URL. **A SaveAs to a different
location than the document lives at is what detaches it and leaves it read-only.** Now
`ByVal`. One word, and it is the difference between the tool damaging decks and not.

**STILL OPEN: cloud persistence is INTERMITTENT and uncharacterised.** On fresh
cloud-hosted decks, identical property writes land roughly half the time. Eight hypotheses
were tested and eliminated — file size, AutoSave, sync latency, URL translation, fixture
poisoning, aggressive polling, wrapper-vs-direct, and the dirty flag (the property write
does mark the deck dirty, `-1 -> 0`). No discriminator left that costs minutes rather than
hours. **Do not re-derive these eight.** The settle window was raised to 30s in 5s steps,
which is the right remedy for an intermittent fault now that retrying is no longer
destructive; the 5s step also matters, because every re-read copies the whole package and
polling a 49MB deck once a second copies half a gigabyte to answer one question.

**SCOPE, so this is not read as worse than it is.** The affected surface is FOUR custom
document properties, all setup writes: `DeckSyncPeriod`, `DeckSyncWorkbookPath`,
`DeckSyncType:<type>`, `DeckSyncId`. Slide CONTENT is unaffected — text and tags wrote
fine to a cloud deck in the same session. The register is Excel, a different application
and code path, and has lived in `OneDrive\Claude\` all project. **Scenario 1 is the
exposed one**, because Start a Quarter writes the period.

**Operational note until it is fixed:** a run that hits this leaves the deck read-only, so
anything done afterwards in the same PowerPoint session silently fails to save too.
Closing and reopening the deck clears it.

**Supersedes the recorded remedy.** `NEXT-SESSION`'s "turning AutoSave ON made the write
land" is wrong and was corrected here; `SetDeckPeriodVerified:844-856`'s comment block
diagnoses the same failure as "SaveAs returns without raising and writes nothing" and
prescribes the escalation that causes it. Both were written from symptoms without
probing the mechanism.

**P. THE "STILL OPEN" ABOVE WAS ITSELF A MEASUREMENT ARTIFACT. FIXED 2026-08-16.**
The "cloud persistence is INTERMITTENT" finding above was real, but the fix chosen for
it — wait passively up to 30s, never escalate to `SaveAs` on a cloud deck — was wrong
for the same reason the original bricking was: the midday measurement that ruled
SaveAs out was taken **before** this same entry's afternoon `ByRef`→`ByVal` fix.
Before that fix, `PeriodOnDisk(path)` (called right before the escalation decision)
silently rewrote the caller's `path` from the `https://` URL to a local mapped path —
so the `SaveAs` that got measured bricking a cloud deck was actually "save this
cloud-open document to a *different* local path," not "save it to itself." Nobody had
re-run the three-phase measurement since the fix.

Re-measured 2026-08-16 with an isolated, self-contained probe module
(`vba/tools/SaveAsSelfProbe.bas` + `vba/tools/onedrive_saveas_self_probe.ps1`, each
trial on its own fresh scratch OneDrive deck): **`pres.SaveAs path, 24` (to self,
unmangled URL) landed 5/5, each in under a second, none flagged read-only
afterward** — verified both by the copy-and-unzip read `PropertyOnDisk` already uses
(independent of PowerPoint's object cache) and cross-checked separately via .NET's
own zip reader. Contrast: the SAME production `SetDeckPeriodVerified`, unmodified,
run head-to-head via `vba/tools/onedrive_write_probe.ps1` immediately beforehand,
**failed 4 for 4**, each exhausting the full 4-attempt/30s-wait budget (~121s) with
nothing landing.

**Fixed, PARTIALLY**: `SaveDeckVerified`, `SetDeckPeriodVerified`, and
`SetWorkbookPathVerified` in `DeckRegistry.bas` now escalate to `SaveAs`-to-self on a
cloud-hosted deck exactly as they already did on a local one — the `IsUrl(path)`
branch that split the two is gone, along with the now-dead
`WaitForFileToMove`/`WaitForPropertyOnDisk`/`PauseSeconds` helpers and the
`SETTLE_SECONDS`/`SETTLE_STEP_SECONDS` constants. Static checks clean, suite 230/0.

**REBUILT AND RE-PROVEN ON THE REAL ADD-IN, 2026-08-16 — AND IT SURFACED A DEEPER,
STRUCTURAL LIMIT THE FIX DOES NOT REACH.** Rebuilt `addin108`, re-ran
`onedrive_write_probe.ps1` (unmodified) against the real production
`SetDeckPeriodVerified`: **8 for 8 failed**, but *instantly* (~0.4s each, not the old
~121s) — a different signature entirely, because the fix removed the wait but the
underlying write genuinely still isn't landing.

Isolated with a new probe (`onedrive_reused_file_probe.ps1`, reusing ONE open
cloud-hosted file across 8 trials instead of a fresh file per trial, unlike the 5/5
success above): **trial 1 landed; trials 2–8 all failed, permanently stuck reporting
trial 1's value.** A follow-up scope check
(`SaveAsSelfProbe.ProbeScopeCheck`) shows this is not per-property-name: writing a
**second, never-before-used** property name in the same session, right after the
first property's write already landed, **also failed on its first attempt**. So the
real shape is: **the first custom-document-property write a session ever lands on a
cloud-hosted deck is the LAST one that ever will, for the life of that open file** —
not a per-property or a timing thing.

Tried the one documented community workaround (close the presentation and reopen it
from the same path, which several Microsoft Q&A threads describe as forcing a genuine
resync) via `SaveAsSelfProbe.ProbeCloseReopenRescue` — **did not rescue it**, including
with a deliberate 15s wait between close and reopen to rule out a local-cache timing
race. Three real attempts at a rescue, all negative; capped there rather than
continuing to guess.

**This looks like a genuine OneDrive Personal limitation**, not a bug in this
project's code — one community report found via search says plainly "there is no
function to set custom properties to OneDrive files," and another describes VBA
losing reliable access to SharePoint content-type properties after the first sync in
a similar shape. `DeckSyncPeriod`/`DeckSyncWorkbookPath`/`DeckSyncType`/`DeckSyncId`
are exactly the kind of value that needs updating every quarter on the SAME long-lived
cloud deck — which this limitation appears to forbid outright via
`CustomDocumentProperties`, no matter how the save/retry/reopen logic is written.

**What still stands from the fix**: it's strictly better than before (a session's
FIRST setup write, e.g. a brand-new deck's initial onboarding, now lands fast and
reliably instead of failing 4/4 over 2 minutes) and it's real dead-code cleanup. What
it does NOT do: make a QUARTERLY period update land on an existing, already-synced
cloud deck — which is the actual Scenario 1 use case.

**FIXED FOR REAL, 2026-08-16 evening — proven on the real add-in, repeated writes,
independently cross-checked.** Moved `DeckSyncPeriod`/`DeckSyncWorkbookPath`/
`DeckSyncType:<type>`/`DeckSyncTemplate:<type>:<letter>`/`DeckSyncId` off
`CustomDocumentProperties` entirely, onto a dedicated hidden slide (`DeckSyncRegistry`,
found by `Slide.Name`), keyed by shape NAME (an attribute of the slide's own XML, not
a separate metadata part — and shape names over invisible tags is this project's own
established preference). `ReadStringProperty` falls back to the old
`CustomDocumentProperties` location when nothing is found on the registry slide, so
decks registered before this change keep reading correctly; `WriteStringProperty`
only ever writes to the new location, so values migrate the next time anything
legitimately updates them — no separate migration step. New `RegistryValueOnDisk`
verifies from the saved file by scanning `ppt/slides/*.xml` for the shape, the same
technique `PropertyOnDisk` always used for `docProps/custom.xml`, deliberately NOT by
slide part number (PowerPoint renumbers those on save).

**Two real bugs found and fixed while building this, both worth remembering as a
class:**
- `sh.Namespace()` needs a `Variant`, not a bare `String` — the exact defect this
  file's own `PropertyOnDisk` already documents once, reintroduced on `outDir` in the
  new function and caught by the test suite. Checked every `Namespace()` call in the
  repo for the same mistake, not just the one that broke.
- A genuine VBA `""` success return marshals as PowerShell `$null` through
  `Application.Run`/`InvokeMember` — an 8/8-failure report turned out to be a false
  negative from the *test script's* `-eq ""` check, not the fix. Caught only by
  checking the saved file's actual bytes independently of the return value, exactly
  the "the file is the evidence" discipline this project has learned before, now
  applied to a test script's own output rather than to VBA code.

**Proven, in order**: static checks clean, suite 230/0 (local, deterministic). Rebuilt
`addin109`, re-ran `onedrive_write_probe.ps1` (the SAME probe that caught the original
bug) against the real, unmodified `SetDeckPeriodVerified` — **8 for 8 landed**, each
under a second, on ONE reused open cloud file across all 8 trials (the exact scenario
that failed 0/8 against the old mechanism). Independently cross-checked via a
separately-written .NET zip reader scanning `ppt/slides/*.xml` — agrees. This closes
the OneDrive write-reliability gap that blocked Scenario 1 (updating the period on an
existing, already-synced deck every quarter).

## Added 2026-08-15 (late morning) — four, all LIVE

**L. A COPIED DECK KEEPS POINTING AT THE ORIGINAL'S REGISTER, AND THE REPAIR WAS
UNREACHABLE.** `DeckRegistry.GetWorkbookPath:170` returns the stored path unchanged
whenever that path exists; the sibling lookup is a fallback for a MISSING file only. So
the known-good snapshot's deck carries
`DeckSyncWorkbookPath = C:\Users\rohan\OneDrive\Claude\register-wide.xlsx` in its own
`docProps/custom.xml` — read from the file, not the object model — and a "test on the
snapshot" would have read and WRITTEN the live register. `DiscoverUI` would have landed
on the live `Field Discovery` fixture. **Half-fixed:** `RepointWorkbookUI` existed but was
reachable only from the sync path and only when the pairing was EMPTY, never when it was
WRONG. A 5th toolbar button (`CAP_REPOINT_WORKBOOK`) now exposes it — written, statically
proven, **not yet in a built add-in**. The detection half is still open: nothing warns
that a deck is paired to a register that isn't the one beside it.

**M. `PROJECT_PROGRESS` GENERATES A FALSE DIFF THAT WOULD CORRUPT THE SLIDE.** In the live
review queue right now: Current (on the slide) `80%`, Proposed (from the register) `0.8`,
flagged as a change. Approving writes the literal string `0.8` onto the slide. Same
formatted-VIEW-vs-stored-VALUE collision recorded at 00:35 for the harvest direction, now
confirmed in the publish direction. Both cells also carry Excel's "number stored as text"
marker, so `0.8` is not even numeric in the register. **Nothing in the queue is safe to
tick until the comparison is format-aware or refuses this field.**

**N. INVISIBLE-CHARACTER DIFFS ARE REPORTED AS CHANGES WITH NO WAY TO SEE THEM.**
`PROJECT_NAME` rows show Current and Proposed identical word for word; `What differs` says
*"differs from character…"* and is cut off at the column edge. A person cannot triage
what they cannot see. Widen the column, or render the differing character by name
(`<space>`, `<nbsp>`, `<CRLF>`).

**O. `build_ppam.ps1` QUITS POWERPOINT WITH THE USER'S DECK OPEN, UNANNOUNCED. FIXED
2026-08-16 (night).** Line 88 calls `Request-GracefulQuit "PowerPoint.Application"`
before building. It closed the live deck mid-session on 2026-08-15; that time nothing was
lost because AutoSave happened to be on and the deck happened to be saved already. The
"safe path" was believed to be the 30-second timeout an unsaved deck's modal save prompt
would trigger — but that assumption was never actually tested against the failure it was
supposed to catch. **It was wrong.** 2026-08-16: a real, deliberate, unsaved edit (a role
tag on a real shape) was silently discarded — no modal, no timeout, no message — because
`Quit()` runs inside a background job's isolated COM apartment, which does not
necessarily surface a UI prompt to anyone. Fixed properly: the job now explicitly
`Save()`s every open presentation before calling `Quit()`, rather than trusting a prompt
that may never appear. Excel is still not in the loop and still never touched.

---

## 1. Excel's real error is thrown away, and the message sends you to the wrong file

**Found 2026-08-09, live, mid-session.** `1. Drafting Sheets` reported:

> Could not open the paired workbook at: `C:\Users\rohan\deck-sync-e2e\register-wide.xlsx`

The file was perfectly good — Excel opened it directly, 13 sheets. The actual cause was
that a **different** `register-wide.xlsx` was already open from another folder, and
Excel refuses two workbooks with the same filename at once. Nothing in the message
said so, so the first five minutes went into "did we corrupt the file?"

`WorkbookBridge.OpenOrGetWorkbook` wraps the open in `On Error Resume Next` and returns
`Nothing`. Every caller can then only say "could not open". Excel supplies a specific,
actionable reason — name clash, lock, permissions, cloud path — and all of it is
discarded at the moment it is generated.

**Fix:** capture `Err.Description` from the failed `Workbooks.Open` and return it, so
callers report *why*. Same shape as the publish save that reported `Err.Number = 0` as
success: the diagnosis existed and was thrown away.

**Cost:** small. One out-parameter, five call sites.

---

## 1a. A message that reports a COUNT must name its SUBJECT

**Three instances in one afternoon, 2026-08-09. Each was true, and each was unusable.**

- `Copy AI to Submit`: *"Nothing to copy: there are no AI drafts on this sheet yet.
  Column F is empty for all 43 row(s)."* It acts on whichever `TPL_` sheet is ACTIVE in
  Excel. The active tab was `TPL_ABOUT_BODY`; the work was on
  `TPL_STRATEGIC_ALIGNMENT_BODY`, where column F held 812 characters. The message never
  names the field, so a correct statement about one sheet reads as a flat contradiction
  of what you just did. Its success message has the same hole — *"1 copied"*, of what?
- `Create Template Slide`: *"5 field(s) set to placeholders."* True. The deck had EIGHT
  tagged fields; the template was cloned from a slide carrying only five. Naming them
  would have shown the three missing at a glance instead of after a byte-level check.
- `OpenOrGetWorkbook` (item 1): *"Could not open the paired workbook at <path>"* — names
  the file that is FINE and not the duplicate filename blocking it.

**The shape:** the code holds the identifying detail at the moment it composes the
message, and drops it. What survives is a number with no subject, which is worse than
silence — it reads as authoritative and sends you to check the wrong thing.

**Fix:** every count in a user-facing message names what it counted. `"Nothing to copy
for ABOUT_BODY"`, `"5 of 8 fields set to placeholders -- missing: ..."`, `"...blocked by
register-wide.xlsx already open from <other path>"`. Mechanical, and it is the single
cheapest reduction in time-to-diagnose available in this codebase.

**Worth guarding, not just fixing:** this is the fourth time a message has been true and
unusable. A test that asserts every `MsgBox` composing a count also interpolates a name
would be crude, but it would hold.

---

## 1b. A drafting sheet for a derived field carries nothing to derive from

**Found 2026-08-09, on the first real drafting run of a new field.** Two faults, one
cheap and one structural.

**FIXED (verified in the workbook 2026-08-14).** The global rules now carry an
explicit empty-column-C clause: *"Where column C is EMPTY, this field has never been
published for that project. Write it fresh against the Length above. Do not leave the row
blank on account of an empty column C."* The account below is kept for the reasoning.

**The cheap one -- the global rules assumed a prior value existed.** They say *"Column C
is the standard, not a draft to improve on… if the text in column C already does its
job, say so and leave the row blank."* On a field added today, column C is empty for
every project that has not been harvested — 42 of 43. Read literally, that clause tells
Copilot to leave every row blank. Those rules were written when every field had a prior
value, which was true of `ABOUT_BODY` and is false of every field added from now on.
It is a cell on the Field Spec sheet, so the fix is an edit, not a build.

**The structural one — the sheet shows a drafter only the project's NAME.** The columns
are: code, name, ORIGINAL (empty for a new field), SUBMIT, tick, AI draft, sources,
counts, notes. `STRATEGIC_ALIGNMENT_BODY` is the "so what" for a project, and it cannot
be written from a title. `ABOUT_BODY` — what the project actually IS — exists for all
43 projects, in the register and on its own sheet, and is not visible here.

So for a derived field the honest output is 42 blank rows, and the rules correctly
forbid inventing the rest. That is not the recipe failing; it is the sheet not carrying
what the recipe needs.

**Fix, and it is Field-Spec-shaped rather than code-shaped:** a `Context fields` column
on the Field Spec saying which other fields to show read-only alongside — Strategic
Alignment and Problem both want `ABOUT_BODY`; Progress wants `KEY_EVENTS_BODY`. Then
`WriteDraftingSheet` renders those as extra read-only columns.

**Why this matters more than it looks:** it is the difference between a tool that can
only UPDATE text that already exists and one that can WRITE a field for the first time.
Every new field, and every new project, hits this.

---

## 1c. FIXED 2026-08-14 (`aff84d6`) — the at-risk scan missed SOURCES and NOTES

> **Fixed, not compiled.** The scan now counts SOURCES and NOTES alongside SUBMIT and
> DRAFT, so the rollover refusal covers all typed work. Original entry follows.


**Found 2026-08-13 while rewriting the two tests the refusal guard turned red.**

`WriteDraftingSheet`'s refusal counts a row as at risk only if `COL_D_SUBMIT` or
`COL_D_DRAFT` holds text (`Drafting.bas`, the `If periodChanged Then` block). But the
rebuild carries **four** columns of human work — DRAFT, SUBMIT, SOURCES, NOTES — and the
rollover drop clears all four. So a row where a person has assigned citations or written
notes, but has not drafted yet, does not trip the refusal, and its work is discarded.

Citations are not incidental: per `project_deck_sync_provenance_is_field_architecture`,
the source IDs are the control on the generative step. Losing them silently is the same
failure the guard was built to stop, one column over.

**Not live on the real workbook right now** — the 129 drafted values sit in columns E and
F, exactly what the scan does cover, so every sheet refuses and everything is protected.
This is the next instance of the class, not an active threat.

**Fix:** count all four columns as at-risk. **Then check whether the per-row cadence
machinery should be deleted rather than fixed** — see below.

**The bigger question this exposes.** Once the refusal covers all four columns, the
rollover drop is reachable only for rows holding *nothing*, where there is nothing to
drop. The whole cadence mechanism — the `cadence` dictionary, `carryThisRow`,
`droppedQuarterly`/`keptStatic`, and the test that covers them — would become dead code
whose most heavily-commented property ("the drop is per row, not per sheet, and that
distinction is the whole of this guard's correctness") no longer decides anything. It was
correct machinery for a design that the refusal replaced. Deleting it is probably right,
and is Rohan's call, not a silent cleanup.

---

## 1d. FIXED 2026-08-14 (`aff84d6`) — by deletion, not repair

> **The late `ParkSheetCopy` call site is gone.** The park now runs before
> `ws.Cells.Clear`, unconditionally, which made this site both unreachable and wrong.
> Note the clear itself is now confined to a layout migration, so on the normal path
> there is nothing to park *from*. Original entry follows.


**Found 2026-08-13, in the code directly beneath the comment warning about this.**

`Drafting.bas`: `ws.Cells.Clear` runs at the top of the rebuild. The `ParkSheetCopy` call
for the rollover case runs in the *reporting* section, hundreds of lines later —
`If lostWithContent > 0 And parkedName = "" ...`. `ParkSheetCopy` does `ws.Copy`, which
copies the sheet **as it is at that moment**: already cleared, already rebuilt, with the
dropped rows' content gone. `ParkedNote` then reports *"The previous sheet was kept as
'<name>' — nothing was lost."*

The archive is real, named, and hidden on the workbook. It just does not contain the
thing it was taken to preserve.

This sits immediately below the comment quoting ReviewQueue's own rule: *"A REPORTED
BACKUP THAT IS NOT ON DISK IS WORSE THAN NO BACKUP: it is the reason you feel safe
running the destructive write that follows."* The layout-mismatch park at the top of the
function is correctly placed, before the clear; only this second call site is wrong —
the same "fixed where it was found, not everywhere the shape exists" pattern the repo has
now logged five times.

**Narrow but reachable:** needs a rollover where a dropped row has NOTES content and no
row anywhere has SUBMIT or DRAFT (otherwise the refusal pre-empts it). Fixing 1c closes
it by making the path unreachable, which is the cleaner fix of the two.

**Fix:** park before `ws.Cells.Clear`, or delete the second call site along with the
cadence machinery if 1c is resolved by deletion.

---

## 2. SUPERSEDED — the button this described no longer exists

> The toolbar had `0`–`4` numbered buttons when this was written; there are now two.
> Kept because the CAUSE is still live: `ContentKindOf` returns `KIND_PROSE` for
> everything outside three hardcoded names, so `HasBatchableWork` is never true for a
> prose field. Under the chain that is simply the normal route to the review sheet.

### As written, against the old toolbar

`AssignBatches` batches only `KIND_CONTROLLED` fields, so for prose —
`ABOUT_BODY`, the field the whole drafting apparatus exists for — `HasBatchableWork` is
never true. The button numbered 4 in a 0–4 sequence therefore **always** ends in a
`vbExclamation` dialog saying it is opening the review sheet instead. The toolbar
presents that as an exception; the code makes it the rule.

And the dialog names `REVIEW_SHEET_NAME` = `"Sync Review"`, while `ReviewSheetNameFor`
produces `Review project-status-3D1B`. The rename's own comment records why it happened:
*"Rohan could not find it and asked where the 'sync review file' was."* The sheet got a
findable name; the sentence pointing at it did not.

**Fix:** when there is nothing batchable *because everything is prose*, that is the
normal route — drop the exclamation, go straight to the review sheet, and name the real
sheet via `ReviewSheetNameFor`. Delete the constant.

---

## 3. The double tick — 86 ticks for 43 pieces of text

You tick `Y` in the drafting sheet's column E, then tick `Y` again against the same 43
paragraphs in the review grid under a different column name.

**Fix:** carry the drafting approval forward. Pre-tick a review row **only** where the
slide's current text is still exactly what the register last wrote — so a slide someone
hand-edited since still arrives blank and demands a read, which is the case the review
grid was built for. Every gate survives: current-vs-proposed is still shown per row,
every row is still untickable, the change hash still drops stale approvals.

**NEEDS ROHAN'S EXPLICIT SIGN-OFF.** It sits adjacent to R13.2 (prose may never be
batch-approved) even though it is not batching. Do not ship on a reviewer's say-so.

---

## 4. `Kind` is answered in two places and they can disagree

`FieldSpec` has a user-editable `Kind (Controlled/Prose/Static)` column. `ReviewQueue.
ContentKindOf` hardcodes the same three values for three field names. The sheet governs
which fields get drafting sheets; the hardcode governs batching. Edit the sheet and
batching does not move — silently.

**Fix:** `ContentKindOf(fieldId, specWs)` reads the sheet, falls back to the built-in
table when there is no row, keeps `Prose` as the unknown default so absence is never
read as permission to batch. Rohan's design, 2026-08-08: the vocabulary comes from the
code as a **dropdown**, the assignment comes from the sheet, and an unrecognised value
is REPORTED rather than silently defaulted.

---

## 5. Two questions asked of a person that nothing reads

`FieldType` (text/number/currency/date) and `FieldVolatility` (static/variable) are
asked once per field during marking, normalised, serialised, round-tripped through the
marking session, written into the review grid — and read by nothing. On the `Setup A`
path that is roughly **104 dialogs that alter nothing**, across ~52 fields.

**Fix:** delete both. Nothing becomes impossible; nothing read them. If field typing is
wanted later, `Kind` on the Field Spec sheet is where it belongs — editable, not a modal
at mark time.

---

## 6. User-facing strings that describe a tool that no longer exists

**Re-audited against the code 2026-08-14** — the list grew rather than shrank. Each entry
below now says whether it is still live.

Fix as a CLASS, with a grep for the shape, not one at a time. The 3de4be8 sheet rename
left four stale readers and they have been found in three separate sessions.

- `FastPathRefusalText` still says sync "does not create slides" — untrue since the
  25% create path landed.
- `IsToolOwnedSheet` still matches the prefix `"Sync Review"`, while
  `ReviewQueue.ReviewSheetNameFor` produces `"Review project-progress-A32C"`. **Re-checked
  2026-08-14: STILL LIVE.** So the review sheet a person is actually working in is not
  recognised as tool-owned, and `START HERE` labels it *"(not created by this tool)"*.
- **`IsToolOwnedSheet` is missing `Run Log` AND `Register`.** Re-checked 2026-08-14: still
  true of both. `LifespanOf` knows `Register` perfectly well two functions away — the two
  disagree about the most important sheet in the workbook.
- `DescribeSheet("Register")` describes the LONG register — one row per project, field
  and quarter, with approval state. That model was retired 2026-08-03.
- **Dead caption constants.** `CommandBarUI` defines 19 `CAP_*` strings; `AddButton` is
  called twice. The other 17 name buttons that do not exist, including the whole
  three-chain design (`1. Start the quarter`, `2. Draft and publish`,
  `3. Put it on the slides`). Any message interpolating one of those is describing a
  button nobody can press — which is the exact failure `TOOLBAR.md` predicted.
- `WORKFLOW.md` still says columns G/I, "no button for roll forward", and "nothing ties
  a source to a period" — all three fixed in code, none in the doc.

---

## 7. Things typed that could be picked — 2 of 3 FIXED 2026-08-13 (addin81/82)

**Status, so this does not read as fully open:** the slide-type picker is **fixed**
(`RibbonUI.PickType` auto-selects a sole type, all three call sites). Roll Forward is
**partly fixed** — it no longer asks at all when the destination already holds rows
(`ExcelOutput.PeriodRowCount`), but when the destination IS empty it still asks for the
source period as free text, so the validated-list fix below still stands. The third item is
**untouched**.


- **`Roll Forward`'s source period** is typed free-hand, into exactly the trap its own
  header warns about for the destination. `Sources.ApplyPeriodValidation` already builds
  this list from the register's own `Quarter` values.
- **The slide-type picker** in `Audit Fields` and `Create Template Slide` asks a question
  with one legal answer on any real deck — `ResolveRegisterSheet` refuses a deck with
  more than one type. Auto-select when there is one; keep the picker for the rest.
- **The second unsaved-workbook prompt** in `ApplyApprovedCore` re-asks after
  `ReviewChangesCore` already saved. Measure once, before the tool dirties anything —
  same fix as the Preview Sync footprint bug.

---

## 8. Three of six context switches exist because the workbook cannot find its deck

Steps 1, 2 and 3 are PowerPoint buttons whose entire effect is in the workbook.
`WORKFLOW.md` says so for each: *"Touches: the workbook only. Never the deck."* They are
PowerPoint-hosted because `DraftingUI` needs `ActivePresentation` for two facts — the
period and the workbook path — and the workbook stores the deck's **opaque ID**, never
its path.

**Fix (own session):** write the deck's full path and current period into the workbook's
custom properties on every `Start a Quarter` and publish, then ship an Excel-hosted
`.xlam` carrying steps 1–3. Same Subs, no forks — `TestRunnerExcel.bas` already proves
the modules run Excel-hosted. Costs a second package to version; saves four alt-tabs per
field per quarter.

---

## Capability gaps — a person cannot do these at all

- **Retire a slide or a project.** Nothing removes a slide the register no longer
  mentions. It shows only as a parity mismatch, resolved by hand.
- **A maintained list of linkage codes.** Strategic Alignment must cite codes it can
  check. **Rohan owns it** (settled; do not re-ask), colleagues later — so it needs an
  `as at` date and a home. **CORRECTED 2026-08-14: that home is a PERMANENT sheet in the
  register workbook, beside `Sources`.** The earlier wording here said "probably outside
  the register workbook, which the tool rebuilds" and that is wrong —
  `WorkbookBridge.LifespanOf` classifies `Register`, `Field Spec` and `Sources` as
  PERMANENT; only `TPL_*` and review sheets are rebuilt. An external file would be a
  second artefact to keep in step, which is the class of problem template-first removes.
- **Provenance that survives a rollover.** Source citations live only on the drafting
  sheet, which is cleared at every period change. The record answering "why does it say
  90%?" has a lifespan of one quarter.
- **The other 42 slides.** Only slide 4 and the template carry the three new panels.
  Extending them is `Setup A2` → `Setup B`, which has never been walked by Rohan
  unaided — and until it is, "is the setup formulaic?" is unanswered.

---

# Loose ends — found in conversation, at risk of being lost

Small, real, and none of them recorded anywhere else as at 2026-08-09.

## Five status values differ by one capital letter

`Q1F27` rows for `3_P001`, `3_P002`, `2_P003`, `2_P004`, `1_P005` hold `'Not started'`
against the vocabulary's `'Not Started'`. Reported on every drafting run as
*"5 value(s) are not in their allowed list. Nothing was changed."* and never chased.

The tool is right not to auto-correct — `ApplyControlledValidation`'s own comment says
so — but the message never names the five, which is fix-list item 1a again. Five cell
edits.

## UNVERIFIED: does `||||` render as a paragraph break on the slide?

`2_P004`'s published Strategic Alignment stores a blank line between paragraphs as
FOUR pipes (`||` is one line break; a blank line is two). Publish encodes it, and
`Drafting` renders `||` back to real breaks in column C — but **nothing has yet
confirmed what reaches the SLIDE.** If it arrives as literal `||||`, every multi-
paragraph field is affected, and it will be visible on 43 slides at once.

**This is the last unverified link in the chain.** Check it the first time anything
syncs to slide 4.

## Duplicate filenames block the add-in, and the machine is full of them

Excel refuses two open workbooks with the same NAME regardless of folder. A review copy
called `register-wide.xlsx` in `OneDrive\Claude\` blocked the live one and surfaced as
*"Could not open the paired workbook"* pointing at the file that was fine.

Also on the machine as at 2026-08-09: `addin52`–`addin56` all in the AddIns folder, plus
stray `addin55.ppam` and `addin56.ppam` in the OneDrive root. Only `56` is needed. Six
near-identical add-ins is how the wrong one gets loaded on a tired evening.

**Rule for review copies from here: never reuse the live filename.**

## One judgement call in the first drafted field, left in deliberately

`2_P004`'s Strategic Alignment contains *"rather than monitoring broadly and intervening
on judgement"*. That edges toward `PROBLEM_BODY`'s territory — the "so what" is hard to
state without a contrast against current practice. Kept, and flagged: if Rohan reads it
as a bleed, the own-job test needs a sharper line, and that is a Field Spec edit rather
than a code change. It is the first live test of whether the boundary wording holds.

## The Run Log fix is not in `addin56`

`addin56` was built before the `NumberFormat = "@"` fix landed. Until the next build,
every Run Log still loses its body — so "read the Run Log" is not usable advice yet.

## Source assets — the state as at 2026-08-09, and what work means

**Sources sheet: 6 rows, every one prefixed `EXAMPLE (not a real document)`.** Invented
2026-08-08 to exercise the citation check. Still there.

**Citations: ZERO, on every field.** This is new today and it is not a cleanup — the 37
fabricated citations on `TPL_ABOUT_BODY` were WIPED when the drafting sheets rebuilt at
the period change, because `cadence` is permanently `Nothing` so every row cleared.
The provenance-dies-at-rollover problem, demonstrating itself on real data. Half the
scaffolding removed itself; the six sheet rows did not.

**Two of the six now point at a period that no longer exists** — `S02` and `S05` are
`Applies to: Q4F26`, and the register holds only `Q3F26` and `Q1F27` after the relabel.

**So "source assets need work" means, concretely:**

1. **Delete `S01`–`S06`.** Nothing cites them now, so it costs nothing — and leaving
   fabricated provenance next to real content is worse than having none.
2. **Build the first real source from the question already asked.** `J13` on
   `TPL_STRATEGIC_ALIGNMENT_BODY` holds it: are `1.4.2, 1.5.2, 2.3.1` DECLARED linkage
   codes, and where does that list live? That answer becomes the first genuine Sources
   row and unblocks the `[TBC]` in `2_P004`'s published text.
3. **Use the form.** `SOURCE-CAPTURE-FORM.md` has the rubric and the fields; it needs no
   tools and can be filled in at work.
4. **Decide where a maintained list lives.** Rohan owns it (settled). It needs an `as at`
   date and a PERMANENT sheet in the register workbook, beside `Sources` — see the
   correction under Capability gaps. The workbook is not rebuilt; only `TPL_*` and review
   sheets are.

**And the structural point stays open:** citations live only on the drafting sheet,
which is cleared at every period change. Until that moves to the register — same grain
and lifespan as the text it explains — provenance has a maximum life of one quarter,
and today proved it empirically rather than theoretically.

---

## PARKED: name each prose field's recipe, so it can be discussed plainly

Rohan, 2026-08-14, and parked by him in the same breath: *"giving the particular
generative agents goblin names associated with their coding task so I can talk about them
more plainly with non-AI, non-code, non-data people. We do not need to get too carried
away by that now."*

**The thing being named already exists -- it is the Field Spec row.** A recipe holds a
purpose, a voice, a length, a Do-NOT list and an *own-job test* whose entire function is
to stop one field wandering into a neighbouring field's territory. That is a job
description, not a config row. Thirteen `Kind = Prose` fields, thirteen of them.

**What it buys, beyond plain speech:** the bug report writes itself. *"The Strategic
Alignment goblin keeps doing Problem's job"* is a precise statement that an own-job test
is failing, and it is intelligible to someone who has never heard of a register. This
project's hardest recurring problem is boundary bleed between adjacent panels; a named
owner per panel makes that boundary a thing people can point at.

**If it is built, it is ONE COLUMN on the Field Spec, not a roster in a document.** A
markdown list of names is a second copy of a machine-knowable fact and will drift -- the
class of defect this repo spent 2026-08-14 removing. As a column it derives out wherever
the field does: prompts, drafting tab labels, review sheets, run reports.

**The risk worth holding:** a name makes the output easier to over-trust. A goblin has a
brief, not judgement. It is why the approve gate exists and why naming must not soften it.

**Not scheduled.** The interesting half of the design session is what each goblin
*refuses* to do, which is the Do-NOT column read aloud.

---

## PARKED, nice-to-have: show/hide control for fields on a slide

Rohan, 2026-08-09, asked for it and then ranked it himself: **nice to have.** Not
scheduled. Recorded so the design conversation is not re-run from scratch.

**The want:** control whether a given field appears on a given slide -- "this
project has no Problem section this quarter".

**Grain:** the register row is already slide x period, so one `Hidden fields`
column on that row gives per-slide, per-period control. Field Spec is the wrong
home: it is per-field globally and cannot say "hidden on this project, not that
one". Values picked from known field IDs, never typed -- same rule as
`Sources.ApplyPeriodValidation`.

**Mechanism:** `Shape.Visible = msoFalse`, not blanking the text and not
deleting the shape. Blanking leaves an empty box with its fill and border still
drawn. Deleting destroys the tag, so sync could never bring the field back --
and it must come back, because next quarter the section returns.

**What Rohan wants to happen to the space, in his words:** *"ultimately leave
it. or replace it. some items could move."* So the behaviour is not one rule:
leaving the gap is the default, some panels would be REPLACED by other content
rather than emptied, and a minority of items would move. That last part is the
expensive one -- PowerPoint does not reflow, so closing a gap means the tool
moves slide furniture, which is a much larger promise than filling in text and
has to handle stacking, columns and grouped shapes without silently damaging a
layout.

**If it is ever built, build it in that order:** leave-the-gap first (cheap,
reversible, no geometry), replace-with second, move-things last or never.

---

## The time-elapsed bar: the best automation target on the slide

Rohan, 2026-08-09, drawing the distinction the model was missing:

> timeline will move towards end but not necessarily every quarter, vs time
> elapsed bar autoshapes that move with the clock regardless of progress

**Judgement versus clockwork.** Milestone markers move when the plan changes and
someone decides. The elapsed bar is a pure function of the date -- nobody
decides it, and it is wrong on every slide the moment the quarter turns, in a
way nobody notices until a reader does.

**It beats the prose panels as an automation target, which is counterintuitive
given where the effort has gone.** Prose needs drafting and review, and
automating delivery still leaves the expensive human part. The elapsed bar needs
NOTHING from a person: the slide already carries Start and End, the period
supplies the third input, and position is arithmetic. It changes every quarter,
guaranteed, on all 43 slides. Rohan's words: fiddling with **its two component
autoshapes is a PITA** -- so it is two shapes, not one, and the manual cost is
per-shape.

**It needs a KIND that does not exist.** `Kind` is how content is decided
(Controlled / Prose / Static); `FieldType` is what it is (Text / Picture /
Shape). The elapsed bar is `FieldType = Shape`, `Kind = Derived` -- computed,
never drafted, never approved, re-derived every run. Nothing in the tool has
that shape today, and it is the only kind that could sync with no human in the
loop at all.

**Design, as agreed:** the register says WHERE, sync just applies it -- keeps
sync dumb, which is what unattended quarterly work needs. Two refinements on
top of that, both about avoiding a fragile coupling:

- **Store a FRACTION, not points.** Raw PowerPoint coordinates in the register
  means the day anyone nudges the timeline, all 43 stored positions are quietly
  wrong and still look synced. Store `0.72` = "72% along", tag the timeline axis
  itself, and let sync read THAT shape's own Left and Width at run time. Moving
  the track then fixes itself.
- **Let Excel compute the fraction, not VBA.** A formula over start, end and
  period end is visible and checkable in the sheet, and it keeps date handling
  out of the VBA -- which matters specifically here, because the code has
  already been bitten by period parsing (`Q4F26` vs `FYnnQn`) and this sidesteps
  the whole class rather than adding to it.

So the register row gains `TIMELINE_START`, `TIMELINE_END`, `TIMELINE_ELAPSED`
(a formula over the first two and the period). Sync reads only the last.

**Open:** whether formatting varies too. Size alone covers "the bar gets
longer". Colour -- amber when the end date falls inside the current quarter, red
once passed -- is a second column AND a second decision, because a bar that
changes colour is making a claim about the project rather than showing a date.

**Not for the week it was raised in:** building it costs longer than dragging
the shapes once. This is next-quarter work, ranked above the Excel polish and
above show/hide.

---

## PARKED WITH ITS DESIGN SETTLED: picture fields

Rohan, 2026-08-10: pictures **"don't change quarterly, set once at project
start, I'd still like them set from link rather than played around with
manually"**. That answer makes this much smaller than the parked note assumed.

**It is not a sync field.** `FieldType = Picture`, `Kind = Static/Given` -- the
two axes the specs keep apart and must not collapse into one word. It carries
across a rollover untouched like any period-invariant content, so it costs
nothing per quarter.

**The cell holds a SOURCE ID, not a path.** Sources already solves one row per
thing, reference by ID, period binding, existence checking, and now feeds the
prompt. An image is evidence-shaped: it came from somewhere, that somewhere has
an owner, and "why is this photo on the slide?" is the same question as "why
does it say 90%?". A raw path per row would duplicate all of that badly -- 43
spellings of one folder and nothing checking any of them.

**Idempotence without comparing images:** when the tool fills a picture it
STAMPS THE SOURCE ID IT USED into the shape's own tag. Sync fills only when the
register's ID and the shape's stamped ID differ. So it fires once at project
start, stays silent forever after, and re-fires by itself if the link is ever
changed. No image comparison, no re-inserting 43 photos a quarter.

**Mechanism, once decided:** insert, copy Top/Left/Width/Height and z-order from
the tagged shape, delete the old, re-apply the tag. Fiddly, not deep -- the
tagged shape already knows its own geometry.

**THE ONE OPEN DECISION, and it is Rohan's:** when the image's proportions do
not match the frame -- fit inside and letterbox, fill and crop the overflow, or
refuse and report. Visibly different slides, no defensible default; it depends
whether the project photos are consistently shaped. Answer it against a few real
ones before anything is built.

---

# THE TIMELINE, READ FROM THE REAL DECK 2026-08-10 -- READ BEFORE TOUCHING BARS

Rohan: *"you can see the extremely accurate positioning? we need to maintain
that and z order etc."* Everything below is measured from `slide4.xml`, not
described from memory.

## What the timeline actually is

- **A VERTICAL bar of two rounded rectangles sharing an origin.** Track
  `Shape 188` is 0.05 x 3.03; fill `Shape 189` is 0.05 x 1.46 at the same x/y.
  The fill grows DOWNWARD from a fixed top, so drawing it changes exactly one
  property -- `Height`. No `Top`, no `Left`, no z-order.
- **`InjectProgressField` IS HORIZONTAL ONLY and refuses to guess the axis.**
  It cannot draw this slide today. That is a blocker, not a tweak.
- **Milestone state is FORMATTING, not content.** Circles are 0.35 except one
  at 0.43, and fills split `005832` (start, 6-month) against `003C23` (the
  rest). No register value can express size and colour.
- **The fill length is DERIVED** -- it runs to the next unachieved circle. So
  the register should carry WHICH MILESTONES ARE ACHIEVED, and the length is
  measured from circle positions the template already defines. The earlier
  `0.25||0.5||1` model assumed independent fractions per bar and is wrong for
  this layout.
- **Counts vary and are therefore data.** Ellipses per slide across the 44:
  mostly 6, but 4, 5, 7, 12 and 14 all occur.

## Three things that make naive manipulation dangerous

1. **Nothing is hidden.** No pre-placed big/small pair exists today, so a
   visibility model needs the TEMPLATE built for it. The tool cannot retrofit
   it.
2. **Shapes are stacked and share names.** THREE separate shapes are called
   `Shape 202`, two of them at the identical position `10.96` -- one filled,
   one not. Name-based addressing is hopeless; tags are the only safe handle.
3. **Structure varies within one slide.** Most circles are groups of
   circle+number; the 12-month one is a bigger loose ellipse with its number
   floating separately on top. Any "assume the pattern" logic breaks here.

## What is safe to change, and what is not

| operation | verdict |
|---|---|
| fill bar `Height` | SAFE -- origin fixed, one property |
| circle **fill colour** | SAFE -- no geometry touched |
| circle **size** | **NOT SAFE** -- resizing about a centre needs Top/Left compensation, which is the arithmetic that hid `LockAspectRatio` for five rounds |
| replacing any shape | NEVER -- a new shape lands on top and z-order is lost |
| writing label text | needs the existing geometry save/restore -- autofit moves the box |

## DECIDED 2026-08-10: THE VISIBILITY MODEL

Rohan: *"I like the idea of the visibility model."*

**Achievement is shown by toggling `.Visible` on pre-placed shapes, never by
resizing or recolouring computed by the tool.** The template carries both
states -- an achieved circle and an unachieved circle, at the same centre --
and the tool shows one and hides the other.

Why this and not the alternatives:

- **Position is preserved exactly**, because nothing moves. The tool computes
  no geometry, which is the standing rule the deck's own precision demands.
- **Z-order is preserved exactly**, because no shape is created, deleted or
  reordered. `.Visible` does not touch the stack.
- **The tool never invents a colour or a size.** Both states were authored by a
  person in the template, which is the same division as everywhere else here:
  the template owns geometry and formatting, the register owns values.

**What it requires next, in order:**

1. **A template decision by Rohan** -- the template gains a second, hidden
   circle per milestone in the achieved styling. Until that exists there is
   nothing to toggle. This is deck work, not code.
2. **A vertical mode for the bar.** Derive the axis from the TRACK's own
   dimensions (0.05 wide x 3.03 tall is unambiguous) and refuse only when the
   track is square-ish -- the existing refusal-to-guess comment is right about
   a square and wrong to give up on a 60:1 rectangle.
3. **Achievement in the register**, per milestone, `||`-separated like the
   values -- then the fill height is measured to the next unachieved circle
   rather than supplied.

**STILL UNKNOWN, needs Rohan's slides:** whether the achieved/unachieved pair
should be two shapes per milestone or one shape per state for the whole
timeline; and whether the 0.43 circle marks ACHIEVED or CURRENT -- on slide 4
it is dark (`003C23`, the unachieved colour) while the two lighter circles
above it are 0.35, which does not fit "big means achieved" on its own.

---

# STATE AT 2026-08-10 07:45 -- READ BEFORE PLANNING ANYTHING

## THE ONE THAT MATTERS: two features are built and unreachable

`InjectPictureField` and `InjectProgressField` are unit-tested and called by
**nothing except their own tests**. The sync path calls
`InjectPrimitive.InjectPrimitive` -- the TEXT injector -- for every field
(`ReviewQueue.bas:1283` and `:1301`, `SyncOperations.bas:157`), so a picture or
progress field is handed to the text writer and refused as "no text frame to
write into".

169 passing tests say nothing about this. Rohan found it by asking whether the
bars had been proven slide to slide.

**In this order:**

1. ~~**Template-clone test FIRST.**~~ **ANSWERED 2026-08-10 08:15 — THEY SURVIVE.
   The track-pair design stands; nothing has to change.** Run against real
   PowerPoint via `vba/tools/tag_cloning_probe.ps1` (the probe had existed since
   31 July with no runner and no recorded result — written, never run, which is
   the same shape as the unreachable injectors it exists to answer for).

   What survived `Slide.Duplicate`: the suffixed role value `BAR_BODY.track`, its
   pair `BAR_BODY`, **both inside a GROUP**, and a second tag name on the same
   shape (`picsrc`). Slide-level tags survive too, and the duplicate gets a new
   SlideID.

   Two things the first run could not have told me, both fixed before the verdict
   was believed: `DumpTags` printed a grouped shape and a never-grouped shape
   identically, so it could not show the group had actually formed (it now prints
   `[Group 4] GROUP of 2` and marks members); and nothing demonstrated the finder
   could return "not found". Two controls now run and **must** read False — a role
   never tagged, and a picture stamp on the track shape. Both did.

   Incidental, not today's question but worth having in writing: **Copy/Paste into
   a different presentation preserved the SlideID exactly.** SlideID is unique
   within a deck, not across decks — do not use it as a cross-deck discriminator.
2. ~~**Dispatch by shape type in the sync path.**~~ **BUILT 2026-08-10.**
   `InjectPrimitive.InjectField` is the one entry point; the type is derived
   from the SHAPE (picture / a `.track` sibling / else text), so it cannot
   disagree with a column the way `Kind` does. All three sync call sites rewired
   (`ReviewQueue` ×2 with the Sources sheet resolved once per run,
   `SyncOperations` without one). New `Sources.LocatorFor`, with a `found`
   out-parameter so "cited a source that does not exist" and "source row has a
   blank locator" get different messages — they send the person to different
   files.

   A non-numeric register cell for a bar is refused **loudly**: `Found` and
   `WouldChange` are both set, because `Found=False` is a SKIP to
   `SyncOperations` and `WouldChange=False` is a NO CHANGE — either would let a
   bar silently fail to draw inside a run reporting success. `Val("done")`
   returning 0 would have drawn an empty bar and called it a success.

3. ~~**A multi-slide test.**~~ **DONE** — `InjectField_TwoSlidesEachGetTheirOwnValue`,
   with the two slides given **different track widths** so a bar drawn against
   the wrong slide's track cannot land on the right number by coincidence.

**172 passed / 0 failed behind the compile gate, and every new test was made to
fail on purpose first, in two separate runs.** Breaking both non-text branches
failed all three new tests while the five existing picture/progress tests stayed
green — which is the blindness itself, since those test the injectors directly
and cannot see that nothing calls them. Breaking the picture branch ALONE then
failed only the picture assertions, with the failure text being the exact defect
this work removes: `shape tagged role=PHOTO has no text frame to write into`.

## 2a. PICTURES ARE NOW REACHABLE. BARS ARE STILL NOT — THERE IS NO WAY TO MARK ONE

**Found 2026-08-10 by Rohan asking "how are these types of fields marked",
while the dispatch was still being written.** The dispatch was necessary and is
not sufficient.

- **Pictures: markable.** `BatchOnboardFlow`'s gate accepts a picture shape
  (added earlier the same day) and `Onboarding.IsCandidateField` offers pictures
  in discovery. With the router in, pictures now work end to end.
- **Progress bars: not markable at all.** Two independent blocks. The marking
  gate (`BatchOnboardFlow.bas:1429-1436`) requires a picture OR a shape with
  non-empty text; a bar's done part and its track are empty rectangles, so both
  are refused — the message says *"progress bars are not supported yet"* in as
  many words. And **nothing outside `InjectPrimitive.bas` mentions `.track`
  anywhere in the codebase**: `Onboarding.ConfirmFieldMatch` is the only thing
  that ever writes a `role` tag, and no path offers a suffixed value.

**DECIDED 2026-08-10, not yet built: bars are tagged ONCE ON THE TEMPLATE, not
per slide.** Rohan's call, and this morning's probe is what makes it safe — the
tags survive duplication, including inside a group. A bar is furniture the
template owns, so there is no per-slide marking and no guessing which of two
shapes is the track.

**Also decided, and it is a boundary worth keeping:** `Behaviour` (Fill / Fit /
Leave as is) **stays on the Field Spec for now**, even though it is a shape fact
and would arguably be better on the tag. The line agreed: the template owns what
is true of the SHAPE, the workbook owns what is true of the CONTENT — and
`Kind`, which is already double-sourced, gets fixed by making the sheet win, not
by moving it to a third place.

**The standing cost of leaning on tags: a tag is invisible.** It cannot be seen,
diffed or repaired without the tool, on the machine with no Python and no
Claude. If real behaviour moves onto the template, a toolbar-reachable "what
does this slide say about itself" dump stops being a nicety —
`read_deck_slide_tags.py` is Python and therefore useless at work.

## 2b. `field_e2e.ps1`'s module list is missing `Readiness.bas`

Pre-existing, found 2026-08-10 by running `vba/tools/check_module_lists.py`
(not caused by the dispatch work; the test-suite and add-in lists both pass).

    FAIL  harness  (field_e2e.ps1)
            missing Readiness.bas -- referenced by DeckRegistry, RibbonUI, WorkbookBridge

The harness will fail at runtime on any path that reaches Readiness. Cost: one
line.

## What is built and verified

> **SUPERSEDED IN PART, 2026-08-14 night.** The two buttons named below are gone.
> The toolbar is now `1. Set up my quarter` / `2. Put it on the slides` /
> `Review changes (writes nothing)`, split by ARTIFACT so neither side can trigger
> the other, and `Rebuild my sheets` is deleted outright. `Layout 4` below is now
> **layout 5**. The verification records are still true of the day they were made.

> - **Two buttons** (`1. Sync Now`, `2. Rebuild my sheets`) with every capability
>   reached from inside the chain; `check_vba_static.py` fails the build on an
>   orphan and caught nine during the refactor. *(Superseded — see the banner above.
>   Kept inside the quote because it is a record of that day, not a description of
>   the tool.)*
- **The whole quarter loop ran through them**, verified from file bytes: period
  set and confirmed on disk, 43 rows rolled forward, a second roll refused,
  `PendingApprovals` fired on its first real press, three fields written to
  `2_P004`, backup taken one second before the write.
- **Layout 4** drafting columns in workflow order, with `ColumnInLayout`
  migrating layout-3 sheets rather than dropping their work.
- **Cited sources reach the prompt** (`Sources.CitedBlockFor`). Before this the
  evidence rule could never be satisfied -- `[TBC]` was permanent no matter how
  many sources were cited.
- **Colour family per field**, from POSITION not a name hash (a hash collided
  ABOUT_BODY with PROGRESS_BODY, the two most-used sheets).
- **Every column letter and button caption in user-facing text is derived**, not
  typed. 29 stale captions swept on 08-09, four stale column letters on 08-10.
- **Build stamp** written by `build_ppam.ps1` into `CommandBarUI.BUILD_STAMP`
  and shown on every report. **Not in `addin62`** -- it was added after.

## Rules this codebase earned the hard way, tonight

- **Green is not evidence; green plus a demonstrated red is.** Five checks
  tonight passed while being incapable of failing. Twice the first attempt to
  break a test failed to break it -- which is itself the finding.
- **Probe the mechanism before modelling the symptom.** Five rounds of geometry
  arithmetic were spent compensating for `LockAspectRatio`, one property, which
  a thirty-second probe settled. A retry loop that re-asserts a value is an
  admission that something else is changing it -- go find out what.
- **The template owns geometry, the register owns values, the tool computes
  neither.** Every fit/fill/scale calculation was the tool deciding something
  the template had already decided.
- **Trace the path from a person's action to the code before calling it done.**
  Both unreachable features would have failed that sentence immediately.

## Rohan's constraints, current

- The workbook can NEVER be `.xlsm` -- work restrictions. No workbook-level VBA
  ever; all behaviour lives in the `.ppam`. Add-ins themselves ARE permitted at
  work (`addin33` ran there).
- Native in-cell checkboxes have no VBA API. The achievable "toggle" is a
  validation dropdown plus conditional formatting.
- Pictures are set once at project start, not quarterly -- so a picture field is
  a `Given` filled from a link, not a sync field.
- Sharing turns on **portability, install, single truth** -- his words. Single
  truth forces a shared register, which forces the cloud path.

---

## PAPERCUTS FROM THE FIRST SUCCESSFUL PUBLISH — 2026-08-13

All found by pressing buttons on the real deck, none visible to 192 passing tests.
Ranked by how much of an evening they cost.

### P1. A dialog opens BEHIND the PowerPoint window, and reads as "nothing happened"

**CODE FIX WRITTEN, NOT YET BUILT/DEPLOYED/LIVE-VERIFIED (2026-08-17,
background session).** `DraftingUI.BringPowerPointToFront` added right
beside `BringExcelToFront`, same marker-caption/`AppActivate` technique,
`Public` so `RibbonUI.SyncNowChainCore` can call it. Wired in at both
points `NEXT-SESSION.md`'s plan named: the very top of
`SyncNowChainCore` (covers the period-confirm `MsgBox` and everything
early in the chain) and immediately after `DraftingUI.RollForwardUI`
returns, before `RefreshDraftingSheets` (undoes Roll Forward's own
correct Excel-activation). `check_vba_static.py` clean across 38
modules; the Python `pytest` suite could not be run in this sandbox (no
`pip`/network here) but nothing Python-side was touched. **Still needed:
build the next addin, deploy, and retry the real Scenario 1 attempt from
scratch on Rohan's machine — this has not been proven live.**

Prior state, for context: reconfirmed live 2026-08-17 evening, during
this project's first-ever real Scenario 1 attempt. Hit a third time,
four different hidden window titles across two presses ("Start a
Quarter", "Roll Forward", "1. Set up my quarter", "PopupHost") — root
cause diagnosed precisely (exactly one call site, `RollForwardUI`, does
this correctly, and it actively LEAVES Excel frontmost afterward,
burying the very next prompt in the chain).

**Three times in one session** (2026-08-13, the original finding). Rohan pressed a button, nothing appeared, and the run
looked dead. Each time a VBA modal was sitting behind another window — twice behind
Excel, once behind PowerPoint itself. It cost two separate diagnostic detours before a
reliable test was found.

**The calibrated test, worth keeping:** while a modal is open PowerPoint stops answering
COM — `ActivePresentation.Name` comes back empty and `Slides.Count` reads 0. Idle, it
answers normally. That distinguishes "waiting for you" from "finished" in one call, and it
was the only thing that settled it.

**Fix:** activate the PowerPoint window immediately before showing any prompt
(`AppActivate`, or `Application.Activate`). A question nobody can see is not a question.

**Do NOT fix by adding waits.** Rohan asked whether the code should "cycle through
applications to allow adequate initialisation time". Tempting and wrong: nothing here is
an initialisation problem — Excel was doing genuine work, building 13 drafting sheets of
~43 rows. A sleep long enough to help is wrong on a faster machine and still wrong on a
slower one, and it would mask the real defect. Zettel
`20260810-compensating-arithmetic-hides-the-mechanism-you-never-probed`.

**Second half of the same defect: nothing says "working".** A long Excel operation with a
silent PowerPoint is indistinguishable from a crash. A status-bar line, or making Excel
visible while it writes, is the honest fix.

### P2. The field-picker InputBox has its text field OFF THE BOTTOM OF THE SCREEN

**PRIMARY PATH FIXED 2026-08-14 (was already stale before tonight's pass found it) --
RESIDUAL FIXED 2026-08-17 late night.** This entry had drifted: it described the typed
InputBox as THE mechanism, but Rohan asked for click-to-pick instead on 2026-08-14
("I don't want to do any secret hidden typing, I need to select the field by clicking on
it"), and `PickFieldByClicking` (`DraftingUI.bas:674`) became the primary path -- it
already does exactly this item's own suggested fix, "put the explanation on a sheet, not
in the dialog," via `ShowSheet`/`BringExcelToFront`. The typed `InputBox` survived only as
a fallback for when clicking is cancelled, but it still built the SAME long message
(drafting fields AND the full "other fields, edit in register" list) that caused the
original off-screen defect -- so the residual was real, just rare (only reachable when
clicking is declined).

**Original description, for the record:** `AskForField`'s prompt listed every field in the
workbook — around 40 lines by the time it named the drafting fields, then the Given/
Derived/Controlled ones with their kinds. On a 1080p screen that pushed the actual entry
box below the screen edge. Rohan: **"what box?"** He could not type into it because he
could not see it — not cosmetic, the box returns "" when dismissed, silently cancelling
the stage.

**Residual fix:** the fallback prompt now only lists the drafting (Prose) fields — the
ones actually worth typing here — and points at the Field Spec sheet (already open in
Excel by the time this fallback can fire) instead of repeating the full "other fields"
list a second time. `check_vba_static.py` clean. No automated test exists for either path
(both are interactive InputBox/Excel-picker UI, same class of untestable-this-way gap as
item AS) — a live look at the fallback next time it's reached is the real proof, same
honest limit as AS.

### P3. The 21 `MS*` "fields with nothing to write into" warning is a FALSE POSITIVE

**FIXED 2026-08-17 late night — three of four suspected consumers were already fine, the
"device registry" needed was one function, not a new module.** The design doc
(`NEXT-SESSION.md`, "A DEVICE REGISTRY") described four broken consumers; re-checked
against the actual current code before building anything (per this project's own "read
the file, don't produce a theory" rule), and the picture had changed since that doc was
written:

- **Discovery** — already fixed (`Discovery.bas:165`, `MilestoneDevice.SlotCount(shp) > 0`
  recognises the whole group as ONE candidate, never descends into its parts).
- **Marking** — already handled reasonably (`BatchOnboardFlow.bas:1503` — asks "tag the
  whole group as ONE field?" before offering to open it up; not a strict refusal, but
  informed, not blind).
- **Template Audit** — traced, not assumed: it builds its candidate list via
  `Discovery.DiscoverSlideWithShapes`, the SAME function Discovery's fix lives in, so it
  never sees the 21 internals as separate candidates either. Confirmed the remaining path
  too — `ShapeText` (`TemplateAudit.bas:150`) wraps `.TextFrame.TextRange.Text` in `On
  Error Resume Next`, and a group shape has no `TextFrame`, so the device candidate
  silently returns `""` and never becomes an audit row, tagged or not. No code change
  needed.
- **`FieldWiring`** — the one real gap, confirmed unchanged: `ScanFieldWiring` still asked
  "does any slide's role-tag set carry this exact column name," a question `MS1_LABEL`
  etc. can never answer yes to by design (they're addressed by shape name inside a tagged
  group, not by individual role tag).

**Fix, narrow, matching the pattern the codebase already chose twice** (Discovery and
Marking both call `MilestoneDevice.SlotCount` directly — no registry indirection): added
`MilestoneDevice.IsColumnForThisDevice(colName)`, and `ScanFieldWiring`'s per-field loop
now skips device-owned columns into a new `DeviceOwnedCount` instead of running them
through the carrier/unmarked/case-mismatch checks that don't apply to them. Building a
separate `DeviceRegistry` module would have been inventing an abstraction the codebase had
already independently avoided for a population of one device.

New test `Test_FieldWiring_DeviceOwnedColumnsAreNotUnmarkedFields` — proves BOTH halves
(device columns excluded AND a genuinely unwired ordinary field alongside them still
caught, so the fix discriminates rather than blanket-suppressing). Made to fail first: with
the fix stashed, the test doesn't even COMPILE (references a result field that doesn't
exist yet) — as strong a "make it fail" proof as this gets. `check_vba_static.py` clean.

### P4. FIXED 2026-08-19 — The 17-column prompt is ALL-OR-NOTHING across a mixed set

"17 field(s) on the Field Spec have no column in the register. Add a column for each?"
Sixteen were uncontroversial (`INDUSTRY_CASH`, `START_DATE`, `PROJECT_LEAD` …). One was
`HIGHLIGHTS_BODY`, which **must not** get a single column — it is three shapes per slide
and needs slot columns like the milestones.

So a real architectural decision was made, silently, by a Yes on a bundled prompt. Declined
twice on 13 Aug for exactly this reason.

**Fixed exactly as the second option proposed: a fourth Renders-as value, `Slots`.**
`FieldSpec.RENDER_SLOTS`, recognised by `RendersAsFor` and offered in the real Excel
dropdown (`ApplyRendersValidation`) — confirmed `RendersAsFor` has zero callers outside
`FieldSpec.bas` itself before adding a case, so the change carries no blast radius into
injection code. `ExcelOutput.MissingRegisterColumns` now excludes any field marked `Slots`
from the bundled list, the same rule already applied to Derived fields. **`HIGHLIGHTS_BODY`
itself set to `Slots` on the live Field Spec sheet**, confirmed saved. New tests prove the
exclusion, made to fail first. Full suite 270/270.

### P5. FIXED 2026-08-19 — Re-running the Template Audit REPLACES the sheet, decisions included

The dialog said so — *"Re-running this REPLACES that sheet, decisions included"* — which was
honest, and still wrong. A 2026-08-15 interim fix stopped the silent loss by refusing to
rebuild over pending decisions, but that still meant the audit could only ever be worked in
one sitting — the actual complaint. Same shape as the Discover Fields grid (item 3), which
still rebuilds from scratch and loses marks (not touched by this fix).

**Fixed as originally proposed: `TemplateAudit.WriteAuditGrid` now carries decisions across
by shape identity** (name + group path + text together — a decision was made about what a
shape SAID, so a genuinely changed text correctly does not inherit it), the same pattern
`Drafting.WriteDraftingSheet` already uses to carry drafts across a rebuild rather than
blocking it. `carriedCount`/`orphanedCount` are new ByRef outputs so a person is told
exactly how many decisions carried and how many couldn't (shape/text gone), rather than
either silently losing them or being blocked outright. New tests prove both halves, made to
fail first (temporarily dropped text from the identity key, confirmed the orphan test fails
on the right assertions before restoring it). Full suite 269/269.

### P6. `PROGRESS_BODY` has more approvals than submitted text — CONFIRMED, not a miscount

**Settled 2026-08-13 against the saved file**, after chat side and Claude Code disagreed on
which column held which number. Authoritative counts, header row excluded, 43 data rows:

```
                       E draft   F submit   G approve
TPL_KEY_EVENTS_BODY          1         43          43
TPL_PROGRESS_BODY            0         34          43
TPL_HIGHLIGHTS_BODY         43         43          43
```

Constants, so nobody re-infers the mapping: `COL_D_DRAFT = 5` (E), `COL_D_SUBMIT = 6` (F),
`COL_D_APPROVED = 7` (G), `DRAFT_HEADER_ROW = 9`, `DRAFT_FIRST_ROW = 10`.

**RETRACTED 13 Aug, later the same evening. The count was mine and it was wrong.**

Column G holds `'0'` on at least some rows, not `Y` — `3_P001` on `TPL_PROGRESS_BODY` is
`SUBMIT` empty, `APPROVED = '0'`. My figure counted **non-empty** G cells, so it counted
`'0'` as an approval. Publishing tests `ReviewQueue.IsApprovalMark`, which is an
*affirmative* test, not a non-empty one.

So "43 approvals against 34 texts" is not established. The real approved count is lower and
unknown until someone counts affirmative marks rather than filled cells.

**This is the exact error this file exists to catch, committed by the person writing the
file.** A check that asks "is there something here" cannot answer "does this say yes".
Recount with `IsApprovalMark`'s own rule before treating the mismatch as real. Publishing correctly requires BOTH, so the nine extra
publish nothing — but the reported count will not match the ticks, and there is no message
explaining why. Either the ticks are stale or the text was lost; nothing currently says
which, and it should.

**A counting habit this cost an exchange to learn:** these sheets are rebuilt by any
`Sync Now`, so a count is only true of one moment. Quote the workbook's mtime beside any
figure taken from it — two correct counts taken an hour apart will disagree and look like a
defect.

### Confirmed FIXED 2026-08-13 (addin81), listed so they are not re-found

- **Publish/Copy-AI could only ever reach the FIRST `Kind = Prose` field.** `FieldForRun`
  now asks inside a chain. Two call sites. Proven on the real deck at 17:23.
- **Roll Forward asked a question whose every answer was refused.** `PeriodRowCount` lets
  the caller check first; inside the chain it is now one line in the report.
- **The slide-type picker demanded typed input for a one-item list** (item 7). `PickType`
  auto-selects a sole type and still asks when there is a real choice. Three call sites.
- **`RefreshDraftingSheets` reported "drafting sheets are ready. Workbook saved." over
  seven refusals.** Refusal now leads, names the fields, and carries a warning icon.

---

## FOUND BY PRESSING THE BUTTON — 2026-08-15, ALL THREE STILL LIVE

Every one of these was caught at a dialog on the real deck. **None was found by the
suite**, which went 194/0 to 199/0 across the same session without seeing any of them.

### A. FIXED 2026-08-15 — the harvest wrote a formatted VIEW into a numeric-contract field

> **Fixed in source, NOT yet in a built add-in.** `Harvest.HarvestSlide` now asks
> `InjectPrimitive.InjectorFor` — the same decision `InjectField` routes on, extracted
> rather than copied so the two cannot drift — and refuses anything that is not
> `INJECTOR_TEXT`, naming it. A group is refused separately, because a group has no text
> of its own even when the router would not call it a device. **Refusal, not conversion:**
> `33%` -> `0.33` is guessable, `33% (est.)` is not, and a wrong guess is permanent.
> Suite 200/0. Original entry follows.

`PROJECT_PROGRESS` reads `33%` off the slide. `InjectPrimitive.bas:340` refuses any
non-numeric progress value and names **`'90%'` as wrong in those exact words**. So the
harvest would write a value the tool itself cannot publish.

**Why it is worse than an ordinary bug:** `Harvest.HarvestSlide` writes only where the
register is empty. Once `'33%'` is in the cell it is no longer empty, so a corrected
harvest **cannot overwrite it** — it has to be cleared by hand, exactly like the four
Excel-coerced cells cleared at 23:40 on 14 Aug.

The harvest assumes "what is displayed" is what the register wants. True for prose, names
and dates; false wherever the slide shows a formatted view of a stored value. It already
refuses devices by name — numeric-contract fields need the same treatment: convert
(`33%` -> `0.33`) or refuse, never write the string.

### B. FIXED 2026-08-15 (same night, doc never updated until 2026-08-19) —
### `OfferHarvestForSelectedSlides`'s prompt mislabels and truncates

Two defects in one dialog, both in `RibbonUI.bas`:

1. **All propagation detail is accumulated into the `collisions` string**, so successful
   stamps print underneath a `Refused -- two fields matched one shape:` header. On
   14 Aug 23:57 that made a run reporting 16 correct stamps read as though it had refused
   everything.
2. **It hits `CapReport`'s 900-character cap mid-word**, so collisions can be present and
   invisible. A person cannot consent to what the dialog does not show them.

Neither is dangerous — the guards bound the write, not the text — but this project has
already paid twice for approving a prompt that could not be fully read.

**Checked 2026-08-19, found already fixed** — `1986c9a`, "Stop the harvest prompt calling
successes refusals, and rename the other harvest," landed less than an hour after this item
was written (00:45 the same night). Both halves confirmed still in place by reading the
current code, not assumed from the doc:
1. `collisions` now accumulates ONLY real collisions (`Harvest.PropagateTemplateTags`'s own
   `.Collisions`, never `.Detail`) — the code's own comment names the exact former bug:
   "mixing the two is what printed 16 successful stamps under a 'Refused' header."
2. `CapReport(ask)` now caps the WHOLE assembled dialog text, not one part while another
   grows unbounded, and the full plan is written to the Run Log sheet BEFORE any capping —
   so a shortened dialog now says so explicitly ("[shortened -- the full list is on the Run
   Log sheet]") instead of silently cutting a line mid-word with no notice.

No code change needed here tonight — this entry itself was the only thing stale.

### C. Slide 27 carries a shape already named `Text 216a` that is not the date

The 2026-08-15 rename pass (55 shapes, so 32 of 44 slides carry both `Text 212a` and
`Text 216a`) **refused** slide 27 rather than create a duplicate name. That slide's
`END_DATE` will keep colliding until a human looks at it. Correct behaviour, still an
open item.

**Slide 27 itself is still open — this class of collision cannot be healed automatically,
by design** (`Matching.bas`'s name tie-break needs the real slide's shape name to match
the template's exactly, and the shipped add-in deliberately never renames a shape itself).
2026-08-19: closed the door for every future template instead. `TemplateSlide.
MakeTemplateFrom` now refuses (and deletes the half-built copy) if any two shapes on the
resulting template share a name, so nothing built from a template from here on can drift
into slide 27's problem. Slide 27's own fix is still the five-minute manual VBA Immediate
Window rename described in this session's chat — rename whatever currently squats on
`Text 216a`, then rename the real `END_DATE` shape onto it.

### D. FIXED 2026-08-19 — `check_vba_static.py`'s reachability check is weaker than its name

It asked whether a procedure's NAME appears in another module, **not** whether anything
reachable calls it. A chain of private orphans was invisible to it. Proven 2026-08-14:
commenting out the call to a wrapper left the callee's name still written inside the
now-orphaned wrapper, and the checker stayed clean.

**Fixed, deliberately narrow rather than a full call-graph rewrite** — the "reachability
is reported, not enforced" design is itself deliberate (this project has already paid for
a checker that reports maybes and gets ignored). The specific, provable gap: a direct call
from `tests/TestRunner.bas` counted as proof a PERSON could reach a capability, which is
exactly the disguise `CreateMissingSlides`/`RollForwardPeriod` wore the first two times this
class of bug was found. Excluded test-only reachability from the "no caller anywhere"
branch specifically (excluding it from the milder "used only inside its own module" branch
too was tried first and produced 41 near-entirely-false notices — VBA has no module-
internal-but-testable visibility, so most Public-for-testing functions would have been
flagged as over-exposed for no real reason).

**A second, deeper bug found sanity-checking the first fix's real output**: the
"over-exposed" heuristic counted a function's own `FuncName = result` return-assignment
lines as if they were calls, so any Function with an exit point already scored above the
threshold before anything real called it. `BatchOnboardFlow.BuildBatchPlan` — genuinely zero
callers anywhere, confirmed by grep — was landing in the milder bucket instead of the
correct one. Fixed by excluding self-return-assignment lines from the count.

Verified against the real corpus (this tool's own established testing convention, it has no
unit-test suite of its own): 41 genuine "make it Private" notes, one genuine "tested but
unreachable by a person" note (`BuildBatchPlan` — see the new item below). Made to fail
first: reverted the self-return-assignment exclusion in a scratch copy, confirmed
`BuildBatchPlan` reverts to the wrong bucket, restored the fix.

### BL. RESOLVED 2026-08-19, same session — `BatchOnboardFlow.BuildBatchPlan` is built, tested,
### and has NO caller anywhere, including its own module

The exact "tested unit behind a locked door" shape this project keeps finding by hand
(`CreateMissingSlides`, `RollForwardPeriod`, the picture/progress-bar findings of
2026-08-10) — this time found by the checker itself, the first time the fixed
`check_vba_static.py` ran against the real corpus. `BuildBatchPlan(templateSld,
otherSlides())` genuinely has zero callers anywhere in production code (confirmed by grep,
not assumed) — it's called only from `tests/TestRunner.bas`. Its sibling,
`BuildBatchPlanFromMarkedFields`, IS reachable (used within `BatchOnboardFlow.bas` itself),
so this is not a case of the whole batch-onboarding mechanism being dead — just this one
specific entry path into it.

**Resolved, not by adding a button — `BuildBatchPlan`'s own header comment already
explains it, read rather than assumed.** *"Kept for the tests that already exercise it
directly; the live 'Bulk Onboard Type' ribbon entry point uses `BuildBatchPlanFromMarkedFields`
instead"* — Discovery-based auto-enumeration produced an unreviewable 87-row grid on the
real deck (2026-07-26), so the marked-fields flow replaced it deliberately. Not dead code
from neglect; a design that was tried and explicitly rejected for a documented reason.
Building a button for it would have resurrected a rejected UX. Added to
`check_vba_static.py`'s `REACHABLE_OTHERWISE` list with the reason instead, so the checker
stops flagging a deliberate decision as if it were an oversight.

### E. FIXED 2026-08-15 — the harvest prompt undercounted what it would write

> **Fixed and verified on the real deck**, `bd63134`. `PropagateTemplateTags` records the
> shape text per role it would stamp; the dry harvest consults that map, so the preview
> counts fields the same press is about to create. Dry-run only by construction — a wet
> run stamps first, so the ordinary lookup finds the shape and the map is never read.
> **Proven 06:02: offer `22 shape(s), 34 value(s)`, result `22 shape(s) labelled, 34
> value(s) written`** — the pair that previously read 10 then 34. Register arithmetic
> agrees: six fields went 38 -> 34 empty, and START_DATE/END_DATE 38 -> 35, the one-row
> difference being slide 8's refused collision. Suite 201/0. Original entry follows.

**Demonstrated at scale 2026-08-15 05:44: the dialog offered `10 value(s)` and the run
wrote `34`.** Every value was correct; the number was wrong by 3.4x.

Pass one runs the harvest dry-run BEFORE the propagation dry-run has taken effect —
neither writes anything, so the harvest can only see fields already carrying a role tag.
Pass two labels and then reads per slide, so the six newly-labelled scalars on each slide
are harvested in the same run and never appear in the count that was consented to.

**Why it matters even though nothing wrong was written:** this is the one approval gate in
front of an operation that changes both files, and its headline number is the thing a
person actually reads. A gate whose number is routinely wrong teaches people to ignore
the number.

**Fix:** the dry run already knows which roles propagation WOULD stamp, and
`PropagateTemplateTags`' Detail already carries each one's shape text. Count a value as
"would be written" when that role's register cell is empty and the shape's text is
non-empty — the same two tests `HarvestSlide` applies — instead of only counting
already-tagged fields.

### AV. FIXED 2026-08-19 — `PairingProblem` existed and was checked in exactly one place

The workbook↔deck identity check (`DeckRegistry.PairingProblem`, stamped by
`StampPairing`) was built 2026-08-14 specifically to close the "wrong register" risk —
but the only caller was `DraftingUI.bas`'s AI-drafting publish path. Every other place
that opens the register and then writes real content never asked whether the workbook
it just opened actually belongs to this deck: `DraftingUI.Resolve` (the shared resolver
behind "1. Set up my quarter" and three other callers), `RibbonUI.BuildAllQueuesCore`
("Review changes" and half of "Put it on the slides"), `RibbonUI.ApplyApprovedCore`
(the actual slide-write step — the highest-stakes gap), `RibbonUI.OfferHarvestForSelectedSlides`
(the reverse direction, slide→register), `AdoptFlow.PromptAdoptExistingSlides`
(onboarding writes), and both branches of `BatchOnboardFlow.ResolveDataWorkbook`
(silently reusing an existing path, and choosing a new one during onboarding — this
second branch also never completed the pairing: it set the deck's half of the pairing
but never stamped the workbook's, so a deck onboarded this way looked "unstamped"
forever). `DiscoverUI.DiscoverFields`'s common case (an already-paired deck) was also
uncovered.

**Fix:** added the identical `PairingProblem` check to all eight sites, matching the
existing pattern exactly. `BatchOnboardFlow`'s new-workbook branch also gained a check
*before* stamping — without it, choosing an existing file there would have silently
overwritten that file's identity with this deck's, hijacking someone else's register.
`RibbonUI.ResolveSyncContext` (found to have zero callers — see item AX's sibling note
below) was fixed too, for consistency, even though nothing currently reaches it.

Full suite 246/246 across all eight edits. `ResolveSyncContext` and `BuildAllQueuesCore`/
`ApplyApprovedCore` still carry independent, duplicated copies of the same
resolve-and-check sequence rather than sharing one — the exact "two copies of a guard
drift" shape `WarnOnDuplicateKeys` was already extracted to fix once. Not consolidated
this pass; flagged as a real follow-up, not done silently.

### AW. FIXED 2026-08-19 — `SaveWorkbookVerified` reported failure for a register with nothing pending, live, blocking

`WorkbookBridge.SaveWorkbookVerified` never had the "nothing pending is not a failure"
fix `DeckRegistry.SaveDeckVerified` got on 2026-08-14 for the identical defect: a plain
`wb.Save()` on a workbook with nothing to save legitimately does not move the file's
mtime, and without a `wasClean` check that read as "THE WORKBOOK WAS NOT SAVED" — hit
live, mid-session, blocking a real Apply Approved run on the real register. Ported
`SaveDeckVerified`'s exact fix: capture `wb.Saved` before attempting the save, exit
clean if there was nothing pending, and only escalate to a forced full rewrite
(`wb.SaveAs path`, no format arg) when a real change genuinely didn't reach disk.

**A second, real bug found applying the same pattern**: the forced-rewrite escalation
needed `wb.Application.DisplayAlerts = False` around the `SaveAs` call, which
`SaveDeckVerified`'s identical-looking PowerPoint version never needed — confirmed live
that Excel prompts "a file named ... already exists ... replace it?" for a `SaveAs` to
its own already-open path, and PowerPoint does not. Restored unconditionally right
after the call, not in an error handler, so a future edit here can't leave it off.

New tests `WorkbookBridge_SaveWorkbookVerifiedProvesTheFileMoved` (mirrors
`Test_DeckRegistry_SaveDeckVerifiedProvesTheFileMoved` exactly) and
`WorkbookBridge_EnsureSavedQuietlySavesWithoutAsking` (see AX). Full suite 246/246.

### AX. FIXED 2026-08-19 — the "unsaved changes, save now?" prompt was an invariant answered "Yes" every time

Four independent copies of "`IsDirty`, then a Yes/No `MsgBox`, then `SaveWorkbookVerified`"
lived across `RibbonUI.bas`. The prompt's own justification (written 2026-08-14) already
argued for doing the save silently — *"go and press Ctrl+S yourself is friction with no
safety benefit over doing it for them"* — without noticing that argument applies to the
prompt itself, not just to who performs the save. An invariant prompt, always answered
the same way, is how the one that matters gets clicked past (the same rule that collapsed
"Put it on the slides" to one press).

**Fix:** new `WorkbookBridge.EnsureSavedQuietly(wb, path)` — saves without asking if
dirty, returns `""` on success or a real error message on genuine failure (a save that
actually fails is never hidden). Replaced all four `RibbonUI.bas` copies with a single
call each. `UnsavedWorkbookText` (the prompt's wording, previously pinned by its own
test) is deleted — nothing shows it anymore, and a written-but-unread string is exactly
this project's own signature defect in miniature.

**A live, real-world byproduct of this whole evening's testing**: running the automated
suite repeatedly in quick succession surfaced roughly 1,000 accumulated leftover test
files in `%TEMP%` (days old, from every prior run that crashed before its own cleanup
ran) and several genuine SaveAs-collision bugs across both the test harness and this
fix's own new tests — full detail in `AGENTS.md`'s "hangs waiting for a click nobody's
there to give" section rather than duplicated here, since it's tooling hygiene, not a
product defect.

### AY. FIXED 2026-08-19 — the milestone timeline was never register-driven in practice: corrupted data, an unconverted line-break delimiter, and a roll-forward gap that was actually a data-freshness gap

`3_P001`'s `MILESTONE_TIMELINE` device (the only project with any milestone data at
all) had three real, separate defects, found by comparing a fresh register-only
rebuild against Rohan's real original deck:

1. **Data entered by hand, packed wrong.** `MS1_LABEL`/`MS7_LABEL` had the milestone's
   date typed straight into the label with `" | "` (e.g. `"Project initiated | Oct
   2023"`), leaving `MS1_DATE`/`MS7_DATE` as literal `"?"` placeholders. `MS3_LABEL`/
   `MS4_LABEL` are genuine two-line labels but used the same `" | "` instead of the
   codebase's real multi-line convention, `InjectPrimitive.LINE_BREAK_DELIMITER`
   (`"||"`, no spaces).
2. **`MilestoneDevice.WriteText` never converted the delimiter at all.** Unlike
   `InjectPrimitive`'s own text writer (`Replace(sourceValue, LINE_BREAK_DELIMITER,
   PPT_LINE_BREAK)`), the milestone writer did a raw, unconverted
   `shp.TextFrame.TextRange.text = value` assignment -- so even a correctly-formatted
   `"||"` label would have shown two literal pipe characters on a real slide instead
   of a line break. This is the actual code defect; everything else this item lists
   is data.
3. **`MS*_DATE` cells corrupted the same way as `PROJECT_PROGRESS`'s percentage cells
   were (item AU's sibling defect, same session).** Assigning `.Value = "Oct 2023"`
   to a cell without first forcing `NumberFormat = "@"` let Excel silently
   reinterpret it as a real date (`1/10/2023`), and the bare-number `MS2`-`MS6`
   dates (`6`, `12`, `24`...) were independently found to be raw Numbers wearing a
   text-looking display, same class as the percentage-column sweep earlier the same
   day.
4. **NOT a roll-forward bug, confirmed rather than assumed.**
   `ExcelOutput.RollForwardPeriod` copies every column verbatim already -- proven by
   reading it, not guessed. `Q4F26`/`Q1F27` (current) were blank in every `MS*`
   column while `Q3F26` had data because the milestone data was typed into `Q3F26`
   *after* the later periods had already forked from an earlier, blank version of
   it -- the same pattern `START_DATE`/`END_DATE` show on the identical three rows
   (blank in `Q3F26`, populated only from `Q4F26` onward). A one-time data copy into
   the current row, not a code change, is the correct fix for this part.

**Fix:** `MilestoneDevice.WriteText` now converts `LINE_BREAK_DELIMITER` to a real
line break, matching `InjectPrimitive`'s own writer -- new test
`MilestoneDevice_TwoLineLabelConvertsThePipeDelimiterToARealBreak`, made to fail
first (asserted the delimiter itself must not survive onto the slide; failed against
the unwrapped code exactly as expected, then passed once fixed). The `Q3F26` data
mistakes were corrected in place (label/date properly split, correct `"||"`
convention, `NumberFormat = "@"` locked on all 7 `MS*_DATE` columns before writing),
then copied into the current period's row -- a real backup was taken first
(`SaveCopyAs`, never `SaveAs`, per this project's own standing rule for touching
real data).

**Proven live, not just in the test suite**: applied through the real "2. Put it on
the slides" button, unaided, twice (once for the data/visibility fix, once more
after the `WriteText` fix reached a built addin) -- verified both times by reading
the saved `.pptx`'s own slide XML directly. First pass confirmed every circle's
visibility state now matches the corrected `_DONE` flags exactly (achieved/current/
not-achieved, all four states). Second pass confirmed `MS3_LABEL`/`MS4_LABEL` are
now genuinely two separate text runs in the saved XML (`'Method exploration'`,
`'Pre-trial package complete'`), not one run containing literal `"||"` -- proof the
line break is real, not a screenshot artifact.

### AZ. FIXED 2026-08-19 — Error 50290's fourth occurrence, diagnostic capture extended to the queue-build/planning phase

Error 50290 (this item's own earlier entries, item V) recurred for a fourth time,
at a fourth different call site: this time during the queue-BUILD/planning phase
(`RibbonUI.BuildAllQueuesCore` -> `ReviewQueue.BuildQueue` -> `SyncOperations.
PlanRoutineSync`), before `ApplyApproved`'s own per-item trap (built for the third
occurrence) was ever reached. No new Sync Log line appeared and PowerPoint was not
mid-macro when checked afterward, confirming the crash landed in the planning chain
and nothing there captured `Err.Description`/`Source` at the point of failure --
the top-level dialog said only "VBAProject" again, same as every prior occurrence
before its own call site got wrapped.

**Root cause is still open, deliberately not chased further** -- four occurrences
in four different call sites across four sessions has already ruled out a
single-function cause, and Office cannot be made to raise the real fault on demand.
This entry is about closing the diagnostic gap at this NEW call site, matching the
pattern already proven for `ApplyApproved`.

**Fix, built by a Fable-model agent from a detailed brief, verified and finished in
the main session:** a shared `ReviewQueue.LogAndReraiseCrash` helper (log-before-
re-raise, enriched `Err.Source`) now backs per-item traps in `PlanRoutineSync`
(the resolve loop, the device-tag scan, the elapsed-bar lookup, and -- the actual
site tonight's crash hit -- the main per-field probe, via a new `GuardedPlanProbe`
wrapper) and per-stage traps in `BuildQueue` (gather, parity resolve) and
`BuildAllQueuesCore`. A new gated test-only hook, `SyncOperations.
mTestForcePlanCrash` (matching `ReviewQueue.mTestForceInjectCrash`'s exact shape --
module-level, declared with the Consts rather than after a procedure, per this
project's own documented `check_vba_static.py` gap), simulates the fault
deterministically since the real one can't be summoned.

**Genuinely proven fail-first, not assumed**: the agent deliberately left the main
field-probe site unwrapped through its own work so
`Test_SyncOperations_PlanRoutineSyncNamesTheItemWhenProbeCrashes` could demonstrate
failing against real uncaptured code first -- run and confirmed failing (raw
`Err.Source` of just `"InjectPrimitive.InjectField"`, nothing logged to the Sync
Log) before the final wrap was applied in the main session, then re-run and
confirmed passing. Full suite 248/248 (one unrelated bug found and fixed along the
way: a test written earlier the same session had its own delimiter collision --
`"||"` used simultaneously as the intra-label line break AND the test helper's
inter-milestone separator -- fixed by building the fixture arrays directly instead
of through the string-splitting helper).

## Added and FIXED 2026-08-19 morning — BA, a consistency audit found FieldWiring's
## own copy of a bug already fixed elsewhere in this codebase

**BA. `FieldWiring.WalkForRoles` ONLY RECURSED INTO A GROUP AND NEVER READ THE
GROUP'S OWN ROLE TAG.** Same shape as `InjectPrimitive.WalkForRoleTag`'s
2026-08-10 fix (item recorded in that function's own header, not lettered here at
the time) -- a device's own role tag (`MILESTONE_TIMELINE`, stamped on the group
shape itself, not on any child) was invisible to `RoleTagsInDeck`/
`RolesByInstance`/`RoleTagsOnSlide`, the whole consistency-scanning family behind
`ScanFieldWiring`. Found by a cold consistency audit asked to check whether recent
field-type work (generic picture/bar/device routing, proven in items building up
to `PROJECT_PHOTO`) was reflected everywhere the codebase branches on field type,
not just where it was built. **Currently masked, not harmless**: nothing today
queries a device's own tag through this scanner, so there is no visible symptom
yet -- but it is the exact "device-owned column" blind spot item P3
(`FieldWiring` had no concept of a device-owned column) was fixed for, recurring
in a sibling walker P3 never touched. Fixed the same way as the 2026-08-10 fix:
test the shape's own tag, then always recurse, instead of `If msoGroup Then
recurse Else test`.

Same audit found `Discovery.bas`'s `IsPicture` function duplicating
`InjectPrimitive.IsPictureShape`'s body verbatim rather than calling it, the one
holdout after `BatchOnboardFlow.bas` already called the shared function directly.
Deduped to the shared function; no behaviour change, just one fewer place the two
copies could silently drift apart.

## Added and FIXED 2026-08-19 morning — BB, a long-running Excel automation
## instance found to silently refuse new workbook opens

**BB. `WorkbookBridge.OpenOrGetWorkbook` HAD NO RECOVERY WHEN THE ATTACHED EXCEL
INSTANCE SILENTLY REFUSED A NEW OPEN.** Probed live: `Workbooks.Open` on an
`Excel.Application` attached via `GetObject` (hours of uptime, many prior
automation calls already made against it that same session) returned Nothing with
`Err.Number = 0` -- no error, no workbook -- against a confirmed-valid,
confirmed-uncorrupted local `.xlsx`. A brand-new `CreateObject("Excel.Application")`
instance opened the identical file instantly, isolating the cause to the
long-running attached instance itself (a known-in-the-wild COM failure class: a
long-lived automation server going quietly deaf to new requests) rather than the
file or the pairing path. Surfaced to a person as "Could not open the paired
workbook at: ..." from every one of `OpenOrGetWorkbook`'s callers (Preview Sync,
Apply Approved, Set up my quarter, and more) with no indication the file itself
was fine. Fixed with one retry against a genuinely fresh instance when the
attached-instance open returns Nothing -- a truly bad/missing path still returns
Nothing through the retry too, so the function's documented contract ("Returns
Nothing if `path` doesn't exist and can't be opened") is unchanged for callers.
No unit test added: this depends on a long-running COM server going quietly deaf,
which the harness cannot simulate -- the live probing that found it (attached
returns Nothing, fresh instance succeeds) is the fail-first proof.

## Added and FIXED 2026-08-19 morning — BC, SVG-inserted picture shapes were
## invisible to the entire picture pipeline, found proving items BA/BB's fix live

**BC. `IsPictureShape` (`InjectPrimitive.bas`) AND `IsCandidateLeafType`
(`Discovery.bas`) ONLY RECOGNISED `msoPicture` AND `msoLinkedPicture`,
MISSING `msoGraphic` ENTIRELY.** Building the four `DELIVERABLE1_PHOTO..
DELIVERABLE4_PHOTO` fields (see `CHECKLIST.md`'s "RESOLVED 2026-08-19" entry)
put two PNG-sourced and two SVG-sourced picture shapes on the same slide, same
row, tagged the identical way. The two PNGs worked correctly end to end on the
first live test. **The two SVGs never even appeared in the review queue --
no error, nothing wrong-looking, just silent absence.** Probed live: PowerPoint
reports `shp.Type = 13` (`msoPicture`) for the PNG shapes and `shp.Type = 28`
(`msoGraphic`) for the SVG ones -- confirmed against the real shapes, not
documentation, since it auto-names SVG-inserted shapes "Graphic N" instead of
"Picture N" for the same underlying reason. `IsPictureShape` backs both
`InjectorFor` (routing) and the review-queue's change-detection scan, so a
shape this function doesn't recognise as a picture is invisible everywhere,
not just at write time -- and `Discovery.IsCandidateLeafType`'s matching
allowlist meant an SVG-backed field could never even be offered as a taggable
candidate via "Tag fields on this slide" either. Fixed by adding `msoGraphic`
to both checks.

**Proven live, not just compiled**: rebuilt the `.ppam` (`build_ppam.ps1`,
addin145), registered it `AutoLoad` in place of the stale `addin139`,
reopened the deck fresh, re-ran "Review changes" (now correctly queued 3
changes instead of 1) and "Put it on the slides" (3 written, 0 failed, 0
stale). Verified from the saved file's own bytes, not the dialog: all five
picture fields on `3_P001` (`PROJECT_PHOTO`, `DELIVERABLE1..4_PHOTO`) now
carry a correct `PICSRC` tag matching their register source ID.

## Added 2026-08-19 morning — BD, a real field-coverage notification (was a
## one-off diagnostic, never a feature), and the uncached scan underneath it

**BD. THE FIELD-COVERAGE QUESTION ("is the register complete for this
slide type") HAD NO REAL ANSWER -- Rohan asked directly, after the BC
incident, why the 2026-08-16 "Field Coverage Matrix" (`CHECKLIST.md`) was
never real code.** Root-caused from the actual commit (`58b9631`): it
touched zero lines outside `CHECKLIST.md`, built by hand into a throwaway
test workbook via ad hoc PowerShell/COM, never a VBA function, button, or
test. The tested, shipped mechanism that sounds similar (`cd9a6a6`,
`FieldWiring.bas`'s per-slide coverage counting) answers a different
question -- breadth across instances, not completeness across fields for
one type.

Two real changes, not one:

1. **`RolesByInstance`/`RoleTagsOnSlide` were doing a live, uncached
`shp.Tags("role")` walk of every shape on every slide of a type, on every
single call** -- and this scan runs from "1. Set up my quarter," pressed
every session, not once ever. The exact cost class `ShapeAddressBook.bas`
exists to eliminate (items W, Y, AT), never applied here. Fixed: both now
route through `ShapeAddressBook`'s shared per-slide tag index first,
falling back to the existing full walk only on a genuine miss and
populating the cache afterward -- shared with, and shared benefit from,
`InjectPrimitive.FindShapeByRoleTag`'s own fast path. Proven the cache
path is actually taken, not just "still correct either way": a scenario
where a stale cache and a fresh walk would disagree, confirming the stale
answer wins (`Test_FieldWiring_ScanUsesTheSharedShapeCache`).

2. **A real, consolidated notification**, reversing the 2026-08-14 modal
removal deliberately -- that removal was correct for what it fixed
(`MS1_LABEL`..`MS7_DONE` false-positiving on every press, since a device's
internals were being asked about individually). That source is now fixed
(`MilestoneDevice.IsColumnForThisDevice` excludes them entirely), so
`RibbonUI.OfferMarkingForUnwiredFields` now shows ONE `MsgBox` across every
registered slide type, naming both the field AND (for partial coverage)
the specific slide instance keys missing it
(`FieldWiringResult.MissingDetail`, new). Never blocks -- still a notice,
sync continues either way.

Full suite run for real (not just compiled) via `run_vba_tests.ps1`:
static checks clean, compile clean, **255 passed / 0 failed / 0 skipped**.
Three new tests, cache-is-used and case-preservation both fail-first
proven the same way `FastPathActuallyFires` was. **Not yet deployed to a
live add-in build** -- `addin145` predates this commit.

## Added 2026-08-19 (still morning) — BE, the BD modal turned out to need
## three more real rounds before it actually worked live

**BE. BD's notification looked right in code and still broke three
separate ways once actually pressed against the real, messy 45-slide
deck** -- each one caught by watching it live, not by re-reading the
code, and each one a genuine defect, not a restyling:

1. **Silent mid-word truncation.** Windows `MsgBox`'s real, undocumented
~1024-character limit -- the same failure class `FieldSpec.
ApplyControlledValidation` was already fixed for -- reproduced here
because the new modal never got the same treatment. Fixed two ways:
`FieldWiringResult.MissingDetail` caps each field's named slide list to 6
keys (count stays exact, "+N more" covers the rest); `RibbonUI.CapReport`
(this project's established "one place that knows about the limit," not
a second answer to the same question) got the call, with its notice
overridden.

2. **A false claim.** The modal said "full detail is also in the Run
Log" -- untrue, because `WorkbookBridge.WriteRunLog` clears the whole
sheet on every call ("REPLACED each run, not appended"), and several
more calls happen later in the same chain. Confirmed live: the entry was
gone from the saved file by the time the run finished. Dropped the claim
rather than leave a promise the code can't keep.

3. **A modal-count regression, and a real (if narrow) structural risk
found investigating it.** BD's notification fired as its own standalone
`MsgBox`, called before `DraftingUI.BeginCollecting` -- pushing "Set up
my quarter" from `LOBBY-DESIGN.md` section 6's documented ~2-modal
target back up to 3. Investigating why that mattered found something not
previously documented: `AppEvents.cls`'s `mApp_SheetChange` handler
watches Excel's events via `WithEvents`, but the handler code runs
*inside this PowerPoint-hosted VBA project* -- so any modal left open in
PowerPoint blocks its single VBA thread from servicing that call, and
ticking an Approve-column cell in Excel while one is open would hit
"source application may be busy," no external automation required.
Fixed by exposing `DraftingUI.AppendCollected` (mirrors `Say()`'s own
collecting-branch behaviour, which is `Private` to `DraftingUI.bas`) and
moving the notification inside the same collecting window
`StartQuarter`/`RollForwardUI`/`RefreshDraftingSheets` already share --
back to 2 modals.

**Then a fourth found live, folding those four sections into one
dialog**: even individually capped, the coverage section still got
chopped by the *outer* `CapReport` call wrapping all four sections
together, because the other three already used most of the shared
900-char budget on their own -- capping one section against the *full*
shared ceiling does nothing to stop the *combined* text from exceeding
that same ceiling. `CapReport` gained an optional `maxLen` (0 = the
shared `REPORT_CAP`, unchanged for every other caller) -- and even a
350-char section-level cap still weren't enough headroom. Asked Rohan
directly: richer detail in its own dialog (back to 3) or a short summary
kept folded (stays at 2). **"Folded, I think."** New `FieldWiring.
CoverageSummaryLine` gives counts only ("11 not on any slide, 16 missing
from the template, 7 partially covered") -- `BlockingText`, the full
detail, is untouched and available if a future persistent view (a real
Field Coverage sheet, the thing `CHECKLIST.md`'s one-off never became)
ever needs it.

**Proven live end to end, five add-in builds (`addin146`-`addin150`),
not just the last one** -- each round rebuilt, redeployed
(`AutoLoad` moved forward each time, prior build disabled), and
re-tested against the real deck before moving to the next fix, so each
claimed fix was actually verified rather than assumed. Full suite run
for real after every code change; final state **258 passed / 0 failed /
0 skipped**. `addin150` is the current live build.

## Added and FIXED 2026-08-19 — BF, the comprehensive readiness audit's own
## test-coverage gap: zero tests existed for tonight's msoGraphic fix or the
## SVG picture pipeline it unblocked

**BF. THE 6-LAYER READINESS AUDIT (run at Rohan's request after item BE,
"ensure these items are what drives us from hereon in") named its own
Layers 5&6 finding explicitly: zero tests reference `msoGraphic`,
`IsPictureShape`, or `IsCandidateLeafType`, and zero tests name
`DELIVERABLE`.** Item BC's fix (`msoGraphic` added to both checks) and
the whole SVG-sourced picture path it unblocked had shipped and gone
live across five add-in builds with no regression coverage at all --
correct today, unprotected against tomorrow.

Closed by adding `TestRunner.MakeTestSvg` (mirrors `MakeTestBitmap`'s
self-contained-file pattern) and three tests: `IsPictureShape`
recognises an SVG-inserted `msoGraphic` shape; `Discovery.DiscoverSlide`
offers it as a taggable candidate; `InjectPrimitive.InjectPictureField`
writes and verifies through a real SVG locator end-to-end. Confirmed
live first, before writing any assertion around it: an isolated COM
probe (`Shapes.AddPicture` on a bare self-contained SVG) reproduced
`Type=28`, matching item BC's original finding rather than assuming it
still held.

Full suite: **261 passed / 0 failed / 0 skipped** (258 prior + 3 new),
and each of the three new tests independently reconfirmed passing on a
filtered re-run rather than trusted from the combined count alone.
Commit `9a0bf6f`.

## Added and FIXED 2026-08-19 evening — BG, six of the 11 audit-flagged
## fields tagged live on the REAL template slide (slide44), not a copy

**BG. THE 11 FIELDS THE READINESS AUDIT FLAGGED AS "NEVER POPULATED" TURNED
OUT TO ALREADY HAVE FIELD SPEC ROWS AND REGISTER COLUMNS -- WHAT THEY
LACKED WAS A SHAPE TAGGED ON THE REAL TEMPLATE.** Probed the live register
(`register-wide.xlsx`) and the live deck's actual registered template
(slide44, SlideID 303, resolved via the old-mechanism `CustomDocumentProperties`
fallback since the registry-slide shape for `DeckSyncType:project-progress`
was never written on this real deck) rather than trust `COLUMNS.md`/
`CHECKLIST.md`, which were both stale on this point. Zero of the 11 had a
role tag anywhere on the real template.

Split cleanly into four groups by what was actually missing:
- **Group A** (`SAAFE_CASH`, `TOTAL_INKIND`, `INDUSTRY_PARTNER`,
  `TERTIARY_INSTITUTION`) plus `DELIVERABLES_BODY` and `PROJECT_PHOTO`:
  the shape already existed on slide44 with the SOURCE project's real
  content sitting on it untagged (a fossil of `MakeTemplateFrom`'s own
  documented limit -- "untagged content is NOT touched, and cannot be").
  Pure tagging work, tagged all six for real.
- `HIGHLIGHTS_BODY`: shapes exist (three bullet groups) but `InjectPrimitive.
  bas` only ever writes one field into one shape -- no one-into-many
  mechanism exists. Needs real code, not a tag.
- `SUBTITLE_B`/`SECTOR`/`TRL`: spec'd to concatenate into `SUBTITLE_A`'s one
  shape; no many-into-one mechanism exists either. Same class of gap,
  mirrored.
- `SCHEDULE_STATUS`: feeds a `STATUS_BADGE` concept that depends on the
  already-known `Kind=Derived` gap (`CHECKLIST.md`, "confirmed NOT built").
- `STRATEGIC_LINKAGES`: its own Field Spec says "NO SHAPE on the
  project-progress slide" -- not a gap on this slide type by design.

**Tagging alone is not the whole of `MakeTemplateFrom`'s own safety
contract, and this closes only half of it for the six real ones.** Every
already-tagged text field on this template carries `<<FIELD>>` placeholder
text, replacing the source project's real value, specifically so a slide
cloned from the template cannot silently carry another project's data
looking correct (`TemplateSlide.bas`'s own docstring: "the worst shape a
reporting-tool defect can take"). Replicated that for the five text
fields -- tag write via `shp.Tags.Add "role", <name>`, then
`shp.TextFrame.TextRange.Text = "<<" & name & ">>"`, matching
`TemplateSlide.PlaceholderFor` exactly.

**`PROJECT_PHOTO` cannot get the same placeholder treatment, and this is a
real, currently-unclosed gap, not an oversight.** `InjectPrimitive`
(the function `MakeTemplateFrom` itself calls to blank a field) refuses on
any shape with `HasTextFrame = False` -- confirmed live, the placeholder
write on `Picture 235` failed with "the specified value is out of range."
A picture shape simply has nowhere to put `<<PROJECT_PHOTO>>` text. The
tag is real and correct, but the template's picture shape still shows the
source project's real photo, untagged-content-camouflage exactly as
`TemplateSlide.bas` describes it -- just for a Kind the module's own
blanking step was never built to reach. No existing mechanism in this
codebase addresses it (checked: this is the first picture field ever
tagged on this specific template). Left open, named here rather than
silently accepted.

**A real near-miss, caught only because of copy-first discipline.** The
first attempt at `DELIVERABLES_BODY` matched a decorative underline
connector 2pt away from the real text box -- both untagged, both within a
3pt position tolerance, and the connector's empty text passed a check that
only tested `HasTextFrame`, not `HasText`. Caught on the copy, before the
real deck was touched; fixed by also requiring non-empty existing text
(and, separately, allowing picture-type shapes through that same filter
for `PROJECT_PHOTO`, since a picture has no text to check). The real deck
never carried the bad tag.

**Every write proven from the saved file's own bytes, in a fresh process,
not an in-process readback** -- each of the six fields: written on a
scratch copy, reopened cold and verified, then the identical write applied
to the real deck (`C:\Users\rohan\OneDrive\Claude\3. Project
Progress.pptx`), reopened cold again and reverified. Backup taken before
the first real-deck write: `OneDrive\Claude\backups\3-Project-Progress.
PRE-GROUPA-TAGGING-20260819-151035.pptx`.

**`DELIVERABLE1_PHOTO..DELIVERABLE4_PHOTO` turned out NOT to be ready for
this same treatment, discovered before any write was attempted.** Checked
the real Field Spec sheet directly: `PROJECT_PHOTO` has a real row (51) and
a real `Register` column. The four `DELIVERABLE*_PHOTO` fields have
**neither** -- last session's Field Spec rows, Sources rows (S12-S15) and
register columns for them were built and proven only on the
`deck-sync-test-deliverables` copy, never brought over to the real
`register-wide.xlsx`. Tagging their template shapes without that backing
would create fields the register cannot drive. Not done here; needs its
own register-authoring pass first, not a mechanical backfill.

## Added and FIXED 2026-08-19 evening — BH, DELIVERABLE1-4_PHOTO's register
## gap closed for real, PROJECT_PHOTO's camouflage gap closed, both proven
## with actual images on the real deck's template AND its live 3_P001 slide

**BH. Rohan's choice, offered as ranked options: close BG's two remaining
gaps (`DELIVERABLE1-4_PHOTO`'s missing register infrastructure,
`PROJECT_PHOTO`'s template camouflage) by porting the ALREADY-BUILT
placeholder mechanism from the test-copy register to the real one, end to
end, rather than leaving them empty or half-wired.** Register-side: added
Field Spec rows 52-55 (`DELIVERABLE1_PHOTO`..`4_PHOTO`, mirroring
`PROJECT_PHOTO` row 51's own pattern exactly -- `Kind=Static`,
`Behaviour=Fit inside`, `Renders=Picture`), 4 new `Register` columns, and
Sources rows S12-S15 -- honestly labelled `TEST/PLACEHOLDER`, matching
`PROJECT_PHOTO`'s own existing S11 row's "Rohan to replace with the real
photo" convention rather than disguising them as real citations. Confirmed
first, not assumed: `PROJECT_PHOTO` already had a real Field Spec row,
register column AND a populated `3_P001`/Q1F27 register value (`S11`) --
only `DELIVERABLE1-4_PHOTO` needed the infrastructure built.

**Placeholder images moved from the scratch test folder to
`OneDrive\Claude\images\`, matching where `S11`'s own placeholder already
lived** -- not left depending on `deck-sync-test-deliverables\`, which is
disposable scratch space.

**Tagged all 4 `DELIVERABLE*_PHOTO` shapes on TWO real slides, not one:**
the template (`slide44`, so future onboarding carries them forward) AND
the real, live `3_P001` slide itself (so today's deck actually shows them
-- template tags do not retroactively apply to an instance that predates
the tagging). Confirmed live: `3_P001`'s own picture shapes sit at
IDENTICAL positions to the template's, confirming the template really was
cloned from a slide just like this one.

**A real near-miss caught by copy-first discipline, again.** The first
`Application.Run "InjectPrimitive.InjectPictureField", ...` attempt failed
with a misleading "Sub or function not defined" -- not a missing function,
but VBA refusing to expose a Function whose return type is a custom `Type`
(`InjectResult`) across the `Application.Run` COM boundary. A throwaway
wrapper Sub was written (`vba/tools/InjectPictureProbe.bas`, unused in the
end, kept as a documented dead end) before abandoning that route.

**`PROJECT_PHOTO` is genuinely cropped; all four `DELIVERABLE*_PHOTO`
shapes are not** -- checked live via `PictureFormat.CropLeft/Right/Top/
Bottom` rather than assumed. That distinction matters because
`InjectPictureField`'s own two branches are NOT interchangeable: an
uncropped frame is fed via `Fill.UserPicture` in place; a cropped one MUST
be rebuilt (`Fill.UserPicture` and cropping are mutually exclusive --
probed and documented at length in `InjectPrimitive.bas` already, "two
passes of arithmetic did not fix it, because the arithmetic was never the
problem"). Since the real function couldn't be called directly, its EXACT
documented sequence was hand-replicated rather than reinvented: capture
Left/Top/Width/Height/ZOrderPosition and the crop values from the old
shape BEFORE anything changes, `AddPicture` at the old position with
`-1,-1` size, `LockAspectRatio = False` FIRST (the step five earlier
attempts in this codebase's own history missed), apply the captured crop
values verbatim, THEN apply the captured Width/Height, re-tag `role` and
`picsrc`, delete the old shape, restore z-order. Verified live: every
shape's final Width/Height matched its pre-injection value exactly, on
both slides, both `PROJECT_PHOTO` (rebuilt) and all four `DELIVERABLE*`
fields (fed in place).

**One real, pre-existing gap surfaced and confirmed harmless for THIS
case, not silently assumed:** `Fill.UserPicture` stretches to the shape's
exact box -- it does not implement `Field Spec`'s own `Fit inside` vs
`Fill the frame` distinction as separate scaling behaviour; that
distinction currently exists only in the Field Spec's prose, not in code.
Confirmed with Rohan directly rather than left as a silent risk: the
placeholder deliverable images share one base aspect ratio by design, so
no distortion occurs here -- but a real deliverable image with a different
aspect ratio than its shape would stretch, and nothing in the injector
would catch it.

**A real OneDrive save hiccup, twice, both recovered without data loss --
matches this project's own documented Office/OneDrive save-quirk
history.** The real deck is cloud-hosted (`FullName` resolves to a
`d.docs.live.net` URL, not a bare local path). Both times `Presentation.
Save` failed on the first attempt with a generic "An error occurred while
PowerPoint was saving the file" -- once recovered by reattaching and
retrying `Save` directly (succeeded second try), once the reported failure
turned out to be cosmetic: the save had actually already gone through,
confirmed by reopening cold and reading the saved bytes. Neither is a
`.ppam`/VBA problem; both are the known cloud-save contention class this
project has hit before.

**Proven from the saved file's own bytes, in fresh processes, on every
target -- register (copy, then real) and both slides (copy, then real,
tagged, then injected).** Nothing here trusted an in-process readback.
Backups taken before each real write: `OneDrive\Claude\backups\register-
wide.PRE-DELIVERABLE-PHOTO-FIELDS-20260819-154008.xlsx`, `...3-Project-
Progress.PRE-DELIVERABLE-TAGGING-20260819-154228.pptx`, `...3-Project-
Progress.PRE-PICTURE-INJECTION-20260819-155430.pptx`.

## Added and FIXED 2026-08-19 late evening — BI, `Fill.UserPicture` silently
## no-ops on `msoGraphic` (SVG) shapes -- a real defect in `InjectPictureField`
## itself, found only by actually looking at the rendered slide

**BI. RENDERED, DID NOT JUST TRUST THE TAG.** Rohan: "open the deck and take
a look." Exported `3_P001` to PNG and looked at it -- `DELIVERABLE1_PHOTO`'s
card rendered blank/white next to three correctly-rendered siblings, despite
BH's own byte-level checks (role tag, `picsrc` stamp, geometry) all passing.
First theory, WRONG: chased it as a colour bug in the placeholder SVG file
(white-on-white via a stray Office theme CSS class) -- fixed the file,
re-fed it, still blank. **Second theory, confirmed live:** fed a
deliberately distinct, unmistakable "MARKER" image (solid red, huge white
text) into `DELIVERABLE2_PHOTO` -- a card that WAS rendering correctly -- via
`Fill.UserPicture`. It reported success. The rendered card did not change
at all. `Fill.UserPicture` does not touch an `msoGraphic` (type=28,
SVG-backed) shape's actual content; it silently does nothing while
reporting success and letting the `picsrc` tag update regardless.

**This is a real defect in the shipped `InjectPrimitive.InjectPictureField`
function, not just tonight's scratch scripts.** Its "uncropped" branch calls
`Fill.UserPicture` unconditionally with no distinction between `msoPicture`
(type=13, proven to work -- see the function's own probed comment, which
correctly scoped its claim to "type 13" and never claimed type 28) and
`msoGraphic`. The existing sibling test
(`Test_InjectPicture_SvgSourceFillsAndStamps`) never caught this because it
only asserts the role tag and `picsrc` stamp -- both written unconditionally
right after the `Fill.UserPicture` call regardless of whether the call
actually did anything. Exactly this project's own named failure class:
"reports success without confirming the effect."

**New regression test written BEFORE the fix, confirmed to fail against the
unfixed code first.** `Test_InjectPicture_SvgGraphicIsRebuiltNotFedInPlace`
-- since a shape's rendered picture content can't be inspected from VBA, it
checks the one thing that genuinely differs between the two code paths: a
rebuilt shape (delete + fresh `AddPicture`) gets a new `Shape.Id`; an
in-place `Fill.UserPicture` feed keeps the same one. Ran it against the
unfixed function first: **FAIL, "same shape Id (3) means Fill.UserPicture
ran and silently did nothing"** -- proof the test is actually sensitive to
the bug, not just a check that happens to pass either way.

**Fix: `msoGraphic` shapes now always take the rebuild branch, regardless of
crop.** One added condition (`isGraphic = (shp.Type = msoGraphic)`,
OR'd into the uncropped-branch guard) -- the rebuild branch already existed
and was already proven correct (it's the same path `PROJECT_PHOTO`'s crop
handling uses). Full suite after the fix: **262 passed, 0 failed, 0
skipped** (261 prior + this one new test, now passing).

**Real deck re-fixed for real, not just the code.** `DELIVERABLE1_PHOTO`
and `DELIVERABLE2_PHOTO` had never actually been written on either the
template or `3_P001` all evening (BH's claimed success was the tag/stamp
only) -- re-ran the injection on both real slides with the fixed logic,
re-exported and looked at `3_P001` again: all four deliverable cards render
correctly now, `PROJECT_PHOTO` unaffected (it was already on the rebuild
path via its own crop). `addin151` built (the actual `.ppam` "Save As"
click is a confirmed-permanent manual step, per this file's `build_ppam.ps1`
header -- Rohan did it), registered `AutoLoad`, `addin150` disabled.

**One thing Rohan corrected me on directly:** `PROJECT_PHOTO`'s crop/fill
behaviour is intentional and correct as-is; the deliverable cards' aspect
ratio matching the source image is a non-concern by design (all four
placeholder images deliberately share one base aspect ratio) -- confirmed
before spending any effort building aspect-preserving logic that was never
needed.

## Added and FIXED 2026-08-19 night — BJ, `1_S004`'s `PROJECT_STATUS` badge
## bled its text into the money grid above it: an 18pt font on a
## 79x20 box, where every other slide's same field renders at 8.5pt

**BJ. Found spot-checking other real slides' card areas after BI, at
Rohan's request ("check the other cards on real slides too") -- not the
defect being hunted, a different one noticed along the way.**
`1_S004`'s "Not Started" badge showed text visibly overflowing upward into
the "SAAFE Cash $ AUD" box above it. Checked `PROJECT_STATUS`'s font size
across 8 other real slides (P and K types, plus one other S type) before
touching anything: every one of them renders at 8.5pt (one at 7pt) --
`1_S004`'s tagged shape was set to 18pt, over double the next-largest.

**Also structurally different from every other slide, not just the font
size:** `1_S004`'s tagged `PROJECT_STATUS` shape is an AutoShape ("Shape
7"); every other checked slide's tagged shape is a plain TextBox ("Text
N"). An untagged plain-text duplicate sitting at the identical position
with correctly-sized text was already there, unrelated and unedited --
this fix touched only the tagged shape's font size, not the duplicate
structure. **Worth a look in the field-by-field formatting check Rohan
named as the next standing activity now that reachability is largely
covered** -- this AutoShape-vs-TextBox difference may be why the font size
drifted in the first place (a different default), and the untagged
duplicate is dead weight that would go stale if `PROJECT_STATUS` ever
changes on this project.

**Fixed: `Font.Size = 8.5`, matching the dominant convention across the
deck.** Proven on a copy first (exported and visually confirmed the bleed
was gone), then applied identically to the real deck and reverified from
the saved file's own bytes. Backup: `OneDrive\Claude\backups\3-Project-
Progress.PRE-S004-FONTFIX-20260819-165132.pptx`.

## Added 2026-08-19 night — BK, `PROJECT_STATUS`'s vocabulary drift is
## real and quantified, confirmed against the live register, not fixed

**BK. `COLUMNS.md`'s standing "Register field questions" item ("`PROJECT_
STATUS` casing disagrees with the deck... may already be moot post the
2026-08-15 casing fix, check before treating as open") was checked
tonight, against the real register, and it is NOT moot.**

Ran `FieldSpec.ApplyControlledValidation` for real (not read about it --
called it directly via `Application.Run` against the live `register-wide.
xlsx`), which is this project's own established mechanism for exactly this
question: it never silently rewrites a controlled field, it only applies
Excel dropdown validation and REPORTS what's outside the declared
vocabulary ("left exactly as they are").

**Real result: 17 values outside `PROJECT_STATUS`'s declared vocabulary**
(`In Progress,Not Started,Project Closed`, from `FieldSpec.bas`'s own
seeded `COL_S_ALLOWED`) **across 258 controlled cells checked.** Two
distinct kinds of drift, not one:
1. **Casing** -- `In progress` (lowercase p) instead of `In Progress`,
   confirmed on at least 2_P003/2_P004/1_P005/1_P006/1_P007 (Q3F26) and
   more not shown (the function's own report caps at 5 examples by
   design, count is exact).
2. **A genuinely different value, not just casing** -- `Not yet commenced`
   appears at least once, which isn't in the 3-item vocabulary at all
   (distinct from `Not Started`, not merely differently capitalised). This
   is a real semantic question, not a typo fix -- may be a deliberate
   distinct state someone typed before the vocabulary was locked to 3
   values, and only a person can say whether it should collapse to `Not
   Started` or stay its own thing.

**Real, positive side effect from running this**: 258 cells across both
controlled columns now carry a live Excel dropdown for the first time,
which stops any NEW drift even though it correctly left the 17 existing
offenders untouched. Confirmed saved.

**Not fixed. `DeriveStatusBadge` (see `SyncOperations.bas`) already
tolerates the casing half** -- its `PROJECT_STATUS`/`SCHEDULE_STATUS`
comparisons use `vbTextCompare`, so `In progress` still correctly resolves
to the `In Progress` branch. It does NOT tolerate `Not yet commenced` --
that value falls through every branch and refuses (returns `""`) rather
than guessing, the same "refuse rather than draw a wrong bar" instinct as
`ElapsedFraction`. **Whether that's the right behaviour for `Not yet
commenced` specifically is Rohan's call, not resolved here.**

**RESOLVED, same session.** Rohan: *"let's just have one way of writing it
through."* All 17 normalized to the declared vocabulary exactly -- 8
`In progress` -> `In Progress`, 8 `Not started` -> `Not Started`, 1
`Not yet commenced` -> `Not Started` (his explicit call, not a guess).

**A real bug caught building the fix, not the fix itself.** The exact
offender list was first pulled with PowerShell's `-contains`, which is
CASE-INSENSITIVE by default -- it silently hid every casing mismatch and
reported only 1 offender instead of 17, the opposite of the 17
`ApplyControlledValidation` had already correctly reported. Cross-checked
against that authoritative count rather than trusting the smaller number,
found the tool mismatch, redid it with `-ccontains`. **The exact same bug
class, self-inflicted, minutes later**: the first normalize script used a
PowerShell hashtable for its old-value/new-value map, whose key lookups
are ALSO case-insensitive by default -- `$map.ContainsKey("In Progress")`
matched the key `"In progress"` and silently rewrote all 112 already-
correct cells too (harmlessly, since the case-insensitive value lookup
happened to return the same text back, but not what was intended, and not
something to wave through). Caught by checking the change count (112)
against the expected count (17) before trusting the log, not by any
special insight -- fixed with an explicit ordinal (case-sensitive) string
comparison, re-tested on a copy, re-verified `112 -> 17`.

Proven on a copy first, then applied to the real register with a fresh
backup taken immediately before
(`OneDrive\Claude\backups\PRE-STATUS-NORMALIZE-20260819-1940\`), then
reverified from the saved file in a fresh process: 0 remaining
out-of-vocabulary, `1_K010`'s `Q3F26` row confirmed reading `Not Started`.

## Added 2026-08-20 morning — BL, queued idea (not built): a hound for
## decision-propagation drift -- a design change lands in code but not in
## every doc/data cell that describes it

**QUEUED, NOT BUILT. Rohan's call, 2026-08-20: "queue it and let's draft"** --
noted so it isn't lost, deliberately deferred rather than built today, to get
to actual content drafting instead of more infrastructure.

**Why it's worth building.** The same defect shape recurred SIX times in one
morning while fixing the drafting sheet's column redesign (removing
`COL_D_CURRENT`, making `REPORTED LAST TIME` a hybrid ferry+register read):
`WorkbookBridge.DescribeSheet`/`LifespanOf` (fixed 2026-08-19 night, item BK's
neighbour), `field_e2e.ps1`'s module list (stale names, twice, one causing a
live VBA compile error that blocked all COM automation), `read_deck_props.py`
and `preflight.py` (both reading the deprecated `docProps/custom.xml`
fallback instead of the real `DeckSyncRegistry` hidden-slide storage that's
been primary since 2026-08-16 -- every deck-period check made with these
tools was silently reading stale data), the Field Spec sheet's single shared
`GLOBAL RULES` cell (told drafters "if column C already does its job, leave
the row blank" when C had just been repurposed to mean something the
instruction no longer fit), and five separate `.md` docs (`WORKFLOW.md`,
`SYSTEM-OVERVIEW.md`, `SOURCE-CAPTURE-FORM.md`, plus two already-correctly-
exempted ones) describing a column layout that no longer existed.

**The shape a hound would hunt:** a hardcoded name, a hand-typed table, a
seeded cell of prose, or a doc's own worked example -- anywhere a fact that a
single source of truth (a VBA constant, a function, a schema) already
determines gets RESTATED somewhere else, with no mechanism forcing the
restatement to track the source when it changes. Distinct from the existing
hounds: not about reachability, provenance, naming conflation, or waste --
specifically about a SECOND COPY of a fact that can silently disagree with
the FIRST.

**Where to start if built:** the six sites above are a ready-made seed
corpus -- a hound built against this pattern should find them (and any
siblings still lurking) as its first real test.

## Added and FIXED 2026-08-20 afternoon — BM, a class-wide "second unguarded
## Excel/PowerPoint handle" bug, found via a diagnostic script's own silent
## write loss

**Found live.** A PowerShell write script opened `register-wide.xlsx` with a
fresh `CreateObject("Excel.Application")` while a PRIOR automation step
(`field_e2e.ps1 -Mode refreshdrafting`) had left the same file open in a
different Excel process. The second open silently succeeded (no error),
the writes appeared to apply in memory, `.Save()` raised nothing -- and
none of it reached disk. Read back via COM directly (bypassing the write
script entirely): every field's History-treatment cell was blank. Classic
"reports success without confirming the effect" (the project's own
standing rule), one level removed -- the CHECK that would have caught it
(reading the saved bytes back) was run, and still nearly missed it,
because the first read used the SAME broken open pattern.

**The fix, rolled out to every site with the same shape, not just the one
that broke.** `WorkbookBridge.OpenOrGetWorkbook` already guarded
PRODUCTION code (`RibbonUI`/`OnboardFlow`) against this exact class --
matches by path against every already-open workbook first, opens fresh
only if none match (FIX-LIST item AU is the same defect one layer up: the
match itself silently failing on a cloud path). It gained a
`wasAlreadyOpen` output param so a caller knows whether it owns the
workbook it got back -- a caller that Closes/Quits a workbook it merely
found already open would discard or kill whatever else had it open. Six
functions in `vba/tools/E2EField.bas` and one in `AuditRealDeck.bas` that
previously opened the register with a raw `CreateObject+Open`, bypassing
the guard entirely, now route through it. Pure-observation tools
(`ReadWidePeriod`, `PreviewRealDeck`, `VerifyRealDeck`, `HiddenFixCheck`,
`SyncRealDeck`, `InProcessTimingProbe`) were deliberately left alone --
they exist specifically to read disk truth, and reusing an already-open
instance would trade that guarantee for a small speed gain that doesn't
matter given these run as standalone cold-start scripts anyway.

**A regression, caused and fixed the same session.** The fail-first test
proving `wasAlreadyOpen` (`Test_WorkbookBridge_
OpenOrGetWorkbookDetectsAnAlreadyOpenFile`) calls `OpenOrGetWorkbook`
twice against a real temp file -- and `OpenOrGetWorkbook` has its own,
pre-existing side effect: it calls `ShapeAddressBook.SetActiveWorkbook`
on every resolve. The test closed its workbook and quit its Excel
instance without resetting that cache to `Nothing`, leaving
`ShapeAddressBook`'s module-level `mWb` pointed at a disposed COM object
for the rest of the suite run. Ten unrelated tests
(`InjectField`/`InjectPicture`/`FieldWiring`/`ReviewQueue`) then failed
or errored with "no single shape found" / "The object invoked has
disconnected from its clients" -- passing clean every time when run in
isolation, which is what pointed at shared state rather than a logic
break. Fixed with the same reset every OTHER test touching that cache
already used (`ShapeAddressBook.SetActiveWorkbook Nothing`). Full suite
green after: 278/278.

**The addin-registration gotcha this did NOT catch**, because it's a
PowerPoint `AddIns` collection issue, not an Excel-workbook one: `File >
Save As` lands a freshly built `.ppam` in `OneDrive\Claude\`, which is
NOT the trusted location every prior addin actually lives in
(`C:\Users\rohan\AppData\Roaming\Microsoft\AddIns\`). `AddIns.Add(path)`
registers AT WHATEVER PATH IT'S GIVEN rather than copying the file in --
registering straight from `OneDrive\Claude\` silently "worked"
(`Loaded=True`) while pointing at the wrong folder. Same underlying shape
(a handle registered somewhere other than where everything else expects
it to live), different surface. **NOT a new discovery -- `AGENTS.md`
already documented this exact failure from two prior nights
(`addin133`/`addin134`), unread before repeating it a third time.**

**Worse, fixing the path masked a second miss of the identical shape:**
having moved the file and set `.Loaded = -1`, reported the addin
registered -- without setting `.AutoLoad`. `.Loaded` is true for the
CURRENT session only; `.AutoLoad` governs the NEXT PowerPoint launch. A
check in a genuinely fresh PowerPoint instance (quit entirely, relaunched,
nothing touched) found `addin155` at `AutoLoad = 0` and the PREVIOUS
build (`addin153`) still at `AutoLoad = -1` -- the very next normal
PowerPoint open would have loaded the OLD addin, silently, with the
session-level check having reported success the whole time. Fixed:
`.AutoLoad = -1` on the new build, `.AutoLoad = 0` explicitly on the
old one, re-verified from a second cold quit-and-relaunch with nothing
touched in between. Full corrected workflow, including this step and the
cold-restart verification requirement, now in `CHECKLIST.md`'s "Before
rebuilding the addin" checklist and reinforced in `AGENTS.md`.

## Added and FIXED 2026-08-20 evening — BN, investigated a backup for the
## ordinary drafting-sheet rewrite, built a test instead

**The question, from Rohan**: the parking mechanism (`ParkSheetCopy`,
`SAVED HH:MM <field>` sheets) only fires on a layout migration or a
period turnover -- never on the ordinary same-layout/same-period rewrite,
which is the path that runs on every single "Set up my quarter" press.
`Drafting.bas`'s own 2026-08-17 comment already flagged this as "still
open, not decided." Is that backwards -- protecting the rare case and
leaving the common one exposed?

**Traced the actual row-write logic before answering.** The ordinary path
writes ONLY `Project code` and `Project name` per row -- confirmed at
`Drafting.bas`'s own header above the loop: "THE PERSON'S COLUMNS ARE NOT
WRITTEN HERE. THAT IS THE FIX." `SOURCES`, `AI DRAFT`, `SUBMIT`,
`APPROVE` and `NOTES` are never touched, never cleared, never read into
memory and written back. So a backup on this path would protect against
nothing that happens today -- only a FUTURE regression that made the
ordinary path start touching those columns again.

**Costed a naive "always park" fix and rejected it.** `ParkSheetCopy`
does a full `Sheet.Copy` per field; doing that on every routine press
across ~13 Prose fields adds real seconds to the most frequently-run path
in the tool (the same class of cost item AB/W already paid for once this
project, on a different function). Worse: since `PruneParked` caps at 2
per field, making the ordinary path ALSO park would mean the workbook
PERMANENTLY carries up to 26 `SAVED` tabs, cycling on every run, instead
of the rare event it is today -- directly reintroducing the exact clutter
found and cleaned from the live working register the same afternoon (26
stray `SAVED` sheets, from two genuine layout-6→7 migrations, at
`PruneParked`'s correct 2-per-field cap -- not a pruning bug).

**Built the cheap thing instead**:
`Test_Drafting_OrdinaryRewriteLeavesThePersonsColumnsUntouched`, which
enforces the invariant Drafting.bas's comment only asserted in prose, at
zero runtime cost, and would catch the actual risk (a future regression)
before it ships rather than paying for a backup against a risk that
isn't live. Fail-first proven: breaks with the exact expected failure
("SUBMIT survives an ordinary rewrite untouched") when a deliberate
`ws.Cells(r, COL_D_SUBMIT).Value = "..."` is added to the ordinary path,
passes clean once reverted. `DOCUMENT-MAP.md` decision 6 updated with a
pointer to this and to the four History treatments (item above this
one's neighbour in `SYSTEM-OVERVIEW.md`) -- the open question that
comment posed now has an answer, not just a fix.

## Added and FIXED 2026-08-21 — BO, `Template Audit`'s own findings had
## been sitting unactioned since before this project's live-leak-cleanup
## work even started

**Rohan asked what the other Excel sheet tabs looked like** while checking
in on tonight's field-propagation work. `Template Audit` (a 51-row, 6-
column sheet: `Shape | Inside group | Text on the slide | Guess | On how
many other slides | Decide: field / chrome / drop`) turned out to already
have found the exact leaked title-label defect (`Shape 229`, "3_P001
Timeline") fixed independently tonight through a completely different
path (visually inspecting the milestone widget) -- with the right guess
("LIKELY PROJECT DATA -- on no other slide of this type") already sitting
in row 12. **Every one of the sheet's 50 data rows has an empty `Decide`
column.** The audit ran, guessed correctly in every case checked, and
nobody ever closed the loop -- not a tooling gap, a process gap: a finding
sitting unactioned reads identically to no finding at all until someone
goes and reads the sheet.

**Working through every "LIKELY PROJECT DATA" row against the live deck
(not trusting the sheet as current -- it's a point-in-time snapshot) found
real, still-live leaks the sheet's own text pointed at directly**: four
untagged standalone shapes on the P template (`slide44`) showing
`3_P001`'s real content in the open -- `"90%"`, two milestone-achievement
sentences ("Found lead compounds...", "Lead candidates were safe..."),
and `"Project closed 2026"`. None were part of the `MILESTONE_TIMELINE`
group already fixed earlier the same night, so the earlier fix never
touched them. Content-verified against the sheet's own quoted text before
writing, then blanked and re-verified from the saved file's bytes -- same
discipline as every other live-file write tonight. Two of the sheet's
other flagged dollar figures (`$275,597`, `$1,371,209`) no longer exist
anywhere in the current deck -- resolved by something else before tonight,
not a live leak, the sheet is simply stale on those two rows.

**A second question the sheet's rows raised, genuinely investigated
rather than assumed either way**: the `MS1/3/4/5/7_LABEL` shapes inside
the milestone widget are correctly hidden (`Visible=0`, from the earlier
fix the same night) but their TEXT is still `3_P001`'s real content --
does a future real sync risk showing another project's milestone story
through them? Read `MilestoneDevice.DrawMilestones`'s actual draw loop
(lines ~594-597) rather than guessing: every slot that becomes visible
gets `WriteText` called in the same operation, before or alongside
`SetVisible`. The stale hidden text is inert -- any real sync overwrites
it before it could ever render. Left as-is; the residue is a `.pptx` file
you'd have to open the XML to see, not a rendering risk.

**Net for `Template Audit` itself**: worth re-running fresh rather than
trusted as current (tonight's fixes have already resolved a chunk of what
it flagged), and worth deciding whether `Decide` gets filled in as a
standing habit going forward or the sheet gets treated as disposable and
rebuilt each time -- currently it's neither, which is how a correct
finding sat for this long.

## Added and FIXED 2026-08-21 — BP, a delegated cleanup silently undid an
## earlier fix, and a non-breaking hyphen hid a leak from every text search

**The regression.** K and S's `MILESTONE_TIMELINE` widgets were correctly
recoloured (green to K's orange / S's lavender) earlier the same night. A
few hours later, a fork was asked to fix a separate, real problem -- the
widget showing too many circles and real `3_P001` data on the templates --
and its fix, correctly scoped to that problem, deleted the (by-then-
recoloured) K/S copies and re-cloned them fresh from the P template,
explicitly not touching colour (its own report said so). That silently
reverted the earlier recolour work, and nothing after it redid the fix --
the 42-real-slide propagation that followed cloned FROM the now-reverted
K/S templates, so the regression spread to all 32 real K/S slides too.
Found only because Rohan looked at the live deck directly and asked why
the bar was "often in front of circles? and wrong colour" -- the z-order
half turned out correct everywhere checked (probably a stale screenshot),
but the colour half was real, confirmed on all 34 widgets (2 templates +
32 real slides), and fixed the same way as the original recolour.

**The lesson, stated as a shape**: delegating a fix for problem A, when
the target already carries an unrelated fix for problem B, needs an
explicit check afterward that B survived -- "fixed A" and "did not touch
B" are two different claims, and a rebuild-from-source step can satisfy
the first while silently failing the second without anyone asking it to.

**Separately, a leaked achievement-sentence on the P template
("...helped stop biofilm formation...") survived two earlier search
passes that should have caught it**, including a search across all 47
slides for the literal string. The text uses `Chr(8209)`, a non-breaking
hyphen (U+2011), not the ordinary ASCII hyphen (U+002D) the search
patterns used -- so "Silver-enhanced" as typed never matched, even though
the word renders identically on screen. Found only by matching on
hyphen-free anchors ("biofilm formation") once the mismatch was suspected.
**Any earlier search this session that matched on a hyphenated phrase has
the same blind spot and has not been re-checked.**

## Added and FIXED 2026-08-21 — BQ, K/S templates were missing 4 fields
## every real K/S slide already carries, and Harvest.bas had E2EField.bas's
## SUBTITLE_A blind spot too

**K/S template gap.** `TERTIARY_INSTITUTION`, `INDUSTRY_PARTNER`,
`TIMELINE_ELAPSED` and `SAAFE_CASH` are tagged and live on all 32 real K/S
slides (from earlier tonight's propagation work) but were never tagged on
the K or S templates themselves -- that propagation targeted real slides
only, and nobody separately checked the templates carried the same set.
Worse than a plain gap: K's `INDUSTRY_PARTNER` shape held a hardcoded
literal `"[TBC]"` string, not a placeholder tag, so no injector could ever
have written into it. Found candidates by geometry match against the P
template's own reference positions (S needed a slightly wider search --
its layout carries an extra "Research Supervisor" row P/K don't have,
shifting the partner/institution block down). All 8 tagged, the literal
`[TBC]` reset to the proper `<<INDUSTRY_PARTNER>>` placeholder, verified
from the saved file's bytes.

**`Harvest.bas`'s `ShapeIsNotHarvestableText`** had the identical blind
spot `E2EField.bas` did before it was fixed a few hours earlier the same
night: `SUBTITLE_A`'s shape structurally looks like an ordinary text
field, so the existing injector-routing check would wave it through as
harvestable. What it displays is a middot-joined composite of four
register columns, not its own raw value -- harvesting it would write the
whole rendered composite into the raw `SUBTITLE_A` column, corrupting it
for the next real sync. Unlike `E2EField.bas` (a diagnostic tool), this
one sits behind a real ribbon button in the main sync chain
(`RibbonUI.OfferHarvestForSelectedSlides`). Added the same explicit
refusal, fail-first proven live (reverted, confirmed the composite string
would have been written and the cell was no longer empty, restored,
confirmed green), full suite 285/285.

**Both found by the same kennel pass, both examples of the same shape**:
a fix applied to one call site of a cross-cutting rule (SUBTITLE_A's
composite nature; a field set needing to exist on every real slide) does
not by itself confirm every OTHER call site or instance got the same
treatment. Worth checking deliberately next time a cross-cutting fix
lands, rather than assuming coverage from where the fix happened to be
applied first.

## Added and BUILT 2026-08-21 — BR, PROGRESS_HEADER and KEY_EVENTS_HEADER
## built as real, data-backed fields on two different Kinds

Rohan, 2026-08-21: "progres sheader can just be 'Last reported quarter
Q4F26' ... Key events header is really any important status like project
closed" -- followed by a clarifying question on how PROGRESS_HEADER's
quarter should be computed, and his answer overriding both proposed
options: "separate field thats either manually adjusted in rare
nonexpected cases or for student projects takes a frozen quarter label
when they report every six months." That settles it as `Kind=Given`, not
a cross-quarter derivation -- the same shape as SECTOR/TRL/SUBTITLE_B
built earlier tonight, not a new mechanism.

**PROGRESS_HEADER (Given).** New Field Spec row + new register column
(`register-wide.xlsx` col 55). Populated `"Last reported quarter Q1F27"`
across all 43 Q1F27 rows -- verified by reading the saved file directly,
not the writing session's own report. Rohan adjusts the rare exceptions
(irregular/six-monthly/student-project reporters) by hand; nothing here
infers or overwrites that judgement.

**KEY_EVENTS_HEADER (Derived).** `SyncOperations.bas`:
`KEY_EVENTS_HEADER_TAG` constant, `DeriveKeyEventsHeader` (blank when
`PROJECT_STATUS = "In Progress"`, otherwise passes the value straight
through -- a plain lookup, not a priority ladder like
`DeriveStatusBadge`, and unlike it there is no Controlled-vocabulary
refusal because there is no invented word to guess), added to
`DerivedFieldTags()` and `ComputeDerivedValue`'s `Case`. No register
column, computed at sync time, same as `TIMELINE_ELAPSED`/
`STATUS_BADGE`. New Field Spec row (`Kind=Derived`), mirroring
`STATUS_BADGE`'s row shape.

**First pass wrote the PROGRESS_HEADER Field Spec row wrong** -- populated
`Voice`/`Length`/`Own-job test`/`Do NOT`/`GLOBAL RULES`, the AI-drafting-
prompt columns SECTOR/TRL (the other `Kind=Given` rows) correctly leave
blank, because those columns only mean something for a field an AI
drafts. Caught by comparing against SECTOR/TRL's actual row content
before moving on, not assumed correct from having followed the STATUS_
BADGE row as a template for the wrong Kind. Fixed in place.

**Fail-first proven**: temporarily removed the `"In Progress"` exclusion
from `DeriveKeyEventsHeader`, confirmed both new/extended tests failed
naming the exact defect (`"the default status headlines nothing, got 'In
Progress'"`), restored, confirmed green. Extended the existing
`BothDerivedFieldsReachableThroughPlanRoutineSync` test with a third
shape rather than writing a separate reachability test -- proves the
shared derived-field loop's `"" = don't write` rule correctly leaves a
found-but-blank `KEY_EVENTS_HEADER` shape untouched, not just that the
pure function works in isolation. Full suite: 286 passed, 0 failed, 0
skipped (285 -> 286: one new test, one extended).

**Still open, not done in this pass**: the Field Spec's own prompt text
for `PROGRESS_BODY`/`KEY_EVENTS_BODY` still instructs the AI to open with
a bold quarter/status line, now duplicated by these dedicated fields and
needs removing.

## Fixed 2026-08-21 — BS, PROGRESS_HEADER default was wrong-quarter AND
## wrong for every S-project, both caught live by Rohan

**Wrong quarter.** BR's PROGRESS_HEADER pass defaulted all 43 Q1F27 rows
to `"Last reported quarter Q1F27"`, reasoning from today's date under an
AU July-June fiscal year. Rohan had already written the correct value
directly in his own instruction -- `"progres sheader can just be 'Last
reported quarter Q4F26'"` -- and it got silently overridden by a computed
guess instead of used. Asked directly rather than guess a second time;
answer was Q4F26. Moved the field from the 43 Q1F27 rows to the 43 Q4F26
rows.

**Wrong for every S-project.** Applied `"Last reported quarter Q4F26"`
uniformly across all 43 Q4F26 rows including the 13 S-series (student)
projects -- directly contradicting what Rohan had ALREADY told me earlier
the same conversation (six-monthly reporters, "not necessarily this
quarter"). He caught it immediately: *"hang on the s projects didnt
report in q4 like I explained."* Cleared `PROGRESS_HEADER` on all 13
S-project rows rather than guess which quarter each actually last
reported -- no data on disk says, and inventing one would be exactly the
kind of unsourced fact `PROVENANCE.md`'s rule exists to prevent. Rohan
fills these in by hand, which was the design from the start for exactly
this case.

**Both wrong in the same pass, both against something Rohan had already
told me in the same conversation** -- a computed default overriding his
literal example, and a uniform default overriding his own stated
exception. Verified the fix from the saved file: 30/30 P/K rows correct,
13/13 S-project rows blank, 0 stray values on any other quarter.

## Propagated 2026-08-21 — BT, PROGRESS_HEADER/KEY_EVENTS_HEADER shapes
## cloned from templates onto all 43 real slides

Same clone-and-tag pattern used for `MILESTONE_TIMELINE`/`PROJECT_PHOTO`
earlier tonight: for each real slide, clone the matching P/K/S template's
`PROGRESS_HEADER`/`KEY_EVENTS_HEADER` shape, position it identically
(confirmed all 3 templates already share identical geometry), tag it,
then shrink `PROGRESS_BODY`/`KEY_EVENTS_BODY` by the same 12pt (bottom
edge held fixed) and clear their bold flag -- the same body-bold bug
fixed on templates only, per the earlier CHECKLIST entry, now fixed on
the real slides too.

**Instance-key naming isn't fully consistent**: 5 real slides
(`P008`, `S009`, `S021`, `S022`, `S023`) carry a bare key with no leading
number/underscore, unlike the `3_P001` pattern everywhere else --
`_P`/`_K`/`_S` substring matching missed all 5 on the first pass (silent
skip, not a crash). Fixed by matching the type letter against the end of
the string instead (`P\d+$` etc.), re-ran the (idempotent) script, all 43
picked up.

**Save failed the first run** (`Presentation.Save : This presentation is
read-only`) -- the script had opened with `ReadOnly=True`. Confirmed the
live file's timestamp was unchanged before touching it again (no partial
write landed), fixed the open call, re-ran clean.

**Verified from the saved file's own XML, not the writing session's COM
report**: parsed the pptx zip directly with Python -- 46/46 `PROGRESS_
HEADER` and 46/46 `KEY_EVENTS_HEADER` role tags present (43 real + 3
template), correct per-type colour confirmed on one P/K/S real slide
each (`003C23`/`F55A2D`/`C0A2F2`), 0 remaining bold runs in `PROGRESS_
BODY`/`KEY_EVENTS_BODY` across all 43 real slides. The bold-detection
code path was positively proven first (found `b="1"` correctly on the
deliberately-bold header shapes themselves) before trusting its "0
found" result on the body text -- a check that only ever reports absence
is worthless until something proves it can also report presence. First
verification attempt scanned 0 real slides (tag lookup used lowercase
`instance_key`/`role`; the actual XML stores tag names uppercase --
`INSTANCE_KEY`/`ROLE`) -- caught before it was reported as a clean
result, not after.

## Fixed 2026-08-21 — BU, PROGRESS_HEADER's S-project value was Q4F26 too

Rohan, right after BS landed: *"hang on the s projects didnt report in
q4 like I explained."* BS had correctly excluded S-projects from the
Q4F26 default but left them entirely blank rather than set to their
actual last-reported quarter -- Rohan's answer was `Q3F26`. Set on all
13 S-project rows. Verified from the saved file: 13/13 correct.

## Closed 2026-08-21 — BV, "SRC_MILESTONES -> MS1-7 migration" was a
## broken CARRY, not a fresh migration

Rohan: *"import all milestone data."* Before writing anything, checked
the Field Spec's own declared shape for `MS1-7`: `Kind=Prose/Given`,
`Cadence=Standing`, `History treatment=CARRY` for the five middle
labels -- "the milestone plan only changes if the project is varied, so
this is drafted ONCE per project and carried forward." That single
check reframed the whole task: 38 projects already had this data,
correctly compressed from `SRC_MILESTONES`'s raw tracker rows, sitting
on **Q3F26**. **Q4F26 had zero of it** -- 0 of 543 non-blank cells
carried across. This is the same CARRY-restoration gap flagged earlier
in the session as still open, not a fresh derivation task.

**Carried 543 cells for 38 projects** from Q3F26 -> Q4F26, verified by
comparing per-project non-blank-cell counts between the two quarters
(0 mismatches) and spot-checking two projects' exact values.

**First attempt crashed** (`Unable to cast object of type System.Double
to type System.String`) -- some `MS*_DATE` cells hold raw numeric month
offsets (e.g. `6`, `12`), not text, and assigning a `.NET Double` back
through `.Value2` without matching type failed via COM interop. Left an
orphaned EXCEL.EXE holding the unsaved partial run -- confirmed the file
on disk was untouched (same timestamp as the prior write) before killing
it and re-running with an explicit type branch (`is [double]` ->
`.Value2 = [double]`, else `[string]`).

**5 of 43 projects still had no MS data anywhere** after the carry.
Checked each against `SRC_MILESTONES` before doing anything further:
`P008` and (initially miscounted, corrected on a direct final check)
`S023` have no raw milestone rows AND no `START_DATE`/`END_DATE` --
nothing to import, left blank. `2_P009`/`1_P010`/`2_P012` have clean
4-5-row source data that maps directly onto MS2-6 with no selection
judgement needed. `S023` alone has 10 raw milestones with no due dates
at all -- exactly the case the Field Spec calls "a judgement, not a
read" (compressing 10 real milestones down to 5 circles), left
untouched rather than guessed.

**Checked whether `MS_DONE` could be mechanically mirrored from the
tracker's own "Deliverable Completion RM %" column before drafting the
3 clean projects -- it can't.** `3_P001`'s own existing (already-CARRY'd)
data has milestones the tracker marks 100% complete that are NOT marked
done on the slide -- proof the DONE flag reflects Rohan's own review,
not a literal read of that column. Asked rather than guess a fourth
field this session; his answer: draft labels/dates, leave every
`MS_DONE` blank on all 3 projects for him to set. Verified from the
saved file: all 3 correctly drafted (MS1 = "Project initiated" +
`START_DATE`, middle slots = ~4-word compressions of the real tracker
milestone names with their due dates, MS7 = "Project end" + `END_DATE`,
every DONE flag blank), `P008`/`S023` confirmed untouched.

**Still open**: `P008` and `S023` have no importable data at all --
`S023` additionally needs the genuine milestone-selection judgement call
before it can carry any MS data.

## Fixed 2026-08-21 — BW, PROGRESS_BODY/KEY_EVENTS_BODY prompt text
## still told the AI to draft the bold header line the new dedicated
## fields now own

Follow-up from BR/BT: once `PROGRESS_HEADER`/`KEY_EVENTS_HEADER` existed
as real fields, the Field Spec rows for the two body fields still carried
the old instruction to open with a bold quarter/status line -- drafting
against the unedited prompt would have produced the header twice, once
in the dedicated shape and once again inside the body text.

`PROGRESS_BODY`: removed "Opens with a header carrying the quarter
label..." from Purpose, "after the quarter-labelled header" from Length,
and "Quarter-labelled header line in bold, then bullets" from Behaviour.
`KEY_EVENTS_BODY`: removed "A status label line in bold, then bullets"
from Behaviour. Both Behaviour cells now name the field that owns the
header (`PROGRESS_HEADER`/`KEY_EVENTS_HEADER`) and say explicitly not to
repeat it, so the next person reading the spec sees why it isn't
mentioned rather than assuming an oversight. Verified from the saved
file.

## Fixed 2026-08-21 — BX, mother-hound kennel: BU's "13/13 correct" was
## wrong, the same bare-key bug repeating in a second script

`excel-hound` re-derived the actual state of `PROGRESS_HEADER` on Q4F26
directly from the saved register rather than trusting BU's own claim,
and found 17 S-series rows, not 13 -- `S009`, `S021`, `S022`, `S023`
(the identical bare-key instance IDs, no leading `N_` prefix, that
FIX-LIST item BT already found breaking a DIFFERENT script's
`_P`/`_K`/`_S` substring match) still read the P/K default `"Last
reported quarter Q4F26"`, wrong for six-monthly student reporters. BT's
fix (match the type letter against the end of the string) was applied
to the propagation script; it was never carried to whatever script
wrote BS/BU's Q3F26 default, so the same bug recurred a second time
under a different name.

Re-verified all 17 S-rows directly before touching anything (confirmed
the exact 4 wrong, 13 already right), fixed the 4, re-verified from the
saved file: 17/17 correct.

**The lesson, stated as a shape (same shape as BQ's "a fix applied to
one call site doesn't confirm every other call site")**: a defect class
found and fixed in one script is not fixed everywhere that class of bug
can occur. The bare-key naming inconsistency (`P008`/`S009`/`S021`/
`S022`/`S023`) has now caused silent skips in at least two independent
scripts this session -- worth treating as a standing hazard for any
future script that matches on instance-key substrings, not a
one-off.

## Fixed 2026-08-21 — BY, mother-hound kennel: DeckAdoption.bas had
## Harvest.bas/E2EField.bas's SUBTITLE_A blind spot, a THIRD live site

`hickey-hound` traced the actual call graph for the `SUBTITLE_A`
composite-field collision (a shape's rendered text is a middot-joined
join of four register columns, not its own raw value) and found a live
path neither of tonight's earlier fixes touched: `RibbonUI.bas:1171` ->
`AdoptFlow.AdoptExistingSlides` -> `DeckAdoption.PlanAdoption`
(`DeckAdoption.bas:193`) harvested every high-confidence matched role's
rendered text unconditionally, with no exemption, and wrote it straight
into a new register row via `ExcelOutput.UpsertRow` with no guard.

Concrete trigger: duplicate an existing composed slide to seed a new
project (a normal way to start one -- the same instinct that caused the
BQ template gap earlier tonight), run "Adopt Existing Slides" -- the
rendered composite (`"Calix ~ UniSA ~ Livestock ~ TRL 3-5"`) gets
written verbatim into the new row's `SUBTITLE_A` cell, and `VerifyLink`
round-trips the same corrupted string back onto the slide, so the write
reports success.

Fixed with the same shape as the other two refusals, adapted to this
call site's structure: `PlanAdoption`'s harvest loop now skips writing
`SUBTITLE_A`'s value into `harvested()` when its role matches
`SyncOperations.SUBTITLE_COMPOSITE_FIELD`, but still counts the role as
a confident structural match (`anyMatch = True`) -- the shape genuinely
IS in the right place, only the VALUE harvest is unsafe, so the field
is left unwritten rather than downgrading the whole slide to
`unclassified`. New test (`Test_DeckAdoption_
RefusesToHarvestSubtitleAAsAComposite`) proves both halves: `Title`
still harvests and lands on the register row normally, `SUBTITLE_A`
does not, and `CommitAdoption`/`VerifyLink` both report success because
they never see the composite. Fail-first proven live (reverted the
`SUBTITLE_A` branch, confirmed the exact two expected failures --
composite harvested, composite written to the register -- restored,
confirmed green).

**Third live call site, same defect class, in one session.** Same
lesson as BQ/BX above, one level up: `SyncOperations.
SUBTITLE_COMPOSITE_FIELD` is the shared constant so the check is
centralized on what "is this field a composite" means, but nothing
forces every NEW code path that reads a shape's text to actually
consult it -- a fourth site would need the same manual find-and-fix,
not a mechanism that catches it automatically. Worth a real look if a
fourth turns up.

## Corrected 2026-08-21 — BZ, the "dialog count" CHECKLIST item was
## stale, not open -- verified by reading the chain, not the doc

Rohan asked to prioritise CHECKLIST.md's "Dialog count across one full
cycle" item, flagged 2026-08-16 (night) as "PRIORITY for next session."
Before writing anything, read `RibbonUI.PutItOnTheSlidesCore` and
`SyncNowChainCore` directly rather than sizing the work from the doc's
own description (the exact mistake this project's own CLAUDE.md already
names: "NEVER SIZE OR SCOPE A CHANGE FROM A HANDOVER DOCUMENT. OPEN THE
MODULE"). Every candidate the CHECKLIST item names was already fixed,
dated in the code's own comments, all AFTER the CHECKLIST entry was
written and never reflected back onto it:

- The two-press build-then-apply requirement Rohan explicitly called out
  ("that should all be one approval step") -- collapsed 2026-08-18. One
  press now builds the queue, asks once with the real pre-ticked list in
  hand, and applies.
- The "unsaved workbook" guard -- made silent 2026-08-19
  (`WorkbookBridge.EnsureSavedQuietly`). Only interrupts on a genuine
  save failure now, never on the routine "had changes, saved them" case.
- `StartQuarter`/`RollForwardUI`/`RefreshDraftingSheets`/the field-
  coverage notice -- folded into ONE combined report dialog
  (`DraftingUI.BeginCollecting`/`AppendCollected`/`EndCollecting`),
  shown only when there is something to report.
- Three separate redundant modals deleted outright: the "Go ahead?"
  confirmation (2026-08-14), the `WhereAmI` redundant status re-check
  (2026-08-17), the original three-way setup gate.

A steady-state cycle today is roughly 2-4 dialogs, not the 10-12 Rohan
counted on 2026-08-16. Nothing was built for this item because there
was nothing left to build -- CHECKLIST.md corrected in place rather
than "fixed" a second time.

## Built 2026-08-21 — CA, the file-per-quarter PRUNE half, gated hard on
## a verified archive per Rohan's explicit design calls

`SCENARIOS.md`: "file-per-quarter deletes work -- it is not hygiene."
The archive half (`WorkbookBridge.ArchiveWorkbookForPeriod`) was already
live, tested, and non-destructive (creates a file, never touches the
live register); the prune half was explicitly not built, and the code's
own comment at the call site said so: *"WHEN THE PRUNE IS BUILT THIS
MUST BECOME A HARD GATE... pruning the old period's rows without a
verified archive is the destructive step the archive exists to make
safe."*

Before writing anything destructive, asked Rohan two scoping questions
directly rather than guess a default for a mechanism that deletes real
rows from the live register:
1. **Timing** -- automatic, immediately after a successful archive
   (his answer, the recommended option), not a separate button.
2. **Scope** -- only the exact period this roll-forward's own
   `fromPeriod` names, never every period older than the newest (his
   answer). The live register held 3 periods at once the night this was
   built (`Q3F26`/`Q4F26`/`Q1F27`) -- a "sweep everything older" design
   would have silently touched `Q3F26` the first time this ran, from a
   roll-forward that had nothing to do with `Q3F26`.

Rohan also asked directly whether pruning breaks a new quarter's ability
to look back at the previous one. It doesn't, and the answer is in the
existing code, not something this fix changes: `ExcelOutput.
RollForwardPeriod` already copies EVERY column, every row, from the old
period into the new one's rows, live, in the SAME roll-forward action,
BEFORE archive or prune run. Prune only ever removes a period's rows
AFTER they have already been carried into the new period and already
sit safely in the archive file -- nothing needed for "look back" is at
risk at the point it is actually needed. (This also explains why
`Q4F26` was missing all its `MS1-7` data earlier tonight, per FIX-LIST
item BV: `RollForwardPeriod` copies unconditionally, so `Q4F26`'s rows
evidently were not created by an actual Roll Forward press -- not a
defect in roll-forward itself.)

**Built:**
- `ExcelOutput.PrunePeriod(ws, period)` -- deletes every row whose
  `Quarter` cell matches `period` exactly, collecting row numbers first
  and deleting bottom-to-top (same shape `RollForwardPeriod` already
  uses to collect source rows before writing) so a shifting row index
  can never skip or double-delete a row mid-loop. Refuses a blank
  period rather than treating it as a wildcard that could match blank-
  Quarter rows. Fail-first proven live: broke the scope filter so it
  deleted every row regardless of period, confirmed the exact two
  expected failures ("Q3F26 -- OLDER... UNTOUCHED, got 0", "Q1F27...
  untouched, got 0"), restored, confirmed green.
- `DraftingUI.RollForwardUI` -- the archive result is now a hard gate
  (`If archiveProblem <> "" Then Say ... : Exit Sub`), exactly as the
  comment it replaces said it must become. Prune only runs after
  confirming `RollForwardPeriod`'s own return string reports a real
  copy (not `"REFUSED:"` or `"Nothing to do:"`) -- pruning the source
  period when nothing was actually proven to land in the new period
  would delete the only live copy of that data.
- New test `Test_ExcelOutput_PrunePeriodRemovesOnlyTheNamedPeriod`,
  built against a 3-period register (matching the real register's own
  shape) to actually exercise the scope decision, not just claim it.
  Full suite: 288 passed, 0 failed.

**Deliberately NOT done this session, and worth reading before this is
treated as finished:**
- **`Sync Log`'s per-period sweep is not built.** `SCENARIOS.md` calls
  for it explicitly, but `Sync Log`'s own columns (`When`/`Run`/
  `EntityCode`/`FieldID`/`Outcome`/`Change ID`) have no `Quarter`
  column to key a prune on -- there is no reliable way to say which
  logged change belongs to which period without a real design decision
  (a date-range heuristic against period boundaries was considered and
  rejected here as too fragile to build silently). `Sync Log` keeps
  growing, unpruned, until this is designed properly.
- **`ParkSheetCopy` is NOT retired.** `DOCUMENT-MAP.md` decision 6 and
  this checklist both say explicitly not to remove it early -- it stays
  load-bearing until the prune has actually run live and been trusted,
  which has not happened yet.
- **No real keyboard run against the live deck.** The checklist's own
  requirement ("Tests + one real keyboard run before the prune touches
  anything live") is only half satisfied -- the destructive primitive is
  unit-tested in isolation, but nobody has pressed the real "Roll
  Forward" button against a real or copied deck since this landed.
  **This is a real behaviour change to a button already wired to the
  live production register**: the next genuine Roll Forward press against
  `register-wide.xlsx` will now also prune the period it rolled out of.
  Worth running once, deliberately, on a copy before it happens for real
  -- not run against the live register in this session on purpose, since
  that is exactly the kind of irreversible action this file's own
  standing rules say to confirm first.

## INCIDENT, 2026-08-21 ~07:15 — a real sync blanked 117 fields of real
## content across 41+ live slides. Restored, root-caused, fixed, verified
## deployed. Full account below -- this is the most serious incident this
## project has had.

**What happened.** The first-ever real "Put it on the slides" sync of the
night wrote 357 items, reported "0 not approved, 0 dropped as stale, 1
failed" -- an apparently clean run. It wasn't: a field-by-field diff of a
verified pre-sync backup against the post-sync file showed **117 of 122
fields that held real, human-authored content were blanked** --
`STRATEGIC_ALIGNMENT_BODY` (41), `PROBLEM_BODY` (41), `PROJECT_PROGRESS`
(35). Full paragraphs and real percentages, gone, on 41+ real project
slides. Caught because Rohan looked at the actual deck and said "look at
the slides, they are trashed" -- not because anything in the tool flagged
it.

**Root cause, per an independent bloodhound (Fable model) investigation,
NOT what I first assumed.** My own first theory -- "Harvest was never run
first, so the register was genuinely empty" -- was wrong, and matters
because the fix it points to (run Harvest before syncing) would not
actually have prevented this:

A DIFFERENT session, the day before (2026-08-20 morning), deliberately
blanked 161 stale `Q4F26` register cells that were unedited carry-copies
of `Q3F26` -- the right *intent* ("clean up so these read as not-yet-
drafted"), but whatever wrote the blanks left **zero-length-string
cells (`""`)**, not truly empty ones. VBA's `IsEmpty("")` is `False`, so
every downstream check read these cells as genuine, intentional content:
`ExcelOutput.ReadSheetForPeriod` includes them in a row's field
dictionary (its own `Not IsEmpty(...)` guard), and -- this is the part
that matters -- **`Harvest.HarvestSlide`'s own "skip if the register
already has a value" check (`rowValues.Exists`) would have skipped every
one of these 117 fields too.** Harvest was not skipped by circumstance;
it was structurally disabled by the exact same collision. Running it
first would have changed nothing.

The write path itself had no equivalent protection. `SyncOperations.
PlanRoutineSync`'s DERIVED-field loop already refuses to write `""`
(comment: `"" = don't write`); the ordinary Given/Prose field path --
what actually wrote these 117 fields -- had no such guard anywhere.
`InjectPrimitive.bas` simply assigned `shp.TextFrame.TextRange.Text =
sourceValue`, blank or not.

Bloodhound's verdict: **DESIGN, triggered by an EVENT.** The event (the
Aug 20 cleanup using the wrong emptiness) is bounded and dated. The
design gap that let one bad data state destroy 117 fields with a green
report has two parts: (1) the one-press "build, ask once, apply" flow's
single consent gate is a *count* ("358 change(s) queued... apply them
now?"), not a per-field before-and-after, and R13's own header already
names count-only confirmation as the exact thing R13 exists to prevent;
(2) two definitions of "empty" coexist in this codebase (structurally-
absent vs. zero-length-string) and nothing anywhere can *display* the
difference to a person, so Harvest's protection silently inverts on it.
Also confirmed: `MILESTONE_TIMELINE`'s 41 similarly-first-ever writes
were NOT destructive -- what was overwritten was cloned donor
boilerplate from the earlier template-propagation work, not per-project
content, verified by direct comparison against the pre-sync backup.

**Immediate response, in order:**
1. Both pre-sync backups (pptx + register) duplicated to a second,
   separate location (`deck-sync-backups/PRE-real-sync-20260821-071027/`)
   before touching anything further, verified byte-identical by hash.
2. Live PowerPoint/Excel closed cleanly (Rohan's own session, no forced
   quit), deck restored from the verified backup via file copy, hash-
   verified, and independently re-read (not just trusted) to confirm real
   content was actually back.
3. Register deliberately NOT reverted -- its husk cells were identically
   blank before and after the sync, so nothing was lost there, and
   reverting it would undo other legitimate writes from the same run.

**The actual fix -- two changes, both built, fail-first proven, and now
DEPLOYED to the live add-in (see below):**

- **`InjectPrimitive.bas`**: a blank `sourceValue` is refused, not
  written, when the shape's current text is real (not empty, not the
  untouched `<<FIELDNAME>>` placeholder) -- the ordinary field path's
  missing counterpart to the derived loop's `"" = don't write` rule.
  Refused writes surface in the review queue with an explanation
  ("REFUSED: the register holds nothing for X, but the slide currently
  shows real content...") rather than either silently writing or
  silently skipping. **First version broke a genuinely different,
  correct behaviour**: `InjectSlotsField` calls `InjectPrimitive` per
  slot, and a slot legitimately clearing when this quarter has fewer
  items than slots (`Test_InjectField_SlotsSplitsOneValueAcrossFixedShapes`'s
  own words: "the unused third slot is blanked, not left holding the
  stale line") is NOT the same shape as this incident -- caught by the
  full suite, fixed with an `Optional refuseBlankOverReal As Boolean =
  True` parameter that `InjectSlotsField` explicitly opts out of. Fail-
  first proven twice: once for the guard itself (reproduced the exact
  incident with the guard disabled), once implicitly by the slots test
  regression and its fix.
- **`RibbonUI.bas`**: `OfferHarvestForSelectedSlides` renamed
  `OfferHarvestAcrossDeck` and rebuilt to walk every real, linked,
  non-template slide in the presentation (new `RealLinkedSlides`
  helper, `Public` for testability), not `Application.ActiveWindow.
  Selection`. The old version's own comment said "in Normal view a
  slide is ALWAYS selected" and reasoned that as fine -- it was checking
  1 of 43 real slides on every press, silently, and the night this was
  found nobody had ever multi-selected the deck first. Rohan: "make it
  so I don't have to select slides for that... shouldn't it just check
  register and/or deck?" Still silent on a steady-state deck (unchanged
  dry-run-first design), so this adds no new invariant prompt. Fail-
  first proven (inverted the template-exclusion filter, confirmed the
  exact three expected failures, restored).

Full suite: 290 passed, 0 failed, both fail-first provals done live.

**Deployment -- and a THIRD finding, arguably as important as the
first two.** Checking whether the fix was live at all surfaced that
**no VBA source change from this entire session had ever reached the
actual add-in Rohan runs.** The "Apply Approved" dialog during the
incident itself showed `build: 2026-08-20 15:34` -- `addin155.ppam`,
confirmed unchanged in the AddIns folder since the afternoon before.
Every fix built tonight (this one, the `SUBTITLE_A` `DeckAdoption` fix,
`KEY_EVENTS_HEADER`, the milestone prune) existed only in source files
and in disposable test-harness presentations. This also fully explains
`<<KEY_EVENTS_HEADER>>` showing on every slide -- not because
`PROJECT_STATUS` was blank (an earlier, now-superseded theory in this
same session), but because the code that even knows `KEY_EVENTS_HEADER`
exists was never loaded.

Rebuilt via `vba/tests/build_ppam.ps1` (all 34 production modules,
stamped `2026-08-21 08:49`), Rohan did the one confirmed-non-automatable
step (File > Save As > PowerPoint Add-in), saved as `addin156.ppam`.
Registered and loaded via COM (`Application.AddIns`: `addin155.AutoLoad
= False`, `addin156` added/`Loaded = True`/`AutoLoad = True`), **verified
persisted across a real close-and-reopen**, not just the in-memory
property -- and finally verified by triggering the live build-stamp
dialog (`RibbonUI.ShowSyncResult`, the exact mechanism this codebase
already had for "is this fix in, answerable from the screen") and
having Rohan read it directly: **`addin156` confirmed live.**

**Still open, not done tonight:**
- The register still holds all ~158 husk cells (`STRATEGIC_ALIGNMENT_
  BODY`/`PROBLEM_BODY`/`PROJECT_PROGRESS`/dormant `PROJECT_STATUS`).
  They are now SAFE to sync against (the guard refuses rather than
  destroys), but the underlying content is still genuinely missing from
  the register -- a real Harvest pass (now deck-wide) or manual entry is
  still needed to actually recover the data into the register itself,
  not just protect the slides from losing it a second time.
- The exact mechanism that minted zero-length-strings instead of true
  Empty cells on 2026-08-20 was not identified (bloodhound: "the scratch
  script wasn't preserved"). Worth naming if it recurs -- a bulk `Range
  = array` assignment with `""` elements mints ZLS; a scalar `.Value =
  ""` does not, per the same investigation's own note.

## Fixed 2026-08-21 (evening) — CB, milestone label boxes were sized for one
## line on 31 real projects that genuinely need two

Rohan: *"fixing milestones important"* — `3_K016`'s `MS1_LABEL` ("Kickstart
initiated | Dec 2025") was overlapping `MS2`'s circle below it. First theory
(wrong, corrected before anything was touched): that the non-standard wording
was corrupted data needing reset to the generic "Project initiated." Checked
against the pre-add-in backup and scanned every real project first, per
Rohan: *"measuere on my original slides, I dont trust what the addin did."**
Confirmed **31 real projects** (all 15 K-type, all 16 S-type) carry
deliberate, type-specific two-line wording ("Kickstart initiated |
[date]"/"PhD Commencement | [date]") — legitimate content, not a defect.

**The real defect**: static XML across all 46 slides (43 real + P/K/S
templates) showed byte-identical `MILESTONE_TIMELINE` geometry everywhere —
slots 1/3/4/7 fixed at H=12.5pt (room for one line), slots 2/5/6 at
H=21.5-22.9pt (room for two). Rohan: *"you shouldn't have geometry and
position [differences] if they are copied elements?"* — correct: since every
real slide is a faithful `Shape.Copy`/`Shapes.Paste` clone of its P/K/S
template (see BT above), the height split is baked into the TEMPLATE, not
per-slide drift. `MilestoneDevice.bas`'s own design rule (Rohan, 2026-08-13:
*"you can see the extremely accurate positioning? we need to maintain
that"*) means `DrawMilestones` never touches label geometry — the fix could
only be a template-level edit, re-propagated the same way the device was
built.

**Fix**: grew slots 1/3/4/7's label boxes to match 2/5/6 (22.92pt), centred
on each label's own existing circle (matching the pattern slot 2 already
used), applied identically across all 46 slides in one pass. **First attempt
silently failed** — `Presentations.Open`'s 3rd parameter (`Untitled`) was
`$true`, so PowerPoint opened the content disconnected from the real path;
`Save()` reported success against a phantom in-memory document while the
real file's mtime never moved. Caught by re-reading the saved file's own
bytes rather than trusting the script, same discipline as the morning
incident — refixed with `Untitled=$false` and a hard guard (abort if the
reopened `.FullName` doesn't match the real path) so this can't repeat
silently. Verified from saved bytes: all 46 slides uniform,
`MS1/3/4/7_LABEL` now 22.92pt.

## Fixed 2026-08-21 (evening) — CC, `KEY_EVENTS_HEADER`'s legitimate blank
## was being treated the same as "cannot compute" and silently never written

`SyncOperations.PlanRoutineSync`'s derived-field loop gates every write on
`derivedVal <> ""`, on the assumption `""` only ever means "no source data."
True for `TIMELINE_ELAPSED`/`STATUS_BADGE`, false for `KEY_EVENTS_HEADER`:
`DeriveKeyEventsHeader` returns `""` as the CORRECT answer whenever
`PROJECT_STATUS = "In Progress"` — the header is supposed to render as
nothing. The existing regression test (`...BothDerivedFieldsReachable
ThroughPlanRoutineSync`) had actually enshrined the bug: its own comment
called the placeholder-shape-left-untouched outcome correct, using a fixture
seeded with arbitrary text ("should stay untouched") rather than the real
`<<KEY_EVENTS_HEADER>>` placeholder every real slide actually starts with —
a test that validated the wrong scenario convincingly.

Practical effect: every "In Progress" project's slide (the overwhelming
majority) showed the raw `<<KEY_EVENTS_HEADER>>` token forever, because the
write was never attempted, not because it failed.

**Fix**: the gate now also attempts the write whenever `PROJECT_STATUS` was
actually read for the row (`rowValues.Exists("PROJECT_STATUS")`), regardless
of whether the computed value is blank — and defers the actual safety
decision to `InjectPrimitive`'s own `refuseBlankOverReal` guard, which is
already placeholder-aware and was being bypassed entirely by this shortcut.
Rewrote the existing test to assert the corrected behaviour (placeholder ->
real blank) and added a new one proving real hand-typed content still
survives (the guard is reached, not merely present in source). Both fail-
first proven; full `InjectField`/`SyncOperations` regression cluster green.

**A second, genuine data gap surfaced verifying this, not a code bug**:
`PROJECT_STATUS` (a `Controlled`-kind field, fixed vocabulary "In
Progress"/"Not Started"/"Project Closed") was blank for Q4F26 on 41 of 43
projects — the fix correctly refused to guess rather than write garbage.
Milestone-done evidence (37/41 show `MS1_DONE=Y`, none show `MS7_DONE=Y` or
any closure signal) supported "In Progress" for all 41; confirmed with
Rohan rather than inferred silently (Field Spec's own text: *"do not infer
or fill the gap"*), then written in. Also added an Excel data-validation
dropdown on the `PROJECT_STATUS` column (the only `Controlled`-kind field in
the whole Field Spec) so future quarters get a constrained in-cell picker
instead of a blank cell nobody remembers to fill correctly.

## Fixed 2026-08-21 (evening) — CD, `PROJECT_PROGRESS` rendered as a raw
## decimal ("0.8027") instead of a percentage on 16 real slides

`PROJECT_PROGRESS` is `Kind=Given`, stored in the register as display text
("80%", not 0.8) per its own Field Spec contract — the opposite convention
from the BAR path, whose own contract is a raw 0-1 fraction
(`InjectProgressVia`'s error message: *"0.9, not 90 and not '90%'"*). Two
different injectors, two different, mutually-inverted conventions for the
same-looking value — worth remembering if this field's routing ever changes
again. On the 16 real slides where `PROJECT_PROGRESS` has no track/rest companion,
`InjectorFor` correctly falls through to the plain-text writer — which had
no idea this one field's register value needs percentage formatting, and
wrote whatever raw text sat in the cell verbatim. Same failure shape already
documented in this file for a trackless bar ("the injector was never the
thing that was wrong") — the value preparation was.

**Fix**: `InjectField`'s plain-text (`Case Else`) branch now runs the source
value through a new `FormatIfProgressText`, scoped exactly to
`PROJECT_PROGRESS` — numeric and in 0-1 range converts to a whole-number
percentage string; anything already formatted, non-numeric, or out of range
passes through unchanged (not this function's job to guess-correct a
data-entry error). New fail-first-proven test through `InjectField` (not
`InjectPrimitive` directly, since the fix lives one layer up in the
dispatch), plus a negative case proving the formatting is scoped to this
one field and doesn't touch anything else.

**Root cause of the 16 wrong values, checked directly rather than assumed**:
the register's `PROJECT_PROGRESS` was genuinely blank for Q4F26 on 41 of 43
projects (all of them, not just these 16) — the stale numbers were leftover
content from an earlier quarter or early testing, never refreshed, and the
same-night `refuseBlankOverReal` guard (see the morning INCIDENT below)
correctly refused to blank them rather than erase real-looking content.
Sourced real Q4F26 values from `SRC_MILESTONES` column M ("Project
Completion %") — a permanent, already-pasted-in sheet the sheet's own row 7
documents as feeding `PROJECT_PROGRESS` directly — for the 16 with the worst
(and, on inspection, duplicated/leaked) stale values; confirmed with Rohan
before writing, not inferred. The other 27 projects' register values were
already correct and untouched.


## INCIDENT, 2026-08-21 (evening) — a diagnostic sync tool silently wrote
## LAST QUARTER's prose onto the live deck, twice, both reporting "SAVED to
## disk (verified)"

The second serious incident this project has had, different shape from the
morning one (destructive blanking vs. silent regression to stale data), same
root lesson: a green report is not evidence.

**What happened.** After fixing CC/CD above, ran `sync_real_deck.ps1
-SaveWhenDone` twice against the real deck to apply them. Both runs reported
"43 corrected... SAVED to disk (verified)." Independent verification (a
byte-copy of the saved file, parsed fresh, never trusting the writer's own
in-process report) showed `PROGRESS_BODY` reading *"Last reported quarter
update – Q3F26"* on the live slides — not Q4F26. The deck had been silently
regressed by a full quarter.

**Root cause, found by reading the actual code, not guessed.**
`SyncRealDeck.bas` (the driver `sync_real_deck.ps1` uses) calls
`RunSync.RunRoutineSync`, which reads the register via
`ExcelOutput.ReadSheet(ws)` — `ReadSheetForPeriod(ws, "")`, no period filter.
`ReadSheetForPeriod`'s own code comment states the collision rule plainly:
*"Same project, same period, twice. First one wins."* With no filter, EVERY
row for a given instance ID passes the keep-check, so "first" isn't "same
period twice" as the comment's own example assumes — it's whichever period's
row sits highest in the sheet. Q3F26 rows sit above Q4F26 rows in every real
register past its first quarter, so this reads last quarter's data for
every field on every project, not just the two just-fixed ones — CC's own
"43 corrected" success for `KEY_EVENTS_HEADER` earlier the same night likely
also ran against Q3F26 `PROJECT_STATUS`, invisible only because that
specific value happened not to differ between quarters for most projects.

**The real "Sync Now" button was never affected.** `RibbonUI.bas` (what
Rohan actually presses) builds its sheet via `ExcelOutput.
ReadSheetForDeckPeriod(ws, period, problem)` — genuinely period-aware,
confirmed by reading the call site directly. `SyncRealDeck.bas`'s own header
comment claimed to be *"the same RunSync.RunRoutineSync"* mirroring
`SyncNow`'s behaviour; that claim was false the whole time this bug existed
— a real, undetected drift between a diagnostic tool and the button it
claimed to twin, same shape as `OnboardFlow.bas`/`AppEvents.cls` staleness
found fixing the tooling earlier the same session (see below).

**Immediate response.**
1. Corrupted state preserved as its own backup before touching anything
   (`PRE-RESTORE-period-bug-corrupted-*.bak`) — evidence kept, not
   discarded.
2. Restored the deck from the last backup confirmed (by direct content
   check, not assumed) to predate BOTH buggy runs and still carry the CB
   milestone-height fix — one clean restore point, nothing re-done.
3. Register was NOT reverted — it was never the corrupted side; the bug was
   in the sync tool's read logic, not the stored data. All of tonight's
   register writes (CC/CD's real values, `PROJECT_STATUS`) survived intact.

**The fix.** `SyncRealDeck.bas` and `HiddenFixCheck.bas` (the only two
callers of the unfiltered path) now build the sheet the same period-aware
way `RibbonUI.SyncNow` already did — `ExcelOutput.ReadSheetForDeckPeriod(ws,
DeckRegistry.GetDeckPeriod(pres), problem)` — and call
`RunSync.RunRoutineSyncWithSheet` directly. `RunRoutineSync` itself (the
unfiltered wrapper) is left in place, since deleting it risks breaking a
caller this search didn't find, but now carries a loud "DO NOT ADD A NEW
CALLER" comment naming the exact failure and pointing at the fix, so a third
caller can't reintroduce this silently.

**Also found and fixed reaching this point, same tool-drift shape as CB's
`Untitled` bug**: `preview_real_deck.ps1`/`sync_real_deck.ps1` both
hand-maintained a second, already-stale copy of the production module list
— missing `IdentityCheck`/`ReviewQueue`/`TemplateSlide`/`CommandBarUI`/
`Sources` (needed by `PreviewRoutineSync`'s R9 duplicate-key check),
referencing a deleted `OnboardFlow.bas`, and copying `AppEvents.cls`
byte-for-byte without the LF->CRLF conversion `run_vba_tests.ps1` already
documents as required for a class module's header to be recognised at all
(silent import as a plain Standard Module, then a generic "Expected: end of
statement" nowhere near the real cause). Both scripts now import the exact
canonical list `build_ppam.ps1` uses for the real shipped add-in.

**Verified, independently, from the saved file's own bytes, after the
re-run**: period text reads Q4F26 throughout; all 16 `PROJECT_PROGRESS`
values match the sourced numbers exactly (tag-scoped read, not a blind text
scan — an earlier "still broken" alarm during this same investigation
turned out to be a different, unrelated stray value on the same slides,
caught by a regex that matched anywhere in the slide rather than the
actually-tagged shape); `KEY_EVENTS_HEADER` placeholder count 0/43 (42
correctly blank, 1 real, matching each project's actual status); milestone
geometry from CB still intact, unaffected throughout.

**Lesson for next time, stated so it's checkable**: a "headless twin" of a
real button is a claim, not a guarantee — the two code paths can drift
apart silently, and the only way to know is to read both call chains, not
trust a header comment. Worth a standing check: before trusting any
`vba/tools/*.bas` script that claims to mirror a real ribbon action, grep
the actual button's handler and diff the call chain.


## Fixed 2026-08-21/22 — CE, `VerifyRealDeck.bas` had the same period bug
## just fixed, plus a full milestone-lifecycle test that was never written

`VerifyRealDeck.bas` — the tool built after the 2026-07-26 incident
specifically to be trusted when a sync looks suspicious — read the register
via the same unfiltered `ExcelOutput.ReadSheet` as `SyncRealDeck.bas`/
`HiddenFixCheck.bas` had (see CC's incident entry above). A mother-hound
audit found it: it would have compared a live Q4F26 slide against its
Q3F26 register row, either a false mismatch on any field that legitimately
changed between quarters, or worse, a false "slideOk" on any that happened
to coincide. **Fixed** the same way as the other two: builds the sheet via
`ExcelOutput.ReadSheetForDeckPeriod(ws, DeckRegistry.GetDeckPeriod(pres),
problem)` instead.

**A second, independent bug in the same tool, same audit**:
`verify_real_deck.ps1` had the exact stale-module-list problem already
found and fixed twice the same night in `preview_real_deck.ps1`/
`sync_real_deck.ps1` -- a hand-picked 3-module import (`Resolve`/
`ExcelOutput`/`VerifyRealDeck`) that never covered what the tool actually
calls (`DeckRegistry`, `InjectPrimitive`, `WorkbookBridge`, and their own
transitive dependencies). The tool had been unable to compile for three
weeks, and its `catch` block never set a failing exit code, so every run
exited 0 regardless. Both fixed: imports the same canonical module list
`build_ppam.ps1` uses, and a driver-level failure now actually `exit 1`s.
Proven by running it against the real live deck end to end for the first
time in three weeks -- its first real output surfaced a large, separate
body of findings (624 "no tagged shape found", 62 mismatches) still
untriaged, likely including a systematic false-positive: the milestone
device's 21 fields per project are deliberately addressed by shape NAME,
not a `role` tag (`MilestoneDevice.bas`'s own design), and this tool only
checks for `role` tags.

**Also built, same session**: a full milestone-lifecycle test. Rohan,
directly: *"have you run a test where on each slide you are able to
basically make the milestone circles go through every part of their life
cycle as if a real project was regularly ticking them off?"* Honest answer
was no -- every existing `MilestoneDevice` test called `DrawMilestones`
exactly once with one static done-pattern, never proving a real project
ticking off milestones one quarter at a time doesn't leave a stale circle
behind (the exact risk `DrawMilestones`'s own header comment names). New
test (`Test_MilestoneDevice_FullLifecycleNoStaleCircles`) calls
`DrawMilestones` six times on the same group (0 done through fully
complete) and checks every slot's three circles at every step, not just
the one that changed. Needed a second test fixture too
(`NewMilestoneDeviceWithNow`) -- the existing one never created `_NOW`
shapes at all, so no test had ever exercised the path every real P/K/S
template actually uses. Fail-first proven: deliberately skipped the
NOW-circle hide call, confirmed the test caught a stale NOW circle
surviving on every previously-current slot, restored, confirmed clean.

## Fixed 2026-08-22 — CF, `TIMELINE_ELAPSED` rendered raw register decimals
## at 18pt inside a ~5.5pt bar on 29 real slides, because the bar shape
## structurally couldn't be drawn as a bar at all there

Found via `ui-hound` and independently corroborated by `mother-hound`: real
slides showing text like `"0.8027"` at 18pt inside the `TIMELINE_ELAPSED`
bar (height ~5.5pt), and the K/S templates themselves still holding the
literal `<<TIMELINE_ELAPSED>>` placeholder token -- meaning every future
slide cloned from either template would inherit it fresh.

**Two layers, both real.** `InjectProgressField` only ever wrote
`.Left`/`.Width` to the bar shape -- it never touched the text frame, so
whatever text happened to be sitting there (a template placeholder, a
stray manual entry) stayed forever, growing more wrong every time the
bar's real width changed underneath it. Fixed: clears both the done and
rest shapes' text as part of every write, folded into `WouldChange` so a
bar already at the right width but still carrying stray text doesn't
report "no change" and skip the clear. Fail-first proven (skipped the
clear, confirmed the new test caught the exact live symptom, restored).

That fix alone didn't reach 29 real slides, because the deeper cause was
structural: `TIMELINE_ELAPSED` had **neither a `.track` nor a `.rest`
companion shape** there, so `InjectorFor` could not route it to the bar
injector at all ("EITHER COMPANION MAKES IT A BAR") -- it fell through to
the plain-text path, which just wrote the raw computed fraction as literal
text. Compare a clean P-type slide, which has a `.rest` companion and
routes correctly. **Fixed** by adding a `.rest` companion (matching the P
template's own convention: grey fill, bar+rest summing to one consistent
extent) to the K/S templates and all 27 real slides missing one, resetting
a few visibly-drifted bar positions in the process (one had wandered to
`Left=871pt`, evidence of an earlier partial/broken write). Verified
independently after a real sync, from the saved file's own bytes: zero
remaining stray text across all 43 real slides, every bar's computed
elapsed fraction in sensible 0-1 range (0.16-1.0).

## Fixed 2026-08-22 — CG, the real "Apply Approved" button dropped
## `STATUS_BADGE` and `KEY_EVENTS_HEADER` on every approval

Found via `mother-hound`: `ReviewQueue.ApplyApproved`'s field-resolution
chain only rescues `TIMELINE_ELAPSED_TAG` as a derived value -- ordinary
register columns and device role tags are the other two branches, and
`STATUS_BADGE`/`KEY_EVENTS_HEADER` fell into none of them, hitting
`haveProposed = False` and getting reported as `"DROPPED ... -- the
register no longer has a value for this field"`. Both fields only ever
reached the real deck via a diagnostic script, never the button Rohan
actually presses (`RibbonUI.bas:1907` -> `ReviewQueue.ApplyApproved`).

**Fixed** by adding an `ElseIf` branch resolving both tags through
`SyncOperations.ComputeDerivedValue` -- the same single function
`PlanRoutineSync` already uses at build time, now wired to apply time too.
Re-deriving is safe here (unlike `TIMELINE_ELAPSED`, which re-reads its
own build-time `ProposedValue` instead, because re-deriving a fraction a
second time would drift from what was actually shown for approval) since
`STATUS_BADGE`/`KEY_EVENTS_HEADER` are pure functions of `rowValues`, not
of anything computed at build time that could have moved on.

Fail-first proven: temporarily disabled the new branch, confirmed the
test caught the exact "dropped as stale" failure matching mother-hound's
real finding, restored, confirmed clean (4/4 `ApplyApproved` tests, 12/12
`SyncOperations` regression tests).

## Fixed 2026-08-22 — CH, `WorkbookBridge.LifespanOf` matched a sheet-name
## format retired since an earlier rename

`LifespanOf` checked for `"Sync Review"` (`Left(sheetName, 11)`) but
`ReviewQueue.ReviewSheetNameFor` has produced `"Review <type>-<tag>"`
sheet names since the rename documented in that function's own header
(the old format truncated mid-word and read like a temp file). Every real
review sheet has shown "unknown" in the workbook index instead of its
actual lifespan ever since.

Impact is cosmetic only: `LifespanOf`'s sole caller is
`WriteWorkbookIndex`, a human-facing report sheet -- nothing acts on the
misclassification. The companion test had been masking the bug the whole
time by hardcoding the same retired `"Sync Review q"` string as its own
fixture, rather than calling the real generator.

**Fixed**: `Left(sheetName, 7) = "Review "`, plus the test corrected to
call `ReviewQueue.ReviewSheetNameFor("q")` for a real sheet name. Fail-
first proven: reverted to the old check, confirmed the test failed with
the exact symptom (`"a review grid ('Review q-B616') is marked consumed,
got 'unknown'"`), restored, confirmed clean (13/13 `WorkbookBridge` tests,
52-module whole-project compile clean).

## Fixed 2026-08-22 — CI, the "Apply them now?" approval dialog could
## silently truncate away its own question

`PutItOnTheSlidesCore`'s confirmation dialog -- the one authorising real
writes to every approved slide -- built its `MsgBox` text as bare
concatenation of `fullReport` and the actual Yes/No question, with no
cap. VBA's `MsgBox` truncates silently, so a large enough `fullReport`
(grows with the number of registered slide types and slides) could cut
the question off the bottom entirely. Same class of defect already fixed
once in this file at `OfferHarvestAcrossDeck`.

**Fixed** by wrapping in `CapReport(fullReport, askApply)`, protecting
the question as `mustKeep` so it survives truncation no matter how long
the report body grows. Also fixed a second, related gap this exposed:
`CapReport`'s default notice claims "the full list is on the Run Log
sheet", which was untrue here -- `BuildAllQueuesCore` never wrote
`fullReport` there. Fixed at the source (`BuildAllQueuesCore` itself),
which closes the same latent lie in `ReviewChangesCore`'s `ShowSyncResult`
call too.

No dedicated unit test -- `RibbonUI`'s Core functions are
integration-only (live `MsgBox`/COM), same as the `OfferHarvestAcrossDeck`
fix this mirrors, which also has none. Verified via clean whole-project
compile and no regression in the existing `RibbonUI`/`CapReport` test
cluster (7/7 passed).

## Fixed 2026-08-22 — CJ, `VerifyRealDeck` reported 624 false positives for
## fields that never carry their own role tag by design

Traced the tool's first-ever run (2026-08-21: 624 "no tagged shape"
findings) field-by-field. Every single one fell into one of two known-by-
design categories: `MS1-7_LABEL`/`_DATE`/`_DONE` (581 -- `MilestoneDevice`'s
own columns, addressed by shape name, never role-tagged; recognised by
the existing `MilestoneDevice.IsColumnForThisDevice`, which already fixed
the identical false-alarm class for `FieldWiring.ScanFieldWiring` once
before), and `PROJECT_STATUS` (43 -- a register-only source for the
Derived `STATUS_BADGE` field, per `FieldSpec.bas`'s own row for it).
**100% false positive, 0% real defects.**

**Fixed** by adding `IsExpectedToCarryNoOwnRoleTag`, reusing
`MilestoneDevice`'s own function plus an explicitly commented
`PROJECT_STATUS`/`SCHEDULE_STATUS` exclusion. Skipped fields get their own
counter ("Fields skipped, no role tag expected by design") rather than
silently vanishing from the report.

Verified live against the real deck (read-only run, no backup needed --
the tool never writes to either file): missing-shape count went from 624
to exactly 0; slides reading fully OK went from 0 to 13.

**A separate, real finding surfaced by the same live run, not fixed
here:** the "fields with a text/value mismatch" count held steady at 62,
identical before and after this fix, and identical in shape to the stale
2026-08-21 report despite being fresh 2026-08-22 data (post-dating
today's `TIMELINE_ELAPSED`/`ApplyApproved` fixes) -- so this is real,
current, unrelated drift, not stale noise. Dominated by two fields
appearing together on nearly every checked slide: `KEY_EVENTS_BODY` (30)
and `PROGRESS_BODY` (27), with `ABOUT_BODY` (3), `STRATEGIC_ALIGNMENT_BODY`
(1) and `PROJECT_PROGRESS` (1) scattered. Not triaged further -- could be
ordinary pending-sync state (drafted/approved in the register, not yet
applied to these slides this quarter) or a genuine defect; the systematic
two-field pattern across nearly every slide is the kind of shape that
turned out to be structural, not incidental, the last three times this
project saw it. See CHECKLIST.md.

## Added 2026-08-22 — CK, STILL OPEN, the milestone timeline doesn't
## reflect project closure -- 8 of 8 closed projects affected

Slide-by-slide gap analysis of the real deck against the last hand-built
version (`2026-08-13-0956-pre-onboard`, 43 matched slides) surfaced this
via Rohan's own live observation ("whats happening with circle colour?")
on `3_P001`, slide 1: the milestone timeline's current-milestone circle
sits at "12 months / Method exploration" -- teal, visually distinct from
the rest -- while the slide's own badge says `PROJECT_STATUS = Project
Closed`, 80% progress, and the narrative says "Final Report submitted".

**Root cause is register data, not code.** `MilestoneDevice.DrawMilestones`
computes `lastAchieved` from the `MS*_DONE` flags and is working exactly
as designed -- it drew slot 3 as current because slot 3 is the last one
flagged `Y`. `3_P001`'s Q4F26 row has `MS1-3_DONE = Y`, `MS4-7_DONE =
blank`, despite the project being closed.

**Checked whether this is an isolated data-entry gap: it is not.** Scanned
every Q4F26 register row with `PROJECT_STATUS = Project Closed` (8 rows
total) for complete `MS*_DONE` coverage against however many milestone
slots that row's labels populate. **Zero of eight have complete flags** --
`3_P001` (3/7), `2_P009` (0/7), `1_P010` (0/5), `1_K1004` (1/6), `1_K1008`
(1/6), `3_K016` (4/6), `1_K022` (3/5), plus one closed project with no
milestone labels at all. There is no "correct" convention anywhere in the
live register to point at as the expected pattern -- closing a project has
never once been paired with backfilling its milestones.

**Independently re-derived by Fable (fresh agent, own read of the same
files): CONFIRMED, with one correction.** Same 8 rows found, same
DONE-flag gaps -- but `P008` has ZERO populated milestone labels, so
"every populated slot marked Y" is vacuously true for it: nothing to
repair there, it needs milestone content authored, not a DONE-flag fix.
**7 real candidates, not 8.**

**NOT A BLANKET FIX -- confirmed independently, this matters.** `3_P001`'s
own `KEY_EVENTS_BODY` says outright: "The industry partner's withdrawal
halted further development. Later-stage milestones not completed." MS4-7
blank is the TRUE state -- this project closed early, genuinely
incomplete, and marking it fully done would fabricate history. Fable's
own slide-side check confirms the rendered timeline (`MS_BAR` stopping at
MS3, `MS3_NOW` oversized) faithfully reflects this. Per-project narrative
review (this session) sorted the remaining 6:
  - **FIXED 2026-08-22**: `1_P010` -- `PROGRESS_BODY`/`KEY_EVENTS_BODY`
    describe full completion of all 5 milestones' activities; was 0/5
    marked. Backed up `register-wide.xlsx` first (`backups/
    PRE-1P010-MILESTONE-FIX-20260822-071...`), written via the LIVE Excel
    COM session already open on the real file (not a second handle --
    this project has a real prior data-loss bug from exactly that
    mistake), saved, and independently verified from the saved file's own
    raw XML bytes (not the writer's in-process report): all 5
    `MS1-5_DONE` read back `Y`. **Register only** -- the deck slide (10)
    still needs a real sync run to pick this up; not done automatically.
**BETTER EVIDENCE FOUND: `SRC_MILESTONES`.** Rohan's own question ("can't
you check the milestone info against the source evidence in the xl
file?") pointed at a sheet not yet consulted -- one row per milestone,
pasted from the real CRC SAAFE Milestone & Deliverable Tracker (source
S04), carrying `Deliverable Completion RM %`, `Project Status (S10)` and
the raw Knack reporting comments per milestone. This is materially better
evidence than inferring from drafted narrative prose, and re-triaged all
5 remaining rows against it:

  - **FIXED 2026-08-22**: `3_K016` -- source's M05 ("Publish findings") is
    explicitly `Complete` ("Paper submitted... final stages of
    preparation"), matching register `MS5` (was blank). `MS6` ("Kickstart
    end") has no source row of its own -- same administrative pattern as
    `MS1` -- and every substantive milestone is now done, so marked `Y`
    too. Backed up, written via the live Excel session (single-project
    write; the combined two-project script was blocked by Claude Code's
    own auto-mode safety classifier and split per Rohan's instruction),
    verified from the saved file's raw bytes: all 6 slots now `Y`.
  - **FIXED 2026-08-22**: `2_P009` -- source M01/M02 map to register
    `MS2`/`MS3`, both `Complete`. `MS6` shows `Completion=0` but its own
    comment overrides the stale number: "has been complete successfully
    and a report has been generated." `MS4`/`MS5` map to genuinely
    incomplete source rows and were left blank. `MS1` was NOT touched --
    outside what was proposed and approved, flagged separately rather
    than silently added; every other project here has `MS1` pre-marked,
    so this is a real inconsistency worth a decision, just not made
    unilaterally. Verified from saved bytes: `MS2`/`MS3`/`MS6` = `Y`,
    `MS1`/`MS4`/`MS5`/`MS7` unchanged.
  - **Confirmed correct as-is, no change**: `1_K022` -- source M04
    ("Finalise and distribute draft report") explicitly still pending
    sign-off ("final report with Dr Kelly Hill and Prof Erica Donner for
    final review"). Register's matching slot was already blank.
  - **Still genuinely ambiguous, left for Rohan**: `1_K1004`, `1_K1008` --
    the tracker's milestone numbering doesn't map cleanly onto the
    register's MS-slots for either (multi-item milestone descriptions
    bundling more than one register slot into a single tracker row).
    Cell locations handed to Rohan directly: `1_K1004` row 59 (`Q59`
    MS2/`T59` MS3/`W59` MS4/`Z59` MS5/`AC59` MS6), `1_K1008` row 63 (same
    column layout).

**Mostly fixed.** `1_P010`, `3_K016`, `2_P009` done and verified in the
register (all need a sync run to reach the deck). `1_K022` needed no
change. `1_K1004`/`1_K1008` await Rohan's own read -- cells identified,
evidence surfaced, decision genuinely his to make. The drawing code
itself was never wrong, confirmed three times now (twice independently).

## Added 2026-08-22 — CL, STILL OPEN, content depth drops sharply for
## S007 onward (13 slides) versus the original hand-built deck

Same gap analysis. P-series, K-series, and `S001`-`S004` all run 80-100%
of the original deck's visible text length per slide (normal quarter-to-
quarter variation). From `S007` onward (slides 31-43 of 43, matched 1:1
by content-similarity and confirmed order-preserving) it drops to a
consistent 55-70%, visibility-aware (hidden shapes excluded from both
sides so the comparison isn't skewed by leftover template cruft -- see
CM below).

Spot-checked `S012` (slide 34): the current `ABOUT` text (272 chars) is
factually accurate but drops the original's explicit deliverables framing
entirely -- "a database of ARGs across waste types, a protocol for
tracking ARG fate during treatment, and a modelling framework to evaluate
AMR risk..." is simply gone, compressed into one general sentence. The
`PROBLEM` box on the same slide is untouched (430 chars, byte-identical),
so this isn't a blanket per-slide issue -- it's specific to how much depth
`ABOUT`-class content gets during drafting for this batch of projects.

**Not triaged further.** Reads like these 13 projects got less drafting/
source depth than the rest of the deck at some point, not like anything
in the sync mechanism is broken -- worth checking `Sources`/extraction
completeness for `S007`-`S023` specifically before assuming a code fix is
even the right lever here.

**Independently re-derived by Fable: CONFIRMED, with a correction.** The
depleted band actually runs 51-70% (a little lower than first measured),
and it's not purely an S007+ pattern -- `4_K021` (position 24, ratio 0.70)
belongs in the same band and was missed by the original sweep. Driver
independently confirmed on `1_S012`: baseline's "Key Events -- Next 12
Months" and "Deliverables for Partner Review" blocks are simply absent
from the current slide, not just shortened.

## Added 2026-08-22 — CM, STILL OPEN (cosmetic), hidden leftover milestone
## donor text on 32 of 43 slides

Same gap analysis, found while investigating what first looked like a
severe cross-project data-contamination bug on slide 43 (S023) -- a raw
text-extraction diff (not visibility-aware) showed `MS1_LABEL`/`MS5_LABEL`
etc. carrying content describing a completely different (antimicrobial)
project. Rohan's instinct to check backups before trusting this ("You
have backups from before we started check your notes") caught it: those
`MS*` shapes all carry `hidden="1"` in the real slide's XML -- confirmed
by checking every shape's `hidden` attribute directly, not just its text.
`DrawMilestones` is working exactly as its own header says ("A SLOT WITH
NO MILESTONE IS HIDDEN, NOT EMPTIED") -- nothing is visible, nothing is
wrong on screen.

Scanned all 43 slides for hidden shapes still carrying text: 32 of 43
have at least one, almost all the harmless `MS7_LABEL="Project end"` /
`MS7_DATE="★"` pair left over on projects using fewer than 7 milestone
slots. Two slides (8 and 43, both projects with zero real milestone data)
carry a full donor set of unrelated antimicrobial-project text across all
14 hidden shapes -- confirmed via a pre-Aug-19 backup that the `MS*`
shapes didn't exist on slide 43 at all before a later milestone-device
retrofit, so the donor content was baked in when that retrofit cloned
from a real, populated slide rather than a blank template.

**Not fixed, low priority.** Purely a data-hygiene item -- real project
text sitting in shapes nobody sees. Worth a cleanup pass (blank the
donor text at template-build time) but not urgent since nothing currently
exposes it.

**Independently re-derived by Fable: CONFIRMED, count corrected to 34.**
Pattern and root cause both check out. The count differs (34 vs 32) --
likely a hidden-shape-detection methodology difference (direct `hidden`
check vs. also treating children of a hidden ancestor as hidden), not new
drift; not chased further given this finding's own low priority.

## Added 2026-08-22 — CN, STILL OPEN, `MS*_DONE` drifts from the real CRC
## tracker with nothing to catch it -- diagnostic tool built, not yet run

Found live: Rohan spotted `2_P012`'s entire milestone timeline rendering as
"not achieved" (every circle the same colour, no achieved/current
distinction) despite the project being well underway, 80% progress, "In
Progress". Register showed `MS1-7_DONE` all blank. `SRC_MILESTONES` (the
permanent pasted extract from the real CRC tracker) showed 4 of 5 tracked
milestones at `Completion=1`. Fixed `2_P012` by hand from that evidence
(`MS1-5=Y`).

**A naive follow-up scan (comparing raw DONE-flag counts against raw
source-completion counts) found 8 more "gaps" -- wrong.** Root-caused by
Fable (fresh agent, independent read of the same files):

  - **`MS*_DONE` has NEVER been code-linked to `SRC_MILESTONES`.** Zero VBA
    references to that sheet anywhere in `vba/`. Its own header claim
    ("Feeds MS1..MS7") is hand-typed process documentation, never
    implemented -- aspirational, not a broken mechanism.
  - **This is a deliberate, already-recorded decision, not an oversight.**
    FIX-LIST item BV (2026-08-21) proved mechanical mirroring from
    `Completion%` is wrong in both directions: `3_P001` correctly shows
    later milestones undone (closed early, genuinely incomplete);
    `2_P009`'s `MS6` showed `Completion=0` while its own comment said the
    work was finished. Rohan's own call at the time: leave it manual.
  - **The naive scan overcounts because `SRC_MILESTONES` is far more
    granular than the register.** One project can have 20+ tracker
    milestone rows collapsing into the register's 7 `MS` slots. Verified
    directly on `2_P004`: all 13 "missing" source completions, once
    grouped by due-month against the register's own slot dates, land
    under slots already marked `Y`. **`2_P004` needs no fix at all** --
    the earlier flagged-list entry for it was a false positive from
    counting raw rows instead of grouping them.

**Built (not yet run against the real files): `vba/tools/
MilestoneEvidenceReport.bas` + `milestone_evidence_report.ps1`, matching
`VerifyRealDeck.bas`'s own pattern exactly** -- opens both real files
READ-ONLY, never writes, not in the shipped add-in. Groups each
`SRC_MILESTONES` row into the first register `MS` slot whose own
`MS<n>_DATE` offset is >= that row's due-month offset (a deliverable due
at month 9, with checkpoints at 6 and 12, belongs to the 12-month
checkpoint). Reports disagreements only, quoting the actual Knack
comment text -- since a comment overrides a stale percentage in either
direction, per BV -- and **never decides or writes anything**; Rohan
applies by hand, same review-before-write norm `ReviewQueue.bas` already
holds for every other field. The grouping logic itself
(`GroupSourceIntoSlots`) is a pure function taking plain arrays, no file
I/O, matching `MilestoneDevice.DrawMilestones`'s own separation of pure
computation from the COM-driving wrapper around it -- but per this repo's
own established convention (confirmed in `run_vba_tests.ps1`'s comment:
tool modules are "compile-only... no test exercises them", same as
`VerifyRealDeck`), it is NOT wired into `TestRunner.bas`; verified by
compile + a live read-only run instead.

**Now run and working, after finding and fixing three real bugs in the
tool itself along the way (2026-08-22, later the same night):**

1. **Perf**: the first version re-scanned all ~560 `SRC_MILESTONES` rows
   via per-cell COM calls for each of ~40 projects (~22,000 round trips)
   -- never finished, this project's own established "bulk read/write"
   lesson paid for again. Fixed: one `Range.Value2` call reads the whole
   block, everything after works in memory.
2. **Appeared hung, wasn't**: the fixed-perf version still looked stuck
   (flat CPU for minutes) -- it was actually sitting at a modal VBA error
   dialog the whole time ("Run-time error '9': Subscript out of range"),
   only found because a screenshot was checked. Added per-project error
   trapping so one project's data can never crash the whole batch or
   leave a run silently stuck again.
3. **The actual bug**: `GroupSourceIntoSlots` had `If bestSlot = 0 Or
   slotOffset(slotNum) < slotOffset(bestSlot) Then` -- **VBA's `Or` does
   not short-circuit.** Both operands always evaluate, so
   `slotOffset(bestSlot)` = `slotOffset(0)` was read even when `bestSlot`
   was still its initial 0, against an array dimensioned `1 To 7`. This
   is why it crashed on literally every single project, every run.
   Fixed to nested `If`/`ElseIf`.

**A fourth bug, found by actually reading the report instead of trusting
"0 errors": the per-project detail was silently accumulating across
projects.** `projectDetail`, `slotGroupN`, `slotGroupComplete`, and
`slotLatestComment` are all declared inside the per-project loop, but a
loop-scoped `Dim` does not reset a variable's value between iterations in
VBA -- only an explicit assignment does. `projectDetail` was only ever
appended to, so `2_P003`'s printed block literally repeated all of
`3_P002`'s lines before adding its own; `slotGroupN`/`slotGroupComplete`
were only ever incremented, never zeroed, which is why the first run's
"N of M complete" figures visibly grew project over project (1 of 2, then
5 of 8, then 11 of 14...) instead of resetting per project. Fixed by
explicitly resetting all four at the top of each iteration. Rohan asked
whether this exact bug shape exists elsewhere in the codebase --
swept every fixed-size array declaration in `vba/` (`Drafting.bas`,
`SyncOperations.bas`, two `tools/` probes): all write every index
unconditionally before use, none carry this risk. Isolated to this tool.

**Verified live against the real deck/register, clean run, 0 errors,
detail no longer repeating: 41 projects checked, 14 with real
disagreements between grouped tracker evidence and recorded `DONE`
flags.** (The earlier "26" was the corrupted count -- discard it.)

**All 14 worked through with Rohan, evidence-based, register cross-
referenced against the tracker's own quarter-by-quarter comment history
(not just the raw Completion% -- pulled full untruncated text per row and
matched grouping against each register slot's actual `MS<n>_DATE` offset,
not label-text guessing):**

  - **Register overstated -- `Y` removed (7 projects, 8 slots), tracker
    evidence showed the grouped item(s) genuinely incomplete**: `3_P002`
    MS3 (a grouped item still "awaiting international isolates"),
    `2_P003` MS3 and `2_P004` MS3 (both an explicitly deferred economic
    assessment, "dependent on first identifying AMR management options"),
    `1_P007` MS3+MS4 (the project's own quarterly self-reports peaked at
    90% and 25% respectively, never reaching 100), `1_K1001` MS2 ("will
    be completed during the next milestone reporting period"), `1_K1002`
    MS2 (team/PAC setup not yet complete), `1_S012` MS2 (a grouped item
    explicitly "in the process of" and "expected" -- future tense).
  - **Register understated -- `Y` added (2 projects, 3 slots), tracker
    showed genuine completion the register never recorded**: `1_K1008`
    MS2+MS3 ("standardised protocol is included... chemistry lab has
    begun conjugating"), `4_K017` MS4 ("EDAR8 training has now been
    structured as a masterclass... published on the EDAR8 webpage").
  - **Left alone, deliberately -- 5 slots across 4 projects where the
    evidence gave no clean signal either way**: `4_K021` MS2 and
    `3_K023` MS4 (blank/"No data" tracker rows, absence isn't evidence of
    absence), `3_S002` MS2+MS3 (same), and notably `3_S003` MS3 and
    `2_S011` MS3, where the comment TEXT explicitly confirmed genuine
    completion ("Completed Successfully") despite a stale `0%` sitting
    next to it -- the register's existing `Y` was correct, the disagreement
    was a false positive from trusting the number over the comment. Left
    untouched rather than "corrected" into being wrong.

Backed up first (`backups/PRE-MILESTONE-EVIDENCE-FIXES-...`), written in
one pass to the register (no live session held open at the time, so no
handle-collision risk), verified independently from the saved file's raw
XML bytes -- all 11 writes confirmed landed exactly as intended. **Register
only, as always -- needs a real sync run to reach the deck.** Prevention
(Fable's proposal, not yet actioned): hook this same comparison into
`RollForwardUI` -- a button Rohan already presses every quarter -- rather
than a separate tool requiring new discipline to remember to run.

## Fixed 2026-08-22 — CO, `Put it on the slides` showed a needless OK-only
## dialog before its one real question

`RibbonUI.PutItOnTheSlides` collected `DraftingUI.PublishAllDraftedFields`'s
report and, if non-empty, showed it as a blocking `MsgBox` before falling
into `PutItOnTheSlidesCore`'s own "Apply them now?" Yes/No. In practice this
fired on every press with anything pinned: `PublishAllDraftedFields`'s own
final `Say()` always emits a summary, even when nothing changed, so
`published <> ""` was true almost always. The summary itself added nothing
-- it's already written to the Run Log unconditionally inside
`PublishAllDraftedFields` (the `WriteRunLog` call), so the dialog was pure
ceremony in front of the actual decision one click later. Same "one
question, not three" call already made for the Apply step, applied here.

Fixed: dropped the `MsgBox`, kept the `EndCollecting` call (still needed to
reset `mCollecting`/`mReport` so a chain that died halfway can't leak into
the next run). **Static check clean; live build+test verification pending
-- PowerPoint/Excel were both open under a live session at fix time, so the
usual house-pattern test run (`vba/tests/run_vba_tests.ps1`, which aborts if
either is already running) couldn't be exercised. Run it next session
before trusting this beyond the static read.**

## RETRACTED 2026-08-22 — CP, milestone circle colours were never a
## regression -- this was Claude's own earlier fix, misread as a bug

Originally logged as "S/K `_OFF` circles regressed to shared teal, losing
type-specific pale tints, confirmed via the pre-retrofit backup." That
citation didn't survive checking: every available backup was swept (18
timestamped `.bak.pptx` snapshots, 6 named backups, the 2026-08-13
pre-onboard original) and in every single one, the only slides carrying
the old `MS1_OFF` shape naming are P-type, and P is **already teal** going
back to the earliest backup on disk (13 Aug). No S or K slide with a
different colour exists anywhere on record.

**Rohan, 2026-08-22: "today you took the real colours and fixed the hidden
templates in our ppt."** Teal is the real, correct `_OFF` colour --
Claude had already propagated it to the hidden P/K/S exemplar templates
earlier this same session, before compaction. What got logged as CP was
that same correct fix, mischaracterized as a regression by the
pre-compaction summary and then written into FIX-LIST/CHECKLIST/
NEXT-SESSION.md without re-verifying it against the actual files first --
exactly the mistake this project has a standing rule against. No colour
fix was ever applied to the real production deck; a dry run was tested
against a scratch copy only and discarded once the citation fell apart.
No action needed. Lesson: a claim inherited from a compacted summary is
exactly as unverified as one from a stale handover doc, and gets the same
scrutiny before landing in a durable file.

## Built 2026-08-22 — CQ, milestone percentage display (Option A), live
## verification still pending

Full design in `MILESTONE-PERCENTAGE-DESIGN.md`. Rohan asked for a
Research Manager's informed percentage next to a not-yet-achieved
milestone circle, without changing which circle counts as "current."
Built as: a new register column per slot, `MS<n>_PCT` (free text, RM-
entered, optional -- blank means no opinion given), folded straight into
the existing `MS<n>_LABEL` text by `MilestoneDevice.DrawFromRow` (e.g.
"Fieldwork complete (75%)"). No new shape, no template change, no 43-slide
retrofit -- `DrawMilestones`'s signature and its 7 existing direct-call
tests in `TestRunner.bas` were untouched; the fold lives entirely in the
row-to-arrays translation layer. `IsColumnForThisDevice` extended to
recognise `_PCT` so it doesn't get flagged as an orphan column. New test:
`Test_MilestoneDevice_PercentageFoldsIntoLabelText`.

**Static check clean; NOT YET run live.** PowerPoint/Excel were both open
under a live session at build time, so `vba/tests/run_vba_tests.ps1`
(aborts if either is already running) couldn't be exercised -- same
constraint as CO above, same fix: run it next session before trusting this
beyond the static read. Also genuinely untested against real data: no
`MS<n>_PCT` value has ever been entered in the live register, so this has
never rendered on an actual slide.

