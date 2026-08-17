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
> -> 6.9s, tabs/index/format 53.7s -> 6.3s. AF added 2026-08-17 evening, still open --
> `PublishAllDraftedFields` redoes press-level work 13x, once per field (~4 min
> redundant register re-reads, ~2-3.5 min redundant saves, no fast-mode wrapper).
> AG added 2026-08-17 evening, still open -- `OfferMarkingForUnwiredFields` costs a
> full register read + full deck shape-walk per press, output destroyed before
> anyone sees it. AH added 2026-08-17 evening, still open -- harvest dry-run reads
> the register up to 3x in one press of button 1. AI added 2026-08-17 evening,
> still open -- `ScanPendingApprovals` computes dead detail for a gate deleted this
> morning, double-reads the review sheet. AJ added 2026-08-17 evening, still open
> -- `SyncNow`/`SyncNowCore` is fully dead code containing the worst call pattern
> found tonight (4x `BuildQueue` per type); recommend deletion, matching the
> bulk-approve precedent. AF-AJ full detail in `HOT-PATH-AUDIT.md`. AK added and
> FIXED 2026-08-17 evening -- `Readiness.bas`/`WhereAmI`/`WhereAmICore` deleted
> entirely (Rohan: "delete the whole thing"); every check it made was independently
> redundant with what the real operations already catch and explain when actually
> run, at a cost of two full deck-file copies plus a full BuildQueue diff per type,
> on every single press of the tool's most-used button. AL added 2026-08-17
> evening, still open -- the Lobby pin watcher (`AppEvents.cls`) taxes every
> single cell write anywhere in the tool (~100-170 COM calls per event) despite
> its own comments claiming "one comparison"; `ApplyApproved`'s fast-mode
> wrapper never disables it. AM added 2026-08-17 evening, still open --
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
> back large, on the real button.**

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

**Three times in one session.** Rohan pressed a button, nothing appeared, and the run
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

`AskForField`'s prompt lists every field in the workbook — around 40 lines by the time it
names the drafting fields, then the Given/Derived/Controlled ones with their kinds. On a
1080p screen that pushes the actual entry box below the screen edge. Rohan: **"what box?"**

He could not type into it because he could not see it. This is not a cosmetic problem: the
box returns "" when dismissed, which silently cancels the stage.

**Fix:** cut the prompt to the drafting fields only (the ones that can be answered), and
put the "these do not need a drafting sheet" explanation on a sheet, not in the dialog.
Same lesson as the Sync Now report: the container was the problem, not the wording.

### P3. The 21 `MS*` "fields with nothing to write into" warning is a FALSE POSITIVE

Fires on **every single run**, and the answer is always No. `FieldWiring.ScanFieldWiring`
compares register columns against individually tagged fields and has no concept of a
device consuming a column set, so every device-driven column reads as orphaned.

Answering Yes would walk the person through tagging 21 timeline internals as ordinary
fields — destroying the device.

**Fix is the device registry.** See NEXT-SESSION.md, "A DEVICE REGISTRY". Do not special-
case this one call site.

### P4. The 17-column prompt is ALL-OR-NOTHING across a mixed set

"17 field(s) on the Field Spec have no column in the register. Add a column for each?"
Sixteen are uncontroversial (`INDUSTRY_CASH`, `START_DATE`, `PROJECT_LEAD` …). One is
`HIGHLIGHTS_BODY`, which **must not** get a single column — it is three shapes per slide
and needs slot columns like the milestones.

So a real architectural decision is made, silently, by a Yes on a bundled prompt. Declined
twice on 13 Aug for exactly this reason.

**Fix:** offer the set per field, or exclude fields whose Renders-as implies slots.

### P5. Re-running the Template Audit REPLACES the sheet, decisions included

The dialog says so — *"Re-running this REPLACES that sheet, decisions included"* — which is
honest, and still wrong. The audit's whole purpose is to record field/chrome/drop
decisions against 50 items; losing them on re-run means the work can only ever be done in
one sitting. Same shape as the Discover Fields grid (item 3) which rebuilds from scratch
and loses marks.

**Fix:** carry decisions across by shape ID, the way `WriteDraftingSheet` carries drafts.

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

### B. `OfferHarvestForSelectedSlides`'s prompt mislabels and truncates

Two defects in one dialog, both in `RibbonUI.bas`:

1. **All propagation detail is accumulated into the `collisions` string**, so successful
   stamps print underneath a `Refused -- two fields matched one shape:` header. On
   14 Aug 23:57 that made a run reporting 16 correct stamps read as though it had refused
   everything.
2. **It hits `CapReport`'s 900-character cap mid-word**, so collisions can be present and
   invisible. A person cannot consent to what the dialog does not show them.

Neither is dangerous — the guards bound the write, not the text — but this project has
already paid twice for approving a prompt that could not be fully read.

### C. Slide 27 carries a shape already named `Text 216a` that is not the date

The 2026-08-15 rename pass (55 shapes, so 32 of 44 slides carry both `Text 212a` and
`Text 216a`) **refused** slide 27 rather than create a duplicate name. That slide's
`END_DATE` will keep colliding until a human looks at it. Correct behaviour, still an
open item.

### D. `check_vba_static.py`'s reachability check is weaker than its name

It asks whether a procedure's NAME appears in another module, **not** whether anything
reachable calls it. A chain of private orphans is invisible to it. Proven 2026-08-14:
commenting out the call to a wrapper left the callee's name still written inside the
now-orphaned wrapper, and the checker stayed clean.

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
