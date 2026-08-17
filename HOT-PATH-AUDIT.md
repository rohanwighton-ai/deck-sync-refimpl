# Hot-path audit — dead weight in the live button chains

Companion to `DRAFTING-SPEED-STRATEGY.md` (item AD) and `FIX-LIST.md` items
AF-AJ. Written 2026-08-17 evening, commissioned after deleting `Readiness.bas`
(the "Where am I" mechanism) turned up a real pattern worth checking for
elsewhere: machinery that survives the deletion of the human-facing surface
it was built to feed. Produced by a fresh research pass (model: fable,
read-only, no code changes), given the real measured numbers from tonight's
other fixes and this codebase's own incident history as context, reviewed
here before acting on it.

**Gist:** the two most-pressed buttons still contain leftover machinery —
mostly parts built to feed dialogs and sheets that have since been deleted —
and that machinery re-reads the same big spreadsheet and re-saves the same
file over and over inside a single button press. Several minutes of a full
publish run is this duplication, not real work.

**Scope and method.** Traced both hot chains end to end: "1. Set up my
quarter" (`RibbonUI.SyncNowChainCore`) and "2. Put it on the slides"
(`RibbonUI.PutItOnTheSlides`), plus every function they call. Cost figures
reuse tonight's own measured numbers from `FIX-LIST.md` (items W, Y, AB, AC,
AE) rather than fresh estimates. Frequency baseline, from the code's own
words (`DraftingUI.bas`): *"'2. Put it on the slides' — the button pressed
many times a session … '1. Set up my quarter' — pressed once at the
start."* Findings below are ranked by frequency × cost.

## Finding 1 (AF) — "2. Put it on the slides" redoes the whole-press work 13
times, once per field

The most-pressed button in the tool treats each of the 13 drafted fields as
its own separate button press: `DraftingUI.PublishAllDraftedFields` loops
the pinned fields and, per field, calls `CopyAiDraftsToSubmit` then
`PublishDraftsForField` — both written as standalone buttons, each still
carrying a full button's opening and closing ceremony. Per pinned field,
inside ONE press:

- **2x `Resolve`** — workbook path lookup, open-or-get, register-sheet
  resolve. The chain's own `Resolve` already did this moments earlier.
- **2x full register read** (`ExcelOutput.ReadSheetForDeckPeriod`) — once
  dry-run, once wet, nothing writes the register between them. Measured
  cost of this exact read on the real register: **9.3s for 43 rows** (item
  AC). 13 fields x 2 = **~4 minutes of redundant register re-reads per
  press.**
- **2x `SaveWorkbookVerified`** — a real disk save plus before/after
  timestamp checks, per field. Item AC measured the log+save stage at
  **8.8s**; 26 saves per press is plausibly **1.5-3.5 minutes**. One
  verified save at the end of the loop protects everything the 26 protect —
  `ApplyApprovedCore` immediately after re-checks dirtiness anyway.
- **2x `WriteRunLog`**, and `WriteRunLog` REPLACES the whole Run Log sheet
  every call (`ws.Cells.Clear`). Of the 26 log writes in a 13-field press,
  each erases the previous, and `ApplyApprovedCore`'s own final
  `WriteRunLog` erases the last one too. **Every one of the 26 is
  unobservable in a normal press** — the completion dialog still says "full
  detail is on the Run Log sheet" about content that is already gone.
- **2x `ShowSheet` tab activations per field, with `ScreenUpdating` ON** —
  this loop has NO fast-mode wrapper at all. Item AE just proved what that
  costs on this exact workbook: 2.67s wrapped vs. 168.6s unwrapped for the
  identical code, ~60x.
- **1x `PairingProblem`** per field — the pairing cannot change mid-loop;
  12 of 13 checks are answering an already-answered question.

**Why the duplication is safe to remove:** every load-bearing check lives
downstream, untouched by this fix. `ApplyApproved` re-validates every item
against the live slide by hash before writing, takes the backup, and
verifies its own writes — "THE REVALIDATION IS THE POINT," its own header
says. Nothing in the per-field ceremony is the last line of defence for
anything.

**Should parts of it exist at all:** the per-field DRY-RUN preview pass
existed to feed the "write these into the register?" consent dialog —
deleted 2026-08-14. In chain mode its only surviving consumers are
`NothingToPublish` (which the wet pass makes true by construction anyway)
and a Run Log entry the same press destroys. The mechanism outlived its
dialog.

**Fix direction:** restructure `PublishAllDraftedFields` to do the
press-level work once — one `Resolve`, one register read shared across
fields (or at minimum one per field instead of two, by deleting the
chain-mode dry pass), one collected Run Log write at the end (the report
buffer already collects exactly this text), one `SaveWorkbookVerified`
after the loop, `ShowSheet` suppressed while collecting, and an AE-style
`ScreenUpdating`/`Calculation` wrapper around the whole loop. Rough bound:
several minutes off every full publish press, on the button pressed many
times a session.

## Finding 2 (AG) — `OfferMarkingForUnwiredFields` costs a full register read
plus a full deck shape-walk per press, output destroyed before anyone sees it

Runs unconditionally on every press of "1. Set up my quarter", before
`StartQuarter`'s own dialog. Per registered type: a full register read
(9.3s measured) plus `FieldWiring.ScanFieldWiring`'s recursive shape-walk of
every slide of that type (item Y: a full walk of one real slide is
seconds-class; across 40+ real slides, plausibly tens of seconds per press).

Its "offer" died 2026-08-14 (own comment) — it never prompts and never
stops the chain. What survives is one `WriteRunLog` call. But
`RefreshDraftingSheets`, later in the SAME press, calls its own
`WriteRunLog`, which starts with `Cells.Clear` — **the note is destroyed
before the completion dialog appears.** Its other claimed surface ("reported
on START HERE instead") was Readiness's sheet, deleted tonight. Post-
deletion: all cost, zero observable output. The genuine remedy path
(tagging) is unaffected — it lives on "Tag fields on this slide"
(`DiscoverUI.DiscoverFields`), as the dying function's own comment says.

**Bearing on item AA:** this runs before `StartQuarter`'s dialog, in the
exact unmeasured gap AA describes. Tonight's `Readiness.bas` deletion
already removes a large share of that gap; this scan plus Finding 3's
harvest dry-run are the surviving pre-dialog weight — plausibly 15-40s of
it. Worth re-measuring AA after both are fixed, before instrumenting
`Resolve()` as AA's entry currently plans.

**Fix direction:** delete the call from the chain. If the unwired-fields
report is worth keeping at all, it belongs behind "Tag fields on this
slide" — the moment someone is actually thinking about tagging — not on
every quarter-setup press.

## Finding 3 (AH) — harvest dry-run gate reads the register up to 3 times in
one press of button 1

Unlike Finding 2, this is legitimately distinct work, not a delete
candidate — a pays-too-much-for-its-gate candidate instead. In one steady-
state press of "1. Set up my quarter" with a slide selected (normal-view
default per the code's own comment):

1. `OfferHarvestForSelectedSlides` -> `Harvest.HarvestSlide` dry-run ->
   `ReadSheetForDeckPeriod`, per selected slide.
2. `OfferMarkingForUnwiredFields` -> `ReadSheetForDeckPeriod` (Finding 2).
3. `RefreshDraftingSheets` -> `ReadSheetForDeckPeriod` (the measured 9.3s
   instance).

**Fix direction:** one register read per press, shared. The chain already
resolves `pres`/workbook once; pass the `Sheet` UDT down to all three
consumers instead of re-reading. Cheaper local alternative: gate the dry
harvest on its PowerPoint-side half first (needs no register) and only read
the register when there's at least one untagged/stampable shape — a
steady-state deck skips the read entirely. Saves ~9-19s per press either
way.

## Finding 4 (AI) — `ScanPendingApprovals` computes detail for a dialog
deleted this morning, and double-reads every review sheet

`PutItOnTheSlidesCore` calls `ScanPendingApprovals`, which per type does a
full row-by-row read of the review sheet (~5 COM reads/row; at Phase 3's
real scale of 221 items, 1,100+ cross-app reads). Moments later
`ApplyApprovedCore` does its own read of the same sheet — that second read
is the load-bearing fresh one. The caller uses only `pending = 0` from the
first read; the `sheetNames`/`stamp` by-refs it fills are NEVER used — they
existed for the Yes/No/Cancel gate deleted this morning (Lobby Phase 3).
Same pattern as Finding 1: machinery that outlived its dialog.

**Fix direction:** shrink `PendingApprovals` to the boolean actually
consumed — early-exit on the first approved, unconsumed row (the consumed-
banner check already short-circuits whole sheets) — and delete the dead
`sheetNames`/`stamp` plumbing. Small win per press at real scale, and it
removes stale-dialog residue that will otherwise confuse the next reader.

## Finding 5 (AJ) — dead code carrying the single most expensive call
pattern in the module

`RibbonUI.SyncNow`/`SyncNowCore` is `Private`, has no caller (the toolbar
targets `SyncNowChain` only), and builds `ReviewQueue.BuildQueue` FOUR
separate times per type in one run (plan/apply/post-apply-rebuild/parity
loops) plus three register reads per type — each `BuildQueue` being the
full register-vs-deck diff item X's live stall implicated as multi-minute-
class on the real deck. Not a live cost today, since nothing can reach it —
but if anyone ever rewires it, it becomes the slowest path in the tool
overnight.

**Fix direction:** this project's own precedent decides it — bulk-approve
was deleted for exactly this state ("unreachable code provides no
capability... git holds it"). Same explicit decision for `SyncNow`/
`SyncNowCore` (keep `WarnOnDuplicateKeys`, still legitimately used by
`ResolveDeckContext`/`ReviewChangesCore`). Delete rather than optimize;
nothing needs its 4x pattern fixed if it doesn't exist.

## Finding 6 (watch, not fixed) — every deck-property read is a full slide
walk

Since the 2026-08-16 move of settings onto a hidden slide, every "what
quarter is this deck?" question walks every slide to find that hidden slide
first (`FindRegistrySlide`, routed through by `GetDeckPeriod`,
`GetWorkbookPath`, `LookupType`, `ListRegisteredTypes`). Cheap per call
(in-process PowerPoint COM, low milliseconds), called dozens of times per
press. **No measurement shows this matters yet** — do not fix ahead of a
number. If a future Timing pass shows unattributed PowerPoint-side time,
caching the registry-slide reference per press is the one-line fix.

## Checked and clean — no action

`ApplyApproved`'s internals (dry probe, hash revalidation, per-run backup)
are load-bearing, not waste — the legitimate "evidence from the far side of
the boundary" discipline. `SaveDeckVerified`/`SaveWorkbookVerified`'s
`wasClean` guard already avoids a gratuitous save on a clean deck — the
*number of calls* is Finding 1's problem, not the functions themselves.
`OfferMarkingForSelectedShape`/`OfferAdoptionForSelectedSlides` are
selection-local and cheap. `ReadSheetForDeckPeriod`'s second read on the
zero-rows path is a legitimate diagnostic re-read, refusal-path only.
`RefreshDraftingSheets` itself is already measured and fixed tonight
(AB/AC/AE); its remaining known cost is item AD, with its own strategy doc.

## Cross-cutting observation

Four of six findings share one shape, worth naming because it predicts
where the next one is: **a mechanism built to feed a human-facing surface
survived the deletion of that surface.** Finding 1's dry pass fed a deleted
consent dialog; Finding 2's scan fed a deleted prompt and a deleted sheet;
Finding 4's detail fed a deleted three-way gate; Finding 1's per-field save/
log ceremony fed a per-field button that became a loop. Tonight's dialog-
deletion campaign correctly removed the stops but left their feeders
running. Cheap sweep for the class: for each `MsgBox`/prompt deleted since
2026-08-14, grep for what computed its inputs and ask whether anything else
still consumes them.

## The larger question this audit sits inside

A separate needs-vs-build comparison, run the same evening, found that
tonight's entire session (6+ hours) went to sync-speed engineering, while
the project's own manual-baseline memory states plainly that speed is NOT
the dominant cost of a quarter — the recipe-poor panels (Strategic
Alignment, Problem, Project Progress) are named as the highest-leverage
remaining work, untouched tonight. The stated finish line (Scenario 1: a
real quarter reviewed, approved, and published UNAIDED, no Claude in the
loop) hasn't moved. This audit's findings are real and the fixes are
individually cheap and low-risk — but doing them is still choosing to
extend tonight's own pattern, not correct it. Worth deciding explicitly,
not by default.
