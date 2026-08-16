# The Lobby — architectural plan

> **CURRENT — Phase 0 built and tested 2026-08-16 (night); phases 1-4 not started.**
> Written at Rohan's explicit request before any of this got built: *"please do a full
> architectural plan for this before we start, large changes."* Supersedes the shorter
> version of this design scattered across `CHECKLIST.md`'s "The Lobby" section — that
> section now points here. Follow `SESSION-PROTOCOL.md`'s documentation discipline: read
> this in full before touching any of the modules it names, and update it in the same
> commit as any code that changes what it describes.
>
> **Phase 0 status:** `vba/DraftingLobby.bas` exists — `PinToLobby`, `ReadLobby`,
> `ClearLobbyEntry`, `LobbyCount`, and the cold-start `BuildLobbyFromScratch`. Three real
> tests in `TestRunner.bas` (`DraftingLobby_PinReadClearRoundTrip`,
> `DraftingLobby_BuildFromScratchFindsOnlyApprovedRows`,
> `DraftingLobby_PinTwiceUpdatesInPlaceNotDuplicate`), full suite green (233/0). Two real
> bugs found and fixed building this — both worth remembering as a CLASS, not an
> instance: a `Collection` cannot hold a VBA `Type` (compile error, not runtime — see
> `AGENTS.md`'s "Known Patterns", hit a fourth time here via `Collection` rather than
> `Dictionary`), and a bare `Application.DisplayAlerts` inside code that is compiled into
> PowerPoint's VBA project resolves to `PowerPoint.Application`, not the Excel instance
> that actually owns the workbook being changed — `wb.Application` is required, exactly
> the pattern already used everywhere else in `DraftingUI.bas`. The second bug was caught
> by the test that proves a rebuild does not trust its own past content, not by
> inspection — worth noting as a small proof that the "make it fail on purpose first"
> discipline pays for itself.
>
> **Not yet built:** phase 1 (the `Application.SheetChange` pin-on-tick event
> mechanism), phase 2 (wiring `PublishAllDraftedFields` to read the Lobby instead of
> crawling), phase 3 (pre-ticked queue items + removing the Yes/No/Cancel apply gate).
> `BuildLobbyFromScratch` is fully usable right now as a manual/on-demand rebuild, but
> nothing calls it automatically yet, and nothing reads the Lobby instead of the 13
> sheets yet — the crawl this design exists to remove is still the live behaviour today.

---

## 1. Why this exists — the problem, with real numbers

Tonight, getting **one** new field (the elapsed-time bar) from "data in the register" to
"visible on a slide" took two full presses of the same toolbar button, a build script
that silently discarded an edit, and roughly ten separate modal dialogs across the
session. That is not one slow step — it is the same expensive work happening two or three
times for what should be one action, plus a pile of ceremony dialogs around it. Rohan's
own words, unprompted: *"it now needs to start speeding up and be far less annoying to
use... I need to start showing it to people."*

Three distinct root causes, confirmed by reading the actual code tonight, not assumed:

1. **The drafting-sheet crawl is unconditional and full-cost, every press.**
   `DraftingUI.PublishAllDraftedFields` loops every Prose field (currently 13), and for
   each one calls `CopyAiDraftsToSubmit` then `PublishDraftsForField` — each of which opens
   that field's *entire* 43-row sheet and reads every relevant cell, whether or not
   anything on that sheet has changed since the last run. That is up to `13 × 2 × 43 ≈
   1,100` cell touches, visibly, with `ScreenUpdating` never suppressed (already logged as
   FIX-LIST item U) — on a press that, most of the time, finds nothing to do ("0 copied,
   38 left alone" was the actual result on nearly every run tonight).

2. **The button conflates two different actions and always redoes both from scratch.**
   "2. Put it on the slides" (`RibbonUI.PutItOnTheSlides` → `PutItOnTheSlidesCore`) means
   *"build me the list of what's changed"* the first press, and *"apply what I ticked"*
   the second — but **both presses re-run the entire chain**, including the crawl above,
   even though the only thing that happened in between is a person ticking one cell in
   Excel. The apply-only code path (`RibbonUI.ApplyApprovedChanges` → `ReviewQueue.
   ApplyApproved`) already exists as an isolated function — it is simply never reachable
   without re-running everything ahead of it (`PutItOnTheSlidesCore`'s own comment: "NO
   LONGER A BUTTON TARGET... Private so the reachability check reports genuine orphans").

3. **Confirmation dialogs fire on ceremony, not on real decisions.** The two "unsaved
   workbook, save?" guards (one in `ReviewChangesCore`, one in `ApplyApproved`) are
   questions with only one sane answer, asked twice per full cycle. The "holds N ticked
   change(s), Yes/No/Cancel" gate fires *after* the one real human decision (the tick)
   already happened, adding a second confirmation for the same choice.

None of this is safety. All of it is re-doing expensive work, and asking about decisions
that were already made, because nothing anywhere summarises "what actually needs
attention right now."

---

## 2. What does NOT change — read this before touching anything

This plan adds a new mechanism. It does not touch, weaken, or route around any of the
following, and any implementation that does should be treated as a bug in the
implementation, not a reinterpretation of the design:

- **`Drafting.WriteDraftingSheet`'s row-addressing is untouched.** This is the single
  most incident-prone function in the codebase — five real data-loss bugs between
  1–14 Aug (`ws.Cells.Clear` on the normal path, the layout-migration wipe, 27 lost
  drafted paragraphs). The Lobby is deliberately designed to solve the crawl and the
  two-loop problem *without* requiring any change to how or where drafting sheets store
  their rows. See §7 for why the earlier "merge the milestone sheets" idea was rejected
  in favour of this.
- **The drift-hash re-check at apply time is untouched.** `ReviewQueue.ApplyApproved`
  still re-reads the live slide and refuses to write anything whose hash no longer
  matches what was approved (`ChangeHash`, the "changed since you approved it;
  re-review" path). Pre-ticking (§5) changes *who has to act* to prevent a write from
  happening — it does not change whether the write is re-verified against live reality
  immediately before it happens.
- **`R13.3`'s diff-only filtering is untouched.** Unchanged, carried-forward content
  still never enters any queue. The Lobby only ever holds things that actually changed.
- **The register still holds content, never workflow state.** The Lobby is workflow
  state (what's pending, for which stage) — a new, parallel concept, not a rename or
  repurposing of the register. See §3 for the precise distinction (this is the question
  Rohan's own "what's the difference between the Lobby and the register" caught).

---

## 3. Where the Lobby sits — three concerns, not two

The register (`ExcelOutput`'s wide sheet) holds **content**: the actual field values,
current truth, no notion of "pending" at all. Historically this project has had exactly
one workflow-state mechanism sitting on top of it: `ReviewQueue`, which tracks "register
value differs from slide value, pending a human tick" — the second half of the pipeline
(register → slide).

**The Lobby is the same pattern, one stage earlier: drafting sheet → register.** It does
not compete with `ReviewQueue`, it complements it, and — this is the finding worth
keeping from tonight's design conversation — **it should reuse `ReviewQueue`'s
build/write/tick/apply shape rather than being built as parallel machinery.** The two
differ only in what they compare and how they write:

| | ReviewQueue (existing) | The Lobby (new) |
|---|---|---|
| Compares | register value vs. live slide value | drafting SUBMIT vs. register value |
| Pending signal | a computed diff (`PlanRoutineSync`) | the APPROVE tick itself |
| Writes via | `InjectPrimitive.InjectField` | `ExcelOutput.UpsertRow` (register write) |
| Sheet | `Review <slideType>-<hash>` | `Drafting Lobby` (one, shared, not per-field) |

```
   DRAFTING SHEETS  --tick-->   [ THE LOBBY ]  --publish-->   REGISTER  --diff-->  [ ReviewQueue ]  --apply-->  SLIDES
   (13 sheets today)                                        (content only)                                    (real deck)
```

---

## 4. The pin mechanism — event-driven, narrow, on purpose

**Rohan's own model, verbatim, because it is the precise spec:** *"it's more like the
work being pinned on a board by the author, and the crawler just looking at the board,
not every author's desk... it's pinned as part of authoring."*

- **Trigger: the APPROVE column only**, on any of the 13 (soon-to-be-however-many)
  drafting sheets. Not every keystroke — the tick is already the designated, deliberate
  "this is ready" signal (Step 5 of every drafting sheet's own printed instructions:
  *"Type Y in column H, save and CLOSE the file, then press '2. Put it on the slides'
  again"*). This is a rare, meaningful event, not a stream.
- **Mechanism: `Application.SheetChange`**, not thirteen separate `Worksheet_Change`
  handlers. A single class module (`WithEvents App As Application`) checks (a) is the
  changed sheet a drafting sheet (`Drafting.DraftSheetNameFor`-shaped name), (b) did the
  change touch `Drafting.COL_D_APPROVED`, (c) does the new value satisfy
  `ReviewQueue.IsApprovalMark` (already exists, already case-insensitive: `"Y"` or
  `"YES"`). If all three, write one row to the Lobby.
- **Bulk-write safety:** `WriteDraftingSheet`'s own rebuilds must run with
  `Application.EnableEvents = False` around any bulk write, restored in a `Finally`-style
  guard (VBA: `On Error GoTo` cleanup, not a try/finally, but the same discipline). This
  matters less than it sounds — `WriteDraftingSheet` already **carries forward an
  existing tick, never rewrites one** (confirmed reading the function tonight), so a
  normal rebuild should not fire the event at all even with events left on. Disable them
  anyway; do not rely on that as the only guard.
- **What gets written to the Lobby:** one row per pin — `SheetName`, `Row`, `FieldID`,
  `EntityCode`, `Timestamp`. Enough for the publish step to go straight to the exact cell
  without re-scanning anything.

---

## 5. The approval default — pre-ticked, opt-out, no exception list

**Scope boundary, settled the same night:** pre-ticking applies to **routine quarterly
sync only**. Onboarding a new deck or a batch of new projects is a critical-mass
event — a flood of genuinely first-time content arriving at once, where "the author
naturally scans it" breaks down regardless of sheet consolidation, for the same reason a
huge PR gets worse review than a small one. That case already has its own answer and
does not need a new one: `BatchOnboardFlow.bas`/`AdoptFlow.bas`/`DeckAdoption.bas` are
already a separate, deliberate, tag-by-tag confirmation flow, structurally distinct from
routine sync. As long as onboarded content goes through that flow before it ever reaches
the register, the critical-mass case never reaches the Lobby in bulk — it is handled
upstream, one deliberate decision at a time, same as today.

**Explicitly rejected: a numeric threshold (e.g. "20+ pending = force full review").**
Considered and rejected the same night — a count is a noisy proxy for "this is
onboarding," not the real signal. A busy routine quarter can legitimately produce a
large batch of small, already-reviewed corrections (safe to auto-flow, would be wrongly
blocked by a threshold); a single new small project can add only a handful of genuinely
first-time fields (unsafe to pre-tick, would wrongly slip through under a threshold).
The real distinguishing signal is categorical — *which flow the content arrived
through* — not countable, and picking a number where a structural signal already exists
is exactly the "threshold picked out of the air" mistake this project has been burned by
before.

Settled the same night, after real back-and-forth, not a snap call:

> **Every field the queue ever shows arrives pre-ticked `Y`.** Working the queue is
> removing ticks from what should *not* sync this round — not adding ticks to bless what
> should. No exception list. `PROJECT_STATUS` and every other field with no
> drafting-sheet gate is pre-ticked too.

**Why this is safe, precisely (this took three passes to get right, recorded so it isn't
re-litigated from scratch):**

- The queue only ever contains **real diffs** (§2, `R13.3` unchanged) — never an
  undifferentiated dump of the whole register. Pre-ticking a list that's already
  filtered down to "things that changed" is a different risk profile than pre-ticking
  everything that exists.
- Most fields in the queue already passed a **real human decision upstream** — the
  drafting sheet's SUBMIT + APPROVE tick. By the time that content reaches the
  register-to-slide stage, someone already deliberately signed off on the text once.
  Requiring a *second*, independent tick for the same content is the exact
  double-approval redundancy Rohan caught earlier the same night with `ABOUT_BODY`
  (approved once in drafting, asked to approve the *identical* diff again with zero new
  information).
- **The one genuine gap:** fields with no drafting-sheet stage at all (`PROJECT_STATUS`,
  typed straight into the register) have never had *any* human gate before reaching this
  queue — this is the literal field, and the literal failure shape, that created R13
  (19 slides changed, nobody looked, before R13 existed). Rohan's counter, considered and
  accepted: consolidating small/device-style fields onto fewer sheets (the direction
  already agreed for the milestone-family fields, §7) means the author sees all of them
  as a natural part of authoring a quarter, not buried in a wide, easy-to-skim-past
  register. That is a genuinely different visibility story than the one that produced the
  incident.
- **Residual risk, on record, not a blocker:** the incident happened because visibility
  was *assumed* and never actually checked. Any implicit-visibility design carries the
  same shape of risk if it recurs — something changes while attention is on a different
  part of the same consolidated sheet. Accepted deliberately, given the structural
  difference from the original register-scale problem. If this needs revisiting, the
  concrete trigger is: a wrong value reaches a real slide because it was pre-ticked and
  nobody noticed it in the pile. That has not happened yet under this design because
  the design has not been built yet.

---

## 6. Modal reduction — the ≤2-per-chain target

Independent of the Lobby but designed to land alongside it, since both attack the same
"too much ceremony" complaint. The governing principle already exists in this project —
`DOCUMENT-MAP.md`: *"buttons are not a safety mechanism — the confirmation is"* (why the
toolbar went 16 → 2 on 9 Aug) — just never applied to the dialogs *inside* a button's own
chain.

**"Set up my quarter" (`DraftingUI.StartQuarter` → `RollForwardUI` → `RefreshDraftingSheets`):**
currently up to 5 modals (type period, confirm Yes/No, period-set result, roll-forward
result, drafting-sheets-ready result). **Target 2:** the type-and-confirm prompt stays (it
is the one real decision); fold the three result messages into a single combined summary,
the same folding technique `PublishAllDraftedFields` already uses *within* one step
(`-- Publish Drafts --` / `-- Copy AI to Submit --` already concatenate into one dialog)
— apply it *across* steps too.

**"Put it on the slides" (with the Lobby built):** currently up to 6 (publish result,
review-build result, save-guard, Yes/No/Cancel apply gate, save-guard, apply result).
**Target 2:**
- The two "unsaved workbook, save?" *questions* become **silent auto-saves** — there is
  no real scenario where the honest answer is "no, don't save." One guard, once, at the
  top of the whole chain, not one per macro invocation.
- The "holds N ticked change(s), apply now? Yes/No/Cancel" gate is **removed outright**
  once pre-ticking (§5) is in place — pressing the button *is* the confirmation, same
  logic as the toolbar cut. This was Rohan's own explicit priority call the same night:
  *"that should all be one approval step, prioritise it after this."*
- Remaining: one combined "here's what's pending, go remove what you don't want" summary,
  one final apply-result summary. **2.**

**Informational-only results should not block at all where avoidable.** Office's
non-blocking `Application.StatusBar` can carry "Published ABOUT_BODY: 1 row" the way a
toast notification would, clearing itself, while full detail stays in the `Run Log` /
`Sync Log` sheets (already written, already durable) — summary-first, detail on demand,
rather than scrolling a person through every step inline.

---

## 7. Why sheet-merging was considered and rejected as the primary fix

Rohan's original question was narrower than the Lobby: *"can't complex shape groups be
drafted on one sheet together? Like the timeline elements?"* — specifically the 5
`MS2_LABEL`..`MS6_LABEL` sheets, structurally identical, a few words of content each.
Reasonable on its face — fewer sheets directly reduces crawl time and tab-clutter — but
checking `Drafting.WriteDraftingSheet` showed it assumes **fixed, absolute row
positions** (`DRAFT_INTRO_ROW=1`, `DRAFT_HEADER_ROW=9`, `DRAFT_FIRST_ROW=10`, always).
Combining fields onto one sheet means every one of those needs to become offset-aware —
real surgery on the exact function with five prior real incidents.

**The Lobby solves the same underlying complaint (crawl, ceremony) without touching that
function at all.** Sheet-merging is not ruled out forever — if the Lobby alone doesn't
fully address the sheet-count/tab-clutter part of the complaint, it is worth revisiting
on its own, as a separate, smaller, carefully-tested change to `WriteDraftingSheet`
specifically (see §9, phase 4) — but it is not the recommended first move, and should not
be bundled into the Lobby's own rollout.

---

## 8. Modules touched, and how

| Module | Change |
|---|---|
| **New: `DraftingLobby.bas`** (or a `Lobby` sheet + functions in `Drafting.bas` — name TBD, see §10) | `BuildLobbyFromScratch` (cold-start / repair path, see §9 phase 0), `PinToLobby(sheetName, row, fieldId, entityKey)`, `ReadLobby() As Collection`, `ClearLobbyEntry(...)` after a successful publish. |
| **New: `AppEvents.bas` (class module)** | `WithEvents App As Application`; `App_SheetChange` handler — the narrow trigger described in §4. Needs to be instantiated once and kept alive (a module-level `Public gAppEvents As New AppEvents`, wired in `Auto_Open`, matching how `CommandBarUI`'s toolbar already gets built there). |
| **`DraftingUI.PublishAllDraftedFields`** | Reads the Lobby instead of looping `ProseFields(wb)` blind. For each Lobby entry: `CopyAiDraftsToSubmit`/`PublishDraftsForField` narrowed to touch only that one row, not the whole sheet. Clears the Lobby entry on success. |
| **`RibbonUI.PutItOnTheSlidesCore`** | The Yes/No/Cancel gate removed (§6). Pending-approvals check becomes "does the Lobby/queue have entries" rather than a fresh `ScanPendingApprovals` crawl. |
| **`ReviewQueue.BuildQueue`** | New items arrive with `Approved = True` by default (§5), not `False`. Everything else about queue-building (diff detection, hashing, batching) is unchanged. |
| **`Drafting.WriteDraftingSheet`** | **Not touched.** Named here only to be explicit that it stays out of scope. |
| **`CommandBarUI.bas`** | No new buttons — the toolbar stays at 2, per the already-ratified "buttons are not a safety mechanism" call. |

---

## 9. Rollout order — build and prove each piece before the next

Matches `SESSION-PROTOCOL.md`'s mandated discipline: static checks + full suite before
every commit touching `.bas`, and for a fix specifically, *prove it fails on purpose
first*. For a new mechanism like this, the equivalent is: prove each phase live, on a
real deck, verified from saved file bytes — the standard this project has actually held
itself to all along, not a green test suite alone.

0. **Cold-start path first, before the event mechanism.** `BuildLobbyFromScratch` — a
   one-time full crawl (the *existing* slow crawl, unchanged) that populates the Lobby
   from whatever's already ticked across the 13 sheets. Needed regardless of the event
   mechanism, for: the very first run after this ships, recovery if the Lobby sheet is
   ever deleted/corrupted, and the at-work no-Claude-no-Python hand-edit gap (§4) —
   someone can always force a full resync by re-running this, even if no pins fired.
   **This alone is testable and shippable independently** — it does not require the event
   mechanism to exist yet, and de-risks the rest.
1. **The event mechanism**, tested in isolation: tick a cell, confirm exactly one Lobby
   row appears, with the right sheet/row/field/entity. Then the negative case: run a full
   `WriteDraftingSheet` rebuild with an existing tick already present, confirm *zero* new
   Lobby rows (proving the "ticks are carried forward, never rewritten" assumption holds
   under events-enabled conditions, not just assumed from reading the code).
2. **Wire `PublishAllDraftedFields` to read the Lobby.** Prove the crawl visibly
   disappears — time a real "Put it on the slides" press before/after, on a deck with a
   realistic number of untouched sheets.
3. **Pre-ticked queue items + the Yes/No/Cancel removal**, together — these are two halves
   of the same UX change and should land in one press-count-verified test: tick one field
   in a drafting sheet, one press of the button, confirm it reaches the slide, count the
   actual dialogs shown.
4. **Only after 0-3 are proven live:** revisit sheet-merging (§7) as its own, separate,
   carefully-scoped change if the sheet-count/tab-clutter complaint still stands once the
   crawl and the two-loop pattern are gone.

---

## 10a. A UserForm authoring front-end — real idea, deliberately separate from this build

Raised the same night: a `UserForm` pulling disparate drafting sheets together into one
unified authoring window (type SUBMIT, tick APPROVE, across every field, without
navigating sheet tabs) — "to aid with user satisfaction." Genuinely good, standard VBA
pattern for exactly this. **Not the same layer as the Lobby**, and not built alongside
it: the Lobby is backend tracking (what's pending, without re-crawling); a UserForm
would be the front end for actually *authoring* the content. Complementary, neither
depends on the other existing first. Logged here so it isn't lost, deliberately not
folded into this build — scope is already large enough to finish cleanly without it.

## 10. Open decisions — still Rohan's call, not resolved here

- **Exact Lobby sheet name and column layout.** Draft above (`SheetName, Row, FieldID,
  EntityCode, Timestamp`) is a starting point, not final.
- **Module organisation** — a new `DraftingLobby.bas`, or functions folded into the
  existing `Drafting.bas`/`DraftingUI.bas`. Given `Drafting.bas`'s incident history (§2),
  there's a case for keeping the Lobby's own code in a clean new module rather than adding
  surface area to the riskiest file in the codebase.
- **What "repair the Lobby" looks like as a button/action** for the at-work,
  no-macro-running hand-edit gap (§4, §9 phase 0) — does phase 0's cold-start function get
  its own toolbar entry, or live inside an existing one?
- **Whether `PROJECT_STATUS`-style fields ever get their own drafting-sheet-equivalent
  gate** instead of relying purely on consolidated-sheet visibility (§5's residual risk) —
  parked, not rejected; revisit if the visibility assumption is ever actually tested and
  found wanting.
