# The nine scenarios

> **CURRENT — the frame development is steered by, with status re-derived from the code
> 2026-08-15.** Origin: `OneDrive\Claude\scenarios-from-claude-code-2026-08-14.md`, which
> proposed the list; that file is the proposal and this one is the live status. Two
> statuses moved on 15 Aug and nobody would have noticed, because the list lived only in
> a handover.

## Why scenarios and not more tests

Every defect that has cost this project a day is the same shape: **the parts worked and a
person could not reach them.** The picture injector, built and called by nothing. The
progress bars, built while every field went to the text injector. Publish's field
selection, working but reachable only for the first `Prose` field. The approve tick, read
by publish and cleared by the step immediately before it. The propagation machinery, fine,
behind a button that asked for a *new slide type*.

The suite was green for all of them. **A unit test asks "does this behave when called".
Nothing in a suite asks "can a person cause it to be called"** — and that second question
is what a scenario is.

So this is not documentation, it is the missing test layer, and the pass condition is
deliberately harsh: **Rohan completes it unaided, on a real deck, with no agent in the
loop.** Anywhere he has to ask is a defect, not a support call — at work there is no
Claude, no Python and no WSL.

## Status

| # | scenario | status |
|---|---|---|
| 1 | Generate a new quarter | **MECHANISM PROVEN 2026-08-15 (~16:50), NOT CLOSED — the pass condition is unaided and this run was not.** Run by button on a local copy of the REAL deck and register (`AppData\Local\deck-sync-quarter-20260815-1623\`): deck period `Q4F26` → `Q1F27` confirmed in the saved `.pptx`'s `docProps/custom.xml`; `43 row(s) copied from Q4F26 to Q1F27`, verified from the `Register` sheet as **43 rows / 43 distinct keys**, with `Q3F26` and `Q4F26` both intact at 43; 13 drafting sheets rebuilt and 13 parked. Content carried forward is real, not stubs (row 89, `3_P002`, `PROGRESS_BODY` cites `(Q3F26)` milestones correctly). **Nothing reached a slide — the `Sync Log` has no entry for the 16:23–16:52 window.** Whether publish belongs to this scenario or to 5 is arguable; what is not arguable is that Claude found the blocker and the row number, so "unaided" is unproven. **Next: press publish on that same rig, then re-run this unaided.** |
| 2 | Add a new project on an existing template | **DONE 2026-08-15 (late morning).** Run by button, unaided, on a local expendable copy: one register row added (`S999` / `Q4F26`) → `Add missing slides` → `1 created, 0 failed`, resequenced to row order, deck and workbook both saved. **Verified from the saved `.pptx`**, not the dialog: the package went from 44 to **45 slide parts** and `ppt/tags/tag666.xml` carries `SLIDE_TYPE=project-progress` + `INSTANCE_KEY=S999`. It was blocked first by unreachability (below) and then by a second one found on the day: a **copied deck keeps pointing at the ORIGINAL's register**, so testing "on the snapshot" would have written to the live file. See scenario 8 — the repair for that is now a button. Original note kept: **Unblocked 2026-08-15,** It was blocked by unreachability, not by a missing feature: `RunSync.CreateMissingSlides` existed and was tested, reachable only via `SyncNowCore`, called only from a `Private SyncNow` that nothing called. Its comment claimed *"the chain is the entry point"*; the chain calls `StartQuarter`, `RollForwardUI`, `RefreshDraftingSheets`, marking and discovery, and never this. It had been made `Private` **specifically so the reachability check would not report it**. Now reachable via the `Add or retire slides` button. Also note the direction: adding a project is **register-driven** — add a row, and the slide is created from the template — not "copy the template slide" as this list originally said. |
| 3 | Add a new slide type or derivative template | **UNBLOCKED for the colour-variant case (K/S/P), verified from saved bytes, not reasoned about.** The old blocker (`TemplateSlide.FindTemplateFor` matching on type alone) is no longer what governs this: `DeckRegistry.RegisterTemplateLetter`/`LookupTemplateLetter`/`LookupTemplateForLetter` (registration), `TemplateSlide.ExistingTemplateForLetter` (the creation guard), `DeckRegistry.RegisterNewTemplateLetter` (claims a letter's own slot, and the type-level fallback only if nothing already holds it) — all built, all tested against real Office via COM, 2026-08-16. `Create template slide` is a real, repeatable toolbar button (was reachable only as a one-time post-onboarding prompt until the same day — found live when it was pressed and nothing happened). **K and S templates built and verified 2026-08-16** on an isolated copy (`scenario3-template-surgery-20260816`, never the live deck): 46 slides (was 44), both new ones correctly hidden/tagged/keyless, `DeckSyncTemplate:project-progress:K`→slide 45 and `:S`→slide 46, `DeckSyncType:project-progress` unchanged at the original P template (303) confirming the fallback wasn't stolen. **Not yet proven: end-to-end routing.** `RunSync.CreateMissingSlides` was given per-row letter-aware template selection (Scenario 3 step 3) and unit-tested with synthetic fixtures, but has never been run against a REAL register row that needs a K or S template — that is the next real test, not a formality: three fake rows (one per letter) into the register, then a real sync, is what actually proves a new K-project row gets cloned from the K template and not the P one. |
| 4 | Add a new field to an existing type | **DONE 2026-08-15.** Tag once on the template -> propagate the role to every slide -> harvest the values. Ran by button across 43 projects: six fields at 43 of 43, `START_DATE`/`END_DATE` at 31 of 43 (12 slides display no dates). The origin doc called this *"the largest stopped data flow in the project"*, and named the cause exactly — the machinery was good and the entry point was wrong. |
| 5 | Draft a field and get it onto the slides | **Closed 2026-08-14 20:45**, `3_P001`/`KEY_EVENTS_BODY`, published by button. The blocker was the rebuild clearing the approve tick immediately before publish read it. |
| 6 | Correct a value already on a slide | **DONE 2026-08-15.** `PROJECT_STATUS` `'Not started'` -> `'Not Started'` on 8 slides, by button, unaided: review -> one `Y` in the batch's Approve cell -> `2. Put it on the slides` -> Yes. One tick covered eight changes via `PropagateBatchApprovals`. Verified from the SAVED `.pptx` against the pre-write backup, not from the dialog: before 8 old-casing / 1 already correct, after 0 old-casing / 9 correct — the untouched ninth is why "8 written" was honest. Backup confirmed on disk, written outside the synced folder. |
| 7 | Retire a project or a slide | **DONE 2026-08-15 (late morning).** Same sitting: the `S999` row removed → `Retire slides with no row` → warning named it exactly (`slide 44 -- S999`) → `1 slide(s) retired (deleted)`. **Verified from the saved `.pptx`**: back to **44 slide parts** and **0 tags carrying S999**. Original note kept: **Built 2026-08-15,** Same button as 2, asked separately: a slide whose key the register has no row for is DELETED, after a warning naming every slide by index and key. Rohan chose delete over hide — the register is the source of truth and last quarter's saved deck is the archive, so hiding would grow the deck forever to avoid a loss already covered. Guards: never the template (it carries a type and no instance key by design), never an unclassified slide. |
| 8 | Take the tool to a different deck or employer | **Untested.** The pairing mechanism exists. Whether a fresh deck plus a fresh register can be brought up from nothing, by Rohan, with no agent, has never been tried. This is the standing requirement — the tool is personal and must travel. |
| 9 | Answer "why does this field say that?" | **Designed, not built** (`PROVENANCE.md`). Provenance lives on the drafting sheet, which is rebuilt every period, so the answer has a shelf life of one quarter. |

## Weighting

- **1, 2 and 5 are "the quarter".** If those three run unaided the tool works and the rest
  is expansion. **2 and 5 are closed. 1's mechanism is proven and its unaided run is
  owed** — see the row above, and note the distinction is the whole point of this file:
  a scenario that works with Claude in the loop is a mechanism, not a scenario.

### THE STALE-STATE CLASS — found by running scenario 1, 2026-08-15

The quarter turn was refused, correctly and silently, by **five leftover `Q1F27` stub rows**
from earlier testing (4 populated cells each against 15 for a real row, text reading
`"New text for update"`). `ExcelOutput.RollForwardPeriod:919` refuses rather than
duplicating all 43 projects — right call, bad discovery. Clearing them let the run
complete.

**The same leftovers are in the LIVE register**, so the real quarter will stop the same
way. Found in the copy on 15 Aug:

| leftover | state |
|---|---|
| `Q1F27` stub rows | 5 — refused the roll-forward |
| `Review project-status-2D3D` | **38 `Y` ticks** from 10 Aug, still live |
| `Review project-progress-A32C` | 49 rows, 0 ticks — clean |
| `SAVED` park sheets | 26, two quarters' worth, of 59 sheets total |

**The gap: nothing checks the register for stale state before a quarter turn.** It refuses
at the point of collision, one item at a time, with a message that reads as a failure
rather than a checklist. A pre-flight naming all four in one dialog turns an hour into a
minute — and this is the class most likely to STALL a real quarter at work, where there is
no Claude to explain the refusal.

**Two more defects from the same run:**
- **`Roll Forward` asks a question it can answer itself.** It requires clicking a cell in a
  row of the source quarter — hunting a 90-row, 49-column sheet — when it already knows the
  deck's period and which periods hold rows. Picking a `Q3F26` row instead of `Q4F26` rolls
  a stale quarter forward and looks exactly like success. Same shape as the field picker
  Rohan deleted: *it shouldn't have to ask.*
- **A bare Excel `"permanently delete this sheet"` prompt** fires during the drafting-sheet
  rebuild — Excel's own alert, `DisplayAlerts` unsuppressed, naming neither the sheet nor
  the reason. Either suppress it or ask in the tool's own words.

### FILE-PER-QUARTER DELETES WORK — it is not hygiene, and it is mis-ranked

Confirmed from the code 2026-08-15: **not built.** `ParkSheetCopy` (`Drafting.bas:446`)
copies sheets *within* one workbook; the only `SaveCopyAs` is for DECK backups
(`ReviewQueue.bas:1153`). `EXPECTED-TRACE-2026-08-14.md:94` calls it *"GAP 4 — THE BIG
ONE."*

If each quarter were its own register file, four open problems close at once:
`ParkSheetCopy` becomes unnecessary (last quarter's file IS the archive — the ruling
already made); the destructive rollover clear becomes unnecessary; **the defect class that
wiped 43 approve ticks disappears**, because nothing rebuilds over work worth keeping; and
the Excel delete prompt goes with it. It keeps being deferred because it looks like
plumbing beside the template layer — that ranking is wrong, and the 26 park sheets now
sitting in a 59-sheet workbook are the evidence.

**A fifth: `Sync Log` is a second, independent append-forever sheet, and the prune
half must sweep it up too.** Checked from the code 2026-08-15, prompted by Rohan asking
whether the run-history sheets are needed at all: `Run Log` is NOT part of this problem —
its own comment says it is `REPLACED each run, not appended`, one bounded sheet, and it
earned its keep the same evening being read live to check what a chain had actually done.
**`Sync Log` is genuinely unbounded**: `ReviewQueue.ApplyApproved` (`ReviewQueue.bas:1403`)
appends one row per approved change — When/Run/EntityCode/FieldID/Outcome/Change ID —
walking to the first empty row, forever, inside the same file as the park sheets and every
quarter's register rows. **Do not delete it outright** — it is the only durable record of
what changed and when, i.e. the machinery scenario 9 (provenance) depends on. The right
fix is the one already planned: when a quarter is pruned out of the live register and
archived, its slice of `Sync Log` goes with it in that frozen file, and the live sheet
starts short again next quarter — same rule as everything else, not a separate mechanism.
**Add `Sync Log` to the prune half's scope now, before it is built, so it is not
rediscovered as a second problem later.**
- **3, 4 and 7 are "next year"** — surviving a growing deck and a change of employer.
  **4 is closed; 7 is built and unrun; 3 is still blocked.**
- **9 quietly matters most and is the easiest to defer forever.**

## What is left, honestly

**Scenario 1 is the last of "the quarter" that is neither closed nor built.** 2 and 7 are
now CLOSED (see the table) — the note that they were built and unrun is superseded.

**Do NOT test against `OneDrive\deck-sync-known-good\`.** That deck still stores the LIVE
register's path in its own `docProps/custom.xml`, so "testing on the snapshot" writes to
the real register. Use `AppData\Local\deck-sync-backups\
PRESERVED-known-good-20260815-1050\` (deck and register together, local, re-pointed and
proven), or re-point a fresh copy with the `Change which workbook this deck uses` button.

**The OneDrive risk is DIAGNOSED, and it is not what this file used to say.** The standing
claim was "nothing is proven on OneDrive; everything was proven on a local copy because
the OneDrive write failed outright". Measured 2026-08-15 midday on a scratch deck:
**plain `pres.Save` works on a OneDrive-hosted file.** What fails is `SaveAs`-to-self,
which raises `0x80CD1001` and leaves the presentation READ-ONLY so every later save fails
— and `DeckRegistry` calls it as its escalation whenever a read-back does not confirm
immediately. The tool breaks its own document and then reports that OneDrive did.
File size, AutoSave and sync latency were each tested and are each innocent. Full
evidence and the three call sites: `FIX-LIST.md` item **P**. Until P is fixed, work on a
local copy — not because OneDrive is unproven, but because a known defect is waiting there.

**Two defects found while closing 6, both still open (2026-08-15).**

1. **The review grid is not safe to rebuild under a live Excel AutoFilter.** With a filter
   active on the sheet, a rebuild left the previous grid's residue: 108 rows instead of 57,
   21 rows carrying a change ID and nothing else, 26 change IDs appearing twice, and a
   second copy of the batch pre-ticked `Y` — approvals no human made, on the sheet publish
   reads. `BuildQueue` and `WriteQueueSheet` are both innocent; `WriteQueueSheet` does
   `ws.Cells.Clear` first and writes all eight columns per item, so it cannot produce those
   rows. Turning the filter off and rebuilding produced a clean 57-row grid. Seen dirty
   twice under a filter and clean once without: strongly implicated, not proven — the
   proving test is to re-apply a filter and rebuild. Evidence workbook kept at
   `OneDrive\Claude\backups\EVIDENCE-doubled-review-sheet-20260815-0713.xlsx`.
2. **`Review project-status-2D3D` is still `OPEN` with 38 rows ticked `Y` from
   `2026-08-10 10:22:28`.** Five-day-old approvals sitting live. The change-hash design
   should stop them applying to values that have since moved, since the hash covers current
   and proposed values — but that claim is untested.

**Also open, cosmetic but the write-it-twice class:** the apply confirmation is titled
"1. Set up my quarter -- slide changes" and is reached from "2. Put it on the slides". A
caption hardcoded where it should be derived from the `CAP_*` constant.

**The pattern that keeps producing these.** Four of the nine were blocked by reachability
rather than by a missing capability: the picture injector, the progress bars, publish's
field selection, and now `CreateMissingSlides`. In every case the code existed and passed
its tests. **Scenario 3 is the only one blocked by an actual defect** —
`TemplateSlide.FindTemplateFor` returning the first match for a slide type.
