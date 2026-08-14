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
| 1 | Generate a new quarter | Partially built, **never run end to end.** Period-setting verifies against the saved file. The rollover refusal and cadence machinery are deleted; the quarter-turn ferry exists and has never touched the real workbook. |
| 2 | Add a new project on an existing template | **Unblocked 2026-08-15, not yet run.** It was blocked by unreachability, not by a missing feature: `RunSync.CreateMissingSlides` existed and was tested, reachable only via `SyncNowCore`, called only from a `Private SyncNow` that nothing called. Its comment claimed *"the chain is the entry point"*; the chain calls `StartQuarter`, `RollForwardUI`, `RefreshDraftingSheets`, marking and discovery, and never this. It had been made `Private` **specifically so the reachability check would not report it**. Now reachable via the `Add or retire slides` button. Also note the direction: adding a project is **register-driven** — add a row, and the slide is created from the template — not "copy the template slide" as this list originally said. |
| 3 | Add a new slide type or derivative template | **BLOCKED.** Re-verified 2026-08-15: `TemplateSlide.FindTemplateFor` matches on slide type alone and `Exit Function`s on the first hit. Three colour templates all tagged `project-progress` means whichever sits earliest in deck order silently wins. Needs a project-type key before the scenario is attemptable. |
| 4 | Add a new field to an existing type | **DONE 2026-08-15.** Tag once on the template -> propagate the role to every slide -> harvest the values. Ran by button across 43 projects: six fields at 43 of 43, `START_DATE`/`END_DATE` at 31 of 43 (12 slides display no dates). The origin doc called this *"the largest stopped data flow in the project"*, and named the cause exactly — the machinery was good and the entry point was wrong. |
| 5 | Draft a field and get it onto the slides | **Closed 2026-08-14 20:45**, `3_P001`/`KEY_EVENTS_BODY`, published by button. The blocker was the rebuild clearing the approve tick immediately before publish read it. |
| 6 | Correct a value already on a slide | **DONE 2026-08-15.** `PROJECT_STATUS` `'Not started'` -> `'Not Started'` on 8 slides, by button, unaided: review -> one `Y` in the batch's Approve cell -> `2. Put it on the slides` -> Yes. One tick covered eight changes via `PropagateBatchApprovals`. Verified from the SAVED `.pptx` against the pre-write backup, not from the dialog: before 8 old-casing / 1 already correct, after 0 old-casing / 9 correct — the untouched ninth is why "8 written" was honest. Backup confirmed on disk, written outside the synced folder. |
| 7 | Retire a project or a slide | **Built 2026-08-15, not yet run.** Same button as 2, asked separately: a slide whose key the register has no row for is DELETED, after a warning naming every slide by index and key. Rohan chose delete over hide — the register is the source of truth and last quarter's saved deck is the archive, so hiding would grow the deck forever to avoid a loss already covered. Guards: never the template (it carries a type and no instance key by design), never an unclassified slide. |
| 8 | Take the tool to a different deck or employer | **Untested.** The pairing mechanism exists. Whether a fresh deck plus a fresh register can be brought up from nothing, by Rohan, with no agent, has never been tried. This is the standing requirement — the tool is personal and must travel. |
| 9 | Answer "why does this field say that?" | **Designed, not built** (`PROVENANCE.md`). Provenance lives on the drafting sheet, which is rebuilt every period, so the answer has a shelf life of one quarter. |

## Weighting

- **1, 2 and 5 are "the quarter".** If those three run unaided the tool works and the rest
  is expansion. **5 is closed. 2 is unblocked and unrun. 1 remains.**
- **3, 4 and 7 are "next year"** — surviving a growing deck and a change of employer.
  **4 is closed; 7 is built and unrun; 3 is still blocked.**
- **9 quietly matters most and is the easiest to defer forever.**

## What is left, honestly

**Scenario 1 is the last of "the quarter" that is neither closed nor built.** 2 and 7 are
built and unrun, and both write to real files — which is what the known-good snapshot at
`OneDrive\deck-sync-known-good\` exists for. Run them against it, not against a hope.

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
