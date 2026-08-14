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
| 2 | Add a new project on an existing template | **Pieces exist, seam never exercised.** `SlideDuplication` strips the template marker so a clone cannot impersonate its template. No 44th project has been added, so "new slide" -> "new register row" is untested. |
| 3 | Add a new slide type or derivative template | **BLOCKED.** Re-verified 2026-08-15: `TemplateSlide.FindTemplateFor` matches on slide type alone and `Exit Function`s on the first hit. Three colour templates all tagged `project-progress` means whichever sits earliest in deck order silently wins. Needs a project-type key before the scenario is attemptable. |
| 4 | Add a new field to an existing type | **DONE 2026-08-15.** Tag once on the template -> propagate the role to every slide -> harvest the values. Ran by button across 43 projects: six fields at 43 of 43, `START_DATE`/`END_DATE` at 31 of 43 (12 slides display no dates). The origin doc called this *"the largest stopped data flow in the project"*, and named the cause exactly — the machinery was good and the entry point was wrong. |
| 5 | Draft a field and get it onto the slides | **Closed 2026-08-14 20:45**, `3_P001`/`KEY_EVENTS_BODY`, published by button. The blocker was the rebuild clearing the approve tick immediately before publish read it. |
| 6 | Correct a value already on a slide | Should work, **never deliberately tested.** The most common real action after a quarter ships, and its mirror — a slide hand-edited after approval — is what the review grid exists for. |
| 7 | Retire a project or a slide | **Not built.** Re-verified 2026-08-15: no routine removes or archives a slide. It surfaces only as a parity mismatch resolved by hand. Projects close every quarter. |
| 8 | Take the tool to a different deck or employer | **Untested.** The pairing mechanism exists. Whether a fresh deck plus a fresh register can be brought up from nothing, by Rohan, with no agent, has never been tried. This is the standing requirement — the tool is personal and must travel. |
| 9 | Answer "why does this field say that?" | **Designed, not built** (`PROVENANCE.md`). Provenance lives on the drafting sheet, which is rebuilt every period, so the answer has a shelf life of one quarter. |

## Weighting

- **1, 2 and 5 are "the quarter".** If those three run unaided the tool works and the rest
  is expansion. **5 is closed; 1 and 2 remain.**
- **3, 4 and 7 are "next year"** — surviving a growing deck and a change of employer.
  **4 is closed.**
- **9 quietly matters most and is the easiest to defer forever.**

## What is left, honestly

Scenarios 1 and 2 are the whole of "the quarter" now. Neither is blocked by a defect; both
are simply unrun, and both write to real files — which is what the known-good snapshot at
`OneDrive\deck-sync-known-good\` exists for.
