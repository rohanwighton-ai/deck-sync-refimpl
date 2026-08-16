# Phase 1 — scoped to three hours

Written by the delivery-check agent, 2026-07-31 evening. This is not a design document —
it is a budget. If it doesn't fit in three hours it isn't in here.

## What Phase 1 is for

Everything built so far proves the machinery can move real text to a real slide. Nobody has
established that the machinery is *worth using* — that question belongs to Rohan alone, and
today's manual-test doc says outright that his half (Steps 1–2) has not been run. Phase 1's
only job is to get that judgement rendered, on real content, and to remove the one mechanical
gap (single-row apply, never a batch) standing between "one demo row" and "a real quarter."

## The exit test

**Rohan has read and, where he chose to, rewritten the `ABOUT_BODY` text for at least 10 real
projects in the drafting sheet, published and applied them to the e2e-rig deck copy, opened
that deck, and can say in one sentence whether this was faster than doing it by hand.**

That sentence is allowed to be "no." A "no" that names why the sheet was clunky is a
completed Phase 1 and a useful spec for Phase 2. A "yes" is also complete. What is NOT
complete is silence — nobody having looked.

Nothing about field count, test count, or slide count is part of this exit test.

## Status, 2026-08-01

**Step 1 has been run.** Rohan read `TPL_ABOUT_BODY` cold and answered all four checkboxes
yes — *"not perfect but ok for test."* The sentence in "What Phase 1 is for" above, that his
half has not been run, is now out of date for Step 1 only.

Remaining, and it is the whole of the exit test: **the 10 real projects.** Draft, tick,
publish, apply, open the deck, say the sentence.

Two things cleared out of the way to get here, both invisible from the outside: the VBA
project did not compile (`Variant` → `As Object` ByRef), and beneath that the register was
being read by tab position and had been silently returning the `START HERE` sheet. See
TRACKER.md. Neither was a Phase 1 item; both stood in front of one.

## The three hours

| Time | Who | What | Why this and not more |
|---|---|---|---|
| 0:00–0:45 | **Agent (code)** | Prove `apply` on a *batch* of approved rows, not one. Today's loop was verified for exactly one approved `ABOUT_BODY` row at a time — R13's "unique prose, never batched" rule means each row is still an individual decision, but the driver's apply/verify loop has never run with more than one row `Approved` in the same pass. Confirm no cross-contamination (row A's text doesn't land on row B's slide), confirm the gate lists all N before-and-afters, confirm verify-count matches N. Budget includes one full 8-minute suite re-run and one compile-error round, because that's today's actual failure rate, not the best case. | This is the one real mechanical unknown left between "one row proven" and "a real quarter." Everything else the agent could build (batching UI, controlled-field batching, ribbon polish, the GUID redesign, chrome-tagging) is generalisation the spec doesn't need yet. |
| 0:45–1:00 | **Agent (code)** | Update `WORKPLAN.md`'s status header, which still reads "ONE field, mid-flight" against a day that closed at five fields and a real drafting loop. Fifteen minutes, not more — this is bookkeeping, not delivery. | Stale status docs are exactly the "artifact accumulates, count doesn't move" smell this review exists to catch. Fifteen minutes closes it; it doesn't need a design pass. |
| 1:00–1:20 | **Rohan** | Open the `TPL_ABOUT_BODY` sheet on `deck-sync-e2e`. Read it cold, no prior framing. Answer the four checkboxes already in `MANUAL-TEST-DRAFTING-LOOP.md` Step 1: legible at a glance, column C readable, F and G obvious, would you work down 43 rows of this. | This is the load-bearing 20 minutes of the whole plan. Everything upstream of it is scaffolding for this judgement; everything downstream depends on the answer being genuine, not rushed to hit a number. |
| 1:20–2:50 | **Rohan** | Draft (or accept-as-is) `ABOUT_BODY` for 10 real projects — pick ones you actually know, not the first 10 alphabetically. Tick, save. Publish, dry-run the gate, read every before-and-after it shows you (not skim — read), apply. | This is the actual content labour the tool is meant to save. Ten rows, not 43: three hours doesn't hold a full pass, and a partial real pass tells you as much about "would I use this" as a full one does. If it's taking meaningfully longer than 90 minutes per row's worth, that itself is the finding — stop and write down why. |
| 2:50–3:00 | **Rohan** | Open the resulting `e2e-deck.pptx`, find the 10 slides, read your own words on them. Say the one sentence the exit test asks for. | The whole point. Nothing after this is Phase 1. |

**Total: 3:00.** Agent code time: 1:00. Rohan content/judgement time: 2:00. They are not
interchangeable and not parallel — the agent's hour has to land before Rohan's can start
cleanly (batch apply needs to actually work before he trusts ticking more than one row).

## What Phase 1 explicitly does NOT include

- A second field (`KEY_EVENTS_BODY` etc. stay at "proven once," not extended).
- The real deck, or `deck-sync-live`. This stays on the e2e copy throughout.
- R13 batching for controlled fields (`PROJECT_STATUS`-style transformations) — untouched,
  parked exactly where it is.
- The chrome-tagging gap (MECE finding #5) — documented, deliberately not fixed here.
- The GUID/identity-key redesign in `DECISIONS.md`.
- Ribbon UI, onboarding new entities, or any Excel-side polish.
- All 43 rows of `ABOUT_BODY` — 10 is the sample size this budget affords.
- Fixing whatever Rohan's answer surfaces. If the sheet is wrong, that's Phase 2's brief, not
  something to patch same-night to protect the number.

## If three hours doesn't get there

If the 0:00–0:45 batch-apply slot blows out — plausible, given today's pattern of one "quick"
fix triggering another — the honest fallback is: **cut Rohan's sample from 10 rows to 3**, not
extend the agent's slot. His judgement doesn't need 10 data points to be genuine; the mechanism
does need to not silently corrupt a row when the agent is out of budget to keep debugging it.
Never trade his two hours to buy the agent a third.
