# NEXT SESSION — start here

**Written 13 August 2026, ~16:15.** Previous version archived as `NEXT-SESSION-2026-08-12.md`.

> ## THE DECK IS ONBOARDED. THE DELIVERY COUNT IS STILL ZERO.
>
> 43 slides tagged and linked, 0 failed verification, register saved, deck period
> `Q4F26` on disk. **No drafted value has reached a slide.**
>
> **The reason was found tonight, and it is one line of code.**

---

## THE BLOCKER, AND THE TWO-ROW WORKAROUND

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

**STILL UNPROVEN.** 192 tests pass and the project compiles, but no test exercises the
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
- **Build `addin80`**, stamp `2026-08-13 14:37`, in `OneDrive\Claude\` and in the trusted
  location `AppData\Roaming\Microsoft\AddIns\`.

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

## FILES CHANGED THIS SESSION (repo `deck-sync-refimpl`, uncommitted)

- `vba/DraftingUI.bas` — refusal count/names surfaced in the dialog; warning icon
- `vba/tests/TestRunner.bas` — two tests rewritten + renamed, runner registrations updated
- `FIX-LIST.md` — new items **1c** and **1d**
- `NEXT-SESSION.md` — this file

Not committed. Nothing else in the repo was touched.

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
