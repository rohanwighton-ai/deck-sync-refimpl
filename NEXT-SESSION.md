# NEXT SESSION — start here

> ## THE SUITE IS RED. FIX THIS FIRST — IT IS 15 MINUTES.
>
> **190 passed / 2 failed**, and both failures are expected and understood:
> `Drafting_PeriodRolloverDropsStaleSubmit` and
> `Drafting_RolloverKeepsEntityStaticRows`.
>
> Those two tests assert the OLD behaviour — that a period mismatch silently DROPS
> drafting. That behaviour was changed on 13 Aug because it was one button press from
> destroying 129 drafted values. `WriteDraftingSheet` now REFUSES and changes nothing
> when a mismatch would discard typed work.
>
> **The two tests are asserting the defect.** They need rewriting to the new contract:
> a mismatch with typed work REFUSES; a mismatch with EMPTY rows may still rebuild and
> drop, which is what those fixtures should exercise.
>
> `Test_Drafting_RefusesRatherThanDiscardOnPeriodChange` is new, passes, and was made
> to fail on purpose first — with the guard disabled it reported the submitted text as
> `''`, which is the loss demonstrated in miniature.
>
> Do NOT revert the guard to make the suite green. The guard is the fix.


**Written 13 August 2026.** Previous version archived as `NEXT-SESSION-2026-08-12.md`.
Bridge copy: `OneDrive\Claude\NEXT-SESSION-deck-sync-v6.md`.

---

## THE ONE THING THAT MATTERS

**Zero fields have ever reached a slide.** 48 recipes, 129 drafted values, a source
pipeline, a timeline device, 191 passing tests — and the delivery count has never left
zero. Publishing one project's three drafted fields is the only item that moves it, and it
keeps getting deferred behind mechanism.

Do that first. Everything below serves it.

---

## IMMEDIATE STATE

**The real deck** is `OneDrive\Claude\3. Project Progress.pptx` (Rohan's ruling, 13 Aug —
his original at work is superseded). 43 slides, one layout.

- **Period `Q3F26`** — written directly into `docProps/custom.xml` from WSL, because
  `Start a Quarter` could not write it (defect 2 below)
- **Nine fields TAGGED on slide 1**, each verified by read-back: `PROJECT_CODE`,
  `PROJECT_NAME`, `PROJECT_PROGRESS`, `PROJECT_STATUS`, `STRATEGIC_ALIGNMENT_BODY`,
  `ABOUT_BODY`, `PROBLEM_BODY`, `PROGRESS_BODY`, `KEY_EVENTS_BODY`
- **NOT onboarded** — no slide type registered, no pairing on disk, no tags on slides 2–43
- **Timeline on slide 1 complete**: `MILESTONE_TIMELINE`, 37 shapes, one level, no
  duplicate names. Slides 2–43 still have their original single-circle timelines.

**The register** is `OneDrive\Claude\register-wide.xlsx` — 17 sheets, 32 columns,
Q3F26 43 / Q4F26 43 / Q1F27 5. **21 `MS*` columns added, 38 of 43 projects populated.**
Drafting sheets hold 129 drafted values in columns E and F.

**Build: `addin79`**, installed here and in the Claude folder for work.

**Backups**: `OneDrive\Claude\backups\2026-08-13-1202-pre-ONBOARD-run2 - *`.

---

## NEXT ACTION

Onboarding should now work in one pass: the fields are already tagged, so the Discover
Fields grid — which lost his marks twice — is bypassed.

1. Open the deck, `1. Sync Now` → **Yes**
2. Period already set → OK
3. Workbook path → `C:\Users\rohan\OneDrive\Claude\register-wide.xlsx`
4. Should skip discovery and go to **Bulk Onboard** → 42 slides
5. Slide type name → `project-progress`
6. Register prompt → use existing `Register` → **Yes**

**LEAVE EVERY EXCEL WINDOW OPEN until the chain finishes.** Closing the scratch review
workbook killed a run with Error 424.

### THEN STOP. DO NOT PRESS SYNC NOW A SECOND TIME.

**All seven drafting sheets are stamped `Q4F26`; the deck declares `Q3F26`.**
`WriteDraftingSheet`'s period guard DROPS drafting when the stamps disagree, and
`Sync Now`'s main chain calls `RefreshDraftingSheets` as step 3. Once the deck is
onboarded, the next press silently discards 129 drafted values.

**Fix before that button is pressed again: the guard must REPORT AND REFUSE, not
discard.** Losing a person's typing to protect them is the wrong trade.

---

## DEFECTS, IN PRIORITY ORDER

**1. The period guard destroys drafting** (above). Not fixed. Blocks the next press.

**2. `SaveAs` returns clean and writes nothing** on this 49MB deck in a synced folder —
four attempts, file untouched, `Err` never set. Ordinary `Save` works. **Partially fixed**:
`SetDeckPeriodVerified` and `SetWorkbookPathVerified` now try `Save` first and escalate.
**Never actually run on the real deck** — the period was written from WSL instead, so the
fix is unverified. `SaveDeckVerified` already had the correct order sixty lines away.

**3. The Discover Fields grid loses marks on every re-run.** Rebuilt from scratch each
time, so a failed chain costs all the typing. This is the intended route at work.

**4. The workbook picker uses PowerPoint's Save As dialog**, which takes no file-type
filter and appends `.pptx`. Fixed with `NormaliseWorkbookPath`, but the dialog still
offers presentation types.

**5. No period dropdown.** Free-typed, exact-matched; a typo reads as a clean run of zero
rows. Needs a UserForm from the register's own periods — "picked, never typed".

**6. A dead pairing cannot be repaired** from the setup path.

---

## FIXED AND SHIPPED TODAY (191 green behind the compile gate)

- **Pairing fix** — onboarding no longer invents an empty register beside a populated one.
  `RegisterShapedSheets` scans the header ROW (not A1 — a test fixture caught that); more
  than one register is refused.
- **Re-onboard guard** — keys off "has this deck been onboarded before", not a value
  comparison, because on a first onboard register and slides legitimately differ.
- **Quarter before onboarding** — a virgin deck could not reach `StartQuarter`.
- **Discover Fields establishes the pairing** instead of refusing and pointing back at the
  button just pressed.
- **An existing workbook is opened, never created over**, and `CreateWorkbook` refuses an
  existing file outright. This was one click from destroying the register.
- **Register column creation from the Field Spec**, with a `Derived` carve-out.
- **Milestone device**: `MSn_NOW` four-state visibility, track shortens to the last USED
  slot, integrity check now verifies circles at all.
- **`START HERE` column letters derived** from constants.

---

## THE TIMELINE

Slide 1 done. Slides 2–43 need the group copied — **only after** their milestone values
are in the register, or the paste destroys them. 38 of 43 harvested; five have no timeline
(`P008`, `2_P009`, `1_P010`, `2_P012`, `S023`) — theirs were off-slide and were deleted.

**Rohan's insight, which removes a job:** those five don't need restoring. Put values in
the register, copy the group, sync — the slide is just a renderer.

**Copy method: programmatic, one-time, NOT an add-in feature.** Measured 13 Aug: a copied
group keeps its shape names AND its group tag. Slide 1 is GREEN; K and S need their own
colourways, read from each slide's existing circles rather than hardcoded.

**Harvest rule that works** (verified against three screenshots, one per palette): the
**oversized circle** marks current; everything at or above it is achieved. The bar tracks
nothing — slides 1 and 7 have identical bars at different stages.

---

## STILL OPEN, NOT BLOCKING

Stale `Q4F26`/`Q1F27` rows predate this deck and should be cleared before rolling forward ·
`HIGHLIGHTS_BODY` is three separate shapes per slide, so it needs slot columns like the
milestones · money/dates/team/subtitle have no register columns · the contribution-scale
question (if backbone Column R is legacy, every linkage inverts).

---

## TWO LESSONS WORTH KEEPING

**When the file and the object model disagree, the file is the evidence.** COM reported
the timeline group as flat; the XML said six nested sub-groups. Believing COM stranded six
shapes. Office also returns `Fill.ForeColor.RGB` as **BGR** — caught only because a
screenshot showed teal where the number said khaki.

**A defect is a class, not an instance.** The quarter-before-onboarding fix was made and
the identical shape two functions away — Discover Fields before pairing — was not looked
for, and cost a second failed run the same hour.
