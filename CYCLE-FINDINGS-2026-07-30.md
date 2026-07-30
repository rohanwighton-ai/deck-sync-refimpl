# First complete cycle — findings (2026-07-30)

Rig: `C:\Users\rohan\deck-sync-cycle\` — `cycle-deck.pptx` (3 slides carved from the
46-slide rehearsal copy, already onboarded) + its own 3-row `SAAFE-Projects-Data.xlsx`.
Add-in: `addin28`. Code at `e18a420` (102/102 green).

Purpose of the run: `Sync Now` and the recurring path had 102 tests and zero real use.
Everything below was found by USE, in the first fifteen minutes.

---

## Rig-building findings (before the cycle even started)

**R1. A carved deck still points at the original workbook.**
Copying a deck and deleting slides leaves `DeckSyncWorkbookPath` aimed at the full
46-row Data sheet. Preview Sync would have reported ~43 slides "would be created", and
a real sync would have created them. Not a code defect — a documentation/onboarding gap
for anyone who makes a deck by copying one. Worth a warning when the row count vastly
exceeds the slide count.

**R2. `DeckSyncId` is identical across rehearsal, sandbox and cycle decks.**
`943EC6E8-1089-4225-BEE0-297A952C3097` in all three, because they are all copies of one
original. Copying a deck clones its identity. Nothing keys off it globally today (the
workbook stores the same id via `WriteDeckReference`, so the pairing still matches), so
this is latent, not live. Related to the known design debt about instance keys wanting
to be immutable GUIDs.

---

## Cycle findings

**Step 1 — baseline Preview Sync: PASS.**
`3 unchanged, 0 would be corrected, 0 new slide(s) would be created, 0 flagged` /
`0 slide(s) are not in Data-sheet row order.` First time the preview has been pointed
at a deck it should say "nothing to do" about and correctly said it.

**C1. Preview Sync never shows the value it would write. (real defect)**
Changed `3_P001`'s Project Status in Excel; the preview reported:

```
would correct: 3_P001
    Project Status:
      now:  'Project Closed'
Summary: 2 unchanged, 1 would be corrected, 0 new slide(s) would be created, 0 flagged
```

Detection is correct — right instance, right field, `0 new`. But it prints only the
slide's CURRENT value and never the incoming Data-sheet value. To know what you are
approving you have to go and read Excel, which is the thing the preview exists to save
you from. The intent is written down and unimplemented — `SyncOperations.bas:26` says of
`ChangedFieldCurrent`: *"the whole point of a dry run is being able to show
before/after"*. It shows before.

Not caught by 102 tests because the tests assert on counts and on the presence of the
instance/field name, never on the report containing the target value.

Fix, sized: `sourceValue` is already in scope at `SyncOperations.bas:153` where
`changedCurrent` is filled. Add `ChangedFieldNew As Object` to `Public Type SyncAction`,
populate it alongside the other three dictionaries, set it on the action, and emit a
`new:  '...'` line under the existing `now:` line in `RunSync.bas:191`. Roughly six
lines across two files, plus a test that asserts the report contains the incoming value.

**C2. There was no Sync Now button. (the big one)**
The cycle reached step 4 — click Sync Now — and the button did not exist. The toolbar
had Preview Sync, Mark Field for Batch, Bulk Onboard Type, Clear Marked Fields.
`RibbonUI.SyncNow` has existed and been tested all along; `CommandBarUI.bas:76` had it
commented out under the rule "only add an operation when I'm fully clear it works."

The rule ate itself. The action could not be tried because it had no button, and it had
no button because it had never been tried. Nine sessions of hardening went into a tool
whose central action was unreachable from the UI, and the only reason that was not
obvious sooner is that nobody had run the recurring path end to end. This is the single
strongest argument for the "one complete cycle beats three hardening sessions" call.

**C3. Sync Now wrote without asking, and mass duplication was one click away.**
Found while wiring C2. `SyncNowCore` went straight to `RunRoutineSync` with no
confirmation. Combined with the fact that an orphaned Data row is classified
`new_record`, a deck with drifted linkage turns a sync into a mass slide duplication —
which `PreviewRoutineSync`'s own header already records as the live state of the real
deck on 2026-07-27 (43 orphaned rows against 46 slides), noting that *"only the button
being absent from the toolbar prevented it."*

So the missing button in C2 was, accidentally, the only safety mechanism. Adding it
without a guard would have removed the protection at the same moment it made the action
reachable.

---

## Fixed in addin29 (103/103 green)

- **C1** — `ChangedFieldNew` added to `SyncAction`, populated from `sourceValue`, and
  `RunSync.bas` now emits a `new:` line beneath `now:`. The existing preview test
  asserted only the "before" value; it now pins both, in their exact rendered form.
- **C2** — `Sync Now` on the toolbar. Toolbar tests updated 4 → 5 buttons and assert it
  by name. Their allowlist match was also tightened: it used a bare `InStr` over a
  pipe-joined string, so any substring passed — delimiters are now included in the match.
- **C3** — `RunSync.PlanCounts` (counts, not parsed prose) plus `RunSync.ConfirmSyncText`,
  wired into `SyncNowCore`, which now plans every registered type BEFORE writing any of
  them and asks first. Slide creation is stated in capitals with its consequence spelled
  out. A new test pins that wording, including that the capitalised warning is ABSENT
  when nothing will be created — a warning that always fires stops being read.

---

## THE FIRST COMPLETE SYNC — 2026-07-30 13:24, on addin29

The recurring path ran end to end for the first time in the project's life:

```
=== RunRoutineSync: q ===
  corrected: 3_P001
Summary: 2 unchanged, 1 corrected, 0 created, 0 failed, 0 flagged
Resequenced 0 slide(s) to match Data-sheet row order.
```

Slide 1's status field changed from `Project Closed` to `SYNC TEST 1248`, confirmed
visually on the slide, not merely reported by the tool. All three addin29 fixes were
exercised live in the same run: the preview showed `now:`/`new:` (C1), Sync Now was
clickable at all (C2), and it asked before writing with `0 new slides created` (C3).

**C4. The real sync's report is thinner than the preview's. (minor, unfixed)**
`RunRoutineSync` prints `corrected: 3_P001` and stops — no field names, no before/after,
where the preview gives all three. The moment you most want a record of what changed is
just after it changed, and that is the one report that does not say. Same fix shape as
C1: the data is already on the action.

**Slide creation works.** Adding row 5 (`3_P004`) produced:

```
=== RunRoutineSync: q ===
  created: 3_P004
Summary: 3 unchanged, 0 corrected, 1 created, 0 failed, 0 flagged
Resequenced 2 slide(s) to match Data-sheet row order.
```

A 4th slide appeared, cloned from the template, in correct row order. Resequencing also
ran for the first time (every previous run reported 0) and put it in the right place.

**C5. Sync reads Excel's UNSAVED in-memory buffer. (real, and invisible)**
The saved workbook on disk had rows 1-4 only, `dimension A1:F4`, last written 12:51 --
yet the 13:32 sync created a slide from row 5. `WorkbookBridge.OpenOrGetWorkbook`
attaches to the running Excel instance, so the sync sees whatever is on screen, saved or
not. The deck now contains a slide whose backing row exists nowhere on disk.

Why this matters more than it first looks, for a tool whose whole premise is "the deck is
fed by tracked data":
- Close Excel without saving and the slide becomes a permanent orphan -- its instance key
  matches no row, so no future sync will ever update it, silently.
- A sync run mid-edit writes half-typed values into the deck.
- Preview Sync's authority is weakened: the data can change between preview and sync
  without any file changing, so the preview is not a promise about the next write.

Suggested fix: refuse to sync a dirty workbook, offering to save it first
(`wb.Saved = False` is the check). Reading the live buffer could be defended as a feature,
but it must not be silent.

Worth noting how it surfaced: the cycle folder was deliberately put OUTSIDE OneDrive to
keep cloud-save behaviour out of the first cycle, which meant AutoSave was off. On a
OneDrive-hosted workbook AutoSave would usually have hidden this. Choosing the quieter
environment is what made the bug visible.

---

## Fixed in addin30 (104/104, commit `8d54832`)

- **C4** — `RunRoutineSync` now prints the same field detail the preview does, in the past
  tense (`was:` / `now:`). The end-to-end test asserted the deck was corrected but never
  the report, which is how the gap survived.
- **C5** — `WorkbookBridge.IsDirty` + `UnsavedWorkbookText`. Sync Now checks before
  planning and offers to save; Preview Sync warns at the top of its report instead of
  blocking, since it writes nothing.

**Recorded while fixing C5:** `Workbooks.Add()` reports `Saved = True`. Excel's flag means
"unmodified since last write", not "exists on disk". Asserted wrongly from memory, caught
by the suite in one run — the same class of mistake `feedback_verify_office_automation_
before_asserting` exists for, this time caught cheaply because it was written as a test
against real Excel rather than as a claim.

## Still unexercised

`New Period` — also not on the toolbar. It duplicates slides, so under the rule
established by C2/C3 it needs a confirmation guard before it goes on, and a live cycle of
its own. It is now the largest untested surface in the tool.

## Scoreboard

Six findings in one evening of real use. **None of them were logic errors**, and none were
reachable by the test suite as it stood: they lived in wiring (C2), messages (C1, C4),
missing guards (C3), file state (C5), and one that turned out to be correct behaviour
misread (C6, retracted). Nine prior sessions of engine hardening produced nothing
comparable, because each exercised a slice that had already been exercised.

The lesson is not "write more tests". It is that a test can only guard what someone
thought to guard, and using the thing is what generates that list.

---

# Progression step 1 — master template slide (same evening, 19:45–19:56)

Built and live-cycled immediately after the above, on the same rig. `addin31`, 109/109.

## What it does

Each slide type gets one **master template slide** that is never a real project: tagged
`is_template`, deliberately keyless, fields set to `<<placeholders>>`, hidden from the
slideshow, parked last, and registered as what the type clones from.

The hazard closed: `DeckSyncType:q` was `256|q` — SlideID 256, i.e. **slide 1, the real
project `3_P001`**. Every slide the tool created was a clone of that real project, so
everything the sync does not manage (figures, chart data, notes, untagged text) arrived on
the new slide belonging to P001 and *looking correct*. Correct tagged fields are what made
it plausible, which is the worst shape a reporting-tool defect can take.

## The live cycle — 6 checks passed, 1 finding (found by looking, not by the checks)

| # | Check | Result |
|---|---|---|
| 1 | Toolbar | 6 buttons, `Create Template Slide` present |
| 2 | Baseline Preview Sync | `4 unchanged, 0 corrected, 0 new, 0 flagged` |
| 3 | Create Template Slide → `q` | confirmation named `3_P001 (slide 1)`, `5 field(s)`, `'q' RE-REGISTERED`; created slide 5, hidden, placeholders written |
| 4 | Preview Sync with the template in the deck | **byte-identical to check 2** — template not counted, not flagged, not out of order |
| 5 | New row `3_P005` → Sync Now | `1 NEW SLIDE(S) WILL BE CREATED` guard fired; `created: 3_P005`, `1 created, 0 failed, 0 flagged`, `Resequenced 1 slide(s)` |
| 6 | Preview Sync after the create | `5 unchanged, 0 corrected, 0 new, 0 flagged` |

**Check 4 is what step 1 exists to make true.** A typed keyless slide is exactly case 6
(`unclassified_slide`), so without the exclusion the template would have been reported as a
flagged problem on every sync forever — permanent noise in the report a human reads for real
problems. The exclusion lives at one choke point (`RunSync.GatherInstances`), so planning,
correcting, counting and resequencing all inherit it.

**Check 6 is the one that could have failed silently.** `Slide.Duplicate` copies slide-level
tags, so the clone could have inherited `is_template` — and an inherited marker makes the new
slide invisible to every future sync, silently, *because the exclusion exists to keep
templates out of reports*. The symptom would have been check 6 reading `4 unchanged, 1 would
be created`, offering to create `3_P005` again forever. `5 unchanged` is the proof the strip
worked on a real deck, not just in the suite.

## Emergent, and it worked

`Resequenced 1 slide(s)` on check 5 was not planned. The new slide is born immediately after
the template (Duplicate's placement), gets packed back into row order at position 5, and
pushes the template to 6. So parking the template last is **self-maintaining** — it re-parks
itself after every create, with nothing managing it. That was the reason for choosing "last"
over "first" (`ResequenceByRowOrder` packs from the lowest keyed index, so a template among
the instances gets shuffled arbitrarily), and the mechanism turned out to be load-bearing in
the good direction.

## S1. Every created slide was HIDDEN. (real defect, fixed)

`Slide.Duplicate` copies `SlideShowTransition.Hidden` as well as the slide-level tags. The
master template is deliberately hidden, so **every record cloned from it arrived hidden from
the slideshow** — a brand-new project silently absent from the presented deck, with nothing
in any report saying so.

All six checks above passed with this defect live. It was found by Rohan looking at the
thumbnail pane and asking why *two* slide numbers were struck through when only the template
should have been.

**The shape of the miss is the useful part.** `DuplicateAndTag` already guarded the
`is_template` TAG against exactly this inheritance — with a postcondition that deletes the
slide rather than trust `Tags.Delete` — and the sibling PROPERTY three lines away was not
considered at all. Reasoning about "what does `Duplicate` copy" one attribute at a time is
what let it through; the question that would have caught it is "what else does it copy?"

Fixed by setting `newSld.SlideShowTransition.Hidden = msoFalse` unconditionally in
`DuplicateAndTag` — not conditionally on the source being a template, because an invisible
new record is a silent failure whatever it was cloned from. The test that covered this path
already existed and passed: it asserted the tag was stripped and never looked at the hidden
flag. It now pins both, plus that the template itself stays hidden.

## Why the "designed out up front" claim held only partly

The previous cycle found six defects in fifteen minutes; this one found one. The reduction is
real and the reason is not luck — the six findings had already told us where this class of
code fails, and each of those classes was designed out up front rather than discovered:

| Last cycle's failure class | Designed out this time |
|---|---|
| wiring (C2 — no button) | button added in the same commit as the action |
| messages (C1, C4 — report didn't say what changed) | result message states slide number, field count, and the re-registration |
| missing guards (C3 — wrote without asking) | `ConfirmTemplateText` written first, and pinned by a test |
| file state (C5 — read Excel's unsaved buffer) | inherited; fired correctly during check 5 |

But S1 is a **fifth** class that list did not contain: **inherited state on a copy.** It is
not wiring, not a message, not a missing guard, not file state — it is an attribute of an
object that `Duplicate` carries across and nobody enumerated. And notice it was introduced
*by* step 1: before there was a hidden template there was nothing to inherit hiddenness from.
So the honest reading is narrower than "we designed the failures out". We designed out the
four classes we had already been shown, and the new feature brought a new class with it.

That is worth holding onto, because it predicts the same thing for step 2: a `period` column
means rows a sync must *not* touch, which is a new class of "invisible by design" state — the
same family as a hidden slide and an excluded template, and the family where failures are
silent by construction.

One thing did work as intended, though: `DuplicateAndTag` and `MakeTemplateFrom` both verify
their own identity writes and **refuse, deleting the slide**, rather than trusting
`Tags.Delete`'s behaviour — which was not established on this build. Those guards can
genuinely fail, which is the whole point (zettel
`20260729-an-always-true-guard-is-worse-than-no-guard`). They just guarded the tag and not
its neighbour.

## Process finding — cost one full suite run

**A module-level VBA `Type` below the module's first `Function` reports its error in a
DIFFERENT module.** `MakeTemplateResult` was declared after `PlaceholderFor` in
`TemplateSlide.bas`; the module imported without complaint (the import log said
`Imported TemplateSlide.bas as component name: TemplateSlide`) and the failure surfaced as
`User-defined type not defined` at `TestRunner.bas`, which reads as "the module didn't
import". Cost one ~8-minute run. **Rohan's screenshot of the VBE dialog named the statement
in seconds** — second time in two days that the screen settled an opaque Office failure
faster than any diagnostic. Recorded in `AGENTS.md`, including the cheap static check
(scan each `.bas` for a module-level `Type`/`Const`/`Enum` after the first procedure) and
the reminder to run `Debug > Compile VBAProject` before a suite run that adds a module.

## Honest scope limit

Step 1 does **not** clean the template. It copies P001 once, blanks the five *tagged*
fields, and stops. P001's untagged content is still on the template and still rides onto
every created slide. What changed is that it now sits on **one hidden slide, cleanable by
hand, once** — instead of being invisibly inherited by every slide the tool ever creates.
That manual clean-out is outstanding on the cycle deck; the result message says so.

## Next

Progression step 2: `period` column + deck-declares-its-period (`DECISIONS.md` 2026-07-30,
dated-row model). Step 4 still blocked on the VBA-vs-Office-JS fork — do not start it
without deciding.
