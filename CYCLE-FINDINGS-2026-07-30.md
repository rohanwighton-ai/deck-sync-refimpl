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

## Still unexercised

`New Period` -- also not on the toolbar.
