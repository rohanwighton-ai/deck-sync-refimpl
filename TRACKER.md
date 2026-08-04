# Deck Sync — tracker

**Finished =** Rohan produces a real quarter's deck and the tool saved him time.
Not a field count. Not a test count.

**8 of 10.**

*Said "7 of 10" for two hours after item 8 was ticked. The file that exists to police
stale counts carried a stale count — caught by the PM agent, not by anyone editing it.*

---

- [x] **1. A field syncs end to end** — register → gate → slide → verified by re-reading.
      *Five fields proven: PROJECT_STATUS, ABOUT_BODY, KEY_EVENTS_BODY, PROJECT_NAME, PROJECT_CODE.*

- [x] **2. Nothing reaches a slide unseen** — every change shown as before-and-after, nothing
      written without approval. *Caught 22 slides of prose about to be corrupted.*
      *2026-08-01: this was ticked against a display that printed `slideLen=376 regLen=21
      firstDiffAt=1` and no text at all — technically "shown", not readably shown. Rohan on
      seeing it: "fix it now. It has to be simple and obvious to use." Now prints the full
      untruncated text of both sides, names invisible characters in words, and keeps the
      character codes for the one case that needs them. The tick stands; it did not before.*

- [x] **3. A drafting sheet exists, and says what to do on it** — instructions on the sheet,
      exemplar beside the input, one tick column.

- [x] **4. Drafted text reaches a slide** — written in the sheet, ticked, published, applied,
      verified. *The loop closes.*

- [x] **5. Several rows apply at once** — 3 in one pass, each to its own slide, confirmed by an
      independent harvest.

- [x] **6. A quarter rolls forward** — deck declares its own period; at FY26Q4 it cannot see
      FY27Q1 rows at all.

- [x] **7. Deck settings survive being written** — 5 consecutive differing updates confirmed
      on disk. *The write itself is still unreliable (SaveAs 4/5, Save far worse); the
      operation is made reliable by `set_deck_period.py` — write, verify offline, retry,
      fail loudly. Never verify in-process: it shares PowerPoint's cache with the writer.*

- [x] **8. Rohan says the sheet is usable** — or says exactly why it isn't.
      *Answered 2026-08-01 ~04:55, on the sheet built at 04:18. All four Step 1 checkboxes
      ticked: the job is clear at a glance, column C readable without widening, F/G obvious,
      and yes he would work down all 43 rows. His words: **"not perfect but ok for test."***
      *"Not perfect" is unspecified and NOT yet a defect list — if item 9 turns up what
      bothered him, capture it then rather than guessing at it now.*

- [ ] **9. Rohan drafts a real quarter's content** — text he actually needed written, not test
      edits. *Done when: 10 projects drafted, ticked, applied on the e2e copy.*

- [ ] **10. One real quarter produced, and it saved time** — on a copy of the live deck, which
      then becomes the deck. *Done when: Rohan says the sentence. "No" is a valid answer and a
      spec for what to fix.*

---

## The first real run — 2026-08-01

Rohan took the add-in to his **work machine** and started onboarding his **real**
deck. Everything before this was the redacted deck on the personal machine.

**The existential risk is answered: the add-in loads there.** An employer
blocking unsigned VBA add-ins would have killed this design outright, with no
workaround inside the architecture. It does not. Policy allows trusted add-ins
and blocks macro-enabled *documents* — which this design already respects, since
decks carry tags and document properties, workbooks carry sheets, and all code
lives in the `.ppam`.

**Thirteen findings in about two hours of use. Nine fixed the same day.** Full
account in `FIRST-REAL-RUN.md`. Four `.ppam` builds (33→37).

The most expensive, and the one worth carrying: **the marking session identified
shapes by NAME.** Measured on the real deck — 158 shapes on slide 1 including
nested, **47 sharing a name, zero sharing an `Id`**. Four currency fields all
recorded against "Shape 16" restored onto one shape. The marking was correct when
made and destroyed at *serialise* time, which also made it unrepairable. An hour
of his work, gone.

**And the suite was not compiling.** `run_vba_tests.ps1` never imported
`Sources.bas`, so "135 tests pass" had not been true all day while work was
reported as verified. `build_ppam.ps1` had never imported `ReviewQueue.bas`,
which meant every shipped add-in failed to compile — silently breaking Sync Now,
Review Changes and Apply Approved. Three hand-maintained module lists, nothing
checking them. `vba/tools/check_module_lists.py` now does. **Run it before
trusting a build.**

Suite is green: 135 passed, 0 failed.

**Where item 9 stands: still zero.** Nothing about the real FY26Q4 content has
moved. The tool is markedly more trustworthy than it was that morning and that
is not the same thing.

## 2026-08-03 — the wide model stopped being theory

`register-wide.xlsx` on the rig, read back through `ExcelOutput.ReadSheetForPeriod`
in real Excel. 220 long rows became **48 wide rows**:

| Period | Rows read |
|---|---|
| (rows on sheet) | 48 |
| FY26Q4 | **43** |
| FY27Q1 | **5** |

`3_P001` appears in both, `PROJECT_STATUS` reading `Project Closed` at FY26Q4 and
`Not started` at FY27Q1, while `PROJECT_CODE` / `PROJECT_NAME` / `ABOUT_BODY` are
identical in both — the `Quarter = ALL` sentinel replaced by copying the value onto
each period row, working, on real data. `KEY_EVENTS_BODY` is absent at FY27Q1,
which is correct: next quarter's events are not written yet.

**Two periods were read because one proves nothing.** A single FY26Q4 read returning
43 is exactly what a broken filter returns too. The first version of the check
compared the filtered read against the unfiltered one and **reported failure on a
correct read** — the unfiltered read collapses one project's two periods onto one
instance, so both came back 43. The usable discriminator is the row count taken off
the sheet itself, which no reader bug can move: 43 + 0 duplicates against 48 rows.

Original `register.xlsx` untouched — `migrate_register_to_wide.py` refuses to write
its own input, and the migration went to a new file.

**Item 9 has NOT moved.** This makes the sheet real; it is not a quarter of content.

## The sync path now reads the deck's period — compiled, not yet exercised

`ExcelOutput.ReadSheetForDeckPeriod(ws, deckPeriod, problem)` replaces the unfiltered
`ReadSheet` at all four sync-side reads in `RibbonUI` (Sync Now's queue build and its
apply, Review Changes, Apply Approved). It refuses two ways, and **each refusal was
watched failing before it was trusted** — the filter was broken on purpose, then the
guards were, and the suite caught both:

- two rows for one slide in the read → whichever sat higher used to win silently
- the sheet has rows and none carry the deck's period → zero rows is a legal state
  that reports as a clean sync of nothing

Sync Now stops the **whole run** on a refusal rather than syncing the readable types;
Review Changes and Apply Approved report and skip the type. 145 VBA tests pass.

**Not yet run once.** No test executes `RibbonUI`, and `run_vba_tests.ps1` does not
compile — it never has. The project compiles clean through `field_e2e.ps1` and the
wide read still works afterwards (a compile failure makes PowerPoint deaf to COM, so
that is real evidence), but the four changed call sites have not been executed. The
add-in was deliberately NOT rebuilt: shipping a build whose changed paths have never
run once is the "don't infer a link works from the two ends looking consistent" trap
this project already wrote down.

`RibbonUI.bas:761` still reads unfiltered. That is `New Period` → `RunPeriodRollover`,
the deck-accumulates model Rohan rejected on 2026-08-02. It is on the removal list,
not the fix list.

## What is now knowingly half-wired

`CreateSheet` still does not write the `Quarter` header, and this is the reason the
one-line change was not made: `UpsertRow` does not write a period either. A sheet
with the header and blank period cells reads as **zero rows** under a filtered read,
and an empty read is a legal state that reads as success — the failure that has
already cost this project two evenings. The header and the write have to land
together, with the period coming from the deck that is being onboarded.

## The rule for this file

**Only tick something when it is observably true**, not when the code for it exists. Six of
these were ticked by watching a slide change, not by a test passing.

If an item can't be checked by looking at something, it's written wrong — rewrite it.

## Broken right now — `FieldSpec` wiring (was filed as "unverified")

**Resolved 2026-08-01 05:00: the wiring was not unverified, it was broken. It has never
executed once.**

`Drafting.WriteDraftingSheet` holds `guidance` as `Variant` (it is `Optional`) and passed it
straight into `FieldSpec.LookupGuidance(ws As Object, ...)`, which is **ByRef**. VBA will not
coerce `Variant -> Object` across a ByRef boundary. `Compile error: ByRef argument type
mismatch`. The whole VBA project therefore failed to compile, the modal error box made
PowerPoint deaf to COM, and every driver run died on `RPC_E_CALL_REJECTED`.

**The 135 unit tests could not have caught this.** They call `LookupGuidance` directly with
an already-typed `Object`. The defect exists only at the single cross-module call site — the
exact seam a unit test does not cross. *Passing tests were the reason this looked safe.*

Two things the previous version of this note got wrong, worth remembering:
- **"no dialog"** — a window scan run *between* driver attempts came back clean, because the
  dialog only exists while the macro is being invoked. Absence of evidence, sampled at the
  wrong moment.
- **"the environment, after several hundred Office launches"** — blamed the machine for a
  one-line type error in our own source. The Office-is-flaky prior was available and wrong.

Both were settled in seconds by Rohan screenshotting the actual error box, after four tool
calls of remote diagnosis had produced a wrong theory.

**FIXED AND WATCHED WORKING, 2026-08-01 05:20.** `Drafting.bas:202` assigns to a typed
`Object` local first, leaving `FieldSpec`'s public signature and its 135 tests untouched.
`compile executed`, 43 rows written, and `I1` carries **no** `--  GENERIC, no Field Spec row`
suffix — read back off the closed file, not from the console. The `FieldSpec` wiring has now
executed. `I2` is not a discriminator and was never treated as one: a plausible field-specific
prompt is produced either way. The `I1` suffix is.

---

## The bug the compile error was hiding — register found by tab position

Fixing the compile bought a *worse* symptom: `0 row(s) written`, down from 43, reported as a
clean run.

`E2EField.bas` read the register as `wb.Worksheets(1)` in five places.
`WorkbookBridge.WriteWorkbookIndex` ends with `ws.Move Before:=wb.Worksheets(1)` — the
`START HERE` sheet puts itself at the front on purpose. So every register read had been
returning the **instructions tab**. No matching columns, no rows, no error, because an empty
register is a legal state.

It survived because the 04:18 run wrote its 43 rows *before* the index sheet was inserted,
and the `FieldSpec` compile error then blocked every run that would have exposed it. Two
defects, each hiding the other.

Fixed via `WorkbookBridge.RegisterSheet(wb)` — by name, and it **raises** on a missing
register rather than returning `Nothing`, because "reportable as zero rows" is what let this
live. All five `E2EField.bas` sites converted; 43 rows confirmed back.

**~~Still unconverted~~ — CONVERTED the same day, and this paragraph was stale for a day
after that.** `E2EFirstField.bas:129`, `VerifyRealDeck.bas:36` and `R13RealDeck.bas:45` all
call `RegisterOrFirstDataSheet` now. Caught by consultant review 2026-08-01, which noted the
irony precisely: *the file that exists to stop stale claims was carrying one, one day after
it carried a stale count.* (`BatchOnboardFlow` and the `R13RealDeck` `gwb` sites index into
*different*, single-sheet workbooks and are genuinely fine.)

**The lesson, which the codebase already knew:** `E2EField.bas` carries the comment
*"Columns by header name, never by position"* directly beneath a line selecting the **sheet**
by position. The rule was understood one level down and never lifted one level up.

## 2026-08-04 — the copy is fully repointed to the wide model, and "q" is retired

The PM checklist for "the sync path actually works" (see the deck-sync-pm agent brief):

1. Drafting reads the wide sheet — still open, `DraftingUI.bas:231` reads long
2. Period-aware `UpsertRow` — still open, the keystone
3. Publish writes to the wide sheet — still open
4. `CreateSheet` stamps `Quarter` — still open, depends on (2)
5. Long register / Seed-Draft-Approved / `RunPeriodRollover` retired — still open

None of those moved today. What did: a copy of the rig deck
(`e2e-deck.wide-test.pptx` — the original `e2e-deck.pptx` untouched) is now fully wired to
the wide model, closing the setup gap that was blocking Sync Now from ever being tried on it:

- **`ExcelOutput.ReadSheetForDeckPeriod` closed an empty-sheet hole.** The guard added
  2026-08-03 refused a sheet with rows that didn't match the period, but a freshly-created
  **empty** sheet (the exact shape `GetOrAddWorksheet` produces when a worksheet name
  resolves wrong) read as zero rows either way and passed as a clean sync of nothing. Now
  checked directly (`IsEmpty(ws.Cells(1,1))`) before any read is attempted. Made to fail on
  purpose first — disabling the guard broke exactly the new test and nothing else, 145/146.
- **The copy's workbook pairing was pointed at `register-wide.xlsx`**, off
  `SAAFE-Projects-Data.xlsx` — a real file in Rohan's OneDrive the deck was still paired to.
  `set_deck_workbook_path.py` (new, mirrors `set_deck_period.py`'s write/verify-offline/retry
  shape) confirmed the property on disk before calling it done.
- **The slide type "q" was renamed to `project-status`, atomically.** Three things were doing
  one job under that one string — the deck's type registration, its worksheet pointer
  (wrongly `"q"`, should be `"Register"`), and a hardcoded literal in `DraftingUI.bas:231`
  filtering the long register's `SlideType` column. Renamed the first two together (a
  worksheet-only repoint would have left every slide's tag saying `"q"` while the
  registration said something else — exactly the shape of the live incident
  `BatchOnboardFlow.bas:937` already records, where a retype without a matching worksheet
  fix stranded an entire Data sheet). `rename_slide_type.py` (new) retags all 43 slides,
  re-registers, deletes the stale property, and verifies FOUR things independently against
  the file's own bytes (new registration, old registration gone, zero straggler tags, new
  tags present) before reporting success. The third "q" — the `DraftingUI.bas` literal —
  was deliberately NOT touched; it lives inside checklist item 1 above and gets fixed once,
  there, not patched twice.
- **Real taxonomy, from Rohan directly:** the deck's slides are Output (program level),
  Milestone, Project Status, Project Progress. Only `project-status` is registered anywhere
  today. **Open, handed to Opus:** whether decks close in style but distinct in program
  (research / kickstart / student) should be the same slide type or different ones sharing
  one register workbook — the spec history (`Round8.md` §6) already argues for one workbook,
  many decks, with `SlideType` as a column, but that reasoning predates the wide-sheet pivot
  and `migrate_register_to_wide.py` currently REFUSES mixed slide types in one sheet. Also
  open: whether onboarding should offer a dropdown of slide "prototypes" with a human suffix,
  once the above is settled — UX for a decision that hasn't been made yet.

Sync Now itself is still **not exercised** — that needs a human at the keyboard (18 `MsgBox`
calls in the sync paths mean a headless run hangs). The copy is now in the state where that
run is actually meaningful to attempt.

## Not on this list, deliberately

The other 38 fields. The GUID key redesign. R13's full review subsystem (built, parked).
Chrome enforcement. Ribbon polish. None of them stand between here and finished; adding them
here would be the "field count as progress" trap this project already fell into once.
