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

**Still unconverted, same defect, not in the item 9/10 path:** `E2EFirstField.bas:129`,
`VerifyRealDeck.bas:36`, `R13RealDeck.bas:45`. All three open the register workbook and will
read `START HERE`. Left alone deliberately — `R13` is parked and the other two are not on the
route to finished. Fix before ever trusting them again. (`BatchOnboardFlow` and the
`R13RealDeck` `gwb` sites index into *different*, single-sheet workbooks and are fine.)

**The lesson, which the codebase already knew:** `E2EField.bas` carries the comment
*"Columns by header name, never by position"* directly beneath a line selecting the **sheet**
by position. The rule was understood one level down and never lifted one level up.

## Not on this list, deliberately

The other 38 fields. The GUID key redesign. R13's full review subsystem (built, parked).
Chrome enforcement. Ribbon polish. None of them stand between here and finished; adding them
here would be the "field count as progress" trap this project already fell into once.
