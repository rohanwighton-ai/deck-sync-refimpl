# The workflow, in English

What a person actually does, what the tool does back, and what it touches — so the
time-and-motion can be argued about on paper instead of discovered in a UI.

Derived from `CommandBarUI.bas` (the toolbar IS the workflow) and the handlers it
calls. Where a step does not work today, it says so.

**Two words that matter before anything else:**

- **Slide type** — a *template you would clone*. Not a category. "Research Project
  Status" and "Kickstart Project Status" are different types if they differ in colour
  or content, because you would not clone one to make the other.
- **Period** — which quarter this deck renders. The register holds every quarter side
  by side; the deck declares which one it is.

---

## SETUP — once per slide type, not per quarter

| | You do | The tool does | Touches |
|---|---|---|---|
| **A. Mark Fields** | Click a field's shape on the template slide, run it. Repeat per field. | Tags the shape. Text shapes only. | The deck (tags only, no text) |
| **A2. Discover Fields** | *Alternative to A.* Tick and name every field in one Excel grid. | Lists every text shape in reading order. Marks nothing until you confirm. | Nothing until you confirm |
| **B. Onboard Slides** | Review the batch in Excel, confirm. | Finds the other slides of that layout, links the batch, writes a row per slide, offers to make the hidden template slide. | Deck (tags) + workbook (rows) |
| **C. Check Coverage** | Run it whenever you wonder what you missed. | Lists everything on the slide NOT tracked, ranked by how likely it is to be data. | Nothing — read-only |
| **Clear Marks** | Start marking over. | Discards every mark. Cannot remove just one. | Marks only |

**Effort:** A2 replaced three dialogs per field with one grid — that was the single
biggest time saving in the project so far. B is where the 43 slides get linked.

**ORDERING PROBLEM, found 2026-08-04 and not yet fixed in the UI.** Onboarding now
stamps every row it writes with the deck's period, and refuses a blank one. So on a
fresh deck **step 0 must run before Setup B** — the toolbar's numbering implies
otherwise. Today that surfaces as a raw error rather than "set your quarter first".

---

## THE QUARTERLY LOOP — every quarter

### 0. Start a Quarter
**You:** type the period (e.g. `FY27Q1`) into a box, confirm.
**Tool:** writes it to the deck, reads it back to verify (deck-property writes are
documented-unreliable here), tells you to save.
**Touches:** the deck's period property. No slide.
**Status:** works. **Its closing advice is stale** — it tells you to hand-copy last
quarter's rows in Excel and set Status to Seed. `ExcelOutput.RollForwardPeriod` now
does that copy, and Status is being retired. The message needs rewriting and
roll-forward needs a button.

### 1. Drafting Sheets
**You:** click it.
**Tool:** builds one sheet per prose field — every project on a row, current text in
column C, an empty box for your new wording in G, Copilot's prompt in L2.
**Touches:** the workbook only. Never the deck.
**Keeps your work:** a rebuild carries across the AI draft, your SUBMIT text, source
IDs and notes. Only ORIGINAL and character counts are re-derived.
**Status:** reads the **WIDE** sheet as of 2026-08-05, through the same guarded reader
Sync Now uses. **Exercised in real Excel 2026-08-05: 43 rows written.**
**Changed with it:** a rollover no longer carries drafting work across. That protection
moved rather than vanishing — `RollForwardPeriod` copies last period's rows, so the
previous text arrives in the ORIGINAL column instead of in your draft box.

### 2. Copy AI to Submit *(optional)*
**You:** click it, then edit column G and tick column I.
**Tool:** copies AI drafts into SUBMIT, filling **only** cells you left empty. Never
overwrites your own words.
**Touches:** the workbook only.
**Status:** works.

### 3. Publish & Preview
**You:** read the list of ticked rows, say go.
**Tool:** writes them into the register for the deck's period, then offers to preview
what that would change on the slides.
**Touches:** the workbook. No slide by itself.
**Status:** writes the **WIDE** sheet as of 2026-08-05, via period-aware `UpsertRow`.
**Exercised in real Excel 2026-08-05:** 2 rows published, 1 refused by name for having
no row in the period. Verified by reading the saved `.xlsx` bytes, not from the report.
**Two things went with the move:** there is no longer a Status column to write
"Approved" into — your tick in column I is the only consent gate, which is what it
always actually was. And publishing into a period where a slide has no row is
**refused**, not created: otherwise publishing would invent slides nobody onboarded.

### 4. Sync Now
**You:** read what it will change, confirm.
**Tool:** applies the register to the slides. If anything needs reading one at a time,
it sends you to the review sheet instead of guessing.
**Touches:** the deck.
**Status:** reads the **WIDE** sheet. **RUN 2026-08-05, dry run then apply, on
`e2e-deck.wide-test.pptx`:** 43 slides, ABOUT_BODY found on all 43, 5 changed, `written
and verified: 5`, `slides matching the register: 43`, `mismatched: 0`, deck saved.
Confirmed independently by unzipping the saved deck and reading its slide XML.

> **THE LOOP CLOSED 2026-08-05, AND IT HAS BEEN RUN.** Steps 1, 3 and 4 all address the
> wide sheet, and the whole loop went end to end on the 43-slide rig copy: drafting
> sheet -> publish -> register -> sync -> deck, **43 slides matching / 0 mismatched**.
> Both ends verified by reading the saved `.pptx` and `.xlsx` BYTES rather than by
> trusting either tool's own report.
>
> That distinction matters here more than usual. "Green in source" has been claimed and
> later found false twice on this project (2026-08-01, 2026-08-04), so the claim above
> is deliberately stated as what was observed out of process, not as what the code says.
>
> Also fixed the same day: `ExcelOutput.ManualSmokeTest` still called `UpsertRow` with
> three arguments after the period became required — *"Argument not optional"*, a
> **compile** error, which stops the whole VBA project rather than just that Sub. The
> suite could not see it, and that was PROVEN by reintroducing the error: it still
> printed "152 passed, 0 failed", exit 0. **VBA compiles per procedure, on demand**, so
> a Sub nothing calls is never compiled. `vba/tests/compile_check.ps1` now gates the
> suite and was made to fail before being trusted.

---

## THE CAREFUL ROUTE — when step 4 is too blunt

| | What it is |
|---|---|
| **Preview Sync** | Everything the register would change. Reads only, writes nothing — the safest thing on the toolbar, and the right first action on an unfamiliar machine. |
| **Review Changes** | Builds a "Sync Review" sheet, current vs proposed per slide. Writes nothing to the deck. |
| **Apply Approved** | Writes what you ticked. Takes a backup, re-checks each change against the slide, skips anything that moved since you approved it. |
| **Review + Approve All** | Scratch copies only. Ticks everything without individual review. Still writes nothing until Apply Approved. |
| **Repoint Workbook** | Only if deck and workbook got separated. Keep them in one folder and the pairing repairs itself. |

---

## Where the human time actually goes

1. **Writing the words.** Steps 1–3. Everything else is minutes; this is the evening.
   43 rows for one field.
2. **Marking fields, once per slide type.** Setup A2 + B.
3. **Everything else is confirmation clicks** — and every one of them exists because
   something was once written to a slide that shouldn't have been.

The tool's claim is not that it writes the words. It is that nothing reaches a slide
unseen, last quarter's text is sitting in front of you while you write this
quarter's, and you never retype a project name.

---

## Sources — how provenance lines up

`Sources` is a **bibliography, not a data store**: ID, what it is, type, where it
lives, notes, added. Its own rule — *"Do not paste document text in here; point at
it."*

- A contract's description block is **one row, cited forever** — it feeds a field that
  does not change between quarters.
- Each quarter's progress report is **its own row**, because it is genuinely a
  different document.

**Nothing ties a source to a period.** Publish checks a cited ID *exists*, not that it
is from the right quarter. Convention covers it for now: **put the period in the
label** (`P001 Progress Report — FY26Q4`) for period-bound sources, leave it off
permanent ones.

**Volume caveat:** the sheet's defence against sprawl is "one report cited by 43
projects is one row." That only helps for *shared* documents. Per-project quarterly
reports give one row per project per quarter — 14 × 4 = **56 rows a year**, and the
saving does not apply.
