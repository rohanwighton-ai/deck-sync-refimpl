# Q4F26 drafting — context handover

**Written 2026-08-10 for work done on another machine, with a different AI assistant.**
Assume the reader knows nothing about this project. Everything needed is below.

---

## 1. What this system is, in one paragraph

Rohan produces a quarterly PowerPoint deck of ~43 project-status slides for SAAFE
CRC. Each slide has the same panels (About, Problem, Key Events, Strategic
Alignment, Progress, a milestone timeline). Doing it by hand costs about **three
weeks of evenings per quarter**, and the dominant cost is not typing — it is
**re-deciding what each panel is for and what belongs in it**, every quarter,
for every project.

The system being built has two halves:

- **The drafting half (Excel)** — a workbook holding a **register** of every
  slide's content, a **recipe** per field saying what that field is for and how
  to write it, and a **sources** list saying what the writing is based on.
- **The delivery half (a PowerPoint add-in)** — pushes approved text onto the
  slides.

**The drafting half is the valuable half.** The recipes are what remove the
re-deciding; the add-in is just delivery. Tomorrow's work is entirely on the
drafting half, and it does not need the add-in to be working.

---

## 2. The files

| file | what it is |
|---|---|
| `register-wide.xlsx` | The workbook. Holds everything below. **One file, forever** — rows accumulate across quarters. |
| the deck `.pptx` | One **new file per quarter**. Last quarter's stays as the record. Not needed tomorrow. |

### Sheets in the workbook

| sheet | what it holds | lifetime |
|---|---|---|
| `START HERE` | A status report the add-in writes. Ignore it when working manually. | rebuilt |
| `Field Spec` | **The recipes.** One row per field. | permanent, hand-edited |
| `Sources` | What the writing is based on. One row per source, ever. | permanent, accumulates |
| `Register` | The content itself. One row per slide per quarter. | permanent, accumulates |
| `TPL_<FIELD>` | A **drafting sheet** per field — the workspace. | rebuilt each run |
| `Run Log`, `Sync Log` | Reports from the add-in. | append |

---

## 3. The register — the record

**One row per slide per quarter.** Columns:

| col | meaning |
|---|---|
| A | `Instance ID` — the slide's permanent key, e.g. `3_P001` |
| B | `Quarter` — e.g. `Q4F26`. **Quarter first, `F`, two-digit FY.** Not `FY26Q4`. |
| C onward | one column per field: `PROJECT_CODE`, `PROJECT_NAME`, `PROJECT_STATUS`, `ABOUT_BODY`, `KEY_EVENTS_BODY`, `STRATEGIC_ALIGNMENT_BODY`, `PROBLEM_BODY`, `PROGRESS_BODY` |

Rows for different quarters sit side by side — `Q3F26`, `Q4F26` and `Q1F27` rows
all live in the same sheet. Nothing is overwritten between quarters.

**Line breaks inside a cell are written as `||`.** Excel cells are single-line;
the add-in converts `||` to real line breaks when it writes to a slide.
`KEY_EVENTS_BODY` uses this heavily (one event per line).

---

## 4. The drafting sheet — the workspace

One per field, named `TPL_ABOUT_BODY`, `TPL_STRATEGIC_ALIGNMENT_BODY` etc.
Rebuilt from the register each time the add-in runs, so **it is a workspace, not
a record**. Column layout ("layout 4"), left to right in workflow order:

| col | name | what goes in it |
|---|---|---|
| A | Project code | derived — do not edit |
| B | Project name | derived — do not edit |
| C | **ORIGINAL** | what the slide says now, from the register. Read-only reference. |
| D | **SOURCES** | the source IDs this row was drafted from, e.g. `S10, S12` |
| E | **AI DRAFT** | where an AI's suggestion is pasted |
| F | **SUBMIT** | **the wording that will be published.** This is the one that counts. |
| G | **APPROVE** | type `Y` to approve the row |
| J | NOTES | free text, questions, `[TBC]` markers |
| L2 | — | the generated prompt for this field (recipe + rules, ready to paste into an AI) |

**A row publishes only when BOTH `SUBMIT` is non-empty AND `APPROVE` is `Y`.**
Either alone is not consent.

**The manual loop, without the add-in:**

1. Read the recipe in `L2` (or on the `Field Spec` row for that field)
2. Read column C to see what the slide currently says
3. Draft into E, or write straight into F
4. Put the source IDs in D
5. Put `Y` in G
6. Copy the approved F value into the matching register row for `Q4F26`

Step 6 is what the add-in's "publish" does. Doing it by hand is fine and is how
to get started before the add-in is ready — **just be careful to match on
`Instance ID` AND `Quarter`, not row position.**

---

## 5. The recipes — `Field Spec` sheet

One row per field. Columns:

| col | meaning |
|---|---|
| FieldID | the field name, matching the register column |
| Kind | `Prose` / `Controlled` / `Static` — how the content is decided |
| Purpose | **the question this field answers** |
| Voice | tone, tense, register |
| Length | target, advisory |
| Own-job test | "does this do THIS field's job, and not the neighbouring field's?" |
| Do NOT | what to avoid — the anti-patterns |
| Allowed values | for `Controlled` fields only, comma separated |
| GLOBAL RULES | one cell, applied to **every** field's prompt |
| Behaviour | how content is PLACED (pictures/objects) |
| Renders as | Text / Picture / Progress bar |

**The recipes are the product.** If a recipe is good, drafting a field stops
being a decision and becomes a transcription. If you still have to work out what
belongs in the panel, the recipe needs improving — and improving it is more
valuable than drafting faster.

**A known recipe defect to fix:** the global rules say *"Column C is the
standard, not a draft to improve on… if the text in column C already does its
job, say so and leave the row blank."* That was written when every field had a
prior value. For a field added recently, column C is **empty for 42 of 43
projects**, so read literally it tells the AI to leave every row blank. Needs
rewording for the no-prior-value case.

**A structural recipe gap:** a drafting sheet shows only the project's *name*.
`STRATEGIC_ALIGNMENT_BODY` is the "so what" of a project and cannot be written
from a title — it needs `ABOUT_BODY` (what the project *is*) visible alongside.
The intended fix is a `Context fields` column on the Field Spec saying which
other fields to show read-only. Not built.

---

## 6. Sources — and the SPOT source list to build

### The Sources sheet

| col | meaning |
|---|---|
| A | `ID` — `S01`, `S02`… |
| B | What it is |
| C | Type |
| D | Where it lives — path or URL |
| E | Notes |
| F | Added |
| G | **Applies to** — `All periods`, or one specific period |

One row per source, ever. If twenty projects use the same report, they all cite
the same ID. **Point at documents; never paste document text into this sheet.**

### ⚠️ Six rows are FABRICATED and must be replaced

`S01`–`S06` were invented by an AI to exercise the citation mechanism. They are
prefixed `EXAMPLE (not a real document)`. **37 real `TPL_ABOUT_BODY` rows
currently cite them.** Clearing them properly needs BOTH: delete rows `S01`–`S06`
*and* blank the citation column on the rows that point at them — deleting only
the source rows leaves citations pointing at IDs that no longer exist.

`S10`, `S11`, `S12` are synthetic test files with real paths — also not real
program documents.

**So: the Sources sheet currently contains no genuine source. Building the real
list is the highest-value thing that can be done tomorrow alongside drafting.**

### What makes something a SPOT source

The test agreed in this project — **and it is a demanding one**:

> **A SPOT source is authoritative AND maintained.** The defining test is that it
> has an **owner** and a **refresh cadence**. Not tidiness, not origin.

Consequences of that test:

- A **document** is a legitimate source, not a weak third grade. A team's
  quarterly Knack report is where new facts *enter* the system.
- A one-off spreadsheet someone made in 2024 and never updated is **not** a SPOT
  source, however authoritative it looked at the time.
- **The register itself is refined truth built FROM sources** — always derived,
  never exempt from provenance.

### The list to build

Start a table with these columns and fill it as you go. The two on the right are
the SPOT test:

| ID | What it is | Type | Where it lives | Owner | Refresh cadence | Applies to |
|---|---|---|---|---|---|---|
| | | | | | | |

**Candidates already identified in this project, to confirm or discard:**

- **Knack** — the project management system. Quarterly team reports are where new
  project facts enter. Likely the single most important SPOT source.
- **Xero** — financials (cash, in-kind, total project value appear on every slide).
- **The Family Tree** — *being retired*. Rohan: "I can start pulling them out of
  the old family tree." **The open question this raises is unresolved and matters:
  it held the declared programme linkage codes (e.g. `1.4.2`, `2.3.1`), and
  nobody has yet said who maintains that list once it is gone.** A drafted field
  already stalled on exactly this — the evidence rule correctly refused to assert
  linkage codes because nothing on hand could distinguish *declared* from
  *inferred*. Ask at work: **who owns the declared-linkage list now?**
- **The research programme plan** — declares the linkage structure.
- **Milestone register / contract schedules** — the milestone dates and
  achievement status the timeline panel needs.
- **PAC / governance minutes** — decisions and status changes.

### The evidence rule this exists to serve

A drafted field must not assert something no source supports. This already
fired correctly once: a Strategic Alignment draft refused to state linkage codes
and ended `Declared linkages: [TBC]`, with the question recorded in the notes
column. **That is the mechanism working, not failing** — `[TBC]` plus a named
question is the correct output when the evidence is not there.

---

## 7. What to do tomorrow — suggested order

1. **Draft `STRATEGIC_ALIGNMENT_BODY` for a second project.** It exists for
   `2_P004` only, 1 of 43. Use the recipe in `L2`.
   **While doing it, notice the thing that matters most: did the recipe tell you
   what to write, or did you still have to work it out?** That question is the
   one the entire project rests on and it has never been tested. If the recipe
   works, the tool is worth finishing. If you still re-decide from scratch, the
   recipes need the work, not the delivery mechanism.
2. **Improve that recipe based on what you just noticed.** This is the real
   deliverable.
3. **Start the SPOT source list** — even five real rows with owners and cadences
   is more than exists today.
4. Draft more fields if there is time.

---

## 8. What NOT to do — would break the add-in later

- **Do not rename the `Quarter` column** — it is matched by that literal string.
- **Do not change `Instance ID` values.** They are the permanent key linking a
  register row to a slide.
- **Do not reorder or rename field columns casually** — the field name is
  matched **exactly and case-sensitively** against the slide. `PROJECT_PROGRESS`
  and `Project_Progress` are different fields and will silently fail to match.
- **Do not save as `.xlsm`.** Work restrictions prohibit macro-enabled workbooks;
  this workbook must stay `.xlsx` forever. All behaviour lives in the add-in.
- **Do not type into a hidden sheet named `SAVED …`** — those are archives of
  previous drafting sheets. The tool never reads them.
- **Periods are matched exactly.** `Q4F26` — not `Q4 F26`, not `q4f26`, not
  `FY26Q4`. A mistyped period reads as a clean run of zero rows.

---

## 9. Known dirty state in the workbook

- Six fabricated `Sources` rows (§6) cited by 37 real rows.
- One `TESTFILL-1256` placeholder left in `ABOUT_BODY` for `3_P001`.
- Three projects have **no real `ABOUT_BODY` text at all** — `3_P001`, `3_P002`,
  `2_P003` — and it is not recoverable from any backup. They need real text.
- `STRATEGIC_ALIGNMENT_BODY`, `PROBLEM_BODY` and `PROGRESS_BODY` each have a
  value on **1 of 43** rows.

---

## 10. Vocabulary, so the assisting tool uses the right words

- **Field** — a register column whose value ends up on a slide.
- **Instance** — one slide, identified by `Instance ID`.
- **Period / Quarter** — `Q4F26` format.
- **Publish** — copy approved drafting-sheet text into the register. Touches no slide.
- **Sync** — write register values onto slides. Not needed tomorrow.
- **Recipe** — a `Field Spec` row: what a field is for and how to write it.
- **SPOT source** — authoritative *and* maintained; has an owner and a cadence.
- **Provenance** — which sources and which recipe produced a given value.
  Currently NOT preserved across quarters; design written up in `PROVENANCE.md`.
