# What this system is — for someone with zero context

> **CURRENT.** Compiled 2026-08-16 from a documentation sweep, migrating the still-true
> parts of `HANDOVER-Q4F26-DRAFTING.md` (2026-08-10, now archived — see
> `archive/HANDOVER-Q4F26-DRAFTING.md`) before that file was archived. Provenance is noted
> per section below, per Rohan's instruction: *"when you find useful items, build them
> into current documentation with their provenance noted."* Column letters and layout
> below are re-verified against `Drafting.bas`'s constants directly, 2026-08-16 — not
> copied forward from the source document, which had gone stale on exactly this point.

## In one paragraph

*Source: `HANDOVER-Q4F26-DRAFTING.md` §1, migrated verbatim — still accurate.*

Rohan produces a quarterly PowerPoint deck of ~43 project-status slides for SAAFE CRC.
Each slide has the same panels (About, Problem, Key Events, Strategic Alignment,
Progress, a milestone timeline). Doing it by hand costs about **three weeks of evenings
per quarter**, and the dominant cost is not typing — it is **re-deciding what each panel
is for and what belongs in it**, every quarter, for every project.

The system has two halves:

- **The drafting half (Excel)** — a workbook holding a **register** of every slide's
  content, a **recipe** per field saying what that field is for and how to write it, and
  a **sources** list saying what the writing is based on.
- **The delivery half (a PowerPoint add-in)** — pushes approved text onto the slides.

**The drafting half is the valuable half.** The recipes are what remove the re-deciding;
the add-in is delivery. See `TRACKER.md`'s baseline framing for why this governs the
order of work.

## The files

*Source: `HANDOVER-Q4F26-DRAFTING.md` §2, still accurate.*

| file | what it is |
|---|---|
| `register-wide.xlsx` | The workbook. Holds everything below. **One file, forever** — rows accumulate across quarters. *(This is decision 2 in `DOCUMENT-MAP.md` — one file pair per quarter is ratified but NOT YET BUILT; today it is genuinely one file forever, not one per quarter.)* |
| the deck `.pptx` | One file, currently reused across quarters (the file-per-quarter design is not yet built either — see `CHECKLIST.md`). |

### Sheets in the workbook

| sheet | what it holds | lifetime |
|---|---|---|
| `START HERE` | A status report the add-in writes. | rebuilt |
| `Field Spec` | **The recipes.** One row per field. | permanent, hand-edited |
| `Sources` | What the writing is based on. One row per source, ever. | permanent, accumulates |
| `Register` | The content itself. One row per slide per quarter. | permanent, accumulates |
| `TPL_<FIELD>` | A **drafting sheet** per field — the workspace. | rebuilt in place each run |
| `Run Log` | What the last run did. | **replaced each run** — see `SCENARIOS.md`'s file-per-quarter section |
| `Sync Log` | Durable record of every applied change. | **append-forever** — flagged in `CHECKLIST.md` as needing to fold into the file-per-quarter archive once that's built |

## The drafting sheet — column layout, RE-VERIFIED 2026-08-16

*Source: re-derived from `Drafting.bas`'s `COL_D_*` constants directly — the handover
document's own table (F=SUBMIT, G=APPROVE) was already one column off, from before the
`REPORTED LAST TIME` column was inserted. Per `DOCUMENT-MAP.md`'s own rule: this table
is provided for orientation only. If it ever disagrees with `Drafting.bas`, the code
wins — don't let this table calcify the same way the one it replaces did.*

| col | constant | what goes in it |
|---|---|---|
| A | `COL_D_ENTITY` | Project code — derived, do not edit |
| B | `COL_D_NAME` | Project name — derived, do not edit |
| C | `COL_D_CURRENT` | **ORIGINAL** — what the slide says now. Read this first. |
| D | `COL_D_PREV` | **REPORTED LAST TIME** — last quarter's text, for voice/narrative consistency only, never storage (see `DOCUMENT-MAP.md` decision 6) |
| E | `COL_D_SOURCES` | source IDs this row was drafted from, e.g. `S10, S12` |
| F | `COL_D_DRAFT` | **AI DRAFT** — never published |
| G | `COL_D_SUBMIT` | **your words — this is what publishes** |
| H | `COL_D_APPROVED` | the tick |
| K | `COL_D_NOTES` | notes back to the tool |
| L2 | `COL_D_PROMPT` | the generated prompt for this field |

**A row publishes only when BOTH `SUBMIT` is non-empty AND `APPROVE` is `Y`.** Either
alone is not consent.

## The recipes — `Field Spec` sheet

*Source: `HANDOVER-Q4F26-DRAFTING.md` §5, still accurate as design intent — the sheet
itself is the authoritative list of columns per `COLUMNS.md`'s own rule, this is only
the reasoning.*

One row per field: `FieldID`, `Kind` (`Prose` / `Controlled` / `Static`), `Purpose`,
`Voice`, `Length`, `Own-job test`, `Do NOT`, `Allowed values` (Controlled only),
`GLOBAL RULES`, `Behaviour`, `Renders as`.

**The recipes are the product.** If a recipe is good, drafting a field stops being a
decision and becomes a transcription. See `project_deck_sync_recipes_are_the_product`
memory and `TRACKER.md`'s item 9 for why this bet already paid off.

## Vocabulary

*Source: `HANDOVER-Q4F26-DRAFTING.md` §10, migrated whole — no current doc holds a
compact glossary and it is genuinely useful.*

- **Field** — a register column whose value ends up on a slide.
- **Instance** — one slide, identified by `Instance ID`.
- **Period / Quarter** — `Q4F26` format. Free text, matched exactly — see
  `reference_saafe_period_convention` memory.
- **Publish** — copy approved drafting-sheet text into the register. Touches no slide.
- **Sync** — write register values onto slides.
- **Recipe** — a `Field Spec` row: what a field is for and how to write it.
- **SPOT source** — authoritative *and* maintained; has an owner and a refresh cadence.
  Not tidiness, not origin — see the full test below.
- **Provenance** — which sources and which recipe produced a given value. Design in
  `PROVENANCE.md`; build steps now on `CHECKLIST.md`.

## What NOT to do — standing rules, not previously written into any CURRENT doc

*Source: `HANDOVER-Q4F26-DRAFTING.md` §8, migrated whole. Checked 2026-08-16 that
`AGENTS.md` does not already cover these — it doesn't; this was a genuine gap.*

- **Do not rename the `Quarter` column** — matched by that literal string.
- **Do not change `Instance ID` values** — the permanent key linking a register row to a
  slide.
- **Do not reorder or rename field columns casually** — field names are matched
  **exactly and case-sensitively** against the slide tag. `PROJECT_PROGRESS` and
  `Project_Progress` are different fields and fail to match silently.
- **Do not save the register as `.xlsm`** — work restrictions prohibit macro-enabled
  workbooks; it stays `.xlsx` forever. All behaviour lives in the add-in.
- **Do not type into a hidden sheet named `SAVED …`** — those are parked archives of
  previous drafting sheets. The tool never reads them.
- **Periods are matched exactly.** `Q4F26` — not `Q4 F26`, not `q4f26`, not `FY26Q4`. A
  mistyped period reads as a clean run of zero rows, not an error.

## The SPOT source test

*Source: `HANDOVER-Q4F26-DRAFTING.md` §6, migrated whole — a real, still-applicable
definition with no other current home.*

> **A SPOT source is authoritative AND maintained.** The defining test is that it has an
> **owner** and a **refresh cadence**. Not tidiness, not origin.

Consequences: a document (e.g. a maintained quarterly report) is a legitimate source, not
a weak third grade — it's where new facts *enter* the system. A one-off spreadsheet
someone made once and never updated is **not** a SPOT source however authoritative it
looked at the time. The register itself is refined truth built FROM sources — always
derived, never exempt from provenance.

**Open, standing question, also on `CHECKLIST.md`:** who owns the declared programme
linkage codes now the Family Tree is being retired.
