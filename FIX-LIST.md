# Fix list

One place for what is known-broken and not yet fixed, so each new review stops
re-deriving the same findings. Three reviews have now paid to rediscover items that
were already known — that cost is what this file exists to stop.

Ranked by how much real work is destroyed, or wasted, before anyone notices.

---

## 1. Excel's real error is thrown away, and the message sends you to the wrong file

**Found 2026-08-09, live, mid-session.** `1. Drafting Sheets` reported:

> Could not open the paired workbook at: `C:\Users\rohan\deck-sync-e2e\register-wide.xlsx`

The file was perfectly good — Excel opened it directly, 13 sheets. The actual cause was
that a **different** `register-wide.xlsx` was already open from another folder, and
Excel refuses two workbooks with the same filename at once. Nothing in the message
said so, so the first five minutes went into "did we corrupt the file?"

`WorkbookBridge.OpenOrGetWorkbook` wraps the open in `On Error Resume Next` and returns
`Nothing`. Every caller can then only say "could not open". Excel supplies a specific,
actionable reason — name clash, lock, permissions, cloud path — and all of it is
discarded at the moment it is generated.

**Fix:** capture `Err.Description` from the failed `Workbooks.Open` and return it, so
callers report *why*. Same shape as the publish save that reported `Err.Number = 0` as
success: the diagnosis existed and was thrown away.

**Cost:** small. One out-parameter, five call sites.

---

## 1b. A drafting sheet for a derived field carries nothing to derive from

**Found 2026-08-09, on the first real drafting run of a new field.** Two faults, one
cheap and one structural.

**The cheap one — the global rules assume a prior value exists.** They say *"Column C
is the standard, not a draft to improve on… if the text in column C already does its
job, say so and leave the row blank."* On a field added today, column C is empty for
every project that has not been harvested — 42 of 43. Read literally, that clause tells
Copilot to leave every row blank. Those rules were written when every field had a prior
value, which was true of `ABOUT_BODY` and is false of every field added from now on.
It is a cell on the Field Spec sheet, so the fix is an edit, not a build.

**The structural one — the sheet shows a drafter only the project's NAME.** The columns
are: code, name, ORIGINAL (empty for a new field), SUBMIT, tick, AI draft, sources,
counts, notes. `STRATEGIC_ALIGNMENT_BODY` is the "so what" for a project, and it cannot
be written from a title. `ABOUT_BODY` — what the project actually IS — exists for all
43 projects, in the register and on its own sheet, and is not visible here.

So for a derived field the honest output is 42 blank rows, and the rules correctly
forbid inventing the rest. That is not the recipe failing; it is the sheet not carrying
what the recipe needs.

**Fix, and it is Field-Spec-shaped rather than code-shaped:** a `Context fields` column
on the Field Spec saying which other fields to show read-only alongside — Strategic
Alignment and Problem both want `ABOUT_BODY`; Progress wants `KEY_EVENTS_BODY`. Then
`WriteDraftingSheet` renders those as extra read-only columns.

**Why this matters more than it looks:** it is the difference between a tool that can
only UPDATE text that already exists and one that can WRITE a field for the first time.
Every new field, and every new project, hits this.

---

## 2. `4. Sync Now` always refuses on the main field, and names a sheet that does not exist

`AssignBatches` batches only `KIND_CONTROLLED` fields, so for prose —
`ABOUT_BODY`, the field the whole drafting apparatus exists for — `HasBatchableWork` is
never true. The button numbered 4 in a 0–4 sequence therefore **always** ends in a
`vbExclamation` dialog saying it is opening the review sheet instead. The toolbar
presents that as an exception; the code makes it the rule.

And the dialog names `REVIEW_SHEET_NAME` = `"Sync Review"`, while `ReviewSheetNameFor`
produces `Review project-status-3D1B`. The rename's own comment records why it happened:
*"Rohan could not find it and asked where the 'sync review file' was."* The sheet got a
findable name; the sentence pointing at it did not.

**Fix:** when there is nothing batchable *because everything is prose*, that is the
normal route — drop the exclamation, go straight to the review sheet, and name the real
sheet via `ReviewSheetNameFor`. Delete the constant.

---

## 3. The double tick — 86 ticks for 43 pieces of text

You tick `Y` in the drafting sheet's column E, then tick `Y` again against the same 43
paragraphs in the review grid under a different column name.

**Fix:** carry the drafting approval forward. Pre-tick a review row **only** where the
slide's current text is still exactly what the register last wrote — so a slide someone
hand-edited since still arrives blank and demands a read, which is the case the review
grid was built for. Every gate survives: current-vs-proposed is still shown per row,
every row is still untickable, the change hash still drops stale approvals.

**NEEDS ROHAN'S EXPLICIT SIGN-OFF.** It sits adjacent to R13.2 (prose may never be
batch-approved) even though it is not batching. Do not ship on a reviewer's say-so.

---

## 4. `Kind` is answered in two places and they can disagree

`FieldSpec` has a user-editable `Kind (Controlled/Prose/Static)` column. `ReviewQueue.
ContentKindOf` hardcodes the same three values for three field names. The sheet governs
which fields get drafting sheets; the hardcode governs batching. Edit the sheet and
batching does not move — silently.

**Fix:** `ContentKindOf(fieldId, specWs)` reads the sheet, falls back to the built-in
table when there is no row, keeps `Prose` as the unknown default so absence is never
read as permission to batch. Rohan's design, 2026-08-08: the vocabulary comes from the
code as a **dropdown**, the assignment comes from the sheet, and an unrecognised value
is REPORTED rather than silently defaulted.

---

## 5. Two questions asked of a person that nothing reads

`FieldType` (text/number/currency/date) and `FieldVolatility` (static/variable) are
asked once per field during marking, normalised, serialised, round-tripped through the
marking session, written into the review grid — and read by nothing. On the `Setup A`
path that is roughly **104 dialogs that alter nothing**, across ~52 fields.

**Fix:** delete both. Nothing becomes impossible; nothing read them. If field typing is
wanted later, `Kind` on the Field Spec sheet is where it belongs — editable, not a modal
at mark time.

---

## 6. Five user-facing strings that describe a tool that no longer exists

Fix as a CLASS, with a grep for the shape, not one at a time. The 3de4be8 sheet rename
left four stale readers and they have been found in three separate sessions.

- `FastPathRefusalText` still says sync "does not create slides" — untrue since the
  25% create path landed.
- `WorkbookBridge.DescribeSheet` / `LifespanOf` / `IsToolOwnedSheet` still match
  `"Sync Review"`, so `START HERE` labels the live review sheet
  *"(not created by this tool)"*, lifespan *"unknown"*.
- `DescribeSheet("Register")` describes the LONG register — one row per project, field
  and quarter, with approval state. That model was retired 2026-08-03.
- `IsToolOwnedSheet` is missing `"Run Log"` entirely.
- `WORKFLOW.md` still says columns G/I, "no button for roll forward", and "nothing ties
  a source to a period" — all three fixed in code, none in the doc.

---

## 7. Things typed that could be picked

- **`Roll Forward`'s source period** is typed free-hand, into exactly the trap its own
  header warns about for the destination. `Sources.ApplyPeriodValidation` already builds
  this list from the register's own `Quarter` values.
- **The slide-type picker** in `Audit Fields` and `Create Template Slide` asks a question
  with one legal answer on any real deck — `ResolveRegisterSheet` refuses a deck with
  more than one type. Auto-select when there is one; keep the picker for the rest.
- **The second unsaved-workbook prompt** in `ApplyApprovedCore` re-asks after
  `ReviewChangesCore` already saved. Measure once, before the tool dirties anything —
  same fix as the Preview Sync footprint bug.

---

## 8. Three of six context switches exist because the workbook cannot find its deck

Steps 1, 2 and 3 are PowerPoint buttons whose entire effect is in the workbook.
`WORKFLOW.md` says so for each: *"Touches: the workbook only. Never the deck."* They are
PowerPoint-hosted because `DraftingUI` needs `ActivePresentation` for two facts — the
period and the workbook path — and the workbook stores the deck's **opaque ID**, never
its path.

**Fix (own session):** write the deck's full path and current period into the workbook's
custom properties on every `Start a Quarter` and publish, then ship an Excel-hosted
`.xlam` carrying steps 1–3. Same Subs, no forks — `TestRunnerExcel.bas` already proves
the modules run Excel-hosted. Costs a second package to version; saves four alt-tabs per
field per quarter.

---

## Capability gaps — a person cannot do these at all

- **Retire a slide or a project.** Nothing removes a slide the register no longer
  mentions. It shows only as a parity mismatch, resolved by hand.
- **A maintained list of linkage codes.** Strategic Alignment must cite codes it can
  check. Rohan owns it, colleagues later — so it needs an `as at` date and probably a
  home outside the register workbook, which the tool rebuilds.
- **Provenance that survives a rollover.** Source citations live only on the drafting
  sheet, which is cleared at every period change. The record answering "why does it say
  90%?" has a lifespan of one quarter.
- **The other 42 slides.** Only slide 4 and the template carry the three new panels.
  Extending them is `Setup A2` → `Setup B`, which has never been walked by Rohan
  unaided — and until it is, "is the setup formulaic?" is unanswered.
