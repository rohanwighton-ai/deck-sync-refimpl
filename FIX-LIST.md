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

## 1a. A message that reports a COUNT must name its SUBJECT

**Three instances in one afternoon, 2026-08-09. Each was true, and each was unusable.**

- `Copy AI to Submit`: *"Nothing to copy: there are no AI drafts on this sheet yet.
  Column F is empty for all 43 row(s)."* It acts on whichever `TPL_` sheet is ACTIVE in
  Excel. The active tab was `TPL_ABOUT_BODY`; the work was on
  `TPL_STRATEGIC_ALIGNMENT_BODY`, where column F held 812 characters. The message never
  names the field, so a correct statement about one sheet reads as a flat contradiction
  of what you just did. Its success message has the same hole — *"1 copied"*, of what?
- `Create Template Slide`: *"5 field(s) set to placeholders."* True. The deck had EIGHT
  tagged fields; the template was cloned from a slide carrying only five. Naming them
  would have shown the three missing at a glance instead of after a byte-level check.
- `OpenOrGetWorkbook` (item 1): *"Could not open the paired workbook at <path>"* — names
  the file that is FINE and not the duplicate filename blocking it.

**The shape:** the code holds the identifying detail at the moment it composes the
message, and drops it. What survives is a number with no subject, which is worse than
silence — it reads as authoritative and sends you to check the wrong thing.

**Fix:** every count in a user-facing message names what it counted. `"Nothing to copy
for ABOUT_BODY"`, `"5 of 8 fields set to placeholders -- missing: ..."`, `"...blocked by
register-wide.xlsx already open from <other path>"`. Mechanical, and it is the single
cheapest reduction in time-to-diagnose available in this codebase.

**Worth guarding, not just fixing:** this is the fourth time a message has been true and
unusable. A test that asserts every `MsgBox` composing a count also interpolates a name
would be crude, but it would hold.

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

---

# Loose ends — found in conversation, at risk of being lost

Small, real, and none of them recorded anywhere else as at 2026-08-09.

## Five status values differ by one capital letter

`Q1F27` rows for `3_P001`, `3_P002`, `2_P003`, `2_P004`, `1_P005` hold `'Not started'`
against the vocabulary's `'Not Started'`. Reported on every drafting run as
*"5 value(s) are not in their allowed list. Nothing was changed."* and never chased.

The tool is right not to auto-correct — `ApplyControlledValidation`'s own comment says
so — but the message never names the five, which is fix-list item 1a again. Five cell
edits.

## UNVERIFIED: does `||||` render as a paragraph break on the slide?

`2_P004`'s published Strategic Alignment stores a blank line between paragraphs as
FOUR pipes (`||` is one line break; a blank line is two). Publish encodes it, and
`Drafting` renders `||` back to real breaks in column C — but **nothing has yet
confirmed what reaches the SLIDE.** If it arrives as literal `||||`, every multi-
paragraph field is affected, and it will be visible on 43 slides at once.

**This is the last unverified link in the chain.** Check it the first time anything
syncs to slide 4.

## Duplicate filenames block the add-in, and the machine is full of them

Excel refuses two open workbooks with the same NAME regardless of folder. A review copy
called `register-wide.xlsx` in `OneDrive\Claude\` blocked the live one and surfaced as
*"Could not open the paired workbook"* pointing at the file that was fine.

Also on the machine as at 2026-08-09: `addin52`–`addin56` all in the AddIns folder, plus
stray `addin55.ppam` and `addin56.ppam` in the OneDrive root. Only `56` is needed. Six
near-identical add-ins is how the wrong one gets loaded on a tired evening.

**Rule for review copies from here: never reuse the live filename.**

## One judgement call in the first drafted field, left in deliberately

`2_P004`'s Strategic Alignment contains *"rather than monitoring broadly and intervening
on judgement"*. That edges toward `PROBLEM_BODY`'s territory — the "so what" is hard to
state without a contrast against current practice. Kept, and flagged: if Rohan reads it
as a bleed, the own-job test needs a sharper line, and that is a Field Spec edit rather
than a code change. It is the first live test of whether the boundary wording holds.

## The Run Log fix is not in `addin56`

`addin56` was built before the `NumberFormat = "@"` fix landed. Until the next build,
every Run Log still loses its body — so "read the Run Log" is not usable advice yet.

## Source assets — the state as at 2026-08-09, and what work means

**Sources sheet: 6 rows, every one prefixed `EXAMPLE (not a real document)`.** Invented
2026-08-08 to exercise the citation check. Still there.

**Citations: ZERO, on every field.** This is new today and it is not a cleanup — the 37
fabricated citations on `TPL_ABOUT_BODY` were WIPED when the drafting sheets rebuilt at
the period change, because `cadence` is permanently `Nothing` so every row cleared.
The provenance-dies-at-rollover problem, demonstrating itself on real data. Half the
scaffolding removed itself; the six sheet rows did not.

**Two of the six now point at a period that no longer exists** — `S02` and `S05` are
`Applies to: Q4F26`, and the register holds only `Q3F26` and `Q1F27` after the relabel.

**So "source assets need work" means, concretely:**

1. **Delete `S01`–`S06`.** Nothing cites them now, so it costs nothing — and leaving
   fabricated provenance next to real content is worse than having none.
2. **Build the first real source from the question already asked.** `J13` on
   `TPL_STRATEGIC_ALIGNMENT_BODY` holds it: are `1.4.2, 1.5.2, 2.3.1` DECLARED linkage
   codes, and where does that list live? That answer becomes the first genuine Sources
   row and unblocks the `[TBC]` in `2_P004`'s published text.
3. **Use the form.** `SOURCE-HARVEST.md` has the rubric and the fields; it needs no
   tools and can be filled in at work.
4. **Decide where a maintained list lives.** Rohan owns it now, colleagues later — so it
   needs an `as at` date, and probably a home outside the register workbook, which the
   tool rebuilds and clears.

**And the structural point stays open:** citations live only on the drafting sheet,
which is cleared at every period change. Until that moves to the register — same grain
and lifespan as the text it explains — provenance has a maximum life of one quarter,
and today proved it empirically rather than theoretically.
