# Fix list

> **CURRENT — the live list of what is known-broken and not yet fixed.** Re-audited
> against the code 2026-08-14; four entries added 2026-08-15 (see the last section).
> Entries say whether they are still live; anything marked fixed names the build it was
> fixed in.

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

**FIXED (verified in the workbook 2026-08-14).** The global rules now carry an
explicit empty-column-C clause: *"Where column C is EMPTY, this field has never been
published for that project. Write it fresh against the Length above. Do not leave the row
blank on account of an empty column C."* The account below is kept for the reasoning.

**The cheap one -- the global rules assumed a prior value existed.** They say *"Column C
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

## 1c. FIXED 2026-08-14 (`aff84d6`) — the at-risk scan missed SOURCES and NOTES

> **Fixed, not compiled.** The scan now counts SOURCES and NOTES alongside SUBMIT and
> DRAFT, so the rollover refusal covers all typed work. Original entry follows.


**Found 2026-08-13 while rewriting the two tests the refusal guard turned red.**

`WriteDraftingSheet`'s refusal counts a row as at risk only if `COL_D_SUBMIT` or
`COL_D_DRAFT` holds text (`Drafting.bas`, the `If periodChanged Then` block). But the
rebuild carries **four** columns of human work — DRAFT, SUBMIT, SOURCES, NOTES — and the
rollover drop clears all four. So a row where a person has assigned citations or written
notes, but has not drafted yet, does not trip the refusal, and its work is discarded.

Citations are not incidental: per `project_deck_sync_provenance_is_field_architecture`,
the source IDs are the control on the generative step. Losing them silently is the same
failure the guard was built to stop, one column over.

**Not live on the real workbook right now** — the 129 drafted values sit in columns E and
F, exactly what the scan does cover, so every sheet refuses and everything is protected.
This is the next instance of the class, not an active threat.

**Fix:** count all four columns as at-risk. **Then check whether the per-row cadence
machinery should be deleted rather than fixed** — see below.

**The bigger question this exposes.** Once the refusal covers all four columns, the
rollover drop is reachable only for rows holding *nothing*, where there is nothing to
drop. The whole cadence mechanism — the `cadence` dictionary, `carryThisRow`,
`droppedQuarterly`/`keptStatic`, and the test that covers them — would become dead code
whose most heavily-commented property ("the drop is per row, not per sheet, and that
distinction is the whole of this guard's correctness") no longer decides anything. It was
correct machinery for a design that the refusal replaced. Deleting it is probably right,
and is Rohan's call, not a silent cleanup.

---

## 1d. FIXED 2026-08-14 (`aff84d6`) — by deletion, not repair

> **The late `ParkSheetCopy` call site is gone.** The park now runs before
> `ws.Cells.Clear`, unconditionally, which made this site both unreachable and wrong.
> Note the clear itself is now confined to a layout migration, so on the normal path
> there is nothing to park *from*. Original entry follows.


**Found 2026-08-13, in the code directly beneath the comment warning about this.**

`Drafting.bas`: `ws.Cells.Clear` runs at the top of the rebuild. The `ParkSheetCopy` call
for the rollover case runs in the *reporting* section, hundreds of lines later —
`If lostWithContent > 0 And parkedName = "" ...`. `ParkSheetCopy` does `ws.Copy`, which
copies the sheet **as it is at that moment**: already cleared, already rebuilt, with the
dropped rows' content gone. `ParkedNote` then reports *"The previous sheet was kept as
'<name>' — nothing was lost."*

The archive is real, named, and hidden on the workbook. It just does not contain the
thing it was taken to preserve.

This sits immediately below the comment quoting ReviewQueue's own rule: *"A REPORTED
BACKUP THAT IS NOT ON DISK IS WORSE THAN NO BACKUP: it is the reason you feel safe
running the destructive write that follows."* The layout-mismatch park at the top of the
function is correctly placed, before the clear; only this second call site is wrong —
the same "fixed where it was found, not everywhere the shape exists" pattern the repo has
now logged five times.

**Narrow but reachable:** needs a rollover where a dropped row has NOTES content and no
row anywhere has SUBMIT or DRAFT (otherwise the refusal pre-empts it). Fixing 1c closes
it by making the path unreachable, which is the cleaner fix of the two.

**Fix:** park before `ws.Cells.Clear`, or delete the second call site along with the
cadence machinery if 1c is resolved by deletion.

---

## 2. SUPERSEDED — the button this described no longer exists

> The toolbar had `0`–`4` numbered buttons when this was written; there are now two.
> Kept because the CAUSE is still live: `ContentKindOf` returns `KIND_PROSE` for
> everything outside three hardcoded names, so `HasBatchableWork` is never true for a
> prose field. Under the chain that is simply the normal route to the review sheet.

### As written, against the old toolbar

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

## 6. User-facing strings that describe a tool that no longer exists

**Re-audited against the code 2026-08-14** — the list grew rather than shrank. Each entry
below now says whether it is still live.

Fix as a CLASS, with a grep for the shape, not one at a time. The 3de4be8 sheet rename
left four stale readers and they have been found in three separate sessions.

- `FastPathRefusalText` still says sync "does not create slides" — untrue since the
  25% create path landed.
- `IsToolOwnedSheet` still matches the prefix `"Sync Review"`, while
  `ReviewQueue.ReviewSheetNameFor` produces `"Review project-progress-A32C"`. **Re-checked
  2026-08-14: STILL LIVE.** So the review sheet a person is actually working in is not
  recognised as tool-owned, and `START HERE` labels it *"(not created by this tool)"*.
- **`IsToolOwnedSheet` is missing `Run Log` AND `Register`.** Re-checked 2026-08-14: still
  true of both. `LifespanOf` knows `Register` perfectly well two functions away — the two
  disagree about the most important sheet in the workbook.
- `DescribeSheet("Register")` describes the LONG register — one row per project, field
  and quarter, with approval state. That model was retired 2026-08-03.
- **Dead caption constants.** `CommandBarUI` defines 19 `CAP_*` strings; `AddButton` is
  called twice. The other 17 name buttons that do not exist, including the whole
  three-chain design (`1. Start the quarter`, `2. Draft and publish`,
  `3. Put it on the slides`). Any message interpolating one of those is describing a
  button nobody can press — which is the exact failure `TOOLBAR.md` predicted.
- `WORKFLOW.md` still says columns G/I, "no button for roll forward", and "nothing ties
  a source to a period" — all three fixed in code, none in the doc.

---

## 7. Things typed that could be picked — 2 of 3 FIXED 2026-08-13 (addin81/82)

**Status, so this does not read as fully open:** the slide-type picker is **fixed**
(`RibbonUI.PickType` auto-selects a sole type, all three call sites). Roll Forward is
**partly fixed** — it no longer asks at all when the destination already holds rows
(`ExcelOutput.PeriodRowCount`), but when the destination IS empty it still asks for the
source period as free text, so the validated-list fix below still stands. The third item is
**untouched**.


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
  check. **Rohan owns it** (settled; do not re-ask), colleagues later — so it needs an
  `as at` date and a home. **CORRECTED 2026-08-14: that home is a PERMANENT sheet in the
  register workbook, beside `Sources`.** The earlier wording here said "probably outside
  the register workbook, which the tool rebuilds" and that is wrong —
  `WorkbookBridge.LifespanOf` classifies `Register`, `Field Spec` and `Sources` as
  PERMANENT; only `TPL_*` and review sheets are rebuilt. An external file would be a
  second artefact to keep in step, which is the class of problem template-first removes.
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
4. **Decide where a maintained list lives.** Rohan owns it (settled). It needs an `as at`
   date and a PERMANENT sheet in the register workbook, beside `Sources` — see the
   correction under Capability gaps. The workbook is not rebuilt; only `TPL_*` and review
   sheets are.

**And the structural point stays open:** citations live only on the drafting sheet,
which is cleared at every period change. Until that moves to the register — same grain
and lifespan as the text it explains — provenance has a maximum life of one quarter,
and today proved it empirically rather than theoretically.

---

## PARKED: name each prose field's recipe, so it can be discussed plainly

Rohan, 2026-08-14, and parked by him in the same breath: *"giving the particular
generative agents goblin names associated with their coding task so I can talk about them
more plainly with non-AI, non-code, non-data people. We do not need to get too carried
away by that now."*

**The thing being named already exists -- it is the Field Spec row.** A recipe holds a
purpose, a voice, a length, a Do-NOT list and an *own-job test* whose entire function is
to stop one field wandering into a neighbouring field's territory. That is a job
description, not a config row. Thirteen `Kind = Prose` fields, thirteen of them.

**What it buys, beyond plain speech:** the bug report writes itself. *"The Strategic
Alignment goblin keeps doing Problem's job"* is a precise statement that an own-job test
is failing, and it is intelligible to someone who has never heard of a register. This
project's hardest recurring problem is boundary bleed between adjacent panels; a named
owner per panel makes that boundary a thing people can point at.

**If it is built, it is ONE COLUMN on the Field Spec, not a roster in a document.** A
markdown list of names is a second copy of a machine-knowable fact and will drift -- the
class of defect this repo spent 2026-08-14 removing. As a column it derives out wherever
the field does: prompts, drafting tab labels, review sheets, run reports.

**The risk worth holding:** a name makes the output easier to over-trust. A goblin has a
brief, not judgement. It is why the approve gate exists and why naming must not soften it.

**Not scheduled.** The interesting half of the design session is what each goblin
*refuses* to do, which is the Do-NOT column read aloud.

---

## PARKED, nice-to-have: show/hide control for fields on a slide

Rohan, 2026-08-09, asked for it and then ranked it himself: **nice to have.** Not
scheduled. Recorded so the design conversation is not re-run from scratch.

**The want:** control whether a given field appears on a given slide -- "this
project has no Problem section this quarter".

**Grain:** the register row is already slide x period, so one `Hidden fields`
column on that row gives per-slide, per-period control. Field Spec is the wrong
home: it is per-field globally and cannot say "hidden on this project, not that
one". Values picked from known field IDs, never typed -- same rule as
`Sources.ApplyPeriodValidation`.

**Mechanism:** `Shape.Visible = msoFalse`, not blanking the text and not
deleting the shape. Blanking leaves an empty box with its fill and border still
drawn. Deleting destroys the tag, so sync could never bring the field back --
and it must come back, because next quarter the section returns.

**What Rohan wants to happen to the space, in his words:** *"ultimately leave
it. or replace it. some items could move."* So the behaviour is not one rule:
leaving the gap is the default, some panels would be REPLACED by other content
rather than emptied, and a minority of items would move. That last part is the
expensive one -- PowerPoint does not reflow, so closing a gap means the tool
moves slide furniture, which is a much larger promise than filling in text and
has to handle stacking, columns and grouped shapes without silently damaging a
layout.

**If it is ever built, build it in that order:** leave-the-gap first (cheap,
reversible, no geometry), replace-with second, move-things last or never.

---

## The time-elapsed bar: the best automation target on the slide

Rohan, 2026-08-09, drawing the distinction the model was missing:

> timeline will move towards end but not necessarily every quarter, vs time
> elapsed bar autoshapes that move with the clock regardless of progress

**Judgement versus clockwork.** Milestone markers move when the plan changes and
someone decides. The elapsed bar is a pure function of the date -- nobody
decides it, and it is wrong on every slide the moment the quarter turns, in a
way nobody notices until a reader does.

**It beats the prose panels as an automation target, which is counterintuitive
given where the effort has gone.** Prose needs drafting and review, and
automating delivery still leaves the expensive human part. The elapsed bar needs
NOTHING from a person: the slide already carries Start and End, the period
supplies the third input, and position is arithmetic. It changes every quarter,
guaranteed, on all 43 slides. Rohan's words: fiddling with **its two component
autoshapes is a PITA** -- so it is two shapes, not one, and the manual cost is
per-shape.

**It needs a KIND that does not exist.** `Kind` is how content is decided
(Controlled / Prose / Static); `FieldType` is what it is (Text / Picture /
Shape). The elapsed bar is `FieldType = Shape`, `Kind = Derived` -- computed,
never drafted, never approved, re-derived every run. Nothing in the tool has
that shape today, and it is the only kind that could sync with no human in the
loop at all.

**Design, as agreed:** the register says WHERE, sync just applies it -- keeps
sync dumb, which is what unattended quarterly work needs. Two refinements on
top of that, both about avoiding a fragile coupling:

- **Store a FRACTION, not points.** Raw PowerPoint coordinates in the register
  means the day anyone nudges the timeline, all 43 stored positions are quietly
  wrong and still look synced. Store `0.72` = "72% along", tag the timeline axis
  itself, and let sync read THAT shape's own Left and Width at run time. Moving
  the track then fixes itself.
- **Let Excel compute the fraction, not VBA.** A formula over start, end and
  period end is visible and checkable in the sheet, and it keeps date handling
  out of the VBA -- which matters specifically here, because the code has
  already been bitten by period parsing (`Q4F26` vs `FYnnQn`) and this sidesteps
  the whole class rather than adding to it.

So the register row gains `TIMELINE_START`, `TIMELINE_END`, `TIMELINE_ELAPSED`
(a formula over the first two and the period). Sync reads only the last.

**Open:** whether formatting varies too. Size alone covers "the bar gets
longer". Colour -- amber when the end date falls inside the current quarter, red
once passed -- is a second column AND a second decision, because a bar that
changes colour is making a claim about the project rather than showing a date.

**Not for the week it was raised in:** building it costs longer than dragging
the shapes once. This is next-quarter work, ranked above the Excel polish and
above show/hide.

---

## PARKED WITH ITS DESIGN SETTLED: picture fields

Rohan, 2026-08-10: pictures **"don't change quarterly, set once at project
start, I'd still like them set from link rather than played around with
manually"**. That answer makes this much smaller than the parked note assumed.

**It is not a sync field.** `FieldType = Picture`, `Kind = Static/Given` -- the
two axes the specs keep apart and must not collapse into one word. It carries
across a rollover untouched like any period-invariant content, so it costs
nothing per quarter.

**The cell holds a SOURCE ID, not a path.** Sources already solves one row per
thing, reference by ID, period binding, existence checking, and now feeds the
prompt. An image is evidence-shaped: it came from somewhere, that somewhere has
an owner, and "why is this photo on the slide?" is the same question as "why
does it say 90%?". A raw path per row would duplicate all of that badly -- 43
spellings of one folder and nothing checking any of them.

**Idempotence without comparing images:** when the tool fills a picture it
STAMPS THE SOURCE ID IT USED into the shape's own tag. Sync fills only when the
register's ID and the shape's stamped ID differ. So it fires once at project
start, stays silent forever after, and re-fires by itself if the link is ever
changed. No image comparison, no re-inserting 43 photos a quarter.

**Mechanism, once decided:** insert, copy Top/Left/Width/Height and z-order from
the tagged shape, delete the old, re-apply the tag. Fiddly, not deep -- the
tagged shape already knows its own geometry.

**THE ONE OPEN DECISION, and it is Rohan's:** when the image's proportions do
not match the frame -- fit inside and letterbox, fill and crop the overflow, or
refuse and report. Visibly different slides, no defensible default; it depends
whether the project photos are consistently shaped. Answer it against a few real
ones before anything is built.

---

# THE TIMELINE, READ FROM THE REAL DECK 2026-08-10 -- READ BEFORE TOUCHING BARS

Rohan: *"you can see the extremely accurate positioning? we need to maintain
that and z order etc."* Everything below is measured from `slide4.xml`, not
described from memory.

## What the timeline actually is

- **A VERTICAL bar of two rounded rectangles sharing an origin.** Track
  `Shape 188` is 0.05 x 3.03; fill `Shape 189` is 0.05 x 1.46 at the same x/y.
  The fill grows DOWNWARD from a fixed top, so drawing it changes exactly one
  property -- `Height`. No `Top`, no `Left`, no z-order.
- **`InjectProgressField` IS HORIZONTAL ONLY and refuses to guess the axis.**
  It cannot draw this slide today. That is a blocker, not a tweak.
- **Milestone state is FORMATTING, not content.** Circles are 0.35 except one
  at 0.43, and fills split `005832` (start, 6-month) against `003C23` (the
  rest). No register value can express size and colour.
- **The fill length is DERIVED** -- it runs to the next unachieved circle. So
  the register should carry WHICH MILESTONES ARE ACHIEVED, and the length is
  measured from circle positions the template already defines. The earlier
  `0.25||0.5||1` model assumed independent fractions per bar and is wrong for
  this layout.
- **Counts vary and are therefore data.** Ellipses per slide across the 44:
  mostly 6, but 4, 5, 7, 12 and 14 all occur.

## Three things that make naive manipulation dangerous

1. **Nothing is hidden.** No pre-placed big/small pair exists today, so a
   visibility model needs the TEMPLATE built for it. The tool cannot retrofit
   it.
2. **Shapes are stacked and share names.** THREE separate shapes are called
   `Shape 202`, two of them at the identical position `10.96` -- one filled,
   one not. Name-based addressing is hopeless; tags are the only safe handle.
3. **Structure varies within one slide.** Most circles are groups of
   circle+number; the 12-month one is a bigger loose ellipse with its number
   floating separately on top. Any "assume the pattern" logic breaks here.

## What is safe to change, and what is not

| operation | verdict |
|---|---|
| fill bar `Height` | SAFE -- origin fixed, one property |
| circle **fill colour** | SAFE -- no geometry touched |
| circle **size** | **NOT SAFE** -- resizing about a centre needs Top/Left compensation, which is the arithmetic that hid `LockAspectRatio` for five rounds |
| replacing any shape | NEVER -- a new shape lands on top and z-order is lost |
| writing label text | needs the existing geometry save/restore -- autofit moves the box |

## DECIDED 2026-08-10: THE VISIBILITY MODEL

Rohan: *"I like the idea of the visibility model."*

**Achievement is shown by toggling `.Visible` on pre-placed shapes, never by
resizing or recolouring computed by the tool.** The template carries both
states -- an achieved circle and an unachieved circle, at the same centre --
and the tool shows one and hides the other.

Why this and not the alternatives:

- **Position is preserved exactly**, because nothing moves. The tool computes
  no geometry, which is the standing rule the deck's own precision demands.
- **Z-order is preserved exactly**, because no shape is created, deleted or
  reordered. `.Visible` does not touch the stack.
- **The tool never invents a colour or a size.** Both states were authored by a
  person in the template, which is the same division as everywhere else here:
  the template owns geometry and formatting, the register owns values.

**What it requires next, in order:**

1. **A template decision by Rohan** -- the template gains a second, hidden
   circle per milestone in the achieved styling. Until that exists there is
   nothing to toggle. This is deck work, not code.
2. **A vertical mode for the bar.** Derive the axis from the TRACK's own
   dimensions (0.05 wide x 3.03 tall is unambiguous) and refuse only when the
   track is square-ish -- the existing refusal-to-guess comment is right about
   a square and wrong to give up on a 60:1 rectangle.
3. **Achievement in the register**, per milestone, `||`-separated like the
   values -- then the fill height is measured to the next unachieved circle
   rather than supplied.

**STILL UNKNOWN, needs Rohan's slides:** whether the achieved/unachieved pair
should be two shapes per milestone or one shape per state for the whole
timeline; and whether the 0.43 circle marks ACHIEVED or CURRENT -- on slide 4
it is dark (`003C23`, the unachieved colour) while the two lighter circles
above it are 0.35, which does not fit "big means achieved" on its own.

---

# STATE AT 2026-08-10 07:45 -- READ BEFORE PLANNING ANYTHING

## THE ONE THAT MATTERS: two features are built and unreachable

`InjectPictureField` and `InjectProgressField` are unit-tested and called by
**nothing except their own tests**. The sync path calls
`InjectPrimitive.InjectPrimitive` -- the TEXT injector -- for every field
(`ReviewQueue.bas:1283` and `:1301`, `SyncOperations.bas:157`), so a picture or
progress field is handed to the text writer and refused as "no text frame to
write into".

169 passing tests say nothing about this. Rohan found it by asking whether the
bars had been proven slide to slide.

**In this order:**

1. ~~**Template-clone test FIRST.**~~ **ANSWERED 2026-08-10 08:15 — THEY SURVIVE.
   The track-pair design stands; nothing has to change.** Run against real
   PowerPoint via `vba/tools/tag_cloning_probe.ps1` (the probe had existed since
   31 July with no runner and no recorded result — written, never run, which is
   the same shape as the unreachable injectors it exists to answer for).

   What survived `Slide.Duplicate`: the suffixed role value `BAR_BODY.track`, its
   pair `BAR_BODY`, **both inside a GROUP**, and a second tag name on the same
   shape (`picsrc`). Slide-level tags survive too, and the duplicate gets a new
   SlideID.

   Two things the first run could not have told me, both fixed before the verdict
   was believed: `DumpTags` printed a grouped shape and a never-grouped shape
   identically, so it could not show the group had actually formed (it now prints
   `[Group 4] GROUP of 2` and marks members); and nothing demonstrated the finder
   could return "not found". Two controls now run and **must** read False — a role
   never tagged, and a picture stamp on the track shape. Both did.

   Incidental, not today's question but worth having in writing: **Copy/Paste into
   a different presentation preserved the SlideID exactly.** SlideID is unique
   within a deck, not across decks — do not use it as a cross-deck discriminator.
2. ~~**Dispatch by shape type in the sync path.**~~ **BUILT 2026-08-10.**
   `InjectPrimitive.InjectField` is the one entry point; the type is derived
   from the SHAPE (picture / a `.track` sibling / else text), so it cannot
   disagree with a column the way `Kind` does. All three sync call sites rewired
   (`ReviewQueue` ×2 with the Sources sheet resolved once per run,
   `SyncOperations` without one). New `Sources.LocatorFor`, with a `found`
   out-parameter so "cited a source that does not exist" and "source row has a
   blank locator" get different messages — they send the person to different
   files.

   A non-numeric register cell for a bar is refused **loudly**: `Found` and
   `WouldChange` are both set, because `Found=False` is a SKIP to
   `SyncOperations` and `WouldChange=False` is a NO CHANGE — either would let a
   bar silently fail to draw inside a run reporting success. `Val("done")`
   returning 0 would have drawn an empty bar and called it a success.

3. ~~**A multi-slide test.**~~ **DONE** — `InjectField_TwoSlidesEachGetTheirOwnValue`,
   with the two slides given **different track widths** so a bar drawn against
   the wrong slide's track cannot land on the right number by coincidence.

**172 passed / 0 failed behind the compile gate, and every new test was made to
fail on purpose first, in two separate runs.** Breaking both non-text branches
failed all three new tests while the five existing picture/progress tests stayed
green — which is the blindness itself, since those test the injectors directly
and cannot see that nothing calls them. Breaking the picture branch ALONE then
failed only the picture assertions, with the failure text being the exact defect
this work removes: `shape tagged role=PHOTO has no text frame to write into`.

## 2a. PICTURES ARE NOW REACHABLE. BARS ARE STILL NOT — THERE IS NO WAY TO MARK ONE

**Found 2026-08-10 by Rohan asking "how are these types of fields marked",
while the dispatch was still being written.** The dispatch was necessary and is
not sufficient.

- **Pictures: markable.** `BatchOnboardFlow`'s gate accepts a picture shape
  (added earlier the same day) and `Onboarding.IsCandidateField` offers pictures
  in discovery. With the router in, pictures now work end to end.
- **Progress bars: not markable at all.** Two independent blocks. The marking
  gate (`BatchOnboardFlow.bas:1429-1436`) requires a picture OR a shape with
  non-empty text; a bar's done part and its track are empty rectangles, so both
  are refused — the message says *"progress bars are not supported yet"* in as
  many words. And **nothing outside `InjectPrimitive.bas` mentions `.track`
  anywhere in the codebase**: `Onboarding.ConfirmFieldMatch` is the only thing
  that ever writes a `role` tag, and no path offers a suffixed value.

**DECIDED 2026-08-10, not yet built: bars are tagged ONCE ON THE TEMPLATE, not
per slide.** Rohan's call, and this morning's probe is what makes it safe — the
tags survive duplication, including inside a group. A bar is furniture the
template owns, so there is no per-slide marking and no guessing which of two
shapes is the track.

**Also decided, and it is a boundary worth keeping:** `Behaviour` (Fill / Fit /
Leave as is) **stays on the Field Spec for now**, even though it is a shape fact
and would arguably be better on the tag. The line agreed: the template owns what
is true of the SHAPE, the workbook owns what is true of the CONTENT — and
`Kind`, which is already double-sourced, gets fixed by making the sheet win, not
by moving it to a third place.

**The standing cost of leaning on tags: a tag is invisible.** It cannot be seen,
diffed or repaired without the tool, on the machine with no Python and no
Claude. If real behaviour moves onto the template, a toolbar-reachable "what
does this slide say about itself" dump stops being a nicety —
`read_deck_slide_tags.py` is Python and therefore useless at work.

## 2b. `field_e2e.ps1`'s module list is missing `Readiness.bas`

Pre-existing, found 2026-08-10 by running `vba/tools/check_module_lists.py`
(not caused by the dispatch work; the test-suite and add-in lists both pass).

    FAIL  harness  (field_e2e.ps1)
            missing Readiness.bas -- referenced by DeckRegistry, RibbonUI, WorkbookBridge

The harness will fail at runtime on any path that reaches Readiness. Cost: one
line.

## What is built and verified

> **SUPERSEDED IN PART, 2026-08-14 night.** The two buttons named below are gone.
> The toolbar is now `1. Set up my quarter` / `2. Put it on the slides` /
> `Review changes (writes nothing)`, split by ARTIFACT so neither side can trigger
> the other, and `Rebuild my sheets` is deleted outright. `Layout 4` below is now
> **layout 5**. The verification records are still true of the day they were made.

> - **Two buttons** (`1. Sync Now`, `2. Rebuild my sheets`) with every capability
>   reached from inside the chain; `check_vba_static.py` fails the build on an
>   orphan and caught nine during the refactor. *(Superseded — see the banner above.
>   Kept inside the quote because it is a record of that day, not a description of
>   the tool.)*
- **The whole quarter loop ran through them**, verified from file bytes: period
  set and confirmed on disk, 43 rows rolled forward, a second roll refused,
  `PendingApprovals` fired on its first real press, three fields written to
  `2_P004`, backup taken one second before the write.
- **Layout 4** drafting columns in workflow order, with `ColumnInLayout`
  migrating layout-3 sheets rather than dropping their work.
- **Cited sources reach the prompt** (`Sources.CitedBlockFor`). Before this the
  evidence rule could never be satisfied -- `[TBC]` was permanent no matter how
  many sources were cited.
- **Colour family per field**, from POSITION not a name hash (a hash collided
  ABOUT_BODY with PROGRESS_BODY, the two most-used sheets).
- **Every column letter and button caption in user-facing text is derived**, not
  typed. 29 stale captions swept on 08-09, four stale column letters on 08-10.
- **Build stamp** written by `build_ppam.ps1` into `CommandBarUI.BUILD_STAMP`
  and shown on every report. **Not in `addin62`** -- it was added after.

## Rules this codebase earned the hard way, tonight

- **Green is not evidence; green plus a demonstrated red is.** Five checks
  tonight passed while being incapable of failing. Twice the first attempt to
  break a test failed to break it -- which is itself the finding.
- **Probe the mechanism before modelling the symptom.** Five rounds of geometry
  arithmetic were spent compensating for `LockAspectRatio`, one property, which
  a thirty-second probe settled. A retry loop that re-asserts a value is an
  admission that something else is changing it -- go find out what.
- **The template owns geometry, the register owns values, the tool computes
  neither.** Every fit/fill/scale calculation was the tool deciding something
  the template had already decided.
- **Trace the path from a person's action to the code before calling it done.**
  Both unreachable features would have failed that sentence immediately.

## Rohan's constraints, current

- The workbook can NEVER be `.xlsm` -- work restrictions. No workbook-level VBA
  ever; all behaviour lives in the `.ppam`. Add-ins themselves ARE permitted at
  work (`addin33` ran there).
- Native in-cell checkboxes have no VBA API. The achievable "toggle" is a
  validation dropdown plus conditional formatting.
- Pictures are set once at project start, not quarterly -- so a picture field is
  a `Given` filled from a link, not a sync field.
- Sharing turns on **portability, install, single truth** -- his words. Single
  truth forces a shared register, which forces the cloud path.

---

## PAPERCUTS FROM THE FIRST SUCCESSFUL PUBLISH — 2026-08-13

All found by pressing buttons on the real deck, none visible to 192 passing tests.
Ranked by how much of an evening they cost.

### P1. A dialog opens BEHIND the PowerPoint window, and reads as "nothing happened"

**Three times in one session.** Rohan pressed a button, nothing appeared, and the run
looked dead. Each time a VBA modal was sitting behind another window — twice behind
Excel, once behind PowerPoint itself. It cost two separate diagnostic detours before a
reliable test was found.

**The calibrated test, worth keeping:** while a modal is open PowerPoint stops answering
COM — `ActivePresentation.Name` comes back empty and `Slides.Count` reads 0. Idle, it
answers normally. That distinguishes "waiting for you" from "finished" in one call, and it
was the only thing that settled it.

**Fix:** activate the PowerPoint window immediately before showing any prompt
(`AppActivate`, or `Application.Activate`). A question nobody can see is not a question.

**Do NOT fix by adding waits.** Rohan asked whether the code should "cycle through
applications to allow adequate initialisation time". Tempting and wrong: nothing here is
an initialisation problem — Excel was doing genuine work, building 13 drafting sheets of
~43 rows. A sleep long enough to help is wrong on a faster machine and still wrong on a
slower one, and it would mask the real defect. Zettel
`20260810-compensating-arithmetic-hides-the-mechanism-you-never-probed`.

**Second half of the same defect: nothing says "working".** A long Excel operation with a
silent PowerPoint is indistinguishable from a crash. A status-bar line, or making Excel
visible while it writes, is the honest fix.

### P2. The field-picker InputBox has its text field OFF THE BOTTOM OF THE SCREEN

`AskForField`'s prompt lists every field in the workbook — around 40 lines by the time it
names the drafting fields, then the Given/Derived/Controlled ones with their kinds. On a
1080p screen that pushes the actual entry box below the screen edge. Rohan: **"what box?"**

He could not type into it because he could not see it. This is not a cosmetic problem: the
box returns "" when dismissed, which silently cancels the stage.

**Fix:** cut the prompt to the drafting fields only (the ones that can be answered), and
put the "these do not need a drafting sheet" explanation on a sheet, not in the dialog.
Same lesson as the Sync Now report: the container was the problem, not the wording.

### P3. The 21 `MS*` "fields with nothing to write into" warning is a FALSE POSITIVE

Fires on **every single run**, and the answer is always No. `FieldWiring.ScanFieldWiring`
compares register columns against individually tagged fields and has no concept of a
device consuming a column set, so every device-driven column reads as orphaned.

Answering Yes would walk the person through tagging 21 timeline internals as ordinary
fields — destroying the device.

**Fix is the device registry.** See NEXT-SESSION.md, "A DEVICE REGISTRY". Do not special-
case this one call site.

### P4. The 17-column prompt is ALL-OR-NOTHING across a mixed set

"17 field(s) on the Field Spec have no column in the register. Add a column for each?"
Sixteen are uncontroversial (`INDUSTRY_CASH`, `START_DATE`, `PROJECT_LEAD` …). One is
`HIGHLIGHTS_BODY`, which **must not** get a single column — it is three shapes per slide
and needs slot columns like the milestones.

So a real architectural decision is made, silently, by a Yes on a bundled prompt. Declined
twice on 13 Aug for exactly this reason.

**Fix:** offer the set per field, or exclude fields whose Renders-as implies slots.

### P5. Re-running the Template Audit REPLACES the sheet, decisions included

The dialog says so — *"Re-running this REPLACES that sheet, decisions included"* — which is
honest, and still wrong. The audit's whole purpose is to record field/chrome/drop
decisions against 50 items; losing them on re-run means the work can only ever be done in
one sitting. Same shape as the Discover Fields grid (item 3) which rebuilds from scratch
and loses marks.

**Fix:** carry decisions across by shape ID, the way `WriteDraftingSheet` carries drafts.

### P6. `PROGRESS_BODY` has more approvals than submitted text — CONFIRMED, not a miscount

**Settled 2026-08-13 against the saved file**, after chat side and Claude Code disagreed on
which column held which number. Authoritative counts, header row excluded, 43 data rows:

```
                       E draft   F submit   G approve
TPL_KEY_EVENTS_BODY          1         43          43
TPL_PROGRESS_BODY            0         34          43
TPL_HIGHLIGHTS_BODY         43         43          43
```

Constants, so nobody re-infers the mapping: `COL_D_DRAFT = 5` (E), `COL_D_SUBMIT = 6` (F),
`COL_D_APPROVED = 7` (G), `DRAFT_HEADER_ROW = 9`, `DRAFT_FIRST_ROW = 10`.

**RETRACTED 13 Aug, later the same evening. The count was mine and it was wrong.**

Column G holds `'0'` on at least some rows, not `Y` — `3_P001` on `TPL_PROGRESS_BODY` is
`SUBMIT` empty, `APPROVED = '0'`. My figure counted **non-empty** G cells, so it counted
`'0'` as an approval. Publishing tests `ReviewQueue.IsApprovalMark`, which is an
*affirmative* test, not a non-empty one.

So "43 approvals against 34 texts" is not established. The real approved count is lower and
unknown until someone counts affirmative marks rather than filled cells.

**This is the exact error this file exists to catch, committed by the person writing the
file.** A check that asks "is there something here" cannot answer "does this say yes".
Recount with `IsApprovalMark`'s own rule before treating the mismatch as real. Publishing correctly requires BOTH, so the nine extra
publish nothing — but the reported count will not match the ticks, and there is no message
explaining why. Either the ticks are stale or the text was lost; nothing currently says
which, and it should.

**A counting habit this cost an exchange to learn:** these sheets are rebuilt by any
`Sync Now`, so a count is only true of one moment. Quote the workbook's mtime beside any
figure taken from it — two correct counts taken an hour apart will disagree and look like a
defect.

### Confirmed FIXED 2026-08-13 (addin81), listed so they are not re-found

- **Publish/Copy-AI could only ever reach the FIRST `Kind = Prose` field.** `FieldForRun`
  now asks inside a chain. Two call sites. Proven on the real deck at 17:23.
- **Roll Forward asked a question whose every answer was refused.** `PeriodRowCount` lets
  the caller check first; inside the chain it is now one line in the report.
- **The slide-type picker demanded typed input for a one-item list** (item 7). `PickType`
  auto-selects a sole type and still asks when there is a real choice. Three call sites.
- **`RefreshDraftingSheets` reported "drafting sheets are ready. Workbook saved." over
  seven refusals.** Refusal now leads, names the fields, and carries a warning icon.

---

## FOUND BY PRESSING THE BUTTON — 2026-08-15, ALL THREE STILL LIVE

Every one of these was caught at a dialog on the real deck. **None was found by the
suite**, which went 194/0 to 199/0 across the same session without seeing any of them.

### A. The harvest writes a formatted VIEW into a field whose contract is numeric

**Blocker. Nothing bulk should run before this is fixed.**

`PROJECT_PROGRESS` reads `33%` off the slide. `InjectPrimitive.bas:340` refuses any
non-numeric progress value and names **`'90%'` as wrong in those exact words**. So the
harvest would write a value the tool itself cannot publish.

**Why it is worse than an ordinary bug:** `Harvest.HarvestSlide` writes only where the
register is empty. Once `'33%'` is in the cell it is no longer empty, so a corrected
harvest **cannot overwrite it** — it has to be cleared by hand, exactly like the four
Excel-coerced cells cleared at 23:40 on 14 Aug.

The harvest assumes "what is displayed" is what the register wants. True for prose, names
and dates; false wherever the slide shows a formatted view of a stored value. It already
refuses devices by name — numeric-contract fields need the same treatment: convert
(`33%` -> `0.33`) or refuse, never write the string.

### B. `OfferHarvestForSelectedSlides`'s prompt mislabels and truncates

Two defects in one dialog, both in `RibbonUI.bas`:

1. **All propagation detail is accumulated into the `collisions` string**, so successful
   stamps print underneath a `Refused -- two fields matched one shape:` header. On
   14 Aug 23:57 that made a run reporting 16 correct stamps read as though it had refused
   everything.
2. **It hits `CapReport`'s 900-character cap mid-word**, so collisions can be present and
   invisible. A person cannot consent to what the dialog does not show them.

Neither is dangerous — the guards bound the write, not the text — but this project has
already paid twice for approving a prompt that could not be fully read.

### C. Slide 27 carries a shape already named `Text 216a` that is not the date

The 2026-08-15 rename pass (55 shapes, so 32 of 44 slides carry both `Text 212a` and
`Text 216a`) **refused** slide 27 rather than create a duplicate name. That slide's
`END_DATE` will keep colliding until a human looks at it. Correct behaviour, still an
open item.

### D. `check_vba_static.py`'s reachability check is weaker than its name

It asks whether a procedure's NAME appears in another module, **not** whether anything
reachable calls it. A chain of private orphans is invisible to it. Proven 2026-08-14:
commenting out the call to a wrapper left the callee's name still written inside the
now-orphaned wrapper, and the checker stayed clean.
