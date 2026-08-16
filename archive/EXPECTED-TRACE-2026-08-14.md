# Expected trace — Rohan's imagined use of the tool

**Captured 14 Aug 2026, in his words, BEFORE scenario 5 was run.** Written down first
deliberately: without a recorded prediction, whatever the tool does on the day will look
like what it was supposed to do. This is the prediction. The run is the observation. The
diff is the finding.

Nothing here is a design. It is what he expects, plus what the code does today.

---

## The five steps, as he described them

1. **I have a slide design I want to house similar content.** The add-in converts a *copy*
   of one of my slides into a template — with guidance on deck management and shape
   naming. (He notes the guidance can be built later: *"for now we just both know
   templates at the start are important."*)

2. **Press a button and, because of the nature of the template, the Excel workbook
   appears** — already holding drafting sheets capable of generating the fields those
   slides need. The human's job here is tuning: *"the human can adjust how good the goblin
   is doing the job, this is about making sure he's in the right spot with the right feeds
   and outfeeds."*

3. **A simpler, compiled, clean view to approve the register values that will end up on
   the slides** — *"a verification that the right fields are connected."*

4. **Static/empirical and shape on-off data also live in these sheets**, so the right
   visibility values get set.

5. **Adjust values, hit sync, the deck adjusts for that period.** On a new quarter, *"the
   old quarter files persist as separate files uncluttered by new data, which flows into
   the new period deck on its sync."*

---

## Diff against what the code does today

### GAP 1 — the template does not produce the workbook. Direction is reversed. (step 2)

He expects: template defines the fields → workbook appears with the right drafting sheets.

Today: the workbook is paired to the deck **by hand** (`DraftingUI.RepointWorkbookUI`,
`DeckRegistry.GetWorkbookPath`), the Field Spec is a sheet a person seeds, and discovery
reads the *deck* to find out what fields exist. The template is not authoritative for
anything.

This is the **template-first pivot, decided in principle 13 Aug and never built**
(NEXT-SESSION, "TEMPLATE-FIRST, NOT DISCOVERY-FIRST"). His walkthrough assumes it is
already the architecture. It is not.

### GAP 2 — there are TWO approval gates. He imagines one. (step 3)

He describes a single compiled view where he confirms the right values are connected to
the right fields.

Today there are two, in series, and they are shaped differently:

| gate | scope | what it asks |
|---|---|---|
| **Publish** (`PublishDraftsForField`) | ~~**ONE FIELD** — `FieldForRun` asks which~~ **EVERY field, asks nothing (2026-08-14 night)** | tick column **H** on the drafting sheet (was G under layout 4; layout 5 inserted `REPORTED LAST TIME` at D) |
| **Review** (`ReviewChangesCore` → `WriteQueueSheet`) | compiled, per entity+field (`ChangeHash`) | approve each proposed slide change |

The compiled view he wants **exists** — it is the review queue. What does not exist is
reaching it without first walking a per-field publish, one field at a time. **He never
once mentions choosing a field.** The entire per-field selection UX is absent from his
mental model of the tool.

This is very likely the root of the tick confusion: two approval moments, and he has room
for one.

### GAP 3 — shape visibility is not supported at all. (step 4)

**THIS PARAGRAPH IS WRONG -- corrected 2026-08-14 night.** There IS a `.Visible` write:
`MilestoneDevice.bas:630` sets `shp.Visible` and it works. It is `Private` to that module
and drives the milestone slot circles, whose ON/NOW/OFF triples are already named on the
real template. So GAP 3 is EXPOSING an existing mechanism as a field behaviour, not
building one -- hours, not the "entirely unbuilt" this section claims. The original text
follows.

Searched `InjectPrimitive` and `FieldSpec`: there is **no `.Visible` write anywhere**, and
`Behaviour` is not what he thinks it is. Its only values are:

```
BEHAVIOUR_FILL = "Fill the frame"
BEHAVIOUR_FIT  = "Fit inside"
BEHAVIOUR_ASIS = "Leave as is"
```

Placement of pictures and objects. Nothing to do with on/off. **Step 4 is entirely
unbuilt** — not partially, not nearly. It is the only one of the five with no
corresponding machinery whatsoever.

### GAP 4 — THE BIG ONE. He expects a file per quarter; the tool holds all quarters in one. (step 5)

He said: *"the old quarter files persist as separate files uncluttered by new data, which
flows into the new period deck on its sync."*

Today:

- **One register**, one `Register` sheet, all periods stacked in a `Quarter` column —
  currently 43 rows `Q3F26` + 43 `Q4F26` + 5 `Q1F27` in the same sheet.
- **One deck**, whose `DeckSyncPeriod` custom property is *changed* to move quarters.
- Roll-forward **copies rows within the same workbook** (`ExcelOutput.RollForwardPeriod`).
- Last quarter is preserved by `SAVED *` parked sheets and `backups\` — archives *inside*
  or *beside* the live artifacts, not a frozen file per quarter.

**Why this matters more than the other three.** Nearly all the machinery that has cost
this project since 1 August exists to stop one quarter's data contaminating another
*inside a shared artifact*:

- `periodChanged` / `carryThisRow` in `WriteDraftingSheet`
- the rollover refusal (added 13 Aug, after *"the Q4 text I generated is real?"*)
- the cadence machinery (already suspected dead code)
- `ParkSheetCopy` and the park that does not fire on a matching layout
- the tick-carry fix of 14 Aug

If each quarter is its own file, **that entire problem class does not arise.** You never
decide which rows to clear on rollover, because you never roll a sheet over — you produce
next quarter's file and freeze this one.

**Stated as a hypothesis, not a conclusion.** It needs testing before anything is deleted,
and it has an obvious cost to weigh: cross-quarter reads (trend, "what did we say last
quarter", column C `CURRENT`) become cross-*file* reads. That may be cheap or may be the
whole reason it was built this way. Nobody has checked.

---

## What he did NOT mention, which is also evidence

- **Choosing a field.** Never comes up. See GAP 2.
- **Copy AI → Submit as a separate press.** He folds generation into the drafting sheet
  itself — *"drafting sheets capable of generating the required fields"*.
- **Tagging or marking shapes after the template exists.** For him it is all step 1, under
  "shape naming".
- **Any second press of anything.** His trace is linear: five steps, each once.

---

## What this changes about the plan

The rename (`RefreshDraftingSheets` → `BuildDraftingSheets`) and the button split were
scoped against the *current* architecture. GAP 4 could make some of that work moot. Do not
start the rename until GAP 4 is settled one way or the other.

Scenario 5 still runs first and still runs unchanged — one field, drafting → register →
slide, by button, on `addin84` as it stands. It is now doing double duty: proving the
publish path, and producing the observation that this prediction gets diffed against.
