# NEXT SESSION — start here

> ## DO THIS FIRST — publish `PROGRESS_BODY`
>
> **State at handoff (19:40, 13 Aug), read from files not dialogs:**
> register `register-wide.xlsx` unchanged since **16:57** · deck **19:27** ·
> `addin82` built, stamped `2026-08-13 19:22`, **confirmed loaded and the only add-in**.
>
> **A `PROGRESS_BODY` run was attempted and did nothing.** Excel never opened, the register
> never changed. Most likely cause: **`Cancel` was pressed on the `MS*` dialog**, which is
> the only button there that ends the whole chain and sits beside the `No` you want. Not
> confirmed — the screenshot channel failed (see below) — so treat it as the first thing to
> rule out, not as fact.
>
> **The sequence.** One button, `1. Sync Now`, pressed twice. First press publishes drafting
> to the register and builds the review; you tick the review; second press applies it.
>
> | prompt | answer |
> |---|---|
> | `MS*` fields warning | **No** — NOT Cancel |
> | 5-step plan | Yes |
> | Quarter | OK (already `Q4F26`) |
> | Roll Forward | should not appear — reports instead |
> | 17 Field Spec columns | **No** (`HIGHLIGHTS_BODY` needs slot columns, not one) |
> | Which field? | **`PROGRESS_BODY`** |
>
> **Two traps.** The `MS*` `Cancel` above. And the field dialog's **text box sits below the
> bottom of a 1080p screen** — drag the dialog up by its title bar before typing, or it gets
> dismissed empty and the publish is silently skipped. That is FIX-LIST P2.
>
> **Watch the published count against 34.** 43 rows are ticked, only 34 have text, so 9
> should be skipped. Whether the report *says* so is a live test of FIX-LIST P6.
>
> **Then expect a small diff.** As with Key Events (21 of 43), many slides may already match.
>
> ---
>
> ### Two things that are true but not yet visible
>
> **`addin82`'s tab ordering has not taken effect.** `ArrangeTabs` runs during a drafting
> rebuild, and no full run has completed since the add-in was swapped. Sheet order is still
> `START HERE, Field Spec, Sources, TPL_*, …` with `Register` unplaced. It will reorder on
> the next successful `Sync Now`.
>
> **Slide 1 is inconsistently de-identified, live, right now.** `KEY_EVENTS_BODY` was
> published and reads *"The industry partner's withdrawal halted further development"* — but
> `PROGRESS_BODY` directly below it still reads *"...ceased following Calix withdrawal"*, and
> the partner is also named in the header subtitle and the Project Team box. Publishing
> `PROGRESS_BODY` fixes one of the three. **The other two are untagged shapes** — part of the
> 50 unmanaged items the Template Audit found, so they need tagging, not publishing. This is
> a funder-facing deck; worth treating as content risk rather than tidiness.
>
> ### Practical note: screenshots stopped reaching Claude Code
>
> From ~19:30 no image reached the clipboard and **no PNG was written anywhere under the
> profile**, including the OneDrive `Claude` folder where earlier ones landed. Verbal
> description plus COM state reads worked fine as a substitute. Useful calibrated check,
> worth keeping: **while a modal is open PowerPoint stops answering COM** —
> `ActivePresentation.Name` returns empty — and answers normally when idle. That single call
> distinguishes "waiting for you behind a window" from "the run has ended".


**Written 13 August 2026, ~16:15.** Previous version archived as `NEXT-SESSION-2026-08-12.md`.

> ## THE DECK IS ONBOARDED. THE DELIVERY COUNT IS STILL ZERO.
>
> 43 slides tagged and linked, 0 failed verification, register saved, deck period
> `Q4F26` on disk. **No drafted value has reached a slide.**
>
> **The reason was found tonight, and it is one line of code.**

---

## THE BLOCKER, AND THE TWO-ROW WORKAROUND

`RibbonUI.SyncNowChainCore` step 4 is `DraftingUI.PublishDraftsForField`, which begins:

```vba
fieldId = ActiveDraftField(wb)          ' whatever TPL_ sheet is ACTIVE in Excel
If fieldId = "" Then fieldId = AskForField(CAP, wb)
```

It only asks **if** the active sheet is not a drafting sheet. But the step immediately
before it — `RefreshDraftingSheets` — ends with `ShowSheet wb, firstSheet`, and
`firstSheet` is the **first `Kind = Prose` row on the Field Spec sheet**
(`DraftingUI.ProseFields`, row order).

That row is `ABOUT_BODY`, which has **0 submitted, 0 approved**. So every run publishes
an empty sheet, reports "0 would be published", and finishes quietly. **The chain cannot
reach any other field.** There is no field picker on the toolbar (two buttons only), so
this is the only publish route.

This is very likely a large part of why this project has never got a field onto a slide.

**A ROW-REORDER WORKAROUND WAS PROPOSED AND REJECTED — DO NOT USE IT.** Moving
`KEY_EVENTS_BODY` above `ABOUT_BODY` on the Field Spec would work, because
`FieldSpec.WriteSpecSheet` only seeds *missing* FieldIDs and never reorders existing rows.
It is still the wrong move, and Rohan stopped it with one question: *"Why are you having
to move register rows manually? Worries me that the code won't work when it needs to."*

He is right, for three reasons:

1. **It makes which field reaches a slide depend on spreadsheet row order** — invisible,
   unstated coupling of exactly the kind that has bitten this project repeatedly.
2. **It is not available at work.** No Claude, no WSL, no Python there — a quarter has to
   be runnable from toolbar buttons. "Reorder rows in a spec sheet so the right field
   publishes" is not a procedure; it is a defect with instructions attached.
3. **It would have hidden the defect behind a successful-looking run**, which is the
   failure mode this project keeps rediscovering.

**FIXED in `addin81`** (build stamp `2026-08-13 16:24`). New `DraftingUI.FieldForRun`:
inside a collected chain it ASKS; standalone it still reads the active sheet, because
there the answer really is on screen. Asked ONCE per run and reused, so the two stages
that need it do not ask twice.

**It had TWO call sites.** `CopyAiDraftsToSubmit` carried the identical line and the
identical consequence — fixed together rather than only where it was noticed.

**PROVEN 2026-08-13 17:23 on the real deck.** 21 KEY_EVENTS_BODY values written and confirmed by reading slide XML from the saved file -- slide 1 now carries the drafted "The industry partner's withdrawal..." wording. The delivery count is no longer zero. Previously read: 192 tests pass and the project compiles, but no test exercises the
chain's field selection — which is precisely the gap that allowed this defect. Green here
means "nothing broke", not "the fix works". Prove it by pressing the button: `Sync Now`
must now ASK which field, and `KEY_EVENTS_BODY` must be selectable.

**Note what the test suite did NOT do here.** 192 tests pass. Not one of them asks "can a
person cause `KEY_EVENTS_BODY` to be published?" — they test that publishing works when
called, not that the chain can reach it. Same "tested unit behind a locked door" shape as
the picture injection and the progress bars, found the same way: by pressing the button.

---

## STATE, VERIFIED FROM FILES (not from dialogs)

- **Deck** `OneDrive\Claude\3. Project Progress.pptx` — 44 slides, 49,247,250 bytes.
  `DeckSyncPeriod = Q4F26` confirmed by property name in `docProps/custom.xml`.
  Slide 44 is the hidden master template, 9 fields set to `<<placeholders>>`.
- **Register** `OneDrive\Claude\register-wide.xlsx` — 308,072 bytes, `Register` sheet has
  92 rows (1 header + 91: 43 Q3F26 + 43 Q4F26 + 5 Q1F27). All 43 Q4F26 instance keys
  match the Q3F26 keys exactly — **the handover's "stale/foreign Q4F26 rows" warning was
  wrong**, they describe the same slides.
- **Backups** `OneDrive\Claude\backups\2026-08-13-1520-post-onboard-Q3F26 - *` — deck and
  register, both verified byte-identical by md5 at the time of copy.
- **Build `addin80`**, stamp `2026-08-13 14:37`, in `OneDrive\Claude\` and in the trusted
  location `AppData\Roaming\Microsoft\AddIns\`.

### Drafting sheets — real counts (header row EXCLUDED)

| sheet | submitted | approved |
|---|---|---|
| `TPL_KEY_EVENTS_BODY` | 43 | 39 |
| `TPL_PROGRESS_BODY` | 34 | **42** |
| `TPL_HIGHLIGHTS_BODY` | 43 | 42 |
| `TPL_ABOUT_BODY` | 0 | 0 |
| `TPL_STRATEGIC_ALIGNMENT_BODY`, `TPL_PROBLEM_BODY`, `TPL_STRATEGIC_LINKAGES` | 0 | 0 |

`PROGRESS_BODY` has **more approvals than submitted text** (42 vs 34). Those 8 rows
publish nothing — both text and tick are required — but the count will look wrong.

`HIGHLIGHTS_BODY` has the most work in it and **cannot publish**: it is not one of the
nine tagged fields and needs slot columns, not one column. See FIX-LIST.

---

## ENVIRONMENT FINDING — SAVES AND ONEDRIVE

Both files are open via **OneDrive URLs**, not local paths:
`https://d.docs.live.net/96b9ec593ee3ba55/Claude/…`

With AutoSave **off**, the deck period write failed **4 verified attempts**, and a manual
`Ctrl+S` did not change the file's mtime either. `SetDeckPeriodVerified` correctly
detected this and refused to continue:

> `THE PERIOD DID NOT REACH THE FILE after 4 attempt(s). Asked for: Q4F26  On disk: Q3F26`

**Turning AutoSave ON made the write land.** This is an environment condition, not a code
defect — and it is the configuration the work machine will be in. The tool's behaviour
here was correct and is what a week ago was missing: it checked the file, not its own
cache, and refused rather than reporting success.

---

## WHAT SHIPPED TONIGHT

- **Suite green: 192 passed, 0 failed** (was 190/2), behind the compile gate.
- The two tests asserting the deleted defect were rewritten **and renamed**, because the
  old names stated the defect as the requirement:
  - `Drafting_PeriodRolloverDropsStaleSubmit` → `Drafting_RolloverRebuildsOnlyWhenNothingIsAtRisk`
  - `Drafting_RolloverKeepsEntityStaticRows` → `Drafting_RolloverCadenceGovernsUntypedRows`
- **`RefreshDraftingSheets` no longer reports success over a refusal.** It collected
  refusals into the Run Log and then said *"drafting sheets are ready. Workbook saved."*
  with an information icon. Now: refusal first (so MsgBox truncation eats the guidance,
  not the warning), refused field names listed, warning icon.

### NOT YET TRUSTED

**Neither rewritten test has been made to fail on purpose.** Green alone is not evidence.
Break each before relying on it — for `...RolloverCadenceGovernsUntypedRows`, put SUBMIT
text back on the fixture and it should stop testing anything, because the refusal
pre-empts the whole path.

---

## FILES CHANGED THIS SESSION (repo `deck-sync-refimpl`, uncommitted)

- `vba/DraftingUI.bas` — refusal count/names surfaced in the dialog; warning icon
- `vba/tests/TestRunner.bas` — two tests rewritten + renamed, runner registrations updated
- `FIX-LIST.md` — new items **1c** and **1d**
- `NEXT-SESSION.md` — this file

Not committed. Nothing else in the repo was touched.

---

## OPEN, IN PRIORITY ORDER

1. **Publish one field.** Field Spec row move → `Sync Now` → `KEY_EVENTS_BODY` → review →
   apply. Then verify by reading the slide XML out of the saved deck, not the dialog.
2. **Fix the publish-target defect properly** (above), then revert the row move.
3. **FIX-LIST 1c/1d** — at-risk scan misses SOURCES/NOTES; the park that reports "nothing
   was lost" runs *after* `ws.Cells.Clear`. Fixing 1c makes 1d unreachable.
4. **The cadence machinery is probably dead code.** The refusal pre-empts it; it now
   governs only SOURCES/NOTES on untyped rows. If 1c is fixed, delete it rather than
   maintain it. Rohan's call.
5. **`MILESTONE_TIMELINE` group tagging is UNVERIFIED.** It was not among the nine tagged
   fields. If the timeline renders blank after a sync, check this first.
6. **Field Spec `Kind` values look wrong for the milestones**: `MS1_LABEL`/`MS7_LABEL` are
   `Given` while `MS2`–`MS6_LABEL` are `Prose`; MS1/MS7 DATE+DONE are `Derived` while
   MS2–MS6 are `Given`. 13 fields are `Prose` but only 7 have drafting sheets, so the next
   refresh will create six more tabs.
7. Slide 44 still carries P001's unmanaged content (figures, photo, team). The audit found
   **50 unmanaged text items on slide 1, 21 of which look like project data** — that is the
   next tagging backlog, and the same set the Field Spec wants columns for.

---

## THREE THINGS WORTH KEEPING

**A warning that only reaches the log is not a warning.** The refusal guard was correct
and invisible; the dialog said "ready" over seven refused sheets. Fixed, but the shape
recurs — check where a message *lands*, not just that it exists.

**The check that found the save failure was the one that read the file.** Four
in-process attempts all "succeeded". Only comparing against `docProps/custom.xml` on disk
told the truth. Evidence must come from the far side of the boundary.

**"Nothing happened" meant a dialog behind the window — twice.** A VBA modal can open
behind PowerPoint. Before diagnosing a dead button, Alt+Tab. A calibrated test:
PowerPoint stops answering COM (`ActivePresentation.Name` comes back empty) while a modal
is open, and answers normally when idle.

---

## ARCHITECTURE — DECIDED IN PRINCIPLE 2026-08-13, NOT YET BUILT

Two calls made at the end of the first successful publish. Both are Rohan's, both are
right, and both should be settled properly before more feature work.

### 1. TEMPLATE-FIRST, NOT DISCOVERY-FIRST

Rohan: *"Are we better off gearing it to be an expert template builder and pushing a
pattern we know? ... maybe that's best rather than a sensory beast that doesn't quite
know what it is trying to be."*

**Yes.** The strongest argument is the WORK MACHINE. A sensory tool needs an operator with
judgement at every step — tonight it needed Claude to decide which of 59 shapes were
fields, whether the `MS*` warning was real, which register to pair, whether `TESTFILL` was
junk, and whether 88 proposed changes were safe. At work there is no Claude. A
template-driven tool needs no run-time judgement, because the judgement was made once, in
advance, and frozen into an artifact.

**The cost asymmetry says the same.** Discovery runs ONCE per slide type, ever. Publishing
runs 43 slides x 9 fields x 4 quarters, forever. Nearly all the code, nearly all the
defects and nearly all of 13 Aug went into the once-ever path.

**And look where the defects actually were:** the grid that loses marks, the blank grid
that unmarks, the 21 `MS*` false positive, 50 "unmanaged" items it can only guess at,
"which register?", the pairing. Every one is a PERCEPTION defect. The publish path — the
thing that runs every quarter — had one defect, four lines long.

**The mechanical diagnosis under "doesn't know what it's trying to be":** the tool holds
THREE sources of truth about what a field is — a tagged shape, a register column, and a
Field Spec row — and reconciles them at run time. Every reconciliation is a place to be
wrong. A template collapses all three into one, decided at design time.

**What that means concretely**
- The template is the authority; register schema derives from it.
- `FieldWiring`'s orphan-column question dissolves: nothing can be orphaned if the
  template defines the set.
- Discovery is DEMOTED to a one-off migration tool. Run once per deck, then never
  developed again. It has already been run — 13 Aug — so for this deck it is done.
- New projects clone slide 44 and are conformant by construction.

**What NOT to throw away:** the verification discipline (the file is the evidence; prove a
check can fail; a defect is a class), the Office/COM/OOXML knowledge, the consent-gate
design, and the register-deck contract. None of it is discovery-specific; all of it
transfers to the next project.

**The honest risk:** real slides are messier than any taxonomy — the finding that made
discovery seem necessary in the first place. But that now cuts the other way: owning the
template BOUNDS the mess instead of trying to perceive it, and the messy migration has
already happened.

### 2. A DEVICE REGISTRY — PROTECT COMPOUND SHAPES FROM THE GENERIC MACHINERY

Rohan: *"I can already see the need for a specialist module spot to protect complex shape
mechanisms like the timeline being pulled into marking etc."*

Three pieces of evidence from one evening:

- `MS1_DATE` … `MS7_LABEL` appeared as 21 rows in the Discover Fields grid, and the only
  thing that stopped them being tagged was **Claude telling Rohan not to**.
- `FieldWiring.ScanFieldWiring` reported those same 21 columns as orphaned on EVERY run —
  the recurring "21 field(s) on the register that no slide carries" warning.
- The Template Audit counted device internals among its "50 unmanaged text items, 21 of
  which look like project data".

Three generic mechanisms seeing the device's PARTS; none of them seeing the device.

**The model already exists, in exactly one place.** Injection has it right:

```vba
InjectPrimitive.InjectField(sld, "MILESTONE_TIMELINE", "", False, Nothing, row)
```

One addressable thing, consuming its own columns off the register row. That understanding
never propagated to discovery, wiring, marking or audit — so a device is a first-class
citizen at write time and a pile of loose shapes everywhere else.

**Principle: the device is the unit of addressing, not its parts.**

**Shape of the fix.** One declaration per device — its role tag, the register columns it
consumes, its internal shape-name pattern — read by four consumers:

| consumer | today | with the registry |
|---|---|---|
| Discovery | lists 21 internals as candidate fields | skips anything inside a declared device |
| `FieldWiring` | reports 21 orphan columns every run | counts them as OWNED by the device |
| Marking | will happily tag a device internal | refuses |
| Template Audit | counts internals as unmanaged project data | classifies as device internals |

One declaration, four consumers — versus four independent special cases, which is what
would get written if this is approached site by site.

**Why it is urgent rather than tidy:** "leave rows 41-65 alone" was ADVICE GIVEN TO A
HUMAN. Zettel `20260719-telling-an-agent-not-to-do-something-isnt-a-control` — and it is
not a control when you tell a person either. At work, with nobody to say it, the next run
of Discover Fields tags 21 timeline shapes as individual fields and quietly destroys the
device.

**It folds into the template pivot.** In a template-owned world the device is PART of the
template, so it is declared by construction rather than looked up — the registry stops
being a side-table and becomes a property of the thing already controlled. The expensive
half, knowing what a device is, is already written.

### 3. ALSO REQUESTED, NOT STARTED

**Logical tab numbering and Excel best practice across all workbooks.** `register-wide.xlsx`
has 18 sheets in arrival order with no scheme (`START HERE`, `Sources`, `SRC_EXTRACTS`,
`Field Spec`, seven `TPL_*`, `Register`, `Run Log`, `Sync Log`, `Field Discovery`,
`Template Audit`, a stale `Review project-status-2D3D`, plus the live review sheet).
`WorkbookBridge.ArrangeTabs` already orders them, so the scheme belongs there rather than
in a manual pass. Do this FIRST next session — it is small, bounded, and was explicitly
asked for.

---

## EXCEL TAB ORDER — DONE BY POSITION, DEFERRED BY NAME (13 Aug, late)

Rohan asked for logical tab numbering and Excel best practice across the workbooks, then
added: **"anything that threatens the data chain fix it"** and **"if you are going to
renumber use logic"**. Both shaped what was and was not done.

### Done: a data-chain fix found while auditing the sheet names

**`"Sync Log"` was a bare literal in SEVEN places across two modules** — alone among the
tool-owned sheets, every one of which otherwise has a constant. Two of the seven are
`GetOrAddWorksheet` calls, which **create** the sheet when the name does not match. So one
divergent literal would not fail loudly: it would quietly open a second log sheet while
`IsToolOwnedSheet` and `ArrangeTabs` went on guarding the first, splitting the audit trail
while looking healthy. Now `WorkbookBridge.SYNC_LOG_SHEET_NAME`, all seven replaced.

### Done: tab order follows the lifecycle of a quarter

`ArrangeTabs` now orders by position — no renaming, so nothing can break a lookup:

```
 1  START HERE     where a person begins; the readiness checklist
 2  Field Spec     what fields exist at all -- configuration before data
 3  Sources        the evidence values may cite
 4  SRC_*          harvested source data
 5  Register       THE DATA. The reason the workbook exists.
 6  TPL_*          where a person works, in Field Spec order
 7  Review *       the approval gate, between work and the deck
 8  Field Discovery, Template Audit    diagnostics, off the normal path
 9  Run Log, Sync Log                  the audit trail
10  SAVED *        parked archives, absolutely last
```

**`Register` was previously unplaced** — it fell into "everything else in its current
order" beside the diagnostics, so the most important sheet in the workbook sat wherever it
happened to land. That, rather than any cosmetic gain, is what this fixes. `SAVED *`
archives now sort last so they cannot be mistaken for the live sheet they were copied from;
typing in one is silent, because publish reads the live sheet only.

### Compile-verified 13 Aug, after Rohan closed Office

```
COMPILE OK: whole project compiled clean (33 modules).
RESULT: OK
=== 192 passed, 0 failed ===
```

Static checks also clean across 34 modules. The new cross-module references
(`DiscoverUI.DISCOVERY_SHEET_NAME` and `TemplateAudit.AUDIT_SHEET_NAME` read from
`WorkbookBridge`) compile.

**NOT YET IN AN ADD-IN.** The ordering and the `SYNC_LOG_SHEET_NAME` constant are in source
only — `addin81` predates them. Nothing changes in the workbook until `addin82` is built
and installed, so do not expect to see the new tab order before then.

### Deferred deliberately: numbering the NAMES

`01_FIELD_SPEC`-style names would be better still, and are a **migration, not a tidy-up**:
sheet names are this tool's addressing mechanism — nine constants, dozens of literals, plus
the `TPL_`, `Review ` and `SAVED ` prefix matches. It needs the constants changed, every
literal found, and a rename pass over a live workbook holding drafted work, with a fallback
for a workbook that has not been migrated yet.

**If it is done, the scheme should be the ordering above**, so position and name agree and
neither can drift from the other. Do it as its own session with the suite green before and
after — not alongside anything else.

---

## REQUESTED 13 Aug: a standard placeholder for projects that have stopped reporting

Rohan: *"if it's missing because they have stopped reporting please include a standard
placeholder for now."*

**The need is real.** `3_P001` is `Project Closed` and its `PROGRESS_BODY` drafting row is
`SUBMIT` empty / `APPROVED = '0'`. A closed project should not silently render last
quarter's prose or an empty box.

**Proposed standard wording:**

> `No update this quarter — project closed Q3F26. Last reported Q1F26.`

States that the absence is intentional, why, and where the last real content is.

**THE TRAP, and why this was not just typed in.** A placeholder typed into `SUBMIT`
publishes as ordinary content and is then indistinguishable from drafted prose — next
quarter nobody can tell which rows are real. That is exactly the `TESTFILL-1256` string
found in the Q4F26 register tonight, which was one bundled "Yes" from reaching a slide.

**So it should be DERIVED, not typed.** The register already knows `PROJECT_STATUS` and the
period. A placeholder generated at publish time for rows where status is closed/not-started
AND the field is empty is:
- consistent by construction, so the wording cannot drift between projects;
- distinguishable from human prose, so a later review can find every one;
- self-correcting — the moment real text is drafted it takes precedence.

Typing it into 43 drafting rows gets the same pixels this quarter and a mess next quarter.

**Open decisions for Rohan:**
1. Which statuses trigger it — `Project Closed` only, or `Not Started` too?
2. Does it apply to every prose field, or only the quarterly ones (`PROGRESS_BODY`,
   `KEY_EVENTS_BODY`) and not entity-static ones like `ABOUT_BODY`?
3. Exact wording, including whether to name the closing quarter dynamically.

**Do not bulk-type placeholders into the drafting sheets before deciding 1–3.** Undoing 43
rows of typed placeholder is harder than generating it once.
