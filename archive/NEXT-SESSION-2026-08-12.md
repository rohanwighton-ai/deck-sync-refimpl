# Next session — inbound from the chat side, 11 Aug 2026 evening

Full text: `OneDrive\Claude\Two messages — they need different.txt`. Read it first.

## Workbook state (done by Rohan + chat Claude, not by this side)

Field Spec is **32 rows, complete** — every Q4F26 field has a recipe, column J
(placement/constraints) filled from the real `1_P006` slide. `SRC_EXTRACTS` replaced the
per-period staging sheet — one permanent sheet, `Applies to` looked up from Sources rather
than typed twice. All 7 `TPL_` sheets rewritten to point at it.

**Cadence is now readable: Field Spec column L, `Per-period` / `Standing`, on every row.**

## Four items, in their priority order

1. **`TOTAL_VALUE` alarm.** Store all four money fields; block publication when
   `TOTAL_VALUE <> INDUSTRY_CASH + SAAFE_CASH + TOTAL_INKIND`. Live slide is out by $646
   and shipped that way. *Check first whether the register stores rounded display values —
   exact equality against rounded inputs would fail permanently, which is the always-firing
   warning that stops being read.*
2. **Linkage-subset check.** `STRATEGIC_ALIGNMENT_BODY` may cite a subset of the codes in
   `STRATEGIC_LINKAGES`, never a code not declared. Extract from both, report the
   difference. Copilot cannot self-check this — the declared codes are on another sheet.
3. **`Kind = Derived`** as a fourth value, plus a `Derivation` column carrying the rule the
   way `Voice`/`Length` do for Prose. Retrofit onto `TOTAL_VALUE`, elapsed-time %, the
   current-milestone marker. First new fields under it: `SUBTITLE_A`/`SUBTITLE_B`.
4. **Wire the orphaned `cadence` parameter** in `Drafting.WriteDraftingSheet` to column L,
   or delete it. It currently reads the retired `Quarter = ALL` sentinel and falls through
   to "unknown" on every field, silently.

## The question asked of this side, answered

> Should elapsed-time-% and the milestone-marker get real FieldIDs?

**No register column. Yes to a Field Spec row.**

A FieldID that implies a register column gives a derived value somewhere to be *stored*,
and a stored copy of a computed value is the drift this project already designed out —
elapsed time must come from `START_DATE`/`END_DATE`, or it goes stale while the dates
beside it stay right.

But the *rule* needs a home, and `Kind = Derived` with a `Derivation` column is exactly
right for it. So: Field Spec row yes, register column no.

**The consequence that must be handled at the same time:** this breaks the bidirectional
completeness check specified in `COLUMNS.md` — every register field column has a Field Spec
row *and vice versa*. Derived rows have no register column **by design**, so without a
carve-out the check reports them as orphans forever. A warning that always fires stops
being read, and that is the failure mode this repo keeps paying for. The check must treat
`Kind = Derived` as "no register column expected", and report a Derived field that *does*
have a register column — that is the real defect worth catching.

## Also open

Three Sources paths still `[TBC]` (S04/S05/S06); `SUBTITLE` derivation awaiting approval;
Copilot investigating authoritative sources for `PROJECT_STATUS`, `SECTOR`, `TRL`,
`START_DATE`, `END_DATE`, `ABOUT_BODY`, `PROBLEM_BODY`.

## Live-deck defects found on `1_P006` (data, not spec)

Money boxes don't sum ($646 out) · `PROJECT_STATUS` renders lowercase against a title-case
controlled vocabulary · two different end dates shown (header vs timeline) · elapsed-time %
matches neither plausible end date.

## Left broken by this side

Two edits to `WorkbookBridge.DescribeSheet` and its test (deriving `START HERE`'s column
letters from constants) **fail the compile gate.** Static checks passed across 34 modules;
compile failed; no tests ran. Either fix or revert before anything else — the gate blocks
the whole suite. PowerPoint names the offending line when the gate runs.

## Late additions, 11 Aug 17:56 (workbook sweep, chat side)

5. **Chars columns must be written as FORMULAS.** `TPL_` columns H and I were static
   numbers — H frozen at a past value, I blank on every row of every sheet — so
   length-against-target has never been checkable. The sheet builder must write
   `=LEN(C{row})` and `=LEN(F{row})`, never a computed literal, or the next rebuild
   silently reintroduces it. This is `Drafting.WriteDraftingSheet`'s job.
6. **`SRC_EXTRACTS` lookup formulas were hardcoded to the current row count** and have been
   widened by hand. Anything that regenerates that sheet must match the wider range.

## CORRECTION to the `PROJECT_STATUS` casing item

All 91 register rows checked and correctly title-cased (56 `In Progress` / 25
`Not Started` / 10 `Project Closed`). So the lowercase `in progress` on the live `1_P006`
slide is NOT a data problem, and the earlier framing of it ("one of the two is wrong") was
wrong.

**Do not go straight to the injection code.** The likelier explanation is that
`PROJECT_STATUS` has simply never been synced to that slide, so it still shows hand-typed
original text — plenty on this deck is unsynced. Settle it the cheap way first: run
Preview Sync and see whether `PROJECT_STATUS` on `1_P006` appears as a pending change. If
it does, there is no bug. Only if it reports the field as already matching is there
something wrong in the write path.

## DECISION, 12 Aug evening — pre-placed shapes + visibility, NOT computed sizing

Rohan's call, and it holds across slide types: **positions are pre-drawn and known; state is
shown by hiding and showing, never by resizing at run time.** His reasoning is the one he
gave when the device was designed — the positioning and z-order are exact and must not be
disturbed. A computed size can drift; a hidden shape cannot.

Colour is still worth applying from a DECLARED spec (the locked palette by asset type,
traffic-light by status never by raw percentage). Applying a stated rule is not the same as
inventing formatting, which is what the module's "does not invent formatting" rule guards
against. **Size stays visibility-driven.**

**This generalises beyond the timeline, and is why the deck should NOT split into three
slide types.** Every type difference found so far is the same shape: K has 6 timeline
circles vs P's 7; K's subtitle carries no Sector or TRL segment; S has a fourth team row
(Research Supervisor). All of it is "pre-place everything, hide what does not apply" — one
template, not three.

### The consequence, and it is the build job

**Nothing propagates template shapes to existing slides.** Searched: no mechanism exists.
The template is cloned only for a NEW project (a register row with no slide). So pre-drawn
variants reach the 43 existing slides only by hand.

**So SHAPE PROPAGATION is now the highest-value build item for the timeline**, and it is the
same mechanism as the missing name-propagation step. One build solves both:

- copy a named shape set from the template to existing slides, positions preserved
- or push names onto shapes already present, matched positionally

Until it exists, the pre-drawn model costs 43 slides of hand-work, which is the thing that
makes or breaks whether the timeline ships at all.

## ARCHITECTURE, settled 13 Aug — central register, drafting per deck

**Adopted 13 August** (Rohan's, and neither Claude proposed it): it dissolves the
rebuild-safety problem rather than guarding against it, and yields the provenance trail as a
by-product.

| Central workbook | Per-period workbook (beside the deck) |
|---|---|
| `Register` — all periods | `TPL_*` drafting sheets |
| `Field Spec` — the recipes | `SRC_EXTRACTS` — **a derived COPY** |
| `Sources` — the bibliography | |
| `SRC_EXTRACTS` master — accumulating, period-tagged | |
| `SRC_MILESTONES` | |

**Only drafting travels with the deck.** A new period means a new deck and a new drafting
workbook, so nothing is ever rebuilt over live content. The guard is not needed because the
failure cannot occur.

**Safety condition, reworded from the period-folder version:** *read during its own period,
never after.* Once a period closes, its workbook is a record and nothing touches it.

### Why `SRC_EXTRACTS` is BOTH central and copied

Two constraints that are not the same, and the distinction is the whole design:

- **The generative step has a LOCALITY requirement.** Copilot cannot follow a pointer into
  another file — that is why `SRC_EXTRACTS` exists at all. Put the extracts in a different
  workbook from the drafting sheet and the original problem returns.
- **The mechanical step does not.** VBA opens arbitrary workbooks by path
  (`WorkbookBridge.OpenOrGetWorkbook`). Register-to-slide can span files freely.

So the master accumulates centrally — **a standing source is fetched once, ever** — and the
relevant rows (standing plus this period's) are **copied into each period's drafting
workbook at creation**. No re-fetch, and Copilot sees everything in the file it is already
in.

**The copy is DERIVED, never authoritative** — same status as column C on a drafting sheet.
New source text goes into the central master and is re-copied. Otherwise there are two
places holding evidence and no rule for which wins.

### Rejected on the way here, with reasons

- **Per-period `SRC_EXTRACTS`** — five of eight sources are `All periods`, and
  `SRC_MILESTONES` feeds Standing fields (`MSn_LABEL` / `MSn_DATE`). Per-period placement
  means re-pasting 370 rows of standing data every quarter to feed fields that never change.
- **Splitting `SRC_EXTRACTS` by cadence across two files** — a file boundary running through
  the middle of one sheet is a boundary nobody can hold in their head, and the column
  deciding which side a row falls on will be got wrong.
- **`Sources` per-period** — a source cited in two periods would exist twice with two IDs.
  The ID-reuse problem in a new costume.

### CORRECTION to an answer given 13 Aug morning

I said onboarding creates the register column for a `Given` field. True for an ordinary one
like `SECTOR`; **wrong for every field the timeline needs.** Onboarding harvests from TAGGED
shapes, and `MSn_DATE` / `MSn_DONE` are names inside one tagged group, never tagged
themselves. So no `MSn_*` column is ever created.

**Fix: the completeness check must OFFER TO CREATE, not merely report.** Give it the ability
to add missing register columns from the Field Spec and all ~30 affected fields are solved
at once. Build it with the pairing work — both touch how the register gets written.

---

# 13 AUGUST — WHERE THINGS ACTUALLY STAND

## The real deck is now `OneDrive\Claude\3. Project Progress.pptx`

Rohan's ruling, 13 Aug. His original at work is SUPERSEDED and must not be edited --
this copy carries the timeline naming and nothing else does. Deck and register both
live in that folder, so the pairing's sibling-fallback makes it work from either machine.

Backups before onboarding: `OneDrive\Claude\backups\2026-08-13-0956-pre-onboard - *`.

## Deck state, verified from the file

- **43 slides**, one layout, 138 shapes / 17 groups / 8 pictures each
- **Virgin**: no deck-sync properties, ZERO tag parts. Nothing to migrate.
- Deck and register agree EXACTLY: 43 codes each, none on either side alone
- Three duplicate slides (`3_P002`, `2_P004`, `1_P006`) found and deleted by Rohan

## The timeline on slide 1 is DONE

`MILESTONE_TIMELINE`, **37 shapes, ONE level, no duplicate names, nothing nested.**
Seven slots x (`_NOW` / `_ON` / `_OFF` / `_LABEL` / `_DATE`) plus `MS_TRACK`, `MS_BAR`.

Colours, as authored by Rohan (**Office returns fill as BGR, not RGB** -- misreading this
nearly produced 14 plausibly-wrong circles):

| State | RGB |
|---|---|
| current / achieved | `003C23` dark green |
| achieved, earlier | `005832` mid green |
| not achieved | `93DCDC` light teal |

**It will LOOK wrong until first sync** -- all three circles per slot are visible at once
until the device hides the inactive ones.

## Device changes shipped today

`MSn_NOW` (four states, exactly one circle visible per used slot) - track shortens to the
last USED slot (the bar already reached the last ACHIEVED one) - `SlotCount` accepts any
circle, not just `_ON` - **the integrity check now verifies circles at all**, which it
never did.

## THE TIMELINE IS BLOCKED ON EXCEL, AND NOTHING BUILDS IT

The register has **no `MS*_LABEL` / `_DATE` / `_DONE` columns**, and no existing path
creates them:

- roll forward copies existing columns -- nothing to copy
- sync READS the register
- onboarding harvests TAGGED shapes; timeline parts are named-inside-a-tagged-group
- publish would create `MSn_LABEL` (Prose, has a drafting sheet) but never `_DATE`/`_DONE`
  (`Given`, never published)

**NEXT BUILD: the completeness check must OFFER TO CREATE missing register columns from
the Field Spec.** Field Spec has all 21 `MS*` rows; the register has none. Same operation
fixes ~30 other `Given` fields.

## Also shipped today (188 tests green, compile + static clean)

- **The pairing fix** -- onboarding no longer invents an empty register beside a populated
  one. `ExcelOutput.RegisterShapedSheets` finds existing registers by scanning the HEADER
  ROW for `Instance ID` (not A1 -- a test fixture caught that); more than one is REFUSED.
- **The re-onboard guard** -- keys off "has this deck been onboarded before"
  (`SlidesCarryingASlideType`), not a value comparison, because on a first onboard the
  register and slides legitimately differ and a comparison would fire loudest when the
  harvest is correct.
- **Quarter-before-onboarding** -- `StartQuarter` now runs inside the setup branch. A
  virgin deck could not reach it, so setup walked the whole marking grid and then hit a
  raw error at the commit with slides already tagged. `WORKFLOW.md` flagged this ordering
  problem on 2026-08-04; it broke when the toolbar collapsed to two buttons and no test
  onboards a virgin deck.

## Current build: `addin75`

## Two mistakes worth not repeating

**COM said the timeline group had no nested sub-groups. The XML said it had six.** I
believed COM, ungrouped one level, stranded six shapes and saved it. Rohan's fix was
better than my plan: ungroup repeatedly until a pass frees nothing, then regroup at ONE
level. Depth discovered, not assumed.

**Office returns `Fill.ForeColor.RGB` as BGR.** Caught only because a screenshot showed
teal where the number said khaki.
