# Milestone percentage display — design

> **DESIGNED, NOT YET BUILT.** Written 2026-08-22, from Rohan's decision the same
> session ("option 2 as ultimately... this is where Research Managers give an
> informed opinion of where the project is at so the newest achieved milestone
> can be the default 'current achieved' big circle position"). Geometry facts
> below are read from the real deck's exemplar slides (44/46/47); see the
> caveat under "DATE geometry" — one shape's numbers didn't parse cleanly and
> are flagged, not guessed.

## The problem

The milestone timeline (`MilestoneDevice.bas`) shows exactly one state per slot:
achieved-and-current (big circle), achieved-earlier (small circle, one colour),
not-achieved (small circle, another colour). A project that's finished MS2 and
MS3, has done nothing visible on MS4/MS5, and finished MS6 renders as
"done, done, gap, gap, done" — accurate, but reads as messy or wrong at a
glance (`2_P009`, this session).

Rohan's read: the binary is honest but throws away real information — a
not-yet-done milestone might be 75% along, not 0%. His fix: **let the
Research Manager say so**, next to the circle, without changing which circle
is "current."

## What does NOT change

- `DrawMilestones`'s "current" logic — `lastAchieved` = the highest-numbered
  slot with `DONE="Y"` — is untouched. This spec was preceded by a proposal to
  cap "current" at the last *contiguous* done slot; Rohan declined it
  explicitly in favour of percentages carrying the nuance instead.
- The four-state model (`_NOW` / `_ON` / `_OFF` / unused) stays exactly as is.
- The module's hard rule stands: **never creates, moves, resizes, or reorders
  a shape.** Whatever this feature does, it does through visibility and text,
  like everything else in this device.

## Data model

New register column(s), one per slot: **`MS<n>_PCT`** (`MS1_PCT` .. `MS7_PCT`),
same naming convention as `_LABEL`/`_DATE`/`_DONE` — deliberately, since
`ColumnFor(i, part)` builds shape names and column names from the same string
and this device's whole design rests on that never drifting.

- **Value**: blank, or a whole number 0-100. Blank means "no RM opinion
  entered" — same absent-means-absent convention `COL_LABEL` already uses.
  Not auto-computed from `SRC_MILESTONES`; a Research Manager enters it,
  optionally reading `MilestoneEvidenceReport`'s grouped tracker evidence
  first as reference material — same relationship that report already has to
  the `DONE` flags (advisory, never authoritative, never auto-written).
- **Where it lives**: **appended after the existing register columns, not
  inserted into the `L..AF` MS block.** `ExcelOutput.ReadSheetForDeckPeriod`
  hands `MilestoneDevice` a Dictionary keyed by field name — confirmed
  directly from `MilestoneDevice.bas`'s own header ("`rowValues` is the
  slide's row as ExcelOutput hands it over — field name to value") and from
  `ValueOr`'s lookup (`rowValues.Exists(key)`) — so nothing reads these
  columns by letter. Appending avoids the exact failure this project has
  already paid for once: reordering six columns broke four sentences and one
  wrong-column instruction elsewhere in hours. **Still confirm before
  building**: grep the whole repo for any hardcoded `MS` column letter
  (`L`, `O`, `R`...) the way `check_docs.py` already does for other roles —
  the header-dictionary read is what `MilestoneDevice` uses, but a stray
  direct-letter reference elsewhere would break silently if one exists.
- **Not a Field Spec entry.** Like `MS1_LABEL`/`_DATE`/`_DONE`, this column is
  addressed by shape name inside the tagged device group, never by role tag —
  confirmed by grep: zero `MS*` references anywhere in `FieldSpec.bas`. It
  needs no History treatment (`CARRY`/`FRESH`/`PART-FROZEN`/`DIFF`); it's a
  plain register value the RM overwrites each quarter, same as `DONE`.
- **No Field Spec means no drafting-sheet workflow either** — it's entered
  directly in the register, the same way `MS1_DONE` already is.

## Two ways to render it — pick one

### Option A (recommended): fold the percentage into the existing `_LABEL` text

No new shape. No template change. No 43-slide retrofit. `DrawFromRow` already
builds each slot's label text from `MS<n>_LABEL` before handing it to
`WriteText` — extend that one line:

```
labelText = labels(i)
If Trim(pcts(i)) <> "" Then labelText = labelText & " (" & pcts(i) & "%)"
```

Renders as e.g. *"Fieldwork complete (75%)"*. Ships the moment `DrawMilestones`
reads a fourth array; every real slide already has a working `_LABEL` shape,
so there is nothing to retrofit and nothing that can be missing on an
individual slide the way a new shape could be.

**Cost: near zero.** Trade-off: the percentage is inside a sentence, not
scannable at a glance across all 7 circles.

### Option B: a new `_PCT` shape per slot, same pattern as `_LABEL`/`_DATE`

Adds `PART_PCT = "_PCT"`, extends `IsColumnForThisDevice` to recognise it as a
fourth suffix, adds a `WriteText`/`SetVisible` pair in `DrawMilestones`
exactly mirroring how `_DATE` is handled today — **optional**, like `_OFF`:
absent on a template, it's simply not shown, reported once, never faked.

**Geometry, read from the real deck's P/K/S exemplar slides (44/46/47 —
identical across all three, confirming they share one master):**

| shape | size | position (slot 1 example) |
|---|---|---|
| `MS1_ON`/`_OFF` | 0.354in × 0.354in | x≈10.96in, y≈3.11in |
| `MS1_NOW` | 0.433in × 0.433in | x≈10.92in, y≈3.08in |
| `MS1_LABEL` | 1.575in × 0.318in | x≈11.60in, y≈3.13in |

The circle's right edge sits at ≈11.31-11.35in; the label's left edge starts
at 11.60in — **roughly 0.25-0.29in of genuinely empty horizontal space
between circle and label**, on every slot, on every template. That gap is
where a `_PCT` caption would go if built.

**`_DATE` is not a calendar date — corrected 2026-08-22, Rohan.** Checked all
7 slots on the P exemplar: `MS1_DATE="▶"`, `MS2_DATE="6"`, `MS3_DATE="12"`,
`MS4_DATE="24"`, `MS5_DATE="36"`, `MS6_DATE="48"`, `MS7_DATE="★"`. It holds
the milestone's month-offset number (the same offset `SRC_MILESTONES` groups
against — matches `MilestoneEvidenceReport`'s own live output this session,
"MS1: no numeric date to group tracker items against"), with the first and
last slots replaced by a start/end glyph instead of a number. Not cruft — a
real, already-shipped design.

**This changes Option B's risk, not its geometry.** The circle-adjacent
footprint (~0.35in) already carries a small numeric badge doing real work.
A `_PCT` badge placed there would be a *third* element competing for the
same tight space (circle + date badge already occupy it), not empty ground —
worth weighing against the 0.25-0.29in circle-to-label gap as the more likely
home if B is built. Still confirm exact `_DATE` placement live via COM before finalizing either —
a raw-XML regex read is workable for a quick check but doesn't resolve group
nesting the way COM does, so treat these numbers as good enough to plan
against, not good enough to build the final geometry from.

**Cost: the big one.** New shape needs adding to 3 exemplar templates *and*
all 43 real slides (same retrofit class as this session's earlier
milestone-device work), plus the code changes above, plus tests. Sized
comparable to the biggest single piece of work done this session.

## Recommendation

Build **Option A** first. It delivers the actual thing Rohan asked for — a
Research Manager's percentage opinion visible next to the milestone circle —
with a code change measured in lines, zero template risk, and zero retrofit.
If it turns out the label-text version isn't scannable enough in practice
(a real slide has to be looked at to know), Option B is a well-scoped
follow-up with its shape geometry already gathered above, not a redesign.

## Open questions for Rohan

1. **Format**: `"(75%)"` appended after the label, or some other punctuation/
   position (before the label, on its own line via the existing `||`
   line-break convention)?
2. **Which slots get a number?** Every slot with a value, or specifically
   discouraged on `DONE=Y` slots (redundant "100%" next to an already-dark
   circle) — a workflow convention for the RM, not a code branch either way,
   since the code just renders whatever `MS<n>_PCT` holds.
3. Whole numbers only, or does a Research Manager ever want to write
   something like "~75%"? (Affects whether the column is validated as numeric
   or left as free text like `_LABEL` already is.)
