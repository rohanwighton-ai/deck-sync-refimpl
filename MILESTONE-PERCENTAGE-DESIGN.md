# Milestone percentage display — design

> **THE SHAPE WAS REMOVED, LATER THE SAME EVENING — 2026-08-22.** Option A
> shipped, was superseded by Option B (a real `MS<n>_PCT` shape) after
> Rohan looked at it on screen, was retrofitted across all 46 slides and
> tested end-to-end live — then Rohan's verdict on seeing it rendered:
> "remove the % tags we added please they didn't really work." All 322
> shapes were removed. **The contiguous-colour and position-gating logic
> below survive and are still live** — the objection was to the visual
> shape specifically, not that underlying design. `COL_PCT`/`PART_PCT` and
> the register columns are left in the code/register, inert without a
> shape to write to. Full history: `FIX-LIST.md` items CQ (Option A) and
> CT (Option B, retrofit, and the removal). Kept below as a design record
> of what was tried and why it didn't survive contact with the real deck —
> if percentage display is revisited, don't re-derive this from scratch.

## The problem

The milestone timeline (`MilestoneDevice.bas`) used to show exactly one state
per slot: achieved-and-current (big circle), achieved-earlier (small circle,
one colour), not-achieved (small circle, another colour). A project that's
finished MS2 and MS3, has done nothing visible on MS4/MS5, and finished MS6
rendered as "done, done, gap, gap, done" — accurate, but reads as messy or
wrong at a glance (`2_P009`, this session).

## What actually shipped — three decisions, in the order they landed

**1. Colour is positional (contiguous), not per-flag.** Originally this doc
said `lastAchieved`'s definition of "current" was the only thing that
mattered and colour stayed per-flag. Rohan, looking at the real timeline:
*"we spoke about making the gaps the same colour as done if they are before
where the big circle currently is... colour are always contiguous."* Any
slot at or before the current slot now renders as achieved (`_ON`/`_NOW`)
regardless of its own `DONE` flag; anything after current still renders
`_OFF` — which it always would anyway, since current is defined as the
*highest* done-flagged slot, so nothing after it could ever have been "done".
The flag still decides WHERE current sits; it stopped deciding each earlier
slot's own colour. `MilestoneDevice.DrawMilestones`, one line:
`isDone = (i <= lastAchieved)`, replacing `IsDoneWord(done(i))`.

**2. The percentage is its own shape, not folded into the label (Option B,
superseding Option A).** First built as Option A — a percentage appended to
`MS<n>_LABEL`'s own text, e.g. *"Fieldwork complete (75%)"*, zero new shapes.
Rohan, after seeing it live: *"isn't that separator needed?"* (a real,
separate question) led to *"as a smaller font separate label under the date
label."* `PART_PCT = "_PCT"` is now a real, optional shape (same pattern as
`_OFF`: absent on a template, it's simply not shown, reported once, never
faked) — `MS<n>_PCT`, positioned directly below that slot's own circle,
small bold centred text, hidden until a value exists.

**3. The percentage is ALSO gated by position, same rule as the colour.**
Rohan, looking at the first real one on screen: *"not worth having on future
ones, should just turn on when big lead circle reaches that point."* A
percentage sitting in the register for a slot after current is suppressed
even if a shape and a value both exist — `Trim(pctValue) <> "" And i <=
lastAchieved`. Only once the current marker reaches or passes a slot does
its percentage become eligible to show.

**What never changed**: the module's hard rule — never creates, moves,
resizes, or reorders a shape, except for the one deliberate exception this
feature *is*: the shapes were created once, by a dedicated retrofit script,
not by the runtime `DrawMilestones` code, which only ever toggles visibility
and writes text, exactly as before.

## Data model

**`MS<n>_PCT`** (`MS1_PCT` .. `MS7_PCT`), a register column, same naming
convention as `_LABEL`/`_DATE`/`_DONE`/`_OFF` — `ColumnFor(i, part)` builds
shape names and column names from the same string, deliberately, so they
cannot drift apart.

- **Value**: blank, or a number (a trailing `%` is stripped and re-added on
  write, so `"75"` and `"75%"` render identically). Blank means "no RM
  opinion entered." Not auto-computed from `SRC_MILESTONES`; a Research
  Manager enters it, optionally reading `MilestoneEvidenceReport`'s grouped
  tracker evidence first as reference material — advisory, never
  authoritative, never auto-written, same relationship that report already
  has to the `DONE` flags.
- **Where it lives**: appended after the existing register columns, not
  inserted into the `L..AF` MS block — confirmed safe because
  `ExcelOutput.ReadSheetForDeckPeriod` hands `MilestoneDevice` a Dictionary
  keyed by field name, never by column letter.
- **Not a Field Spec entry.** Addressed by shape name inside the tagged
  device group, never by role tag — needs no History treatment; a plain
  register value the RM overwrites each quarter, same as `DONE`.

## The shape

Real, optional, created once by `vba/tools/add_pct_shapes.vbs` (kept in the
repo for the next slide that needs it — see its own header). Geometry is
**measured per-slide from that slide's own `MS<n>_ON` circle**, never a
hardcoded absolute position, matching the "measured, not computed"
convention `DrawMilestones` already uses for the bar/track:

- Width/left match the circle's own width/left.
- Top = circle bottom + 0.015in gap.
- Height 0.10in, font 5.5pt bold, centred, `AutoSize` explicitly disabled.

**One real bug found and fixed during the build**: an empty textbox with
`WordWrap=False` and no `AutoSize` override auto-fits to ~0 width the moment
it's created — every shape shipped at `cx=65` EMU (invisible) on the first
attempt, caught only by reading the saved file's own XML rather than
trusting the "7 shapes added" success message. Fixed by disabling `AutoSize`
before setting any text/font property, then re-asserting geometry
explicitly afterward as a safety net.

**Regrouping to add the shape destroys the device group's own name and role
tag** (confirmed empirically against a scratch copy before touching the real
file) — both are captured before every regroup and explicitly restored
after.

**Chaining more than one slide's regroup work in a single PowerPoint session
produced a reproducible "Type mismatch"** on the second slide's device-group
lookup, even though the identical logic against that slide in total
isolation worked cleanly. Never root-caused further — the retrofit runs one
slide per PowerPoint launch instead (driven by a bash loop), which sidesteps
it entirely and matches the fresh-instance pattern already used elsewhere in
this project.

## Verification

- **Code**: fail-first proven twice — the contiguous-colour condition and
  the position-gate condition were each deliberately reverted, confirmed to
  fail the right assertions for the right reason, then restored. Full
  automated suite (`run_vba_tests.ps1`) green before the retrofit began.
- **Retrofit**: dry-run against a scratch copy first; real deck backed up
  before every write; every one of the 46 target slides re-verified from
  the *saved file's own bytes* afterward — 7 PCT shapes each, correct
  non-zero width, `MILESTONE_TIMELINE` group name and role tag intact,
  `MS_BAR`/`MS_TRACK` counts unchanged. `VerifyRealDeck` re-run clean
  afterward: 0 mismatches, 0 unwired fields, 43/43 real slides fully OK.
- **Not yet verified**: nobody has entered a real `MS<n>_PCT` value in the
  live register and run an actual sync — the mechanism is proven, the real
  workflow end-to-end (draft → register → sync → slide) hasn't been
  exercised with real data yet.
