# Excel Control Layer — Round 13

**From:** the VBA side, 31 July 2026
**Replying to:** `ABOUT_BODY_Field_Package.md` (§1 R13, §2–§4 the field package)
**Contains:** R13 accepted and specified as a mechanism (§1–§3), one column we need from
you and cannot derive (§4), and the two dependencies that gate seeding `ABOUT_BODY` (§5).

---

## 0. Position

R13 is accepted without argument. Your reading of the `PROJECT_STATUS` run is correct:
it passed `Status = Approved` and skipped the question of whether *this change to this
slide* was right, and 19 slides moved with nobody seeing a before-and-after. That is a
gate that was never built, not a gate that failed.

R13 is being built **before** `ABOUT_BODY` is wired in, per your §5.2. Nothing from §2–§4
is being started until it lands.

### 0.1 One setting the RM has loosened, deliberately and temporarily

**Ruled 31 July 2026: while the work runs on a carved COPY of the deck, the review grid is
built and read in full, but may be approved wholesale via a named `Approve All` action.**

The reasoning: R13 exists to protect the real deck, and a throwaway copy has nothing to
protect. Building the full mechanism now and running it loose keeps velocity without
leaving the gate unbuilt — and tightening then costs the removal of one button rather than
the construction of a mechanism.

Two things make this safe rather than a quiet weakening:

- The real deck (`test1.pptx`) is never opened for write. That rule is not relaxed, and it
  is the entire basis for relaxing this one. If it ever is relaxed, this reverts first.
- `Approve All` is a separately named action someone presses, not a default and not an
  absent gate. Your own §13.2 warns that bulk approval "teaches the operator to click
  through" — that risk is real, so the loose mode is a decision taken each time and
  recorded in the run report, rather than a default that outlives the copy it was
  justified by.

**Trigger to revert to full R13:** the first run that points at anything other than a
scratch copy.

---

## 1. The mechanism

**A review grid written into the paired workbook, not a dialogue box.**

We already run this pattern for batch onboarding (`WriteInstanceKeyGrid` /
`WriteReviewGrid` / `ReadReviewGrid`): the tool writes a worksheet, the human edits it at
their own pace, the tool reads it back and acts only on what is there. It fits R13.4
exactly, because a worksheet is inherently leavable and returnable — resumability is not a
feature we have to add to it, it is what a worksheet already is.

The sync path gains a review surface, and **which surface you get is decided by the change
set, not by which button you press**:

| Change set | Surface | Why |
|---|---|---|
| Every change belongs to a uniform batch, ≤12 batches | One dialog listing every transformation, count and affected entities | R13.2: a verified uniform batch **is** one decision |
| Anything needs reading individually — prose, one-offs | The `Sync Review` worksheet | R13.4: prose review is real reading, and must be resumable |

`Sync Now` survives, and this is a correction to our own first answer. We initially removed
it outright, reasoning that its confirmation showed a *count* and R13 demands a
before-and-after. That was right about the count and wrong as a conclusion: it ignored your
own R13.2, which says approving the members of a uniform batch individually "teaches the
operator to click through, which is worse than not asking." Nineteen identical
`In progress` → `In Progress` corrections are **one** transformation. A dialog showing that
transformation in full is not a count-based confirmation — it is a complete before-and-after
that happens to fit on a screen.

So `Sync Now` now builds the same queue as everything else, and **refuses to the worksheet
the moment one individual decision exists**. One prose row disqualifies the fast path for
the whole run; there is no partial modal.

What has genuinely gone is the *unreviewed* write. There is exactly one place in the add-in
that writes a field to a slide, everything reaching it was approved against a visible
before-and-after, and every path revalidates against the live slide immediately before
writing. A policy that says "don't press that one" is not a control — so there is no button
left that could be pressed.

**Grid shape**, one row per changed field per entity:

| A | B | C | D | E | F |
|---|---|---|---|---|---|
| `EntityCode` | `FieldID` | `Current` | `Proposed` | `Batch` | `Approve (Y/N)` |

Column E is blank for individually-reviewed rows, or a batch label shared by every row in
a uniform batch. Approving any one row of a batch approves the batch.

---

## 2. How each of R13.1–R13.5 is actually held

| Rule | Held by |
|---|---|
| **R13.1** no replacement without a before-and-after | Columns C and D are the row. A row with no C→D difference is not written at all (see R13.3), so "shown" and "exists" are the same condition. |
| **R13.2** only uniform batches may be batched | §3 below — and it needs one thing from you. |
| **R13.3** unchanged rows never enter the queue | The queue is built from `in_place_correction` actions only. `no_change` never reaches the grid. The 46-row run would have produced a 19-row grid. |
| **R13.4** readable, resumable, not a modal | It is a worksheet. Close it, reopen it, finish it Thursday. |
| **R13.5** approval does not carry across runs | Two independent mechanisms, below. |

**On R13.5, deliberately belt and braces**, because this is the rule most likely to be
quietly violated by a convenience later:

1. **Run stamp.** The grid carries the run identity that built it. `Apply Approved`
   refuses a grid stamped by any other run, and marks it consumed once applied. A second
   apply against the same grid does nothing.
2. **Per-row content hash.** Each row carries a hash of
   `EntityCode | FieldID | Current | Proposed`. At apply time the plan is recomputed and
   each approval is matched by hash. A row whose hash no longer matches a live pending
   change is **dropped, not applied**.

The stamp is what enforces your rule as written. The hash covers the case your rule
implies but does not name: someone approves a row on Tuesday, edits that slide by hand on
Wednesday, and applies on Thursday. The approval was for a before-and-after that no
longer exists, and it dies rather than overwriting the hand edit.

---

## 3. Batching, and why it is verified rather than declared

Your R13.2 wording is the load-bearing part: *"a **verified** batch of uniform changes"*,
and *"every instance shares one FieldID **and** one transformation."*

We are implementing uniformity as a **measured property of the actual change set**, not
as a claim made in a spec:

> A batch exists where two or more rows share the same `FieldID`, the same `Current`
> value, and the same `Proposed` value.

All three, exactly equal. The 19 `In progress` → `In Progress` corrections satisfy it and
present as one decision: *"PROJECT_STATUS: 'In progress' → 'In Progress', 19 entities:
3_P001, 3_P002, …"*. Nineteen slides with nineteen different current values could never
group, whatever any column claimed about them.

This matters because the alternative — trusting a declared class — means a mislabelled
field silently converts N unreviewed writes into one click. Deriving uniformity from the
data means a wrong label can only ever cost extra reviewing.

---

## 4. Content kind — noted, deferred, **no action wanted**

**Uniformity is necessary but not sufficient, and the missing half is not derivable from
the register.** Recording it here so it is on the record, not to open a round.

Your class table has four rows: Controlled, Quarterly (prose), Entity-static, Chrome. The
register can only distinguish **cadence** — `Quarter = ALL` versus a period literal. It
carries nothing that separates a *controlled vocabulary* field from a *prose* field, and
those are the two you treat most differently:

- `PROJECT_STATUS` — quarterly, controlled, batchable
- `KEY_EVENTS_BODY` — quarterly, prose, **never** batchable

Both are `Quarter = <period>`. In the register as specified they are indistinguishable.
Worse, prose can be *accidentally* uniform — eleven projects all reading "No key events
this quarter" and all getting the same rewrite measure as a uniform batch of 11, while
R13.2 says prose may never be batched.

**What we have done instead of asking you for a column:** hardcoded the lookup in
`ReviewQueue.ContentKindOf` — `PROJECT_STATUS` and the two identifier fields named, and
**everything else defaults to Prose**. With two fields in flight a lookup table beats a
schema change, and a schema change costs an exchange round we would rather spend on
`ABOUT_BODY`.

**The default is the load-bearing part:**

> Absent or unrecognised content-kind ⇒ **not batchable** ⇒ every row reviewed
> individually.

Failing that direction costs a longer review. Failing the other converts N unreviewed
writes into one click, which is the thing R13 exists to prevent. The absence of a label is
never read as permission.

**Revisit at the third field**, not before. If the register grows a `ContentKind` column
(`Controlled` / `Prose` / `Static`) at some later point, we will read it and delete the
hardcoded table — it is deliberately one `Select Case` so that is a small change.

---

## 5. Seeding `ABOUT_BODY` — two dependencies, one of them ours

Your §5.1 asks us to seed column C and the register `Value` from the current slides, 43
entities. Both blockers are on our side and are named here so the sequencing is visible:

1. **`DumpFieldValues` is written but not yet callable.** In a freshly imported VBA
   project a function only resolves through `Application.Run` if the cross-module Public
   UDTs it declares were touched by an earlier `Application.Run` in the same session — it
   fails as "Sub or function not defined" while a manual compile passes cleanly. The
   proper fix is forcing a compile after import; warm-up probes are the current
   workaround. Until this is done we cannot dump 43 slides' current values reliably.
2. **Not by parsing the .pptx XML.** We tried. The XML disagrees with the object model on
   nearly every field, so a register seeded that way would be wrong on day one and the
   exemplar column in your §4 — the whole point of that layout — would be anchored to
   fiction. The seed comes from the object model or it does not come.

**And the ordering point:** the three `-2` slides (`3_P002-2`, `2_P004-2`, `1_P006-2`) must
be deleted **before** the seed runs, or the dump produces 46 rows against your 43 and the
extras are duplicates of exactly the entities most likely to be edited.

---

## 6. Summary

| | |
|---|---|
| R13 | Accepted, being built now, before `ABOUT_BODY` |
| Loosened, temporarily | `Approve All` permitted **while on a scratch copy only** (§0.1) |
| `Sync Now` | Kept, batch-aware — refuses to the worksheet on any individual change |
| Batching | Measured from the change set, gated by a hardcoded kind table |
| Content-kind column | **Deferred — do not build.** Revisit at the third field (§4) |
| Blocking us | `DumpFieldValues` callability; the three `-2` deletions |
| **Needed from you** | **Nothing.** No decisions requested, no columns wanted |

### One change to your runbook

`Sync Now` still exists and still does the obvious thing for a routine controlled-field
run — it just shows you the transformations first. For anything with prose in it (so, every
`ABOUT_BODY` run), the flow is **`Review Changes` → tick the `Sync Review` sheet →
`Apply Approved`**, and `Sync Now` will send you there itself rather than proceeding. The
`TPL_ABOUT_BODY` worksheet and its Copilot prompt block in your §4 are unaffected — that is
a drafting surface and sits upstream of all of this.

---

## 7. `Status` now has a vocabulary, and it is not optional

**Added 31 July 2026, after measuring your register.** All 46 `ABOUT_BODY` rows read
`Approved`. None had been read and agreed by anyone — they were copied off the slides per
your §3 seeding step. Copying is not approving, and the column could not tell the
difference.

| Status | Meaning | Reaches a slide |
|---|---|---|
| `Approved` | A person read this and meant it | **yes** |
| `Seed` | Copied off the slide as a baseline or exemplar | never |
| `Draft` | Synthesised, not yet agreed | never |
| *anything else* | A typo in the column that decides whether words reach a slide | never, and it warns |

**Why it could not wait.** It is harmless today only because seed and slide are identical —
which is exactly what makes it easy to leave. The moment Copilot drafts replacement text,
`Approved` becomes the thing deciding whether words land on a slide, and a row copied off
that same slide is indistinguishable from one a human genuinely read. The distinction
cannot be recovered afterwards.

**What this changes for you:** `ABOUT_BODY`'s 46 rows are being relabelled `Seed`, so the
field is inert until someone approves text. That is the correct state — the drafting round
in your §4 is what produces `Approved` values, and the exemplar in column C is exactly what
`Seed` is for. `PROJECT_STATUS` is untouched: those are controlled-vocabulary values applied
and verified against a prediction.

**One thing to confirm:** the publish step in your §4 says approved column-F values are
pasted into the register. Those rows should be written `Approved`; everything the seeding
step produces should be written `Seed`. If your publish macro hardcodes `Approved`, that is
the one line to change.

