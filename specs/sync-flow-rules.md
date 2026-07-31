# Sync flow — the decision rules

**Settled 31 July 2026.** Governs what happens between "the workbook disagrees with the
deck" and "the slide has changed". Supersedes the flow described in any earlier round doc.

These are rules about **the change set**, not about buttons. Which surface a human sees is
derived from what is actually pending, never from which button was pressed — a button that
picks its own level of scrutiny is a button that can pick the wrong one.

---

## The shape

```
register + deck
      |
      v
  [ build ]  ------> change set (only genuine differences)
      |
      v
 [ classify ]  ----> each item is BATCHED or INDIVIDUAL
      |
      +---- batched ------> confirmation dialog ---+
      |                                            |
      +---- individual ---> review worksheet ------+
                                                   |
                                                   v
                                          [ one apply path ]
                                     backup -> revalidate -> write -> log
```

---

## F1 — Only genuine changes enter the flow

An item exists only where the slide's current value differs from what the register would
write. Unchanged rows appear on no surface and in no count.

*Why:* the first real run was 46 rows, of which 27 were already correct. Putting those 27
in front of a human is how a review becomes a formality. (R13.3)

## F2 — Every item is classified exactly once: content kind first, measurement second

An item is **batched** only if **both** hold:

1. its `FieldID`'s content kind is `Controlled`, **and**
2. two or more items share the same `FieldID` **and** the same `Current` **and** the same
   `Proposed` — all three exactly equal.

Everything else is **individual**.

*Why both halves:* measurement alone cannot see that prose which happens to be uniform is
still prose (eleven projects reading "No key events this quarter" is a perfect batch by
measurement and forbidden by R13.2). A declared kind alone cannot verify that the changes
really are one transformation. Kind decides eligibility; measurement decides membership.

**A group of one is never a batch.** That is an individual review wearing a different word.

## F3 — Two surfaces, chosen by the change set

| Items | Surface |
|---|---|
| batched | confirmation dialog: transformation, count, affected entities |
| individual | the `Sync Review` worksheet |

A single run may use both. The dialog is a legitimate review surface *only* for batched
items, because a uniform batch is one decision (R13.2) and there is nothing to come back
to. Prose is real reading and gets the worksheet, which is leavable and resumable (R13.4).

## F4 — The dialog is offered only while it is readable

At least one batch, and no more than `MAX_BATCHES_IN_MODAL` (12) distinct batches. Above
that, the batched items go to the worksheet too.

*Why:* a wall of transformations in a dialog is dismissed, not read — which would recreate
the rubber stamp R13 exists to remove, with more text on it. The number is a judgement
about what fits on a screen. Being wrong costs a trip to the worksheet, which is the safe
direction.

## F5 — Partial application is a normal, complete outcome

Applying the batched half and deferring the individual half is valid. A run is not
all-or-nothing.

**But the dialog must state what it is not covering, before it asks the question.** A
partial run that reads as a whole one leaves prose silently unapplied while the operator
believes they are finished.

*Origin:* the first version of these rules required the whole change set to be batchable,
so one prose row sent everything to the worksheet. That was extra strictness with no rule
behind it, and it made the large easy part wait behind the small hard one.

## F6 — Order within a run

```
build -> dialog (if batched work) -> apply approved batches
      -> REBUILD -> write worksheet for what remains -> open it
```

The rebuild is not optional. The worksheet must describe the deck as it is *after* the
batch writes, not as it was before them — otherwise a human reviews a before-and-after
whose "before" the same click just invalidated.

## F7 — Exactly one write path

Both surfaces converge on the same apply: backup, revalidate, write, log. What differs
between them is the *provenance of the approval*, never the write itself.

*Why:* two write paths become two sets of guards, and the second one always ends up with
fewer.

## F8 — Revalidate at the moment of writing, not at the moment of approving

Immediately before each write, recompute the change's identity from the **live slide** and
the **current register value**, and compare it to what was approved. Any mismatch drops the
row — it is never written, and it is reported.

This covers: a slide hand-edited after approval, a register value changed after approval,
and any edit made to the worksheet's own `Current`/`Proposed` cells (which are display, not
input — the place to change a value is the register).

## F9 — Approval never outlives what it was about

- A new build clears every tick.
- An apply consumes the sheet; the same approvals cannot be applied twice.
- A changed before-or-after voids that row (F8).

## F10 — Every unknown fails toward more review

| Unknown | Resolves to |
|---|---|
| content kind not recognised | `Prose` → individual review |
| approval cell not an affirmative `Y`/`Yes` | not approved |
| no review sheet for a type | "you never reviewed", **not** "you approved nothing" |
| deck path unreadable / cloud-hosted | write proceeds, **loudly** reported as unbacked |

*Why the last one differs:* refusing there would block cloud decks entirely while adding
nothing, because the real protection is that the original deck is never opened for write.
Every other unknown costs only extra reading, which is why it fails closed.

## F11 — A trailing paragraph mark is not a difference

When deciding whether a value needs writing, trailing `CR` / `LF` / `Chr(11)` are ignored
on **both** sides. What gets written is still exactly the value supplied — this rule
decides *whether* to write, never *what* to write.

**Measured, not assumed.** `ABOUT_BODY` against the real 46-slide deck, after the
register's encoding was repaired: 22 pending changes, of which **20 were this and nothing
else**. `TextRange.Text` returns a shape's trailing paragraph mark; the harvested register
value does not carry it. Not one of the 20 was a wording change. Applying them would have
rewritten 20 slides of real prose to delete a character nobody can see — and no human
reading that list would have approved it, which means the tool would have been generating
review work that existed only to be dismissed.

**Scope is deliberately narrow.** *Trailing* only. An internal paragraph break that the
register does not carry is still a real difference and still surfaces — that is the
remaining 2 rows (`2_P004`, `2_P004-2`), and they are a separate question about how `||`
maps to paragraph breaks, not this one.

**The write path is untouched on purpose.** Normalising what gets written, rather than what
gets compared, would strip real formatting out of a deck one sync at a time — invisibly,
and with every run making the next run's baseline slightly more wrong.

The post-write verification uses the identical rule. If the two disagreed, a value whose
only difference was a trailing break would be judged "needs writing" by one check and
"write did not take" by the other — a permanent, self-inflicted failure on a field that is
in fact correct.

## F12 — Seeding is not approving

`Status` carries **one** meaning: *a human read this text and means it*. It is the gate
between synthesis and population, and it only works if it says exactly one thing.

| Status | Meaning | Writable |
|---|---|---|
| `Approved` | A person read this and meant it | **yes** |
| `Seed` | Copied off the slide to establish a baseline or exemplar | never |
| `Draft` | Synthesised, not yet agreed | never |
| *anything else* | A typo in the column that decides whether words reach a slide | never, **and reported loudly** |

**Why this needed its own status rather than leaving seeds as `Approved`.** All 46
`ABOUT_BODY` rows in the real register read `Approved`. Not one had been read and agreed by
anybody — they were copied off the slides so the column would not be empty, exactly as the
field package specifies for seeding. Copying something is not approving it.

That was harmless only while seed and slide were identical, which is precisely why it was
easy to miss. The moment drafted text arrives, `Approved` becomes the thing deciding
whether words reach a slide, and a row copied off that same slide is **indistinguishable**
from one a human genuinely read. There is no way to recover the distinction after the
fact — so it has to exist before the first drafting round, not after it.

**An unrecognised status is reported even on an otherwise healthy read.** Every other
rejection is the system working as intended: `Seed` is meant to be refused, `Draft` is
meant to be refused, a different period is meant to be refused. A misspelt status is none
of those — it silently removes a row from the sync, and it looks identical to a deliberate
hold. Returning early once some rows were accepted would hide it behind exactly the runs
that appear fine.

**The counters discriminate, deliberately.** "3 rows not approved" is true of a seed, a
draft and a typo alike, and useless about all three. `RejectedSeed` / `RejectedDraft` /
`RejectedUnknownStatus` are reported separately because they are three different situations
wearing one number — the same failure as F1's ambiguous zero, in the status column.

