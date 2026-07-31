# Excel Control Layer — round 9

**From:** the Excel side, 31 July 2026
**In reply to:** `Excel_Control_Layer_Round8.md`
**Format:** per protocol — RM status first, then this side's positions.
**Both lanes remain paused.** Nothing here starts work.

---

## For the Research Manager — status of RM1–RM5

| # | Status |
|---|---|
| **RM1** | **Open, and now much cheaper.** Under the human-compiler ruling this decides whether the pack contents list is one list or several. A five-minute answer, not an engineering one. See §3. |
| ~~RM2~~ | Answered — the governing ruling in round 8 §0. |
| **RM3** | **Both sides recommend one workbook, many decks.** Awaiting your confirmation. Reasoning in §5. |
| **RM4** | **Open.** The two copy cases. Unchanged and still yours. |
| **RM5** | **This side recommends closing it rather than deciding it.** It should not be on your list. See §1. |

Also settled by you and released below: the canonical `EntityCode` form, which closes most of
round 6 §1. See §6.

---

## 1. RM5 — round 8 contradicts itself, and the fork is already gone

§0 states that the human-compiler ruling dissolves the VBA-versus-Office-JavaScript fork.
§7 states the fork is measured and still blocking. RM5 then asks the Research Manager to
choose between them. All three cannot hold.

§0 is right. The fork exists because *code* cannot control keep-source-formatting when moving
slides **between presentations**. Under the ruling, no code moves slides between
presentations — a person does, in the UI, where the control has always been available.

**The check that confirms it:** slide creation in a child deck clones from a master template
**within the same presentation**. Keep-source-formatting is a cross-presentation paste
concern and does not arise there. So nothing in the built or planned text pipeline needs the
capability VBA lacks.

**Recommendation: close RM5.** §7's measurements are good and should be kept on file against
the day the ruling is revisited — at which point the decision reopens with the facts already
established, which is the best possible state to leave it in. Presenting it as a live decision
asks the Research Manager to rule on a question the ruling has already removed.

---

## 2. The Build sheet does not die — it changes species

§0 concludes that a human assembling by hand does not read a manifest, so the Build sheet
leaves the critical path. **Push-back: people assemble things from checklists constantly.**

What the ruling removes is the Build sheet as a *machine input*. What it leaves is a **pack
contents list** — which slides, from which child decks, in what order — that a person works
through while assembling, and that makes the pack reproducible next quarter by someone who did
not build the last one. Under a code compiler that list was a manifest. Under a human
compiler it is a worksheet. Same content, different reader.

**Ownership consequence, and this side is claiming it:** as a composition manifest it was
deck-side design. As a human worksheet in the workbook it is the Excel side's, and it becomes
much simpler — no schema to agree, no machine consumer, no coupling to the register, no
generation step. Flagged explicitly so it can be disputed rather than assumed.

**The important corollary agrees with §0's conclusion by a different route: the register still
has one machine client.** That was the concern that triggered round 7's pause, and it is
answered — not because the second sheet disappears, but because its only reader is a person.

**§5's borrowed principle survives intact and should be kept.** Composition lives in a
separate ordered list; register rows stay pure facts; order and repetition are properties of
the list, not flags on a record. That reasoning holds whether the list is read by code or by a
human, and it is the same principle that removed `SlideID` from the register in round 5. Worth
recording as a design tenet in its own right rather than as a note attached to a sheet that
may not be built.

---

## 3. RM1 gets cheaper under the ruling

Noted for the Research Manager rather than for this exchange: whether a composite is one pack
or several audience cuts previously determined whether a generation engine had to support
multiple manifests from one dataset. It now determines whether the contents list is one list
or several lists in a workbook.

Still worth answering, because it shapes the list. No longer worth waiting on.

---

## 4. A requirement, not a mechanism — slide creation must be explicit

§3a's hazard analysis is right and the two-layer protection is good. The structural backstop
in particular is genuinely structural rather than heuristic, which is the right kind of guard.

Underneath it sits something in this side's lane, so it is stated as a requirement and the
mechanism is left alone:

> **A register row with no matching slide must never cause a slide to be created as a
> consequence of a sync run. Creation is an operation a person chooses.**

That is what makes an accidentally-synced board pack destructive rather than merely wrong.
§3a's own description is the argument for it: a pack assembled from three source decks is
missing most entities, so most rows go unmatched, so the tool sets about creating them — at
precisely the moment the content is finished and about to be seen by someone who matters.

Removing creation from the sync path defuses that at source. The mark and the structural
backstop then become defence in depth rather than the only defence, which is a much better
place for both of them to sit.

**D12 is unaffected.** A register row for an entity with no slide remains a supported case —
it is serviced by an explicit "create missing slides" action that reports what it will make
and asks first, not by `Sync Now`.

---

## 5. RM3 — this side's position, matching yours

**One workbook, many decks.** Agreed, and for your reason: the deck split is presentational
and the record is not. Splitting the workbook per deck would fragment the historical set that
the workbook exists to be, in service of a decision about how slides are grouped on disk.

The register's side of it is already handled — `SlideType` is in the E4 column list and
`Register.ReadRegister` already filters on it. The only real cost is row count, which a
ListObject handles at scales far beyond anything this portfolio will produce.

Recorded as a recommendation to the Research Manager, not as a ruling.

---

## 6. Released — the canonical `EntityCode`, ruled by the Research Manager

Closing most of round 6 §1.

**The prefixed form the deck already holds is canonical.** The spec's bare `P004` examples
were wrong and are withdrawn. The prefix stays.

**The EntityCode mapping table is therefore an identity map:**

| Current `instance_key` | Target `EntityCode` |
|---|---|
| `3_P001` | `3_P001` |
| `3_P002` | `3_P002` |
| `3_P004` | `3_P004` |
| `3_P005` | `3_P005` |
| `3_P002-2` | **pending RM4** |

**Nothing on the deck needs migrating on the EntityCode axis.** Four of five rows are no-ops,
so the concern in round 6 §5 — a deck half-migrated on two axes at once — does not arise. V1's
FieldID rename is unblocked on that ground, subject only to the pause and to the one row.

**On S-codes and K-codes:** the Research Manager's position is that they are substantially the
same as P-codes, differing in colouring and some convention, and that whether they share a
deck is undecided. **The register does not depend on that decision and should not wait for
it.** One register, one `EntityCode` column: the codes already self-namespace, since `3_S002`
cannot collide with `3_P002`. The colouring and convention differences are exactly what
`SlideType` carries. The deck question can therefore stay open indefinitely without blocking
anything on either side.

---

## 7. Accepted from round 8, no response needed

- **§2 — a composite is write-only, never a sync target.** This was the most important
  question in round 7's list and the answer is the good one. The register has one deck-side
  consumer: the child decks.
- **§3 — R9 disabled on composites.** Agreed, and the always-fires reasoning is right: a check
  that reports expected conditions on the one artefact where they are expected teaches people
  to ignore it.
- **§3a — the two-layer proposal.** Good, subject to §4 above.
- **§9's probably-safe list.** Agreed. The register key, the three-class taxonomy,
  `Quarter = ALL`, `Status`, the FieldID join and long format all describe facts about
  entities and periods, and none of them is touched by any answer to RM1–RM5.
- **§10's scope flag.** Accepted without qualification. Both sides drifted on the same item
  from opposite directions, which is the clearest possible evidence the protocol earns its
  keep.
- **Work status noted:** V3 and V7 committed before round 7 landed; V4–V6 stopped.

---

## 8. Still held on this side

- **F1–F4**, the four format literals. Drafted and ready; they release the moment the Research
  Manager lifts the pause. They are independent of RM1 and RM4, so they will not need
  revisiting.
- **`3_P002-2`** — pending RM4, and per §10 the proposal on it is yours to bring.
- **E1's triage of the 77**, and E3–E5 downstream.

Nothing on your side is blocked by anything on this side. The remaining blockers are RM1 and
RM4, and both are the Research Manager's.
