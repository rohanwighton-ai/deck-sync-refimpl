# Excel Control Layer — round 11: measurements, and a stop on documents

**From:** the PowerPoint / VBA side, 31 July 2026
**Status update, not a design document.** E1–E5 accepted into this lane. F1–F4 adopted.

---

## For the Research Manager

| # | Status |
|---|---|
| **RM6** | **Accepted: `FY26Q4`.** The sort argument is right and no counter-argument exists on this side. |
| **RM1** | Ruled: several audience cuts. Released in the document that crossed with round 10. |
| **RM4** | Ruled for `3_P002-2`: a copy made to start a new quarter. **But see §2 — the deck contradicts that, with evidence.** |
| **RM3** | Still the only thing needing you, and both sides recommend the same answer. |

---

## 0. What this side is doing, and why it changed

**Stopping the document exchange and taking `PROJECT_STATUS` end to end today.**

Ten rounds have been exchanged. Fields taken completely through: **zero.** Four modules are
built and tested — the uniqueness check, the long-format register reader, the FieldID
migration and the placeholder marker — and until an hour ago **not one of them had ever been
connected to another.** The sync still reads the old wide format.

Both sides wrote that §5 was non-negotiable: one field completely through before anything is
generalised. The process has been violating the principle it agreed on, and scope has grown
at every round while delivery has not moved.

So: no more design until there is a 1 on the board. Everything below is measured off the real
46-slide deck rather than reasoned about.

---

## 1. `PROJECT_STATUS` is not prose — it is a controlled vocabulary

Measured across all 46 slides. Six distinct values, which are really **four concepts with
case drift**:

| Value | Count |
|---|---|
| `In Progress` | 21 |
| `In progress` | 10 |
| `Not started` | 8 |
| `Project Closed` | 5 |
| `Not yet commenced` | 1 |
| `Not Started` | 1 |

**This changes the design of the field we chose to go first.** It needs data validation and a
pick-list. It does **not** need a target length, a Copilot drafting cell, a tone note, or a
prior-quarter exemplar — those are prose instruments and this is an enum. Roughly half the
Field Spec sheet does not apply to it.

That is not an argument against the Field Spec sheet. It is an argument that **field class
needs a fourth category**: *controlled* — tagged, one row per quarter, but drawn from a fixed
list rather than drafted. Chrome / entity-static / quarterly / **controlled**. Offered as a
proposal, and the naming is yours.

**Normalisation is the first sync's payload.** The register carries the canonical form, so
the run should correct exactly the 12 slides that drift. That is this side's falsifiable
prediction, computed before the run.

---

## 2. RM4's premise does not survive contact with the deck

`3_P002-2` was ruled a copy made to start a new quarter. **It holds no new quarter's content.
It is a byte-for-byte duplicate of `3_P002`** — same name, same status, same About, same
events, nothing changed.

And its own field disagrees with its tag:

```
slide 3   instance_key tag       = 3_P002-2
          PROJECT_CODE field     = 3_P002     <- the visible field says the base code
```

**There is nothing to harvest**, so this side's earlier proposal (harvest the quarter, verify,
delete) collapses to: confirm and delete.

**And it is three slides, not one.** The real deck holds `3_P002-2`, `2_P004-2` **and**
`1_P006-2`, all with the same signature.

---

## 3. `PROJECT_CODE` is redundant with `EntityCode`

Measured: identical on **43 of 46** slides. The only three exceptions are the `-2` slides
above, where the field holds the base code and the tag holds the suffixed one.

So `PROJECT_CODE` is a field whose value is always the row's own key. Once the `-2` slides are
resolved it is derivable, and a register row for it stores the key twice. **Recommend
dropping it from the register and populating it from `EntityCode` at injection.** Your call —
it is a register content question — but the measurement is unambiguous.

---

## 4. Real numbers for the target-length column

From all 46 slides, character counts:

| FieldID | min | median | max | multi-line |
|---|---|---|---|---|
| `PROJECT_CODE` | 4 | 6 | 7 | 0 |
| `PROJECT_STATUS` | 11 | 11 | 17 | 0 |
| `PROJECT_NAME` | 22 | 91 | 241 | 0 |
| `ABOUT_BODY` | 195 | 272 | 759 | 3 |
| `KEY_EVENTS_BODY` | 33 | 248 | 281 | **46 of 46** |

**`KEY_EVENTS_BODY` is multi-line on every single slide** — median 5 paragraphs, max 6. So the
`||` rule is not an edge case for that field, it is the normal case, and V5 must land before
it is onboarded. Confirmed rather than inferred from the `_x000D_` artefacts.

Also note `ABOUT_BODY`'s range: 195 to 759 characters is nearly 4×. A single target length
across the portfolio would be wrong for most of it.

---

## 5. Your three observations from Appendix A

- **Timeline group** — not yet checked. Will come with the E1 triage.
- **Four financial values** — not yet checked.
- **Three highlight statements carrying their own quarter markers** — not yet checked, and
  agreed this is the interesting one. `(Q1F25)` and `(Q1F26)` inside the prose while a
  `Quarter` column exists outside it is a genuine modelling question.

All three are E1 work and E1 comes after the first field is through. Flagged so they are not
assumed done.

---

## 6. Running now

`PROJECT_STATUS`, on a **copy** of the real 46-slide deck:

1. rename all five role tags to the E2 FieldIDs *(V1, first live run)*
2. read a 46-row long-format register, `FY26Q4`, `Approved` *(V3, first live run)*
3. feed it to the sync engine
4. **verify by re-reading the deck**, not by trusting the run's own report
5. expect: **12 corrected, 0 created, 0 failed**

Result in the next document. If the number is not 12, that is the finding and it will be
reported as one.
