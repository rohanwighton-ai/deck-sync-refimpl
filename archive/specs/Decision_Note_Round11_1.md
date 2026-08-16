# Decision note — four rulings for round 11

**From:** the Excel side, 31 July 2026
**Not a design document.** Rulings only, answering what round 11 asked for. Nothing new
proposed.

**Read §2 before running.** It changes your predicted number.

---

## For the Research Manager

| # | Status |
|---|---|
| **RM3** | Still the only open item. Both sides recommend one workbook, many decks. |
| **RM4** | Superseded by evidence: three `-2` slides, all byte-for-byte duplicates, nothing to harvest. The decision is now simply confirm-and-delete. |

---

## 1. The fourth class — accepted, named **Controlled**

Chrome / entity-static / quarterly / **controlled**.

Definition: tagged, one row per quarter, value drawn from a fixed list rather than drafted.

Field Spec consequence: a controlled field carries an **`AllowedValues`** column and no target
length, tone note, or exemplar. Roughly half the sheet is `n/a` for it, as you say — that is
the class doing its job rather than a defect in the sheet.

You were right that this changes the field we chose to go first. It also means
`PROJECT_STATUS` proves the join, the register read and injection, but **exercises the Copilot
drafting surface not at all.** See §4.

---

## 2. Canonical vocabulary — and it changes your prediction from 12 to 19

Three values, title case throughout:

```
In Progress
Not Started
Project Closed
```

`Not yet commenced` maps to `Not Started`.

**The reason this matters before you run:** a canonical set picked by frequency would be
`In Progress` / `Not started` / `Project Closed` — title case for one, sentence case for
another, which is the drift being fixed, preserved into the thing doing the fixing. Consistent
casing is worth one extra correction per slide now and never thinking about it again.

Recomputed against your counts:

| Current | Count | Action |
|---|---|---|
| `In Progress` | 21 | keep |
| `In progress` | 10 | correct |
| `Not started` | 8 | correct |
| `Not Started` | 1 | keep |
| `Not yet commenced` | 1 | correct |
| `Project Closed` | 5 | keep |

**Revised prediction: 27 unchanged, 19 corrected, 0 created, 0 failed.**

Stated before the run so it is a prediction rather than a reading. If the number is not 19,
that is the finding.

---

## 3. `PROJECT_CODE` — agreed, drop it from the register

Measurement is unambiguous and the field stores the row's own key. Drop it from the register
and populate it from `EntityCode` at injection.

One condition, which is yours already: **after the three `-2` slides are resolved.** Until
then the field and the tag genuinely disagree on those slides, and deriving would overwrite
the visible value with the suffixed one.

---

## 4. Target length — one number per field will not work, and the fix is already in the design

`ABOUT_BODY` at 195 to 759 characters is close to 4×, and a portfolio-wide target would be
wrong for most of the portfolio. So:

- the Field Spec's target length is **a range with a median anchor**, not a single number, and
  it is **advisory** — which F1's answer already established, since nothing can enforce it at
  injection;
- the real anchor is **the entity's own prior-quarter value**, sitting in the Copilot template
  sheet beside the drafting cell. A project whose About text ran to 700 characters last
  quarter should not be pulled toward a portfolio median of 272.

The portfolio figures you measured are still useful — as a drift signal for `CharCount`, and
for spotting the 759 as either legitimate or as something that needs editing on its own merits.

---

## 5. Consequence of §1 — the second field matters more than usual

Because `PROJECT_STATUS` is an enum, the first field through proves the mechanical half of the
pipeline and none of the content half. Nothing about the drafting surface, the exemplar
column, or the Copilot loop gets tested by it.

**Recommend `ABOUT_BODY` second, not `KEY_EVENTS_BODY`.** Your §4 measurement is the argument:
`ABOUT_BODY` is multi-line on 3 of 46, `KEY_EVENTS_BODY` on 46 of 46. Taking `ABOUT_BODY`
second proves the content half against mostly single-line prose, leaving `KEY_EVENTS_BODY` —
which needs V5 and is the hardest field on the slide — to run third through a path proven
twice.

---

Nothing else needed from this side. Report the number.
