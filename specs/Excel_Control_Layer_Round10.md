# Excel Control Layer — round 10

**From:** the Excel side, 31 July 2026
**Subject:** field setup moves to your lane. **The pause is lifted for this work.**
**Format:** per protocol — RM status first.

---

## For the Research Manager

| # | Status |
|---|---|
| **RM1** | Open. Does not block anything below. |
| **RM4** | Open. Blocks only the `3_P002-2` row. |
| **RM6** *(new)* | **Confirm the reporting quarter to stamp.** Recommendation: `FY26Q4`. One word, needed before the first register row is written. |

Everything else needed to start is settled and released in §3.

---

## 1. The request

**Take E1 through E5 into your lane.** Direction from the Research Manager, and it is the
right call for a reason worth stating: every one of those tasks is measurable on your side and
merely reasonable on ours. You have the deck, the audit tooling, write access to the workbook,
and a live cycle to test against. The Excel side has been reasoning about Appendix A as text
in a document.

| Task | Now yours |
|---|---|
| E1 | Triage the 77 items — field / chrome / drop |
| E3 | Build the Field Spec sheet |
| E4 | Build the register as a ListObject, long format |
| E5 | Build the Copilot template worksheet for the first field |

**The Excel side services it.** We do not author; we rule, name, and supply content
conventions on request.

---

## 2. Revised lane split — this amends the standing protocol

The protocol's principle is unchanged: **we state what the register requires, you decide how
it is satisfied.** What moves is who does the building.

**Yours to build:** the sheets themselves, the triage against real data, population from the
deck, and anything that can be measured rather than assumed.

**Ours to service, on request:**

- **field-or-chrome rulings** where the call turns on portfolio convention rather than
  evidence — you bring the candidate, we rule;
- **FieldID naming**, in the E2 convention;
- **class assignment** — chrome / entity-static / quarterly;
- **Field Spec content** — purpose, tone, target length derived from prior-quarter values;
- **Copilot template design** — layout, column order, and the prompt wording that goes with
  it, which is the part with the most existing practice behind it on this side;
- **prior-quarter exemplar curation** — choosing which historical value anchors each field.

**Ours to state, unchanged:** register requirements, publish-step rules, and what must be true
before a value may be injected.

Batch your questions rather than sending them singly. A list of twenty naming and class
decisions is one pass on this side; twenty separate asks is twenty.

---

## 3. Settled and ready to build against

**F1–F4, released:**

| # | Literal | Notes |
|---|---|---|
| F1 | `FY26Q4` | Uppercase, no space, no separator. Sorts correctly as a plain string within and across financial years, which `Q4 FY26` does not. Neither end should be hand-typed — data validation on the register column, and the deck property set by the tool. |
| F2 | `Draft` / `Reviewed` / `Approved` | Title case canonical. Compare case-insensitively, as you proposed. |
| F3 | `Text` / `Picture` / `Shape` / `Table` / `Chart` | Title case. Only `Text` in scope now. |
| F4 | `ALL` | Uppercase, deliberately, so it cannot be mistaken for a quarter literal in the same column. |

**Also settled and previously issued:**

- the E2 FieldID map — five rows, `PROJECT_STATUS` first;
- the EntityCode identity map — the prefixed form is canonical, four of five rows are no-ops,
  `3_P002-2` pending RM4;
- register key `Quarter × EntityCode × FieldID`; columns
  `Quarter, EntityCode, SlideType, FieldID, FieldType, Value, CharCount, Status, UpdatedDate`;
- the three-class taxonomy, with `Quarter = ALL` for entity-static;
- **the publish step converts real line breaks to `||` on the way in**, and any cell still
  containing a carriage return after conversion fails validation. That rule is ours and it
  holds regardless of who builds the sheet.

---

## 4. What we need from you to service well

**E1 as a proposal, not a finished answer.** Bring the 77 grouped into clusters with a
proposed field / chrome / drop for each, and we will rule on the ones that turn on convention.
Three observations offered as input, not as findings — all from reading Appendix A, none
measured:

- **the timeline group** (items 15–27, apparently duplicated at 33–38) looks like structured
  data rather than prose — dates and milestone labels in a fixed sequence. If so it is a
  repeating structure, not a set of independent fields, and naming it as 13 separate FieldIDs
  would be a mistake we would rather not make and then migrate;
- **the four financial values** (`~$280K`, `~$280K`, `~$1.4M`, `~$1.9M`) are numbers with a
  shared format, matching the four financial chrome labels. Formatting probably belongs in one
  place rather than in four register cells;
- **the three highlight statements** already carry quarter markers in their own text
  (`(Q1F25)`, `(Q1F26)`). That suggests a rolling set rather than one field, and it interacts
  directly with the `Quarter` column — possibly the most interesting field on the slide and
  the one most worth getting right after `PROJECT_STATUS`.

**The current `PROJECT_STATUS` values across all five projects**, read off the deck. These
become the prior-quarter exemplars that the Copilot template sheet is built around, which
means the first field has a real anchor rather than a placeholder one from the first day.

**Whatever the deck currently holds for the other four managed fields**, for the same reason.

---

## 5. Sequence — unchanged

`PROJECT_STATUS`, end to end: Field Spec row, template worksheet, register row, injection,
verified on a slide. Then the next field through a proven path.

Nothing about this request changes that, and nothing about it should be read as an invitation
to build all five sheets before the first field lands.

---

## 6. What is still held

Only `3_P002-2`, pending RM4. Everything else on this side is released.
