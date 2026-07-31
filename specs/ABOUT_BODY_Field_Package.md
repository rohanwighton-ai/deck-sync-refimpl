# ABOUT_BODY — field package

**From:** the Excel side, 31 July 2026
**Contains:** a new standing requirement (§1), then the second field ready to wire in (§2–§4).
**Source for all field content:** Notion Prompt Library, Prompt 18 — the six-field definitions,
locked 1 July 2026. Not authored here; transcribed.

---

## For the Research Manager

| # | Status |
|---|---|
| **RM3** | **Ruled: one workbook, many decks.** Closed. |
| **RM4** | Ruled: delete all three `-2` slides. `3_P002-2`, `2_P004-2`, `1_P006-2`. |
| **RM7** *(new)* | Nothing. §1 below is a requirement issued, not a question. |

---

## 1. R13 — every replacement is seen by a human before it lands

**Direction from the Research Manager, and it is not optional.**

> A human must observe and approve every automated replacement before it reaches a slide,
> except where they are approving a verified batch of uniform changes.

**This is a second gate, not a stricter version of the first.** `Status = Approved` in the
register means *this text is right*. R13 means *this change to this slide is right*. Different
question, different moment. The `PROJECT_STATUS` run passed the first gate and skipped the
second — 19 slides changed and nobody saw a single before-and-after.

### R13.1 — no replacement without a before-and-after

Every value that would differ from what is currently on the slide is shown as
`current → proposed`, per slide, and is not written until approved.

### R13.2 — what may be batched, and it falls out of the class column

| Class | Batchable? | Why |
|---|---|---|
| **Controlled** | **Yes** | One transformation applied N times. The 19 `In progress` → `In Progress` corrections are a single decision. Approving them individually teaches the operator to click through, which is worse than not asking. |
| **Quarterly (prose)** | **No, never** | Every replacement is unique text. There is no transformation to approve, only content to read. |
| **Entity-static** | Individually, but rare | A project name changing is a real event and should feel like one. |
| **Chrome** | n/a | Never touched. |

A batch is approvable only when every instance in it shares one FieldID **and** one
transformation. Present it as: the transformation, the count, and the list of affected
entities.

### R13.3 — unchanged rows never enter the review queue

The 46-row run was 27 unchanged, 19 changed. Only the 19 are decisions. Putting the other 27
in front of someone is how a review becomes a formality.

### R13.4 — the review must be readable, not a modal

Prose review is real reading. It must be possible to leave it, come back, and finish it later
without re-running. A dialogue box that must be cleared in one sitting will be cleared without
being read.

### R13.5 — approval does not carry across runs

Per run, per change. An approval given last quarter approves nothing this quarter.

**Mechanism is yours.** `Preview Sync` already exists; the requirement is that it stops being
optional and gains the batch rule above.

---

## 2. Field Spec row — `ABOUT_BODY`

Transcribed from Prompt 18's **About** definition and the global prompt tenets.

| Column | Value |
|---|---|
| `FieldID` | `ABOUT_BODY` |
| `Class` | **entity-static** — `Quarter = ALL` |
| `FieldType` | `Text` |
| `Purpose` | The "what". A neutral, factual description of what the project is and does: the approach or technology, the target systems, and the aim. |
| `Voice` | Descriptive, present tense. Taciturn. No promotional language, no hedging, no padding. Assumes AMR literacy. |
| `Length` | ~1 paragraph. Measured across 46 slides: min 195, median 272, max 759 characters. **Target 200–350, advisory.** The real anchor is this project's own prior value, not the portfolio median. |
| `Multi-line` | Rare — 3 of 46. Uses `\|\|` where present. |
| `Source` | The workbook only. Where a needed field is missing or ambiguous, say so and ask — never infer or fill the gap. |
| `Own-job test` | Does it describe **what the project is**, without straying into why it is needed (Problem) or what it is worth (Strategic Alignment)? |
| `Do not` | Justify the project. Pitch its strategic value. Describe this quarter's activity. Invent facts, figures, organisations or outcomes not in the workbook. |

**One honest note on sequencing.** As an entity-static field, `ABOUT_BODY` exercises prose
drafting, the exemplar column and R13's individual review — but **not** the quarter-to-quarter
cadence, because it carries one row with `Quarter = ALL`. That half is proved by
`KEY_EVENTS_BODY` third. `ABOUT_BODY` is still the right second field: it is prose, it is
mostly single-line, and it does not need V5.

---

## 3. Register rows

One row per entity. 43 rows after the three `-2` slides are deleted.

```
Quarter  EntityCode  SlideType  FieldID      FieldType  Value       CharCount  Status  UpdatedDate
ALL      3_P001      project    ABOUT_BODY   Text       <text>      <calc>     Draft   2026-07-31
ALL      3_P002      project    ABOUT_BODY   Text       <text>      <calc>     Draft   2026-07-31
...
```

Seed `Value` from what is currently on each slide. That is not a shortcut — it is the prior
value that the drafting surface in §4 anchors against, and it means the exemplar column is
real from the first day rather than empty until next quarter.

---

## 4. Copilot template worksheet — `TPL_ABOUT_BODY`

One row per entity, left to right in reading order.

| A | B | C | D | E | F | G |
|---|---|---|---|---|---|---|
| `EntityCode` | `Project name` | `Current ABOUT text` | `Chars` | `RM source notes` | `Draft` | `Approved` |
| `3_P001` | *(from register)* | *(read-only exemplar)* | *(calc)* | *(RM types here)* | *(Copilot writes here)* | *(Y/N)* |

Column C is read-only and is the whole point of the layout — Copilot rewrites against a
visible exemplar far better than it generates cold.

**Prompt block, pasted above the sheet:**

```
Read the attached workbook and the ABOUT text already in column C.

ABOUT answers one question: what is this project and what does it do?

Write a neutral, factual description of the approach or technology, the target
systems, and the aim. Present tense. Descriptive. Roughly one paragraph, and close
in length to what is already in column C for that project — that text is the
standard, not a draft to improve on.

Do not justify the project — that is the PROBLEM field. Do not describe what it is
worth to SAAFE or Australia — that is STRATEGIC ALIGNMENT. Do not describe this
quarter's activity — that is LATEST PROGRESS COMMENTS. If the text in column C
already does its job, say so and leave it.

The workbook is the sole source of truth. Do not introduce facts, figures,
organisations or outcomes that are not in it. Where something needed is missing or
ambiguous, say so and ask — do not infer or fill the gap.

Taciturn voice. No promotional language, no hedging, no padding. Assume AMR
literacy.

Write into column F only. Leave every other column untouched.
```

**Publish step, unchanged:** column F values, where column G is `Y`, are pasted into the
register with `Quarter = ALL`, real line breaks converted to `||`, `CharCount` calculated, the
workbook saved. Any cell still containing a carriage return after conversion fails validation.

---

## 5. What is needed from your side

Nothing to decide. Two things to do:

1. **Seed column C and the register `Value` field** from the current slides — 43 entities.
2. **R13 before `ABOUT_BODY` is wired in**, not after. This is the first field where
   individual review actually bites, and running it once without the gate would establish
   exactly the habit R13 exists to prevent.
