# Worked example — STRATEGIC_ALIGNMENT, end to end

> **HISTORICAL — written 2026-08-08.** Names the old multi-button toolbar and the
> drafting layout of the time. Kept because the METHOD is the point: one field, one real
> project, recipe through to text. Do not follow its button names or column letters.

Gist: one field, one real project, taken all the way from writing the recipe to producing the text — so we can see whether the recipe actually works before doing it 43 times.

Project: **`2_P004` — Bayesian Network Models for Effective AMR Management in Water
Systems**, slide 4 of the rig deck, period `Q4F26`. Everything below came from the rig's
own files, read as bytes from copies. Nothing was written to the register or the deck.

---

## Step 1 — the recipe (a Field Spec row)

Transcribed from Prompt 18's locked definition, plus the four evidence columns.

| Column | Value |
|---|---|
| **FieldID** | `STRATEGIC_ALIGNMENT` |
| **Kind** | `Prose` |
| **Purpose** | The "so what" for SAAFE CRC and Australia. Connects the project's outcomes to SAAFE's strategic objectives, the declared Commonwealth milestone/output codes, and the broader sector/AMR benefit. |
| **Voice** | Taciturn, interpretive, professional. Forward-looking, value and impact oriented. No promotional language, no hedging, no padding. Assumes AMR literacy. |
| **Length** | About two short paragraphs. Target 600–800 characters. |
| **Own-job test** | Does it say why this MATTERS to SAAFE and Australia — without describing the technology (that is ABOUT) or setting out the sector gap (that is PROBLEM)? |
| **Do NOT** | Describe the approach or technology in technical detail. Restate the sector gap. Present an inferred linkage as a declared one. Invent codes, figures, organisations or outcomes not in the workbook. |
| **Evidence required** | Linkage codes, from the maintained codes list (see note below) |
| **Cite at least** | 1 |
| **Period rule** | Any period — strategic alignment is standing, not quarterly |
| **When absent** | `[TBC]` and flag. Never infer a code. |

The last four are the new part. They say what has to be *true* before this field can be
published, not just how it should read.

### Note on the codes list — and on the word "declared"

**The list can be built by harvesting the codes off the existing slides.** Rohan's call,
2026-08-08, and it is right: the register is the curated truth and sync obeys it, so a
harvest that a person reviews is the same move as onboarding harvesting slide text into
register rows. A human curates in the middle; the register governs afterwards; the slides
become output rather than an independent thing being checked.

What the harvest does NOT establish is that a code was ever *declared* in a contract. It
establishes what the deck currently says. Those are two claims, and the field should only
make the one it can support:

- `Codes in use, curated by <owner>, as at <date>` — honest, and startable today.
- `Declared Commonwealth linkage codes` — a stronger claim, needing the contractual list.

Start with the first. If a contractual list later appears, reconcile against it — which is
a real check, and only possible because the register never claimed to be it.

**"Declared" is a word doing two jobs** (declared in a contract / what we have been using),
which is the shape that has cost this project twice already — `static` meaning both "not
authored by a person" and "does not change between periods". Keep the two apart in the
column name, not just in your head.

**Owner:** Rohan initially, a few selected colleagues later. Once there is more than one
maintainer, `as at` stops being decoration, and the list probably wants to live outside
the register workbook — which the tool rebuilds, clears sheets in, and never backs up.

---

## Step 2 — the prompt the tool generates from that row

This is what `FieldSpec.PromptFrom` would emit into cell L2, carrying the whole recipe so
it works with Copilot at work or Claude at home, with nothing held elsewhere.

```
Read the workbook and the existing text in column C.

WHAT THIS FIELD IS FOR
The "so what" for SAAFE CRC and Australia. Connects the project's outcomes to
SAAFE's strategic objectives, the declared Commonwealth milestone/output codes,
and the broader sector/AMR benefit.

VOICE
Taciturn, interpretive, professional. Forward-looking, value and impact oriented.
No promotional language, no hedging, no padding. Assumes AMR literacy.

LENGTH
About two short paragraphs. Target 600-800 characters.

ITS OWN JOB
Does it say why this MATTERS to SAAFE and Australia -- without describing the
technology (that is ABOUT) or setting out the sector gap (that is PROBLEM)?

DO NOT
Describe the approach or technology in technical detail. Restate the sector gap.
Present an inferred linkage as a declared one. Invent codes, figures,
organisations or outcomes not in the workbook.

EVIDENCE
This field requires: declared linkage codes, from an authoritative list.
Cite at least 1 source ID from the Sources sheet, in column G.
A source from any period is acceptable for this field.
If the evidence is not present, write [TBC] and say so. Never infer a code.

Write an updated version of STRATEGIC_ALIGNMENT for each project into column F ONLY.
Leave every other column untouched.
```

---

## Step 3 — the input material

What the assistant is given, all of it from the workbook and the slide:

**Register row (`2_P004`, `Q4F26`)**
- `PROJECT_NAME` — Bayesian Network Models for Effective AMR Management in Water Systems through various customer uses
- `PROJECT_STATUS` — In Progress
- `ABOUT_BODY` — Develops Bayesian network models that assess antimicrobial resistance risk across water systems. Expert-derived causal diagrams, monitoring data and machine learning are combined to model transmission through wastewater treatment, water reuse and biosolids pathways, and to identify control points and data gaps.
- `KEY_EVENTS_BODY` — Causal diagrams completed for all priority water pathways (Q3F26) · Screening risk assessments completed (Q3F26) · Economic assessment milestone adjusted (Q3F26) · EDAR8 workshop, Ozwater26 presentation (Q1F27), PAC meeting scheduled (Q4F26)

**Current slide text for this field** (1,113 characters), which mentions
`(1.4.2, 1.5.2, 2.3.1)` as declared linkages and `(2.1.1)`, `(2.2.1)` as contributions.

---

## Step 4 — the output

```
STRATEGIC_ALIGNMENT  --  2_P004

Gives regulators and water utilities a defensible basis for deciding WHERE to act
on AMR, rather than monitoring broadly and intervening on judgement. Quantified
transmission pathways and identified control points turn wastewater treatment,
water reuse and biosolids decisions into ranked options that can be defended to a
regulator and revisited as evidence changes.

For SAAFE it supplies the risk-assessment layer other programs draw on: a shared,
testable account of how resistance moves between environmental, agricultural and
human-health systems, against which surveillance designs and intervention
proposals can be checked before they are funded.

Declared linkages: [TBC]
```

**Characters:** 723, inside the 600–800 target.

### The flags it raised

**1. The linkage codes could not be confirmed, so they were not written.**
The recipe requires declared codes from an authoritative list. Prompt 18 says declared
codes come from the workbook's Strategic Alignment sheet. **That sheet does not exist in
this register** — it holds `Instance ID`, `Quarter`, `PROJECT_CODE`, `PROJECT_NAME`,
`PROJECT_STATUS`, `ABOUT_BODY`, `KEY_EVENTS_BODY` and nothing else. The codes on the
current slide may well be right, but nothing available here can distinguish a declared
code from one somebody inferred two years ago, and the recipe explicitly forbids
presenting the second as the first. So: `[TBC]`.

**2. No source could be cited, because no real source exists.**
`Cite at least 1` cannot be satisfied — the Sources sheet holds six fabricated rows and
nothing else.

**3. The existing slide text fails its own recipe, in two ways.**
It runs 1,113 characters against a 600–800 target. And its first sentence — *"developing
integrated, system-level risk assessment models"* — is ABOUT's job, not this field's. The
own-job test catches it: the paragraph describes the technology before it says why it
matters. This is exactly the bleed Prompt 18 warns about between Strategic Alignment,
About and Problem.

---

## What this proves, and what it doesn't

**Proved.** The recipe is transcribable from the library into a Field Spec row without
loss. The generated prompt is self-contained — it names no external document and would
work pasted into Copilot at work. The evidence rules fired on the first real project
tried, and they fired *correctly*: they stopped two codes being restated as authoritative
when nothing on hand could confirm them. That is the control working before any content
run, not after.

**Not proved.** Nothing has been through the tool. This did not touch `Setup A2`, a
drafting sheet, publish, or sync — no shape on slide 4 is tagged `STRATEGIC_ALIGNMENT`,
so the field does not exist to the add-in yet. The draft above is text in a document.

**And the thing worth arguing about:** the output is *shorter and emptier* than what is
on the slide today. It refuses to assert the codes. If those codes are in fact declared
and simply live somewhere this register cannot see, then the recipe is right and the
**register is missing a field** — declared linkages belong in it, as data, not buried in
a prose paragraph where nothing can check them. If they were never declared, the current
slide has been overstating them.

Either way the question is now visible, which it was not an hour ago. That is the
argument for recipes in one example.

---

## To take it the rest of the way

1. Add the Field Spec row above to the `Field Spec` sheet.
2. `Setup A2: Discover Fields` on slide 4 — tag the Strategic Alignment body shape as `STRATEGIC_ALIGNMENT`.
3. `1. Drafting Sheets` — a `TPL_STRATEGIC_ALIGNMENT` sheet appears with the prompt in L2.
4. Paste the draft into column D, tick column E, `3. Publish`, then `4. Sync Now`.
5. Settle the linkage-codes question — it is now the blocker for this field, and it is a data question, not a writing one.
