# Columns — what the register holds, and why

> **CURRENT — the design record.** Written 2026-08-11 from a real work-deck slide. Once
> the Field Spec sheet carries the column list, THE SHEET is authoritative for the list
> and this file keeps only the reasoning.

Derived from a **real work-deck slide** (`1_P006`, screenshot 2026-08-11), not from the
rig. The rig's nine columns were an artefact of two weeks of scaffolding; the real slide
carries roughly 35–40 addressable values.

**This file is the DESIGN RECORD — the decisions and the reasoning.** Once the Field Spec
sheet carries the column list, *the sheet is authoritative for the list* and this file
keeps only the why. Do not maintain the same fact in both: a sentence cannot fail a test,
and the sheet can.

---

## The columns

| Column | Where the value comes from | Varies per period? |
|---|---|---|
| `PROJECT_CODE` | given | no |
| `PROJECT_NAME` | given | no |
| `SUBTITLE_A` | Rohan | rarely |
| `SUBTITLE_B` | Rohan | rarely |
| `SECTOR` | Rohan | rarely |
| `TRL` | Rohan | rarely |
| `PROJECT_STATUS` | Rohan | **yes** |
| `STRATEGIC_LINKAGES` | dedicated source | **yes** |
| `STRATEGIC_ALIGNMENT_BODY` | drafted | no |
| `ABOUT_BODY` | drafted | no |
| `PROBLEM_BODY` | drafted | no |
| `HIGHLIGHTS_BODY` | drafted | **yes** |
| `PROGRESS_BODY` | drafted | **yes** |
| `KEY_EVENTS_BODY` | drafted | **yes** |
| `PROJECT_PROGRESS` | project deliverables file | **yes** |
| `MS1..MSn` (label / date / done) | project deliverables file | **yes** |
| `START_DATE` | given | no |
| `END_DATE` | given | no |
| `INDUSTRY_CASH`, `SAAFE_CASH`, `IN_KIND`, `TOTAL_VALUE` | Rohan | no |
| team ×3 (role + name) | Rohan | no |
| `DELIVERABLES_BODY` | Rohan | no |
| `PROJECT_PHOTO` | Sources citation (picture) | rarely |
| `DELIVERABLE1_PHOTO`..`DELIVERABLE4_PHOTO` | Sources citation (picture) | **yes**, as new outputs appear |

> **2026-08-19: this table was already stale before tonight** — `PROJECT_PHOTO`
> (added `d013b64`, 2026-08-18) was missing entirely. Added both rows now, but
> per this file's own header, Field Spec is the authoritative list once
> populated — don't trust this table over the live sheet.

## Derived — never a column

- **Time elapsed** (the horizontal bar and its percentage) is computed from `START_DATE`
  and `END_DATE`. An hour was spent treating it as a synced field on 2026-08-10 before
  this was noticed. If it is stored, it goes stale silently while the dates beside it
  stay correct.

---

## Decisions, 2026-08-11

**The two percentages are different things.** The top-right figure (58% on `1_P006`) is
milestone/deliverable achievement and comes from the project deliverables file, entered
by hand. The bar underneath the dates is time elapsed and is derived. They had been
conflated; the drafting surface splits them:

- *milestone progress* — `PROJECT_PROGRESS` + the `MS*` timeline entries, entered each period
- *time elapsed* — `START_DATE` / `END_DATE` entered once, everything else computed

**`HIGHLIGHTS_BODY` is a real field that was missing.** Prompt 18 defines it, the slide
renders it as 2–4 quarter-tagged bullets, and the register had no column. Not scaffolding
— a gap.

**Strategic linkages become their own per-period field, from a dedicated source.**
Rohan, 2026-08-11: linkages "are apt to change / grow". Previously the codes lived inside
the Strategic Alignment prose, which meant the whole panel went stale whenever a linkage
was added, and no tool could diff a code buried in a paragraph.

Splitting them means `STRATEGIC_ALIGNMENT_BODY` stays **standing** — the argument for why
a project matters does not change quarterly — while the codes that do change carry their
own column, their own source and their own owner. It also gives the long-running `[TBC]`
a home: the declared codes are already visible on the real slide (`1.1.1`, `2.1.1`,
`3.3.1`, `3.4.1`), they simply had nowhere to live.

**The two `[TBC]` subtitle values are kept as static, rarely-changing fields.** Their
original purpose is no longer remembered, which is itself the argument for giving them a
row in the Field Spec: a field nobody can name is a field nobody can maintain.

**Standing versus per-period is the split that earns the register its keep.** Most of the
real slide is standing — money, dates, team, deliverables, TRL, sector. Entered once per
project, barely moves, needs no drafting sheet. Only six columns genuinely change each
quarter, and those are the ones the quarterly evening is actually spent on.

---

## Open

- **`PROJECT_STATUS` casing disagrees with the deck.** The slide reads `In progress`; the
  register and the Field Spec's allowed values read `In Progress`. One is wrong and the
  mismatch is silent. Settle against the real deck, not the rig.
- **Which source is the dedicated one for `STRATEGIC_LINKAGES`**, and who maintains it now
  the Family Tree is going. Costs a minute to ask at work.
- **Whether Copilot can open a SharePoint path** handed to it. The Sources sheet points at
  documents rather than holding their text, so if it cannot, a citation means "I read
  this" rather than "the model read this" — an attestation, not a trace.
- The rig's `Q4F26` rows are all character-identical to `Q3F26`. Nothing there is real Q4
  content; do not read a full-looking register as progress.
