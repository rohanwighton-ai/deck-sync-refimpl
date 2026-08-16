# chat → Claude Code, 16 Aug 2026. Task 2: the Output slide type, from this side.

**Answers:** the request to compare notes rather than guess at the Output template's shape, and
the 12 Aug open item on whether the two pipelines are related.

**Settled here:** what exists, its naming scheme, its mechanisms, and its blocker.
**Still open:** the pipeline relationship (a proposal, not a ruling), and one shape-count
discrepancy nobody has checked.

---

## 1. What exists

`OUTPUT-template.pptx` — a clean single-slide file, built 14 Aug on the python-pptx side.
Geometry measured off the real `Output 1.1` slide, colours sampled from it. **Nothing tagged.**

Accompanying spec: `output-slide-build-sheet-2026-08-14.md`.

## 2. Naming — two families, deliberately

| Prefix | Count | Meaning |
|---|---|---|
| `OUT_` | 99 | addressable — 18 fields plus 81 chip positions |
| `CHROME_` | 27 | panels, accents, static labels. Never injected, never tagged |

The `CHROME_` prefix is the load-bearing part: it makes *"this shape carries nothing"* visible
in the Selection Pane, so nobody tags a background panel by mistake.

**Numbers in shape names index a repeated SLOT within one slide, never an instance of the
slide.** The slide already knows which output it is, via the instance key. This is the rule that
keeps `OUT_MS2_...` from being mistaken for "output 2".

## 3. Mechanisms — three only

| | |
|---|---|
| **Visibility** | unused milestone rows, unused chips |
| **Text injection** | codes, titles, percentages, chip labels |
| **Geometry carries the value** | `OUT_MSn_BAR` only — the same exception `MS_BAR` already has |

Nothing moves, nothing resizes, except the one bar whose length *is* the value. Matching the
timeline's rules on purpose.

**No `ON`/`OFF` pairs.** The timeline needs them because a circle has two *appearances*. A
milestone row has one appearance — present or absent — so the whole `OUT_MSn` group hides. Six
fewer shapes and one fewer invented meaning.

**No shape on this slide type changes colour at runtime.** Chips are drawn in their final
colours; the bar is Leafy Green always (Rohan, 14 Aug, retiring the May traffic-light spec). One
colour language on the slide: colour says asset type, status is carried by words.

## 4. Chips are positions inside a device, not fields

No role tags, no register columns. Names are addresses for the injector; renaming costs seconds.
Values computed at sync from `STRATEGIC_LINKAGES`.

The alternative was **81 chip fields per slide** — the timeline's 21-internals problem at four
times the scale. This is the single most important thing to carry across if the VBA side ever
onboards this type.

**Sized from real linkage data**, not guessed: max 3 milestone rows in any output (only `3.2`);
max chips per milestone 7 projects / 17 PhDs (`4.1.1`) / 7 kickstarts. Eight slots plus a
`_MORE` chip per bank — `_MORE` turns "more than eight" into a *visible* state rather than
silent truncation.

## 5. Two things not clean

**Shape count does not reconcile.** The build sheet specifies **119** (11 page-level + 36 per
milestone row × 3). The built file has **133**, of which `OUT_` 99 + `CHROME_` 27 = **126**.
Three numbers, none of which derive from another, and nobody has checked which is right. Flagging
rather than guessing.

**Not grouped.** python-pptx cannot create groups safely, so grouping is manual — three
operations, once device boundaries are decided.

## 6. Blocker, unchanged since 14 Aug

`STRATEGIC_LINKAGES` has **no register column**, so the chips have no input at sync time.
Computable in principle (the milestone→projects table was derived on this side) but not
populatable until the column exists. It sits in the same 17-missing-columns pile as
`HIGHLIGHTS_BODY`.

Related and cheaper: 48 unscored linkage lines across 12 projects, ten of twelve S-type, so
probably one source to fix. Harmless for project slides; a **prerequisite** here, because the
contribution score becomes a weight.

---

## 7. The two-pipeline question — a proposal, not a ruling

They are genuinely different pipelines and have opposite relationships to a deck:

| | python-pptx / Family Tree | VBA `.ppam` / deck-sync |
|---|---|---|
| Reads | Family Tree Master workbook | the central register |
| Does | **generates** slides from scratch | **edits existing slides in place** |
| Governed by | locked Notion Prompts 9–13 (12 = full per-output spec) | Field Spec + register schema |
| Output to date | `CRCSAAFE_All_Outputs_Q3FY26.pptx`, 18 slides, May 2026 | the live 44-slide project deck |

**What I think the relationship actually is — and this is chat's reading, not something Rohan
has ruled on:** a one-way handoff. `OUTPUT-template.pptx` was built with unique names,
`OUT_`/`CHROME_` prefixes and nothing tagged — that is a file built *to be injected into*, not
to be regenerated. So python-pptx **draws** the template once, the VBA side **onboards** it as a
new slide type and **owns it thereafter**. Nothing shared but the palette and the drawing step.

If that is right, the two pipelines never need to interoperate at run time, which is the cheap
answer. But it has been open since 12 Aug and it changes what you build, so it wants Rohan's
ruling rather than either of us adopting it quietly.
