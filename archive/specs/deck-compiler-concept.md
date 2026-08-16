# The deck compiler concept

Status: **design, not built.** Written 2026-07-30, the evening the first complete sync
cycle ran. Supersedes nothing; extends `sync-operations.md` and `deck-registry.md`.

---

## 1. The idea

Rohan's framing: keep source decks **simple and singular to slide type**, then use markup
in the spreadsheet plus a **deck compiler** to assemble composite output decks — which
child slides get pulled in from which source decks (Output, Milestone, Project Progress),
in what order, and how period-over-period reporting is handled.

The reframe underneath it: **a deck stops being a document and becomes a build artifact.**
You never edit the composite. You edit data and rebuild. This is the same instinct as the
project-wide preference for "the source IS the artifact" (zettel
`20260704-editable-representation-beats-rendered-output`), applied one level up.

## 2. Two engines, currently conflated

| Engine | Answers | Status |
|---|---|---|
| **Content** | what values go in the fields of a slide? | built, ran end to end 2026-07-30 |
| **Composition** | which slides exist, from where, in what order? | not built |

The content engine already has a toe in composition: `ResequenceByRowOrder` lets the Data
sheet dictate slide order, and it demonstrably works (`Resequenced 2 slide(s)`, first
observed 2026-07-30). That is the whole concept in miniature, already running.

## 3. What a row can say — the MECE decomposition

Six categories. Mutually exclusive; believed collectively exhaustive over "things the
spreadsheet can assert about a slide."

| # | Category | Question | Today |
|---|---|---|---|
| 1 | Identity | which thing is this? | done — `instance_key` |
| 2 | Content | what values go in its fields? | done |
| 3 | Existence | should a slide for this exist at all? | not started |
| 4 | Provenance | which source deck / slide type? | partial — `slide_type` |
| 5 | Position | where does it sit in the output? | partial — row order |
| 6 | Time | which period, and what does a new one create? | not started |

## 4. Decisions made

### D1. Time model: one dated row per period (2026-07-30)

A row is an **observation** (project-at-a-point-in-time), not an **entity** (project).
Chosen over overwrite-in-place and over per-period sheet snapshots.

Consequences:
- Any past deck is rebuildable exactly. Under overwrite it never can be — reversing that
  choice later means reconstructing history that no longer exists.
- Trends become possible at all.
- Every read needs a period filter. The **deck declares its own period**; rows do not have
  to be filtered by hand.
- `New Period` stops being a slide operation and becomes "add rows".

### D2. Identity splits into three columns

`3_P002-2` is hand-rolled duplicate management — one project needing two slides — and it
overloads one field with three jobs. Split them:

| Column | Means | Stable |
|---|---|---|
| `instance_key` | this *slide's* identity | forever |
| `project` | which project it is about | forever |
| `period` | which quarter this row describes | per row |

This also retires the long-standing design debt that instance keys want to be immutable
GUIDs: they can be, because the human-readable identity moves to `project`.

### D3. One deck per slide type

Endorsed. It collapses more than it first appears: one deck = one type = one template =
one worksheet. The type picker disappears; `SyncNowCore`'s multi-type loop disappears.
And every source deck then has exactly **one** template slide — which is where the
master-template-that-is-never-a-real-project lives naturally. The "created slide inherits
another project's untagged figures" hazard stops existing by construction rather than
being fixed.

## 5. Prior art (scanned 2026-07-30)

### think-cell has shipped this concept, and formalised it

`.ppttc` — IANA-registered media type `application/vnd.think-cell.ppttc+json`, registered
2018 by Arno Schoedl. There is a published `ppttc-schema.json` and a CLI (`ppttc.exe`).
Structure:

```json
[
  { "template": "quarterly.pptx",
    "data": [ { "name": "SlideTitle", "table": [[{"string": "Competition: Germany"}]] } ] },
  { "template": "quarterly.pptx",
    "data": [ { "name": "SlideTitle", "table": [[{"string": "Competition: France"}]] } ] }
]
```

- Top-level **array**; each object is one copy of a template.
- Elements are bound **by name** — the same design this project already uses via shape tags.
- Cell types: `string`, `number`, `date`, `percentage`, `fill`.
- **"Templates appear in the order that you specify in the array."**
- Repetition = list the same template again with different data.

**The design lesson, and it corrects our instinct.** The original framing here was
"switches ganged at the left of the Excel row". think-cell does not do that. Composition
lives in a **separate ordered manifest**; the data rows stay pure facts. Order and
repetition are properties of the *list*, not flags on a record.

Recommended shape for us — two sheets, not more columns:

| Sheet | Holds | MECE categories |
|---|---|---|
| **Data** | one row per project per period, facts only | 1, 2, 6 |
| **Build** | one row per output slide, in output order | 3, 4, 5 |

The Build sheet can be generated from Data for the common case, so the normal quarterly
run stays one click. Hand-editing Build is how you get audience cuts (board / funder /
internal) from identical Data.

### UpSlide and Empower occupy the same territory

Both sell slide libraries plus Excel-linked content that updates in one click. Neither
appears to expose a data-driven *manifest* as the assembly mechanism; the assembly is
interactive. That is the gap this concept sits in, and it is a real one.

## 6. The technical ceiling — this matters most

**VBA cannot control formatting when moving slides between presentations.**

- `Slides.InsertFromFile` has no keep-source-formatting option. The UI offers it as a
  smart tag, and smart tags are not exposed to the object model at all.
- The available workaround is assigning the source `.Design` to the pasted slide. Its
  cost is documented and real: **duplicated slide masters and file bloat**, which is
  actively bad when every source deck shares one corporate template.

**The modern Office JavaScript API has exactly the missing capability:**

```js
const opts: PowerPoint.InsertSlideOptions = {
  formatting: "KeepSourceFormatting",   // or "UseDestinationTheme"
  targetSlideId: "267#",                 // exact insertion point
  sourceSlideIds: ["267#763315295", "256#"]   // which slides to take
};
context.presentation.insertSlidesFromBase64(base64Pptx, opts);
```

Caveats found: inserted slides keep the **source's relative order regardless of array
order**, so output ordering needs either one call per slide or a reposition pass. The
source presentation must be supplied base64-encoded.

Microsoft's own guidance for this API states that the add-in "must create and maintain a
data source that correlates the selection criterion with slide IDs" — which is precisely
the Build sheet above. The instinct matches Microsoft's documented pattern for the API.

**Therefore: the compiler probably does not belong in VBA.** The content engine is
correctly a `.ppam` and works. The composition engine hits a wall in VBA that Microsoft
has explicitly filled in the JS add-in API. This is a genuine fork and should be decided
before step 4, not during it.

## 7. Progression ladder

Ordered by dependency and risk.

1. **Master template slide, never a real project.** Cheap, fixes a live hazard, made
   natural by D3.
2. **`period` column + deck-declares-its-period.** Makes D1 real. Unblocks the rest.
3. **Existence + Position (Build sheet).** Position is half-built — resequencing works.
   Still entirely within one deck, still VBA, still an extension.
4. **Provenance / cross-deck composition.** The genuinely new engine, and the one with the
   technology fork above. Spike before designing.
5. **Compiler.** Assembles the composite from 1–4.

Steps 1–3 extend what runs today. **Step 4 does not** — the honest correction to "this is
just an extension of what you are already doing."

## 8. What it unlocks

- Multiple audience cuts from one dataset, as different Build sheets over identical Data.
- Any past deck rebuildable exactly — impossible under the overwrite model.
- Trends, because there is finally a time series.
- Consistency by construction: a figure cannot be current in one deck and stale in another.
- **The sleeper:** reporting content stops being trapped in PowerPoint. Once it is rows it
  is queryable by anything — which is the last mile of the SAAFE three-layer data map
  pitch actually delivering, and the stated reason this tool exists.

### D4. Composite decks are strictly generated; child decks may be two-way (2026-07-30)

Rohan's call. The composite is a build output — never edited, always rebuilt. Two-way on
the composite would invite the "which one is truth" question that kills tools in this
category, and it is the one place where the answer is genuinely unrecoverable, because a
composite slide has no single owning row.

**Child decks are a different case and two-way there is reasonable** — that is where a
human legitimately edits a real slide.

Status of that today: **the machinery exists, the operation does not.** Slide values are
harvested into the sheet by `OnboardFlow.CommitOnboarding` (via `ExcelOutput.UpsertRow`)
and by `DeckAdoption`, but **only at onboard/adopt time**. Nothing takes a slide edit made
afterwards and pushes it back. So this is a new operation built from two halves that both
already work and are tested.

**The hard part is not the write, it is the conflict policy.** Today
`SyncOperations.PlanRoutineSync` has exactly one interpretation of "slide text differs
from row value": the slide is stale, correct it. Excel always wins. Two-way needs to tell
apart three states, and cannot do it from the current two values alone:

| Slide | Row | Correct action |
|---|---|---|
| changed | unchanged | push slide → row |
| unchanged | changed | push row → slide (today's only behaviour) |
| changed | changed | genuine conflict — must ask, never guess |

Distinguishing them needs a **last-synced value per field**, stored somewhere durable.
Without it the engine can only see that the two sides differ, not who moved — and an
engine that guesses here silently discards someone's typing.

This lands well with D1: under the dated-row model, the previous period's row is already
an immutable record of what was last agreed, so the snapshot may be mostly free. Worth
checking before designing a separate mechanism.

**Resolution is human, and it is a grid — not a prompt.** Rohan's rule: a disagreement
between a slide and its row must be settled by a human looking at *both* values; but that
must not become per-instance questioning during batch operations, "where repeated
questioning at every instance kills the user experience."

Those two constraints only conflict if the interface is a dialog. They are both satisfied
by the pattern this project already built and proved on 2026-07-29/30: a **review grid**
written into the workbook, listing every conflict at once with both values side by side
and a decision column, read back in one pass. `WriteInstanceKeyGrid` /
`ReadInstanceKeyGrid` / `SlidePreviewText` in `BatchOnboardFlow.bas` are the working
precedent — built for exactly this reason, to replace a 45-prompt InputBox chain where a
single mistake discarded the lot.

So the conflict UI is a **conflict review grid**:

| slide | field | value on slide | value in row | keep |
|---|---|---|---|---|
| 3 | Project Status | In Progress | Closed | slide / row |

Design rules that follow:
- Never auto-resolve a both-changed conflict. Present it.
- One-sided changes need no prompt at all — direction is unambiguous, so they apply
  silently and are reported afterwards, not asked about beforehand.
- The grid appears only when there is at least one genuine conflict. A grid that opens
  every sync to say "nothing to decide" trains people to close it unread, and is the same
  failure as an always-true guard (zettel
  `20260729-an-always-true-guard-is-worse-than-no-guard`).

## 9. Open questions

- **Q1.** Where does the compiler run — VBA `.ppam` with `.Design` and master bloat, or an
  Office JS add-in with `insertSlidesFromBase64`? Blocks step 4.
- ~~**Q2.** Is the composite ever editable?~~ **Decided — see D4.**
- **Q3.** Does the Build sheet get generated from Data by default, or authored? (Suggest:
  generated, hand-editable.)
- **Q4.** How are retired projects handled — an Existence flag, or absence of a current
  period row? The second is cheaper and self-maintaining.
- **Q5.** Do we read the `ppttc-schema.json` and borrow its vocabulary wholesale rather
  than inventing column names?
- **Q6.** For child-deck two-way (D4): can the previous period's row serve as the
  last-synced snapshot, or is a separate per-field record needed?
