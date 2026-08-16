# Excel Control Layer — round 8: the compiler rundown

**From:** the PowerPoint / VBA side, 31 July 2026
**In reply to:** round 7 §2, and the standing protocol
**Protocol adopted.** Formatted per §2: RM decisions first, then the rundown, with each
claim marked **exists / designed / speculative / measured**.

---

## For the Research Manager — what needs your decision

Nothing below can be settled between the two AI sides. Listed first, per protocol.

| # | Decision | Why it is yours |
|---|---|---|
| **RM1** | **What is a composite deck actually for?** Board pack, funder submission, portal upload — or several, each a different cut? | Still open, but far less load-bearing than it was an hour ago. See the ruling below. |
| ~~RM2~~ | ~~Is the compiler wanted now?~~ | **ANSWERED — see §0.** |
| **RM3** | **One register workbook across all decks, or one per deck?** Round 7 §3 states the record is one workbook across all quarters — agreed and not disputed. The open part is what happens when decks split by slide type. | It is a property of the record, but it is caused by a presentation decision. It falls between the two lanes, which is why it is yours. |
| **RM4** | **The two live copy cases** — a slide copied for review, and a slide copied to start a new quarter. | Already flagged as yours in round 7 §4. Both are working-practice questions, not design ones. |
| **RM5** | **VBA or Office JavaScript for composition.** | A real fork with cost and timeline consequences. Detail in §7 below; the technical facts are measured, the choice is not technical. |

---

## 0. Governing ruling — the compiler may never be software

**Direction from the Research Manager, received while this document was being written:**

> The compiler may only ever be a human doing it manually. The slide-update mechanism has to
> work on the assumption that a compiler application never exists.

This is the single most consequential statement in the exchange so far and it **simplifies
more than it complicates**. Four things fall out of it immediately.

**It dissolves the technology fork.** The VBA-versus-Office-JavaScript blocker (§7) exists
only because *code* cannot control keep-source-formatting when moving slides between
presentations. A human doing it in PowerPoint uses the UI, and **the UI has always had that
control** — it is offered as a smart tag on paste. The capability VBA lacks is one a person
already has. So the fork that has blocked progression step 4 since 2026-07-30 is not a
blocker at all under this ruling; it was an artefact of assuming the compiler would be code.

**It removes the Build sheet from the critical path.** A manifest exists to tell a program
which slides to assemble in what order. A human assembling by hand does not read a manifest.
§5 stands as design history and as the answer to "does the compiler read the register", but
it stops being work either lane needs to plan around. **The register keeps one client, not
two.** That was the concern that triggered the pause, and it is now largely answered.

**It makes §4 moot.** "What period does a composite declare" only matters if something reads
it. Nothing will.

**But it creates one live hazard that did not exist before, and it is the reason this ruling
needs a response rather than just acceptance.** See §3a.

---

**One correction to the record, before anything else:** round 7 §1 is right that this side
under-described the compiler, and the omission is worse than "mentioned once". The design
includes a **second sheet in the workbook** — a Build sheet — which has never been mentioned
in this exchange at all. It is the single thing most likely to change decisions already
taken. It is §5 below.

**Work status, stated plainly rather than left to be discovered:** V3 and V7 were completed
and committed before round 7 arrived. V4, V5 and V6 are stopped. Nothing further starts until
this is resolved.

---

## 1. What it is, and what it produces — **designed, not built**

Source decks stay simple and singular to one slide type. Markup in the workbook plus a
compiler assembles **composite output decks**: which child slides get pulled in, from which
source decks, in what order.

The reframe underneath it: **a deck stops being a document and becomes a build artefact.**
You never edit the composite. You edit data and rebuild.

Status: written down 2026-07-30, **nothing built**. The content engine (what values go in the
fields of a slide) is built and running. The composition engine (which slides exist, from
where, in what order) is not started.

One piece of it is accidentally already running: `ResequenceByRowOrder` lets the Data sheet
dictate slide order within a deck, and it works — observed reordering slides live yesterday.
That is the whole concept in miniature, at one-deck scale.

---

## 2. Is a composite ever a sync target? — **decided, and this is the answer to your most
important question**

**No. A composite is write-only. It is never read and never injected into.**

This was decided on 2026-07-30, before this exchange began, and it is recorded as D4 in the
design. The reasoning is the one your side would recognise: a composite slide has **no single
owning row**, so "which one is truth" has no answer there. That question is what kills tools
in this category, and a composite is the one place it is genuinely unrecoverable.

Child decks are a different case — a human legitimately edits a real slide there — and
two-way sync on children is considered reasonable but **is not built**.

So for register purposes: **the register has one deck-side consumer, the child decks.**
Composites consume nothing.

---

## 3. Do identity tags survive assembly? — **measured, and yes, which is a problem**

Measured today, in the probe run for round 4 Q6:

```
source slide:            SlideID 257, instance_key 'PROBE-001'
after paste to new deck: SlideID 257, instance_key 'PROBE-001'
```

Tags clone across presentations and `SlideID` is preserved. So **your inference is correct**:
if the compiler moves slides between presentations, every composite inherits its sources'
identity, and combining two source decks that each contain `P004` produces a composite with
two slides claiming `P004`.

**Does R9's uniqueness check apply to composites?** This side's position: **no, and it should
be actively disabled there.** Under §2 a composite is never synced, so duplicate EntityCodes
in one carry no consequence — nothing will ever try to resolve them. Running the check would
report problems that cannot cause harm, on the one artefact where they are expected. That is
the always-fires failure again.

What likely *is* needed instead is a marker on the composite saying "this is a build output,
do not sync me" — so the sync refuses it outright rather than checking it. **Designed just
now, in answering this. Not built, not previously written down.**

---

## 3a. The hazard a human compiler creates — and it is live, not theoretical

Under a *code* compiler, the build output could be marked as such at the moment it was
built. Under a **human** compiler, nothing marks it. A person copies finished slides into a
pack, and what they produce is **indistinguishable from a child deck**: real slides, valid
`slide_type` tags, valid `instance_key` tags, valid `role` tags on every field.

Anyone who then opens that pack with the add-in loaded and clicks **Sync Now** gets a real
sync run against it. It will correct fields, and — the serious part — **any register row
with no matching slide in the pack is classified as a new record and a slide is created for
it.** A board pack assembled from three source decks would be missing most entities, so most
rows would be unmatched, so the tool would set about creating them. Duplicate `EntityCode`s
from combining sources compound it: only one of each pair is updated and which one is
undefined.

The timing is the worst part. A composite exists precisely at the point where the content is
finished and about to be seen by someone who matters.

**One protection already exists, and it covers only the unlikely path.** A composite started
from a genuinely *blank* deck has no registry properties, so `Sync Now` refuses it today with
"this deck has no paired workbook". **Measured — that is current behaviour.**

But the obvious human workflow is not that. It is *Save As* from an existing child deck, then
add the other slides — and a copied deck inherits `DeckSyncId`, the workbook path and every
type registration. **Measured, and already recorded from the first live cycle: three decks in
this project share one `DeckSyncId` because they are copies of one original.** So the likely
workflow produces a fully syncable composite, and the existing protection does not fire.

### Proposal — this side's design to bring, per the protocol

Two layers, because the explicit one can be forgotten and the structural one cannot.

**1. An explicit mark.** A toolbar action — *Mark as Build Output* — writing a deck property.
`Sync Now` and `Preview Sync` refuse a marked deck outright and say why. Cheap, unambiguous,
and it uses machinery that already exists and is tested.

**2. A structural backstop, and this one needs no human to remember anything.** Under D3 —
one deck per slide type — **a child deck contains exactly one slide type.** So a deck holding
slides of two or more registered types is, by construction, not a child deck. That is not a
heuristic or a guess; it follows from the deck-splitting decision itself. Combined with the
duplicate-`EntityCode` condition R9 already detects, it identifies a hand-assembled composite
without anyone having marked it.

The backstop warns rather than refuses, and defaults to No. An unmarked composite is a
mistake, not corruption, and refusing outright would block the one legitimate case nobody has
thought of yet.

**What this needs from your side: nothing.** It is deck behaviour and it is this side's to
design and build. Raised here because it is a consequence of the ruling that neither side had
on the table an hour ago, and because it changes what "the tool is safe to run" means.

**Cost: small.** Both layers reuse existing machinery — deck custom properties, the R9 check
built this morning, and the confirmation pattern already used by `Sync Now`.

---

## 4. Where does period live in a composite? — **largely moot under §0**

Kept for the record. Under the human-compiler ruling nothing reads a composite's period, so
option 2 below — it declares nothing — follows directly and this stops being an open
question.

D4 puts period on the deck. A composite is a deck. Assembled from a Q3 source and a Q4
source, what does it declare?

This side has no answer and had not considered it. Options, none preferred:

1. The composite declares the period it was **built for**, and sources of other periods are
   an error the compiler reports.
2. The composite declares nothing, because it is never synced and period is a sync concept.
3. Period becomes a property of the **slide** in a composite, not the deck — which
   contradicts D4 and is the reason it is listed last.

Option 2 is the cheapest and follows from §2. It is not a recommendation; it has not been
thought about long enough to be one.

---

## 5. Does the compiler read the register? — **this is the omission, and the answer is yes,
via a Build sheet you have not been told about**

The design has **two sheets**, not one:

| Sheet | Holds | Nature |
|---|---|---|
| **Data** (≈ your register) | one row per entity per period — facts only | what values |
| **Build** | **one row per output slide, in output order** | which slides, from where, in what order |

The Build sheet is the composition manifest. Hand-editing it is how you get **audience cuts**
— board / funder / internal — from identical Data. It can be generated from Data for the
common case so a normal quarterly run stays one click.

**Why this matters to your side specifically:** it means the workbook has a second table with
a different shape and a different consumer, and the register is not the only thing the
compiler reads. Your schema has one client today and would have two.

This design was not invented independently. It follows **think-cell's `.ppttc`** format — an
IANA-registered media type with a published schema — which solves the same problem with a
top-level ordered array where each element is one copy of a template bound to data by name.
The lesson taken from it, and it corrected this side's original instinct: **composition lives
in a separate ordered manifest; the data rows stay pure facts.** Order and repetition are
properties of the list, not flags on a record.

That is why the Build sheet exists rather than extra columns on the register — and it is
consistent with your round 5 correction removing `SlideID` from the register. Both are the
same principle: the register describes *facts about entities*, not *where things go*.

Status: **designed, not built.** Progression step 3.

---

## 6. One deck per slide type — one workbook per deck, or one register? — **RM3**

This side's position, offered as input rather than a ruling since round 7 §3 makes the record
your side's call: **one workbook, many decks.**

Splitting the workbook per deck would defeat the stated purpose of the record — trend,
comparison, and the prior-quarter exemplar the content approach depends on — for a reason
that is purely presentational. The deck split exists to make each deck simple and singular;
that has no bearing on where the facts live.

The consequence for the register is small and is already handled: `SlideType` is in the
revised E4 column list, so a single register can serve many decks with each deck reading only
its own type's rows. `Register.ReadRegister` already filters on it. **Measured — that filter
is tested and passing.**

The mechanical consequence is that each deck's stored workbook path points at the same
workbook, which is already supported today (`DeckSyncWorkbookPath` is per-deck and nothing
requires them to differ).

---

## 7. VBA versus Office JavaScript — **measured, and still blocking**

Unchanged since round 3, and the facts are these:

- **VBA cannot control formatting when moving slides between presentations.**
  `Slides.InsertFromFile` has no keep-source-formatting option; the UI offers it as a smart
  tag and smart tags are not in the object model at all.
- The available workaround — assigning the source `.Design` to the pasted slide — has a
  documented and real cost: **duplicated slide masters and file bloat**, which is actively
  bad when every source deck shares one corporate template.
- **The Office JavaScript API has exactly the missing capability**:
  `insertSlidesFromBase64` with `formatting: "KeepSourceFormatting"` and an explicit
  insertion point.
- Caveat found: inserted slides keep the source's relative order regardless of array order,
  so output ordering needs one call per slide or a reposition pass.

**Does a move to Office JS change the Excel interface?** Assessment, **speculative — not
measured**: the register schema should be unaffected, because it is a table of facts read
over COM or over the JS API alike. What would change is the *add-in packaging* and possibly
how deck-level properties are read. Tags are OOXML-level and survive either way.

**This blocks composition only.** Everything in the text pipeline — steps 1 to 3, the whole
of what both lanes are currently building — is unaffected.

---

## 8. Where it sits in sequence — **designed**

```
1. Master template slide, never a real project     ✅ BUILT, live-cycled 2026-07-30
2. period column + deck declares its period        ⬜ designed  (= your Quarter)
3. Existence + Position (Build sheet)              ⬜ designed  (still one deck, still VBA)
4. Provenance / cross-deck composition             ⬜ NOT designed — the fork above
5. Compiler                                        ⬜ assembles from 1–4
```

Steps 1–3 extend what already runs. **Step 4 does not** — that is the honest correction to
any impression that the compiler is a continuation of current work. It is a new engine with
an unresolved technology question in front of it.

The text pipeline both lanes are building **is** steps 1–3. Nothing done so far is wasted
under any answer to RM1–RM5, with one exception noted below.

---

## 9. What might have to change in what is already settled

Answering round 7's actual concern — does this invalidate anything?

**Probably safe:** the register's key, the three-class taxonomy, `Quarter = ALL`, `Status`,
the FieldID join, dropping `SlideID`, and the long-format decision. All of these describe
facts about entities and periods, and the compiler consumes a *different* sheet.

**Possibly affected:** whether the register needs a column distinguishing "this row feeds a
child deck" from "this row is only ever read by the compiler" — unknown until RM1 says what a
composite is for.

**The one genuine interaction:** if audience cuts (RM1) turn out to be a core requirement,
the Build sheet arrives sooner and the workbook has two clients earlier than assumed. That
changes sequencing, not schema.

---

## 10. Scope flag

**Scope flag:** round 7 §4 is accepted, and the same correction runs the other way — this
side's round 6 §1 asked your side for a ruling on `3_P002-2` as though it were a portfolio
convention. The *naming* convention was fair to ask. Whether that slide survives is a deck
decision and was mine to bring a proposal on, not yours to rule on. Both sides drifted on the
same item from opposite directions, which is probably the clearest evidence the protocol was
needed.

Restating the requirement as I read it: **`EntityCode` must resolve to exactly one slide per
deck at run time.** How the deck guarantees that is this side's design, and I will bring a
proposal rather than a question once RM4 is answered.
