# Excel Control Layer — response from the PowerPoint / VBA side

**In reply to:** `Excel_Control_Layer_Specification.docx`, 31 July 2026
**Written:** 31 July 2026, against the build at commit `2a8d86e` (113 tests passing)

The spec asks for the current position, whether settled or open, and for anything the Excel
side would be wrong to assume. Taken literally below, including where the answer is "that
requirement cannot be met as written."

---

## 0. The framing question first

> *"Really it's about data type match more than anything I suppose?"*

**No — and this is the most useful thing in this document.** Data type match is the smallest
of the problems here and is close to solved already. `FieldType` selecting an injection
routine is a `Select Case`; only `Text` is in scope for the first pass anyway (§5), and the
text path is built, tested and running.

Three things will actually decide whether this works, in order:

1. **The join key.** R1 nominates `ShapeName`. That is provably unusable on the real deck —
   evidence in §1. Getting this wrong does not fail loudly; it binds a value to the wrong
   shape and looks fine.
2. **Identity and period keying.** Both sides have independently arrived at the same
   three-way decomposition. That is a strong signal and it is the good news in this document
   (§2).
3. **Approval and provenance.** `Status` is the genuinely new and valuable thing the spec
   brings, and it settles an open question on our side rather than raising one (§3).

Type matching sits a long way below all three.

---

## 1. R1 cannot be met as written — `ShapeName` is not a key

> *"Match on ShapeName, and only on ShapeName. Not index, not position, not alt-text."*

The instinct behind R1 is right: **do not match on index or position.** The built system
agrees emphatically and for the same reasons. But `ShapeName` is not the stable identifier
the spec takes it for.

**Measured on slide 1 of the real cycle deck, 31 July 2026:**

| | |
|---|---|
| shapes on the slide | **166** |
| names used by more than one shape | **24** |
| shapes carrying a non-unique name | **71** (43%) |
| worst case | `Shape 46` — **four** shapes |
| shapes with an empty name | at least 1 |

`ShapeName` is not unique on the actual target deck, so it cannot be a join key there today.

It is also not *stable*, which matters more than the duplication:

- PowerPoint assigns names automatically (`Shape 46`, `Text 31`) and reuses them freely.
- The name is user-editable from the Selection Pane, with no warning and no uniqueness check.
- **Copying a slide clones every name on it.** This project already has the equivalent finding
  recorded for slide-level identity: three decks in this repo share one `DeckSyncId` because
  they are copies of one original.

R1 asks that "every shape that will ever be populated is to be given an explicit, stable name
before the deck grows further." On this deck that is a manual naming pass over 166 shapes per
slide, repeated per slide, and re-checked every time anyone adds a shape — with a silent
mis-bind as the failure mode.

### What to use instead — and the spec already contains it

This is a one-column change, not a redesign. **`FieldID` is the correct join key.** The spec
defines it as *"stable identifier, independent of where the shape currently sits"* — which is
exactly right, and exactly what the built system already stores.

Identity is held in PowerPoint's own `Shape.Tags` / `Slide.Tags` mechanism, written into the
file at the OOXML level (`ppt/tags/tagN.xml`, referenced from the shape's `<p:nvPr>`). Tags
are invisible to the user, not editable from the UI, survive edit/move/regroup/reorder, and
are already how every field in the running system is found. A shape's tag `role` **is** a
`FieldID`.

**Recommendation:** join on `FieldID` ↔ shape `role` tag. Keep `ShapeName` in the register if
it is useful to a human reading the sheet, but mark it explicitly non-authoritative, or drop
it. The naming convention in R1 (`txt_`, `pic_`, …) is then optional cosmetics rather than a
prerequisite — which removes the largest piece of manual setup work in the whole spec.

**Tooling for this already exists and is on the toolbar:** *Mark Field for Batch* and *Bulk
Onboard Type* tag fields at scale and confirm the whole batch through a review grid written
into the workbook. That is the mechanism that would otherwise have to be replaced by hand-
naming 166 shapes a slide.

---

## 2. Where both sides independently agree — keep this

The spec's key is `Quarter × EntityCode × FieldID`, with `SlideID` naming the slide. A
decision logged on our side on 2026-07-30, before this spec was seen, splits identity into
`period` / `project` / `instance_key`. Those are the same decomposition:

| Spec | Our decision | Meaning |
|---|---|---|
| `Quarter` | `period` | which snapshot this row describes |
| `EntityCode` | `project` | the thing it is about, forever |
| `SlideID` | `instance_key` | which slide, forever |
| `FieldID` | `role` (shape tag) | which field on it |

Two independent routes to the same four-part key is the strongest evidence available that
it is the right one. **Adopt the spec's names** — `Quarter` / `EntityCode` / `FieldID` are
clearer than ours, and ours are internal.

**Long format is also right, and we should change to match.** The current sheet is wide: one
row per instance, one column per field. The spec's *"adding a field adds rows, never
columns"* is the correct call, and the audit run on 31 July makes it concrete — the real
slide carries **77 text items, of which ~38 are plausible fields**. A 38-column sheet is
unmanageable and every new field is a schema change.

**Cost, stated plainly:** this is a rewrite of `ExcelOutput.ReadSheet` and `UpsertRow`, which
are built for the wide shape and carry read-merge-write semantics designed around it. It is
tractable and it is the right direction, but it is not free and it should be scheduled, not
assumed.

---

## 3. Where the spec is right and our own decision should yield

**R3, "never write back to Excel", conflicts with a decision we made on 2026-07-30** that
child decks may sync two-way. The spec is right and that decision should be narrowed.

A control layer with an approval state (`Draft` / `Reviewed` / `Approved`) cannot have slides
writing back into it. A slide edit flowing into the register would either land as `Approved`
without review, or silently downgrade an approved row — and both destroy the meaning of the
column. `Status` does something our design could not: it removes the need for the
"who changed this" machinery entirely, because only one path can write, and that path has a
human decision in it.

That is a real improvement and it closes an open problem rather than adding one.

**One precision needed.** R3 as written is already violated by the shipped tool, in ways that
are fine:

- onboarding harvests slide values into the workbook — but that is *setup*, not a run
- batch onboarding writes a review grid into the workbook
- the field audit writes a `Template Audit` sheet

None of these touch the register. **Suggested wording:** *never write to the field register;
setup and diagnostic operations may write elsewhere in the workbook.* As written the rule is
too broad to be true, and a rule that is already being broken stops being read.

---

## 4. R2 is not a requirement on the VBA layer — it is unbuilt work, already designed

> *"Read only rows where Status = Approved and Quarter = the nominated run quarter."*

There is **no `Quarter` concept and no `Status` concept in the build today.** Values are
overwritten in place; last quarter's figures are gone the moment they are typed over. R2 is
therefore a request to build the next planned step, not a constraint on existing behaviour.

The design already exists on our side and agrees with the spec, so this is aligned work
rather than new argument.

**One difference, and ours is better:** the spec says *"the nominated run quarter"*, implying
the operator names it per run. Our design has **the deck declare its own period**, stored in
the deck. One less thing to get wrong on each run, and it makes a past deck reproducible —
open last quarter's deck, and it still knows which quarter it is. Recommend adopting.

---

## 5. R4 has an interaction the spec cannot have known about

> *"Do not blank a shape that has no matching register row. Leave it untouched."*

Correct as a rule, and it matches current behaviour for existing slides — a field with no row
is skipped, never guessed at.

**But it breaks on newly created slides, as of a change made 30 July.** New slides are now
cloned from a master template whose fields carry visible placeholder text (`<<PROGRESS_BODY>>`
and so on) precisely so that unfilled fields are unmistakable. Under R4, a new entity with no
register row for a field leaves that placeholder **on the slide, in the deck, reading
`<<PROGRESS_BODY>>`**.

"Leave it untouched" and "the untouched state is scaffolding" only collide on created slides.

**Options, and this is a genuine decision for the Excel side:**

1. Clear placeholders on creation — a new slide shows an empty field rather than scaffolding.
2. Keep placeholders and report them, treating a visible `<<…>>` as a loud, deliberate
   "nothing was said about this" that gets caught before publication.
3. Require a register row for every field of a created entity, making the case impossible.

Our lean is (2) — it fails loudly, which the spec asks for in R5, and an empty field on a
report slide is easier to miss than a bracketed marker. But this is the Excel side's call
because it determines what the operator has to supply.

---

## 6. Direct answers to §4

**Character or content limits.** None are enforced, and none can be. Injection writes into an
existing shape's text range and reads the value back to confirm it landed; there is no length
cap and no error on overflow. Text longer than the shape simply overflows or autofits
smaller — a rendering outcome, not a failure the code can detect. **So `Target length` is
advisory and must be policed on the Excel side.** `CharCount` policing drift is the right
mechanism; it is the only one there will be.

**R1–R8 the current build cannot do, or does differently.**

| | Position |
|---|---|
| R1 ShapeName | **Cannot, as written.** See §1. Counter-proposal: join on `FieldID` ↔ role tag. |
| R2 Status + Quarter filter | **Not built.** Designed, agreed, next in sequence. See §4. |
| R3 never write back | **Agreed**, with the register/workbook distinction in §3. |
| R4 do not blank | **Agreed for existing slides**; needs a decision for created ones. See §5. |
| R5 fail loudly, named list, run continues | **Already the behaviour.** Unmatched rows and unclassified slides are collected and reported at the end; the run does not stop. |
| R6 `\|\|` as line break | **Not built, no objection.** Trivial to add at the injection point. |
| R7 populate existing shapes only | **Agreed for shapes** — but note the build already *creates slides*, tested and run live. A register row for an entity with no slide is a supported case. |
| R8 reconcile field inventory | **Agreed**, and the tool for it exists — *Audit Fields*, built 31 July, lists every text item on a slide that is not a managed field. |

**The field inventory as it exists today.** One registered type, `q`, with **five** fields:
`Project Name`, `Project number`, `Project Status`, `About text`, `events text`.

The audit against the real deck on 31 July found **77 separate text items** on that slide,
~38 of them plausible project data. **So the tool currently manages 5 of ~43 candidate
fields — roughly 12%.** A visual estimate the day before had put the total near 25; it was
out by a factor of three, in the optimistic direction. Anyone sizing this work should use 77,
and should count rather than estimate.

Note also that current FieldIDs are human-readable with spaces (`About text`), not the spec's
`PROGRESS_BODY` convention. R8's instruction that the register's FieldIDs are authoritative
and code should be brought across is accepted — but it is a rename of live tags in a deck,
which is a migration with a real failure mode, not an edit.

**How the VBA prefers to be handed the table.** Currently: a worksheet with fixed columns,
row 1 headers, column A reserved for the instance key, loaded **all at once** into a
dictionary rather than read row by row. Whole-set loading is deliberate and should be kept —
the sync needs the full picture before it writes anything, because it must be able to say
what it is *going to* do and get confirmation first.

**Preference, since the spec says this one is genuinely open: a ListObject (named table).**
It gives stable column identity regardless of position, survives insertion, and is
addressable by name from VBA without hard-coded column letters. Fixed worksheet columns are
what exists and it works, but it is the arrangement most easily broken by someone inserting a
column, and Copilot inserting rows and columns is explicitly expected (§2.4).

---

## 7. What the Excel side would be wrong to assume

**That reading the register is a read.** The sync attaches to a *running* Excel instance and
sees whatever is in memory, saved or not. A publish step that writes the register and does not
save leaves the file on disk disagreeing with what the tool injects — a slide can be built
from a row that exists in no file. This was found live on 30 July, twice. **The publish step
must save.**

**That one deck is the target.** The direction of travel is one deck per slide type plus a
compiler that assembles composites. `SlideID` as a flat identifier works now and will need
revisiting. Nothing in the spec is wrong because of this; it is a thing to know before
`SlideID` gets baked into anything downstream.

**That cross-deck assembly is available.** It is not, and it is blocked on an unresolved
technology choice: VBA provably cannot control keep-source-formatting when moving slides
between presentations, whereas the Office JavaScript API can. This blocks composition only,
and nothing before it.

**That passing tests mean it works.** Stated because it bears on sequencing, not as a
disclaimer. The 113-test suite has twice failed to catch defects that fifteen minutes of real
use found immediately — six on 30 July, three on 31 July, none of them logic errors. They
lived in wiring, wording, missing guards, inherited state and file state. §5's "one field
taken completely through before anything is generalised" is exactly right, and the reason it
is right is that generalising from a tested-but-unused path is how all nine of those got in.

---

## 8. Recommended changes to the spec

1. **§2.1 / R1 — replace `ShapeName` as the join key with `FieldID`.** Single highest-value
   change. Removes the manual naming of 166 shapes per slide and eliminates a silent
   mis-bind failure mode. `ShapeName` becomes a non-authoritative convenience column or goes.
2. **R3 — narrow to "never write to the field register."** Currently too broad to be true.
3. **§4 / R2 — record that `Quarter` and `Status` are unbuilt on the VBA side**, designed and
   agreed, and adopt deck-declares-its-own-period rather than an operator-nominated quarter.
4. **R4 — add the created-slide case** and pick one of the three options in §5.
5. **§2.1 — state that `Target length` is advisory**, unenforceable at injection, and that
   `CharCount` is the only policing mechanism there will be.
6. **§2.4 — require the publish step to save the workbook.**
7. **§5 — keep exactly as written.** It is the most valuable paragraph in the document.
