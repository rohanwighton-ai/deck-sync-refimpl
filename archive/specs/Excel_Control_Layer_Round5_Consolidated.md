# Excel Control Layer — round 5 (consolidated)

**In reply to:** `Excel_Control_Layer_Round4.md`, 31 July 2026
**Contains:** E2 as a deliverable, a decision on each amendment, sequencing confirmed, and one
correction on the Excel side. Single file — supersedes any earlier round 5 and its addendum.
**No new design questions.**

---

## 1. E2 — the mapping table

Convention: `UPPER_SNAKE`, matching `PROGRESS_BODY`. `Project Status` first, as requested.

| Order | Current `role` value | Target `FieldID` | Class |
|---|---|---|---|
| 1 | `Project Status` | `PROJECT_STATUS` | quarterly |
| 2 | `Project Name` | `PROJECT_NAME` | entity-static |
| 3 | `Project number` | `PROJECT_CODE` | entity-static |
| 4 | `About text` | `ABOUT_BODY` | entity-static |
| 5 | `events text` | `KEY_EVENTS_BODY` | quarterly |

Inverse map for rollback is this table read right to left. Nothing else is in scope for V1.

Two notes on the choices, neither of which changes the table:

- `Project number` becomes `PROJECT_CODE` rather than `PROJECT_NUMBER` because the portfolio
  runs on P-codes, S-codes and K-codes. "Number" is already wrong for two of the three.
- The `Class` column is not decoration — see §3. It is what makes Amendment B work.

---

## 2. Amendment A — accepted as recommended

**The placeholder is the draft marker.** No watermark, no filename suffix, no hidden-slide
use. Your argument holds: a second marker invented before the first has been tested produces
two that nobody reads.

Backed by:

- a deck custom property carrying the last run's placeholder count;
- the run report's headline being the placeholder list itself, not a summary line.

One addition, and it is one string not one mechanism: **store the period and run timestamp in
that same property alongside the count.** A bare count cannot be told apart from a stale count
left by a previous run, and the whole value of the property is that it is machine-readable.

D6 is amended accordingly: a surviving `<<…>>` blocks publication by being visible, and the
deck property records that it happened. The word "draft" comes out of D6 — there is no draft
state and we are not building one.

---

## 3. Amendment B — mechanism accepted, taxonomy refined

**Accepted:** the Field Spec sheet's class column determines whether a tag is applied at all,
not merely whether a row is expected. A tagged field with no row keeping its placeholder
forever is a real trap and you are right to have caught it before E3 was authored.

**Refined:** the binary is wrong, and it would have caused a different failure. Amendment B
collapses "does not change quarter to quarter" into "is template chrome". Those are not the
same thing. `PROJECT_NAME` does not change quarter to quarter, but it is not chrome either —
leaving it untagged means every created slide carries the master template's project name until
someone retypes it by hand, which is precisely the failure the tool exists to remove.

**Three classes, not two:**

| Class | Tagged? | Register rows | Example |
|---|---|---|---|
| **Chrome** | No | none | `ABOUT`, `months`, `Time elapsed` |
| **Entity-static** | Yes | exactly one, `Quarter = ALL` | `PROJECT_NAME`, `PROJECT_CODE` |
| **Quarterly** | Yes | one per quarter | `PROJECT_STATUS`, `KEY_EVENTS_BODY` |

`ALL` is a sentinel in the `Quarter` column, not a null. R2's filter becomes
`Status = Approved AND (Quarter = <deck period> OR Quarter = 'ALL')`. That is the only change
this taxonomy imposes on V4, and it is one clause.

Amendment B's danger case is then closed for the right reason: every tagged field has a row,
so no tagged field can hold a placeholder indefinitely. Untagged means unmanaged, and
unmanaged means the audit will list it forever — which, as you say, is correct behaviour and
not something the audit should learn to suppress.

`ABOUT_BODY` is classed entity-static on the evidence that it describes the project rather
than the quarter. If it turns out to be edited between quarters it moves to quarterly, and
that is a one-cell change in the Field Spec sheet with no code consequence. Classes are
intended to be cheap to revise; that is part of why they live in Excel.

---

## 4. Correction — `SlideID` comes out of the register

Round 4 §3 said not to bake `SlideID` into anything durable, and Q6b measured why: it is
reassigned on duplicate within a deck and preserved on paste across decks. Our round-3
confirmation put it in the register key regardless. **That was our error and it is withdrawn.**

`SlideID` is also redundant, which is the better reason to remove it. The join resolves
entirely from the deck's own tags — the slide carries `instance_key` (EntityCode), the shape
carries `role` (FieldID), and the deck carries the period. The register never needs to name a
slide instance, and naming it duplicates identity the deck already holds using the one piece
of it you have measured as unreliable.

**Revised register key:** `Quarter × EntityCode × FieldID`.

**Revised E4 column list:**

```
Quarter, EntityCode, SlideType, FieldID, FieldType, Value, CharCount, Status, UpdatedDate
```

`SlideType` replaces `SlideID` and is not part of the key. It names a *kind* of slide, not an
instance, so it survives duplication, cross-deck paste and the one-deck-per-slide-type
direction in §7 of round 3. It exists for one purpose: the supported case in D12 of a register
row for an entity with no slide yet, where creation needs to know which template to clone.
It maps to the existing `slide_type` slide tag and to `DeckSyncType:<type>`.

If `SlideType` turns out to be better held per entity than repeated per field row, say so and
it moves to the Field Spec sheet — it is an attribute, not a key, so moving it is free.

---

## 5. Further confirmations from round 4 §3 and §4

**§3 — period lives in a deck custom property, not a slide tag. Confirmed.** Your reasoning
stands on its own: one read against N things that can disagree with each other, and it makes
D5's plausibility check a single property read at run start. Adopted.

**§3 — `period_key` is specified and never implemented. Note taken, and it should be closed
rather than recorded.** A slide tag documented in `specs/identity-tags.md` that no code reads
or writes will be assumed live by whoever reads the spec next — you have just demonstrated
that by having to check. Given the decision above it should be deleted from the spec, not
implemented. Your call on timing; it belongs with V2 since that is the task that touches tag
documentation anyway.

**§4 — tag enumeration returns names uppercased, matching must be case-insensitive.
Acknowledged.** It does not touch E2, since the mapping table operates on the `role` tag's
*value* rather than its name, but it is recorded here so it is not rediscovered later.

**§3 — `DeckSyncType:<type>` holds `templateSlideID|worksheetName`.** Noted only because Q12
scanned slide text runs, shared strings and audit rows and reported zero pipes, while a pipe
demonstrably exists in a deck custom property. R6 governs injected field values, so there is
no conflict and nothing to change. Flagged so the scan's scope is remembered accurately if it
is re-run before a bulk import, as §1 of round 4 recommends.

---

## 6. Sequencing — confirmed, rename first

**Confirmed: rename before rewrite, with V5 and V6 as fill-ins.** The argument is right and it
is right for the reason you give rather than the reason it first appears: the cost scales with
the tagged-field count, and every other task raises that count. 30 writes today against ~258
after onboarding is not a close call, and the probes have removed the grounds for caution —
idempotent overwrite, enumerable, reversible, and a partial migration degrades to
non-injection rather than corruption.

**E2 is delivered above, ahead of dump order.** V1 is unblocked as of this document.

**First field confirmed: `Project Status`.** Your reasoning is better than ours was. Proving a
pipeline on something answerable at a glance, then running the hard field through a proven
path, is the correct order and it is the same principle as §5 itself.

### One ordering query, internal to your §6

§6 orders **V1 → V2 → V3 → V4**, then says V2 "guards every subsequent operation that touches
tags in bulk — **including V1**."

Both cannot hold. V2 depends on nothing, so if it guards V1 there is no cost to running it
first, and a duplicate-EntityCode deck is exactly the condition under which a bulk tag rewrite
does the most damage.

**Proposed: V2 → V1 → V3 → V4.** Your call — you may have meant V2 guards everything after V1
specifically, in which case the order stands as written and this is nothing.

### One correction to the workplan

**V5 is not a fill-in.** Appendix A shows `_x000D_` in items 3, 28, 33, 34, 35 and 37 —
existing content already carries embedded carriage returns. The first multi-line field
onboarded needs `||` handling on the day it is onboarded, not later. `PROJECT_STATUS` is
single-line so V1 and the §5 meet point are unaffected, but V5 should land before `ABOUT_BODY`
or `KEY_EVENTS_BODY` are taken through.

---

## 7. Excel lane status

E1 triage of the 77 is under way and is ours; nothing on your lane waits on it. E3, E4 and E5
follow from it and will arrive together, built for `PROJECT_STATUS` first.

Two observations from Appendix A, recorded because they are evidence rather than requests:

- `Shape 16` appears four times with four different values (`~$280K`, `~$280K`, `~$1.4M`,
  `~$1.9M`), matching the four financial chrome labels. That is R1's failure mode sitting in
  plain sight in the real deck, and it is the clearest vindication available of the decision
  to drop `ShapeName`.
- The banding caveat in §C is understood and accepted. We will treat `Guess` as a sort order
  and make the field-or-chrome call ourselves, which is where it belongs.

Nothing further is needed from your side to start V1.
