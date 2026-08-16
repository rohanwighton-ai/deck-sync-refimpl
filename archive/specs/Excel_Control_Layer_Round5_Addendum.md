# Excel Control Layer — round 5 addendum

**In reply to:** `Excel_Control_Layer_Round4.md` §3, §4 and §6
**Status:** one correction on the Excel side, three confirmations, one ordering query.
**No new design questions.** Issued because round 5 under-responded to §3 and §6.

---

## 1. Correction — `SlideID` comes out of the register

Round 4 §3 said not to bake `SlideID` into anything durable, and Q6b measured why: it is
reassigned on duplicate within a deck and preserved on paste across decks. Round 5 confirmed
an E4 column list containing it anyway, and the round-3 confirmation put it in the register
key. **That was our error and it is withdrawn.**

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

## 2. Confirmations that round 5 should have carried

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
is re-run before a bulk import, as §1 recommends.

---

## 3. One ordering query, and it is internal to your §6

§6 orders **V1 → V2 → V3 → V4**, then says V2 "guards every subsequent operation that touches
tags in bulk — **including V1**."

Both cannot hold. V2 depends on nothing, so if it guards V1 there is no cost to running it
first, and a duplicate-EntityCode deck is exactly the condition under which a bulk tag
rewrite does the most damage.

**Proposed: V2 → V1 → V3 → V4.** Your call — you may have meant V2 guards everything after
V1 specifically, in which case the order stands as written and this is nothing. Either way E2
is delivered and V1 is unblocked whenever it runs.

---

Nothing else from round 4 remains unanswered on our side. E1 continues here.
