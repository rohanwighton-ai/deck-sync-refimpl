# Excel Control Layer — confirmation and follow-up

**In reply to:** `Excel_Control_Layer_Response.md`, 31 July 2026
**Supersedes:** the relevant clauses of `Excel_Control_Layer_Specification.docx`, same date

The response did the useful thing — it refused a requirement rather than building around it,
and it brought measurements. R1 is withdrawn on the evidence. Rulings below are settled unless
flagged otherwise. Section 3 is what we need back.

---

## 1. Rulings

**D1 — R1 is withdrawn. Join on `FieldID` ↔ shape `role` tag.**
Accepted in full. The counter-proposal is better than the rule it replaces.

**D2 — `ShapeName` is dropped from the register entirely**, rather than retained as a marked
non-authoritative column. A column that looks like a key, sits beside the real key, and is
documented as "do not join on this" will eventually be joined on. The human-readable locator
moves to the Field Spec sheet, which is the human-facing surface in any case. The `txt_` /
`pic_` naming convention is withdrawn with it.

**D3 — R3 narrows to "never write to the field register."**
Accepted, in your wording. Setup, onboarding and diagnostic writes elsewhere in the workbook
are in scope and always were.

**D4 — the deck declares its own period.** Accepted, conditional on D5.

**D5 (new, R9) — identity tags must be checked for uniqueness at the start of every run.**
This follows from your own finding and extends it. Tags clone on copy the same way names do —
that is what the three decks sharing one `DeckSyncId` already demonstrate. So:

- duplicating a project slide clones its `EntityCode`, producing two slides both claiming P004;
- copying last quarter's deck to start this quarter clones the declared period, so a Q4 deck
  reports itself as Q3.

`FieldID` cloning is correct and wanted — every project slide should carry `PROGRESS_BODY`.
It is the identity tags that cannot be allowed to duplicate. Requirement: before any run
writes, check `EntityCode` and `SlideID` tags for duplicates and the declared period for
plausibility, and report both. Period is set by explicit action on deck creation, never
inherited, and is displayed for confirmation at run start.

**D6 — R4, created slides: option 2, with a gate.**
Placeholders stay visible and are reported. Added condition: **a surviving `<<…>>` blocks
publication.** If any placeholder remains after a run, that is the headline of the run report
and the output is marked draft by whatever mechanism you have. Reporting alone degrades into
option 1 the first time someone skims the log.

Option 3 was rejected because it forces a register row for fields that genuinely have nothing
to say in a given quarter, which is how reporting prose gets padded.

**D7 (new, R10) — an empty `Value` on an otherwise complete row is a validation failure at
publish, and is never injected.** "No row" means leave the shape alone (R4). "Row with empty
value" means someone left a cell blank. These must not resolve to the same behaviour.

**D8 — `Target length` is advisory, `CharCount` is the only policing mechanism.**
Accepted. This makes the prior-quarter exemplar column in the template worksheets
load-bearing rather than merely convenient, which is noted on our side.

**D9 (new, R11) — the publish step saves the workbook.**
Accepted from §7. Not a spec change so much as a defect you found for us.

**D10 — ListObject, whole-set load.** Accepted, both.

**D11 — long format.** Confirmed as the direction, and the `ExcelOutput` rewrite is
acknowledged as scheduled work rather than assumed. See Q9 on ordering.

**D12 — R5, R7, R8 stand as you have them**, including that slide creation is in scope and
a register row for an entity with no slide is supported.

---

## 2. What this changes on the Excel side

The register loses `ShapeName` and gains nothing — the key becomes
`Quarter × EntityCode × SlideID × FieldID`. The Field Spec sheet absorbs the human-readable
locator and the per-field target length. Neither is blocked on anything from your side except
the field inventory in Q5.

---

## 3. Questions

**Q5 — the audit output, as raw data.** The `Audit Fields` list for slide 1: all 77 text
items, with current tag state where one exists, and which of the ~38 you judge to be
project data. This is the input the Field Spec sheet is built from, and it is the single
thing most blocking work on this side. A dump is fine; it does not need interpreting.

**Q6 — confirm the tag cloning behaviour directly.** D5 is reasoned from the `DeckSyncId`
finding rather than measured. Do `Shape.Tags` and `Slide.Tags` survive slide copy and deck
copy identically to names? And is there anything PowerPoint-native that would let a copied
slide be distinguished from its original — a creation stamp, a modification marker, anything —
or does the uniqueness check have to be the whole defence?

**Q7 — where do `EntityCode` and the period actually live?** Slide tag, deck-level custom
property, or both? Asking because it determines the scope of the D5 check, and because the
one-deck-per-slide-type direction in §7 will move whatever the answer is.

**Q8 — the R8 rename migration.** `About text` → `PROGRESS_BODY` is a rename of live tags in
a working deck. What is the plan and what is the rollback? Specifically: is there a mapping
table we should author on this side, is the migration reversible, and what happens to a deck
that is half-migrated when someone opens it.

**Q9 — sequencing, and your recommendation.** Three pieces of unbuilt work are interdependent:
the long-format `ExcelOutput` rewrite (§2), the `Quarter` / `Status` build (§4), and the
FieldID rename migration (§6). Doing the rename before the rewrite means migrating tags twice
if the rewrite changes how fields are addressed; doing it after means the long-format sheet is
authored against names we intend to discard. What order do you want them in, and what is the
dependency chain? We will schedule the Excel side around your answer rather than the reverse.

**Q10 — static fields.** Some fields do not change quarter to quarter. Current intent is that
they carry no register row and are populated once at slide creation from the master template.
Does creation-from-template already do this, and does R4's "leave untouched" then hold for
them on every subsequent run?

**Q11 — the draft-marking mechanism for D6.** What is available: filename suffix, a visible
watermark shape, a flag written to the run report only, something else? Naming what exists so
the gate is specified against a real mechanism rather than an assumed one.

**Q12 — pipes in content.** R6 nominates `||` as the line-break delimiter. Does any current
or historical field value contain a pipe character? If so we will pick a different delimiter
now rather than discover it in a published deck.

---

## 4. Sequencing

§5 stands unchanged and is not negotiable from this side: one text field taken completely
through — spec row, template layout, register row, injection, verified on a slide — before
anything is generalised. Your §7 note on the nine defects that tests did not catch is the
argument for it, better made than we made it.

The number that should drive scheduling is 5 of ~43, not 113 of 113.
