# Deck Sync — current workplan

**Updated:** 31 July 2026, after round 5 of the Excel Control Layer exchange.
Supersedes the workplan in `specs/excel-control-layer-round4.md` §6, which is now history.

The design conversation is **closed** — round 5 delivered E2, decided both amendments, and
raised no new questions. What follows is build work against settled decisions.

---

## Status

| | Task | State |
|---|---|---|
| **V2** | R9 duplicate identity-tag check | ✅ **done**, 115/115, pushed |
| **V1** | FieldID rename migration | ✅ built, 120/120 — **not run on the real deck**, blocked on the EntityCode ruling (round 6 §1) |
| **V3** | Long-format register reader | ✅ **done**, 120/120 — `Register.bas` |
| **V4** | Wire the register reader into the sync + period property | ⬜ **next** — V3 built the reader, V4 connects it |
| **V5** | `\|\|` → line break | ⬜ **promoted — see below** |
| **V6** | Empty `Value` is a validation failure | ⬜ after V3 |
| **V7** | Placeholder-count deck property | ⬜ new, from Amendment A |
| **E1** | Triage the 77 items | 🔨 Excel lane, in progress |
| **E2** | Mapping table | ✅ **delivered** |
| **E3/E4/E5** | Field Spec, register, template worksheet | ⬜ Excel lane, `PROJECT_STATUS` first |

---

## What changed in round 5, and what each change costs

**E2 delivered — V1 unblocked.** Five rows, `PROJECT_STATUS` first:

| Current `role` | Target `FieldID` | Class |
|---|---|---|
| `Project Status` | `PROJECT_STATUS` | quarterly |
| `Project Name` | `PROJECT_NAME` | entity-static |
| `Project number` | `PROJECT_CODE` | entity-static |
| `About text` | `ABOUT_BODY` | entity-static |
| `events text` | `KEY_EVENTS_BODY` | quarterly |

**Amendment B was wrong and has been corrected — by them.** My version said static fields go
untagged. That collapses "does not change quarter to quarter" into "is template chrome", and
those are not the same thing: `PROJECT_NAME` does not change quarterly, but leaving it
untagged means every created slide carries the **master template's** project name until a
human retypes it — the exact failure this tool exists to remove. I would have written that
into the Field Spec sheet's semantics.

Three classes, not two:

| Class | Tagged? | Register rows |
|---|---|---|
| Chrome | no | none |
| Entity-static | yes | exactly one, `Quarter = ALL` |
| Quarterly | yes | one per quarter |

**Cost: one clause in V4.** The filter becomes
`Status = Approved AND (Quarter = <deck period> OR Quarter = 'ALL')`.

**`SlideID` withdrawn from the register — their correction, on my measurement.** Q6b showed
it is reassigned on within-deck duplicate and preserved on cross-deck paste, so it is
unreliable as an identifier; it is also redundant, since the join resolves entirely from the
deck's own tags. Revised key: **`Quarter × EntityCode × FieldID`**. `SlideType` replaces it as
a non-key attribute, for the D12 case of a register row for an entity with no slide yet.

**Cost: V3's target schema is now settled** — `Quarter, EntityCode, SlideType, FieldID,
FieldType, Value, CharCount, Status, UpdatedDate`. This is the reason V3 is safe to start.

**V5 is promoted out of "fill-in", and they found it in my own data.** Appendix A of round 4
shows `_x000D_` in items 3, 28, 33, 34, 35 and 37 — existing content already carries embedded
carriage returns. I generated that appendix and did not read it closely enough.

**Cost: ordering.** `PROJECT_STATUS` is single-line so the first field through is unaffected,
but **V5 must land before `ABOUT_BODY` or `KEY_EVENTS_BODY`** are taken through.

**Amendment A accepted, with one addition.** The placeholder is the draft marker; no
watermark, no filename suffix. Their addition is right: **store the period and run timestamp
alongside the placeholder count**, because a bare count cannot be told apart from a stale one
left by an earlier run. The word "draft" comes out of D6 — there is no draft state and none is
being built. That is V7.

**Their ordering query is resolved by fact.** Round 4 §6 said `V1 → V2` while claiming V2
guards V1 — a genuine contradiction on my part. Moot in practice: V2 shipped first.

---

## Order, and why

```
V2 ✅ ──> V1 ──> V3 ──> V4 ──> §5 meet point: PROJECT_STATUS end to end
                  │
                  ├─> V6 (needs V3)
                  └─> V7 (independent)

V5 ─── independent, but MUST precede ABOUT_BODY / KEY_EVENTS_BODY
```

**V1 before V3** because the rename's cost scales with the tagged-field count and every other
task raises it: ~30 tag writes today against ~258 once the inventory is onboarded. It never
gets cheaper. The probes removed the grounds for caution — idempotent overwrite, enumerable,
reversible, and a partial migration degrades to non-injection rather than corruption.

**V3 before V4** because `Quarter` and `Status` are columns in the long format; filtering
cannot precede the shape being filtered.

---

## V1 execution — the actual steps, in order

Not "run the migration". This touches every managed shape in a real deck.

1. Suite green. *(in progress)*
2. **Dry run against a COPY** of the cycle deck. Confirm: 5 mapped values found, count matches
   expectation, template included, zero unmapped surprises.
3. **Live run on the copy.** Verify with `Preview Sync` — if the join still resolves, every
   field still matches its row.
4. **Rollback on the copy**, to prove reversibility on real material rather than on the
   synthetic fixture the test uses.
5. Only then: backup the rig, run it live, verify, save.

Step 4 is the one worth not skipping. Reversibility is proven by a unit test against a
constructed slide; it is not yet proven against a deck with grouped and nested fields.

---

## Known gaps, carried deliberately

**Period plausibility is undefended.** R9 asks for it; it cannot be built because there is no
period concept until V4. This is the half of D5 with **no native discriminator behind it** —
`SlideID` is preserved on cross-deck paste, so a deck copied to start next quarter has nothing
but the check standing between it and reporting itself as last quarter. Closes with V4.

**The audit's chrome-vs-data heuristic is weak on cloned decks.** It fired its strongest
verdict zero times out of 77. Rule for reading it: 5/5 is chrome, below 5/5 is a candidate;
treat `Guess` as a sort order. The Excel side has accepted this and will make the call
themselves, which is where it belongs.

**One deck, one slide type, 5 of ~43 fields.** The honest measure of progress is 5 of ~43,
not 115 of 115.
