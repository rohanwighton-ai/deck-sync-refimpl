# Drafting sheet speed strategy — FIX-LIST item AD

Companion to `LOBBY-DESIGN.md` (the Lobby architecture) and `FIX-LIST.md` item AD.
Written 2026-08-17 evening, after item AB's live proof (`BuildLobbyFromScratch`,
503.4s -> 2.67s, ~188x) made `Drafting.WriteDraftingSheet` the next largest measured
cost in `RefreshDraftingSheets`. Produced by a fresh research pass (model: fable,
read-only, no code changes) given the real measured numbers and this codebase's own
incident history as context, then reviewed here.

**Gist:** each drafting sheet is slow because the code talks to Excel one cell at a
time across a process boundary, roughly 600 times per sheet; the fix is to read the
whole sheet in one call, do the thinking in memory, write back in a handful of
column-sized calls, and stop repainting the decoration that hasn't changed. Bulk
arrays alone only buy ~2-3x -- the decoration is 40% of the cost and has to be
skipped-when-unchanged to hit 10x.

## What's actually slow (measured against the real source, `vba/Drafting.bas`)

`WriteDraftingSheet` (lines 497-1173) makes roughly 570-600 individual cross-process
COM calls per field, for a 43-row sheet:

| Stage | Ops | Scales with rows? |
|---|---|---|
| Layout detection, period stamp, intro wipe | ~14 | no |
| Instruction block + header row | ~29 | no |
| `rowOf` index scan (2 reads/row) | ~87 | yes |
| Main entity loop (4 writes/row) | ~172 | yes |
| Widths, wrap, interior, valign, freeze | ~35 | no |
| `FieldSpec.LookupGuidance` | ~25-40 | with spec rows |
| `Sources.CitedBlockFor` (2 reads/row + a linear Sources-sheet scan) | ~87+ | yes |
| Prompt writes, row heights | ~17 | no |
| **`ApplyDraftingLook`** (borders/interior/font property sets) | **~100** | no |
| Trailing cosmetics | ~6 | no |

**The constant-overhead portion (~240 of ~600 ops -- instructions, headers, widths,
and above all `ApplyDraftingLook`) is paid in full on every rebuild even when
nothing cosmetic changed.** This is the finding that sets the strategy: bulk arrays
alone remove the row-scaling ~350 ops (roughly 2-3x); hitting 10x requires the
cosmetic pass to be skippable too.

**On preservation**: the ordinary path never writes the typed columns at all
(SOURCES, DRAFT, SUBMIT, SUBCHARS, APPROVED, NOTES are read-never-written for
existing rows). The load-bearing "don't lose a person's drafted work" guarantee is
structural (which columns get touched), not value-by-value -- this maps cleanly onto
a bulk approach at column granularity without re-deriving any per-cell logic.

**A doc/code mismatch found while reading, independent of the speed work**: a
comment at `Drafting.bas:714-728` says parking (backup-before-write) is now
"unconditional on 'there was a sheet here'" -- the code still only parks on stranded
rows, layout mismatch, migration, or period change. The ordinary same-layout-
same-period rebuild -- the exact case that lost 27 paragraphs to a mid-write failure
on 2026-08-14 -- still takes no copy. Worth fixing the comment (or the behavior)
independent of anything below.

## The fix, two phases

**Phase A -- bulk data, no layout change.** Read the whole sheet in one
`Range.Value` call into a 2D array (using the `End(XL_UP)` idiom already proven by
items W and AB for the bound). Everything that currently reads piecemeal --
`DetectSheetLayout`, the period stamp, `isNewSheet`, the `rowOf` index, the
stranded-row count, the rollover-ferry SUBMIT values, `CitedBlockFor`'s per-row scan
-- reads the array instead. The main loop computes into in-memory arrays, then
writes back as ~4 column-sized `Range.Value = arr` operations (ENTITY, NAME,
CURRENT, CHARS) instead of ~172 per-cell writes. On a quarter turn: one PREV column
write, read it back in bulk to confirm (evidence from the far side of the save
boundary, this project's own rule), THEN clear the six work columns -- same
ordering discipline as today, just at column instead of cell grain.

Two real subtleties, not glossed over:
- **Orphan rows** (on the sheet, no longer in the register) must carry their
  just-read values back through the bulk write. The hazard: a Variant round-trip can
  flip a numeric-looking text value (e.g. `'0123`) back to a number. Today's writer
  already `'`-prefixes NAME and CURRENT -- the bulk writer keeps that per column.
  The parity harness (below) checks text-ness cell-by-cell, not just value.
- Rare paths (`MigrateSheetLayout`, the stranded-row fallback) stay untouched --
  they run a handful of times a year; touching them spends risk on the exact code
  that wiped 129 paragraphs for negligible time savings.

**Phase B -- cosmetic skip.** Stamp a look-version + row-count marker into a
tool-owned intro cell, written ONLY after `ApplyDraftingLook` and friends actually
complete. On a rebuild where the stamp matches, skip the whole cosmetic pass (~100
ops, ~1-1.9s at measured rates); any mismatch (new sheet, appended rows, a bumped
version constant, migration) pays the full pass. This is the piece that gets 10x
rather than 2-3x.

**Expected arithmetic**: Phase A alone ~2.5x. Phase A+B on a normal refresh:
~600 -> ~25-40 ops, an estimated **0.15-0.35 sec/field against the measured
2.5-4.6s/field -- 10-20x**. The 13-field stage: ~40s -> ~2-4s. A brand-new sheet
still pays cosmetics once (~1-1.5s), which is honest and correct.

## One-sheet-per-type consolidation (Lobby Phase 4) -- evaluated, not bundled

Explicitly asked: does merging the 13 per-field drafting sheets into one sheet per
slide type also serve the speed goal? **Marginally, and not on the critical path.**
After Phase A+B the 13-sheet stage is already ~2-4s; consolidation might save
another second or two by removing per-sheet constant overhead, but that overhead is
mostly what Phase B already eliminates.

**What it costs**: `WriteDraftingSheet` assumes fixed absolute row positions
(`DRAFT_INTRO_ROW`, `DRAFT_HEADER_ROW`, `DRAFT_FIRST_ROW`, used ~30 times) --
consolidation means every one of those becomes offset-aware, in the function with
five prior real data-loss incidents including the 2026-08-14 wipe (129 paragraphs,
43 ticks, 75 notes). `LOBBY-DESIGN.md`'s own gate for Phase 4 -- "whether Phases 0-3
alone solve the tab-clutter complaint once used for real" -- has not been exercised
yet; Phase 3 only deployed the morning of 2026-08-17.

**Recommendation: do not bundle.** Ship the speed fix now, per-field sheets intact.
Revisit consolidation on UX grounds only, after a real quarter on Phases 0-3 --
exactly the existing gate, not a new one. Real sweetener: Phase A makes a future
Phase 4 substantially cheaper, since row logic keyed off an in-memory array with
sheet coordinates isolated to one read makes "offset-aware" a `firstRow` parameter
instead of surgery across ~30 call sites. Speed fix first is the argument for doing
them in that order, not together.

## Risks and validation path

Named hazards: silent type/format drift through the Variant round-trip; a bulk
write landing one column off; ferry-before-clear ordering inverted at column level;
the cosmetic-skip stamp passing vacuously; a mid-write failure between column
writes (still bounded to tool-owned columns on the ordinary path -- strictly smaller
blast radius than today's per-cell equivalent, not larger).

**What existing tests already pin** (`vba/tests/TestRunner.bas`, all against real
Excel): same-period preservation, rollover ferry + clear, idempotence of a second
rollover, stale-value bleed, layout-stamp bootstrap, layout-3->4 migration, cited
source reaching the prompt, publish gating. All exercise `WriteDraftingSheet`
end-to-end, so they run against the bulk rewrite automatically. But every existing
assertion is on ONE row (P001).

**New, before any rewrite lands**:
1. A multi-row preservation test -- 5+ rows, all five typed columns populated,
   including one orphan row and one numeric-looking text value. Rebuild, assert
   typed columns byte-identical and no type flips.
2. A parity harness -- legacy vs. new implementation on cloned sheets across four
   states (new sheet, same-period with typed work + orphans, rollover, layout-4
   source), comparing the full used range value-by-value including text-ness.
3. A cosmetic-stamp test.

Per this project's own rule, each gets made to fail on purpose before being
trusted -- point the new code at the wrong column and confirm the harness names it;
skip one column write and confirm the multi-row test catches it.

**Rollout**: build `WriteDraftingSheetBulk` alongside the legacy function, switch
behind one constant; full suite green + static checks; deliberate-fail demonstrated;
swap the one call site (`DraftingUI.bas:825`); keep legacy callable for one addin
release as fallback; live proof on the real test deck via the `Timing` sheet's own
per-field numbers, before/after, plus one preserved SUBMIT and one ferried rollover
verified from the saved workbook's own bytes -- never an in-place live rewrite.

## Sequencing

0. **First (~30 min): close the measurement gap.** By item AB's own numbers,
   665.4s total - 503.4s Lobby - ~40s drafting sheets leaves ~120s inside the same
   button entirely unattributed -- `WriteSpecSheet`, `WriteSourcesSheet`,
   validations, tab arranging, the workbook index, register formatting are all
   inside `tRefresh` with no `Timing.LogTiming` of their own. One line per stage
   settles whether drafting sheets are really the #2 cost or just the largest
   *measured* one. Cheap, and de-risks everything below from being aimed at the
   wrong target.
1. **Phase A** (bulk data + the two new tests + parity harness) -- one focused
   session.
2. **Phase B** (cosmetic stamp/skip) -- roughly an hour, same or next session.
3. **Live proof** as part of a real "1. Set up my quarter" press; retire the
   legacy function the session after.
4. **Follow-on, priced by step 0's data**: `FieldSpec.WriteSpecSheet` (62
   `.Cells(` in-module), `WriteSourcesSheet` (41), `FormatRegisterSheet`/
   `WriteWorkbookIndex` (20), `ExcelOutput` (44) are all suspected to share the
   same per-cell-COM-call shape. If step 0 shows the missing ~120s living there,
   the identical bulk-array fix applies with LESS risk than drafting sheets, since
   none of those hold irreplaceable typed content.
5. **Phase 4 consolidation** -- parked behind its existing gate in
   `LOBBY-DESIGN.md`, revisited after a real quarter on Lobby 0-3, made cheaper by
   Phase A having already landed.
