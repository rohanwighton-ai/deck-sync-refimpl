# Checklist

> Flat, current, one line per item, each linked to where it actually comes from.
> Compiled 2026-08-16 from a full pass of every document `DOCUMENT-MAP.md` marks
> CURRENT — not just the three files usually cited. This is now the primary
> handover surface: `NEXT-SESSION.md`'s CURRENT block points here first. Add to
> it, tick it here, together — don't let it drift back into prose.

## Immediate — nothing else is real until this happens

- [ ] Build `addin104`: run `build_ppam.ps1`, Save As `addin104`, tick it, untick
      `addin102`, restart PowerPoint. Four fixes from 2026-08-15 (Q, R, the
      drafting-report labels, the readiness partial-quarter check) exist only in
      source. *Source: `NEXT-SESSION.md`, "FIRST ACTION" block.*

## The actual finish line

- [ ] **Item 10 of `TRACKER.md`'s own 10-item list — "one real quarter produced,
      and it saved time."** This is the project's real completion criterion, above
      and separate from the nine scenarios: *"Done when: Rohan says the sentence.
      'No' is a valid answer and a spec for what to fix."* Items 1-9 are all
      ticked. *Source: `TRACKER.md:47-111`.*

## Scenario 1 — the last piece of "the quarter"

- [ ] Review the 43 rolled-forward drafting rows for `Q1F27` on the rig — content
      decision, not a coding one. *Source: `SCENARIOS.md` row 1, `NEXT-SESSION.md`
      critical path #1.*
- [ ] Tick `APPROVE` for what's correct, leave the rest blank.
- [ ] Publish, **unaided, no Claude in the loop** — the only action that moves the
      count past 5/9. *Source: `SCENARIOS.md`'s own pass condition, line 22-25.*

## Milestone device — Q and R fixed in source, two things still owed

- [ ] Confirm Q + R actually work on slide 1 of the real rig deck once `addin104`
      is loaded — evidence so far is tests, not a live slide. *Source:
      `NEXT-SESSION.md`, "What was not done" block.*
- [ ] `InjectDeviceVia` always reports `WouldChange = True`, even when nothing on
      the slide differs from the register. Will pollute the review queue once real
      milestone data exists. Needs a real current-vs-proposed comparison. *Found
      2026-08-16, not previously written anywhere — see `InjectPrimitive.bas`,
      `InjectDeviceVia`, every branch sets `WouldChange = True` unconditionally.*

## Provenance (scenario 9) — fully designed, five concrete steps, none built

*Source: `PROVENANCE.md` in full — the design already answers "how", these are
just its own stated build steps, restated as a checklist.*

- [ ] Step 1 — `Drafting.PublishDrafts` writes `<FIELD>_SOURCES` and
      `<FIELD>_RECIPE` alongside the value, in the same `UpsertRow` call. One
      write, not three — a citation written when the value wasn't (or the
      reverse) is worse than neither.
- [ ] Step 2 — `FieldSpec.RecipeHash(specWs, fieldId)`, a hash over the row's
      guidance cells, stable across whitespace/case so a cosmetic edit doesn't
      read as a changed recipe.
- [ ] Step 3 — `ExcelOutput.UpsertRow` — no change expected; provenance columns
      are ordinary fields to it.
- [ ] Step 4 — `FieldWiring` must NOT report provenance columns as fields with no
      shape. Register-only, narrow enough to derive from the column suffix.
- [ ] Step 5 — drafting sheet needs no change; the SOURCES column already exists
      and is already read at publish.
- [ ] **Refusal rule 1:** a value published without its provenance must not be
      written at all — a half-written pair looks complete and isn't.
- [ ] **Refusal rule 2:** an uncomputable recipe hash refuses the publish, same
      reasoning.
- [ ] **The proof, once built:** publish a field with citations in period A, roll
      forward to B, assert A's row still carries its sources/hash unchanged, and
      that B's row inherits A's *values* but NOT A's *provenance* — the subtle
      failure mode named directly in the design.

## Source capture — two real, separate jobs

*Source: `SOURCE-CAPTURE-FORM.md`, "Before you start: clear the scaffolding."*

- [ ] Delete rows 6-11 of the `Sources` sheet — `S01`-`S06` are fabricated example
      sources from 2026-08-08, not real documents.
- [ ] Blank the SOURCES column on `TPL_ABOUT_BODY` for all 37 rows that currently
      cite them. **Doing only the deletion is worse than doing neither** — the
      citations would survive pointing at IDs that no longer exist, and publish
      reports them as unknown refs on every run. *Corroborated independently in
      `archive/HANDOVER-Q4F26-DRAFTING.md` §6 (2026-08-10, two days earlier than
      `SOURCE-CAPTURE-FORM.md`) — same defect, same 37-row count, named twice
      before either doc knew about the other.*
- [ ] Whether Copilot can open a SharePoint path handed to it — answerable at work
      in a minute. Determines whether a citation is a real trace (the model read
      it) or an attestation (a person read it and typed the ID). *Also in
      `COLUMNS.md`'s open list.*

## A stale claim caught while compiling this — correct it, don't act on it

- [x] **`SOURCE-CAPTURE-FORM.md`'s "are three fields missing from sync?" question
      is answered and the doc is stale, not open.** It says only `ABOUT_BODY` and
      `KEY_EVENTS_BODY` sync. As of 2026-08-15, `TPL_` drafting sheets exist for
      at least nine prose/label fields (`ABOUT_BODY`, `STRATEGIC_ALIGNMENT_BODY`,
      `PROBLEM_BODY`, `PROGRESS_BODY`, `HIGHLIGHTS_BODY`, `STRATEGIC_LINKAGES`,
      `DELIVERABLES_BODY`, `MS2`-`MS6_LABEL`, `KEY_EVENTS_BODY`) — confirmed
      directly from the workbook's own sheet list, not inferred. **Someone should
      correct the doc's banner to CURRENT-but-this-question-answered**, but there
      is no real work item behind it.

## Register field questions — small, concrete, standing

*Source: `COLUMNS.md`, "Open."*

- [ ] `PROJECT_STATUS` casing disagrees with the deck: slide reads `In progress`,
      register/Field Spec read `In Progress`. Silent mismatch — settle against the
      real deck, not the rig. (May already be moot post the 8-slide casing fix on
      2026-08-15 — check before treating as open.)
- [ ] Which source is "the dedicated one" for `STRATEGIC_LINKAGES`, and who
      maintains it now the Family Tree is going. A work question, costs a minute.

## File-per-quarter — the prune half (critical path #3)

- [ ] Design + build the prune: drop the old period's rows from the live register
      once it's archived.
- [ ] Retire `ParkSheetCopy` once the prune lands (load-bearing until then — do
      not delete early). *Source: `DOCUMENT-MAP.md` decision 6, "the live gap this
      exposes."*
- [ ] Sweep `Sync Log` into the same per-quarter archiving. *Source:
      `SCENARIOS.md`'s file-per-quarter section.*
- [ ] Tests + one real keyboard run before the prune touches anything live.

## ARCHITECTURE FORK — Rohan's decision needed before scenario 3 continues

- [ ] **A ratified 12 Aug decision may conflict with the Scenario 3 plan below.**
      *Source: `archive/NEXT-SESSION-2026-08-12.md`, "DECISION, 12 Aug evening —
      pre-placed shapes + visibility, NOT computed sizing," found 2026-08-16.*
      Verbatim: *"positions are pre-drawn and known; state is shown by hiding and
      showing, never by resizing at run time... this generalises beyond the
      timeline, and is why the deck should NOT split into three slide types...
      one template, not three."* Scenario 3's plan (below) does the opposite —
      builds SEPARATE K and S templates. The plan doesn't reference this
      decision; it may have been written without seeing it. **Not resolved —
      needs Rohan's call before step 2 of Scenario 3 starts:** either the
      per-letter-template plan is wrong and should become "one template,
      pre-placed K/S variant shapes, shown/hidden like the milestone circles,"
      or this decision was specific to the timeline device and doesn't
      generalise the way its own text claims. Either answer is fine; not
      deciding and building both is not.

## Register/field items found in the 12 Aug cross-surface handover

*Source: `archive/NEXT-SESSION-2026-08-12.md`, all items checked against current
code 2026-08-16, not assumed.*

- [ ] **`TOTAL_VALUE` alarm — confirmed NOT built.** Block publication when
      `TOTAL_VALUE <> INDUSTRY_CASH + SAAFE_CASH + TOTAL_INKIND`. Live slide was
      out by $646 and shipped that way. Check first whether the register stores
      rounded display values — exact equality against rounded inputs would fail
      permanently, which is the always-firing warning that stops being read.
- [ ] **Linkage-subset check — not built.** `STRATEGIC_ALIGNMENT_BODY` may cite a
      subset of the codes in `STRATEGIC_LINKAGES`, never a code not declared
      there. Extract from both, report the difference — Copilot can't self-check
      this since the declared codes live on a different sheet.
- [ ] **`Kind = Derived` — confirmed NOT built.** No `KIND_DERIVED` anywhere in
      `FieldSpec.bas`. A fourth Kind value plus a `Derivation` column, for values
      like elapsed-time-% and the current-milestone marker that must be computed
      from other fields, never stored (a stored copy of a computed value is the
      drift this project already designed out once). **Must land with a carve-out
      in the same change**: `COLUMNS.md`'s bidirectional completeness check
      (every register column has a Field Spec row and vice versa) would report
      every Derived row as an orphan forever otherwise, and an always-firing
      warning stops being read.
- [ ] **The orphaned `cadence` parameter — status unclear, worth a real look.**
      The handover said it lived in `Drafting.WriteDraftingSheet`, read a retired
      `Quarter = ALL` sentinel, and fell through to "unknown" silently on every
      field. It's gone from `Drafting.bas` entirely now (checked), but "cadence"
      still appears in `ExcelOutput.bas`, `DraftingUI.bas`, `ReviewQueue.bas` —
      could mean it was properly relocated, could mean something else. Not
      confirmed either way.
- [ ] **Chars columns must be written as live formulas, not static numbers.**
      `TPL_` columns H and I (character counts) were static — H frozen at a past
      value, I blank on every row of every sheet, so length-against-target has
      never been checkable. `Drafting.WriteDraftingSheet` must write
      `=LEN(C{row})`/`=LEN(F{row})`, never a computed literal, or the next
      rebuild silently reintroduces the bug. Not verified against current code.
- [ ] **`SRC_EXTRACTS` lookup formulas were hardcoded to a row count** and
      hand-widened once. Anything that regenerates that sheet needs to match the
      wider range, or it silently reverts. Not verified against current code.

## Scenario 3 — per-letter templates (blocked on a real defect, not reachability)

- [x] Step 1 — `TemplateSlide.CodeLetterOf`, done and tested.
- [ ] Step 2 — per-letter registration property (`DeckSyncTemplate:<type>:<letter>`
      alongside the existing `DeckSyncType:`, with fallback for untouched decks).
- [ ] Step 3 — choose the template **per row**, inside `RunSync.CreateMissingSlides`
      (re-run scenario 2 after touching this — same code path).
- [ ] Step 4 — relax the one-per-type guards (`RibbonUI.bas:2375` and
      `MakeTemplateFrom`) to one-per-type-**per-letter**, before step 5 or it's
      refused outright.
- [ ] Step 5 — the deck surgery: make the `K` and `S` templates from real slides,
      on a copy, never the live deck.

## Scenario 8 — portability (never tried)

- [ ] Bring up a genuinely fresh deck + fresh register from nothing, unaided —
      the standing requirement, the tool has to travel with Rohan.

## Known open defects — not urgent, not forgotten

*Source: `FIX-LIST.md`, live entries.*

- [ ] `FIX-LIST` P: cloud persistence on the 4 setup document-properties is
      intermittent (~50% land rate), uncharacterised beyond "eight causes
      eliminated." Workaround stands: setup writes on a local copy.
- [ ] Review grid may not be safe to rebuild under a live Excel AutoFilter —
      strongly implicated, not proven. Proving test: re-apply a filter, rebuild,
      check for duplicate rows.
- [ ] `Roll Forward` requires clicking a cell to name the source quarter, when the
      tool already knows the deck's period and which periods have rows.
- [ ] A bare Excel "permanently delete this sheet" prompt fires during the
      drafting-sheet rebuild — no context, not the tool's own words.
- [ ] The apply-confirmation dialog title is hardcoded wrong when reached from
      `2. Put it on the slides`.

## Explicitly out of scope — from `TRACKER.md`, do not re-add without a reason

The other ~30 unwired fields. A GUID-based key redesign. R13's full review
subsystem (built, parked). Chrome/UI enforcement. Ribbon polish. Adding these
here would be the "field count as progress" trap the project already fell into
once. *Source: `TRACKER.md`, "Not on this list, deliberately."*

## Parked / explicitly decided — do not reopen without a new decision

- [x] ~~Merge the 13 per-field drafting sheets into one long-format sheet.~~
      **Reversed by Rohan, 2026-08-16**: too much time already lost to redesign
      detours on this project; staying with per-field sheets to keep momentum.
      Never written into any doc before this entry.
