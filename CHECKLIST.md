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

## Win ledger — reviewed 2026-08-16, one open call for Rohan

*Source: `archive/correspondence/chat-to-code-2026-08-16-win-ledger.md`, sent
unprompted alongside the architecture reply. Chat's own reconstruction from
documents, flagged as hypothesis, not read from the repo — treated that way here.*

- [x] Confirmed the stated rule matches what this project actually enforces:
      proven **by button**, on **real files**, verified from the **saved
      bytes** — a green suite alone is not a win. Matches
      `feedback_tested_unit_behind_locked_door` and the reachability-defect
      pattern this session already knows well.
- [x] **`DELETIONS.md` created** — one line per killed thing (what, the
      question that killed it, what it would have cost), per chat's suggestion
      that rejections leave no trace the way decisions do. Seeded from the
      ledger's known casualties list.
- [ ] **STILL OPEN, needs Rohan's call, not resolved here:** the "5 of 9"
      scenario count has a moving denominator — 7 scenarios on 14 Aug, 9 on 15
      Aug, unnumbered plain-language framing as of 16 Aug — so the fraction
      isn't comparable across dates and can read as flat when it isn't. Chat's
      own recommendation is to drop the fraction and keep only a dated list of
      byte-verified closures (which is what `NEXT-SESSION.md` and this ledger
      section already do in practice). **Not changed here** — "5 of 9" is
      embedded through `NEXT-SESSION.md`'s history and `SCENARIOS.md`'s own
      title, and a reframe touches both; Rohan's call on whether to freeze the
      nine as the denominator or drop it.
- [ ] Two classes of win chat flagged as leaving no trace at all: **deletions**
      (now addressed by `DELETIONS.md`) and **correct refusals** — a refusal
      that names no alternative reads downstream exactly like a dead end. No
      entry format proposed for the refusal class yet; open.

## ARCHITECTURE FORK — RESOLVED 2026-08-16. "Template" was doing two jobs.

*Full exchange preserved: `archive/correspondence/code-to-chat-2026-08-16-
template-architecture-question.md` (the question) and
`archive/correspondence/chat-to-code-2026-08-16-template-architecture-colour-vs-
structure.md` (the answer).*

- [x] **The 12 Aug "one template, not three" ruling and Scenario 3's per-letter
      plan do NOT conflict — they're answering different questions.** The 12 Aug
      ruling is about slide **TYPE**: one project-progress design. Scenario 3 is
      about template-slide **VARIANTS** of that one type (K/S/P colour). Checked,
      not assumed: the 12 Aug ruling was built by enumerating actual K/S/P
      differences — slot count, subtitle segments, team rows — and colour was
      never in that list. The ruling's own text carves colour out explicitly:
      *"colour MAY be applied from a declared spec... SIZE stays
      visibility-driven"* — two mechanisms, colour assigned to the other one.
      **Scenario 3's existing plan (per-letter registration, separate K/S
      template slides) is correct as written. Proceed with it.**
- [x] **Vocabulary, so this doesn't recur:** *type* decides which fields exist
      (register columns, Field Spec rows). *Variant* decides how a type looks —
      owns nothing of its own. *Template slide* is one registered `is_template`
      slide implementing one type/variant pair. The test that keeps type and
      variant apart: adding the K or S template must touch **zero** field
      definitions. If it ever needs the 29 fields declared three times, the
      design has failed and decision (1) has actually been violated.
- [ ] **The condition, recorded so it isn't silently violated or re-litigated
      blind:** three drawn templates beats runtime recolour **only because no
      colour-role declaration mechanism exists today.** If the parked
      `=FIELD[#INDEX][.PART]` naming thread is ever built, that would supply
      exactly such a declaration and the right answer for colour specifically
      flips to runtime recolour. **Not a live alternative right now** — checked
      2026-08-16, that thread's own first prerequisite (does `=`/`$` survive
      copy/paste, group/ungroup, Reset Slide in PowerPoint) has never been run;
      the test protocol sits with an empty results table. Nothing today argues
      for switching away from cloning.
- [x] `SCENARIOS.md` scenario 3 corrected — it described three colour templates
      as already colliding; only one (`P`, green) exists today. Fixed 2026-08-16.

**Constraint confirmed by Rohan, 2026-08-16, applies regardless:** circle/shape
SIZE must stay pre-drawn, never computed at runtime. Matches the milestone
device's own rule exactly — `SetVisible` only ever toggles which pre-drawn shape
shows, never resizes one.

## Output slide type — real, mostly designed, one ruling still needed from Rohan

*Source: `archive/correspondence/chat-to-code-2026-08-16-template-output-slide-
notes.md`, comparing notes per Rohan's request. A separate, python-pptx-based
pipeline, not the VBA add-in this checklist otherwise tracks.*

- [x] `OUTPUT-template.pptx` exists (built 14 Aug, single slide, untagged), with
      a genuinely clean naming scheme: `OUT_` for addressable shapes (99),
      `CHROME_` for decoration nothing ever touches (27). Matches the milestone
      device's own rules on purpose — visibility-only state changes, one named
      exception where geometry carries the value (`OUT_MSn_BAR` length), no
      runtime colour changes at all on this slide type.
- [ ] **Shape count doesn't reconcile** — the build sheet specifies 119, the
      built file has 133 (`OUT_` 99 + `CHROME_` 27 = 126). Three different
      numbers, nobody has checked which is right. Not urgent, not forgotten.
- [ ] **Real blocker: `STRATEGIC_LINKAGES` has no register column**, so the
      chips this slide type needs have nothing to inject from at sync time.
      Same missing-column pile as `HIGHLIGHTS_BODY`. A cheaper prerequisite
      sits under it: 48 unscored linkage lines across 12 projects (10 of 12
      S-type) need scoring, since the contribution score becomes a chip weight
      here specifically.
- [ ] **STILL OPEN, needs Rohan's ruling specifically — not chat's, not mine:**
      whether the python-pptx pipeline and the VBA pipeline are a one-way
      handoff (python-pptx draws the Output template once from the Family Tree,
      VBA onboards it as a new slide type and owns it forever after, the two
      never talk again) or something else. Chat's explicit that this is a
      proposal, open since 12 Aug, not something either AI should adopt
      quietly. If the handoff reading is right, the two pipelines never need to
      interoperate at runtime — but it decides real build work, so it needs to
      actually be decided.

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

## Found while archiving `FIRST-REAL-RUN.md`, 2026-08-16

- [ ] **Team distribution / multi-user — deliberately parked, genuinely unresolved,
      not previously on this checklist.** *Source: `archive/FIRST-REAL-RUN.md`,
      "Open, parked deliberately," 2026-08-01.* Rohan raised it the same day it
      was parked: *"hang on we dont have to do it now."* It reverses the
      "personal tool, not org adoption" decision the whole current architecture
      rests on (see `project_deck_sync` memory, answered 2026-07-28). Forces two
      things if ever picked up: code-signing the `.ppam` becomes required (not
      optional), and the register becomes shared mutable state with no
      concurrency control — two people publishing at once can overwrite each
      other, a rebuild can wipe a sheet someone is typing into, OneDrive's
      conflict-copy resolution could make the add-in silently read the wrong
      file. **Unanswered, and decides the shape of any solution:** how many
      people, and one shared register or one each? Not urgent — genuinely
      parked — but real, and adjacent to Scenario 8 (portability), which
      currently assumes one person only.
- [ ] **No self-service way to unmark a single field** — `Clear Marks` discards
      every mark, cannot remove just one (confirmed still true today, per
      `WORKFLOW.md`'s own SETUP table). *Source: `archive/FIRST-REAL-RUN.md`
      finding 7, 2026-08-01.* The three specific buttons it proposed
      (`List Marked Fields`, `Unmark Field`, `Unmark By Name`) were never
      built and are themselves now stale against the current chain-based
      toolbar model — but the underlying gap persists. Minor, not urgent.

## Drafting sheets — make them easier to work with (flagged by Rohan, 2026-08-16)

Specifics not yet defined — flagged as a goal, not a spec. Related papercuts
already on this list belong under this heading rather than scattered:

- [ ] A bare Excel "permanently delete this sheet" prompt fires during the
      drafting-sheet rebuild — no context, not the tool's own words. (moved
      from "Known open defects" below — same goal)
- [ ] `Roll Forward` requires clicking a cell to name the source quarter, when
      the tool already knows the deck's period and which periods have rows.
      (moved from "Known open defects" below — same goal)
- [ ] Per-field drafting sheets stay separate (reversed from a long-format
      merge, see "Parked" below) — so whatever "easier to work with" means, it
      works within that shape, not by re-opening the merge question.
- [ ] Rohan to add what's actually slow or annoying day to day — this section
      exists to collect it.

## Known open defects — not urgent, not forgotten

*Source: `FIX-LIST.md`, live entries.*

- [ ] `FIX-LIST` P: cloud persistence on the 4 setup document-properties is
      intermittent (~50% land rate), uncharacterised beyond "eight causes
      eliminated." Workaround stands: setup writes on a local copy.
- [ ] Review grid may not be safe to rebuild under a live Excel AutoFilter —
      strongly implicated, not proven. Proving test: re-apply a filter, rebuild,
      check for duplicate rows.
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
