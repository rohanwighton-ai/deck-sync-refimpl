# Checklist

> Flat, current, one line per item, each linked to where it actually comes from.
> Compiled 2026-08-16 from a full pass of every document `DOCUMENT-MAP.md` marks
> CURRENT — not just the three files usually cited. This is now the primary
> handover surface: `NEXT-SESSION.md`'s CURRENT block points here first. Add to
> it, tick it here, together — don't let it drift back into prose.

## Immediate — nothing else is real until this happens

- [x] Build `addin104`: run `build_ppam.ps1`, Save As `addin104`, tick it, untick
      `addin102`, restart PowerPoint. Four fixes from 2026-08-15 (Q, R, the
      drafting-report labels, the readiness partial-quarter check) exist only in
      source. *Source: `NEXT-SESSION.md`, "FIRST ACTION" block.* **Done
      2026-08-16** — 33 modules imported clean, build stamped `2026-08-16
      11:02`, Rohan confirmed `addin104` loaded and `addin102` unticked.

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

## Scenario 3 step 5 — the deck surgery. Needs Rohan at the keyboard.

*Not automatable past this point* — `Create Template Slide` now drives
`InputBox`/`MsgBox` (the new `PickTemplateSource` picker from step 4), and a
modal blocks headless COM the same way the `.ppam` Save As step always has.
Scripting past it would also defeat the point: this step is what proves the
human-facing flow works, not just the code behind it.

- [x] `addin104` predates steps 2-4 entirely. Rebuilt 2026-08-16 12:14 — 34
      modules imported clean. Superseded by `addin105` (below), then by
      `addin106` (below) — this line is only about clearing addin104.
- [x] A fresh, dedicated copy made for this operation — never the live deck,
      and not reusing `PRESERVED-known-good-20260815-1050` in place (other
      scenarios may depend on that one staying as it is). Copied to:
      `AppData\Local\deck-sync-backups\scenario3-template-surgery-20260816\`
      (the Project Progress deck + `register-wide.xlsx`).
- [x] **THE COPY'S WORKBOOK PAIRING WAS ACTUALLY WRONG, confirmed live, not
      just from bytes.** Rohan opened the copy, pressed "Change which
      workbook this deck uses", and the dialog itself showed the pairing
      still pointing at `...\PRESERVED-known-good-20260815-1050\`, matching
      what the saved bytes said before anyone touched anything. Retyped to
      `...\scenario3-template-surgery-20260816\register-wide.xlsx` and the
      tool reported back **"Confirmed in the saved file, and this deck's
      slide type still finds its sheet there"** — verified against disk, not
      a cache read-back, per `SetWorkbookPathVerified`'s own design. Fixed.
- [x] **`addin105` built with steps 2-4 but "Create Template Slide" had no
      button at all — a second real defect, found live, not by inspection.**
      Rohan pressed it and "nothing happened" because there was nothing to
      press: `CreateTemplateSlide` was reachable only as a one-time MsgBox
      at the end of Bulk Onboard Type, an assumption (one template per type,
      made once) Scenario 3 breaks on purpose. This is the SAME bug class
      `CommandBarUI.bas`'s own header already names once — *"Readiness
      offered 'Create Template Slide' as a remedy for a button the toolbar
      has never carried"* — recurring in a different module. Fixed: added
      `CAP_CREATE_TEMPLATE`, wired a real repeatable toolbar button, flipped
      the two reachability tests that had asserted its absence (they were
      right when written, in 2026-08-01), pointed every dialog title in
      `CreateTemplateSlideCore`/`PickTemplateSource` at the new constant
      instead of a hardcoded literal. Suite 222/0 both before and after.
      Add-in rebuilt as `addin106` (superseding the button-less `addin105`,
      which should not be ticked). **Still needed from Rohan:** File > Save
      As > PowerPoint Add-in, name it `addin106`, tick it, untick whichever
      of `addin104`/`addin105` is currently ticked, restart PowerPoint.
- [ ] **Then, in order, on the copy (pairing already fixed above):**
      1. Press **"Create template slide"** (now a real button — bottom
         right of the Add-in ribbon group). Type auto-picks (only one type
         registered: `project-progress`).
      2. The source picker lists every real onboarded instance by key +
         derived letter. **Pick a K-lettered instance** (15 exist). Confirm
         the summary dialog.
      3. Press **"Create template slide"** again. **Pick an S-lettered
         instance** this time (17 exist). Confirm.
      4. Do **not** pick a P-lettered instance in either pass — see the known
         gap noted below.
- [x] **Both done, 2026-08-16.** K built from `1_K1001` (slide 12), S from
      `1_S001` (slide 27). Both dialogs correctly said the new slide "will
      not appear in Preview Sync or Sync Now reports" and named which
      instance each was copied from.
- [x] **Found a THIRD real defect live, mid-flow: the confirmation dialog's
      own wording was stale.** It unconditionally said `'project-progress'
      RE-REGISTERED to clone this new slide from now on` — true for the
      old one-template-per-type world, false the moment a type already has
      a template and a second letter is being added (only that letter's own
      slot gets registered; the existing fallback is deliberately left
      alone). The underlying WRITE was already correct — this was a
      text-only defect, caught by reading the actual dialog before Rohan
      clicked through it, not by the pinned test, which had only ever
      exercised the letter-less case. Fixed: `ConfirmTemplateText` takes
      `letter`/`willClaimFallback` now and states the real scope — "ONLY
      letter 'K' rows" vs "the FIRST template for this type, so it ALSO
      becomes the default." 2 new tests. Suite 222→224/0.
- [x] Verified from the SAVED file, not a dialog: 46 slides (was 44). Slides
      45/46 both hidden, both tagged `is_template=1`/`slide_type=project-
      progress`, neither carries an `instance_key`.
      `DeckSyncTemplate:project-progress:K` → slide 45,
      `:S` → slide 46. `DeckSyncType:project-progress` **unchanged**, still
      `303|Register` — the original P template, confirming the fallback was
      not stolen. Workbook pairing still correctly self-referential.
- [x] **Known gap, still accepted, not fixed here:** the existing green `P`
      template still only holds the plain `DeckSyncType:project-progress`
      property, never a per-letter one — see the reasoning recorded above
      this list before step 5 ran. Unaffected by anything in this pass;
      P rows still resolve correctly via the fallback path.
- [ ] **`SCENARIOS.md`'s scenario 3 row rewrite** (flagged since step 4) —
      still pending. Now genuinely ready to write from observed behaviour
      rather than reasoning about it.
- [ ] **Not yet decided: does this land on the LIVE deck, or does the copy
      stay the proof-of-concept?** Everything above happened on
      `scenario3-template-surgery-20260816\`, never the real deck — that was
      the point. Rohan's call, not something to do unilaterally: the real
      deck getting K/S templates is a separate, deliberate step whenever he
      wants it, using the same now-fixed button and now-accurate confirmation
      text.

## Real end-to-end tests, tracked explicitly — Rohan, 2026-08-16

Unit tests prove a function behaves when called; nothing in this list is
satisfied by one. Each entry needs a person pressing a real button against
real (or deliberately fake-but-real) data and the result checked from saved
bytes. Started because the K/S template build above was mechanism-tested
(step 2-4's suite) but never proven end-to-end until real rows forced it.

- [x] **#1, DONE and proven, 2026-08-16.** Three fake rows
      (`K900`/`S900`/`P900`, `Q4F26`) added via real Excel, verified from a
      fresh read-only re-open. Rohan pressed "Add missing slides" on the
      copy — "3 created, 0 failed." Verified from the SAVED file which
      template each actually cloned from, via a structural fingerprint
      (shape count) rather than trusting the dialog: P/K/S templates have
      genuinely different shape counts (136/114/121 — real structural
      differences, not coincidence), and each new slide's count matched
      its own letter's template exactly and no other (K900=114=K-template,
      S900=121=S-template, P900=136=P-template-via-fallback). A wrong
      routing (e.g. everything silently defaulting to P) would have shown
      up immediately as two mismatched counts — it didn't. Cross-checked via
      stable `SlideID` (303/304/305), not part filenames, because this save
      renumbered every `slideNN.xml` part -- the exact trap a filename-based
      check would have fallen into. **`RunSync.CreateMissingSlides`'s
      per-row letter resolution (Scenario 3 step 3) is now proven outside a
      synthetic fixture, on a real register, through the real button.**
- [ ] **#2:** Scenario 1 (generate a new quarter) end-to-end, Rohan alone,
      no agent in the loop — already the project's own stated finish line
      (`TRACKER.md` item 10), listed here too so it isn't only remembered
      in one place.
- [ ] **#3:** Scenario 8 (portability) — bring up a genuinely fresh deck +
      register from nothing, unaided. Never attempted once.
- [ ] **#4:** OneDrive-hosted write reliability. Proven only on a local
      copy; the one real attempt on OneDrive failed outright (AutoSave on,
      4 verified attempts, mtime never moved). The work machine is
      OneDrive-hosted, so this is not a papercut.
- [ ] Add more here as they're identified — this list is the record, not a
      one-off.

## Scenario 3 — per-letter templates (blocked on a real defect, not reachability)

- [x] Step 1 — `TemplateSlide.CodeLetterOf`, done and tested.
- [x] Step 2 — per-letter registration property (`DeckSyncTemplate:<type>:<letter>`
      alongside the existing `DeckSyncType:`, with fallback for untouched decks).
      **Done 2026-08-16** — `DeckRegistry.RegisterTemplateLetter` /
      `LookupTemplateLetter` / `LookupTemplateForLetter` added. Traced the real
      call path first: `RunSync.CreateMissingSlides` gets its `templateSld`
      from `DeckRegistry.LookupType`, not `TemplateSlide.FindTemplateFor` (that
      one's only used by Audit Fields and the MakeTemplateFrom guard) — so the
      fix belongs in `DeckRegistry.bas`, confirming the plan's own scoping.
      `LookupTemplateForLetter` tries the letter first, falls back to the
      plain type registration when the letter is `""` or unregistered — proven
      by 6 new tests, including the two-letters-don't-collide and
      prefers-letter-over-unlettered cases. Suite 203→217/0. Not yet wired to
      any caller — that's step 3.
- [x] Step 3 — choose the template **per row**, inside `RunSync.CreateMissingSlides`
      (re-run scenario 2 after touching this — same code path). **Done
      2026-08-16.** Each `new_record` row now derives its own letter via
      `CodeLetterOf(actions(i).RowInstanceKey)` and resolves its OWN template
      via `DeckRegistry.LookupTemplateForLetter`, instead of every row in the
      batch reusing the single `templateSld` the caller resolved once for the
      whole type. Also re-added the `IsTemplateSlide` check per row — a
      per-letter registration can point at a slide never actually marked
      `is_template`, the same defect class the type-level guard already
      existed to prevent, now reachable per-letter too. New test
      (`RunSync_CreateMissingSlidesChoosesTemplateByRowLetter`) proves it by
      deliberately passing the WRONG template as the type-level fallback and
      confirming each row still gets cloned from its own letter's template,
      not the passed-in one. Suite 217→218/0. Scenario 2 re-verified: its own
      dedicated tests (`RunSync_EndToEndCreatesSlidesFromFreshSheet`,
      `RunSync_CreateMissingRefusesWhileSlidesAreUnclassified`) still pass
      unchanged.
- [x] Step 4 — relax the one-per-type guard to one-per-type-**per-letter**.
      **Done 2026-08-16.** `MakeTemplateFrom` itself never actually had a
      one-per-type guard (checked before touching it) — it only refuses when
      the SOURCE slide is already a template or the wrong type; the real
      one-per-type block was entirely in `RibbonUI.CreateTemplateSlideCore`
      (`FindTemplateFor(slideType)`, type-only). Found and fixed a real
      design gap in the process: the source slide `Create Template Slide`
      cloned from was picked via `DeckRegistry.LookupType`, which points at
      the type's real onboarded slide **only until the first template is
      made** — `RegisterType` always overwrites that single property, so a
      SECOND letter could never find a representative real slide to clone
      from. Fixed by having the human pick the real source slide directly
      (new `PickTemplateSource` picker, lists real non-template instances by
      key + derived letter) — the letter then comes from that slide's own
      instance key via `CodeLetterOf`, no separate "which letter" prompt.
      Guard and registration logic pulled OUT of the untestable MsgBox/
      InputBox-driven Sub into two new testable functions:
      `TemplateSlide.ExistingTemplateForLetter` (the guard) and
      `DeckRegistry.RegisterNewTemplateLetter` (claims the letter's own slot,
      and ALSO the type-level fallback but only if nothing already holds it
      as a real template — so the FIRST letter made becomes what every
      letter-less row still resolves through, and a SECOND letter can't
      steal it). 4 new tests. Also found and fixed, in the same pass: 12
      instances of `Assert(Not x Is Nothing And x.Foo = y)` across this
      session's own tests — VBA's `And` isn't short-circuit, so that form
      raises "Object variable not set" instead of failing cleanly the moment
      `x` is genuinely `Nothing`. One had already gone live (caught by the
      suite going 221 passed / 1 ERRORED); the other 11 were dormant
      landmines never yet triggered. Added a shared `AssertSameSlide` helper
      so the whole class can't recur. Suite 218→222/0.
- [ ] `SCENARIOS.md`'s scenario 3 row still names `TemplateSlide.FindTemplateFor`
      as the blocking mechanism. That was already slightly wrong when
      written (see step 2's note) and is more wrong now — the real guard as
      of step 4 is `TemplateSlide.ExistingTemplateForLetter` /
      `DeckRegistry.RegisterNewTemplateLetter`, not a live scan. Needs a
      rewrite, not a patch — folding in with step 5 rather than done here,
      since step 5 (drawing the real K/S templates) is what will actually
      exercise this path for the first time and is the natural point to
      re-verify the row against real behaviour rather than reasoning about it.
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
