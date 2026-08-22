# Checklist

> Flat, current, one line per item, each linked to where it actually comes from.
> Compiled 2026-08-16 from a full pass of every document `DOCUMENT-MAP.md` marks
> CURRENT — not just the three files usually cited. This is now the primary
> handover surface: `NEXT-SESSION.md`'s CURRENT block points here first. Add to
> it, tick it here, together — don't let it drift back into prose.

## INCIDENT 2026-08-21 (evening) — a diagnostic sync tool silently wrote LAST QUARTER's data onto the live deck, twice, both reporting a clean save. READ FIRST.

**Second incident of the day, different shape.** `sync_real_deck.ps1` (not
the real Sync Now button — a standalone diagnostic twin of it) reads the
register via an unfiltered `ExcelOutput.ReadSheet`, whose own code comment
says "first row wins" on a duplicate instance ID — and Q3F26 rows sit above
Q4F26 rows in the register, so it silently synced last quarter's prose onto
43 live slides, twice, both times reporting "SAVED to disk (verified)."
Caught only by re-reading the saved file's own bytes independently. The real
"Sync Now" button was never affected — `RibbonUI.bas` already reads
period-aware. Full account: `FIX-LIST.md`, "INCIDENT, 2026-08-21 (evening)."

- [x] Deck restored to the last state confirmed (by direct content check) to
      predate both buggy runs and still carry the same night's milestone
      geometry fix.
- [x] Root cause found by reading `ExcelOutput.ReadSheetForPeriod`'s actual
      collision-handling code, not guessed.
- [x] Fix: `SyncRealDeck.bas`/`HiddenFixCheck.bas` now build the sheet the
      same period-aware way the real button does
      (`ReadSheetForDeckPeriod` + `RunRoutineSyncWithSheet`).
      `RunRoutineSync` (the unfiltered wrapper) left in place but loudly
      marked do-not-call, since deleting it risks an unknown caller.
- [x] Same session, same tool-drift shape: `preview_real_deck.ps1`/
      `sync_real_deck.ps1` were importing a second, stale copy of the
      production module list (missing 5 modules, one deleted-module
      reference, no CRLF handling for the one `.cls`) — now import the
      same canonical list `build_ppam.ps1` uses.
- [x] Re-verified independently from saved bytes after the fix: Q4F26
      throughout, all 16 `PROJECT_PROGRESS` values correct (tag-scoped
      check, not a blind text scan), `KEY_EVENTS_HEADER` 0/43 remaining
      placeholders, milestone geometry intact.

**Follow-on: a full mother-hound kennel audit of the live deck (Opus),
same night — real findings, real fixes, all verified from saved bytes:**

- [x] 883 stray shapes removed across 38 slides (`FIX-LIST.md` CE-adjacent —
      logged under CB) — a complete second, hand-built pre-deck-sync
      milestone timeline sitting underneath the real one on every slide,
      invisible to any tag-based check. Rohan's hypothesis ("P slides are
      clean, reset K/S to match") was half right — 6 P slides had the same
      problem. Fixed by Fable, fail-first proven (planted a text corruption
      and a height regression, confirmed both caught), one genuine near-miss
      caught (slide 2's stray timeline was bundled with a real annotation —
      separated, not deleted wholesale).
- [x] `PROJECT_STATUS` bulk-fill from earlier in the night created 8 slides
      contradicting their own body text ("In Progress" badge, "Project
      closed. No reporting submitted..." in the prose) — corrected to
      "Project Closed" for all 8, verified against the literal text on each
      slide (not inferred a second time).
- [x] `VerifyRealDeck.bas`'s own period bug fixed (same class as the
      incident above) plus its driver's stale module list — see FIX-LIST
      CE. Tool now runs for the first time in three weeks; its first real
      output (624/62 findings) is untriaged, separate work.
- [x] A full milestone-lifecycle test built — see FIX-LIST CE. Every prior
      test called `DrawMilestones` once; nothing proved a real project
      ticking off milestones over time doesn't leave a stale circle behind.
- [x] `TIMELINE_ELAPSED` fixed — see FIX-LIST CF. Raw decimals at 18pt in a
      5.5pt bar on 29 real slides, root cause was structural (no
      `.rest`/`.track` companion, so the bar injector was never reached at
      all), not just the stray-text symptom.

- [x] The real "Apply Approved" button dropping `KEY_EVENTS_HEADER` and
      `STATUS_BADGE` — see FIX-LIST CG. Both now resolved through
      `SyncOperations.ComputeDerivedValue` at apply time, same as build
      time; fail-first proven against the exact real "DROPPED" symptom.
- [x] Review-approval sheet lookup name-matching bug — see FIX-LIST CH.
      `LifespanOf` matched the retired "Sync Review" format; cosmetic only
      (feeds a human-facing report sheet, nothing acted on it), but the
      same "a fact restated somewhere its source has moved on" shape as
      the rest of this file's rot-prone areas. Fail-first proven.

- [x] The approval dialog authorising writes to all 43 slides could
      silently truncate — see FIX-LIST CI. Now capped with the actual
      Yes/No question protected as `mustKeep`; also fixed `CapReport`'s
      "full list is on the Run Log sheet" notice, which was untrue until
      `BuildAllQueuesCore` was fixed to actually write it there.

**Still open from the same audit, not yet actioned:**
- [x] `DeckAdoption.bas:101`'s own unfiltered-read — **investigated,
      false positive, no fix needed.** `PlanAdoption` calls
      `ExcelOutput.ReadSheet(ws)` but uses exactly one member of the
      result, `sheet.Fields` — and `.Fields` is built purely from the
      header row (`ReadSheetForPeriod`, before any period filter is
      applied), identical whether read filtered or unfiltered. The actual
      row-level work this function does (`ReadKeylessRows`, matching
      pre-existing untagged legacy data into the bootstrap) is a separate,
      deliberate raw-cell scan across every row regardless of period —
      correct, because that data can predate the period column entirely
      or span periods, and this is the one-time historical bootstrap, not
      ongoing sync. Same surface shape as the real period-collision bug
      (`ExcelOutput.ReadSheet` with no period arg) but the vulnerable path
      — silent first-wins row collision — is never reached, since `.Rows`
      is never consumed here. Swapping to `ReadSheetForDeckPeriod` would
      add nothing (Fields is unaffected) and would be actively wrong if
      anyone later extended it to also filter Rows.
- [x] `VerifyRealDeck`'s first real 624/62 findings — see FIX-LIST CJ.
      The 624 "no tagged shape" bucket was 100% false positive (581
      milestone-device columns addressed by name not tag, 43
      PROJECT_STATUS feeding the Derived STATUS_BADGE) — fixed, verified
      live: 624 -> 0. **The 62 mismatch bucket is real and NOT fixed** —
      held steady at 62 on a fresh 2026-08-22 run (post-dating today's
      other fixes), dominated by `KEY_EVENTS_BODY` (30) and
      `PROGRESS_BODY` (27) co-occurring on nearly every checked slide.
      Not yet triaged as pending-sync-state vs. a real defect — flagging
      rather than chasing further in this session; the systematic
      two-field pattern is worth its own look (ppt-hound or a fresh
      Sync Now against these slides would settle it fastest).

## INCIDENT 2026-08-21 (morning) — real content blanked by a real sync, restored, root-caused, fixed, deployed.

**The most serious incident this project has had.** A real "Put it on the slides"
sync blanked 117 fields of real, human-authored content across 41+ live slides
(`STRATEGIC_ALIGNMENT_BODY`/`PROBLEM_BODY`/`PROJECT_PROGRESS`), reporting a clean
run. Root cause (independent bloodhound/Fable investigation, not my first theory):
a prior session's cleanup left zero-length-string cells instead of truly empty
ones, which every downstream check — including `Harvest`'s own protection — read
as real content. Full account: FIX-LIST.md, "INCIDENT, 2026-08-21."

- [x] Deck restored from a verified pre-sync backup, real content confirmed back.
- [x] Backups duplicated to a second location, hash-verified.
- [x] Root cause found (DESIGN, triggered by an EVENT — not "harvest wasn't run,"
      which would not have prevented this).
- [x] Fix 1: `InjectPrimitive.bas` refuses to write a blank value over real slide
      content. Fail-first proven, including a real regression it caused and fixed
      (`InjectSlotsField`'s legitimate slot-clearing).
- [x] Fix 2: Harvest coverage is deck-wide (`RibbonUI.RealLinkedSlides`), not
      scoped to whatever's selected in PowerPoint (was silently checking 1 of 43
      slides). Fail-first proven.
- [x] **Found while verifying: no VBA change from this entire session had ever
      reached the live add-in** — `addin155.ppam` was unchanged since the day
      before. Rebuilt (`addin156.ppam`), deployed, verified persisted across a
      real close/reopen, verified live via the build-stamp dialog with Rohan
      reading it directly off his own screen.
- [x] **Partially closed, same night**: `PROJECT_PROGRESS` (16 of the worst-
      affected projects, sourced from `SRC_MILESTONES` column M, confirmed
      with Rohan) and `PROJECT_STATUS` (all 41 blank Q4F26 rows, set to
      "In Progress" on milestone-done evidence, also confirmed with Rohan —
      plus a new Excel dropdown so future quarters can't leave it blank the
      same way) are now real. `STRATEGIC_ALIGNMENT_BODY`/`PROBLEM_BODY`
      husk cells and the remaining 27 projects' pre-existing
      `PROJECT_PROGRESS` values were untouched (already fine) — still
      genuinely empty where the source material itself hasn't been drafted
      yet (see "extraction-to-quality loop" below).
- [ ] The exact mechanism that minted zero-length-strings instead of true Empty
      cells on 2026-08-20 was not identified — worth naming if it recurs.

## The extraction-to-quality loop — Rohan's plan, 2026-08-20, PRIORITY

**The real state, checked from `SRC_EXTRACTS` directly, not assumed:** across all 43
real projects, only the linkage-code and quarterly-comment columns (fed by `S01`) are
actually populated at scale (46/52 rows). Every other source — `PROJECT_PROGRESS` %,
milestones/deliverables (`S04`), the money fields (`S05`/`S06`), lead/partner (`S07`),
the deliverables judgement (`S08`), and the new contract source (`S16`) — is populated
for **3-4 projects out of 52**. The drafting mechanism itself is proven and correct
(fail-first tested, fixed today); the actual blocker for real content is that almost
none of the source material behind it has been extracted yet.

**The plan, stated so it survives today's volume of other work:**

1. **Rohan runs the AI extraction prompts** against the real source documents (the
   Notion prompt library, per `S01`-`S16`'s own "Where it lives" entries) to populate
   `SRC_EXTRACTS` for the empty columns, across all 43 projects. Not automatable from
   here — needs the actual source documents and the actual prompts, which live outside
   this repo.
2. **Once extraction is complete**, draft the real Q4F26 content through the (now
   fixed) drafting-sheet mechanism — `SRC_EXTRACTS` -> AI DRAFT -> SUBMIT -> Publish,
   per field, per project.
3. **Measure the generated Q4F26 material against Q3F26's** — quality, completeness,
   whether it actually reads as this quarter's own words rather than a restatement of
   last quarter's. **Not yet defined: the actual rubric or comparison method.** This is
   the first open question for whoever picks this up — "measure quality" needs an
   actual yardstick before it can be done, not just eyeballing.
4. **Triage which prompts need work** based on what step 3 finds, and iterate — refine
   the AI prompts (not the drafting mechanism, which is a separate, already-solid
   layer), re-run extraction/drafting for the fields that came out weak, re-measure.

**Explicitly not done tonight, and not blocking anything:** no real button-press test
was run against a broad set of fields — deliberately paused (Rohan: "why press put it
on the slide before that is ready?") once the extraction gap was found, rather than
prove the mechanism again on a narrow, already-well-tested slice. The mechanism proof
that matters is steps 1-2 actually running for real, at scale, which hasn't happened
yet.

## Before rebuilding the addin — standing checklist (Rohan, 2026-08-20)

Run through this every time, not just when something feels off:

- [ ] **Re-run "Refresh Drafting Sheets" on the working copy if Field Spec was
      edited directly** (outside the button flow) since the sheets were last
      refreshed. The row-2 AI-draft prompt is generated FROM Field Spec AT
      REFRESH TIME — a direct edit (e.g. via COM) after the last refresh
      leaves the prompt cell showing stale/generic text even though the
      sheet's own columns are correct. Caught live 2026-08-20: the History
      treatment column had the right values, but every prompt still said the
      generic fallback until refresh ran again.
- [ ] **Confirm `build_ppam.ps1`'s module list is current** — it went stale
      once already this session (missing `FormattingAudit.bas`), which
      silently blocked all COM compilation. No new `.bas` files were added
      today, so the risk is low, but it costs nothing to check.
- [ ] **Full VBA suite green** (`run_vba_tests.ps1`, no filter) — necessary,
      not sufficient. It proves the code compiles and each unit behaves; it
      does not prove a person's button press actually reaches the changed
      code (see `feathers-hound` / the "tested unit behind a locked door"
      pattern already burned this project twice).
- [ ] **After building, verify the NEW addin is actually loaded** before
      trusting any test against it — register it, then call a real function
      through it (`Application.Run`), not just confirm the file exists or
      `AddIns.Loaded = True`.
- [ ] **`OneDrive\Claude\` is NOT the trusted location — this is DOCUMENTED,
      not newly discovered, and got relearned live a THIRD time 2026-08-20
      (`addin155`) because `AGENTS.md`'s own COM-automation section wasn't
      checked first. Read that section before touching `AddIns` at all.**
      `File > Save As` lands the new `.ppam` in `OneDrive\Claude\`, and
      `AddIns.Add(path)` registers it AT WHATEVER PATH IT'S GIVEN — it does
      NOT auto-copy into the trusted folder every prior addin (99-155)
      actually lives in: `C:\Users\rohan\AppData\Roaming\Microsoft\AddIns\`.
      Fix: quit PowerPoint entirely first (`AddIn` COM objects expose no
      `.Delete()`, so a mis-registered entry can't be removed any other
      way), `Copy-Item` the file into the AddIns folder, `AddIns.Add()` from
      THAT path.
- [ ] **Set BOTH `.Loaded = -1` AND `.AutoLoad = -1` — setting only `.Loaded`
      is a trap that looks like success.** `.Loaded` is true for the CURRENT
      session only; `.AutoLoad` is what actually governs the NEXT PowerPoint
      launch. `addin155` was reported "loaded" 2026-08-20 with `.Loaded`
      set and `.AutoLoad` untouched — the next normal PowerPoint restart
      would have silently loaded the OLD addin instead. Also set
      `.Registered = -1` (does not default to match) and explicitly disable
      `.Loaded`/`.AutoLoad` on whichever PREVIOUS addin was active — three
      builds have raced to load simultaneously before (`addin131`/`132`,
      per `AGENTS.md`) when this step was skipped.
- [ ] **Verify by fully quitting PowerPoint and relaunching fresh, checking
      state WITHOUT touching either property again.** A reading taken in
      the same session you just set `.Loaded`/`.AutoLoad` in proves nothing
      — it's checking the value you just wrote, not whether it persisted.
      Only a genuinely cold restart proves the new addin actually auto-loads
      and the old one doesn't.
- [ ] **`MsoTriState` enum names (`msoTrue`/`msoFalse`) fail to resolve in a
      fresh PowerShell COM session** ("Unable to find type") — this is a
      RECURRING gotcha (see also `reference_vba_office_gotchas` memory). Use
      the raw integers instead: `-1` for True/Loaded/Registered, `0` for
      False. A failed enum-type assignment does not throw a script-stopping
      error by default, so it is easy to believe a `.Loaded = msoTrue` line
      succeeded when it silently didn't.

## Immediate — nothing else is real until this happens

- [x] Build `addin104`: run `build_ppam.ps1`, Save As `addin104`, tick it, untick
      `addin102`, restart PowerPoint. Four fixes from 2026-08-15 (Q, R, the
      drafting-report labels, the readiness partial-quarter check) exist only in
      source. *Source: `NEXT-SESSION.md`, "FIRST ACTION" block.* **Done
      2026-08-16** — 33 modules imported clean, build stamped `2026-08-16
      11:02`, Rohan confirmed `addin104` loaded and `addin102` unticked.

### STATUS_BADGE built (`61bd2b7`) and its shape retagged on the real deck
Rohan: "retag the badge shapes on the real deck." Confirmed first that PROJECT_STATUS
has no other reader (not harvested from a slide anywhere, not special-cased in
`FieldWiring.bas`) before touching anything — retagging its shape doesn't strand any
other mechanism.

- [x] **Retagged all 42 real `PROJECT_STATUS`-tagged shapes to `STATUS_BADGE`** (41
      real project slides + the one template, slide 44) — proven on a throwaway copy
      first, then applied to the real deck, then verified in a fresh reopen: 42
      `STATUS_BADGE`, 0 `PROJECT_STATUS` remaining, template placeholder text updated
      to `<<STATUS_BADGE>>`. **DONE 2026-08-19.**
- [ ] **The visible badge TEXT on all 42 shapes still shows the OLD raw
      `PROJECT_STATUS` value** (e.g. "In Progress") — only the TAG changed tonight.
      Nothing has actually computed and written the new derived word yet, because
      tonight's `DeriveStatusBadge` code isn't in any built `.ppam` — the Save-As step
      is still the confirmed-manual one. **Next real "1. Set up my quarter" / sync,
      once a fresh add-in build is made and loaded, is what actually refreshes these
      42 shapes to their correct derived word.** Until then the deck is correctly
      *tagged* but not yet correctly *displayed*.
- [ ] **Real, foreseeable side effect for next session, not yet handled**: `PROJECT_
      STATUS` (the underlying data field, still real and still set in the Register)
      now has ZERO shapes tagged anywhere on the deck — nothing special-cases this in
      `FieldWiring.bas`'s coverage check the way Derived/device-owned columns are
      excluded. The next field-coverage notice will very likely report `PROJECT_
      STATUS` as "0/42 wired," which is now expected and correct (it feeds
      `STATUS_BADGE`'s computation rather than being displayed itself), not a real
      gap. Worth adding the same kind of exclusion `MilestoneDevice.
      IsColumnForThisDevice` already has for device-owned columns, if the noise turns
      out to actually bother anyone — not built yet, deliberately, since it's
      speculative until it's confirmed annoying in practice.

## Workbook modernisation — surgical pass, started 2026-08-19 18:37, budget until 02:00

Rohan's framing: the Q3F26 harvest-and-rebuild plan (see chat, not yet written up as its
own doc) is NOT ready to start. Before that, the live drafting-sheet workbook needs a
**surgical, bit-by-bit** pass against what the CODE currently requires of it — we now know
those requirements precisely, which we didn't when most of the workbook's structure was
laid down. This is an audit-then-fix pass, not a rebuild. **Never write to the live file
directly — work on a copy, real backup first, per this project's own standing rule.**

### Phase 0 — before touching anything
- [x] Get the real workbook's path and take a fresh, md5-verified backup, same
      convention as every `PRE-*` backup this session already used. **DONE
      2026-08-19 18:43** — Rohan: "just use the most developed file we have been
      working on." Located via the known `OneDrive\Claude\backups` folder:
      `OneDrive\Claude\3. Project Progress.pptx` (deck) + `OneDrive\Claude\
      register-wide.xlsx` (register), both modified within the hour. Backed up to
      `OneDrive\Claude\backups\PRE-MODERNISATION-AUDIT-20260819-1843\`, md5-verified
      identical to the originals. PowerPoint was running with **zero presentations
      open** (nothing at risk) — confirmed via `GetActiveObject` before touching
      anything, not assumed.
- [x] Snapshot the live workbook's actual sheet list, column headers, and row counts
      as the "before" baseline. **DONE 2026-08-19** — opened read-only via Excel COM
      (never through `/mnt/c` directly, per this project's own OneDrive-reads-lie
      finding). **54 sheets total.** Full snapshot at
      `[scratchpad]/register_snapshot.log` this session — not copied into the repo,
      it's a point-in-time dump of Rohan's real data.

### Phase 1 — concrete drift already found by reading the code tonight, checked against the live file
- [x] **`Drafting.bas`'s column layout — CHECKED, NOT drifted.** FIX-LIST item P6
      (2026-08-13) describes `COL_D_DRAFT=5(E)/SUBMIT=6(F)/APPROVED=7(G)` — stale
      relative to the code, which now has `COL_D_SOURCES=5(E)/DRAFT=6(F)/
      SUBMIT=7(G)/APPROVED=8(H)`. **Verified live 2026-08-19: the current
      `TPL_PROGRESS_BODY` sheet's own row-9 header text already reads "E SOURCES...
      F AI DRAFT... G SUBMIT... H APPROVE" — the LIVE FILE already matches the
      CURRENT code.** P6 was stale documentation only, not a live bug. Confirmed by
      contrast: the archived `SAVED 0814-1911 PROGRESS_BODY` snapshot (13 cols, one
      fewer) still carries the OLD D/E/F/G layout P6 describes — real, in-file
      evidence of exactly when the migration happened, sitting right next to the
      current sheet. *Source: `vba/Drafting.bas:82-95`, live header rows this
      session.*
- [x] **26 `SAVED <timestamp> <field>` archive sheets found accumulated in the live
      register — since REMOVED, see "Cleanup done" below.** A real, deliberate
      mechanism (`Drafting.bas:466` creates them, `WorkbookBridge.IsToolOwnedSheet`
      recognises them), not orphaned debris — but nothing capped or pruned them.
      **Still genuinely open, separate from the cleanup**: nothing stops the SAME
      accumulation happening again from here — check against "File-per-quarter — the
      prune half (critical path #3)," elsewhere in this file, for whether a real
      retention policy exists yet or still needs building.
- [x] **The two hex-suffixed `Review ...` sheets are NOT clutter — checked against
      `ReviewQueue.bas` and left alone.** `ReviewSheetNameFor` (`ReviewQueue.bas:220`)
      derives the hex tag from a hash of the SLIDE TYPE NAME, not a per-run
      timestamp — one permanent sheet per type, cleared and rewritten on every
      "Review changes" run, never accumulating. So there will only ever be as many
      of these as there are registered slide types. **One real thing worth Rohan's
      attention, not deleted**: `Review project-status-2D3D` is still marked `OPEN`
      from **2026-08-10** (9+ days), with a genuine un-actioned pending change
      (`1_P006`'s `ABOUT_BODY`). Either it's still worth reviewing, or it's stale
      relative to newer edits and safe to let a fresh "Review changes" run overwrite
      it — Rohan's call, not touched either way. `Review project-progress-A32C` is
      `CONSUMED` (2026-08-18) — already actioned, nothing to do.
- [x] **The `Register` sheet's 54 columns — cross-checked against every FieldID on
      the `Field Spec` sheet. Clean, with one real, good finding.** 52 of the Field
      Spec's 52 listed FieldIDs match a Register column 1:1 — no orphans, no
      unexplained gaps. The two apparent "missing" columns are both deliberate:
      `TIMELINE_ELAPSED` is `Kind=Derived` (`FieldSpec.bas:231`, "fresh every sync
      from START_DATE/END_DATE, never drafted, never a register column" — by
      design). **`STATUS_BADGE` is also `Kind=Derived` but genuinely UNBUILT** — its
      Field Spec row (row 49) carries a fully-written derivation rule (combines
      `PROJECT_STATUS` + `SCHEDULE_STATUS` into one badge word, 5-branch priority
      table, one flagged `UNVERIFIED` edge case about `SCHEDULE_STATUS`'s "Complete"
      branch) with a note addressed directly to future work: *"the Field Spec has no
      dedicated Derivation column yet — that is on Claude Code's pile, deliberately
      NOT added here to avoid colliding with what they build."* Not yet referenced
      anywhere in `FieldSpec.bas`/`InjectPrimitive.bas`/`SyncOperations.bas` — real,
      specified, unbuilt work, separate from tonight's BJ fix (which only touched
      the badge's font size, not its logic). *Source: live Field Spec row 49,
      2026-08-19.*

### Cleanup done — 26 `SAVED` archive sheets removed, live register only
Rohan: "get rid of old stuff like those tabs, keeping what you want as systemic
timestamps in repo memory." Full content of every one of these sheets already lives
untouched in the Phase-0 backup — nothing here needed to be transcribed, only the
fact that it existed and when.

**Manifest, before deletion** (14 from 2026-08-14 ~19:10-19:11, in the pre-migration
13-column layout; 12 from 2026-08-17 ~18:55-19:29, in the current 14-column layout):

| Field | Timestamps archived |
|---|---|
| ABOUT_BODY | 0814-1910, 0817-1855 |
| KEY_EVENTS_BODY | 0814-1911, 0817-1857 |
| STRATEGIC_ALIGNMENT_BODY | 0814-1911, 0817-1858 |
| PROBLEM_BODY | 0814-1911, 0817-1900 |
| PROGRESS_BODY | 0814-1911, 0817-1904 |
| HIGHLIGHTS_BODY | 0814-1911, 0817-1929 |
| STRATEGIC_LINKAGES | 0814-1911, 0817-1929 |
| DELIVERABLES_BODY | 0814-1911, 0817-1929 |
| MS2_LABEL | 0814-1911, 0817-1929 |
| MS3_LABEL | 0814-1911, 0817-1929 |
| MS4_LABEL | 0814-1911, 0817-1929 |
| MS5_LABEL | 0814-1911, 0817-1929 |
| MS6_LABEL | 0814-1911, 0817-1929 |

- [x] Tested on a throwaway copy first — 54 sheets to 28, reopened clean, `Register`
      untouched at 130x54, every live/working sheet intact. **DONE 2026-08-19.**
- [x] Applied to the real working file (`register-wide.xlsx`), full backup already
      taken in Phase 0 immediately before. **DONE 2026-08-19** — see commit for the
      verified before/after sheet count. Full original content recoverable from
      `OneDrive\Claude\backups\PRE-MODERNISATION-AUDIT-20260819-1843\
      register-wide.xlsx` indefinitely.
- [x] **`COL_S_` naming collision — RENAMED.** `Sources.bas`'s seven constants
      (`COL_S_ID` etc.) are now `COL_SRC_*`; `FieldSpec.bas`'s eleven (`COL_S_
      FIELDID` etc.) are now `COL_SPEC_*`. Pure rename, every reference across
      `ExcelOutput.bas`, `DraftingUI.bas`, and `tests/TestRunner.bas` updated
      alongside the two declaring modules — confirmed zero `COL_S_` references
      remain anywhere in `vba/`. Full suite 268/268, no behaviour change. **DONE
      2026-08-19.**
- [ ] **Confirm no legacy long-register/`Status`/`ALL`-period structure remains in the
      live workbook.** Grepped the shipped VBA (`DeckRegistry.bas`, `ExcelOutput.bas`) —
      nothing references those constructs anymore, so the code side is clean. The open
      question is whether the live FILE still carries dead sheets/columns from before
      the wide-sheet, roll-forward object model (see `project_deck_sync_object_model`
      memory) that nothing reads anymore and could be archived out.
- [x] **`FIX-LIST` P4 — FIXED.** A fourth Renders-as value (`Slots`) now excludes a
      multi-shape field from the bundled column prompt, same rule as Derived fields —
      full detail in `FIX-LIST.md` item P4. `HIGHLIGHTS_BODY` itself set to `Slots` on
      the real register (its single existing column is pre-existing, not from this
      prompt — Field Spec confirms it was never referenced by injection code, so
      nothing is fed wrong data by it sitting there dormant). Full suite 270/270.
      **DONE 2026-08-19.**
- [x] **`FIX-LIST` P5 — FIXED.** `TemplateAudit.WriteAuditGrid` now carries decisions
      forward by shape identity instead of refusing to rebuild over them — full detail
      in `FIX-LIST.md` item P5. Full suite 269/269. **DONE 2026-08-19.**

### Explicitly OUT of scope tonight
- [ ] Sheet-name migration to `01_FIELD_SPEC`-style numbered names. Already ruled on in
      `NEXT-SESSION.md`: "a migration, not a tidy-up... do it as its own session, suite
      green before and after — not alongside anything else." Don't bundle it in here.

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

## Milestone device — Q and R fixed in source, one thing still owed

- [x] Confirm Q + R actually work on slide 1 of the real rig deck. **DONE
      2026-08-19**, and with genuine register data, not a fixture: found and
      fixed three real defects along the way (FIX-LIST item AY — corrupted
      hand-typed date/label data, a real code bug in `MilestoneDevice.
      WriteText` never converting the `"||"` line-break delimiter, and a
      separate `NumberFormat` corruption on the date columns). Proven live
      twice through the real "2. Put it on the slides" button, unaided,
      verified both times from the saved `.pptx`'s own XML bytes.
- [ ] `InjectDeviceVia` always reports `WouldChange = True`, even when nothing on
      the slide differs from the register. Will pollute the review queue once real
      milestone data exists. Needs a real current-vs-proposed comparison. *Found
      2026-08-16, not previously written anywhere — see `InjectPrimitive.bas`,
      `InjectDeviceVia`, every branch sets `WouldChange = True` unconditionally.*
      **Confirmed still true 2026-08-19** with real milestone data finally in
      the register: every "Put it on the slides" run queues the timeline as
      1 pending change regardless of whether anything actually differs.
      Not blocking (one device = one queue entry, not one per milestone, so
      "pollute" so far means a single always-pending row, not a flood) but
      still a real gap — worth a proper current-vs-proposed comparison
      before this is called fully closed.

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

- [x] **`PROJECT_STATUS` casing — CHECKED tonight against the live register, NOT
      moot.** Ran `FieldSpec.ApplyControlledValidation` for real (its own
      established report-only mechanism — never rewrites, only applies dropdown
      validation and reports drift). **17 values outside the declared vocabulary**
      across 258 controlled cells: mostly casing (`In progress` vs `In Progress`),
      but also a genuinely different, non-vocabulary value (`Not yet commenced`) that
      isn't a casing issue at all. Real side effect: 258 cells now carry a live Excel
      dropdown for the first time, so no NEW drift can accumulate — the 17 existing
      ones were correctly left untouched. Full detail, including which fields
      `DeriveStatusBadge` does and doesn't tolerate: `FIX-LIST.md` item BK. **DONE
      2026-08-19** — quantified and recorded, not fixed (a person's call on whether
      `Not yet commenced` should collapse to `Not Started` or stay distinct).
- [x] **All 17 normalized to the declared vocabulary, real register, verified.**
      Rohan: "let's just have one way of writing it through." 8 `In progress` ->
      `In Progress`, 8 `Not started` -> `Not Started`, 1 `Not yet commenced` ->
      `Not Started`. A real case-sensitivity bug was caught building the fix itself
      (PowerShell hashtable lookups are case-insensitive by default, silently
      rewrote 112 already-correct cells on the first attempt) — full account in
      `FIX-LIST.md` item BK. Proven on a copy, applied to the real register with a
      fresh backup, reverified from the saved file: 0 remaining out-of-vocabulary.
      **DONE 2026-08-19.**
- [ ] Which source is "the dedicated one" for `STRATEGIC_LINKAGES`, and who
      maintains it now the Family Tree is going. A work question, costs a minute.

## File-per-quarter — the prune half (critical path #3) — BUILT 2026-08-21, not yet run live

`ExcelOutput.PrunePeriod` built and unit-tested (FIX-LIST item CA): deletes a
named period's rows from the live register, scoped to exactly that period
(never everything older — the live register held 3 periods at once the night
this was built, and a broader sweep would have silently touched the oldest
one on an unrelated roll-forward). `DraftingUI.RollForwardUI`'s archive step
is now a hard gate — no verified archive, no roll-forward, no prune — per
Rohan's two explicit design calls: prune runs automatically right after a
successful archive, scoped to only the period this roll-forward actually
rolled out of. Confirmed with Rohan directly: this does NOT break a new
quarter's ability to look back at the previous one — `RollForwardPeriod`
already copies every column forward, live, in the same action, before
archive or prune ever run.

- [x] Design + build the prune: drop the old period's rows from the live register
      once it's archived. — **Done.** See FIX-LIST item CA for the full account.
- [ ] Retire `ParkSheetCopy` once the prune lands (load-bearing until then — do
      not delete early). *Source: `DOCUMENT-MAP.md` decision 6, "the live gap this
      exposes."* **Still not done — deliberately.** The prune hasn't run live yet;
      this should not move until it has and is trusted.
- [ ] Sweep `Sync Log` into the same per-quarter archiving. *Source:
      `SCENARIOS.md`'s file-per-quarter section.* **Still not done.** `Sync Log` has
      no `Quarter` column to key a prune on — needs a real design decision, not a
      fragile date-range guess.
- [ ] Tests + one real keyboard run before the prune touches anything live.
      **Half done** — `PrunePeriod` is unit-tested and fail-first proven in
      isolation, but nobody has pressed the real "Roll Forward" button since this
      landed. **The next genuine Roll Forward press against the live
      `register-wide.xlsx` will now also prune** — worth running once, deliberately,
      on a copy first.

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
- [x] **#4, CLOSED 2026-08-16 evening.** The full arc: `SaveAs`-to-self fix
      (source), rebuild+re-prove surfaced a deeper limit (only a session's
      FIRST `CustomDocumentProperties` write ever lands on a cloud deck),
      three rescue attempts failed (close+reopen, even with a 15s wait),
      real fix built — moved `DeckSyncPeriod`/`DeckSyncWorkbookPath`/
      `DeckSyncType`/`DeckSyncTemplate`/`DeckSyncId` off
      `CustomDocumentProperties` onto a dedicated hidden slide, keyed by
      shape name, with a read-fallback to the old location so existing
      decks keep working. **Proven on the real add-in (`addin109`): 8 for 8
      repeated writes landed on ONE reused open cloud file** — the exact
      scenario that failed 0/8 before — independently cross-checked via a
      separately-written .NET zip reader. Found and fixed a real `Variant`
      vs `String` bug in the new code along the way (same class already
      documented once in this file — checked every `Namespace()` call in
      the repo, not just the one that broke). Also found and corrected a
      false negative in the TEST SCRIPT itself: a genuine VBA `""` success
      return marshals as PowerShell `$null` through `Application.Run`,
      which an early "8/8 failed" report turned out to be — caught only by
      checking the saved file's actual bytes, not the return value.
      Static checks clean, suite 230/0. Full evidence: `FIX-LIST.md` item
      P's final 2026-08-16 update. This was the last blocker on Scenario 1
      (updating the period on an existing, already-synced deck every
      quarter).
- [x] **#4b, CLOSED 2026-08-16 evening.** Prompted by "check the register
      too" — the register had the same class of defect as #4, on a
      different application: `ExcelOutput.WriteDeckReference` used
      `Workbook.CustomDocumentProperties`, and a probe found a NEW property
      lands fine (even a second one, same session — narrower than the
      deck's version) but RE-WRITING an existing one never persists, via
      `.Value=` (the real function's actual pattern) or Delete+Add. Called
      on every repoint via `StampPairing`, so a workbook re-paired to a
      different deck after its first stamp would silently keep reporting
      the OLD deck's identity forever. Fixed the same shape as #4: moved
      `DeckReference` onto a cell on a dedicated very-hidden `DeckSyncMeta`
      sheet, with a read-fallback to the old location. **Proven against the
      real re-pairing scenario across two genuinely separate sessions**
      (stamp, close, reopen, confirm; re-stamp with a different value,
      close, reopen, confirm the NEW value landed, not stuck on the old
      one). Static checks clean, suite 230/0. Full evidence: `FIX-LIST.md`
      item S.
- [ ] **#5:** `Tag fields on this slide` run fresh against `K900` and `S900`
      to get CURRENT field coverage ground truth, replacing the
      cross-referenced-from-tags inference in the Field Coverage Matrix
      below with a direct discovery pass. Would also settle the open
      `PROJECT_PROGRESS` bar-vs-text routing question.
- [ ] Add more here as they're identified — this list is the record, not a
      one-off.

## Field coverage — real gaps found while checking, 2026-08-16

Rohan asked whether Field Spec is the right place to hold a per-template target
field list (yes) and flagged that pictures didn't carry across to the new K/S
templates. Built a **Field Coverage Matrix** sheet (flanking `Field Spec` in
`scenario3-template-surgery-20260816\register-wide.xlsx`) cross-referencing
Field Spec's 48 FieldIDs against the REAL `ROLE` tags read directly from the
saved P/K/S template slides — not the stale `Template Audit`/`Field Discovery`
sheets, which predate today and were checked but not trusted as current.

- [x] **Confirmed, from real tags, not guessed:** P/K/S all carry the SAME 15
      real fields (`ABOUT_BODY`, `KEY_EVENTS_BODY`, `PROJECT_STATUS`,
      `STRATEGIC_ALIGNMENT_BODY`, `PROBLEM_BODY`, `PROGRESS_BODY`,
      `PROJECT_CODE`, `PROJECT_NAME`, `SUBTITLE_A`, `START_DATE`, `END_DATE`,
      `PROJECT_PROGRESS`, `INDUSTRY_CASH`, `TOTAL_VALUE`, `PROJECT_LEAD`) —
      cloning preserved tagging correctly for everything that WAS tagged.
- [x] **Real defect found and root-caused: this was never a K/S-vs-P
      gap.** Scanned all 46 real slides across every letter — **zero** carry
      a `MILESTONE_TIMELINE` role tag. Only the P **template** has it. The
      real source slides K/S were cloned from (`1_K1001`, `1_S001`) never had
      it either, so `MakeTemplateFrom` faithfully copied their actual state —
      not a cloning bug. The prototype, `3_P001`, has a fully-built,
      correctly-**named** `MILESTONE_TIMELINE` group (all `MS1..MS7` parts
      present) with **zero tags on it** — confirmed from the raw XML
      (`<p:cNvPr id="17" name="MILESTONE_TIMELINE">`, no `r:id`, no tags
      relationship at all). **Fixed 2026-08-16**, see below.
- [x] **Confirmed, from Field Spec's own text, not inferred:** the three
      deliverable thumbnail picture cards have "no Field Spec row yet" —
      untagged by design-so-far, so `MakeTemplateFrom`/sync cannot touch
      them. This is why pictures didn't come across — a pre-existing,
      already-named gap that today's cloning made visible, not a regression.
- [x] **RESOLVED 2026-08-19: "three" was undercounting.** All 45 slides in
      the real deck checked directly — the count is 0-4+ per project, not a
      fixed three; `3_P001` itself (the prototype this note is about) has
      **four**. Built as `DELIVERABLE1_PHOTO..DELIVERABLE4_PHOTO` (fixed
      max slots, matching `3_P001`'s own real count), tagged, wired
      (register columns, Sources rows, Field Spec rows), and **proven live
      end to end** on a copy of the real deck/register — see FIX-LIST items
      BA-BC. Two of the four turned out to be SVG-inserted shapes
      (`shp.Type = msoGraphic`, not `msoPicture`), which `IsPictureShape`
      didn't recognise at all -- a second real defect found doing this,
      also fixed (BC). **Not yet applied to the real deck/register** — the
      copy this was proven against is
      `C:\Users\rohan\deck-sync-test-deliverables\`, not the live files.
- [ ] **Not yet resolved:** whether `PROJECT_PROGRESS` actually routes
      through the bar injector (`InjectProgressVia`) or the plain-text one
      on these templates. `InjectorFor()` auto-detects this from the shape's
      own structure at sync time, not from Field Spec's "Renders as" column
      — so the stale audit calling the shape "text" doesn't settle it either
      way. Needs a fresh check, not an assumption in either direction.
- [ ] Fresh `Tag fields on this slide` runs against `K900` and `S900`
      (real test #5, tracked above) would give current ground truth on all
      of this instead of the cross-referenced-from-tags inference used here.

## Field Discovery cross-slide bug, and a repo-wide Clear audit — 2026-08-16

Real test #5 (running `Tag fields on this slide` against K900) surfaced this.
`Field Discovery`'s "existing marks" for two K900 shapes turned out to be
**fabricated** — a shape whose real text is "Industry" was reported as
tagged `PROJECT_CODE`; a section-header label was reported as
`PROJECT_PROGRESS`. Neither shape carries a tag at all (confirmed from the
raw XML — no `r:id`, no tags relationship).

- [x] **Root cause found: `DiscoverUI.ExistingRowsById` keys purely on the
      bare shape `id`.** PowerPoint restarts shape-ID numbering on every
      slide, so it isn't unique across the deck. Running "Tag fields on this
      slide" against a NEW slide whose shapes happen to share ID numbers
      with a PRIOR run's marked shapes silently attaches the old marks to
      the new, unrelated shapes.
- [x] **First fix attempt would have introduced a NEW silent-loss bug —
      caught by Rohan, not by me, before it shipped.** Naively clearing the
      sheet on every slide switch would have discarded any pending,
      un-applied marks from whatever slide was analysed previously, with no
      warning — trading the cross-slide misattribution bug for exactly the
      "why is clear needed there? It is not" class of loss DOCUMENT-MAP
      decision 1 already eliminated once, for the drafting sheets.
- [x] **Real fix:** a slide-identity marker (the slide's own stable
      `SlideID`, stored off to the side) decides whether `rowById` carries
      forward at all. Different slide + pending marks still on the sheet →
      **refuse** ("!..." message, same convention as every other refusal in
      this codebase), naming the count, telling the person to apply or
      clear first. Different slide + nothing pending → proceeds, but with a
      **scoped tail-clear** (only the specific leftover rows beyond the new
      slide's own content), not a blanket `ws.Cells.Clear` — `ws.Cells.Clear`
      is now called ZERO times in `DiscoverUI.bas`, down from one
      unconditional call.
- [x] **Repo-wide audit of every `.Clear`/`.ClearContents` in production
      code, per Rohan's direct instruction** ("verify every use of it and
      justify its existence... everywhere in the repo"). `Err.Clear`
      excluded (a different thing — the VBA error object, not data).
      Checked and confirmed already-justified: `Drafting.bas` (layout
      migration guarded by `Not layoutMatches` + unconditional backup; the
      ratified quarter-rollover ferry, decision 6; scoped intro-row stamps),
      `Readiness.bas`/`WorkbookBridge.bas` (pure generated reports, nothing
      ever typed into them), `ReviewQueue.bas` (ratified R13.5 — approvals
      are deliberately single-use per build, not carried; flagged as the
      SAME risk shape as this bug, by design rather than by accident, worth
      Rohan knowing).
- [x] **One real match found and fixed the same way:** `TemplateAudit.bas`'s
      `WriteAuditGrid` — its own comment already admitted *"a re-run
      discards decisions typed into the last column"* for a worklist
      explicitly meant to be filled in *"across several sittings,"* and the
      only protection was a warning shown to the human **after** the sheet
      had already been cleared. Now a `Function` returning `""`/`"!..."`
      (was a `Sub`), refuses when the sheet has recorded field/chrome/drop
      decisions not yet acted on. `RibbonUI.AuditFieldsCore` updated to
      check the result instead of assuming success. `SummaryText`'s own
      wording corrected to match (was describing the old, dangerous
      behaviour).
- [x] **6 new/updated tests**, all passing: `DiscoverUI`'s cross-slide
      refusal and its own safety rail, `TemplateAudit`'s refusal proving the
      pending decision survives untouched and the second write never
      happens. Suite 229 → (pending full run — Office apps still open,
      static checks clean across 35 modules).

## Milestone device reachability — the real fix, 2026-08-16

**The bug:** `WalkForDeviceRoleTags` (`InjectPrimitive.bas`, the FIX-LIST R
reachability fix from 2026-08-15) required the timeline group to carry a
`role` **tag** before `PlanRoutineSync` would ever attempt to sync it. But
this device's own convention — stated by Rohan, 2026-08-10, and baked into
`MilestoneDevice.SlotCount`/`PartsOf`/`DeviceIntegrity` — is that its parts
are addressed by **name**, not tag: *"they are simply addressed by name...
this is what stands in for the robustness a tag would have given."* The
group on every real slide is correctly *named* `MILESTONE_TIMELINE`. It was
never *tagged* that way, because tagging it was never the convention. The
reachability fix contradicted the device's own design and made it
unreachable everywhere, including the one slide (`3_P001`) it was actually
built and tested on.

- [x] **Fixed.** `WalkForDeviceRoleTags` now falls back to the group's own
      `.Name` when no tag is present, before adding it as a discovered device
      identity. `ShapeHasRoleTag` (used by `FindShapeByRoleTag`, which
      `InjectorFor` and the actual injection call both re-run) got the same
      fallback, scoped **strictly** to a structurally-confirmed device
      (`msoGroup` + `MilestoneDevice.SlotCount > 0` + no tag already present)
      — an explicit tag still wins when one exists, and an ordinary
      text/bar/picture field that happens to share a shape's name is
      untouched, proven by a dedicated negative-control test.
- [x] **5 new tests, all passing, plus the existing tagged-device test
      re-confirmed unchanged:** discovery-by-name, `FindShapeByRoleTag`-by-
      name, `InjectorFor` routing an untagged device correctly, the safety
      rail (untagged textbox/non-device group NOT matched by name), and tag
      still winning over name when both exist. Suite 224 → 229/0.
- [x] **Proven against the real file, 2026-08-16 14:00.** `addin107` built,
      Rohan pressed "Review changes (writes nothing)" on the copy against
      `3_P001`. Verified from the SAVED workbook, not the dialog (the macro
      writes the review sheet to memory but never saves it — caught a
      real save-timing gap doing this check, asked Rohan to save Excel
      before re-reading): `Review project-progress-A32C`, run stamp
      `2026-08-16 14:00:02`, row 4 —
      `3_P001 | MILESTONE_TIMELINE | 7 slot(s), all parts present |
      (redrawn from its register columns)`. That row could not have existed
      before this fix; the device was completely invisible to
      `PlanRoutineSync`. **Unrelated, pre-existing defect noticed in the same
      sheet, flagged to Rohan, not touched:** row 3 is the already-known
      `PROJECT_PROGRESS` false-diff bug (`80%` → `0.8`) — do not approve it.
- [ ] **Separate, later decision, NOT this fix's job:** whether/how to
      backfill the milestone device onto the 45 real slides that don't have
      the shapes at all (only `3_P001` was ever built out) — this fix makes
      an EXISTING, correctly-built device reachable; it cannot invent circles
      that were never drawn. Content work, not sync work.

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

### Field propagation — CLOSED 2026-08-21, after being written wrong twice on 2026-08-20

**History, kept because the lesson is worth more than a clean page.** This section went
through three versions in one night before the actual fix landed. v1 called 46/47
"dormant" from the code's fallback logic, unchecked. v2 corrected that but kept the frame
that this was a 46/47 problem — a 4-slide sample looked consistent with that frame and
nobody widened it before writing it down. v3 ran a full 43-slide census and found the real
shape: 42 of 43 real project slides were missing fields, not two templates. The standing
lesson from that: **never state what's live on the real deck from code behaviour, a doc,
or a small sample — read the actual file, at the actual scale of the claim.**

**Now genuinely closed, verified from the saved file's own bytes, not from a script's own
report.** Every field v3 found missing has been tagged/built and propagated to all 43 real
slides plus the P/K/S templates:

- `SAAFE_CASH`, `TOTAL_INKIND`, `INDUSTRY_PARTNER`, `TERTIARY_INSTITUTION`,
  `DELIVERABLES_BODY`, `START_DATE`, `END_DATE` — tag-only gaps, closed by geometry/content
  matching against an already-tagged reference, same discipline each time (refusal-guarded,
  verified against saved bytes, never trusted from the writing session). **Correction, same
  day**: this line originally claimed "plus P/K/S templates" for all seven — wrong for four
  of them. That night's propagation batch targeted the 43 real slides only; `SAAFE_CASH`,
  `INDUSTRY_PARTNER`, `TERTIARY_INSTITUTION` and `TIMELINE_ELAPSED` were never actually
  checked on the K/S templates themselves, and K's `INDUSTRY_PARTNER` shape held a
  hardcoded literal `"[TBC]"` string rather than a placeholder tag — unreachable by any
  injector. Found by a later kennel pass and fixed (FIX-LIST item BQ). The lesson this
  section already states applies to itself here: a claim of "closed... plus templates" is
  exactly the kind of small-sample generalisation the rest of this section exists to warn
  against, and it slipped through anyway.
- `MILESTONE_TIMELINE` and `PROJECT_PHOTO` — these needed actual shape creation, not just
  tagging (genuinely absent from 41 real slides and from the K/S templates). Built on the
  P/K/S templates first (colour-corrected per type — the milestone device's internal
  colours turned out to be universal device-state formatting, not P-branding, a wrong
  assumption caught and reverted before it shipped), then cloned from the now-correct
  templates onto all 43 real slides via `Shape.Copy`/`Shapes.Paste` (tags survive the
  copy, confirmed by probe on a scratch file first).
- `SECTOR`, `TRL`, `SUBTITLE_B` — never had their own shape tags at all; they're inputs to
  the `SUBTITLE_A` composite (`SyncOperations.ComposeSubtitleLine`), not separately
  rendered. Turned out `SUBTITLE_A` already held real per-project data for 29 of 43
  projects, hand-typed as one 4-part blob before these three columns existed — split into
  its proper four columns rather than overwritten, preserving the real data. The other 14
  (`[TBC]` on both halves) got best-effort inferred values, explicitly flagged as such.
- `PROJECT_STATUS`/`SCHEDULE_STATUS` — also never need their own tag; they feed
  `STATUS_BADGE` as inputs. `PROJECT_STATUS` was already 100% filled in the register (a
  false-positive gap). `SCHEDULE_STATUS` genuinely can't be filled by inference — the Field
  Spec is explicit that it's read from an external Milestone & Deliverable Tracker
  document's own formula, not selected — left blank pending that source, with one
  clearly-flagged test value (`3_P001` = "At Risk") written to prove the
  `SCHEDULE_STATUS` → `STATUS_BADGE` link actually works end to end.

**A second, unrelated leaked-content class turned up along the way and got fixed too**:
real per-project text (milestone dates, a partner name, dollar figures, two named
individuals) sitting in untagged shapes on templates and on 15 real K-project slides —
found by scanning broadly rather than trusting one earlier "confirmed clean" check, fixed
with the same content-verified-before-delete discipline throughout. One of those fixes
was wrong on the first attempt (deleted real per-project data mistaking a shared
milestone-phase name for duplication) and was caught and restored from backup before it
shipped — logged here because the near-miss is as instructive as the fix.

**Genuinely still open, by design, not oversight:** `HIGHLIGHTS_BODY` and
`STRATEGIC_LINKAGES` have no shape anywhere in the deck — not the exemplar, not any
template. Proving they're linked means inventing shape geometry from scratch first, which
needs Rohan's own placement/design call, not a blind clone. Deliberately deferred.

- [ ] Bring up a genuinely fresh deck + fresh register from nothing, unaided —
      the standing requirement, the tool has to travel with Rohan.

## SECTOR/TRL inferred rather than sourced on 14 projects, 2026-08-20/21 — needs verification, not a defect in the mechanism

`SECTOR`/`TRL` are `Kind=Given` — the Field Spec's own text says "do not infer or fill
the gap." Rohan explicitly authorised an exception for this one pass ("happy for you to
take a stab at values based on what you can see"), but the register carries no flag
distinguishing which values came from real source data and which were inferred from
`ABOUT_BODY` prose. Recorded here rather than in the cell itself — a marker in the value
would render on the actual slide subtitle line, which is worse than not flagging it at
all.

**29 of 43 projects had real data already** (hand-typed into the old `SUBTITLE_A` blob
before it was split into columns) and were preserved, not inferred. **These 14 were not**
— genuinely Claude-inferred from project descriptions, not sourced, and should be treated
as a first draft needing Rohan's (or the real source's) confirmation before they're relied
on: `1_K1001, 1_K1002, 1_K1003, 1_K1004, 1_K1005, 1_K1006, 1_K1007, 1_K1008, 1_K010,
3_K016, 4_K017, 4_K021, 1_K022, 3_K023`.

## Bold header line for PROGRESS_BODY/KEY_EVENTS_BODY — CLOSED 2026-08-21

Rohan's original ask, from earlier the same night: a dedicated bold header line
(quarter on the left, project status on the right when not "In Progress") above
each section body, replacing the old stale first-run-formatting inheritance that
made the whole body render bold. Design fully confirmed with him: new shapes, not
a formatting-reset fix.

**Built and confirmed on the P/K/S templates**: `PROGRESS_HEADER`/`KEY_EVENTS_HEADER`
shapes exist, correctly positioned (body shapes shrunk ~12pt and shifted down to make
room, bottom edge held fixed), bold, Calibri 8pt, coloured per type (P `#003C23`,
K `#F55A2D`, S `#C0A2F2` — matching the same per-type convention as the milestone
widget). The pre-existing body-bold bug is fixed alongside it (`Font.Bold = 0` on
`PROGRESS_BODY`/`KEY_EVENTS_BODY`).

**Update, 2026-08-21 — both fields real, per Rohan's own clarification.**
He rejected the computed-quarter design this section originally called for:
*"progres sheader can just be 'Last reported quarter Q4F26' ... separate
field thats either manually adjusted in rare nonexpected cases or for
student projects takes a frozen quarter label when they report every six
months."* So `PROGRESS_HEADER` is `Kind=Given` (plain, directly-editable
register column, same pattern as `SECTOR`/`TRL`/`SUBTITLE_B`), not a
computed derivation. `KEY_EVENTS_HEADER` IS a real `Kind=Derived` field
(`SyncOperations.DeriveKeyEventsHeader`: blank when `PROJECT_STATUS = "In
Progress"`, otherwise shows it verbatim), wired into `DerivedFieldTags()`/
`ComputeDerivedValue`, fail-first proven, full suite 286/286.

**First default was wrong twice, both caught live by Rohan (FIX-LIST item
BS)**: defaulted to `Q1F27` (computed from today's date) instead of the
`Q4F26` he'd written directly in his own instruction — *"why Q1F27?!!"* —
then, after fixing that, applied `Q4F26` uniformly to all 43 rows
including the S-project (student, six-monthly) rows he'd already told
me report on a different cadence — *"hang on the s projects didnt report
in q4 like I explained."* Rohan confirmed `Q3F26` is their actual
last-reported quarter (FIX-LIST item BU) — but that fix also missed 4 of
the 17 S-rows (`S009`/`S021`/`S022`/`S023`, the bare-key instance IDs
with no leading `N_` prefix), caught by the same mother-hound kennel run
that found BY below (FIX-LIST item BX). Final state, verified from the
saved file: `"Last reported quarter Q4F26"` on the 30 P/K rows, `"Last
reported quarter Q3F26"` on all 17 S-project rows.

**Propagated to all 43 real slides (FIX-LIST item BT)**: both header
shapes cloned from the matching P/K/S template onto every real slide,
`PROGRESS_BODY`/`KEY_EVENTS_BODY` shrunk by the same 12pt and un-bolded.
Verified by parsing the saved `.pptx`'s own XML directly: 46/46 role tags
for each header (43 real + 3 template), correct per-type colour on a
sample P/K/S real slide each, 0 remaining bold runs in either body field.

**Duplicate prompt instruction removed, 2026-08-21.** `PROGRESS_BODY`'s
Purpose/Length/Behaviour no longer say to open with a quarter-labelled
header; `KEY_EVENTS_BODY`'s Behaviour no longer says to open with a bold
status line. Both now point at the dedicated header field by name and
say not to repeat it. Verified from the saved file.

- [ ] Verify or correct the 14 inferred SECTOR/TRL values above against real source data.
- [ ] Decide whether `SECTOR`/`TRL` need a lightweight provenance flag of their own —
      `PROVENANCE.md`'s design is scoped to prose fields going through the real publish
      pipeline (`Drafting.PublishDrafts`) and wouldn't have caught this even if built,
      since these were hand-written directly to a `Given` field outside that path entirely.

## Milestone data (MS1-7) — CLOSED 2026-08-21, was a broken CARRY not a fresh migration

Rohan: *"import all milestone data."* Checked the Field Spec first rather than
guess a design: `MS1-7`'s five middle labels are `Kind=Prose`, `Cadence=Standing`,
`History treatment=CARRY` — drafted once per project from `SRC_MILESTONES`, then
carried forward unless the milestone plan itself changes. 38 projects already had
this (verified, correctly compressed from the raw tracker data) sitting on
**Q3F26**; **Q4F26 had 0 of 543 non-blank cells carried across**. Not a fresh
`SRC_MILESTONES → register` migration — the known "still open" gap from earlier
in the session was this carry, already broken.

**Carried 543 cells for 38 projects**, Q3F26 → Q4F26, verified (0 mismatches,
spot-checked). See FIX-LIST item BV for the full account, including a crashed
first attempt (numeric `MS*_DATE` cells vs. a `.Value2` type-cast COM error,
caught and fixed before any partial write landed).

**5 remaining projects had no MS data anywhere.** `P008`/`S023` have no source
data (`SRC_MILESTONES` rows or `START_DATE`/`END_DATE`) to import — left blank.
`2_P009`/`1_P010`/`2_P012` have clean 4-5-row source data, drafted directly
(labels compressed to ~4 words from the real tracker milestone names, dates
set, MS1/MS7 from `START_DATE`/`END_DATE`) — but `MS_DONE` was NOT auto-filled:
checked first whether it mechanically mirrors the tracker's completion %, found
it doesn't (`3_P001`'s own existing data has 100%-complete tracker milestones
marked not-done on the slide, proving the flag reflects Rohan's own review),
asked rather than guess. His answer: leave every `MS_DONE` blank on all 3,
he sets them by hand.

- [ ] `P008` and `S023` still have no milestone data. `S023` additionally needs
      a real selection judgement (10 raw milestones, no due dates, down to 5
      circle slots) before it can carry any — not something to auto-pick.
- [ ] `MS_DONE` blank on `2_P009`/`1_P010`/`2_P012` (7-14 cells) — Rohan to set.

## mother-hound kennel, 2026-08-21 — CLOSED, both findings fixed

Ran after the header propagation and milestone-carry work above. Two real
defects, both fixed same day — see FIX-LIST items BX and BY:

- **BX**: BU's "13/13 correct" S-project quarter fix was wrong — 17
  S-rows exist, not 13; the 4 bare-key ones (`S009`/`S021`/`S022`/`S023`)
  still held the P/K default. Same bare-key bug BT already found and fixed
  in a different script, recurring in this one. Fixed, verified 17/17.
- **BY**: `DeckAdoption.PlanAdoption` (the bulk "Adopt Existing Slides"
  button) had the identical `SUBTITLE_A` composite-field blind spot
  `Harvest.bas`/`E2EField.bas` were fixed for earlier tonight — a third
  live call site, reachable via a normal workflow (duplicate a slide,
  adopt it). Fixed with the same shape, new test, fail-first proven.

Also confirmed clean by the kennel, independently re-derived: the 543-cell
MS carry (full tuple match, not just counts), Field Spec ↔ Register column
parity, controlled vocabulary, cross-row prose duplication, full field-tag
parity across every real slide vs. its template, the milestone-widget
colour fix standing on all 46 widgets, and the leaked-sentence/bold-run
fixes holding uniformly across all 43 real slides.

**Confirmed as expected, not a defect**: `PROGRESS_HEADER`/
`KEY_EVENTS_HEADER` are tagged on all 43 real slides but still show
literal `<<...>>` placeholder text — tag propagation happened, but no
real sync has run to inject actual values yet. Needs a real sync before
either field is genuinely live.

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

## Test-suite viewing aid — "Watch", 2026-08-16

Rohan: "make key bits obvious to a human viewer while still obviously
completing," then, once built, "use this viewing aid pause concept with me
for the rest of our existence together." Standing practice now — see memory
`feedback_viewing_aid_pauses_for_observable_effects`.

- [x] `TestRunner.bas`'s `Watch(sld, label, [seconds])` helper: places a big
      bold orange on-slide banner naming the effect, pauses ~1.2s
      (`Timer`-based busy-wait — `Application.Wait` is Excel-only and does
      not exist on `PowerPoint.Application`; that mismatch compiled-error
      the whole project until fixed), removes the banner, resumes. Wired
      into 2 shortlisted tests so far: `MilestoneDevice_DrawsFromDataAndCreatesNothing`
      (ON/OFF/NOW visibility) and `InjectProgress_MeasuresAgainstTheTrackNotItself`
      (bar resize). "Shape colour change" was NOT added — checked first:
      nothing in production code writes `Fill.ForeColor` at runtime yet, so
      there is no real effect to shortlist there.
- [x] **Real lesson, not just a feature build:** the first two implementation
      attempts tried to force the PowerPoint WINDOW to the OS foreground
      (`AppActivate`, then `SendKeys`+`AppActivate`). Both failed for real,
      live reasons — `AppActivate` alone only flashed the taskbar (Windows'
      focus-stealing prevention), and adding `SendKeys` to work around that
      HUNG the run for 2 minutes under this PowerShell→COM→VBA driver.
      Capped at 2 attempts per standing practice, then stopped rather than
      trying a third variant. **Rohan's own fix was simpler and correct:**
      label the effect on the artifact itself, stop fighting window focus
      entirely. `AppActivate` is kept as a harmless best-effort call but is
      no longer load-bearing for "obvious."
- [ ] **Found while checking this live, not yet cleaned up:** the reused
      `deck_sync_test_run_presentation` fixture file has accumulated **133
      slides**, mostly orphaned blanks — each filtered test run this session
      adds throwaway slides via `NewBlankSlide()` and deletes them at the
      end, but a run that gets cut off (like the `SendKeys` hang above)
      never reaches its own cleanup, leaving slides behind permanently in a
      file that persists across runs. Low urgency (throwaway test fixture,
      not real content) but real — worth either deleting that file so a
      fresh one gets created, or adding startup cleanup to
      `run_vba_tests.ps1` for a stale/oversized fixture.
- [x] **Full suite re-run against the final banner-based `Watch()`, 2026-08-16:
      230 passed, 0 failed, 0 skipped.** Confirms the on-slide-banner
      implementation (after the two failed foreground-focus attempts above)
      introduced no regressions anywhere in the suite, not just in the 2
      tests it's wired into.

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

## The milestone-family drafting sheets could share one sheet — superseded by The Lobby, above

Rohan's original question, live: *"can't complex shape groups be drafted on one sheet
together? Like the timeline elements?"* — worked through the same night into the Lobby
design above, which solves the actual complaint (the crawl, the two-loop pattern)
without the sheet-merge's risk. See "The Lobby" section, and "Milestone-sheet-merge —
superseded" immediately under it, for the full reasoning and why sheet-merging was ruled
out as the primary fix.

## The Lobby — full design in `LOBBY-DESIGN.md`, PRIORITY, phases 0-3 done and deployed

**Full architecture, reasoning, and the pre-ticked/opt-out approval rule now live in
`LOBBY-DESIGN.md`** — read that, not this, for the design. This entry is just the
tickable build status so it isn't duplicated in two places (this project's own rule:
a machine-knowable fact lives once).

- [x] **Phase 0 — cold-start crawl + core mechanics.** `vba/DraftingLobby.bas`
      (`PinToLobby`, `ReadLobby`, `ClearLobbyEntry`, `LobbyCount`,
      `BuildLobbyFromScratch`). Three real tests, full suite green (233/0). Built and
      tested 2026-08-16 (night). Two real bugs found and fixed building it — see
      `LOBBY-DESIGN.md`'s status banner and `AGENTS.md`'s Known Patterns.
- [x] **Phase 1 — the `Application.SheetChange` pin-on-tick event mechanism.**
      `vba/AppEvents.cls` (first class module in the codebase). Proven LIVE
      2026-08-17, both directions (tick pins, un-tick clears), zero direct
      `PinToLobby` calls in the test — a real cell write via COM caused a real,
      correct pin. Two more real bugs found and fixed, both logged as classes in
      `AGENTS.md` — see `LOBBY-DESIGN.md`'s status banner for detail.
- [x] **Phase 2 — `PublishAllDraftedFields` now reads the Lobby instead of crawling
      the 13 sheets directly.** `DraftingUI.DistinctPinnedFields` (pure, tested --
      proven to catch a deliberately-broken dedupe before the real fix was restored).
      Proven LIVE 2026-08-17 on `PRESERVED-known-good-20260815-1050`: with 39 rows
      pinned across two fields (`ABOUT_BODY` x1, `PROGRESS_BODY` x38), "2. Put it on
      the slides" ran Copy+Publish for exactly those two fields and nothing else --
      confirmed from the saved workbook's own `Drafting Lobby` sheet, not just the
      dialog. Full suite green (234/0). Safety valve: `RefreshDraftingSheets`
      ("1. Set up my quarter") now silently repairs the Lobby from ground truth every
      run, at no extra cost (it already reads every row of every sheet) -- closes the
      at-work hand-edit gap without a third button.

      **Found 2026-08-19, not yet fixed:** `DraftingLobby.ClearLobbyEntry` is tested
      but has zero production callers -- its own comment says it should run "after
      `PublishDraftsForField` has actually published" a pinned entry, and the LIVE
      publish path (`DraftingUI.PublishAllDraftedFields`) never calls it. Practical
      effect: a Lobby entry is never cleared after its field is actually published,
      so it keeps looking like pending work indefinitely. Phase 0/2 above are still
      correctly marked done for what they DO cover (pin, read, publish-from-Lobby);
      this is a gap in what happens AFTER a successful publish, not in those.
- [x] **Phase 3 — pre-ticked queue items + remove the Yes/No/Cancel apply gate.**
      Built, tested (236/0), held back overnight, reviewed by Rohan the next morning
      (asked real questions about the layered no-overwrite mechanism and a proposed
      "register as memory" shortcut, correctly rejected), approved, deployed
      (`addin119`, confirmed loaded live). See `LOBBY-DESIGN.md` section 5 (the rule)
      and its own Phase 3 status block (the full review record).
- [ ] Phase 4 — revisit sheet-merging, only if the Lobby alone doesn't fully address
      the sheet-count/tab-clutter complaint once 0-3 are proven live.

## Milestone-sheet-merge — superseded by the Lobby above, kept for the reasoning only

Original idea (still true, just no longer the recommended path): combine the 5
`MS2_LABEL`..`MS6_LABEL` drafting sheets into one. Ruled out as the primary fix once the
Lobby design landed — merging sheets means reworking `Drafting.WriteDraftingSheet`'s
fixed row-addressing (`DRAFT_FIRST_ROW=10` etc.), the single most incident-prone
function in the codebase (five real data-loss bugs, 1–14 Aug). The Lobby solves the same
underlying complaint (the crawl, the two-loop pattern) without touching that function at
all. Revisit sheet-merging only if the Lobby alone doesn't fully address the sheet-count
friction.

## Dialog count across one full cycle — CLOSED 2026-08-21, was already fixed

Flagged 2026-08-16 (night) as PRIORITY, after Rohan's live *"too many message boxes
and confirmations across that chain and excel takes ages"* and roughly 10-12 dialogs
counted across one full cycle. **Checked the actual chain code before doing any work
(FIX-LIST item BZ) — every candidate this item named was already fixed, across four
separate sessions since this was written, and the doc was simply never updated:**

- The two-press build-then-apply requirement Rohan explicitly called out — collapsed
  2026-08-18. One press builds, asks once with the real pre-ticked list in hand,
  applies.
- The unsaved-workbook guard — made silent 2026-08-19
  (`WorkbookBridge.EnsureSavedQuietly`), only interrupts on a real save failure.
- `StartQuarter`/`RollForwardUI`/`RefreshDraftingSheets`/field-coverage — folded into
  one combined report dialog, shown only when there's something to report.
- Three separate redundant modals deleted outright (2026-08-14, 2026-08-17).

A steady-state cycle today is roughly 2-4 dialogs, not 10-12. Nothing built here —
there was nothing left to build.

## Explicitly out of scope — from `TRACKER.md`, do not re-add without a reason

The other ~30 unwired fields. A GUID-based key redesign. R13's full review
subsystem (built, parked). Chrome/UI enforcement. Ribbon polish. Adding these
here would be the "field count as progress" trap the project already fell into
once. *Source: `TRACKER.md`, "Not on this list, deliberately."*

## Gap analysis, 2026-08-22 — current deck vs. last hand-built version
(`2026-08-13-0956-pre-onboard`, 43 slides matched 1:1 by content similarity,
order-preserving). Full detail in `FIX-LIST.md` CK/CL/CM.

- [~] **CK — milestone timeline doesn't reflect project closure, 7 real
      candidates (of 8 closed projects; P008 has no milestones to mark).**
      Register data gap, not code. Fable independently confirmed.
      `SRC_MILESTONES` (the real CRC tracker extract, per Rohan's own
      question) gave much better evidence than narrative prose. FIXED and
      verified: `1_P010` (all 5), `3_K016` (`MS5`/`MS6`), `2_P009`
      (`MS2`/`MS3`/`MS6`). Confirmed correct as-is, no change: `3_P001`
      (closed early), `1_K022` (final report explicitly still pending
      sign-off). Still open, genuinely ambiguous, cells identified for
      Rohan: `1_K1004` (row 59), `1_K1008` (row 63) -- tracker milestone
      numbering doesn't map cleanly to register slots for either. Also
      noted, not acted on: `2_P009`'s `MS1` is blank unlike every other
      project here -- outside what was proposed/approved, flagged not
      fixed.
- [x] **CN — `MS*_DONE` drifts from the real CRC tracker with nothing to
      catch it, found via `2_P012` showing every milestone as "not
      achieved" despite 80% progress.** Root-caused by Fable: never
      code-linked to `SRC_MILESTONES` (deliberately manual, per BV), and a
      naive gap scan overcounts because the tracker is far more granular
      than the register's 7 slots (`2_P004` needs no fix once grouped
      properly). Built `vba/tools/MilestoneEvidenceReport.bas` +
      `milestone_evidence_report.ps1` (read-only, reports only, matches
      `VerifyRealDeck`'s pattern) to do the grouping properly and surface
      disagreements with the actual comment text. **Now run and working**
      -- found and fixed 4 real bugs in the tool itself first: a bulk-read
      perf fix, per-project error trapping, the actual crash (VBA's `Or`
      doesn't short-circuit, so `bestSlot = 0 Or slotOffset(...) <
      slotOffset(bestSlot)` read `slotOffset(0)` unconditionally against a
      `1 To 7` array -- crashed on every project until fixed), and a
      silent-accumulation bug (loop-scoped `Dim`s never reset between
      projects, so each project's printed detail and "N of M" counts kept
      compounding forward from every project before it -- caught by
      actually reading the report instead of trusting "0 errors"). Swept
      the rest of `vba/` for the same shape (Rohan's own question) --
      isolated to this tool. Clean live run: 41 projects checked, 14
      real disagreements (the earlier "26" was the corrupted count).
      **All 14 worked through with Rohan and resolved -- see FIX-LIST
      CN for the full evidence per project.** 8 slots corrected `Y`->blank
      (register overstated), 3 slots corrected blank->`Y` (register
      understated), 5 slots deliberately left alone (no clean evidence
      either way, or the comment text confirmed the register was already
      right despite a stale 0%). Backed up, written in one pass, verified
      from saved bytes. **Register only -- needs a sync run to reach the
      deck.** Prevention proposed, not actioned: hook into
      `RollForwardUI` rather than a separate tool to remember.
- [ ] **CL — content depth drops sharply for S007 onward (13 slides).**
      Rest of the deck runs 80-100% of original text length; `S007`-`S023`
      sits at 55-70%. Reads like a drafting/sourcing depth gap for this
      batch, not a sync-mechanism bug. Not yet investigated further.
- [ ] **CM — hidden leftover milestone donor text, 32/43 slides, cosmetic.**
      All correctly hidden, nothing visible on screen. Low priority
      cleanup.
- [~] **CO — `Put it on the slides` showed a needless OK-only dialog before
      its one real question.** Fixed: dropped the `MsgBox`, its content was
      already unconditionally in the Run Log. Static check clean; live
      build+test verification still pending -- Excel/PowerPoint were both
      open under a live session at fix time. Run `run_vba_tests.ps1` next
      session before trusting this beyond the static read.
- [ ] **CP — milestone circle colours regressed to one shared teal across
      P/K/S, losing type-specific distinction.** Confirmed against the
      pre-retrofit backup: S was `C0A2F2`, K was `scheme:accent3`. P's
      original colour is unknown (already retrofitted by the earliest
      available backup). Not actioned -- Rohan hasn't said whether he wants
      the pale tints restored or the shared teal kept as the new baseline.
- [~] **CQ — milestone percentage display, Option A built.** `MS<n>_PCT`
      register column, folded into the existing `MS<n>_LABEL` text by
      `DrawFromRow` -- no new shape, no template retrofit. Full design in
      `MILESTONE-PERCENTAGE-DESIGN.md`. Static check clean; live
      verification pending (Excel/PowerPoint open under a live session at
      build time) and no real `MS<n>_PCT` value has ever been entered in
      the register, so it hasn't rendered on an actual slide yet.

## Parked / explicitly decided — do not reopen without a new decision

- [x] ~~Merge the 13 per-field drafting sheets into one long-format sheet.~~
      **Reversed by Rohan, 2026-08-16**: too much time already lost to redesign
      detours on this project; staying with per-field sheets to keep momentum.
      Never written into any doc before this entry.
