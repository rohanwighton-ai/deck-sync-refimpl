# NEXT SESSION — start here

> ## 14 AUG, 10:15 — THE APPROVE-TICK DEFECT IS FIXED AND PROVEN. NOT COMMITTED.
>
> **Working tree carries UNCOMMITTED changes to `vba/Drafting.bas` and
> `vba/tests/TestRunner.bas`.** Commit them. Nothing else is modified.
>
> **The fix.** `WriteDraftingSheet` no longer clears `COL_D_APPROVED`. The tick is
> harvested into a fifth `keptApproved` dictionary under the SAME `carryThisRow` test
> as SUBMIT, so it travels with the text it approves: same quarter keeps both, rollover
> drops both. Defect 1 below is closed — **publish is reachable**.
>
> **Rohan's framing is what produced it**, and it beat both proposals on the table.
> Chat side argued approval cannot live on the register row (the register is downstream
> of approval); Claude Code proposed reversing the chain order. Rohan asked *"why are
> drafting sheets getting rebuilt??? I thought they were staying as part of the record?"*
> — and the answer is that the FREQUENCY was the defect. The code already computed
> `periodChanged` and then ignored its own answer for one column.
>
> **Evidence — five suite runs, every check watched failing:**
>
> | run | state | verdict |
> |---|---|---|
> | 1 | tick fix in | 192/1 · tick test PASS |
> | break 1 | unconditional clear restored | 191/2 · *"SAME-PERIOD REBUILD KEEPS THE TICK ... got `''`"* |
> | break 2 | harvest outside `carryThisRow` | 191/2 · *"ROLLOVER DROPS THE TICK ... got `'Y'`"* |
> | 5 | + index test derived | **193 passed, 0 failed** |
> | break 3 | `DescribeSheet` points at draft column | 192/1 · both index assertions red, naming column E |
>
> Each break named a DIFFERENT assertion, so both halves are independently pinned.
>
> **THE SUITE HAD BEEN RED SINCE `a6e57af`, AND THIS FILE SAID 192/0 THE WHOLE TIME.**
> `WorkbookBridge_IndexExplainsEachSheet` hardcoded layout-3 letters (read C, type D,
> tick E) while `DescribeSheet` was correctly fixed to derive them from the `COL_D_*`
> constants. The test had already been written once to catch exactly this drift, and
> held the next generation of it in place for the same reason: **it typed a
> machine-knowable fact instead of deriving it.** Now derived on both sides. This is
> the write-it-twice class landing in a test, where no document checker can see it.
>
> **`addin84` IS PREPARED BUT NOT SAVED.** Build stamp `2026-08-14 10:09`, all 32
> production modules imported. The `File > Save As > PowerPoint Add-in (*.ppam)` click
> is a permanent manual step (see `build_ppam.ps1` header — proven impossible to
> automate, twice, for independent reasons). **Verify the live build by the stamp
> reading `10:09`** — 15 `.ppam` files are on that machine.
>
> **Still open, unchanged:** defect 2 (chain dies with `Error 50290` when the register
> is open in Excel, then names the wrong file) and defect 3 (three invariant prompts).
> Both sit in scenario 5's path, so it is walkable-with-traps, not walkable.
>
> **Roll-forward spec received** — `OneDrive\Claude\rollforward-spec-2026-08-14.md`,
> scenario 1, all 48 fields with five treatments. Not yet filed into this repo. Findings
> returned to Rohan, highest first:
> 1. **"Placeholder" is doing two jobs** — treatment D means "drafting cell arrives
>    empty awaiting a human", the MECE rule means "standard wording printed on a
>    funder-facing slide". Collapsing them would placeholder active projects that should
>    be reported as GAPS. Pin before building anything on the table.
> 2. Chat's §0 approval argument is overtaken by the fix above.
> 3. §2.3 is a real catch — one milestone device holds three roll-forward cadences;
>    write that test before the device registry lands.
> 4. Totals line does not close (states 51 across treatments, table holds 47 rows,
>    `STRATEGIC_LINKAGES` is argued in §2.1 but has no row).
> 5. Open for Rohan: does the per-project-state rule LAYER with the per-field table or
>    override it? Chat assumed layering and asked to be told.
>
> **Rohan, 14 Aug: he owns the workbook and Claude Code may work that side directly**
> when it is the right tool — holding `feedback_write_office_files_through_office`
> (never hand-write the `.xlsx`) and backing up before any write to the live register.

> ## TWO FIELDS ARE ON REAL SLIDES. The delivery count is 2, not 0.
>
> **14 Aug 2026, 08:21 — `PROGRESS_BODY` written to 43 slides.** Verified from the saved
> `.pptx`, not from a dialog: 43 match the register, **0 differ**, 27 carry real paragraph
> breaks, no literal `||` anywhere. Build `addin83`, stamp `2026-08-14 08:09`.
>
> The workbook's **`Sync Log` is the durable record** (it appends; `Run Log` is replaced
> every run and cannot answer "what has this ever delivered"):
>
> ```
> 2026-08-13 16:55   KEY_EVENTS_BODY   written  21
> 2026-08-14 08:21   PROGRESS_BODY     written  43
> ```
>
> **Closed with it:** whether `||` reaches a slide as a real paragraph break. It does.
> FIX-LIST called that "the last unverified link in the chain".
>
> ### But read this before claiming the loop works
>
> **The 43 values were written into the register BY HAND** (Excel COM), because the
> publish path cannot be reached — see 1 below. So what is proven is *register → slides*.
> *Drafting → register* has still never happened through the tool. They are different
> claims and only one of them is true.
>
> Those 43 values also carry no sources and no recipe hash, because publish is where
> provenance would be written.
>
> ---
>
> ## STATE
>
> - **Deck** `OneDrive\Claude\3. Project Progress.pptx` — 44 slides, period `Q4F26`,
>   nine `ROLE` tags per slide. Tags live in `ppt/tags/tagN.xml`, **not** in the slide XML;
>   grepping `slideN.xml` for them returns nothing and looks exactly like an untagged deck.
> - **Register** `OneDrive\Claude\register-wide.xlsx` — 25 sheets, 43 rows each at
>   `Q3F26` / `Q4F26`, 5 at `Q1F27`. `PROGRESS_BODY` and `ABOUT_BODY` both 43/43 at Q4F26.
> - **Build `addin83` is the ONLY registered add-in** (`AutoLoad=1`; 82 was unticked and
>   removed from the registry). Thirteen stale `.ppam` files still clutter
>   `AppData\Roaming\Microsoft\AddIns` — delete them before one gets loaded by accident.
> - Backups from 14 Aug in `OneDrive\Claude\backups\`: `PREPUBLISH-20260813-204031` is
>   the last Excel-written baseline and is guaranteed openable.
>
> ---
>
> ## THE THREE DEFECTS THAT COST THE MORNING
>
> ### 1. The publish path cannot be reached. Still unfixed.
>
> `Drafting.bas:694` clears `COL_D_APPROVED` on every drafting-sheet rebuild. Draft,
> submit, sources and notes are carried; the tick is not. `SyncNowChainCore` runs
> **rebuild (step 3) immediately before publish (step 4)**, and nothing but a person ever
> writes that column. So a tick can never survive to be read.
>
> Demonstrated, not inferred: 43 rows ticked `Y`, saved to disk, `Sync Now` pressed,
> publish reported `0 would be published, 43 drafted but not ticked`.
>
> **Do not "fix" it by never clearing the column** — a tick approves a *specific* text.
> The real question is whether approval belongs on the **register row** instead of on a
> surface that is rebuilt by design. That is a design call, not a patch.
>
> ### 2. The chain fails when the register is open in Excel, and does not say so
>
> `Refresh Drafting Sheets` raises **`Error 50290`** when Excel already holds the
> workbook; every later stage then reports *"Could not open the paired workbook at
> C:\...\register-wide.xlsx"* — naming the file that is fine and discarding Excel's real
> reason. Self-defeating too: step 3 leaves the workbook dirty, and the apply's own guard
> then correctly refuses to read values that exist only in Excel's memory.
>
> **Workaround until fixed: close Excel before pressing `1. Sync Now`.** Cost four failed
> runs on 14 Aug before the pattern was visible.
>
> ### 3. Three prompts have an invariant answer
>
> Two presses produced ~20 dialogs; four were real decisions. The `MS*` warning, the
> 17-column question and Roll Forward can only ever be answered the same way. Rohan:
> *"way too many msgboxes and having to answer no is confusing."* This is not polish —
> it trains click-through past the one dialog where `No` destroys 43 ticks.
>
> ---
>
> ## FIXED IN `addin83`
>
> `DeckRegistry.SaveDeckVerified` and `WorkbookBridge.SaveWorkbookVerified` now resolve a
> OneDrive URL through `LocalPathForUrl` before touching the filesystem. Previously
> `fso.FileExists(url)` returned **False** for an `https://` path, so both functions
> returned "this workbook has never been saved to a file" and **exited before calling
> Save** — the one function written to guarantee the save was the only thing skipping it.
> Nine call sites for the workbook, two for the deck. Where the path cannot be resolved
> the save is now *attempted* and reported as unverified, rather than silently refused.
>
> ---
>
> ## NEXT, IN ORDER
>
> 1. **Approve-tick placement** (defect 1). Unblocks drafting → register, which is the
>    whole point of the recipes.
> 2. **Delete the invariant prompts** (defect 3), and make defect 2 say "close Excel first".
> 3. **Tag the sixteen standing `Given` fields on the template.** Propagation to
>    already-linked slides **is supported and tested** — `ExistingInstanceKey` protects
>    their keys, hardened after the 2026-07-26 incident that orphaned 46 slides. Only the
>    entry point is wrong: it asks *"Name for this new slide type"*. Verify what it does
>    on re-registering `project-progress`, on a carved copy, first.
> 4. **Then the template pivot** — but not making the template authoritative for the field
>    set until a quarter has been produced from it.
>
> **ANSWERED 2026-08-14, and do not re-ask it.** The load-bearing question of the whole
> project -- did the recipe tell him what to write, or did he still work it out himself --
> got its answer minutes after he was shown the `PROGRESS_BODY` Field Spec row:
> **"It worked and we will continue to refine over time."** The central bet holds: the
> recipes remove the re-deciding. A first affirmative read across two fields, not a full
> pass over thirteen; refinement is expected and he said so. `TRACKER.md` item 9 is ticked
> on the strength of it -- 9 of 10.
>
> **NOT blocked, and both halves are settled: Rohan owns the declared-linkage list, and
> its home is a PERMANENT sheet in the register workbook beside `Sources`.** It was put
> back to him as an open question on 14 Aug after both parts had already been answered --
> twice for ownership, and once by reading `WorkbookBridge.LifespanOf`. What remains is a
> task, not a decision: create the sheet, give it an `as at` date, and put the four codes
> already visible on the real slide (`1.1.1`, `2.1.1`, `3.3.1`, `3.4.1`) into it.
>
> Original note follows. FIX-LIST
> records this twice (*"Rohan owns it, colleagues later"*) and it was still written up as
> a question for him on 14 Aug, from a file that had been read in full the same morning.
> The real gap is that the list does not exist as an artifact yet: it needs a home
> **outside** the register workbook, which the tool rebuilds and clears, plus an `as at`
> date. Four declared codes are already on the real slide (`1.1.1`, `2.1.1`, `3.3.1`,
> `3.4.1`) with nowhere to live.
>
> ---
>
> ## RULES THIS MORNING EARNED
>
> **Never hand-write an `.xlsx`.** Re-serialising a worksheet with ElementTree produced a
> file that passed zip integrity, XML parsing, relationship, content-type, cell-order and
> inline-string checks — and Excel refused to open it, because only two namespaces were
> registered and `mc`/`x14ac`/`xr` got invented prefixes. **The only valid test of "can
> Excel open this" is Excel opening it**, A/B'd against a known-good file in the same
> session. Drive Excel over COM to write; read by hand is still fine.
>
> **A defect is a class, and so is a diagnostic.** "Could not open the paired workbook"
> appeared three times this morning with three different underlying causes, all hidden by
> the same discarded error. FIX-LIST item 1 is now the highest-value cheap fix in the repo.
>
> ---

**Written 13 August 2026, ~16:15.** Previous version archived as `NEXT-SESSION-2026-08-12.md`.


> ## SUPERSEDED — written mid-session, kept for the reasoning below it
>
> This banner read "THE DELIVERY COUNT IS STILL ZERO". **It is no longer true.** 21 drafted
> `KEY_EVENTS_BODY` values reached real slides at 17:23 on 13 Aug, verified three ways: the
> apply dialog, the slide XML in the saved file, and a screenshot of the rendered slide.
>
> 43 slides tagged and linked, 0 failed verification, register saved, deck period `Q4F26`
> on disk. The section below explains the defect that had blocked it, which is now fixed
> and proven — keep it for the reasoning, not as a description of current state.

---

## THE BLOCKER — FIXED AND PROVEN. Kept for the reasoning and the rejected workaround.

`RibbonUI.SyncNowChainCore` step 4 is `DraftingUI.PublishDraftsForField`, which begins:

```vba
fieldId = ActiveDraftField(wb)          ' whatever TPL_ sheet is ACTIVE in Excel
If fieldId = "" Then fieldId = AskForField(CAP, wb)
```

It only asks **if** the active sheet is not a drafting sheet. But the step immediately
before it — `RefreshDraftingSheets` — ends with `ShowSheet wb, firstSheet`, and
`firstSheet` is the **first `Kind = Prose` row on the Field Spec sheet**
(`DraftingUI.ProseFields`, row order).

That row is `ABOUT_BODY`, which has **0 submitted, 0 approved**. So every run publishes
an empty sheet, reports "0 would be published", and finishes quietly. **The chain cannot
reach any other field.** There is no field picker on the toolbar (two buttons only), so
this is the only publish route.

This is very likely a large part of why this project has never got a field onto a slide.

**A ROW-REORDER WORKAROUND WAS PROPOSED AND REJECTED — DO NOT USE IT.** Moving
`KEY_EVENTS_BODY` above `ABOUT_BODY` on the Field Spec would work, because
`FieldSpec.WriteSpecSheet` only seeds *missing* FieldIDs and never reorders existing rows.
It is still the wrong move, and Rohan stopped it with one question: *"Why are you having
to move register rows manually? Worries me that the code won't work when it needs to."*

He is right, for three reasons:

1. **It makes which field reaches a slide depend on spreadsheet row order** — invisible,
   unstated coupling of exactly the kind that has bitten this project repeatedly.
2. **It is not available at work.** No Claude, no WSL, no Python there — a quarter has to
   be runnable from toolbar buttons. "Reorder rows in a spec sheet so the right field
   publishes" is not a procedure; it is a defect with instructions attached.
3. **It would have hidden the defect behind a successful-looking run**, which is the
   failure mode this project keeps rediscovering.

**FIXED in `addin81`** (build stamp `2026-08-13 16:24`). New `DraftingUI.FieldForRun`:
inside a collected chain it ASKS; standalone it still reads the active sheet, because
there the answer really is on screen. Asked ONCE per run and reused, so the two stages
that need it do not ask twice.

**It had TWO call sites.** `CopyAiDraftsToSubmit` carried the identical line and the
identical consequence — fixed together rather than only where it was noticed.

**PROVEN 2026-08-13 17:23 on the real deck.** 21 KEY_EVENTS_BODY values written and confirmed by reading slide XML from the saved file -- slide 1 now carries the drafted "The industry partner's withdrawal..." wording. The delivery count is no longer zero. Previously read: 192 tests pass and the project compiles, but no test exercises the
chain's field selection — which is precisely the gap that allowed this defect. Green here
means "nothing broke", not "the fix works". Prove it by pressing the button: `Sync Now`
must now ASK which field, and `KEY_EVENTS_BODY` must be selectable.

**Note what the test suite did NOT do here.** 192 tests pass. Not one of them asks "can a
person cause `KEY_EVENTS_BODY` to be published?" — they test that publishing works when
called, not that the chain can reach it. Same "tested unit behind a locked door" shape as
the picture injection and the progress bars, found the same way: by pressing the button.

---

## STATE, VERIFIED FROM FILES (not from dialogs)

- **Deck** `OneDrive\Claude\3. Project Progress.pptx` — 44 slides, 49,247,250 bytes.
  `DeckSyncPeriod = Q4F26` confirmed by property name in `docProps/custom.xml`.
  Slide 44 is the hidden master template, 9 fields set to `<<placeholders>>`.
- **Register** `OneDrive\Claude\register-wide.xlsx` — 308,072 bytes, `Register` sheet has
  92 rows (1 header + 91: 43 Q3F26 + 43 Q4F26 + 5 Q1F27). All 43 Q4F26 instance keys
  match the Q3F26 keys exactly — **the handover's "stale/foreign Q4F26 rows" warning was
  wrong**, they describe the same slides.
- **Backups** `OneDrive\Claude\backups\2026-08-13-1520-post-onboard-Q3F26 - *` — deck and
  register, both verified byte-identical by md5 at the time of copy.
- **Build `addin82`**, stamp `2026-08-13 19:22` — CURRENT, confirmed loaded and the only
  add-in registered. **Source is one commit ahead of it** (`20dc2a9`, six sheet-name
  literals replaced by their constants). The values are identical, so behaviour is
  unchanged and a rebuild is not required — but `addin82` is not byte-equivalent to `HEAD`. In `OneDrive\Claude\` and the trusted location
  `AppData\Roaming\Microsoft\AddIns\`. (Superseded: `addin80` stamp 14:37, `addin81`
  stamp 16:24.)

### Drafting sheets — real counts (header row EXCLUDED)

| sheet | submitted | approved |
|---|---|---|
| `TPL_KEY_EVENTS_BODY` | 43 | 39 |
| `TPL_PROGRESS_BODY` | 34 | **42** |
| `TPL_HIGHLIGHTS_BODY` | 43 | 42 |
| `TPL_ABOUT_BODY` | 0 | 0 |
| `TPL_STRATEGIC_ALIGNMENT_BODY`, `TPL_PROBLEM_BODY`, `TPL_STRATEGIC_LINKAGES` | 0 | 0 |

`PROGRESS_BODY` has **more approvals than submitted text** (42 vs 34). Those 8 rows
publish nothing — both text and tick are required — but the count will look wrong.

`HIGHLIGHTS_BODY` has the most work in it and **cannot publish**: it is not one of the
nine tagged fields and needs slot columns, not one column. See FIX-LIST.

---

## ENVIRONMENT FINDING — SAVES AND ONEDRIVE

Both files are open via **OneDrive URLs**, not local paths:
`https://d.docs.live.net/96b9ec593ee3ba55/Claude/…`

With AutoSave **off**, the deck period write failed **4 verified attempts**, and a manual
`Ctrl+S` did not change the file's mtime either. `SetDeckPeriodVerified` correctly
detected this and refused to continue:

> `THE PERIOD DID NOT REACH THE FILE after 4 attempt(s). Asked for: Q4F26  On disk: Q3F26`

**Turning AutoSave ON made the write land.** This is an environment condition, not a code
defect — and it is the configuration the work machine will be in. The tool's behaviour
here was correct and is what a week ago was missing: it checked the file, not its own
cache, and refused rather than reporting success.

---

## WHAT SHIPPED TONIGHT

- **Suite green: 192 passed, 0 failed** (was 190/2), behind the compile gate.
- The two tests asserting the deleted defect were rewritten **and renamed**, because the
  old names stated the defect as the requirement:
  - `Drafting_PeriodRolloverDropsStaleSubmit` → `Drafting_RolloverRebuildsOnlyWhenNothingIsAtRisk`
  - `Drafting_RolloverKeepsEntityStaticRows` → `Drafting_RolloverCadenceGovernsUntypedRows`
- **`RefreshDraftingSheets` no longer reports success over a refusal.** It collected
  refusals into the Run Log and then said *"drafting sheets are ready. Workbook saved."*
  with an information icon. Now: refusal first (so MsgBox truncation eats the guidance,
  not the warning), refused field names listed, warning icon.

### NOT YET TRUSTED

**Neither rewritten test has been made to fail on purpose.** Green alone is not evidence.
Break each before relying on it — for `...RolloverCadenceGovernsUntypedRows`, put SUBMIT
text back on the fixture and it should stop testing anything, because the refusal
pre-empts the whole path.

---

## FILES CHANGED THIS SESSION — ALL COMMITTED AND PUSHED

`deck-sync-refimpl` (main, in sync with origin):
- `vba/DraftingUI.bas` — refusal surfaced in the dialog; `FieldForRun`; roll-forward skip
- `vba/RibbonUI.bas` — `PickType`; sync-log constant at two creators
- `vba/AdoptFlow.bas` — third `PickType` call site
- `vba/ExcelOutput.bas` — `PeriodRowCount`
- `vba/WorkbookBridge.bas` — `SYNC_LOG_SHEET_NAME`; lifecycle tab ordering
- `vba/tests/TestRunner.bas` — two tests rewritten + renamed
- `FIX-LIST.md`, `NEXT-SESSION.md`

`claude-brain`: `DECISIONS.md` — the template-first decision.
`Zettelkasten`: `20260813-the-device-is-the-unit-of-addressing-not-its-parts.md`.
`OneDrive\Claude`: `reply-from-claude-code-2026-08-13-evening.md` — answers chat side's
three questions; not a repo, so not versioned.

Builds: `addin80` (14:37) and `addin81` (16:24), both superseded. **`addin82` (19:22) is loaded and
the only add-in)**. Suite green at 192 passed / 0 failed with the compile gate clean across
33 modules.

---

## OPEN, IN PRIORITY ORDER

1. **Publish one field.** Field Spec row move → `Sync Now` → `KEY_EVENTS_BODY` → review →
   apply. Then verify by reading the slide XML out of the saved deck, not the dialog.
2. **Fix the publish-target defect properly** (above), then revert the row move.
3. **FIX-LIST 1c/1d** — at-risk scan misses SOURCES/NOTES; the park that reports "nothing
   was lost" runs *after* `ws.Cells.Clear`. Fixing 1c makes 1d unreachable.
4. **The cadence machinery is probably dead code.** The refusal pre-empts it; it now
   governs only SOURCES/NOTES on untyped rows. If 1c is fixed, delete it rather than
   maintain it. Rohan's call.
5. **`MILESTONE_TIMELINE` group tagging is UNVERIFIED.** It was not among the nine tagged
   fields. If the timeline renders blank after a sync, check this first.
6. **Field Spec `Kind` values look wrong for the milestones**: `MS1_LABEL`/`MS7_LABEL` are
   `Given` while `MS2`–`MS6_LABEL` are `Prose`; MS1/MS7 DATE+DONE are `Derived` while
   MS2–MS6 are `Given`. 13 fields are `Prose` but only 7 have drafting sheets, so the next
   refresh will create six more tabs.
7. Slide 44 still carries P001's unmanaged content (figures, photo, team). The audit found
   **50 unmanaged text items on slide 1, 21 of which look like project data** — that is the
   next tagging backlog, and the same set the Field Spec wants columns for.

---

## THREE THINGS WORTH KEEPING

**A warning that only reaches the log is not a warning.** The refusal guard was correct
and invisible; the dialog said "ready" over seven refused sheets. Fixed, but the shape
recurs — check where a message *lands*, not just that it exists.

**The check that found the save failure was the one that read the file.** Four
in-process attempts all "succeeded". Only comparing against `docProps/custom.xml` on disk
told the truth. Evidence must come from the far side of the boundary.

**"Nothing happened" meant a dialog behind the window — twice.** A VBA modal can open
behind PowerPoint. Before diagnosing a dead button, Alt+Tab. A calibrated test:
PowerPoint stops answering COM (`ActivePresentation.Name` comes back empty) while a modal
is open, and answers normally when idle.

---

## ARCHITECTURE — DECIDED IN PRINCIPLE 2026-08-13, NOT YET BUILT

Two calls made at the end of the first successful publish. Both are Rohan's, both are
right, and both should be settled properly before more feature work.

### 1. TEMPLATE-FIRST, NOT DISCOVERY-FIRST

Rohan: *"Are we better off gearing it to be an expert template builder and pushing a
pattern we know? ... maybe that's best rather than a sensory beast that doesn't quite
know what it is trying to be."*

**Yes.** The strongest argument is the WORK MACHINE. A sensory tool needs an operator with
judgement at every step — tonight it needed Claude to decide which of 59 shapes were
fields, whether the `MS*` warning was real, which register to pair, whether `TESTFILL` was
junk, and whether 88 proposed changes were safe. At work there is no Claude. A
template-driven tool needs no run-time judgement, because the judgement was made once, in
advance, and frozen into an artifact.

**The cost asymmetry says the same.** Discovery runs ONCE per slide type, ever. Publishing
runs 43 slides x 9 fields x 4 quarters, forever. Nearly all the code, nearly all the
defects and nearly all of 13 Aug went into the once-ever path.

**And look where the defects actually were:** the grid that loses marks, the blank grid
that unmarks, the 21 `MS*` false positive, 50 "unmanaged" items it can only guess at,
"which register?", the pairing. Every one is a PERCEPTION defect. The publish path — the
thing that runs every quarter — had one defect, four lines long.

**The mechanical diagnosis under "doesn't know what it's trying to be":** the tool holds
THREE sources of truth about what a field is — a tagged shape, a register column, and a
Field Spec row — and reconciles them at run time. Every reconciliation is a place to be
wrong. A template collapses all three into one, decided at design time.

**What that means concretely**
- The template is the authority; register schema derives from it.
- `FieldWiring`'s orphan-column question dissolves: nothing can be orphaned if the
  template defines the set.
- Discovery is DEMOTED to a one-off migration tool. Run once per deck, then never
  developed again. It has already been run — 13 Aug — so for this deck it is done.
- New projects clone slide 44 and are conformant by construction.

**What NOT to throw away:** the verification discipline (the file is the evidence; prove a
check can fail; a defect is a class), the Office/COM/OOXML knowledge, the consent-gate
design, and the register-deck contract. None of it is discovery-specific; all of it
transfers to the next project.

**The honest risk:** real slides are messier than any taxonomy — the finding that made
discovery seem necessary in the first place. But that now cuts the other way: owning the
template BOUNDS the mess instead of trying to perceive it, and the messy migration has
already happened.

### 2. A DEVICE REGISTRY — PROTECT COMPOUND SHAPES FROM THE GENERIC MACHINERY

Rohan: *"I can already see the need for a specialist module spot to protect complex shape
mechanisms like the timeline being pulled into marking etc."*

Three pieces of evidence from one evening:

- `MS1_DATE` … `MS7_LABEL` appeared as 21 rows in the Discover Fields grid, and the only
  thing that stopped them being tagged was **Claude telling Rohan not to**.
- `FieldWiring.ScanFieldWiring` reported those same 21 columns as orphaned on EVERY run —
  the recurring "21 field(s) on the register that no slide carries" warning.
- The Template Audit counted device internals among its "50 unmanaged text items, 21 of
  which look like project data".

Three generic mechanisms seeing the device's PARTS; none of them seeing the device.

**The model already exists, in exactly one place.** Injection has it right:

```vba
InjectPrimitive.InjectField(sld, "MILESTONE_TIMELINE", "", False, Nothing, row)
```

One addressable thing, consuming its own columns off the register row. That understanding
never propagated to discovery, wiring, marking or audit — so a device is a first-class
citizen at write time and a pile of loose shapes everywhere else.

**Principle: the device is the unit of addressing, not its parts.**

**Shape of the fix.** One declaration per device — its role tag, the register columns it
consumes, its internal shape-name pattern — read by four consumers:

| consumer | today | with the registry |
|---|---|---|
| Discovery | lists 21 internals as candidate fields | skips anything inside a declared device |
| `FieldWiring` | reports 21 orphan columns every run | counts them as OWNED by the device |
| Marking | will happily tag a device internal | refuses |
| Template Audit | counts internals as unmanaged project data | classifies as device internals |

One declaration, four consumers — versus four independent special cases, which is what
would get written if this is approached site by site.

**Why it is urgent rather than tidy:** "leave rows 41-65 alone" was ADVICE GIVEN TO A
HUMAN. Zettel `20260719-telling-an-agent-not-to-do-something-isnt-a-control` — and it is
not a control when you tell a person either. At work, with nobody to say it, the next run
of Discover Fields tags 21 timeline shapes as individual fields and quietly destroys the
device.

**It folds into the template pivot.** In a template-owned world the device is PART of the
template, so it is declared by construction rather than looked up — the registry stops
being a side-table and becomes a property of the thing already controlled. The expensive
half, knowing what a device is, is already written.

### 3. ALSO REQUESTED, NOT STARTED

**Logical tab numbering and Excel best practice across all workbooks.** `register-wide.xlsx`
has 18 sheets in arrival order with no scheme (`START HERE`, `Sources`, `SRC_EXTRACTS`,
`Field Spec`, seven `TPL_*`, `Register`, `Run Log`, `Sync Log`, `Field Discovery`,
`Template Audit`, a stale `Review project-status-2D3D`, plus the live review sheet).
`WorkbookBridge.ArrangeTabs` already orders them, so the scheme belongs there rather than
in a manual pass. Do this FIRST next session — it is small, bounded, and was explicitly
asked for.

---

## EXCEL TAB ORDER — DONE BY POSITION, DEFERRED BY NAME (13 Aug, late)

Rohan asked for logical tab numbering and Excel best practice across the workbooks, then
added: **"anything that threatens the data chain fix it"** and **"if you are going to
renumber use logic"**. Both shaped what was and was not done.

### Done: a data-chain fix found while auditing the sheet names

**`"Sync Log"` was a bare literal in SEVEN places across two modules** — alone among the
tool-owned sheets, every one of which otherwise has a constant. Two of the seven are
`GetOrAddWorksheet` calls, which **create** the sheet when the name does not match. So one
divergent literal would not fail loudly: it would quietly open a second log sheet while
`IsToolOwnedSheet` and `ArrangeTabs` went on guarding the first, splitting the audit trail
while looking healthy. Now `WorkbookBridge.SYNC_LOG_SHEET_NAME`, all seven replaced.

### Done: tab order follows the lifecycle of a quarter

`ArrangeTabs` now orders by position — no renaming, so nothing can break a lookup:

```
 1  START HERE     where a person begins; the readiness checklist
 2  Field Spec     what fields exist at all -- configuration before data
 3  Sources        the evidence values may cite
 4  SRC_*          harvested source data
 5  Register       THE DATA. The reason the workbook exists.
 6  TPL_*          where a person works, in Field Spec order
 7  Review *       the approval gate, between work and the deck
 8  Field Discovery, Template Audit    diagnostics, off the normal path
 9  Run Log, Sync Log                  the audit trail
10  SAVED *        parked archives, absolutely last
```

**`Register` was previously unplaced** — it fell into "everything else in its current
order" beside the diagnostics, so the most important sheet in the workbook sat wherever it
happened to land. That, rather than any cosmetic gain, is what this fixes. `SAVED *`
archives now sort last so they cannot be mistaken for the live sheet they were copied from;
typing in one is silent, because publish reads the live sheet only.

### Compile-verified 13 Aug, after Rohan closed Office

```
COMPILE OK: whole project compiled clean (33 modules).
RESULT: OK
=== 192 passed, 0 failed ===
```

Static checks also clean across 34 modules. The new cross-module references
(`DiscoverUI.DISCOVERY_SHEET_NAME` and `TemplateAudit.AUDIT_SHEET_NAME` read from
`WorkbookBridge`) compile.

**NOT YET IN AN ADD-IN.** The ordering and the `SYNC_LOG_SHEET_NAME` constant are in source
only at the time of writing — `addin81` predated them. **`addin82` has since been built,
installed and confirmed loaded**, so the ordering is live in code; it takes visible effect
on the next completed drafting rebuild, since `ArrangeTabs` runs during one.

### Deferred deliberately: numbering the NAMES

`01_FIELD_SPEC`-style names would be better still, and are a **migration, not a tidy-up**:
sheet names are this tool's addressing mechanism — nine constants, dozens of literals, plus
the `TPL_`, `Review ` and `SAVED ` prefix matches. It needs the constants changed, every
literal found, and a rename pass over a live workbook holding drafted work, with a fallback
for a workbook that has not been migrated yet.

**If it is done, the scheme should be the ordering above**, so position and name agree and
neither can drift from the other. Do it as its own session with the suite green before and
after — not alongside anything else.

---

## REQUESTED 13 Aug: a standard placeholder for projects that have stopped reporting

Rohan: *"if it's missing because they have stopped reporting please include a standard
placeholder for now."*

**The need is real.** `3_P001` is `Project Closed` and its `PROGRESS_BODY` drafting row is
`SUBMIT` empty / `APPROVED = '0'`. A closed project should not silently render last
quarter's prose or an empty box.

**Proposed standard wording:**

> `No update this quarter — project closed Q3F26. Last reported Q1F26.`

States that the absence is intentional, why, and where the last real content is.

**THE TRAP, and why this was not just typed in.** A placeholder typed into `SUBMIT`
publishes as ordinary content and is then indistinguishable from drafted prose — next
quarter nobody can tell which rows are real. That is exactly the `TESTFILL-1256` string
found in the Q4F26 register tonight, which was one bundled "Yes" from reaching a slide.

**So it should be DERIVED, not typed.** The register already knows `PROJECT_STATUS` and the
period. A placeholder generated at publish time for rows where status is closed/not-started
AND the field is empty is:
- consistent by construction, so the wording cannot drift between projects;
- distinguishable from human prose, so a later review can find every one;
- self-correcting — the moment real text is drafted it takes precedence.

Typing it into 43 drafting rows gets the same pixels this quarter and a mess next quarter.

**Open decisions for Rohan:**
1. Which statuses trigger it — `Project Closed` only, or `Not Started` too?
2. Does it apply to every prose field, or only the quarterly ones (`PROGRESS_BODY`,
   `KEY_EVENTS_BODY`) and not entity-static ones like `ABOUT_BODY`?
3. Exact wording, including whether to name the closing quarter dynamically.

**Do not bulk-type placeholders into the drafting sheets before deciding 1–3.** Undoing 43
rows of typed placeholder is harder than generating it once.

### ANSWERED 13 Aug — the placeholder rule

Rohan on the three open decisions: *"1) whatever makes MECE logical sense 2) only the
quarterly ones, good instinct 3) not sure it should be one of a set selection yeah? like
validated data?"*

**2 is settled: quarterly fields only** (`PROGRESS_BODY`, `KEY_EVENTS_BODY`). Entity-static
fields are excluded — a closed project still *is* what it always was, so a placeholder on
`ABOUT_BODY` would be false.

**3 is settled: a validated set, not free text.** The wording is SELECTED, never typed —
same mechanism as the Controlled fields and `ApplyControlledValidation`. Selecting it is
the consent, and it cannot drift across 43 projects.

**1, the MECE partition. It is over `status x has-text`, not status alone** — that is what
makes it collectively exhaustive:

| project state | field empty | treatment |
|---|---|---|
| Closed | yes | placeholder — no further update will ever come |
| Not yet commenced | yes | placeholder — none due yet |
| Active, explicitly marked nothing-to-report | yes | placeholder — a deliberate statement |
| **Active, no explicit mark** | yes | **NOT placeholdered. Reported as a GAP.** |
| any | no | real text always wins |

**Row four is the whole point.** An active project with an empty quarterly field is a
MISSING UPDATE, not a non-update. Auto-placeholdering it would silently convert "chase this
up" into a tidy sentence on a funder-facing slide — the same shape as every
reports-success-without-confirming-the-effect defect in this file, pointed at content
instead of code.

So the vocabulary needs a fourth member that a person CHOOSES, e.g. `Nothing to report`.
Its presence is what authorises a placeholder on an active project; its absence is a gap
the run should name.

**Prerequisite: normalise `PROJECT_STATUS` first.** The vocabulary is currently inconsistent
(`In progress`/`In Progress`, `Not started`/`Not Started`, plus `Not yet commenced`) — see
"Five status values differ by one capital letter". A MECE rule keyed on status cannot be
built on a status set that has near-duplicate members. Fix that before implementing this.
