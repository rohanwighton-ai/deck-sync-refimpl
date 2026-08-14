# NEXT SESSION — start here

> ## 14 AUG, ~23:10. **STATUS: CURRENT.** Everything below is historical.
>
> ### THE HARVEST DID NOT EXIST. IT DOES NOW, AND IT HAS NEVER RUN ON THE DECK.
>
> The block below says the harvest is a press away. **It is not, and it never
> was.** `DeckAdoption.PlanAdoption` skips any slide carrying `slide_type` +
> `instance_key` at `DeckAdoption.bas:149`, **before any matching runs**. All 43
> project slides carry both. So adoption on a live deck reports 43 already-linked
> and reads nothing — by construction, not by accident. Adoption is for slides
> the tool has never seen; every slide here was adopted months ago, which is how
> the nine prose roles reached all 44.
>
> Searched for any other route: `PlanAdoption`/`CommitAdoption` are the only two
> functions that read a value off a slide into the register. The `CurrentValue`
> reads in `InjectPrimitive` are what-was-there-before records inside the
> INJECTION path, not a harvest.
>
> ### FIRST ACTION: REBUILD THE ADD-IN, THEN DRY-RUN PROPAGATION ON **ONE** SLIDE
>
> Nothing below is reachable until the modules are re-imported and
> `File > Save As > PowerPoint Add-in` is clicked. The loaded add-in is still the
> pre-harvest build.
>
> Then: select ONE project slide, press `1. Set up my quarter`, and answer **No**
> at the prompt. The dry run writes nothing. **Read what it offers, not what it
> claims** — and note the suite cannot substitute for this: every test fixture is
> a two-shape slide, and slide 1 has **136 shapes, four of them named
> `Shape 46`**. The matcher has never been asked anything that hard.
>
> ### THE TAGGING WAS NOT 29 FIELDS. IT WAS 7, AND IT IS DONE.
>
> Read from the deck's XML, not from the field list:
>
> | | |
> |---|---|
> | 21 `MS*_LABEL/_DATE/_DONE` | ONE tag. `Discovery.bas:165` recognises the timeline by counting slots, not by name, so the device is one candidate |
> | 6 scalars | real shapes, tagged |
> | `SECTOR`, `TRL` | **no shape at all** — substrings of one composed subtitle line |
>
> **Written to slide 44 and verified from the saved bytes**: tag parts 440 → 447,
> each of the seven ROLE values appearing exactly once, file moved 19:20 → 22:33.
> `START_DATE`, `END_DATE`, `PROJECT_LEAD`, `SUBTITLE_A`, `INDUSTRY_CASH`,
> `TOTAL_VALUE`, `MILESTONE_TIMELINE`. Backup: `backups/PRE-TAG-20260814-222900/`,
> md5-verified.
>
> ### WHAT THE DECK ACTUALLY LOOKS LIKE — settled from the file, do not re-derive
>
> - **Slide 44 is the template** (the one tag file with `SLIDE_TYPE` and no
>   `INSTANCE_KEY`), and it is a **clone of slide 1 with only the nine prose
>   fields placeholder-ised**. Every Given value on it is still project
>   `3_P001`'s real data, including a shape named `3_P001 Timeline`.
> - **The milestone device is named on 2 of 44 slides** — 44 and 1. Nowhere else.
> - **There is a stray `MS2_ON` OUTSIDE the timeline group**, at slot 3's
>   position, on both slides. `PartsOf` only walks the group, so the device can
>   never control it. Uncontrolled dark circle over slot 3.
> - **The slide shows four money figures and the register names two.** SAAFE Cash
>   and In-Kind are labelled on the slide and have no field.
>
> ### WHAT SHIPPED — suite **197 passed / 0 failed**, compile clean, 34 modules
>
> - **`vba/Harvest.bas` (new).** `HarvestSlide` reads role-tagged shapes into the
>   register row for the slide and period. **A cell is written ONLY when the
>   register holds nothing**, read structurally as `Not rowValues.Exists(field)`
>   because `ReadSheetForPeriod` only dictionaries non-empty cells. Refuses on
>   duplicate instance rows, never creates a column, never invents a row, and
>   **skips a tagged group by name** rather than reading a device's empty text as
>   a blank field.
> - **`PropagateTemplateTags`** — the missing middle link. Carries the template's
>   roles onto an already-linked slide, reusing `MatchSlideAgainstTemplate`
>   unchanged. **The caller must filter the roles first**: that function loops
>   EVERY template role while filtering the target to UNTAGGED shapes, so asking
>   about a role the slide already carries scores it against some unrelated
>   leftover. Stamps only on `high`; refuses when two roles claim one shape.
> - **`AdoptFlow.AdoptExistingSlides` was orphaned** — no button, no caller, for
>   the whole life of the three-button toolbar. Now called from
>   `RibbonUI.OfferAdoptionForSelectedSlides`. Its header said "Toolbar entry
>   point" the entire time.
> - **`check_vba_static.py` could not see it**: `AdoptFlow.bas` was missing from
>   `UI_MODULES`. Added. **And the checker is weaker than its name** — it asks
>   whether a NAME appears in another module, not whether anything reachable
>   calls it, so a chain of private orphans is still invisible to it.
> - Both new doors are gated on there being something to do. In Normal view a
>   slide is ALWAYS selected, so an ungated "you have slides selected" prompt
>   would fire on nearly every press.
>
> ### PROVEN BY A DELIBERATE BREAK — AND ONE THAT WAS WEAKER THAN CLAIMED
>
> - **Harvest's empty-cell rule: real damage.** Guard removed → all three
>   assertions failed, `FIELD_A was NOT overwritten, got 'FROM THE SLIDE A'`,
>   `Written` 2 not 1. The device test stayed green, isolating it.
> - **Propagation's role filter: bookkeeping only.** Filter removed → **one**
>   assertion failed (`already on the slide, got 0`). `FIELD_B` still landed
>   correctly and `FIELD_A` was not corrupted. **So the test proves the guard is
>   reached, NOT that removing it puts a role on the wrong shape.** The fixture
>   has two easily-distinguished shapes; the real slide has 136. This is the
>   weaker of the two guards and it is the one that writes to the deck.
> - **The first break attempt did not break anything** — commenting the call to
>   the wrapper left the callee's name still written inside it, and the checker
>   only greps for the name. That is what exposed the checker's real contract.
>
> ### ALSO: A SKIPPED SUITE RUN EXITS 0
>
> `run_vba_tests.ps1` refuses to close an open PowerPoint (correctly, since
> 14 Aug) and **exits 0 having done nothing**. Its own header still describes the
> old "ask it to Quit()" behaviour, which is what produced a wrong prediction to
> Rohan. Read the output file, never the exit code.
>
> ### OPEN, NAMED
>
> 1. **Propagation has never touched a real slide.** Dry run one first.
> 2. **The device covers 21 of the 29 fields and cannot be read back.** Slot
>    state lives in shape visibility — GAP 3, unbuilt. Harvest skips it by name.
> 3. **`SECTOR`/`TRL`: Rohan chose to split them into their own shapes.** Not
>    built. Note the split does NOT reach the other 42 slides, and `Text 4`'s
>    other two facts already exist as their own shapes (`Text 110`, `Text 112`) —
>    so `SUBTITLE_A` may be a composed output, not a field.
> 4. The stray `MS2_ON` on slides 1 and 44.
> 5. `STRATEGIC_LINKAGES` has a register column, no values, no template tag.
> 6. `PROJECT_STATUS` still has nothing enforcing its vocabulary.
>
> **Delivery count is 2.** Nothing reached a slide, and none of tonight's code
> has run against the real deck. Everything is committed to the working tree
> only — `Harvest.bas` and `backups/` are untracked.
>
> ---

> ## 14 AUG, LATE NIGHT (~22:15). **SUPERSEDED** by the block above.
>
> ### THE THING THAT CHANGED TONIGHT: DRAFTING → REGISTER WORKS.
>
> `3_P001`, `KEY_EVENTS_BODY`, 272 chars, published **by button**, 20:45.
> That link had never once run in this project's life. S5 is closed.
>
> ### FIRST ACTION TOMORROW: **THE HARVEST.** Nothing else comes first.
>
> There are roughly **1,247 values** — 29 `Given` fields × 43 projects — sitting on the
> slides and **nowhere else**. Read from the register tonight: `START_DATE`, `END_DATE`,
> `INDUSTRY_CASH`, `TOTAL_VALUE`, `PROJECT_LEAD`, `SECTOR`, `TRL`, `SUBTITLE_A`, every
> `MS*_LABEL/_DATE/_DONE` — **0 of 43 rows populated**. `PROJECT_STATUS` (43) and
> `PROJECT_PROGRESS` (1) are the only non-prose fields with anything in them.
>
> Tagging is not a chore, it is what makes harvesting possible: an untagged shape is
> anonymous, so the tool can neither read from it nor write to it. **Tag once on the
> template, harvest, and the register fills itself from the deck.** Nobody types 1,247
> values.
>
> **Do not do the quarter turn before the harvest** — it would roll forward emptiness.
>
> ### STATE, READ FROM FILES
>
> - **`addin88` is the only registered add-in**, `AutoLoad=1`, confirmed loaded.
> - Suite **194 passed / 0 failed**. Compile clean, 33 modules. Static clean.
> - Commits tonight: `a9f52b7`, `6c74912`, `3e7db62`, `db364b3`, `d1278b4`, `368d873`.
>   All pushed.
> - Register **394,082 bytes**, zip integrity OK, 39 sheets. Drafting sheets at
>   **layout 5**, 43 rows / 43 SUBMIT each; `PROGRESS_BODY` holds its 43 approve ticks.
> - Review queue **open, nothing applied**: 32 rows ticked = the 19 INVISIBLE changes only
>   (punctuation and casing). The 36 visible ones are deliberately unticked.
> - Backups tonight: `PRE-ADDIN86-20260814-185040`, `GOOD-STATE-20260814-202633`,
>   `register-wide.PRE-RESTORE-MIGRATION-*`, `register-wide.PRE-TICKSET-*`.
>
> ### THE TOOLBAR IS THREE BUTTONS, SPLIT BY ARTIFACT
>
> `1. Set up my quarter` (workbook) · `2. Put it on the slides` (deck) ·
> `Review changes (writes nothing)`. **Neither side can trigger the other** — that coupling
> is what wiped 43 approve ticks in the morning. Seven dialogs became **one approval (the
> review tick) plus two selections**. The field picker is **deleted**, not improved:
> publish walks every Prose field. `Rebuild my sheets` is deleted — "press this when a
> sheet looks wrong" is a defect with instructions attached.
>
> ### WHAT COST THE EVENING, AND WHAT IT TAUGHT
>
> **`MigrateSheetLayout` wiped every drafting sheet on the real workbook at 19:11** — 129
> drafted paragraphs, 43 approve ticks, 75 notes. **All recovered**, 412 values verified
> through Excel against the backup, 0 overwritten.
>
> The defect was a **bootstrap error**: `sheetLayout` was read at `COL_D_LAYOUT` and
> `sheetPeriod` at `COL_D_PERIOD` — the CURRENT layout's stamp columns. *You need the
> layout to know where the layout stamp is.* On a layout-4 sheet those cells hold the
> prompt and nothing, so the version came back 0, `layoutMatches` went False, the
> whole-sheet clear fired, and the migration — guarded on the same flag — never ran.
> **Fixed:** `DetectLayoutFromRow` searches the intro row right-to-left; the period is then
> read at *that* layout's position.
>
> **The test that should have caught it encoded the bug.** Its fixture wrote layout-3 DATA
> under layout-5 STAMPS — a sheet that has never existed — and passed only because the
> reading code made the same mistake. That is why 192 green tests certified a defect that
> destroys real work.
>
> ### ALSO SHIPPED, EACH PROVEN BY A DELIBERATE BREAK
>
> - **The pairing cross-wiring hole.** `specs/deck-registry.md` claimed the deck↔workbook
>   pairing "closes the cross-wiring risk". It did not: the workbook's `DeckReference` GUID
>   was written once at onboarding and **never read for its purpose**. Now stamped on both
>   ends at repoint and checked before every register write. A blank stamp is NOT a
>   mismatch — every pre-existing register has one.
> - **Discovery is device-aware.** A group with slots is ONE candidate. The timeline used to
>   present as 21 fields to hand-tag.
> - **The test runner stopped closing your Office windows** and now clears more than one
>   husk (`Count -ne 1` meant one husk self-heals and two are permanent).
>
> ### WHAT ROHAN'S QUESTIONS DELETED — the pattern, again
>
> - *"It shouldn't have to ask"* → deleted the field picker, its wording, and the
>   Excel-focus bug just built for it.
> - *"What happens when I tag them?"* → reordered the plan. Tagging enables the HARVEST.
> - *"I don't have 26 duplicate slides"* → stopped a deck-integrity hunt. Read the deck:
>   43 `INSTANCE_KEY`, 44 `SLIDE_TYPE`, clean. The duplicates are 30 **queue rows** with
>   identical ChangeIDs.
> - *"I thought the only generative text was coming from drafting sheets"* → exposed that
>   the review queue shows every register↔slide difference regardless of origin. Of 55
>   queued changes, **one** came from tonight's drafting.
>
> ### OPEN, NAMED
>
> 1. **The harvest** — the 1,247 values. Everything queues behind it.
> 2. **`GAP 3` is hours, not a build.** `MilestoneDevice.bas:630` already writes
>    `shp.Visible`; the template already carries `MS1_ON/_NOW/_OFF` … `MS7`. Expose it as a
>    field behaviour. `EXPECTED-TRACE`'s "no `.Visible` write anywhere" is **wrong** and now
>    annotated as such.
> 3. **The review queue emits duplicate rows** — 85 rows for 55 changes, identical
>    ChangeIDs. Cosmetic; Rohan ranked it below connectivity.
> 4. **`PROJECT_STATUS` has no vocabulary enforcement.** 17 values out of list tonight,
>    reported and not written.
> 5. **`STRATEGIC_LINKAGES` now HAS a register column** (added tonight with 16 others) but
>    no values and no template tag.
> 6. The unsaved-workbook prompt before review always answers Yes — a save-and-report, not
>    a question. And something after publish dirties the workbook without saving.
>
> ### DOCUMENT CONTROL, DONE THIS SESSION
>
> `TOOLBAR.md`, `WORKFLOW.md`, `FIX-LIST.md` and `EXPECTED-TRACE-2026-08-14.md` were
> scanned and corrected where tonight made them false — button names, layout 4 → 5 column
> letters, and the wrong `.Visible` claim. **Historical records were annotated, not
> rewritten.** `NEXT-SESSION.md`'s older blocks quote past suite counts; those are dated
> observations and stay as they are.
>
> **Delivery count is 2.** Tonight's 8 `PROJECT_STATUS` writes applied ticks made at 16:51,
> for a field already delivered on 8 August. The machinery moved a long way; the count did
> not.
>
> ---

> ## 14 AUG, NIGHT (~19:00). **SUPERSEDED** by the block above.
>
> ### FIRST ACTION: SAVE `addin86`, THEN PRESS `1. Sync Now` AND WALK THE LOOP.
>
> `addin86` is imported and waiting on the one permanent manual step
> (`File > Save As > PowerPoint Add-in`). `addin85` is registered and loaded but
> **predates everything below it.** Nothing in this block is reachable until 86 is saved,
> ticked, and 85 is unticked.
>
> Deck and register are **backed up and md5-verified** at
> `backups/PRE-ADDIN86-20260814-185040/`.
>
> ### THE SIZING IN THE BLOCK BELOW WAS WRONG. ROHAN CAUGHT IT. DO NOT REPEAT IT.
>
> Earlier tonight this file's plan cut Rohan's five imagined steps ("A") from the weekend
> as **weeks** of work. He challenged it and he was right. **The estimate was made without
> opening the modules** — the failure this project has logged nine times.
>
> | the claim he made | what the source says |
> |---|---|
> | marking/confirmation runs against the template, manual **or** matrix review | `BatchOnboardFlow.BuildBatchPlan(templateSld, otherSlides())` and `BuildBatchPlanFromMarkedFields(...)` — **both exist and are tested** |
> | onboarding runs off slides matching the template | that IS `BuildBatchPlan` |
> | templates are provided at the start | `TemplateSlide.MakeTemplateFrom(...)` exists |
>
> **`EXPECTED-TRACE-2026-08-14.md` is WRONG on GAP 3.** It says there is "no `.Visible`
> write anywhere". `MilestoneDevice.bas:630` writes `shp.Visible` and it works — it is
> private to that device. The job is EXPOSING it as a field behaviour, not inventing it.
>
> **Revised: GAP 1 hours (spec + wiring), GAP 3 hours, GAP 2 already plumbing, GAP 4 days
> and deferrable.** The scoring doc itself said "3 of 5 substantially built" and this file
> still said weeks. **A handover document is a summary of the code and goes stale in hours.
> Read the module.**
>
> ### THE "PUBLISH ONE FIELD BY BUTTON" MILESTONE IS DISSOLVED, NOT SKIPPED
>
> Rohan: *"no we have proved we can push to slides over and over. finish this against
> target"*. S6 is proven repeatedly. The target loop IS tweak → sync → tweak → sync, so
> **the first real sync exercises S5 as a by-product.** Building a separate ceremony for it
> was making a step out of something that happens anyway.
>
> ### WHAT SHIPPED TONIGHT — suite 192/0, static clean, compile gate passed
>
> **1. The pairing defect: the GUID was written and never read.**
> `specs/deck-registry.md` claims the pairing is "mutually verifiable" and closes the
> cross-wiring risk. It did not. Deck stores the workbook PATH; workbook stores the deck's
> `DeckSyncId` GUID (path one way, identity the other — deliberate, since OneDrive breaks
> paths and not GUIDs). But the GUID was written once at onboarding and **nothing ever
> compared the two.** Only readers were a struct assignment and a `MsgBox` in a demo
> expecting the literal `"deck-v1"`.
> - `ExcelOutput.WriteDeckReference`/`ReadDeckReference` made **Public**
> - `DeckRegistry.StampPairing` — a repoint now writes BOTH ends, save-verified
> - `DeckRegistry.PairingProblem` + pure `PairingVerdict` — wired into
>   `PublishDraftsForField`, the write path into the register
> - **Blank is not a mismatch.** Every pre-existing register has a blank `DeckReference`;
>   refusing those would strand him at work. A DIFFERENT GUID refuses.
> - **Proven by two deliberate breaks**, each failing a different assertion by name.
>
> **2. Seven dialogs → one approval + two selections.** Deleted: "This will, in order…",
> "Write these into the register?", "Ready to build the list" (3-way), the Roll Forward
> confirm. **Bulk-approve is not lost** — it keeps its own toolbar entry, so it stays chosen
> by name rather than reachable by answering "No" to a question about something else.
> KEPT: quarter selection, field selection, review tick.
>
> **3. Two typed boxes became clicks, and one was a DATA HAZARD.** The field picker needed
> an exact match against a 30-item list that pushed its own text box off screen. The
> roll-forward source period was free text — and periods are matched EXACTLY, so `"q3f26"`
> or a quarter not present produces a clean run that copies nothing and reports success. The
> period is now read from a row the person clicks, so a typo is impossible **by construction
> rather than by validation**. Range-picker chosen over a UserForm because `build_ppam.ps1`
> imports `.bas` only.
>
> **4. `check_vba_static.py` caught a real defect in my own code** (a retry that recursed
> with identical arguments) hours after I called it "nearly worthless". Blind in one
> direction is not blind.
>
> ### THE PREDICTION FOR THE FIRST PRESS — recorded before it happens
>
> 1 quarter prompt · 2 ~~"Go ahead?"~~ gone · 3 roll-forward reports non-modally ·
> **4 drafting rebuild — the layout 4→5 migration AND the ferry, first time ever on a real
> workbook** · 5 click a row to pick the field · 6 ~~register confirm~~ gone ·
> 7 ~~"ready to build"~~ gone · 8 **review tick** · 9 apply.
>
> Step 4 is the only one that has never touched real files. If what he sees differs, **the
> diff is the finding** — take what he saw, not what was expected.
>
> ### STILL OPEN
>
> 1. **`STRATEGIC_LINKAGES` has no register column** — the whole output slide type depends
>    on it.
> 2. **`PROJECT_STATUS` drifted clean → dirty in three days.** A Controlled field with a
>    declared vocabulary has **nothing enforcing it**. Normalising fixes today; only
>    enforcement fixes next time.
> 3. **The template specification is now the critical path** — GAP 1 is hours of wiring
>    *given a spec*, so the spec is the bottleneck. Handed to chat.
>
> **Delivery count is still 2.** It moves on the next real sync, not on a ceremony.
>
> ---

> ## 14 AUG, EVENING. **SUPERSEDED** by the block above where they disagree.
>
> ### THE BUILD IS GREEN. THE QUARTER-TURN FERRY IS IN.
>
> Compile gate **passed** (proven: the gate blocks the suite entirely, and 191 tests ran).
> Suite **191 passed / 0 failed**. 191 = the previous 193 minus the two tests deleted below.
>
> ### FIRST ACTION: PRESS `Sync Now` AND PUBLISH ONE FIELD, END TO END, BY BUTTON
>
> Chat proposed this ordering and it is right. **The publish path has still never run
> through the tool** — the 43 `PROGRESS_BODY` values were written into the register by hand
> over COM. *Register -> slides* is proven; ***drafting -> register* is not.** Those are
> different claims and only one is true.
>
> Do it BEFORE restructuring the chain (rename, split, delete the three invariant prompts).
> If publish then fails, a restructuring bug is indistinguishable from the original defect.
>
> **Green does not mean reachable.** No test asks "can a person cause this to run". This
> project has shipped a tested picture injector, tested progress bars and a tested publish
> path all behind locked doors. Only the press answers it, and the answer is read from the
> SAVED FILE, never from a dialog.
>
> ### ROHAN'S RULE, 14 AUG: **ONE APPROVAL STEP, PLUS SELECTION.** NOT YET BUILT.
>
> *"there should just be one approval step (besides any selection)"* — after walking the
> chain himself and hitting seven dialogs.
>
> **The one approval is the REVIEW TICK.** That is where consent belongs: it is the only
> gate whose answer varies, and the only one standing between the register and his slides.
> Everything else becomes non-blocking and goes to the Run Log.
>
> **Dialogs to delete, in order met when pressing `1. Sync Now`:**
>
> | dialog | why it dies |
> |---|---|
> | `MS*` "fields with nothing to write into" | **DONE this session** — answer never varied; the 21 names are `MILESTONE_TIMELINE` internals, and "Yes" would have destroyed the device the warning pointed at. Now `WriteRunLog`. |
> | "This will, in order... Go ahead?" | Invariant. Orientation belongs on START HERE, not in the way. |
> | Roll Forward offer | Invariant when the destination already holds rows — it says so itself and then asks anyway. |
> | "Write these into the register for Q4F26?" | The register is not the deck. Nothing reaches a slide from it without the review tick. Backed up, reversible. |
> | "Ready to build the list of slide changes" | Pure ceremony in front of the real gate. |
> | "Nothing was changed" confirmations | Report, never a modal. |
>
> **KEEP:** the field selection (it is a selection, not an approval) and the review tick.
>
> **The field picker is a defect in its own right.** It is a typed `InputBox` requiring an
> exact match against a 30-item list, and the list is so long it pushes the text field OFF
> THE BOTTOM OF THE SCREEN. Rohan: *"I don't want to do any secret hidden typing, I need to
> select the field by clicking on it."* Make it a click.
>
> **COMPILED AND GREEN.** The `MS*` deletion is verified: compile clean across 33 modules,
> suite 191 passed / 0 failed.
>
> ---
>
> ### WHAT SHIPPED THIS SESSION
>
> - **Layout migration restored** (`MigrateSheetLayout`). The five `kept*` dictionaries WERE
>   the migration, as a side effect of clear-and-rebuild; deleting them deleted it, and a
>   layout-3 sheet was being read with current column numbers -- SUBMIT text becoming a
>   source ID, the AI draft landing in the column that publishes.
> - **The quarter-turn ferry.** `DRAFT_LAYOUT_VERSION` 4 -> 5, new `COL_D_PREV` at column 4
>   ("REPORTED LAST TIME"). On a period change SUBMIT moves sideways, the working columns are
>   handed to the new quarter empty, and the sheet is parked first because sources and notes
>   ARE cleared.
> - **The rollover refusal is DELETED** -- it was a deadlock. It told people to publish
>   first; publish never cleared what it objected to; the only `ClearContents` sat past the
>   exit. The tool could not start a new quarter. The per-row cadence machinery and the
>   `cadence` parameter went with it.
> - **Every column letter in headers and instructions is now DERIVED.** They were literals
>   and the one-column shift made all six name the wrong column while still reading correctly.
> - Tests: two deleted (`RolloverCadenceGovernsUntypedRows`,
>   `RefusesRatherThanDiscardOnPeriodChange` -- they asserted the deadlock), one replaced by
>   `Drafting_QuarterTurnFerriesSubmitIntoReportedLastTime`, which also asserts idempotence
>   on a second rebuild.
> - **Your register was repaired**: the two 11:40-damaged tabs had no layout/period stamp,
>   which would have triggered `Cells.Clear` on both. Re-stamped and verified from the file.
> - **The 27 lost paragraphs are RECOVERED** -- both sheets at 43 rows / 43 SUBMIT, verified
>   from the saved bytes. Restored by writing ONLY into cells empty in the live file, which
>   made the backup predating the 08:21 publish irrelevant.
>
> ### THE COMPILE ERROR, AND WHY IT MATTERS BEYOND ITSELF
>
> `ByRef argument type mismatch`. `DraftingUI:521` passed a bare positional `Nothing` into
> the removed `cadence` slot, so `Nothing` bound to `srcWs` and `srcWs` to a Long. The
> comment three lines above said *"No cadence argument"* and was READ INSTEAD OF THE CODE.
> **It described the argument BY NAME for a call that passes POSITIONALLY** -- true about
> intent, false about the call.
>
> ### KNOWN GAPS, NAMED
>
> 1. **The ferry has never touched a real workbook.** The test builds its own two-project
>    fixture. Your live sheets are layout 4 and will migrate to 5 on the next refresh -- that
>    is the first real exercise of BOTH new things at once. Back up first.
> 2. **No test for corrected text.** Chat's point: if column F is corrected mid-quarter the
>    ferry must carry the CORRECTED value. The code reads SUBMIT at the moment of the turn so
>    it should be right; nothing proves it.
> 3. **`STRATEGIC_LINKAGES` has no register column.** 32 columns and it is not one. The
>    output slide's chips are computable in principle and have **no input at sync time**.
>    Fold into the sixteen-column pass; it is the one that whole slide type depends on.
> 4. **`PROJECT_STATUS` drifted clean -> dirty between 11 and 14 Aug.** 17 rows, all Q3F26,
>    16 pure casing. The finding underneath: **a Controlled field with a declared vocabulary
>    has nothing enforcing it.** A defect in the control, not the data.
>
> **Delivery count is 2.** It moves when a drafted field reaches a slide through the tool.
>
> ---


> ## 14 AUG, LATEST. **FIRST ACTION: COMPILE `Drafting.bas` IN THE VBE AND READ THE DIALOG.**
>
> The reported-last-time build is committed and pushed and **THE COMPILE GATE IS RED.
> NO TESTS RAN. NOTHING IN IT IS VERIFIED.** `vba/Drafting.bas` carries a warning banner
> on line 1; delete it once green. Committed only because losing it was the bigger risk.
>
> The gate prints no line number — the VBE dialog does, and PowerPoint is left on screen
> holding it. Open it and read it. Suspects, in order: the deleted `droppedQuarterly` /
> `keptStatic` counters (report block ~line 1050), the removed `cadence` parameter from
> `WriteDraftingSheet` (a caller may still pass it), and the new ferry block in the row loop.
> `check_vba_static.py` says clean and CANNOT see this — it does not check block balance.
>
> ### FROM CHAT, 14 AUG 14:49 — TWO FINDINGS TO CARRY
>
> `OneDrive\Claude\chat-status-for-claude-code-2026-08-14-night.md`.
>
> 1. **`STRATEGIC_LINKAGES` HAS NO REGISTER COLUMN.** 32 columns and it is not one of them.
>    This is the reopener named in the chips reply and it is REAL: the output slide's chips
>    are computable in principle but have **no input at sync time**. It does not stop Rohan
>    drawing; it stops the slide being populated. Fold it into the sixteen-column pass — it
>    is not just another missing column, it is the one the whole output slide type depends on.
>
> 2. **`PROJECT_STATUS` drifted from clean to dirty between 11 and 14 Aug.** 17 non-conformant
>    rows, ALL in Q3F26 (Q4F26 and Q1F27 clean), 16 pure casing. Legacy, so normalising cannot
>    break anything live. **The finding underneath it matters more: a Controlled field with a
>    declared vocabulary has NOTHING ENFORCING IT.** That will recur, and it is a defect in the
>    control, not in the data.
>
> Also settled by Rohan: progress bars stay GREEN (retires the May traffic-light spec), and
> `Project Closed` is TERMINAL. Chat's operational claims are hypotheses to check against the
> source, not instructions — it has twice advised something impossible in the tool.
>
> ---
>
> ### WHAT THE BUILD DOES, once it compiles
>
> Rohan, 14 Aug: *"whenever a 1/4 changes at the top, the ferries belonging to that system
> deal with information for the new quarter. The last previous set that runs move info into
> 'reported last time' column."*
>
> **This REVERSED the earlier "derived from the register" decision, deliberately.** The
> ferry runs at the moment the update notices the period changed, so it needs no predecessor
> tracking, no second register read, and no period ordering — and it preserves text that was
> never published, which the derived version could not show.
>
> - `DRAFT_LAYOUT_VERSION` **4 → 5**, new `COL_D_PREV` at column 4; everything from SOURCES
>   rightwards shifts one right. `ColumnInLayout` gains an explicit `Case 4`.
> - **The rollover refusal is DELETED** — it was the deadlock. So is the per-row cadence
>   machinery and the `cadence` parameter.
> - On a quarter turn: SUBMIT moves to `COL_D_PREV`, the working columns are cleared, and
>   the sheet is parked first because sources and notes are genuinely destroyed.
> - Every column letter in the headers and instructions is now DERIVED. They were literals
>   and the shift made all six name the wrong column.
> - Intro rows 1–2 are wiped before rewrite so a layout bump cannot strand an old stamp.
> - Tests: `Drafting_RolloverCadenceGovernsUntypedRows` and
>   `Drafting_RefusesRatherThanDiscardOnPeriodChange` **deleted** (they asserted the
>   deadlock). `RolloverRebuildsOnlyWhenNothingIsAtRisk` **replaced** by
>   `Drafting_QuarterTurnFerriesSubmitIntoReportedLastTime`, which also asserts idempotence
>   on a second rebuild — none of it has ever run.
>
> **The invariant was restated, and this matters.** It said "if the carry dictionaries come
> back, so has the clear". Rohan: *"can't they just sate the need then run on update?"* They
> can. Holding values in a variable was never the defect; opening a gap was. The real
> invariant is **no `ws.Cells.Clear` on the normal path**.
>
> ### ALSO DONE 14 AUG, AND VERIFIED
>
> - Layout migration restored (`MigrateSheetLayout`), suite 192/1 at commit `a5156f8`.
> - **The 27 lost paragraphs are RECOVERED** — both sheets at 43 rows / 43 SUBMIT, read
>   back from the saved file. `PROGRESS_BODY` also carries 43 restored approve ticks.
> - The live register's two damaged tabs were re-stamped to layout 4 / `Q4F26`.
> - **Chat-side template handover written:**
>   `OneDrive\Claude\handover-to-chat-templates-2026-08-14.md`.
>
> **Delivery count is still 2.**
>
> ---


> ## 14 AUG, LATE. **SUPERSEDED** by the blocks above — it was current on 14 Aug
> and its banner was never demoted, so this file carried TWO blocks claiming to
> be current until the 14 Aug 23:10 block above. Kept for reasoning only.
>
> ### THE COMPILE IS DONE AND IT WAS GREEN
>
> `aff84d6`'s restructure **compiles clean, whole project, 33 modules**. That was the
> previous block's blocking first action and it is closed. Suite came back **191 passed,
> 2 failed** — both failures real, both diagnosed below.
>
> ### YOUR REGISTER WAS REPAIRED. VERIFIED FROM THE SAVED FILE.
>
> `TPL_KEY_EVENTS_BODY` and `TPL_PROGRESS_BODY` had **no layout stamp and no period
> stamp** — the 11:40 abort died before writing them. A blank K1 reads as layout 0,
> `ColumnInLayout(0,"SUBMIT")` returns 0, `layoutMatches` is False, and **`ws.Cells.Clear`
> would have fired on both sheets** at the next rebuild. Stamped back to `4` / `Q4F26`
> through Excel. Confirmed from the saved XML: both sheets `K1=4`, `M1=Q4F26`, nothing
> beyond column M, `maxrow=52` matching every healthy sheet, counts intact at 23 rows /
> 22 SUBMIT and 37 / 37. Backup: `backups/register-wide.PRE-STAMPREPAIR-20260814-132207.xlsx`.
>
> **A COM diagnostic written as read-only WROTE to the live register** — `Close($false)`
> cannot discard while AutoSave is on. The AutoSave hazard was already documented in this
> file for PowerPoint and was not carried across to Excel. It left a `Z999` probe cell on
> both sheets; removed and verified gone. Logged to FRICTION.md.
>
> ### DEFECT: THE LAYOUT MIGRATION WAS DELETED WITH THE CARRY DICTIONARIES. FIXED.
>
> The five `kept*` dictionaries were the layout migration, as a side effect of clearing
> and rebuilding. Deleting them deleted it. `layoutMatches` is
> `(ColumnInLayout(sheetLayout,"SUBMIT") > 0)` — **True for a layout-3 sheet** — so the
> clear never fires, the row loop reads the sheet with layout-4 numbers, and SUBMIT text
> becomes a source ID while the AI draft lands in the column that publishes. The sheet is
> then stamped forward, so it can never self-correct.
>
> **Fixed** by a new `MigrateSheetLayout` — in place, per row, read-all-then-write
> (3 -> 4 is a permutation of the same six columns), old positions emptied between, parked
> first. **No clear, no carry dictionaries: the structural proof that the clear is gone
> still holds.**
>
> **This was NOT urgent until it was.** Nothing on the live register is layout 3. But the
> "reported last quarter" column below is a layout bump to 5, which runs this exact path
> against sheets stamped 4. Fix had to land first.
>
> ### DEFECT: A ROLLOVER CANNOT HAPPEN AT ALL. NOT YET FIXED.
>
> Traced end to end, not inferred:
> - `Start a Quarter` sets the deck period; does not touch drafting sheets.
> - `Roll Forward` copies register rows; does not touch drafting sheets.
> - `Refresh Drafting Sheets` with `periodChanged` → at-risk scan finds typed work →
>   **REFUSED, nothing changed.**
> - The refusal says *"publish this sheet's work before rebuilding"*. `PublishDrafts` only
>   READS `COL_D_SUBMIT` and `COL_D_APPROVED` (`Drafting.bas:1123,1125`). It never clears
>   them. The next attempt refuses identically.
> - The only `ClearContents` for those columns is **inside the branch the refusal already
>   exited past**. No other route clears or deletes a drafting sheet.
>
> **Once a sheet holds typed work — the wanted state — it can never be rolled forward.**
> Predates tonight: the refusal landed 13 Aug in `ddf867b`. Tonight's FIX-LIST 1c only
> widened the trigger to SOURCES and NOTES. Never hit because no rollover has been
> attempted since. **It bites at Q1F27.**
>
> ### ROHAN'S ANSWER DISSOLVES IT — AND DELETES CODE
>
> *"an owner should be able to edit their quarter's slides by editing their quarter's
> spreadsheet and hitting sync. When the deck is set to a new quarter, items that were
> current last quarter but now need replacing will move to a 'reported last quarter'
> column in the drafting sheet, adding fodder for the AI prompt re style and narrative
> consistency."*
>
> If last quarter's text moves aside instead of being destroyed, **the refusal has nothing
> left to protect** — republishing is prevented structurally. That kills the refusal, the
> cadence machinery and the deadlock together, and adds prompt material. Same move as the
> in-place fix and the per-quarter split. `Drafting.bas:228` already states the principle:
> *"nothing is destroyed, only superseded."*
>
> **SETTLED: the column is DERIVED from the register, not moved from the sheet.** The
> register already holds it — 43 rows at `Q3F26`, 43 at `Q4F26`. Derived shows what
> actually reached a slide, which is the honest prompt fodder; moved would mix in drafts
> that were never published and make the heading false.
>
> **The one cost, named:** periods are free text with **no ordering** (`Q3F26` is just a
> string). `RollForwardPeriod` is told both periods by the UI and does not record the
> link. So deriving needs the predecessor stored at roll-forward time. That is the
> prerequisite, and it is small.
>
> ### THE SECOND SUITE FAILURE IS A DESIGN COLLISION, NOT A BUG
>
> `Drafting_RolloverCadenceGovernsUntypedRows` asserts the old rebuild-and-clear
> behaviour; FIX-LIST 1c's widened scan now refuses first, so its assertions never run.
> Rohan called it: the scan wins, delete the cadence machinery. **NOT DONE, deliberately**
> — that machinery holds the only `ClearContents` in the rollover path, so deleting it
> before the "reported last quarter" column exists would weld the deadlock shut. Delete it
> as part of that build, not before.
>
> Also stale prose: the refusal says *"N row(s) have drafted or submitted text"* when it
> now also fires on sources or notes alone. Machine fact written into prose.
>
> ### NEXT, IN ORDER
>
> 1. ~~**Recover the 27 texts.**~~ **DONE 14 Aug, verified from the saved file.** Both
>    sheets are back to **43 rows / 43 SUBMIT**, nothing missing, every text byte-identical
>    to the backup. The 27 = 20 appended rows on `KEY_EVENTS_BODY` + `4_K017`, whose row
>    survived with an emptied SUBMIT (309 chars), + 6 appended rows on `PROGRESS_BODY`.
>
>    **The rule that made a wholesale restore unnecessary: a cell was written ONLY when
>    empty in the live file.** Nothing existing could be overwritten by construction rather
>    than by care, so the backup predating the 08:21 publish stopped mattering. Backup:
>    `backups/register-wide.PRE-RESTORE27-20260814-134147.xlsx`.
>
>    **`PROGRESS_BODY`'s 43 approve ticks were restored too** — they were wiped at 11:40 and
>    every SUBMIT on that sheet is byte-identical to what they approved, so this restores
>    consent rather than inventing it. `KEY_EVENTS_BODY` has 0; its backup had none. Flagged
>    to Rohan as a consent gate touched without asking.
> 2. **Build the "reported last quarter" column**: store the predecessor period at roll
>    forward, derive the column, bump `DRAFT_LAYOUT_VERSION` to 5, delete the refusal and
>    the cadence machinery, feed the prior text into the L2 prompt.
> 3. **Scenario 5**, once the sheets are whole.
>
> **Delivery count is still 2.** Nothing reached a slide on 14 Aug evening.
>
> ---
>
> ## 14 AUG, EVENING — SUPERSEDED by the block above. Kept for reasoning.
>
> ### FIRST ACTION: COMPILE AND RUN THE SUITE IN POWERPOINT. NOTHING ELSE.
>
> `vba/Drafting.bas` carries a substantial restructure that is **committed and pushed
> (`aff84d6`) but has NEVER BEEN COMPILED**. Everything else waits on that. It is
> committed because losing it was the bigger risk, not because it is trusted.
>
> **Do not trust `check_vba_static.py` as compile evidence.** It was handed a
> `Drafting.bas` containing a deliberate `If True Then` with no `End If` and printed
> `static checks clean across 34 module(s)`. It has a function named
> `check_structural_sanity` that does not check block balance. Proven by breaking it
> on purpose, 14 Aug.
>
> If the compile is red it is almost certainly one of five edits, each isolated and
> easy to bisect (listed below). A known-good copy of the file as it stood before the
> harvest deletion is at
> `scratchpad/Drafting.bas.pre-harvest-delete` — likely cleaned up; the real fallback
> is `git checkout vba/Drafting.bas`.
>
> ---
>
> ### WHAT HAPPENED: THE 11:40 RUN DESTROYED DRAFTED TEXT, NOT JUST TICKS
>
> The morning block below says *"No text was lost."* **That was wrong.** It counted the
> 22 and 37 SUBMIT cells that survived and never compared them to the 43/43 baseline
> sitting in that morning's own backup.
>
> | sheet | 14 Aug 07:16 | after 11:40 | lost |
> |---|---|---|---|
> | `TPL_KEY_EVENTS_BODY` | 43 rows / 43 drafted | 23 / 22 | **20 rows, 21 texts** |
> | `TPL_PROGRESS_BODY` | 43 rows / 43 drafted | 37 / 37 | **6 rows, 6 texts** |
>
> **Named, because a count is not evidence.** `KEY_EVENTS_BODY` lost every S-coded
> project plus `4_K021`, `1_K022`, `3_K023`. `PROGRESS_BODY` lost a strict subset:
> `1_S018, 3_S019, 3_S020, S021, S022, S023`.
>
> **Established from the files:**
> - Both sheets stopped **mid-function**. Rows present; prompt, layout stamp and period
>   stamp all absent — and all three are written *after* the row loop.
> - Truncation is contiguous at the tail. KEY_EVENTS stopped after register row 72,
>   PROGRESS after row 86 (the Q4F26 block spans physical rows 50–92).
> - **The register never changed.** 92 rows in all four backups and live.
> - **Not a cell-length problem.** Longest value 609 chars against Excel's 32,767. This
>   hypothesis was tested and is dead.
>
> **NOT established: why the write stopped where it did — AND DO NOT GO LOOKING.**
>
> Rohan, 14 Aug: *"please don't blow my budget on mysteries that are not part of the
> existing architecture and current decision making."* He is right. The in-place fix
> bounds this failure mode — a mid-write abort now leaves untouched rows untouched — so
> the cause is a curiosity, not a blocker. It is a one-off on a specific machine state
> and may never recur.
>
> **Reopen it only if it happens again**, and then with the evidence in hand rather than
> by hunting. Do not write a speculative cause into this file.
>
> **RECOVERY STILL OUTSTANDING.** The 27 texts exist in
> `backups/register-wide.PRE-ABOUTFIX-20260814-071658.xlsx`. **A wholesale file restore
> is wrong** — that backup predates the 08:21 PROGRESS_BODY publish. The rows must go
> back through Excel once the rebuild is proven safe. A fresh full backup of both files
> is at `backups/PRESCENARIO5-20260814/`, md5-verified against the originals.
>
> ---
>
> ### THE FIX: THE DRAFTING SHEET IS UPDATED IN PLACE, NOT REBUILT
>
> Rohan, after being shown that the park-before-clear change made the destruction
> survivable rather than removing it: *"why is clear still happening? For a fresh
> project or new quarter?"* He was right — it was patch six in a family of five.
>
> **`ws.Cells.Clear` now runs in exactly one branch: `If Not layoutMatches`** — a column
> renumbering, which has happened three times in the tool's life
> (`DRAFT_LAYOUT_VERSION` is 4). Everything else updates in place.
>
> **The proof condition, and it is structural rather than a claim:** the five `kept*`
> carry dictionaries are **deleted**. They existed only to ferry a person's work across
> the gap `Cells.Clear` opened. *If they ever come back, so has the clear.*
>
> The five edits, in the order they were made:
> 1. Park before every clear, not only on a layout mismatch (the old guard was gated on
>    `strandedRows > 0`, counted only inside `If Not layoutMatches` — so the ordinary
>    case took **no copy at all**, which is why 11:40 had no archive).
> 2. Layout + period stamps written **immediately after** the clear, not at the end. A
>    mid-write failure used to leave rows with no stamp, which reads as "unknown layout"
>    next time — and unknown layout carries nothing. A crash armed the *next* run to
>    discard everything that survived.
> 3. At-risk scan now counts **SOURCES and NOTES** (FIX-LIST 1c).
> 4. FIX-LIST 1d's late `ParkSheetCopy` call **removed** — it copied the already-cleared
>    sheet and reported "nothing was lost".
> 5. Rows addressed **by project code, not by position** (`rowOf` index, `appendAt`
>    cursor); human columns never written; harvest block and dictionaries deleted;
>    `r = appendAt` restored after the loop for the chrome's benefit.
>
> **Behaviour now:** an existing project keeps its row and its typed columns are not
> touched at all. A new project appends one row at the bottom. On a **period change**
> only, and per row, the work columns are cleared — the one case that must still
> destroy, because last quarter's text must not be republishable as this quarter's.
>
> **Known loose end:** `lostWithContent` (line ~530) is now declared and never
> incremented — dead, harmless, mine to clean.
>
> ---
>
> ### THE COLUMN CONTRACT, SETTLED 14 AUG — do not re-derive it
>
> - **Nothing in the add-in ever writes column E (AI DRAFT).** The only two references
>   are the header text and the rollover clear. Copilot writes it, externally, from the
>   prompt the tool puts in `L2`. There is no AI call inside the tool.
> - **Column F (SUBMIT) is written in exactly one place** — `CopyAiToSubmit`,
>   `Drafting.bas:1272` — and it skips any row where F already holds text.
> - **Publish reads F and G. It never reads E.** That is what makes "NEVER published"
>   a mechanism rather than a label.
>
> So AI text reaches the register only through the person's own column, via an explicit
> press that cannot overwrite them, plus a tick.
>
> ---
>
> ### ROHAN'S IMAGINED USE OF THE TOOL — captured before scenario 5 ran
>
> **`EXPECTED-TRACE-2026-08-14.md`** (untracked, in the repo root). His five steps in
> his own words, plus the diff against what the code does. Written down first
> deliberately: without a recorded prediction, whatever the tool does will look like
> what it was supposed to do.
>
> Scoring: **3 of 5 substantially built, 1 half, 1 absent.** The four gaps:
> 1. The template does not produce the workbook — direction is reversed (this is the
>    13 Aug template-first pivot, decided and never built).
> 2. **There are TWO approval gates and he imagines one.** Publish is per-field
>    (`FieldForRun` asks which); the compiled view he describes is the review queue
>    behind it. **He never once mentions choosing a field.** Very likely the root of the
>    tick confusion.
> 3. **Shape visibility is entirely unbuilt.** No `.Visible` write anywhere; `Behaviour`
>    is fill/fit/as-is for pictures only. The only one of the five with zero machinery.
> 4. **He expects a file per quarter.** **SETTLED 14 Aug — see `DECISIONS.md`.**
>
> ### GAP 4 IS SETTLED: ONE FILE PAIR PER QUARTER, PLUS A DERIVED CENTRAL REGISTER
>
> Deck **and** register split per quarter. A central register for multi-period analysis
> is **derived and never authored** — rebuilt by reading the quarter files, read-only,
> nothing ever flows back out of it into a quarter. Rohan's addition, and it closes the
> only real objection to splitting.
>
> **The cost was checked, not assumed, and it is nearly zero.** All twelve register
> reads are `ReadSheetForDeckPeriod(ws, <the deck's CURRENT period>)`. Exactly one
> function crosses periods: `RollForwardPeriod`. Column C reads the current period, not
> the previous one — the feared cross-quarter reads **do not exist**. An earlier version
> of this file named that as the unpriced cost; that was wrong.
>
> **The actual argument** is that the drafting sheets live *inside* the register
> workbook. If the register does not split, they do not split, and clearing last
> quarter's work out of a live sheet stays a normal-path operation — the operation that
> lost 27 paragraphs on 14 Aug. Per-quarter, the old quarter survives *by construction*.
> Same move as the in-place fix, one level up.
>
> **Not designed yet:** the aggregate's rebuild trigger and location; and the naming
> convention, which stops being cosmetic — `GetWorkbookPath` is already a bug source
> ("could not open the paired workbook", three causes in one morning) and more files
> means more pairing, on a machine with no Claude, Python or WSL.
>
> ### GAP 2 IS SMALLER THAN IT LOOKED — plumbing, not architecture
>
> Checked 14 Aug: **the review queue is already matrix-shaped.** `ChangeHash` is keyed
> per entity *and* field, so the compiled projects-x-fields view Rohan describes already
> exists downstream. The per-field bottleneck is only in **publish** (`FieldForRun`
> asking "which field?"), sitting in front of a machine that already handles the matrix.
> So this is closer to "publish across all fields" than to building a new surface.
>
> ### ON SHAPES vs TEXT BOXES — asked 14 Aug, answered from the code
>
> Injection already handles four kinds: text (`InjectPrimitive`), pictures
> (`InjectPictureVia`), progress bars (`InjectProgressVia`) and **devices**
> (`InjectDeviceVia` — a group consuming several register columns as one addressable
> thing). Visibility on/off is the one with no machinery at all (GAP 3).
>
> **Writing is solved; RECOGNISING is not.** Injection treats a device as one thing;
> discovery, wiring, marking and the audit all see its parts — which is why the timeline
> appeared as 21 candidate fields. That is Rohan's "load shape modules as prenamed per
> slide entities", and it is the device registry decided 13 Aug.
>
> ---
>
> ### DEFERRED, DELIBERATELY
>
> - **Scenario 5 has not run.** It should not, until the sheets are recovered and the
>   rebuild is proven — the chain rebuilds drafting sheets as a step, which is the
>   operation that caused this, and the baseline is currently 27 rows short.
> - **The rename** (`RefreshDraftingSheets` → `BuildDraftingSheets`) and the two-button
>   split. Reasoning is sound and recorded below; blocked on GAP 4.
> - The two-button shape agreed with Rohan: the boundary is **where he stops typing**,
>   not workbook-vs-deck. `[1. Start the quarter]` / `[2. Put it on the slides]`, one
>   add-in, count unchanged. Chat side had no quarrel with it.
>
> ### PROCESS NOTE
>
> Rohan's plain questions did the work again tonight, twice: *"aren't they permanent
> now?"* exposed the clear-and-restore shape, and *"why is clear still happening?"*
> caught the sixth local patch being applied to the same defect. Two FRICTION entries
> logged (the wrong "no text was lost" claim, and the static checker's blind spot).

> ## 14 AUG, ~12:00 — READ THIS BLOCK FIRST. It supersedes the 10:15 block below.
>
> ### STATE
>
> - **`addin84` IS LOADED.** Verified from the registry, not from a dialog:
>   `HKCU\...\Office\16.0\PowerPoint\AddIns\addin84`, `AutoLoad=1`, and **`addin83` is
>   gone**. Build stamp `2026-08-14 10:09`. File is byte-identical in `OneDrive\Claude\`
>   and the trusted `AppData\Roaming\Microsoft\AddIns\` (md5 `2b5514d6…`).
> - **The tick fix is committed and pushed** — `644ed16` on `main`.
> - **Delivery count is still 2.** No new field reached a slide today.
>
> ### WHAT HAPPENED AT 11:40, AND WHAT IT COST
>
> Rohan pressed `1. Sync Now` **while `addin83` was still the loaded add-in**. The chain
> rebuilt the drafting sheets and **wiped every approve tick on the two sheets it
> rebuilt** — defect 1 demonstrating itself on the real workbook, an hour after being
> fixed in source. Read from the saved file:
>
> | sheet | SUBMIT cells | APPROVE cells (data rows) |
> |---|---|---|
> | `TPL_KEY_EVENTS_BODY` | 22 | **0** |
> | `TPL_PROGRESS_BODY` | 37 | **0** |
> | the other eleven `TPL_*` | 43 | 43 present |
>
> **No text was lost. The deck was never touched** (still 08:21). **Nothing was
> published** — the Sync Log's newest entry is still `2026-08-14 08:21`, and that check
> was calibrated: the same search finds `07:48` and `08:21`, so its silence about 11:xx
> is evidence rather than absence.
>
> **NO BACKUP WAS TAKEN before that rebuild.** That is FIX-LIST 1d — the park runs
> *after* `ws.Cells.Clear`, not before. Still unfixed.
>
> **TO RE-TICK: `TPL_KEY_EVENTS_BODY` and `TPL_PROGRESS_BODY`.** Those two are certain.
> The other eleven have approve cells PRESENT, which is **not** the same as holding `Y` —
> a cell containing `0` counts identically, and this project has already recorded one
> wrong conclusion from exactly that (FIX-LIST P6). Unverified; check column G by eye.
> Under `addin84` those ticks now survive a same-quarter rebuild.
>
> ### ARCHITECTURE DECIDED THIS SESSION — Rohan's calls, all four
>
> 1. **The tick default INVERTS.** Everything auto-ticked at drafting; you untick what
>    you do not want shipped. Writing the text is the declaration; the exception should
>    carry the marking cost, not the majority case. **Two conditions attached:** the
>    review grid keeps FIX-LIST 3's conditional pre-tick (pre-tick only where the slide's
>    current text still exactly matches what the register last wrote, so a hand-edited
>    slide still demands a read); and the period reset becomes load-bearing rather than
>    incidental, because auto-tick plus a rollover would otherwise re-approve last
>    quarter's prose. Today's fix already drops the tick with the text on a period
>    change — that property now needs a test that says so out loud.
>
> 2. **THE CHAIN SPLITS BY ARTIFACT, NOT BY STEP.** One set of actions touches the
>    workbook, one touches the deck, and **neither can trigger the other**. This reverses
>    the 2026-08-09 two-button chain decision. The chain was right about the problem
>    (orientation — "which step am I up to?") and wrong about the remedy (removing the
>    choice), and the coupling it created is what wiped the ticks above: a person pressed
>    a button to PUBLISH and it REBUILT first.
>
> 3. **The deck side runs fully independently** of whether the workbook side has just
>    run. It reads the register as it stands on disk. This is what makes each half
>    testable alone — scenario 5 currently cannot be exercised without walking scenarios
>    1 and 2's machinery first.
>
> 4. **Review is its own action, not a step inside something else** — with its purpose
>    AND its exclusions stated ("for reading current vs proposed and ticking; not for
>    writing to slides"). Rohan: *"clear what it is and isn't for, part of a sequence, or
>    not."*
>
> ### THE ONLY THING TO BUILD NEXT — PHASE 1, AND NOTHING ELSE
>
> **Split the chain into independent buttons, and delete the three invariant prompts.**
>
> This is mostly DELETION. `RefreshDraftingSheets`, `PublishDraftsForField`,
> `ReviewChangesCore` and `ApplyApprovedCore` already exist as independent Subs;
> `RibbonUI.SyncNowChainCore` is a wrapper that calls them in order, and
> `CommandBarUI.AddButton` is called exactly twice. Removing the wrapper and adding
> buttons is small and subtractive.
>
> **DEFERRED, DELIBERATELY — do not build these as part of phase 1.** Scope grew in
> consecutive rounds this session and was cut back at Rohan's prompt (*"are we
> complicating this or making it simpler"*): self-describing action objects, promoting
> the `Readiness` surface as an orientation map, Excel-hosting the workbook half as an
> `.xlam` (FIX-LIST 8 — needs the deck's path and period written into the workbook's
> custom properties first), template-first, the device registry. All still wanted. None
> before scenario 5 has been walked once.
>
> ### PROCESS NOTE, WORTH KEEPING
>
> **Four rounds of chain design happened and the delivery count did not move.** Rohan
> asked "are we complicating this?" — the question this session should have asked itself
> at round three. The tell was scope growing in consecutive rounds while nothing new
> could be demonstrated. Consider running `deck-sync-pm` before the next design round.
>
> ---
>
> ## 14 AUG, 10:15 — THE APPROVE-TICK DEFECT IS FIXED AND PROVEN.
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
> **SUPERSEDED — `addin84` has since been saved, installed and confirmed loaded; see the
> block above.** As written: build stamp `2026-08-14 10:09`, all 32
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
