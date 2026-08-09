# Five buttons, four chains

Target: **five buttons on the PowerPoint toolbar, no capability deleted.**

The mechanism is a CHAIN, not a menu. A chain reads the current state, runs the
sequence, and stops only where the person has work to do or a decision to make.

The rule is already in this codebase, written by Rohan at `BatchOnboardFlow.bas:2886`:

> a boundary earns its place only where a person has to do work or make a decision
> in the gap.

A numbered menu puts a boundary at every stage, including the ones with no decision
in them. That is filing capability, not merging it. **Number of stops = number of
real decisions**, and every stop below names the decision it exists for.

---

## Why four chains and not one

`Readiness.bas:9-20` states the shape of the work, and it is not a ladder:

> The toolbar numbers its buttons 0, 0b, 1, 2, 3, 4 and that ordering is a LIE about
> the shape of the work. The true shape is: a deck-level prologue -> N per-field
> lanes that do NOT wait for each other -> a deck-level sync.

`RefreshDraftingSheets` builds every prose field's sheet at once and is safe at any
time; `PublishDraftsForField` publishes ONE field. So `ABOUT_BODY` can be published
and synced while `KEY_EVENTS_BODY` is half drafted.

A single 1-to-4 chain would re-tell that lie in a new form. So:

| Button | Chain | Runs |
|---|---|---|
| `Where am I?` | none — it is the read | any time |
| `1. Start the quarter` | **A — deck prologue** | once per quarter |
| `2. Draft and publish` | **B — one field lane** | once per field, lanes independent |
| `3. Put it on the slides` | **C — deck sync** | once per sync |
| `Set up a slide type` | **D — setup** | once per slide type, ever |

Chain B is scoped by its field and must never warn about another field's progress.
That independence is the point of the whole shape.

---

## Every chain branches on `Readiness.Build`

`Build(pres, wb)` already computes exactly the preconditions the chains need:
period on disk, period-not-saved drift, paired workbook, register readable, slide
type registered, template slide present and marked, register sheet present, rows for
the period, deck/register parity. Each returns `ok` / `BLOCKED` / `CANNOT TELL`.

**One `Build` per button press.** Recompute only after a stage that writes. The
chains and the `START HERE` sheet are then two renderings of one computation and
cannot disagree.

`Readiness.bas:51` still governs: **it offers, it does not gate.** A chain may
report a `BLOCKED` line and stop, but no button is ever disabled.

---

## Chain A — `1. Start the quarter`

| Stop | Condition | The decision |
|---|---|---|
| A1 | period not set on disk | **Which quarter?** — input nothing can derive |
| A2 | 0 rows for the period, previous period has rows | **Copy N rows forward from `<prev>`?** — he may want to start empty |

No other stop. Building the drafting sheets has no decision in the gap, so the chain
does it and reports it.

Period is written with `SetDeckPeriodVerified` and confirmed **from the saved file's
bytes** — never from PowerPoint's cache. That defect is why this project has the rule.

Everything already done:

```
Q3F26 is set up. 43 rows, 8 drafting sheets built 9 Aug.
Nothing to do here. Next: press "2. Draft and publish".
```

---

## Chain B — `2. Draft and publish`

| Stop | Condition | The decision |
|---|---|---|
| B1 | field not derivable from the active sheet | **Which field?** — scopes the whole chain |
| B2 | always, before writing the register | **The publish confirmation** — the consent gate |

The field resolves ONCE at entry (`ActiveDraftField`, else `AskForField`) and becomes
the title of every dialog in the chain: `2. Draft and publish — ABOUT_BODY`.

That is FIX-LIST 1a fixed structurally rather than by remembering: the subject is
carried by the container, so a message cannot report a count without it. The worst
1a instance — *"Column F is empty for all 43 row(s)"* fired against the wrong sheet —
becomes impossible, because the chain owns the sheet it is working on.

Copying AI drafts to submit has no decision in the gap. The chain does it.

**B2 is not a formality.** Per the ratified model the tick is a SELECTION and the
publish confirmation is the CONSENT GATE. It names the field, the row count, and the
projects with nothing to publish.

---

## Chain C — `3. Put it on the slides`

The only button that changes slide text.

| Stop | Condition | The decision |
|---|---|---|
| C1 | a review sheet holds ticks not yet applied | **Apply those / start fresh and lose them / cancel** |
| C2 | otherwise | **The write-authorising confirmation**, capped, detail in the Run Log |

**C1 is the safety property of this whole design.** `ReviewChangesCore`
unconditionally calls `WriteQueueSheet`, which rewrites the sheet. A chain that
opened by rebuilding would destroy an evening's approvals — strictly worse than
today, where `Apply Approved` is picked directly. The chain must DETECT and branch,
never rebuild first.

Closes FIX-LIST item 2. `ContentKindOf` hardcodes three field names and returns
`KIND_PROSE` for everything else; `AssignBatches` batches only `KIND_CONTROLLED`, so
`HasBatchableWork` is never true for `ABOUT_BODY` and today's button 4 ALWAYS ends in
a `vbExclamation`. Under a chain, prose going to the review sheet is simply the
normal route: drop the exclamation, name the sheet via `ReviewSheetNameFor`, delete
`REVIEW_SHEET_NAME`.

`Review + Approve All` survives as a choice inside C2, never the default, and
`ReviewChangesCore` still prepends the `APPROVE-ALL` banner to the report and the Run
Log. `RibbonUI.bas:573` argues for a separate button so bulk approval stays "a
decision taken each time and visible in the report" — the banner is the mechanism
doing that work, not the button.

Ends with `PersistBothFiles` / `SaveDeckVerified` / `SaveWorkbookVerified`.

---

## Chain D — `Set up a slide type`

Discover fields -> mark -> onboard slides -> check coverage, chaining as onboarding
already does. Its stops are inherent: marking a field requires selecting a shape on
the canvas, and VBA modals block the whole application, so each mark is its own act
(`MarkFieldForBatchCore:1261`).

**Do FIX-LIST item 5 with this chain.** `MarkShapeForBatch:1112` stores `fieldType`
and `fieldVolatility` in memory, never writes them to a shape tag, and the type's own
declaration calls them *"human-declared hint only, not wired into sync behavior yet."*
Nothing reads them. Deleting those two prompts takes marking from **3 modals per
field to 1** — on a 52-field pass, 104 dialogs that change nothing.

`Check coverage` is a first-class choice at the top of this chain, not something
reached by re-running onboarding — `CommandBarUI.bas:151` is right that a read-only
diagnostic you can only reach by re-running a setup step is one nobody will run.

---

## The three hazards a chain creates and a menu did not

**1. A chain can do more than the person expected.**
Every chain shows its plan before its first write — the stages it will run, named,
with the state each was derived from — and the end report lists every write that
happened. Respect `REPORT_CAP = 900` (`RibbonUI.bas:7`) via `CapReport`, detail to
the Run Log. A chain that writes silently is this project's defining defect wearing
a new hat.

**2. Cancel must leave a describable state.**
Every stop sits BEFORE a write, so Cancel leaves the state from the last completed
stage. The chain says which: *"Stopped. Q3F26 is set and 43 rows were copied
forward. The drafting sheets were not built. Press this again to finish."*
Re-pressing resumes, because the chain reads state rather than remembering it.

**3. A chain must never rebuild a surface holding human ticks.**
See C1. This is the one place the design can be worse than what exists today.

---

## The real work item: 37 strings, not 5 buttons

28 instruction strings across seven modules tell the person to press a button by
caption (`RibbonUI` x7, `BatchOnboardFlow` x5, `DraftingUI` x4, `ReviewQueue` x3,
`DiscoverUI` x2, `AdoptFlow`, `RunSync`, `WorkbookBridge` x1 each), plus 9
`Readiness` remedy values. Every one names a caption that ceases to exist.

Fix as a CLASS with a grep, per FIX-LIST item 6 — this is the same defect twice
already. `Readiness` today offers **"Create Template Slide"** as a remedy, naming a
button that does not exist on the current toolbar (`Readiness.bas:205,209`). That is
the failure this redesign exists to prevent, already live.

`Remedy` should become an enum the code switches on, never free text, so a remedy
naming a dead button becomes a compile-time impossibility rather than a string.

---

## What this breaks

1. **`Test_CommandBarUI_EveryDeclaredCapabilityHasAButton`** (`TestRunner.bas:6013`)
   hard-asserts 16 Subs have buttons; 13 will not. Split it: the runtime half
   asserts the 5 dispatchers are wired and visible; the REACHABILITY half — each
   demoted Sub is called from its chain — moves to `check_vba_static.py`. Be honest
   about the cost: that is Python, so **dev-machine only**, and the strongest guard
   against silently orphaning a capability gets weaker at work.
   Also `Test_CommandBarUI_ShowToolbarCreatesWiredButtons`: 16 -> 5, one bar.
2. **`HideToolbar` must keep deleting the three old bar names** so upgrading from
   build 40 leaves no orphan bar (`CommandBarUI.bas:226` already documents this).
3. **`Readiness.Build` per press, and `PropertyOnDisk` copies the deck to unzip it.**
   Sub-second on the 49MB deck — but that was a PowerShell probe of the TECHNIQUE,
   not this code at chain frequency. UNVERIFIED at 5-10 calls an evening.
4. **Assumes one slide type per deck.** `ResolveRegisterSheet` refuses more than one,
   so the type pickers auto-select (FIX-LIST item 7). A two-type deck brings them
   back and adds a stop to chains C and D.
5. **`TooltipText` raises past 255 chars and takes the whole toolbar down mid-build**
   (`CommandBarUI.bas:259`). Five tooltips, all under.

Per the standing rule: **make every new assertion fail on purpose once** — delete a
stage call and watch the static check go red — before trusting any of it.
