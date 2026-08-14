# The toolbar: three buttons, three chains

> **DESIGN DOC. ITS UI DESCRIPTION IS SUPERSEDED — rewritten 2026-08-14 (night).** The toolbar is **three
> buttons again** -- but not the three below. It now splits by ARTIFACT, not by step:
>
> | caption constant | caption | action |
> |---|---|---|
> | `CAP_SET_UP_QUARTER` | `1. Set up my quarter` | `RibbonUI.SyncNowChain` -- workbook side |
> | `CAP_PUT_ON_SLIDES` | `2. Put it on the slides` | `RibbonUI.PutItOnTheSlides` -- deck side |
> | `CAP_REVIEW_ONLY` | `Review changes (writes nothing)` | `RibbonUI.ReviewChanges` |
>
> **NEITHER SIDE CAN TRIGGER THE OTHER.** The boundary is where the person stops typing.
> That coupling is what wiped 43 approve ticks on 2026-08-14: a press meant to PUBLISH
> rebuilt the sheets first.
>
> `CAP_REBUILD_SHEETS` and its button are **deleted**. Its tooltip said "use this when a
> drafting sheet looks wrong", which is a defect with instructions attached rather than a
> capability; `RefreshDraftingSheets` still runs inside button 1.
>
> **The reasoning below is still the governing rule**, and the chain it criticised has now
> been cut to ONE approval (the review tick) plus two selections. The field picker it
> would also have condemned is deleted outright: publish now walks every field.


Was 16 buttons across three bars. This is the design that replaces them, and
nothing is deleted -- every capability is reached, just not from the bar.

## The rule the whole thing follows

Rohan, `BatchOnboardFlow.bas:2886`:

> a boundary earns its place only where a person has to do work or make a
> decision in the gap.

A **chain** reads the current state, runs the sequence, and stops only at those
points. A **menu** puts a boundary at every stage including the ones with no
decision in them -- that is filing capability, not merging it, and it was
proposed and rejected. **Number of stops = number of real decisions**, and every
stop below names the decision it exists for.

Applied to the bar itself, the same rule removes buttons: if pressing a button
is not a decision, it should not be a button.

## Why three

`Readiness.bas:9-20` states the shape of the work, and it is not a ladder:

> The toolbar numbers its buttons 0, 0b, 1, 2, 3, 4 and that ordering is a LIE
> about the shape of the work. The true shape is: a deck-level prologue -> N
> per-field lanes that do NOT wait for each other -> a deck-level sync.

| Button | Chain | Runs |
|---|---|---|
| `1. Start the quarter` | **A — deck prologue** (absorbs setup) | once per quarter |
| `2. Draft and publish` | **B — one field lane** | once per field, lanes independent |
| `3. Put it on the slides` | **C — deck sync** | once per sync |

Chain B is scoped by its field and must never warn about another field's
progress. That independence is the point of the shape.

## The two buttons that were removed, and why

**`Where am I?` is gone.** It rebuilt a readiness sheet you had to remember to
press. Every chain now opens by reading `Readiness.Build` and stating what it
found, so the check arrives at the moment it matters instead of being a separate
act. This also removes a hazard the module warns about in its own header -- a
readiness sheet is *"designed to be the thing consulted INSTEAD of checking...
it teaches a person to stop looking."* A button made that more likely.

To look without starting anything: press a chain and cancel at its first stop.
It leads with what it found and writes nothing until told.

**`Set up a slide type` is gone as a button.** It is a PRECONDITION, not an
activity: it runs once ever per slide type, and only on a deck that has none.
Chain A detects that and walks it. On a configured deck the chain never mentions
setup, because there is no decision in that gap.

**The risk, named rather than glossed:** a first run where "Start the quarter"
becomes a 52-field marking session is a surprise. Chain A must therefore state
what it is about to walk, and how long it is, BEFORE the first step -- and
`Check coverage`, which `CommandBarUI.bas:151` correctly argues is a RECURRING
read-only diagnostic and must not require re-running setup, becomes what chain A
offers when the deck is already configured.

## Every chain branches on one computation

`Readiness.Build(pres, wb)` already computes the preconditions: period on disk,
period-not-saved drift, paired workbook, register readable, slide type
registered, template present and marked, register sheet present, rows for the
period, deck/register parity. Each returns `ok` / `BLOCKED` / `CANNOT TELL`.

**One `Build` per button press.** Recompute only after a stage that writes.

`Readiness.bas:51` still governs: **it offers, it does not gate.** A chain may
report a `BLOCKED` line and stop, but nothing is ever disabled.

This only became safe on 2026-08-09. Before that, `Build` reported `BLOCKED` for
a cloud-hosted deck whose file it had never read -- so chains branching on it
would have insisted you re-set a period that was already correct.

---

## Chain A — `1. Start the quarter`

| Stop | Condition | The decision |
|---|---|---|
| A0 | no slide type registered | **Walk setup now?** -- named, and sized, before it starts |
| A1 | period not set on disk | **Which quarter?** -- input nothing can derive |
| A2 | 0 rows for the period, previous period has rows | **Copy N rows forward from `<prev>`?** -- he may want to start empty |

No other stop. Building the drafting sheets has no decision in the gap, so the
chain does it and reports it.

Period is written with `SetDeckPeriodVerified` and confirmed **from the saved
file's bytes**, never from PowerPoint's cache.

Already configured and current:

```
Q3F26 is set up. 43 rows, 8 drafting sheets built 9 Aug.
Nothing to do here. Next: press "2. Draft and publish".
```

## Chain B — `2. Draft and publish`

| Stop | Condition | The decision |
|---|---|---|
| B1 | field not derivable from the active sheet | **Which field?** -- scopes the whole chain |
| B2 | always, before writing the register | **The publish confirmation** -- the consent gate |

The field resolves ONCE at entry (`ActiveDraftField`, else `AskForField`) and
becomes the title of every dialog: `2. Draft and publish — ABOUT_BODY`.

That is FIX-LIST 1a fixed structurally rather than by remembering: the subject is
carried by the container, so a message cannot report a count without it. The
worst 1a instance -- *"Column F is empty for all 43 row(s)"* fired against the
wrong sheet -- becomes impossible.

Copying AI drafts to submit has no decision in the gap. The chain does it.

## Chain C — `3. Put it on the slides`

The only button that changes slide text.

| Stop | Condition | The decision |
|---|---|---|
| C1 | a review sheet holds ticks not yet applied | **Apply those / start fresh and lose them / cancel** |
| C2 | otherwise | **The write-authorising confirmation**, capped, detail in the Run Log |

**C1 is the safety property of the whole design.** `ReviewChangesCore` calls
`WriteQueueSheet` unconditionally, which rewrites the sheet and takes every tick
with it. A chain that opened by rebuilding would destroy an evening's approvals
-- strictly worse than today, where `Apply Approved` is picked directly. Detect
and branch, never rebuild first. **Built and tested: `ReviewQueue.PendingApprovals`.**

Closes FIX-LIST item 2. `ContentKindOf` returns `KIND_PROSE` for everything
outside three hardcoded names, `AssignBatches` batches only `KIND_CONTROLLED`,
so `HasBatchableWork` is never true for `ABOUT_BODY` and today's button 4 ALWAYS
ends in a `vbExclamation`. Under a chain, prose going to the review sheet is
simply the normal route: drop the exclamation, name the sheet via
`ReviewSheetNameFor`, delete `REVIEW_SHEET_NAME`.

`Review + Approve All` survives as a choice inside C2, never the default, with
`ReviewChangesCore` still prepending the `APPROVE-ALL` banner to the report and
the Run Log -- the banner is the mechanism `RibbonUI.bas:573` actually wanted,
not the separate button.

Ends with `PersistBothFiles` / `SaveDeckVerified` / `SaveWorkbookVerified`.

---

## The three hazards a chain creates and a menu did not

1. **A chain can do more than expected.** Every chain states its plan before its
   first write -- the stages it will run, named, with the state each was derived
   from -- and the end report lists every write. Respect `REPORT_CAP = 900` via
   `CapReport`, detail to the Run Log.
2. **Cancel must leave a describable state.** Every stop sits BEFORE a write, so
   Cancel leaves the last completed stage, and the chain says which: *"Stopped.
   Q3F26 is set and 43 rows were copied forward. The drafting sheets were not
   built. Press this again to finish."* Re-pressing resumes, because the chain
   reads state rather than remembering it.
3. **Never rebuild a surface holding human ticks.** See C1.

## What this costs

**The work is 37 strings, not 3 buttons.** 28 instruction strings across seven
modules name a button by caption (`RibbonUI` x7, `BatchOnboardFlow` x5,
`DraftingUI` x4, `ReviewQueue` x3, `DiscoverUI` x2, `AdoptFlow`, `RunSync`,
`WorkbookBridge` x1 each), plus 9 `Readiness` remedy values. Every one names a
caption that ceases to exist.

`Readiness` today offers **"Create Template Slide"** as a remedy -- a button that
does not exist on the current toolbar. That is the failure this redesign exists
to prevent, already live.

**`Remedy` must become an enum the code switches on, never free text**, so a
caption cannot be renamed out from under a message again. Fix the strings as a
CLASS with a grep, per FIX-LIST item 6.

## What this breaks

1. **`Test_CommandBarUI_EveryDeclaredCapabilityHasAButton`** (`TestRunner.bas`)
   hard-asserts 16 Subs have buttons. Split it: the runtime half asserts the 3
   dispatchers are wired and visible; the REACHABILITY half -- each demoted Sub
   is called from its chain -- moves to `check_vba_static.py`. Honest cost: that
   is Python, so **dev-machine only**, and the strongest guard against silently
   orphaning a capability gets weaker at work.
2. **`HideToolbar` must keep deleting the three old bar names**, or upgrading
   from build 40 leaves orphan bars (`CommandBarUI.bas:226`).
3. **`Readiness.Build` per press, and `PropertyOnDisk` copies the deck to unzip
   it.** Sub-second on the 49MB deck -- but measured as a PowerShell probe of the
   TECHNIQUE, not this code at chain frequency. UNVERIFIED.
4. **Assumes one slide type per deck.** `ResolveRegisterSheet` refuses more than
   one, so type pickers auto-select. A two-type deck brings them back.
5. **`TooltipText` raises past 255 chars and takes the whole toolbar down
   mid-build** (`CommandBarUI.bas:259`).

Per the standing rule: **every new assertion gets made to fail on purpose once**
before it is trusted.
