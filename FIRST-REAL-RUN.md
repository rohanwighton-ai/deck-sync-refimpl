# First real run — findings

Rohan onboarding his **real** deck for the first time, on the work machine,
2026-08-01. Everything before this was the redacted deck on the personal
machine.

**This file is the log of what he hits, as he hits it.** Findings get written
down and the run continues — they are not fixed mid-run unless they actually
block typing. The point of a first real run is to produce a list, and stopping
to fix each item is how you end up with a good tool and no finished quarter.

## Starting state, established before touching anything

- The add-in **loads on the work machine**. `addin33` active. This was the
  existential risk for the whole design — an employer blocking unsigned VBA
  add-ins would have killed it outright — and it is now answered. Yes.
- The real deck had **never been touched by the tool**: no paired workbook, no
  `DeckSyncId`, no tags. `Preview Sync` said so plainly. Cleanest possible
  starting state — nothing to migrate, nothing to undo.
- The register must be a **NEW workbook**, not `SAAFE-Projects-Data.xlsx` —
  that one was seeded from the *redacted* deck, so pairing the real deck to it
  would stage redacted text to sync onto real slides.

---

## Findings

### 1. No picture or icon field type when marking a field
`BatchOnboardFlow.NormalizeFieldType` accepts exactly four answers: `text`,
`number`, `currency`, `date`. There is no way to declare a field as a picture,
an icon, a bar or a timeline marker.

**Deeper than it looks.** That type is currently *cosmetic only* — its own
comment says it affects "a bonus NumberFormat, never the synced value itself".
So `FieldType` is an Excel formatting hint, NOT a rendering contract.

That corrects an earlier claim of mine that `FieldType` was already the right
hook for picture/bar rendering. The column exists; it means something else.
Widening its meaning silently would be the same class of mistake as the column
renumbering earlier today — a value whose meaning changed underneath the code
that reads it.

Whoever builds picture/bar rendering must decide deliberately: widen
`FieldType` and document the change, or add a separate rendering dimension.
Not both, and not by accident.

*Status: recorded, not fixed.*

### 2. Cancelling the field-type dialog marked the field anyway, as "text"
Rohan: *"cancelling the dialogue seemed to mark the field anyway incorrectly
text"*.

`InputBox` returns `""` both for **Cancel** and for **OK with nothing typed**,
and `NormalizeFieldType` maps anything unrecognised to `"text"`. So backing out
silently marked the field and carried on. The volatility prompt one step later
had the identical bug — it would have defaulted to `"variable"` — found by
reading rather than by being hit.

**The idiom was already in the same function.** The field-NAME prompt, eleven
lines above, does exactly the right thing: `If Trim(typedName) = "" Then …
Exit Sub`. It simply was not carried to the next two prompts. Same shape as
"columns by header name, never by position" sitting directly above a line that
read a sheet by position.

Two situations — "I changed my mind" and "I pressed OK" — collapsing into one
indistinguishable value. That is the recurring failure of this whole project,
now at the UI layer.

*Status: FIXED in source. Needs a new .ppam to reach the work machine.*
*Workaround until then: do not cancel those two dialogs. To undo a mis-marked
field use `Clear Marked Fields` and re-mark — marking again on top adds a
second identity rather than replacing the first.*

### 3. Marking an icon killed the whole batch, and the error blamed the user
`Bulk Onboard Type` failed with:

> Marked shape 'Graphic 285' could not be re-found on the template slide (was
> it deleted or moved into/out of a group after marking?). Clear marked fields
> and mark again.

**Nothing was deleted or regrouped.** `Discovery` only returns shapes with a
text frame that has non-empty text. `BuildBatchPlanFromMarkedFields` reconciles
marked shapes against that candidate list **by object identity** (`Is`), so an
icon or picture can be marked happily and can never be reconciled.

Three separate faults stacked:
1. **Marking accepted a shape the batch cannot process.** No check at the point
   of marking, where it costs one dialog.
2. **The failure surfaced at the END**, after a slide's worth of marking, and
   invalidated all of it.
3. **The message named the wrong cause** and implied user error. A person
   following it would go looking for a shape they had moved, and find nothing.

Cost: every mark on the slide, plus the time spent looking for a non-existent
regrouping.

*Status: FIXED in source — marking now refuses a shape with no text up front and
says why, naming pictures/icons/bars as unsupported-for-now rather than
implying a mistake. Needs a new .ppam.*
*Workaround until then: mark TEXT shapes only. Skip icons, photos, graphics and
the progress bar.*

### 4. No way to unmark a single field — one bad mark costs every mark
Marks are held as a marking-session list. The only removal is
`Clear Marked Fields`, which discards **all** of them. There is no "unmark this
shape".

On its own that is a small gap. Combined with finding 3 it is what turned one
wrong click into a lost slide's work: the bad mark could not be removed, so the
only route forward was to discard every good mark alongside it.

Marking already knows how to *update* an existing mark (`existingIdx > 0`
re-marks in place), so the list is addressable — removal is a small addition,
not a redesign.

*Status: recorded, not fixed. Would be a fourth button ("Unmark Field") or a
re-mark answer of "-" meaning remove.*

### 5. My own ListMarks truncated at 18 of 52 (MsgBox limit)
The first version of the rescue tool used a `MsgBox`, which truncates around
1000 characters **silently** — no ellipsis, no count mismatch, nothing to
indicate 34 marks were missing. Rohan had 52 marked fields and saw 18.

An auditing tool that cannot show the whole list is worse than no tool: it
looks like the full picture. Rewritten to write a text file and open it, and to
flag malformed lines explicitly.

Same failure as everything else today — the wrong answer presenting as a
complete one.

### 6. Two malformed entries in the marking session
`ListMarks` showed `1. shape: 256` (no field name, no type) and
`18. shape:` (no shape name at all). Both are lines that did not split into the
expected four `|`-separated parts — most likely a field name containing a line
break, splitting one record across two lines and leaving fragments either side.

Probably harmless: restore matches by shape name and skips marks it cannot
find, counting them as missing rather than aborting. So they drop out on the
next save/close/reopen along with the icon.

Worth guarding at write time though — the serializer should refuse or strip a
field name containing `|` or a line break, rather than writing a record that
cannot be read back.

*Status: recorded, not fixed.*

### 7. Anything needing the USER'S OWN FILE to hold macros is unusable here
The rescue script was imported into the presentation's VBA project. That made a
`.pptx` macro-bearing, which company policy blocks — and the block presented as
a **silent save no-op**, not a message.

The add-in itself is fine and loads normally, so macros are not blocked
wholesale: trusted add-ins are allowed, macro-enabled *documents* are not. That
is a sensible split and the design already respects it — the deck holds tags and
document properties, the workbook holds sheets, all code lives in the `.ppam`.

The rescue script was the only thing that violated it, and that was a bad shape
regardless of policy. Maintenance tooling belongs in the add-in, where it is
already trusted.

*Status: recovered by removing the module. The rescue functions should become
add-in buttons — `List Marked Fields`, `Unmark Field`, `Unmark By Name` — so
there is never a reason to put code in a user's document.*

---

## Open, parked deliberately

**Team distribution is now a requirement** (Rohan, 2026-08-01), via OneDrive —
which reverses the 2026-07-28 "personal tool, not org adoption" decision that a
lot of the current design rests on. Parked the same day: *"hang on we dont have
to do it now."*

Two things it forces, recorded so they are not rediscovered:

1. **Code-signing the `.ppam` moves from optional to required.** Unsigned, every
   teammate needs an individual Trust Center exception — a per-person IT
   conversation that will not survive a policy refresh.
2. **The register becomes shared mutable state, and there is no concurrency
   control.** Two people publishing at once can overwrite each other; a rebuild
   can wipe a sheet someone is typing into; and OneDrive resolves simultaneous
   edits by making *conflict copies*, so the add-in could silently read the
   wrong file. That last one fails the same way everything else here does —
   quietly, looking correct.

Unanswered, and it decides the shape of any solution: **how many people, and are
they on one shared register or each on their own?** Three people with their own
project decks is nearly free. Six on one register is a different tool.

Worth a DECISIONS.md entry when it is picked up, because it reverses a call the
architecture was built on.

### 8. Bulk Onboard never restored the saved marking session
After reopening the deck, `Bulk Onboard Type` reported **"No fields marked
yet"** while 33 marks sat intact in the `DeckSyncMarkingSession` document
property.

The restore lived only inside `MarkFieldForBatchCore`. So the one button a
person presses *after reopening a file* was the only one that never looked for
the saved session — and the message told them to start over.

*Status: FIXED — `PromptBatchOnboardType` now attempts a restore before
concluding nothing is marked.*
*Workaround until rebuilt: select a TEXT shape, run `Mark Field for Batch`, and
Cancel at the name prompt. The restore runs before the prompts, so cancelling
leaves the session restored and adds nothing.*

### 9. The session parser crashed on the malformed records — and I called them harmless
`RestoreMarkingSession` read `parts(1)`, `parts(2)` and `parts(3)` with **no
check on how many parts the line had**. So a record that did not split into four
took down the entire restore, losing every good mark with it.

Rohan's session contained exactly two such records (finding 6). I looked at them
and said they were "probably harmless: restore skips marks it cannot find". That
was wrong — it skips marks whose SHAPE cannot be found, which is a different
branch entirely. A short record that *does* match a shape name goes straight
into the unguarded read.

Two fixes, because one is not enough:
- **Read side:** a record with fewer than four parts is skipped and counted, not
  repaired. Guessing the missing parts would mark a shape under a name nobody
  chose.
- **Write side:** `SafePart` strips `|` and line breaks from every part before
  serialising, so a field name containing either can no longer produce a record
  that cannot be read back. That is where the two bad records came from.

*Status: FIXED both sides.*

### 10. Error 13 — my rescue macro corrupted the session, and an unguarded CLng turned that into a dead end
`Bulk Onboard Type` and `Mark Field for Batch` both died with *"Error 13: Type
mismatch"*, on a text shape and a graphic alike, with 33 marks intact and
unreachable in the document property.

Cause, in two halves:

1. **Mine.** The rescue macro rebuilt the session as
   `kept = kept & vbCrLf & line`, starting from an empty string — which puts a
   **blank line at the front**. The session's first line is a *slide ID*, not a
   record, so every line was shifted by one.
2. **The add-in's.** `RestoreMarkingSession` read `slideId = CLng(lines(0))`
   with no validation. `CLng("")` raises 13. Any session whose first line was
   not a number — a leading blank, a stray newline, anything hand-edited — was
   unrecoverable, and the message blamed something unanticipated.

Both were needed for the failure. Only one is fixable in the tool, so that is
the one fixed: the header line is now **located** (first non-blank) rather than
assumed to sit at index 0, a non-numeric header degrades to slide 0 instead of
raising, and the record loop starts after the header wherever it is. A session
that cannot name its slide is still a session full of marks; losing all of them
over a header is the wrong trade.

The rescue macro is fixed too, but that matters less — it should never have been
the thing writing to a user's file.

**The pattern, for the fourth time today:** unvalidated input meeting a
conversion that assumes success. `CLng` here, `parts(1..3)` in finding 9,
`NormalizeFieldType("")` in finding 2, `Worksheets(1)` this morning. Every one
presented as something other than what it was.

*Status: FIXED. Needs addin35.*

### 11. The marking session identified shapes by NAME, and names are not unique
The Field Review grid showed `Industry Cash Value`, `SAAFE Cash value`,
`In-Kind Value` and `Total Project Value` **all reading `$275,598`** — one
value in four fields. `About Text` and `Problem Text` both showed the About
paragraph. `Start Date` sampled 2028 while `End Date` sampled 2024.

The evidence was already in `ListMarks` and I read past it:

```
12. shape: Shape 16   field: Industry Cash Value
13. shape: Shape 16   field: SAAFE Cash value
15. shape: Shape 16   field: In-Kind Value
17. shape: Shape 16   field: Total Project Value
```

Four fields recorded against one shape name. Three more against `Text 35`.

**Measured on the real deck rather than assumed:**

```
slide 1 shapes (incl. nested): 158
duplicate NAMES: 47
duplicate IDs:   0
```

`Shape 13` appears twice, with `Id=125` and `Id=148`. **47 of 158 shapes share
a name.** Restore matched `allShapes(ai).Name = parts(0)` and bound every
duplicate to whichever came first.

**The marking was correct when made** — in memory it holds real Shape
references. The information was destroyed at *serialise* time, by writing a key
that could not distinguish the shapes. Which means the stored session cannot be
repaired: the distinction is already gone.

Cost: Rohan's hour of marking, unrecoverable.

*Status: FIXED. Records are now `Id | Name | Field | Type | Volatility`, matched
on `Shape.Id`. The name is still written, second, purely so a human can read the
session — it is a label now, not a key. Legacy four-part records still load, with
their original ambiguity, so an old session is not thrown away.*

**The lesson, and it is the day's biggest:** every other finding today was a
value that could not distinguish two situations. This one is an *identifier*
that could not distinguish two objects — the same failure at the level where it
does the most damage, because it corrupts data silently and looks like data.

### 12. Removing the modules did not shed the VBA project — only an explicit .pptx Save As did
After finding 7, the two rescue modules were deleted and the deck saved. It later
stopped saving again, silently, with AutoSave off.

Cause: **PowerPoint keeps the VBA project part even when every module is
removed.** The file still counted as macro-bearing, so policy still blocked it,
and the symptom was the same silent no-op.

Fix, which is also the test: `File > Save As`, explicitly choose
**PowerPoint Presentation (\*.pptx)**, new filename. That writes a fresh package
with no VBA part. It saved.

Two things worth carrying:
- **"I deleted the code" is not the same as "the file has no code."** The
  container outlives the contents.
- **The marking session survived**, because it is a document *property* — data,
  not code. That distinction is what made the design compatible with this
  policy in the first place, and it held.

**And a trap on the far side:** Save As to a cloud-backed location turns AutoSave
back ON by default, which is the combination that caused the earlier silent
failures — AutoSave disables manual save, and an upload that cannot complete
leaves edits unpersisted with nothing said. Turn it off for any deck the add-in
writes to.

*Status: resolved. Root cause was my rescue macro (finding 7); this is the tail
of it.*

### 13. The toolbar name is shared across add-in versions, so old builds silently overwrite new ones
Twice today the toolbar showed the wrong buttons: four buttons this morning
(`addin28` loaded alongside `addin33`), and the pre-reorder set this evening
(`addin35` alongside `addin36`).

Every version names its toolbar `"Deck Sync"`. `Auto_Open` deletes any bar of
that name and rebuilds; `Auto_Close` deletes it. So whichever add-in loads or
unloads **last** wins, and unloading an old one removes the new one's toolbar
by name.

No error either time. The tool simply presented an older interface, and the only
signal was the buttons looking wrong — which requires already knowing what they
should say.

*Status: recorded, not fixed. Putting the build number in the toolbar name
("Deck Sync 36") turns an invisible collision into two visible toolbars.
Flagged this morning and not acted on; it then cost a second diagnosis.*

---

## Built the same evening, in response

### `Setup A2: Discover Fields` — the marking grid
Finding 3, 4 and 11 were all consequences of marking being a **prompt chain**:
52 fields x 3 dialogs, about 150 prompts and an hour, where one wrong click cost
the lot and there was no way to remove a single entry.

This project already made this journey once. Instance keys were a modal prompt
per slide — 45 of them — until 2026-07-29, on the recorded reasoning that *modal
prompts are irreplaceable input, while a sheet is editable, re-openable, and
survives a validation failure with the human's edits intact.* Field marking never
got the same treatment.

`DiscoverUI` lists every text shape on the template slide in **one grid, in
reading order** — top of the slide first, not z-order, which is an artefact of
how the deck was built and correlates with nothing a person can see. Tick column
F, name in column G, one confirm.

Keyed on `Shape.Id` throughout.

**Additive on purpose.** `Setup A: Mark Fields` is untouched. A flow written in
one evening and never run against the real deck should not be the only road.

**Exercised, not just compiled.** Split into `BuildDiscoverySheet` /
`ApplyDiscoverySheet` so it runs with no human in it, then driven on the
43-slide rig:

```
build: 82 text shape(s) listed
grid rows: 82   blank ids: 0   DUPLICATE ids: 0
apply: 2 marked, 1 ticked-but-unnamed REPORTED not guessed
read back: SELFTEST_ONE, SELFTEST_TWO
```

82 shapes with zero duplicate ids, on the same slide where 47 of 158 *names*
collide. Both branches observed — the marking and the refusal — rather than only
the happy one.

### Toolbar name now carries the build number (finding 13)
`"Deck Sync 37"`. Two loaded add-ins now produce two visible toolbars instead of
one silently winning.

### 14. A read-only workbook is never detected — found by consultant review, verified
`WorkbookBridge.OpenOrGetWorkbook` does `xl.Workbooks.Open(path)` inside
`On Error Resume Next`. **`.ReadOnly` is never checked anywhere in production
code** — verified: the only hits in the tree are two diagnostic prints in
`tools/E2EField.bas`, and those are on the *presentation*, not the workbook.

So a register that is locked (open elsewhere, or on a read-only share) opens
**read-only**, everything proceeds normally, and the failure surfaces — or
doesn't — at `wb.Save`. The user is told nothing.

This is the project's signature failure one layer up: **a returned object that
cannot distinguish "I opened it" from "I opened a copy you cannot save"**. Same
family as `Worksheets(1)`, `CLng("")`, `NormalizeFieldType("")` and the shape
names.

It matters more the moment a second person exists, but it is already reachable
today — anyone with the register open in Excel triggers it.

*Status: recorded, not fixed. Roughly five lines in one place: after opening,
if `wb.ReadOnly` then refuse loudly and name the path. This is the one I would
fix first.*

### 15. `RepointWorkbook` exists but has no button
`DeckRegistry.RepointWorkbook` (`DeckRegistry.bas:210`) is the escape hatch for a
deck whose paired workbook has moved. It is referenced **zero times** in
`CommandBarUI.bas`.

Today the sibling-fallback in `GetWorkbookPath` covers the common case, so this
has never bitten. It would become the first support call the moment anyone else
moves a file — with no self-service fix.

*Status: recorded, not fixed.*

---

## Fixed the same night, after consultant review

**16. The publish button never saved the register.** `DraftingUI.PublishDraftsForField`
wrote `Value` and `Status` in memory and stopped; the harness path had always
called `wb.Save`. So `published: 12` described Excel's buffer and nothing on
disk. **Rohan's refusal to run a real quarter before the tool worked was
correct** — that exercise would have ended in lost approvals. Now saves and
*reports the outcome*, because a read-only workbook (finding 14) would otherwise
fail here in silence too.

**17. Two live bugs in `IsToolOwnedSheet`.** `Left(sheetName, 13) = "Template Audit"`
— that literal is **fourteen** characters, so the comparison could never be true:
an always-false guard shipped in production, inside the list written to prevent
this very class of mistake. And `"Field Discovery"` was missing, so
`RegisterOrFirstDataSheet` could hand back the discovery grid and call it the
register. Both were possible because the list duplicated names that already exist
as public constants. It now uses the constants, and prefix lengths are derived
with `Len()` instead of counted by hand.

**18. Macro-enabled files now warn before a write.** Rohan: *"remember the pptx vs
pptm thing too (pptm won't save on work machine)."* Nothing in this tool creates
a macro-enabled file — no `.pptm`/`.xlsm` constant exists in the source — but it
can be pointed at one, and on a managed machine that save fails **silently**.
`WorkbookBridge.MacroEnabledWarning` is surfaced before the confirmation in both
`Publish` and `Sync Now`, where it can still change the answer. It warns rather
than refuses: these files save fine on the personal machine, and the person knows
which machine they are on.

**19. Unticking a row now unmarks the field (finding 4, properly this time).**
`ApplyDiscoverySheet` only ever ADDED. A sheet that presents as declarative state
— *what is ticked is what is tracked* — behaved as an in-tray, which meant the
grid did **not** fix finding 4 despite looking like it should. I described it to
Rohan as fixing that, and it did not.

`BatchOnboardFlow.UnmarkShapeForBatch` removes one field by its shape, rebuilding
the position-keyed dictionaries rather than mutating them — an off-by-one there
would reattach a field name to a different shape, which is precisely what
destroyed the original marking. Reported by name, so an accidental untick is
visible. Covered by test.

---

## What review found that the suite could not

Two consultants and the PM reviewed the repo the same evening. Between them they
produced findings 14–19, and the single most uncomfortable observation:

**`test-fixtures/crc-real-deck-redacted.pptx` has been committed since 25 July
and is opened by ZERO VBA tests.** Measured: 46 slides, 5,828 shapes, **1,630
duplicate shape names, 0 duplicate ids.** Finding 11 — the defect that destroyed
an hour of real marking and could not be repaired — was sitting in a fixture in
this repository the whole time. The suite could not see it because *every fixture
it uses is one it wrote itself, from the same assumptions as the code.*

The systemic framing worth keeping: **for almost every fact this system depends
on, the component that produces the fact is also the component that vouches for
it.** The three exceptions — `read_deck_props.py` reading the zip with no Office,
counting `PASS` lines instead of grepping for `FAIL`, and Rohan's screenshots —
were each bought with an expensive incident, and each one works. The remedy is
not to distrust more values; it is to add independent oracles where none exists.

And the closing rule, which this project has earned: **an incident is not closed
until it has produced a check that has been made to fail once, on purpose.**
Prose belongs in the commit message; the defence belongs in a script.

---

## 2026-08-02 — three things are called "rollover", and they disagree

Found by Rohan asking a plain question about a guard being written for the
drafting sheet: *"rollover — what do you mean by it, and why would you delete
static text when it happens? Variable text is the text that gets updated on
each new time horizon."*

The guard was clearing every carried column on a period change. The defence
offered for that was: a drafting sheet only ever holds **Prose** fields, so
everything on it is quarterly. **That is wrong, and it conflates two different
axes:**

- `FieldSpec.Kind` — Controlled / Prose / Static — is *how a value is produced*.
- The register's `Quarter` column — a period, or `ALL` — is *when it applies*.

A project description is prose **and** static. Nothing stops a Prose field
carrying a `Quarter = ALL` row, and the guard would destroy that person's draft
while the register went on serving the ALL row. The Prose-only filter lives in
`DraftingUI.AskForField` (`:124`); the guard that depends on it is in
`Drafting.WriteDraftingSheet`. **Neither one said so.** Same shape as the
`E2EField.bas` sheet-by-position defect: a rule understood at one level and
never lifted to the level that relies on it.

**~~Recorded rather than fixed, because nothing can currently put an ALL row on
a drafting sheet.~~ WRONG, and stale within the hour.** Round 5 §3 classes
`ABOUT_BODY` — the flagship prose field, the one the drafting sheet was built
around — as **entity-static**, and Rohan confirmed the same evening that he
writes it once and edits it rarely. So the destroyed-work case was never
hypothetical: it was the main field, on every rollover, for no safety at all —
an ALL row's previous text *is* its current text, so there is nothing stale to
republish.

**FIXED the same night.** `Register.RegisterRead` now keeps the cadence map it
was already computing and discarding (`EntityCode & Chr(1) & FieldID -> True`
when the value came from a period row); `Drafting.WriteDraftingSheet` takes it
and drops **per row** instead of per sheet. A sheet with no period stamp is
treated as a rollover rather than as current — strictly better than both earlier
options, since it drops what might be stale and keeps what cannot be. Both
directions asserted, and the report counts what it cleared *and* what it kept,
because "nothing was carried across" would now be a lie on exactly the rows that
matter.

### The larger one, which the question flushed out

Searching for what `ALL` actually does turned up **three separate mechanisms
named for period rollover, built to two incompatible models:**

| Mechanism | Model |
|---|---|
| `DeckRegistry.SetDeckPeriod` — deck declares its period, register filters to it | deck renders ONE quarter |
| Register `Quarter` column + `QUARTER_ALL` — rows accumulate, ALL rows match any period | register is the ARCHIVE |
| `SyncOperations.PlanPeriodRollover` / `RunSync.RunPeriodRollover` | deck ACCUMULATES quarters |

The third duplicates a project's slide **inside the same deck**, tags the copy
with a new instance key, and leaves the original "untouched as history"
(`RunSync.bas:505`). It is the most thoroughly built of the three — a written
spec (`specs/sync-operations.md` case 2), a plan/execute split, its own tests,
and a ribbon button (`RibbonUI.bas:721`).

**Rohan's call, 2026-08-02: the deck must not accumulate.** Fourteen projects
over four quarters is 56 slides in a year, which he named as the reason without
being shown the code: *"the deck growing across time, ie gaining more quarters,
[is] less sensible for a multi project report."* The register grows; the deck
is a per-quarter **render** of it, and each quarter produces a new deck file
while last quarter's stays on disk as the record.

**Nothing implements that.** No per-quarter deck copy exists anywhere in the
source. The most-specified rollover subsystem in the codebase builds the model
that has now been rejected, and the model that was chosen has no code at all.

Not touched tonight: `RunPeriodRollover` and its button are now dead weight at
best and a trap at worst, but removing a spec'd, tested, wired subsystem is a
decision to take deliberately and in daylight, not as a side effect of writing
a test for something else.

### What this cost, and what it was worth

One question from Rohan, asked cold about a variable name, reached further into
the architecture than 138 passing tests. The tests all agree with the code
because they were written from the same assumptions; **he was the only oracle in
the room that had not read the source.** Same lesson as the fixture nobody
opened, arriving from the other direction.

### The gap underneath all of it — you already answered, three times

Rohan, on being shown the ABOUT_BODY reasoning: *"why just focussing on about?"*
Correct, and the honest answer was "because the spec's example did", which is
not a reason. The question is per field, and there are ~33 of them.

**Static-vs-variable is asked or stated in three places, and only one of them
does anything:**

| Where | Effect |
|---|---|
| Discovery grid `Static/Variable` (`DiscoverUI.bas:176`, defaults to `variable`) | **none** — `BatchOnboardFlow.bas:129`: *"human-declared hint only, not wired into sync behavior yet"* |
| Field Spec `Kind` = Controlled / Prose / **Static** | decides whether a field gets a drafting sheet at all (`DraftingUI.bas:124`) |
| Register `Quarter` column — a period, or `ALL` | **the only one that changes sync behaviour**, and it is typed by hand |

So the person marking fields answers the question for every one of them, the
answer is serialised, carried through onboarding, and read by nothing; while the
column that actually governs the behaviour is the one no part of the tool helps
them fill in. Rohan did not know `ALL` existed — *"I don't know where ALL came
from or what its for"* — which is unsurprising: it entered via
`Excel_Control_Layer_Round5_Consolidated.md` §3 and was implemented straight out
of a spec exchange he never had to ratify.

Worse, `Kind = Static` and `Quarter = ALL` sound like the same statement and are
not: the first switches OFF drafting for a field, the second keeps its value
across quarters. Marking `ABOUT_BODY` "static" per Round 5 in the wrong sheet
silently removes it from the drafting menu.

**Next item, and it is bigger than tonight's fix:** the marking answer should
seed the register's `Quarter` column, so classification happens once, in the grid
already being filled in. Same shape as the instance-key fix (45 prompts → one
grid) and the marking grid after it: *the answer was already given; the tool
just never used it.*
