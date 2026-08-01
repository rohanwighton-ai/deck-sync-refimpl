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
