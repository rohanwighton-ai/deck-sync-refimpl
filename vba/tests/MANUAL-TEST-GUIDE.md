# Manual test guide

> **HISTORICAL — written 2026-07-29**, when the toolbar had sixteen buttons. The
> interactive surface it walks no longer exists in that shape; the reasoning about what
> only a human can test still holds.

For the parts no automated test can reach: the interactive layer (every `InputBox`
and `MsgBox` flow), the toolbar, and the install. The 93 automated tests cover
everything underneath these.

**Work through in order.** Later tests assume earlier ones passed, and a failure at
step 3 makes step 9 meaningless.

---

## Setup — do not skip

**Never run these against `test1.pptx`.** A sandbox is prepared:

```
OneDrive\Claude\test-sandbox\
    sandbox-deck.pptx              (copy of the real 46-slide deck)
    SAAFE-Projects-Data.xlsx       (copy of the real Data workbook)
```

The sandbox deck has already been re-pointed at the sandbox workbook, so nothing in
here can touch live data. **Verify that before starting**: open `sandbox-deck.pptx`,
run Bulk Onboard or Preview, and confirm any path shown says `test-sandbox`.

Backups of the live files (taken 2026-07-27) are in `backup-20260727-pre-edittest\`.

If a test corrupts the sandbox, delete the folder and ask for a fresh copy — do not
attempt repair mid-run, it invalidates everything after it.

---

## 1. Install

1. In the PowerPoint window left open by `build_ppam.ps1`: **File > Save As**
2. Set *Save as type* to **PowerPoint Add-in (\*.ppam)** — pick it from the dropdown,
   don't type the extension
3. Save as `addin22.ppam` in `OneDrive\Claude\`
4. If warned that macro-free formats can't save VBA: **No**, then fix the dropdown
5. Close PowerPoint entirely
6. Reopen PowerPoint → **File > Options > Add-ins > Manage: PowerPoint Add-ins > Go >
   Add New** → select `addin22.ppam`

**Expect:** under the **Add-Ins** ribbon tab, a group labelled **Custom Toolbars**
containing **3 buttons**: Mark Field for Batch, Bulk Onboard Type, Clear Marked Fields.

**The toolbar is NOT labelled "Deck Sync" in the UI** — confirmed 2026-07-28. Office
gives legacy `CommandBars` its own generic grouping and never displays the toolbar's
own name, so `TOOLBAR_NAME` is only ever visible to code. Judge this test on the three
buttons being present and wired, not on finding the words "Deck Sync" anywhere. A
genuinely branded ribbon tab is impossible for a `.ppam` (proven 2026-07-26) and would
need a COM/VSTO add-in.

**If the toolbar doesn't appear:** `Auto_Open` didn't fire. Run `CommandBarUI.ShowToolbar`
from the VBE (Alt+F11) — if that works, the add-in loaded but the lifecycle hook
didn't, which is a different (smaller) problem.

**Delete `addin20.ppam` / `addin21.ppam` once this works.** Old builds still contain
the re-keying bug that corrupted the deck on 2026-07-27.

---

## 2. Marking — a top-level field

Open `sandbox-deck.pptx`, go to slide 1.

1. Click the **Project number** text (`3_P001`)
2. Click **Mark Field for Batch**
3. Name prompt appears

**Expect:** the prompt shows the field's current value **truncated to ~20 characters**,
on one line. Not a wall of text.

4. Name it `Project number` → choose type **2 (Number)** → choose **1 (Static)**

**Expect:** confirmation reading `Marked field 1: 'Project number'`.

---

## 3. Marking — a grouped field (the important one)

Still slide 1. The **Project Name** text is inside a group — all 46 of them are.

1. Click the Project Name text (may select the whole group)
2. Click **Mark Field for Batch**

**Expect:** if PowerPoint reported the group rather than the shape, a **numbered
picker** lists the group's members, each with a short text preview. Not shape names
alone — `TextBox 51` tells you nothing.

3. Pick the right number → name it `Project Name` → type **1 (Text)** → **2 (Variable)**

**Why this matters:** grouped fields were invisible to the engine until 2026-07-27. If
marking inside groups fails, the tool cannot work on this deck at all.

---

## 4. Clear Marked Fields

1. Click **Clear Marked Fields**

**Expect:** confirmation that marks were discarded.

2. Click **Bulk Onboard Type**

**Expect:** refusal — *"No fields marked yet."* Not a crash, not an empty grid.

---

## 5. Marking survives a close

1. Mark 2 fields (repeat 2 and 3)
2. **Save and close** the deck, reopen it
3. Click **Bulk Onboard Type**

**Expect:** your marks are still there — it proceeds rather than saying nothing is
marked. Marking state persists in a document property.

**This is the highest-value test in the guide and it is known to have been failing.**
On 2026-07-28 the restore guard was fixed for the wrong case: it compared the deck's
`FullName`, so reopening *the same* deck reported an identical value, the guard never
fired, and the marks stayed in the file while the add-in said nothing was marked. Only
switching to a *different* deck worked. Fixed 2026-07-29 by probing whether the held
Shape references are dead.

The lesson worth keeping: this test was in the guide the whole time and would have
caught it. A test that exists but is never run is worth exactly as much as no test.

---

## 6. Bulk Onboard — auto-select by layout

With fields marked and no explicit multi-slide selection:

1. Click **Bulk Onboard Type**
2. Enter a type name — use `sandbox-test`, **not** `q`

**Expect:** it switches to Slide Sorter, selects the template plus every same-layout
sibling, and asks you to confirm the count via MsgBox. On this deck that should be
**46 slides**.

**Expect:** declining ("No") leaves the deck alone and tells you to select manually.

---

## 7. The Field Review grid

Continue from 6, accepting the auto-selection.

**Expect:** Excel opens a **Field Review** sheet with one row per marked field, showing:
field name, a suggested include Y/N, the template's value, sample values from other
slides, the type and static/variable hints you set.

Test the editing round-trip:

1. Rename a field in the grid
2. Flip one field's **Include** to `N`
3. Return to PowerPoint and click **Yes**

**Expect:** the renamed field is used, and the excluded field is not written at all.

---

## 8. Instance keys — reuse (the 2026-07-27 fix)

Immediately after 7.

**Expect: you are NOT prompted for instance keys at all.** Every slide already carries
one, so all 46 are reused.

**Expect:** the result dialog includes
`Kept existing instance key (already linked, not re-keyed): 46`

**This is the regression test for the bug that orphaned all 46 slides.** If it prompts
you for keys here — **stop immediately** and don't accept. That means the fix isn't in
the build you installed, and continuing will corrupt the sandbox exactly as it
corrupted the live deck.

---

## 9. Commit result

**Expect:** `Linked: 46`, `Skipped: 0`, `FAILED verification: 0`.

Any non-zero FAILED count is a real defect — capture the slide numbers listed.

---

## 10. Preview Sync — read-only

Not on the toolbar yet. Alt+F11 → Immediate window (Ctrl+G):

```vba
Application.Run "RibbonUI.SyncPreview"
```

**Expect:** a report saying `PREVIEW (nothing written)`, then
`46 unchanged, 0 would be corrected, 0 new slide(s) would be created`.

**Then verify it really wrote nothing:** close the deck **without saving** and reopen.
Everything should be as it was.

**If it says any number of slides "would be created" — stop.** That means linkage has
drifted, and a real sync would duplicate slides.

---

## 11. Preview catches a real change

1. In the sandbox workbook, change one **Project Status** cell (column D, row 2)
2. Save and close Excel
3. Run `Application.Run "RibbonUI.SyncPreview"` again

**Expect:** `1 would be corrected`, naming the instance, the field, and the slide's
**current** value — so you can see before/after.

---

## 12. Sync Now — the real write

Only after 11 behaved.

```vba
Application.Run "RibbonUI.SyncNow"
```

**Expect:** `1 corrected`, `0 created`, `0 failed`. The slide text now matches the cell
you edited.

**Expect:** the count is **slides**, not fields — editing two fields on one slide
reports `1 corrected`.

Save the deck, close, reopen, and confirm the change persisted.

---

## 13. Portability — the work-machine test

The reason this matters: on a work PC the stored absolute path won't exist.

1. Close everything
2. Copy **both** sandbox files into a brand-new folder (e.g. `Documents\portability-test\`)
3. Open the copied deck from there
4. Click **Preview Sync** (a toolbar button as of addin28 — no longer a VBE call)

**Expect:** it works — finding the workbook beside the deck rather than at the stored
path. This is the fix that makes the tool usable at work.

**If it says "Could not open the paired workbook":** the sibling lookup failed. Capture
the exact path in the message.

---

## 14. Toolbar lifecycle

1. Unload the add-in (File > Options > Add-ins, untick it)

**Expect:** the Deck Sync toolbar disappears (`Auto_Close`).

2. Re-tick it

**Expect:** toolbar returns, buttons still wired.

---

## 15. Four buttons, Preview Sync first (addin28 sanity check)

Look at the Deck Sync toolbar.

**Expect:** `Preview Sync | Mark Field for Batch | Bulk Onboard Type | Clear Marked Fields`

**Three buttons means you are running addin27 or older.** Stop and check the add-in
list rather than debugging anything downstream — every test below assumes addin28.

---

## 16. Office keeps its own Save command (the 2026-07-28 regression)

On a **cloud-hosted** deck (OneDrive/SharePoint — this does not reproduce locally):

1. Mark one field
2. Open the **File** menu

**Expect:** Office's own Save is still there and still works; the AutoSave toggle still
reads normally.

**Also expect** the confirmation to report a real timestamp: *"Deck last saved to disk:
2026-07-29 21:14."*

**Why both halves matter:** the add-in used to force a save on every mark, which took
saving away from Office and hid its Save command. Trusting AutoSave instead fixed the UI
and broke persistence — a deck was found 2.6 hours stale with marks only in memory. The
add-in now saves only when the deck's real save timestamp has gone stale, so this test
and test 5 pull in opposite directions **and both must pass**. Either one alone is easy.

**If the timestamp reads hours old:** AutoSave has stalled. That is the failure this
guard exists to catch — note it, hit Ctrl+S, and say so.

---

## 17. Workbook comes first, and can be browsed

On a deck with **no paired workbook yet**:

1. Mark a field, run **Bulk Onboard Type**, get through the Field Review grid

**Expect:** you are asked for the Data workbook **before** any instance-key prompts, via
a real **Save As dialog** — not a typed path, and not after 45 prompts.

2. Cancel the dialog

**Expect:** a clean "Cancelled — no workbook path given. Nothing was written." No VBA
error, no Debug/End.

**Why:** the prompt used to come *after* the key loop, so one bad path discarded 45
hand-confirmed keys (2026-07-29). Ordering is the fix; the browser is the convenience.

---

## 18. A web address is refused with an explanation

Only reachable if the Save As dialog is unavailable and you get the typed prompt — worth
trying deliberately if you can:

1. Paste an `https://` SharePoint URL as the workbook path

**Expect:** *"That's a web address, not a file path"*, with the explanation that cloud
decks report their own location as a URL and you want the synced folder instead. Then it
offers to try again.

**Not expected:** runtime error 52, "Bad file name or number", Debug/End. That was the
2026-07-29 failure, and its cause was an unguarded `Dir()` inside the routine that was
supposed to *confirm* the save had worked.

---

## 19. Two slides cannot share an instance key

During the instance-key prompts, deliberately give a slide a key you already used.

**Expect:** it refuses, names the slide already holding it, explains that both would
point at one row in the Data sheet, and re-prompts. A blank skips the slide instead.

**Why:** nothing checked this until 2026-07-29. Two slides sharing a key both resolve to
one row — the second overwrites the first, nothing errors, and the deck quietly starts
showing another project's numbers.

Also watch the commit summary for **"duplicate instance keys ALREADY in this deck"**.
That reports pre-existing collisions, most likely from the 2026-07-28 key re-derivation
bug. It is a real finding about your deck, not a complaint about this run.

---

## Results

| # | Test | Pass/Fail | Notes |
|---|---|---|---|
| 1 | Install + toolbar | | |
| 2 | Mark top-level field | | |
| 3 | Mark grouped field | | |
| 4 | Clear marked fields | | |
| 5 | Marking survives close | | |
| 6 | Auto-select by layout | | |
| 7 | Field Review grid | | |
| 8 | **Instance key reuse** | | |
| 9 | Commit result | | |
| 10 | Preview (no-op) | | |
| 11 | Preview (real change) | | |
| 12 | Sync Now | | |
| 13 | **Portability** | | |
| 14 | Toolbar lifecycle | | |

**8 and 13 are the ones that matter most** — 8 is the regression test for the bug that
corrupted the live deck, and 13 is what decides whether this can be used at work.

Untested by design: `New Period`, `Onboard New Slide Type`, `Resolve Unmatched Fields`
and `Adopt Existing Slides` remain off the toolbar and unexercised.
