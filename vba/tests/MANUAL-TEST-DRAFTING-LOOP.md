# Manual test — the drafting loop, end to end

> **REWRITTEN 2026-08-01 for layout 2.** The sheet gained a third text column
> that day: **C = ORIGINAL** (read-only), **F = AI DRAFT** (what Copilot wrote,
> *never* published), **G = SUBMIT** (your text — the only column that reaches a
> slide), **I = the tick**.
>
> This document previously said "column F is where the new text goes and G is
> the tick", which was true of the older layout. Following it as written would
> have put real text into the column that is never published, and a `Y` into the
> column that now holds text.
>
> **TRACKER item 8 was ticked against the OLD layout and these OLD
> instructions.** That approval does not carry over — the sheet Rohan approved
> no longer exists. Re-run Step 1 against the current sheet before treating item
> 8 as evidence of anything.

**What this tests:** the one thing no automated test can — whether the loop is
*usable by a person*. The 132 automated tests prove the machinery is correct. They
cannot tell you whether the sheet is readable, whether the gate tells you what you
need, or whether you'd trust it on a real deck at 9pm before a deadline.

**Time:** about 20 minutes.

**Run it in one PowerShell session**, staying in the repo directory throughout —
every command below assumes `(Get-Location)` is the repo root.

**You will need:** nothing prepared. Step 0 builds its own sandbox.

---

## The rule that is not negotiable

**Never point any of this at `OneDrive\Claude\test1.pptx`.**

Every command below defaults to `C:\Users\rohan\deck-sync-e2e\`, a scratch copy.
If you ever find yourself typing a `-DeckPath` that is not under `deck-sync-e2e`,
stop.

---

## Step 0 — Fresh sandbox

In PowerShell:

```powershell
$rig = "C:\Users\rohan\deck-sync-e2e"
Copy-Item "$rig\e2e-deck.pptx"  "$rig\e2e-deck.pptx.manual-$(Get-Date -f HHmmss).bak.pptx"
Copy-Item "$rig\register.xlsx"  "$rig\register.xlsx.manual-$(Get-Date -f HHmmss).bak"
```

**Expect:** two new `.bak` files. If anything below goes wrong, you restore from these.

**Close Excel and PowerPoint completely before continuing.** Every driver refuses
to run while either is open, deliberately — it will not attach to your live session.

---

## Step 1 — Build a drafting sheet

```powershell
cd \\wsl.localhost\Ubuntu\home\snadger77\deck-sync-refimpl
Copy-Item vba\tools\field_e2e.ps1 $env:TEMP
& "$env:TEMP\field_e2e.ps1" -RepoRoot (Get-Location) -Mode draft -FieldId ABOUT_BODY
```

**Expect:** `43 row(s) written for ABOUT_BODY`, then a prompt block printed.

**Now open `register.xlsx` and look at the `TPL_ABOUT_BODY` sheet.** This is the
real test, and it is a judgement call only you can make:

- [ ] Can you tell at a glance what you are being asked to do?
- [ ] Is the current text readable in column C, or do you have to widen it?
- [ ] Is it obvious that **G (SUBMIT)** is where your text goes and **I** is the tick?
- [ ] **Would you be willing to work down 43 rows of this?** If the honest answer
      is no, the sheet is wrong and the rest of this test does not matter.

Write down anything that annoyed you. That list is worth more than a pass.

---

## Step 2 — Draft one row by hand

Pick any project. In its row:

- Put something in **column G (SUBMIT)** — genuinely rewrite the sentence in column C, or
  just change a word. Do not invent facts about the project.
- Put `Y` in **column I**.

Save and close the workbook.

- [ ] Did you have to think about which column was which? (You should not have.)

---

## Step 3 — Publish, preview first

```powershell
& "$env:TEMP\field_e2e.ps1" -RepoRoot (Get-Location) -Mode publish -FieldId ABOUT_BODY
```

**Expect:** `=== PREVIEW: publish ABOUT_BODY ===` and `would publish: <your project>`.

- [ ] **It says PREVIEW, and the summary says "would be published".** If it says
      `published` without you passing `-Write`, stop — that is a serious bug and
      the exact one that hid here on 2026-07-31.

Now for real:

```powershell
& "$env:TEMP\field_e2e.ps1" -RepoRoot (Get-Location) -Mode publish -FieldId ABOUT_BODY -Write
```

**Expect:** `=== Publish ABOUT_BODY ===`, `published: <your project>`.

- [ ] Open `register.xlsx`, find that project's `ABOUT_BODY` row. Its `Status`
      should now read **`Approved`**, and every other `ABOUT_BODY` row should
      still read `Seed`. Close the workbook.

---

## Step 4 — The gate

```powershell
& "$env:TEMP\field_e2e.ps1" -RepoRoot (Get-Location) -Mode dryrun -FieldId ABOUT_BODY
```

**Expect:**

```
WOULD CHANGE:          1
not writable (held back by Status): <the rest>
```

- [ ] **`WOULD CHANGE` equals the number of rows you ticked — not 43.** Everything
      you did not approve is held back. That is the whole point of the
      Seed/Approved split.

**Do not expect a specific held-back number.** It is 42 on a pristine rig and
lower once earlier runs have left rows approved — normal, not a failure. An
earlier draft of this guide printed 42 as gospel; run it a second time and you
see 41, and the only two conclusions available are "it broke" or "ignore the
numbers". The second is the rubber-stamp habit this whole tool exists to
prevent, so the guide must not teach it.

**You may also see `N drafted but not ticked`.** Rebuilding a drafting sheet
preserves drafts and clears approvals deliberately — a rebuild can move the
exemplar, and an approval is against a specific pairing of exemplar and draft.
So a draft from an earlier session survives with its tick cleared. Correct
behaviour, not a bug.
- [ ] The before-and-after line names your project and shows where the text first
      differs.
- [ ] The last line says `DRY RUN -- nothing was written to the deck.`

---

## Step 5 — Watch it land

```powershell
& "$env:TEMP\field_e2e.ps1" -RepoRoot (Get-Location) -Mode apply -FieldId ABOUT_BODY
```

**Expect:** `written and verified: 1`, `failed verification: 0`, and
`slides matching the register: 1  /  mismatched: 0`.

- [ ] **Open `e2e-deck.pptx` and find that project's slide. Your text is on it.**

That is the loop. Everything before this was machinery.

---

## Step 6 — The part most likely to be wrong

Run the gate again without changing anything:

```powershell
& "$env:TEMP\field_e2e.ps1" -RepoRoot (Get-Location) -Mode dryrun -FieldId ABOUT_BODY
```

- [ ] **`WOULD CHANGE: 0` and `already correct: 1`.**

If this says 1 again, the tool wants to rewrite a slide it just wrote — a
permanent phantom correction, and the single most valuable thing this test can
catch. It means the value written and the value read back disagree, usually over
whitespace or line breaks.

---

## Step 7 — Try to break it

Do these in any order. Each one *should* be refused or reported, never silently
accepted.

| Do this | Expect |
|---|---|
| Tick `Y` on a row with an **empty** column G, publish | `SKIPPED ... SUBMIT is empty`, nothing published |
| Tick `Y` on a row with text in **F (AI DRAFT)** but empty **G**, publish | `SKIPPED ... there IS an AI draft, but SUBMIT is empty`. The AI's words never publish on their own. |
| Write a draft but leave G **blank**, publish | counted as `drafted but not ticked`, nothing published |
| Publish **twice in a row** with `-Write` | second run publishes the same row again — harmless, but note whether that surprises you |
| Edit the slide **by hand** in PowerPoint, save, close, then run `-Mode dryrun` | the change shows up as pending again, with your hand edit as the "now" side |
| Put `Aproved` (typo) in a register `Status` cell, run `-Mode dryrun` | a `WARNING` about an unrecognised status; that row does **not** sync |
| Blank a `SlideType` cell in the register, run `-Mode dryrun` | `WARNING` about blank SlideType; that row does **not** match |

- [ ] Anything on that list that was silently accepted is a bug. Write it down.

---

## Step 8 — Restore

```powershell
$rig = "C:\Users\rohan\deck-sync-e2e"
Get-ChildItem "$rig\*.manual-*.bak*" | Sort-Object LastWriteTime | Select-Object -Last 2
```

Copy those two back over `e2e-deck.pptx` and `register.xlsx` if you want the rig
returned to where you found it. Or leave it — it is a scratch copy and the next
run rebuilds what it needs.

---

## Verified mechanically on 2026-07-31

Steps 0, 1, 3, 4, 5, 6 and the three guard rows of Step 7 were run exactly as
written and all passed — including Step 6 (no phantom correction) and every
guard: ticked-but-empty skipped, typo'd `Status` warned and skipped, blank
`SlideType` warned and skipped.

**That establishes nothing about Steps 1 and 2.** Those ask whether a person
would actually use this, which is the question that decides it, and no machine
can answer it.

---

## What a pass actually means

Every box ticked means the machinery works. It does **not** mean the tool is good.

The questions that decide that are in Step 1 and Step 2, and only you can answer
them: would you sit down and work through 43 rows of this on a Tuesday night? If
not, say what would have to change. That is the finding worth having — the other
132 tests already cover whether the code is right.
