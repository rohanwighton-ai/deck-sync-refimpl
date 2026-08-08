# Run sheet — full manual rehearsal on the personal machine

**Gist:** a step-by-step script for doing one complete fake quarter by hand on this
machine, so every button gets pressed once before the real deck at work is involved.

Rules for this run:
- Fake data only. Nothing here needs to be real text.
- Work on `e2e-deck.wide-test.pptx` (the copy). **Never** `e2e-deck.pptx` — the
  original is still paired to a real OneDrive workbook.
- After every step, say what you saw. If a dialog says something confusing, that IS
  the finding — don't work around it, report it.
- Claude's job: read files off disk between steps and confirm what actually changed,
  independent of what the dialog claimed.

---

**Route: the toolbar, not PowerShell.** The command-line drivers test a path that
does not exist at work. Click what you will actually click.

---

## Before you start — DONE 2026-08-07

| | Do this | Result |
|---|---|---|
| 0.1 | `preflight.py` on the deck | **READY**, 0 blocking. FY26Q4, type `project-status` -> sheet `Register`, paired to `register-wide.xlsx`, 43 slides all tagged, 43 rows at FY26Q4 / 5 at FY27Q1. One warning: slide text still mentions Q3F26, Q1F26, Q2F26, Q1F27 — historical content, expected here. |
| 0.2 | Backups taken | `e2e-deck.wide-test.pptx.manualrun-20260807-101833.bak.pptx` and `register-wide.xlsx.manualrun-20260807-101833.bak.xlsx` |
| 0.3 | You: open `e2e-deck.wide-test.pptx`, check slide 1 | no `LOOPTEST-…` marker text left over |

**The work-machine gap this exposed:** `preflight.py` is a Python script on the
personal machine. At work there is no Python, no WSL and no Claude — so today there
is NO way to ask "are these two files actually linked?" on the machine where it
matters. Queued: promote that check to a toolbar button. Workaround meanwhile: copy
the real deck + workbook here and run it, which works because it reads closed files.

---

## The toolbar, exactly as captioned

`Setup A: Mark Fields` · `Setup A2: Discover Fields` · `Setup B: Onboard Slides` ·
`Setup C: Check Coverage` · `Setup: Clear Marks`

`0. Start a Quarter` · `1. Drafting Sheets` · `2. Copy AI to Submit` ·
`3. Publish & Preview` · `4. Sync Now`

`Preview Sync` · `Review Changes` · `Apply Approved` · `Review + Approve All` ·
`Repoint Workbook`

There is **no "check connection" button**. That is the gap.

---

## FIRST — fix the period spelling (do this before clicking anything)

The slides say **`Q4F26`**. The machinery says `FY26Q4`. That spelling came from
Claude and leaked into the deck property and the register's `Quarter` column.

**Periods are free text, matched exactly, with no validation.** Two spellings of one
quarter match nothing, the drafting sheet comes up empty, sync writes nothing, and
both report success. Fix it before it costs an evening.

**1. In Excel** — open `register-wide.xlsx`, Register sheet, Ctrl+H:

| Find | Replace | Expect exactly |
|---|---|---|
| `FY26Q4` | `Q4F26` | **43 replacements** |
| `FY27Q1` | `Q1F27` | **5 replacements** |

Verified safe: those two strings appear **only** as Quarter values, only on the
Register sheet, nowhere else in the workbook. Every other quarter reference in the
body text is already in your convention.

**If the counts are not 43 and 5, stop and say so** — that means the strings are
somewhere they shouldn't be. Save and close.

**2. In PowerPoint** — `0. Start a Quarter`, type `Q4F26`, save. This keeps the deck
on the quarter its 43 rows are in.

**3. Claude re-runs preflight** and confirms both ends agree before you draft.

---

## Part A — Setup (only if testing onboarding from scratch)

The rig copy is **already onboarded**. Skip Part A unless you want to prove setup
works on a virgin deck — in which case Claude carves a fresh copy first.

| | You do | Tool does | Watch for |
|---|---|---|---|
| A.0 | **Start a Quarter** first | writes the period onto the deck | must be done BEFORE onboarding, or you get a raw error |
| A.1 | **Discover Fields** — tick and name each field in one Excel grid | lists every text shape, marks nothing yet | is the grid readable? is reading order sane? |
| A.2 | **Onboard Slides** — review the batch, confirm | links all slides of that layout, writes one row each | row count = slide count |
| A.3 | **Check Coverage** | lists what's on the slide but not tracked | anything important missing |

---

## Part B — The quarterly loop (this is the real rehearsal)

### 1. Start a Quarter
- **You:** click it, type `FY27Q1`, confirm, save the deck.
- **Expect:** it writes the period and reads it back.
- **Known stale:** the closing message tells you to hand-copy last quarter's rows in
  Excel and set Status to Seed. Ignore it — that's on the fix list. Note whether it
  confused you.
- **Claude checks:** reads the deck's period off the saved file.

### 2. Drafting Sheets
- **You:** click it.
- **Expect:** one sheet per prose field. Every project a row. Last quarter's text in
  column C, your new wording goes in column G, tick column I when done.
- **Claude checks:** row count and that column C is populated from the register.

### 3. Write fake content — 10 rows
- **You:** type obviously-fake text into column G for 10 projects. Tick column I.
- **This is the step that takes the evening in real life.** Time it. That number is
  the whole business case.
- **Copilot role:** if you want, dictate the gist per project and Claude drafts the
  cell text; you edit and tick. Tick is yours, always.

### 4. Copy AI to Submit *(optional — test it once)*
- **Expect:** fills only empty SUBMIT cells. Never overwrites your words.
- **Test it properly:** put your own text in one row first, then run it, and confirm
  that row is untouched.

### 5. Publish & Preview
- **You:** read the list of ticked rows, say go.
- **Expect:** writes those rows into the register for the deck's period. Nothing
  touches a slide yet.
- **Deliberate behaviour:** a slide with no row for this period is **refused**, not
  created. If you see a refusal, that's correct.
- **Claude checks:** reads the saved `.xlsx` bytes and confirms the 10 rows landed on
  the right period.

### 6. Preview Sync
- **Read-only.** Everything the register would change, before anything is written.
- **Expect:** your 10 changes listed, nothing else.

### 7. Sync Now
- **You:** read what it will change, confirm.
- **Expect:** before-and-after for each change, then it writes and re-reads the slides.
- **Claude checks:** unzips the saved `.pptx` and reads the slide XML — confirms your
  fake text is on the slides and the counts match.

---

## Part C — The careful route (test it once, separately)

Only after B works end to end. Same idea, three steps instead of one:

1. **Review Changes** — builds a "Sync Review" sheet, current vs proposed per slide.
2. Tick what you approve on that sheet.
3. **Apply Approved** — takes a backup, re-checks each change, skips anything that
   moved since you ticked it.

Worth proving once because this is the route you'll want on the real deck at work.

---

## Part D — Second quarter (proves the model)

Do a short version of Part B again at `FY27Q2` with 2 rows.

**The thing being proved:** FY27Q1's rows are still there afterwards, untouched. A
quarter's approved text surviving the next quarter is the single most expensive bug
this project has had.

---

## What to capture as you go

- Any dialog whose wording you had to re-read.
- Any moment you didn't know what to click next.
- Minutes spent on step 3, honestly.
- Anything you expected to be able to do and couldn't.

That list is the output of tonight — more than the deck is.

---

## Then, and only then

Move to the work machine: install the add-in, run `preflight.py` against the real
deck, and do Part A properly on a **copy** of it.
