# Document map — what each file is authoritative for

> **CURRENT. This is the index every other document is governed by**, and it is
> enforced by `vba/tests/check_docs.py`, which fails the build on a document that
> contradicts the code. Written 2026-08-14 after six documents were found stating
> facts the source disproved — several of them read in full the same morning.

Document control is not a habit here, it is a check. The rule underneath all of it:

> **A machine-knowable fact must be DERIVED, never written into prose.** A sentence
> cannot fail a test, so it drifts silently — and a stale sentence with a plausible
> orphan beside it is worse than an obviously wrong one, because it reads as
> authoritative and sends a person to the wrong place.

---

## 0. ROHAN'S DECISIONS THAT ELIMINATE WHOLE CATEGORIES

**Read these before proposing anything.** Each one removed a class of problem rather
than fixing an instance, and each was reached by him asking a plain question about a
word or a shape — not by the test suite, which has been green through every one of them.
When a proposal conflicts with one of these, the proposal is wrong.

**1. "It's a build, not a rebuild." "Why is clear still happening?"** *(14 Aug)*
The drafting sheet is UPDATED IN PLACE. `ws.Cells.Clear` survives only for a column
renumbering. **Eliminates:** the carry dictionaries, the restore step, the park's reason
to exist, layout stranding, and every future variant of "the rebuild lost my work" —
which had produced five separate patches between 1 and 14 August and still lost 27
drafted paragraphs. A destructive operation made survivable is not a fix; the tell is
adding a backup instead of deleting a call.

**2. One file pair per quarter, and the central register is DERIVED, never authored.**
*(14 Aug — `claude-brain/DECISIONS.md`)* **Eliminates:** cross-period contamination as a
category. The old quarter is safe because it is a different file, not because the code
was careful. The derived aggregate eliminates the second-source-of-truth drift that an
editable one would guarantee.

**3. Fields are declared at template time; template NAMES are the identity.**
*(13–14 Aug, decided and NOT YET BUILT)* **Eliminates:** the entire "recognising"
surface — discovery, the marking grid, orphan-column warnings, and device internals
appearing as candidate fields. If the template says what exists, nothing has to perceive
it. A name is repairable by a person in the Selection Pane; an invisible tag is not, and
there is no Claude on the work machine. May also delete the planned device registry: if
names carry device membership, the naming convention IS the registry.

**4. "We are simplifying, why split?"** *(14 Aug)* Buttons are not a safety mechanism —
the confirmation is. The toolbar went 16 → 2 on 9 August and stays at 2. **Eliminates:**
recurring proposals to add a button for a problem already fixed at the root.

**5. One add-in. The deck is the anchor.** *(14 Aug)* `ActivePresentation` yields the
deck; the deck's properties yield the period and the workbook path. **Eliminates:** an
Excel-hosted half needing the deck's path and period copied into the workbook — a second
copy of a machine-knowable fact, which is the same defect class this whole document
exists to prevent.

**6. THE ARCHIVE IS LAST QUARTER'S FILE. `REPORTED LAST TIME` IS NOT STORAGE.** *(15 Aug,
Rohan, verbatim)*

> *"I can go back to previous quarters non destroyed drafting sheets to see field progeny,
> it is only in the new quarter for some of the ai tools and human to use its structure
> and narrative consistency."*

The quarter turn ferries `SUBMIT` into `D — REPORTED LAST TIME` and clears `DRAFT`,
`SUBMIT`, `SUBCHARS`, `APPROVED`, `SOURCES` and `NOTES` on the row. Read alone, that looks
like the ferry carries the claim and drops its evidence — the citations being the control
on the generative step. **It does not.** Under decision 2 (one file pair per quarter) the
previous quarter's drafting sheets still exist, undestroyed, in the previous quarter's
workbook: sources, notes, drafts and approvals all intact, which is where **field progeny**
is read. The ferried column is not an archive and must never be designed as one — it is a
working aid carried into the new quarter so a human and the AI tools have last quarter's
voice in front of them, for **structure and narrative consistency**, while drafting.

**Eliminates:** every proposal to ferry more columns "so nothing is lost", to widen
`REPORTED LAST TIME` into a history, or to treat the rollover clear as data loss needing a
carry, a park or a refusal. Nothing is lost — it is in another file.

**The live gap this exposes, and the only real one:** decision 2 is **NOT YET BUILT**. The
register still holds `Q3F26`, `Q4F26` and `Q1F27` in a single workbook, so the rollover
clear currently fires in the *only* copy rather than in a fresh one. Until file-per-quarter
lands, `ParkSheetCopy` is load-bearing and must not be removed — it is holding that gap
shut. Once it lands, the park has nothing left to guard.

---

## 1. The one rule that decides where a fact lives

| the fact | lives in | never in |
|---|---|---|
| what a column is | `Drafting.bas` constants | prose, as a letter |
| what a button is called | `CommandBarUI` `CAP_*` | prose, as a caption |
| what a sheet is called | the `*_SHEET_NAME` constants | prose, as a literal |
| which sheets survive a rebuild | `WorkbookBridge.LifespanOf` | prose |
| what the tool has delivered | the workbook's `Sync Log` | any doc's headline |
| what a field is for | the `Field Spec` sheet | any doc |

If a document needs to state one of these, it names the **thing** (`the SOURCES
column`), not the **coordinate** (`column D`). Coordinates drift; names do not.

---

## 2. The live documents

Every one carries a status banner in its first 40 lines. Anything marked
`HISTORICAL` / `SUPERSEDED` / `STALE` is exempt from the content checks — it is a
record of what was believed, and rewriting it to match today's code would destroy
the record.

| document | authoritative for | status |
|---|---|---|
| `NEXT-SESSION.md` | **where the project actually is.** Read first, always. | CURRENT — rewritten each session |
| `CHECKLIST.md` | **what's actually left, flat and tickable.** The primary handover surface as of 2026-08-16. | CURRENT — ticked together, in session |
| `SYSTEM-OVERVIEW.md` | what this system is, for a reader with zero context | CURRENT — added 2026-08-16, migrated from `archive/HANDOVER-Q4F26-DRAFTING.md` with provenance |
| `DOCUMENT-MAP.md` | this index and the review protocol | CURRENT |
| `FIX-LIST.md` | what is known-broken and not yet fixed | CURRENT — re-audited 14 Aug |
| `TRACKER.md` | the framing and the manual baseline | CURRENT for framing; defers status to `NEXT-SESSION.md` |
| `COLUMNS.md` | the *reasoning* behind the register's columns | CURRENT — the Field Spec owns the list |
| `PROVENANCE.md` | the provenance design | CURRENT — designed, not built; build steps now on `CHECKLIST.md` |
| `SOURCE-CAPTURE-FORM.md` | the paper form for capturing sources | CURRENT — tool-independent |
| `AGENTS.md` | platform gotchas earned the hard way | CURRENT — append-only |
| `TOOLBAR.md` | the rule that a boundary needs a decision in it | design doc; its UI description is SUPERSEDED, its REASONING is still the governing rule — kept active, not archived (re-affirmed 2026-08-16) |
| `WORKFLOW.md` | what each stage *does* | STALE in part — its reach-it-by instructions are wrong; the one specific "not yet fixed" note it carried is itself superseded by the Aug 14 toolbar rebuild (checked 2026-08-16) |
| `archive/*` | historical record, moved 2026-08-16 from repo root: `WORKPLAN.md`, `RUNSHEET.md`, `HANDOVER-Q4F26-DRAFTING.md`, `WORKED-EXAMPLE-STRATEGIC-ALIGNMENT.md` | HISTORICAL — never edit to match today. Anything still-useful from these was migrated into `SYSTEM-OVERVIEW.md` or `CHECKLIST.md` with provenance before archiving. |
| `specs/*`, `IMPLEMENTATION_PLAN.md`, `FIRST-REAL-RUN.md`, `CYCLE-FINDINGS-*`, `SPIKE_NOTES_*` | the design exchange and its findings, as they stood | HISTORICAL by design — never edit to match today |

**The bridge folder** (`OneDrive\Claude\`) is transport, not storage. Documents sent
to or received from the chat side live there; anything durable is copied into this
repo. It is not covered by the checker.

---

## 3. The review protocol

`python3 vba/tests/check_docs.py` — exit 0 clean, 1 with findings, 2 if it could not
run. It checks five things, and **each has been made to fail on purpose**:

1. **Dead button caption** — a doc names a button that is not on the toolbar, and
   does not say it was removed
2. **Wrong column letter** — a doc names a letter for a role that contradicts the
   `COL_D_*` constants
3. **Path does not exist** — a doc cites a file that is not there
4. **Symbol not in source** — a doc cites `Module.Thing` that no module defines
5. **No status marker** — a doc does not say how current it is

**When it runs:**

- before committing any `.md` change — it is seconds
- at the start of any session that will rely on the docs to make a decision
- immediately after any rename, any toolbar change, any column change
- after a build, alongside `check_vba_static.py` and `check_module_lists.py`

**When a finding appears, fix the document, never the check.** The one legitimate
exception is a false positive proven as one — and then the fix is to narrow the
check and record why, as was done on 14 Aug when a first version reported 77
findings including correct ones. A checker that cries wolf is worse than no checker
(`AGENTS.md`, 2026-07-31).

**What it cannot check, and therefore what a human review is still for:**

- whether a statement is *true* — only whether it agrees with the code
- whether a document is still *useful*
- claims about the deck or the register, which live in files the checker cannot open
- reasoning, judgement, and everything in `NEXT-SESSION.md`'s narrative

That last gap is real: *"three projects' text is not recoverable from any backup"*
passed every mechanical check and was false, because the text was on the slides.
**No check replaces opening the file.**

---

## 4. Why this exists

On 14 August 2026, in one morning:

- `WORKFLOW.md` gave the drafting columns as G and I — layout 3, two layouts stale,
  pointing a person at the character-count column to put their tick in
- `TOOLBAR.md` described three toolbar buttons; there are two, and 17 of 19 caption
  constants named buttons nobody could press — four of them reached users as
  "Press '<button>'" remedies
- `FIX-LIST.md` said the register workbook is one "the tool rebuilds and clears".
  It is not: `Register`, `Field Spec` and `Sources` are PERMANENT. That single wrong
  sentence was repeated twice more in fresh documents before anyone opened
  `LifespanOf`
- `HANDOVER-Q4F26-DRAFTING.md` said three projects' text was "not recoverable from
  any backup". It was on the slides
- `SOURCE-CAPTURE-FORM.md` told a person to put source IDs in column G, which is the
  approve column
- `WorkbookBridge.DescribeSheet` — in the *code* — told a person to type into D and
  tick E, and described a register model retired on 3 August

Rohan found the first one by asking a plain question. He then said: *"read through
notes on the project etc and fix dodgy records like that — it's not my tool, it's
yours."* This file and its checker are the answer to that.
