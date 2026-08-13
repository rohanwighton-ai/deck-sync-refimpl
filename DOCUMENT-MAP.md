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
| `DOCUMENT-MAP.md` | this index and the review protocol | CURRENT |
| `FIX-LIST.md` | what is known-broken and not yet fixed | CURRENT — re-audited 14 Aug |
| `TRACKER.md` | the framing and the manual baseline | CURRENT for framing; defers status to `NEXT-SESSION.md` |
| `COLUMNS.md` | the *reasoning* behind the register's columns | CURRENT — the Field Spec owns the list |
| `PROVENANCE.md` | the provenance design | CURRENT — designed, not built |
| `SOURCE-HARVEST.md` | the paper form for capturing sources | CURRENT — tool-independent |
| `AGENTS.md` | platform gotchas earned the hard way | CURRENT — append-only |
| `WORKPLAN.md` | the 31 July build plan | superseded in part; kept for the record |
| `WORKFLOW.md` | what each stage *does* | STALE in part — its reach-it-by instructions are wrong |
| `TOOLBAR.md` | the rule that a boundary needs a decision in it | design doc; its UI description is SUPERSEDED |
| `RUNSHEET.md`, `HANDOVER-Q4F26-DRAFTING.md`, `WORKED-EXAMPLE-*.md` | method, not steps | HISTORICAL |
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
- `SOURCE-HARVEST.md` told a person to put source IDs in column G, which is the
  approve column
- `WorkbookBridge.DescribeSheet` — in the *code* — told a person to type into D and
  tick E, and described a register model retired on 3 August

Rohan found the first one by asking a plain question. He then said: *"read through
notes on the project etc and fix dodgy records like that — it's not my tool, it's
yours."* This file and its checker are the answer to that.
