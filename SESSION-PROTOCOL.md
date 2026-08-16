# Session protocol — mandated, not optional

> **CURRENT. Read this before touching anything else in the repo.** Written
> 2026-08-16 after a full documentation sweep found a real, ratified architecture
> decision (12 Aug) sitting undiscovered for four days because no prior session had
> done what this file now mandates. `NEXT-SESSION.md` says "read first, always" —
> that is necessary but not sufficient. This is what "read first" actually means.

## 1. Familiarization order — every session, no exceptions

1. **`DOCUMENT-MAP.md` first, before `NEXT-SESSION.md`.** It is the index of what
   is actually current versus historical, and it is the tool that catches drift —
   it has a checker (`check_docs.py`) behind it for a reason. Skipping straight to
   `NEXT-SESSION.md` is how a real decision sat unread for four days on
   2026-08-16 (see `CHECKLIST.md`'s "ARCHITECTURE FORK" item, found only because
   a full sweep was finally done).
2. **`CHECKLIST.md`** — the primary handover surface. Flat, tickable, linked to
   source. This is "what's actually left," not `NEXT-SESSION.md`'s narrative.
3. **`NEXT-SESSION.md`'s CURRENT block only** — the top block, not the whole
   history underneath it. The rest is a journal, useful for *why*, not *what's
   left*.
4. **`SCENARIOS.md`** — the pass/fail frame. Read the harsh pass condition at the
   top every time (*"Rohan completes it unaided... anywhere he has to ask is a
   defect"*) — it is easy to let this soften into "Claude helped, close enough."
5. **`FIX-LIST.md`** — known-broken, not yet fixed. Check before diagnosing
   anything that feels new; it has been rediscovered from scratch at least three
   times before this file existed.
6. **`SYSTEM-OVERVIEW.md`** — only if genuinely new to the project. Skip if
   already oriented.

**Do not treat any document's stated status as final without checking it against
the code when a decision hangs on it.** `DOCUMENT-MAP.md`'s own classifications
have been wrong before (see `WORKFLOW.md`'s one live-looking note that turned out
to already be superseded, caught by grep rather than assumed). A document says
what someone believed when they wrote it. The code says what is true now.

## 2. Documentation upkeep — the discipline, not just the rule

**Before archiving anything:** read it fully (or, for a large homogeneous batch
where individual reading is genuinely disproportionate, sample it — but say so
explicitly, name what was sampled, and give the reasoning, the way the 2026-08-16
sweep did for `specs/*` and `SPIKE_NOTES_*`). Extract anything still true and not
already captured elsewhere. Migrate it into a CURRENT document **with provenance
noted** — which source document, which date — before moving the source to
`archive/`. Never silently drop a still-true fact because its container went
stale.

**When you find a genuine open question or unresolved decision while reading a
historical document** (this happens — see the 12 Aug architecture fork, or the
team-distribution question found in `FIRST-REAL-RUN.md`), **do not resolve it
yourself.** Surface it plainly, add it to `CHECKLIST.md`, and let Rohan decide.
Archiving is not a license to quietly settle open questions on the way past.

**After any documentation change — not just code changes:**
```
python3 vba/tests/check_docs.py
python3 vba/tests/check_vba_static.py
```
Moving files can silently break a checker's own logic. The 2026-08-16 `specs/`
move broke `check_docs.py`'s historical-exemption regex (`^specs/` stopped
matching once the directory became `archive/specs/`) — caught only because this
was run before committing, not after.

**When a document's status changes or a file moves, update `DOCUMENT-MAP.md`'s
table in the same commit.** An index that lags the thing it indexes is worse than
no index — it reads as authoritative and isn't.

## 3. Storing test results — the actual mechanism, not an aspirational one

There is no `LAST_TEST_RUN.md` or equivalent tracking file, and do not invent one.
A prior prompt template (now archived, `archive/PROMPT_build.md`) referenced one;
it was never actually built or used. Adding a second place to record results is
exactly the "second copy of a machine-knowable fact" anti-pattern this project has
been burned by repeatedly — a tracking file drifts the moment someone forgets to
update it, and nothing enforces it.

**The actual, correct mechanism, already in use and mandated going forward:**

1. Run `check_vba_static.py`, then the full suite (`run_vba_tests.ps1`), before
   every commit that touches `.bas` files. Both close Office first if it's open.
2. **Quote the exact pass/fail counts in the commit message itself.** Git history
   is the durable record — `git log` is the real "test results over time," and it
   can't drift because it's immutable once pushed.
3. For a **fix**, prove it properly before committing: make the test fail on
   purpose first (revert the fix, confirm the test catches the original defect),
   then restore the fix, then commit. State this in the commit message too. A
   green suite alone is not evidence a test can fail — see
   `feedback_assiduousness_by_default` memory and this project's own repeated
   "a check that cannot fail looks exactly like one that passed" lesson.
4. For a **known-broken item**, it belongs in `FIX-LIST.md`, not a separate test
   log — with the evidence (what was measured, when) inline.

## 4. Working with Rohan on this project specifically

His general working style, communication preferences, and standing feedback live
in his global `CLAUDE.md` — read that once per account, not per project. This
section is only what's specific to working **on deck-sync**, observed directly
across the 2026-08-15/16 sessions. Don't duplicate the global profile here; it
will drift independently of the real one.

- **He drives the keyboard; you verify.** The scenario pass condition is
  literal — he presses the buttons, you read the saved file afterward and report
  what actually happened. Doing it for him defeats the point, even when it would
  be faster. See "Press the button yourself" memory.
- **A workaround is not an answer to "I need it."** When he says he needs
  something demonstrated or chased, a clever way around the obstacle is not the
  same as finding out why the obstacle exists. The milestone-visibility gap
  (2026-08-15 night) is the example: a wrapper function would have shown him a
  circle move; what he actually wanted was why the sync loop couldn't reach the
  device at all, and that turned out to be the more important finding (FIX-LIST
  R). When in doubt, chase the real cause.
- **He'll park a hard question rather than need it resolved immediately** — but
  only if it's genuinely written down, not dropped. "We can answer that later" is
  real permission to move on, not permission to forget. Confirm it landed on
  `CHECKLIST.md` before moving on.
- **He wants the honest count, not the flattering one.** Scenario 1's mechanism
  being proven and the scenario count staying at 5/9 are both true at once, and
  he wants both said plainly, in that order, without either one softened.
- **Long live sessions get honestly fatiguing** ("I'm burnt" — 2026-08-15,
  after several hours of live debugging plus a documentation sweep). When a
  session has been running long and the natural unit of work is done, say so
  plainly and offer a clean stopping point rather than starting something new.
  This is a project-specific instance of the global "prefer long sessions,
  don't refresh" guidance, not a contradiction of it — the point is noticing
  fatigue in the moment, not ending sessions early by default.
- **Plain questions are architecture probes, not requests for a definition** —
  already in global memory (`feedback_plain_questions_beat_the_test_suite`), and
  it has been true on this project specifically, repeatedly: "why not one sheet,"
  "have you toggled visibility through inject," and "well you tell me, what's the
  best answer" have each either found a real defect or forced a genuine design
  decision. Treat them that way, not as reassurance-seeking.
