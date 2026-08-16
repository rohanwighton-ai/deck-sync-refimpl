# chat → Claude Code, 16 Aug 2026. Task 3: the win ledger, and why it's been leaking.

**Answers:** Rohan's observation that wins on this project haven't been recorded properly.

**Settled here:** a rule for what counts as a win, and the ledger as chat can reconstruct it.
**Still open:** two whole classes of win that currently leave no trace at all (§4), and a
denominator that keeps moving (§5).

**Every line below is dated and is chat's reconstruction from documents, not read from the
repo.** Standing rule applies — treat these as hypotheses and correct anything wrong. Where I am
unsure I have said so rather than rounding up.

---

## 1. The rule for what counts

The project already has one and it is stricter than most:

> Proven **by button**, on **real files**, verified from the **saved bytes** afterward — not
> from a dialog.

A green suite is not a win under this rule, and that is correct. The testing rule this project
learned the hard way: *a unit test asks whether a part behaves when called; nothing asks whether
a person can cause it to be called.* Every defect that cost a day — picture injector, progress
bars, publish field selection, cleared tick, propagation button, and now the milestone device —
was working machinery nobody could reach, with ~193 green tests throughout.

So the ledger has two columns that matter: **reachable** and **verified from bytes**. Everything
else is a supporting note.

---

## 2. Closed by button, on real files, verified from saved bytes

| Date | What | Evidence |
|---|---|---|
| 13 Aug | `KEY_EVENTS_BODY` synced register→slides, **21 slides** | saved `.pptx` bytes |
| 14 Aug 08:21 | `PROGRESS_BODY` synced register→slides, **43 slides** | saved `.pptx` bytes |
| 14 Aug | `\|\|` proven to convert to real paragraph breaks | inspected in the saved file |
| 14 Aug 20:45 | **Publish drafting→register by button** — `3_P001` / `KEY_EVENTS_BODY` | the step that had never run through the tool |
| 15 Aug ~08:40 | **Value corrected on 8 real slides by button** | saved `.pptx` vs pre-write backup |
| by 15 Aug ~13:00 | **5 of 9 scenarios closed** — add a project, add a field, draft and publish a field, correct a value, retire a project | each by button, each verified from the file |

**Not counted, correctly:** generating a new quarter. Mechanism-proven, but the pass condition
is Rohan alone with no agent in the loop, and that run hasn't happened. Holding this open rather
than claiming it is itself the discipline working.

---

## 3. Defects found and killed

| Date | Defect | Why it mattered |
|---|---|---|
| 14 Aug | Rebuild destroyed **27 drafted paragraphs** | fix was to stop destroying, not to back up better — became the "nothing is destroyed, only superseded" rule |
| 14 Aug | Publish had **never seen a tick** — rebuild cleared `COL_D_APPROVED` immediately before publish read it | `carryThisRow`: approval travels with the text it approves, byte-identical test |
| 15 Aug | Review grid rebuilt under a live AutoFilter didn't clear — duplicate rows and **13 approvals no human made**, on the sheet publish reads by change id | silent false approval on the authoritative path |
| ~14 Aug | Cloud-save defect: `SaveAs`-to-self **bricking a cloud-hosted deck read-only** | root-caused and fixed |
| 15 Aug | Milestone-device **reachability** — real data, a tested writer, and nothing in the ordinary sync path ever called it | same shape as the five before it; the hide/show device is now proven, not hypothetical |
| 26 Jul | Propagation orphaned **46 slides** | hardened via `ExistingInstanceKey`; the reason propagation is now safe to run |

The pattern in that table is the win. Five of six are *reachability*, not correctness.

---

## 4. The two classes of win that leave no trace — this is the actual leak

**Deletions.** The project's own rule already names this: *a decision leaves a trace in the
workbook, a rejection leaves nothing.* Code has observed repeatedly that **Rohan's plain
questions have deleted more work than they created**. Known casualties, none of which has a
record anywhere:

- the packed-cell milestone format
- the propagation tool (as originally scoped)
- the rebuild guard
- the three-button split
- the traffic-light bar rule (retired 14 Aug — this one *was* recorded, with reasoning, and is
  the model for how the rest should have been)
- five steps called "weeks away" and cut, three of which turned out to be already built

Every one of those is a saved evening or a saved week and none of them moved a counter.

**Correct refusals.** The ordered-fallback rule for template resolution — *refuse naming the
type and letter, rather than silently cloning a green template onto a purple project* — is a
prevented defect. Prevented defects are invisible by construction. So is the standing rule that
a correct refusal naming no alternative is indistinguishable downstream from a dead end.

**Suggested fix, cheap:** one file, `DELETIONS.md`, one line per killed thing — what it was, who
killed it, the question that killed it, what it would have cost. It is the same argument as
recording rejections with reasoning, applied to build scope.

---

## 5. Why the count keeps looking wrong: the denominator moves

On 14 Aug there were **seven** scenarios. On 15 Aug there were **nine**. On 16 Aug they are
named in plain language and not numbered at all.

That is not a mistake — the frame genuinely got better — but it means "5 of 9" is not comparable
to anything recorded before it, and progress reads as flat when it isn't. Two options:

- **freeze the nine** as the delivery denominator until the quarter closes, and record any
  further reframing as a separate note; or
- **drop the count** and keep only the dated list in §2, which is comparable regardless of how
  the frame changes.

Chat's view: the second. The count exists to answer "is this drifting", and the drift tell is
already defined more precisely — *artifacts accumulating while delivery doesn't move*. A dated
list of byte-verified closures answers that directly. A fraction with a moving denominator
doesn't.

---

## 6. Supporting work — real, but not wins by the rule above

Recording these separately so the ledger doesn't inflate:

- Suite **191 → 203 passing / 0 failing**
- `CodeLetterOf` built and tested 15 Aug, 10 assertions
- Field Spec built out **6 → 48 fields**, all columns populated
- Contribution scale conversion, **21 cells / 103 lines**, recorded in S01 so it is never redone
- **370 milestone rows** reconciled, `K1009`/`K1023` naming resolved
- Roll-forward spec covering all 48 fields, five treatments
- Ten taggable standing-field shapes mapped; **nine of ten pixel-identical across all 44 slides**
- `SUBTITLE_A/B` validated against a real slide (`3_P001`); `STATUS_BADGE` validated on `3_K016`
- `OUTPUT-template.pptx` built, 133 shapes, every name unique
- Full documentation sweep, 15 Aug — historical docs *checked*, not just filed, still-true
  content migrated forward with provenance

---

## 7. Chat's own error record, since it belongs in the same ledger

Three wrong operational claims to date, all from reasoning about code from documents rather than
from source. The most recent: asserting scenario 5 was still open from a handover already stale
when read. Cost five minutes, because the standing discipline caught it — *handovers describe
code and go stale within hours; date-stamped documents are not present state.*

Two chat claims that **were** verified true against source: the §0.1 name-reversal flag and the
`TemplateAudit.bas:377` note.

Keeping both halves visible is the point. A ledger with only wins in it stops being evidence.
