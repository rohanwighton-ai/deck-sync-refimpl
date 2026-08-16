# Deletions — what got killed, by what question, and what it would have cost

CURRENT — created 2026-08-16, prompted by chat's win-ledger read
(`archive/correspondence/chat-to-code-2026-08-16-win-ledger.md`): *"a decision
leaves a trace in the workbook, a rejection leaves nothing... every one of these
is a saved evening or a saved week and none of them moved a counter."* This file
is that trace, applied to build scope instead of register decisions.

Only actual kills go here — things built, drafted, or seriously planned, then
removed. Not: features never started, ideas mentioned once and dropped.

| Date | What was killed | The question that killed it | What it would have cost |
|---|---|---|---|
| 14 Aug | Packed-cell milestone format (`label~date~Y\|\|...`) | *"wouldn't the timeline be one column per interval not one row?"* | An invisible-tag problem rebuilt inside a spreadsheet cell — the exact defect class shape names were chosen to avoid, re-introduced one layer down. |
| 14 Aug | Propagation tool, as originally scoped | — | Orphaned 46 slides before being hardened via `ExistingInstanceKey`; the original scope is what made propagation unsafe to run. |
| 14 Aug | The rebuild guard / restore-and-carry machinery | *"it's a build, not a rebuild"* / *"why is clear still happening?"* | Five separate patches (1–14 Aug) papering over the same destructive `ws.Cells.Clear`, still losing 27 drafted paragraphs. Deleting the call, not backing it up better, was the fix — see `DOCUMENT-MAP.md` §0.1. |
| ~14 Aug | The three-button split | *"we are simplifying, why split?"* | A button standing in for a confirmation that should have done the job — recorded as `DOCUMENT-MAP.md` §0.4. |
| 14 Aug | Traffic-light bar rule | (recorded with reasoning at the time — the model for how the rest of this file should have looked) | — |
| ~14 Aug | Five steps called "weeks away" and cut from a plan | Rohan challenging a handover-derived size estimate | Three of the five turned out to be already built (`BuildBatchPlan`, `BuildBatchPlanFromMarkedFields`, `MakeTemplateFrom`) — the cost was a mis-scoped plan, not lost work. See `DOCUMENT-MAP.md`-adjacent handover discipline. |
| 16 Aug | Category exemption for a check that only knew one addressing mode | *"I'm not sure I understand the difference between that and a field"* | Would have silently switched three real checks off across two dozen columns. |
| 16 Aug | Inferred "same construction" prompt for track-vs-remainder | *"can't they both be the same if we make them?"* | The two constructions were visually identical — the question the prompt existed to ask had no answer worth asking for. |

**Not yet logged, flagged by the win-ledger read as a real gap:** correct
refusals — the ordered-fallback rule that refuses naming type+letter rather
than silently cloning a green template onto a purple project is a *prevented*
defect, and prevented defects don't have a "what it would have cost" line the
way a killed build does. Chat's framing: a correct refusal naming no
alternative is indistinguishable downstream from a dead end. No entry format
proposed yet for this class — open.
