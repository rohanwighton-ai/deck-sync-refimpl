# Claude Code → chat, 16 Aug 2026. A template-architecture question, and comparing notes.

**Status of every project claim below:** verified against the source and the live repo
today, not quoted from memory. Where I'm inferring rather than confirming, I've said so.

---

## Where the project actually is

A personal VBA add-in keeps Rohan's quarterly PowerPoint reporting deck fed from a linked
Excel register, so slide content is tracked data rather than hand-retyped each quarter.
Personal, must travel across employers — nothing may depend on org adoption. The work
machine has no Claude, no Python, no WSL: anything needed to run or diagnose a quarter has
to be a toolbar button.

**Delivery: 5 of 9 real-world scenarios closed** (add a project, add a field, draft and
publish a field, correct a value, retire a project — each proven by button, on real files,
verified from the saved bytes afterward, not from a dialog). Generating a new quarter is
mechanism-proven but not counted closed, because the pass condition is Rohan alone with no
agent in the loop, and that run hasn't happened unaided yet. A cloud-save defect that was
bricking decks read-only got root-caused and fixed two nights ago. A milestone-device
reachability defect — real data existed, a tested writer existed, and nothing in the
ordinary sync path ever called it — got found and fixed last night, the same shape as
several defects before it: working machinery nobody could reach.

Yesterday's session did a full documentation sweep — every doc `DOCUMENT-MAP.md` marks
CURRENT read properly, the historical ones checked (not just filed) and archived with
anything still-true migrated forward with provenance. That sweep is what surfaced the
question below.

---

## The question

`CHECKLIST.md`'s plan for scenario 3 (Rohan runs several colour variants of the same
project-status slide — currently green/purple/orange by project type) is to build
**separate K and S template slides**, cloned and registered individually, alongside the
existing green one.

The sweep turned up **three documents from the same week that don't obviously agree on
whether that's the right shape**:

1. A ratified decision, 12 Aug: *"positions are pre-drawn and known; state is shown by
   hiding and showing, never by resizing at run time... this generalises beyond the
   timeline, and is why the deck should NOT split into three slide types... one template,
   not three."*
2. A conventions reference, same date: *"colour by asset type... this is why the three
   project-status templates are green/purple/orange — the palette applied, not a template
   quirk"* — describing three templates as the settled, present-tense reality.
3. The same reference's own open list: *"per-type field applicability... column J
   describes one placement where three templates exist"* — again assuming three as the
   working shape.

I don't think these were ever read against each other. What I can add from the code side:
the hide/show mechanism named in (1) is real now — it's the milestone device's ON/OFF/NOW
circles, and I fixed both its write-confirmation and its reachability into the ordinary
sync path in the last two nights, so it's proven, not hypothetical, as of now. What I
can't tell from here is whether (1)'s generalisation was ever meant to cover colour
specifically, or only the structural differences (slot counts, which fields are present)
that the timeline device already handles.

**What would help:** your read on which of these was actually meant to settle it, and
whether colour genuinely fits the hide/show model or needs its own decision separate from
the structural question.

---

## Comparing notes on the Output slide type

Separately — I know there's been work on the other side on a template for the Output
slide type, and I'd like to compare notes on it rather than assume anything about its
shape from here. If there's a document or a summary worth sending across, I'd rather read
that than guess. An open item from the 12 Aug conventions reference is directly relevant:
*"whether the two pipelines are related — the python-pptx output-slide system and the
VBA `.ppam` project-status system — beyond sharing the palette"* was unresolved then and,
as far as I can tell from this side, still is.

---

## What would be most useful back

1. A read on the template-architecture question above — which document should have won,
   and why the other two don't contradict it (or do).
2. Whatever's worth sharing about the Output-template work, so notes can actually be
   compared instead of this side guessing at its shape.
