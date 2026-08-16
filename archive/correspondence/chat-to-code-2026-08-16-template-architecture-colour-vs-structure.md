# chat → Claude Code, 16 Aug 2026. Task 1: which document settles scenario 3.

**Answers:** "which of these was actually meant to settle it, and whether colour genuinely fits
the hide/show model."

**Settled here:** none of the three settles it, and they do not contradict each other. The
document that settles it is the 15 Aug per-letter plan, which postdates all three. Colour does
not fit hide/show and was never inside decision (1)'s scope.

**Still open:** one condition on the answer, in §5, that should be recorded with the decision
or this argument comes back in a month.

I cannot see the repo. §1–§3 are verified against the 12 Aug session record on this side.
§4–§5 are reasoning, marked as such.

---

## 1. Decision (1) was scoped to STRUCTURE, and the record shows how it got there

Your instinct is right, and it is checkable rather than a judgement call.

The 12 Aug decision was reached by *enumerating the type differences* and finding they were all
one shape. The enumeration, verbatim from the decision record on this side:

> K has 6 timeline circles vs P's 7; K's subtitle carries no Sector or TRL segment; S has a
> fourth team row (Research Supervisor).

Slot counts, and which fields are present. Exactly the two things you guessed. **Colour is not
in the list**, and the list is not illustrative — it is the whole basis for the conclusion. The
generalisation was "every type difference *found* is the same pattern", and colour was not
among the differences found.

## 2. The same decision explicitly carves colour OUT of hide/show

This is the part that makes it not a matter of interpretation. From the same 12 Aug ruling:

> Colour MAY be applied from a declared spec (applying a stated rule is not the same as
> inventing formatting). SIZE stays visibility-driven.

So the ruling names two mechanisms in one breath and assigns colour to the *other* one. Reading
(1) as covering colour requires reading past a sentence that says it does not.

## 3. Documents (2) and (3) are not disagreeing — they are wrong about a fact

(2) and (3) describe three project-status templates as the settled present-tense reality. They
state the *palette* rule correctly (P green, S purple, K orange, "the palette applied, not a
template quirk") and then assume the artifacts exist.

They do not. As of 15 Aug the live deck has **exactly one** `is_template` slide and it is the
green `P` one — which is the whole of scenario 3. So (2) and (3) are a documentation defect
against present state, not a design position competing with (1).

Worth noting for the sweep: `SCENARIOS.md` also carried a version of this, asserting three
templates already collide when only one exists.

## 4. The reconciliation: "template" is a word doing two jobs

This is your own rule and it applies cleanly here. `Template` currently means both:

- **the one slide type / the one design** — what (1) is about, and
- **a registered `is_template` slide** — what (2), (3) and `CHECKLIST.md` are about.

Split those and nothing disagrees:

| | |
|---|---|
| **One TYPE** | decision (1). Still stands, untouched. |
| **Three TEMPLATE SLIDES** | (2), (3), `CHECKLIST.md`. Variants of that one type. |

The 15 Aug plan already encodes this as the axis rule: **type decides which FIELDS exist;
variant decides how it LOOKS and owns nothing of its own.** The test is there too — adding the
orange template must touch **zero** field definitions. If the implementation ever wants the 29
fields declared three times, that is the design failing, and it is also the moment decision (1)
would genuinely be violated.

So `CHECKLIST.md`'s plan is right, and (1) does not forbid it. Building K and S templates is not
splitting into three types.

## 5. Colour: cloning is right, but not for the reason `CHECKLIST.md` gives — and it is conditional

Three ways colour could have gone, and the honest comparison:

**Hide/show — three stacked shapes per coloured element, show one.** What would follow if (1)'s
generalisation covered colour. Triples the shape count on every coloured element and hands the
device machinery three parts where it needs one. **Already rejected once in the analogous
case**: the output slide's status bar, where "three stacked and switched" lost to "draw one",
on the reasoning that splitting later is a duplicate-and-recolour (cheap) and drawing three and
being wrong is a delete.

**Recolour at runtime from the declared palette.** Explicitly *permitted* by §2, and cheapest in
shapes. It fails on something else: it needs a per-shape declaration of *which* shapes carry the
type colour. That declaration does not exist, and building it means an invisible per-shape
attribute — against "declaration beats perception, repairability beats robustness", the same
reasoning that chose shape names over hidden tags in the naming thread.

**Three drawn templates.** Each drawn once, correctly, by hand. Repairable by eye in the
Selection Pane on a machine with no tooling. Zero runtime colour machinery. Costs three slides
of drawing and, thanks to per-letter registration, nothing in the field definitions.

Third one wins — **on repairability, and only because per-letter registration keeps it a
variant.** The colour difference by itself is an argument *for* runtime recolour, not against
it.

**The condition, and please record it with the decision:** cloning wins *given that no
colour-role declaration exists*. If one is ever built — and the parked `$FIELD[#INDEX][.PART]`
naming thread would arguably supply exactly that — the calculus flips and three drawn templates
become three things to keep in sync for no gain. Recorded as a permanent truth, this becomes the
fourth document in this pile a month from now.

---

## 6. What I would do to the three documents

1. **(1) stays ratified, unchanged**, with a scope note added: *structure only — colour is
   governed by the declared-palette carve-out in the same decision.*
2. **(2) and (3) get corrected to present tense**: three templates are *planned* under scenario
   3, one exists today. Their palette content is correct and stays.
3. **Add the type/variant vocabulary** so the next reader cannot make this mistake: *type*,
   *variant*, *template slide* — three words, three meanings, no overlap.
