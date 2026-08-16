# Excel Control Layer — round 7

**From:** the Excel side, 31 July 2026
**Status:** **both lanes paused by direction from the Research Manager.** No rulings issued in
this document. One request, one correction, one boundary restated.

---

## 1. Stop before starting anything further

Direction from the Research Manager: neither side starts work until both sides are clear on
the whole system. That includes V3, which round 6 §5 has under way as unblocked.

The trigger is that the Excel side has been designing against a picture that does not include
the compiler. It has appeared once in this exchange — round 3 §7, "one deck per slide type
plus a compiler that assembles composites", flagged as blocked on a VBA versus Office
JavaScript question. That single sentence is the entirety of what this side knows about it,
and it has been treated as a future concern rather than as something that shapes the register.

That was a mistake. It may well change decisions already taken as settled.

---

## 2. Request — a full rundown of the compiler

Not a summary. The same treatment round 4 gave the field inventory: what exists, what is
designed, what is speculative, and what has been measured. Specifically:

1. **What it is and what it produces.** Source decks in, composite deck out? What is the
   composite *for* — a board pack, a portal submission, something else?
2. **Is a composite ever a sync target, or is it write-only?** That is, does the tool ever
   read a compiled deck and inject into it, or is a composite a terminal artefact that gets
   rebuilt rather than updated? This is the single most important question in the list for
   the register's design.
3. **Do identity tags survive assembly?** Q6 measured that cross-deck paste clones tags and
   preserves `SlideID`. If the compiler moves slides between presentations, every composite
   inherits the identity of its sources — including duplicate `EntityCode`s the moment two
   source decks are combined. Does R9's uniqueness check apply to composites, and if so what
   is it checking against?
4. **Where does period live in a composite** assembled from decks that may each declare a
   different one? D4 puts period on the deck. A composite is a deck. If it is assembled from
   a Q3 deck and a Q4 deck, what does it declare?
5. **Does the compiler read the register, or only move already-populated slides?** If it
   reads, it is a second consumer of the register and the schema has two clients, not one.
6. **One deck per slide type — does that mean one workbook per deck, or one register across
   all decks?** This is the question that most affects this side. See §3.
7. **Status of the VBA versus Office JavaScript blocker.** Round 3 called it unresolved and
   said it blocks composition only. Still true? And does a move to Office JS change anything
   about tags, custom properties or the Excel interface that has been specified since?
8. **Where it sits in sequence** relative to the text-field pipeline.

---

## 3. Correction — the register is one workbook across all quarters

Stated because a remark in the last exchange implied otherwise, and it needs to be
unambiguous:

**There is one Excel workbook. It holds every quarter, and it is intended to be used
historically** — trend, comparison, and the prior-quarter exemplar that the whole content
approach depends on. It is not one file per quarter and it will not become one.

The deck is not the historical record. Old decks are outputs that happen to persist. If a
deck and the register ever disagree about a past quarter, the register is right.

This is what makes question 6 above matter so much. One deck per slide type is a
presentation decision this side is happy to leave alone — but "one workbook holding all
quarters and all entities" is a property of the record, and it should not quietly become one
workbook per deck as a side effect of splitting decks.

---

## 4. Boundary — this side overstepped and is pulling back

Round 6 §1 asked for portfolio conventions. The reply that followed went further and started
designing deck behaviour: what should happen when someone duplicates a slide, whether copies
should be harvested or blocked, what a human approval step should look like, and — worst of
it — a ruling that `3_P002-2` should not survive as a slide at all.

None of that was this side's to decide.

The legitimate interest is narrow and stops early: **the register keys on `EntityCode`, so
`EntityCode` must resolve to exactly one slide per deck at run time.** That is the
requirement. How the deck guarantees it — prevention, detection, harvesting, a prompt, or
something not yet thought of — is the PowerPoint side's design, and the answer may look
quite different once the compiler is on the table.

Both live cases still need answering by the Research Manager rather than by either of us, and
both are being held pending the compiler rundown:

- a slide copied so that edits could be reviewed before being applied;
- a slide copied in order to add a new quarter.

The second is the one with a register consequence, since the Quarter column is how a new
quarter is meant to be added.

---

## 5. What is held, not withdrawn

Drafted on this side and deliberately not issued, pending the rundown:

- answers to F1–F4, the four format literals;
- the EntityCode mapping table;
- any ruling on `3_P002-2`;
- E1's triage of the 77 items, and E3–E5 downstream of it.

Round 5 stands as accepted on both sides and nothing in it is retracted. It may need
revisiting once the compiler is understood, which is the reason for the pause rather than an
argument against anything in it.

---

Nothing to build until §2 is answered.
