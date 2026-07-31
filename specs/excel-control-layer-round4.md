# Excel Control Layer — round 4: answers, two amendments, and the workplan

**In reply to:** `Excel_Control_Layer_Confirmation.md`, 31 July 2026
**From:** the PowerPoint / VBA side, against commit `452f952`

D1–D12 accepted as settled except where §2 amends them. Everything in §1 was measured
today against real Office rather than reasoned — including two things I would otherwise
have asserted and got wrong.

---

## 1. Confirmations — measured, not assumed

| # | Question | Answer | How |
|---|---|---|---|
| Q5 | Field inventory | **5 managed, 77 untracked, 38 candidates** | `Audit Fields` run on the real deck |
| Q6 | Do tags clone? | **Yes — slide tags and shape tags, on duplicate AND across decks** | live probe |
| Q6b | Native discriminator? | **Partially — see below** | live probe |
| Q8 | Is a rename safe? | **Yes — overwrite, enumerable, reversible** | live probe |
| Q12 | Pipes in content? | **Zero occurrences** | scanned all slides + workbook + audit |

### Q6 — tag cloning, measured

```
source slide:            SlideID 257, instance_key 'PROBE-001', shape role 'PROGRESS_BODY'
after Slide.Duplicate:   SlideID 258, instance_key 'PROBE-001', shape role 'PROGRESS_BODY'
after paste to new deck: SlideID 257, instance_key 'PROBE-001', shape role 'PROGRESS_BODY'
```

**D5 is confirmed.** Identity tags clone exactly as names do. Duplicating a project slide
produces two slides both claiming the same EntityCode, and copying a deck carries the
declared period with it. The uniqueness check is required.

**On "is there anything PowerPoint-native that distinguishes a copy" — partially, and the
split matters for how much the check has to carry:**

- **Within a deck: yes.** `SlideID` is reassigned on duplicate (257 → 258). So two slides
  claiming one EntityCode inside one deck are natively distinguishable, and the check has
  help.
- **Across decks: no.** A slide pasted into another presentation kept SlideID **257**. So
  SlideID is not a discriminator across files — and it is not reliably *stable* either,
  since it was only free to keep 257 because the target deck was empty. **Do not treat
  SlideID as a cross-deck identifier.**

Consequence: for the copied-deck case — which is the likely one, since starting next
quarter from last quarter's file is the obvious workflow — **the uniqueness and period
checks are the whole defence.** R9 should be treated as load-bearing, not belt-and-braces.

### Q12 — pipes

Scanned every slide's text runs, the workbook's shared strings, and all 77 audit rows.
**Zero occurrences of `|`.** `||` is safe. Worth re-running this check before any bulk
import of new prose, since it costs seconds.

---

## 2. Two amendments to D6 and Q10

### Amendment A — D6's gate is specified against a mechanism that does not exist

D6 requires that a surviving `<<…>>` blocks publication and marks the output draft. Q11
then asks what the draft-marking mechanism is. **There isn't one.** No draft state, no
publishable flag, no watermark. D6 as written cannot be implemented.

The condition is right and the intent is right — reporting alone does degrade into option 1
the first time someone skims a log, exactly as D6 says. But it needs a real mechanism.

**What actually exists, all proven in this build:**

| Mechanism | Status | Assessment |
|---|---|---|
| The placeholder text itself | **working today** | Strongest available. `<<PROGRESS_BODY>>` on a slide is unmissable in a way no log line or filename is. |
| Deck-level custom property | proven (`DeckSyncId` uses it) | Can carry `DeckSyncLastRunPlaceholders = N`. Machine-readable, invisible to the reader. |
| Adding a visible shape | proven | A real watermark is possible. Intrusive, and someone will delete it. |
| Slide hidden flag | proven last night | Available but wrong tool — hiding an incomplete slide makes the problem less visible, not more. |
| Filename suffix | trivial | Weakest. People rename files. |

**Recommendation:** the placeholder IS the draft marker. Back it with a deck property
recording the count from the last run, and make the run report's headline the placeholder
list rather than a summary line. Do not build a watermark until the placeholder has
demonstrably failed to stop someone — inventing a second marker before the first has been
tested is how you end up with two nobody reads.

### Amendment B — Q10: static fields must be UNTAGGED, not "tagged with no row"

Q10 asks whether creation-from-template populates static fields and whether R4's "leave
untouched" holds for them.

**It does — but only if static fields carry no FieldID at all**, and this is a correction
to the model rather than an answer within it.

- Untagged content rides across from the master template as ordinary slide content when a
  slide is created, and is never touched by any subsequent run. Confirmed by code and by
  last night's live cycle.
- **But a field that is TAGGED and has no register row keeps its placeholder** — that is
  exactly the D6 case. So if a static field were given a FieldID and no row, its
  `<<…>>` would survive every run and, under D6, **block publication forever**.

So "static" and "managed" have to be mutually exclusive. The Field Spec sheet's
static-vs-quarterly column should therefore determine **whether a tag is applied at all**,
not merely whether a row is expected. A static field is content on the template; a managed
field is a tag plus a row.

One consequence to accept deliberately: static fields will show up in every `Audit Fields`
run as untracked, forever. That is correct — the audit's job is to ensure nothing is
missed, and "we looked at this and chose not to track it" is a decision the Field Spec
sheet records, not something the audit should learn to suppress.

---

## 3. Q7 — where identity actually lives today

| Thing | Where | Status |
|---|---|---|
| `slide_type` | slide tag | live |
| `instance_key` (≈ EntityCode) | slide tag | live |
| `role` (≈ FieldID) | **shape** tag | live |
| `DeckSyncId` | deck custom property | live |
| `DeckSyncWorkbookPath` | deck custom property | live |
| `DeckSyncType:<type>` | deck custom property | live — holds `templateSlideID\|worksheetName` |
| `period_key` | slide tag | **specified but never implemented** — no code reads or writes it |

That last row matters for planning: the period concept exists on paper in
`specs/identity-tags.md` and in comments, and nowhere in behaviour. Anyone reading the
specs would reasonably assume otherwise.

**Recommendation for where period goes: deck-level custom property, not a slide tag.**
D4 already says the deck declares its own period, and one property per deck is a single
read and a single thing to check, where a per-slide tag is N things that can disagree with
each other. It also makes the D5 period-plausibility check trivial — one read at run start,
displayed for confirmation.

**Note the scope consequence for R9:** EntityCode uniqueness is a *scan of every slide's
tags*; period plausibility is a *single property read*. Two different checks, different
costs, both cheap.

Also flagged, since §7 of the previous round raised it: the one-deck-per-slide-type
direction will move `DeckSyncType:<type>` and make `SlideID` ambiguous across decks. Q6's
finding that SlideID is preserved on cross-deck paste makes that worse than it looked.
Not urgent; do not bake SlideID into anything durable.

---

## 4. Q8 — the rename migration, now that it has been measured

**The key fact, and it makes this much smaller than it looked: `FieldID` is the *value* of
the `role` tag, not the tag's name.** Renaming `About text` → `PROGRESS_BODY` is overwriting
one string. It is not a structural change.

Measured today:

```
Tags.Add on an existing name:  Count 1 -> 1, value replaced      => OVERWRITES, idempotent
Enumeration (Count/Name/Value): works, no error                  => ENUMERABLE
Set the old value straight back: succeeds, Count unchanged       => REVERSIBLE
```

One detail worth having: **enumeration returns tag names UPPERCASED** (`ROLE`, not `role`),
matching what is stored in the file. Any migration or audit tool must match tag names
case-insensitively. Reads by name are already case-insensitive, which is why nothing has
tripped on this yet.

**The plan:**

1. Excel side authors a mapping table: current tag value → target FieldID. Five rows.
2. A migration tool enumerates every shape carrying a `role` tag across the deck, applies
   the map, and reports anything unmapped rather than guessing.
3. Verification is a `Preview Sync` — if the join still resolves, every field still matches.
4. **Rollback is the inverse map, run the same way.** Measured as reversible.

**Scope today: 5 distinct FieldIDs across 6 slides — 30 tag writes, one deck.**

**What happens to a half-migrated deck** (asked explicitly, and the answer is reassuring):
nothing corrupts. The join is FieldID ↔ role tag, so an unmigrated shape simply fails to
match its register row, and R5 reports it in the named list as "in register, no matching
shape". The field is not injected and nothing else is affected. A half-migrated deck is
*incomplete*, not *wrong* — which is the right failure mode for a migration and is why it
can safely be run in stages.

---

## 5. Q9 — sequencing. My recommendation, and it inverts the obvious order

**Do the FieldID rename FIRST, today, before the Field Spec work lands.**

The intuitive order is rewrite → build → rename, on the grounds that renaming last is safer.
That is wrong here, for one reason that is only visible from the measurements:

> **The rename's cost scales with the number of tagged fields, and every other task on the
> list increases that number.**

- **Today:** 5 FieldIDs, 30 tag writes, one deck, one type.
- **After the Field Spec work onboards the ~38 candidates:** ~43 FieldIDs, ~258 writes, and
  by then plausibly more than one deck.

It never gets cheaper than this morning. And the probes have removed the reasons to fear it:
the operation is an idempotent value overwrite, it is enumerable, it is reversible, and a
partial migration degrades to non-injection rather than corruption.

**Dependency chain:**

```
Q5 dump (DONE)
   └─> mapping table (Excel, ~15 min)
          └─> rename migration (VBA)  ─┐
                                        ├─> §5 one field end to end
   long-format ExcelOutput rewrite (VBA)┘
          └─> Quarter / Status filter (VBA)

R9 uniqueness check (VBA) — independent, no dependencies
```

**Answering the specific worry in Q9:** doing the rename before the rewrite does *not* mean
migrating tags twice. The rewrite changes how the *sheet* is read (wide → long); it does not
change how *fields are addressed in the deck*. Those are separate layers and the rename only
touches the deck. There is no double migration.

---

## 6. Workplan for today

Two lanes. They are independent except at one point, and that point is a 15-minute
handover, not a blocker.

### Excel lane — can start now, blocked on nothing

| # | Task | Input | Output |
|---|---|---|---|
| E1 | Triage the 77 items: field / chrome / drop | `Q5_field_inventory_dump.md` | decided list |
| E2 | **Author the mapping table** — 5 current tags → target FieldIDs | E1's naming decisions | 5-row table → hand to VBA lane |
| E3 | Author the Field Spec sheet from E1 | E1 | one row per FieldID, incl. static-vs-quarterly |
| E4 | Build the register as a ListObject, long format | E3 | `Quarter, EntityCode, SlideID, FieldID, FieldType, Value, CharCount, Status, UpdatedDate` |
| E5 | Build one Copilot template worksheet for the §5 field | E3 | drafting surface |

**E2 is the handover and it is small — do it early**, not in dump order. The VBA lane's
first task is blocked on it and on nothing else.

**E3 must record static-vs-quarterly per field** (Amendment B) — it now determines whether a
tag is applied at all, so it is no longer just documentation.

### VBA lane

| # | Task | Depends on | Size |
|---|---|---|---|
| V1 | **Rename migration tool + run it** | E2 | small — 30 writes, reversible |
| V2 | **R9 uniqueness check** — EntityCode duplicates, period plausibility, reported at run start | nothing | small |
| V3 | **Long-format `ExcelOutput` rewrite** | nothing | **large — the main piece** |
| V4 | `Quarter` / `Status` filtering | V3 | medium |
| V5 | `\|\|` → line break at injection | nothing | trivial |
| V6 | R10: empty `Value` is a validation failure, distinct from no-row | V3 | small |

**Order: V1 (when E2 lands) → V2 → V3 → V4.** V5 and V6 are fill-ins.

V2 is placed early deliberately: it is small, it depends on nothing, and it guards every
subsequent operation that touches tags in bulk — including V1.

### The meet point

§5, unchanged and not negotiable: **one text field taken completely through** — Field Spec
row, template worksheet, register row, injection, verified on a slide — before anything is
generalised.

**First field: `Project Status`. Decided.** Short, its value is visibly different per
project, it is already tagged and syncing today, and it appears on the slide header where a
wrong value is obvious rather than buried in prose. `PROGRESS_BODY` is the more valuable
field but it is long prose with a character-count target, which makes it a poor choice for
proving a pipeline — prove the pipeline on something where "did it work" is answerable at a
glance, then run the hard field through the proven path.

So E2's mapping table needs `Project Status` → its target FieldID **first**; the other four
can follow.

---

## 7. What I am NOT confident about

Stated because the previous round's most useful move was refusing a requirement rather than
building around it.

**The audit's banding is weak on this deck and I have not fixed it.** The chrome-vs-data
heuristic assumes sibling slides were independently authored; four of six here were cloned,
so identical untagged content proves nothing. The strongest verdict fired **zero times out
of 77**. Use the rule "5/5 is chrome, below 5/5 is a candidate" and treat the `Guess` column
as a sort order, not a finding. This will improve on a deck of genuinely hand-authored
slides, which is the wrong way round for confidence.

**The number that should drive scheduling is 5 of ~43, not 113 of 113** — agreed, and worth
repeating because it is the honest measure of how far along this is.
