# Excel Control Layer — round 4

**In reply to:** `Excel_Control_Layer_Confirmation.md`, 31 July 2026
**From:** the PowerPoint / VBA side
**Contains:** answers to Q5–Q12, two amendments, the sequencing call, the workplan,
and the full field inventory as Appendix A. Single file — nothing else to read alongside it.

---

## What I need back — read this first

Four things, in priority order. Only the first has to come back as *content*.

**1. The E2 mapping table. This is a deliverable, not a question.**
Five rows: current tag value → target FieldID. `Project Status` first — it is the field
chosen to go end to end (§6). Confirming that a mapping table *should exist* does not
unblock anything; V1, the rename migration, does not start without the actual rows.

**2. Amendment B — static fields must be UNTAGGED** (§2). This one has a downstream cost if
it is skipped: it changes what the Field Spec sheet's static-vs-quarterly column *does*, so
it is needed before E3 is authored, not after.

**3. Amendment A — D6's publication gate has no mechanism behind it** (§2). Needs a
decision; blocks nothing today.

**4. Confirm the sequencing in §5.** Rename-first only works if E2 is scheduled early rather
than in dump order.

**What "finished" looks like:** the table, a yes or a push-back on each amendment, and
sequencing confirmed. No new questions are needed. If this comes back with fresh design
questions instead of E2, that is the signal the exchange has stopped converging — at which
point we should stop writing documents and take the one field through end to end, which is
what §5 of the original spec says to do and it applies to this conversation as much as to
the code.

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
| E1 | Triage the 77 items: field / chrome / drop | **Appendix A** (below) | decided list |
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


---

# Appendix A — Q5, the full field inventory

Source: `Audit Fields` run against the real cycle deck, 31 July 2026.
Subject slide: the master template (a copy of 3_P001). Compared against 5 sibling slides.

## A. Fields the tool manages today (tagged)

| Current tag (`role`) | Notes |
|---|---|
| `Project Name` | live tag, in use, sync-managed |
| `Project number` | live tag, in use, sync-managed |
| `Project Status` | live tag, in use, sync-managed |
| `About text` | live tag, in use, sync-managed |
| `events text` | live tag, in use, sync-managed |

**5 managed fields.** These are the tags that R8's rename migration would touch.

## B. Untracked text items (no tag)

`Seen on` = how many of the 5 sibling slides carry identical text.
5/5 means universal across siblings — almost certainly fixed furniture.
Anything below 5/5 varies between slides and is a field candidate.

**Read the banding with the caveat in §C.**

| # | Seen on | Band | Shape | Inside group | Text |
|---|---|---|---|---|---|
| 1 | 3 of 5 | CANDIDATE | `Text 110` | Group 50 | Industry partner (withheld) |
| 2 | 3 of 5 | CANDIDATE | `Text 112` | Group 54 | University of South Australia (UniSA) |
| 3 | 3 of 5 | CANDIDATE | `TextBox 206` |  | Last reported quarter update – Q1F26 _x000D_Work has shifted to consolidation of outputs, … |
| 4 | 3 of 5 | CANDIDATE | `Text 4` | Group 55 | an industry partner (withdrawn) · UniSA · Animal health / Livestock · TRL 3–5 |
| 5 | 3 of 5 | CANDIDATE | `Text 6` |  | 80% |
| 6 | 3 of 5 | CANDIDATE | `Text 12` |  | Successful commercialisation of MgO-based treatments would give Australian producers a fir… |
| 7 | 3 of 5 | CANDIDATE | `Shape 16` |  | ~$280K |
| 8 | 3 of 5 | CANDIDATE | `Shape 16` |  | ~$280K |
| 9 | 3 of 5 | CANDIDATE | `Shape 16` |  | ~$1.4M |
| 10 | 3 of 5 | CANDIDATE | `Shape 16` |  | ~$1.9M |
| 11 | 3 of 5 | CANDIDATE | `Text 212a` | Group 211 | 30 Oct 2023 |
| 12 | 3 of 5 | CANDIDATE | `Text 216a` | Group 211 | 30 Sep 2026 |
| 13 | 3 of 5 | CANDIDATE | `Text 219a` | Group 211 | 90% |
| 14 | 3 of 5 | CANDIDATE | `Shape 229` |  | 3_P001 Timeline |
| 15 | 3 of 5 | CANDIDATE | `Text 11` | Group 154 | Project initiated |
| 16 | 3 of 5 | CANDIDATE | `Text 12` | Group 154 | Oct 2023 |
| 17 | 3 of 5 | CANDIDATE | `Text 16` | Group 154 | Scope & framework defined |
| 18 | 3 of 5 | CANDIDATE | `Text 20` | Group 154 | Method development |
| 19 | 3 of 5 | CANDIDATE | `Text 21` | Group 154 | Pre-trial package complete  formulation pe… |
| 20 | 3 of 5 | CANDIDATE | `Text 24` | Group 154 | Validation phase |
| 21 | 3 of 5 | CANDIDATE | `Text 25` | Group 154 | Stable delivery system established  formul… |
| 22 | 3 of 5 | CANDIDATE | `Text 28` | Group 154 | Deployment |
| 23 | 3 of 5 | CANDIDATE | `Text 29` | Group 154 | First in vivo proof  S. aureus wound infec… |
| 24 | 3 of 5 | CANDIDATE | `Text 33` | Group 154 | Reports / tools delivered |
| 25 | 3 of 5 | CANDIDATE | `Text 36` | Group 154 | Project end |
| 26 | 3 of 5 | CANDIDATE | `Text 37` | Group 154 | Sep 2026 |
| 27 | 3 of 5 | CANDIDATE | `TextBox 118` | Group 154 | Project closed 2026 |
| 28 | 3 of 5 | CANDIDATE | `TextBox 170` |  | Deliverables Include:_x000D_Journal Article – Silver doped magnesium hydroxide particles h… |
| 29 | 3 of 5 | CANDIDATE | `Text 33` |  | Intensive livestock systems require effective alternatives to antibiotics for managing ski… |
| 30 | 3 of 5 | CANDIDATE | `Text 39` | Group 252 | Found lead compounds that were much more effective than standard materials (Q1F25)) |
| 31 | 3 of 5 | CANDIDATE | `Text 39` | Group 266 | Lead candidates were safe for skin cells at working doses (>80% viability) (Q1F26) |
| 32 | 3 of 5 | CANDIDATE | `Text 39` | Group 270 | Silver‑enhanced compounds helped stop biofilm formation (Q1F26). |
| 33 | 3 of 5 | CANDIDATE | `Text 11` | Group 274 | Project initiated_x000D_Oct 2023 |
| 34 | 3 of 5 | CANDIDATE | `Text 19` | Group 274 | Method exploration_x000D_Pre-trial package complete |
| 35 | 3 of 5 | CANDIDATE | `Text 23` | Group 274 | Validation_x000D_Stable delivery system established |
| 36 | 3 of 5 | CANDIDATE | `Text 28` | Group 274 | First in vivo proof S . aureus |
| 37 | 3 of 5 | CANDIDATE | `Text 32` | Group 274 | Project end_x000D_Sep 2026 |
| 38 | 3 of 5 | CANDIDATE | `TextBox 278` |  | Project closed 2026 |
| 39 | 5 of 5 | CHROME | `Text 31` |  | ABOUT |
| 40 | 5 of 5 | CHROME | `Text 107` | Group 48 | CI / PI |
| 41 | 5 of 5 | CHROME | `Text 108` | Group 48 | Lead researcher (name withheld) |
| 42 | 5 of 5 | CHROME | `Text 109` | Group 50 | Industry |
| 43 | 5 of 5 | CHROME | `Text 111` | Group 54 | University |
| 44 | 5 of 5 | CHROME | `Text 11` |  | STRATEGIC ALIGNMENT |
| 45 | 5 of 5 | CHROME | `Text 211a` | Group 211 | Start |
| 46 | 5 of 5 | CHROME | `Text 213a` | Group 211 | End |
| 47 | 5 of 5 | CHROME | `Text 218a` | Group 211 | Time elapsed |
| 48 | 5 of 5 | CHROME | `Text 6` |  | months |
| 49 | 5 of 5 | CHROME | `Text 15` | Group 154 | System design complete |
| 50 | 5 of 5 | CHROME | `Text 32` | Group 154 | Final outputs |
| 51 | 5 of 5 | CHROME | `Text 10` | Group 154 | ▶ |
| 52 | 5 of 5 | CHROME | `Text 14` | Group 154 | 6 |
| 53 | 5 of 5 | CHROME | `Text 18` | Group 154 | 12 |
| 54 | 5 of 5 | CHROME | `Text 23` | Group 154 | 24 |
| 55 | 5 of 5 | CHROME | `Text 27` | Group 154 | 36 |
| 56 | 5 of 5 | CHROME | `Text 31` | Group 154 | 48 |
| 57 | 5 of 5 | CHROME | `Text 35` | Group 154 | ★ |
| 58 | 5 of 5 | CHROME | `Text 35` |  | PROJECT PROGRESS |
| 59 | 5 of 5 | CHROME | `Text 35` |  | PROJECT DELIVERABLES |
| 60 | 5 of 5 | CHROME | `Text 35` |  | KEY PROJECT EVENTS AND STATUS |
| 61 | 5 of 5 | CHROME | `Text 35` |  | PROJECT TEAM |
| 62 | 5 of 5 | CHROME | `Text 31` |  | PROBLEM |
| 63 | 5 of 5 | CHROME | `Rectangle 132` |  | SAAFE Cash $ AUD |
| 64 | 5 of 5 | CHROME | `Rectangle 126` |  | Industry Cash $ AUD |
| 65 | 5 of 5 | CHROME | `Rectangle 135` |  | In-Kind $ AUD |
| 66 | 5 of 5 | CHROME | `Rectangle 136` |  | Total Project Value $ AUD |
| 67 | 5 of 5 | CHROME | `Shape 229` |  | Project Highlights |
| 68 | 5 of 5 | CHROME | `Text 15` | Group 274 | System Design Complete |
| 69 | 5 of 5 | CHROME | `Text 191` | Group 274 | ▶ |
| 70 | 5 of 5 | CHROME | `Text 194` | Group 274 | 6 |
| 71 | 5 of 5 | CHROME | `Text 207` | Group 274 | 36 |
| 72 | 5 of 5 | CHROME | `Text 211` | Group 274 | ★ |
| 73 | 5 of 5 | CHROME | `Text 203` | Group 274 | 9 |
| 74 | 5 of 5 | CHROME | `Text 207` | Group 274 | 24 |
| 75 | 5 of 5 | CHROME | `Text 199` | Group 274 | 12 |
| 76 | 5 of 5 | CHROME | `Text 207` | Group 274 | 48 |
| 77 | 5 of 5 | CHROME | `Text 28` | Group 274 | Final outputs |

**77 untracked items — 38 candidates, 39 probable chrome.**
Total text items on the slide: 82. Managed: 5 (6%).

## C. Caveat on the banding — read this before using it

The band is a heuristic and it is weakened on THIS deck specifically. It compares each
text against sibling slides and assumes those were independently authored. Four of the
six slides here were produced by CLONING (the tool's own slide creation, plus the master
template itself), so their untagged content is identical by construction. Genuinely
project-specific text therefore still scores above zero, and the strongest verdict
("on no other slide") fired **zero times out of 77**.

Practical consequence: **treat 5/5 as chrome and anything below 5/5 as a candidate.**
Do not read the absence of a 0/5 row as meaning there is no project data here — there
plainly is (`~$280K`, `30 Oct 2023`, `90%`, the timeline milestones).

The judgement column is deliberately left to the Excel side. The audit's job is to make
sure nothing is missed, not to decide what is tracked.