# Provenance at register grain

**Status:** designed, not built. Raised by Rohan 2026-08-10:

> "I'm a bit worried with the write-over element I will not be able to see the
> drafting building blocks used to synthesise prior results that end up in the
> slide? Surely I should be able to audit a field's progeny back through the
> provision of source data into its particular recipe?"

---

## The gap

The register accumulates: one row per slide per period, and rolling forward
copies a period's rows into the next. So **what a field said** survives forever,
per quarter, and can be read back years later.

Everything that explains *how it came to say that* lives on the drafting sheet,
which is rebuilt every period.

| | survives a rollover |
|---|---|
| the published text | **yes** — register row, per period |
| which sources it was drafted from | **no** — the drafting sheet's SOURCES column, cleared |
| the AI draft before it was edited | **no** — drafting sheet, cleared |
| the recipe as it stood at the time | **no** — Field Spec is permanent but edited **in place** |

This was already recorded in the project notes as *"provenance dies at
rollover ... fine while Rohan is the only user; not fine the moment someone else
asks."* His framing on 2026-08-10 supersedes that: it is not fine for **him**,
auditing his own work. That is a higher standard and the right one.

## What the sheet-parking change is NOT

`Drafting.ParkSheetCopy` (added the same evening) copies a drafting sheet aside
before a rebuild that would lose unpublished work. That is a **safety net against
an accidental rebuild**, not an audit trail:

- hidden, so it is not stumbled upon
- **capped at two per field**, so it is not history
- unstructured — a whole sheet, not a queryable record

Using it as the provenance record would repeat the mistake this design exists to
avoid: the right fact in the wrong place, invisible, and pruned.

---

## The design

**One rule: provenance lives at the same grain as the value it explains.** The
value is a cell in a register row keyed by (instance, period). So is its
provenance. It then rolls forward, accumulates and survives identically, with no
second mechanism to keep in step.

### 1. `<FIELD>_SOURCES` — which documents it came from

A column beside each prose field's value column, written **by publish, in the
same operation** that writes the value.

Written together or not at all. A citation that can be written when the value
was not (or vice versa) is a provenance record that can lie, and this project
already has a rule for that shape: the pair is one write.

Content is the cited source IDs exactly as the drafting sheet held them —
`S10, S12`. Resolution against the Sources sheet stays a *read-time* concern, so
a source row that is later corrected does not rewrite history.

**Cost:** one column per prose field (5 today, 8 with the new ones). The register
already carries ~11 columns; this roughly doubles it. Acceptable — it is the
same shape as the milestone columns and, per Rohan 2026-08-10, "one column per
interval" is the readable form.

### 2. `<FIELD>_RECIPE` — which instructions produced it

A short hash of the Field Spec row (purpose, voice, length, own-job test,
do-nots, allowed values) as it stood **at publish time**, written the same way.

This answers a question that cannot be asked at all today: *was this written
under the current recipe, or the one from two quarters ago?* The Field Spec is
permanent but edited in place, so a recipe improved in Q1 leaves every earlier
field looking as though it was written under the new rules.

A hash rather than the text, because the text is a paragraph per field per
period and would dominate the register. The hash is a discriminator, not a
record: it tells you the recipe CHANGED, and the Field Spec's own history (git,
or a dated copy) tells you how.

**Open question for Rohan:** a hash says "different" but not "how different". If
that is not enough, the alternative is a dated snapshot sheet of the whole Field
Spec written on each publish where the hash changes — bounded, since it only
writes when something actually changed.

### 3. The AI draft — deliberately NOT stored

The pre-edit draft is bulky (a paragraph per field per project per period) and
its value decays fast: what matters six months later is what was published, from
what sources, under what instructions. Left out unless Rohan asks.

The parked-sheet archive covers the short-term case — "what did the AI actually
suggest last week" — which is the window where it is genuinely wanted.

---

## What has to change

1. **`Drafting.PublishDrafts`** — write `<FIELD>_SOURCES` and `<FIELD>_RECIPE`
   alongside the value in the same `UpsertRow` call. One write, not three.
2. **`FieldSpec`** — a `RecipeHash(specWs, fieldId)` function over the row's
   guidance cells. Must be stable across whitespace and case so a cosmetic edit
   does not read as a changed recipe.
3. **`ExcelOutput.UpsertRow`** — no change expected: provenance columns are
   ordinary fields as far as it is concerned.
4. **`FieldWiring`** — provenance columns must NOT be reported as fields with no
   shape. They are register-only: they explain a value, they never reach a
   slide. This is the one genuine exemption in the model, and it is narrow
   enough to be derived from the suffix rather than configured.
5. **Drafting sheet** — the SOURCES column already exists and is already read at
   publish. No change to what a person types.

## What it must refuse

- **A value published without its provenance.** If the sources column cannot be
  written, the value is not written either. A half-written pair is worse than a
  refusal, because it looks complete.
- **A recipe hash that cannot be computed** — reported, and the publish refused,
  for the same reason.

## How it gets proven

The check that matters is not that provenance is written, but that it **survives
a rollover**. The test: publish a field with citations in period A, roll forward
to period B, and assert period A's row still carries its sources and recipe hash
unchanged — and that B's row carries A's values as its starting point without
inheriting A's provenance as though it were its own.

That last clause is the subtle one. Rolling forward copies values so the next
quarter starts from last quarter's text. It must **not** copy the provenance,
or every quarter would claim to have been drafted from the sources of the first
one. Provenance is written by publish, and only by publish.
