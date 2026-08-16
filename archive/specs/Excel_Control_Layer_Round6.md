# Excel Control Layer — round 6

**From:** the PowerPoint / VBA side, 31 July 2026
**Status:** E2 received and V1 is built. One blocker found while building it, plus four
format questions that V4 cannot be written without. No disagreements — round 5 is accepted
in full, including both corrections to my own positions.

---

## 1. Blocker — `EntityCode` has the same problem `FieldID` had, and nobody has done E2 for it

E2 mapped the `role` tag values. **Nothing has mapped the `instance_key` values**, and they
do not match the register's `EntityCode` as specified.

What the real deck actually holds today, read off the slides:

```
3_P001    3_P002    3_P002-2    3_P004    3_P005
```

What the register expects, per §2.1 of the original spec and the examples throughout:

```
P004
```

**These will not join.** The register keys on `Quarter × EntityCode × FieldID`; the deck
carries `3_P001` where the register says `P004`. Every row misses. This is exactly the
failure E2 was written to prevent, one column over — and it went unnoticed because E2 solved
the FieldID half so cleanly that the EntityCode half looked handled.

Three things to decide, and they are yours because they are portfolio conventions, not
implementation:

1. **Is the canonical EntityCode `P001` or `3_P001`?** The `3_` prefix appears to be a
   programme or theme number. If it is meaningful, the register should carry it and the
   examples in the spec are wrong. If it is not, the deck needs migrating.
2. **What is `3_P002-2`?** It is one project occupying two slides, keyed by hand with a
   `-2` suffix. Under `Quarter × EntityCode × FieldID` those two slides now collide on the
   same key — the suffix was doing work the new key structure has to do a different way.
   This is the case D2 of round 3 was written to retire, and it is now live and blocking.
3. **Do S-codes and K-codes appear in the same register as P-codes**, or is EntityCode
   namespaced by type?

**What I need back: an E2-equivalent mapping table for EntityCode** — current
`instance_key` → target `EntityCode`, five rows, same shape as E2. Plus a ruling on
`3_P002-2`, which cannot be a straight rename.

**Cost of doing this now versus later is identical to the FieldID argument** and it is the
same 30-writes-versus-258 curve. The migration tooling built for V1 handles tag values
generically, so an EntityCode map costs nothing beyond authoring it.

---

## 2. Four format questions V4 cannot be written without

All four are string comparisons that must match exactly across two systems. None is a design
question; each just needs the literal.

| # | Question | Why it matters |
|---|---|---|
| F1 | **Exact `Quarter` literal** — `Q4 FY26`? `2026-Q4`? `FY26Q4`? | The deck's period property must string-match the register's `Quarter` column. A mismatch silently returns zero rows, which reads as "nothing to sync" rather than as an error. |
| F2 | **Exact `Status` values and case** — `Approved`, `approved`, `APPROVED`? | R2 filters on it. I will compare case-insensitively unless told otherwise, but the canonical form should be recorded. |
| F3 | **Exact `FieldType` values** — `Text`, `Picture`, `Shape`, `Table`, `Chart`? | Selects the injection routine. Only `Text` is in scope now, but the literal is needed. |
| F4 | **Exact entity-static sentinel** — `ALL`, `All`, `*`? | Round 5 §3 said `ALL`. Confirming because it is a magic value in a data column and those get typed by hand. |

F1 is the one with a real trap in it: **an unmatched period is indistinguishable from an
empty quarter.** Both produce zero rows. Whatever the format, it should be one that a human
typing into the deck's property and a human typing into the register are unlikely to disagree
about — which argues against anything with a space or an optional prefix.

---

## 3. Yours from round 5's own findings — the publish step must normalise line breaks

Round 5 spotted `_x000D_` in Appendix A items 3, 28, 33, 34, 35 and 37: existing slide content
already contains embedded carriage returns. Good catch, and it was in data I produced and did
not read closely enough.

The consequence is on the Excel side, not mine. **R6 says no multi-line cells will appear
anywhere in Excel and `||` is the delimiter — so the publish step has to convert real line
breaks to `||` as it writes.** If existing prose is pasted into the register with its CRs
intact, R6 is violated on day one and V5 has nothing to convert.

My side handles `||` → carriage return on injection. Nothing on my side can enforce the
inbound direction.

---

## 4. Accepted from round 5, no response needed

- **E2** received. V1 built against it: walks every slide including the master template
  (excluding it would leave the slide every new record is cloned from holding the old names),
  dry-run by default, idempotent, and rollback is the same operation with the map reversed.
- **Amendment B's three-class taxonomy.** Mine was wrong and would have caused a real bug —
  under my version `PROJECT_NAME` goes untagged, so every created slide carries the master
  template's project name until someone retypes it. `Quarter = ALL` costs V4 one clause.
- **`SlideID` withdrawn**, revised key `Quarter × EntityCode × FieldID`. Better than my
  version and it removes a redundant column.
- **Amendment A** with period and timestamp stored alongside the placeholder count — correct,
  a bare count cannot be told apart from a stale one.
- **Sequencing confirmed**, and your ordering query was right: my §6 contradicted itself. Moot
  in fact — V2 shipped before V1 was built.
- **V5 promoted** out of fill-in status. It must land before `ABOUT_BODY` or
  `KEY_EVENTS_BODY` go through; `PROJECT_STATUS` is single-line so the first field is
  unaffected.

---

## 5. What I am doing while you answer

V3, the long-format `ExcelOutput` rewrite. It is unblocked, it is the largest remaining
piece, and round 5 settled its target schema — so it is safe to start without any of the
above.

V1 does not run against the real deck until §1 is answered. Renaming FieldIDs while
EntityCodes are still wrong would produce a deck that is half-migrated on two axes at once,
and I would rather do one migration than two.
