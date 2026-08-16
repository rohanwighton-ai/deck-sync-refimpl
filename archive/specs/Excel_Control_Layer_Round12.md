# Excel Control Layer — round 12: the number is 19

**From:** the PowerPoint / VBA side, 31 July 2026
**One field is through, end to end, on the real 46-slide deck.**

---

## The result

```
--- V1 rename ---
  scanned:  230      renamed: 230      already: 0      unmapped: 0

--- V3 register read ---
  rows seen: 46      accepted: 46      (period-matched 46, entity-static 0)
  rejected status: 0     rejected period: 0     missing columns: none

--- sync, fed from the register ---
Summary: 27 unchanged, 19 corrected, 0 with no slide, 0 failed, 0 flagged

--- VERIFIED BY RE-READING THE DECK ---
  slides matching the register: 46
  mismatched:                    0
  register rows with no slide:   0
```

**19 corrected. The predicted number, computed independently on both sides before the run.**

The verification is a second pass that re-reads all 46 slides and compares them against the
register. It is deliberately not the engine's own report — that says what the engine believes
it did, and only the re-read is evidence.

Corrections were exactly the drift your §2 ruling identified: ten `In progress` → `In
Progress`, eight `Not started` → `Not Started`, one `Not yet commenced` → `Not Started`.

## What this proves, and what it does not

**Proved, all previously built and never connected to each other:**

| | |
|---|---|
| FieldID rename | 230 tags across 46 slides, including the master template |
| Long-format register read | period + status filtering, columns located by name |
| Register → sync engine | the same `Sheet` structure feeding the existing engine unchanged |
| Creation removed from sync | `0 with no slide`, nothing created as a side effect |

**Not proved, and worth being explicit since you named it first:** `PROJECT_STATUS` is a
controlled field, so this run exercised the mechanical half of the pipeline and none of the
content half. No drafting surface, no exemplar column, no Copilot loop, no target length. Your
round-11-decision-note §5 called that before the run and it is exactly right.

**`ABOUT_BODY` second, accepted.** Your argument is our own measurement turned around — 3 of
46 multi-line against 46 of 46 — so it proves the content half against mostly single-line
prose and leaves the hardest field to run third through a path proven twice.

---

## Two things found on the way

**1. `Not yet commenced` was the only one of its kind.** One slide, `1_K010`. It is the case
that argues hardest for the `AllowedValues` list: a single unconstrained cell drifted to a
phrase nobody else used, and nothing would have caught it. Under the controlled class it
cannot recur, which makes these 19 corrections a one-off historical cleanup rather than a
recurring transformation. **That is the better outcome — a constraint that prevents bad values
beats a function that repairs them.**

**2. The `-2` slides synced identically to their base.** `2_P004-2` and `1_P006-2` both
appear in the corrected list alongside `2_P004` and `1_P006`, receiving the same value from
what is effectively the same project. Harmless here because the value is identical — but it
is the phantom-entity problem visible in real output rather than argued about. Three register
rows currently exist for two projects. Still pending RM4, and now with a worked example.

---

## State

**Blocked on: nothing from your side.**

`F1`–`F4` are in the code as constants (`FY26Q4` used for real above). `E2` is applied. The
identity map is applied. The three-class taxonomy is implemented, with `Quarter = ALL`
carrying entity-static rows into any period.

**Next on this side:** the full test suite over today's changes, then `ABOUT_BODY` — which
needs V5 (`||` conversion) exercised for the first time, since 3 of 46 of its values are
multi-line.

**Still open and still the Research Manager's:** RM3 (one workbook, many decks — both sides
recommend it) and RM4 (the three `-2` slides).
