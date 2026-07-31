# Deck Sync — tracker

**Finished =** Rohan produces a real quarter's deck and the tool saved him time.
Not a field count. Not a test count.

**7 of 10.**

---

- [x] **1. A field syncs end to end** — register → gate → slide → verified by re-reading.
      *Five fields proven: PROJECT_STATUS, ABOUT_BODY, KEY_EVENTS_BODY, PROJECT_NAME, PROJECT_CODE.*

- [x] **2. Nothing reaches a slide unseen** — every change shown as before-and-after, nothing
      written without approval. *Caught 22 slides of prose about to be corrupted.*

- [x] **3. A drafting sheet exists, and says what to do on it** — instructions on the sheet,
      exemplar beside the input, one tick column.

- [x] **4. Drafted text reaches a slide** — written in the sheet, ticked, published, applied,
      verified. *The loop closes.*

- [x] **5. Several rows apply at once** — 3 in one pass, each to its own slide, confirmed by an
      independent harvest.

- [x] **6. A quarter rolls forward** — deck declares its own period; at FY26Q4 it cannot see
      FY27Q1 rows at all.

- [x] **7. Deck settings survive being written** — 5 consecutive differing updates confirmed
      on disk. *The write itself is still unreliable (SaveAs 4/5, Save far worse); the
      operation is made reliable by `set_deck_period.py` — write, verify offline, retry,
      fail loudly. Never verify in-process: it shares PowerPoint's cache with the writer.*

- [ ] **8. Rohan says the sheet is usable** — or says exactly why it isn't.
      *Done when: the four checkboxes in `MANUAL-TEST-DRAFTING-LOOP.md` Step 1 are answered.*
      *Blocks everything below. Five minutes. Nobody else can do it.*

- [ ] **9. Rohan drafts a real quarter's content** — text he actually needed written, not test
      edits. *Done when: 10 projects drafted, ticked, applied on the e2e copy.*

- [ ] **10. One real quarter produced, and it saved time** — on a copy of the live deck, which
      then becomes the deck. *Done when: Rohan says the sentence. "No" is a valid answer and a
      spec for what to fix.*

---

## The rule for this file

**Only tick something when it is observably true**, not when the code for it exists. Six of
these were ticked by watching a slide change, not by a test passing.

If an item can't be checked by looking at something, it's written wrong — rewrite it.

## Unverified right now

`FieldSpec` (per-field drafting guidance) is proven by 135 unit tests — the prompt it
builds, the fallback when a field has no row, and that two fields get materially different
instructions. What has NOT been watched happening is the wiring: that `WriteDraftingSheet`
receives the spec sheet and writes the new prompt into the cell.

PowerPoint stopped responding to COM before that could be demonstrated (no stale process, no
dialog — the environment, after several hundred Office launches). **First thing next session:
rebuild a drafting sheet and read the prompt cell.** Until then treat the wiring as untested.

## Not on this list, deliberately

The other 38 fields. The GUID key redesign. R13's full review subsystem (built, parked).
Chrome enforcement. Ribbon polish. None of them stand between here and finished; adding them
here would be the "field count as progress" trap this project already fell into once.
