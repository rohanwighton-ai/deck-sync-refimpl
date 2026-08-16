# Checklist

> Flat, current, one line per item. Add to it, tick items off in place — do not
> rewrite history here. `NEXT-SESSION.md` is the narrative; `SCENARIOS.md` is the
> nine-scenario frame; `FIX-LIST.md` is defect detail. This is just the list of
> what's actually left, so nothing has to be re-derived from prose to know what to
> do next.

## Immediate — nothing else is real until this happens

- [ ] Build `addin104`: run `build_ppam.ps1`, Save As `addin104`, tick it in
      PowerPoint's Add-ins manager, untick `addin102`, restart PowerPoint. Four real
      fixes (Q, R, the drafting-report labels, the readiness partial-quarter check)
      exist only in source until this happens.

## Scenario 1 — the last piece of "the quarter"

- [ ] Review the 43 rolled-forward drafting rows for `Q1F27` on the rig
      (`AppData\Local\deck-sync-quarter-20260815-1623\`) — content decision, not a
      coding one.
- [ ] Tick `APPROVE` for what's correct, leave the rest blank.
- [ ] Publish, **unaided, no Claude in the loop** — this is the only action that
      actually moves the count from 5/9 to 6/9. The mechanism is already proven;
      only the unaided run is owed.

## Milestone device — Q and R fixed in source, two things still owed

- [ ] Confirm Q + R actually work on slide 1 of the real rig deck once `addin104`
      is loaded (evidence so far is tests, not a live slide — see "What was not
      done" in `NEXT-SESSION.md`).
- [ ] **New finding, 2026-08-16:** `InjectDeviceVia` always reports `WouldChange =
      True`, even when nothing on the slide actually differs from the register.
      Once real milestone data exists, every sync run will re-flag the device as
      "changed" whether or not it is — will pollute the review queue over time.
      Needs a real current-vs-proposed comparison before this is fully correct, not
      just reachable. Not started.

## File-per-quarter — the prune half (critical path #3)

- [ ] Design + build the prune: drop the old period's rows from the live register
      once it's been archived.
- [ ] Retire `ParkSheetCopy` once the prune lands (it stays load-bearing until then
      — do not delete early).
- [ ] Sweep `Sync Log` into the same per-quarter archiving (append-forever sheet,
      same shape as the park-sheet problem — see `SCENARIOS.md`).
- [ ] Tests + one real keyboard run before the prune touches anything live — this
      is the destructive half; the archive half already proved that discipline
      matters here.

## Scenario 3 — per-letter templates (blocked on a real defect, not reachability)

- [x] Step 1 — `TemplateSlide.CodeLetterOf`, done and tested.
- [ ] Step 2 — per-letter registration property (`DeckSyncTemplate:<type>:<letter>`
      alongside the existing `DeckSyncType:`, with fallback for untouched decks).
- [ ] Step 3 — choose the template **per row**, inside `RunSync.CreateMissingSlides`
      (re-run scenario 2 after touching this — same code path).
- [ ] Step 4 — relax the one-per-type guards (`RibbonUI.bas:2375` and
      `MakeTemplateFrom`) to one-per-type-**per-letter**, before step 5 or it's
      refused outright.
- [ ] Step 5 — the deck surgery: make the `K` and `S` templates from real slides,
      on a copy, never the live deck.

## Scenario 8 — portability (never tried)

- [ ] Bring up a genuinely fresh deck + fresh register from nothing, unaided. This
      is the standing requirement — the tool is personal and has to travel with
      Rohan, not just work on this one deck.

## Scenario 9 — provenance (designed, not built)

- [ ] Build the "why does this field say that?" answer per `PROVENANCE.md`. Named
      repeatedly as the thing most likely to get deferred forever.

## Known open defects — not urgent, not forgotten

- [ ] `FIX-LIST` P: cloud persistence on the 4 setup document-properties is
      intermittent (~50% land rate), uncharacterised beyond "eight causes
      eliminated." Workaround stands: do setup writes on a local copy.
- [ ] Review grid may not be safe to rebuild under a live Excel AutoFilter —
      strongly implicated (dirty twice under a filter, clean once without), not
      proven. Proving test: re-apply a filter, rebuild, check for duplicate rows.
- [ ] `Roll Forward` requires clicking a cell to name the source quarter, when the
      tool already knows the deck's period and which periods have rows. Picking
      the wrong quarter looks exactly like success.
- [ ] A bare Excel "permanently delete this sheet" prompt fires during the
      drafting-sheet rebuild — no context, not the tool's own words.
- [ ] The apply-confirmation dialog title is hardcoded wrong when reached from
      `2. Put it on the slides` (says "1. Set up my quarter -- slide changes").

## Parked / explicitly decided — do not reopen without a new decision

- [x] ~~Merge the 13 per-field drafting sheets into one long-format sheet.~~
      **Reversed by Rohan, 2026-08-16**: too much time already lost to redesign
      detours on this project; staying with per-field sheets to keep momentum.
      Never written into any doc before this entry, so nothing else needed fixing.
