# Ribbon / UI Layer

Every module through `vba-port.md`'s port order plus the orchestration layer
(`RunSync.bas`) is now written and proven against real Office (29/29 tests), but every
entry point is a Sub a developer calls from the VBE — there is no surface a real user
opens PowerPoint and clicks. This spec governs the thin ribbon-and-forms layer that
exposes the existing engine to an actual user, full workflow (routine sync, explicit
period rollover, onboarding a new slide type, and resolving an ambiguous match) — not a
core-only slice. It builds no new sync/matching/verification decision logic; every
button is a call into an existing, already-tested Sub.

**Host: PowerPoint only.** `RunSync.bas` is already PowerPoint-hosted and reaches Excel
cross-app via COM (confirmed in code) — the deck is the primary artifact a user has
open, so the ribbon lives there. No duplicate entry points on the Excel side.

## Requirements

- **Ribbon tab, two always-visible buttons:**
  - **"Sync Now"** → `RunSync.RunRoutineSync` (cases 1/3/4/6) against the currently open
    deck and its paired workbook (`input-contract.md`'s deck-workbook pairing).
  - **"New Period"** → case 2, explicit rollover. Not global — per
    `run-sync.md`'s Step 3, it's per type+record — so this opens a small picker (type,
    then the specific record) rather than acting on the whole deck.
- **Onboard New Slide Type** button, following `onboard-slide-type.md`'s six steps
  exactly:
  1. Takes the user's current slide selection (`Application.ActiveWindow.Selection.
     SlideRange`, 1-2 slides) as the example. Immediately duplicates it into a working
     copy per the workflow's own invariant — the original the user had selected is never
     opened for write again; every later step in this flow operates on the duplicate.
  2. Runs `Discovery` against the working copy (existing call, no new logic).
  3. **Review form — the phase gate.** Lists discovered fields (proposed name, shape
     reference, harvested value), editable: rename, exclude, and mark one field as the
     period-key or the type as evergreen. Nothing is written until this form is
     confirmed.
  4. On confirm: commits tags/template/Data-sheet columns/seed row via the existing
     `Onboarding.bas`/`ExcelOutput.bas` calls, then runs the existing verify-the-link
     pass (`onboard-slide-type.md` Step 6) and reports pass/fail per field before
     declaring the type onboarded.
- **Resolve Unmatched Fields** button/flow, for medium-confidence matches found when a
  *subsequent* slide is checked against an already-established template
  (`onboarding.md`'s matching case — distinct from first-time onboarding above):
  - User clicks the ambiguous shape directly in the open deck
    (`Application.ActiveWindow.Selection.ShapeRange` — the exact mechanism
    `onboarding.md`'s own Non-goals names as this UI's job to build), then picks the
    field role from a list of the template's defined roles.
  - Calls the existing `confirm_field_match` primitive; no new matching logic.
- **One shared result form**, reused after Sync Now, New Period, and the onboarding
  verify step — not a bespoke dialog per action: counts (no-op / created / corrected /
  flagged unclassified / flagged conflict), with flagged items listed by slide
  name/index so the user can actually locate them in the deck.
- **Ribbon defined via `customUI14.xml`** injected into the `.ppam` package itself — the
  ribbon is add-in-native from the start, not something designed against a raw `.pptm`
  and ported later. Ties directly to the COM-add-in packaging decision (DECISIONS.md,
  2026-07-25).

## Non-goals

- **Settings/preferences UI** (paths, matching thresholds, managing deck-workbook
  pairings beyond what onboarding itself establishes) — deferred.
- **Case 5 (`record_retired`) and case 7 (`deck_side_conflict`) resolution UI** — both
  stay manual-Excel-edit workflows per `run-sync.md`'s own Step 4 ("a human resolves it
  by hand and resets `Status` back to `Active` themselves"); no dialog builds a
  resolution flow for either.
- **Multi-deck dashboard / cross-deck management** — parked per the multi-deck note in
  `IMPLEMENTATION_PLAN.md` Priority 19; this spec is single-deck throughout.
- **Reconciling a template that has itself drifted** — not a UI concern; out of scope
  per `onboarding.md`'s own Non-goals.
- **Excel-side ribbon or entry points** — see Host note above.

## Reference

- Engine being surfaced, unchanged by this spec: `vba/RunSync.bas`, `vba/Onboarding.bas`,
  `vba/SlideDuplication.bas`, `vba/ExcelOutput.bas`.
- Workflows this UI must match, not reinvent:
  `~/.claude/skills/crc-vba-deck-sync/workflows/run-sync.md`,
  `~/.claude/skills/crc-vba-deck-sync/workflows/onboard-slide-type.md`.
- Confirmation primitive and its named selection mechanism: `specs/onboarding.md`.
- Packaging context: `DECISIONS.md`, 2026-07-25 ("COM add-in first, Office.js kept
  open").
