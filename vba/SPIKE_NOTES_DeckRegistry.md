# VBA implementation: `DeckRegistry.bas`

Implements `specs/deck-registry.md`: the missing lookup a one-click ribbon button
needs and no prior module provided -- "for this open deck, which workbook is
paired with it, and where does each known slide type's template/worksheet live."
Every existing engine entry point (`RunSync.RunRoutineSync`/`RunPeriodRollover`,
`DeckAdoption`'s `templateSld` param) takes these as caller-supplied parameters,
correct for a developer at the VBE but not for a button with no caller to ask.

**Executed against real Office 2026-07-26** (WSL host, `run_vba_tests.ps1`). All
8 tests pass. Full account of that pass's two real bugs found and fixed
(`run_vba_tests.ps1`'s stale import lists, `NewBlankSlide()`'s missing view
navigation) is in `SPIKE_NOTES_ResolveFields.md`/`SPIKE_NOTES_DeckAdoption.md` --
this module's own tests were unaffected by either.

## Design

Storage is `Presentation.CustomDocumentProperties`, the same mechanism
`ExcelOutput.bas` already uses on the `Workbook` side (`WriteDeckReference`/
`ReadDeckReference`) -- confirmed via `grep -rn WriteDeckReference vba/*.bas` that
those are never actually called by anything, so this module is also the first
real caller closing that half of `input-contract.md`'s `deck_workbook_pairing`
rule.

Three property shapes:
- `DeckSyncId` -- a GUID (`Scriptlet.TypeLib`'s `.Guid`, the standard VBA idiom;
  VBA has no native GUID generator), generated once and stable thereafter.
- `DeckSyncWorkbookPath` -- full path to the paired `.xlsx`.
- `DeckSyncType:<slideType>` -- one property per registered type, value
  `<templateSlideID>|<worksheetName>`. Keyed by `Slide.SlideID` (stable across
  reorder/insert/delete of *other* slides), never `SlideIndex`.

`BuildTypeRegistration`/`ParseTypeRegistration` are pure functions (no
`CustomDocumentProperties` access) -- testable without a live `Presentation`,
same posture as `ResolveFields.bas`'s split between interactive entry points and
pure logic.

`LookupType` returns `False` (never raises) both when a type was never
registered and when its stored SlideID no longer resolves (template slide
deleted) -- a ribbon button needs to tell "not onboarded yet" apart from a
genuine error without catching an exception to find out.

## Divergence from the spec

None of substance -- `specs/deck-registry.md` was written this same pass,
directly against what got built, rather than speculatively ahead of it.

## Manual verification recipe

1. In a presentation with `DeckRegistry.bas` imported, run `ManualSmokeTest`
   (F5 or `Application.Run "DeckRegistry.ManualSmokeTest"`).
2. Expect a message box showing a freshly-generated `DeckId` and
   `LookupType found=True templateSld.SlideID=<n> ws=Sheet1` for the
   `manual-smoke-test-type` type registered against slide 1.
3. Re-run `GetOrCreateDeckId` a second time (e.g. via Immediate Window) and
   confirm it returns the identical GUID, not a new one.
