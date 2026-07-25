# VBA implementation: `WorkbookBridge.bas`

Small shared primitive both `RibbonUI.bas` (Sync Now, New Period) and
`OnboardFlow.bas` (Onboard New Slide Type, which establishes the deck-workbook
pairing in the first place) need: given a workbook path, get a live `Workbook`
object -- reusing an already-open instance if one matches, otherwise driving
Excel via COM the same way this project's engine already does everywhere else
(per `vba-port.md`: "VBA runs against a live Worksheet... via COM automation
driven from the PowerPoint side"). Not a new sync/matching primitive -- pure
plumbing, split out once two ribbon-layer callers needed the exact same few
lines rather than duplicating them.

**Executed against real Office 2026-07-26.** `SanitizeSheetName` (the one
genuinely new piece of logic here -- Excel sheet-name character/length rules)
has 4 direct tests, all passing. `GetExcelApp`/`OpenOrGetWorkbook`/
`CreateWorkbook`/`GetOrAddWorksheet` are COM plumbing, exercised indirectly via
`OnboardFlow_CommitAndVerifyOnboardingRoundTrip` (which creates a real throwaway
workbook end-to-end) rather than given dedicated tests of their own -- same
posture DeckAdoption/RunSync took for their own Excel-COM-driving code.

## Divergence from any spec

No dedicated spec -- this is plumbing extracted once duplication appeared
across `RibbonUI.bas` and `OnboardFlow.bas`, not a designed-ahead module.

## Manual verification recipe

No dedicated manual recipe beyond what `SPIKE_NOTES_OnboardFlow.md`'s already
covers (Onboard New Slide Type exercises every function here as part of its own
flow) -- this module has no standalone user-facing entry point.
