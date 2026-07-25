# Deck Registry

Closes a gap `ribbon-ui.md` assumes is already solved and isn't: every existing
engine entry point (`RunSync.RunRoutineSync`, `RunPeriodRollover`, `DeckAdoption`'s
`templateSld` param, etc.) takes `ws` (the paired worksheet) and `templateSld` as
caller-supplied parameters. That's correct for a developer driving the VBE directly,
but a one-click ribbon button has no caller to supply them — it has only "the deck
that's currently open." This module is what a button asks instead: *for this deck,
what types exist, and where does each type's template/worksheet/workbook live.*

Confirmed via `grep -rn WriteDeckReference vba/*.bas` that `ExcelOutput.bas`'s
`WriteDeckReference`/`ReadDeckReference` (the only prior art for this) are never
actually called by anything — they cover the workbook→deck direction only
(`input-contract.md`'s `deck_workbook_pairing` rule) and nothing establishes the
reverse: given an open deck, which workbook is paired with it.

## Storage

`Presentation.CustomDocumentProperties`, same mechanism `ExcelOutput.bas` already
uses on the `Workbook` side — no new persistence layer, no external file.

- `DeckSyncId` (String): a GUID generated once, on first registration. Written into
  the paired workbook's `DeckReference` property (via `ExcelOutput.CreateSheet`) so
  the pairing is mutually verifiable, not just "any workbook with matching headers"
  — directly closes `input-contract.md`'s cross-wiring risk.
- `DeckSyncWorkbookPath` (String): full path to the paired `.xlsx`.
- `DeckSyncType:<slideType>` (String, one property per registered type):
  `<templateSlideID>|<worksheetName>`. Keyed by Slide**ID** (stable across reorder/
  insert/delete of *other* slides), never by index or position.

## Public surface

- `GetOrCreateDeckId(pres) As String` — read-or-generate-and-write.
- `GetWorkbookPath(pres) As String` / `SetWorkbookPath(pres, path As String)`.
- `RegisterType(pres, slideType As String, templateSld As Object, worksheetName As String)`
  — writes one `DeckSyncType:<slideType>` property.
- `LookupType(pres, slideType As String, ByRef templateSld As Object, ByRef worksheetName As String) As Boolean`
  — resolves the stored SlideID back to a live `Slide` via `Slides.FindBySlideID`;
  `False` (with both ByRef args left `Nothing`/`""`) if the type isn't registered, or
  if its stored SlideID no longer resolves (template slide was deleted) — never
  raises, since "not registered yet" is an expected, routine state for a ribbon
  button to handle (e.g. show "Onboard New Slide Type" instead of failing).
- `ListRegisteredTypes(pres) As String()` — every registered type name, for the New
  Period picker's type dropdown and any other "what types does this deck know about"
  need.

Parsing/building the `<templateSlideID>|<worksheetName>` value is split into pure
functions (`BuildTypeRegistration`/`ParseTypeRegistration`) taking/returning plain
values, no `CustomDocumentProperties` access — testable without a live `Presentation`,
same posture as `ResolveFields.bas`'s split between interactive entry points and
pure logic helpers.

## Non-goals

- No UI here — this is a lookup/storage primitive the ribbon layer calls into, same
  relationship `ExcelOutput.bas` has to `RunSync.bas`.
- No migration/versioning for the property schema — first version, nothing to migrate
  from.
- Doesn't decide *when* `SetWorkbookPath`/`RegisterType` get called — that's each
  ribbon flow's job (Onboard New Slide Type registers on commit; first-ever onboarding
  on a deck also sets the workbook path and deck ID).
