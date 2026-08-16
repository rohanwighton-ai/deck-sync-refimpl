# VBA port: `excel_output` module

Module 6 of `specs/vba-port.md`'s port order -- the last one. Ports
`create_sheet()`/`read_sheet()`/`upsert_row()` (`ExcelOutput.bas`),
operating on a live `Worksheet` object instead of rebuilding a `.xlsx` zip
from scratch on every write.

**Executed against real Office (2026-07-25)** -- all 8 tests in
`vba/tests/TestRunnerExcel.bas` pass for real, run inside Excel's own VBA
project via `run_vba_tests.ps1`. Same limitation the Python side itself
notes: no `.xlsx` fixture exists in `test-fixtures/` for this spec, so both
languages' own tests are necessarily round-trip/self-consistency checks
(write then read), not checks against an externally-produced file.

## A real cross-app finding (2026-07-25): named Excel constants don't resolve outside Excel's own project

`LastUsedColumn`/`LastUsedRow` originally used the named constants
`xlToLeft`/`xlUp` (from Excel's `XlDirection` enum). These resolve fine
when this module runs *inside Excel's own VBA project* (which is how
`TestRunnerExcel.bas` has always exercised it, and why every prior test
pass here looked completely clean) -- but `vba-port.md`'s stated real
target also includes driving Excel *from PowerPoint* via COM automation
("VBA runs inside Excel or drives it via COM from the PowerPoint side").
When `RunSync.bas` (a PowerPoint-hosted module) actually did that for the
first time, `xlToLeft` raised a hard compile error: "Variable not defined"
-- a PowerPoint-hosted VBA project has no reference to Excel's type
library, so Excel-specific named constants simply don't exist there, even
though the exact same code compiles and runs fine inside Excel itself. This
had been silently true since `ExcelOutput.bas` was first written; nothing
had ever exercised the cross-app path until `RunSync.bas` did.

**Fixed** by replacing both named constants with their numeric literal
values (`XL_TO_LEFT = -4159`, `XL_UP = -4162`, module-level `Private
Const`s) -- stable, documented Office constants unaffected by which host
application's VBA project this runs in. See `SPIKE_NOTES_RunSync.md` for
where this was actually found (a real PowerPoint-driven end-to-end test,
not a code review).

## What was ported

- `ExcelOutput.CreateSheet(ws, deckReference)`: writes the `"Instance ID"`
  header into A1 and stores `deckReference`, refusing to touch an
  already-set-up sheet.
- `ExcelOutput.ReadSheet(ws) As Sheet`: recovers fields (row 1, columns B..),
  instance rows (column A, rows 2..), and the deck reference, into a `Sheet`
  UDT.
- `ExcelOutput.UpsertRow(ws, instanceId, values)`: appends a new field
  column or a new instance row as needed, writes only the given fields'
  cells -- never touches a field/instance this call doesn't mention.

## Deliberate divergences from the Python semantics

1. **Direct incremental `Cells` writes, not read-whole/mutate/rewrite-whole-file.**
   `excel_output.py`'s `_write_xlsx` regenerates the entire `.xlsx` package
   (six parts) on every single call, because a headless zip has no other
   write primitive. A live `Worksheet` does: `UpsertRow` finds-or-appends
   the target column and row directly and writes only those specific cells.
   This is exactly what `specs/vba-port.md` calls out in advance for this
   module -- "VBA is strictly simpler... plain Range/Cells reads and
   writes... don't port the zip-rebuilding approach" -- not a corner cut,
   the intended port shape.
2. **Deck reference lives in a native `CustomDocumentProperties` entry, not
   a hand-rolled `docProps/custom.xml` part.** Same reasoning as
   `identity_tags.md`'s `Shape.Tags`/`Slide.Tags`: Python reverse-engineers
   a custom-properties OOXML part because it has no host application;
   `Workbook.CustomDocumentProperties` is the exact same underlying
   mechanism exposed natively. `WriteDeckReference` checks-then-`Add`s
   rather than blindly `.Add`ing, since `CustomDocumentProperties.Add`
   errors if a property with that name already exists -- the object model's
   own upsert idiom, not read-merge-write.
3. **"Refuse to overwrite" is reinterpreted as "refuse to re-initialize a
   non-empty sheet."** `create_sheet` refuses when the target *file* already
   exists; a live `Worksheet` has no equivalent "exists on disk yet"
   concept -- what actually matters is not silently discarding an
   already-set-up sheet's data, so `CreateSheet` checks `A1` is empty
   instead. A blank new worksheet and a "file that doesn't exist yet" are
   the same real-world situation this guard protects against.
4. **`IsEmpty(cell.Value)`, not `cell.Value = ""`, distinguishes "no value
   harvested" from "harvested value happens to be an empty string."**
   `read_sheet` makes this same distinction structurally (a cell's `<c>`
   XML element either exists or doesn't); `IsEmpty` is the VBA object-model
   equivalent of that presence check, and a plain string-equality test
   would have silently collapsed the two cases.
5. **`Sheet.Fields`/`Sheet.InstanceOrder` are `Collection`s, `Sheet.Rows` is
   a `Scripting.Dictionary` of `Scripting.Dictionary`s** -- not, e.g., a
   single `Dictionary` keyed by field/instance name for everything. Ordered
   lists use `Collection` (matching `SyncOperations.bas`'s existing
   `instanceOrder` convention) rather than leaning on `Scripting.Dictionary`'s
   de-facto (undocumented) key-insertion order. The nested-Dictionary shape
   for `Rows` is legal specifically because `Dictionary`/`Collection` are
   *Objects*, not UDTs -- see `SPIKE_NOTES_Onboarding.md`'s finding on why a
   UDT (`Candidate`, `MatchResult`, `InjectResult`) could never have gone
   into one.
6. **No SHA-256 hashing, no re-open-from-disk verification** for the same
   reason every object-model-based module lacks it: writes and reads both
   go through the live `Worksheet`, not a reopened file.

## The port-order-4 interface contract is satisfied exactly, not adapted

`SPIKE_NOTES_Resolve.md`'s divergence 4 documented `SyncOperations.
PlanRoutineSync`'s `dataRows`/`instanceOrder` parameters as "a documented
interface contract for that future module to satisfy." No adapter was
needed: `ReadSheet(ws).Rows` **is** a `Scripting.Dictionary` of instance ID
-> `Scripting.Dictionary` of field name -> value, and
`ReadSheet(ws).InstanceOrder` **is** a `Collection` of instance-ID strings
in row order -- pass them straight into `PlanRoutineSync` as-is.

## What was deliberately left out of scope

- Formatting/styling of the sheet -- `specs/excel-output.md`'s own Non-goal.
- The archival-copy behavior for period rollover -- a deck-side concern
  (`specs/slide-duplication-trigger.md`), not this module's job; nothing
  here decides when a new row should be created, only how to write one.
- Any orchestration that decides *when* to call `UpsertRow` (onboarding's
  seed-row step, a future "add new quarterly record" flow) -- this module
  builds the primitive, not the driver script that calls it, same boundary
  every other module in this port draws for itself.
- Opening/creating the workbook file itself -- `CreateSheet`/`ReadSheet`/
  `UpsertRow` all take an already-obtained `Worksheet` object, matching
  every other module's "operates on what the caller already has open," not
  a file path.

## Manual verification recipe

Run from the VBA IDE (Alt+F11) with the Immediate window open (Ctrl+G), in
an Excel workbook this time (or driven via COM from PowerPoint against a
live Excel `Application`/`Workbook`/`Worksheet` -- the code doesn't care
which host it's running in, only that it has a real `Worksheet` object).
Import `ExcelOutput.bas` into the workbook's VBA project.

### 1. `ManualSmokeTest` -- the full round trip in one pass

1. Add a new blank worksheet (so A1 is genuinely empty) and note its name.
2. Run `ExcelOutput.ManualSmokeTest ThisWorkbook.Worksheets("<that name>")`.
3. Expected output, cross-checked against
   `tests/test_excel_output.py::test_upsert_row_new_field_appends_a_column_without_touching_existing_data`
   and `::test_upsert_row_new_instance_does_not_disturb_existing_rows`:
   `DeckReference=deck-v1`, `Fields=Title,Date,Region`,
   `InstanceOrder=slide-1,slide-2`, `slide-1: Title=Q3 Revenue Date=2026-07
   Region=APAC`, `slide-2: Title=Q4 Revenue HasDate=False`.
4. Visually confirm in the worksheet itself: A1 = `Instance ID`, B1 =
   `Title`, C1 = `Date`, D1 = `Region` (per
   `tests/test_excel_output.py::test_header_row_reserves_column_a_for_instance_id`),
   and that row 3 (slide-2) has genuinely blank cells in C and D, not empty
   strings.

### 2. `CreateSheet` refuses to re-initialize

1. On the same worksheet used above (now non-empty), run
   `ExcelOutput.CreateSheet ws, "deck-v2"` directly from the Immediate
   window.
2. Expected, per
   `tests/test_excel_output.py::test_create_sheet_refuses_to_overwrite_an_existing_file`:
   raises an error; re-run `ExcelOutput.ReadSheet(ws).DeckReference` and
   confirm it still reads `"deck-v1"`, not `"deck-v2"`.

### 3. Partial update merges, never replaces

1. On a fresh blank sheet, run `CreateSheet` then `UpsertRow` with
   `{"Title": "Q3 Revenue", "Date": "2026-07"}` for `"slide-1"`.
2. Run `UpsertRow` again for `"slide-1"` with only `{"Title": "Q3 Revenue
   (revised)"}`.
3. Expected, per
   `tests/test_excel_output.py::test_upsert_row_partial_update_merges_rather_than_replacing`:
   `ReadSheet(ws).Rows("slide-1")("Date")` still reads `"2026-07"` --
   untouched by the call that didn't mention it.

### 4. Field/instance order is first-seen, not re-sorted

1. On a fresh blank sheet, run `CreateSheet` then, in this exact order:
   `UpsertRow ws, "slide-3", {"Zeta": "z"}`, `UpsertRow ws, "slide-1",
   {"Alpha": "a"}`, `UpsertRow ws, "slide-2", {"Zeta": "z2", "Alpha": "a2"}`.
2. Expected, per
   `tests/test_excel_output.py::test_read_sheet_preserves_field_and_instance_order_across_many_writes`:
   `ReadSheet(ws).Fields` = `Zeta, Alpha` (first-seen, not alphabetical);
   `ReadSheet(ws).InstanceOrder` = `slide-3, slide-1, slide-2`.

### 5. Interface-contract check against `SyncOperations.bas`

1. After step 1's `ManualSmokeTest` setup, call:
   ```
   Dim sheet As Sheet
   sheet = ExcelOutput.ReadSheet(ws)
   Dim instances(1 To 0) As Object ' no onboarded slides in this smoke test -- everything should come back new_record
   Dim actions() As SyncAction
   actions = SyncOperations.PlanRoutineSync(instances, sheet.InstanceOrder, sheet.Rows)
   ```
2. Expected: two `new_record` actions, `RowInstanceKey` = `"slide-1"` and
   `"slide-2"` respectively, `Values` matching each row's harvested fields --
   confirms `ReadSheet`'s output plugs directly into `PlanRoutineSync`
   with no adapter, per the interface-contract note above.
