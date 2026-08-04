Attribute VB_Name = "TestRunnerExcel"
Option Explicit

' First real-execution test harness for ExcelOutput.bas, mirroring
' TestRunner.bas's shape (pure assertion functions, no MsgBox, one
' machine-readable report). Runs against a live Worksheet in a fresh
' Excel workbook the driver script creates -- no .xlsx fixture, same
' round-trip/self-consistency approach as tests/test_excel_output.py.

Public Function RunAllTests() As String
    Dim report As String
    report = "=== deck-sync-refimpl VBA test run (Excel) ===" & vbCrLf

    Dim r As String

    r = "": On Error Resume Next: Err.Clear
    r = Test_CreateSheet_SeedsDeckReferenceNoFieldsOrRows()
    AppendResult report, "CreateSheet_SeedsDeckReferenceNoFieldsOrRows", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    r = Test_CreateSheet_RefusesToReinitialize()
    AppendResult report, "CreateSheet_RefusesToReinitialize", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    r = Test_UpsertRow_SeedsNewInstanceFromHarvestedValues()
    AppendResult report, "UpsertRow_SeedsNewInstanceFromHarvestedValues", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    r = Test_UpsertRow_NewFieldAppendsColumnWithoutTouchingExisting()
    AppendResult report, "UpsertRow_NewFieldAppendsColumnWithoutTouchingExisting", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    r = Test_UpsertRow_PartialUpdateMergesNotReplaces()
    AppendResult report, "UpsertRow_PartialUpdateMergesNotReplaces", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    r = Test_UpsertRow_NewInstanceDoesNotDisturbExistingRows()
    AppendResult report, "UpsertRow_NewInstanceDoesNotDisturbExistingRows", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    r = Test_ReadSheet_PreservesFieldAndInstanceOrderAcrossManyWrites()
    AppendResult report, "ReadSheet_PreservesFieldAndInstanceOrderAcrossManyWrites", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    r = Test_HeaderRow_ReservesColumnsAAndBForIdentityAndPeriod()
    AppendResult report, "HeaderRow_ReservesColumnsAAndBForIdentityAndPeriod", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    r = Test_UpsertRow_NextPeriodAppendsAndLeavesLastQuarterIntact()
    AppendResult report, "UpsertRow_NextPeriodAppendsAndLeavesLastQuarterIntact", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    r = Test_UpsertRow_SamePeriodUpdatesThatRowInPlace()
    AppendResult report, "UpsertRow_SamePeriodUpdatesThatRowInPlace", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    r = Test_UpsertRow_RefusesABlankPeriodOnAPeriodSheet()
    AppendResult report, "UpsertRow_RefusesABlankPeriodOnAPeriodSheet", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    r = Test_UpsertRow_RefusesAFieldNamedLikeAStructuralColumn()
    AppendResult report, "UpsertRow_RefusesAFieldNamedLikeAStructuralColumn", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    r = Test_UpsertRow_LegacySheetWithNoPeriodColumnStillMatchesOnInstance()
    AppendResult report, "UpsertRow_LegacySheetWithNoPeriodColumnStillMatchesOnInstance", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    r = Test_ReadForDeckPeriod_KeepsOnlyThatPeriodsRows()
    AppendResult report, "ReadForDeckPeriod_KeepsOnlyThatPeriodsRows", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    r = Test_ReadForDeckPeriod_RefusesTwoRowsForOneSlideInOnePeriod()
    AppendResult report, "ReadForDeckPeriod_RefusesTwoRowsForOneSlideInOnePeriod", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    r = Test_ReadForDeckPeriod_RefusesAPeriodTheSheetDoesNotHave()
    AppendResult report, "ReadForDeckPeriod_RefusesAPeriodTheSheetDoesNotHave", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    r = Test_ReadForDeckPeriod_SilentOnASheetWithNoQuarterColumn()
    AppendResult report, "ReadForDeckPeriod_SilentOnASheetWithNoQuarterColumn", r
    On Error GoTo 0

    r = "": On Error Resume Next: Err.Clear
    r = Test_ReadForDeckPeriod_RefusesANeverInitializedSheet()
    AppendResult report, "ReadForDeckPeriod_RefusesANeverInitializedSheet", r
    On Error GoTo 0

    RunAllTests = report
End Function

Private Sub AppendResult(ByRef report As String, testName As String, testResult As String)
    If Err.Number <> 0 Then
        report = report & "ERROR " & testName & " :: " & Err.Description & " (line context lost -- VBA has no stack trace)" & vbCrLf
    ElseIf testResult = "" Then
        report = report & "PASS  " & testName & vbCrLf
    Else
        report = report & "FAIL  " & testName & vbCrLf & testResult
    End If
End Sub

Private Function Assert(cond As Boolean, msg As String) As String
    If cond Then
        Assert = ""
    Else
        Assert = "    FAIL: " & msg & vbCrLf
    End If
End Function

' Adds a fresh, genuinely blank worksheet to the active workbook -- each
' test gets its own, so no test's data can bleed into another's.
Private Function NewBlankSheet() As Object
    Dim ws As Object
    Set ws = Application.ActiveWorkbook.Worksheets.Add
    Set NewBlankSheet = ws
End Function

Private Function DictOf1(k As String, v As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d(k) = v
    Set DictOf1 = d
End Function

Private Function DictOf2(k1 As String, v1 As String, k2 As String, v2 As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d(k1) = v1
    d(k2) = v2
    Set DictOf2 = d
End Function

Private Function Test_CreateSheet_SeedsDeckReferenceNoFieldsOrRows() As String
    Dim result As String
    Dim ws As Object
    Set ws = NewBlankSheet()

    ExcelOutput.CreateSheet ws, "deck-v1"
    Dim sheet As Sheet
    sheet = ExcelOutput.ReadSheet(ws)

    result = result & Assert(sheet.DeckReference = "deck-v1", "DeckReference round-trips, got '" & sheet.DeckReference & "'")
    result = result & Assert(sheet.Fields.count = 0, "no fields yet, got " & sheet.Fields.count)
    result = result & Assert(sheet.InstanceOrder.count = 0, "no rows yet, got " & sheet.InstanceOrder.count)

    Test_CreateSheet_SeedsDeckReferenceNoFieldsOrRows = result
End Function

Private Function Test_CreateSheet_RefusesToReinitialize() As String
    Dim result As String
    Dim ws As Object
    Set ws = NewBlankSheet()

    ExcelOutput.CreateSheet ws, "deck-v1"

    Dim raised As Boolean
    On Error Resume Next
    Err.Clear
    ExcelOutput.CreateSheet ws, "deck-v2"
    raised = (Err.Number <> 0)
    Err.Clear
    On Error GoTo 0

    result = result & Assert(raised, "CreateSheet on an already-set-up sheet raises rather than silently overwriting")

    Dim sheet As Sheet
    sheet = ExcelOutput.ReadSheet(ws)
    result = result & Assert(sheet.DeckReference = "deck-v1", "original deck reference survives the refused re-init, got '" & sheet.DeckReference & "'")

    Test_CreateSheet_RefusesToReinitialize = result
End Function

Private Function Test_UpsertRow_SeedsNewInstanceFromHarvestedValues() As String
    Dim result As String
    Dim ws As Object
    Set ws = NewBlankSheet()
    ExcelOutput.CreateSheet ws, "deck-v1"

    ExcelOutput.UpsertRow ws, "slide-1", DictOf2("Title", "Q3 Revenue", "Date", "2026-07"), "FY26Q4"

    Dim sheet As Sheet
    sheet = ExcelOutput.ReadSheet(ws)
    result = result & Assert(sheet.Fields.count = 2, "2 fields, got " & sheet.Fields.count)
    result = result & Assert(sheet.InstanceOrder.count = 1 And sheet.InstanceOrder(1) = "slide-1", "one instance, 'slide-1'")
    result = result & Assert(sheet.Rows("slide-1")("Title") = "Q3 Revenue", "Title value correct")
    result = result & Assert(sheet.Rows("slide-1")("Date") = "2026-07", "Date value correct")

    Test_UpsertRow_SeedsNewInstanceFromHarvestedValues = result
End Function

Private Function Test_UpsertRow_NewFieldAppendsColumnWithoutTouchingExisting() As String
    Dim result As String
    Dim ws As Object
    Set ws = NewBlankSheet()
    ExcelOutput.CreateSheet ws, "deck-v1"
    ExcelOutput.UpsertRow ws, "slide-1", DictOf2("Title", "Q3 Revenue", "Date", "2026-07"), "FY26Q4"

    ExcelOutput.UpsertRow ws, "slide-1", DictOf1("Region", "APAC"), "FY26Q4"

    Dim sheet As Sheet
    sheet = ExcelOutput.ReadSheet(ws)
    result = result & Assert(sheet.Fields.count = 3, "3 fields after append, got " & sheet.Fields.count)
    result = result & Assert(sheet.Fields(1) = "Title" And sheet.Fields(2) = "Date" And sheet.Fields(3) = "Region", "append order preserved, not reordered")
    result = result & Assert(sheet.Rows("slide-1")("Title") = "Q3 Revenue" And sheet.Rows("slide-1")("Date") = "2026-07", "existing data untouched")
    result = result & Assert(sheet.Rows("slide-1")("Region") = "APAC", "new field written")

    Test_UpsertRow_NewFieldAppendsColumnWithoutTouchingExisting = result
End Function

Private Function Test_UpsertRow_PartialUpdateMergesNotReplaces() As String
    Dim result As String
    Dim ws As Object
    Set ws = NewBlankSheet()
    ExcelOutput.CreateSheet ws, "deck-v1"
    ExcelOutput.UpsertRow ws, "slide-1", DictOf2("Title", "Q3 Revenue", "Date", "2026-07"), "FY26Q4"

    ExcelOutput.UpsertRow ws, "slide-1", DictOf1("Title", "Q3 Revenue (revised)"), "FY26Q4"

    Dim sheet As Sheet
    sheet = ExcelOutput.ReadSheet(ws)
    result = result & Assert(sheet.Rows("slide-1")("Title") = "Q3 Revenue (revised)", "Title updated")
    result = result & Assert(sheet.Rows("slide-1")("Date") = "2026-07", "Date survives untouched, got '" & sheet.Rows("slide-1")("Date") & "'")

    Test_UpsertRow_PartialUpdateMergesNotReplaces = result
End Function

Private Function Test_UpsertRow_NewInstanceDoesNotDisturbExistingRows() As String
    Dim result As String
    Dim ws As Object
    Set ws = NewBlankSheet()
    ExcelOutput.CreateSheet ws, "deck-v1"
    ExcelOutput.UpsertRow ws, "slide-1", DictOf2("Title", "Q3 Revenue", "Date", "2026-07"), "FY26Q4"
    ExcelOutput.UpsertRow ws, "slide-2", DictOf1("Title", "Q4 Revenue"), "FY26Q4"

    Dim sheet As Sheet
    sheet = ExcelOutput.ReadSheet(ws)
    result = result & Assert(sheet.InstanceOrder.count = 2 And sheet.InstanceOrder(1) = "slide-1" And sheet.InstanceOrder(2) = "slide-2", "append order preserved")
    result = result & Assert(sheet.Rows("slide-1")("Title") = "Q3 Revenue" And sheet.Rows("slide-1")("Date") = "2026-07", "slide-1 untouched")
    result = result & Assert(sheet.Rows("slide-2")("Title") = "Q4 Revenue", "slide-2 Title correct")
    result = result & Assert(Not sheet.Rows("slide-2").Exists("Date"), "slide-2 has no Date entry at all -- absent, not forced to empty string")

    Test_UpsertRow_NewInstanceDoesNotDisturbExistingRows = result
End Function

Private Function Test_ReadSheet_PreservesFieldAndInstanceOrderAcrossManyWrites() As String
    Dim result As String
    Dim ws As Object
    Set ws = NewBlankSheet()
    ExcelOutput.CreateSheet ws, "deck-v1"

    ExcelOutput.UpsertRow ws, "slide-3", DictOf1("Zeta", "z"), "FY26Q4"
    ExcelOutput.UpsertRow ws, "slide-1", DictOf1("Alpha", "a"), "FY26Q4"
    ExcelOutput.UpsertRow ws, "slide-2", DictOf2("Zeta", "z2", "Alpha", "a2"), "FY26Q4"

    Dim sheet As Sheet
    sheet = ExcelOutput.ReadSheet(ws)
    result = result & Assert(sheet.Fields.count = 2 And sheet.Fields(1) = "Zeta" And sheet.Fields(2) = "Alpha", "first-seen field order, not alphabetical")
    result = result & Assert(sheet.InstanceOrder.count = 3 And sheet.InstanceOrder(1) = "slide-3" And sheet.InstanceOrder(2) = "slide-1" And sheet.InstanceOrder(3) = "slide-2", "first-seen instance order")

    Test_ReadSheet_PreservesFieldAndInstanceOrderAcrossManyWrites = result
End Function

Private Function Test_HeaderRow_ReservesColumnsAAndBForIdentityAndPeriod() As String
    Dim result As String
    Dim ws As Object
    Set ws = NewBlankSheet()
    ExcelOutput.CreateSheet ws, "deck-v1"
    ExcelOutput.UpsertRow ws, "slide-1", DictOf1("Title", "Q3 Revenue"), "FY26Q4"

    result = result & Assert(CStr(ws.Cells(1, 1).Value) = ExcelOutput.INSTANCE_ID_HEADER, "A1 holds the Instance ID header, got '" & CStr(ws.Cells(1, 1).Value) & "'")
    result = result & Assert(CStr(ws.Cells(1, 2).Value) = ExcelOutput.QUARTER_HEADER, "B1 holds the Quarter header, got '" & CStr(ws.Cells(1, 2).Value) & "'")
    result = result & Assert(CStr(ws.Cells(1, 3).Value) = "Title", "C1 holds the first field name -- fields start AFTER both structural columns, got '" & CStr(ws.Cells(1, 3).Value) & "'")
    result = result & Assert(CStr(ws.Cells(2, 1).Value) = "slide-1", "A2 holds the first instance id")
    result = result & Assert(CStr(ws.Cells(2, 2).Value) = "FY26Q4", "B2 holds the period the row was written for")

    Test_HeaderRow_ReservesColumnsAAndBForIdentityAndPeriod = result
End Function

' --- ReadSheetForDeckPeriod: the sync path's guarded read ------------------
'
' Built by hand rather than through CreateSheet/UpsertRow, because neither of
' those writes a Quarter yet -- these describe the shape the sync path must
' cope with TODAY, which is a sheet a person or a migration produced.

Private Function WideSheet() As Object
    Dim ws As Object
    Set ws = NewBlankSheet()
    ws.Cells(1, 1).Value = ExcelOutput.INSTANCE_ID_HEADER
    ws.Cells(1, 2).Value = ExcelOutput.QUARTER_HEADER
    ws.Cells(1, 3).Value = "PROJECT_STATUS"
    Set WideSheet = ws
End Function

Private Sub WideRow(ws As Object, r As Long, instanceId As String, quarter As String, status As String)
    ws.Cells(r, 1).Value = instanceId
    ws.Cells(r, 2).Value = quarter
    ws.Cells(r, 3).Value = status
End Sub

Private Function Test_ReadForDeckPeriod_KeepsOnlyThatPeriodsRows() As String
    Dim ws As Object
    Set ws = WideSheet()
    WideRow ws, 2, "P1", "FY26Q4", "In Progress"
    WideRow ws, 3, "P1", "FY27Q1", "Not Started"
    WideRow ws, 4, "P2", "FY26Q4", "Project Closed"

    Dim problem As String
    Dim s As Sheet
    s = ExcelOutput.ReadSheetForDeckPeriod(ws, "FY26Q4", problem)

    Dim result As String
    result = result & Assert(problem = "", "a clean two-period sheet raises no problem, got: " & problem)
    result = result & Assert(s.InstanceOrder.count = 2, "FY26Q4 has 2 slides, got " & s.InstanceOrder.count)
    ' The discriminator: a read that ignored the period would say In Progress
    ' for P1 either way, because FY26Q4's row sits higher. The OTHER period is
    ' what proves the filter ran.
    result = result & Assert(s.Rows("P1")("PROJECT_STATUS") = "In Progress", "P1 @ FY26Q4 is In Progress")

    s = ExcelOutput.ReadSheetForDeckPeriod(ws, "FY27Q1", problem)
    result = result & Assert(problem = "", "FY27Q1 raises no problem, got: " & problem)
    result = result & Assert(s.InstanceOrder.count = 1, "FY27Q1 has 1 slide, got " & s.InstanceOrder.count)
    result = result & Assert(s.Rows("P1")("PROJECT_STATUS") = "Not Started", "P1 @ FY27Q1 is Not Started")

    ws.Delete
    Test_ReadForDeckPeriod_KeepsOnlyThatPeriodsRows = result
End Function

Private Function Test_ReadForDeckPeriod_RefusesTwoRowsForOneSlideInOnePeriod() As String
    Dim ws As Object
    Set ws = WideSheet()
    WideRow ws, 2, "P1", "FY26Q4", "In Progress"
    WideRow ws, 3, "P1", "FY26Q4", "Project Closed"

    Dim problem As String
    Dim s As Sheet
    s = ExcelOutput.ReadSheetForDeckPeriod(ws, "FY26Q4", problem)

    Dim result As String
    result = result & Assert(problem <> "", "two rows for one slide in one period must be refused")
    result = result & Assert(InStr(problem, "repeat") > 0, "the refusal says what is wrong, got: " & problem)

    ws.Delete
    Test_ReadForDeckPeriod_RefusesTwoRowsForOneSlideInOnePeriod = result
End Function

Private Function Test_ReadForDeckPeriod_RefusesAPeriodTheSheetDoesNotHave() As String
    Dim ws As Object
    Set ws = WideSheet()
    WideRow ws, 2, "P1", "FY26Q4", "In Progress"

    Dim problem As String
    Dim s As Sheet
    s = ExcelOutput.ReadSheetForDeckPeriod(ws, "FY28Q3", problem)

    Dim result As String
    result = result & Assert(s.InstanceOrder.count = 0, "nothing matches FY28Q3")
    result = result & Assert(problem <> "", "an empty filtered read of a non-empty sheet must be refused, not returned as success")

    ws.Delete
    Test_ReadForDeckPeriod_RefusesAPeriodTheSheetDoesNotHave = result
End Function

Private Function Test_ReadForDeckPeriod_SilentOnASheetWithNoQuarterColumn() As String
    ' Every sheet this tool wrote before 2026-08-03. It has one row per slide
    ' and no opinion about periods, so a deck declaring one must still read it.
    '
    ' BUILT BY HAND, deliberately: CreateSheet now always writes a Quarter
    ' header, so it can no longer produce this shape. The shape still exists on
    ' disk in every workbook made before that change, which is exactly why it
    ' still needs a test -- it is the one case that cannot be regenerated.
    Dim ws As Object
    Set ws = NewBlankSheet()
    ws.Cells(1, 1).Value = ExcelOutput.INSTANCE_ID_HEADER
    ws.Cells(1, 2).Value = "PROJECT_STATUS"
    ws.Cells(2, 1).Value = "P1"
    ws.Cells(2, 2).Value = "In Progress"

    Dim problem As String
    Dim s As Sheet
    s = ExcelOutput.ReadSheetForDeckPeriod(ws, "FY26Q4", problem)

    Dim result As String
    result = result & Assert(problem = "", "an old-shape sheet is not refused, got: " & problem)
    result = result & Assert(s.InstanceOrder.count = 1, "its row is still read, got " & s.InstanceOrder.count)

    ws.Delete
    Test_ReadForDeckPeriod_SilentOnASheetWithNoQuarterColumn = result
End Function

Private Function Test_ReadForDeckPeriod_RefusesANeverInitializedSheet() As String
    ' Mirrors WorkbookBridge.GetOrAddWorksheet creating a brand-new tab because
    ' the name it looked up didn't exist -- the exact shape of the wrong-sheet
    ' hazard (a registered type's worksheet name not matching the sheet a
    ' person actually built). A1 is empty because CreateSheet never ran.
    Dim ws As Object
    Set ws = NewBlankSheet()

    Dim problem As String
    Dim s As Sheet
    s = ExcelOutput.ReadSheetForDeckPeriod(ws, "FY26Q4", problem)

    Dim result As String
    result = result & Assert(problem <> "", "a never-initialized sheet must be refused, not read as a clean empty sync")
    result = result & Assert(InStr(problem, "never been set up") > 0, "the refusal says what is wrong, got: " & problem)

    ws.Delete
    Test_ReadForDeckPeriod_RefusesANeverInitializedSheet = result
End Function

' --- UpsertRow keyed on (instance, period) --------------------------------

' THE ONE THAT MATTERS. Until 2026-08-04 UpsertRow matched on instance alone,
' so writing FY27Q1 for a project found its FY26Q4 row and overwrote it --
' a real quarter's approved text destroyed silently, on rollover.
'
' Asserts BOTH directions, because a version that appended without matching
' would also produce two rows: last quarter's values must still be readable
' AT last quarter, and this quarter's at this quarter.
Private Function Test_UpsertRow_NextPeriodAppendsAndLeavesLastQuarterIntact() As String
    Dim result As String
    Dim ws As Object
    Set ws = NewBlankSheet()
    ExcelOutput.CreateSheet ws, "deck-v1"

    ExcelOutput.UpsertRow ws, "P1", DictOf1("PROJECT_STATUS", "In Progress"), "FY26Q4"
    ExcelOutput.UpsertRow ws, "P1", DictOf1("PROJECT_STATUS", "Project Closed"), "FY27Q1"

    Dim q4 As Sheet, q1 As Sheet
    Dim problem As String
    ' Presence is checked BEFORE dereferencing, so the regression this test
    ' exists to catch reports as a FAIL with a readable message. Reading a
    ' missing key straight out of a Scripting.Dictionary silently creates it as
    ' Empty, and indexing into Empty raises Type mismatch -- which surfaced as
    ' ERROR, not FAIL, when this was first broken on purpose.
    q4 = ExcelOutput.ReadSheetForDeckPeriod(ws, "FY26Q4", problem)
    result = result & Assert(problem = "", "FY26Q4 reads cleanly, got: " & problem)
    result = result & Assert(q4.InstanceOrder.count = 1, "FY26Q4 has one row, got " & q4.InstanceOrder.count)
    If Not q4.Rows.Exists("P1") Then
        result = result & Assert(False, "P1 is missing from the FY26Q4 read -- last quarter's row was overwritten")
    Else
        result = result & Assert(q4.Rows("P1")("PROJECT_STATUS") = "In Progress", _
            "LAST QUARTER'S TEXT SURVIVED -- got '" & q4.Rows("P1")("PROJECT_STATUS") & "'")
    End If

    q1 = ExcelOutput.ReadSheetForDeckPeriod(ws, "FY27Q1", problem)
    result = result & Assert(problem = "", "FY27Q1 reads cleanly, got: " & problem)
    result = result & Assert(q1.InstanceOrder.count = 1, "FY27Q1 has one row, got " & q1.InstanceOrder.count)
    If Not q1.Rows.Exists("P1") Then
        result = result & Assert(False, "P1 is missing from the FY27Q1 read -- no row was appended for the new period")
    Else
        result = result & Assert(q1.Rows("P1")("PROJECT_STATUS") = "Project Closed", _
            "this quarter's text is its own -- got '" & q1.Rows("P1")("PROJECT_STATUS") & "'")
    End If

    ' Two rows on the sheet, not one overwritten and not three.
    result = result & Assert(CStr(ws.Cells(3, 1).Value) = "P1" And CStr(ws.Cells(3, 2).Value) = "FY27Q1", _
        "row 3 is P1 @ FY27Q1, got '" & CStr(ws.Cells(3, 1).Value) & "' / '" & CStr(ws.Cells(3, 2).Value) & "'")
    result = result & Assert(IsEmpty(ws.Cells(4, 1).Value), "no fourth row was created")

    ws.Delete
    Test_UpsertRow_NextPeriodAppendsAndLeavesLastQuarterIntact = result
End Function

' The other half: re-syncing the SAME period must not accumulate rows.
Private Function Test_UpsertRow_SamePeriodUpdatesThatRowInPlace() As String
    Dim result As String
    Dim ws As Object
    Set ws = NewBlankSheet()
    ExcelOutput.CreateSheet ws, "deck-v1"

    ExcelOutput.UpsertRow ws, "P1", DictOf1("PROJECT_STATUS", "Not Started"), "FY26Q4"
    ExcelOutput.UpsertRow ws, "P1", DictOf1("PROJECT_STATUS", "In Progress"), "FY26Q4"

    Dim problem As String
    Dim s As Sheet
    s = ExcelOutput.ReadSheetForDeckPeriod(ws, "FY26Q4", problem)

    result = result & Assert(problem = "", "no duplicate is reported, got: " & problem)
    result = result & Assert(s.InstanceOrder.count = 1, "still one row, got " & s.InstanceOrder.count)
    result = result & Assert(s.Rows("P1")("PROJECT_STATUS") = "In Progress", "the row was updated in place")
    result = result & Assert(IsEmpty(ws.Cells(3, 1).Value), "no second row was appended")

    ws.Delete
    Test_UpsertRow_SamePeriodUpdatesThatRowInPlace = result
End Function

' A blank period on a period-bearing sheet writes a row no filtered read can
' ever see -- present, correct, and invisible. Refused rather than defaulted.
Private Function Test_UpsertRow_RefusesABlankPeriodOnAPeriodSheet() As String
    Dim result As String
    Dim ws As Object
    Set ws = NewBlankSheet()
    ExcelOutput.CreateSheet ws, "deck-v1"

    Dim raised As Boolean
    On Error Resume Next
    Err.Clear
    ExcelOutput.UpsertRow ws, "P1", DictOf1("PROJECT_STATUS", "In Progress"), ""
    raised = (Err.Number <> 0)
    Err.Clear
    On Error GoTo 0

    result = result & Assert(raised, "a blank period must be refused on a sheet that has a Quarter column")
    result = result & Assert(IsEmpty(ws.Cells(2, 1).Value), "and nothing was written")

    ws.Delete
    Test_UpsertRow_RefusesABlankPeriodOnAPeriodSheet = result
End Function

' A slide really can carry a field called "Quarter". Before this it went
' straight into the period cell and the row vanished from every filtered read.
Private Function Test_UpsertRow_RefusesAFieldNamedLikeAStructuralColumn() As String
    Dim result As String
    Dim ws As Object
    Set ws = NewBlankSheet()
    ExcelOutput.CreateSheet ws, "deck-v1"
    ExcelOutput.UpsertRow ws, "P1", DictOf1("PROJECT_STATUS", "In Progress"), "FY26Q4"

    Dim raised As Boolean
    On Error Resume Next
    Err.Clear
    ExcelOutput.UpsertRow ws, "P1", DictOf1("Quarter", "clobbered"), "FY26Q4"
    raised = (Err.Number <> 0)
    Err.Clear
    On Error GoTo 0

    result = result & Assert(raised, "a field named 'Quarter' must be refused, not written into the period column")
    result = result & Assert(CStr(ws.Cells(2, 2).Value) = "FY26Q4", _
        "the period cell is untouched, got '" & CStr(ws.Cells(2, 2).Value) & "'")

    ws.Delete
    Test_UpsertRow_RefusesAFieldNamedLikeAStructuralColumn = result
End Function

' Sheets built before 2026-08-03 have one row per slide and no period column.
' UpsertRow must keep updating them in place rather than appending a row per
' period into a sheet that cannot express periods.
Private Function Test_UpsertRow_LegacySheetWithNoPeriodColumnStillMatchesOnInstance() As String
    Dim result As String
    Dim ws As Object
    Set ws = NewBlankSheet()
    ws.Cells(1, 1).Value = ExcelOutput.INSTANCE_ID_HEADER
    ws.Cells(1, 2).Value = "PROJECT_STATUS"
    ws.Cells(2, 1).Value = "P1"
    ws.Cells(2, 2).Value = "Not Started"

    ExcelOutput.UpsertRow ws, "P1", DictOf1("PROJECT_STATUS", "In Progress"), "FY26Q4"

    result = result & Assert(CStr(ws.Cells(2, 2).Value) = "In Progress", _
        "the existing row was updated, got '" & CStr(ws.Cells(2, 2).Value) & "'")
    result = result & Assert(IsEmpty(ws.Cells(3, 1).Value), "no period row was appended to a sheet with no period column")

    ws.Delete
    Test_UpsertRow_LegacySheetWithNoPeriodColumnStillMatchesOnInstance = result
End Function
