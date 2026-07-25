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
    r = Test_HeaderRow_ReservesColumnAForInstanceId()
    AppendResult report, "HeaderRow_ReservesColumnAForInstanceId", r
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

    ExcelOutput.UpsertRow ws, "slide-1", DictOf2("Title", "Q3 Revenue", "Date", "2026-07")

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
    ExcelOutput.UpsertRow ws, "slide-1", DictOf2("Title", "Q3 Revenue", "Date", "2026-07")

    ExcelOutput.UpsertRow ws, "slide-1", DictOf1("Region", "APAC")

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
    ExcelOutput.UpsertRow ws, "slide-1", DictOf2("Title", "Q3 Revenue", "Date", "2026-07")

    ExcelOutput.UpsertRow ws, "slide-1", DictOf1("Title", "Q3 Revenue (revised)")

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
    ExcelOutput.UpsertRow ws, "slide-1", DictOf2("Title", "Q3 Revenue", "Date", "2026-07")
    ExcelOutput.UpsertRow ws, "slide-2", DictOf1("Title", "Q4 Revenue")

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

    ExcelOutput.UpsertRow ws, "slide-3", DictOf1("Zeta", "z")
    ExcelOutput.UpsertRow ws, "slide-1", DictOf1("Alpha", "a")
    ExcelOutput.UpsertRow ws, "slide-2", DictOf2("Zeta", "z2", "Alpha", "a2")

    Dim sheet As Sheet
    sheet = ExcelOutput.ReadSheet(ws)
    result = result & Assert(sheet.Fields.count = 2 And sheet.Fields(1) = "Zeta" And sheet.Fields(2) = "Alpha", "first-seen field order, not alphabetical")
    result = result & Assert(sheet.InstanceOrder.count = 3 And sheet.InstanceOrder(1) = "slide-3" And sheet.InstanceOrder(2) = "slide-1" And sheet.InstanceOrder(3) = "slide-2", "first-seen instance order")

    Test_ReadSheet_PreservesFieldAndInstanceOrderAcrossManyWrites = result
End Function

Private Function Test_HeaderRow_ReservesColumnAForInstanceId() As String
    Dim result As String
    Dim ws As Object
    Set ws = NewBlankSheet()
    ExcelOutput.CreateSheet ws, "deck-v1"
    ExcelOutput.UpsertRow ws, "slide-1", DictOf1("Title", "Q3 Revenue")

    result = result & Assert(CStr(ws.Cells(1, 1).Value) = ExcelOutput.INSTANCE_ID_HEADER, "A1 holds the Instance ID header, got '" & CStr(ws.Cells(1, 1).Value) & "'")
    result = result & Assert(CStr(ws.Cells(1, 2).Value) = "Title", "B1 holds the first field name")
    result = result & Assert(CStr(ws.Cells(2, 1).Value) = "slide-1", "A2 holds the first instance id")

    Test_HeaderRow_ReservesColumnAForInstanceId = result
End Function
