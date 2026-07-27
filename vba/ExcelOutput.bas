Attribute VB_Name = "ExcelOutput"
Option Explicit

' VBA port of src/excel_output.py, per specs/vba-port.md's port order
' (module 6 of 6, the last one -- after discovery/identity_tags/matching/
' resolve+sync_operations/onboarding).
'
' This is the one module vba-port.md itself says is strictly SIMPLER than
' the Python it replaces, not just a mechanism swap: Python hand-rebuilds
' the whole .xlsx zip (Content_Types, workbook.xml, styles.xml, the sheet
' part, custom.xml) from scratch on every write, because it has no host
' application to lean on. VBA runs against a live Worksheet the calling
' context already has open (in Excel directly, or via COM automation driven
' from the PowerPoint side) -- so this is plain Cells/Range reads and
' writes, no XML, no zip, no "regenerate the whole file" step. Per
' vba-port.md: "Don't port excel_output.py's zip-rebuilding approach -- it
' only existed to work around Python lacking a live Excel instance."
'
' Layout convention preserved exactly from excel-output.md/excel_output.py:
' column A is the reserved instance-identity column (header "Instance ID");
' columns B.. hold confirmed fields, one per column, in first-seen (append)
' order; row 1 is the header, rows 2.. are data keyed by instance ID (never
' by row position). Field identity is the column's header TEXT, looked up
' by name on every read -- never assumed from position.
'
' See SPIKE_NOTES_ExcelOutput.md for deliberate divergences and the manual
' verification recipe -- there is no .xlsx test fixture for this spec (same
' as the Python side: this module is both writer and reader, so its own
' tests are round-trip/self-consistency checks).

Public Const INSTANCE_ID_HEADER As String = "Instance ID"
Private Const DECK_REFERENCE_PROPERTY_NAME As String = "DeckReference"

' XlDirection enum values, as numeric literals rather than the named
' constants (xlToLeft/xlUp) -- confirmed real (2026-07-25) that the named
' forms only resolve when this module runs inside Excel's own VBA project
' (which has the Excel type library referenced natively). Driven cross-app
' from PowerPoint (RunSync.bas's actual real usage, per vba-port.md's "VBA
' runs inside Excel or drives it via COM from the PowerPoint side"), the
' PowerPoint-hosted project has no such reference, and the named constants
' raise a compile error ("Variable not defined") -- found via a real
' PowerPoint-driven end-to-end test, not caught by any of ExcelOutput's own
' prior tests since those all ran inside Excel's own project, where the
' names happened to resolve. Numeric values are stable, documented Office
' constants, unaffected by which host application's project this runs in.
Private Const XL_TO_LEFT As Long = -4159
Private Const XL_UP As Long = -4162

' Sheet.Fields/InstanceOrder are Collections (ordered, append-only), not
' Dictionary keys -- matches this project's existing convention for ordered
' lists (SyncOperations.bas's instanceOrder) rather than relying on
' Scripting.Dictionary's de-facto-but-undocumented key-insertion order.
' Sheet.Rows is a Scripting.Dictionary of Scripting.Dictionaries
' (instanceId -> fieldName -> value) -- legal because Dictionary/Collection
' are Objects, not UDTs; see Onboarding.bas's SPIKE_NOTES for why a
' Dictionary could NOT hold this data if the values were a UDT instead of
' another Dictionary.
Public Type Sheet
    DeckReference As String
    Fields As Collection
    InstanceOrder As Collection
    Rows As Object
End Type

' ---------------------------------------------------------------------
' Create
' ---------------------------------------------------------------------

' Set up `ws` as a fresh, empty data sheet bound to `deckReference`.
' Refuses to reinitialize a sheet that already has a header in A1 -- a
' second "create" against an already-set-up sheet is almost certainly a
' mistake (mirrors create_sheet's FileExistsError; the "file" here is
' represented by the sheet already carrying content, since a live Worksheet
' has no separate "exists on disk yet" concept the way a path does).
Public Sub CreateSheet(ws As Object, deckReference As String)
    If Not IsEmpty(ws.Cells(1, 1).Value) Then
        Err.Raise vbObjectError + 1, "ExcelOutput.CreateSheet", _
            "refusing to initialize an already-set-up sheet (A1 is not empty) -- possible accidental double-create"
    End If

    ws.Cells(1, 1).Value = INSTANCE_ID_HEADER
    WriteDeckReference ws.Parent, deckReference
End Sub

' ---------------------------------------------------------------------
' Read
' ---------------------------------------------------------------------

' Read `ws` back into a Sheet. Fields are recovered from the header row
' (row 1, columns B..), instance rows from column A (rows 2..) -- both via
' Excel's End(xlToLeft)/End(xlUp) "walk from the far side" idiom rather than
' a stored count, the standard reliable way to find a used range's true
' extent in the object model.
Public Function ReadSheet(ws As Object) As Sheet
    Dim result As Sheet
    Set result.Fields = New Collection
    Set result.InstanceOrder = New Collection
    Set result.Rows = CreateObject("Scripting.Dictionary")

    result.DeckReference = ReadDeckReference(ws.Parent)

    Dim lastCol As Long
    lastCol = LastUsedColumn(ws)
    Dim c As Long
    For c = 2 To lastCol
        result.Fields.Add CStr(ws.Cells(1, c).Value)
    Next c

    Dim lastRow As Long
    lastRow = LastUsedRow(ws)
    Dim r As Long
    For r = 2 To lastRow
        Dim instanceId As String
        instanceId = CStr(ws.Cells(r, 1).Value)
        If instanceId <> "" Then
            result.InstanceOrder.Add instanceId

            Dim rowValues As Object
            Set rowValues = CreateObject("Scripting.Dictionary")
            For c = 2 To lastCol
                ' IsEmpty (not "= """"") to distinguish "field never
                ' harvested for this instance" from "harvested value happens
                ' to be an empty string" -- mirrors read_sheet's structural
                ' cell-presence check (Python looks at whether a <c> element
                ' exists at all, not whether its text is falsy).
                If Not IsEmpty(ws.Cells(r, c).Value) Then
                    rowValues(result.Fields(c - 1)) = CStr(ws.Cells(r, c).Value)
                End If
            Next c
            Set result.Rows(instanceId) = rowValues
        End If
    Next r

    ReadSheet = result
End Function

Private Function LastUsedColumn(ws As Object) As Long
    If IsEmpty(ws.Cells(1, 1).Value) Then
        LastUsedColumn = 0
        Exit Function
    End If
    LastUsedColumn = ws.Cells(1, ws.Columns.count).End(XL_TO_LEFT).Column
End Function

Private Function LastUsedRow(ws As Object) As Long
    If IsEmpty(ws.Cells(1, 1).Value) Then
        LastUsedRow = 0
        Exit Function
    End If
    LastUsedRow = ws.Cells(ws.Rows.count, 1).End(XL_UP).Row
End Function

Private Sub WriteDeckReference(wb As Object, deckReference As String)
    Dim prop As Object
    On Error Resume Next
    Set prop = wb.CustomDocumentProperties(DECK_REFERENCE_PROPERTY_NAME)
    On Error GoTo 0

    If prop Is Nothing Then
        wb.CustomDocumentProperties.Add Name:=DECK_REFERENCE_PROPERTY_NAME, _
            LinkToContent:=False, Type:=msoPropertyTypeString, Value:=deckReference
    Else
        prop.Value = deckReference
    End If
End Sub

Private Function ReadDeckReference(wb As Object) As String
    Dim prop As Object
    On Error Resume Next
    Set prop = wb.CustomDocumentProperties(DECK_REFERENCE_PROPERTY_NAME)
    On Error GoTo 0

    If prop Is Nothing Then
        ReadDeckReference = ""
    Else
        ReadDeckReference = CStr(prop.Value)
    End If
End Function

' ---------------------------------------------------------------------
' Upsert
' ---------------------------------------------------------------------

' Add any of `values`'s keys not already a known field as a new column
' (appended after the last used column, never replacing/reordering existing
' ones), then create or update `instanceId`'s row -- direct incremental
' Cells writes, not a read-whole-sheet/mutate/rewrite-whole-file cycle
' (unlike upsert_row, which must rebuild the entire .xlsx because that's the
' only write primitive a headless zip has; a live Worksheet doesn't need
' that). A new instance is appended as a new row, seeded entirely from
' `values`. An existing instance only has the given keys' cells written --
' any field this call doesn't mention is left completely untouched, so a
' partial re-sync of one changed field can never blank out the rest.
'
' `values` is a Scripting.Dictionary (fieldName -> value String), matching
' the shape SyncOperations.bas's `dataRows` entries already use.
'
' Deliberately does NOT apply real typed Excel formatting (a Date/Double
' value + NumberFormat) even though BatchOnboardFlow.bas now captures a
' field type at mark time -- traced 2026-07-26 that this Sheet's own values
' feed directly back into SyncOperations.PlanRoutineSync -> InjectPrimitive,
' which WRITES Excel's value onto the live PowerPoint slide on every
' routine sync. A typed cell's read-back (ReadSheet's CStr(.Value)) is not
' guaranteed to equal the exact string that was written -- a Date reads
' back locale-formatted, a Double can drop a trailing zero -- so writing a
' real typed value here would risk a routine sync silently rewriting a
' slide's date/number text into a reformatted (though "equal") version the
' human never asked for. That's a slide-content mutation, not a formatting
' nicety, and conflicts with this project's founding invariant that nothing
' gets silently mutated (InjectPrimitive/Verification exist specifically to
' guard it). Every value here is always written and read back as the exact
' harvested string, unconditionally -- the field type is still captured and
' shown to a human (BatchOnboardFlow's Field Review grid), just not acted
' on here. Revisit only alongside making PlanRoutineSync's own comparison
' type-aware, not by touching this function in isolation.
Public Sub UpsertRow(ws As Object, instanceId As String, values As Object)
    Dim rowNum As Long
    rowNum = FindOrAppendInstanceRow(ws, instanceId)

    Dim fieldName As Variant
    For Each fieldName In values.Keys
        Dim colNum As Long
        colNum = FindOrAppendFieldColumn(ws, CStr(fieldName))
        ws.Cells(rowNum, colNum).Value = CStr(values(fieldName))
    Next fieldName
End Sub

Private Function FindOrAppendFieldColumn(ws As Object, fieldName As String) As Long
    Dim lastCol As Long
    lastCol = LastUsedColumn(ws)
    Dim c As Long
    For c = 2 To lastCol
        If CStr(ws.Cells(1, c).Value) = fieldName Then
            FindOrAppendFieldColumn = c
            Exit Function
        End If
    Next c

    FindOrAppendFieldColumn = lastCol + 1
    ws.Cells(1, lastCol + 1).Value = fieldName
End Function

Private Function FindOrAppendInstanceRow(ws As Object, instanceId As String) As Long
    Dim lastRow As Long
    lastRow = LastUsedRow(ws)
    Dim r As Long
    For r = 2 To lastRow
        If CStr(ws.Cells(r, 1).Value) = instanceId Then
            FindOrAppendInstanceRow = r
            Exit Function
        End If
    Next r

    FindOrAppendInstanceRow = lastRow + 1
    ws.Cells(lastRow + 1, 1).Value = instanceId
End Function

' ---------------------------------------------------------------------
' Manual smoke test -- not a real test harness, same as every other module
' here. See SPIKE_NOTES_ExcelOutput.md for the full recipe and expected
' values, cross-checked against tests/test_excel_output.py's already-proven
' round-trip results.
' ---------------------------------------------------------------------

' Run against a blank worksheet in the active workbook (e.g. add a new
' sheet first so A1 is genuinely empty).
Public Sub ManualSmokeTest(ws As Object)
    CreateSheet ws, "deck-v1"

    Dim v1 As Object
    Set v1 = CreateObject("Scripting.Dictionary")
    v1("Title") = "Q3 Revenue"
    v1("Date") = "2026-07"
    UpsertRow ws, "slide-1", v1

    Dim v2 As Object
    Set v2 = CreateObject("Scripting.Dictionary")
    v2("Region") = "APAC"
    UpsertRow ws, "slide-1", v2 ' new field, existing instance -- appends a column, doesn't disturb Title/Date

    Dim v3 As Object
    Set v3 = CreateObject("Scripting.Dictionary")
    v3("Title") = "Q4 Revenue"
    UpsertRow ws, "slide-2", v3 ' new instance -- new row, no Date/Region yet

    Dim sheet As Sheet
    sheet = ReadSheet(ws)

    Dim msg As String
    msg = "DeckReference=" & sheet.DeckReference & vbCrLf
    msg = msg & "Fields=" & JoinCollection(sheet.Fields) & vbCrLf
    msg = msg & "InstanceOrder=" & JoinCollection(sheet.InstanceOrder) & vbCrLf
    msg = msg & "slide-1: Title=" & sheet.Rows("slide-1")("Title") & " Date=" & sheet.Rows("slide-1")("Date") & " Region=" & sheet.Rows("slide-1")("Region") & vbCrLf
    msg = msg & "slide-2: Title=" & sheet.Rows("slide-2")("Title") & " HasDate=" & sheet.Rows("slide-2").Exists("Date")

    Debug.Print msg
    MsgBox msg & vbCrLf & "(expected: DeckReference=deck-v1, Fields=Title,Date,Region, InstanceOrder=slide-1,slide-2, slide-2 HasDate=False)"
End Sub

Private Function JoinCollection(coll As Collection) As String
    Dim i As Long, result As String
    For i = 1 To coll.count
        If i > 1 Then result = result & ","
        result = result & coll(i)
    Next i
    JoinCollection = result
End Function
