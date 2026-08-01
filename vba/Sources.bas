Attribute VB_Name = "Sources"
Option Explicit

' WHERE THE WORDS CAME FROM.
'
' A drafting sheet records what a field was changed TO. It has never recorded
' what that change was based on, so a quarter later "why does it say 90%?" has
' no answer except somebody's memory. This sheet is that answer.
'
' ONE ROW PER SOURCE, EVER. Drafting rows refer to sources by ID, not by
' description. A milestone report used by all 43 projects is written down once
' and referenced 43 times -- which is the difference between a record and a
' mess. Rohan's constraint, 2026-08-01: "Sources should be managed so it
' doesn't get silly."
'
' The three things that make it stay manageable:
'   1. IDs are short and stable (S01, S02 ...). Retyping a label 43 times gives
'      43 spellings; retyping an ID gives an ID.
'   2. The sheet is never rebuilt from scratch -- rows accumulate and survive
'      every drafting refresh, because a source outlives the field it was used
'      for.
'   3. A reference to an ID that does not exist is REPORTED, not silently
'      ignored. An unnoticed typo turns the record into decoration.
'
' Deliberately NOT here: source content. This holds pointers -- what it is,
' where it lives. Pasting document text into a spreadsheet is how the sheet
' becomes unreadable, which is the failure mode being designed against.

Public Const SOURCES_SHEET_NAME As String = "Sources"

Public Const COL_S_ID As Long = 1
Public Const COL_S_LABEL As Long = 2
Public Const COL_S_TYPE As Long = 3
Public Const COL_S_LOCATOR As Long = 4
Public Const COL_S_NOTES As Long = 5
Public Const COL_S_ADDED As Long = 6

Public Const SRC_INTRO_ROW As Long = 1
Public Const SRC_HEADER_ROW As Long = 5
Public Const SRC_FIRST_ROW As Long = 6

' Creates the sheet if it is absent and (re)writes only the furniture --
' headings, instructions, widths. EXISTING ROWS ARE NEVER TOUCHED: this runs on
' every drafting build, and a function that ran 40 times a session and could
' delete the provenance record would be a liability rather than a convenience.
Public Function WriteSourcesSheet(ws As Object) As String
    ws.Cells(SRC_INTRO_ROW, 1).Value = "SOURCES  --  what the drafting sheets are based on"
    ws.Cells(SRC_INTRO_ROW, 1).Font.Bold = True
    ws.Cells(SRC_INTRO_ROW, 1).Font.Size = 9

    ws.Cells(2, 1).Value = "Add a row for anything you drafted from. Give it the next free ID, then put that ID in column E of the drafting sheet."
    ws.Cells(3, 1).Value = "One row per source, ever -- if 20 projects use the same report, they all reference the same ID. Do not paste document text in here; point at it."

    ws.Cells(4, 1).Value = "A drafting row referring to an ID that is not listed here is reported when you publish. It will not stop the publish, but it means the record is wrong."
    ws.Cells(4, 1).Font.Italic = True

    ws.Cells(SRC_HEADER_ROW, COL_S_ID).Value = "ID"
    ws.Cells(SRC_HEADER_ROW, COL_S_LABEL).Value = "What it is"
    ws.Cells(SRC_HEADER_ROW, COL_S_TYPE).Value = "Type"
    ws.Cells(SRC_HEADER_ROW, COL_S_LOCATOR).Value = "Where it lives"
    ws.Cells(SRC_HEADER_ROW, COL_S_NOTES).Value = "Notes"
    ws.Cells(SRC_HEADER_ROW, COL_S_ADDED).Value = "Added"
    ws.Rows(SRC_HEADER_ROW).Font.Bold = True

    ws.Cells.Font.Size = 8
    ws.Cells(SRC_INTRO_ROW, 1).Font.Size = 9
    ws.Columns(COL_S_ID).ColumnWidth = 7
    ws.Columns(COL_S_LABEL).ColumnWidth = 44
    ws.Columns(COL_S_TYPE).ColumnWidth = 10
    ws.Columns(COL_S_LOCATOR).ColumnWidth = 46
    ws.Columns(COL_S_NOTES).ColumnWidth = 34
    ws.Columns(COL_S_ADDED).ColumnWidth = 11
    ws.Cells.VerticalAlignment = -4160        ' xlTop

    Dim n As Long
    n = CountSources(ws)
    WriteSourcesSheet = "Sources: " & n & " source(s) on record. Next free ID is " & NextSourceId(ws) & "."
End Function

Public Function CountSources(ws As Object) As Long
    Dim r As Long
    r = SRC_FIRST_ROW
    Do While Trim(CStr(ws.Cells(r, COL_S_ID).Value)) <> ""
        CountSources = CountSources + 1
        r = r + 1
    Loop
End Function

' The lowest unused S-number, so two people adding a source on the same day do
' not both reach for S04. Numeric, not "count + 1": a deleted row in the middle
' would otherwise hand out an ID that is already in use further down.
Public Function NextSourceId(ws As Object) As String
    Dim highest As Long
    Dim r As Long
    r = SRC_FIRST_ROW
    Do While Trim(CStr(ws.Cells(r, COL_S_ID).Value)) <> ""
        Dim raw As String
        raw = UCase(Trim(CStr(ws.Cells(r, COL_S_ID).Value)))
        If Left(raw, 1) = "S" Then
            Dim numPart As String
            numPart = Mid(raw, 2)
            If IsNumeric(numPart) Then
                If CLng(numPart) > highest Then highest = CLng(numPart)
            End If
        End If
        r = r + 1
    Loop
    NextSourceId = "S" & Format(highest + 1, "00")
End Function

' Every ID currently on the sheet, upper-cased, as a Dictionary for lookup.
Public Function KnownSourceIds(ws As Object) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim r As Long
    r = SRC_FIRST_ROW
    Do While Trim(CStr(ws.Cells(r, COL_S_ID).Value)) <> ""
        d(UCase(Trim(CStr(ws.Cells(r, COL_S_ID).Value)))) = CStr(ws.Cells(r, COL_S_LABEL).Value)
        r = r + 1
    Loop
    Set KnownSourceIds = d
End Function

' Which of the IDs in a drafting row's Sources cell are not on the Sources
' sheet. Returns "" when they all check out.
'
' Pure and separated deliberately: it is the only part of this module with a
' decision in it, and a decision that cannot be driven both ways from a test is
' one nobody ever watches fail.
Public Function UnknownRefs(refs As String, known As Object) As String
    If Trim(refs) = "" Then Exit Function
    If known Is Nothing Then Exit Function

    Dim parts As Variant
    parts = Split(Replace(Replace(refs, ";", ","), " ", ","), ",")

    Dim bad As String
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim one As String
        one = UCase(Trim(CStr(parts(i))))
        If one <> "" Then
            If Not known.Exists(one) Then
                If InStr(bad & ",", one & ",") = 0 Then
                    bad = bad & IIf(bad = "", "", ",") & one
                End If
            End If
        End If
    Next i
    UnknownRefs = bad
End Function
