Attribute VB_Name = "ExcelPropertyProbe"
Option Explicit

' EXPERIMENTAL, STANDALONE PROBE -- NOT part of the production build. Checks
' whether Excel's CustomDocumentProperties on a OneDrive-hosted workbook has
' the same "only the first write per session lands" limit PowerPoint showed
' (FIX-LIST P, DeckRegistry.bas). ExcelOutput.WriteDeckReference/
' ReadDeckReference use the exact same mechanism and have never been tested
' against this. Driven from VBA, not PowerShell directly -- PowerShell's own
' .NET COM interop cannot even introspect Excel's CustomDocumentProperties
' object (a NullReferenceException deep in
' ComRuntimeHelpers.GetTypeAttrForTypeInfo, confirmed live 2026-08-16, twice,
' with two different call shapes) -- a limitation of that interop layer, not
' a fact about the underlying mechanism. VBA's own native COM dispatch has no
' such problem, so the probe lives here instead.

' MIRRORS THE FIXED ExcelOutput.WriteDeckReference/ReadDeckReference EXACTLY
' (meta sheet, very-hidden, cell A1) -- proves the SPECIFIC new
' implementation survives a real re-pairing (write, then write again with a
' DIFFERENT value, the exact scenario that broke the old
' CustomDocumentProperties mechanism), not just "cells work on OneDrive in
' general" which this project's own extensive UpsertRow test coverage
' already established.
Private Const META_SHEET_NAME As String = "DeckSyncMetaProbe"

Private Function FindMetaSheetNew(wb As Object) As Object
    Dim ws As Object
    On Error Resume Next
    For Each ws In wb.Worksheets
        If ws.Name = META_SHEET_NAME Then
            Set FindMetaSheetNew = ws
            Exit Function
        End If
    Next ws
    On Error GoTo 0
End Function

Private Function FindOrCreateMetaSheetNew(wb As Object) As Object
    Dim existing As Object
    Set existing = FindMetaSheetNew(wb)
    If Not existing Is Nothing Then
        Set FindOrCreateMetaSheetNew = existing
        Exit Function
    End If

    Dim ws As Object
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
    ws.Name = META_SHEET_NAME
    ws.Visible = 2   ' xlSheetVeryHidden, numeric for cross-app safety
    Set FindOrCreateMetaSheetNew = ws
End Function

Public Sub WriteDeckReferenceNew(wb As Object, deckReference As String)
    Dim ws As Object
    Set ws = FindOrCreateMetaSheetNew(wb)
    ws.Cells(1, 1).Value = deckReference
End Sub

Public Function ReadDeckReferenceNew(wb As Object) As String
    Dim ws As Object
    Set ws = FindMetaSheetNew(wb)
    If ws Is Nothing Then Exit Function
    Dim v As Variant
    v = ws.Cells(1, 1).Value
    If Not IsEmpty(v) Then ReadDeckReferenceNew = CStr(v)
End Function

' MIRRORS ExcelOutput.WriteDeckReference EXACTLY, including the update path
' this probe exists to test: Add for a new property, `.Value = ` (never
' Delete+Add) for an existing one -- the real function's real behaviour.
Public Sub WriteLikeTheRealFunction(wb As Object, propName As String, value As String)
    Dim prop As Object
    On Error Resume Next
    Set prop = wb.CustomDocumentProperties(propName)
    On Error GoTo 0

    If prop Is Nothing Then
        wb.CustomDocumentProperties.Add Name:=propName, _
            LinkToContent:=False, Type:=4, Value:=value   ' 4 = msoPropertyTypeString
    Else
        prop.Value = value
    End If
End Sub

' The known-safe VBA pattern DeckRegistry.bas used before moving off this
' mechanism entirely -- delete and re-add rather than assign to an existing
' property.
Public Sub WriteViaDeleteAndAdd(wb As Object, propName As String, value As String)
    Dim prop As Object
    On Error Resume Next
    Set prop = wb.CustomDocumentProperties(propName)
    On Error GoTo 0

    If Not prop Is Nothing Then
        On Error Resume Next
        wb.CustomDocumentProperties(propName).Delete
        On Error GoTo 0
    End If

    wb.CustomDocumentProperties.Add Name:=propName, _
        LinkToContent:=False, Type:=4, Value:=value
End Sub

' Reads a custom document property straight from the SAVED FILE's
' docProps/custom.xml -- same OOXML part, same technique, as
' DeckRegistry.PropertyOnDisk uses for a .pptx. Copies the file, opens it as
' a zip via Shell.Application, extracts and reads the XML, never trusts
' Excel's own object cache.
Public Function PropertyOnDiskXlsx(ByVal wbPath As String, propertyName As String, _
                                   Optional ByRef trace As String) As String
    On Error GoTo Failed

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    trace = "fso ok"

    If LCase$(Left$(wbPath, 4)) = "http" Then
        trace = trace & " | cannot map URL in this standalone probe -- pass a local path"
        Exit Function
    End If
    If Not fso.FileExists(wbPath) Then
        trace = trace & " | file missing: " & wbPath
        Exit Function
    End If

    Dim work As String
    work = fso.BuildPath(fso.GetSpecialFolder(2).Path, "dsxlverify_" & Format(Now, "hhnnss") & Int(Rnd * 10000))
    fso.CreateFolder work

    Dim zipPath As String, outDir As String
    zipPath = fso.BuildPath(work, "wb.zip")
    outDir = fso.BuildPath(work, "out")
    fso.CreateFolder outDir
    fso.CopyFile wbPath, zipPath
    trace = trace & " | copied"

    Dim sh As Object
    Set sh = CreateObject("Shell.Application")

    Dim zipVar As Variant, outVar As Variant, propsVar As Variant
    zipVar = zipPath
    outVar = outDir

    Dim zipNs As Object
    Set zipNs = sh.Namespace(zipVar)
    If zipNs Is Nothing Then
        trace = trace & " | Namespace(zip) returned Nothing"
        GoTo Cleanup
    End If

    Dim props As Object
    Set props = zipNs.ParseName("docProps")
    If props Is Nothing Then
        trace = trace & " | docProps NOT FOUND"
        GoTo Cleanup
    End If

    propsVar = props.Path
    Dim propsNs As Object
    Set propsNs = sh.Namespace(propsVar)
    If propsNs Is Nothing Then
        trace = trace & " | Namespace(docProps) returned Nothing"
        GoTo Cleanup
    End If

    Dim customFile As Object
    Set customFile = propsNs.ParseName("custom.xml")
    If customFile Is Nothing Then
        trace = trace & " | custom.xml NOT FOUND"
        GoTo Cleanup
    End If

    sh.Namespace(outVar).CopyHere customFile, 16

    Dim extracted As String
    extracted = fso.BuildPath(outDir, "custom.xml")

    Dim waited As Long
    Do While Not fso.FileExists(extracted) And waited < 100
        DoEvents
        waited = waited + 1
    Loop
    If Not fso.FileExists(extracted) Then
        trace = trace & " | EXTRACT TIMED OUT"
        GoTo Cleanup
    End If

    Dim xml As String
    xml = fso.OpenTextFile(extracted, 1).ReadAll

    Dim atProp As Long
    atProp = InStr(1, xml, propertyName, vbTextCompare)
    If atProp = 0 Then
        trace = trace & " | property name not in xml"
        GoTo Cleanup
    End If

    Dim openTag As Long, closeTag As Long
    openTag = InStr(atProp, xml, "<vt:lpwstr>")
    If openTag = 0 Then GoTo Cleanup
    openTag = openTag + Len("<vt:lpwstr>")
    closeTag = InStr(openTag, xml, "</vt:lpwstr>")
    If closeTag = 0 Then GoTo Cleanup

    PropertyOnDiskXlsx = Mid$(xml, openTag, closeTag - openTag)

Cleanup:
    On Error Resume Next
    fso.DeleteFolder work, True
    On Error GoTo 0
    Exit Function

Failed:
    trace = trace & " | ERROR " & Err.Number & ": " & Err.Description
    On Error Resume Next
    If Not fso Is Nothing Then fso.DeleteFolder work, True
    On Error GoTo 0
End Function
