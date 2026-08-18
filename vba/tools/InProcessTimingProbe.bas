Attribute VB_Name = "InProcessTimingProbe"
Option Explicit

' THROWAWAY DIAGNOSTIC, not part of the production build list -- exists only
' to settle whether ShapeAddressBook's AT fix (mSheet/mNextRow caching, then
' the in-memory negative cache) behaves the same way under a real in-project
' call loop as it does when driven via cross-project Application.Run hops.
' Delete after use.
Public Function ProbeInProcessLookup(registerPath As String) As String
    Dim xl As Object, wb As Object
    Dim lastRow As Long
    Dim i As Long
    Dim t0 As Double, elapsed As Double
    Dim bookSheet As Object

    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Open(registerPath, False, True)

    ShapeAddressBook.SetActiveWorkbook wb

    Set bookSheet = wb.Worksheets("Shape Address Book")
    lastRow = bookSheet.Cells(bookSheet.Rows.Count, 2).End(-4162).Row

    t0 = Timer
    For i = 1 To 100
        Dim result As String
        result = ShapeAddressBook.Lookup("ZZZ_INPROC_PROBE_TYPE", "ZZZ_INPROC_PROBE_FIELD_" & i)
    Next i
    elapsed = Timer - t0

    ShapeAddressBook.SetActiveWorkbook Nothing
    wb.Close False
    xl.Quit

    ProbeInProcessLookup = "INPROC lastRow=" & lastRow & " elapsedSec=" & Format(elapsed, "0.00") & " msPerCall=" & Format((elapsed * 1000) / 100, "0.0")
End Function

' ProbeNegativeCacheScaling was deleted 2026-08-18 along with its subject:
' ShapeAddressBook.RecordAbsent (the per-TYPE miss cache) no longer exists
' -- confirmed absence now lives in the per-slide tag index. See
' ShapeAddressBook.mSlideTagIndex's header for the incident and design.
