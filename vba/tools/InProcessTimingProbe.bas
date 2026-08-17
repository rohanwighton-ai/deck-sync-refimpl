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

' THE REAL TEST OF THE NEGATIVE-CACHE FIX. Simulates production's actual
' shape: many DISTINCT (slideType, fieldId) misses get recorded via
' RecordAbsent (as InjectorFor's up-to-four-suffix walk does per field),
' THEN a fresh Lookup is timed. Under the OLD (pre-fix) design, every
' distinct miss grew the "Shape Address Book" SHEET by one row, so a later
' Lookup's FindBookRow scan got slower as more distinct misses piled up.
' Under the FIX, RecordAbsent never touches the sheet at all -- the sheet
' stays exactly as large as the real HIT count, regardless of how many
' misses this session has recorded. This is the number the earlier direct-
' Lookup-of-a-genuine-miss probe could NOT show, because it never called
' RecordAbsent at all.
Public Function ProbeNegativeCacheScaling(registerPath As String) As String
    Dim xl As Object, wb As Object
    Dim i As Long
    Dim t0 As Double, elapsed As Double
    Dim bookSheetBefore As Long, bookSheetAfter As Long
    Dim bookSheet As Object

    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Open(registerPath, False, True)

    ShapeAddressBook.SetActiveWorkbook wb

    Set bookSheet = wb.Worksheets("Shape Address Book")
    bookSheetBefore = bookSheet.Cells(bookSheet.Rows.Count, 2).End(-4162).Row

    ' Simulate 300 DISTINCT misses -- roughly what fable's audit estimated
    ' for a few real slide types' worth of absent fields x suffix variants.
    For i = 1 To 300
        ShapeAddressBook.RecordAbsent "ZZZ_SCALE_TYPE_" & (i Mod 5), "ZZZ_SCALE_FIELD_" & i
    Next i

    ' Sheet extent must not have grown -- re-read it fresh from the file,
    ' not from any in-memory row counter, so this genuinely proves nothing
    ' was written, not just that mNextRow wasn't incremented.
    bookSheetAfter = bookSheet.Cells(bookSheet.Rows.Count, 2).End(-4162).Row

    ' NOW time 100 lookups AFTER 300 distinct misses exist in memory --
    ' this is the number that would have degraded under the old design.
    t0 = Timer
    For i = 1 To 100
        Dim result As String
        result = ShapeAddressBook.Lookup("ZZZ_INPROC_PROBE_TYPE", "ZZZ_INPROC_PROBE_FIELD_" & i)
    Next i
    elapsed = Timer - t0

    ShapeAddressBook.SetActiveWorkbook Nothing
    wb.Close False
    xl.Quit

    ProbeNegativeCacheScaling = "SCALING sheetRowsBefore=" & bookSheetBefore & _
        " sheetRowsAfter300Misses=" & bookSheetAfter & _
        " elapsedSecFor100LookupsAfter=" & Format(elapsed, "0.00") & _
        " msPerCall=" & Format((elapsed * 1000) / 100, "0.0")
End Function
