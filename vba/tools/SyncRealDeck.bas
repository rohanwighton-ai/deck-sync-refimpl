Attribute VB_Name = "SyncRealDeck"
Option Explicit

' Headless twin of RibbonUI.SyncNow, for the first real run against a real
' deck. NOT part of the shipped add-in -- same posture as VerifyRealDeck.bas
' and PreviewRealDeck.bas beside it.
'
' Exists for the same reason PreviewRealDeck does: SyncNow ends in
' ShowSyncResult, an unconditional MsgBox, which blocks an automated run
' indefinitely. This resolves the registry identically and calls the same
' RunSync.RunRoutineSync, returning the report instead of showing it.
'
' UNLIKE its two sibling diagnostics, this one is NOT read-only -- it is the
' real sync. RunRoutineSync writes corrected field text, duplicates the
' template slide for any Data row with no matching slide, and reorders slides
' to match row order. Run PreviewRealDeck first, every time, and read its
' "new slide(s) would be created" number before running this.
'
' `saveWhenDone` defaults to False: the deck is closed WITHOUT saving, so
' nothing reaches disk. A first run wants to prove the decision path works
' without persisting anything, and discarding is always the safe direction.
' The returned report says whether the presentation was left dirty, which is
' the honest signal for "did this actually change anything in memory" -- a
' sync the preview called a no-op should leave a clean presentation, and if it
' doesn't, that discrepancy is worth seeing rather than silently saving.
Public Function SyncRealDeck(deckPath As String, Optional saveWhenDone As Boolean = False) As String
    Dim report As String

    Dim pres As Object
    ' Read-WRITE (ReadOnly:=msoFalse) and WITH a window: RunSync.GatherInstances
    ' reads Application.ActivePresentation, so the deck must be the active one.
    Set pres = Application.Presentations.Open(deckPath, msoFalse, msoFalse, msoTrue)
    pres.Windows(1).Activate

    Dim wasSavedOnOpen As Boolean
    wasSavedOnOpen = pres.Saved

    Dim workbookPath As String
    workbookPath = DeckRegistry.GetWorkbookPath(pres)
    If workbookPath = "" Then
        SyncRealDeck = "No paired workbook is registered on this deck -- nothing to sync."
        pres.Saved = msoTrue
        pres.Close
        Exit Function
    End If

    Dim types() As String
    types = DeckRegistry.ListRegisteredTypes(pres)

    Dim lo As Long, hi As Long, hasTypes As Boolean
    On Error Resume Next
    lo = LBound(types)
    hi = UBound(types)
    hasTypes = (Err.Number = 0)
    On Error GoTo 0

    If Not hasTypes Then
        SyncRealDeck = "This deck has no registered slide types -- nothing to sync."
        pres.Saved = msoTrue
        pres.Close
        Exit Function
    End If

    Dim slidesBefore As Long
    slidesBefore = pres.Slides.count

    Dim xl As Object, wb As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    ' Opened READ-ONLY on purpose: a routine sync reads the workbook and writes
    ' only to the deck. If anything here ever tries to write to the workbook,
    ' this makes it fail loudly instead of quietly succeeding.
    Set wb = xl.Workbooks.Open(workbookPath, 0, True)

    report = "Deck:      " & deckPath & vbCrLf & _
             "Workbook:  " & workbookPath & vbCrLf & _
             "Slides in: " & slidesBefore & vbCrLf & _
             "Saved flag on open: " & wasSavedOnOpen & vbCrLf & vbCrLf

    Dim i As Long
    For i = lo To hi
        Dim templateSld As Object
        Dim wsName As String
        If DeckRegistry.LookupType(pres, types(i), templateSld, wsName) Then
            report = report & RunSync.RunRoutineSync(wb.Worksheets(wsName), types(i), templateSld) & vbCrLf
        Else
            report = report & "SKIPPED " & types(i) & ": registered type's template slide no longer resolves." & vbCrLf
        End If
    Next i

    wb.Close False
    xl.Quit

    report = report & vbCrLf & _
        "Slides out: " & pres.Slides.count & " (was " & slidesBefore & ")" & vbCrLf & _
        "Presentation dirty after sync: " & (Not pres.Saved) & vbCrLf

    If saveWhenDone Then
        pres.Save
        report = report & "SAVED to disk." & vbCrLf
    Else
        pres.Saved = msoTrue   ' discard: suppresses the close-time save prompt
        pres.Close
        report = report & "Closed WITHOUT saving -- nothing written to disk." & vbCrLf
    End If

    SyncRealDeck = report
End Function
