Attribute VB_Name = "PreviewRealDeck"
Option Explicit

' One-off, read-only diagnostic -- NOT part of the shipped add-in (not in
' build_ppam.ps1's module list, not wired to CommandBarUI), same posture as
' VerifyRealDeck.bas next to it.
'
' Exists because RibbonUI.SyncPreview cannot be driven by automation: it ends
' in ShowSyncResult, an unconditional MsgBox, which blocks an automated run
' indefinitely (see AGENTS.md's Testing section). This is the headless twin --
' it resolves the deck's registry exactly as SyncPreview does, then calls the
' same RunSync.PreviewRoutineSync and RETURNS the report as a string instead of
' showing it.
'
' Read-only on both files, twice over: the deck is opened ReadOnly and the
' workbook is opened ReadOnly, AND everything underneath is already
' non-mutating by construction (PreviewRoutineSync suppresses all three of a
' routine sync's write sites). Both are closed without saving.
'
' The deck is opened WITH a window and explicitly activated, unlike
' VerifyRealDeck's windowless open: RunSync.GatherInstances reads
' Application.ActivePresentation rather than taking a presentation argument,
' so the deck under test has to actually be the active one for the preview to
' see its slides at all. A windowless open here would silently report zero
' instances -- and "zero instances" against a populated workbook is exactly the
' shape of a false mass-duplication warning.
Public Function PreviewRealDeck(deckPath As String) As String
    Dim report As String

    Dim pres As Object
    Set pres = Application.Presentations.Open(deckPath, msoTrue, msoFalse, msoTrue)
    pres.Windows(1).Activate

    Dim workbookPath As String
    workbookPath = DeckRegistry.GetWorkbookPath(pres)
    If workbookPath = "" Then
        PreviewRealDeck = "No paired workbook is registered on this deck -- nothing to preview."
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
        PreviewRealDeck = "This deck has no registered slide types -- nothing to preview."
        pres.Saved = msoTrue
        pres.Close
        Exit Function
    End If

    Dim xl As Object, wb As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    ' positional args: UpdateLinks:=0, ReadOnly:=True -- named args aren't
    ' reliable on a late-bound Object (VerifyRealDeck.bas hit this first)
    Set wb = xl.Workbooks.Open(workbookPath, 0, True)

    report = "Deck:      " & deckPath & vbCrLf & _
             "Workbook:  " & workbookPath & vbCrLf & _
             "Slides:    " & pres.Slides.count & vbCrLf & _
             "Types:     " & (hi - lo + 1) & vbCrLf & vbCrLf

    Dim i As Long
    For i = lo To hi
        Dim templateSld As Object
        Dim wsName As String
        If DeckRegistry.LookupType(pres, types(i), templateSld, wsName) Then
            report = report & RunSync.PreviewRoutineSync(wb.Worksheets(wsName), types(i)) & vbCrLf
        Else
            report = report & "SKIPPED " & types(i) & ": registered type's template slide no longer resolves." & vbCrLf
        End If
    Next i

    wb.Close False
    xl.Quit

    ' Belt and braces: the deck was opened ReadOnly and nothing above writes to
    ' it, but marking it Saved guarantees Close cannot raise a save prompt --
    ' a modal dialog would hang the whole automated run.
    pres.Saved = msoTrue
    pres.Close

    PreviewRealDeck = report
End Function
