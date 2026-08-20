Attribute VB_Name = "AuditRealDeck"
Option Explicit

' One-off diagnostic -- NOT part of the shipped add-in (not in build_ppam.ps1's
' module list, not wired to CommandBarUI), same posture as PreviewRealDeck.bas
' and VerifyRealDeck.bas beside it.
'
' Exists because RibbonUI.AuditFields cannot be driven by automation: it opens
' with an InputBox type picker and ends in ShowSyncResult's MsgBox, either of
' which blocks an automated run indefinitely (AGENTS.md's Testing section).
' This is the headless twin -- it makes the same subject choice the button
' makes, calls the same TemplateAudit.BuildAudit and WriteAuditGrid, and
' RETURNS the report as a string instead of showing it.
'
' Unlike PreviewRealDeck this one WRITES: the audit grid goes into the paired
' workbook, which is the whole point of it. The deck is still opened ReadOnly
' and closed unsaved -- the audit reads slides and never touches them. Run it
' against COPIES of a real deck and workbook, never the originals.
' `workbookPathOverride` exists for one specific and easily-missed reason.
' Copying a deck does NOT re-point it: DeckRegistry stores the workbook's
' ABSOLUTE path, and GetWorkbookPath returns that stored path whenever it
' still resolves. Its sibling-of-the-deck fallback only engages when the
' stored path is broken. So running this against a copy of a real deck, in a
' scratch folder, with a copy of the workbook beside it, would still open and
' WRITE THE AUDIT SHEET INTO THE ORIGINAL WORKBOOK -- the copy would be
' ignored, and the protection of having made one would be entirely illusory.
'
' Passing the path explicitly is the only reliable way to keep a diagnostic
' run off the real files. Leave it "" to use the deck's own registration.
Public Function AuditRealDeck(deckPath As String, Optional workbookPathOverride As String = "") As String
    Dim report As String

    Dim pres As Object
    ' WITH a window and activated, for the same reason PreviewRealDeck does it:
    ' RunSync.GatherInstances reads Application.ActivePresentation rather than
    ' taking a presentation argument, so a windowless open would silently
    ' report zero instances -- and zero instances means every text on the
    ' subject slide scores "on no other slide", i.e. the whole audit would come
    ' back reading LIKELY PROJECT DATA. A false answer, not an obvious failure.
    Set pres = Application.Presentations.Open(deckPath, msoTrue, msoFalse, msoTrue)
    pres.Windows(1).Activate

    Dim workbookPath As String
    If workbookPathOverride <> "" Then
        workbookPath = workbookPathOverride
    Else
        workbookPath = DeckRegistry.GetWorkbookPath(pres)
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
        AuditRealDeck = "This deck has no registered slide types -- nothing to audit."
        pres.Saved = msoTrue
        pres.Close
        Exit Function
    End If

    ' Routed through WorkbookBridge.OpenOrGetWorkbook, not a raw
    ' CreateObject+Open, for the same reason every writer in E2EField.bas
    ' was moved onto it 2026-08-20: a second, unguarded Excel instance
    ' opening a file someone already has open can silently "save"
    ' successfully into a copy that never reaches disk. This module's own
    ' header already says never to point it at Rohan's real files, but the
    ' guard costs nothing and the header comment is not itself a control.
    Dim wb As Object
    Dim wasOpen As Boolean
    Set wb = WorkbookBridge.OpenOrGetWorkbook(workbookPath, wasOpen)

    report = "Deck:      " & deckPath & vbCrLf & _
             "Workbook:  " & workbookPath & vbCrLf & _
             "Slides:    " & pres.Slides.count & vbCrLf & vbCrLf

    Dim i As Long
    For i = lo To hi
        Dim slideType As String
        slideType = types(i)

        Dim instances() As Object
        instances = RunSync.GatherInstances(slideType)

        ' Same preference chain as RibbonUI.AuditFieldsCore: master template,
        ' else the registered slide, else the first instance.
        Dim subjectSld As Object
        Dim subjectLabel As String
        Set subjectSld = TemplateSlide.FindTemplateFor(slideType)
        If Not subjectSld Is Nothing Then
            subjectLabel = "master template (slide " & subjectSld.SlideIndex & ")"
        Else
            Dim registeredSld As Object
            Dim wsNameTmp As String
            If DeckRegistry.LookupType(pres, slideType, registeredSld, wsNameTmp) Then
                Set subjectSld = registeredSld
                subjectLabel = "slide " & subjectSld.SlideIndex & " (no master template; registered slide)"
            End If
        End If

        If subjectSld Is Nothing Then
            report = report & "=== " & slideType & ": no slides of this type ===" & vbCrLf
            GoTo NextType
        End If

        Dim comparisons() As Object
        ' RibbonUI.ExcludeSlide is Private, so this module could never see it and
        ' AuditRealDeck has not compiled for as long as that has been true --
        ' undetected because nothing in this repo compiles vba/tools/ (fixed the
        ' same day, 2026-08-20). AdoptFlow.ExcludeTemplateSlide is the Public
        ' twin, same signature, and a dead-code audit the same evening flagged it
        ' as having no callers -- which is exactly what a caller pointed at the
        ' wrong module looks like from the other end.
        comparisons = AdoptFlow.ExcludeTemplateSlide(instances, subjectSld)

        Dim cLo As Long, cHi As Long, hasComp As Boolean
        On Error Resume Next
        cLo = LBound(comparisons): cHi = UBound(comparisons)
        hasComp = (Err.Number = 0)
        On Error GoTo 0
        Dim compCount As Long
        If hasComp Then compCount = cHi - cLo + 1

        Dim rowCount As Long
        Dim trackedFields As String
        Dim rows() As AuditRow
        rows = TemplateAudit.BuildAudit(subjectSld, comparisons, rowCount, trackedFields)

        Dim trackedCount As Long
        If trackedFields <> "" Then trackedCount = UBound(Split(trackedFields, "|")) + 1

        Dim likelyData As Long, checkCount As Long, chromeCount As Long
        Dim r As Long
        For r = 1 To rowCount
            If TemplateAudit.IsLikelyProjectData(rows(r).Verdict) Then
                likelyData = likelyData + 1
            ElseIf InStr(rows(r).Verdict, "CHECK") = 1 Then
                checkCount = checkCount + 1
            Else
                chromeCount = chromeCount + 1
            End If
        Next r

        report = report & "=== " & slideType & " ===" & vbCrLf & _
            "Subject:        " & subjectLabel & vbCrLf & _
            "Compared with:  " & compCount & " other slide(s)" & vbCrLf & _
            "Tracked fields: " & trackedCount & "  (" & trackedFields & ")" & vbCrLf & _
            "Untracked text: " & rowCount & vbCrLf & _
            "   likely project data: " & likelyData & vbCrLf & _
            "   check (partial match): " & checkCount & vbCrLf & _
            "   probably chrome:       " & chromeCount & vbCrLf & vbCrLf

        ' Every row, in the grid's own order. This is a diagnostic -- the
        ' whole list is the finding, and truncating it here would hide exactly
        ' the misclassifications worth looking at.
        For r = 1 To rowCount
            Dim preview As String
            preview = Replace(Replace(rows(r).Text, vbCr, " "), vbLf, " ")
            If Len(preview) > 70 Then preview = Left(preview, 70) & "..."
            report = report & "  [" & rows(r).SeenOn & "/" & rows(r).InstanceCount & "] " & _
                preview & vbCrLf
        Next r
        report = report & vbCrLf

        Dim ws As Object
        Set ws = WorkbookBridge.GetOrAddWorksheet(wb, TemplateAudit.AUDIT_SHEET_NAME)
        TemplateAudit.WriteAuditGrid ws, rows, rowCount

NextType:
    Next i

    wb.Save
    If Not wasOpen Then
        Dim app As Object
        Set app = wb.Application
        wb.Close False
        app.Quit
    End If

    pres.Saved = msoTrue
    pres.Close

    AuditRealDeck = report
End Function
