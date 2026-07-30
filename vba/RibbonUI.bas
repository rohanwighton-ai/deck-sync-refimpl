Attribute VB_Name = "RibbonUI"
Option Explicit

' Action logic for specs/ribbon-ui.md's four buttons -- gather what the
' engine needs (via DeckRegistry lookups + WorkbookBridge), call an
' existing Sub, report the result. New Period's picker and Resolve
' Unmatched Fields' role picker are the only "new" pieces, and both are
' InputBox chains, not new sync/matching logic.
'
' Each action is a plain, parameterless Public Sub (SyncNow/NewPeriod/
' OnboardNewType/ResolveUnmatchedFields) rather than a ribbon-style
' `(control As IRibbonControl)` callback -- see CommandBarUI.bas's header
' comment for why: a real customUI14.xml ribbon turned out to be
' impossible to ship for a .ppam add-in (confirmed 2026-07-26 -- its loader
' rejects the package if it contains anything beyond its exact expected
' part set, ribbon or not), so CommandBarUI.bas's toolbar buttons drive
' these directly. Kept as zero-argument Subs (not folded into
' CommandBarUI.bas itself) so a future real ribbon-hosted add-in (the
' "Office.js kept open" option DECISIONS.md's 2026-07-25 entry already
' flags as a later move) can still wire a `(control As IRibbonControl)`
' wrapper straight back to these without touching this module's logic.
'
' Shared result reporting (ribbon-ui.md's "one shared result form... not a
' bespoke dialog per action") is deliberately NOT a UserForm here -- see
' OnboardFlow.bas's header comment for why (no proven .frm/.frx precedent
' yet). ShowSyncResult below is the shared *reporting* logic ribbon-ui.md
' asks for (one function, called from every action, not divergent MsgBox
' calls scattered per-button) -- just backed by MsgBox instead of a form.
' Upgrading the display mechanism later only touches this one function.

' ---------------------------------------------------------------------
' Sync Now
' ---------------------------------------------------------------------

' Toolbar entry point. The real work is in SyncNowCore; this exists only to
' catch anything that escapes it.
'
' A WRAPPER rather than an inline "On Error GoTo" on purpose. In VBA,
' "On Error GoTo 0" disables the enabled handler for the whole procedure, and
' these bodies are full of "On Error Resume Next / On Error GoTo 0" pairs -- an
' inline handler would be switched off by the first of them and read as
' protection while providing none. Putting the handler in a separate frame
' means nothing inside the body can turn it off, now or after a later edit.
Public Sub SyncNow()
    On Error GoTo Failed
    SyncNowCore
    Exit Sub
Failed:
    RibbonUI.ShowSyncResult "Sync Now", RibbonUI.UnexpectedErrorText("Sync Now", Err.Number, Err.Description, Err.Source)
End Sub

Private Sub SyncNowCore()
    Dim pres As Object
    Set pres = Application.ActivePresentation

    Dim workbookPath As String
    workbookPath = DeckRegistry.GetWorkbookPath(pres)
    If workbookPath = "" Then
        MsgBox "This deck has no paired workbook yet -- use 'Onboard New Slide Type' first.", vbExclamation, "Sync Now"
        Exit Sub
    End If

    Dim types() As String
    types = DeckRegistry.ListRegisteredTypes(pres)

    Dim lo As Long, hi As Long, hasTypes As Boolean
    On Error Resume Next
    lo = LBound(types): hi = UBound(types)
    hasTypes = (Err.Number = 0)
    On Error GoTo 0

    If Not hasTypes Then
        MsgBox "This deck has no registered slide types yet -- use 'Onboard New Slide Type' first.", vbExclamation, "Sync Now"
        Exit Sub
    End If

    Dim wb As Object
    Set wb = WorkbookBridge.OpenOrGetWorkbook(workbookPath)
    If wb Is Nothing Then
        MsgBox "Could not open the paired workbook at: " & workbookPath, vbCritical, "Sync Now"
        Exit Sub
    End If

    ' Refuse to sync out of Excel's unsaved buffer -- see WorkbookBridge.IsDirty
    ' for the live incident. Checked BEFORE planning, not after: the plan itself
    ' reads the sheet, so a plan built on unsaved data would already be the
    ' thing being prevented, and the confirmation would be describing values
    ' that exist in no file.
    If WorkbookBridge.IsDirty(wb) Then
        If MsgBox(WorkbookBridge.UnsavedWorkbookText(workbookPath), _
                  vbYesNo + vbExclamation, "Sync Now") <> vbYes Then
            Exit Sub
        End If
        wb.Save
    End If

    ' Plan every registered type BEFORE writing any of them, so the
    ' confirmation covers the whole deck rather than the first type only.
    ' Sync Now is the one toolbar action that can change a deck at scale, and
    ' until 2026-07-30 it was only kept safe by not being on the toolbar --
    ' see RunSync.PreviewRoutineSync's header for the 43-orphaned-row
    ' near-miss this guard exists for.
    Dim i As Long
    Dim totCorrect As Long, totCreate As Long, totFlag As Long
    Dim c1 As Long, c2 As Long, c3 As Long, c4 As Long
    For i = lo To hi
        Dim planTemplate As Object
        Dim planWsName As String
        If DeckRegistry.LookupType(pres, types(i), planTemplate, planWsName) Then
            RunSync.PlanCounts WorkbookBridge.GetOrAddWorksheet(wb, planWsName), types(i), c1, c2, c3, c4
            totCorrect = totCorrect + c2
            totCreate = totCreate + c3
            totFlag = totFlag + c4
        End If
    Next i

    If totCorrect = 0 And totCreate = 0 Then
        MsgBox "Nothing to sync -- every linked slide already matches the Data sheet.", _
               vbInformation, "Sync Now"
        Exit Sub
    End If

    If MsgBox(RunSync.ConfirmSyncText(totCorrect, totCreate, totFlag), _
              vbYesNo + vbQuestion, "Sync Now") <> vbYes Then
        Exit Sub
    End If

    Dim fullReport As String
    For i = lo To hi
        Dim templateSld As Object
        Dim wsName As String
        If DeckRegistry.LookupType(pres, types(i), templateSld, wsName) Then
            Dim ws As Object
            Set ws = WorkbookBridge.GetOrAddWorksheet(wb, wsName)
            fullReport = fullReport & RunSync.RunRoutineSync(ws, types(i), templateSld) & vbCrLf
        Else
            fullReport = fullReport & "SKIPPED " & types(i) & ": registered type's template slide no longer resolves (was it deleted?)" & vbCrLf
        End If
    Next i

    ShowSyncResult "Sync Now", fullReport
End Sub

' Toolbar entry point. The real work is in SyncPreviewCore; this exists only to
' catch anything that escapes it.
'
' A WRAPPER rather than an inline "On Error GoTo" on purpose. In VBA,
' "On Error GoTo 0" disables the enabled handler for the whole procedure, and
' these bodies are full of "On Error Resume Next / On Error GoTo 0" pairs -- an
' inline handler would be switched off by the first of them and read as
' protection while providing none. Putting the handler in a separate frame
' means nothing inside the body can turn it off, now or after a later edit.
Public Sub SyncPreview()
    On Error GoTo Failed
    SyncPreviewCore
    Exit Sub
Failed:
    RibbonUI.ShowSyncResult "Preview Sync", RibbonUI.UnexpectedErrorText("Preview Sync", Err.Number, Err.Description, Err.Source)
End Sub

' Read-only twin of SyncNow: identical resolution path (same registry lookups,
' same workbook, same worksheets), but every registered type is run through
' RunSync.PreviewRoutineSync instead of RunRoutineSync, so the deck is never
' touched. Deliberately shares SyncNow's structure line for line -- a preview
' that resolves its inputs differently from the real thing can disagree with it
' about what would happen, which defeats the point.
Private Sub SyncPreviewCore()
    Dim pres As Object
    Set pres = Application.ActivePresentation

    Dim workbookPath As String
    workbookPath = DeckRegistry.GetWorkbookPath(pres)
    If workbookPath = "" Then
        MsgBox "This deck has no paired workbook yet -- nothing to preview.", vbExclamation, "Preview Sync"
        Exit Sub
    End If

    Dim types() As String
    types = DeckRegistry.ListRegisteredTypes(pres)

    Dim lo As Long, hi As Long, hasTypes As Boolean
    On Error Resume Next
    lo = LBound(types): hi = UBound(types)
    hasTypes = (Err.Number = 0)
    On Error GoTo 0

    If Not hasTypes Then
        MsgBox "This deck has no registered slide types yet -- nothing to preview.", vbExclamation, "Preview Sync"
        Exit Sub
    End If

    Dim wb As Object
    Set wb = WorkbookBridge.OpenOrGetWorkbook(workbookPath)
    If wb Is Nothing Then
        MsgBox "Could not open the paired workbook at: " & workbookPath, vbCritical, "Preview Sync"
        Exit Sub
    End If

    Dim fullReport As String
    Dim i As Long
    For i = lo To hi
        Dim templateSld As Object
        Dim wsName As String
        If DeckRegistry.LookupType(pres, types(i), templateSld, wsName) Then
            Dim ws As Object
            Set ws = WorkbookBridge.GetOrAddWorksheet(wb, wsName)
            fullReport = fullReport & RunSync.PreviewRoutineSync(ws, types(i)) & vbCrLf
        Else
            fullReport = fullReport & "SKIPPED " & types(i) & ": registered type's template slide no longer resolves (was it deleted?)" & vbCrLf
        End If
    Next i

    ' Warns where Sync Now refuses. The preview writes nothing, so unsaved data
    ' cannot damage the deck here -- but a preview of values that exist in no
    ' file is still a preview of something that might never be synced, and the
    ' whole worth of this report is that it can be trusted. Stated at the TOP:
    ' a caveat below a long report is a caveat nobody reads.
    If WorkbookBridge.IsDirty(wb) Then
        fullReport = "NOTE: the Data workbook has unsaved changes, so this preview " & _
            "reflects what is on screen in Excel, not what is in the file." & vbCrLf & _
            "Save it before syncing." & vbCrLf & vbCrLf & fullReport
    End If

    ShowSyncResult "Preview Sync (nothing written)", fullReport
End Sub

' ---------------------------------------------------------------------
' New Period (case 2, explicit rollover) -- per-type-and-record, not
' global (run-sync.md Step 3). The new period's Data-sheet row must
' already exist (user enters it in Excel first, same as any other
' spreadsheet-first workflow this project assumes) -- this picker locates
' the source slide to duplicate and the already-entered row to inject from,
' it does not invent new field values.
' ---------------------------------------------------------------------

' Toolbar entry point. The real work is in NewPeriodCore; this exists only to
' catch anything that escapes it.
'
' A WRAPPER rather than an inline "On Error GoTo" on purpose. In VBA,
' "On Error GoTo 0" disables the enabled handler for the whole procedure, and
' these bodies are full of "On Error Resume Next / On Error GoTo 0" pairs -- an
' inline handler would be switched off by the first of them and read as
' protection while providing none. Putting the handler in a separate frame
' means nothing inside the body can turn it off, now or after a later edit.
Public Sub NewPeriod()
    On Error GoTo Failed
    NewPeriodCore
    Exit Sub
Failed:
    RibbonUI.ShowSyncResult "New Period", RibbonUI.UnexpectedErrorText("New Period", Err.Number, Err.Description, Err.Source)
End Sub

Private Sub NewPeriodCore()
    Dim pres As Object
    Set pres = Application.ActivePresentation

    Dim types() As String
    types = DeckRegistry.ListRegisteredTypes(pres)
    Dim lo As Long, hi As Long, hasTypes As Boolean
    On Error Resume Next
    lo = LBound(types): hi = UBound(types)
    hasTypes = (Err.Number = 0)
    On Error GoTo 0

    If Not hasTypes Then
        MsgBox "This deck has no registered slide types yet -- use 'Onboard New Slide Type' first.", vbExclamation, "New Period"
        Exit Sub
    End If

    Dim slideType As String
    slideType = InputBox(BuildTypePickerPrompt(types), "New Period -- Choose Type")
    slideType = ResolveTypeAnswer(slideType, types)
    If slideType = "" Then
        Exit Sub
    End If

    Dim templateSld As Object
    Dim wsName As String
    If Not DeckRegistry.LookupType(pres, slideType, templateSld, wsName) Then
        MsgBox "Could not resolve type '" & slideType & "'.", vbExclamation, "New Period"
        Exit Sub
    End If

    Dim instances() As Object
    instances = RunSync.GatherInstances(slideType)
    Dim iLo As Long, iHi As Long, hasInstances As Boolean
    On Error Resume Next
    iLo = LBound(instances): iHi = UBound(instances)
    hasInstances = (Err.Number = 0)
    On Error GoTo 0

    If Not hasInstances Then
        MsgBox "No existing instances of '" & slideType & "' to roll forward -- use 'Onboard New Slide Type' or Sync Now to create the first one.", vbExclamation, "New Period"
        Exit Sub
    End If

    Dim recordAnswer As String
    recordAnswer = InputBox(BuildRecordPickerPrompt(instances), "New Period -- Choose Record")
    Dim sourceSld As Object
    Set sourceSld = ResolveRecordAnswer(recordAnswer, instances)
    If sourceSld Is Nothing Then
        Exit Sub
    End If

    Dim newInstanceKey As String
    newInstanceKey = InputBox("Enter the new period's instance key (must already have a row in the Data sheet):", "New Period -- New Instance Key")
    If Trim(newInstanceKey) = "" Then
        Exit Sub
    End If

    Dim workbookPath As String
    workbookPath = DeckRegistry.GetWorkbookPath(pres)
    Dim wb As Object
    Set wb = WorkbookBridge.OpenOrGetWorkbook(workbookPath)
    If wb Is Nothing Then
        MsgBox "Could not open the paired workbook at: " & workbookPath, vbCritical, "New Period"
        Exit Sub
    End If
    Dim ws As Object
    Set ws = WorkbookBridge.GetOrAddWorksheet(wb, wsName)

    Dim sheet As Sheet
    sheet = ExcelOutput.ReadSheet(ws)

    If Not sheet.Rows.Exists(newInstanceKey) Then
        MsgBox "No Data-sheet row found for instance key '" & newInstanceKey & "' -- add the new period's row in Excel first, then retry.", vbExclamation, "New Period"
        Exit Sub
    End If

    Dim newValues As Object
    Set newValues = sheet.Rows(newInstanceKey)

    Dim dr As DuplicateResult
    dr = RunSync.RunPeriodRollover(sourceSld, slideType, newInstanceKey, newValues)

    Dim report As String
    If dr.Ok Then
        report = "Created new period '" & newInstanceKey & "' for type '" & slideType & "'."
        If dr.MissingFieldCount > 0 Then
            Dim m As Long
            report = report & vbCrLf & "Missing " & dr.MissingFieldCount & " field(s):"
            For m = 1 To dr.MissingFieldCount
                report = report & " " & dr.MissingFields(m)
            Next m
        End If
    Else
        report = "FAILED to create new period '" & newInstanceKey & "': " & dr.Reason
    End If

    ShowSyncResult "New Period", report
End Sub

Public Function BuildTypePickerPrompt(types() As String) As String
    Dim s As String
    s = "Choose a slide type (enter the number or the name):" & vbCrLf
    Dim lo As Long, hi As Long, i As Long
    lo = LBound(types): hi = UBound(types)
    For i = lo To hi
        s = s & i & ") " & types(i) & vbCrLf
    Next i
    BuildTypePickerPrompt = s
End Function

' Same number-or-name convention ResolveFields.PickRoleFromList already
' established -- kept as a separate function here (rather than a shared
' generic picker) since the two operate on different array element types
' with no natural common signature in VBA without a Variant-typed generic,
' which this project avoids per its own established style.
Public Function ResolveTypeAnswer(answer As String, types() As String) As String
    If Trim(answer) = "" Then
        ResolveTypeAnswer = ""
        Exit Function
    End If

    Dim lo As Long, hi As Long
    lo = LBound(types): hi = UBound(types)

    If IsNumeric(answer) Then
        Dim idx As Long
        idx = CLng(answer)
        If idx >= lo And idx <= hi Then
            ResolveTypeAnswer = types(idx)
            Exit Function
        End If
    End If

    Dim i As Long
    For i = lo To hi
        If LCase(types(i)) = LCase(Trim(answer)) Then
            ResolveTypeAnswer = types(i)
            Exit Function
        End If
    Next i

    ResolveTypeAnswer = ""
End Function

Public Function BuildRecordPickerPrompt(instances() As Object) As String
    Dim s As String
    s = "Choose the record to roll forward (enter the number):" & vbCrLf
    Dim lo As Long, hi As Long, i As Long
    lo = LBound(instances): hi = UBound(instances)
    For i = lo To hi
        Dim inst As SlideInstance
        inst = Resolve.ResolveSlideInstance(instances(i))
        s = s & i & ") " & inst.InstanceKey & vbCrLf
    Next i
    BuildRecordPickerPrompt = s
End Function

Public Function ResolveRecordAnswer(answer As String, instances() As Object) As Object
    If Not IsNumeric(answer) Then
        Set ResolveRecordAnswer = Nothing
        Exit Function
    End If

    Dim lo As Long, hi As Long, idx As Long
    lo = LBound(instances): hi = UBound(instances)
    idx = CLng(answer)

    If idx >= lo And idx <= hi Then
        Set ResolveRecordAnswer = instances(idx)
    Else
        Set ResolveRecordAnswer = Nothing
    End If
End Function

' ---------------------------------------------------------------------
' Onboard New Slide Type / Resolve Unmatched Fields -- thin wrappers over
' OnboardFlow.bas / ResolveFields.bas, the only new pieces being the
' DeckRegistry lookup Resolve Unmatched Fields needs to find its template
' (ribbon-ui.md's spec text assumed a caller would supply templateSld;
' DeckRegistry is that caller now).
' ---------------------------------------------------------------------

' Toolbar entry point. The real work is in OnboardNewTypeCore; this exists only to
' catch anything that escapes it.
'
' A WRAPPER rather than an inline "On Error GoTo" on purpose. In VBA,
' "On Error GoTo 0" disables the enabled handler for the whole procedure, and
' these bodies are full of "On Error Resume Next / On Error GoTo 0" pairs -- an
' inline handler would be switched off by the first of them and read as
' protection while providing none. Putting the handler in a separate frame
' means nothing inside the body can turn it off, now or after a later edit.
Public Sub OnboardNewType()
    On Error GoTo Failed
    OnboardNewTypeCore
    Exit Sub
Failed:
    RibbonUI.ShowSyncResult "Onboard New Slide Type", RibbonUI.UnexpectedErrorText("Onboard New Slide Type", Err.Number, Err.Description, Err.Source)
End Sub

Private Sub OnboardNewTypeCore()
    Dim report As String
    report = OnboardFlow.PromptOnboardNewSlideType()
    If report <> "" Then
        ShowSyncResult "Onboard New Slide Type", report
    End If
End Sub

' Toolbar entry point. The real work is in ResolveUnmatchedFieldsCore; this exists only to
' catch anything that escapes it.
'
' A WRAPPER rather than an inline "On Error GoTo" on purpose. In VBA,
' "On Error GoTo 0" disables the enabled handler for the whole procedure, and
' these bodies are full of "On Error Resume Next / On Error GoTo 0" pairs -- an
' inline handler would be switched off by the first of them and read as
' protection while providing none. Putting the handler in a separate frame
' means nothing inside the body can turn it off, now or after a later edit.
Public Sub ResolveUnmatchedFields()
    On Error GoTo Failed
    ResolveUnmatchedFieldsCore
    Exit Sub
Failed:
    RibbonUI.ShowSyncResult "Resolve Unmatched Fields", RibbonUI.UnexpectedErrorText("Resolve Unmatched Fields", Err.Number, Err.Description, Err.Source)
End Sub

Private Sub ResolveUnmatchedFieldsCore()
    Dim pres As Object
    Set pres = Application.ActivePresentation

    Dim sel As Object
    Set sel = Application.ActiveWindow.Selection
    If sel.Type <> ppSelectionShapes Or sel.ShapeRange.count <> 1 Then
        MsgBox "Select exactly one shape on the slide first.", vbExclamation, "Resolve Unmatched Fields"
        Exit Sub
    End If

    ' .Parent is the containing Slide only for a top-level shape -- a shape
    ' selected from inside a group would resolve to the GroupShape instead.
    ' Not handled here: field shapes are expected to be top-level per this
    ' project's existing discovery convention (Discovery.bas recurses into
    ' groups to find candidates, but a human directly clicking one they
    ' want to resolve is the common case this flow targets); flagged as a
    ' known gap rather than a silently wrong assumption.
    Dim sld As Object
    Set sld = sel.ShapeRange(1).Parent

    Dim instance As SlideInstance
    instance = Resolve.ResolveSlideInstance(sld)
    If Not instance.HasTypeTag Then
        MsgBox "This slide has no slide type tag -- Resolve Unmatched Fields only applies to a slide already matched to a type.", vbExclamation, "Resolve Unmatched Fields"
        Exit Sub
    End If

    Dim templateSld As Object
    Dim wsName As String
    If Not DeckRegistry.LookupType(pres, instance.TypeTag, templateSld, wsName) Then
        MsgBox "Could not find a registered template for type '" & instance.TypeTag & "'.", vbExclamation, "Resolve Unmatched Fields"
        Exit Sub
    End If

    Dim result As String
    result = ResolveFields.PromptResolveUnmatchedField(templateSld)
    MsgBox result, vbInformation, "Resolve Unmatched Fields"
End Sub

' ---------------------------------------------------------------------
' Shared result reporting -- ribbon-ui.md's "one shared result form...
' reused after Sync Now, New Period, and the onboarding verify step."
' ---------------------------------------------------------------------

Public Sub ShowSyncResult(title As String, report As String)
    MsgBox report, vbInformation, title
End Sub

' What the human sees when an action dies of something nobody anticipated.
'
' Every toolbar action is wrapped in a handler that ends here (see any entry
' point's own header for why a WRAPPER rather than an inline handler). Before
' 2026-07-29 there were none at all, so an unguarded raise anywhere below a
' button produced VBA's own Debug/End dialog -- which is not just ugly: End
' discards whatever the run had collected, and on 2026-07-29 that meant 45
' instance keys confirmed one prompt at a time.
'
' Pure, so the wording is testable without provoking a real error.
'
' Deliberately does NOT claim nothing was written. An error partway through a
' commit can leave real changes in the deck and the Data sheet, and a
' reassuring "no changes were made" would be a lie exactly when the human most
' needs the truth. Saying "check before re-running" is less comforting and
' actually correct.
Public Function UnexpectedErrorText(actionName As String, errNumber As Long, errDescription As String, errSource As String) As String
    Dim where As String
    where = Trim(errSource)
    If where = "" Then where = "an unidentified step"

    UnexpectedErrorText = _
        actionName & " stopped early -- something went wrong that this add-in didn't anticipate." & vbCrLf & vbCrLf & _
        "Error " & errNumber & ": " & errDescription & vbCrLf & _
        "Reported by: " & where & vbCrLf & vbCrLf & _
        "IMPORTANT: this run may have already changed your deck or its Data sheet before it stopped. Check both before running it again." & vbCrLf & vbCrLf & _
        "Nothing here needs the VBA editor -- if this keeps happening, the text above is the useful part to pass on."
End Function
