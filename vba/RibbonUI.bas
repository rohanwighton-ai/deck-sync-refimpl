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

    ' R9: duplicate identity tags, checked BEFORE planning and before any
    ' write. Placed here rather than inside the planner because the planner
    ' cannot report it usefully -- to PlanRoutineSync two slides sharing a key
    ' simply means one of them matches the row and the other does not exist,
    ' which is indistinguishable from a normal unmatched slide. The condition
    ' is only visible by looking across instances, which is what this does.
    '
    ' Warns rather than refuses: a duplicate key is a data-entry mistake in the
    ' deck, not a corruption, and the sync will still do something sensible to
    ' one of the two slides. Refusing outright would block a whole quarter's
    ' reporting over a fixable typo. But the consequence is stated plainly and
    ' the default is to stop.
    Dim dupType As Long
    For dupType = lo To hi
        Dim dupReport As DuplicateKeyReport
        dupReport = IdentityCheck.FindDuplicateKeys(types(dupType))
        If dupReport.HasDuplicates Then
            If MsgBox(IdentityCheck.DuplicateKeyWarningText(types(dupType), dupReport) & _
                      vbCrLf & vbCrLf & "Continue anyway?", _
                      vbYesNo + vbExclamation + vbDefaultButton2, "Sync Now") <> vbYes Then
                Exit Sub
            End If
        End If
    Next dupType

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

' ---------------------------------------------------------------------
' Audit Fields -- "what on this slide is the tool not tracking?"
' See TemplateAudit.bas for the reasoning. Read-only; the only write is
' the audit worksheet.
' ---------------------------------------------------------------------

Public Sub AuditFields()
    On Error GoTo Failed
    AuditFieldsCore
    Exit Sub
Failed:
    RibbonUI.ShowSyncResult "Audit Fields", RibbonUI.UnexpectedErrorText("Audit Fields", Err.Number, Err.Description, Err.Source)
End Sub

' Picks the subject slide by preference, and never requires a template.
'
' The fallback chain is the whole point (Rohan, 2026-07-30): this operation and
' field marking must each work whatever the other has or hasn't done, because
' decks arrive at different maturities. A deck that has never run Create
' Template Slide still needs to know which fields it is missing -- arguably
' more than a mature one does, since knowing the fields is what makes a
' template worth building. So:
'   1. the type's master template, if step 1 has been run
'   2. otherwise the registered slide (pre-step-1 decks: a real project slide)
'   3. otherwise the first instance of the type
' Each is a legitimate subject; only the ranking differs.
Private Sub AuditFieldsCore()
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
        MsgBox "This deck has no registered slide types yet -- use 'Onboard New Slide Type' first.", vbExclamation, "Audit Fields"
        Exit Sub
    End If

    Dim slideType As String
    slideType = InputBox(BuildTypePickerPrompt(types), "Audit Fields -- Choose Type")
    slideType = ResolveTypeAnswer(slideType, types)
    If slideType = "" Then Exit Sub

    Dim instances() As Object
    instances = RunSync.GatherInstances(slideType)

    Dim subjectSld As Object
    Dim subjectLabel As String

    Set subjectSld = TemplateSlide.FindTemplateFor(slideType)
    If Not subjectSld Is Nothing Then
        subjectLabel = "the master template (slide " & subjectSld.SlideIndex & ")"
    End If

    Dim wsName As String
    If subjectSld Is Nothing Then
        Dim registeredSld As Object
        If DeckRegistry.LookupType(pres, slideType, registeredSld, wsName) Then
            Set subjectSld = registeredSld
            subjectLabel = "slide " & subjectSld.SlideIndex & " (no master template yet -- this type's registered slide)"
        End If
    Else
        Dim ignoredSld As Object
        DeckRegistry.LookupType pres, slideType, ignoredSld, wsName
    End If

    If subjectSld Is Nothing Then
        Dim iLo As Long, iHi As Long, hasInstances As Boolean
        On Error Resume Next
        iLo = LBound(instances): iHi = UBound(instances)
        hasInstances = (Err.Number = 0)
        On Error GoTo 0
        If hasInstances Then
            Set subjectSld = instances(iLo)
            subjectLabel = "slide " & subjectSld.SlideIndex & " (first slide of this type)"
        End If
    End If

    If subjectSld Is Nothing Then
        MsgBox "Nothing to audit -- type '" & slideType & "' has no slides in this deck.", vbExclamation, "Audit Fields"
        Exit Sub
    End If

    ' The subject is excluded from its own comparison set. Without this, a
    ' pre-step-1 deck audits a real project slide against a list that includes
    ' that same slide, so every text scores at least 1 and nothing can ever
    ' read "on no other slide" -- the verdict that carries all the signal.
    Dim comparisons() As Object
    comparisons = ExcludeSlide(instances, subjectSld)

    Dim rowCount As Long
    Dim trackedFields As String
    Dim rows() As AuditRow
    rows = TemplateAudit.BuildAudit(subjectSld, comparisons, rowCount, trackedFields)

    Dim trackedCount As Long
    trackedCount = 0
    If trackedFields <> "" Then trackedCount = UBound(Split(trackedFields, "|")) + 1

    Dim likelyDataCount As Long
    Dim cLo As Long, cHi As Long, hasComparisons As Boolean
    On Error Resume Next
    cLo = LBound(comparisons): cHi = UBound(comparisons)
    hasComparisons = (Err.Number = 0)
    On Error GoTo 0
    Dim comparisonCount As Long
    If hasComparisons Then comparisonCount = cHi - cLo + 1

    ' Asks TemplateAudit rather than re-implementing the prefix match. The two
    ' were separate hand-written comparisons until 2026-07-30 and only this one
    ' was right -- the other had the literal's length wrong by one and silently
    ' mis-sorted the whole grid.
    Dim i As Long
    For i = 1 To rowCount
        If TemplateAudit.IsLikelyProjectData(rows(i).Verdict) Then likelyDataCount = likelyDataCount + 1
    Next i

    ' Write the grid, if there is a workbook to write it to. A deck with no
    ' paired workbook still gets the counts -- the audit reads the DECK, so
    ' refusing outright would withhold an answer it already has.
    Dim wroteGrid As Boolean
    Dim workbookPath As String
    workbookPath = DeckRegistry.GetWorkbookPath(pres)
    If workbookPath <> "" And rowCount > 0 Then
        Dim wb As Object
        Set wb = WorkbookBridge.OpenOrGetWorkbook(workbookPath)
        If Not wb Is Nothing Then
            Dim ws As Object
            Set ws = WorkbookBridge.GetOrAddWorksheet(wb, TemplateAudit.AUDIT_SHEET_NAME)
            TemplateAudit.WriteAuditGrid ws, rows, rowCount
            wroteGrid = True
        End If
    End If

    Dim report As String
    report = TemplateAudit.SummaryText(slideType, subjectLabel, trackedCount, rowCount, likelyDataCount, comparisonCount)
    If rowCount > 0 And Not wroteGrid Then
        report = report & vbCrLf & vbCrLf & "COULD NOT WRITE THE LIST: no paired workbook was reachable, so only the counts above are available."
    End If

    ShowSyncResult "Audit Fields", report
End Sub

' `slides` minus `dropSld`, matched by SlideID rather than object identity --
' same reasoning as AdoptFlow.ExcludeTemplateSlide, whose shape this mirrors
' (two references to one slide are not guaranteed to compare equal).
Public Function ExcludeSlide(slides() As Object, dropSld As Object) As Object()
    Dim result() As Object
    Dim n As Long
    n = 0

    Dim lo As Long, hi As Long, hasAny As Boolean
    On Error Resume Next
    lo = LBound(slides): hi = UBound(slides)
    hasAny = (Err.Number = 0)
    On Error GoTo 0
    If Not hasAny Then
        ExcludeSlide = result
        Exit Function
    End If

    Dim i As Long
    For i = lo To hi
        If slides(i).SlideID <> dropSld.SlideID Then
            n = n + 1
            ReDim Preserve result(1 To n)
            Set result(n) = slides(i)
        End If
    Next i

    ExcludeSlide = result
End Function

' ---------------------------------------------------------------------
' Create Template Slide -- specs/deck-compiler-concept.md progression
' step 1. Gives a type a master template slide that is never a real
' project, and re-points the type's registration at it, so new records
' stop being cloned from whichever real slide happened to be onboarded
' first. See TemplateSlide.bas's header for the hazard.
' ---------------------------------------------------------------------

' Toolbar entry point. The real work is in CreateTemplateSlideCore; this
' exists only to catch anything that escapes it.
'
' A WRAPPER rather than an inline "On Error GoTo" on purpose. In VBA,
' "On Error GoTo 0" disables the enabled handler for the whole procedure, and
' these bodies are full of "On Error Resume Next / On Error GoTo 0" pairs -- an
' inline handler would be switched off by the first of them and read as
' protection while providing none. Putting the handler in a separate frame
' means nothing inside the body can turn it off, now or after a later edit.
Public Sub CreateTemplateSlide()
    On Error GoTo Failed
    CreateTemplateSlideCore
    Exit Sub
Failed:
    RibbonUI.ShowSyncResult "Create Template Slide", RibbonUI.UnexpectedErrorText("Create Template Slide", Err.Number, Err.Description, Err.Source)
End Sub

Private Sub CreateTemplateSlideCore()
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
        MsgBox "This deck has no registered slide types yet -- use 'Onboard New Slide Type' first.", vbExclamation, "Create Template Slide"
        Exit Sub
    End If

    Dim slideType As String
    slideType = InputBox(BuildTypePickerPrompt(types), "Create Template Slide -- Choose Type")
    slideType = ResolveTypeAnswer(slideType, types)
    If slideType = "" Then Exit Sub

    ' Already has one: stop here rather than at MakeTemplateFrom's own guard,
    ' so the message can name the existing template's slide number. Both
    ' checks stay -- this one is for the human, that one is the invariant.
    Dim existing As Object
    Set existing = TemplateSlide.FindTemplateFor(slideType)
    If Not existing Is Nothing Then
        MsgBox "Type '" & slideType & "' already has a master template: slide " & existing.SlideIndex & "." & vbCrLf & vbCrLf & _
               "A type must have exactly one. Nothing was changed.", vbInformation, "Create Template Slide"
        Exit Sub
    End If

    Dim sourceSld As Object
    Dim wsName As String
    If Not DeckRegistry.LookupType(pres, slideType, sourceSld, wsName) Then
        MsgBox "Type '" & slideType & "' is registered but its slide no longer resolves (was it deleted?)." & vbCrLf & _
               "Re-onboard the type before creating its template.", vbExclamation, "Create Template Slide"
        Exit Sub
    End If

    ' Label the source by its instance key where it has one, falling back to
    ' the slide number -- the key is what the human recognises from the Data
    ' sheet, and "slide 3" is meaningless once the deck is reordered.
    Dim sourceInstance As SlideInstance
    sourceInstance = Resolve.ResolveSlideInstance(sourceSld)
    Dim sourceLabel As String
    sourceLabel = "slide " & sourceSld.SlideIndex
    If sourceInstance.HasInstanceKey Then sourceLabel = sourceInstance.InstanceKey & " (slide " & sourceSld.SlideIndex & ")"

    Dim sourceRoles() As String
    Dim sourceShapes() As Candidate
    sourceShapes = Onboarding.BuildTemplateFieldShapes(sourceSld, sourceRoles)
    Dim fLo As Long, fHi As Long, hasFields As Boolean
    On Error Resume Next
    fLo = LBound(sourceRoles): fHi = UBound(sourceRoles)
    hasFields = (Err.Number = 0)
    On Error GoTo 0
    Dim fieldCount As Long
    If hasFields Then fieldCount = fHi - fLo + 1

    If MsgBox(TemplateSlide.ConfirmTemplateText(slideType, sourceLabel, fieldCount), _
              vbYesNo + vbQuestion, "Create Template Slide") <> vbYes Then
        Exit Sub
    End If

    Dim mr As MakeTemplateResult
    mr = TemplateSlide.MakeTemplateFrom(sourceSld, slideType)

    Dim report As String
    If Not mr.Ok Then
        report = "FAILED to create a template for '" & slideType & "': " & mr.Reason
        ShowSyncResult "Create Template Slide", report
        Exit Sub
    End If

    ' Registration is the step that actually changes behaviour -- without it
    ' the template exists but nothing clones it, which is the quietest
    ' possible half-finished state. Done here rather than inside
    ' MakeTemplateFrom so that function stays testable with no registry.
    DeckRegistry.RegisterType pres, slideType, mr.NewSlide, wsName

    report = "Master template created for '" & slideType & "'." & vbCrLf & vbCrLf & _
        "    slide " & mr.NewSlide.SlideIndex & ", hidden from the slideshow" & vbCrLf & _
        "    " & mr.FieldCount & " field(s) set to placeholders" & vbCrLf & _
        "    new records will now be cloned from it, not from " & sourceLabel & vbCrLf & vbCrLf & _
        "It will not appear in Preview Sync or Sync Now reports -- a template" & vbCrLf & _
        "is not a record, so it is neither counted nor corrected." & vbCrLf & vbCrLf & _
        "Worth doing now: open it and clear anything the sync does not manage" & vbCrLf & _
        "(figures, chart data, notes, untagged text) that belonged to " & sourceLabel & "."
    ShowSyncResult "Create Template Slide", report
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
