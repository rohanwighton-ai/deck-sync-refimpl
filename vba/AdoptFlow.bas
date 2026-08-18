Attribute VB_Name = "AdoptFlow"
Option Explicit

' The ribbon-facing entry point for specs/deck-adoption.md's bulk retroactive
' linking engine (vba/DeckAdoption.bas) -- deferred by that spec's own
' Non-goals ("the UI itself... a future pass"), built now that
' CommandBarUI.bas's toolbar exists to hang it on. Mirrors OnboardFlow.bas's
' relationship to Onboarding.bas exactly: this module is pure interactive
' glue (selection validation, the phase-gate InputBox review, instance-key
' prompts), DeckAdoption.bas's Plan/Commit split does every real decision --
' no new matching/adoption logic lives here.
'
' Per DeckAdoption.bas's own documented precondition, this flow requires an
' ALREADY-REGISTERED type (the greenfield "pick a template from scratch" path
' is explicitly out of scope for adoption -- use Setup B: Onboard Slides
' first, same boundary DeckAdoption.bas's header comment draws).

' Validates the current selection is >= 2 slides (a template plus at least
' one slide to adopt) and returns them in deck order (ascending SlideIndex)
' -- DeckAdoption.PlanAdoption's own documented precondition, since a Ctrl-
' click multi-select is not guaranteed to already be in that order. `outSlides`
' is left unallocated on any validation failure.
Public Function ValidateAdoptionSelection(sel As Object, ByRef outSlides() As Object) As String
    If sel.Type <> ppSelectionSlides Then
        ValidateAdoptionSelection = "Select the slides to adopt (plus one example already-onboarded slide is fine to include) first."
        Exit Function
    End If

    If sel.SlideRange.count < 1 Then
        ValidateAdoptionSelection = "Select at least one slide to adopt."
        Exit Function
    End If

    ' Populate the ByRef out-parameter directly (ReDim + per-element Set) --
    ' not a separate local array bulk-assigned at the end, which isn't a
    ' pattern this codebase uses anywhere else for a ByRef array parameter
    ' (PlanAdoption's own harvestedValues() out-param follows this same
    ' ReDim-then-Set-per-element convention). A first version of this
    ' function built a local `unsorted()` array and did
    ' `outSlides = unsorted` at the end -- real-Office run confirmed
    ' (2026-07-26) that leaves the caller's array unallocated ("Subscript
    ' out of range" reading it back), not merely a style inconsistency.
    Dim n As Long
    n = sel.SlideRange.count
    ReDim outSlides(1 To n)
    Dim i As Long
    For i = 1 To n
        Set outSlides(i) = sel.SlideRange(i)
    Next i

    ' Simple insertion sort by SlideIndex -- n is a UI selection, never large
    ' enough to need anything smarter.
    Dim j As Long
    For i = 2 To n
        Dim current As Object
        Set current = outSlides(i)
        j = i - 1
        Do While j >= 1
            If outSlides(j).SlideIndex <= current.SlideIndex Then Exit Do
            Set outSlides(j + 1) = outSlides(j)
            j = j - 1
        Loop
        Set outSlides(j + 1) = current
    Next i

    ValidateAdoptionSelection = ""
End Function

' Removes `templateSld` from `slides` if present (by SlideID, not object
' identity, since a fresh COM reference to "the same" slide is a distinct
' object) -- a user adopting a whole child deck in one Select All will
' naturally include the template slide itself, which must be excluded before
' calling PlanAdoption (it isn't "a slide to adopt", it's the reference).
' Order-preserving.
Public Function ExcludeTemplateSlide(slides() As Object, templateSld As Object) As Object()
    Dim lo As Long, hi As Long
    lo = LBound(slides): hi = UBound(slides)

    Dim result() As Object
    Dim n As Long
    n = 0

    Dim i As Long
    For i = lo To hi
        If slides(i).SlideID <> templateSld.SlideID Then
            n = n + 1
            ReDim Preserve result(1 To n)
            Set result(n) = slides(i)
        End If
    Next i

    ExcludeTemplateSlide = result
End Function

' Builds the phase-gate review summary shown before any write -- counts by
' disposition plus every non-"ready" slide's label and reason, so a human can
' actually see what's being skipped and why before confirming.
Public Function BuildAdoptionReviewSummary(plans() As AdoptionSlidePlan) As String
    Dim lo As Long, hi As Long, hasPlans As Boolean
    On Error Resume Next
    lo = LBound(plans): hi = UBound(plans)
    hasPlans = (Err.Number = 0)
    On Error GoTo 0

    If Not hasPlans Then
        BuildAdoptionReviewSummary = "No slides in scope."
        Exit Function
    End If

    Dim readyCount As Long, alreadyCount As Long, needsConfirmCount As Long, unclassifiedCount As Long
    Dim details As String

    Dim i As Long
    For i = lo To hi
        Select Case plans(i).Disposition
            Case "ready"
                readyCount = readyCount + 1
                details = details & "  READY  " & plans(i).SlideLabel & " -- " & plans(i).Reason & vbCrLf
            Case "already_linked"
                alreadyCount = alreadyCount + 1
            Case "needs_confirmation"
                needsConfirmCount = needsConfirmCount + 1
                details = details & "  NEEDS CONFIRMATION  " & plans(i).SlideLabel & " -- use Resolve Unmatched Fields, then re-run Adopt" & vbCrLf
            Case "unclassified"
                unclassifiedCount = unclassifiedCount + 1
                details = details & "  UNCLASSIFIED  " & plans(i).SlideLabel & " -- excluded, not forced in" & vbCrLf
        End Select
    Next i

    Dim s As String
    s = readyCount & " ready to link, " & alreadyCount & " already linked (skipped), " & _
        needsConfirmCount & " need confirmation first, " & unclassifiedCount & " unclassified." & vbCrLf & vbCrLf
    s = s & details
    BuildAdoptionReviewSummary = s
End Function

' Ribbon entry point. Requires the selection to include an already-registered
' type's template (via DeckRegistry) plus the slides to adopt -- greenfield
' (no template yet) is out of scope here, per DeckAdoption.bas's own
' precondition; use Setup B: Onboard Slides first.
' PRIVATE since 2026-08-14. It was Public for a caller that never existed --
' nothing outside this module has ever referenced it -- and a Public with no
' outside caller reads as a capability someone can reach.
Private Function PromptAdoptExistingSlides() As String
    Dim pres As Object
    Set pres = Application.ActivePresentation

    Dim types() As String
    types = DeckRegistry.ListRegisteredTypes(pres)
    Dim tLo As Long, tHi As Long, hasTypes As Boolean
    On Error Resume Next
    tLo = LBound(types): tHi = UBound(types)
    hasTypes = (Err.Number = 0)
    On Error GoTo 0

    If Not hasTypes Then
        PromptAdoptExistingSlides = "This deck has no registered slide types yet. Press '" & CommandBarUI.CAP_SET_UP_QUARTER & "' -- on a deck with no slide type it walks the setup, then come back and Adopt the rest."
        Exit Function
    End If

    Dim slides() As Object
    Dim selErr As String
    selErr = ValidateAdoptionSelection(Application.ActiveWindow.Selection, slides)
    If selErr <> "" Then
        PromptAdoptExistingSlides = selErr
        Exit Function
    End If

    Dim slideType As String
    slideType = RibbonUI.PickType(types, "Adopt Existing Slides -- Choose Type")
    If slideType = "" Then
        PromptAdoptExistingSlides = "Cancelled -- no type selected."
        Exit Function
    End If

    Dim templateSld As Object
    Dim wsName As String
    If Not DeckRegistry.LookupType(pres, slideType, templateSld, wsName) Then
        PromptAdoptExistingSlides = "Could not resolve type '" & slideType & "' (its template slide may have been deleted)."
        Exit Function
    End If

    Dim slidesToAdopt() As Object
    slidesToAdopt = ExcludeTemplateSlide(slides, templateSld)

    Dim lo As Long, hi As Long, hasSlides As Boolean
    On Error Resume Next
    lo = LBound(slidesToAdopt): hi = UBound(slidesToAdopt)
    hasSlides = (Err.Number = 0)
    On Error GoTo 0
    If Not hasSlides Then
        PromptAdoptExistingSlides = "Nothing to adopt -- the selection only contained the template slide itself."
        Exit Function
    End If

    Dim workbookPath As String
    workbookPath = DeckRegistry.GetWorkbookPath(pres)
    Dim wb As Object
    Set wb = WorkbookBridge.OpenOrGetWorkbook(workbookPath)
    If wb Is Nothing Then
        PromptAdoptExistingSlides = "Could not open the paired workbook at: " & workbookPath
        Exit Function
    End If

    ' IS THIS EVEN OUR REGISTER? Same check DraftingUI.Resolve now makes for
    ' its own callers, added 2026-08-19 -- adoption creates new rows in
    ' whatever workbook this resolves to, so a mismatch here means a
    ' stranger's register gains rows describing this deck's real projects.
    Dim pairNote As String
    pairNote = DeckRegistry.PairingProblem(pres, wb)
    If pairNote <> "" Then
        PromptAdoptExistingSlides = pairNote
        Exit Function
    End If

    Dim ws As Object
    Set ws = WorkbookBridge.GetOrAddWorksheet(wb, wsName)

    Dim harvestedValues() As Object
    Dim plans() As AdoptionSlidePlan
    plans = DeckAdoption.PlanAdoption(slidesToAdopt, templateSld, ws, harvestedValues)

    If MsgBox(BuildAdoptionReviewSummary(plans), vbYesNo + vbQuestion, "Adopt Existing Slides -- Review") <> vbYes Then
        PromptAdoptExistingSlides = "Cancelled at review -- nothing written."
        Exit Function
    End If

    ' Instance-key prompts: one per "ready" slide, in the same index range as
    ' plans()/slidesToAdopt()/harvestedValues() (CommitAdoption's own
    ' documented requirement). A blank answer means "" -- CommitAdoption
    ' itself treats that as not-yet-confirmed and skips the slide, never
    ' guessing, so this loop doesn't need its own separate skip handling.
    Dim confirmedInstanceKeys() As String
    ReDim confirmedInstanceKeys(lo To hi)

    Dim i As Long
    For i = lo To hi
        If plans(i).Disposition = "ready" Then
            Dim prompt As String
            prompt = "Instance key for " & plans(i).SlideLabel & ":" & vbCrLf
            Dim fieldName As Variant
            If Not harvestedValues(i) Is Nothing Then
                For Each fieldName In harvestedValues(i).Keys
                    prompt = prompt & "  " & fieldName & " = '" & harvestedValues(i)(fieldName) & "'" & vbCrLf
                Next fieldName
            End If
            If plans(i).MatchedKeylessRowId <> "" Then
                prompt = prompt & "(matches existing keyless Data-sheet row " & plans(i).MatchedKeylessRowId & " -- will link into it)" & vbCrLf
            End If
            prompt = prompt & "Leave blank to skip this slide this pass."

            confirmedInstanceKeys(i) = InputBox(prompt, "Adopt Existing Slides -- Instance Key")
        Else
            confirmedInstanceKeys(i) = ""
        End If
    Next i

    Dim result As AdoptionResult
    result = DeckAdoption.CommitAdoption(plans, slidesToAdopt, harvestedValues, confirmedInstanceKeys, slideType, templateSld, ws)

    Dim report As String
    report = "Linked: " & result.LinkedCount & vbCrLf & _
        "Already linked (skipped): " & result.AlreadyLinkedCount & vbCrLf & _
        "Excluded/unclassified/unconfirmed: " & result.ExcludedUnclassifiedCount & vbCrLf & _
        "FAILED verification: " & result.FailedVerificationCount

    If result.FailedVerificationCount > 0 Then
        Dim m As Long
        report = report & vbCrLf & "Failed slides (harvest bug this pass -- fix before re-running, not something a later sync corrects):"
        For m = 1 To result.FailedVerificationCount
            report = report & vbCrLf & "  " & result.FailedVerificationLabels(m)
        Next m
    End If

    PromptAdoptExistingSlides = report
End Function

' ENTRY POINT. Called by RibbonUI.OfferAdoptionForSelectedSlides, which is the
' only route to it -- it has NO toolbar button of its own, deliberately.
'
' This header said "Toolbar entry point" from the day it was written until
' 2026-08-14, and the button it named was deleted in the three-button split on
' that same day. The sentence stayed true-sounding and became false, which is
' why it is now written as the caller's NAME rather than as a button's caption:
' a caption can be deleted out from under a comment, a caller cannot without the
' compiler saying so.
'
' The real work is in AdoptExistingSlidesCore; this exists only to catch
' anything that escapes it.
'
' A WRAPPER rather than an inline "On Error GoTo" on purpose. In VBA,
' "On Error GoTo 0" disables the enabled handler for the whole procedure, and
' these bodies are full of "On Error Resume Next / On Error GoTo 0" pairs -- an
' inline handler would be switched off by the first of them and read as
' protection while providing none. Putting the handler in a separate frame
' means nothing inside the body can turn it off, now or after a later edit.
Public Sub AdoptExistingSlides()
    On Error GoTo Failed
    AdoptExistingSlidesCore
    Exit Sub
Failed:
    RibbonUI.ShowSyncResult "Adopt Existing Slides", RibbonUI.UnexpectedErrorText("Adopt Existing Slides", Err.Number, Err.Description, Err.Source)
End Sub

' Thin wrapper -- a plain, parameterless Sub (same shape as RibbonUI.bas's
' actions), reporting via the same shared ShowSyncResult rather than a bespoke
' dialog.
Private Sub AdoptExistingSlidesCore()
    Dim report As String
    report = PromptAdoptExistingSlides()
    If report <> "" Then
        RibbonUI.ShowSyncResult "Adopt Existing Slides", report
    End If
End Sub
