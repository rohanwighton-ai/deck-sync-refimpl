Attribute VB_Name = "ResolveFields"
Option Explicit

' specs/ribbon-ui.md's "Resolve Unmatched Fields" flow: a human resolving a
' medium-confidence match found when a *subsequent* slide is checked against
' an already-established template (onboarding.md's matching case, distinct
' from first-time onboarding, which needs no code here at all -- see
' Onboarding.bas's own header comment).
'
' Per the spec, the only genuinely new code this flow needs is the
' shape-click capture and the role picker -- Onboarding.ConfirmFieldMatch
' already exists, is already tested (SPIKE_NOTES_Onboarding.md), and is
' called here unchanged; no new matching logic lives in this module.
'
' Split into a thin interactive entry point (PromptResolveUnmatchedField,
' which reads Application.ActiveWindow.Selection and drives an InputBox) and
' pure-logic helpers below it (ValidateSingleShapeSelection,
' BuildRolePickerPrompt, PickRoleFromList) -- mirrors DeckAdoption.bas's own
' posture of keeping engine/decision logic testable by taking objects as
' parameters rather than reading Application.ActiveWindow.Selection deep
' inside a function TestRunner.bas can't otherwise drive. The helpers ARE
' real, automatable tests here (TestRunner.bas runs against a visible,
' real PowerPoint window per run_vba_tests.ps1, so `shp.Select` + reading
' `Application.ActiveWindow.Selection` back is a genuine round trip, not a
' mock) -- only the InputBox prompt itself is manual-verification-only, per
' the same constraint every other UserForm/InputBox-driven interaction in
' this project has (TestRunner.bas cannot click through a live dialog).
'
' Deliberately InputBox-based, not a UserForm ListBox: this project has
' never authored a VBA UserForm (.frm/.frx pair) before, and unlike every
' other VBA gotcha logged in AGENTS.md, that one genuinely cannot be
' confirmed or refuted without a real VBE to export/import against -- this
' container has no Office install at all (see IMPLEMENTATION_PLAN.md's
' Priority 20 note on `powershell.exe` being unreachable here). Rather than
' hand-author an untested, unverifiable binary-adjacent file, this flow
' picks the lower-risk, already-proven mechanism (InputBox, used the same
' way SyncOperations/RunSync's own manual smoke tests already lean on
' MsgBox/InputBox-style interaction) and leaves the nicer ListBox picker as
' a future upgrade once a real Office session is available to build and
' verify a UserForm against. See specs/ribbon-ui.md's shared-result-form
' bullet, which hits this same open question for a different dialog.

' Resolves the currently selected shape's field role by prompting the user
' to choose from templateSld's defined roles, then calling
' Onboarding.ConfirmFieldMatch. Returns a human-readable outcome ("Assigned
' ..." on success, an explanation otherwise) rather than raising -- this is
' meant to be called directly off a ribbon button click with no caller in a
' position to catch a VBA error.
Public Function PromptResolveUnmatchedField(templateSld As Object) As String
    Dim shp As Object
    Dim selErr As String
    selErr = ValidateSingleShapeSelection(Application.ActiveWindow.Selection, shp)
    If selErr <> "" Then
        PromptResolveUnmatchedField = selErr
        Exit Function
    End If

    Dim roles() As String
    Dim fieldShapes() As Candidate
    fieldShapes = Onboarding.BuildTemplateFieldShapes(templateSld, roles)

    Dim answer As String
    answer = InputBox(BuildRolePickerPrompt(roles), "Resolve Unmatched Field")
    If answer = "" Then
        PromptResolveUnmatchedField = "Cancelled -- no role selected."
        Exit Function
    End If

    Dim role As String
    role = PickRoleFromList(answer, roles)
    If role = "" Then
        PromptResolveUnmatchedField = "'" & answer & "' is not a valid role number or name."
        Exit Function
    End If

    Onboarding.ConfirmFieldMatch shp, role
    PromptResolveUnmatchedField = "Assigned role '" & role & "' to shape '" & shp.Name & "'."
End Function

' ---------------------------------------------------------------------
' Pure-logic helpers -- exercised directly by TestRunner.bas
' ---------------------------------------------------------------------

' Confirms `sel` is exactly one selected shape (not zero, not a slide
' selection, not a text-editing selection, not multiple) -- ribbon-ui.md's
' mechanism is explicitly "the ambiguous shape" (singular), and
' Onboarding.ConfirmFieldMatch itself takes exactly one shape. Returns ""
' and sets `outShp` when valid; a human-readable error otherwise, leaving
' `outShp` unset.
Public Function ValidateSingleShapeSelection(sel As Object, ByRef outShp As Object) As String
    If sel.Type <> ppSelectionShapes Then
        ValidateSingleShapeSelection = "Select exactly one shape first."
        Exit Function
    End If

    If sel.ShapeRange.count <> 1 Then
        ValidateSingleShapeSelection = "Select exactly one shape (selected " & sel.ShapeRange.count & ")."
        Exit Function
    End If

    Set outShp = sel.ShapeRange(1)
    ValidateSingleShapeSelection = ""
End Function

' Builds the numbered-list prompt text shown in the InputBox -- an
' InputBox has no listbox, so the prompt enumerates every role with a
' 1-based number, and PickRoleFromList below accepts either that number or
' the role name typed directly.
Public Function BuildRolePickerPrompt(roles() As String) As String
    Dim lo As Long, hi As Long, hasAny As Boolean
    On Error Resume Next
    lo = LBound(roles)
    hi = UBound(roles)
    hasAny = (Err.Number = 0)
    On Error GoTo 0

    Dim s As String
    s = "Choose a role for the selected shape (enter the number or the name):" & vbCrLf

    If Not hasAny Then
        BuildRolePickerPrompt = s & "(template has no defined roles)"
        Exit Function
    End If

    Dim i As Long
    For i = lo To hi
        s = s & i & ") " & roles(i) & vbCrLf
    Next i
    BuildRolePickerPrompt = s
End Function

' Resolves a raw InputBox answer against `roles()`: accepts either a
' 1-based list number ("2") or the role name itself, matched
' case-insensitively (a typed name is a reasonable thing to expect once the
' prompt has already shown it). Returns "" (never a valid role name) if
' `answer` matches neither a number in range nor a known role, or if
' `roles()` is unallocated (a template with no defined roles).
Public Function PickRoleFromList(answer As String, roles() As String) As String
    Dim lo As Long, hi As Long, hasAny As Boolean
    On Error Resume Next
    lo = LBound(roles)
    hi = UBound(roles)
    hasAny = (Err.Number = 0)
    On Error GoTo 0

    If Not hasAny Then
        PickRoleFromList = ""
        Exit Function
    End If

    If IsNumeric(answer) Then
        Dim idx As Long
        idx = CLng(answer)
        If idx >= lo And idx <= hi Then
            PickRoleFromList = roles(idx)
            Exit Function
        End If
    End If

    Dim i As Long
    For i = lo To hi
        If LCase(roles(i)) = LCase(answer) Then
            PickRoleFromList = roles(i)
            Exit Function
        End If
    Next i

    PickRoleFromList = ""
End Function
