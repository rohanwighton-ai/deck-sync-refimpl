VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' -----------------------------------------------------------------------
' THE PIN. See LOBBY-DESIGN.md section 4 for the full design.
' -----------------------------------------------------------------------
'
' Rohan, 2026-08-16: "it's pinned as part of authoring."
'
' WithEvents App As Excel.Application, EARLY-BOUND ON PURPOSE. VBA's
' WithEvents cannot sink events from a late-bound `Object` -- this is why
' this project now carries a reference to Excel's own object library
' (added programmatically by build_ppam.ps1 / run_vba_tests.ps1 /
' compile_check.ps1 / field_e2e.ps1, every build; see any of them for the
' GUID and why it has to be added fresh each time). This is the ONLY class
' module in the codebase and the only reason that reference exists.
'
' PRIVATE, WITH Watch()/IsWatching() AS THE PUBLIC INTERFACE, BY DESIGN --
' NOT because "Public WithEvents" itself is illegal. It was first suspected
' as the cause of a real "Expected: end of statement" compile error
' 2026-08-16, but that diagnosis was wrong: the actual cause was this
' entire .cls file importing as a plain Standard Module instead of a Class
' Module (see AGENTS.md's "A .cls file with LF-only line endings imports
' silently as the wrong component type" -- WithEvents cannot compile in
' EITHER visibility outside a real class module, which is exactly what
' made Public-vs-Private look like the variable when it never was). Kept
' Private anyway, on its own merits: it keeps the assignment path to a
' single, named entry point (Watch()) rather than letting any caller repoint
' the sink directly.
Private WithEvents mApp As Excel.Application

Public Sub Watch(target As Excel.Application)
    Set mApp = target
End Sub

Public Function IsWatching(target As Excel.Application) As Boolean
    If mApp Is Nothing Then Exit Function
    IsWatching = (ObjPtr(mApp) = ObjPtr(target))
End Function

' Fires for EVERY sheet change in EVERY workbook open on this Excel
' instance -- there is no way to scope WithEvents any narrower than the
' Application itself. FieldIdForSheet's own check is what keeps this cheap
' and safe: anything that is not a real drafting sheet name exits in one
' comparison, so editing an unrelated spreadsheet on the same machine costs
' nothing more than that one check.
Private Sub mApp_SheetChange(ByVal Sh As Object, ByVal Target As Excel.Range)
    On Error GoTo CleanExit

    Dim wb As Object
    Set wb = Sh.Parent

    Dim fieldId As String
    fieldId = DraftingLobby.FieldIdForSheet(wb, Sh.Name)
    If fieldId = "" Then Exit Sub

    ' wb.Application (or Sh.Application), never the bare `Application` --
    ' this class module compiles into the same PowerPoint-hosted project as
    ' everything else, and the bare keyword resolves to PowerPoint's
    ' Application there. Exactly the bug DraftingLobby.BuildLobbyFromScratch
    ' hit and was fixed for the same night this class was written.
    Dim approveCol As Object
    Set approveCol = Sh.Columns(Drafting.COL_D_APPROVED)

    Dim hit As Object
    Set hit = Sh.Application.Intersect(Target, approveCol)
    If hit Is Nothing Then Exit Sub

    Dim r As Object
    For Each r In hit.Rows
        Dim rowNum As Long
        rowNum = r.Row
        If rowNum >= Drafting.DRAFT_FIRST_ROW Then
            Dim entityKey As String
            entityKey = Trim(CStr(Sh.Cells(rowNum, Drafting.COL_D_ENTITY).Value))
            If entityKey <> "" Then
                Dim mark As String
                mark = CStr(Sh.Cells(rowNum, Drafting.COL_D_APPROVED).Value)
                If ReviewQueue.IsApprovalMark(mark) Then
                    DraftingLobby.PinToLobby wb, Sh.Name, rowNum, fieldId, entityKey
                Else
                    ' UN-ticking must clear a pin too, not just leave it stale --
                    ' the Lobby has to agree with the sheet in both directions,
                    ' same "never trust a maintained record over reality" rule
                    ' BuildLobbyFromScratch already enforces on a full rebuild.
                    DraftingLobby.ClearLobbyEntry wb, Sh.Name, fieldId, entityKey
                End If
            End If
        End If
    Next r

CleanExit:
    ' AN EVENT HANDLER MUST NEVER RAISE INTO EXCEL. An uncaught error here
    ' would surface as a modal dialog on an ordinary keystroke in the
    ' Approve column, which would be alarming out of all proportion to what
    ' this does. Swallowed on purpose -- BuildLobbyFromScratch is the safety
    ' net if a pin is ever missed because of it, per LOBBY-DESIGN.md's
    ' stated residual-risk reasoning.
End Sub
