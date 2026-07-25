Attribute VB_Name = "CommandBarUI"
Option Explicit

' The actual shipped UI surface for specs/ribbon-ui.md, after a real
' customUI14.xml ribbon turned out to be impossible for a .ppam add-in
' (confirmed 2026-07-26 against real Office -- see customUI/customUI14.xml's
' header comment and SPIKE_NOTES_RibbonUI.md for the full account: the
' loader rejects the package outright if it contains anything beyond its
' exact expected part set, proven with a harmless unrelated dummy part
' producing the identical failure). CommandBars is the pre-Ribbon VBA UI
' mechanism -- still fully supported for exactly this reason: it's pure
' runtime code, not a package part, so it needs zero changes to the .ppam
' structure that's already proven to load. Modern (Ribbon-era) Office hosts
' a VBA-created CommandBar toolbar under its built-in "Add-Ins" tab
' ("Menu Commands"/"Custom Toolbars" group) -- less visually polished than
' a dedicated branded ribbon tab, but a real, clickable, testable surface,
' not a step back to the VBE/Immediate-Window experience.
'
' Auto_Open/Auto_Close are the classic legacy add-in lifecycle Subs --
' PowerPoint still runs Auto_Open automatically when an add-in is loaded
' (AddIns.Add + .Loaded = True) and Auto_Close when unloaded, confirmed
' 2026-07-26 against a real loaded add-in in this pass. ShowToolbar is a
' manual fallback entry point (same "ManualSmokeTest" convention every
' other module uses) in case Auto_Open didn't fire, e.g. code added to an
' add-in that was already loaded before this module existed.

Private Const TOOLBAR_NAME As String = "Deck Sync"

Public Sub Auto_Open()
    ShowToolbar
End Sub

Public Sub Auto_Close()
    HideToolbar
End Sub

' Idempotent: deletes any stale toolbar of the same name first (a prior
' unclean shutdown can leave one behind, same "clean slate" posture
' run_vba_tests.ps1's own process-cleanup step already takes), then builds
' it fresh. Safe to call directly from the VBE if Auto_Open didn't fire.
Public Sub ShowToolbar()
    HideToolbar

    Dim bar As Object
    Set bar = Application.CommandBars.Add(Name:=TOOLBAR_NAME, Position:=1, Temporary:=True)  ' msoBarTop = 1

    AddButton bar, "Sync Now", "RibbonUI.SyncNow", 1004  ' faceId: refresh-like icon
    AddButton bar, "New Period", "RibbonUI.NewPeriod", 297  ' faceId: insert-row-like icon
    AddButton bar, "Onboard New Slide Type", "RibbonUI.OnboardNewType", 1697  ' faceId: new-item-like icon
    AddButton bar, "Resolve Unmatched Fields", "RibbonUI.ResolveUnmatchedFields", 594  ' faceId: find-like icon

    bar.Visible = True
End Sub

Public Sub HideToolbar()
    On Error Resume Next
    Application.CommandBars(TOOLBAR_NAME).Delete
    On Error GoTo 0
End Sub

' faceId values are best-guess built-in icon indices, not verified against
' a live toolbar render yet -- an unresolved faceId shows a blank/default
' icon, not a load failure (same cosmetic-only risk customUI14.xml's
' imageMso values carried, per that file's own now-superseded comment).
Private Sub AddButton(bar As Object, caption As String, onAction As String, faceId As Long)
    Dim btn As Object
    Set btn = bar.Controls.Add(1)  ' msoControlButton = 1
    btn.Caption = caption
    btn.OnAction = onAction
    btn.FaceId = faceId
    btn.Style = 2  ' msoButtonIconAndCaption
End Sub
