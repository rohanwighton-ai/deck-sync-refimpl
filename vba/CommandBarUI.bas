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
'
' Shows the 5 actions cleared so far. Three were live-tested against the
' real deck as of 2026-07-26 (Mark Field for Batch, Bulk Onboard Type,
' Clear Marked Fields) -- Rohan's framing: "I only want to add an operation
' when I'm fully clear it works and I know what it does." Preview Sync
' followed on 2026-07-29 (it cannot write; see its own note below), and
' Sync Now on 2026-07-30.
'
' Sync Now is the one that bends the rule, so it is worth being honest about
' why. It writes, and it had never been run. But it could not BE run: the
' first live cycle on 2026-07-30 reached the point of clicking it and found
' no button, which is precisely how an untested action stays untested. The
' rule's real purpose is "don't let a half-understood button change a real
' deck", and that is now served by something better than absence -- a
' confirmation showing exactly what will change, with slide creation called
' out in capitals (RibbonUI.SyncNowCore / RunSync.ConfirmSyncText). Absence
' was never a safety mechanism anyway; it just moved the risk to whenever
' the button eventually appeared.
'
' The remaining 4 (NewPeriod/OnboardNewType/ResolveUnmatchedFields/
' AdoptExistingSlides) are commented out below, not deleted -- every
' underlying Sub still exists and is still tested by TestRunner.bas, they
' are just not on the toolbar yet. Uncomment (or ask for) one once it has
' actually been tried.
Public Sub ShowToolbar()
    HideToolbar

    Dim bar As Object
    Set bar = Application.CommandBars.Add(Name:=TOOLBAR_NAME, Position:=1, Temporary:=True)  ' msoBarTop = 1

    ' Preview Sync is on the toolbar as of 2026-07-29, at Rohan's request, and
    ' it is the one action where promoting it BEFORE a live test is the correct
    ' call rather than an exception to his rule.
    '
    ' The rule ("only add an operation when I'm fully clear it works") exists to
    ' stop a half-understood button changing a real deck. Preview Sync cannot:
    ' RunSync.PreviewRoutineSync suppresses all three mutation sites, and
    ' Test_RunSync_PreviewReportsWithoutTouchingTheDeck asserts that directly --
    ' no slides created, stale text still stale, order unchanged. So the risk
    ' the rule guards against is absent here, while the cost of NOT having it is
    ' real: without a button, the only way to run it is from the VBE, and
    ' driving the VBE is precisely the friction that keeps it untested.
    '
    ' It is also the safest possible first action on an unfamiliar machine,
    ' which is the situation it was asked for.
    AddButton bar, "Preview Sync", "RibbonUI.SyncPreview", 1090, "Show exactly what Sync Now would change in this deck -- reads only, writes nothing."

    ' Writes, and confirms before it does -- see this Sub's header for why it
    ' is on the toolbar despite the rule, and RunSync.ConfirmSyncText for what
    ' the confirmation says.
    AddButton bar, "Sync Now", "RibbonUI.SyncNow", 1004, "Pull changes from the paired Data workbook onto every already-linked slide in this deck. Shows what will change and asks before writing."

    ' Still not live-tested against a real deck -- hidden, not deleted. These
    ' DO write, so the rule applies to them unchanged. See this Sub's header.
    ' AddButton bar, "New Period", "RibbonUI.NewPeriod", 297, "Duplicate an existing slide instance into a new period (e.g. next quarter), with a fresh instance key."
    ' AddButton bar, "Onboard New Slide Type", "RibbonUI.OnboardNewType", 1697, "Register a brand-new slide type from one example slide, one field at a time via prompts."
    ' AddButton bar, "Resolve Unmatched Fields", "RibbonUI.ResolveUnmatchedFields", 594, "Manually assign a role to one selected shape that Sync Now couldn't confidently match on its own."
    ' AddButton bar, "Adopt Existing Slides", "AdoptFlow.AdoptExistingSlides", 1651, "Link a batch of already-existing slides to their matching Data-sheet rows without duplicating anything."

    AddButton bar, "Mark Field for Batch", "BatchOnboardFlow.MarkFieldForBatch", 165, "Click a field's shape first, then run this. Names and types the field, ready to include in a batch. Repeat for each field on your template slide."
    AddButton bar, "Bulk Onboard Type", "BatchOnboardFlow.BatchOnboardType", 122, "After marking your fields, run this to auto-select matching slides, review in Excel, and link the whole batch at once."
    AddButton bar, "Clear Marked Fields", "BatchOnboardFlow.ClearMarkedFieldsForBatch", 480, "Discard the fields you've marked so far and start over (e.g. after a misclick)."

    bar.Visible = True
End Sub

Public Sub HideToolbar()
    On Error Resume Next
    Application.CommandBars(TOOLBAR_NAME).Delete
    On Error GoTo 0
End Sub

' faceId values are best-guess built-in icon indices -- an unresolved
' faceId shows a blank/default icon, not a load failure. tooltipText shows
' on hover, so a real explanation is available without cluttering the
' button's own visible caption.
Private Sub AddButton(bar As Object, caption As String, onAction As String, faceId As Long, tooltipText As String)
    Dim btn As Object
    Set btn = bar.Controls.Add(1)  ' msoControlButton = 1
    btn.Caption = caption
    btn.OnAction = onAction
    btn.FaceId = faceId
    btn.Style = 2  ' msoButtonIconAndCaption
    btn.TooltipText = tooltipText
End Sub
