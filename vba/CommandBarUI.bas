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
' Shows the 7 actions cleared so far. Three were live-tested against the
' real deck as of 2026-07-26 (Mark Field for Batch, Bulk Onboard Type,
' Clear Marked Fields) -- Rohan's framing: "I only want to add an operation
' when I'm fully clear it works and I know what it does." Preview Sync
' followed on 2026-07-29 (it cannot write; see its own note below),
' Sync Now on 2026-07-30, and Create Template Slide on 2026-07-30 (evening).
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
'
' New Period is a deliberate hold rather than a queue position, as of
' 2026-07-30: DECISIONS.md's dated-row decision the same day retires it.
' Under one dated row per period, "new period" becomes "add rows to the Data
' sheet" and the slides arrive through the ordinary new_record path that now
' works -- so hardening it (guard + its own live cycle) is likely work that
' gets deleted at progression step 2. Left off the toolbar, where it is
' harmless, instead of invested in.
Public Sub ShowToolbar()
    HideToolbar

    Dim bar As Object
    Set bar = Application.CommandBars.Add(Name:=TOOLBAR_NAME, Position:=1, Temporary:=True)  ' msoBarTop = 1

    ' ---------------------------------------------------------------------
    ' ORDERED BY THE WORK, AND NUMBERED. Rohan, 2026-08-01: "The ribbon should
    ' be organised in line with the workflow and numbered in steps."
    '
    ' The previous order was accretion -- buttons sat wherever they were added,
    ' so Preview Sync came first, the setup steps came last, and nothing told a
    ' person what to press after what. On a toolbar this is not cosmetic: it is
    ' the only instruction the tool gives before somebody clicks something.
    '
    ' Two tracks, deliberately distinguished. SETUP runs once per slide type,
    ' ever. The numbered steps run every quarter. Numbering all of them 1..7
    ' would say "do these seven things each time", which is wrong and would
    ' send somebody back through onboarding they have already done.
    ' ---------------------------------------------------------------------

    ' --- SETUP: once per slide type -------------------------------------
    AddButton bar, "Setup A: Mark Fields", "BatchOnboardFlow.MarkFieldForBatch", 165, _
        "SETUP, once per slide type. Click a field's shape on your template slide, then run this. Repeat for each field. Text shapes only -- pictures, icons and bars are not supported yet."
    AddButton bar, "Setup B: Onboard Slides", "BatchOnboardFlow.BatchOnboardType", 122, _
        "SETUP, after marking. Finds the other slides of the same layout, shows every field in Excel for review, and links the whole batch at once."
    AddButton bar, "Setup C: Check Coverage", "RibbonUI.AuditFields", 1000, _
        "SETUP, after onboarding. Lists everything on a slide of this type that is NOT being tracked, ranked by how likely it is to be project data. Writes a checklist to a 'Template Audit' sheet; never changes the deck. Run it to find fields you meant to mark and missed."
    AddButton bar, "Setup D: Template Slide", "RibbonUI.CreateTemplateSlide", 26, _
        "SETUP, last, once per slide type. Adds a hidden master slide carrying this type's placeholder fields, so a NEW project gets a clean slide built from the template instead of a clone of somebody else's project. Needs the type onboarded first -- it cannot know which fields to place until then. Asks before writing."
    AddButton bar, "Setup: Clear Marks", "BatchOnboardFlow.ClearMarkedFieldsForBatch", 480, _
        "Discard every field marked so far and start the marking over. Cannot remove just one."

    ' --- THE QUARTERLY LOOP ---------------------------------------------
    AddButton bar, "1. Drafting Sheets", "DraftingUI.RefreshDraftingSheets", 1697, _
        "STEP 1. Build or refresh the drafting sheets -- one per prose field, every project on a row, current text beside a box for your new wording, Copilot's prompt in L2. Keeps everything you have already written. Writes nothing to the deck.", True
    AddButton bar, "2. Copy AI to Submit", "DraftingUI.CopyAiDraftsToSubmit", 122, _
        "STEP 2, optional. Copy the AI's drafts into the SUBMIT column, filling ONLY cells you left empty. Never overwrites your own words. Then edit column G and tick column I."
    AddButton bar, "3. Publish Drafts", "DraftingUI.PublishDraftsForField", 3, _
        "STEP 3. Show every ticked SUBMIT row, then on your say-so write them into the register as Approved. Touches no slide."
    AddButton bar, "4. Preview Sync", "RibbonUI.SyncPreview", 1090, _
        "STEP 4. Show everything the register would change in this deck. Reads only, writes nothing. The safest thing on this toolbar.", True
    AddButton bar, "5. Sync Now", "RibbonUI.SyncNow", 1004, _
        "STEP 5. Apply the register's changes to the slides. If every change is the same transformation repeated, it shows them and asks once; if anything needs reading one at a time, it sends you to the review sheet instead."

    ' --- The careful route to slides, when step 5 is too blunt -----------
    AddButton bar, "5a. Review Changes", "RibbonUI.ReviewChanges", 1090, _
        "INSTEAD OF STEP 5, when you want to read each change. Builds a 'Sync Review' sheet showing current vs proposed per slide. Writes nothing to the deck.", True
    AddButton bar, "5b. Apply Approved", "RibbonUI.ApplyApprovedChanges", 3, _
        "After 5a. Writes the changes you ticked onto the slides. Takes a backup first, re-checks each change against the slide, and skips anything that has moved since you approved it."
    AddButton bar, "Review + Approve All", "RibbonUI.ReviewChangesApproveAll", 463, _
        "SCRATCH COPIES ONLY: builds the review sheet and ticks every row without individual review. Still writes nothing until you run 5b."

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
Private Sub AddButton(bar As Object, caption As String, onAction As String, faceId As Long, tooltipText As String, _
                      Optional beginGroup As Boolean = False)
    Dim btn As Object
    Set btn = bar.Controls.Add(1)  ' msoControlButton = 1
    btn.Caption = caption
    btn.OnAction = onAction
    btn.FaceId = faceId
    btn.Style = 2  ' msoButtonIconAndCaption
    btn.TooltipText = tooltipText
    ' A separator bar before this button -- the CommandBars idiom for grouping.
    If beginGroup Then btn.BeginGroup = True
End Sub
