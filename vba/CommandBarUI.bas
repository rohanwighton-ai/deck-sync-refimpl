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

' THE BUILD NUMBER IS IN THE NAME, ON PURPOSE.
'
' Every version used to name its toolbar "Deck Sync". Auto_Open deletes any bar
' of that name and rebuilds; Auto_Close deletes it. So whichever add-in loaded
' or unloaded LAST won, and unloading an old add-in removed the NEW one's
' toolbar by name -- silently, with the only symptom being buttons that looked
' wrong, which requires already knowing what they should say.
'
' It happened twice on 2026-08-01: addin28 alongside addin33 in the morning
' (four buttons), addin35 alongside addin36 in the evening (the pre-reorder
' set). I flagged it after the first and did not act, and it cost a second
' diagnosis.
'
' With the build in the name, two loaded add-ins produce TWO VISIBLE TOOLBARS
' instead of one quietly winning. The collision stops being invisible, which is
' the entire fix -- you can see you have two, and go untick one.
'
' BUMP THIS when building a new .ppam. It is deliberately manual: a version that
' derives itself from something automatic would drift out of step with the file
' the user actually loaded, which is the thing being disambiguated.
Private Const TOOLBAR_BUILD As String = "40"
Private Const TOOLBAR_NAME As String = "Deck Sync " & TOOLBAR_BUILD

' The toolbar's name, for anything that needs to find it.
'
' Exposed because three tests hardcoded the literal "Deck Sync" and broke the
' moment the build number went into the name -- a test that duplicates a
' constant is a second place to update, and it fails for a reason that has
' nothing to do with what it is testing.
Public Function ToolbarName() As String
    ToolbarName = TOOLBAR_NAME
End Function

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
        "Use to tag one field at a time by clicking its shape. SETUP, once per slide type: click the shape on your template slide, then run this, and repeat per field. Text shapes only -- pictures and icons are not supported yet."
    AddButton bar, "Setup A2: Discover Fields", "DiscoverUI.DiscoverFields", 1697, _
        "Use to tag every field on a slide at once, in one Excel grid. Lists every text shape in reading order -- tick and name the ones you want, instead of three dialogs per field. Marks nothing until you confirm. Setup A still works."
    AddButton bar, "Setup B: Onboard Slides", "BatchOnboardFlow.BatchOnboardType", 122, _
        "Use to link all the other slides of this layout to the register. Finds them, shows every field in Excel for review, links the batch -- then offers to check what is not tracked and to create the hidden template slide."
    ' KEPT AS A BUTTON, unlike Template Slide, and the distinction is real.
    ' Both were merged into onboarding on 2026-08-01 -- then the test suite
    ' pointed out both had become unreachable except by re-running onboarding.
    ' Template Slide is genuinely once per slide type, so the offer at the end
    ' of onboarding covers it. "What am I not tracking?" is a RECURRING question
    ' -- asked again every time a field is added or a slide is redesigned -- and
    ' a read-only diagnostic you can only reach by re-running a setup step is
    ' one nobody will run. It is offered at onboarding AND available here.
    AddButton bar, "Setup C: Check Coverage", "RibbonUI.AuditFields", 1000, _
        "Use to find out what on the slide is NOT being tracked. Ranks everything untracked by how likely it is to be project data, and writes a checklist to a 'Template Audit' sheet. Never changes the deck."
    AddButton bar, "Setup: Clear Marks", "BatchOnboardFlow.ClearMarkedFieldsForBatch", 480, _
        "Use to throw away the marking so far and start again. Discard every field marked so far and start the marking over. Cannot remove just one."

    ' --- THE QUARTERLY LOOP ---------------------------------------------
    AddButton bar, "0. Start a Quarter", "DraftingUI.StartQuarter", 297, _
        "Use to tell this deck which period it now reports. Everything downstream -- drafting sheets, publish and sync -- then reads that period's rows only. Saves the deck and confirms the value reached the file.", True
    AddButton bar, "0b. Roll Forward", "DraftingUI.RollForwardUI", 1017, _
        "Use to copy last period's register rows into the new period. One row per slide, stamped with the period this deck declares. Refuses if that period already has rows. Touches no slide."
    AddButton bar, "1. Drafting Sheets", "DraftingUI.RefreshDraftingSheets", 1697, _
        "Use to build the sheets you write this period's text on. One per prose field: every project a row, current text in column C, your wording in G, the tick in I. Keeps work you have already done.", True
    AddButton bar, "2. Copy AI to Submit", "DraftingUI.CopyAiDraftsToSubmit", 122, _
        "Use to move Copilot's drafts into the column that publishes. STEP 2, optional. Copy the AI's drafts into the SUBMIT column, filling ONLY cells you left empty. Never overwrites your own words. Then edit column G and tick column I."
    AddButton bar, "3. Publish & Preview", "DraftingUI.PublishDraftsForField", 3, _
        "Use to send your ticked rows to the register. STEP 3. Show every ticked SUBMIT row, then on your say-so write them into the register as Approved -- and offer to preview what that would change on the slides. Touches no slide by itself."
    AddButton bar, "4. Sync Now", "RibbonUI.SyncNow", 1004, _
        "Use to put the register's text onto the slides. STEP 4. Apply the register's changes to the slides. Shows exactly what it will change and asks first. If anything needs reading one at a time, it sends you to the review sheet instead."

    ' --- The careful route to slides, when step 4 is too blunt -----------
    AddButton bar, "Preview Sync", "RibbonUI.SyncPreview", 1090, _
        "Use to see what would change, without changing anything. Show everything the register would change in this deck. Reads only, writes nothing -- the safest thing on this toolbar, and the right first action on an unfamiliar machine.", True
    AddButton bar, "Review Changes", "RibbonUI.ReviewChanges", 1090, _
        "Use to read each proposed change one at a time. INSTEAD OF STEP 4, when you want to read each change. Builds a 'Sync Review' sheet showing current vs proposed per slide. Writes nothing to the deck."
    AddButton bar, "Apply Approved", "RibbonUI.ApplyApprovedChanges", 3, _
        "Use to write only the changes you ticked. After Review Changes. Writes the changes you ticked onto the slides. Takes a backup first, re-checks each change against the slide, and skips anything that has moved since you approved it."
    AddButton bar, "Review + Approve All", "RibbonUI.ReviewChangesApproveAll", 463, _
        "Use to tick everything at once, on scratch copies only. SCRATCH COPIES ONLY: builds the review sheet and ticks every row without individual review. Still writes nothing until you run Apply Approved."
    AddButton bar, "Repoint Workbook", "DraftingUI.RepointWorkbookUI", 23, _
        "Use to point this deck at a different Excel workbook. Point this deck at a different Excel workbook. Only needed if the deck and its workbook have been separated -- keep them in the same folder and the pairing repairs itself.", True

    ' CommandBars.Add CREATES THE BAR HIDDEN. Without this line the toolbar is
    ' built correctly, wired correctly, and invisible -- and PowerPoint shows no
    ' "Add-Ins" ribbon tab at all, because that tab only appears once a VISIBLE
    ' custom bar exists. So the symptom is "the add-in did nothing", which sends
    ' you looking at loading, macro security and trusted locations, none of which
    ' are involved. Found 2026-08-08 by attaching to the live instance: the bar
    ' was there with all 15 controls, Visible = False.
    '
    ' The three toolbar tests could not catch it. They assert the bar exists, has
    ' 15 controls, and that every button resolves to the right Sub -- all true of
    ' a bar nobody can see.
    bar.Visible = True
End Sub

Public Sub HideToolbar()
    On Error Resume Next
    ' By our own exact name only. Deleting anything that merely looks like ours
    ' is how the collision above worked.
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
    ' OFFICE CAPS TooltipText AT 255 CHARACTERS, and a longer one does not get
    ' trimmed -- it RAISES "Invalid procedure call or argument", which happens
    ' part-way through building the bar and so takes the WHOLE TOOLBAR with it.
    ' Adding one sentence to seven tooltips on 2026-08-08 did exactly that: the
    ' add-in loaded, Auto_Open ran, and no toolbar appeared. Same symptom as the
    ' invisible-toolbar bug earlier the same day, from an unrelated cause, which
    ' is precisely why it must not be possible.
    '
    ' Truncated here rather than trusted to every future caller: this is a UI
    ' that must not be destroyable by prose.
    If Len(tooltipText) > 255 Then
        btn.TooltipText = Left$(tooltipText, 252) & "..."
    Else
        btn.TooltipText = tooltipText
    End If
    ' A separator bar before this button -- the CommandBars idiom for grouping.
    If beginGroup Then btn.BeginGroup = True
End Sub
