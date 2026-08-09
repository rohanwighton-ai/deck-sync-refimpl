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

' THREE BARS, NOT ONE, because sixteen buttons do not fit a row.
'
' 2026-08-08: the last four -- Apply Approved, Review + Approve All, Repoint
' Workbook and half of Review Changes -- sat behind an overflow chevron, which
' means they could not be found by someone who did not already know they
' existed. Hover help does not reach a button you cannot see.
'
' Shortening the captions was the other option and was rejected: the caption is
' the only instruction this tool gives before somebody clicks. The split follows
' the division the code already makes -- setup runs once per slide type, the
' numbered steps run every period, and the careful route is what you use when
' step 4 is too blunt.
Private Const BAR_SETUP As String = "Deck Sync " & TOOLBAR_BUILD & " -- Setup"
Private Const BAR_STEPS As String = "Deck Sync " & TOOLBAR_BUILD & " -- Each period"
Private Const BAR_CAREFUL As String = "Deck Sync " & TOOLBAR_BUILD & " -- One at a time"

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
    Set bar = NewBar(BAR_SETUP)

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
        "Use to tag one field by clicking its shape. Repeat per field. Text shapes only."
    AddButton bar, "Setup A2: Discover Fields", "DiscoverUI.DiscoverFields", 1697, _
        "Use to tag every field on a slide at once, in one Excel grid. Marks nothing until you confirm."
    AddButton bar, "Setup B: Onboard Slides", "BatchOnboardFlow.BatchOnboardType", 122, _
        "Use to link the other slides of this layout to the register. You review them in Excel first."
    ' KEPT AS A BUTTON, unlike Template Slide, and the distinction is real.
    ' Both were merged into onboarding on 2026-08-01 -- then the test suite
    ' pointed out both had become unreachable except by re-running onboarding.
    ' Template Slide is genuinely once per slide type, so the offer at the end
    ' of onboarding covers it. "What am I not tracking?" is a RECURRING question
    ' -- asked again every time a field is added or a slide is redesigned -- and
    ' a read-only diagnostic you can only reach by re-running a setup step is
    ' one nobody will run. It is offered at onboarding AND available here.
    AddButton bar, "Setup C: Check Coverage", "RibbonUI.AuditFields", 1000, _
        "Use to see what on the slide is not being tracked. Writes a checklist; changes nothing."
    AddButton bar, "Setup: Clear Marks", "BatchOnboardFlow.ClearMarkedFieldsForBatch", 480, _
        "Use to discard all marking and start again. Cannot remove just one."

    ' --- THE QUARTERLY LOOP ---------------------------------------------
    Set bar = NewBar(BAR_STEPS)
    ' FIRST, because it is the only button that answers "what should I press?".
    ' Rebuilds the readiness sheet from the saved files and shows it.
    AddButton bar, "Where am I?", "RibbonUI.WhereAmI", 1000, _
        "Use to see what is set, what is missing, and what to press next. Writes one sheet; changes nothing else."
    AddButton bar, "0. Start a Quarter", "DraftingUI.StartQuarter", 297, _
        "Use to tell this deck which period it reports. Saves the deck and confirms it landed.", True
    AddButton bar, "0b. Roll Forward", "DraftingUI.RollForwardUI", 1017, _
        "Use to copy last period's rows into this one. Refuses if this period already has rows."
    AddButton bar, "1. Drafting Sheets", "DraftingUI.RefreshDraftingSheets", 1697, _
        "Use to build the sheets you write on. Your text goes in column D, the tick in column E.", True
    AddButton bar, "2. Copy AI to Submit", "DraftingUI.CopyAiDraftsToSubmit", 122, _
        "Use to move Copilot's drafts into the column that publishes. Never overwrites your own words."
    AddButton bar, "3. Publish & Preview", "DraftingUI.PublishDraftsForField", 3, _
        "Use to send your ticked rows to the register. No slide is touched."
    AddButton bar, "4. Sync Now", "RibbonUI.SyncNow", 1004, _
        "Use to put the register's text onto the slides. Shows you what will change first."

    ' --- The careful route to slides, when step 4 is too blunt -----------
    Set bar = NewBar(BAR_CAREFUL)
    ' PREVIEW SYNC LOST ITS BUTTON 2026-08-09 (Rohan: "pre sync review can be part
    ' of sync? option to cancel for user, one less button?"). Its output was a
    ' count plus a Run Log dump -- a LINE, not an action. The standing half of
    ' that question is now a readiness line ("Parity: deck and register agree");
    ' the per-run half is Sync Now's own confirmation, which you can cancel.
    ' The Sub stays: DraftingUI offers it at the end of Publish.
    AddButton bar, "Review Changes", "RibbonUI.ReviewChanges", 1090, _
        "Use to read each change one at a time, on a sheet. Writes nothing."
    AddButton bar, "Apply Approved", "RibbonUI.ApplyApprovedChanges", 3, _
        "Use to write only the changes you ticked. Takes a backup first."
    AddButton bar, "Review + Approve All", "RibbonUI.ReviewChangesApproveAll", 463, _
        "Use to tick everything without reading it. Scratch copies only."
    AddButton bar, "Repoint Workbook", "DraftingUI.RepointWorkbookUI", 23, _
        "Use to point this deck at a different workbook. Only needed if they got separated.", True

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
End Sub

' One bar, created visible. CommandBars.Add makes a bar HIDDEN; forgetting the
' Visible line is what made the whole toolbar disappear on 2026-08-08.
Private Function NewBar(barName As String) As Object
    Dim b As Object
    Set b = Application.CommandBars.Add(Name:=barName, Position:=1, Temporary:=True)  ' msoBarTop = 1
    b.Visible = True
    Set NewBar = b
End Function

' Every bar we own, so a caller can find or check them all.
Public Function ToolbarNames() As Variant
    ToolbarNames = Array(BAR_SETUP, BAR_STEPS, BAR_CAREFUL)
End Function

Public Sub HideToolbar()
    On Error Resume Next
    ' By our own exact names only. Deleting anything that merely looks like ours
    ' is how the collision above worked. The pre-split single bar is removed too,
    ' so upgrading from an older build does not leave an orphan.
    Application.CommandBars(TOOLBAR_NAME).Delete
    Dim n As Variant
    For Each n In ToolbarNames()
        Application.CommandBars(CStr(n)).Delete
    Next n
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
