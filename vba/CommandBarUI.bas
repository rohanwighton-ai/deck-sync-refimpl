Attribute VB_Name = "CommandBarUI"
Option Explicit

' ---------------------------------------------------------------------
' THE BUTTON CAPTIONS, DEFINED ONCE.
'
' Every user-facing message that tells someone to press a button used to spell
' its caption out as a literal -- 88 of them across seven modules. So a rename
' was 88 edits, and missing one left a message pointing at a button that does
' not exist. That is not hypothetical: Readiness offered 'Create Template
' Slide' as a remedy for a button the toolbar has never carried.
'
' The toolbar builder owns the vocabulary; everything else refers to it. Same
' split as Kind and the Sources period list -- the code owns the words, the
' caller owns the choice. Renaming a button is now one edit here.
' ---------------------------------------------------------------------
' ONE CAPTION PER BUTTON, AND NO CAPTION WITHOUT ONE.
'
' This block held 19 CAP_* constants while AddButton was called twice. The other
' 17 named buttons removed on 2026-08-09 -- and they were not inert: Readiness
' built four remedies from them, so the tool told a person to press things that
' were not on the toolbar. Deleted 2026-08-14 rather than left "in case", since
' a caption constant with no button is a stale string waiting for a reader.
'
' CAP_REBUILD_SHEETS went the same way later that night, under the same rule and
' for the same reason: its button was removed, so the constant was a caption for
' a button nobody could press.
'
' The two dialog TITLES that used to live here ("Review Changes", "Apply
' Approved") are now STAGE_* -- they name a stage of the chain, not a button,
' and the CAP_ prefix is what made them read as buttons.
' THE TOOLBAR SPLITS BY ARTIFACT, NOT BY STEP. Rohan, 2026-08-14.
'
' One set of actions touches the WORKBOOK, one touches the DECK, and neither can
' trigger the other. The boundary is WHERE YOU STOP TYPING: button 1 gets your
' sheets ready, you write, button 2 takes what you wrote all the way to a slide.
'
' "1. Sync Now" was renamed because it did not say which artifact it changed, and
' it changed both -- which is how pressing a button to PUBLISH rebuilt the sheets
' first and wiped 43 approve ticks on 2026-08-14.
Public Const CAP_SET_UP_QUARTER As String = "1. Set up my quarter"
Public Const CAP_PUT_ON_SLIDES As String = "2. Put it on the slides"

' REVIEW IS ITS OWN ACTION, AND ITS CAPTION SAYS WHAT IT IS **NOT** FOR.
' Rohan: "clear what it is and isn't for, part of a sequence, or not." A reader
' who cannot tell whether a button writes to their deck will not press it, or
' will press it and be surprised -- both worse than four extra words.
Public Const CAP_REVIEW_ONLY As String = "Review changes (writes nothing)"

' DECK MEMBERSHIP -- which slides exist, against which rows the register holds.
'
' It does BOTH directions, so it is named for both. Rohan chose one membership
' button over two, and chose delete over hide for retirement (2026-08-15) --
' the register is the source of truth and last quarter's saved deck is the
' archive, so hiding would grow the deck forever to avoid a loss that is
' already covered.
'
' Adding and removing are asked SEPARATELY inside it. One "make the deck match"
' confirmation would buy consent for the destructive half using the safe half's
' reasoning.
' SPLIT FROM ONE "Add or retire slides" BUTTON, 2026-08-15, per Rohan: "can't we
' just declare intent at the start and run the appropriate half?"
'
' The old button computed both halves and asked about each in turn. That was
' built to stop one "make the deck match" confirmation buying consent for the
' destructive half -- a real risk, solved by asking TWICE rather than by never
' asking about what you did not come for. Adding a project meant answering "no"
' to a delete prompt every single time, which is precisely how the prompt that
' matters gets clicked past.
'
' Declaring intent kills that. What it must NOT do is hide the other direction,
' or the deck grows forever: each half REPORTS the other's count without asking,
' which costs nothing because the scan already computes both.
Public Const CAP_ADD_SLIDES As String = "Add missing slides"
Public Const CAP_RETIRE_SLIDES As String = "Retire slides with no row"

' FIELD DISCOVERY IS AN ACTIVITY, NOT A PRECONDITION -- given a button 2026-08-15.
'
' RibbonUI.bas:1549 gates the setup question on `Not hasTypes`, correctly: setting
' up a slide type IS once-ever. But DiscoverFields was reachable ONLY from inside
' that gate, so a CONFIGURED deck could never tag another field -- with 32 fields
' untagged on the real deck and waiting for exactly this. The gate is right; the
' entry point was missing. DiscoverFields is self-contained (active slide, paired
' workbook, writes a sheet, marks in memory) and safe to run at any time.
Public Const CAP_DISCOVER_FIELDS As String = "Tag fields on this slide"
Public Const CAP_REPOINT_WORKBOOK As String = "Change which workbook this deck uses"

' The standing formatting-consistency check (FormattingAudit.bas) -- built
' and tested 2026-08-19, wired here so it is reachable at all. Read-only,
' same "safe to run at any time" class as CAP_DISCOVER_FIELDS above.
Public Const CAP_CHECK_FORMATTING As String = "Check field formatting"

' MISSING ENTIRELY UNTIL 2026-08-16, and this is the SAME bug this file's own
' header describes fixing once already: "Readiness offered 'Create Template
' Slide' as a remedy for a button the toolbar has never carried." It had been
' reachable only as a one-time MsgBox prompt at the end of Bulk Onboard Type
' (BatchOnboardFlow.bas), on the assumption a type only ever gets ONE
' template, made once, right after its first onboarding -- true until
' Scenario 3 (per-letter colour templates) needed a SECOND and THIRD
' template added to a type that has been running for weeks. Rohan asked for
' this button directly, live, after being sent to press one that did not
' exist. The Bulk Onboard prompt stays -- this is an additional, repeatable
' entry point, not a replacement.
Public Const CAP_CREATE_TEMPLATE As String = "Create template slide"

Public Const STAGE_REVIEW_CHANGES As String = "Review Changes"
Public Const STAGE_APPLY_APPROVED As String = "Apply Approved"


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
' WHEN THIS ADD-IN WAS BUILT, written by build_ppam.ps1 at build time.
'
' Three times on 2026-08-10 Rohan pressed a button to check a fix and got the
' old behaviour, because the .ppam predated the commit -- and nothing on screen
' could tell him which build he was running. The question "is this fix in?" cost
' a round trip every time.
'
' NOT hand-maintained, deliberately. TOOLBAR_BUILD below is "40" and has been
' since build 40, which is exactly how a version constant goes stale: it depends
' on someone remembering. This one is stamped by the script that makes the file,
' so it cannot disagree with what is in it.
'
' A build that says (unbuilt) is running straight from source -- worth knowing.
Public Const BUILD_STAMP As String = "(unbuilt)"

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
' The bar the three buttons live on. BAR_SETUP/BAR_STEPS/BAR_CAREFUL are kept
' ONLY so HideToolbar still deletes them -- upgrading from build 40 would
' otherwise leave orphan bars on screen with dead buttons on them.
Private Const BAR_MAIN As String = "Deck Sync " & TOOLBAR_BUILD
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
    Set bar = NewBar(BAR_MAIN)

    ' ---------------------------------------------------------------------
    ' TWO BUTTONS. Was 16 across three bars.
    '
    ' Rohan applied his own boundary rule to the bar itself -- "a boundary earns
    ' its place only where a person has to do work or make a decision in the
    ' gap" (BatchOnboardFlow.bas:2886) -- and then asked whether two would do.
    ' They do, and the reason three felt necessary was a mistake: buttons were
    ' being treated as the safety mechanism. They are not. The CONFIRMATION is
    ' the consent gate -- the tick is a selection, the dialog is the consent --
    ' so "publish to the register" and "write to the slides" do not need separate
    ' buttons. They need separate STOPS, which a chain gives you anyway.
    '
    ' NOTHING IS UNREACHABLE. Every capability is called from the chain or
    ' offered by the step before it, and check_vba_static.py fails the build if
    ' that stops being true -- it caught exactly that regression when this was
    ' three buttons and the chains called the private Cores instead.
    '
    ' Button 2 is NOT called "Reset". Reset could mean clearing the marking,
    ' rebuilding the sheets, discarding the quarter's rows, or restoring the deck
    ' from backup -- four different consequences, one word, pressed four months
    ' after you last used the tool. It rebuilds the drafting sheets, and
    ' WriteDraftingSheet harvests drafts, notes, submit text and sources before
    ' clearing and restores them after, so typed work survives. Verified before
    ' this button was put one click away.
    ' ---------------------------------------------------------------------
    AddButton bar, CAP_SET_UP_QUARTER, "RibbonUI.SyncNowChain", 1004, _
        "Use to start a quarter, BEFORE you write: sets the period and gets your drafting sheets ready. Workbook only -- it cannot change a slide.", True
    AddButton bar, CAP_PUT_ON_SLIDES, "RibbonUI.PutItOnTheSlides", 1017, _
        "Use to put what you wrote onto the slides, AFTER you have finished writing: publishes every field you ticked, then shows each change and asks once.", True
    AddButton bar, CAP_REVIEW_ONLY, "RibbonUI.ReviewChanges", 1000, _
        "Use to read the register against your slides and tick what should change. It does NOT write to a slide -- button 2 does that."
    ' Tooltips open with "Use to " on purpose, and a test enforces it -- so the
    ' first thing a person reads is the verb, not the occasion. Broken by both
    ' of these on their first write, 2026-08-15, and caught by that test.
    AddButton bar, CAP_ADD_SLIDES, "RibbonUI.AddMissingSlides", 1959, _
        "Use to create a slide, copied from the template and tagged, for every register row that has none. It never deletes. It tells you if any slides have no row."
    AddButton bar, CAP_RETIRE_SLIDES, "RibbonUI.RetireSlides", 358, _
        "Use to DELETE every slide whose key the register no longer lists, after naming each one by index and key. It never creates. Last quarter's saved deck is where they still exist."
    AddButton bar, CAP_DISCOVER_FIELDS, "RibbonUI.DiscoverFieldsOnSlide", 1758, _
        "Use to tag fields on the slide you are looking at: writes a grid of its shapes into the workbook, you mark what to track, and it tags them. Safe to re-run -- existing marks are kept."
    AddButton bar, CAP_CHECK_FORMATTING, "RibbonUI.CheckFormatting", 472, _
        "Use to compare every field's formatting (font size, shape type) against every other real slide with the same field, and name whichever ones disagree. Writes nothing."

    ' ---------------------------------------------------------------------
    ' WHY THIS IS A BUTTON, AND WHY IT IS NOT THE "REBUILD MY SHEETS" CLASS
    ' KILLED DIRECTLY BELOW.
    '
    ' The pairing is a FACT ABOUT THIS DECK that only a person can supply --
    ' which workbook feeds it. It is not a repair for a defect, so exposing it
    ' does not stop a bug being reported; a deck legitimately changes workbook
    ' when the file moves, when the deck is copied, and on the first day at a
    ' new employer.
    '
    ' DraftingUI.RepointWorkbookUI already existed and was reachable from ONE
    ' place: the sync path, and only when GetWorkbookPath returned "". That is
    ' a deck paired with NOTHING. A deck paired with the WRONG workbook could
    ' not reach it at all -- and that is the case that actually occurs, because
    ' GetWorkbookPath returns the stored path unchanged whenever that path
    ' exists, so a COPIED deck keeps pointing at the original's register while
    ' its own sits unread beside it. Caught 2026-08-15 before a "test on the
    ' known-good snapshot" wrote to the live register. The function's own
    ' comment predicted it: "the first support call the moment the two are
    ' separated, with no self-service fix."
    ' ---------------------------------------------------------------------
    AddButton bar, CAP_REPOINT_WORKBOOK, "RibbonUI.ChangePairedWorkbook", 23, _
        "Use to see which workbook this deck is paired with, and point it at a different one. Needed whenever a deck and its register are separated -- a copy, a moved file, a new machine. Changes the pairing only; it writes nothing to a slide."
    AddButton bar, CAP_CREATE_TEMPLATE, "RibbonUI.CreateTemplateSlide", 942, _
        "Use to turn a real, already-onboarded slide into a hidden master template: you pick which one, it strips the fields to placeholders and hides it from the slideshow. Repeatable -- one per type per colour letter."

    ' "REBUILD MY SHEETS" IS GONE, 2026-08-14, and it is not coming back as a
    ' button. Its tooltip said "use this when a drafting sheet looks wrong" --
    ' which is not a feature, it is a DEFECT WITH INSTRUCTIONS ATTACHED. Rohan
    ' made the same call on the row-reorder workaround: "why are you having to
    ' move register rows manually? Worries me that the code won't work when it
    ' needs to."
    '
    ' A sheet that looks wrong is a bug to fix, and a repair button is how that
    ' bug stops being reported. Button 1 rebuilds the sheets anyway, so nothing
    ' is lost -- RefreshDraftingSheets keeps its caller and is not orphaned.

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
    ToolbarNames = Array(BAR_MAIN, BAR_SETUP, BAR_STEPS, BAR_CAREFUL)
End Function

' THE BARS THAT ACTUALLY EXIST, as opposed to the ones HideToolbar must delete.
'
' ToolbarNames had quietly become two things: the delete list (which must keep
' naming build 40's three bars so an upgrade leaves no orphan) and the list of
' live bars. Anything iterating it to read controls asked PowerPoint for a bar
' that was never created and got "Invalid procedure call" -- which surfaced as
' two tests ERRORING rather than failing, plus a third erroring downstream off
' the same left-behind state. One word, two jobs.
Public Function ActiveToolbarNames() As Variant
    ActiveToolbarNames = Array(BAR_MAIN)
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
