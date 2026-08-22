' Adds a real MS<n>_CALDATE textbox to every milestone slot on the given
' slide, positioned directly below that slot's own DATE label, measured
' from the DATE label's ACTUAL geometry on THIS slide (never a hardcoded
' absolute position) -- same "measured, not computed" convention
' MilestoneDevice.bas's bar/track already use, and the same technique
' add_pct_shapes.vbs proved (2026-08-22, same session). Positioned under
' the DATE label specifically, not the circle -- Rohan: "date could be
' added instead under each circle like we were going to with % deliverable
' complete", and that design was "as a smaller font separate label under
' the date label" (MILESTONE-PERCENTAGE-DESIGN.md). Hidden by default;
' DrawMilestones shows/writes it when a register MS<n>_CALDATE value
' exists -- NOT gated by achieved state (see COL_CALDATE's own header).
'
' Gist: this script bolts one hidden date label under each of the seven
' milestone circles on one slide, and is built to be safe to re-run on a
' slide where an earlier run died halfway -- it repairs leftovers instead
' of doubling them up.
'
' ============================ HARDENED 2026-08-22 ============================
' Running the first version against the real OneDrive-synced deck surfaced
' five real failure modes, each now handled here:
'
' 1. IDENTITY: Presentation.FullName on a fully OneDrive-synced file
'    sometimes reports the https://d.docs.live.net/... cloud alias instead
'    of the local path Open() was given (confirmed live: 36/36 slides hit
'    it and aborted safely). Open() itself guarantees the right file; the
'    check compares Presentation.Name (filename only) and exists purely to
'    catch a genuinely different file.
'
' 2. SAVE: Presentation.Save() intermittently raises "An error occurred
'    while PowerPoint was saving the file" on OneDrive-synced files even
'    when the LOCAL write landed (Presentation.Saved reads True right after
'    the error; the on-disk bytes were confirmed correct) -- the error is
'    most likely OneDrive's cloud-sync confirmation step, not the save.
'    So: on a Save error, trust Presentation.Saved. True -> proceed.
'    False -> wait 5s and retry Save() once on the SAME open presentation
'    (never reopen); one retry succeeded every time it was needed live.
'    Only a failed retry with Saved still False is fatal (exit 2, and the
'    presentation is deliberately LEFT OPEN so a human can save by hand).
'
' 3. QUIT: after Save+Close, POWERPNT.EXE sometimes lingers even with
'    Saved=True. Passive waiting never once resolved it (80s of polling,
'    live) -- so this script polls only briefly (~9s) after Quit(), then
'    logs and exits anyway rather than hanging the automation. Force-close
'    is the CALLER's job, and only ever after confirming Saved=True via
'    COM first; killing a POWERPNT that still holds unsaved work loses it.
'
' 4. BAD PATH = SILENT HANG, NOT AN ERROR: Presentations.Open on a
'    malformed/nonexistent path does not throw -- PowerPoint puts up a
'    modal "Sorry, we couldn't open..." dialog and the script blocks on it
'    FOREVER, with nothing in the log (found live 2026-08-22 when a
'    caller's quoting mangled the deckPath argument; diagnosed only by
'    UI-Automation-reading the dialog). So: the file's existence is
'    checked with FileSystemObject BEFORE Open, and DisplayAlerts is set
'    to ppAlertsNone so remaining alert-class dialogs error instead of
'    blocking. Callers: quote the deckPath argument carefully.
'
' 5. STRAYS/DUPLICATES: a run killed (or a Save that failed) between
'    AddTextbox and the regroup leaves the new shape STRAY -- right name,
'    top level of the slide, NOT in the device group. The old existence
'    check only looked inside GroupItems, so a re-run added a second shape:
'    real duplicate MS4_CALDATE (and MS1_CALDATE) on the real deck,
'    confirmed from the saved file's raw XML. Now every slot checks BOTH
'    GroupItems AND the slide's top-level Shapes; a lone stray is adopted
'    into the group (re-styled + regrouped, same code path as creation),
'    duplicates are deduped keeping a grouped copy over a stray one, and a
'    final per-slide self-check pass re-counts every name before Save and
'    repairs + reports loudly if this very run somehow left a duplicate.
'
' The device group is RE-ACQUIRED FRESH at the start of every slot's
' processing (and again in the self-check): after a member deletion or a
' regroup, a held group reference goes stale and the next use raises
' "Type mismatch" (confirmed live in dedupe_caldate.vbs, 2026-08-22).
'
' Regrouping the device to add a shape destroys the group's own name and
' role tag (confirmed empirically 2026-08-22) -- both are captured before
' and explicitly restored after every regroup.
'
' ONE SLIDE PER INVOCATION, DELIBERATELY -- chaining multiple slides'
' regroups in one PowerPoint session raised "Type mismatch" on the second
' slide's device-group lookup, reproducibly (see add_pct_shapes.vbs).
'
' Usage:
'   cscript //Nologo add_caldate_shapes.vbs <slideIndex> [deckPath]
'
' The optional deckPath override exists for testing against a SCRATCH COPY
' (the copy must keep the filename "3. Project Progress.pptx" or the
' identity check aborts). Without it, the real deck path below is used.
'
' Exit codes: 0 = done (file saved; PowerPoint may still be lingering --
' see log), 1 = argument/identity/device-group problem, nothing saved,
' 2 = save genuinely failed even after retry; presentation LEFT OPEN.
'
' WRAPPER IDIOM (the caller looping this over many slides), from WSL --
' this exact quoting is the tested one; looser quoting is what triggered
' failure mode 4's hang:
'   for n in 40 41 42 43 44 46 47; do
'     # 1. only ever force-close a leftover POWERPNT after confirming via
'     #    COM that its presentation's Saved=True (a GetActiveObject
'     #    one-liner) -- never blind;
'     powershell.exe -NoProfile -Command "& cscript.exe //Nologo 'C:\...\add_caldate_shapes.vbs' $n"
'     # 2. verify from the SAVED FILE'S OWN BYTES, never the log above:
'     python3 vba/tools/verify_caldate.py "/mnt/c/.../3. Project Progress.pptx" $n || break
'   done
' verify_caldate.py reads the .pptx as a zip and counts every
' MS<n>_CALDATE in the slide's raw XML, grouped vs stray, per slide -- the
' same from-disk check that caught the real duplicates tonight.

Option Explicit

' ---------- configuration ----------
Dim REAL_DECK_PATH, EXPECTED_NAME
REAL_DECK_PATH = "C:\Users\rohan\OneDrive\Claude\3. Project Progress.pptx"
EXPECTED_NAME = "3. Project Progress.pptx"

Dim GAP_PT, HEIGHT_PT, FONT_SIZE
GAP_PT = 0.015 * 72
HEIGHT_PT = 0.10 * 72
FONT_SIZE = 5.5

' ---------- arguments ----------
If WScript.Arguments.Count < 1 Then
    WScript.Echo "Usage: cscript //Nologo add_caldate_shapes.vbs <slideIndex> [deckPath]"
    WScript.Quit 1
End If
Dim slideIdx, deckPath
slideIdx = CInt(WScript.Arguments(0))
If WScript.Arguments.Count > 1 Then
    deckPath = WScript.Arguments(1)
Else
    deckPath = REAL_DECK_PATH
End If

' ---------- helpers ----------

' The group carrying MS_BAR is the milestone device. GroupItems enumerates
' nested members too (COM presents the group flatter than the XML nests it
' -- the 2026-08-13 lesson), so this finds the bar at any nesting depth.
Function FindDeviceGroup(sld)
    Dim shp, inner
    Set FindDeviceGroup = Nothing
    For Each shp In sld.Shapes
        If shp.Type = 6 Then ' msoGroup
            For Each inner In shp.GroupItems
                If inner.Name = "MS_BAR" Then
                    Set FindDeviceGroup = shp
                    Exit Function
                End If
            Next
        End If
    Next
End Function

' Top-level Shapes index of the shape with this Id. Grouping is done via
' Range(Array(index, index)) resolved from shape IDs, not names -- this
' script exists to repair duplicate-NAME states, so names are exactly the
' thing that cannot be trusted as selectors here.
Function TopLevelIndexOfId(sld, shapeId)
    Dim i
    TopLevelIndexOfId = 0
    For i = 1 To sld.Shapes.Count
        If sld.Shapes.Item(i).Id = shapeId Then
            TopLevelIndexOfId = i
            Exit Function
        End If
    Next
End Function

' Count of live POWERPNT.EXE processes, via WMI (works without any COM
' handle -- usable before CreateObject and after Quit).
Function PowerPointProcessCount()
    Dim wmi, procs
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Set procs = wmi.ExecQuery("SELECT ProcessId FROM Win32_Process WHERE Name = 'POWERPNT.EXE'")
    PowerPointProcessCount = procs.Count
End Function

' Full styling for a CALDATE shape, applied identically to a freshly
' created shape and to an adopted stray (a stray may have been interrupted
' mid-styling, so adoption re-asserts everything rather than trusting it).
Sub StyleCaldateShape(shp, leftPt, topPt, widthPt)
    ' AutoSize MUST be turned off before anything else -- an empty textbox
    ' otherwise auto-fits to its (empty) content and collapses to ~0 width
    ' (shipped cx="65" EMU shapes once; see add_pct_shapes.vbs).
    shp.TextFrame.AutoSize = 0 ' ppAutoSizeNone
    shp.TextFrame.MarginLeft = 0
    shp.TextFrame.MarginRight = 0
    shp.TextFrame.MarginTop = 0
    shp.TextFrame.MarginBottom = 0
    shp.TextFrame.TextRange.Text = ""
    shp.TextFrame.TextRange.Font.Size = FONT_SIZE
    shp.TextFrame.TextRange.Font.Bold = False
    shp.TextFrame.TextRange.ParagraphFormat.Alignment = 2 ' ppAlignCenter
    ' Re-assert geometry explicitly, after every text/font property is set
    ' -- the one proven place autofit can still silently override it.
    shp.Left = leftPt
    shp.Top = topPt
    shp.Width = widthPt
    shp.Height = HEIGHT_PT
    shp.Visible = False
End Sub

' ---------- open ----------
Dim preExistingPpt
preExistingPpt = PowerPointProcessCount()
If preExistingPpt > 0 Then
    WScript.Echo "WARNING: " & preExistingPpt & " POWERPNT process(es) already running before this script started -- the end-of-run exit poll cannot tell them apart from this script's own instance."
End If

' Presentations.Open on a bad path does NOT throw -- it blocks forever on
' a modal dialog (failure mode 4). Refuse a nonexistent path up front.
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(deckPath) Then
    WScript.Echo "FATAL: deck file does not exist (check the caller's quoting): [" & deckPath & "]"
    WScript.Quit 1
End If

Dim ppt
Set ppt = CreateObject("PowerPoint.Application")
ppt.Visible = True
ppt.DisplayAlerts = 1 ' ppAlertsNone -- alert-class dialogs must error, not block automation

Dim pres
Set pres = ppt.Presentations.Open(deckPath, False, False, True)

' Checked by NAME, not the full path. Presentations.Open was called with
' the exact path above, so PowerPoint cannot have opened a different file
' -- but a fully OneDrive-synced file's own FullName sometimes reports as
' the cloud https://d.docs.live.net/... alias instead of the local path
' afterward (confirmed live 2026-08-22: every one of 36 slides hit this
' and aborted safely, having written nothing). That is not a wrong-file
' risk, it is OneDrive's own path virtualization -- Open() itself is what
' guarantees the right file; this check exists only to catch a genuinely
' different file being opened by mistake.
If pres.Name <> EXPECTED_NAME Then
    WScript.Echo "WARNING -- opened file does not match, aborting: " & pres.Name
    pres.Saved = True ' discard-close: nothing was written, keep it that way
    pres.Close()
    ppt.Quit()
    WScript.Quit 1
End If

' ---------- locate the device and capture its identity ----------
Dim sld
Set sld = pres.Slides.Item(slideIdx)

Dim deviceGroup
Set deviceGroup = FindDeviceGroup(sld)
If deviceGroup Is Nothing Then
    WScript.Echo "Slide " & slideIdx & " : no milestone device group found -- nothing to do"
    pres.Saved = True
    pres.Close()
    ppt.Quit()
    WScript.Quit 0
End If

' Capture name + every tag ONCE, before any regroup touches them; restored
' after every regroup below (regrouping wipes both).
Dim origName, tagNames(), tagValues(), tagCount, i
origName = deviceGroup.Name
tagCount = deviceGroup.Tags.Count
If tagCount > 0 Then
    ReDim tagNames(tagCount - 1)
    ReDim tagValues(tagCount - 1)
    For i = 1 To tagCount
        tagNames(i - 1) = deviceGroup.Tags.Name(i)
        tagValues(i - 1) = deviceGroup.Tags.Value(i)
    Next
Else
    ' A group with no tags at all is itself the signature of an earlier
    ' interrupted run (regroup wiped them, restore never ran). Not fixable
    ' from here -- the true tag values are gone -- so say so loudly.
    WScript.Echo "WARNING: device group '" & origName & "' carries NO tags -- likely leftover damage from an earlier interrupted regroup; tags cannot be restored by this script."
End If

' Restores identity onto whatever group now carries MS_BAR (the regroup
' that just ran replaced the group object, so re-acquire, never reuse).
Sub RestoreGroupIdentity()
    Dim g, j
    Set g = FindDeviceGroup(sld)
    g.Name = origName
    For j = 0 To tagCount - 1
        g.Tags.Add tagNames(j), tagValues(j)
    Next
End Sub

' Groups the current device group with one top-level shape, then restores
' the group's name and tags.
Sub AbsorbIntoGroup(memberShape)
    Dim g, gIdx, sIdx
    Set g = FindDeviceGroup(sld)
    gIdx = TopLevelIndexOfId(sld, g.Id)
    sIdx = TopLevelIndexOfId(sld, memberShape.Id)
    sld.Shapes.Range(Array(gIdx, sIdx)).Group
    RestoreGroupIdentity
End Sub

' ---------- per-slot processing ----------
' All loop-body variables are declared HERE, outside the loops: a
' fixed-size array Dim inside a VBScript loop body is re-executed on the
' second iteration and raises "Type mismatch" (found live, first test
' run, slide 40 slot 2 -- slot 1 worked, slot 2 died on the Dim line).
' Arrays are declared dynamic and ReDim'd per iteration instead.
Dim added, adopted, deduped, slot
Dim dateName, dateShp, inner, caldName, shp
Dim groupedMatches(), groupedCount, strayMatches(), strayCount
Dim leftPt, widthPt, topPt, newShape
added = 0
adopted = 0
deduped = 0

For slot = 1 To 7
    ' RE-ACQUIRE the device group fresh EVERY slot -- a reference held
    ' across a member deletion or a regroup goes stale and the next use
    ' raises "Type mismatch" (confirmed live, dedupe_caldate.vbs).
    Set deviceGroup = FindDeviceGroup(sld)
    If deviceGroup Is Nothing Then
        WScript.Echo "FATAL: device group lost mid-run at slot " & slot & " -- discarding in-memory changes (nothing saved this run; re-run is safe)"
        pres.Saved = True
        pres.Close()
        ppt.Quit()
        WScript.Quit 1
    End If

    dateName = "MS" & slot & "_DATE"
    Set dateShp = Nothing
    For Each inner In deviceGroup.GroupItems
        If inner.Name = dateName Then
            Set dateShp = inner
            Exit For
        End If
    Next

    If dateShp Is Nothing Then
        ' No DATE label -> this slot doesn't exist on this device; nothing
        ' to add and nothing expected. (Slots are contiguous from 1 in
        ' practice, but that isn't relied on.)
    Else
        caldName = "MS" & slot & "_CALDATE"

        ' Existence check widened (failure mode 5): look in BOTH the
        ' group's items AND the slide's top-level shapes. A stray with the
        ' right name at top level is the interrupted-run signature.
        ReDim groupedMatches(19)
        ReDim strayMatches(19)
        groupedCount = 0
        For Each inner In deviceGroup.GroupItems
            If inner.Name = caldName Then
                Set groupedMatches(groupedCount) = inner
                groupedCount = groupedCount + 1
            End If
        Next
        strayCount = 0
        For Each shp In sld.Shapes
            If shp.Name = caldName Then
                Set strayMatches(strayCount) = shp
                strayCount = strayCount + 1
            End If
        Next

        leftPt = dateShp.Left
        widthPt = dateShp.Width
        topPt = dateShp.Top + dateShp.Height + GAP_PT

        If groupedCount = 1 And strayCount = 0 Then
            WScript.Echo "Slide " & slideIdx & " slot " & slot & " : " & caldName & " already grouped -- skipped"
        ElseIf groupedCount >= 1 Then
            ' At least one good grouped copy exists; everything else with
            ' this name (extra grouped copies, strays) is a duplicate.
            WScript.Echo "Slide " & slideIdx & " slot " & slot & " : " & caldName & " DUPLICATE (grouped=" & groupedCount & ", stray=" & strayCount & ") -- keeping first grouped copy Id=" & groupedMatches(0).Id
            For i = 1 To groupedCount - 1
                WScript.Echo "  deleting extra grouped copy Id=" & groupedMatches(i).Id
                groupedMatches(i).Delete
                deduped = deduped + 1
            Next
            For i = 0 To strayCount - 1
                WScript.Echo "  deleting stray copy Id=" & strayMatches(i).Id
                strayMatches(i).Delete
                deduped = deduped + 1
            Next
        ElseIf strayCount >= 1 Then
            ' No grouped copy, but a stray survivor exists -- an earlier
            ' run died between AddTextbox and the regroup. Adopt it: it may
            ' also have died mid-styling, so re-style fully (idempotent --
            ' these shapes are empty and hidden until DrawMilestones writes
            ' them), then regroup it in, exactly as creation would.
            WScript.Echo "Slide " & slideIdx & " slot " & slot & " : " & caldName & " found STRAY (ungrouped, x" & strayCount & ") -- adopting Id=" & strayMatches(0).Id & " into the device group"
            For i = 1 To strayCount - 1
                WScript.Echo "  deleting extra stray Id=" & strayMatches(i).Id
                strayMatches(i).Delete
                deduped = deduped + 1
            Next
            StyleCaldateShape strayMatches(0), leftPt, topPt, widthPt
            AbsorbIntoGroup strayMatches(0)
            adopted = adopted + 1
        Else
            ' Genuinely missing -- create it (the original, proven path).
            Set newShape = sld.Shapes.AddTextbox(1, leftPt, topPt, widthPt, HEIGHT_PT)
            newShape.Name = caldName
            StyleCaldateShape newShape, leftPt, topPt, widthPt
            AbsorbIntoGroup newShape
            added = added + 1
            WScript.Echo "Slide " & slideIdx & " slot " & slot & " : " & caldName & " added"
        End If
    End If
Next

' ---------- final per-slide self-check (failure mode 5's backstop), BEFORE save ----------
' Re-count every name from scratch. After the per-slot pass above this
' must find exactly 0 or 1 of each; more than one means THIS run just
' introduced (or failed to catch) a duplicate -- repair it, but shout,
' because that is a bug in this script, not in the deck.
Dim selfCheckFixes, n, total, kept
Dim scName, scGrouped(), scGroupedCount, scStray(), scStrayCount
selfCheckFixes = 0
For n = 1 To 7
    Set deviceGroup = FindDeviceGroup(sld) ' fresh again -- same staleness rule
    If deviceGroup Is Nothing Then
        WScript.Echo "FATAL: device group lost during self-check -- discarding in-memory changes (nothing saved this run)"
        pres.Saved = True
        pres.Close()
        ppt.Quit()
        WScript.Quit 1
    End If
    scName = "MS" & n & "_CALDATE"
    ReDim scGrouped(19)
    ReDim scStray(19)
    scGroupedCount = 0
    For Each inner In deviceGroup.GroupItems
        If inner.Name = scName Then
            Set scGrouped(scGroupedCount) = inner
            scGroupedCount = scGroupedCount + 1
        End If
    Next
    scStrayCount = 0
    For Each shp In sld.Shapes
        If shp.Name = scName Then
            Set scStray(scStrayCount) = shp
            scStrayCount = scStrayCount + 1
        End If
    Next
    total = scGroupedCount + scStrayCount
    If total > 1 Then
        WScript.Echo "SELF-CHECK: " & scName & " has " & total & " copies AFTER processing (grouped=" & scGroupedCount & ", stray=" & scStrayCount & ") -- this run has a bug; repairing (keeping a grouped copy) but DO NOT trust this run without the external XML verification"
        kept = False
        For i = 0 To scGroupedCount - 1
            If Not kept Then
                kept = True
            Else
                scGrouped(i).Delete
                selfCheckFixes = selfCheckFixes + 1
            End If
        Next
        For i = 0 To scStrayCount - 1
            If Not kept Then
                kept = True
            Else
                scStray(i).Delete
                selfCheckFixes = selfCheckFixes + 1
            End If
        Next
    End If
Next
WScript.Echo "Slide " & slideIdx & " : added=" & added & " adopted=" & adopted & " deduped=" & deduped & " selfCheckFixes=" & selfCheckFixes

' ---------- save, with retry (failure mode 2) ----------
Dim saveOutcome
saveOutcome = ""
On Error Resume Next
Err.Clear
pres.Save
If Err.Number <> 0 Then
    Dim firstErr
    firstErr = Err.Description
    Err.Clear
    ' The error is NOT reliably a real failure on OneDrive-synced files:
    ' Presentation.Saved is the discriminator (confirmed live 2026-08-22
    ' -- Saved=True right after the error, on-disk bytes already correct).
    If pres.Saved Then
        saveOutcome = "Save() raised [" & firstErr & "] but Presentation.Saved=True -- local write landed (OneDrive cloud-confirm noise), proceeding"
    Else
        WScript.Echo "Save() raised [" & firstErr & "] and Presentation.Saved=False -- genuine failure; waiting 5s, then retrying ONCE on the same open presentation"
        WScript.Sleep 5000
        Err.Clear
        pres.Save
        If pres.Saved Then
            saveOutcome = "retry Save() succeeded (Saved=True)"
        Else
            On Error GoTo 0
            WScript.Echo "FATAL: retry Save() also failed and Presentation.Saved=False -- the file on disk does NOT have this run's changes."
            WScript.Echo "PowerPoint is deliberately LEFT OPEN with the unsaved presentation so a human can save it by hand. Do not force-kill it without checking Saved first."
            WScript.Quit 2
        End If
    End If
Else
    saveOutcome = "Save() clean"
End If
On Error GoTo 0
WScript.Echo "Slide " & slideIdx & " : " & saveOutcome

' ---------- close robustly (failure mode 3) ----------
' Close + Quit, then poll BRIEFLY for process exit. Long passive waits are
' pointless (80s of polling never once resolved a lingering POWERPNT,
' confirmed live) -- so after ~9s this script logs and exits anyway. A
' lingering process at this point is holding a file already confirmed
' saved (Saved=True above), so force-closing it is safe FOR THE CALLER to
' do -- but only the caller, and only with that confirmation in hand.
On Error Resume Next
pres.Close
ppt.Quit
On Error GoTo 0

Dim attempt, remaining
remaining = -1
For attempt = 1 To 6
    WScript.Sleep 1500
    remaining = PowerPointProcessCount()
    If remaining <= preExistingPpt Then Exit For
Next
If remaining <= preExistingPpt Then
    WScript.Echo "PowerPoint exited cleanly."
Else
    WScript.Echo "NOTE: POWERPNT still running after Quit() + ~9s poll (known lingering behaviour; Saved=True was already confirmed). Leaving it for the caller to force-close if needed."
End If
WScript.Echo "Done. Verify from disk: python3 vba/tools/verify_caldate.py ""<deck>"" " & slideIdx
