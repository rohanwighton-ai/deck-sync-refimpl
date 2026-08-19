Attribute VB_Name = "TemplateAudit"
Option Explicit

' Answers one question about a slide type: WHAT ON THIS SLIDE IS THE TOOL NOT
' TRACKING?
'
' Why this exists. Step 1 gave each type a master template and blanked its
' tagged fields to placeholders. It could not blank anything else, because
' nothing distinguishes a project's data from the slide's furniture --
' "~$280K" must go, "STRATEGIC ALIGNMENT" and the SAAFE logo must stay, and
' no available rule separates them. So untagged project content still rides
' from the template onto every created slide.
'
' The reframe that makes that tractable, and it came from Rohan on
' 2026-07-30: blanking is the wrong target. If a value differs per project it
' wants to be a FIELD; if it does not, it is chrome and should stay exactly
' as it is. So "clean the template" is really "finish onboarding the type",
' and it has a definable end state -- the template is finished when nothing
' project-specific on it is untagged.
'
' Scale of the gap on the real deck, MEASURED by running this module against
' it 2026-07-31: type `q` tracks 5 fields on a slide carrying **77** separate
' text items -- 38 plausible project data, 39 fixed furniture. A visual
' estimate the day before had put it near 25, out by a factor of three, which
' is itself the argument for the tool existing: this is a number people guess
' and guess low. It had never bitten because every existing slide was
' hand-authored -- inheritance only happens on CREATED slides, and nothing had
' ever been created into that deck from a real template until step 1 shipped.
'
' KNOWN LIMITATION, found on that same first live run and not yet addressed.
' The comparison below assumes sibling slides were INDEPENDENTLY AUTHORED. In
' a deck whose slides were produced by cloning -- which is what this tool's own
' slide creation does -- untagged content is identical across siblings by
' construction, so genuinely project-specific text still scores above zero and
' the "on no other slide" verdict never fires. On the real deck it fired zero
' times out of 77. The ordering still separates usefully (universal text ranks
' below non-universal), so the LIST is sound; it is the headline count that
' misleads. The likely correction is to treat "on all other slides" as chrome
' and anything less as a candidate, rather than reserving the strong verdict
' for zero -- deliberately not applied yet, because it should be judged against
' a hand-authored deck as well as this one.
'
' The classification signal, and why it is computable. For each untagged text
' on the template, count how many OTHER slides of the same type carry the
' identical text. Identical on all of them => almost certainly chrome. On none
' of them => almost certainly this project's own data. This needs no shape
' correspondence between slides (which hand-authored decks cannot be trusted
' to preserve) -- only a set of strings per slide. It is a heuristic and the
' report says so; it exists to ORDER the list, not to decide for anyone.
'
' Output goes to a WORKSHEET, not a MsgBox. A real slide has dozens of text
' shapes, and MsgBox truncates -- this is the same failure as the 45-InputBox
' chain replaced on 2026-07-29 by a review grid in the workbook
' (BatchOnboardFlow.WriteInstanceKeyGrid, the precedent this follows). A grid
' can also be sorted, worked down over several sittings, and re-run to see
' progress, none of which a dialog can do.
'
' Reads only. Nothing here writes to the deck; the only write is the audit
' worksheet itself.

Public Const AUDIT_SHEET_NAME As String = "Template Audit"

Private Const COL_SHAPE As Long = 1
Private Const COL_GROUP As Long = 2
Private Const COL_TEXT As Long = 3
Private Const COL_VERDICT As Long = 4
Private Const COL_SEEN As Long = 5
Private Const COL_DECISION As Long = 6

Private Const AUDIT_HEADER_ROW As Long = 1
Private Const AUDIT_FIRST_ROW As Long = 2

' The verdict prefixes, declared once. Every place that needs to recognise a
' verdict compares against these, so a wording change moves one line.
Public Const VERDICT_DATA As String = "LIKELY PROJECT DATA"
Public Const VERDICT_CHECK As String = "CHECK"
Public Const VERDICT_UNKNOWN As String = "UNKNOWN"

Public Type AuditRow
    ShapeName As String
    GroupPath As String        ' "" when top-level; nested fields are common on real decks
    Text As String
    Verdict As String
    SeenOn As Long             ' how many OTHER instances carry this identical text
    InstanceCount As Long
End Type

' The heuristic, isolated and pure so its wording and its boundaries are
' testable without a deck.
'
' Deliberately does NOT return a boolean. "Chrome or data" is a judgement
' about intent that only Rohan can make, and a two-valued answer would invite
' the caller to act on it. Three graded verdicts plus an explicit UNKNOWN
' keeps the decision where it belongs and keeps the middle case visible
' instead of rounded away.
Public Function Classify(seenOn As Long, instanceCount As Long) As String
    If instanceCount = 0 Then
        ' The honest answer when there is nothing to compare against. A
        ' single-slide type is a real state (a type onboarded from one
        ' example), and guessing "project data" here would be presenting a
        ' coin toss as a finding.
        Classify = VERDICT_UNKNOWN & " -- no other slides of this type to compare against"
    ElseIf seenOn = 0 Then
        Classify = VERDICT_DATA & " -- on no other slide of this type"
    ElseIf seenOn = instanceCount Then
        Classify = "probably chrome -- identical on all " & instanceCount & " other slide(s)"
    Else
        Classify = VERDICT_CHECK & " -- on " & seenOn & " of " & instanceCount & " other slide(s)"
    End If
End Function

' True when `verdict` starts with `prefix`. Uses InStr(...) = 1 rather than
' Left(verdict, Len(prefix)) -- not a style preference: the first version of
' VerdictRank below hard-coded the lengths and got one wrong
' (Left(verdict, 18) against a 19-character literal), so the comparison could
' never be true, the rank silently fell into the same band as chrome, and the
' furniture sorted above the actionable rows. Caught by the suite 2026-07-30.
' Deriving the length from the string makes that class of mistake unavailable.
Private Function StartsWith(verdict As String, prefix As String) As Boolean
    StartsWith = (InStr(verdict, prefix) = 1)
End Function

' The one question callers actually ask of a verdict, exposed so the count in
' RibbonUI and the ordering here cannot drift apart -- they were separately
' written prefix matches until 2026-07-30, and only one of them was correct.
Public Function IsLikelyProjectData(verdict As String) As Boolean
    IsLikelyProjectData = StartsWith(verdict, VERDICT_DATA)
End Function

' Sort key for the report: the actionable rows first, the furniture last.
' A 60-row grid ordered by shape name is a wall; ordered by this it is a
' worklist that stops mattering partway down.
Private Function VerdictRank(verdict As String) As Long
    If StartsWith(verdict, VERDICT_DATA) Then
        VerdictRank = 1
    ElseIf StartsWith(verdict, VERDICT_CHECK) Then
        VerdictRank = 2
    ElseIf StartsWith(verdict, VERDICT_UNKNOWN) Then
        VerdictRank = 3
    Else
        VerdictRank = 4
    End If
End Function

' A shape's text, or "" for anything without readable text. Guarded rather
' than gated on Candidate.HasText alone: HasText comes from Discovery's own
' inspection, and a shape can satisfy it and still raise on TextRange (a
' picture with a caption frame, a chart, a table cell). A raise here would
' kill an audit that is supposed to be the safest thing on the toolbar.
Private Function ShapeText(shp As Object) As String
    On Error Resume Next
    ShapeText = shp.TextFrame.TextRange.Text
    On Error GoTo 0
End Function

' Normalised form used for cross-slide comparison. Trimmed, case-folded, and
' internal whitespace collapsed -- hand-authored slides differ by a trailing
' space or a line break far more often than they differ in meaning, and a
' comparison that treats "Project Closed " as different from "Project Closed"
' would report chrome as project data on almost every row.
Private Function NormaliseText(raw As String) As String
    Dim s As String
    s = raw
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    NormaliseText = LCase(Trim(s))
End Function

' Every normalised text string present anywhere on `sld`, as a Dictionary
' used as a set. Recurses groups, because real decks stack fields inside
' groups (a lesson already paid for -- see AGENTS.md / the discovery walk).
Private Function CollectSlideTexts(sld As Object) As Object
    Dim texts As Object
    Set texts = CreateObject("Scripting.Dictionary")

    Dim shapes() As Object
    Dim candidates() As Candidate
    candidates = Discovery.DiscoverSlideWithShapes(sld, shapes)

    Dim lo As Long, hi As Long, hasAny As Boolean
    On Error Resume Next
    lo = LBound(candidates)
    hi = UBound(candidates)
    hasAny = (Err.Number = 0)
    On Error GoTo 0

    If hasAny Then
        Dim i As Long
        For i = lo To hi
            Dim t As String
            t = NormaliseText(ShapeText(shapes(i)))
            If t <> "" Then texts(t) = True
        Next i
    End If

    Set CollectSlideTexts = texts
End Function

' The audit itself.
'
' `subjectSld` is ANY slide of the type -- deliberately not "the template".
' Nothing below needs it to be one: the audit walks one slide's shapes and
' compares their text against other slides, and a template is merely the most
' useful subject once one exists.
'
' That independence is the point, and it is a decision not an accident
' (Rohan, 2026-07-30). This operation and field marking must each work
' regardless of whether the other has been used, because users arrive at
' different points with decks of different maturity:
'   - never ran step 1, no template  -> audit any onboarded slide
'   - mid-onboarding, marking fields -> audit the slide being marked
'   - step 1 done                    -> audit the template
' Had this stayed typed to a template, a deck would have had to complete step
' 1 before it could find out which fields it was missing -- which is backwards,
' since knowing the fields is what makes the template worth building.
'
' Neither operation stores state the other reads. The only thing they share is
' this function, and it has no preconditions.
'
' `instances` is the type's other slides -- pass RunSync.GatherInstances(
' slideType), which excludes any template by construction (step 1's choke
' point). Passing a genuinely unallocated array is fine and means "no
' comparison available", which Classify reports as UNKNOWN rather than
' guessing.
'
' `trackedFields` comes back as a "|"-joined list of the role names the tool
' DOES manage, so the report can state what is covered as well as what is
' not. Reporting only the gap would make a well-onboarded type look identical
' to an un-onboarded one.
Public Function BuildAudit(subjectSld As Object, instances() As Object, ByRef rowCount As Long, ByRef trackedFields As String) As AuditRow()
    Dim results() As AuditRow
    rowCount = 0
    trackedFields = ""

    ' --- Which texts exist on the other instances -------------------------
    Dim iLo As Long, iHi As Long, hasInstances As Boolean
    On Error Resume Next
    iLo = LBound(instances)
    iHi = UBound(instances)
    hasInstances = (Err.Number = 0)
    On Error GoTo 0

    Dim instanceCount As Long
    instanceCount = 0
    Dim instanceTexts() As Object
    If hasInstances Then
        instanceCount = iHi - iLo + 1
        ReDim instanceTexts(1 To instanceCount)
        Dim n As Long
        For n = 1 To instanceCount
            Set instanceTexts(n) = CollectSlideTexts(instances(iLo + n - 1))
        Next n
    End If

    ' --- Walk the subject slide -------------------------------------------
    Dim shapes() As Object
    Dim candidates() As Candidate
    candidates = Discovery.DiscoverSlideWithShapes(subjectSld, shapes)

    Dim lo As Long, hi As Long, hasAny As Boolean
    On Error Resume Next
    lo = LBound(candidates)
    hi = UBound(candidates)
    hasAny = (Err.Number = 0)
    On Error GoTo 0

    If Not hasAny Then
        BuildAudit = results
        Exit Function
    End If

    Dim i As Long
    For i = lo To hi
        Dim role As String
        role = shapes(i).Tags("role")

        If role <> "" Then
            ' Tracked. Recorded for the summary, never audited -- a field is
            ' by definition already accounted for.
            trackedFields = trackedFields & IIf(trackedFields = "", "", "|") & role
        Else
            Dim raw As String
            raw = ShapeText(shapes(i))
            If Trim(raw) <> "" Then
                Dim norm As String
                norm = NormaliseText(raw)

                Dim seenOn As Long
                seenOn = 0
                Dim k As Long
                For k = 1 To instanceCount
                    If instanceTexts(k).Exists(norm) Then seenOn = seenOn + 1
                Next k

                rowCount = rowCount + 1
                ReDim Preserve results(1 To rowCount)
                results(rowCount).ShapeName = candidates(i).Name
                results(rowCount).GroupPath = candidates(i).GroupPath
                results(rowCount).Text = raw
                results(rowCount).SeenOn = seenOn
                results(rowCount).InstanceCount = instanceCount
                results(rowCount).Verdict = Classify(seenOn, instanceCount)
            End If
        End If
    Next i

    ' --- Actionable first -------------------------------------------------
    ' Insertion sort: rowCount is dozens at most, and a stable sort keeps
    ' document order within each verdict band, so the grid still reads
    ' top-to-bottom down the slide within a group.
    Dim a As Long, b As Long
    Dim tmp As AuditRow
    For a = 2 To rowCount
        tmp = results(a)
        b = a - 1
        Do While b >= 1
            If VerdictRank(results(b).Verdict) > VerdictRank(tmp.Verdict) Then
                results(b + 1) = results(b)
                b = b - 1
            Else
                Exit Do
            End If
        Loop
        results(b + 1) = tmp
    Next a

    BuildAudit = results
End Function

' Writes the grid. A "Decide" column is left empty on purpose: this is a
' worklist Rohan fills in, and having somewhere to record "field" / "chrome"
' / "drop" per row is what lets the job be done across several sittings
' instead of held in his head in one.
'
' Nothing reads this column back yet. Said plainly rather than implied,
' because a column that looks like an input but is ignored is worse than no
' column -- if a later step consumes it, that is the step to say so in.
'
' CLEARS the sheet first, and that is not incidental. Writing rows 1..N over a
' previous run's rows 1..M leaves M-N stale rows sitting below the new ones,
' indistinguishable from current findings -- so the audit would get LESS
' trustworthy the more of the work you did, which is the worst possible
' direction for a progress tool. Merging instead would need a stable per-row
' key, and the only candidates are shape name and text -- shape name is
' explicitly not an identity key in this project (specs/identity-tags.md),
' and text is the thing being changed. So a real merge is a design job, not a
' line of code, and pretending otherwise would silently drop decisions on
' rows whose text had been edited.
'
' REFUSES rather than silently discarding, 2026-08-16 -- found auditing this
' exact class of bug in DiscoverUI.bas the same day. SummaryText's own
' warning ("Re-running this REPLACES that sheet, decisions included") used to
' be the ONLY protection, then a straight refusal replaced it (2026-08-15 --
' "copy them out first"). FIX-LIST item P5, 2026-08-19: a refusal stops
' silent loss but still means the audit can only ever be worked in one
' sitting, which was the actual complaint. **Now carries every decision
' forward by shape identity instead**, the way `Drafting.WriteDraftingSheet`
' carries a person's drafts across a rebuild rather than either destroying
' them or blocking the rebuild entirely.
'
' Returns "" always now (nothing left to refuse) -- `carriedCount` and
' `orphanedCount` are ByRef so the caller can report what happened, the same
' convention as `FieldSpec.ApplyControlledValidation`'s `outOfVocabulary`.
Public Function WriteAuditGrid(ws As Object, rows() As AuditRow, rowCount As Long, _
                               Optional ByRef carriedCount As Long, _
                               Optional ByRef orphanedCount As Long) As String
    carriedCount = 0
    orphanedCount = 0

    ' Read every prior decision BEFORE anything is cleared, keyed by shape
    ' identity -- (shape name, group path, text) together, not shape name
    ' alone. Text is part of the key deliberately: a decision was made about
    ' what a shape SAID, and if the template's text has since changed, that
    ' decision may no longer apply to carry forward blindly.
    Dim priorDecisions As Object
    Set priorDecisions = CreateObject("Scripting.Dictionary")

    Dim priorLastRow As Long
    priorLastRow = AUDIT_FIRST_ROW - 1
    Do While Trim(CStr(ws.Cells(priorLastRow + 1, COL_SHAPE).Value)) <> ""
        priorLastRow = priorLastRow + 1
    Loop

    Dim pr As Long
    For pr = AUDIT_FIRST_ROW To priorLastRow
        Dim priorDecision As String
        priorDecision = Trim(CStr(ws.Cells(pr, COL_DECISION).Value))
        If priorDecision <> "" Then
            ' Stored text carries a leading apostrophe (forced-text guard
            ' below) -- stripped here so the key matches a freshly-read
            ' AuditRow's own Text, which never has one.
            Dim priorText As String
            priorText = CStr(ws.Cells(pr, COL_TEXT).Value)
            If Left(priorText, 1) = "'" Then priorText = Mid(priorText, 2)
            priorDecisions(AuditRowKey(CStr(ws.Cells(pr, COL_SHAPE).Value), _
                                        CStr(ws.Cells(pr, COL_GROUP).Value), priorText)) = priorDecision
        End If
    Next pr

    ws.Cells.Clear

    ws.Cells(AUDIT_HEADER_ROW, COL_SHAPE).Value = "Shape"
    ws.Cells(AUDIT_HEADER_ROW, COL_GROUP).Value = "Inside group"
    ws.Cells(AUDIT_HEADER_ROW, COL_TEXT).Value = "Text on the slide"
    ws.Cells(AUDIT_HEADER_ROW, COL_VERDICT).Value = "Guess"
    ws.Cells(AUDIT_HEADER_ROW, COL_SEEN).Value = "On how many other slides"
    ws.Cells(AUDIT_HEADER_ROW, COL_DECISION).Value = "Decide: field / chrome / drop"

    Dim seenKeys As Object
    Set seenKeys = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To rowCount
        Dim r As Long
        r = i + AUDIT_FIRST_ROW - 1
        ws.Cells(r, COL_SHAPE).Value = rows(i).ShapeName
        ws.Cells(r, COL_GROUP).Value = rows(i).GroupPath
        ' Leading apostrophe: a cell of text lifted off a slide can begin with
        ' "=", "-" or "+" and Excel would take it as a formula and show
        ' #NAME?. Forcing text keeps the audit readable for exactly the rows
        ' most likely to be numeric project data.
        ws.Cells(r, COL_TEXT).Value = "'" & rows(i).Text
        ws.Cells(r, COL_VERDICT).Value = rows(i).Verdict
        ws.Cells(r, COL_SEEN).Value = rows(i).SeenOn & " of " & rows(i).InstanceCount

        Dim key As String
        key = AuditRowKey(rows(i).ShapeName, rows(i).GroupPath, rows(i).Text)
        seenKeys(key) = True
        If priorDecisions.Exists(key) Then
            ws.Cells(r, COL_DECISION).Value = priorDecisions(key)
            carriedCount = carriedCount + 1
        Else
            ws.Cells(r, COL_DECISION).Value = ""
        End If
    Next i

    ' Named, not silently dropped -- a prior decision whose shape/text no
    ' longer appears in this rebuild (renamed, retyped, or the shape itself
    ' is gone) genuinely cannot be carried forward. Counted so the caller can
    ' say so, the same "say when it did nothing, and why" rule
    ' ApplyControlledValidation already follows.
    Dim k As Variant
    For Each k In priorDecisions.Keys
        If Not seenKeys.Exists(CStr(k)) Then orphanedCount = orphanedCount + 1
    Next k

    WriteAuditGrid = ""
End Function

' The shape's identity for carrying a decision across a rebuild -- not a
' shape ID PowerPoint assigns (this project's own AGENTS.md already records
' that auto-generated shape names keep resolving after a rename, so a raw
' name alone is not a safe identity key on its own), but name + group path +
' text together, which is exactly what a person is looking at when they type
' a decision next to a row.
Private Function AuditRowKey(shapeName As String, groupPath As String, text As String) As String
    AuditRowKey = shapeName & "|" & groupPath & "|" & text
End Function

' What the human sees in the dialog. Counts and a pointer, never the list --
' the list is the whole reason this writes a sheet.
'
' Pure, so the wording is testable.
Public Function SummaryText(slideType As String, subjectLabel As String, trackedCount As Long, rowCount As Long, likelyDataCount As Long, instanceCount As Long) As String
    Dim s As String
    s = "Field audit for '" & slideType & "'." & vbCrLf & _
        "Looked at: " & subjectLabel & vbCrLf & vbCrLf & _
        "    " & trackedCount & " field(s) the tool already manages" & vbCrLf & _
        "    " & rowCount & " text item(s) on that slide it does NOT" & vbCrLf & _
        "    " & likelyDataCount & " of those look like project data" & vbCrLf & vbCrLf

    If rowCount = 0 Then
        s = s & "Nothing untracked -- every text item on that slide is a managed field." & vbCrLf & vbCrLf
    Else
        s = s & "The full list is on the '" & AUDIT_SHEET_NAME & "' sheet of the paired" & vbCrLf & _
            "workbook, most-likely-project-data first, with a column to record" & vbCrLf & _
            "your decision on each: field / chrome / drop." & vbCrLf & vbCrLf & _
            "Re-running this REBUILDS that sheet -- any decision you've already" & vbCrLf & _
            "recorded carries forward automatically onto the matching row." & vbCrLf & _
            "Only a decision whose shape or text has actually changed since" & vbCrLf & _
            "won't carry, and you'll be told exactly how many, if any." & vbCrLf & vbCrLf
    End If

    If instanceCount = 0 Then
        ' Said loudly, because with nothing to compare against every verdict
        ' in the grid reads UNKNOWN, and a grid of UNKNOWNs looks like a
        ' broken tool rather than an honest one.
        s = s & "NOTE: this type has no other slides to compare against, so every" & vbCrLf & _
            "guess is UNKNOWN. The list is still complete; only the ranking is absent." & vbCrLf & vbCrLf
    Else
        s = s & "The 'Guess' column compares each text against the " & instanceCount & " other slide(s)" & vbCrLf & _
            "of this type: identical on all of them usually means chrome, on none" & vbCrLf & _
            "of them usually means this project's own data. It is a hint for" & vbCrLf & _
            "ordering the list, not a decision." & vbCrLf & vbCrLf
    End If

    s = s & "Nothing was written to the deck."
    SummaryText = s
End Function
