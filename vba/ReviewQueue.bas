Attribute VB_Name = "ReviewQueue"
Option Explicit

' R13: every replacement is seen by a human before it lands.
'
' Issued by the Research Manager 31 July 2026 after the first real sync run
' (PROJECT_STATUS, 46 slides) changed 19 of them without anyone seeing a single
' before-and-after. It passed the register's `Status = Approved` gate, which
' answers "is this text right?" -- and skipped the question R13 asks, which is
' "is this change to THIS SLIDE right?". Two different questions asked at two
' different moments; only the first had ever been built.
'
' WHY A SECOND MODULE RATHER THAN A FLAG ON RunSync
'
' Because the gate has to be structural. RunRoutineSyncWithSheet used to plan
' and write in one pass -- there was no moment between "decided" and "written"
' for a human to stand in, so R13 could not be expressed as a conditional
' inside it without inventing that moment first. This module is that moment:
' a change set that exists, durably, before anything touches a slide.
'
' THE REVIEW SURFACE IS A WORKSHEET, NOT A DIALOG (R13.4)
'
' Same call as BatchOnboardFlow's instance-key grid, for the same reason and on
' the same evidence. Rohan, 2026-07-29, on the modal version of that flow:
' "too time consuming for organised slides" -- forty-five prompts to confirm
' what was already known. Worse, modal input is irreplaceable: a single bad
' workbook path that day destroyed 45 hand-confirmed keys, because there was
' nowhere for them to live except in the sequence of dialogs.
'
' R13.4 asks for a review that can be left and returned to. A worksheet already
' is that -- resumability is not a feature added here, it is what a worksheet
' inherently has. Prose review is real reading, and ABOUT_BODY is prose across
' 43 entities; nobody finishes that in one sitting, and a dialog that must be
' cleared in one sitting will be cleared without being read.
'
' WHAT HOLDS R13.5 (approval does not carry across runs)
'
' Two independent mechanisms, deliberately:
'
'   1. BuildQueue REWRITES the sheet whole, with a fresh run stamp and every
'      approval blank. An approval can therefore only ever survive from the
'      build that created it through to the apply that consumes it.
'   2. Apply marks the sheet CONSUMED, and refuses a consumed sheet. The same
'      approvals cannot be applied twice.
'
' Plus a per-row content hash, which covers the case R13.5 implies but does not
' name: approve a row Tuesday, hand-edit that slide Wednesday, apply Thursday.
' The approval was for a before-and-after that no longer exists. The hash is
' recomputed at apply time from the LIVE slide text and the REGISTER's value;
' if it no longer matches the stored one, the row is dropped rather than
' applied -- so a stale approval can never overwrite a hand edit.
'
' A consequence worth stating, because it looks like a bug otherwise: editing
' the Current or Proposed columns in the sheet does nothing. Both are recomputed
' at apply time from their real sources. This is a review surface, not an
' editing surface -- the place to change a value is the register.

' ---------------------------------------------------------------------
' Declarations. All of them, above the first procedure -- a module-level
' Type or Const below the first Sub/Function reports its error in a
' DIFFERENT module (AGENTS.md, hit 2026-07-30, cost a full suite run).
' ---------------------------------------------------------------------

Public Type ReviewItem
    EntityKey As String        ' the slide's instance key
    FieldID As String
    CurrentValue As String     ' what is on the slide now
    ProposedValue As String    ' what the register would write
    ChangeHash As String
    BatchLabel As String       ' "" = reviewed individually
    Approved As Boolean
End Type

Public Type ReviewQueueSet
    Items() As ReviewItem      ' may be genuinely UNALLOCATED -- never assume
    Count As Long              ' LBound/UBound are safe. Trust this instead.
    RunStamp As String
    SlideType As String
    Consumed As Boolean
End Type

' Grid layout. Row 1 is the banner, row 2 the headers, row 3 the first item.
Private Const ROW_BANNER As Long = 1
Private Const ROW_HEADER As Long = 2
Private Const ROW_FIRST_ITEM As Long = 3

Private Const COL_ENTITY As Long = 1
Private Const COL_FIELDID As Long = 2
Private Const COL_CURRENT As Long = 3
Private Const COL_PROPOSED As Long = 4
Private Const COL_BATCH As Long = 5
Private Const COL_APPROVE As Long = 6
Private Const COL_HASH As Long = 7        ' hidden; the row's identity

' APPENDED, not inserted between Proposed and Batch where it reads better.
' These constants are shared by the writer and the reader, so shifting Approve
' from 6 to 7 would make every review sheet already on disk read its ticks out
' of the wrong column -- silently, and in the direction of approving things
' nobody approved. Placement is worth less than that.
Private Const COL_DIFF As Long = 8

Public Const REVIEW_SHEET_NAME As String = "Sync Review"

' Above this many distinct batches, the modal stops being readable and the fast
' path is refused.
'
' A wall of transformations in a dialog is dismissed, not read -- which would
' recreate the rubber stamp R13 exists to remove, just with more text on it. The
' number is a judgement, not a measurement: about what fits on screen without
' scrolling. Being wrong here costs a trip to the worksheet, which is the safe
' direction.
Public Const MAX_BATCHES_IN_MODAL As Long = 12

Private Const STATE_OPEN As String = "OPEN"
Private Const STATE_CONSUMED As String = "CONSUMED"

' Content kinds. Governs BATCHABILITY, and is orthogonal to the register's
' `Quarter` column -- Quarter carries cadence (ALL vs a period), this carries
' kind. PROJECT_STATUS and KEY_EVENTS_BODY are both quarterly and must be
' treated oppositely, so cadence cannot stand in for kind.
Public Const KIND_CONTROLLED As String = "Controlled"   ' batchable when uniform
Public Const KIND_PROSE As String = "Prose"             ' never batchable (R13.2)
Public Const KIND_STATIC As String = "Static"           ' individually, but rare

' ---------------------------------------------------------------------
' Sheet naming
' ---------------------------------------------------------------------

' One review sheet PER SLIDE TYPE, not one per deck.
'
' A single shared sheet would work today -- the real deck has one registered
' type -- and would break silently the moment a second is onboarded: building
' the queue for type B would clear type A's un-applied approvals, and the only
' symptom would be approvals that quietly stopped existing. Naming per type
' costs one function and removes the failure entirely.
Public Function ReviewSheetNameFor(slideType As String) As String
    ' READABLE FIRST, UNIQUE SECOND. This produced
    ' "Sync Review project-st-43212D3D" -- Rohan could not find it and asked
    ' where the "sync review file" was, which is the name's fault: it reads like
    ' a temp file and truncates the slide type mid-word.
    '
    ' "Review project-status-3D1B" fits Excel's 31-character cap with the type
    ' intact. The tag shrinks from 8 hex digits to 4 and stays, because it is not
    ' decoration: two slide types sharing a truncated prefix would otherwise
    ' collapse onto ONE sheet, and WriteQueueSheet clears the sheet it writes --
    ' silently destroying the other type's un-applied ticks.
    Dim tag As String
    tag = Right("0000" & Hex(TypeNameHash(slideType)), 4)
    ReviewSheetNameFor = WorkbookBridge.SanitizeSheetName( _
        Left("Review " & slideType, 26) & "-" & tag)
End Function

' Small non-cryptographic hash, only to keep truncated sheet names distinct.
'
' Excel caps a sheet name at 31 characters, so SanitizeSheetName's Left(.., 31)
' left just 19 characters of slideType -- two types sharing a 19-character
' prefix produced the SAME sheet, and WriteQueueSheet's Cells.Clear on the
' second wiped the first's un-applied ticks. That is exactly the silent
' approval-loss the per-type naming exists to prevent, reintroduced by the
' truncation underneath it. Found by review 2026-07-31; latent today because
' the only registered type is "q", which is why it survived being written and
' tested.
Private Function TypeNameHash(text As String) As Long
    Dim h As Double
    h = 5381#
    Dim i As Long
    For i = 1 To Len(text)
        h = (h * 33#) + AscW(Mid(text, i, 1))
        h = h - (Int(h / 2147483647#) * 2147483647#)
    Next i
    TypeNameHash = CLng(h)
End Function

' ---------------------------------------------------------------------
' Content kind
' ---------------------------------------------------------------------

' HARDCODED ON PURPOSE, 31 July 2026, and this is the one shortcut in the
' module.
'
' R13.2 says batchability "falls out of the class column". There is no class
' column -- the register carries Quarter, EntityCode, SlideType, FieldID,
' FieldType, Value, Status and nothing that separates a controlled vocabulary
' from prose. Asking for one starts another exchange round, and the honest
' position is that with two fields in flight a lookup table beats a schema
' change. Round 13 §4 makes the ask for when there is a third.
'
' THE DEFAULT IS THE LOAD-BEARING PART. Anything not named here is Prose, and
' Prose never batches. Getting this wrong in the Prose direction costs a longer
' review; getting it wrong in the Controlled direction converts N unreviewed
' writes into one click, which is the exact failure R13 exists to prevent. So
' the absence of a label is never read as permission to batch.
Public Function ContentKindOf(fieldId As String) As String
    Select Case UCase(Trim(fieldId))
        Case "PROJECT_STATUS":  ContentKindOf = KIND_CONTROLLED
        Case "PROJECT_NAME":    ContentKindOf = KIND_STATIC
        Case "PROJECT_CODE":    ContentKindOf = KIND_STATIC
        Case Else:              ContentKindOf = KIND_PROSE
    End Select
End Function

' ---------------------------------------------------------------------
' Hashing
' ---------------------------------------------------------------------

' A row's identity: the exact before-and-after a human approved.
'
' Two independent polynomial hashes rather than one, concatenated. A single
' 32-bit hash over a 19-row queue is fine on collision odds, but the cost of a
' collision here is applying a change nobody approved, so the cheap insurance
' is worth taking.
'
' Arithmetic in Double, not Long: VBA raises overflow on Long multiplication,
' and hash * 31 against a modulus near 1e9 exceeds Long's range every
' iteration. Double carries 53 bits of mantissa, so 1e9 * 31 stays exact.
' What actually differs between two values, in words.
'
' 2026-08-08 on the rig: the review sheet showed
'     now:  'Applying quantitative microbial risk assessment ... composts usage '
'     new:  'Applying quantitative microbial risk assessment ... composts usage'
' and asked for approval. 198 characters against 197 -- the slide had a TRAILING
' SPACE. Two rows of eleven were like that. A reviewer cannot approve what they
' cannot see, and a sheet that asks them to is training them to tick blind.
'
' Returns "" only when the difference is genuinely visible on its face.
Public Function DescribeDifference(currentValue As String, proposedValue As String) As String
    If currentValue = proposedValue Then Exit Function

    ' Whitespace-only: the case that looks like no difference at all.
    If Trim(currentValue) = Trim(proposedValue) Then
        Dim note As String
        If Len(currentValue) - Len(LTrim(currentValue)) <> Len(proposedValue) - Len(LTrim(proposedValue)) Then
            note = "leading space"
        End If
        If Len(currentValue) - Len(RTrim(currentValue)) <> Len(proposedValue) - Len(RTrim(proposedValue)) Then
            note = note & IIf(note = "", "", " and ") & "trailing space"
        End If
        If note = "" Then note = "spacing"
        DescribeDifference = "INVISIBLE: " & note & " only (" & _
            Len(currentValue) & " chars on the slide vs " & Len(proposedValue) & ")"
        Exit Function
    End If

    If LCase(currentValue) = LCase(proposedValue) Then
        DescribeDifference = "INVISIBLE at a glance: capitalisation only"
        Exit Function
    End If

    ' Where they part company, and whether the character there is one you can see.
    Dim i As Long, shorter As Long
    shorter = Len(currentValue)
    If Len(proposedValue) < shorter Then shorter = Len(proposedValue)
    For i = 1 To shorter
        If Mid(currentValue, i, 1) <> Mid(proposedValue, i, 1) Then Exit For
    Next i

    If i > shorter Then
        DescribeDifference = "one is longer: the slide has " & Len(currentValue) & _
            " chars, the register " & Len(proposedValue)
        Exit Function
    End If

    Dim cNow As Long, cNew As Long
    cNow = AscW(Mid(currentValue, i, 1))
    cNew = AscW(Mid(proposedValue, i, 1))
    If cNow < 33 Or cNew < 33 Or cNow = 160 Or cNew = 160 Or cNow > 8191 Or cNew > 8191 Then
        DescribeDifference = "INVISIBLE: differs at character " & i & _
            " (slide has code " & cNow & ", register has code " & cNew & ")"
        Exit Function
    End If

    DescribeDifference = "differs from character " & i
End Function

Public Function ChangeHash(entityKey As String, fieldId As String, _
                           currentValue As String, proposedValue As String) As String
    Dim material As String
    ' Chr(1) as the separator, not "|" -- "|" appears in real field text, and a
    ' separator that can occur in the material lets two different rows produce
    ' identical material.
    material = entityKey & Chr(1) & fieldId & Chr(1) & currentValue & Chr(1) & proposedValue

    Dim h1 As Double, h2 As Double
    h1 = 2166136261#
    h2 = 5381#

    Dim i As Long
    For i = 1 To Len(material)
        Dim c As Double
        c = AscW(Mid(material, i, 1))
        h1 = ((h1 * 31#) + c)
        h1 = h1 - (Int(h1 / 1000000007#) * 1000000007#)
        h2 = ((h2 * 131#) + c)
        h2 = h2 - (Int(h2 / 999999937#) * 999999937#)
    Next i

    ChangeHash = Hex(CLng(h1)) & "-" & Hex(CLng(h2)) & "-" & Len(material)
End Function

' ---------------------------------------------------------------------
' Building the queue
' ---------------------------------------------------------------------

' Every change a sync would make to `slideType`, and nothing else.
'
' R13.3: unchanged rows never enter. The 46-row PROJECT_STATUS run was 27
' unchanged and 19 changed; only the 19 are decisions. Putting the other 27 in
' front of someone is how a review becomes a formality -- so this reads only
' `in_place_correction` actions and ignores no_change entirely.
'
' new_record and flagged are also excluded, and that is not an oversight:
' neither writes to a slide. Creation is a separate operation a person chooses
' (RunSync.CreateMissingSlides), and a flag is a report. R13 governs
' REPLACEMENT, so the queue holds replacements.
'
' dryRun:=True on the plan is load-bearing. PlanRoutineSync writes corrected
' text WHILE planning (see its header); building the queue wet would perform
' the very changes it exists to hold back, and the grid would be describing
' work already done.
Public Function BuildQueue(sheet As Sheet, slideType As String) As ReviewQueueSet
    Dim q As ReviewQueueSet
    q.SlideType = slideType
    q.RunStamp = MakeRunStamp()
    q.Count = 0
    q.Consumed = False

    Dim instances() As Object
    instances = RunSync.GatherInstances(slideType)

    Dim actions() As SyncAction
    actions = SyncOperations.PlanRoutineSync(instances, sheet.InstanceOrder, sheet.Rows, True)

    Dim lo As Long, hi As Long, hasActions As Boolean
    On Error Resume Next
    lo = LBound(actions)
    hi = UBound(actions)
    hasActions = (Err.Number = 0)
    On Error GoTo 0

    If hasActions Then
        Dim i As Long
        For i = lo To hi
            If actions(i).Kind = "in_place_correction" Then
                Dim fieldName As Variant
                For Each fieldName In actions(i).ChangedFieldCurrent.Keys
                    q.Count = q.Count + 1
                    ' Never pre-ReDim to (1 To 0) -- it throws Err 9 at runtime
                    ' (AGENTS.md). Growing from the first real item allocates
                    ' cleanly, and a queue with no items stays unallocated.
                    ReDim Preserve q.Items(1 To q.Count)
                    q.Items(q.Count).EntityKey = actions(i).InstanceKey
                    q.Items(q.Count).FieldID = CStr(fieldName)
                    q.Items(q.Count).CurrentValue = CStr(actions(i).ChangedFieldCurrent(fieldName))
                    q.Items(q.Count).ProposedValue = CStr(actions(i).ChangedFieldNew(fieldName))
                    q.Items(q.Count).ChangeHash = ChangeHash( _
                        q.Items(q.Count).EntityKey, q.Items(q.Count).FieldID, _
                        q.Items(q.Count).CurrentValue, q.Items(q.Count).ProposedValue)
                    q.Items(q.Count).BatchLabel = ""
                    q.Items(q.Count).Approved = False
                Next fieldName
            End If
        Next i
    End If

    AssignBatches q

    BuildQueue = q
End Function

' A run stamp that a human can read and two runs cannot share.
Public Function MakeRunStamp() As String
    MakeRunStamp = Format(Now, "yyyy-mm-dd hh:nn:ss")
End Function

' ---------------------------------------------------------------------
' Batching (R13.2)
' ---------------------------------------------------------------------

' Uniformity is MEASURED from the change set, not declared in a spec.
'
' R13.2's wording is the load-bearing part: "a VERIFIED batch of uniform
' changes", and "every instance shares one FieldID AND one transformation". So
' a batch here requires all three of FieldID, CurrentValue and ProposedValue to
' be exactly equal across its members -- which is what "one transformation"
' actually means when you have to compute it. The 19 "In progress" -> "In
' Progress" corrections satisfy it and present as a single decision. Nineteen
' slides with nineteen different current values could never group, whatever any
' column claimed about them.
'
' A declared class alone would not be safe: a mislabelled field would silently
' convert N unreviewed writes into one click. Measuring means a wrong label can
' only ever cost extra reviewing.
'
' But measurement alone is not safe either, which is why ContentKindOf gates
' it. Prose can be ACCIDENTALLY uniform -- if eleven projects all currently read
' "No key events this quarter" and all get the same rewrite, the measurement
' says "uniform batch of 11" while R13.2 says prose may never be batched. The
' measurement cannot catch that on its own, so kind decides eligibility and
' measurement decides membership. Neither is sufficient alone.
'
' Singletons never get a label. A "batch of one" is an individual review
' wearing a different word, and labelling it as a batch would inflate the count
' of things that look pre-agreed.
Public Sub AssignBatches(ByRef q As ReviewQueueSet)
    If q.Count = 0 Then Exit Sub

    Dim groupCounts As Object
    Set groupCounts = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To q.Count
        If ContentKindOf(q.Items(i).FieldID) = KIND_CONTROLLED Then
            Dim gk As String
            gk = GroupKeyOf(q.Items(i))
            If groupCounts.Exists(gk) Then
                groupCounts(gk) = CLng(groupCounts(gk)) + 1
            Else
                groupCounts(gk) = 1
            End If
        End If
    Next i

    Dim labels As Object
    Set labels = CreateObject("Scripting.Dictionary")
    Dim nextLabel As Long
    nextLabel = 0

    For i = 1 To q.Count
        If ContentKindOf(q.Items(i).FieldID) = KIND_CONTROLLED Then
            Dim k As String
            k = GroupKeyOf(q.Items(i))
            If CLng(groupCounts(k)) >= 2 Then
                If Not labels.Exists(k) Then
                    nextLabel = nextLabel + 1
                    labels(k) = "B" & nextLabel
                End If
                q.Items(i).BatchLabel = CStr(labels(k))
            End If
        End If
    Next i
End Sub

Private Function GroupKeyOf(item As ReviewItem) As String
    GroupKeyOf = item.FieldID & Chr(1) & item.CurrentValue & Chr(1) & item.ProposedValue
End Function

' Human-readable description of each batch: the transformation, the count, and
' the affected entities -- exactly the three things R13.2 asks a batch to be
' presented as.
Public Function BatchSummaryText(q As ReviewQueueSet) As String
    If q.Count = 0 Then
        BatchSummaryText = ""
        Exit Function
    End If

    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    Dim s As String

    Dim i As Long
    For i = 1 To q.Count
        If q.Items(i).BatchLabel <> "" Then
            If Not seen.Exists(q.Items(i).BatchLabel) Then
                seen(q.Items(i).BatchLabel) = True

                Dim members As String, memberCount As Long
                members = ""
                memberCount = 0
                Dim j As Long
                For j = 1 To q.Count
                    If q.Items(j).BatchLabel = q.Items(i).BatchLabel Then
                        memberCount = memberCount + 1
                        If members <> "" Then members = members & ", "
                        members = members & q.Items(j).EntityKey
                    End If
                Next j

                s = s & q.Items(i).BatchLabel & "  " & q.Items(i).FieldID & ": '" & _
                    BatchOnboardFlow.FieldPreview(q.Items(i).CurrentValue) & "' -> '" & _
                    BatchOnboardFlow.FieldPreview(q.Items(i).ProposedValue) & "'" & vbCrLf & _
                    "     " & memberCount & " entities: " & members & vbCrLf
            End If
        End If
    Next i

    BatchSummaryText = s
End Function

' ---------------------------------------------------------------------
' The grid
' ---------------------------------------------------------------------

' Writes the queue to `ws`, replacing whatever was there.
'
' Replacing rather than merging is what enforces R13.5: a new build cannot
' inherit an old build's approvals, because the old rows are gone. There is no
' code path that preserves a tick across a rebuild.
Public Sub WriteQueueSheet(ws As Object, q As ReviewQueueSet)
    ws.Cells.Clear

    ws.Cells(ROW_BANNER, 1).Value = "SYNC REVIEW -- " & q.SlideType
    ws.Cells(ROW_BANNER, 2).Value = "Run: " & q.RunStamp
    ws.Cells(ROW_BANNER, 3).Value = IIf(q.Consumed, STATE_CONSUMED, STATE_OPEN)
    ws.Cells(ROW_BANNER, 4).Value = "Put Y in the Approve column. Nothing is written until you run Apply Approved."
    ws.Rows(ROW_BANNER).Font.Bold = True

    ws.Cells(ROW_HEADER, COL_ENTITY).Value = "EntityCode"
    ws.Cells(ROW_HEADER, COL_FIELDID).Value = "FieldID"
    ws.Cells(ROW_HEADER, COL_CURRENT).Value = "Current (on the slide now)"
    ws.Cells(ROW_HEADER, COL_PROPOSED).Value = "Proposed (from the register)"
    ws.Cells(ROW_HEADER, COL_BATCH).Value = "Batch"
    ws.Cells(ROW_HEADER, COL_APPROVE).Value = "Approve (Y/N)"
    ws.Cells(ROW_HEADER, COL_HASH).Value = "Change ID (do not edit)"
    ws.Cells(ROW_HEADER, COL_DIFF).Value = "What differs"
    ws.Rows(ROW_HEADER).Font.Bold = True

    Dim i As Long
    For i = 1 To q.Count
        Dim r As Long
        r = ROW_FIRST_ITEM + i - 1
        ws.Cells(r, COL_ENTITY).Value = q.Items(i).EntityKey
        ws.Cells(r, COL_FIELDID).Value = q.Items(i).FieldID
        ' Apostrophe-prefixed so Excel never coerces a value that happens to
        ' look like a date or a number, and never reads a leading "=" as a
        ' formula. The cell is for reading, and it must show exactly what the
        ' slide holds.
        ws.Cells(r, COL_CURRENT).Value = "'" & q.Items(i).CurrentValue
        ws.Cells(r, COL_PROPOSED).Value = "'" & q.Items(i).ProposedValue
        ' NAMED, not left for the reader to spot. On 2026-08-08 two of eleven
        ' rows differed only by a trailing space and displayed as identical
        ' text -- 198 characters against 197. A reviewer cannot approve what
        ' they cannot see.
        ws.Cells(r, COL_DIFF).Value = DescribeDifference( _
            q.Items(i).CurrentValue, q.Items(i).ProposedValue)
        ws.Cells(r, COL_BATCH).Value = q.Items(i).BatchLabel
        ws.Cells(r, COL_APPROVE).Value = IIf(q.Items(i).Approved, "Y", "")
        ws.Cells(r, COL_HASH).Value = q.Items(i).ChangeHash
    Next i

    ' 8pt, matching every other sheet the tools write. Two 55-wide text
    ' columns side by side do not fit on a screen at 11pt, and comparing
    ' them is the entire job of this grid.
    ws.Cells.Font.Size = 8
    ws.Cells.VerticalAlignment = -4160        ' xlTop
    ws.Columns(COL_ENTITY).ColumnWidth = 14
    ws.Columns(COL_FIELDID).ColumnWidth = 20
    ws.Columns(COL_CURRENT).ColumnWidth = 55
    ws.Columns(COL_PROPOSED).ColumnWidth = 55
    ws.Columns(COL_DIFF).ColumnWidth = 34
    ws.Columns(COL_BATCH).ColumnWidth = 8
    ws.Columns(COL_APPROVE).ColumnWidth = 14
    ws.Columns(COL_HASH).Hidden = True
    ' Wrapped, because prose is the point. ABOUT_BODY's median is 272
    ' characters and its max 759; an unwrapped single-line cell shows the first
    ' dozen words and hides the rest behind a click, which is a review nobody
    ' completes.
    ws.Columns(COL_CURRENT).WrapText = True
    ws.Columns(COL_PROPOSED).WrapText = True
End Sub

' Reads the sheet back, with the human's ticks.
'
' Current and Proposed are deliberately NOT read from the sheet -- only the
' hash, the ticks and the identity columns are. See the module header: the
' sheet is a review surface, not an editing surface, and the apply path
' recomputes both halves from their real sources.
Public Function ReadQueueSheet(ws As Object) As ReviewQueueSet
    Dim q As ReviewQueueSet
    q.Count = 0

    Dim banner As String
    banner = CStr(ws.Cells(ROW_BANNER, 1).Value)
    If InStr(banner, "--") > 0 Then
        q.SlideType = Trim(Mid(banner, InStr(banner, "--") + 2))
    End If

    Dim stampCell As String
    stampCell = CStr(ws.Cells(ROW_BANNER, 2).Value)
    If Left(stampCell, 5) = "Run: " Then
        q.RunStamp = Mid(stampCell, 6)
    Else
        q.RunStamp = stampCell
    End If

    q.Consumed = (UCase(Trim(CStr(ws.Cells(ROW_BANNER, 3).Value))) = STATE_CONSUMED)

    ' Walks until the first blank hash rather than asking Excel for the used
    ' range: xlUp and friends are Excel's own named constants, and this module
    ' is driven from PowerPoint where they do not resolve (AGENTS.md, hit
    ' 2026-07-25 in ExcelOutput.bas).
    Dim r As Long
    r = ROW_FIRST_ITEM
    Do While Trim(CStr(ws.Cells(r, COL_HASH).Value)) <> ""
        q.Count = q.Count + 1
        ReDim Preserve q.Items(1 To q.Count)
        q.Items(q.Count).EntityKey = Trim(CStr(ws.Cells(r, COL_ENTITY).Value))
        q.Items(q.Count).FieldID = Trim(CStr(ws.Cells(r, COL_FIELDID).Value))
        q.Items(q.Count).ChangeHash = Trim(CStr(ws.Cells(r, COL_HASH).Value))
        q.Items(q.Count).BatchLabel = Trim(CStr(ws.Cells(r, COL_BATCH).Value))
        q.Items(q.Count).Approved = IsApprovalMark(CStr(ws.Cells(r, COL_APPROVE).Value))
        r = r + 1
    Loop

    PropagateBatchApprovals q

    ReadQueueSheet = q
End Function

' Anything that is not an affirmative Y is a no. Never silently approve on an
' unrecognised answer -- same rule ReadReviewGrid applies to its Include column.
Public Function IsApprovalMark(raw As String) As Boolean
    Dim v As String
    v = UCase(Trim(raw))
    IsApprovalMark = (v = "Y" Or v = "YES")
End Function

' Approving any one row of a batch approves the batch -- that is what makes it
' one decision rather than N. Applied on READ rather than on write so it
' reflects what the human actually did, not what the grid was seeded with.
Public Sub PropagateBatchApprovals(ByRef q As ReviewQueueSet)
    If q.Count = 0 Then Exit Sub

    Dim approvedBatches As Object
    Set approvedBatches = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To q.Count
        If q.Items(i).BatchLabel <> "" And q.Items(i).Approved Then
            approvedBatches(q.Items(i).BatchLabel) = True
        End If
    Next i

    For i = 1 To q.Count
        If q.Items(i).BatchLabel <> "" Then
            If approvedBatches.Exists(q.Items(i).BatchLabel) Then
                q.Items(i).Approved = True
            End If
        End If
    Next i
End Sub

' The approved change IDs, as a set the apply path can test against.
Public Function ApprovedHashSet(q As ReviewQueueSet) As Object
    Dim s As Object
    Set s = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To q.Count
        If q.Items(i).Approved Then s(q.Items(i).ChangeHash) = True
    Next i

    Set ApprovedHashSet = s
End Function

' Marks the sheet consumed. Called after a successful apply, and it is the
' second half of R13.5 -- the same approvals cannot be applied twice.
Public Sub MarkConsumed(ws As Object)
    ws.Cells(ROW_BANNER, 3).Value = STATE_CONSUMED
End Sub

' ---------------------------------------------------------------------
' Approve-all
' ---------------------------------------------------------------------

' Ticks every row.
'
' THIS IS THE LOOSENED SETTING, and it is deliberately a separate, explicitly
' named action rather than a default or an absent gate. Rohan's call as
' Research Manager, 31 July 2026: while the work runs on a carved COPY of the
' deck, the review grid is built and read in full but may be approved wholesale,
' because R13 exists to protect the real deck and a throwaway copy has nothing
' to protect. Tightening is then the removal of one button, not the building of
' a mechanism.
'
' The honest counterargument, from R13.2 itself, is that approving in bulk
' "teaches the operator to click through". That risk is real and it is why this
' is a button someone has to press by name -- so the loose mode is a decision
' taken each time, visible in the report, rather than a quiet default that
' outlives the copy it was justified by.
Public Sub ApproveAllInSheet(ws As Object)
    Dim r As Long
    r = ROW_FIRST_ITEM
    Do While Trim(CStr(ws.Cells(r, COL_HASH).Value)) <> ""
        ws.Cells(r, COL_APPROVE).Value = "Y"
        r = r + 1
    Loop
End Sub

' ---------------------------------------------------------------------
' The fast path: when a whole change set IS one decision
' ---------------------------------------------------------------------
'
' Rohan's design call, 31 July 2026, and it corrects mine. I had removed
' `Sync Now` outright on the grounds that its confirmation showed a count and
' R13 demands a before-and-after. That reasoning was sound for a count and wrong
' as a conclusion, because it ignored R13.2: a verified uniform batch is ONE
' decision, and the RM's own text says approving its members individually
' "teaches the operator to click through, which is worse than not asking".
'
' So when every queued change belongs to a uniform batch, the whole change set
' reduces to a handful of transformations -- and a dialog showing those, with
' current -> proposed and the affected entities, is not a count-based
' confirmation. It is a complete before-and-after that happens to fit in a
' modal. The 19 `In progress` -> `In Progress` corrections are exactly this: one
' line to read, not nineteen.
'
' His analogy is the right one: the same move as flattening a grouped shape to
' its leaves and then addressing the group as a single thing.
'
' R13.4 is not violated, because R13.4 is about PROSE -- "prose review is real
' reading", and its requirement is that you can leave and come back. There is
' nothing to come back to in three lines of uniform transformation. The moment
' one individual decision exists, this path is refused and the worksheet is the
' only way through.


Public Function IndividualCount(q As ReviewQueueSet) As Long
    Dim n As Long, i As Long
    For i = 1 To q.Count
        If q.Items(i).BatchLabel = "" Then n = n + 1
    Next i
    IndividualCount = n
End Function

Public Function DistinctBatchCount(q As ReviewQueueSet) As Long
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 1 To q.Count
        If q.Items(i).BatchLabel <> "" Then seen(q.Items(i).BatchLabel) = True
    Next i
    DistinctBatchCount = seen.Count
End Function

' Is there batchable work that can honestly be approved in one modal?
'
' PARTIAL IS ALLOWED, and this is Rohan's correction to my rule, 31 July 2026.
' I originally required the WHOLE change set to be batchable -- one prose row
' disqualified the entire run and sent everything to the worksheet. Checked
' against R13, that strictness is mine and not the RM's: the batched changes are
' still shown as current -> proposed before writing (R13.1), only uniform groups
' are batched (R13.2), and the individual ones still go to the worksheet where
' prose belongs (R13.4). Nothing in R13 requires a run to be all-or-nothing.
'
' So a mixed run does the uniform part on one confirmation and sends the
' remainder to the grid. What that buys is real: a quarterly run is typically a
' large controlled-field correction plus a handful of genuine edits, and the
' old rule made the large easy part wait behind the small hard one.
'
' Two conditions now:
'   - there is at least one uniform batch to approve
'   - the batches fit on a screen
Public Function HasBatchableWork(q As ReviewQueueSet) As Boolean
    Dim n As Long
    n = DistinctBatchCount(q)
    If n = 0 Then Exit Function
    If n > MAX_BATCHES_IN_MODAL Then Exit Function
    HasBatchableWork = True
End Function

' Approves ONLY the batched rows. The individual ones stay unapproved and are
' picked up by the review sheet afterwards.
'
' Deliberately not ApproveAllInMemory with a filter at the call site -- the
' distinction between "approved because it was one visible decision" and
' "approved because someone ticked it" is the whole of R13, and it should not
' depend on a caller remembering to pass the right flag.
Public Sub ApproveBatchedOnly(ByRef q As ReviewQueueSet)
    Dim i As Long
    For i = 1 To q.Count
        q.Items(i).Approved = (q.Items(i).BatchLabel <> "")
    Next i
End Sub

' The fast path's confirmation. Pure, so the wording is testable without a live
' deck or a dialog -- same reason RunSync.ConfirmSyncText was pure.
'
' Shows every transformation in full. It must never degrade to a count, because
' the count is precisely what R13 rejected.
Public Function ConfirmBatchText(q As ReviewQueueSet) As String
    Dim batchedRows As Long
    batchedRows = q.Count - IndividualCount(q)

    Dim s As String
    s = "This will change " & batchedRows & " slide field(s) in " & DistinctBatchCount(q) & _
        " uniform change(s)." & vbCrLf & vbCrLf & _
        "Every change below is the SAME transformation applied to several slides:" & vbCrLf & vbCrLf & _
        BatchSummaryText(q) & vbCrLf & _
        "A backup is taken first, and each change is re-checked against its slide" & vbCrLf & _
        "immediately before writing." & vbCrLf

    ' STATED BEFORE THE QUESTION, not after the run.
    '
    ' A partial run must not read as a whole one. Someone who clicks Yes here
    ' believing the deck is now finished would walk away with the prose changes
    ' silently unapplied -- and "I pressed Sync Now" would be a true statement
    ' about an incomplete deck. The remainder is part of what is being agreed to.
    If IndividualCount(q) > 0 Then
        s = s & vbCrLf & IndividualCount(q) & " further change(s) are NOT covered by this and are not" & vbCrLf & _
            "being applied -- they need reading one at a time. The review sheet" & vbCrLf & _
            "opens straight afterwards." & vbCrLf
    End If

    s = s & vbCrLf & "Apply the " & DistinctBatchCount(q) & " uniform change(s) above?"
    ConfirmBatchText = s
End Function

' Why the fast path was refused, in the words a human needs to act on it.
Public Function FastPathRefusalText(q As ReviewQueueSet) As String
    If q.Count = 0 Then
        FastPathRefusalText = "Nothing to sync -- every linked slide already matches the workbook."
    ElseIf IndividualCount(q) > 0 Then
        FastPathRefusalText = IndividualCount(q) & " change(s) are not part of a uniform batch and" & vbCrLf & _
            "have to be read one at a time -- prose, or a one-off correction." & vbCrLf & vbCrLf & _
            "Opening the review sheet instead."
    Else
        FastPathRefusalText = DistinctBatchCount(q) & " separate transformations is too many to read in a" & vbCrLf & _
            "dialog box without skimming." & vbCrLf & vbCrLf & "Opening the review sheet instead."
    End If
End Function

Public Sub ApproveAllInMemory(ByRef q As ReviewQueueSet)
    Dim i As Long
    For i = 1 To q.Count
        q.Items(i).Approved = True
    Next i
End Sub

' ---------------------------------------------------------------------
' Apply -- the only part of R13 that writes
' ---------------------------------------------------------------------

' Takes a local backup before any write batch.
'
' TURNING A DISCIPLINE INTO A MECHANISM. "Always work on a copy" has been a rule
' Rohan remembers, and a rule that depends on remembering is not a control. This
' makes the tool take its own .bak, so the protection survives a tired evening.
'
' Cloud-hosted decks report Presentation.FullName as a URL rather than a path
' (found 2026-07-28, item 4), and SaveCopyAs cannot write a sibling file next to
' a URL. That case WARNS LOUDLY and proceeds rather than refusing: the real
' protection is that the original deck is never opened for write in the first
' place, and refusing here would block cloud decks entirely while adding nothing
' that rule does not already provide. The warning goes in the report so an
' unbacked run is visible rather than assumed.
Public Function BackupBeforeWrite(pres As Object, ByRef backupPath As String) As String
    backupPath = ""

    Dim fullName As String
    On Error Resume Next
    fullName = pres.fullName
    If Err.Number <> 0 Then
        On Error GoTo 0
        BackupBeforeWrite = "WARNING: could not read this deck's path -- NO BACKUP TAKEN."
        Exit Function
    End If
    On Error GoTo 0

    If InStr(fullName, "://") > 0 Then
        BackupBeforeWrite = "WARNING: this deck is cloud-hosted (" & fullName & ")." & vbCrLf & _
            "         NO LOCAL BACKUP COULD BE TAKEN. Office version history is the only undo."
        Exit Function
    End If

    Dim candidate As String
    candidate = fullName & ".r13-" & Format(Now, "yyyymmdd-hhnnss") & ".bak.pptx"

    On Error Resume Next
    pres.SaveCopyAs candidate
    If Err.Number <> 0 Then
        Dim e As String
        e = Err.Description
        On Error GoTo 0
        BackupBeforeWrite = "WARNING: backup failed (" & e & ") -- NO BACKUP TAKEN."
        Exit Function
    End If
    On Error GoTo 0

    backupPath = candidate
    BackupBeforeWrite = "Backup: " & candidate
End Function

' Applies exactly the approved changes and nothing else.
'
' THE REVALIDATION IS THE POINT, not a defensive extra. Every row's hash is
' recomputed here from the LIVE slide text and the REGISTER's current value, and
' compared against the hash the human approved against. Three things that would
' otherwise write silently are caught by that one comparison:
'
'   - the slide was hand-edited between review and apply
'   - the register value changed between review and apply
'   - the sheet's Current/Proposed cells were edited (they are ignored entirely;
'     both halves come from their real sources, never from the grid)
'
' A row that fails revalidation is DROPPED and reported, never applied. A stale
' approval can therefore never overwrite a hand edit -- it can only fail to fire.
'
' `logWs` receives one appended line per attempted change, written BEFORE the
' next change is attempted, so a run that dies at slide 12 of 19 leaves a record
' of exactly what landed. Pass Nothing to skip logging.
Public Function ApplyApproved(sheet As Sheet, slideType As String, ws As Object, _
                              logWs As Object) As String
    Dim q As ReviewQueueSet
    q = ReadQueueSheet(ws)

    Dim report As String
    report = "=== Apply Approved: " & slideType & " ===" & vbCrLf & _
             "Run: " & q.RunStamp & vbCrLf

    ' R13.5, second half. A consumed sheet's approvals have already been spent.
    If q.Consumed Then
        ApplyApproved = report & vbCrLf & _
            "REFUSED: this review has already been applied." & vbCrLf & _
            "Run 'Review Changes' again to build a fresh queue." & vbCrLf
        Exit Function
    End If

    If q.Count = 0 Then
        ApplyApproved = report & "Nothing in the queue." & vbCrLf
        Exit Function
    End If

    Dim approved As Object
    Set approved = ApprovedHashSet(q)
    If approved.Count = 0 Then
        ApplyApproved = report & q.Count & " change(s) queued, none approved. Nothing written." & vbCrLf
        Exit Function
    End If

    ' Backup BEFORE the first write, not after the plan looks good.
    Dim backupPath As String
    report = report & BackupBeforeWrite(Application.ActivePresentation, backupPath) & vbCrLf & vbCrLf

    ' key -> live slide, the same resolve-and-index walk ResequenceByRowOrder does.
    Dim instances() As Object
    instances = RunSync.GatherInstances(slideType)

    Dim lo As Long, hi As Long, hasAny As Boolean
    On Error Resume Next
    lo = LBound(instances)
    hi = UBound(instances)
    hasAny = (Err.Number = 0)
    On Error GoTo 0

    Dim byKey As Object
    Set byKey = CreateObject("Scripting.Dictionary")
    If hasAny Then
        Dim i As Long
        For i = lo To hi
            Dim resolved As SlideInstance
            resolved = Resolve.ResolveSlideInstance(instances(i))
            If resolved.HasInstanceKey Then Set byKey(resolved.InstanceKey) = instances(i)
        Next i
    End If

    Dim writtenCount As Long, skippedCount As Long, staleCount As Long, failedCount As Long

    Dim n As Long
    For n = 1 To q.Count
        If Not q.Items(n).Approved Then
            skippedCount = skippedCount + 1
        ElseIf Not byKey.Exists(q.Items(n).EntityKey) Then
            staleCount = staleCount + 1
            report = report & "  DROPPED " & q.Items(n).EntityKey & "/" & q.Items(n).FieldID & _
                " -- no slide in this deck carries that key any more" & vbCrLf
            AppendLogLine logWs, q.RunStamp, q.Items(n), "dropped: slide missing"
        Else
            Dim sld As Object
            Set sld = byKey(q.Items(n).EntityKey)

            ' The register's value NOW, not the grid's copy of it.
            Dim proposed As String
            Dim haveProposed As Boolean
            haveProposed = False
            If sheet.Rows.Exists(q.Items(n).EntityKey) Then
                Dim rowValues As Object
                Set rowValues = sheet.Rows(q.Items(n).EntityKey)
                If rowValues.Exists(q.Items(n).FieldID) Then
                    proposed = CStr(rowValues(q.Items(n).FieldID))
                    haveProposed = True
                End If
            End If

            If Not haveProposed Then
                staleCount = staleCount + 1
                report = report & "  DROPPED " & q.Items(n).EntityKey & "/" & q.Items(n).FieldID & _
                    " -- the register no longer has a value for this field" & vbCrLf
                AppendLogLine logWs, q.RunStamp, q.Items(n), "dropped: register row gone"
            Else
                ' Dry inject reads the slide's current text without touching it.
                Dim probe As InjectResult
                probe = InjectPrimitive.InjectPrimitive(sld, q.Items(n).FieldID, proposed, True)

                Dim liveHash As String
                liveHash = ChangeHash(q.Items(n).EntityKey, q.Items(n).FieldID, _
                                      probe.CurrentValue, proposed)

                If liveHash <> q.Items(n).ChangeHash Then
                    staleCount = staleCount + 1
                    report = report & "  DROPPED " & q.Items(n).EntityKey & "/" & q.Items(n).FieldID & _
                        " -- changed since you approved it; re-review" & vbCrLf
                    AppendLogLine logWs, q.RunStamp, q.Items(n), "dropped: changed since approval"
                ElseIf Not probe.Found Then
                    failedCount = failedCount + 1
                    report = report & "  FAILED " & q.Items(n).EntityKey & "/" & q.Items(n).FieldID & _
                        " -- " & probe.ErrorMessage & vbCrLf
                    AppendLogLine logWs, q.RunStamp, q.Items(n), "failed: " & probe.ErrorMessage
                Else
                    Dim wrote As InjectResult
                    wrote = InjectPrimitive.InjectPrimitive(sld, q.Items(n).FieldID, proposed, False)
                    If wrote.Verified Then
                        writtenCount = writtenCount + 1
                        report = report & "  written: " & q.Items(n).EntityKey & "/" & q.Items(n).FieldID & vbCrLf
                        AppendLogLine logWs, q.RunStamp, q.Items(n), "written"
                    Else
                        failedCount = failedCount + 1
                        report = report & "  FAILED " & q.Items(n).EntityKey & "/" & q.Items(n).FieldID & _
                            " -- " & wrote.ErrorMessage & vbCrLf
                        AppendLogLine logWs, q.RunStamp, q.Items(n), "failed: " & wrote.ErrorMessage
                    End If
                End If
            End If
        End If
    Next n

    MarkConsumed ws

    report = report & vbCrLf & "Summary: " & writtenCount & " written, " & _
        skippedCount & " not approved, " & staleCount & " dropped as stale, " & _
        failedCount & " failed" & vbCrLf

    If staleCount > 0 Then
        report = report & vbCrLf & "Dropped changes were NOT written. Run 'Review Changes'" & vbCrLf & _
            "again to see them with their current before-and-after." & vbCrLf
    End If

    ApplyApproved = report
End Function

' One line per attempted change, appended as it happens.
'
' Written BEFORE the next change is attempted rather than batched at the end,
' which is the whole value of it: PowerPoint dying at slide 12 of 19 leaves a
' record of the 11 that landed. A log assembled in memory and written at the end
' is exactly zero use in the only scenario it exists for.
Public Sub AppendLogLine(logWs As Object, runStamp As String, item As ReviewItem, outcome As String)
    If logWs Is Nothing Then Exit Sub

    On Error Resume Next
    If Trim(CStr(logWs.Cells(1, 1).Value)) = "" Then
        logWs.Cells(1, 1).Value = "When"
        logWs.Cells(1, 2).Value = "Run"
        logWs.Cells(1, 3).Value = "EntityCode"
        logWs.Cells(1, 4).Value = "FieldID"
        logWs.Cells(1, 5).Value = "Outcome"
        logWs.Cells(1, 6).Value = "Change ID"
        logWs.Rows(1).Font.Bold = True
    End If

    Dim r As Long
    r = 2
    Do While Trim(CStr(logWs.Cells(r, 1).Value)) <> ""
        r = r + 1
    Loop

    logWs.Cells(r, 1).Value = Format(Now, "yyyy-mm-dd hh:nn:ss")
    logWs.Cells(r, 2).Value = runStamp
    logWs.Cells(r, 3).Value = item.EntityKey
    logWs.Cells(r, 4).Value = item.FieldID
    logWs.Cells(r, 5).Value = outcome
    logWs.Cells(r, 6).Value = item.ChangeHash
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------
' Reporting
' ---------------------------------------------------------------------

' What is waiting to be reviewed, for the human who just pressed the button.
Public Function QueueSummaryText(q As ReviewQueueSet) As String
    If q.Count = 0 Then
        QueueSummaryText = "Nothing to review -- this deck already matches the register." & vbCrLf
        Exit Function
    End If

    Dim batched As Long, individual As Long
    Dim i As Long
    For i = 1 To q.Count
        If q.Items(i).BatchLabel <> "" Then batched = batched + 1 Else individual = individual + 1
    Next i

    Dim s As String
    s = q.Count & " change(s) need review." & vbCrLf & _
        "    " & individual & " to review one at a time" & vbCrLf & _
        "    " & batched & " grouped into uniform batches" & vbCrLf & vbCrLf

    If batched > 0 Then
        s = s & "Batches:" & vbCrLf & BatchSummaryText(q) & vbCrLf
    End If

    s = s & "Nothing has been written. Review the '" & REVIEW_SHEET_NAME & _
        "' sheet, put Y against what you approve, then run Apply Approved." & vbCrLf

    QueueSummaryText = s
End Function
