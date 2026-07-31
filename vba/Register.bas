Attribute VB_Name = "Register"
Option Explicit

' V3: reads the Excel Control Layer's FIELD REGISTER -- the long-format table
' that replaces the wide Data sheet.
'
'   wide (ExcelOutput.ReadSheet):  one row per instance, one COLUMN per field
'   long (here):                   one row per Quarter x EntityCode x FieldID
'
' Long is the Excel side's call and it is right: adding a field adds rows, not
' columns, so the schema stays fixed as the deck grows. The audit measured why
' it matters -- the real slide carries 77 text items, ~38 of them plausible
' fields, and a 38-column sheet was never going to be maintainable.
'
' A NEW MODULE rather than a rewrite of ExcelOutput, for three reasons:
'   1. the register is read-only by R3, where ExcelOutput is read-write --
'      different contracts should not share a module
'   2. the wide reader is still live and stays live until migration completes;
'      deleting it now would strand every deck that has not moved
'   3. blast radius. Everything downstream consumes the `Sheet` UDT, and this
'      returns exactly that UDT -- so SyncOperations, RunSync, the preview and
'      the planner are all completely unchanged by the format switch. The
'      format changes how the sheet is READ, not what the sync consumes.
'
' Columns are located BY HEADER NAME, never by position. The Excel side is
' authoring this as a ListObject and Copilot is explicitly expected to insert
' rows and columns (round 3 §2.4), so any position assumption has a shelf life.

' --- Column headers ---------------------------------------------------
' Per the round 5 revised E4 list. `SlideID` is deliberately absent: it was
' withdrawn after the probe showed it is reassigned on within-deck duplicate
' and preserved on cross-deck paste, making it unreliable as an identifier --
' and it was redundant anyway, since the join resolves entirely from the deck's
' own tags.
Public Const COL_QUARTER As String = "Quarter"
Public Const COL_ENTITY As String = "EntityCode"
Public Const COL_SLIDETYPE As String = "SlideType"
Public Const COL_FIELDID As String = "FieldID"
Public Const COL_FIELDTYPE As String = "FieldType"
Public Const COL_VALUE As String = "Value"
Public Const COL_STATUS As String = "Status"

' --- Magic values ------------------------------------------------------
' UNCONFIRMED pending F2 and F4 of round 6 -- the exact literals were asked for
' and have not come back. Declared as constants precisely so that when they do,
' this is a one-line change rather than a hunt. All comparisons below are
' case-insensitive, so only the spelling is at risk, not the casing.
Public Const STATUS_APPROVED As String = "Approved"

' SEEDING IS NOT APPROVING, and until 2026-07-31 the Status column could not
' tell the difference.
'
' All 46 ABOUT_BODY rows in the real register read "Approved". None of them had
' been read and agreed by anyone -- they were copied OFF THE SLIDES so the
' column would not be empty, exactly as the field package specifies for
' seeding. Copying something is not approving it.
'
' It was harmless only while seed and slide were identical. The moment drafted
' text arrives, "Approved" becomes the thing deciding whether words reach a
' slide, and a row copied off that same slide is indistinguishable from one a
' human genuinely read. There is no way to tell them apart after the fact, so
' the distinction has to exist BEFORE the first drafting round.
'
' Only Approved is writable. Seed and Draft are recognised and refused; an
' unrecognised value is refused AND reported loudly, because that is a typo in
' the one column standing between a draft and a real slide.
Public Const STATUS_SEED As String = "Seed"      ' copied from the slide -- never writable
Public Const STATUS_DRAFT As String = "Draft"    ' synthesised, not yet agreed -- never writable

Public Const QUARTER_ALL As String = "ALL"

' What a read produced, and -- as importantly -- what it discarded.
'
' The diagnostics are not instrumentation. They exist because of a specific
' trap identified in round 6 §2: an unmatched period literal returns zero rows,
' which is INDISTINGUISHABLE from a genuinely empty quarter. Both mean "nothing
' to sync", and one of them is a typo. Reporting what was filtered and which
' periods actually exist in the register is what lets a caller tell a mismatch
' from an empty quarter, so the failure is loud instead of silent.
Public Type RegisterRead
    Data As Sheet              ' the SAME UDT the wide reader returns -- nothing downstream changes
    RowsSeen As Long           ' data rows in the register, before any filter
    RejectedStatus As Long     ' dropped because Status <> Approved (total)
    RejectedSeed As Long       ' ...of which: Seed, a value copied off the slide
    RejectedDraft As Long      ' ...of which: Draft, synthesised but not agreed
    RejectedUnknownStatus As Long  ' ...of which: not a recognised status at all -- a typo
    RejectedPeriod As Long     ' dropped because Quarter matched neither the deck period nor ALL
    RejectedType As Long       ' dropped because SlideType is a different type
    Accepted As Long
    AcceptedPeriod As Long     ' accepted because Quarter matched the DECK PERIOD
    AcceptedStatic As Long     ' accepted because Quarter = ALL -- these match ANY period
    PeriodsPresent As String   ' distinct Quarter values found, comma-joined
    MissingColumns As String   ' required headers not found -- "" when the shape is right
End Type

' Reads `ws` as a field register, filtered to one deck period and one slide
' type, and returns it in the shape the sync already consumes.
'
' `deckPeriod` comes from the deck's own custom property (D4: the deck declares
' its own period), not from an operator prompt.
Public Function ReadRegister(ws As Object, deckPeriod As String, slideType As String) As RegisterRead
    Dim result As RegisterRead

    Set result.Data.Rows = CreateObject("Scripting.Dictionary")
    Set result.Data.Fields = New Collection
    Set result.Data.InstanceOrder = New Collection

    ' --- Locate columns by header name ---------------------------------
    Dim lastCol As Long, lastRow As Long
    lastCol = ws.Cells(1, ws.Columns.count).End(-4159).Column   ' xlToLeft
    lastRow = ws.Cells(ws.Rows.count, 1).End(-4162).Row         ' xlUp

    Dim cQuarter As Long, cEntity As Long, cSlideType As Long
    Dim cFieldId As Long, cValue As Long, cStatus As Long
    Dim i As Long
    For i = 1 To lastCol
        Select Case Trim(CStr(ws.Cells(1, i).Value))
            Case COL_QUARTER:   cQuarter = i
            Case COL_ENTITY:    cEntity = i
            Case COL_SLIDETYPE: cSlideType = i
            Case COL_FIELDID:   cFieldId = i
            Case COL_VALUE:     cValue = i
            Case COL_STATUS:    cStatus = i
        End Select
    Next i

    ' Missing columns are reported, never worked around. A register missing
    ' Status would otherwise be read as "everything is approved", which is the
    ' single worst way this could fail -- it would publish drafts.
    If cQuarter = 0 Then result.MissingColumns = result.MissingColumns & COL_QUARTER & " "
    If cEntity = 0 Then result.MissingColumns = result.MissingColumns & COL_ENTITY & " "
    If cFieldId = 0 Then result.MissingColumns = result.MissingColumns & COL_FIELDID & " "
    If cValue = 0 Then result.MissingColumns = result.MissingColumns & COL_VALUE & " "
    If cStatus = 0 Then result.MissingColumns = result.MissingColumns & COL_STATUS & " "
    If result.MissingColumns <> "" Then
        result.MissingColumns = Trim(result.MissingColumns)
        ReadRegister = result
        Exit Function
    End If

    ' --- Walk the rows --------------------------------------------------
    Dim periods As Object
    Set periods = CreateObject("Scripting.Dictionary")

    Dim seenFields As Object
    Set seenFields = CreateObject("Scripting.Dictionary")

    Dim r As Long
    For r = 2 To lastRow
        Dim entity As String, fieldId As String
        entity = Trim(CStr(ws.Cells(r, cEntity).Value))
        fieldId = Trim(CStr(ws.Cells(r, cFieldId).Value))

        ' A row with no entity or no field is not a row -- blank spacer lines
        ' are normal in a hand-maintained sheet and must not be counted as
        ' rejections, or the diagnostics stop meaning anything.
        If entity <> "" And fieldId <> "" Then
            result.RowsSeen = result.RowsSeen + 1

            Dim q As String
            q = Trim(CStr(ws.Cells(r, cQuarter).Value))
            If q <> "" Then periods(q) = True

            Dim rowType As String
            rowType = ""
            If cSlideType > 0 Then rowType = Trim(CStr(ws.Cells(r, cSlideType).Value))

            ' Filter order matters for the diagnostics, not for the result:
            ' each row is attributed to the FIRST reason it was dropped, so the
            ' counts sum to RowsSeen and a human can read them as a funnel.
            If cSlideType > 0 And rowType <> "" And StrComp(rowType, slideType, vbTextCompare) <> 0 Then
                result.RejectedType = result.RejectedType + 1
            ElseIf StrComp(Trim(CStr(ws.Cells(r, cStatus).Value)), STATUS_APPROVED, vbTextCompare) <> 0 Then
                result.RejectedStatus = result.RejectedStatus + 1
                ' WHICH kind of not-approved. "12 rows are seed values" and
                ' "12 rows have a status nobody recognises" are the same number
                ' and completely different situations -- one is the system
                ' working, the other is a typo in the column that decides
                ' whether words reach a slide.
                Dim statusRaw As String
                statusRaw = Trim(CStr(ws.Cells(r, cStatus).Value))
                If StrComp(statusRaw, STATUS_SEED, vbTextCompare) = 0 Then
                    result.RejectedSeed = result.RejectedSeed + 1
                ElseIf StrComp(statusRaw, STATUS_DRAFT, vbTextCompare) = 0 Then
                    result.RejectedDraft = result.RejectedDraft + 1
                Else
                    result.RejectedUnknownStatus = result.RejectedUnknownStatus + 1
                End If
            ElseIf StrComp(q, deckPeriod, vbTextCompare) <> 0 And StrComp(q, QUARTER_ALL, vbTextCompare) <> 0 Then
                ' The entity-static case: Quarter = ALL rows carry forward into
                ' every period. Round 5 §3 -- a field like PROJECT_NAME does not
                ' change quarterly but is still managed, and leaving it untagged
                ' would mean every created slide keeps the template's value.
                result.RejectedPeriod = result.RejectedPeriod + 1
            Else
                result.Accepted = result.Accepted + 1
                ' Split, because a mistyped period does NOT produce zero rows:
                ' the Quarter = ALL rows match any period at all, so a typo
                ' yields a PARTIAL sync -- entity-static fields update, every
                ' quarterly field silently does not. That reads as success.
                ' Found by the test 2026-07-31; the first version only warned
                ' when nothing at all was accepted, which never happened.
                If StrComp(q, QUARTER_ALL, vbTextCompare) = 0 Then
                    result.AcceptedStatic = result.AcceptedStatic + 1
                Else
                    result.AcceptedPeriod = result.AcceptedPeriod + 1
                End If

                If Not result.Data.Rows.Exists(entity) Then
                    Dim fresh As Object
                    Set fresh = CreateObject("Scripting.Dictionary")
                    Set result.Data.Rows(entity) = fresh
                    result.Data.InstanceOrder.Add entity
                End If

                result.Data.Rows(entity)(fieldId) = CStr(ws.Cells(r, cValue).Value)

                If Not seenFields.Exists(fieldId) Then
                    seenFields(fieldId) = True
                    result.Data.Fields.Add fieldId
                End If
            End If
        End If
    Next r

    Dim k As Variant
    For Each k In periods.Keys
        result.PeriodsPresent = result.PeriodsPresent & IIf(result.PeriodsPresent = "", "", ", ") & k
    Next k

    ReadRegister = result
End Function

' The line a human needs when a read comes back empty.
'
' Pure, so the wording is testable. This exists for one failure: a period
' literal that does not match returns zero rows and reads as "nothing to sync".
' Saying WHICH periods the register actually contains turns a silent no-op into
' an obvious typo.
Public Function ReadDiagnostic(r As RegisterRead, deckPeriod As String) As String
    If r.MissingColumns <> "" Then
        ReadDiagnostic = "The register is missing required column(s): " & r.MissingColumns & _
            vbCrLf & "Nothing was read. Check the header row."
        Exit Function
    End If

    ' Fires on ZERO PERIOD ROWS, not on zero rows. Those differ, and the
    ' difference is the whole point: Quarter = ALL rows match any period, so a
    ' mistyped period still accepts the entity-static ones. Testing Accepted > 0
    ' would therefore stay silent through exactly the failure this guards --
    ' a partial sync that looks like a successful one.
    ' An UNRECOGNISED status is reported even on an otherwise healthy read, and
    ' before the success line rather than after it.
    '
    ' Every other rejection here is the system working: a Seed row is meant to
    ' be refused, a Draft row is meant to be refused, a different period is
    ' meant to be refused. An unrecognised value is none of those -- it is a
    ' typo in the single column that decides whether words reach a slide, and
    ' its row is silently absent from the sync. Returning early on
    ' AcceptedPeriod > 0 would hide it behind exactly the runs that look fine.
    Dim warn As String
    If r.RejectedUnknownStatus > 0 Then
        warn = "WARNING: " & r.RejectedUnknownStatus & " row(s) have a Status that is not " & _
            STATUS_APPROVED & ", " & STATUS_SEED & " or " & STATUS_DRAFT & "." & vbCrLf & _
            "Those rows were SKIPPED. A misspelt status is indistinguishable from" & vbCrLf & _
            "a deliberate hold, and neither reaches a slide." & vbCrLf & vbCrLf
    End If

    ' Held-back seed rows are stated in BOTH exit paths, not just the healthy
    ' one. Built here, above the branch, for that reason: the first version put
    ' the sentence inside the AcceptedPeriod > 0 arm only, so a register whose
    ' accepted rows were all entity-static reported the seed count as a number
    ' in a list and never said what it meant. Caught by its own test.
    Dim held As String
    If r.RejectedSeed > 0 Then
        held = r.RejectedSeed & " seed row(s) held back -- seeding is not approving." & vbCrLf
    End If

    If r.AcceptedPeriod > 0 Then
        ReadDiagnostic = warn & r.Accepted & " row(s) accepted for period '" & deckPeriod & "'" & _
            IIf(r.AcceptedStatic > 0, " (" & r.AcceptedStatic & " of them entity-static)", "") & "." & _
            IIf(held <> "", vbCrLf & held, "")
        Exit Function
    End If

    Dim s As String
    s = warn & held & "NO ROWS matched this deck's period '" & deckPeriod & "'." & vbCrLf & vbCrLf & _
        "  rows in the register:   " & r.RowsSeen & vbCrLf & _
        "  wrong slide type:       " & r.RejectedType & vbCrLf & _
        "  not Approved:           " & r.RejectedStatus & _
            IIf(r.RejectedStatus > 0, "  (seed " & r.RejectedSeed & ", draft " & r.RejectedDraft & _
                ", UNRECOGNISED " & r.RejectedUnknownStatus & ")", "") & vbCrLf & _
        "  different period:       " & r.RejectedPeriod & vbCrLf & _
        "  entity-static accepted: " & r.AcceptedStatic & vbCrLf

    If r.AcceptedStatic > 0 Then
        s = s & vbCrLf & "WARNING: " & r.AcceptedStatic & " entity-static row(s) DID match, because" & vbCrLf & _
            "Quarter = " & QUARTER_ALL & " matches any period. So a sync would still write" & vbCrLf & _
            "those fields while silently skipping every quarterly one -- a partial" & vbCrLf & _
            "update that looks like a successful run." & vbCrLf
    End If

    If r.PeriodsPresent <> "" Then
        s = s & vbCrLf & "Periods present in the register: " & r.PeriodsPresent & vbCrLf
        If r.RejectedPeriod > 0 Then
            s = s & vbCrLf & "If one of those is the period you meant, the deck's declared" & vbCrLf & _
                "period does not match it -- that is a spelling mismatch, not an" & vbCrLf & _
                "empty quarter. They are indistinguishable by row count alone," & vbCrLf & _
                "which is why this message exists."
        End If
    End If

    ReadDiagnostic = s
End Function
