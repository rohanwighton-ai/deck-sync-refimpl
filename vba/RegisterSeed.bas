Attribute VB_Name = "RegisterSeed"
Option Explicit

' Creates the register rows for fields that have been marked and onboarded.
'
' WHY THIS EXISTS: nothing in this tool has ever written a register row. Adding
' one was a manual Excel job -- ~33 fields x 14 projects is around 460 rows for
' the real deck, hand-built, after the marking work is already done. Meanwhile
' the marking session ALREADY holds the answer to the only question that decides
' a row's Quarter, and read it back for nothing:
'
'     BatchOnboardFlow.bas:129 -- "human-declared hint only, not wired into
'     sync behavior yet"
'
' Same shape as the two fixes before it. Instance keys went from 45 prompts to
' one grid; marking went from three dialogs per field to one grid. Both times
' the answer had already been given and the tool simply did not use it.
'
' THE QUARTER COLUMN IS THE ONLY PLACE CADENCE IS REAL.
'
' A field marked "static" gets ONE row, Quarter = ALL, which matches every
' period (Register.bas:232). A field marked "variable" gets one row per period,
' stamped with the deck's own period. That is the whole mapping, and it is why
' the marking answer is worth carrying: it is the difference between typing
' PROJECT_NAME once and typing it again every quarter forever.
'
' Do not confuse this with FieldSpec.Kind ("Controlled/Prose/Static"). Kind is
' about HOW a value is produced -- picked from a vocabulary, written by a
' person, or fed from elsewhere -- and its "Static" switches OFF drafting for
' the field. Cadence is about WHEN a value applies. A project description is
' prose AND static; the register rig has ABOUT_BODY on Quarter = ALL for all 43
' entities right now while it is also the flagship drafting field.

Public Const CADENCE_STATIC As String = "static"
Public Const CADENCE_VARIABLE As String = "variable"

' One row this seeding run wants to add. Kept as a type rather than a
' Dictionary so a missing member is a compile error, not a silent Empty.
Public Type SeedRowPlan
    EntityCode As String
    FieldID As String
    Quarter As String
    Reason As String            ' populated only when Skip is True
    Skip As Boolean
End Type

Public Type SeedPlan
    Rows() As SeedRowPlan
    ToAdd As Long
    SkippedExisting As Long
    SkippedCadenceClash As Long
End Type

' Decides WHICH register rows are needed. Pure: no Office, no worksheet, no
' harvesting -- so the decision can be tested without driving PowerPoint, which
' is the only reason any of this project's rules ever get asserted twice.
'
' `entities`   ordered Collection of instance keys (one per project slide)
' `fields`     ordered Collection of FieldIDs that were marked
' `cadence`    Dictionary fieldId -> "static"|"variable"
' `existing`   Dictionary of already-present rows, keyed by RowKey() below
' `deckPeriod` the deck's own declared period, e.g. "FY26Q4"
Public Function PlanSeedRows(entities As Collection, fields As Collection, _
                             cadence As Object, existing As Object, _
                             deckPeriod As String) As SeedPlan
    Dim result As SeedPlan

    If deckPeriod = "" Then
        Err.Raise vbObjectError + 1, "RegisterSeed.PlanSeedRows", _
            "the deck does not declare a period, so a variable field has no quarter to be stamped with"
    End If

    Dim n As Long
    n = entities.count * fields.count
    If n = 0 Then
        result.ToAdd = 0
        PlanSeedRows = result
        Exit Function
    End If
    ReDim result.Rows(1 To n)

    Dim i As Long
    i = 0

    Dim e As Variant, f As Variant
    For Each e In entities
        For Each f In fields
            i = i + 1
            Dim ent As String, fid As String
            ent = CStr(e): fid = CStr(f)

            ' UNKNOWN CADENCE IS TREATED AS VARIABLE, matching
            ' NormalizeFieldVolatility's own default. Variable is the safe
            ' guess: a per-period row is merely more typing if it was really
            ' static, whereas a wrong ALL row silently carries one quarter's
            ' value into every quarter that follows.
            Dim cad As String
            cad = CADENCE_VARIABLE
            If Not cadence Is Nothing Then
                If cadence.Exists(fid) Then cad = LCase$(Trim$(CStr(cadence(fid))))
            End If
            If cad <> CADENCE_STATIC Then cad = CADENCE_VARIABLE

            Dim q As String
            If cad = CADENCE_STATIC Then q = Register.QUARTER_ALL Else q = deckPeriod

            result.Rows(i).EntityCode = ent
            result.Rows(i).FieldID = fid
            result.Rows(i).Quarter = q

            If existing.Exists(RowKey(ent, fid, q)) Then
                result.Rows(i).Skip = True
                result.Rows(i).Reason = "already has a " & q & " row"
                result.SkippedExisting = result.SkippedExisting + 1

            ' A STATIC ROW MUST NOT BE ADDED BESIDE AN EXISTING PERIOD ROW.
            '
            ' Both would pass every filter in ReadRegisterCore, which counts
            ' that as a CadenceCollision and resolves it by "the period row
            ' wins". Resolved is not the same as harmless: the ALL row then sits
            ' there permanently shadowed, and the next period it silently
            ' becomes the value. Refuse, and say which row is in the way.
            ElseIf cad = CADENCE_STATIC And existing.Exists(RowKey(ent, fid, deckPeriod)) Then
                result.Rows(i).Skip = True
                result.Rows(i).Reason = "marked static, but a " & deckPeriod & _
                    " row already exists -- adding an ALL row beside it would be a cadence clash"
                result.SkippedCadenceClash = result.SkippedCadenceClash + 1

            Else
                result.ToAdd = result.ToAdd + 1
            End If
        Next f
    Next e

    PlanSeedRows = result
End Function

' The register's identity for a row. Entity + Field + Quarter, because the same
' entity and field legitimately appear once per period.
'
' Chr(1) as the separator, not "|": the register stores line breaks as "||" and
' a FieldID or EntityCode carrying one would otherwise collide two different
' rows onto one key. That is the same defect that destroyed an hour of marking
' when shapes were keyed by name.
Public Function RowKey(entityCode As String, fieldId As String, quarter As String) As String
    RowKey = UCase$(Trim$(entityCode)) & Chr(1) & UCase$(Trim$(fieldId)) & Chr(1) & UCase$(Trim$(quarter))
End Function

' Reads the register worksheet into the `existing` map PlanSeedRows expects.
' Columns by header NAME, never by position -- the rule this codebase has
' already broken twice, once by reading the sheet by position and once by
' reading the wrong sheet entirely.
Public Function ExistingRowKeys(ws As Object) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")

    Dim cEntity As Long, cField As Long, cQuarter As Long
    Dim c As Long
    For c = 1 To 30
        Select Case Trim$(CStr(ws.Cells(1, c).Value))
            Case Register.COL_ENTITY:  cEntity = c
            Case Register.COL_FIELDID: cField = c
            Case Register.COL_QUARTER: cQuarter = c
        End Select
    Next c

    If cEntity = 0 Or cField = 0 Or cQuarter = 0 Then
        Err.Raise vbObjectError + 2, "RegisterSeed.ExistingRowKeys", _
            "register is missing one of " & Register.COL_ENTITY & "/" & _
            Register.COL_FIELDID & "/" & Register.COL_QUARTER & " in its header row"
    End If

    Dim r As Long
    r = 2
    Do While Trim$(CStr(ws.Cells(r, cEntity).Value)) <> ""
        result(RowKey(CStr(ws.Cells(r, cEntity).Value), _
                      CStr(ws.Cells(r, cField).Value), _
                      CStr(ws.Cells(r, cQuarter).Value))) = True
        r = r + 1
    Loop

    Set ExistingRowKeys = result
End Function

' Human-readable account of what a plan would do, for the confirmation dialog.
' NOTHING REACHES THE REGISTER UNSEEN is this project's standing rule, and a
' count alone is not seeing -- the clashes are named individually because they
' are the ones a person needs to go and look at.
Public Function PlanDiagnostic(p As SeedPlan, deckPeriod As String) As String
    Dim s As String
    s = p.ToAdd & " row(s) would be added for " & deckPeriod & "." & vbCrLf

    If p.SkippedExisting > 0 Then
        s = s & p.SkippedExisting & " skipped -- already present, nothing to do." & vbCrLf
    End If

    If p.SkippedCadenceClash > 0 Then
        s = s & vbCrLf & p.SkippedCadenceClash & " REFUSED as a cadence clash:" & vbCrLf
        Dim i As Long
        If p.ToAdd + p.SkippedExisting + p.SkippedCadenceClash > 0 Then
            For i = LBound(p.Rows) To UBound(p.Rows)
                If p.Rows(i).Skip And p.Rows(i).Reason <> "" Then
                    If InStr(p.Rows(i).Reason, "cadence clash") > 0 Then
                        s = s & "    " & p.Rows(i).EntityCode & " / " & p.Rows(i).FieldID & vbCrLf
                    End If
                End If
            Next i
        End If
        s = s & "These are marked static but already carry a period row. Decide in" & vbCrLf & _
                "Excel which cadence is right; the tool will not guess." & vbCrLf
    End If

    PlanDiagnostic = s
End Function
