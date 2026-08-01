Attribute VB_Name = "E2EFirstField"
Option Explicit

' The §5 meet point, driven headlessly: ONE FIELD taken completely through.
'
' PROJECT_STATUS, on a copy of the real 46-slide deck, from a long-format
' register, via the FieldID rename, injected and verified.
'
' Exists because after ten rounds of design documents the count of fields taken
' end to end was zero. Every piece below was built and tested in isolation and
' none of them had ever been connected to another. This connects them.
'
' Three previously-unexercised modules are proved in one pass:
'   V1  TagMigration     -- role tag values renamed to register FieldIDs
'   V3  Register         -- long-format read, filtered by period and status
'       RunSync.RunRoutineSyncWithSheet -- the register feeding the real engine
'
' The prediction is falsifiable and was computed BEFORE the run: the register
' normalises the deck's case drift (`In progress` -> `In Progress`,
' `Not Started` / `Not yet commenced` -> `Not started`), which differs from what
' the slides currently hold on EXACTLY 12 of 46. So a correct run corrects 12,
' creates 0, and fails 0. Any other number is a finding.
' Isolation probe: if this is callable and E2EFirstField is not, the fault is
' in that function's signature rather than in the module or the call mechanism.

' Second probe: three required String args, same shape as the real function's
' signature. Isolates "3 string arguments" from "this function specifically".

' --- Bisection probes: same 3-String-arg / String-return signature as RunE2E,
' each adding one more UDT-typed local Dim (declared, never assigned/used).
' Goal: find which local UDT declaration (if any) makes the enclosing
' Public Function invisible to Application.Run.




' Dumps every managed field's value AS POWERPOINT READS IT, TSV, one row per
' EntityCode x FieldID. Real line breaks are encoded as the register delimiter.
'
' Exists because the register was first built by parsing the .pptx XML in
' Python -- a reimplementation of how PowerPoint reads text, which disagreed
' with the object model on nearly every field and produced a register that
' "corrected" 46 values to themselves. The only trustworthy source for what a
' slide currently says is the object model that will later be asked to compare
' against it.
Public Function DumpFieldValues(deckPath As String) As String
    Dim pres As Object
    Set pres = Application.Presentations.Open(deckPath, msoTrue, msoFalse, msoTrue)
    pres.Windows(1).Activate

    Dim out As String
    Dim sld As Object
    For Each sld In pres.Slides
        Dim inst As SlideInstance
        inst = Resolve.ResolveSlideInstance(sld)
        If inst.HasInstanceKey And Not inst.IsTemplate Then
            DumpShapes sld.Shapes, inst.InstanceKey, out
        End If
    Next sld

    pres.Saved = msoTrue
    pres.Close
    DumpFieldValues = out
End Function

Private Sub DumpShapes(shapesColl As Object, key As String, ByRef out As String)
    Dim shp As Object
    For Each shp In shapesColl
        If shp.Type = msoGroup Then
            DumpShapes shp.GroupItems, key, out
        Else
            Dim role As String
            role = shp.Tags("role")
            If role <> "" Then
                Dim txt As String
                txt = ""
                On Error Resume Next
                txt = shp.TextFrame.TextRange.Text
                On Error GoTo 0
                ' vbCr is PowerPoint's paragraph separator; vbLf appears inside
                ' some runs. Both become the register delimiter so the value is
                ' single-line in Excel, per R6.
                txt = Replace(txt, vbCrLf, "||")
                txt = Replace(txt, vbCr, "||")
                txt = Replace(txt, vbLf, "||")
                out = out & key & vbTab & role & vbTab & txt & vbCrLf
            End If
        End If
    Next shp
End Sub

Public Function RunE2E(deckPath As String, registerPath As String, period As String) As String
    Dim r As String

    Dim pres As Object
    Set pres = Application.Presentations.Open(deckPath, msoFalse, msoFalse, msoTrue)
    pres.Windows(1).Activate

    r = "Deck:     " & deckPath & vbCrLf & _
        "Register: " & registerPath & vbCrLf & _
        "Period:   " & period & vbCrLf & _
        "Slides:   " & pres.Slides.count & vbCrLf & vbCrLf

    ' --- V1: rename role tag values to register FieldIDs -------------------
    Dim fromV(1 To 5) As String
    Dim toV(1 To 5) As String
    fromV(1) = "Project Status": toV(1) = "PROJECT_STATUS"
    fromV(2) = "Project Name":   toV(2) = "PROJECT_NAME"
    fromV(3) = "Project number": toV(3) = "PROJECT_CODE"
    fromV(4) = "About text":     toV(4) = "ABOUT_BODY"
    fromV(5) = "events text":    toV(5) = "KEY_EVENTS_BODY"

    Dim mig As MigrationReport
    mig = TagMigration.MigrateRoleTags(fromV, toV, False)
    r = r & "--- V1 rename ---" & vbCrLf & _
        "  scanned:  " & mig.Scanned & vbCrLf & _
        "  renamed:  " & mig.Renamed & vbCrLf & _
        "  already:  " & mig.AlreadyDone & vbCrLf & _
        "  unmapped: " & mig.Unmapped & vbCrLf & vbCrLf

    ' --- V3: read the long-format register --------------------------------
    Dim xl As Object, wb As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Open(registerPath, 0, True)

    Dim reg As RegisterRead
    reg = Register.ReadRegister(WorkbookBridge.RegisterOrFirstDataSheet(wb), period, "q")

    r = r & "--- V3 register read ---" & vbCrLf & _
        "  rows seen:        " & reg.RowsSeen & vbCrLf & _
        "  accepted:         " & reg.Accepted & vbCrLf & _
        "  (period-matched:  " & reg.AcceptedPeriod & ", entity-static: " & reg.AcceptedStatic & ")" & vbCrLf & _
        "  rejected status:  " & reg.RejectedStatus & vbCrLf & _
        "  rejected period:  " & reg.RejectedPeriod & vbCrLf & _
        "  missing columns:  '" & reg.MissingColumns & "'" & vbCrLf & _
        "  " & Register.ReadDiagnostic(reg, period) & vbCrLf & vbCrLf

    If reg.MissingColumns <> "" Or reg.Accepted = 0 Then
        wb.Close False
        xl.Quit
        pres.Saved = msoTrue
        pres.Close
        RunE2E = r & "STOPPED: nothing usable read from the register."
        Exit Function
    End If

    ' --- The join: register -> the real sync engine ------------------------
    Dim templateSld As Object
    Dim wsName As String
    DeckRegistry.LookupType pres, "q", templateSld, wsName

    r = r & "--- sync, fed from the register ---" & vbCrLf & _
        RunSync.RunRoutineSyncWithSheet(reg.Data, "q", templateSld) & vbCrLf

    wb.Close False
    xl.Quit

    ' --- Verify by re-reading the DECK, not by trusting the report ---------
    ' The report says what the engine believes it did. This counts what the
    ' slides actually hold now. They are different claims and only the second
    ' one is evidence.
    Dim matched As Long, mismatched As Long, missing As Long
    Dim k As Variant
    For Each k In reg.Data.Rows.Keys
        Dim want As String
        want = CStr(reg.Data.Rows(k)("PROJECT_STATUS"))

        Dim found As Boolean
        found = False
        Dim sld As Object
        For Each sld In pres.Slides
            Dim inst As SlideInstance
            inst = Resolve.ResolveSlideInstance(sld)
            If inst.HasInstanceKey And inst.InstanceKey = CStr(k) And Not inst.IsTemplate Then
                Dim shp As Object
                Set shp = FindByRole(sld.Shapes, "PROJECT_STATUS")
                If Not shp Is Nothing Then
                    found = True
                    If Trim(shp.TextFrame.TextRange.Text) = want Then
                        matched = matched + 1
                    Else
                        mismatched = mismatched + 1
                        r = r & "  MISMATCH " & k & ": slide has '" & shp.TextFrame.TextRange.Text & "', register says '" & want & "'" & vbCrLf
                    End If
                End If
            End If
        Next sld
        If Not found Then missing = missing + 1
    Next k

    r = r & vbCrLf & "--- VERIFIED ON THE DECK ---" & vbCrLf & _
        "  slides matching the register: " & matched & vbCrLf & _
        "  mismatched:                   " & mismatched & vbCrLf & _
        "  register rows with no slide:  " & missing & vbCrLf

    pres.Save
    r = r & vbCrLf & "Deck saved." & vbCrLf

    RunE2E = r
End Function

Private Function FindByRole(shapesColl As Object, role As String) As Object
    Dim shp As Object
    For Each shp In shapesColl
        If shp.Type = msoGroup Then
            Dim inner As Object
            Set inner = FindByRole(shp.GroupItems, role)
            If Not inner Is Nothing Then
                Set FindByRole = inner
                Exit Function
            End If
        ElseIf shp.Tags("role") = role Then
            Set FindByRole = shp
            Exit Function
        End If
    Next shp
End Function
