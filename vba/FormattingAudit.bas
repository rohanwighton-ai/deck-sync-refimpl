Attribute VB_Name = "FormattingAudit"
Option Explicit

' FIELD-BY-FIELD FORMATTING CHECK -- the standing activity Rohan named
' 2026-08-19, now that field REACHABILITY (tagging/wiring) is largely
' covered. Reachability checks (FieldWiring.bas) answer "can this field
' receive a value". This answers a different question: once a field HAS a
' value, does it render the way every other slide's same field does.
'
' The specimen that named this activity: 1_S004's PROJECT_STATUS badge at
' 18pt, next to 8 other real slides' same field all at 8.5pt (one at 7pt)
' -- invisible to every wiring/tag/geometry check, because all of those
' measure whether a field CAN be written to, never what it looks like once
' it has been.
'
' A role seen on fewer than two real slides has no "usual" to differ from
' and is skipped -- not silence, just nothing to compare yet. Groups
' (msoGroup, devices like MILESTONE_TIMELINE) are walked into but never
' themselves compared -- a device's internal consistency is a different,
' harder question, deliberately out of scope for this first pass.

' One shape's specimen, "|"-joined rather than a UDT: VBA Collections
' cannot hold a user-defined Type directly (Collection.Add needs a Variant-
' compatible item), the same reason DeckRegistry's registration strings
' use this encoding. Fields: instanceKey|shapeName|hasText|fontSize|shapeType
Private Function MakeSpecimen(instanceKey As String, shp As Object) As String
    Dim hasText As String, fontSize As String
    hasText = "0"
    fontSize = "0"
    On Error Resume Next
    If shp.HasTextFrame Then
        If shp.TextFrame.HasText Then
            hasText = "1"
            fontSize = CStr(shp.TextFrame.TextRange.Font.Size)
        End If
    End If
    On Error GoTo 0
    MakeSpecimen = instanceKey & "|" & shp.Name & "|" & hasText & "|" & fontSize & "|" & CStr(shp.Type)
End Function

' Scans every real (instance_key-tagged, not-template) slide in `pres`,
' groups every role-tagged leaf shape by role name, and reports any
' specimen whose font size or shape Type differs from that role's own most
' common value.
Public Function ScanFormattingOutliers(pres As Object) As String
    Dim byRole As Object
    Set byRole = CreateObject("Scripting.Dictionary")

    Dim sld As Object
    For Each sld In pres.Slides
        Dim inst As SlideInstance
        inst = Resolve.ResolveSlideInstance(sld)
        If inst.HasInstanceKey And Not inst.IsTemplate Then
            CollectSpecimens sld.Shapes, inst.InstanceKey, byRole
        End If
    Next sld

    ScanFormattingOutliers = ReportOutliers(byRole)
End Function

Private Sub CollectSpecimens(shapesColl As Object, instanceKey As String, ByRef byRole As Object)
    Dim shp As Object
    For Each shp In shapesColl
        Dim role As String
        role = ""
        On Error Resume Next
        role = shp.Tags("role")
        On Error GoTo 0

        If role <> "" And shp.Type <> msoGroup Then
            ' Grouped by (role, template LETTER), not role alone. K/S/P are
            ' genuinely different template variants -- confirmed live
            ' 2026-08-19: PROJECT_NAME's font size is a legitimate ~13/11/
            ' ~11 split by letter, and comparing across the whole deck
            ' flagged nearly every real slide as an "outlier" against a
            ' majority that was really just "whichever letter has the most
            ' slides". A field's usual formatting is scoped to its own
            ' variant, the same way Scenario 3 scoped colour to the letter.
            Dim letter As String
            letter = TemplateSlide.CodeLetterOf(instanceKey)
            Dim key As String
            key = UCase(Trim(role)) & "|" & letter
            If Not byRole.Exists(key) Then Set byRole(key) = New Collection
            byRole(key).Add MakeSpecimen(instanceKey, shp)
        End If

        If shp.Type = msoGroup Then
            CollectSpecimens shp.GroupItems, instanceKey, byRole
        End If
    Next shp
End Sub

Private Function ReportOutliers(byRole As Object) As String
    Dim report As String
    Dim roleKey As Variant
    For Each roleKey In byRole.Keys
        Dim specimens As Collection
        Set specimens = byRole(roleKey)
        If specimens.Count >= 2 Then
            Dim displayName As String
            displayName = Replace(CStr(roleKey), "|", " [") & "]"
            report = report & ReportOutliersForRole(displayName, specimens)
        End If
    Next roleKey

    If report = "" Then
        ReportOutliers = "No formatting outliers found (" & byRole.Count & " role(s) checked)."
    Else
        ReportOutliers = report
    End If
End Function

Private Function ReportOutliersForRole(roleName As String, specimens As Collection) As String
    ' First pass: the mode of font size (text specimens only) and shape type.
    Dim sizeCounts As Object, typeCounts As Object
    Set sizeCounts = CreateObject("Scripting.Dictionary")
    Set typeCounts = CreateObject("Scripting.Dictionary")

    Dim s As Variant
    For Each s In specimens
        Dim parts() As String
        parts = Split(CStr(s), "|")
        Dim hasText As String, fontSize As String, shapeType As String
        hasText = parts(2): fontSize = parts(3): shapeType = parts(4)

        If hasText = "1" Then
            If sizeCounts.Exists(fontSize) Then
                sizeCounts(fontSize) = sizeCounts(fontSize) + 1
            Else
                sizeCounts(fontSize) = 1
            End If
        End If

        If typeCounts.Exists(shapeType) Then
            typeCounts(shapeType) = typeCounts(shapeType) + 1
        Else
            typeCounts(shapeType) = 1
        End If
    Next s

    Dim modeSize As String, modeType As String
    modeSize = ModeOf(sizeCounts)
    modeType = ModeOf(typeCounts)

    ' Second pass: flag every specimen that disagrees with its role's mode.
    Dim out As String
    For Each s In specimens
        parts = Split(CStr(s), "|")
        Dim instanceKey As String, shapeName As String
        instanceKey = parts(0): shapeName = parts(1): hasText = parts(2): fontSize = parts(3): shapeType = parts(4)

        Dim reasons As String
        reasons = ""
        If hasText = "1" And modeSize <> "" And fontSize <> modeSize Then
            reasons = reasons & "font size " & fontSize & " (usual: " & modeSize & ")"
        End If
        If modeType <> "" And shapeType <> modeType Then
            If reasons <> "" Then reasons = reasons & ", "
            reasons = reasons & "shape type " & shapeType & " (usual: " & modeType & ")"
        End If

        If reasons <> "" Then
            out = out & "  " & roleName & " on " & instanceKey & " (" & shapeName & "): " & reasons & vbCrLf
        End If
    Next s

    ReportOutliersForRole = out
End Function

' Most-frequent key in a Dictionary of value->count. First-seen wins a tie
' -- Dictionary iteration order is insertion order in VBA/Scripting, so this
' is deterministic, not arbitrary, even though a tie has no single "right"
' answer.
Private Function ModeOf(counts As Object) As String
    Dim bestKey As String, bestCount As Long
    bestCount = 0
    Dim k As Variant
    For Each k In counts.Keys
        If counts(k) > bestCount Then
            bestCount = counts(k)
            bestKey = CStr(k)
        End If
    Next k
    ModeOf = bestKey
End Function
