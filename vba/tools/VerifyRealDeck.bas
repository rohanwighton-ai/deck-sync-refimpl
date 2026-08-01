Attribute VB_Name = "VerifyRealDeck"
Option Explicit

' One-off, read-only diagnostic -- NOT part of the shipped add-in (not in
' build_ppam.ps1's module list, not wired to CommandBarUI). Exists to
' answer one question after the 2026-07-26 "Linked: 0 / FAILED
' verification: 46" incident (see SPIKE_NOTES_BatchOnboardFlow.md): did the
' underlying Shape.Tags/text writes from that run actually land correctly
' on the real deck, even though CommitBatch's own verification step
' reported failure for every slide (root cause: InjectPrimitive.bas's
' pre-fix FindShapeByRoleTag couldn't see into groups -- the writes always
' worked, only the read-back that CommitBatch used to confirm them
' couldn't find what it had just written).
'
' Opens both real files READ-ONLY (Presentations.Open ReadOnly:=True,
' Workbooks.Open ReadOnly, and closes both without saving) and never
' writes to either. Cross-checks purely by reading: for every slide with a
' slide_type/instance_key (Resolve.ResolveSlideInstance, the same read
' path RunSync/DeckAdoption already trust), walks its shapes recursively
' (mirrors InjectPrimitive.bas's WalkForRoleTag) collecting role-tagged
' shapes' current text, and compares that against the matching row in the
' paired Data workbook (ExcelOutput.ReadSheet, the same read path
' SyncOperations already trusts) field-by-field.
Public Function VerifyRealDeck(deckPath As String, workbookPath As String) As String
    Dim report As String
    Dim detail As String

    Dim pres As Object
    Set pres = Application.Presentations.Open(deckPath, msoTrue, msoFalse, msoFalse)

    Dim xl As Object, wb As Object, ws As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Open(workbookPath, 0, True) ' positional: UpdateLinks:=0, ReadOnly:=True -- named args aren't reliable on a late-bound Object
    Set ws = WorkbookBridge.RegisterOrFirstDataSheet(wb)

    Dim sheet As ExcelOutput.Sheet
    sheet = ExcelOutput.ReadSheet(ws)

    Dim slidesChecked As Long, slidesOk As Long, slidesMissingTags As Long, slidesNoWorkbookRow As Long
    Dim fieldsChecked As Long, fieldsMissingShape As Long, fieldsMismatch As Long

    Dim sld As Object
    For Each sld In pres.Slides
        slidesChecked = slidesChecked + 1

        Dim inst As Resolve.SlideInstance
        inst = Resolve.ResolveSlideInstance(sld)

        If Not inst.HasTypeTag Or Not inst.HasInstanceKey Then
            slidesMissingTags = slidesMissingTags + 1
            detail = detail & "Slide " & sld.SlideIndex & ": MISSING slide_type and/or instance_key tag" & vbCrLf
        ElseIf Not sheet.Rows.Exists(inst.InstanceKey) Then
            slidesNoWorkbookRow = slidesNoWorkbookRow + 1
            detail = detail & "Slide " & sld.SlideIndex & " (instance_key=" & inst.InstanceKey & "): no matching row in workbook" & vbCrLf
        Else
            Dim roleTags As Object
            Set roleTags = CreateObject("Scripting.Dictionary")
            CollectRoleTags sld.Shapes, roleTags

            Dim rowValues As Object
            Set rowValues = sheet.Rows(inst.InstanceKey)

            Dim slideOk As Boolean
            slideOk = True

            Dim fieldName As Variant
            For Each fieldName In rowValues.Keys
                fieldsChecked = fieldsChecked + 1
                If Not roleTags.Exists(CStr(fieldName)) Then
                    fieldsMissingShape = fieldsMissingShape + 1
                    slideOk = False
                    detail = detail & "Slide " & sld.SlideIndex & " (" & inst.InstanceKey & "): field '" & fieldName & "' -- no shape on the slide carries this role tag" & vbCrLf
                ElseIf CStr(roleTags(fieldName)) <> CStr(rowValues(fieldName)) Then
                    fieldsMismatch = fieldsMismatch + 1
                    slideOk = False
                    detail = detail & "Slide " & sld.SlideIndex & " (" & inst.InstanceKey & "): field '" & fieldName & "' -- shape text does not match workbook value" & vbCrLf
                End If
            Next fieldName

            If slideOk Then slidesOk = slidesOk + 1
        End If
    Next sld

    wb.Close False
    xl.Quit
    pres.Close

    report = "=== Real Deck Verification Report ===" & vbCrLf & _
        "Deck: " & deckPath & vbCrLf & _
        "Workbook: " & workbookPath & vbCrLf & _
        "Run at: " & Now & vbCrLf & vbCrLf & _
        "Slides checked: " & slidesChecked & vbCrLf & _
        "Slides fully OK (all fields present + matching): " & slidesOk & vbCrLf & _
        "Slides missing slide_type/instance_key tags: " & slidesMissingTags & vbCrLf & _
        "Slides with no matching workbook row: " & slidesNoWorkbookRow & vbCrLf & _
        "Fields checked (across all OK-tagged slides): " & fieldsChecked & vbCrLf & _
        "Fields with no tagged shape found: " & fieldsMissingShape & vbCrLf & _
        "Fields with a text/value mismatch: " & fieldsMismatch & vbCrLf & vbCrLf & _
        "--- Per-slide detail (only non-OK slides/fields listed) ---" & vbCrLf & detail

    VerifyRealDeck = report
End Function

' Same recursion shape as InjectPrimitive.bas's WalkForRoleTag, but
' collects every role-tagged shape's text instead of searching for one
' specific tag -- this diagnostic needs the whole set per slide, not a
' single lookup.
Private Sub CollectRoleTags(shapesColl As Object, ByRef roleTags As Object)
    Dim shp As Object
    For Each shp In shapesColl
        If shp.Type = msoGroup Then
            CollectRoleTags shp.GroupItems, roleTags
        Else
            Dim r As String
            r = ""
            On Error Resume Next
            r = shp.Tags("role")
            On Error GoTo 0
            If r <> "" And shp.HasTextFrame Then
                roleTags(r) = shp.TextFrame.TextRange.Text
            End If
        End If
    Next shp
End Sub
