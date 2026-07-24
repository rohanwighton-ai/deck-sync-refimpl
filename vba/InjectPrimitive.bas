Attribute VB_Name = "InjectPrimitive"
Option Explicit

' VBA port of src/verification.py's inject_primitive, scoped to a single
' primitive write. Ports the semantics, not the implementation: the Python
' side hand-rolls OOXML zip/XML surgery to read/write a shape's text and to
' resolve a shape by its hidden identity tag (custDataLst -> Tags Part,
' reverse-engineered from ECMA-376 since python-pptx has no support for it
' -- see src/identity_tags.py). VBA/the PowerPoint object model already
' implements that exact mechanism natively as Shape.Tags, so this port does
' not re-derive any XML; it uses the object model directly. See
' SPIKE_NOTES.md for the full list of deliberate divergences from the
' Python semantics and why each one is safe.

Public Type InjectResult
    Found As Boolean       ' False if no shape (or >1 shapes) carried the tag
    Written As Boolean     ' False if the current value already matched (no-op)
    Verified As Boolean    ' current value re-read from the shape equals source_value
    ErrorMessage As String ' non-empty iff Found=False or Verified=False
End Type

' Locate the single shape on `sld` whose Tags("role") equals `identityTag`.
' Mirrors identity_tags.py's read_shape_tags/upsert_shape_tags convention:
' the identity tag lives under the fixed tag name "role" (specs/identity-
' tags.md: "Shape-level tags (role)"), not under the shape's visible Name --
' real-deck validation upstream found 5 shapes on one real slide all named
' the literal string "Rounded Rectangle 1", which is exactly why tag-based
' lookup exists instead of name-based lookup.
'
' Returns Nothing if zero or more than one shape carries the tag -- multiple
' matches are a data-integrity problem (tags are supposed to be unique per
' role on a slide) and are never silently resolved by picking the first one,
' consistent with the source project's stated philosophy of proving links
' rather than assuming them.
Private Function FindShapeByRoleTag(sld As Object, identityTag As String) As Object
    Dim shp As Object
    Dim match As Object
    Dim matchCount As Long
    matchCount = 0

    For Each shp In sld.Shapes
        If ShapeHasRoleTag(shp, identityTag) Then
            matchCount = matchCount + 1
            Set match = shp
        End If
    Next shp

    If matchCount = 1 Then
        Set FindShapeByRoleTag = match
    Else
        Set FindShapeByRoleTag = Nothing
    End If
End Function

' Shape.Tags(name) returns "" both when the tag is absent and when it is
' present with an empty string value -- a real VBA API quirk with no
' equivalent in the Python side's explicit dict-based tag storage (identity_
' tags.py's _parse_tag_list returns a real {} when a tag is missing). Since
' an empty-string role value is not a meaningful identity tag in this
' project's scheme, treating "absent" and "empty" the same way here is a
' safe simplification, not a correctness gap -- documented in SPIKE_NOTES.md.
Private Function ShapeHasRoleTag(shp As Object, identityTag As String) As Boolean
    ShapeHasRoleTag = (shp.Tags("role") = identityTag) And (identityTag <> "")
End Function

' Port of inject_primitive(path, part_name, shape, source_value). Takes a
' live Slide object and the identity tag instead of a file path + z_order,
' since the calling context here is a shape already open in the PowerPoint
' object model, not a closed .pptx being surgically edited byte-for-byte.
'
' Semantics preserved from the Python original:
'   1. No-op if the shape's current value already equals source_value.
'   2. Otherwise write source_value, then re-read the value back and
'      confirm it matches, rather than assuming the write call succeeded.
'   3. Error (Found=False) if the shape has no text frame / no text runs to
'      write into, mirroring inject_primitive's raise on a shape with no
'      text runs.
'
' Semantics NOT preserved (see SPIKE_NOTES.md for why each is judged safe
' for this spike's scope):
'   - Hash comparison (SHA-256) is replaced with direct VBA string equality.
'   - "Verified" re-reads the value from the live Shape object, which only
'     proves the object-model state changed -- it does not prove the write
'     persisted to the underlying OOXML part the way the Python version's
'     re-open-the-zip-from-disk check does. Manual verification recipe
'     below closes and reopens the file to close this gap by hand.
Public Function InjectPrimitive(sld As Object, identityTag As String, sourceValue As String) As InjectResult
    Dim result As InjectResult
    Dim shp As Object

    Set shp = FindShapeByRoleTag(sld, identityTag)
    If shp Is Nothing Then
        result.Found = False
        result.Written = False
        result.Verified = False
        result.ErrorMessage = "no single shape found tagged role=" & identityTag & _
            " (zero matches, or more than one -- ambiguous tag, refusing to guess)"
        InjectPrimitive = result
        Exit Function
    End If

    result.Found = True

    If Not shp.HasTextFrame Then
        result.Written = False
        result.Verified = False
        result.ErrorMessage = "shape tagged role=" & identityTag & " has no text frame to write into"
        InjectPrimitive = result
        Exit Function
    End If

    Dim currentValue As String
    currentValue = shp.TextFrame.TextRange.Text

    If currentValue = sourceValue Then
        result.Written = False
        result.Verified = True
        result.ErrorMessage = ""
        InjectPrimitive = result
        Exit Function
    End If

    ' Assigning .Text replaces the entire text-range contents in one run,
    ' which matches _set_text's "write the first run, clear the rest" net
    ' effect (formatting of the original first run is retained by
    ' PowerPoint; any additional runs are removed rather than left stale).
    shp.TextFrame.TextRange.Text = sourceValue

    Dim reReadValue As String
    reReadValue = shp.TextFrame.TextRange.Text

    result.Written = True
    result.Verified = (reReadValue = sourceValue)
    If Not result.Verified Then
        result.ErrorMessage = "wrote " & sourceValue & " but re-read " & reReadValue & " -- write did not take"
    Else
        result.ErrorMessage = ""
    End If

    InjectPrimitive = result
End Function

' Manual smoke-test entry point -- not a real test harness (no VBA unit-test
' framework wired up here, see SPIKE_NOTES.md). Run this from the VBA IDE
' (F5) against a slide that already has a shape tagged role="demo_field",
' per the manual verification recipe in SPIKE_NOTES.md.
Public Sub ManualSmokeTest()
    Dim sld As Object
    Set sld = Application.ActivePresentation.Slides(1)

    Dim r As InjectResult
    r = InjectPrimitive(sld, "demo_field", "hello from VBA " & Now)

    Dim msg As String
    msg = "Found=" & r.Found & " Written=" & r.Written & " Verified=" & r.Verified
    If r.ErrorMessage <> "" Then msg = msg & " Error=" & r.ErrorMessage
    MsgBox msg
End Sub
