Attribute VB_Name = "Resolve"
Option Explicit

' VBA port of src/resolve.py's resolve_slide_instance(), per specs/vba-
' port.md's port order (module 4, after discovery/identity_tags/matching).
' Resolves a live Slide's identity (slide_type, instance_key) via native
' Slide.Tags -- the same object-model mechanism InjectPrimitive.bas already
' uses for shape-level role tags, just one level up.
'
' Real gap found and closed here, not carried over from the spike:
' InjectPrimitive.bas (module 2, "identity_tags") only ever implemented
' Shape.Tags("role") lookups, despite specs/vba-port.md's port order
' describing module 2 as "already done (InjectPrimitive.bas's
' Shape.Tags/Slide.Tags reads)". Grepping InjectPrimitive.bas confirms it
' never once reads Slide.Tags. This module is the first one that actually
' needs slide-level tags (type_tag/instance_key), so the missing read is
' added here rather than left open further -- exactly the "extend only if
' a gap is found" instruction module 2's own port-order entry gives. Scoped
' to reads only: sync-dispatch (this module's actual consumer) never writes
' slide-level tags -- that's onboarding's job (port-order step 5, not yet
' built), so a write/upsert equivalent stays open for that module rather
' than being invented speculatively here. See SPIKE_NOTES_Resolve.md.
'
' Field-shape resolution (the other half of resolve.py's job) is
' deliberately NOT ported as a separate pre-resolution step here -- see
' SPIKE_NOTES_Resolve.md's divergence list for why SyncOperations.bas calls
' InjectPrimitive.InjectPrimitive() directly per field instead of building
' a field-name -> Candidate map the way resolve.py's field_shapes dict
' does.

Public Type SlideInstance
    HasTypeTag As Boolean
    TypeTag As String        ' "" / HasTypeTag=False if the slide has no slide_type tag
    HasInstanceKey As Boolean
    InstanceKey As String    ' "" / HasInstanceKey=False if the slide has no instance_key tag
End Type

' Reads `sld`'s slide-level identity tags. A blank Slide.Tags(name) result
' is indistinguishable from "tag genuinely present but empty" (the same
' real VBA API quirk InjectPrimitive.bas's ShapeHasRoleTag already
' documents for Shape.Tags) -- since a blank slide_type/instance_key is not
' meaningful in this project's scheme, both read as "absent" here.
'
' `sld` is typed As Object so this also accepts a CustomLayout (needed to
' resolve against mst-slide-layouts.pptx-style fixtures, which have no
' Slides at all) -- see SPIKE_NOTES_Resolve.md for why whether
' CustomLayout.Tags actually exists is an unconfirmed, open item in this
' environment.
Public Function ResolveSlideInstance(sld As Object) As SlideInstance
    Dim result As SlideInstance

    result.TypeTag = sld.Tags("slide_type")
    result.HasTypeTag = (result.TypeTag <> "")

    result.InstanceKey = sld.Tags("instance_key")
    result.HasInstanceKey = (result.InstanceKey <> "")

    ResolveSlideInstance = result
End Function

' Manual smoke-test entry point -- not a real test harness, same as every
' other module here. See SPIKE_NOTES_Resolve.md for the full recipe
' (tagging a fixture, expected values cross-checked against
' tests/test_resolve.py).
Public Sub ManualSmokeTest()
    Dim sld As Object
    Set sld = Application.ActivePresentation.Slides(1)

    Dim instance As SlideInstance
    instance = ResolveSlideInstance(sld)

    Dim msg As String
    msg = "HasTypeTag=" & instance.HasTypeTag & " TypeTag=" & instance.TypeTag & _
        " HasInstanceKey=" & instance.HasInstanceKey & " InstanceKey=" & instance.InstanceKey
    Debug.Print msg
    MsgBox msg
End Sub
