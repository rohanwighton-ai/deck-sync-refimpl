Attribute VB_Name = "SlideDuplication"
Option Explicit

' Implements specs/slide-duplication-trigger.md's DuplicateAndTag primitive
' -- the "how" for sync_operations' case 2/3 duplication decisions. This
' module does not decide *whether* to duplicate (SyncOperations does, via
' PlanRoutineSync/PlanPeriodRollover); it only implements what happens once
' that decision is made. Deck-order placement (row order as a standing
' invariant) is deliberately NOT this module's job either -- per the spec,
' resequencing covers both new and already-existing slides uniformly each
' sync, which is a whole-type, driver-level operation; RunSync.bas owns it.
'
' See SPIKE_NOTES_SlideDuplication.md for divergences and the manual/real
' test recipe.

Public Type DuplicateResult
    Ok As Boolean
    Reason As String            ' populated on refusal/failure
    NewSlide As Object          ' the created, tagged slide, if Ok
    MissingFieldCount As Long
    MissingFields() As String   ' fields the template defines that `values` didn't supply -- possibly unallocated when MissingFieldCount = 0
End Type

' Duplicates `sourceSld` (an already-onboarded template/reference slide of
' `slideType`), tags the duplicate with `slideType`/`newInstanceKey`, and
' injects `values` (Scripting.Dictionary fieldName -> value String) into
' the duplicate's tagged fields.
'
' `existingInstances` is the same shape as SyncOperations.PlanRoutineSync's
' `instances()` parameter (already gathered by the caller -- gathering is
' explicitly not this module's job, same non-goal every other module in
' this port draws for itself) and is used only for the instance-key
' collision guard.
Public Function DuplicateAndTag(sourceSld As Object, slideType As String, newInstanceKey As String, values As Object, existingInstances() As Object) As DuplicateResult
    Dim result As DuplicateResult

    ' --- Instance-key collision guard -----------------------------------
    ' Two Data-sheet rows sharing a key by mistake (bad paste, copy-drag
    ' error) must refuse and flag, never silently produce two slides for
    ' one key -- same never-silently-guess posture as case 6.
    Dim lo As Long, hi As Long, hasInstances As Boolean
    On Error Resume Next
    lo = LBound(existingInstances)
    hi = UBound(existingInstances)
    hasInstances = (Err.Number = 0)
    On Error GoTo 0

    If hasInstances Then
        Dim i As Long
        For i = lo To hi
            Dim existing As SlideInstance
            existing = Resolve.ResolveSlideInstance(existingInstances(i))
            If existing.HasInstanceKey And existing.InstanceKey = newInstanceKey Then
                result.Ok = False
                result.Reason = "instance_key '" & newInstanceKey & "' already exists on another slide -- refusing to create a duplicate rather than silently double-creating"
                DuplicateAndTag = result
                Exit Function
            End If
        Next i
    End If

    ' --- Duplicate --------------------------------------------------------
    ' Slide.Duplicate's own default placement (immediately after the
    ' source) is deliberately not corrected here -- per specs/slide-
    ' duplication-trigger.md, row-order placement is a whole-type
    ' resequencing concern RunSync.bas applies uniformly after this call,
    ' not a per-call correction.
    Dim newSlides As Object
    Set newSlides = sourceSld.Duplicate()
    Dim newSld As Object
    Set newSld = newSlides(1)

    ' --- Mandatory structural/z-order verification before any tag write --
    ' A malformed duplicate must never receive an instance_key. Deleted
    ' rather than left sitting untagged in the deck -- an unlabeled
    ' duplicate is confusing debris a human would have to notice and clean
    ' up by hand; refusing and removing it keeps the deck in the same
    ' state as if this call had never run. (A judgment call the spec
    ' itself doesn't dictate explicitly -- documented in
    ' SPIKE_NOTES_SlideDuplication.md.)
    Dim structCheck As StructuralVerification
    structCheck = Verification.VerifyStructure(sourceSld, newSld)
    Dim zOrderCheck As ZOrderVerification
    zOrderCheck = Verification.VerifyZOrder(sourceSld, newSld)

    If Not structCheck.Ok Or Not zOrderCheck.Ok Then
        Dim reasonMsg As String
        reasonMsg = "structural/z-order verification failed after duplication -- malformed duplicate removed, never tagged. "
        If Not structCheck.Ok Then reasonMsg = reasonMsg & structCheck.MismatchCount & " structural mismatch(es). "
        If Not zOrderCheck.Ok Then reasonMsg = reasonMsg & zOrderCheck.MismatchCount & " z-order mismatch(es)."
        newSld.Delete
        result.Ok = False
        result.Reason = reasonMsg
        DuplicateAndTag = result
        Exit Function
    End If

    ' --- Tag slide-level identity, unconditionally ------------------------
    newSld.Tags.Add "slide_type", slideType
    newSld.Tags.Add "instance_key", newInstanceKey

    ' --- Inject fields, per partial-row handling ---------------------------
    ' The type's known field set comes from the source template's own
    ' tagged fields (already verified structurally identical to the
    ' duplicate above) -- reuses Onboarding.BuildTemplateFieldShapes rather
    ' than re-deriving "what fields does this type have" a second way.
    ' A row missing a value for one of the type's fields still gets a
    ' slide created (never withheld pending "completeness", since no
    ' completeness rule exists anywhere in the spec chain) -- injection is
    ' skipped only for the missing field, and it's flagged for visibility.
    Dim templateRoles() As String
    Dim templateFieldShapes() As Candidate
    templateFieldShapes = Onboarding.BuildTemplateFieldShapes(sourceSld, templateRoles)

    Dim rLo As Long, rHi As Long, hasRoles As Boolean
    On Error Resume Next
    rLo = LBound(templateRoles)
    rHi = UBound(templateRoles)
    hasRoles = (Err.Number = 0)
    On Error GoTo 0

    Dim missingCount As Long
    missingCount = 0
    If hasRoles Then
        Dim r As Long
        For r = rLo To rHi
            Dim role As String
            role = templateRoles(r)
            If values.Exists(role) Then
                InjectPrimitive.InjectPrimitive newSld, role, CStr(values(role))
            Else
                missingCount = missingCount + 1
                ReDim Preserve result.MissingFields(1 To missingCount)
                result.MissingFields(missingCount) = role
            End If
        Next r
    End If

    result.MissingFieldCount = missingCount
    result.Ok = True
    Set result.NewSlide = newSld
    DuplicateAndTag = result
End Function
