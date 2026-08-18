Attribute VB_Name = "TagMigration"
Option Explicit

' V1 of the Excel Control Layer workplan: rename the `role` tag VALUE on every
' managed shape, from the current human-readable names to the register's
' FieldIDs.
'
' Why this is small, and why it is being done first. FieldID is the VALUE of
' the role tag, not the tag's name, so a rename is one string overwrite per
' shape -- not a structural change. Measured against real Office 2026-07-31:
'
'   Tags.Add on an existing name:    Count 1 -> 1, value replaced  => OVERWRITES
'   Enumeration (Count/Name/Value):  works                          => ENUMERABLE
'   Setting the old value back:      succeeds, Count unchanged      => REVERSIBLE
'
' The first of those is the one that mattered. Had Add appended rather than
' replaced, every migrated shape would carry two role tags, and
' InjectPrimitive's ambiguity guard would then refuse to write to any of them --
' turning a rename into a deck-wide outage. It cannot happen.
'
' Sequenced first because the cost scales with the tagged-field count and every
' other task raises it: 5 FieldIDs and ~30 writes today, against ~43 and ~258
' once the field inventory is onboarded. It never gets cheaper than now.
'
' Walks EVERY slide, deliberately not RunSync.GatherInstances. The gather
' excludes the master template by design, and the template's fields carry role
' tags too -- skipping it would leave the one slide every new record is cloned
' from holding the old names, so every future creation would reintroduce them.
'
' Rollback is this same operation with the map read right to left.

Public Type MigrationReport
    Scanned As Long          ' shapes carrying any role tag
    Renamed As Long          ' shapes whose role value was in the map and changed
    AlreadyDone As Long      ' shapes already carrying the target value -- idempotent re-runs
    Unmapped As Long         ' shapes with a role value the map does not mention
    UnmappedDetail As String
    Detail As String
End Type

' `fromValues` / `toValues` are parallel arrays, 1-based and equal length --
' the project's standard substitute for a Dictionary, which cannot hold the
' pairs usefully here and cannot hold UDTs at all (AGENTS.md).
'
' dryRun writes nothing and reports exactly what a real run would do. Default
' TRUE: this is a bulk tag rewrite across a whole deck, and the one operation
' where "I meant to preview that" is expensive. The caller must ask for the
' write explicitly.
Public Function MigrateRoleTags(fromValues() As String, toValues() As String, _
                                Optional dryRun As Boolean = True) As MigrationReport
    Dim result As MigrationReport

    Dim mLo As Long, mHi As Long, hasMap As Boolean
    On Error Resume Next
    mLo = LBound(fromValues): mHi = UBound(fromValues)
    hasMap = (Err.Number = 0)
    On Error GoTo 0
    If Not hasMap Then
        result.Detail = "No mapping supplied -- nothing to do."
        MigrateRoleTags = result
        Exit Function
    End If

    Dim sld As Object
    For Each sld In Application.ActivePresentation.Slides
        WalkShapes sld.Shapes, fromValues, toValues, mLo, mHi, dryRun, result, sld.SlideIndex
    Next sld

    If result.Unmapped > 0 Then
        result.Detail = result.Detail & vbCrLf & _
            "UNMAPPED -- these carry a role tag the mapping does not mention, and were left alone:" & vbCrLf & _
            result.UnmappedDetail
    End If

    MigrateRoleTags = result
End Function

' Recurses groups. Real decks stack managed fields inside groups -- the field
' inventory for this deck shows role-tagged shapes nested two levels down, and
' a non-recursive walk would silently migrate only the top-level ones, leaving
' a deck that is half-renamed in a way nothing reports.
Private Sub WalkShapes(shapesColl As Object, fromValues() As String, toValues() As String, _
                       mLo As Long, mHi As Long, dryRun As Boolean, _
                       ByRef result As MigrationReport, slideIndex As Long)
    Dim shp As Object
    For Each shp In shapesColl
        If shp.Type = msoGroup Then
            WalkShapes shp.GroupItems, fromValues, toValues, mLo, mHi, dryRun, result, slideIndex
        Else
            Dim current As String
            current = shp.Tags("role")
            If current <> "" Then
                result.Scanned = result.Scanned + 1

                Dim matched As Boolean
                matched = False
                Dim i As Long
                For i = mLo To mHi
                    ' Case-insensitive on the VALUE: these are human-typed names
                    ' ("About text"), and a migration that silently skips a shape
                    ' over a capital letter is worse than one that refuses.
                    If StrComp(current, fromValues(i), vbTextCompare) = 0 Then
                        matched = True
                        If Not dryRun Then
                            shp.Tags.Add "role", toValues(i)
                            ' The renamed tag is a NEW identity key on this
                            ' slide -- tell the per-slide tag index, or a
                            ' slide walked before the rename keeps answering
                            ' "absent" for the new name all session
                            ' (ShapeAddressBook.NoteRoleTagAdded's header).
                            ShapeAddressBook.NoteRoleTagAdded shp, toValues(i)
                        End If
                        result.Renamed = result.Renamed + 1
                        result.Detail = result.Detail & _
                            "  slide " & slideIndex & ": '" & current & "' -> '" & toValues(i) & "'" & vbCrLf
                        Exit For
                    ElseIf StrComp(current, toValues(i), vbTextCompare) = 0 Then
                        ' Already carries the target. Counted separately rather
                        ' than as unmapped, so a re-run reads as "nothing left to
                        ' do" instead of "N shapes I do not recognise" -- the
                        ' second would look like a failure on an idempotent run.
                        matched = True
                        result.AlreadyDone = result.AlreadyDone + 1
                        Exit For
                    End If
                Next i

                If Not matched Then
                    result.Unmapped = result.Unmapped + 1
                    result.UnmappedDetail = result.UnmappedDetail & _
                        "  slide " & slideIndex & ": '" & current & "'" & vbCrLf
                End If
            End If
        End If
    Next shp
End Sub

' The report a human reads before authorising the write.
' Pure, so the wording is testable -- same reason RunSync.ConfirmSyncText is.
Public Function MigrationSummary(r As MigrationReport, dryRun As Boolean) As String
    Dim s As String
    s = IIf(dryRun, "=== PREVIEW (nothing written) ===", "=== MIGRATION APPLIED ===") & vbCrLf & vbCrLf & _
        "  role-tagged shapes found: " & r.Scanned & vbCrLf & _
        "  " & IIf(dryRun, "would be renamed:        ", "renamed:                 ") & r.Renamed & vbCrLf & _
        "  already correct:          " & r.AlreadyDone & vbCrLf & _
        "  unmapped (left alone):    " & r.Unmapped & vbCrLf & vbCrLf & _
        r.Detail

    If r.Unmapped > 0 Then
        s = s & vbCrLf & "An unmapped tag is not an error -- it is a field the mapping" & vbCrLf & _
            "does not cover yet. It keeps its current value and will simply not" & vbCrLf & _
            "match a register row until it is mapped." & vbCrLf
    End If

    If Not dryRun And r.Renamed > 0 Then
        s = s & vbCrLf & "To roll back, run the same migration with the mapping reversed." & vbCrLf
    End If

    MigrationSummary = s
End Function
