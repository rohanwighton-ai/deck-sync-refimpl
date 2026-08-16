Attribute VB_Name = "ExcelOutput"
Option Explicit

' VBA port of src/excel_output.py, per specs/vba-port.md's port order
' (module 6 of 6, the last one -- after discovery/identity_tags/matching/
' resolve+sync_operations/onboarding).
'
' This is the one module vba-port.md itself says is strictly SIMPLER than
' the Python it replaces, not just a mechanism swap: Python hand-rebuilds
' the whole .xlsx zip (Content_Types, workbook.xml, styles.xml, the sheet
' part, custom.xml) from scratch on every write, because it has no host
' application to lean on. VBA runs against a live Worksheet the calling
' context already has open (in Excel directly, or via COM automation driven
' from the PowerPoint side) -- so this is plain Cells/Range reads and
' writes, no XML, no zip, no "regenerate the whole file" step. Per
' vba-port.md: "Don't port excel_output.py's zip-rebuilding approach -- it
' only existed to work around Python lacking a live Excel instance."
'
' Layout convention preserved exactly from excel-output.md/excel_output.py:
' column A is the reserved instance-identity column (header "Instance ID");
' columns B.. hold confirmed fields, one per column, in first-seen (append)
' order; row 1 is the header, rows 2.. are data keyed by instance ID (never
' by row position). Field identity is the column's header TEXT, looked up
' by name on every read -- never assumed from position.
'
' See SPIKE_NOTES_ExcelOutput.md for deliberate divergences and the manual
' verification recipe -- there is no .xlsx test fixture for this spec (same
' as the Python side: this module is both writer and reader, so its own
' tests are round-trip/self-consistency checks).

Public Const INSTANCE_ID_HEADER As String = "Instance ID"
Private Const DECK_REFERENCE_PROPERTY_NAME As String = "DeckReference"
' STORAGE MOVED OFF CustomDocumentProperties, 2026-08-16 -- same class of
' defect as DeckRegistry.bas's registry slide (see that file's
' REGISTRY_SLIDE_NAME comment for the full PowerPoint-side story), confirmed
' independently on THIS application rather than assumed by analogy: a probe
' against a real OneDrive-hosted workbook found a brand-new custom property
' lands, and a SECOND, different new property in the same session also
' lands (narrower than PowerPoint's version) -- but RE-WRITING an EXISTING
' property never persists, via `.Value =` (the pattern WriteDeckReference
' used) or via Delete+Add. `StampPairing` calls WriteDeckReference on every
' repoint, so a workbook ever re-paired to a different deck after its first
' stamp would have silently kept reporting the OLD deck's identity forever
' -- exactly the cross-wiring risk this mechanism exists to close.
'
' Now a cell on a dedicated, very-hidden meta sheet, using the SAME plain
' Cells/Range writes this whole module already relies on for every other
' piece of register data -- proven reliable on OneDrive by this project's
' entire operating history, not freshly probed, since that is the
' mechanism the register has always used.
Private Const META_SHEET_NAME As String = "DeckSyncMeta"

' XlDirection enum values, as numeric literals rather than the named
' constants (xlToLeft/xlUp) -- confirmed real (2026-07-25) that the named
' forms only resolve when this module runs inside Excel's own VBA project
' (which has the Excel type library referenced natively). Driven cross-app
' from PowerPoint (RunSync.bas's actual real usage, per vba-port.md's "VBA
' runs inside Excel or drives it via COM from the PowerPoint side"), the
' PowerPoint-hosted project has no such reference, and the named constants
' raise a compile error ("Variable not defined") -- found via a real
' PowerPoint-driven end-to-end test, not caught by any of ExcelOutput's own
' prior tests since those all ran inside Excel's own project, where the
' names happened to resolve. Numeric values are stable, documented Office
' constants, unaffected by which host application's project this runs in.
Private Const XL_TO_LEFT As Long = -4159
Private Const XL_UP As Long = -4162
' Same reason as XL_TO_LEFT/XL_UP above -- numeric, not xlSheetVeryHidden,
' for cross-app safety when this module is driven from PowerPoint.
Private Const XL_SHEET_VERY_HIDDEN As Long = 2

' Sheet.Fields/InstanceOrder are Collections (ordered, append-only), not
' Dictionary keys -- matches this project's existing convention for ordered
' lists (SyncOperations.bas's instanceOrder) rather than relying on
' Scripting.Dictionary's de-facto-but-undocumented key-insertion order.
' Sheet.Rows is a Scripting.Dictionary of Scripting.Dictionaries
' (instanceId -> fieldName -> value) -- legal because Dictionary/Collection
' are Objects, not UDTs; see Onboarding.bas's SPIKE_NOTES for why a
' Dictionary could NOT hold this data if the values were a UDT instead of
' another Dictionary.
Public Type Sheet
    DeckReference As String
    Fields As Collection
    InstanceOrder As Collection
    Rows As Object
    ' Two rows for the SAME instance in the SAME period. A data error, and the
    ' kind this project keeps being bitten by: whichever sat lower in the sheet
    ' silently won, and a partial read looked like a whole one. Counted so a
    ' caller can refuse rather than quietly pick one.
    DuplicateInstances As Long
End Type

' THE WIDE SHEET CARRIES ITS OWN PERIOD, one row per slide per period.
'
' Rohan's model, 2026-08-03: "the row that has the fields from the selected
' slide in it" is what a person sees and manages. Rows accumulate -- FY26Q4 and
' FY27Q1 for the same project sit side by side -- and a deck picks up the rows
' for the period it declares.
'
' This is what replaces the long register's per-row Quarter. It is also why
' cadence stops being a storage concept: rolling forward COPIES a period's rows
' into the next period, so a project name carries across by being copied rather
' than by a Quarter = ALL sentinel nobody had been told about. See the
' project-deck-sync-object-model memory for the reasoning.
Public Const QUARTER_HEADER As String = "Quarter"

' Kind values that must NEVER get a register column.
'
' A Derived field is computed from other fields -- elapsed time from the start
' and end dates, a status badge from two other values. Giving it a column gives
' it somewhere to be STORED, and a stored copy of a computed value is exactly
' the drift the Derived kind exists to prevent: it goes stale while the fields
' beside it stay right, and nothing can tell.
Public Const KIND_DERIVED As String = "Derived"

' ---------------------------------------------------------------------
' Create
' ---------------------------------------------------------------------

' Set up `ws` as a fresh, empty data sheet bound to `deckReference`.
' Refuses to reinitialize a sheet that already has a header in A1 -- a
' second "create" against an already-set-up sheet is almost certainly a
' mistake (mirrors create_sheet's FileExistsError; the "file" here is
' represented by the sheet already carrying content, since a live Worksheet
' has no separate "exists on disk yet" concept the way a path does).
Public Sub CreateSheet(ws As Object, deckReference As String)
    If Not IsEmpty(ws.Cells(1, 1).Value) Then
        Err.Raise vbObjectError + 1, "ExcelOutput.CreateSheet", _
            "refusing to initialize an already-set-up sheet (A1 is not empty) -- possible accidental double-create"
    End If

    ws.Cells(1, 1).Value = INSTANCE_ID_HEADER
    ' The period column is written HERE, at creation, and never retrofitted.
    ' It cannot be added on its own: UpsertRow refuses to write a row into a
    ' sheet that has this header without saying which period the row is for,
    ' and a sheet carrying the header with blank period cells returns ZERO rows
    ' from every filtered read -- reported as a clean sync of nothing. Header
    ' and period-aware write are one change; see UpsertRow.
    ws.Cells(1, 2).Value = QUARTER_HEADER
    WriteDeckReference ws.Parent, deckReference
End Sub

' Every sheet in this workbook that is ALREADY A REGISTER, named with its size.
'
' WHY THIS EXISTS. Onboarding derives its worksheet name from the SLIDE TYPE
' name (BatchOnboardFlow: `GetOrAddWorksheet(wb, SanitizeSheetName(slideType))`)
' and GetOrAddWorksheet CREATES a missing sheet. So onboarding a real deck as
' type `project-status` against a workbook whose register is named `Register`
' invents a second, empty register and pairs the deck to THAT -- while every
' drafted row sits in the original, untouched and unread.
'
' Nothing is destroyed. Every subsequent read returns zero rows, which this
' codebase reports as a clean run of nothing. That is the worst shape a defect
' can take here and it is the one this project keeps paying for: a success
' message over a file nobody wrote to.
'
' Found 2026-08-13 before it fired, against a register holding 129 drafted rows
' that exist in exactly one file and have never been published anywhere.
'
' A sheet is a register if its HEADER ROW carries INSTANCE_ID_HEADER in ANY
' column -- not if A1 does.
'
' The first version checked A1 only, because that is where CreateSheet writes
' it. Test_ExcelOutput_PeriodRowsAndRollForward caught it: that fixture puts
' `Instance ID` in column B deliberately, to prove the reader finds structural
' columns by NAME and never by position. A locator that disagrees with the
' reader is the exact failure class this file already carries a comment about
' (the sheet read by tab position, the register read by index), so this scans
' the row the way ReadSheetForPeriod does.
'
' RETURNS A DESCRIPTION, NOT A COUNT. Fix-list 1a: a true count with no subject
' sends people to check the wrong thing. Each line is
' `SheetName|dataRows|periods` so the caller can name what it found.
Public Function RegisterShapedSheets(wb As Object) As String
    Dim out As String

    Dim sh As Object
    For Each sh In wb.Worksheets
        Dim isRegister As Boolean
        isRegister = False
        Dim hc As Long
        On Error Resume Next
        For hc = 1 To LastUsedColumn(sh)
            If StrComp(Trim$(CStr(sh.Cells(1, hc).Value)), INSTANCE_ID_HEADER, vbTextCompare) = 0 Then
                isRegister = True
                Exit For
            End If
        Next hc
        On Error GoTo 0

        If isRegister Then
            Dim lastRow As Long
            lastRow = LastUsedRow(sh)

            Dim dataRows As Long
            dataRows = 0
            If lastRow > 1 Then dataRows = lastRow - 1

            ' Periods present, so the caller can say WHICH register it found
            ' rather than merely that there was one. A person choosing between
            ' two sheets needs the quarters to tell them apart.
            Dim seen As Object
            Set seen = CreateObject("Scripting.Dictionary")
            Dim cQuarter As Long
            cQuarter = 0
            Dim c As Long
            For c = 1 To LastUsedColumn(sh)
                If StrComp(Trim$(CStr(sh.Cells(1, c).Value)), QUARTER_HEADER, vbTextCompare) = 0 Then cQuarter = c
            Next c

            Dim periods As String
            periods = ""
            If cQuarter > 0 Then
                Dim r As Long
                For r = 2 To lastRow
                    Dim q As String
                    q = Trim$(CStr(sh.Cells(r, cQuarter).Value))
                    If q <> "" Then
                        If Not seen.Exists(UCase(q)) Then
                            seen(UCase(q)) = True
                            If periods <> "" Then periods = periods & ", "
                            periods = periods & q
                        End If
                    End If
                Next r
            End If
            If periods = "" Then periods = "no periods"

            If out <> "" Then out = out & vbLf
            out = out & sh.Name & "|" & dataRows & "|" & periods
        End If
    Next sh

    RegisterShapedSheets = out
End Function

' How many data rows this sheet holds for `period`. 0 for a blank period, a
' sheet with no period column, or a sheet that simply has none.
'
' Exists so a caller can say WHAT is at risk before overwriting it, rather than
' warning in the abstract. A person deciding whether to let a harvest run needs
' to know there are 43 rows behind the question, not that "some data exists".
Public Function RowCountForPeriod(ws As Object, period As String) As Long
    If Trim$(period) = "" Then Exit Function

    Dim lastRow As Long
    lastRow = LastUsedRow(ws)
    If lastRow < 2 Then Exit Function

    Dim cQuarter As Long
    cQuarter = 0
    Dim c As Long
    For c = 1 To LastUsedColumn(ws)
        If StrComp(Trim$(CStr(ws.Cells(1, c).Value)), QUARTER_HEADER, vbTextCompare) = 0 Then cQuarter = c
    Next c
    If cQuarter = 0 Then Exit Function

    Dim n As Long
    Dim r As Long
    For r = 2 To lastRow
        If StrComp(Trim$(CStr(ws.Cells(r, cQuarter).Value)), Trim$(period), vbTextCompare) = 0 Then n = n + 1
    Next r

    RowCountForPeriod = n
End Function

' Field Spec fields that have no column in the register yet, comma-separated.
'
' WHY THIS IS NEEDED AT ALL. The only thing that creates a register column is
' UpsertRow appending one on first write -- which happens when a field is
' PUBLISHED from a drafting sheet, or HARVESTED from a tagged shape during
' onboarding. A `Given` field is neither: nobody drafts it, and it is typed
' straight into the register. So a Given field declared in the Field Spec has
' nowhere to be typed, forever, and the Field Spec cheerfully describes a
' column that does not exist.
'
' It bites hardest on the milestone timeline. `MSn_DATE` and `MSn_DONE` are
' Given, and they are addressed by NAME inside a tagged group rather than being
' tagged themselves -- so onboarding never sees them either. 14 of the columns
' that drive the timeline had no route into the register at all.
'
' DERIVED FIELDS ARE EXCLUDED, and that exclusion is the reason this returns a
' list rather than just creating them: a completeness check that demanded a
' column for every Field Spec row would demand one for every Derived field too,
' and then report them as missing forever. A warning that always fires stops
' being read.
Public Function MissingRegisterColumns(specWs As Object, regWs As Object) As String
    If specWs Is Nothing Or regWs Is Nothing Then Exit Function

    ' What the register already has, by header name.
    Dim have As Object
    Set have = CreateObject("Scripting.Dictionary")
    Dim lastCol As Long
    lastCol = LastUsedColumn(regWs)
    Dim c As Long
    For c = 1 To lastCol
        Dim h As String
        h = UCase(Trim$(CStr(regWs.Cells(1, c).Value)))
        If h <> "" Then have(h) = True
    Next c

    Dim out As String
    Dim r As Long
    r = FieldSpec.SPEC_FIRST_ROW
    Do While Trim$(CStr(specWs.Cells(r, FieldSpec.COL_S_FIELDID).Value)) <> ""
        Dim fid As String, kind As String
        fid = Trim$(CStr(specWs.Cells(r, FieldSpec.COL_S_FIELDID).Value))
        kind = Trim$(CStr(specWs.Cells(r, FieldSpec.COL_S_KIND).Value))

        If StrComp(kind, KIND_DERIVED, vbTextCompare) <> 0 Then
            If Not have.Exists(UCase(fid)) Then
                If out <> "" Then out = out & ","
                out = out & fid
            End If
        End If
        r = r + 1
    Loop

    MissingRegisterColumns = out
End Function

' Append a header for each named field. Returns what it actually wrote.
'
' Headers only -- no rows, no values. A column with a header and empty cells is
' a field waiting to be filled, which is the honest state; inventing values
' would be the tool deciding what a Given field says.
Public Function AddRegisterColumns(regWs As Object, fieldNames As String) As String
    If Trim$(fieldNames) = "" Then Exit Function

    Dim parts() As String
    parts = Split(fieldNames, ",")

    Dim added As String
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim nm As String
        nm = Trim$(parts(i))
        If nm <> "" Then
            ' Re-locating the end each time, because the previous iteration moved
            ' it. Caching lastCol outside the loop writes every column on top of
            ' the one before and reports success for all of them.
            Dim nextCol As Long
            nextCol = LastUsedColumn(regWs) + 1
            regWs.Cells(1, nextCol).Value = nm

            ' READ IT BACK. A header that did not land leaves the field exactly
            ' as unwritable as before, behind a message saying it was fixed.
            If StrComp(Trim$(CStr(regWs.Cells(1, nextCol).Value)), nm, vbTextCompare) = 0 Then
                If added <> "" Then added = added & ", "
                added = added & nm
            End If
        End If
    Next i

    AddRegisterColumns = added
End Function

' Locate the sheet's two structural columns by header name.
'
' SHARED BY THE READER AND THE WRITER ON PURPOSE. They used to locate columns
' separately -- the reader by name, the writer by hardcoded position 1 -- and a
' reader and writer that disagree about which column means what is the failure
' class this project has paid for repeatedly (the sheet read by tab position,
' the register read by index). One function, one answer, both callers.
'
' `cQuarter` comes back 0 for a sheet built before 2026-08-03, which has no
' period column and one row per slide. That is a legal shape and both callers
' handle it; it is not an error.
' PUBLIC since 2026-08-05, for FieldSpec.ApplyControlledValidation. It is the
' third caller that needs to know which columns are structural and which are
' fields, and the alternative was a third implementation of that answer --
' which is the exact drift this function was made shared to prevent. Anything
' that needs to tell a field column from a structural one comes here.
Public Sub LocateStructuralColumns(ws As Object, ByRef cInstance As Long, ByRef cQuarter As Long)
    cInstance = 0
    cQuarter = 0

    Dim lastCol As Long
    lastCol = LastUsedColumn(ws)

    Dim c As Long
    For c = 1 To lastCol
        Dim h As String
        h = Trim$(CStr(ws.Cells(1, c).Value))
        If StrComp(h, INSTANCE_ID_HEADER, vbTextCompare) = 0 Then
            cInstance = c
        ElseIf StrComp(h, QUARTER_HEADER, vbTextCompare) = 0 Then
            cQuarter = c
        End If
    Next c

    ' A sheet whose first column is the instance but is not headed as such is
    ' every sheet this tool wrote before headers were read by name. Falling back
    ' to column 1 keeps those readable; it is not a guess about arbitrary sheets,
    ' because CreateSheet has always written INSTANCE_ID_HEADER into A1.
    If cInstance = 0 Then cInstance = 1
End Sub

' ---------------------------------------------------------------------
' Read
' ---------------------------------------------------------------------

' Read `ws` back into a Sheet. Fields are recovered from the header row
' (row 1, columns B..), instance rows from column A (rows 2..) -- both via
' Excel's End(xlToLeft)/End(xlUp) "walk from the far side" idiom rather than
' a stored count, the standard reliable way to find a used range's true
' extent in the object model.
' Unfiltered read -- every row, whatever period it carries. Correct for a sheet
' with no Quarter column at all (which is every sheet built before 2026-08-03).
Public Function ReadSheet(ws As Object) As Sheet
    ReadSheet = ReadSheetForPeriod(ws, "")
End Function

' Reads `ws` back into a Sheet, keeping only the rows for `deckPeriod`.
'
' `deckPeriod = ""` means "no filter" -- used by the unfiltered reader above and
' by any sheet that predates the Quarter column. A sheet WITHOUT a Quarter
' column is never filtered: it has one row per instance and no opinion about
' periods, so filtering it would return nothing and an empty read is a legal
' state that reads as success. That failure has already cost this project two
' evenings; it does not get a third.
'
' COLUMNS BY HEADER NAME, NEVER BY POSITION. The previous version took the
' instance from column 1 and every other column as a field, which meant adding
' any column anywhere shifted what a field was called. That rule is written in
' three places in this codebase and has been broken twice by the code beneath
' it -- once reading the sheet by tab position, once reading a sheet by index.
Public Function ReadSheetForPeriod(ws As Object, deckPeriod As String) As Sheet
    Dim result As Sheet
    Set result.Fields = New Collection
    Set result.InstanceOrder = New Collection
    Set result.Rows = CreateObject("Scripting.Dictionary")

    result.DeckReference = ReadDeckReference(ws.Parent)

    Dim lastCol As Long
    lastCol = LastUsedColumn(ws)
    If lastCol = 0 Then
        ReadSheetForPeriod = result
        Exit Function
    End If

    ' Locate the two structural columns by name; everything else is a field.
    Dim cInstance As Long, cQuarter As Long
    LocateStructuralColumns ws, cInstance, cQuarter

    Dim c As Long
    For c = 1 To lastCol
        If c <> cInstance And c <> cQuarter Then
            result.Fields.Add CStr(ws.Cells(1, c).Value)
        End If
    Next c

    Dim filtering As Boolean
    filtering = (deckPeriod <> "") And (cQuarter > 0)

    Dim lastRow As Long
    lastRow = LastUsedRow(ws)
    Dim r As Long
    For r = 2 To lastRow
        Dim instanceId As String
        instanceId = Trim$(CStr(ws.Cells(r, cInstance).Value))
        If instanceId <> "" Then
            Dim keep As Boolean
            keep = True
            If filtering Then
                keep = (StrComp(Trim$(CStr(ws.Cells(r, cQuarter).Value)), deckPeriod, vbTextCompare) = 0)
            End If

            If keep Then
                If result.Rows.Exists(instanceId) Then
                    ' Same project, same period, twice. First one wins so the
                    ' result is at least deterministic, and it is COUNTED so a
                    ' caller can refuse -- silently taking the lower row is how
                    ' the cadence collision in the long register used to behave,
                    ' and that was judged a defect there too.
                    result.DuplicateInstances = result.DuplicateInstances + 1
                Else
                    result.InstanceOrder.Add instanceId

                    Dim rowValues As Object
                    Set rowValues = CreateObject("Scripting.Dictionary")
                    Dim fi As Long
                    fi = 0
                    For c = 1 To lastCol
                        If c <> cInstance And c <> cQuarter Then
                            fi = fi + 1
                            ' IsEmpty (not "= """"") to distinguish "field never
                            ' harvested for this instance" from "harvested value
                            ' happens to be an empty string" -- mirrors
                            ' read_sheet's structural cell-presence check.
                            If Not IsEmpty(ws.Cells(r, c).Value) Then
                                rowValues(result.Fields(fi)) = CStr(ws.Cells(r, c).Value)
                            End If
                        End If
                    Next c
                    Set result.Rows(instanceId) = rowValues
                End If
            End If
        End If
    Next r

    ReadSheetForPeriod = result
End Function

Public Function LastUsedColumn(ws As Object) As Long
    If IsEmpty(ws.Cells(1, 1).Value) Then
        LastUsedColumn = 0
        Exit Function
    End If
    LastUsedColumn = ws.Cells(1, ws.Columns.count).End(XL_TO_LEFT).Column
End Function

Public Function LastUsedRow(ws As Object) As Long
    If IsEmpty(ws.Cells(1, 1).Value) Then
        LastUsedRow = 0
        Exit Function
    End If
    LastUsedRow = ws.Cells(ws.Rows.count, 1).End(XL_UP).Row
End Function

' PUBLIC since 2026-08-14. These were Private and called from exactly one place
' (CreateSheet, at onboarding), which is why the GUID was written once per
' workbook and then never maintained or consulted again. DeckRegistry needs both
' directions to keep the pairing mutually verifiable across a repoint.

' Nothing (never creates) -- a read-only lookup must not spring a hidden
' sheet into existence on a workbook that has never been stamped.
Private Function FindMetaSheet(wb As Object) As Object
    Dim ws As Object
    On Error Resume Next
    For Each ws In wb.Worksheets
        If ws.Name = META_SHEET_NAME Then
            Set FindMetaSheet = ws
            Exit Function
        End If
    Next ws
    On Error GoTo 0
End Function

' Appended at the END so it never disturbs the position of a real sheet
' anything else in this project addresses by name or index. Very-hidden
' (not just hidden) deliberately: this holds the GUID the cross-wiring check
' depends on, and that check is worth more protection from an accidental
' right-click-unhide-and-delete than a normal hidden sheet gets.
Private Function FindOrCreateMetaSheet(wb As Object) As Object
    Dim existing As Object
    Set existing = FindMetaSheet(wb)
    If Not existing Is Nothing Then
        Set FindOrCreateMetaSheet = existing
        Exit Function
    End If

    Dim ws As Object
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
    ws.Name = META_SHEET_NAME
    ws.Visible = XL_SHEET_VERY_HIDDEN
    Set FindOrCreateMetaSheet = ws
End Function

Public Sub WriteDeckReference(wb As Object, deckReference As String)
    Dim ws As Object
    Set ws = FindOrCreateMetaSheet(wb)
    ws.Cells(1, 1).Value = deckReference
End Sub

' Tries the meta sheet first; falls back to the pre-2026-08-16
' CustomDocumentProperties location so workbooks stamped before this change
' keep reading correctly until the next repoint moves them onto the new
' mechanism for good.
Public Function ReadDeckReference(wb As Object) As String
    Dim ws As Object
    Set ws = FindMetaSheet(wb)
    If Not ws Is Nothing Then
        Dim v As Variant
        v = ws.Cells(1, 1).Value
        If Not IsEmpty(v) And Not IsNull(v) Then
            If CStr(v) <> "" Then
                ReadDeckReference = CStr(v)
                Exit Function
            End If
        End If
    End If

    ReadDeckReference = ReadOldDeckReferenceProperty(wb)
End Function

' THE OLD MECHANISM. Read-only fallback now -- see META_SHEET_NAME's
' comment for why nothing writes here any more.
Private Function ReadOldDeckReferenceProperty(wb As Object) As String
    Dim prop As Object
    On Error Resume Next
    Set prop = wb.CustomDocumentProperties(DECK_REFERENCE_PROPERTY_NAME)
    On Error GoTo 0

    If prop Is Nothing Then
        ReadOldDeckReferenceProperty = ""
    Else
        ReadOldDeckReferenceProperty = CStr(prop.Value)
    End If
End Function

' ---------------------------------------------------------------------
' Upsert
' ---------------------------------------------------------------------

' Add any of `values`'s keys not already a known field as a new column
' (appended after the last used column, never replacing/reordering existing
' ones), then create or update `instanceId`'s row -- direct incremental
' Cells writes, not a read-whole-sheet/mutate/rewrite-whole-file cycle
' (unlike upsert_row, which must rebuild the entire .xlsx because that's the
' only write primitive a headless zip has; a live Worksheet doesn't need
' that). A new instance is appended as a new row, seeded entirely from
' `values`. An existing instance only has the given keys' cells written --
' any field this call doesn't mention is left completely untouched, so a
' partial re-sync of one changed field can never blank out the rest.
'
' `values` is a Scripting.Dictionary (fieldName -> value String), matching
' the shape SyncOperations.bas's `dataRows` entries already use.
'
' Deliberately does NOT apply real typed Excel formatting (a Date/Double
' value + NumberFormat) even though BatchOnboardFlow.bas now captures a
' field type at mark time -- traced 2026-07-26 that this Sheet's own values
' feed directly back into SyncOperations.PlanRoutineSync -> InjectPrimitive,
' which WRITES Excel's value onto the live PowerPoint slide on every
' routine sync. A typed cell's read-back (ReadSheet's CStr(.Value)) is not
' guaranteed to equal the exact string that was written -- a Date reads
' back locale-formatted, a Double can drop a trailing zero -- so writing a
' real typed value here would risk a routine sync silently rewriting a
' slide's date/number text into a reformatted (though "equal") version the
' human never asked for. That's a slide-content mutation, not a formatting
' nicety, and conflicts with this project's founding invariant that nothing
' gets silently mutated (InjectPrimitive/Verification exist specifically to
' guard it). Every value here is always written and read back as the exact
' harvested string, unconditionally -- the field type is still captured and
' shown to a human (BatchOnboardFlow's Field Review grid), just not acted
' on here. Revisit only alongside making PlanRoutineSync's own comparison
' type-aware, not by touching this function in isolation.
' `period` IS REQUIRED, AND THIS IS THE WHOLE POINT OF THE PARAMETER.
'
' A row is a SLIDE IN A PERIOD, so a slide's identity on this sheet is
' (instance, period) and never instance alone. Matching on instance alone --
' which is what this did until 2026-08-04 -- means syncing FY27Q1 finds
' FY26Q4's row and OVERWRITES IT. That is a real quarter's approved text
' destroyed silently, on rollover, which is the single most expensive thing
' this code could do.
'
' Passing "" is refused rather than defaulted, on a sheet that has a period
' column. A blank period cell is invisible to every filtered read, so the row
' would exist, contain the right text, and never appear in a sync again --
' reported as a clean sync of nothing. An Optional parameter defaulting to ""
' would make that the failure a caller gets by FORGETTING, which is exactly
' backwards for the one operation that can lose work.
'
' A sheet with NO period column (everything built before 2026-08-03) keeps the
' old behaviour: one row per slide, matched on instance, `period` ignored. Such
' a sheet has no opinion about periods and retrofitting one is a migration, not
' something an upsert should do behind the caller's back.
' `asText` forces the destination cell to Text format BEFORE the value is
' assigned, so Excel stores the characters it was given instead of parsing them.
'
' Added 2026-08-14, from the first real harvest. A slide reading "30 Oct 2023"
' arrived in the register as 45229 and "$275,598" as 275598 -- the right values,
' with the formatting silently gone, so publishing them back would put a serial
' number on a funder-facing slide. `.Value = CStr(...)` looks like a string write
' and is not: Excel parses anything that looks like a date or a number.
'
' OPT-IN rather than the default, because this is only obviously correct for a
' caller whose contract is "store what was on the slide". The harvest's is
' exactly that. Existing callers keep the old behaviour and are unaffected.
'
' The damage this prevents is PERMANENT once done: the harvest writes only into
' empty cells, so a coerced value cannot be corrected by re-harvesting -- the
' cell is no longer empty. Getting this right before a bulk run matters more
' than it would for an ordinary bug.
Public Sub UpsertRow(ws As Object, instanceId As String, values As Object, period As String, _
                     Optional asText As Boolean = False)
    Dim cInstance As Long, cQuarter As Long
    LocateStructuralColumns ws, cInstance, cQuarter

    If cQuarter > 0 And Trim$(period) = "" Then
        Err.Raise vbObjectError + 4, "ExcelOutput.UpsertRow", _
            "this sheet has a '" & QUARTER_HEADER & "' column, so every row must say which " & _
            "period it belongs to. A row written with a blank period is invisible to every " & _
            "filtered read and would report as a clean sync of nothing."
    End If

    Dim rowNum As Long
    rowNum = FindOrAppendInstanceRow(ws, instanceId, period, cInstance, cQuarter)

    Dim fieldName As Variant
    For Each fieldName In values.Keys
        Dim colNum As Long
        colNum = FindOrAppendFieldColumn(ws, CStr(fieldName), cInstance, cQuarter)
        ' BEFORE the assignment, not after -- setting Text format on a cell that
        ' already holds a parsed date shows the serial, it does not recover the
        ' original characters.
        If asText Then ws.Cells(rowNum, colNum).NumberFormat = "@"
        ws.Cells(rowNum, colNum).Value = CStr(values(fieldName))
    Next fieldName
End Sub

' Refuses a field whose name collides with a structural column, rather than
' writing a slide's text over the row's identity or its period. A slide really
' can carry a field called "Quarter"; before this, that value went straight
' into the period cell and the row vanished from every filtered read.
Private Function FindOrAppendFieldColumn(ws As Object, fieldName As String, _
                                         cInstance As Long, cQuarter As Long) As Long
    If StrComp(Trim$(fieldName), INSTANCE_ID_HEADER, vbTextCompare) = 0 Or _
       StrComp(Trim$(fieldName), QUARTER_HEADER, vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 5, "ExcelOutput.UpsertRow", _
            "'" & fieldName & "' is the name of a structural column on this sheet, not a " & _
            "field. Writing a slide's value into it would overwrite the row's identity or " & _
            "the period it belongs to."
    End If

    Dim lastCol As Long
    lastCol = LastUsedColumn(ws)
    Dim c As Long
    For c = 1 To lastCol
        If c <> cInstance And c <> cQuarter Then
            If CStr(ws.Cells(1, c).Value) = fieldName Then
                FindOrAppendFieldColumn = c
                Exit Function
            End If
        End If
    Next c

    FindOrAppendFieldColumn = lastCol + 1
    ws.Cells(1, lastCol + 1).Value = fieldName
End Function

' The row for THIS slide IN THIS PERIOD, or a new one carrying both.
'
' Instance is matched exactly (Trim only) to agree with the reader, whose
' Dictionary keys are case-sensitive -- two keys differing by case are two
' slides, and merging them here would silently join two projects' rows. Period
' is matched case-insensitively, also to agree with the reader's filter, so
' "fy26q4" and "FY26Q4" resolve to one row rather than quietly becoming two.
Private Function FindOrAppendInstanceRow(ws As Object, instanceId As String, period As String, _
                                         cInstance As Long, cQuarter As Long) As Long
    Dim lastRow As Long
    lastRow = LastUsedRow(ws)

    Dim r As Long
    For r = 2 To lastRow
        If Trim$(CStr(ws.Cells(r, cInstance).Value)) = Trim$(instanceId) Then
            If cQuarter = 0 Then
                FindOrAppendInstanceRow = r
                Exit Function
            ElseIf StrComp(Trim$(CStr(ws.Cells(r, cQuarter).Value)), Trim$(period), vbTextCompare) = 0 Then
                FindOrAppendInstanceRow = r
                Exit Function
            End If
        End If
    Next r

    FindOrAppendInstanceRow = lastRow + 1
    ws.Cells(lastRow + 1, cInstance).Value = instanceId
    If cQuarter > 0 Then ws.Cells(lastRow + 1, cQuarter).Value = period
End Function

' ---------------------------------------------------------------------
' Manual smoke test -- not a real test harness, same as every other module
' here. See SPIKE_NOTES_ExcelOutput.md for the full recipe and expected
' values, cross-checked against tests/test_excel_output.py's already-proven
' round-trip results.
' ---------------------------------------------------------------------

' Run against a blank worksheet in the active workbook (e.g. add a new
' sheet first so A1 is genuinely empty).
'
' EVERY UPSERT NAMES ITS PERIOD, because CreateSheet above writes the Quarter
' header and UpsertRow refuses a blank period on a sheet that has one. These
' three calls passed three arguments until 2026-08-05 -- left behind when the
' period became required -- which is "Argument not optional" at COMPILE time,
' and a compile error in one module stops the whole project, not just the smoke
' test. The suite could not see it: run_vba_tests.ps1 has no compile step, so it
' reported 152 passed against a project that would not start. Same blind spot as
' the Optional-Variant-into-ByRef-Object bug on 2026-08-01.
Public Sub ManualSmokeTest(ws As Object)
    Const SMOKE_PERIOD As String = "FY26Q4"

    CreateSheet ws, "deck-v1"

    Dim v1 As Object
    Set v1 = CreateObject("Scripting.Dictionary")
    v1("Title") = "Q3 Revenue"
    v1("Date") = "2026-07"
    UpsertRow ws, "slide-1", v1, SMOKE_PERIOD

    Dim v2 As Object
    Set v2 = CreateObject("Scripting.Dictionary")
    v2("Region") = "APAC"
    UpsertRow ws, "slide-1", v2, SMOKE_PERIOD ' new field, existing instance -- appends a column, doesn't disturb Title/Date

    Dim v3 As Object
    Set v3 = CreateObject("Scripting.Dictionary")
    v3("Title") = "Q4 Revenue"
    UpsertRow ws, "slide-2", v3, SMOKE_PERIOD ' new instance -- new row, no Date/Region yet

    Dim sheet As Sheet
    sheet = ReadSheet(ws)

    Dim msg As String
    msg = "DeckReference=" & sheet.DeckReference & vbCrLf
    msg = msg & "Fields=" & JoinCollection(sheet.Fields) & vbCrLf
    msg = msg & "InstanceOrder=" & JoinCollection(sheet.InstanceOrder) & vbCrLf
    msg = msg & "slide-1: Title=" & sheet.Rows("slide-1")("Title") & " Date=" & sheet.Rows("slide-1")("Date") & " Region=" & sheet.Rows("slide-1")("Region") & vbCrLf
    msg = msg & "slide-2: Title=" & sheet.Rows("slide-2")("Title") & " HasDate=" & sheet.Rows("slide-2").Exists("Date")

    Debug.Print msg
    MsgBox msg & vbCrLf & "(expected: DeckReference=deck-v1, Fields=Title,Date,Region, InstanceOrder=slide-1,slide-2, slide-2 HasDate=False)"
End Sub

Private Function JoinCollection(coll As Collection) As String
    Dim i As Long, result As String
    For i = 1 To coll.count
        If i > 1 Then result = result & ","
        result = result & coll(i)
    Next i
    JoinCollection = result
End Function

' Read `ws` as the deck's declared period, and say why the result must NOT be
' used -- `problem` is "" only when it is safe.
'
' THE SYNC PATH MUST USE THIS, NOT ReadSheet. Every sync-side read in RibbonUI
' called the unfiltered ReadSheet, which was correct while a sheet held one row
' per slide and became wrong the moment rows started accumulating per period:
' the same slide now appears once per period, the unfiltered read keeps
' whichever sits higher, and the rest land in DuplicateInstances -- a counter
' nothing looked at. A deck would render one period's rows mixed with a
' silently-discarded period's, and report a clean sync.
'
' Two ways a read is wrong while looking fine, and this project has paid for
' both already:
'
'   - TWO ROWS FOR ONE SLIDE. Whichever sat lower used to win silently. Counted
'     by ReadSheetForPeriod; refused here.
'   - THE SHEET HAS ROWS AND NONE OF THEM MATCH. Zero rows is a legal state that
'     reads as a clean sync of nothing -- the exact failure that cost two
'     evenings when a register was read by tab position. Detected by comparing
'     against the unfiltered read, which is only meaningful in this direction:
'     a sheet with no Quarter column is never filtered, so the two are equal and
'     this stays silent. It fires only when filtering genuinely happened and
'     matched nothing.
'
' `problem` is a STRING rather than a Boolean because the caller has to be able
' to tell a person which sheet, and why.
Public Function ReadSheetForDeckPeriod(ws As Object, deckPeriod As String, _
                                       ByRef problem As String) As Sheet
    Dim wanted As Sheet
    problem = ""

    ' A1 empty means this worksheet has never been through CreateSheet -- it is
    ' not "a Data sheet with nothing synced yet", it is GetOrAddWorksheet having
    ' just CREATED a blank tab under a name nothing set up. That is exactly what
    ' happens when a registered type's worksheet name (e.g. "q") does not match
    ' the sheet a person actually built (e.g. "Register"): the read silently
    ' succeeds against an empty sheet it just invented. Checked before reading,
    ' because an empty read of THIS sheet and an empty read of a genuinely
    ' freshly-onboarded one are indistinguishable once both come back as zero
    ' rows -- only the header row tells them apart.
    If IsEmpty(ws.Cells(1, 1).Value) Then
        problem = "worksheet '" & ws.Name & "' has never been set up (A1 is empty) -- " & _
            "this is not a Data sheet with nothing on it yet, it looks like the wrong " & _
            "sheet was resolved and an empty one was created in its place."
        ReadSheetForDeckPeriod = wanted
        Exit Function
    End If

    wanted = ReadSheetForPeriod(ws, deckPeriod)

    If wanted.DuplicateInstances > 0 Then
        problem = wanted.DuplicateInstances & " row(s) repeat a slide already read for '" & _
            deckPeriod & "'. One row per slide per period is the whole model, so this " & _
            "sheet cannot be synced until the repeats are removed -- otherwise whichever " & _
            "row sits higher wins and nothing says so."
    ElseIf wanted.InstanceOrder.count = 0 Then
        Dim everything As Sheet
        everything = ReadSheetForPeriod(ws, "")
        If everything.InstanceOrder.count > 0 Then
            problem = "the sheet holds " & everything.InstanceOrder.count & " slide(s), but not " & _
                "one of them is stamped '" & deckPeriod & "'. Reading zero rows is a legal " & _
                "state and would report as a clean sync of nothing, so it is refused instead."
        End If
    End If

    ReadSheetForDeckPeriod = wanted
End Function

' Copies every row for `fromPeriod` into a new set of rows stamped `toPeriod`.
'
' THIS IS WHAT REPLACES Quarter = ALL. The sentinel existed so a project name
' would not need retyping every quarter. Copying the row forward achieves the
' same thing with no concept attached: static values arrive already correct,
' variable values arrive as last period's text and get rewritten through the
' drafting sheet -- where last period's value is exactly the exemplar a drafter
' wants in front of them anyway.
'
' Refuses when `toPeriod` already has rows. Rolling forward twice would double
' every project, and the second run would look identical to the first.
' HOW MANY ROWS A PERIOD ALREADY HAS, so a caller can decide BEFORE asking a
' person anything. RollForwardPeriod refuses outright when the destination is
' already populated -- correctly -- but it only discovers that after the caller
' has demanded an answer to "which period should they be copied FROM?". On the
' real deck that meant a modal and a free-text prompt whose every possible
' answer led to the same refusal. Rohan, 2026-08-13, on the run that found it:
' "get rid of needless popup messages."
'
' Returns 0 on a sheet with no Quarter column, which is the honest answer: such a
' sheet holds no periods, so this period has no rows on it.
' Which column holds the period, found by reading the header rather than assuming
' a position. Returns 0 when this sheet has no Quarter column at all.
'
' The search itself is not new -- it is written out inline in five places in this
' module. This is the same search with a name, added 2026-08-14 for a caller
' outside this module; the five existing sites are deliberately left alone rather
' than swept up in an unrelated change.
Public Function QuarterColumn(ws As Object) As Long
    Dim c As Long
    For c = 1 To 64
        If StrComp(Trim$(CStr(ws.Cells(1, c).Value)), QUARTER_HEADER, vbTextCompare) = 0 Then
            QuarterColumn = c
            Exit Function
        End If
    Next c
End Function

Public Function PeriodRowCount(ws As Object, period As String) As Long
    If Trim$(period) = "" Then Exit Function

    Dim lastCol As Long
    lastCol = LastUsedColumn(ws)

    Dim cQuarter As Long, c As Long
    For c = 1 To lastCol
        If StrComp(Trim$(CStr(ws.Cells(1, c).Value)), QUARTER_HEADER, vbTextCompare) = 0 Then cQuarter = c
    Next c
    If cQuarter = 0 Then Exit Function

    Dim lastRow As Long, r As Long, n As Long
    lastRow = LastUsedRow(ws)
    For r = 2 To lastRow
        If StrComp(Trim$(CStr(ws.Cells(r, cQuarter).Value)), period, vbTextCompare) = 0 Then n = n + 1
    Next r

    PeriodRowCount = n
End Function

' The largest number of rows any single period holds -- i.e. what a COMPLETE
' quarter looks like in this register.
'
' Exists so a PARTIAL period can be recognised before it causes a refusal.
' 2026-08-15: five leftover Q1F27 stub rows from earlier testing silently refused
' a quarter turn, because RollForwardPeriod correctly declines to duplicate 43
' projects into a period that already holds rows. Five is not a quarter and not
' nothing, and there was no way to see that without counting by hand. Comparing
' against the largest period is what makes "5 of 43" legible as partial.
Public Function LargestPeriodRowCount(ws As Object) As Long
    Dim cQuarter As Long
    cQuarter = QuarterColumn(ws)
    If cQuarter = 0 Then Exit Function

    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")

    Dim lastRow As Long, r As Long
    lastRow = LastUsedRow(ws)
    For r = 2 To lastRow
        Dim q As String
        q = Trim$(CStr(ws.Cells(r, cQuarter).Value))
        If q <> "" Then
            If seen.Exists(q) Then
                seen(q) = seen(q) + 1
            Else
                seen.Add q, 1
            End If
        End If
    Next r

    Dim k As Variant
    For Each k In seen.Keys
        If seen(k) > LargestPeriodRowCount Then LargestPeriodRowCount = seen(k)
    Next k
End Function

Public Function RollForwardPeriod(ws As Object, fromPeriod As String, toPeriod As String) As String
    If Trim$(fromPeriod) = "" Or Trim$(toPeriod) = "" Then
        Err.Raise vbObjectError + 3, "ExcelOutput.RollForwardPeriod", _
            "both the period being copied from and the period being created must be named"
    End If
    If StrComp(fromPeriod, toPeriod, vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 3, "ExcelOutput.RollForwardPeriod", _
            "cannot roll " & fromPeriod & " forward into itself"
    End If

    Dim lastCol As Long
    lastCol = LastUsedColumn(ws)

    Dim cQuarter As Long
    Dim c As Long
    For c = 1 To lastCol
        If StrComp(Trim$(CStr(ws.Cells(1, c).Value)), QUARTER_HEADER, vbTextCompare) = 0 Then cQuarter = c
    Next c

    If cQuarter = 0 Then
        Err.Raise vbObjectError + 3, "ExcelOutput.RollForwardPeriod", _
            "this sheet has no '" & QUARTER_HEADER & "' column, so it holds no periods to roll forward"
    End If

    Dim lastRow As Long
    lastRow = LastUsedRow(ws)

    ' Collect the source rows FIRST. Appending while walking would re-read the
    ' rows just written and copy them again, forever.
    Dim src As Collection
    Set src = New Collection
    Dim existingTarget As Long
    Dim r As Long
    For r = 2 To lastRow
        Dim q As String
        q = Trim$(CStr(ws.Cells(r, cQuarter).Value))
        If StrComp(q, fromPeriod, vbTextCompare) = 0 Then src.Add r
        If StrComp(q, toPeriod, vbTextCompare) = 0 Then existingTarget = existingTarget + 1
    Next r

    If existingTarget > 0 Then
        RollForwardPeriod = "REFUSED: " & toPeriod & " already has " & existingTarget & _
            " row(s). Rolling forward again would duplicate every project."
        Exit Function
    End If

    If src.count = 0 Then
        RollForwardPeriod = "Nothing to do: no rows found for " & fromPeriod & "."
        Exit Function
    End If

    Dim outRow As Long
    outRow = lastRow + 1

    Dim v As Variant
    For Each v In src
        Dim from As Long
        from = CLng(v)
        For c = 1 To lastCol
            If c = cQuarter Then
                ws.Cells(outRow, c).Value = toPeriod
            Else
                ' Value, not formula or format: this is data being carried, and
                ' a copied format would drag the source row's styling with it.
                ws.Cells(outRow, c).Value = ws.Cells(from, c).Value
            End If
        Next c
        outRow = outRow + 1
    Next v

    RollForwardPeriod = src.count & " row(s) copied from " & fromPeriod & " to " & toPeriod & "." & vbCrLf & _
        "Static text arrives already correct. Everything that changes this period is" & vbCrLf & _
        "last period's text until you rewrite it -- which is what the drafting sheet" & vbCrLf & _
        "shows you in the ORIGINAL column."
End Function
