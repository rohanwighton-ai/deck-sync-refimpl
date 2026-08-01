Attribute VB_Name = "Timeline"
Option Explicit

' A TIME-ELAPSED BAR: two rectangles whose widths always sum to the same total.
'
' Rohan's design, 2026-08-01: "the calculation just takes the time difference
' between the two dates and calculates elapsed and remaining bar lengths...
' The two parts added together always equal the same total bar length, the
' elapsed gets longer and remaining gets shorter."
'
' NOBODY AUTHORS THIS FIELD. It is a pure function of (start, end, as-at), so it
' is a fourth Kind -- Derived. It needs no drafting sheet, no approval, and no
' register value of its own. The register holds the two DATES; the bar is a
' rendering of them.
'
' AS-AT IS THE REPORTING PERIOD'S END, NOT TODAY, and that is the one decision
' here worth arguing about. Rohan's first instinct was to update on file open,
' since dates advance whether or not anything else does. But this is a QUARTERLY
' REPORT: if the bar tracks the viewing date, a deck approved in August shows a
' different bar when opened in October, and a slide that was signed off silently
' stops saying what it said. Pinning to the period makes it reproducible -- open
' it in 2030 and it still shows the truth as at FY26Q4.
'
' It also avoids a practical trap: a deck that recomputes on open is DIRTY on
' open, so every close prompts to save -- on a 49MB deck whose saves this project
' already documents as unreliable enough to need a write-verify-retry wrapper.
' Recomputing during sync keeps one write path and one approval surface.
'
' The two shapes are tagged as ORDINARY FIELDS, one each. No new grouping
' concept is needed precisely because the total is invariant: it can be
' recovered from the deck as (elapsed.Width + remaining.Width), so nothing has
' to be stored and resizing the template cannot silently mis-scale anything.

' How far through, as a fraction from 0 to 1.
'
' CLAMPED, because a project can be reported on before it starts and after it
' ends, and both are ordinary rather than exceptional. Unclamped, the fill runs
' off the end of its track and the slide is visibly broken -- or worse, it is a
' negative width, which PowerPoint refuses, so the whole sync dies on a date.
'
' Returns -1 for input it cannot honour, rather than a plausible number. A
' zero-length or backwards project is a data error and must be reported, not
' rendered as 0% -- 0% is a real, meaningful answer and would hide the fault.
Public Function ElapsedFraction(startDate As Date, endDate As Date, asAt As Date) As Double
    If endDate < startDate Then
        ElapsedFraction = -1
        Exit Function
    End If

    ' Same start and end: degenerate but not an error -- a milestone rather than
    ' a span. Anything at or after it is complete; before it, not started.
    If endDate = startDate Then
        If asAt < startDate Then
            ElapsedFraction = 0
        Else
            ElapsedFraction = 1
        End If
        Exit Function
    End If

    Dim total As Double, done As Double
    total = CDbl(endDate) - CDbl(startDate)
    done = CDbl(asAt) - CDbl(startDate)

    If done <= 0 Then
        ElapsedFraction = 0
    ElseIf done >= total Then
        ElapsedFraction = 1
    Else
        ElapsedFraction = done / total
    End If
End Function

' The two widths, given the invariant total. Returned together so they cannot
' drift apart: computing "remaining" as its own fraction would let rounding put
' the pair a hair over or under the total, and the join would show.
Public Sub BarWidths(totalWidth As Double, fraction As Double, _
                     ByRef elapsedWidth As Double, ByRef remainingWidth As Double)
    If fraction < 0 Then fraction = 0
    If fraction > 1 Then fraction = 1
    elapsedWidth = totalWidth * fraction
    remainingWidth = totalWidth - elapsedWidth      ' SUBTRACTED, never recomputed
End Sub

' The as-at date a period means. "FY26Q4" -> 30 June 2026, Australian financial
' year: FY26 ends 30 June 2026, so Q1 = Jul-Sep 2025 ... Q4 = Apr-Jun 2026.
'
' Returns 0 when it cannot parse, and callers must treat 0 as "no date" rather
' than as 30 December 1899. A period this does not understand is a reason to
' refuse, not to guess a date and draw a confident bar from it.
Public Function PeriodEndDate(period As String) As Date
    Dim p As String
    p = UCase(Trim(period))
    If Len(p) <> 6 Then Exit Function
    If Left(p, 2) <> "FY" Then Exit Function
    If Mid(p, 5, 1) <> "Q" Then Exit Function

    Dim yy As String, q As String
    yy = Mid(p, 3, 2)
    q = Right(p, 1)
    If Not IsNumeric(yy) Or Not IsNumeric(q) Then Exit Function

    Dim fyEndYear As Long
    fyEndYear = 2000 + CLng(yy)

    Select Case CLng(q)
        Case 1: PeriodEndDate = DateSerial(fyEndYear - 1, 9, 30)
        Case 2: PeriodEndDate = DateSerial(fyEndYear - 1, 12, 31)
        Case 3: PeriodEndDate = DateSerial(fyEndYear, 3, 31)
        Case 4: PeriodEndDate = DateSerial(fyEndYear, 6, 30)
    End Select
End Function

' Self-test. Counts PASS lines rather than grepping for FAIL, because a suite
' that runs zero tests passes a grep-for-FAIL check -- which this project has
' been caught by.
Public Function SelfTest() As String
    Dim r As String
    Dim passed As Long, failed As Long

    r = r & Chk("start of span", ElapsedFraction(#1/1/2026#, #12/31/2026#, #1/1/2026#), 0, passed, failed)
    r = r & Chk("end of span", ElapsedFraction(#1/1/2026#, #12/31/2026#, #12/31/2026#), 1, passed, failed)
    r = r & Chk("before start clamps to 0", ElapsedFraction(#1/1/2026#, #12/31/2026#, #6/1/2025#), 0, passed, failed)
    r = r & Chk("after end clamps to 1", ElapsedFraction(#1/1/2026#, #12/31/2026#, #6/1/2030#), 1, passed, failed)
    r = r & Chk("halfway", ElapsedFraction(#1/1/2026#, #1/11/2026#, #1/6/2026#), 0.5, passed, failed)
    r = r & Chk("backwards dates refuse", ElapsedFraction(#12/31/2026#, #1/1/2026#, #6/1/2026#), -1, passed, failed)
    r = r & Chk("same day, after", ElapsedFraction(#1/1/2026#, #1/1/2026#, #1/2/2026#), 1, passed, failed)
    r = r & Chk("same day, before", ElapsedFraction(#1/1/2026#, #1/1/2026#, #12/1/2025#), 0, passed, failed)

    ' The invariant, stated as a test rather than trusted: the parts must sum to
    ' the total at an awkward fraction, not just at 0, 0.5 and 1.
    Dim e As Double, rem_ As Double
    BarWidths 317.5, ElapsedFraction(#1/1/2026#, #12/31/2026#, #4/17/2026#), e, rem_
    r = r & Chk("parts sum to total", e + rem_, 317.5, passed, failed)

    r = r & Chk("FY26Q4 ends 30 Jun 2026", CDbl(PeriodEndDate("FY26Q4")), CDbl(#6/30/2026#), passed, failed)
    r = r & Chk("FY26Q1 ends 30 Sep 2025", CDbl(PeriodEndDate("FY26Q1")), CDbl(#9/30/2025#), passed, failed)
    r = r & Chk("FY27Q1 ends 30 Sep 2026", CDbl(PeriodEndDate("FY27Q1")), CDbl(#9/30/2026#), passed, failed)
    r = r & Chk("unparseable period is 0", CDbl(PeriodEndDate("WRAP-5")), 0, passed, failed)
    r = r & Chk("empty period is 0", CDbl(PeriodEndDate("")), 0, passed, failed)

    SelfTest = r & vbCrLf & "Timeline: " & passed & " passed, " & failed & " failed." & vbCrLf
End Function

Private Function Chk(label As String, got As Double, want As Double, _
                     ByRef passed As Long, ByRef failed As Long) As String
    ' Tolerance, because these are floating point. Exact equality on a division
    ' would fail on arithmetic that is correct.
    If Abs(got - want) < 0.000001 Then
        passed = passed + 1
        Chk = "  PASS  " & label & vbCrLf
    Else
        failed = failed + 1
        Chk = "  FAIL  " & label & "  -- got " & got & ", wanted " & want & vbCrLf
    End If
End Function
