# First real-execution driver for deck-sync-refimpl's VBA port.
#
# Every module up to this point was "not executed or verified in this
# environment" per its own SPIKE_NOTES_*.md -- there was no Windows/Office
# install available. This script closes that gap: it drives real PowerPoint
# and Excel instances via COM automation, imports the .bas modules
# (production + the assertion-based TestRunner*.bas), runs the tests, and
# prints one machine-readable report per host app to stdout.
#
# Deliberately NOT interactive: Office is launched, driven, and quit with
# no human input required, so this can be re-run repeatedly as part of a
# hardening loop (run -> read real failures -> fix vba/*.bas -> re-run).
#
# Everything is staged into a Windows-native temp folder first (not run
# directly against the WSL \\wsl.localhost\... path) -- COM automation
# across that boundary is a known source of flakiness; plain file copies
# across it are not, so staging is the only thing that crosses it.

param(
    [string]$RepoRoot = $(Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
)

$ErrorActionPreference = "Stop"

# A prior interrupted run (or you just having Office open) can leave a live
# POWERPNT.EXE/EXCEL.EXE process behind. New-Object -ComObject would attach
# to that instance instead of spawning a fresh one, and a stale instance's
# VBProject/VBComponents access can come back genuinely null even with Trust
# access enabled (found the hard way on this script's first real run -- not
# a hypothetical). So any existing instance must be closed first.
#
# This script is now also invoked unattended after autonomous commits, so
# force-killing (the original approach) is no longer safe -- it would
# silently discard unsaved work in a document you happen to have open at the
# time. Instead: ask each running instance to Quit() normally. If it has
# unsaved changes, Office shows its own native "Save changes?" dialog and
# Quit() blocks until you respond. That's bounded by a timeout so an
# unattended run doesn't hang forever -- if you don't respond in time, this
# run is skipped (exit 2) rather than forcing the issue.
#
# --- The husk problem, and why the discriminator changed (2026-07-30) -------
# This suite LEAKS an Excel process on most runs: several tests do
# CreateObject("Excel.Application") / wb.Close False / xl.Quit, and an
# outstanding COM reference stops the process actually terminating. The
# leftover is an empty Excel holding no workbook -- and it survives a clean
# Quit(), so the $stillRunning check below rejected it and every SUBSEQUENT
# run exited 2 with no output until a human killed it by hand. That cost two
# 8-minute runs in one evening, and the note about it had already been
# written down before it happened again.
#
# The obvious fix -- "force-kill windowless instances" -- does not work. The
# same leftover process reported MainWindowTitle "Excel" at 19:56 and blank
# at 20:12, zero workbooks throughout: window presence flips on its own and
# is not evidence of anything.
#
# What does not flip is whether the instance holds a DOCUMENT. Zero documents
# means there is nothing a human can lose, which is the only question that
# actually matters here. So: count documents first, and force-kill only a
# confirmed-empty instance.
#
# Two asymmetries make that safe, and both are load-bearing:
#
#   1. UNKNOWN NEVER PERMITS A KILL. Attaching or counting can fail or hang
#      (a modal dialog blocks COM), and every such case returns -1, which
#      falls through to the graceful path. Only a confirmed 0 kills.
#   2. MORE THAN ONE PROCESS NEVER PERMITS A KILL. GetActiveObject attaches
#      to whichever instance registered in the Running Object Table -- with
#      two EXCEL.EXE running it may well be the husk, so "0 documents" would
#      say nothing about the other one. Killing by name on that basis is
#      precisely how you would discard the open workbook this comment block
#      exists to protect. So the kill requires exactly one process of that
#      name, and targets it by PID, not by name.
#
# Net effect: the leaked husk gets cleared automatically, a real session of
# yours is still asked politely and still aborts the run if it will not go,
# and anything ambiguous is treated as a real session.

# Documents open in the running instance of $ProgId.
#   0  = confirmed empty -- an automation husk, nothing to lose
#   >0 = a real session with documents in it
#   -1 = unknown (no instance, attach failed, or the call hung) -> treat as real
#
# Runs in a job for the same reason Request-GracefulQuit does: a modal dialog
# in Office blocks COM calls indefinitely, and a pre-flight check that can
# hang forever is worse than the problem it solves.
function Get-OpenDocumentCount {
    param([string]$ProgId, [int]$TimeoutSeconds = 30)

    $job = Start-Job -ScriptBlock {
        param($ProgId)
        try {
            $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject($ProgId)
        }
        catch {
            return -1
        }
        try {
            if ($ProgId -eq "Excel.Application") { return [int]$app.Workbooks.Count }
            else { return [int]$app.Presentations.Count }
        }
        catch {
            return -1
        }
    } -ArgumentList $ProgId

    $completed = Wait-Job $job -Timeout $TimeoutSeconds
    if (-not $completed) {
        Stop-Job $job -ErrorAction SilentlyContinue
        Remove-Job $job -Force -ErrorAction SilentlyContinue
        return -1
    }
    $result = Receive-Job $job
    Remove-Job $job -ErrorAction SilentlyContinue
    if ($null -eq $result) { return -1 }
    return [int]$result
}

# Clears a confirmed-empty leftover.
#
# Returns NOTHING, deliberately. The first version returned $true/$false so the
# caller could tell "handled" from "not applicable" -- but in PowerShell a
# Write-Output inside a function is emitted onto the same pipeline as the
# return value, so the caller's `| Out-Null` (added to discard the boolean)
# also swallowed the "Clearing a leftover..." message. The script would then
# force-kill a process and leave no trace of it in the output, which is the
# single behaviour an unattended script must never have. Caught by running the
# function for real rather than reading it, 2026-07-30.
#
# Since no caller ever used the boolean, dropping it removes the conflict
# outright instead of working around it. Bare `return` exits without emitting.
function Clear-EmptyHusk {
    param([string]$ProgId, [string]$ProcessName)

    $procs = @(Get-Process $ProcessName -ErrorAction SilentlyContinue)
    if ($procs.Count -ne 1) { return }   # asymmetry 2: ambiguity is not a kill

    $docs = Get-OpenDocumentCount -ProgId $ProgId
    if ($docs -ne 0) { return }          # asymmetry 1: unknown (-1) is not a kill

    Write-Output "=== Clearing a leftover empty $ProcessName (pid $($procs[0].Id), 0 documents open) -- an automation husk from a previous run. ==="
    Stop-Process -Id $procs[0].Id -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
}

function Request-GracefulQuit {
    param([string]$ProgId, [int]$TimeoutSeconds = 180)

    $job = Start-Job -ScriptBlock {
        param($ProgId)
        try {
            $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject($ProgId)
        }
        catch {
            return "no-instance"
        }
        try {
            $app.Quit()
            return "quit-ok"
        }
        catch {
            return "quit-error: $($_.Exception.Message)"
        }
    } -ArgumentList $ProgId

    $completed = Wait-Job $job -Timeout $TimeoutSeconds
    if (-not $completed) {
        Stop-Job $job -ErrorAction SilentlyContinue
        Remove-Job $job -Force -ErrorAction SilentlyContinue
        return "timeout"
    }
    $result = Receive-Job $job
    Remove-Job $job -ErrorAction SilentlyContinue
    return $result
}

$processNameFor = @{ "PowerPoint.Application" = "POWERPNT"; "Excel.Application" = "EXCEL" }
foreach ($progId in @("PowerPoint.Application", "Excel.Application")) {
    # Husk first: an empty leftover cannot be cleared by Quit() (that is the
    # whole failure), so asking politely and then rejecting it is a dead end.
    # If this clears something, the checks below simply find nothing running.
    # No `| Out-Null` here on purpose -- see Clear-EmptyHusk's header. Piping
    # its output away is what would silence the record of a force-kill.
    Clear-EmptyHusk -ProgId $progId -ProcessName $processNameFor[$progId]

    $outcome = Request-GracefulQuit -ProgId $progId
    if ($outcome -eq "timeout" -or $outcome -like "quit-error*") {
        Write-Output "=== SKIPPED: $progId did not close cleanly ($outcome) -- likely an unresolved save prompt or in active use. Not forcing it closed. ==="
        exit 2
    }
    # Give Office a moment to actually release the process after Quit()
    Start-Sleep -Seconds 1
    $stillRunning = Get-Process $processNameFor[$progId] -ErrorAction SilentlyContinue
    if ($stillRunning) {
        Write-Output "=== SKIPPED: $($processNameFor[$progId]) still running after a clean Quit() -- not forcing it closed. ==="
        exit 2
    }
}

$staging = Join-Path $env:TEMP ("deck-sync-vba-test-" + [guid]::NewGuid().ToString("N").Substring(0, 8))
$fixturesStaging = Join-Path $staging "fixtures\"
New-Item -ItemType Directory -Path $staging | Out-Null
New-Item -ItemType Directory -Path $fixturesStaging | Out-Null
$staging = (Get-Item $staging).FullName + "\"

Write-Output "Staging directory: $staging"

$vbaSourceDir = Join-Path $RepoRoot "vba"
$fixturesSourceDir = Join-Path $RepoRoot "test-fixtures"
Write-Output "RepoRoot: $RepoRoot"
Write-Output "vbaSourceDir: $vbaSourceDir"
Write-Output "fixturesSourceDir: $fixturesSourceDir"

$pptModules = @("Discovery.bas", "InjectPrimitive.bas", "Matching.bas", "Resolve.bas", "SyncOperations.bas", "Onboarding.bas", "Verification.bas", "SlideDuplication.bas", "TemplateSlide.bas", "TemplateAudit.bas", "RunSync.bas", "DeckAdoption.bas", "ResolveFields.bas", "DeckRegistry.bas", "WorkbookBridge.bas", "OnboardFlow.bas", "RibbonUI.bas", "AdoptFlow.bas", "BatchOnboardFlow.bas", "CommandBarUI.bas", "tests\TestRunner.bas")
foreach ($m in $pptModules) {
    Copy-Item (Join-Path $vbaSourceDir $m) -Destination $staging
}
$excelModules = @("ExcelOutput.bas", "tests\TestRunnerExcel.bas")
foreach ($m in $excelModules) {
    Copy-Item (Join-Path $vbaSourceDir $m) -Destination $staging
}
Copy-Item (Join-Path $fixturesSourceDir "shp-groupshape.pptx") -Destination $fixturesStaging

$pptReport = ""
$excelReport = ""
$pptError = $null
$excelError = $null

# --- PowerPoint pass -------------------------------------------------
$ppt = $null
$pres = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $ppt.Visible = -1  # msoTrue -- visible on purpose for a first real run, see script header
    $pres = $ppt.Presentations.Add()

    foreach ($m in @("Discovery.bas", "InjectPrimitive.bas", "Matching.bas", "Resolve.bas", "SyncOperations.bas", "Onboarding.bas", "ExcelOutput.bas", "Verification.bas", "SlideDuplication.bas", "TemplateSlide.bas", "TemplateAudit.bas", "RunSync.bas", "DeckAdoption.bas", "ResolveFields.bas", "DeckRegistry.bas", "WorkbookBridge.bas", "OnboardFlow.bas", "RibbonUI.bas", "AdoptFlow.bas", "BatchOnboardFlow.bas", "CommandBarUI.bas", "TestRunner.bas")) {
        $comp = $pres.VBProject.VBComponents.Import((Join-Path $staging $m))
        Write-Output ("Imported $m as component name: " + $comp.Name)
    }

    # $ppt.Run("...", $a, $b) directly fails PowerShell's COM overload
    # resolution for Application.Run's VARIANT-optional-parameter signature
    # ("Cannot find an overload for 'Run' and the argument count: 3") --
    # a well-known PowerShell/COM interop gap, not a VBA-side issue.
    # InvokeMember via reflection is the standard workaround: it bypasses
    # PowerShell's own (buggy, for this case) method-binding logic.
    $fixturesArg = [string]$fixturesStaging
    $stagingArg = [string]$staging
    $pptReport = $ppt.GetType().InvokeMember("Run", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ppt, @([string]"TestRunner.RunAllTests", $fixturesArg, $stagingArg))
}
catch {
    $pptError = $_.Exception.Message + " [line " + $_.InvocationInfo.ScriptLineNumber + ": " + $_.InvocationInfo.Line.Trim() + "]"
}
finally {
    if ($pres) { $pres.Saved = $true; $pres.Close() }
    if ($ppt) { $ppt.Quit() }
    # Same COM-release gotcha as the Excel pass below (see that finally
    # block's comment) -- Quit() alone is not reliable here either,
    # confirmed 2026-07-26 (a windowless POWERPNT.EXE survived a clean,
    # non-erroring run until this was added).
    if ($pres) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($pres) | Out-Null }
    if ($ppt) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null }
    $pres = $null; $ppt = $null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    [System.GC]::Collect()
    Start-Sleep -Milliseconds 500
    Get-Process POWERPNT -ErrorAction SilentlyContinue | Where-Object { -not $_.MainWindowTitle } | Stop-Process -Force -ErrorAction SilentlyContinue
}

# --- Excel pass --------------------------------------------------------
$xl = $null
$wb = $null
try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $true
    $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Add()

    foreach ($m in @("ExcelOutput.bas", "TestRunnerExcel.bas")) {
        $wb.VBProject.VBComponents.Import((Join-Path $staging $m)) | Out-Null
    }

    $excelReport = $xl.Run("TestRunnerExcel.RunAllTests")
}
catch {
    $excelError = $_.Exception.Message + " [line " + $_.InvocationInfo.ScriptLineNumber + ": " + $_.InvocationInfo.Line.Trim() + "]"
}
finally {
    if ($wb) { $wb.Saved = $true; $wb.Close() }
    if ($xl) { $xl.Quit() }
    # $xl.Quit() alone reliably leaves the EXCEL.EXE process running
    # (confirmed 2026-07-26, unrelated to the AutoSave/BatchOnboardFlow
    # investigation that day: several sequential no-window EXCEL.EXE
    # processes were still alive well after a completed, non-erroring test
    # run and had to be force-killed by hand) -- releasing the RCW and
    # forcing a GC pass, same idiom used elsewhere in this script's own
    # PowerPoint-pass cleanup probes, actually lets the process exit.
    if ($wb) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null }
    if ($xl) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null }
    $wb = $null; $xl = $null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    [System.GC]::Collect()  # a single pass reliably left some RCWs alive in testing -- a second pass after WaitForPendingFinalizers cleared the rest
    Start-Sleep -Milliseconds 500
    # Belt-and-braces: this script's own opening section already treats a
    # leftover EXCEL.EXE as something to self-heal from, not tolerate --
    # if COM release still didn't let this run's own process exit, force
    # it rather than leaving it for the *next* run to clean up. Restricted
    # to windowless processes specifically (every zombie observed while
    # fixing this had a blank MainWindowTitle) so a real, separate,
    # visible Excel session a human opened during the run is never touched.
    Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { -not $_.MainWindowTitle } | Stop-Process -Force -ErrorAction SilentlyContinue
}

# --- Report --------------------------------------------------------
if ($pptError) {
    Write-Output "=== PowerPoint pass: DRIVER ERROR ==="
    Write-Output $pptError
}
else {
    Write-Output $pptReport
}

if ($excelError) {
    Write-Output "=== Excel pass: DRIVER ERROR ==="
    Write-Output $excelError
}
else {
    Write-Output $excelReport
}

if (-not $env:DECK_SYNC_KEEP_STAGING) {
    Remove-Item -Path $staging -Recurse -Force -ErrorAction SilentlyContinue
}

# --- Did anything actually RUN? ------------------------------------
#
# A VBA compile error anywhere in the project makes Application.Run fail without
# producing a single test result -- and until 2026-07-28 this script reported that
# as a clean run: zero PASS lines, zero FAIL lines, exit 0. A grep for "^FAIL"
# found nothing and the run looked green. The compile error that caused it
# (Application.PathSeparator, an Excel-only member used in a PowerPoint-hosted
# module) therefore reached Rohan's screen instead of this script's output.
#
# Silence is not success. An empty result set is the loudest possible failure --
# it means the project did not even compile -- so it must exit non-zero.
$allOutput = "$pptReport`n$excelReport"
$passCount = ([regex]::Matches($allOutput, '(?m)^PASS')).Count
$failCount = ([regex]::Matches($allOutput, '(?m)^FAIL')).Count

Write-Output ""
Write-Output "=== $passCount passed, $failCount failed ==="

if ($pptError -or $excelError) {
    Write-Output "=== DRIVER ERROR (see above) ==="
    exit 3
}
if ($passCount -eq 0) {
    Write-Output "=== NO TESTS RAN. This is a FAILURE, not a pass. ==="
    Write-Output "Almost always a VBA compile error somewhere in the project:"
    Write-Output "  Application.Run reports 'Sub or function not defined' for a"
    Write-Output "  compile error anywhere, not only for a missing macro."
    Write-Output "  Open the project and use Debug > Compile VBAProject -- its"
    Write-Output "  dialog names the offending line directly."
    exit 2
}
if ($failCount -gt 0) { exit 1 }
exit 0
