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

$pptModules = @("Discovery.bas", "InjectPrimitive.bas", "Matching.bas", "Resolve.bas", "SyncOperations.bas", "Onboarding.bas", "Verification.bas", "SlideDuplication.bas", "RunSync.bas", "DeckAdoption.bas", "ResolveFields.bas", "DeckRegistry.bas", "WorkbookBridge.bas", "OnboardFlow.bas", "RibbonUI.bas", "AdoptFlow.bas", "BatchOnboardFlow.bas", "CommandBarUI.bas", "tests\TestRunner.bas")
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

    foreach ($m in @("Discovery.bas", "InjectPrimitive.bas", "Matching.bas", "Resolve.bas", "SyncOperations.bas", "Onboarding.bas", "ExcelOutput.bas", "Verification.bas", "SlideDuplication.bas", "RunSync.bas", "DeckAdoption.bas", "ResolveFields.bas", "DeckRegistry.bas", "WorkbookBridge.bas", "OnboardFlow.bas", "RibbonUI.bas", "AdoptFlow.bas", "BatchOnboardFlow.bas", "CommandBarUI.bas", "TestRunner.bas")) {
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
