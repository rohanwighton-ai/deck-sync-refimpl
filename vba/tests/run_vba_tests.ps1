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

# A prior interrupted run can leave a zombie POWERPNT.EXE/EXCEL.EXE process
# behind; New-Object -ComObject then attaches to that stale instance
# instead of spawning a fresh one, and its VBProject/VBComponents access
# can come back genuinely null even though Trust access is enabled (found
# the hard way on this script's first real run -- not a hypothetical).
# Always start from a clean process table.
Get-Process POWERPNT, EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 1

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

$pptModules = @("Discovery.bas", "InjectPrimitive.bas", "Matching.bas", "Resolve.bas", "SyncOperations.bas", "Onboarding.bas", "Verification.bas", "SlideDuplication.bas", "RunSync.bas", "tests\TestRunner.bas")
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

    foreach ($m in @("Discovery.bas", "InjectPrimitive.bas", "Matching.bas", "Resolve.bas", "SyncOperations.bas", "Onboarding.bas", "ExcelOutput.bas", "Verification.bas", "SlideDuplication.bas", "RunSync.bas", "TestRunner.bas")) {
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
