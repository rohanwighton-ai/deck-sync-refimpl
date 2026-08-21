# Read-only diagnostic driver for MilestoneEvidenceReport.bas. Opens the
# real deck and its paired Data workbook READ-ONLY, cross-checks every
# registered project's MS<n>_DONE flags against SRC_MILESTONES's own
# tracker evidence, and closes both without saving. Never writes to
# either file, never decides anything -- report only, human applies.
#
# Same pattern as verify_real_deck.ps1 (that script's own header explains
# the module-list history this one reuses verbatim).
#
# Report is written in full to $OutFile (local only -- may contain real
# business field names/instance keys, kept out of chat) plus a summary
# printed to stdout.

param(
    [string]$DeckPath = "C:\Users\rohan\OneDrive\Claude\test1.pptx",
    [string]$WorkbookPath = "C:\Users\rohan\OneDrive\Claude\SAAFE-Projects-Data.xlsx",
    [string]$RepoRoot = $(Split-Path -Parent (Split-Path -Parent $PSScriptRoot)),
    [string]$OutFile = (Join-Path $env:TEMP "milestone_evidence_report.txt")
)

$ErrorActionPreference = "Stop"

foreach ($progId in @("PowerPoint.Application", "Excel.Application")) {
    try {
        $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject($progId)
        Write-Output "=== ABORTING: $progId is already running -- close it first so this diagnostic doesn't attach to or interfere with a live session. ==="
        exit 2
    } catch {}
}

$staging = Join-Path $env:TEMP ("deck-sync-milestone-evidence-" + [guid]::NewGuid().ToString("N").Substring(0, 8))
New-Item -ItemType Directory -Path $staging | Out-Null
$vbaSourceDir = Join-Path $RepoRoot "vba"

# Same canonical module list build_ppam.ps1/verify_real_deck.ps1 use --
# see verify_real_deck.ps1's own header for why this is not hand-maintained
# as a separate subset.
$moduleNames = @(
    "Discovery.bas", "InjectPrimitive.bas", "Matching.bas", "Resolve.bas",
    "SyncOperations.bas", "Onboarding.bas", "ExcelOutput.bas", "Verification.bas",
    "SlideDuplication.bas", "TemplateSlide.bas", "TemplateAudit.bas", "IdentityCheck.bas", "TagMigration.bas", "PlaceholderCheck.bas", "RunSync.bas", "DeckAdoption.bas", "ResolveFields.bas",
    "DeckRegistry.bas", "WorkbookBridge.bas", "Harvest.bas", "RibbonUI.bas",
    "AdoptFlow.bas", "BatchOnboardFlow.bas", "CommandBarUI.bas",
    "ReviewQueue.bas", "FieldWiring.bas", "MilestoneDevice.bas",
    "Drafting.bas", "FieldSpec.bas", "Sources.bas", "DraftingUI.bas", "DiscoverUI.bas",
    "DraftingLobby.bas", "AppEvents.cls", "ShapeAddressBook.bas", "Timing.bas", "FormattingAudit.bas"
)
foreach ($m in $moduleNames) {
    $srcPath = Join-Path $vbaSourceDir $m
    $dstPath = Join-Path $staging $m
    if ($m -like "*.cls") {
        ((Get-Content $srcPath -Raw) -replace "`r`n", "`n" -replace "`n", "`r`n") | Set-Content $dstPath -NoNewline
    } else {
        Copy-Item $srcPath -Destination $dstPath
    }
}
Copy-Item (Join-Path $vbaSourceDir "tools\MilestoneEvidenceReport.bas") -Destination $staging

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = -1
$pres = $null
$driverFailed = $false
try {
    $pres = $ppt.Presentations.Add()
    foreach ($m in ($moduleNames + @("MilestoneEvidenceReport.bas"))) {
        $comp = $pres.VBProject.VBComponents.Import((Join-Path $staging $m))
        Write-Output ("Imported $m as: " + $comp.Name)
    }

    $report = $ppt.GetType().InvokeMember("Run", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ppt, @([string]"MilestoneEvidenceReport.MilestoneEvidenceReport", [string]$DeckPath, [string]$WorkbookPath))

    Set-Content -Path $OutFile -Value $report -Encoding UTF8

    $summaryEnd = $report.IndexOf("--- Per-project detail")
    if ($summaryEnd -gt 0) {
        Write-Output $report.Substring(0, $summaryEnd)
    } else {
        Write-Output $report
    }
    Write-Output "Full report (incl. per-project detail -- kept local, not printed) written to: $OutFile"
}
catch {
    Write-Output ("=== DRIVER ERROR === " + $_.Exception.Message + " [line " + $_.InvocationInfo.ScriptLineNumber + ": " + $_.InvocationInfo.Line.Trim() + "]")
    $driverFailed = $true
}
finally {
    if ($pres) { $pres.Saved = $true; try { $pres.Close() } catch {} }
    if ($ppt) { try { $ppt.Quit() } catch {} }
    if ($pres) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($pres) | Out-Null }
    if ($ppt) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null }
    $pres = $null; $ppt = $null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    [System.GC]::Collect()
    Start-Sleep -Milliseconds 500
    Get-Process POWERPNT -ErrorAction SilentlyContinue | Where-Object { -not $_.MainWindowTitle } | Stop-Process -Force -ErrorAction SilentlyContinue
    Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { -not $_.MainWindowTitle } | Stop-Process -Force -ErrorAction SilentlyContinue
}

if ($driverFailed) { exit 1 }
