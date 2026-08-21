# Read-only diagnostic driver for VerifyRealDeck.bas. Opens the real deck
# and its paired Data workbook READ-ONLY, cross-checks every tagged
# slide's shapes against the workbook's harvested values, and closes both
# without saving. Never writes to either file.
#
# Report is written in full to $OutFile (local only -- may contain real
# business field names/instance keys, kept out of chat) plus a summary
# printed to stdout.

param(
    [string]$DeckPath = "C:\Users\rohan\OneDrive\Claude\test1.pptx",
    [string]$WorkbookPath = "C:\Users\rohan\OneDrive\Claude\SAAFE-Projects-Data.xlsx",
    [string]$RepoRoot = $(Split-Path -Parent (Split-Path -Parent $PSScriptRoot)),
    [string]$OutFile = (Join-Path $env:TEMP "verify_real_deck_report.txt")
)

$ErrorActionPreference = "Stop"

foreach ($progId in @("PowerPoint.Application", "Excel.Application")) {
    try {
        $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject($progId)
        Write-Output "=== ABORTING: $progId is already running -- close it first so this diagnostic doesn't attach to or interfere with a live session. ==="
        exit 2
    } catch {}
}

$staging = Join-Path $env:TEMP ("deck-sync-verify-" + [guid]::NewGuid().ToString("N").Substring(0, 8))
New-Item -ItemType Directory -Path $staging | Out-Null
$vbaSourceDir = Join-Path $RepoRoot "vba"

# CORRECTED 2026-08-21: this used to import a hand-picked 3-module subset
# (Resolve/ExcelOutput/VerifyRealDeck) that never covered what VerifyRealDeck
# actually calls (DeckRegistry, InjectPrimitive, WorkbookBridge, and their
# own transitive dependencies) -- a mother-hound audit found this script had
# been unable to compile for three weeks while still exiting 0. Same failure
# shape already fixed the same night in preview_real_deck.ps1/
# sync_real_deck.ps1: stop hand-maintaining a second, silently-driftable
# copy of the dependency graph, use the canonical list build_ppam.ps1 uses
# for the real shipped add-in instead.
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
        # CRLF, NOT A PLAIN COPY -- see run_vba_tests.ps1's own comment on this
        # exact line. LF-only makes Import() treat the class header as
        # unrecognised and silently import it as a plain Standard Module.
        ((Get-Content $srcPath -Raw) -replace "`r`n", "`n" -replace "`n", "`r`n") | Set-Content $dstPath -NoNewline
    } else {
        Copy-Item $srcPath -Destination $dstPath
    }
}
Copy-Item (Join-Path $vbaSourceDir "tools\VerifyRealDeck.bas") -Destination $staging

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = -1
$pres = $null
$driverFailed = $false
try {
    $pres = $ppt.Presentations.Add()
    foreach ($m in ($moduleNames + @("VerifyRealDeck.bas"))) {
        $comp = $pres.VBProject.VBComponents.Import((Join-Path $staging $m))
        Write-Output ("Imported $m as: " + $comp.Name)
    }

    $report = $ppt.GetType().InvokeMember("Run", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ppt, @([string]"VerifyRealDeck.VerifyRealDeck", [string]$DeckPath, [string]$WorkbookPath))

    Set-Content -Path $OutFile -Value $report -Encoding UTF8

    $summaryEnd = $report.IndexOf("--- Per-slide detail")
    if ($summaryEnd -gt 0) {
        Write-Output $report.Substring(0, $summaryEnd)
    } else {
        Write-Output $report
    }
    Write-Output "Full report (incl. per-slide/per-field detail -- kept local, not printed) written to: $OutFile"
}
catch {
    Write-Output ("=== DRIVER ERROR === " + $_.Exception.Message + " [line " + $_.InvocationInfo.ScriptLineNumber + ": " + $_.InvocationInfo.Line.Trim() + "]")
    # CORRECTED 2026-08-21: this used to fall through to a normal (0) exit
    # even on a genuine failure -- "unable to compile for three weeks and
    # exits 0 anyway" (mother-hound). A DRIVER ERROR line in the output is
    # not the same as a caller checking $LASTEXITCODE.
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
