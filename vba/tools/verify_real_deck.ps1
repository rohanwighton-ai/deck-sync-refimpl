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

$modules = @(
    (Join-Path $vbaSourceDir "Resolve.bas"),
    (Join-Path $vbaSourceDir "ExcelOutput.bas"),
    (Join-Path $vbaSourceDir "tools\VerifyRealDeck.bas")
)
foreach ($m in $modules) { Copy-Item $m -Destination $staging }

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = -1
$pres = $null
try {
    $pres = $ppt.Presentations.Add()
    foreach ($m in @("Resolve.bas", "ExcelOutput.bas", "VerifyRealDeck.bas")) {
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
