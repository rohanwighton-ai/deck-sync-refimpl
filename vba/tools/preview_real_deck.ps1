# Read-only driver for PreviewRealDeck.bas: reports exactly what a real
# Sync Now would do to the deck, and does none of it. Opens the deck and its
# paired Data workbook READ-ONLY and closes both without saving.
#
# Report is written in full to $OutFile (local only -- contains real business
# field values and instance keys, deliberately kept out of chat) with a
# summary printed to stdout.
#
# Run this from a Windows-local path, not \\wsl.localhost: PowerShell treats
# the UNC path as remote and refuses to load the script -- while still exiting
# 0. See AGENTS.md's Known Patterns.

param(
    [string]$DeckPath = "C:\Users\rohan\OneDrive\Claude\test1.pptx",
    [string]$RepoRoot = $(Split-Path -Parent (Split-Path -Parent $PSScriptRoot)),
    [string]$OutFile = (Join-Path $env:TEMP "preview_real_deck_report.txt")
)

$ErrorActionPreference = "Stop"

foreach ($progId in @("PowerPoint.Application", "Excel.Application")) {
    try {
        $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject($progId)
        Write-Output "=== ABORTING: $progId is already running -- close it first so this diagnostic doesn't attach to or interfere with a live session. ==="
        exit 2
    } catch {}
}

$staging = Join-Path $env:TEMP ("deck-sync-preview-" + [guid]::NewGuid().ToString("N").Substring(0, 8))
New-Item -ItemType Directory -Path $staging | Out-Null
$vbaSourceDir = Join-Path $RepoRoot "vba"

# PreviewRoutineSync pulls in most of the add-in: RunSync -> SyncOperations /
# ExcelOutput / SlideDuplication / InjectPrimitive / Resolve, and its report
# formatting currently reaches into BatchOnboardFlow.FieldPreview, which drags
# in that module's own dependency tree (Discovery / Matching / Onboarding /
# WorkbookBridge / RibbonUI). Importing the same set run_vba_tests.ps1 imports
# is simpler and known to compile together.
$modules = @(
    "Discovery.bas", "InjectPrimitive.bas", "Matching.bas", "Resolve.bas",
    "SyncOperations.bas", "Onboarding.bas", "ExcelOutput.bas", "Verification.bas",
    "SlideDuplication.bas", "RunSync.bas", "DeckRegistry.bas", "WorkbookBridge.bas",
    "OnboardFlow.bas", "RibbonUI.bas", "BatchOnboardFlow.bas"
)
foreach ($m in $modules) { Copy-Item (Join-Path $vbaSourceDir $m) -Destination $staging }
Copy-Item (Join-Path $vbaSourceDir "tools\PreviewRealDeck.bas") -Destination $staging

$ppt = $null
$scratch = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $ppt.Visible = -1
    $scratch = $ppt.Presentations.Add()

    foreach ($m in ($modules + @("PreviewRealDeck.bas"))) {
        $leaf = Split-Path $m -Leaf
        $null = $scratch.VBProject.VBComponents.Import((Join-Path $staging $leaf))
    }
    Write-Output ("Imported " + ($modules.Count + 1) + " modules.")

    # InvokeMember rather than $ppt.Run(...): PowerShell cannot bind
    # Application.Run's VARIANT-optional-parameter signature directly.
    $report = $ppt.GetType().InvokeMember("Run", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ppt, @([string]"PreviewRealDeck.PreviewRealDeck", [string]$DeckPath))

    Set-Content -Path $OutFile -Value $report -Encoding UTF8
    Write-Output "=== Full report written to: $OutFile ==="
    Write-Output ""

    # Summary only to stdout -- the per-field before/after lines carry real
    # business content.
    $report -split "`r?`n" | Where-Object {
        $_ -match "^(Deck:|Workbook:|Slides:|Types:|=== |Summary:|WARNING:|A real Sync Now|If that is not)" -or
        $_ -match "slide\(s\) are not in Data-sheet row order" -or
        $_ -match "WOULD CREATE A NEW SLIDE"
    } | ForEach-Object { Write-Output $_ }
}
finally {
    if ($scratch -ne $null) { try { $scratch.Saved = $true; $scratch.Close() } catch {} }
    if ($ppt -ne $null) { try { $ppt.Quit() } catch {} }
    Remove-Item -Recurse -Force $staging -ErrorAction SilentlyContinue
}
