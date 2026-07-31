# Driver for R13RealDeck.bas -- the R13 review/apply flow against the deck COPY.
#
# Two phases, selected by -Phase:
#   review  (default)  opens the deck READ-ONLY, builds the queue, writes the
#                      review grid to its own workbook. Cannot change a slide.
#   apply              opens the deck read-write and applies what the grid ticks.
#
# Defaults point at C:\Users\rohan\deck-sync-e2e\ -- the scratch rig. The real
# deck at OneDrive\Claude\test1.pptx is never a default here and should never be
# passed; work on a copy.
#
# Report goes to $OutFile (local only -- it contains real business field values
# and instance keys, deliberately kept out of chat) and to stdout.
#
# Run from a Windows-local path, not \\wsl.localhost: PowerShell treats the UNC
# path as remote and refuses to load the script while still exiting 0. See
# AGENTS.md's Known Patterns.

param(
    [string]$DeckPath = "C:\Users\rohan\deck-sync-e2e\e2e-deck.pptx",
    [string]$RegisterPath = "C:\Users\rohan\deck-sync-e2e\register.xlsx",
    [string]$Period = "FY26Q4",
    [string]$GridPath = "C:\Users\rohan\deck-sync-e2e\review-grid.xlsx",
    [ValidateSet("review","apply","syncnow")][string]$Phase = "review",
    [string]$ApproveAll = "no",
    [string]$RepoRoot = $(Split-Path -Parent (Split-Path -Parent $PSScriptRoot)),
    [string]$OutFile = (Join-Path $env:TEMP "r13_real_deck_report.txt")
)

$ErrorActionPreference = "Stop"

foreach ($progId in @("PowerPoint.Application", "Excel.Application")) {
    try {
        $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject($progId)
        Write-Output "=== ABORTING: $progId is already running -- close it first so this run does not attach to a live session. ==="
        exit 2
    } catch {}
}

$staging = Join-Path $env:TEMP ("r13-run-" + [guid]::NewGuid().ToString("N").Substring(0,8))
New-Item -ItemType Directory -Path $staging | Out-Null
$vbaSourceDir = Join-Path $RepoRoot "vba"

# The same set run_vba_tests.ps1 imports, which is known to compile together.
$modules = @(
    "Discovery.bas","InjectPrimitive.bas","Matching.bas","Resolve.bas",
    "SyncOperations.bas","Onboarding.bas","ExcelOutput.bas","Verification.bas",
    "SlideDuplication.bas","TemplateSlide.bas","TemplateAudit.bas","IdentityCheck.bas",
    "TagMigration.bas","Register.bas","PlaceholderCheck.bas","RunSync.bas","ReviewQueue.bas",
    "DeckAdoption.bas","ResolveFields.bas","DeckRegistry.bas","WorkbookBridge.bas",
    "OnboardFlow.bas","RibbonUI.bas","AdoptFlow.bas","BatchOnboardFlow.bas","CommandBarUI.bas"
)
foreach ($m in $modules) { Copy-Item (Join-Path $vbaSourceDir $m) -Destination $staging }
Copy-Item (Join-Path $vbaSourceDir "tools\R13RealDeck.bas") -Destination $staging

$ppt = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $ppt.Visible = -1
    $scratch = $ppt.Presentations.Add()
    foreach ($m in ($modules + "R13RealDeck.bas")) {
        $scratch.VBProject.VBComponents.Import((Join-Path $staging $m)) | Out-Null
    }

    # Warm-up probe FIRST -- see R13RealDeck.PingR13. In a freshly Imported
    # project a Public Function is only reachable via Application.Run once the
    # cross-module Public UDTs it declares have been touched by an earlier
    # Application.Run in the same session. Skipping this fails as "Sub or
    # function not defined", which reads as a compile error in the new code.
    $probe = $ppt.GetType().InvokeMember("Run",[System.Reflection.BindingFlags]::InvokeMethod,$null,$ppt,@([string]"R13RealDeck.PingR13"))
    Write-Output "probe: $probe"

    if ($Phase -eq "review") {
        $report = $ppt.GetType().InvokeMember("Run",[System.Reflection.BindingFlags]::InvokeMethod,$null,$ppt,
            @([string]"R13RealDeck.ReviewOnly",[string]$DeckPath,[string]$RegisterPath,[string]$Period,[string]$GridPath))
    } elseif ($Phase -eq "syncnow") {
        $report = $ppt.GetType().InvokeMember("Run",[System.Reflection.BindingFlags]::InvokeMethod,$null,$ppt,
            @([string]"R13RealDeck.SyncNowPhase",[string]$DeckPath,[string]$RegisterPath,[string]$Period,[string]$GridPath))
    } else {
        $report = $ppt.GetType().InvokeMember("Run",[System.Reflection.BindingFlags]::InvokeMethod,$null,$ppt,
            @([string]"R13RealDeck.ApplyPhase",[string]$DeckPath,[string]$RegisterPath,[string]$Period,[string]$GridPath,[string]$ApproveAll))
    }

    Set-Content -Path $OutFile -Value $report
    Write-Output $report
} catch {
    Write-Output "=== DRIVER ERROR ==="
    Write-Output $_.Exception.Message
    exit 3
} finally {
    if ($ppt) { try { $ppt.Quit() } catch {} }
    Get-Process POWERPNT,EXCEL -ErrorAction SilentlyContinue | Where-Object { -not $_.MainWindowTitle } | Stop-Process -Force -ErrorAction SilentlyContinue
    Remove-Item -Recurse -Force $staging -ErrorAction SilentlyContinue
}
