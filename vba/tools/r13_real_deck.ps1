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
    "TagMigration.bas","Register.bas","PlaceholderCheck.bas","RunSync.bas","ReviewQueue.bas","Drafting.bas","FieldSpec.bas",
    "DeckAdoption.bas","ResolveFields.bas","DeckRegistry.bas","WorkbookBridge.bas",
    "OnboardFlow.bas","RibbonUI.bas","AdoptFlow.bas","BatchOnboardFlow.bas","CommandBarUI.bas"
)
foreach ($m in $modules) { Copy-Item (Join-Path $vbaSourceDir $m) -Destination $staging }
Copy-Item (Join-Path $vbaSourceDir "tools\R13RealDeck.bas") -Destination $staging

function Invoke-ForceCompile {
    param($App)
    # OPTIONAL, and it must never cost more than it saves.
    #
    # It was added to turn a would-be runtime mystery into a compile error at a
    # known point. On 2026-08-01 it became the mystery: touching
    # $App.VBE.MainWindow threw "The property 'Visible' cannot be found", and
    # PowerPoint then rejected every subsequent call (RPC_E_CALL_REJECTED) --
    # with no stale Office process anywhere. The suite, which does NOT compile,
    # ran 135 tests clean at the same time.
    #
    # So: reach for the VBE once, abandon the whole idea on any error, and do
    # not touch it again. The warm-up trap this was hedging against did not
    # reproduce when tested directly, so skipping the compile costs nothing
    # known -- while wedging the app costs the run.
    try {
        $vbe = $App.VBE
        if ($null -eq $vbe) { return "compile skipped (no VBE)" }
        $win = $vbe.MainWindow
        if ($null -eq $win) { return "compile skipped (no VBE window)" }
        $wasVisible = $win.Visible
        $win.Visible = $true
        $ctl = $vbe.CommandBars.Item("Menu Bar").Controls.Item("Debug").Controls.Item("Compile VBAProject")
        if ($null -eq $ctl) { $win.Visible = $wasVisible; return "compile control not found" }
        if (-not $ctl.Enabled) { $win.Visible = $wasVisible; return "already compiled" }
        $ctl.Execute()
        $win.Visible = $wasVisible
        return "compile executed"
    } catch {
        return "compile SKIPPED (VBE unreachable: " + $_.Exception.Message + ")"
    }
}

$ppt = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $ppt.Visible = -1
    $scratch = $ppt.Presentations.Add()
    foreach ($m in ($modules + "R13RealDeck.bas")) {
        $scratch.VBProject.VBComponents.Import((Join-Path $staging $m)) | Out-Null
    }

    Write-Output ("compile: " + (Invoke-ForceCompile -App $ppt))

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
