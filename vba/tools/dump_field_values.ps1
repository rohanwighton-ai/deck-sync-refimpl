# Dump every managed field's value AS POWERPOINT READS IT, to TSV.
#
# THE POINT OF THIS SCRIPT IS THE FORCED COMPILE, not the dump.
#
# The trap it removes: in a freshly VBComponents.Import()ed project, a Public
# Function is reachable via Application.Run ONLY if the cross-module Public UDTs
# it declares were already touched by an EARLIER Application.Run in the same
# session. Otherwise it fails as "Sub or function not defined" -- which reads as
# a compile error in the new code, while a human pressing Compile in the VBE
# sees a clean project. E2EFirstField.DumpFieldValues has been uncallable since
# 2026-07-31 morning for exactly this reason, and the workaround was a ladder of
# warm-up probes (Ping .. PingE) whose only job was to declare the UDTs first.
#
# The real fix is to compile the project after importing, rather than to trick
# it into compiling by calling something. Microsoft's documented route is the
# VBE's own Debug > Compile VBAProject menu control:
#
#   Application.VBE.CommandBars("Menu Bar").Controls("Debug") _
#       .Controls("Compile VBAProject").Execute
#
# Two documented caveats, both handled below:
#   - the VBE main window must be visible for the control to be executable
#   - Execute returns NOTHING. It does not report success, and the menu item is
#     greyed out when the project is already compiled. So this script does not
#     trust it: it proves the fix by OUTCOME -- calling DumpFieldValues with no
#     probes at all and seeing whether it resolves.
#
# Requires "Trust access to the VBA project object model" (already enabled on
# this machine, 2026-07-25, with explicit sign-off).

param(
    [string]$DeckPath = "C:\Users\rohan\deck-sync-e2e\e2e-deck.pptx",
    [string]$OutTsv   = "C:\Users\rohan\deck-sync-e2e\field_values.tsv",
    [string]$RepoRoot = $(Split-Path -Parent (Split-Path -Parent $PSScriptRoot)),
    [switch]$SkipCompile   # to demonstrate the failure this script exists to fix
)

$ErrorActionPreference = "Stop"

foreach ($progId in @("PowerPoint.Application","Excel.Application")) {
    try {
        $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject($progId)
        Write-Output "=== ABORTING: $progId is already running -- close it first. ==="
        exit 2
    } catch {}
}

$staging = Join-Path $env:TEMP ("dump-run-" + [guid]::NewGuid().ToString("N").Substring(0,8))
New-Item -ItemType Directory -Path $staging | Out-Null
$vbaSourceDir = Join-Path $RepoRoot "vba"

$modules = @(
    "Discovery.bas","InjectPrimitive.bas","Matching.bas","Resolve.bas",
    "SyncOperations.bas","Onboarding.bas","ExcelOutput.bas","Verification.bas",
    "SlideDuplication.bas","TemplateSlide.bas","TemplateAudit.bas","IdentityCheck.bas",
    "TagMigration.bas","Register.bas","PlaceholderCheck.bas","RunSync.bas","ReviewQueue.bas","Drafting.bas","FieldSpec.bas",
    "DeckAdoption.bas","ResolveFields.bas","DeckRegistry.bas","WorkbookBridge.bas",
    "OnboardFlow.bas","RibbonUI.bas","AdoptFlow.bas","BatchOnboardFlow.bas","CommandBarUI.bas"
)
foreach ($m in $modules) { Copy-Item (Join-Path $vbaSourceDir $m) -Destination $staging }
Copy-Item (Join-Path $vbaSourceDir "tools\E2EFirstField.bas") -Destination $staging

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
    foreach ($m in ($modules + "E2EFirstField.bas")) {
        $scratch.VBProject.VBComponents.Import((Join-Path $staging $m)) | Out-Null
    }

    if ($SkipCompile) {
        Write-Output "compile: SKIPPED on purpose (expecting the historic failure)"
    } else {
        Write-Output ("compile: " + (Invoke-ForceCompile -App $ppt))
    }

    # NO WARM-UP PROBES. Calling DumpFieldValues directly is the whole test:
    # if it resolves, the forced compile replaced the Ping ladder.
    $report = $ppt.GetType().InvokeMember("Run",[System.Reflection.BindingFlags]::InvokeMethod,$null,$ppt,
        @([string]"E2EFirstField.DumpFieldValues",[string]$DeckPath))

    Set-Content -Path $OutTsv -Value $report -Encoding UTF8
    $lineCount = ($report -split "`n").Count
    Write-Output "DumpFieldValues RESOLVED. $lineCount line(s) written to $OutTsv"
} catch {
    Write-Output "=== FAILED ==="
    Write-Output $_.Exception.Message
    exit 3
} finally {
    if ($ppt) { try { $ppt.Quit() } catch {} }
    Get-Process POWERPNT,EXCEL -ErrorAction SilentlyContinue | Where-Object { -not $_.MainWindowTitle } | Stop-Process -Force -ErrorAction SilentlyContinue
    Remove-Item -Recurse -Force $staging -ErrorAction SilentlyContinue
}
