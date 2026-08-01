# Driver for E2EField.bas -- take ONE field end to end on the deck COPY.
#
#   -Mode dryrun  (default)  lists every ABOUT_BODY change with before/after,
#                            writes nothing. This is the R13 gate.
#   -Mode apply              writes, verifies by re-reading the deck, saves.
#
# Read the dryrun output before running apply. That reading IS the review.
#
# Defaults point at the scratch rig. Never pass the real deck.

param(
    [string]$DeckPath = "C:\Users\rohan\deck-sync-e2e\e2e-deck.pptx",
    [string]$RegisterPath = "C:\Users\rohan\deck-sync-e2e\register.xlsx",
    [string]$Period = "FY26Q4",
    [ValidateSet("migrate","dryrun","apply","reseed","verifyharvest","deleteentities","draft","copyai","publish","setperiod","setperiodvariant")][string]$Mode = "dryrun",
    [string]$Variant = "save",
    [string]$FieldId = "ABOUT_BODY",
    [string]$TsvPath = "C:\Users\rohan\deck-sync-e2e\field_values.tsv",
    [string]$Entities = "",
    # Publishing writes with -Write. It briefly reused $ApproveAll, which is a
    # parameter of the OTHER driver and was never declared here -- so it was
    # undefined, and PowerShell evaluated it as false rather than erroring,
    # which made every publish run silently preview. Set-StrictMode below turns
    # that class of mistake into a failure instead of a wrong answer.
    [switch]$Write,
    [string]$RepoRoot = $(Split-Path -Parent (Split-Path -Parent $PSScriptRoot)),
    [string]$OutFile = (Join-Path $env:TEMP "about_body_e2e_report.txt")
)

Set-StrictMode -Version Latest   # an undefined variable must ERROR, not read as false
$ErrorActionPreference = "Stop"

foreach ($progId in @("PowerPoint.Application","Excel.Application")) {
    try {
        $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject($progId)
        Write-Output "=== ABORTING: $progId is already running -- close it first. ==="
        exit 2
    } catch {}
}

$staging = Join-Path $env:TEMP ("ab-run-" + [guid]::NewGuid().ToString("N").Substring(0,8))
New-Item -ItemType Directory -Path $staging | Out-Null
$vbaSourceDir = Join-Path $RepoRoot "vba"

$modules = @(
    "Discovery.bas","InjectPrimitive.bas","Matching.bas","Resolve.bas",
    "SyncOperations.bas","Onboarding.bas","ExcelOutput.bas","Verification.bas",
    "SlideDuplication.bas","TemplateSlide.bas","TemplateAudit.bas","IdentityCheck.bas",
    "TagMigration.bas","Register.bas","PlaceholderCheck.bas","RunSync.bas","ReviewQueue.bas","Drafting.bas","FieldSpec.bas","Sources.bas",
    "DeckAdoption.bas","ResolveFields.bas","DeckRegistry.bas","WorkbookBridge.bas",
    "OnboardFlow.bas","RibbonUI.bas","AdoptFlow.bas","BatchOnboardFlow.bas","CommandBarUI.bas"
)
foreach ($m in $modules) { Copy-Item (Join-Path $vbaSourceDir $m) -Destination $staging }
Copy-Item (Join-Path $vbaSourceDir "tools\E2EField.bas") -Destination $staging

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
    foreach ($m in ($modules + "E2EField.bas")) {
        $scratch.VBProject.VBComponents.Import((Join-Path $staging $m)) | Out-Null
    }

    # Was a warm-up probe (E2EAboutBody.PingAB) until 2026-07-31. The probe
    # existed to work around "Application.Run cannot see a function that
    # declares cross-module Public UDTs unless an earlier Run touched them".
    # That trap did NOT reproduce when tested directly: DumpFieldValues resolved
    # with no probe and no compile. All three "Sub or function not defined"
    # failures found today had ordinary causes (a module missing from the
    # driver's import list, a reserved word, a rename that missed a return
    # line), so the probes look like a workaround for a misattributed symptom.
    # Compiling at a known point is kept because it is cheap and turns a
    # would-be runtime mystery into a compile error where it belongs.
    Write-Output ("compile: " + (Invoke-ForceCompile -App $ppt))

    if ($Mode -eq "reseed") {
        $report = $ppt.GetType().InvokeMember("Run",[System.Reflection.BindingFlags]::InvokeMethod,$null,$ppt,
            @([string]"E2EField.ReseedFromSlides",[string]$DeckPath,[string]$RegisterPath,[string]$Period,[string]$Entities,[string]$FieldId))
    } elseif ($Mode -eq "setperiod") {
        $report = $ppt.GetType().InvokeMember("Run",[System.Reflection.BindingFlags]::InvokeMethod,$null,$ppt,
            @([string]"E2EField.SetPeriod",[string]$DeckPath,[string]$Period))
    } elseif ($Mode -eq "setperiodvariant") {
        $report = $ppt.GetType().InvokeMember("Run",[System.Reflection.BindingFlags]::InvokeMethod,$null,$ppt,
            @([string]"E2EField.SetPeriodVariant",[string]$DeckPath,[string]$Period,[string]$Variant))
    } elseif ($Mode -eq "draft") {
        $report = $ppt.GetType().InvokeMember("Run",[System.Reflection.BindingFlags]::InvokeMethod,$null,$ppt,
            @([string]"E2EField.BuildDraftSheet",[string]$RegisterPath,[string]$Period,[string]$FieldId))
    } elseif ($Mode -eq "copyai") {
        $report = $ppt.GetType().InvokeMember("Run",[System.Reflection.BindingFlags]::InvokeMethod,$null,$ppt,
            @([string]"E2EField.CopyAiToSubmitSheet",[string]$RegisterPath,[string]$FieldId))
    } elseif ($Mode -eq "publish" -or $Mode -eq "publishapply") {
        $report = $ppt.GetType().InvokeMember("Run",[System.Reflection.BindingFlags]::InvokeMethod,$null,$ppt,
            @([string]"E2EField.PublishDraftSheet",[string]$RegisterPath,[string]$FieldId,[string]$(if ($Write) { "apply" } else { "preview" })))
    } elseif ($Mode -eq "deleteentities") {
        $report = $ppt.GetType().InvokeMember("Run",[System.Reflection.BindingFlags]::InvokeMethod,$null,$ppt,
            @([string]"E2EField.DeleteEntities",[string]$DeckPath,[string]$RegisterPath,[string]$Entities))
    } elseif ($Mode -eq "verifyharvest") {
        $report = $ppt.GetType().InvokeMember("Run",[System.Reflection.BindingFlags]::InvokeMethod,$null,$ppt,
            @([string]"E2EField.VerifyHarvest",[string]$DeckPath,[string]$TsvPath,[string]$FieldId))
    } else {
        $report = $ppt.GetType().InvokeMember("Run",[System.Reflection.BindingFlags]::InvokeMethod,$null,$ppt,
            @([string]"E2EField.RunField",[string]$DeckPath,[string]$RegisterPath,[string]$Period,[string]$Mode,[string]$FieldId))
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
