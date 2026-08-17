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
    [ValidateSet("migrate","dryrun","apply","reseed","verifyharvest","deleteentities","draft","copyai","publish","timelinetest","discovertest","setperiod","setperiodvariant","readwide","repointworkbook","renameslidetype")][string]$Mode = "dryrun",
    [string]$Variant = "save",
    [string]$FieldId = "ABOUT_BODY",
    [string]$SheetName = "Register",
    # NO DEFAULT. It defaulted to "q" -- the rig's original slide-type name,
    # renamed to "project-status" on 2026-08-04 (80fe9af). The rename changed one
    # deck property and left this default, so any run that did not pass
    # -SlideType silently addressed a type that no longer exists. Same defect as
    # DraftingUI's hardcoded "q", which had been rejecting every register row.
    # A slide type is deck-specific: there is no sane default, so there is none.
    [string]$SlideType = "",
    [string]$NewSlideType = "",
    [string]$TsvPath = "C:\Users\rohan\deck-sync-e2e\field_values.tsv",
    [string]$Entities = "",
    # Publishing writes with -Write. It briefly reused $ApproveAll, which is a
    # parameter of the OTHER driver and was never declared here -- so it was
    # undefined, and PowerShell evaluated it as false rather than erroring,
    # which made every publish run silently preview. Set-StrictMode below turns
    # that class of mistake into a failure instead of a wrong answer.
    [switch]$Write,
    [string]$NewWorkbookPath = "",
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
    "TagMigration.bas","FieldWiring.bas","MilestoneDevice.bas","Register.bas","RegisterSeed.bas","PlaceholderCheck.bas","RunSync.bas","ReviewQueue.bas","Drafting.bas","FieldSpec.bas","Sources.bas","Timeline.bas",
    "DeckAdoption.bas","ResolveFields.bas","DeckRegistry.bas","WorkbookBridge.bas","Harvest.bas",
    "OnboardFlow.bas","RibbonUI.bas","AdoptFlow.bas","BatchOnboardFlow.bas","CommandBarUI.bas","DraftingUI.bas","DiscoverUI.bas",
    "DraftingLobby.bas","AppEvents.cls","ShapeAddressBook.bas","Timing.bas"
)
foreach ($m in $modules) {
    $srcPath = Join-Path $vbaSourceDir $m
    if ($m -like "*.cls") {
        # CRLF, NOT A PLAIN COPY -- see run_vba_tests.ps1's identical step:
        # Import() silently mis-types an LF-only .cls as a Standard Module,
        # and WithEvents then fails to compile nowhere near this cause.
        ((Get-Content $srcPath -Raw) -replace "`r`n", "`n" -replace "`n", "`r`n") | Set-Content (Join-Path $staging $m) -NoNewline
    } else {
        Copy-Item $srcPath -Destination $staging
    }
}
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
    # Same reference build_ppam.ps1/run_vba_tests.ps1/compile_check.ps1 all add
    # -- AppEvents.cls's WithEvents needs it (LOBBY-DESIGN.md).
    $scratch.VBProject.References.AddFromGuid("{00020813-0000-0000-C000-000000000046}", 1, 9) | Out-Null
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

    if ($Mode -eq "renameslidetype") {
        if ([string]::IsNullOrWhiteSpace($SlideType)) {
            throw "-SlideType is required for renameslidetype. It has no default: a slide type belongs to a deck, and guessing one silently addresses a type that may not exist."
        }
        $report = $ppt.GetType().InvokeMember("Run",[System.Reflection.BindingFlags]::InvokeMethod,$null,$ppt,
            @([string]"E2EField.RenameSlideType",[string]$DeckPath,[string]$SlideType,[string]$NewSlideType,[string]$SheetName,[string]$Variant))
    } elseif ($Mode -eq "repointworkbook") {
        $report = $ppt.GetType().InvokeMember("Run",[System.Reflection.BindingFlags]::InvokeMethod,$null,$ppt,
            @([string]"E2EField.RepointWorkbookVariant",[string]$DeckPath,[string]$NewWorkbookPath,[string]$Variant))
    } elseif ($Mode -eq "readwide") {
        # Read-only. -Entities carries the instance to print in full, because a
        # count alone has never been enough evidence on this project.
        $sample = if ($Entities -ne "") { $Entities } else { "" }
        $report = $ppt.GetType().InvokeMember("Run",[System.Reflection.BindingFlags]::InvokeMethod,$null,$ppt,
            @([string]"E2EField.ReadWidePeriod",[string]$RegisterPath,[string]$SheetName,[string]$Period,[string]$sample))
    } elseif ($Mode -eq "reseed") {
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
    } elseif ($Mode -eq "discovertest") {
        $report = $ppt.GetType().InvokeMember("Run",[System.Reflection.BindingFlags]::InvokeMethod,$null,$ppt,
            @([string]"E2EField.DiscoverSelfTest",[string]$DeckPath,[string]$RegisterPath))
    } elseif ($Mode -eq "timelinetest") {
        $report = $ppt.GetType().InvokeMember("Run",[System.Reflection.BindingFlags]::InvokeMethod,$null,$ppt,
            @([string]"Timeline.SelfTest"))
    } elseif ($Mode -eq "copyai") {
        $report = $ppt.GetType().InvokeMember("Run",[System.Reflection.BindingFlags]::InvokeMethod,$null,$ppt,
            @([string]"E2EField.CopyAiToSubmitSheet",[string]$RegisterPath,[string]$FieldId))
    } elseif ($Mode -eq "publish") {
        $report = $ppt.GetType().InvokeMember("Run",[System.Reflection.BindingFlags]::InvokeMethod,$null,$ppt,
            @([string]"E2EField.PublishDraftSheet",[string]$RegisterPath,[string]$Period,[string]$FieldId,[string]$(if ($Write) { "apply" } else { "preview" })))
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
