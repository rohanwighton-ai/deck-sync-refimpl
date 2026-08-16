# Prepares deck-sync-refimpl's VBA engine for packaging as a real .ppam
# PowerPoint add-in, per DECISIONS.md's 2026-07-25 "COM add-in first" call.
#
# This is NOT a fully automated build -- confirmed 2026-07-26 against real
# Office that it can't be, for two independent reasons:
#   1. Presentation.SaveAs(path, 30) [ppSaveAsOpenXMLAddin] throws a bare
#      COMException in this environment, reproducing even on a completely
#      blank, VBA-free presentation. Every other OOXML SaveAs enum
#      (pptx/pptm/potx/potm/ppsx/ppsm, 24-29) works fine via the same call
#      shape -- this one specifically doesn't, for reasons that don't
#      appear to be discoverable from outside Office.
#   2. A real .ppam's add-in loader rejects the package outright if it
#      contains anything beyond its exact expected part set (confirmed by
#      reverse-engineering a real hand-saved .ppam's structure, then
#      proving even a harmless, completely unrelated, unreferenced dummy
#      part breaks loading identically) -- so a vbaProject.bin built by
#      importing modules into a fresh, never-through-the-UI presentation
#      and spliced into a verified-real shell still doesn't load. Whatever
#      makes a real "Save As Add-in" produce a loadable package isn't
#      reproducible by assembling the same-looking bytes from outside.
#
# So this script does the one half that DOES fully automate reliably
# (importing every production .bas module into a fresh presentation, which
# has been proven solid all project), and stops there -- the actual
# File > Save As > PowerPoint Add-in (*.ppam) click is a real, permanent
# manual step. Per the 2026-07-26 conversation record: this is a
# release/packaging step, not a dev-loop tax -- the dev loop (edit .bas,
# run run_vba_tests.ps1) stays fully automated regardless. Once a .ppam
# exists, updating its code needs this same full cycle again: there is no
# way to edit an already-loaded add-in's VBA project via automation either
# (confirmed: AddIn.VBProject and VBE.VBProjects both expose a
# VBComponents property that is genuinely null, not just empty, for a
# loaded add-in -- almost certainly deliberate, the same class of
# self-modifying-macro restriction AccessVBOM's own default-off posture
# guards against).
#
# The ribbon (customUI14.xml) is NOT part of this pass at all -- see
# vba/customUI/customUI14.xml's header comment: it's provably impossible
# for a .ppam (same "anything beyond the exact expected part set" rejection
# above). CommandBarUI.bas's Auto_Open-driven toolbar is the shipped UI
# instead, and needs no packaging step of its own -- it's pure runtime VBA
# code, already inside vbaProject.bin once imported.

param(
    [string]$RepoRoot = $(Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
)

$ErrorActionPreference = "Stop"
$vbaSourceDir = Join-Path $RepoRoot "vba"

# Production modules only -- no tests\*.bas. Order doesn't matter for VBA
# compilation (cross-module references resolve after every component is
# loaded), kept in the same port/build order as run_vba_tests.ps1 for
# consistency.
$productionModules = @(
    "Discovery.bas", "InjectPrimitive.bas", "Matching.bas", "Resolve.bas",
    "SyncOperations.bas", "Onboarding.bas", "ExcelOutput.bas", "Verification.bas",
    "SlideDuplication.bas", "TemplateSlide.bas", "TemplateAudit.bas", "IdentityCheck.bas", "TagMigration.bas", "PlaceholderCheck.bas", "RunSync.bas", "DeckAdoption.bas", "ResolveFields.bas",
    "DeckRegistry.bas", "WorkbookBridge.bas", "Harvest.bas", "RibbonUI.bas",
    "AdoptFlow.bas", "BatchOnboardFlow.bas", "CommandBarUI.bas",
    # ReviewQueue was MISSING from this list while RibbonUI.SyncNow called it in
    # nine places. An undefined module reference is a COMPILE error in VBA, and a
    # compile error takes out the entire project -- so the built .ppam could not
    # run Sync Now, Review Changes or Apply Approved. Found 2026-08-01 while
    # adding the drafting buttons, not by anyone using the add-in, which is its
    # own finding: nothing verifies that this list covers what the modules
    # actually reference.
    "ReviewQueue.bas", "Readiness.bas", "FieldWiring.bas", "MilestoneDevice.bas",
    # The drafting half. Present in the repo since 2026-07-31 and never shipped
    # in the add-in at all -- reachable only from the PowerShell test harness.
    "Drafting.bas", "FieldSpec.bas", "Sources.bas",  "DraftingUI.bas", "DiscoverUI.bas",
    "DraftingLobby.bas", "AppEvents.cls"
)

function Request-GracefulQuit {
    param([string]$ProgId, [int]$TimeoutSeconds = 30)
    $job = Start-Job -ScriptBlock {
        param($ProgId)
        try { $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject($ProgId) }
        catch { return "no-instance" }
        # SAVE EVERY OPEN PRESENTATION BEFORE QUITTING. FIX-LIST item O: this
        # used to call Quit() directly, on the documented assumption that an
        # unsaved deck would raise a modal and safely time this job out. Found
        # 2026-08-16: that assumption was wrong at least once -- a real,
        # deliberate tag edit (unsaved) was silently discarded with no prompt
        # and no timeout, because Quit() inside a background job's isolated
        # COM apartment doesn't necessarily surface a UI prompt at all. An
        # explicit Save() here is the actual fix; the old "it'll time out"
        # behavior was never a real safety net, just an assumption that
        # happened not to be tested against the failure it was meant to catch.
        try {
            foreach ($pres in @($app.Presentations)) {
                if (-not $pres.Saved) { $pres.Save() }
            }
        } catch { }
        try { $app.Quit(); return "quit-ok" } catch { return "quit-error: $($_.Exception.Message)" }
    } -ArgumentList $ProgId
    $completed = Wait-Job $job -Timeout $TimeoutSeconds
    if (-not $completed) { Stop-Job $job -ErrorAction SilentlyContinue; Remove-Job $job -Force -ErrorAction SilentlyContinue; return "timeout" }
    $result = Receive-Job $job
    Remove-Job $job -ErrorAction SilentlyContinue
    return $result
}
foreach ($progId in @("PowerPoint.Application")) {
    $outcome = Request-GracefulQuit -ProgId $progId
    if ($outcome -eq "timeout" -or $outcome -like "quit-error*") {
        Write-Output "=== ABORTED: $progId did not close cleanly ($outcome). ==="
        exit 2
    }
    Start-Sleep -Seconds 1
}

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = -1
$pres = $ppt.Presentations.Add()
$pres.Slides.Add(1, 12) | Out-Null

# STAMP THE BUILD, so the add-in can say how old it is.
#
# Modules are staged and the stamp written into the COPY -- the repo file keeps
# saying "(unbuilt)", which is itself the signal that something is running
# straight from source rather than from a built package.
$stamp = Get-Date -Format "yyyy-MM-dd HH:mm"
$stageDir = Join-Path $env:TEMP ("deck-sync-ppam-" + [guid]::NewGuid().ToString("N").Substring(0,8))
New-Item -ItemType Directory -Path $stageDir | Out-Null
foreach ($m in $productionModules) {
    $src = Join-Path $vbaSourceDir $m
    $dst = Join-Path $stageDir $m
    if ($m -eq "CommandBarUI.bas") {
        (Get-Content $src -Raw) -replace 'BUILD_STAMP As String = "\(unbuilt\)"', ('BUILD_STAMP As String = "' + $stamp + '"') | Set-Content $dst -NoNewline
        Write-Output ("Stamped build: " + $stamp)
    } elseif ($m -like "*.cls") {
        # CRLF, NOT A PLAIN COPY -- see run_vba_tests.ps1's identical step for
        # why: Import() silently mis-types an LF-only .cls as a Standard
        # Module instead of a Class Module, and WithEvents then fails to
        # compile nowhere near this actual cause.
        ((Get-Content $src -Raw) -replace "`r`n", "`n" -replace "`n", "`r`n") | Set-Content $dst -NoNewline
    } else {
        Copy-Item $src $dst
    }
}

# EXCEL OBJECT LIBRARY REFERENCE -- required for AppEvents.cls's
# `WithEvents App As Excel.Application`. VBA's WithEvents needs an
# EARLY-BOUND type to sink events from; a late-bound `As Object` simply does
# not compile. This VBA project lives inside PowerPoint (COM-add-in-first,
# per DECISIONS.md), so watching an EXCEL event (SheetChange, for the
# Drafting Lobby's pin-on-tick mechanism -- see LOBBY-DESIGN.md) means the
# PowerPoint VBA project needs a reference to Excel's own object library.
# Added programmatically, every build, because build_ppam.ps1 always starts
# from a brand-new presentation -- a reference added by hand in the VBE
# would not survive the next build. GUID and version confirmed against this
# machine's actual registered type library (2026-08-16):
#   HKLM\SOFTWARE\Classes\TypeLib\{00020813-0000-0000-C000-000000000046}\1.9
#   = "Microsoft Excel 16.0 Object Library"
try {
    $pres.VBProject.References.AddFromGuid("{00020813-0000-0000-C000-000000000046}", 1, 9) | Out-Null
    Write-Output "Added reference: Microsoft Excel 16.0 Object Library"
} catch {
    Write-Output ("=== ABORTED: could not add the Excel object library reference (" + $_.Exception.Message + "). ===")
    exit 2
}

foreach ($m in $productionModules) {
    $comp = $pres.VBProject.VBComponents.Import((Join-Path $stageDir $m))
    Write-Output ("Imported $m as: " + $comp.Name)
}

Write-Output ""
Write-Output "=== Ready. In PowerPoint: File > Save As > PowerPoint Add-in (*.ppam). ==="
Write-Output "If PowerPoint warns that macro-free formats can't save VBA, click No, then explicitly pick 'PowerPoint Add-in (*.ppam)' in the Save as type dropdown before saving."
