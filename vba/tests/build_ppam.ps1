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
    "SlideDuplication.bas", "TemplateSlide.bas", "TemplateAudit.bas", "IdentityCheck.bas", "TagMigration.bas", "Register.bas", "RunSync.bas", "DeckAdoption.bas", "ResolveFields.bas",
    "DeckRegistry.bas", "WorkbookBridge.bas", "OnboardFlow.bas", "RibbonUI.bas",
    "AdoptFlow.bas", "BatchOnboardFlow.bas", "CommandBarUI.bas"
)

function Request-GracefulQuit {
    param([string]$ProgId, [int]$TimeoutSeconds = 30)
    $job = Start-Job -ScriptBlock {
        param($ProgId)
        try { $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject($ProgId) }
        catch { return "no-instance" }
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

foreach ($m in $productionModules) {
    $comp = $pres.VBProject.VBComponents.Import((Join-Path $vbaSourceDir $m))
    Write-Output ("Imported $m as: " + $comp.Name)
}

Write-Output ""
Write-Output "=== Ready. In PowerPoint: File > Save As > PowerPoint Add-in (*.ppam). ==="
Write-Output "If PowerPoint warns that macro-free formats can't save VBA, click No, then explicitly pick 'PowerPoint Add-in (*.ppam)' in the Save as type dropdown before saving."
