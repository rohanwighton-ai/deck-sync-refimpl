# Driver for E2EAboutBody.bas -- FIELD 2 end to end on the deck COPY.
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
    [ValidateSet("migrate","dryrun","apply","reseed")][string]$Mode = "dryrun",
    [string]$Entities = "",
    [string]$RepoRoot = $(Split-Path -Parent (Split-Path -Parent $PSScriptRoot)),
    [string]$OutFile = (Join-Path $env:TEMP "about_body_e2e_report.txt")
)

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
    "TagMigration.bas","Register.bas","PlaceholderCheck.bas","RunSync.bas","ReviewQueue.bas",
    "DeckAdoption.bas","ResolveFields.bas","DeckRegistry.bas","WorkbookBridge.bas",
    "OnboardFlow.bas","RibbonUI.bas","AdoptFlow.bas","BatchOnboardFlow.bas","CommandBarUI.bas"
)
foreach ($m in $modules) { Copy-Item (Join-Path $vbaSourceDir $m) -Destination $staging }
Copy-Item (Join-Path $vbaSourceDir "tools\E2EAboutBody.bas") -Destination $staging

$ppt = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $ppt.Visible = -1
    $scratch = $ppt.Presentations.Add()
    foreach ($m in ($modules + "E2EAboutBody.bas")) {
        $scratch.VBProject.VBComponents.Import((Join-Path $staging $m)) | Out-Null
    }

    $probe = $ppt.GetType().InvokeMember("Run",[System.Reflection.BindingFlags]::InvokeMethod,$null,$ppt,@([string]"E2EAboutBody.PingAB"))
    Write-Output "probe: $probe"

    if ($Mode -eq "reseed") {
        $report = $ppt.GetType().InvokeMember("Run",[System.Reflection.BindingFlags]::InvokeMethod,$null,$ppt,
            @([string]"E2EAboutBody.ReseedFromSlides",[string]$DeckPath,[string]$RegisterPath,[string]$Period,[string]$Entities))
    } else {
        $report = $ppt.GetType().InvokeMember("Run",[System.Reflection.BindingFlags]::InvokeMethod,$null,$ppt,
            @([string]"E2EAboutBody.RunAboutBody",[string]$DeckPath,[string]$RegisterPath,[string]$Period,[string]$Mode))
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
