# Tests the documented fix for "stuck after the first write": close the
# cloud-hosted presentation and reopen it from its own path before retrying,
# which several Microsoft Q&A reports describe as forcing a genuine resync.
# Uses SaveAsSelfProbe.bas's ProbeCloseReopenRescue, which deliberately
# reproduces the bug (first write lands, second write on the same open
# session gets stuck) before testing whether the rescue frees it.
#
#   powershell.exe -ExecutionPolicy Bypass -File onedrive_close_reopen_probe.ps1

param(
    [string]$SourceDir = $PSScriptRoot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

foreach ($progId in @("PowerPoint.Application")) {
    try {
        [System.Runtime.InteropServices.Marshal]::GetActiveObject($progId) | Out-Null
        Write-Output "=== ABORTING: $progId is already running -- close it first. ==="
        exit 2
    } catch {}
}

$modulePath = Join-Path $SourceDir "SaveAsSelfProbe.bas"
if (-not (Test-Path $modulePath)) { throw "missing module: $modulePath" }
$localModule = Join-Path $env:TEMP "SaveAsSelfProbe.bas"
Copy-Item -Path $modulePath -Destination $localModule -Force

$scratchDir = Join-Path "$env:USERPROFILE\OneDrive\Claude" "close-reopen-probe-$(Get-Date -Format yyyyMMdd-HHmmss)"
New-Item -ItemType Directory -Path $scratchDir | Out-Null
$scratchFile = Join-Path $scratchDir "probe.pptx"

$ppt = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $ppt.Visible = -1

    $blank = $ppt.Presentations.Add()
    $blank.SaveAs($scratchFile, 24)
    $blank.Close()
    Start-Sleep -Seconds 5
    $pres = $ppt.Presentations.Open($scratchFile)
    Write-Output "Recognised as cloud URL: $($pres.FullName -like 'http*')"
    Write-Output ""

    $pres.VBProject.VBComponents.Import($localModule) | Out-Null

    $result = $ppt.GetType().InvokeMember("Run", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ppt,
        @([string]"SaveAsSelfProbe.ProbeScopeCheck", $pres, [string]"DeckSyncPeriod", [string]"DeckSyncWorkbookPath"))

    Write-Output $result
} catch {
    Write-Output "=== DRIVER ERROR ==="
    Write-Output $_.Exception.Message
    exit 3
} finally {
    if ($ppt) {
        try { foreach ($p in $ppt.Presentations) { $p.Saved = -1 } } catch {}
        try { $ppt.Quit() } catch {}
    }
    Get-Process POWERPNT -ErrorAction SilentlyContinue |
        Where-Object { -not $_.MainWindowTitle } |
        Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 1
    Remove-Item -Path $scratchDir -Recurse -Force -ErrorAction SilentlyContinue
}
