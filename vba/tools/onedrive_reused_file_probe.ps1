# Isolates ONE variable left dangling by onedrive_saveas_self_probe.ps1's 5/5
# success: that test used a FRESH scratch file per trial. The production
# functions (SetDeckPeriodVerified etc.) reuse ONE open presentation across
# many repeated saves in a session -- and a live run against the real fixed
# add-in (addin108) failed 8/8, fast (~0.4s each), a completely different
# signature than the old wait-based failures. This tests whether REPEATED
# SaveAs-to-self against the SAME open cloud file is the actual discriminator,
# using the same already-proven SaveAsSelfProbe.bas module so the only thing
# that changes is fresh-file-per-trial vs one-file-for-all-trials.
#
#   powershell.exe -ExecutionPolicy Bypass -File onedrive_reused_file_probe.ps1

param(
    [int]$Trials = 8,
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

$scratchDir = Join-Path "$env:USERPROFILE\OneDrive\Claude" "reused-file-probe-$(Get-Date -Format yyyyMMdd-HHmmss)"
New-Item -ItemType Directory -Path $scratchDir | Out-Null
$scratchFile = Join-Path $scratchDir "probe.pptx"

$ppt = $null
$results = @()
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $ppt.Visible = -1

    $blank = $ppt.Presentations.Add()
    $blank.SaveAs($scratchFile, 24)
    $blank.Close()
    Start-Sleep -Seconds 5
    $pres = $ppt.Presentations.Open($scratchFile)
    Write-Output "FullName after reopen: $($pres.FullName)"
    Write-Output "Recognised as cloud URL: $($pres.FullName -like 'http*')"
    Write-Output ""

    $pres.VBProject.VBComponents.Import($localModule) | Out-Null

    for ($i = 1; $i -le $Trials; $i++) {
        $value = "REUSEPROBE-$i-$(Get-Random -Maximum 9999)"
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $result = $ppt.GetType().InvokeMember("Run", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ppt,
            @([string]"SaveAsSelfProbe.ProbeSaveAsToSelfVerified", $pres, [string]"DeckSyncPeriod", [string]$value, [long]4))
        $sw.Stop()
        $ok = ($result -eq "" -or $result -like "(landed*")
        $results += [PSCustomObject]@{ Trial = $i; Landed = $ok; Seconds = [math]::Round($sw.Elapsed.TotalSeconds,1) }
        Write-Output ("Trial {0}: {1} in {2}s -- {3}" -f $i, $(if ($ok) {"LANDED"} else {"FAILED"}), [math]::Round($sw.Elapsed.TotalSeconds,1), $result)
    }

    Write-Output ""
    Write-Output "=== SUMMARY ==="
    $landedCount = ($results | Where-Object { $_.Landed }).Count
    Write-Output "$landedCount / $Trials trials confirmed on disk (SAME file/presentation reused across all trials)"
    $results | Format-Table -AutoSize

    $pres.Saved = -1
    $pres.Close()
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
