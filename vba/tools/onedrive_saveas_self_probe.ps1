# Tests the hypothesis behind SaveAsSelfProbe.bas: that `pres.SaveAs
# pres.FullName, 24` (SaveAs-TO-SELF, unmangled path) reliably lands a custom
# document property write on a OneDrive-hosted deck, now that the ByRef path-
# corruption bug (FIX-LIST P, fixed 2026-08-15 afternoon) can no longer make
# that SaveAs silently target the wrong location. FIX-LIST P's own "SaveAs
# bricks a cloud deck" measurement was taken BEFORE that fix and is suspected
# to be an artifact of it, not a fact about a clean SaveAs-to-self.
#
# Each trial gets its OWN fresh scratch presentation and its own OneDrive
# subfolder -- if SaveAs-to-self really does brick a document, that damage
# stays contained to one trial's throwaway file, never touches Rohan's real
# deck, and never poisons the next trial. Every subfolder is deleted on exit.
#
#   powershell.exe -ExecutionPolicy Bypass -File onedrive_saveas_self_probe.ps1

param(
    [int]$Trials = 5,
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

# VBComponents.Import needs a genuine Windows-local path, not a \\wsl.localhost\
# UNC path -- copy it local first (same caveat tag_cloning_probe.ps1 already
# documents).
$localModule = Join-Path $env:TEMP "SaveAsSelfProbe.bas"
Copy-Item -Path $modulePath -Destination $localModule -Force

$claudeDir = "$env:USERPROFILE\OneDrive\Claude"
$results = @()
$ppt = $null

try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $ppt.Visible = -1

    for ($i = 1; $i -le $Trials; $i++) {
        $scratchDir = Join-Path $claudeDir "saveas-self-probe-$(Get-Date -Format yyyyMMdd-HHmmss)-$i"
        New-Item -ItemType Directory -Path $scratchDir | Out-Null
        $scratchFile = Join-Path $scratchDir "probe.pptx"

        $blank = $ppt.Presentations.Add()
        $blank.SaveAs($scratchFile, 24)
        $blank.Close()
        Start-Sleep -Seconds 5
        $pres = $ppt.Presentations.Open($scratchFile)

        $isCloud = $pres.FullName -like "http*"
        if (-not $isCloud) {
            Start-Sleep -Seconds 8
            $pres.Close()
            $pres = $ppt.Presentations.Open($scratchFile)
            $isCloud = $pres.FullName -like "http*"
        }

        $pres.VBProject.VBComponents.Import($localModule) | Out-Null

        $value = "SASPROBE-$i-$(Get-Random -Maximum 9999)"
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $result = $null
        try {
            $result = $ppt.GetType().InvokeMember("Run", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ppt,
                @([string]"SaveAsSelfProbe.ProbeSaveAsToSelfVerified", $pres, [string]"DeckSyncPeriod", [string]$value, [long]4))
        } catch {
            $result = "DRIVER-LEVEL ERROR calling into VBA: $($_.Exception.Message)"
        }
        $sw.Stop()

        $ok = ($result -eq "" -or $result -like "(landed*")
        $results += [PSCustomObject]@{
            Trial   = $i
            IsCloud = $isCloud
            Landed  = $ok
            Seconds = [math]::Round($sw.Elapsed.TotalSeconds, 1)
            Detail  = $result
        }
        Write-Output ("Trial {0} (cloud={1}): {2} in {3}s -- {4}" -f $i, $isCloud, $(if ($ok) {"LANDED"} else {"FAILED"}), [math]::Round($sw.Elapsed.TotalSeconds,1), $result)

        # Independent cross-check outside VBA: read straight from the saved package.
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        try {
            $zip = [System.IO.Compression.ZipFile]::OpenRead($scratchFile)
            $entry = $zip.Entries | Where-Object { $_.FullName -eq "docProps/custom.xml" }
            $xmlText = ""
            if ($entry) {
                $reader = New-Object System.IO.StreamReader($entry.Open())
                $xmlText = $reader.ReadToEnd()
                $reader.Close()
            }
            $zip.Dispose()
            $independentlyConfirmed = $xmlText -match [regex]::Escape($value)
            Write-Output ("  independent .NET zip read confirms value present: {0}" -f $independentlyConfirmed)
        } catch {
            Write-Output "  independent zip read could not open the file (possibly still syncing): $($_.Exception.Message)"
        }

        try { $pres.Saved = -1; $pres.Close() } catch {}
        Start-Sleep -Seconds 1
        Remove-Item -Path $scratchDir -Recurse -Force -ErrorAction SilentlyContinue
    }

    Write-Output ""
    Write-Output "=== SUMMARY ==="
    $landedCount = ($results | Where-Object { $_.Landed }).Count
    Write-Output "$landedCount / $Trials trials confirmed on disk"
    $results | Format-Table -AutoSize
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
    Get-ChildItem $claudeDir -Directory -Filter "saveas-self-probe-*" -ErrorAction SilentlyContinue |
        Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
}
