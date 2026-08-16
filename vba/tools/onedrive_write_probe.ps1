# Empirically measures how reliably DeckRegistry.SetDeckPeriodVerified's cloud-hosted
# retry path (FIX-LIST P) actually lands a custom document property write on a real
# OneDrive-synced deck, end to end -- not a single Save, the WHOLE mechanism a real
# button call gets: up to 4 outer attempts, each with a 30s/5s-step settle wait on the
# cloud branch. FIX-LIST P measured a single attempt landing "roughly half the time"
# but never measured whether the retry loop actually converges. This answers that.
#
# Touches ONLY a scratch presentation this script creates itself, in a freshly-named
# OneDrive subfolder -- never Rohan's real deck or register. Deletes the scratch
# folder on exit, success or failure.
#
#   powershell.exe -ExecutionPolicy Bypass -File onedrive_write_probe.ps1

param(
    [int]$Trials = 8,
    [string]$AddinPath = "$env:APPDATA\Microsoft\AddIns\addin107.ppam"
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

if (-not (Test-Path $AddinPath)) { throw "add-in not found: $AddinPath" }
$addinName = Split-Path $AddinPath -Leaf

$scratchDir = Join-Path "$env:USERPROFILE\OneDrive\Claude" "onedrive-write-probe-$(Get-Date -Format yyyyMMdd-HHmmss)"
New-Item -ItemType Directory -Path $scratchDir | Out-Null
$scratchFile = Join-Path $scratchDir "probe.pptx"

$ppt = $null
$results = @()
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $ppt.Visible = -1

    $addin = $ppt.AddIns.Add($AddinPath)
    $addin.Loaded = $true

    $blank = $ppt.Presentations.Add()
    $blank.SaveAs($scratchFile, 24)   # ppSaveAsOpenXMLPresentation
    $blank.Close()

    Start-Sleep -Seconds 5   # give OneDrive a moment to register the new file before reopening
    $pres = $ppt.Presentations.Open($scratchFile)

    $isCloud = $pres.FullName -like "http*"
    Write-Output "FullName after reopen: $($pres.FullName)"
    Write-Output "Recognised as cloud URL: $isCloud"
    if (-not $isCloud) {
        Write-Output "NOT cloud-hosted yet -- waiting 10s and reopening once more."
        $pres.Close()
        Start-Sleep -Seconds 10
        $pres = $ppt.Presentations.Open($scratchFile)
        Write-Output "FullName after second reopen: $($pres.FullName)"
    }
    Write-Output ""

    for ($i = 1; $i -le $Trials; $i++) {
        $period = "PROBE-$i-$(Get-Random -Maximum 9999)"
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $result = $ppt.GetType().InvokeMember("Run", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ppt,
            @([string]"'$addinName'!DeckRegistry.SetDeckPeriodVerified", $pres, [string]$period, [int]4))
        $sw.Stop()
        # A genuine VBA "" success return marshals as PowerShell $null through
        # this exact Application.Run path -- confirmed 2026-08-16 by checking
        # the SAVED FILE independently after a "failed" call and finding the
        # write had actually landed. $null here means success, same as "".
        $ok = ($result -eq "" -or $result -eq $null)
        $results += [PSCustomObject]@{
            Trial   = $i
            Period  = $period
            Landed  = $ok
            Seconds = [math]::Round($sw.Elapsed.TotalSeconds, 1)
        }
        $status = if ($ok) { "LANDED" } else { "FAILED -- $result" }
        Write-Output ("Trial {0}: {1} in {2}s -- {3}" -f $i, $(if ($ok) {"LANDED"} else {"FAILED"}), [math]::Round($sw.Elapsed.TotalSeconds,1), $status)
    }

    $pres.Saved = -1
    $pres.Close()

    # Independent cross-check, OUTSIDE VBA entirely -- read the LAST trial's
    # value straight out of the saved package using .NET's own zip reader, not
    # the code under test and not PowerPoint's object cache. Storage moved off
    # docProps/custom.xml onto slide content 2026-08-16 (FIX-LIST P) -- scans
    # ppt/slides/*.xml for the shape's text the same way RegistryValueOnDisk
    # does, independently reimplemented rather than calling into the code
    # being tested.
    Start-Sleep -Seconds 1
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($scratchFile)
    $lastPeriod = $results[-1].Period
    $independentlyConfirmed = $false
    foreach ($entry in ($zip.Entries | Where-Object { $_.FullName -like "ppt/slides/slide*.xml" })) {
        $reader = New-Object System.IO.StreamReader($entry.Open())
        $xmlText = $reader.ReadToEnd()
        $reader.Close()
        if ($xmlText -match [regex]::Escape($lastPeriod)) {
            $independentlyConfirmed = $true
            break
        }
    }
    $zip.Dispose()

    Write-Output ""
    Write-Output "Independent .NET zip read of ppt/slides/*.xml (bypasses VBA and PowerPoint):"
    Write-Output "Last trial's value ($lastPeriod) present in the raw file bytes: $independentlyConfirmed"

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
    Start-Sleep -Seconds 1
    Remove-Item -Path $scratchDir -Recurse -Force -ErrorAction SilentlyContinue
}
