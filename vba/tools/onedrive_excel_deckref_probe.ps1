# Proves the FIXED WriteDeckReference/ReadDeckReference implementation
# (mirrored in ExcelPropertyProbe.bas as WriteDeckReferenceNew/
# ReadDeckReferenceNew) survives the exact scenario that broke the old
# CustomDocumentProperties mechanism: a re-pairing, where the SAME workbook
# gets stamped with a DIFFERENT deck reference after its first stamp.
# Verified by closing and reopening the workbook between writes -- two
# genuinely separate sessions, not just an in-memory read of the same open
# object, closer to how a real repoint actually happens.
#
#   powershell.exe -ExecutionPolicy Bypass -File onedrive_excel_deckref_probe.ps1

param(
    [string]$SourceDir = $PSScriptRoot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

foreach ($progId in @("Excel.Application")) {
    try {
        [System.Runtime.InteropServices.Marshal]::GetActiveObject($progId) | Out-Null
        Write-Output "=== ABORTING: $progId is already running -- close it first. ==="
        exit 2
    } catch {}
}

$modulePath = Join-Path $SourceDir "ExcelPropertyProbe.bas"
if (-not (Test-Path $modulePath)) { throw "missing module: $modulePath" }
$localModule = Join-Path $env:TEMP "ExcelPropertyProbe.bas"
Copy-Item -Path $modulePath -Destination $localModule -Force

$scratchDir = Join-Path "$env:USERPROFILE\OneDrive\Claude" "excel-deckref-probe-$(Get-Date -Format yyyyMMdd-HHmmss)"
New-Item -ItemType Directory -Path $scratchDir | Out-Null
$scratchFile = Join-Path $scratchDir "probe.xlsx"

$xl = $null
try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $true
    $xl.DisplayAlerts = $false

    $wb = $xl.Workbooks.Add()
    $wb.SaveAs($scratchFile, 51)
    $wb.Close($false)
    Start-Sleep -Seconds 5
    $wb = $xl.Workbooks.Open($scratchFile)
    Write-Output "Recognised as cloud URL: $($wb.FullName -like 'http*')"

    $wb.VBProject.VBComponents.Import($localModule) | Out-Null
    $flags = [System.Reflection.BindingFlags]::InvokeMethod

    # SESSION 1: first-ever stamp.
    $xl.GetType().InvokeMember("Run", $flags, $null, $xl,
        @([string]"ExcelPropertyProbe.WriteDeckReferenceNew", $wb, [string]"DECK-A-original")) | Out-Null
    $wb.Save()
    $wb.Saved = $true
    $wb.Close($false)
    Write-Output "Session 1: stamped DECK-A-original, saved, closed."

    Start-Sleep -Seconds 3
    $wb = $xl.Workbooks.Open($scratchFile)
    $wb.VBProject.VBComponents.Import($localModule) | Out-Null   # plain .xlsx has no VBA of its own, same as the real register -- re-import the DRIVER each reopen
    $readA = $xl.GetType().InvokeMember("Run", $flags, $null, $xl,
        @([string]"ExcelPropertyProbe.ReadDeckReferenceNew", $wb))
    Write-Output "Session 2 (reopened): reads back [$readA]  MATCHES=$($readA -eq 'DECK-A-original')"

    # SESSION 2 (same reopen): RE-PAIR to a DIFFERENT deck -- the exact
    # scenario that broke the old mechanism.
    $xl.GetType().InvokeMember("Run", $flags, $null, $xl,
        @([string]"ExcelPropertyProbe.WriteDeckReferenceNew", $wb, [string]"DECK-B-repointed")) | Out-Null
    $wb.Save()
    $wb.Saved = $true
    $wb.Close($false)
    Write-Output "Session 2: re-stamped DECK-B-repointed, saved, closed."

    Start-Sleep -Seconds 3
    $wb = $xl.Workbooks.Open($scratchFile)
    $wb.VBProject.VBComponents.Import($localModule) | Out-Null
    $readB = $xl.GetType().InvokeMember("Run", $flags, $null, $xl,
        @([string]"ExcelPropertyProbe.ReadDeckReferenceNew", $wb))
    Write-Output "Session 3 (reopened): reads back [$readB]  MATCHES-NEW=$($readB -eq 'DECK-B-repointed')  STILL-STUCK-ON-OLD=$($readB -eq 'DECK-A-original')"

    $wb.Saved = $true
    $wb.Close($false)
} catch {
    Write-Output "=== DRIVER ERROR ==="
    Write-Output $_.Exception.ToString()
    exit 3
} finally {
    if ($xl) {
        try { foreach ($w in $xl.Workbooks) { $w.Saved = $true } } catch {}
        try { $xl.Quit() } catch {}
    }
    Get-Process EXCEL -ErrorAction SilentlyContinue |
        Where-Object { -not $_.MainWindowTitle } |
        Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 1
    Remove-Item -Path $scratchDir -Recurse -Force -ErrorAction SilentlyContinue
}
