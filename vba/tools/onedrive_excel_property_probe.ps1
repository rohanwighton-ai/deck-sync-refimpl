# Checks whether Excel's CustomDocumentProperties on a OneDrive-hosted
# workbook has the same "only the first write per session lands" limit
# PowerPoint showed (FIX-LIST P). ExcelOutput.WriteDeckReference/
# ReadDeckReference use the exact same mechanism, unverified until now.
#
# Driven via VBA (ExcelPropertyProbe.bas), not direct PowerShell/.NET COM
# calls on CustomDocumentProperties -- that path throws a NullReferenceException
# deep in PowerShell's own COM interop (confirmed live 2026-08-16, twice, two
# different call shapes) trying to introspect that COM object's type info. A
# limitation of .NET's dynamic COM binder, not a fact about Excel or OneDrive.
#
#   powershell.exe -ExecutionPolicy Bypass -File onedrive_excel_property_probe.ps1

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

$scratchDir = Join-Path "$env:USERPROFILE\OneDrive\Claude" "excel-property-probe-$(Get-Date -Format yyyyMMdd-HHmmss)"
New-Item -ItemType Directory -Path $scratchDir | Out-Null
$scratchFile = Join-Path $scratchDir "probe.xlsx"

$xl = $null
try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $true
    $xl.DisplayAlerts = $false

    $wb = $xl.Workbooks.Add()
    $wb.SaveAs($scratchFile, 51)   # xlOpenXMLWorkbook
    $wb.Close($false)
    Start-Sleep -Seconds 5
    $wb = $xl.Workbooks.Open($scratchFile)
    Write-Output "FullName after reopen: $($wb.FullName)"
    Write-Output "Recognised as cloud URL: $($wb.FullName -like 'http*')"
    Write-Output ""

    $wb.VBProject.VBComponents.Import($localModule) | Out-Null

    $flags = [System.Reflection.BindingFlags]::InvokeMethod

    # TRIAL 1: brand-new property, Add (first-ever write to this file).
    $xl.GetType().InvokeMember("Run", $flags, $null, $xl,
        @([string]"ExcelPropertyProbe.WriteLikeTheRealFunction", $wb, [string]"RegProbeA", [string]"VALUE-A1")) | Out-Null
    $wb.Save()
    Start-Sleep -Seconds 1
    $onDiskA1 = $xl.GetType().InvokeMember("Run", $flags, $null, $xl,
        @([string]"ExcelPropertyProbe.PropertyOnDiskXlsx", [string]$scratchFile, [string]"RegProbeA"))
    Write-Output "Trial 1 -- new prop 'RegProbeA'=VALUE-A1, first write: on disk=[$onDiskA1] LANDED=$($onDiskA1 -eq 'VALUE-A1')"

    # TRIAL 2: a SECOND, never-before-used property, same session.
    $xl.GetType().InvokeMember("Run", $flags, $null, $xl,
        @([string]"ExcelPropertyProbe.WriteLikeTheRealFunction", $wb, [string]"RegProbeB", [string]"VALUE-B1")) | Out-Null
    $wb.Save()
    Start-Sleep -Seconds 1
    $onDiskB1 = $xl.GetType().InvokeMember("Run", $flags, $null, $xl,
        @([string]"ExcelPropertyProbe.PropertyOnDiskXlsx", [string]$scratchFile, [string]"RegProbeB"))
    Write-Output "Trial 2 -- new prop 'RegProbeB'=VALUE-B1, same session: on disk=[$onDiskB1] LANDED=$($onDiskB1 -eq 'VALUE-B1')"

    # TRIAL 3: UPDATE RegProbeA via the exact WriteDeckReference pattern
    # (.Value= on an existing property).
    $xl.GetType().InvokeMember("Run", $flags, $null, $xl,
        @([string]"ExcelPropertyProbe.WriteLikeTheRealFunction", $wb, [string]"RegProbeA", [string]"VALUE-A2-UPDATED")) | Out-Null
    $wb.Save()
    Start-Sleep -Seconds 1
    $onDiskA2 = $xl.GetType().InvokeMember("Run", $flags, $null, $xl,
        @([string]"ExcelPropertyProbe.PropertyOnDiskXlsx", [string]$scratchFile, [string]"RegProbeA"))
    Write-Output "Trial 3 -- UPDATE 'RegProbeA' via .Value= (real WriteDeckReference pattern): on disk=[$onDiskA2] LANDED=$($onDiskA2 -eq 'VALUE-A2-UPDATED')"

    # TRIAL 4: UPDATE RegProbeA via Delete+Add instead.
    $xl.GetType().InvokeMember("Run", $flags, $null, $xl,
        @([string]"ExcelPropertyProbe.WriteViaDeleteAndAdd", $wb, [string]"RegProbeA", [string]"VALUE-A3-DELETEADD")) | Out-Null
    $wb.Save()
    Start-Sleep -Seconds 1
    $onDiskA3 = $xl.GetType().InvokeMember("Run", $flags, $null, $xl,
        @([string]"ExcelPropertyProbe.PropertyOnDiskXlsx", [string]$scratchFile, [string]"RegProbeA"))
    Write-Output "Trial 4 -- UPDATE 'RegProbeA' via Delete+Add: on disk=[$onDiskA3] LANDED=$($onDiskA3 -eq 'VALUE-A3-DELETEADD')"

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
