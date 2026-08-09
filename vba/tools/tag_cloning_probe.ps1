# Driver for TagCloningProbe.bas.
#
# The probe existed since 31 July and had NO RUNNER and no recorded result --
# written, never run, referenced nowhere. That is the same shape as the two
# unreachable injectors it exists to answer for.
#
# It touches no deck and no register: it builds its own slide in a scratch
# presentation, duplicates it, and reports what survived. Safe to re-run.
#
# Copy this file and the two .bas files it names into a Windows-local folder
# and run it from there -- VBComponents.Import does not reliably read a UNC or
# WSL path.
#
#   powershell.exe -ExecutionPolicy Bypass -File tag_cloning_probe.ps1
#
# The verdict is the "=== VERDICTS ===" block. Anything else -- including a
# clean exit -- is not an answer.

param(
    [string]$SourceDir = $PSScriptRoot,
    [string]$OutFile = (Join-Path $env:TEMP "tag_cloning_probe.txt")
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# A probe that attaches to an Office instance someone else is using would both
# disturb their work and read state this probe did not set.
foreach ($progId in @("PowerPoint.Application")) {
    try {
        [System.Runtime.InteropServices.Marshal]::GetActiveObject($progId) | Out-Null
        Write-Output "=== ABORTING: $progId is already running -- close it first. ==="
        exit 2
    } catch {}
}

$modules = @("InjectPrimitive.bas", "TagCloningProbe.bas")
foreach ($m in $modules) {
    $p = Join-Path $SourceDir $m
    if (-not (Test-Path $p)) { throw "missing module: $p" }
}

$ppt = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $ppt.Visible = -1
    $scratch = $ppt.Presentations.Add()
    foreach ($m in $modules) {
        $scratch.VBProject.VBComponents.Import((Join-Path $SourceDir $m)) | Out-Null
    }

    $report = $ppt.GetType().InvokeMember("Run", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ppt,
        @([string]"TagCloningProbe.TagCloningProbe"))

    Set-Content -Path $OutFile -Value $report
    Write-Output $report
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
}
