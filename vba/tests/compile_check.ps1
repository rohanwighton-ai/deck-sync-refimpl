# Force a WHOLE-PROJECT VBA compile and report whether it succeeded.
#
# WHY THIS EXISTS, demonstrated 2026-08-05 rather than assumed: the suite
# reported "152 passed, 0 failed", exit 0, with a genuine compile error sitting
# in ExcelOutput.bas (three UpsertRow calls left on the old 3-argument signature
# after b60de0b made the period required). The error was reintroduced on purpose
# and the suite passed anyway. The VBE dialog it had been hiding read
# "Compile error: Argument not optional" at ExcelOutput line 461.
#
# THE MECHANISM: VBA COMPILES PER-PROCEDURE, ON DEMAND -- not per module, and
# not per project. ManualSmokeTest is called by nothing, so its body was never
# compiled, so its error was never seen, while every other procedure in the same
# module compiled and passed. The suite's "NO TESTS RAN" guard only fires when
# the broken code is on the call path to RunAllTests. Code nothing calls can be
# broken indefinitely and every test still passes.
#
# So a probe call is NOT a compile check -- it only compiles what it reaches.
# Debug > Compile VBAProject is the only thing that compiles everything.
#
# TWO DOCUMENTED CAVEATS, both load-bearing:
#   1. Execute() RETURNS NOTHING. It runs the compiler and reports neither
#      success nor failure, so trusting it would be an always-true guard -- the
#      exact shape this project has a zettel about. What IS observable: the menu
#      control is DISABLED once a project is fully compiled. So the check is
#      "was it enabled before, and is it disabled after".
#   2. A COMPILE ERROR OPENS A MODAL. Measured behaviour: the call does not
#      hang, it comes back as RPC_E_CALL_REJECTED within seconds, because a
#      modal-blocked Office app refuses every subsequent COM call. The caller
#      still applies a timeout as a backstop for the case where it does block.
#
# NOTHING IS EVER FORCE-KILLED (Rohan's call, 2026-08-05). On failure PowerPoint
# is left on screen with the VBE dialog open -- that dialog names the offending
# line, which is the most useful diagnostic available. The suite refuses to
# start while Office is running, so a left-open modal also blocks the NEXT run.
# That is the point: somebody has to look.
#
# THE VERDICT IS A LINE OF OUTPUT, NOT AN EXIT CODE. Start-Process -PassThru
# does not retain the process handle, so .ExitCode reads back empty in the
# parent no matter how it waits -- the gate printed "COMPILE GATE FAILED
# (exit )" for exactly that reason. The last line here is RESULT: OK or
# RESULT: FAIL, and the caller requires RESULT: OK to proceed. Absence of the
# line for ANY reason -- crash, hang, timeout -- is a failure, so the gate fails
# closed rather than on the strength of a value that may never arrive.

param(
    [Parameter(Mandatory = $true)][string]$Staging,
    # Exactly the modules the PowerPoint pass imports, comma-separated, handed
    # down by the caller.
    #
    # THIS USED TO IMPORT EVERY .bas IN STAGING, to avoid maintaining a second
    # list. That invented a compile error on its first honest run: staging also
    # holds the EXCEL pass's modules, and TestRunnerExcel.NewBlankSheet calls
    # Application.ActiveWorkbook -- which does not exist on PowerPoint's
    # Application. Reported as "the project does not compile" against code that
    # was perfectly fine, which is worse than the blindness it replaced. The two
    # module sets are two projects and were never meant to compile together.
    #
    # It is not a second list: the caller already computes this one to do its
    # own importing, and passes the same value here.
    [Parameter(Mandatory = $true)][string]$Modules
)

$ErrorActionPreference = "Stop"

function Fail-Compile {
    param([string[]]$Lines)
    foreach ($l in $Lines) { Write-Output $l }
    Write-Output "RESULT: FAIL"
    exit 1
}

# Never attach to an Office instance somebody is using, and never take one over.
foreach ($progId in @("PowerPoint.Application")) {
    try {
        [System.Runtime.InteropServices.Marshal]::GetActiveObject($progId) | Out-Null
        Fail-Compile @("COMPILE CHECK SKIPPED: $progId is already running.")
    } catch [System.Management.Automation.MethodInvocationException] {
        # GetActiveObject throws when nothing is running -- that is the good case.
    }
}

$ppt = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $ppt.Visible = -1
    $pres = $ppt.Presentations.Add()

    # EXACTLY the caller's PowerPoint set -- see the $Modules note above for why
    # "every .bas in staging" was wrong rather than merely loose.
    $names = $Modules -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
    if ($names.Count -eq 0) { Fail-Compile @("COMPILE CHECK FAILED: no modules given") }
    foreach ($n in $names) {
        $p = Join-Path $Staging $n
        if (-not (Test-Path $p)) { Fail-Compile @("COMPILE CHECK FAILED: $n is not in $Staging") }
        $pres.VBProject.VBComponents.Import($p) | Out-Null
    }
    Write-Output "modules imported: $($names.Count)"

    $vbe = $ppt.VBE
    $win = $vbe.MainWindow
    $wasVisible = $win.Visible
    $win.Visible = $true

    $ctl = $vbe.CommandBars.Item("Menu Bar").Controls.Item("Debug").Controls.Item("Compile VBAProject")

    # A freshly imported project must NOT already be compiled. If it is, this
    # check cannot tell a clean compile from a no-op and must say so rather than
    # report a pass it did not earn.
    if (-not $ctl.Enabled) {
        $win.Visible = $wasVisible
        Fail-Compile @("COMPILE CHECK INCONCLUSIVE: 'Compile VBAProject' was already disabled before compiling.")
    }

    $ctl.Execute()

    $stillEnabled = $ctl.Enabled
    $win.Visible = $wasVisible

    if ($stillEnabled) {
        Fail-Compile @(
            "COMPILE FAILED: the project did not compile clean.",
            "  'Compile VBAProject' is still enabled after compiling, which means",
            "  compilation did not complete. Open the project and use",
            "  Debug > Compile VBAProject -- its dialog names the offending line."
        )
    }

    Write-Output "COMPILE OK: whole project compiled clean ($($names.Count) modules)."
    Write-Output "RESULT: OK"
    exit 0
} catch {
    $msg = $_.Exception.Message

    # RPC_E_CALL_REJECTED IS THE COMPILE FAILURE, not an interop glitch.
    #
    # Observed 2026-08-05 against a known-bad project: the compile does not
    # hang. Debug > Compile opens its error dialog, and a modal-blocked Office
    # app then REFUSES every subsequent COM call with 0x80010001. The rejection
    # arrives in seconds and is a better signal than any timeout.
    #
    # This exact HRESULT was recorded in field_e2e.ps1 on 2026-08-01, where it
    # was read as the VBE being unreliable and force-compiling was abandoned
    # because of it. It was not flakiness. It was the compiler reporting an
    # error the only way it can through automation -- and reading it as noise
    # is what left this suite blind for four days.
    if ($msg -match "0x80010001" -or $msg -match "rejected by callee") {
        Fail-Compile @(
            "COMPILE FAILED: the project does not compile.",
            "  PowerPoint rejected the call (RPC_E_CALL_REJECTED). That is how a",
            "  modal announces itself through COM: Debug > Compile VBAProject opened",
            "  its error dialog, and a modal-blocked app refuses everything after it.",
            "  PowerPoint has been LEFT OPEN -- the dialog names the offending line."
        )
    }

    Fail-Compile @("COMPILE CHECK ERROR: $msg")
} finally {
    if ($ppt) { try { $ppt.Quit() } catch {} }
}
