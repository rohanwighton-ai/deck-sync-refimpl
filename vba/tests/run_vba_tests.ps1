# First real-execution driver for deck-sync-refimpl's VBA port.
#
# Every module up to this point was "not executed or verified in this
# environment" per its own SPIKE_NOTES_*.md -- there was no Windows/Office
# install available. This script closes that gap: it drives real PowerPoint
# and Excel instances via COM automation, imports the .bas modules
# (production + the assertion-based TestRunner*.bas), runs the tests, and
# prints one machine-readable report per host app to stdout.
#
# Deliberately NOT interactive: Office is launched, driven, and quit with
# no human input required, so this can be re-run repeatedly as part of a
# hardening loop (run -> read real failures -> fix vba/*.bas -> re-run).
#
# Everything is staged into a Windows-native temp folder first (not run
# directly against the WSL \\wsl.localhost\... path) -- COM automation
# across that boundary is a known source of flakiness; plain file copies
# across it are not, so staging is the only thing that crosses it.

param(
    [string]$RepoRoot = $(Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
)

$ErrorActionPreference = "Stop"

# A prior interrupted run (or you just having Office open) can leave a live
# POWERPNT.EXE/EXCEL.EXE process behind. New-Object -ComObject would attach
# to that instance instead of spawning a fresh one, and a stale instance's
# VBProject/VBComponents access can come back genuinely null even with Trust
# access enabled (found the hard way on this script's first real run -- not
# a hypothetical). So any existing instance must be closed first.
#
# This script is now also invoked unattended after autonomous commits, so
# force-killing (the original approach) is no longer safe -- it would
# silently discard unsaved work in a document you happen to have open at the
# time. Instead: ask each running instance to Quit() normally. If it has
# unsaved changes, Office shows its own native "Save changes?" dialog and
# Quit() blocks until you respond. That's bounded by a timeout so an
# unattended run doesn't hang forever -- if you don't respond in time, this
# run is skipped (exit 2) rather than forcing the issue.
#
# --- The husk problem, and why the discriminator changed (2026-07-30) -------
# This suite LEAKS an Excel process on most runs: several tests do
# CreateObject("Excel.Application") / wb.Close False / xl.Quit, and an
# outstanding COM reference stops the process actually terminating. The
# leftover is an empty Excel holding no workbook -- and it survives a clean
# Quit(), so the $stillRunning check below rejected it and every SUBSEQUENT
# run exited 2 with no output until a human killed it by hand. That cost two
# 8-minute runs in one evening, and the note about it had already been
# written down before it happened again.
#
# The obvious fix -- "force-kill windowless instances" -- does not work. The
# same leftover process reported MainWindowTitle "Excel" at 19:56 and blank
# at 20:12, zero workbooks throughout: window presence flips on its own and
# is not evidence of anything.
#
# What does not flip is whether the instance holds a DOCUMENT. Zero documents
# means there is nothing a human can lose, which is the only question that
# actually matters here. So: count documents first, and force-kill only a
# confirmed-empty instance.
#
# Two asymmetries make that safe, and both are load-bearing:
#
#   1. UNKNOWN NEVER PERMITS A KILL. Attaching or counting can fail or hang
#      (a modal dialog blocks COM), and every such case returns -1, which
#      falls through to the graceful path. Only a confirmed 0 kills.
#   2. MORE THAN ONE PROCESS NEVER PERMITS A KILL. GetActiveObject attaches
#      to whichever instance registered in the Running Object Table -- with
#      two EXCEL.EXE running it may well be the husk, so "0 documents" would
#      say nothing about the other one. Killing by name on that basis is
#      precisely how you would discard the open workbook this comment block
#      exists to protect. So the kill requires exactly one process of that
#      name, and targets it by PID, not by name.
#
# Net effect: the leaked husk gets cleared automatically, a real session of
# yours is still asked politely and still aborts the run if it will not go,
# and anything ambiguous is treated as a real session.

# Documents open in the running instance of $ProgId.
#   0  = confirmed empty -- an automation husk, nothing to lose
#   >0 = a real session with documents in it
#   -1 = unknown (no instance, attach failed, or the call hung) -> treat as real
#
# Runs in a job for the same reason Request-GracefulQuit does: a modal dialog
# in Office blocks COM calls indefinitely, and a pre-flight check that can
# hang forever is worse than the problem it solves.
function Get-OpenDocumentCount {
    param([string]$ProgId, [int]$TimeoutSeconds = 30)

    $job = Start-Job -ScriptBlock {
        param($ProgId)
        try {
            $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject($ProgId)
        }
        catch {
            return -1
        }
        try {
            if ($ProgId -eq "Excel.Application") { return [int]$app.Workbooks.Count }
            else { return [int]$app.Presentations.Count }
        }
        catch {
            return -1
        }
    } -ArgumentList $ProgId

    $completed = Wait-Job $job -Timeout $TimeoutSeconds
    if (-not $completed) {
        Stop-Job $job -ErrorAction SilentlyContinue
        Remove-Job $job -Force -ErrorAction SilentlyContinue
        return -1
    }
    $result = Receive-Job $job
    Remove-Job $job -ErrorAction SilentlyContinue
    if ($null -eq $result) { return -1 }
    return [int]$result
}

# Clears a confirmed-empty leftover.
#
# Returns NOTHING, deliberately. The first version returned $true/$false so the
# caller could tell "handled" from "not applicable" -- but in PowerShell a
# Write-Output inside a function is emitted onto the same pipeline as the
# return value, so the caller's `| Out-Null` (added to discard the boolean)
# also swallowed the "Clearing a leftover..." message. The script would then
# force-kill a process and leave no trace of it in the output, which is the
# single behaviour an unattended script must never have. Caught by running the
# function for real rather than reading it, 2026-07-30.
#
# Since no caller ever used the boolean, dropping it removes the conflict
# outright instead of working around it. Bare `return` exits without emitting.
function Clear-EmptyHusk {
    param([string]$ProgId, [string]$ProcessName)

    $procs = @(Get-Process $ProcessName -ErrorAction SilentlyContinue)
    if ($procs.Count -ne 1) { return }   # asymmetry 2: ambiguity is not a kill

    $docs = Get-OpenDocumentCount -ProgId $ProgId
    if ($docs -ne 0) { return }          # asymmetry 1: unknown (-1) is not a kill

    Write-Output "=== Clearing a leftover empty $ProcessName (pid $($procs[0].Id), 0 documents open) -- an automation husk from a previous run. ==="
    Stop-Process -Id $procs[0].Id -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
}

function Request-GracefulQuit {
    param([string]$ProgId, [int]$TimeoutSeconds = 180)

    $job = Start-Job -ScriptBlock {
        param($ProgId)
        try {
            $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject($ProgId)
        }
        catch {
            return "no-instance"
        }
        try {
            $app.Quit()
            return "quit-ok"
        }
        catch {
            return "quit-error: $($_.Exception.Message)"
        }
    } -ArgumentList $ProgId

    $completed = Wait-Job $job -Timeout $TimeoutSeconds
    if (-not $completed) {
        Stop-Job $job -ErrorAction SilentlyContinue
        Remove-Job $job -Force -ErrorAction SilentlyContinue
        return "timeout"
    }
    $result = Receive-Job $job
    Remove-Job $job -ErrorAction SilentlyContinue
    return $result
}

$processNameFor = @{ "PowerPoint.Application" = "POWERPNT"; "Excel.Application" = "EXCEL" }
foreach ($progId in @("PowerPoint.Application", "Excel.Application")) {
    # Husk first: an empty leftover cannot be cleared by Quit() (that is the
    # whole failure), so asking politely and then rejecting it is a dead end.
    # If this clears something, the checks below simply find nothing running.
    # No `| Out-Null` here on purpose -- see Clear-EmptyHusk's header. Piping
    # its output away is what would silence the record of a force-kill.
    Clear-EmptyHusk -ProgId $progId -ProcessName $processNameFor[$progId]

    $outcome = Request-GracefulQuit -ProgId $progId
    if ($outcome -eq "timeout" -or $outcome -like "quit-error*") {
        Write-Output "=== SKIPPED: $progId did not close cleanly ($outcome) -- likely an unresolved save prompt or in active use. Not forcing it closed. ==="
        exit 2
    }
    # Give Office a moment to actually release the process after Quit()
    Start-Sleep -Seconds 1
    $stillRunning = Get-Process $processNameFor[$progId] -ErrorAction SilentlyContinue
    if ($stillRunning) {
        Write-Output "=== SKIPPED: $($processNameFor[$progId]) still running after a clean Quit() -- not forcing it closed. ==="
        exit 2
    }
}

$staging = Join-Path $env:TEMP ("deck-sync-vba-test-" + [guid]::NewGuid().ToString("N").Substring(0, 8))
$fixturesStaging = Join-Path $staging "fixtures\"
New-Item -ItemType Directory -Path $staging | Out-Null
New-Item -ItemType Directory -Path $fixturesStaging | Out-Null
$staging = (Get-Item $staging).FullName + "\"

Write-Output "Staging directory: $staging"

$vbaSourceDir = Join-Path $RepoRoot "vba"
$fixturesSourceDir = Join-Path $RepoRoot "test-fixtures"
Write-Output "RepoRoot: $RepoRoot"
Write-Output "vbaSourceDir: $vbaSourceDir"
Write-Output "fixturesSourceDir: $fixturesSourceDir"

$pptModules = @("Discovery.bas", "InjectPrimitive.bas", "Matching.bas", "Resolve.bas", "SyncOperations.bas", "Onboarding.bas", "Verification.bas", "SlideDuplication.bas", "TemplateSlide.bas", "TemplateAudit.bas", "IdentityCheck.bas", "TagMigration.bas", "PlaceholderCheck.bas", "RunSync.bas", "ReviewQueue.bas", "Readiness.bas","FieldWiring.bas","MilestoneDevice.bas","Drafting.bas","FieldSpec.bas","Sources.bas","DraftingUI.bas","DiscoverUI.bas", "DeckAdoption.bas", "ResolveFields.bas", "DeckRegistry.bas", "WorkbookBridge.bas", "RibbonUI.bas", "AdoptFlow.bas", "BatchOnboardFlow.bas", "CommandBarUI.bas", "tests\TestRunner.bas")
foreach ($m in $pptModules) {
    Copy-Item (Join-Path $vbaSourceDir $m) -Destination $staging
}
$excelModules = @("ExcelOutput.bas", "tests\TestRunnerExcel.bas")
foreach ($m in $excelModules) {
    Copy-Item (Join-Path $vbaSourceDir $m) -Destination $staging
}
Copy-Item (Join-Path $fixturesSourceDir "shp-groupshape.pptx") -Destination $fixturesStaging

# --- STATIC CHECKS, ahead of the compile gate ------------------------
#
# check_vba_static.py finds in a second what the compile gate finds in minutes
# and what Office attributes to the wrong module entirely. AGENTS.md said to run
# it before every suite run -- which made it a thing to REMEMBER, and it was
# forgotten often enough that a reserved-word parameter sat in TestRunner.bas
# undetected until 2026-08-09.
#
# It also prints reachability notes: capabilities with no toolbar button. Those
# are REPORTED and do not stop the run -- the enforcing half is the declared
# roster in Test_CommandBarUI_EveryDeclaredCapabilityHasAButton, because whether
# an unlisted Public proc should be a capability is a judgement, and a gate that
# is red for a judgement is a gate people learn to bypass.
#
# Skipped, loudly, if python is absent -- this must never silently not run.
# THE REPO LIVES IN WSL AND THE INTERPRETER LIVES THERE WITH IT.
#
# First attempt used Get-Command python on Windows. It "found" one -- the
# Microsoft Store execution-alias shim, which resolves happily and then prints
# "Python was not found" and exits non-zero. A check that locates a thing which
# cannot run is worse than finding nothing, because the caller believes it ran.
#
# RepoRoot arrives as \\wsl.localhost\Ubuntu\home\... when invoked from WSL, so
# the Linux path is recoverable from it. Verified by RUNNING the interpreter,
# not by resolving its name.
$staticOk = $false
$staticRan = $false
if ($RepoRoot -match '^\\\\wsl\.localhost\\[^\\]+\\(.+)$') {
    $linuxRoot = '/' + ($Matches[1] -replace '\\', '/')
    $linuxScript = "$linuxRoot/vba/tests/check_vba_static.py"
    Write-Output "--- static checks ---"
    & wsl.exe -- python3 $linuxScript
    if ($LASTEXITCODE -eq 0) { $staticOk = $true }
    $staticRan = $true
}

if (-not $staticRan) {
    # Windows-hosted checkout: use a real interpreter, proven by running it.
    foreach ($cand in @('python3', 'python')) {
        $c = Get-Command $cand -ErrorAction SilentlyContinue
        if (-not $c) { continue }
        & $c.Source --version > $null 2>&1
        if ($LASTEXITCODE -ne 0) { continue }        # the Store shim lands here
        Write-Output "--- static checks ---"
        & $c.Source (Join-Path $RepoRoot "vba\tests\check_vba_static.py")
        if ($LASTEXITCODE -eq 0) { $staticOk = $true }
        $staticRan = $true
        break
    }
}

if (-not $staticRan) {
    Write-Output "=== STATIC CHECKS SKIPPED: no working python found. THIS IS NOT A PASS. ==="
} elseif (-not $staticOk) {
    Write-Output "=== STATIC CHECKS FAILED. NO TESTS WERE RUN. ==="
    exit 4
}

# --- COMPILE GATE, before a single test runs -------------------------
#
# Added 2026-08-05 after this suite was PROVEN blind to a real compile error:
# the 3-argument UpsertRow calls in ExcelOutput.ManualSmokeTest were
# reintroduced on purpose and this script still printed "152 passed, 0 failed"
# and exited 0. VBA compiles PER PROCEDURE, ON DEMAND -- ManualSmokeTest is
# called by nothing, so it was never compiled and its error was never seen.
# The "NO TESTS RAN" guard below only catches a compile error that happens to
# sit on the call path to RunAllTests.
#
# A CHILD PROCESS, not an inline call, because a compile error opens a modal in
# the VBE and the Execute() that raised it never returns. A blocking COM call
# cannot be timed out from the thread that made it -- so the timeout goes on
# the process instead, and a hang becomes the failure signal rather than a
# mystery. (It cost an evening as a mystery on 2026-07-29.)
#
# NOTHING IS KILLED ON TIMEOUT. Rohan's call: the stuck PowerPoint keeps its
# VBE dialog on screen, and that dialog names the offending line. This script
# refuses to start while Office is running, so a left-open modal also blocks
# the next run -- which is the intended pressure to go and look at it.
$compileScript = Join-Path $staging "compile_check.ps1"
Copy-Item (Join-Path $RepoRoot "vba\tests\compile_check.ps1") -Destination $compileScript
$compileLog = Join-Path $staging "compile_check.log"
$compileTimeoutSeconds = 180

# THE POWERPOINT SET, and only it. Computed here rather than inside the
# PowerPoint pass below, because the gate has to import exactly what that pass
# imports -- the same value now feeds both.
#
# It matters that this is not "all the .bas files in staging": staging also
# holds the EXCEL pass's TestRunnerExcel.bas, whose NewBlankSheet calls
# Application.ActiveWorkbook. That is a compile error in a PowerPoint project
# and has nothing to do with the code under test. The first version of this
# gate did exactly that and reported a healthy project as broken.
$pptImports = @("ExcelOutput.bas") + @($pptModules | ForEach-Object { Split-Path $_ -Leaf })

$compileProc = Start-Process -FilePath "powershell.exe" -PassThru -WindowStyle Hidden `
    -RedirectStandardOutput $compileLog `
    -ArgumentList @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $compileScript, `
                    "-Staging", $staging, "-Modules", ($pptImports -join ","))

if (-not $compileProc.WaitForExit($compileTimeoutSeconds * 1000)) {
    Write-Output ""
    Write-Output "=== COMPILE CHECK DID NOT RETURN after $compileTimeoutSeconds seconds. ==="
    Write-Output "This is a FAILURE, not a slow machine. Debug > Compile VBAProject blocks on"
    Write-Output "a modal dialog when the project does not compile, and that is the only way"
    Write-Output "a compile failure can announce itself."
    Write-Output ""
    Write-Output "PowerPoint has been LEFT RUNNING on purpose. Switch to it: the VBE dialog"
    Write-Output "names the offending line. Close it, fix the line, run this again."
    Write-Output ""
    Write-Output "NO TESTS WERE RUN."
    exit 3
}

$compileOutput = if (Test-Path $compileLog) { (Get-Content $compileLog -Raw) } else { "" }
Write-Output $compileOutput

# THE VERDICT IS A LINE, NOT AN EXIT CODE, and the test is for the PASS marker
# rather than against a failure one.
#
# $compileProc.ExitCode reads back EMPTY here however this waits on it --
# Start-Process -PassThru does not retain the process handle. The first version
# tested `-ne 0` against that empty value, which happened to be truthy and so
# happened to fail closed; it printed "COMPILE GATE FAILED (exit )" and would
# have blocked every run forever, including good ones. Right outcome, no
# reasoning behind it.
#
# Requiring RESULT: OK means a crash, a hang, a killed child, an empty log or a
# child that never got as far as printing anything all land on "did not
# compile" -- which is the safe direction for a gate standing in front of a
# test suite that is otherwise willing to report 152 passed on a dead project.
if ($compileOutput -notmatch '(?m)^RESULT: OK\s*$') {
    Write-Output ""
    Write-Output "=== COMPILE GATE FAILED. NO TESTS WERE RUN. ==="
    Write-Output "Every test below this point would have been meaningless: a project that does"
    Write-Output "not compile can still pass every test whose code path avoids the bad line."
    exit 3
}

$pptReport = ""
$excelReport = ""
$pptError = $null
$excelError = $null

# --- PowerPoint pass -------------------------------------------------
$ppt = $null
$pres = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $ppt.Visible = -1  # msoTrue -- visible on purpose for a first real run, see script header
    $pres = $ppt.Presentations.Add()

    # Derived from $pptModules rather than repeated. These two lists were
    # separate literals until 2026-07-31, when adding ReviewQueue.bas to one and
    # not the other produced "Application.Run : Invalid request. Sub or function
    # not defined." on TestRunner.RunAllTests -- which reads as a compile error
    # in the new code, not as a missing import, and sent the diagnosis straight
    # past the actual cause. Exactly the failure AGENTS.md already records for a
    # missing ExcelOutput.bas import; a second literal is a second place to
    # forget. ExcelOutput.bas is staged by $excelModules but PowerPoint needs it
    # too (RunSync calls into it), so it is added here explicitly. Import order
    # does not affect compilation -- every module is in the project before
    # Application.Run is called.
    # $pptImports is computed above, before the compile gate, so the gate and
    # this pass cannot import different sets.
    foreach ($m in $pptImports) {
        $comp = $pres.VBProject.VBComponents.Import((Join-Path $staging $m))
        Write-Output ("Imported $m as component name: " + $comp.Name)
    }

    # $ppt.Run("...", $a, $b) directly fails PowerShell's COM overload
    # resolution for Application.Run's VARIANT-optional-parameter signature
    # ("Cannot find an overload for 'Run' and the argument count: 3") --
    # a well-known PowerShell/COM interop gap, not a VBA-side issue.
    # InvokeMember via reflection is the standard workaround: it bypasses
    # PowerShell's own (buggy, for this case) method-binding logic.
    $fixturesArg = [string]$fixturesStaging
    $stagingArg = [string]$staging
    $pptReport = $ppt.GetType().InvokeMember("Run", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ppt, @([string]"TestRunner.RunAllTests", $fixturesArg, $stagingArg))
}
catch {
    $pptError = $_.Exception.Message + " [line " + $_.InvocationInfo.ScriptLineNumber + ": " + $_.InvocationInfo.Line.Trim() + "]"
}
finally {
    if ($pres) { $pres.Saved = $true; $pres.Close() }
    if ($ppt) { $ppt.Quit() }
    # Same COM-release gotcha as the Excel pass below (see that finally
    # block's comment) -- Quit() alone is not reliable here either,
    # confirmed 2026-07-26 (a windowless POWERPNT.EXE survived a clean,
    # non-erroring run until this was added).
    if ($pres) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($pres) | Out-Null }
    if ($ppt) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null }
    $pres = $null; $ppt = $null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    [System.GC]::Collect()
    Start-Sleep -Milliseconds 500
    Get-Process POWERPNT -ErrorAction SilentlyContinue | Where-Object { -not $_.MainWindowTitle } | Stop-Process -Force -ErrorAction SilentlyContinue
}

# --- Excel pass --------------------------------------------------------
$xl = $null
$wb = $null
try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $true
    $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Add()

    foreach ($m in @("ExcelOutput.bas", "TestRunnerExcel.bas")) {
        $wb.VBProject.VBComponents.Import((Join-Path $staging $m)) | Out-Null
    }

    $excelReport = $xl.Run("TestRunnerExcel.RunAllTests")
}
catch {
    $excelError = $_.Exception.Message + " [line " + $_.InvocationInfo.ScriptLineNumber + ": " + $_.InvocationInfo.Line.Trim() + "]"
}
finally {
    if ($wb) { $wb.Saved = $true; $wb.Close() }
    if ($xl) { $xl.Quit() }
    # $xl.Quit() alone reliably leaves the EXCEL.EXE process running
    # (confirmed 2026-07-26, unrelated to the AutoSave/BatchOnboardFlow
    # investigation that day: several sequential no-window EXCEL.EXE
    # processes were still alive well after a completed, non-erroring test
    # run and had to be force-killed by hand) -- releasing the RCW and
    # forcing a GC pass, same idiom used elsewhere in this script's own
    # PowerPoint-pass cleanup probes, actually lets the process exit.
    if ($wb) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null }
    if ($xl) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null }
    $wb = $null; $xl = $null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    [System.GC]::Collect()  # a single pass reliably left some RCWs alive in testing -- a second pass after WaitForPendingFinalizers cleared the rest
    Start-Sleep -Milliseconds 500
    # Belt-and-braces: this script's own opening section already treats a
    # leftover EXCEL.EXE as something to self-heal from, not tolerate --
    # if COM release still didn't let this run's own process exit, force
    # it rather than leaving it for the *next* run to clean up. Restricted
    # to windowless processes specifically (every zombie observed while
    # fixing this had a blank MainWindowTitle) so a real, separate,
    # visible Excel session a human opened during the run is never touched.
    # A BLANK TITLE IS NOT THE ONLY SHAPE A ZOMBIE TAKES. On 2026-08-09 a run
    # left an EXCEL.EXE whose MainWindowTitle was the bare word "Excel" -- an
    # empty frame with no workbook -- so this filter did not match it and the
    # self-heal quietly did nothing. Four tests in the NEXT run then ERRORED
    # while the runner still printed RESULT: OK, which is how the leak presents:
    # never as itself, always as unrelated tests failing later.
    #
    # The discriminator is now "no document loaded", which covers both shapes.
    # A human's Excel showing a workbook always carries its name in the title
    # ("register-wide.xlsx - Excel"), so a real working session is still never
    # touched. An Excel a human left open on the start screen with nothing
    # loaded does match and will be closed -- it has no workbook, so there is
    # nothing unsaved to lose.
    Get-Process EXCEL -ErrorAction SilentlyContinue |
        Where-Object { -not $_.MainWindowTitle -or $_.MainWindowTitle -eq 'Excel' } |
        Stop-Process -Force -ErrorAction SilentlyContinue
}

# --- Report --------------------------------------------------------
if ($pptError) {
    Write-Output "=== PowerPoint pass: DRIVER ERROR ==="
    Write-Output $pptError
}
else {
    Write-Output $pptReport
}

if ($excelError) {
    Write-Output "=== Excel pass: DRIVER ERROR ==="
    Write-Output $excelError
}
else {
    Write-Output $excelReport
}

if (-not $env:DECK_SYNC_KEEP_STAGING) {
    Remove-Item -Path $staging -Recurse -Force -ErrorAction SilentlyContinue
}

# --- Did anything actually RUN? ------------------------------------
#
# A VBA compile error anywhere in the project makes Application.Run fail without
# producing a single test result -- and until 2026-07-28 this script reported that
# as a clean run: zero PASS lines, zero FAIL lines, exit 0. A grep for "^FAIL"
# found nothing and the run looked green. The compile error that caused it
# (Application.PathSeparator, an Excel-only member used in a PowerPoint-hosted
# module) therefore reached Rohan's screen instead of this script's output.
#
# Silence is not success. An empty result set is the loudest possible failure --
# it means the project did not even compile -- so it must exit non-zero.
$allOutput = "$pptReport`n$excelReport"
$passCount = ([regex]::Matches($allOutput, '(?m)^PASS')).Count
$failCount = ([regex]::Matches($allOutput, '(?m)^FAIL')).Count
# A test that RAISES reports as "ERROR <name>" and was counted as neither a
# pass nor a failure until 2026-08-04 -- so a blown-up test printed its error
# line and the run still summarised as "N passed, 0 failed" and exited 0.
# Found by breaking UpsertRow's row matching on purpose: the test that exists
# to catch that regression errored, and the summary line called it clean.
# Same "silence is not success" rule as the empty-result check below; it had
# simply never been applied to this case.
$errorCount = ([regex]::Matches($allOutput, '(?m)^ERROR')).Count

Write-Output ""
if ($errorCount -gt 0) {
    Write-Output "=== $passCount passed, $failCount failed, $errorCount ERRORED ==="
    Write-Output "    An ERRORED test raised before it could assert. It is NOT a pass."
} else {
    Write-Output "=== $passCount passed, $failCount failed ==="
}

if ($pptError -or $excelError) {
    Write-Output "=== DRIVER ERROR (see above) ==="
    exit 3
}
if ($passCount -eq 0) {
    Write-Output "=== NO TESTS RAN. This is a FAILURE, not a pass. ==="
    Write-Output "Almost always a VBA compile error somewhere in the project:"
    Write-Output "  Application.Run reports 'Sub or function not defined' for a"
    Write-Output "  compile error anywhere, not only for a missing macro."
    Write-Output "  Open the project and use Debug > Compile VBAProject -- its"
    Write-Output "  dialog names the offending line directly."
    exit 2
}
if ($failCount -gt 0 -or $errorCount -gt 0) { exit 1 }
exit 0
