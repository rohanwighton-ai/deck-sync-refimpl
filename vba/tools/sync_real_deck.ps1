# Driver for SyncRealDeck.bas -- THIS IS THE REAL SYNC, NOT A PREVIEW.
#
# It opens the deck read-WRITE and runs RunSync.RunRoutineSync, which corrects
# field text, duplicates the template slide for any Data row that has no
# matching slide, and reorders slides. Run preview_real_deck.ps1 first and read
# its "new slide(s) would be created" number before running this.
#
# By default SyncRealDeck.bas closes the deck WITHOUT saving, so a run cannot
# reach disk; pass -SaveWhenDone to persist. The paired workbook is opened
# read-only either way.
#
# Report is written in full to $OutFile (local only -- contains real business
# field values and instance keys, deliberately kept out of chat) with a
# summary printed to stdout.
#
# Run this from a Windows-local path, not \\wsl.localhost: PowerShell treats
# the UNC path as remote and refuses to load the script -- while still exiting
# 0. See AGENTS.md's Known Patterns.

param(
    [string]$DeckPath = "C:\Users\rohan\OneDrive\Claude\test1.pptx",
    [string]$RepoRoot = $(Split-Path -Parent (Split-Path -Parent $PSScriptRoot)),
    [string]$OutFile = (Join-Path $env:TEMP "sync_real_deck_report.txt"),
    [switch]$SaveWhenDone
)

$ErrorActionPreference = "Stop"

foreach ($progId in @("PowerPoint.Application", "Excel.Application")) {
    try {
        $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject($progId)
        Write-Output "=== ABORTING: $progId is already running -- close it first so this diagnostic doesn't attach to or interfere with a live session. ==="
        exit 2
    } catch {}
}

$staging = Join-Path $env:TEMP ("deck-sync-run-" + [guid]::NewGuid().ToString("N").Substring(0, 8))
New-Item -ItemType Directory -Path $staging | Out-Null
$vbaSourceDir = Join-Path $RepoRoot "vba"

# A hand-picked subset here used to be "known to compile together" -- until
# it wasn't. RunSync.PreviewRoutineSync grew a call to IdentityCheck (the R9
# duplicate-key check) with nothing here ever re-verifying the subset still
# covered what RunSync actually reaches. VBA compiles the WHOLE PROJECT, not
# just the lines that run -- one undefined module reference anywhere in the
# imported set fails Application.Run for EVERY entry point, including ones
# that never touch the missing module. Same failure shape build_ppam.ps1's
# own header already documents for ReviewQueue (2026-08-01). Found
# 2026-08-21 running preview_real_deck.ps1 against the real deck: "Sub or
# function not defined" on an entry point that doesn't even call
# IdentityCheck directly -- only RunSync does, three calls deep.
# Fix: stop hand-maintaining a second, silently-driftable copy of the
# dependency graph. Use the SAME canonical list build_ppam.ps1 uses for the
# real shipped add-in -- one list, one place it can go stale, and it's
# already proven to compile because it's what Rohan actually runs.
$modules = @(
    "Discovery.bas", "InjectPrimitive.bas", "Matching.bas", "Resolve.bas",
    "SyncOperations.bas", "Onboarding.bas", "ExcelOutput.bas", "Verification.bas",
    "SlideDuplication.bas", "TemplateSlide.bas", "TemplateAudit.bas", "IdentityCheck.bas", "TagMigration.bas", "PlaceholderCheck.bas", "RunSync.bas", "DeckAdoption.bas", "ResolveFields.bas",
    "DeckRegistry.bas", "WorkbookBridge.bas", "Harvest.bas", "RibbonUI.bas",
    "AdoptFlow.bas", "BatchOnboardFlow.bas", "CommandBarUI.bas",
    "ReviewQueue.bas", "FieldWiring.bas", "MilestoneDevice.bas",
    "Drafting.bas", "FieldSpec.bas", "Sources.bas", "DraftingUI.bas", "DiscoverUI.bas",
    "DraftingLobby.bas", "AppEvents.cls", "ShapeAddressBook.bas", "Timing.bas", "FormattingAudit.bas"
)
foreach ($m in $modules) {
    $srcPath = Join-Path $vbaSourceDir $m
    $dstPath = Join-Path $staging (Split-Path $m -Leaf)
    if ($m -like "*.cls") {
        # CRLF, NOT A PLAIN COPY -- see run_vba_tests.ps1's own comment on this
        # exact line. LF-only (this repo's native format) makes Import() treat
        # the class header as unrecognised and silently import it as a plain
        # Standard Module instead, which then fails to compile with a generic
        # "Expected: end of statement" nowhere near the real cause (WithEvents).
        ((Get-Content $srcPath -Raw) -replace "`r`n", "`n" -replace "`n", "`r`n") | Set-Content $dstPath -NoNewline
    } else {
        Copy-Item $srcPath -Destination $dstPath
    }
}
Copy-Item (Join-Path $vbaSourceDir "tools\SyncRealDeck.bas") -Destination $staging

$ppt = $null
$scratch = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $ppt.Visible = -1
    $scratch = $ppt.Presentations.Add()

    foreach ($m in ($modules + @("SyncRealDeck.bas"))) {
        $leaf = Split-Path $m -Leaf
        $null = $scratch.VBProject.VBComponents.Import((Join-Path $staging $leaf))
    }
    Write-Output ("Imported " + ($modules.Count + 1) + " modules.")

    # InvokeMember rather than $ppt.Run(...): PowerShell cannot bind
    # Application.Run's VARIANT-optional-parameter signature directly.
    if ($SaveWhenDone) { Write-Output "*** -SaveWhenDone SET: this run WILL write to $DeckPath ***" }
    $report = $ppt.GetType().InvokeMember("Run", [System.Reflection.BindingFlags]::InvokeMethod, $null, $ppt, @([string]"SyncRealDeck.SyncRealDeck", [string]$DeckPath, [bool]$SaveWhenDone.IsPresent))

    Set-Content -Path $OutFile -Value $report -Encoding UTF8
    Write-Output "=== Full report written to: $OutFile ==="
    Write-Output ""

    # Summary only to stdout -- the per-field before/after lines carry real
    # business content.
    $report -split "`r?`n" | Where-Object {
        $_ -match "^(Deck:|Workbook:|Slides in:|Slides out:|Saved flag|=== |Summary:|Resequenced|Presentation dirty|Closed WITHOUT|SAVED to disk)" -or
        $_ -match "^  (created|FAILED|flagged)"
    } | ForEach-Object { Write-Output $_ }
}
finally {
    if ($scratch -ne $null) { try { $scratch.Saved = $true; $scratch.Close() } catch {} }
    if ($ppt -ne $null) { try { $ppt.Quit() } catch {} }
    Remove-Item -Recurse -Force $staging -ErrorAction SilentlyContinue
}
