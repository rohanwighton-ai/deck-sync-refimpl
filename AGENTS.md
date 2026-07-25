# Operational Learnings

This file contains project-specific guidance that Ralph has learned through observation.

Start minimal. Add entries only when Ralph exhibits repeated failures or needs specific
guidance.

## Build/Test Commands

- Tests: `python3 -m pytest tests/ -v`
- Type check: `python3 -m mypy src/`
- No pip packages beyond pytest/mypy are installed in this image. Prefer stdlib
  (`zipfile`, `xml.etree.ElementTree`) over adding new dependencies unless a spec
  genuinely can't be satisfied without one — this project deliberately stays
  dependency-light so it's trivially runnable.

## Known Patterns

- OOXML (`.pptx`) shape trees: groups are `<p:grpSp>`, not automatically flattened —
  must be walked recursively. See `src/discovery.py` for the reference walk.
- `xml.etree.ElementTree.Element.__bool__` is based on child count, not None-ness, and is
  deprecated for exactly this reason. A chain like `el.find(a) or el.find(b)` will
  silently pick the wrong result if the first match has no children (e.g. `<p:cNvPr
  name="..."/>` is a valid, real match with zero children). Always use explicit
  `is not None` checks per candidate, never `or`-chain `Element.find()` calls. (Hit and
  fixed during initial design, 2026-07-19 — see `src/discovery.py`'s `_shape_name`.)

- VBA: a user-defined `Type` cannot be assigned to a `Variant` (compile-time
  "Invalid use of type"). This means a `Scripting.Dictionary` -- whose `Item`
  is `Variant`-typed -- cannot hold a UDT value, even though it happily
  holds strings/numbers/objects. The natural VBA port of a Python
  `dict[str, SomeDataclass]` is a `Scripting.Dictionary`; it isn't viable
  once the value type is a UDT (`Candidate`, `MatchResult`, `InjectResult`,
  ...) -- use parallel arrays (same index across a `String()` key array and
  a `TheType()` value array) instead. Found while building `Onboarding.bas`
  (2026-07-25); `SyncOperations.bas`'s existing `PlanRoutineSync`
  (`changed(fieldName) = r`, `r As InjectResult`) appears to hit this same
  restriction and was fixed the same day (split into two Dictionaries of
  primitives, `ChangedFieldVerified`/`ChangedFieldError`).

- **VBA: `ReDim arr(1 To 0)` throws "Subscript out of range" (Err 9) at
  runtime** -- confirmed real via multiple clean, isolated repros against a
  genuine PowerPoint 16.0 install (2026-07-25), not a hypothetical or a
  version quirk assumed from documentation. This directly contradicts the
  "(1 To 0) means an empty 1-based array" convention every VBA module in
  this project used from the start (`Discovery.bas`, `Matching.bas`,
  `SyncOperations.bas`, `Onboarding.bas` all pre-ReDim'd to `(1 To 0)`
  before growing via `ReDim Preserve` -- none of it had ever been executed
  until this pass). `ReDim Preserve arr(1 To 0)` fails identically (the
  `Preserve` keyword doesn't help); the type doesn't matter either (Object,
  String, Variant, and UDT arrays all fail the same way).
  **The fix**: never pre-ReDim to an empty range. Just `Dim arr() As
  SomeType` and let the *first* `ReDim Preserve arr(1 To 1)` (when the
  first real item appears) allocate it fresh -- confirmed this works
  cleanly, including for UDT arrays. A function whose loop never finds
  anything then returns a genuinely **unallocated** array, not an empty
  `(1 To 0)` one -- callers must never assume `UBound`/`LBound` are safe to
  call on a returned array; guard with `On Error Resume Next` +
  `Err.Number` check first (same technique used to detect the failure
  itself). This guard **cannot go through a generic `Variant`-parameter
  helper function** for UDT arrays specifically -- passing a UDT array as
  a `Variant` argument to another function is itself a separate compile
  failure (confirmed), so the allocation check has to be inlined at each
  call site, in the same scope as the array variable, not factored into
  one shared utility. Fixed across all 6 production modules plus the test
  harness the same day this was found; see each module's own
  `SPIKE_NOTES_*.md` for the specific sites.

- **VBA: `Shell.Application.Namespace(zipPath)` (the standard "no zip
  library in stock VBA" escape hatch) reliably returns `Nothing` under COM
  automation** (`Application.Run` from an external client), even against
  an independently-verified-valid, non-corrupt zip -- confirmed 2026-07-25
  against a real `.pptx`, reproduced 3x including with a delay to rule out
  a shell-notification race. Root cause not fully isolated (plausible
  theory: the Explorer "compressed folder" namespace extension behaves
  differently under automation vs. genuine interactive use). **Prefer
  shelling out to `tar.exe`** instead (bsdtar, bundled with Windows 10
  1803+/Windows 11 by default, auto-detects zip format despite the name):
  `CreateObject("WScript.Shell").Run("cmd.exe /c tar -xf ""<zip>"" -C
  ""<destdir>"" ""<entry>""", 0, True)` -- windowStyle 0 keeps it
  invisible, the `True` waits for completion and returns tar's real exit
  code, avoiding the shell namespace layer (and its polling-for-async-
  completion dance) entirely. Fixed in `Matching.bas`'s `LoadPartXml`.

- **VBA: `CreateObject("MSXML2.DOMDocument60")` (no dots) can raise Err 429
  "ActiveX component can't create object" even when MSXML 6.0 is genuinely
  installed** -- confirmed 2026-07-25: the dotted ProgID form,
  `"MSXML2.DOMDocument.6.0"`, works and returns the identical real
  `DOMDocument60` object (`TypeName` confirmed) on the same machine where
  the no-dot form fails. This bug had been silently masked in
  `Matching.bas` for a full session because the `Shell.Application`
  zip-extraction failure above always short-circuited first, so this line
  had never actually been reached before both were found and fixed the
  same day. If a future module needs any other versioned MSXML/ActiveX
  ProgID, try the dotted form first rather than assuming the no-dot form
  is equivalent.

## Testing

- **A real, headless test harness now exists and has actually run against
  real Office** (2026-07-25) -- `vba/tests/TestRunner.bas` (PowerPoint:
  Discovery/InjectPrimitive/Matching/Resolve/SyncOperations/Onboarding) and
  `vba/tests/TestRunnerExcel.bas` (Excel: ExcelOutput), driven by
  `vba/tests/run_vba_tests.ps1` via COM automation from PowerShell
  (`powershell.exe` is reachable from this WSL environment and can drive a
  real Windows-side Office install directly -- confirmed, not assumed).
  Every `SPIKE_NOTES_*.md`'s "not executed or verified in this environment"
  disclaimer is now stale for the specific cases these tests cover; each
  file's own text still needs updating per-module as that debt gets paid
  down (`SPIKE_NOTES_Matching.md` already updated as the first case).
  Re-run via `powershell.exe -File <path to run_vba_tests.ps1 via
  wslpath -w>`. Requires "Trust access to the VBA project object model"
  enabled for both PowerPoint and Excel (`HKCU:\Software\Microsoft\
  Office\16.0\{PowerPoint,Excel}\Security\AccessVBOM` = 1) -- off by
  default, deliberately enabled on this machine 2026-07-25 with Rohan's
  explicit sign-off (a real security-relevant setting, not flipped
  silently). The script self-heals one real gotcha found the hard way: a
  zombie POWERPNT.EXE/EXCEL.EXE left over from an interrupted prior run
  makes `New-Object -ComObject` attach to the stale instance instead of
  spawning fresh, and that stale instance's `VBProject`/`VBComponents`
  access can come back genuinely null even with trust enabled -- the
  script kills any stray process first, always.
  Two more real COM-automation gotchas hit while building this: (1)
  `$app.Run("Mod.Func", arg1, arg2)` directly fails PowerShell's own method
  overload resolution for `Application.Run`'s VARIANT-optional-parameter
  signature ("Cannot find an overload for 'Run' and the argument count") --
  use `$app.GetType().InvokeMember("Run", [System.Reflection.BindingFlags]
  ::InvokeMethod, $null, $app, @(...))` instead, with every argument
  explicitly cast `[string]` first (implicit typing failed identically,
  "Sub or function not defined", even though the exact same call with
  literal string arguments worked). (2) An uncaught VBA runtime error
  inside code with no `On Error` handling can pop a real modal "Microsoft
  Visual Basic" dialog that blocks the whole automated run indefinitely --
  every test function in `TestRunner*.bas` wraps risky calls in `On Error
  Resume Next`/`Err.Number` checks specifically to avoid this, and any
  future ad hoc diagnostic code should too.

## Cost / Tool Selection

- This repo is small (specs/ and src/ are each under a dozen short files). A direct
  `Read` is strictly cheaper than a subagent spawn for a file this size — every subagent
  is a fresh, uncached API call, and pays a full context-establishment cost to read
  something a direct Read gets for a fraction of that. Reserve parallel subagents for
  genuinely large fan-out (dozens-to-hundreds of files, or slow independent searches),
  not as a default "study the codebase" step. (Root-caused 2026-07-25: an earlier
  version of `PROMPT_build.md`/`PROMPT_plan.md` authorized "up to 500 parallel Sonnet
  subagents" for exactly this small a repo, which was almost certainly the dominant
  cost driver during real iterations, independent of the retry-storm bug logged below.)
- If a task's obvious approach would be expensive relative to what it accomplishes, say
  so — in the commit message or a plan note — rather than silently paying the cost.
  Flagging "this would cost too much, here's a cheaper way" is a valid, wanted outcome
  of an iteration, not a failure to complete the task.

## Constraints

- This is a reference/test implementation, not the production sync engine — the real
  target is VBA. Python code here exists to harden the discovery/matching/verification
  *logic* against a growing fixture corpus (see `test-fixtures/`), not to become a
  shipped tool. Don't add production concerns (CLI, packaging, distribution) unless a
  spec asks for it.
- Keep this repo separate from both `claude-brain` and any CRC system — see the initial
  commit message for why.
- `vba/InjectPrimitive.bas` was a single de-risking spike, not yet governed by a spec
  the way every `src/*.py` file is. Further VBA porting (discovery, matching,
  sync-dispatch logic) is planned and now has a governing spec, `specs/vba-port.md` —
  written 2026-07-24, mirroring how every other module here began from a spec rather
  than growing VBA scope ad hoc after the fact. Follow its port order and
  manual-verification-recipe requirement for every subsequent VBA module.
