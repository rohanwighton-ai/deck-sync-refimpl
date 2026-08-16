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
  **Hit a third time, 2026-07-26**, in `BatchOnboardFlow.bas`'s
  `BuildBatchPlan` (a `Candidate()` array cached per-slide in a
  Dictionary) -- cost a long, confusing live-debugging session (VBE's
  error-highlight under `Application.Run` landed on the calling function's
  signature or an unrelated `Dim` line, not the actual offending
  statement, making it look like a totally different bug each time) before
  a fresh Fable-model agent's own from-scratch empirical probe reproduced
  the real compile error text directly. **Takeaway for next time: check
  this file's Known Patterns *before* writing new Dictionary-caching code
  that might hold a UDT array, not after chasing the symptom for hours.**
  Fixed by flattening the per-slide arrays into one array-of-primitives-
  indexed set (`allOtherCandidates()`/`otherSlideCandStart()`/
  `otherSlideCandCount()`) instead of a Dictionary -- see
  `SPIKE_NOTES_BatchOnboardFlow.md`.

  **Hit a fourth time, 2026-08-16, via `Collection` rather than
  `Dictionary`** -- same restriction, different container: a `Collection`
  is also late-bound/`Variant`-based internally, so `coll.Add someUDT`
  fails at compile time with "Only user-defined types defined in public
  object modules can be coerced to or from a variant or passed to
  late-bound functions" (`DraftingLobby.ReadLobby`, caught by the compile
  gate in `run_vba_tests.ps1` before it ever reached a real test). Same
  fix as always: an array of the UDT (`LobbyEntry()`), not a `Collection`
  or `Dictionary` of it -- matching how `SyncAction()`/`ReviewItem()`
  already do this throughout the codebase. **The existing "check this
  file before writing new Dictionary-caching code" takeaway above should
  read `Dictionary` **or** `Collection`** -- this file was not checked
  first this time either, and would have caught it before the compile
  gate did.

- **A `.cls` FILE WITH LF-ONLY LINE ENDINGS IMPORTS SILENTLY AS THE WRONG
  COMPONENT TYPE.** `VBComponents.Import()` needs the `VERSION 1.0 CLASS` /
  `BEGIN` / `END` header block terminated with CRLF to recognise a class
  module at all. This repo's files are LF-only (WSL, git) -- fine for every
  `.bas` file, because a standard module needs no header signature to
  detect. `AppEvents.cls` (2026-08-16, the first class module this project
  has ever had) imported with NO ERROR, but as `vbext_ct_StdModule` (Type
  1) instead of `vbext_ct_ClassModule` (Type 2) -- confirmed by querying
  `.Type` directly on the imported component, not assumed. The only visible
  symptom was `WithEvents` failing to compile with a generic "Compile
  error: Expected: end of statement", reported on the `WithEvents` line
  itself, which reads exactly like a `WithEvents`-specific problem (tried
  `Public` vs `Private`, qualified `Excel.Application` vs bare
  `Application` -- neither made any difference, because neither was ever
  the real cause). Confirmed root cause by reading the STAGED file's raw
  bytes (`Get-Content -Encoding Byte`) and finding byte 10 with no
  preceding byte 13 right after "VERSION 1.0 CLASS". Fixed in all three
  places a `.cls` gets staged (`build_ppam.ps1`, `run_vba_tests.ps1`,
  `field_e2e.ps1`): CRLF-normalise `.cls` files specifically during
  staging, leave `.bas` files untouched. **Takeaway: when a symptom is on
  the line that looks obviously relevant, check the file actually compiled
  as the component type you think it is before trusting the symptom's own
  framing of the problem.**

- **ONE WRITER ON THE RIG AT A TIME. Delegating an Office task means not doing
  it yourself.** 2026-08-01: a Fable agent was put on the property-persistence
  bug, and the main agent then ran the same experiments concurrently -- same
  49MB deck at `deck-sync-e2e\e2e-deck.pptx`, same uncommitted working tree,
  and repeated `Stop-Process -Force` on POWERPNT/EXCEL between its own runs.
  Those kills landed on the agent's live Office session. It reported
  "unexplained crashes" and only diagnosed the cause by sweeping for stray
  PowerPoint processes and finding a second `wsl.exe` parent. It was reading the
  main agent's half-finished edits to `E2EField.bas` off the shared filesystem
  at the same time.
  **Why this is worse here than in ordinary parallel work:** the shared state is
  not just the repo. It is a single external file that Office holds open, plus a
  global process namespace where "kill all PowerPoint" cannot be scoped to one
  caller. Two agents cannot both own that.
  **The rule:** one active writer on `deck-sync-e2e/` at a time. If a second
  session genuinely needs to run, give it its own rig copy AND accept that
  process kills are still global -- so in practice, wait. Results obtained while
  another actor was writing the same rig should be treated as unreliable
  regardless of what they say; several of that night's confusing measurements
  are now suspect for exactly this reason.

- **RESOLVED 2026-08-01, was an open bug: writing a PowerPoint CustomDocumentProperty often
  does not persist, and reporting success proves nothing.** `pres.Save` returns
  cleanly, `pres.Saved` flips to True, and reading the property back in the same
  session returns the NEW value -- while the file on disk still holds the old
  one. Confirmed by reading `docProps/custom.xml` out of the .pptx directly.
  Observed sequence: PROBE-Q9 -> FY26Q4 failed, FY26Q4 -> FY27Q1 SUCCEEDED,
  FY27Q1 -> FY26Q4 failed, -> FY27Q2 failed, -> FY26Q4 failed. Not value-
  specific, not obviously timing-specific; the one success is the anomaly. A
  property that does not yet EXIST does persist when Added -- which is why this
  hid for so long: every first write worked and only updates were lost.
  **Two things make it survivable rather than dangerous.** `E2EField.SetPeriod`
  now closes, reopens and compares, and prints `*** DID NOT PERSIST ***` rather
  than claiming success. And `vba/tools/read_deck_props.py` reads the property
  straight out of the zip with no Office involved -- **that is the authoritative
  check**, because an in-process reopen can be served from PowerPoint's cache
  and is untrustworthy in BOTH directions.
  **Do not trust any deck-stored setting** (period, workbook path, deck id)
  after an update until this is resolved. Verify with the offline reader.
  Untried next steps: `SaveAs` to the same path instead of `Save`; writing the
  property with the presentation's window hidden; checking whether a second
  PowerPoint instance holds a lock.

- **The "Application.Run cannot see a function that declares cross-module Public
  UDTs" trap DID NOT REPRODUCE when tested directly (2026-07-31), and the
  warm-up probes built around it are deleted.** `E2EFirstField.DumpFieldValues`
  had been recorded as uncallable all day for this reason. Run against a freshly
  imported project with NO probes and NO compile, it resolved first try and
  dumped 230 rows (5 fields x 46 entities). The A/B was run failure-first on
  purpose: proving something works after a fix means nothing unless the failure
  was demonstrated before it.
  Treat the original diagnosis as unproven. Every "Sub or function not defined"
  found today had an ordinary cause -- a module in the driver's staging list but
  missing from its import list, a VBA reserved word used as a variable name, a
  rename that updated a Function's declaration but not its return line. All
  three surface as that same message naming the WRONG thing, usually the entry
  point rather than the fault, which is exactly how a boring cause gets
  attributed to an exotic one.
  **What replaced the probes:** an explicit compile after import, via the VBE's
  documented `Debug > Compile VBAProject` control. The VBE main window must be
  `Visible` for it to execute, and `Execute` returns NOTHING -- it cannot report
  success, and the item is disabled when the project is already compiled. Never
  trust its return; judge by whether the real call afterwards resolves. Kept not
  because the trap was real, but because compiling at a known point turns a
  would-be runtime mystery into a compile error where it belongs.

- **VBA's `IIf` evaluates BOTH branches, always.** It is a function call, not a
  short-circuiting operator, so `IIf(cond, F(x), -1)` calls `F(x)` even when
  `cond` is False. Hit 2026-07-31 with
  `IIf(dpos <= Len(s), AscW(Mid(s, dpos, 1)), -1)`, written specifically to
  avoid `AscW("")`: the guard was False exactly when it mattered, `AscW` was
  called anyway, and run-time error 5 followed. Headless that is a modal dialog
  and a hung run with no output.
  **This is the always-true-guard failure in a new costume** -- a conditional
  that reads as protection and provides none. Use a real `If ... Then` whenever
  the non-taken branch would error; reserve `IIf` for choosing between two
  values that are both already safe to compute.

- **Three consecutive headless runs were lost to COMPILE errors on 2026-07-31**,
  each ~8 minutes, each presenting as a hang with no output rather than a
  message: a module missing from the driver's import list, a VBA reserved word
  used as a variable name (`Dim empty As ...`), and a rename that updated a
  Function's declaration but not its `Name = result` return line. All three are
  now caught by `check_vba_static.py` in under a second -- run it before every
  suite run, not after a failure.
  **The meta-lesson cost more than the bugs.** Two of those checks were written
  and declared good against a handful of synthetic cases, then reported 27 and
  then 37 false positives the moment they met the real 28 modules -- because
  `Public Type Foo` binds a naive name-capture to the word "Type", and because
  `Set Name = obj` and `Case "x": Name = y` are both perfectly ordinary
  assignments. **A new static check is not finished until it has been run
  against the whole corpus and come back clean.** Synthetic cases prove it fires;
  only the corpus proves it does not cry wolf, and a checker that cries wolf is
  worse than no checker (this file's own docstring says so, which did not
  prevent it).

- **VBA: a module-level `Type`/`Const` below the module's first `Sub`/`Function`
  reports its error in a DIFFERENT module.** Declarations must all sit above the
  first procedure. Break that and the offending module compiles quietly; the
  failure appears as `User-defined type not defined` at whichever *other* module
  referenced the Type — which reads as "the module didn't import" rather than
  "the declaration is in the wrong place", and the import log says it imported
  fine. Hit 2026-07-30 with `TemplateSlide.bas`'s `MakeTemplateResult` placed
  after `PlaceholderFor`: cost one full ~8-minute suite run, then was identified
  in seconds from a screenshot of the VBE dialog (which names the line). **Two
  takeaways**: put declarations first as a matter of course, and when a run dies
  with `Sub or function not defined` / `User-defined type not defined`, ask for
  the VBE screenshot before grepping — the dialog already names the statement.
  A cheap static check, worth running before any suite run that adds a module:
  scan each `.bas` for a `^(Public |Private )?(Type|Const|Enum)\s` line
  appearing after the first `^(Public |Private )?(Sub|Function|Property)\s`
  (procedure-local `Const`s inside a body are legal and will show up as false
  positives — check the indentation).

  **The SAME rule applies to a bare module-level variable, and the checker
  did not know that until 2026-08-17.** `Public foo As Bar` / `Private
  WithEvents mApp As Excel.Application` at column 0 is a declaration exactly
  like `Type`/`Const`/`Enum` — it must precede the first procedure or VBA
  reports the error somewhere else entirely. Hit twice in one night:
  `DraftingLobby.mAppEvents` (building the Lobby's event mechanism) and
  `ReviewQueue.mTestForceInjectCrash` (building this file's own 50290 fix,
  two hours later, same mistake) — both times `check_vba_static.py` reported
  "clean" immediately before the live compile failure, because its
  `DECL_RE` only matched the `Type|Const|Enum` keyword, never a plain
  variable. Fixed by adding `VAR_DECL_RE` alongside it; proven by
  deliberately reintroducing the real `ReviewQueue.bas` defect and
  confirming the checker now names it before restoring the fix — the same
  "make it fail once before trusting it" discipline this project applies to
  its own VBA, now applied to the checker that watches the VBA.

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

- **VBA: an application's named constants (e.g. Excel's `xlToLeft`/`xlUp`)
  only resolve inside that application's own VBA project.** A module built
  to run cross-app (e.g. `ExcelOutput.bas`, driven from PowerPoint per
  `vba-port.md`'s "runs inside Excel or drives it via COM from the
  PowerPoint side") cannot rely on the foreign app's named enum constants
  being defined -- confirmed 2026-07-25: `xlToLeft` compiled and ran fine
  inside Excel's own project (every `TestRunnerExcel.bas` run before this
  looked completely clean) but raised "Variable not defined" the moment
  `RunSync.bas` (PowerPoint-hosted) called into `ExcelOutput.bas` for the
  first time. **Use the numeric literal instead** (e.g. `XL_TO_LEFT =
  -4159`, `XL_UP = -4162`, as module-level `Private Const`s with a comment)
  for any constant a module might use from outside its "home" application.
  A module tested only in its home app can carry this bug invisibly for an
  arbitrarily long time -- it only surfaces once something genuinely drives
  it cross-app, which is exactly what happened here.

- **Reading slide text offline (straight from the `.pptx` OOXML) must join
  paragraphs with CR, or it silently disagrees with VBA on every
  multi-paragraph field.** `TextRange.Text` returns a shape's paragraphs
  CR-separated (`vbCr`) and its soft line breaks as `Chr(11)`; the XML has
  no separator at all -- each `<a:p>` is just another element, and the
  obvious `''.join(all <a:t> text)` concatenates "Project Closed" and
  "Results packaged in..." into one run-on string. Confirmed 2026-07-27:
  an offline verifier built this way reported **49 false mismatches**
  between a workbook and the deck it had just been generated from, all of
  them on multi-paragraph fields, none of them real. Iterate `<a:p>`
  elements and join their text with `\r`, mapping `<a:br/>` to `\r` too.
  This matters well beyond verification -- any offline harvest whose output
  is written into the Data sheet will otherwise differ from what VBA reads
  back off the slide, and a routine sync would then rewrite every affected
  slide with a reformatted version nobody asked for.

- **Excel escapes control characters in shared strings as literal `_xHHHH_`
  text**, so a cell holding a CR reads back from the raw XML as the 7
  characters `_x000D_`, not as `\r`. Any offline join between a workbook
  value and a deck value has to decode those first (or the comparison
  fails on exactly the multi-paragraph fields that matter). Decode for
  COMPARISON only -- write the untouched original string back, so the
  value continues to round-trip byte-identically to what VBA harvested.
  Both halves of this were hit within minutes of each other 2026-07-27.

- **Reading a OneDrive-backed file through `/mnt/c` from WSL can return STALE
  bytes while Windows/Office sees the current file.** Confirmed 2026-07-27
  the hard way: right after a cell was written and saved by Excel, a Python
  read of the same `.xlsx` through `/mnt/c` still reported the *previous*
  value -- specifically the exact state the file had been in one save earlier.
  Excel, reading the same path from the Windows side, returned the correct
  current value. Both readers were internally correct; the bytes visible to
  each differed. The files carry the `ReparsePoint` attribute (OneDrive Files
  On-Demand), and the placeholder WSL reads through is not guaranteed to be
  materialized in step with a Windows-side write.
  **Why this is dangerous rather than merely annoying**: it fails in both
  directions. It can invent a discrepancy that does not exist (this cost a
  real detour hunting a "lost write" that had never been lost), and it can
  equally hide one that does. An offline verifier is only trustworthy here as
  a SECOND opinion -- when an offline read disagrees with what an in-Office
  run reports, believe Office and re-read, rather than believing the file.
  Cheap disambiguation: read the same cell back through Excel COM
  (`$wb.Workbooks.Open(path, 0, $true)` read-only) and compare. Note the
  session that hit this had already, unknowingly, been protected by exactly
  that cross-check: the offline repair verification agreed with what the
  real-Office preview and sync independently reported, which is the only
  reason its conclusions stood up.

- **`powershell.exe -File` on a `\\wsl.localhost\...` path can refuse to run
  and still exit 0.** PowerShell treats the UNC path as remote, so the
  default execution policy blocks the unsigned script -- it prints a
  `SecurityError` to stderr and exits **cleanly**. Confirmed 2026-07-27:
  a `run_vba_tests.ps1` invocation reported success having never started
  Office. **A green exit code from this script proves nothing; count the
  `PASS` lines.** The fix needs no execution-policy change: copy the
  `.ps1` alone to a Windows-native path and pass the repo back in --
  `cp vba/tests/run_vba_tests.ps1 "$WINTMP/" && powershell.exe -NoProfile
  -File "$(wslpath -w "$WINTMP/run_vba_tests.ps1")" -RepoRoot "$(wslpath -w .)"`
  -- since the script already stages every `.bas` into Windows temp itself
  (see its own header); only loading the script file crossed the boundary.
  Related trap from the same run: piping the output through `tail -45`
  truncated away the entire PowerPoint section, leaving a partial result
  that looked like a complete one. Capture the full log to a file and
  grep it.

- **The cross-application trap runs BOTH ways, and it is not only constants:
  `Application` itself is a different object in each host.** Already recorded
  below for Excel's `xlToLeft`/`xlUp` failing in a PowerPoint-driven module;
  hit from the opposite direction 2026-07-28 with
  **`Application.PathSeparator`, which exists on Excel's Application object but
  NOT on PowerPoint's** -- a hard compile error ("Method or data member not
  found") in `DeckRegistry.bas`. Before reaching for any `Application.<member>`,
  check which host that module actually runs in. This add-in is Windows-only, so
  a literal `"\"` is the correct answer for a path separator, not a lookup.
  Notable: this was hit within hours of re-reading this very file. The written
  warning did not prevent it; a real compile error did.

- **A test run that produces ZERO results is a FAILURE, not a pass -- and until
  2026-07-28 `run_vba_tests.ps1` reported it as a pass.** A VBA compile error
  anywhere in the project makes `Application.Run` fail before any test executes,
  so the run emitted no `PASS` lines, no `FAIL` lines, and exit 0. Grepping for
  `^FAIL` found nothing and the run looked green; the compile error reached
  Rohan's screen instead of the script's output. The script now counts `PASS`
  lines, prints an explicit `=== N passed, M failed ===` verdict, and exits **2**
  when nothing ran (1 = real failures, 3 = driver error). **Never infer success
  from the absence of failures** -- require positive evidence that tests actually
  executed. Same false-green shape as the `powershell.exe` UNC refusal logged
  above, which also exited 0 while doing nothing.

- **A hung headless run (no output, process still alive) is very often a
  VBA *compile* error, not a runtime hang** -- `On Error` cannot catch
  compile-time issues (see the Testing section's modal-dialog note below),
  so they can manifest as an indefinite wait rather than a clean COM
  exception. Don't keep guessing blindly at the cause: take an actual
  screenshot. `Add-Type -AssemblyName System.Windows.Forms,
  System.Drawing` + `[Graphics]::CopyFromScreen` captures the full screen
  to a PNG reliably; bring the right window to front first with
  `SetForegroundWindow`/`ShowWindow` (P/Invoke `user32.dll`) since a hung
  Office process is very likely sitting behind whatever terminal window
  triggered it, not on top. The compile-error dialog's own text names the
  exact problem directly -- confirmed 2026-07-25, both bugs in this list
  ("Sub or function not defined" from a missing `ExcelOutput.bas` import,
  and this one) were found this way after other diagnostic techniques
  stalled.

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
