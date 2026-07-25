# VBA implementation: `ResolveFields.bas`

Implements the "Resolve Unmatched Fields" flow from `specs/ribbon-ui.md`: a
human resolving a medium-confidence match found when a *subsequent* slide is
checked against an already-established template (`onboarding.md`'s matching
case, distinct from first-time onboarding, which needs no code at all -- see
`Onboarding.bas`'s own header comment). This is one of five independent
sub-tasks `IMPLEMENTATION_PLAN.md` Priority 21 splits `ribbon-ui.md` into;
the other four (ribbon XML/packaging, the New Period picker, the Onboard New
Slide Type flow, and the shared result form) are untouched by this pass.

**Not executed or verified in this environment** -- same constraint as every
module built this project since the 2026-07-25 pass began: this container has
no `powershell.exe` (confirmed via `which powershell.exe`), so there is no
Windows/Office install reachable to run `run_vba_tests.ps1` against. The
manual verification recipe below is how to actually prove it against a real
Office install; the 5 tests added to `vba/tests/TestRunner.bas` are ready to
run the next time this project is picked up on the WSL/Windows host.

## Scope: exactly what the spec calls "the only new code"

`ribbon-ui.md`'s own text: "the only new code is the shape-click capture and
role picker, both pure UI; the existing `confirm_field_match` primitive... no
new matching logic." Confirmed by reading `Onboarding.bas` directly:
`ConfirmFieldMatch(shp As Object, role As String)` (line 182) already exists,
already tags `shp.Tags.Add "role", role`, and is already exercised indirectly
by `SPIKE_NOTES_Onboarding.md`'s tests. `ResolveFields.bas` calls it
unchanged -- no new tagging or matching logic was written here.

## Design decision: InputBox, not a UserForm ListBox

The natural UI for "pick one role from a list" is a `UserForm` with a
`ListBox` control. This project has never authored a VBA UserForm before
(confirmed via `find vba -iname "*.frm"` returning nothing prior to this
pass), and unlike every other VBA gotcha logged in `AGENTS.md`'s Known
Patterns, this one genuinely can't be confirmed or refuted by careful
reading -- it depends on exactly how the VBE serializes a `UserForm`'s
control tree (some VBA/VB dialects keep the whole layout as plain
`Begin`/`End` text in the `.frm`; others move it into an opaque binary
`.frx` referenced via `OleObjectBlob`), and this container has no VBE to
export/import against and check which applies here. Hand-authoring a file in
a format that can't be verified risks committing something that silently
fails to import -- worse than not building it, since nothing in this
project's toolchain would catch that failure before a human hits it on a
real machine.

Given that risk, this flow uses `InputBox` instead: a numbered-list prompt
(`BuildRolePickerPrompt`) built from the template's own role list, and an
answer parser (`PickRoleFromList`) that accepts either the number or the
role name typed directly. Lower-fidelity than a real picker, but built from
mechanisms this project has already used successfully elsewhere (plain
`String` prompts/reports, e.g. `RunSync.RunRoutineSync`'s own report). A
future pass with real Office access can build the nicer `UserForm` version
and verify the `.frm`/`.frx` question for real rather than guessing --
`specs/ribbon-ui.md`'s shared-result-form bullet will hit this identical
open question for its own dialog, so whichever pass resolves it first should
leave a note here (and there) on what the real answer turned out to be.

## Design decision: pure-logic helpers separated from the interactive entry point

Mirrors `DeckAdoption.bas`'s own posture (see its SPIKE_NOTES) of keeping
decision logic testable by taking objects as parameters rather than reading
`Application.ActiveWindow.Selection`/driving `InputBox` deep inside a
function `TestRunner.bas` can't otherwise exercise:

- `PromptResolveUnmatchedField(templateSld)` -- the actual ribbon-button
  entry point. Reads the live selection, drives the `InputBox`, calls
  `ConfirmFieldMatch`. Not covered by an automated test (no headless harness
  can click through a live `InputBox` -- confirmed by reading
  `run_vba_tests.ps1` and `TestRunner.bas`, neither references `InputBox` or
  any interactive dialog anywhere else in this project either). Covered by
  the manual verification recipe below instead.
- `ValidateSingleShapeSelection(sel, ByRef outShp)`, `BuildRolePickerPrompt
  (roles())`, `PickRoleFromList(answer, roles())` -- pure logic, all `Public`
  so `TestRunner.bas` can call them directly.

`ValidateSingleShapeSelection` turned out to be **genuinely testable against
a real selection, not just a stub**: `run_vba_tests.ps1` runs PowerPoint
visible (`$ppt.Visible = -1`, confirmed by reading the script), so
`Application.ActiveWindow` is real and `shp.Select` / multi-shape
`Shapes.Range(Array(...)).Select` calls actually change
`Application.ActiveWindow.Selection` the same way a live user click would.
`Test_ResolveFields_ValidateSingleShapeSelectionAcceptsOneShape` and
`Test_ResolveFields_ValidateSingleShapeSelectionRejectsMultiple` exercise
this for real, not through a mock -- the one part of this flow that
initially looked like it would have to stay manual-only turned out not to.

## What `Test_ResolveFields_EndToEndTagsSelectedShapeViaPickedRole` proves that the unit tests don't

The four unit-level tests (`ValidateSingleShapeSelection` x2,
`BuildRolePickerPrompt`, `PickRoleFromList`) each check one function in
isolation. This fifth test wires `ValidateSingleShapeSelection` ->
`Onboarding.BuildTemplateFieldShapes` -> `PickRoleFromList` ->
`Onboarding.ConfirmFieldMatch` together exactly as
`PromptResolveUnmatchedField` does internally (skipping only the `InputBox`
call itself, substituting a literal `"1"` answer) and asserts the selected
shape actually ends up tagged with the picked role -- the actual claim this
whole flow exists to make true, not just that each piece works alone.

## Manual verification recipe

Import `Discovery.bas`, `InjectPrimitive.bas`, `Matching.bas`, `Resolve.bas`,
`SyncOperations.bas`, `Onboarding.bas`, and `ResolveFields.bas` into the same
VBA project (`ResolveFields.bas` calls into `Onboarding` directly).
Automated coverage lives in `vba/tests/TestRunner.bas`'s 5 new
`Test_ResolveFields_*` functions -- run via `vba/tests/run_vba_tests.ps1` on
a real Windows/Office host, same as every other module's tests, per
`AGENTS.md`'s Testing section.

For the one piece the harness can't cover (the live `InputBox` interaction
inside `PromptResolveUnmatchedField` itself):

1. Build a template slide with at least one tagged field (e.g. a textbox
   with `shp.Tags.Add "role", "Title"`).
2. Add a second slide with an untagged textbox meant to be that same field.
3. Click the untagged textbox in the open deck to select it (exactly one
   shape).
4. In the Immediate window: `? ResolveFields.PromptResolveUnmatchedField
   (ActivePresentation.Slides(1))` (or whichever slide index is the
   template).
5. An `InputBox` should appear listing the template's roles numbered from 1;
   enter either the number or the exact role name.
6. Confirm the function's return string reports "Assigned role '...' to
   shape '...'", and that the clicked shape now carries that `role` tag
   (`? ActiveWindow.Selection.ShapeRange(1).Tags("role")`).
7. Re-run with nothing selected (click empty slide background first) and
   confirm it returns "Select exactly one shape first." without raising.
8. Re-run with two shapes selected and confirm it returns "Select exactly
   one shape (selected 2)." without raising.
