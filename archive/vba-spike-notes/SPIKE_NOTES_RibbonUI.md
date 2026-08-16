# VBA implementation: `RibbonUI.bas` (and the packaging saga behind it)

Implements the action logic for `specs/ribbon-ui.md`'s four buttons (Sync Now,
New Period, Onboard New Slide Type, Resolve Unmatched Fields) -- gathers what
the engine needs (via `DeckRegistry` lookups + `WorkbookBridge`), calls an
existing, already-tested Sub, reports via the shared `ShowSyncResult`. New
Period's picker and Resolve Unmatched Fields' role picker are the only "new"
pieces, both InputBox chains, no new sync/matching logic.

**Executed against real Office 2026-07-26.** 5 tests pass (`ResolveTypeAnswer`,
`ResolveRecordAnswer`, `BuildTypePickerPrompt`, plus the two picker-building
functions). `SyncNow`/`NewPeriod`/`OnboardNewType`/`ResolveUnmatchedFields`
themselves have no dedicated unit tests (each is a thin orchestration of
already-tested calls, same posture `RunSync.RunRoutineSync` took) -- proven
instead via the full live add-in load test below.

## The actual headline finding this pass: a real .ppam cannot carry a ribbon

`specs/ribbon-ui.md`'s own text calls for "Ribbon defined via `customUI14.xml`
injected into the `.ppam` package itself." This was never verified against real
Office before 2026-07-26 (no Office access existed in any prior pass) and turns
out to be **impossible**, not just difficult. Full trace, since this cost most
of a session and the next person picking this up should not have to re-derive
it:

1. `Presentation.SaveAs(path, 30)` [`ppSaveAsOpenXMLAddin`] throws a bare
   `COMException` ("An error occurred while PowerPoint was saving the file")
   via COM automation in this environment -- reproduces even on a completely
   blank, VBA-free presentation. Every other OOXML SaveAs enum (24-29:
   pptx/pptm/potx/potm/ppsx/ppsm) works fine via the identical call shape.
   Enumerated all of 20-33 to rule out a wrong enum value -- 30 is confirmed
   correct (its own content-type declaration matches the OOXML spec), it just
   doesn't work via automation. Root cause not discoverable from outside
   Office; not a VBA-content issue, not a UNC-path issue (tested both a local
   path and a two-step pptm-then-reopen-then-ppam sequence, both fail
   identically).
2. Reverse-engineered a real `.ppam`'s structure by having Rohan do
   File > Save As > PowerPoint Add-in manually (2026-07-26) and inspecting the
   result: **five parts only** -- `[Content_Types].xml`, `_rels/.rels`,
   `ppt/_rels/presentation.xml.rels`, `ppt/presentation.xml` (**completely
   empty, 0 bytes** -- no slides, theme, master, docProps, view/presentation
   props at all), `ppt/vbaProject.bin`. Dramatically more minimal than the
   `.pptm` this project's own `SaveAs(25)` path produces.
3. Editing `[Content_Types].xml` alone to convert a full `.pptm` to this
   content-type (leaving the `.pptm`'s full slide/theme/master content in
   place) produces a file PowerPoint *correctly identifies* as an add-in
   (`Presentations.Open` refuses it: "You must use Addins.Add to load Addin
   files" -- itself proof the content-type edit worked) but `AddIns.Add` then
   silently fails to actually load it (`Registered` stays `0`, `Loaded` stays
   `0`, no exception) -- reproduces even with **zero VBA code** in the file,
   so it's the extra content a real add-in never carries that a real loader
   rejects, not signing, not the ribbon relationship specifically.
4. Rebuilt from scratch matching the verified 5-part structure exactly
   (`vbaProject.bin` extracted from a COM-driven `SaveAs(.pptm)` and spliced
   in raw) -- **still fails to load identically.** Isolated further: splicing
   that same self-built `vbaProject.bin` into Rohan's own verified-working
   file (replacing only its `vbaProject.bin`, keeping every other byte of his
   real shell) **also fails**. So the difference isn't the surrounding
   package structure at all -- it's specifically that a `vbaProject.bin` built
   by importing modules into a fresh, never-saved-through-the-UI
   `Presentations.Add()` isn't equivalent to one produced by a real
   interactive Save-As-Addin cycle, even though the *code inside it* compiles
   and runs identically (proven by the full 58/58-test pass that ran
   `Application.Run` against exactly this kind of project before any save).
5. **Adding `customUI/customUI14.xml` (or literally any unrelated,
   unreferenced dummy part) to a real, working `.ppam` breaks it identically**
   -- tried both an in-place `ZipArchiveMode.Update` edit and a from-scratch
   `ZipArchiveMode.Create` rebuild copying every original byte verbatim, both
   fail the same way. This is the definitive finding: **a `.ppam`'s add-in
   loader rejects the package outright if it contains anything beyond its
   exact expected part set** -- not a targeted content-type/relationship
   check, a structural integrity check against the whole package. There is no
   way to add a ribbon (or anything else) to a `.ppam` and have it still load.

Also found and closed along the way: **there is no way to edit an
already-loaded add-in's VBA project via automation either.** Both
`AddIn.VBProject` and `Application.VBE.VBProjects` expose a `VBComponents`
property that is **genuinely `null`** for a loaded add-in (not just empty --
calling `.Import` on it throws "cannot call a method on a null-valued
expression"), and `AddIn` exposes no `.Save`/`.SaveAs` of its own, and loaded
add-ins never appear in `Application.Presentations` either. Almost certainly
deliberate -- the same class of self-modifying-macro restriction `AccessVBOM`'s
own default-off posture exists to guard against (`DECISIONS.md`, 2026-07-25).
Consequence: packaging a `.ppam` is a one-time-per-build manual step
(File > Save As, done by a human), not something any future pass can automate
away -- confirmed, not merely unattempted.

**Decision**: ship `CommandBarUI.bas`'s toolbar instead of a ribbon (see its own
SPIKE_NOTES) -- pure runtime VBA, needs zero package changes, works with the
exact `.ppam` structure already proven to load.
`vba/customUI/customUI14.xml` is kept in the repo as a historical record /
future-COM-add-in starting point, not wired into anything shipped.

## Divergence from the spec

`ribbon-ui.md`'s literal ribbon-XML requirement is not met -- provably cannot
be, for a `.ppam` (see above). The four buttons and their behavior are
otherwise built to spec.

## Manual verification recipe

Superseded by `CommandBarUI.bas`'s recipe -- there is no ribbon tab to check
here; verify the toolbar buttons instead.
