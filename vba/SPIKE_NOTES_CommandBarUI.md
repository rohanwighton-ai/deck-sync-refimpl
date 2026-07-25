# VBA implementation: `CommandBarUI.bas`

The actual shipped UI surface for `specs/ribbon-ui.md`, after a real
`customUI14.xml` ribbon was proven impossible for a `.ppam` add-in -- full
account of why in `SPIKE_NOTES_RibbonUI.md`. `CommandBars` is the pre-Ribbon
(pre-Office-2007) VBA UI mechanism, still fully functional today specifically
because it's pure runtime code, not an OOXML package part -- it needs zero
changes to the `.ppam` structure already proven to load, sidestepping the
entire packaging dead-end.

**Executed against real Office 2026-07-26**, both as unit tests and as a full
live add-in load: 3 tests pass (create-with-4-wired-buttons, idempotent
re-call, remove). Then proven end-to-end for real: imported all 17 production
modules into a fresh presentation, Rohan did File > Save As > PowerPoint
Add-in (`Addin2.ppam`), loaded via `AddIns.Add` + `.Loaded = True` --
**`Auto_Open` fired automatically and the "Deck Sync" toolbar appeared with all
4 buttons, correctly wired** (`Caption`/`OnAction` verified via COM, and
visually confirmed by Rohan). This is the first real, complete, working proof
of the whole ribbon-ui.md flow being reachable by an actual user, not just
callable from the VBE.

## Design

`Auto_Open`/`Auto_Close` are the classic legacy add-in lifecycle Subs --
PowerPoint still runs `Auto_Open` automatically when an add-in is loaded
(`AddIns.Add` + `.Loaded = True`) and `Auto_Close` when unloaded, confirmed
directly this pass (not assumed from documentation). `ShowToolbar` is
idempotent (deletes any stale toolbar of the same name first) and doubles as a
manual fallback entry point (same `ManualSmokeTest` convention every other
module uses) in case `Auto_Open` didn't fire, e.g. code added to an add-in that
was already loaded before this module existed.

Modern (Ribbon-era) Office hosts a VBA-created `CommandBar` toolbar under its
built-in "Add-Ins" tab. Less visually polished than a dedicated branded ribbon
tab with large icons, but a real, clickable, testable surface -- confirmed
visually, not just structurally.

`faceId` values (the `CommandBarButton` equivalent of `imageMso`) are
best-guess built-in icon indices, not verified against a live render -- an
unresolved `faceId` shows a blank/default icon, cosmetic risk only, same
posture `customUI14.xml`'s `imageMso` guesses carried.

## Divergence from the spec

Visual form only -- a classic toolbar under "Add-Ins," not a dedicated Fluent
ribbon tab. Every button, its label, and its behavior match `ribbon-ui.md`
exactly; only the chrome differs, and only because the chrome `ribbon-ui.md`
specified turned out to be unbuildable for this package format.

## Manual verification recipe

1. Rebuild the add-in per `vba/tests/build_ppam.ps1`'s instructions (import
   all production modules into a fresh presentation, then File > Save As >
   PowerPoint Add-in manually).
2. Load it: `Application.AddIns.Add(path)` then set `.Loaded = True` (or via
   File > Options > Add-ins > PowerPoint Add-ins > Add in a real interactive
   session).
3. Confirm a "Deck Sync" toolbar appears (typically docked near the "Add-Ins"
   tab) with 4 buttons: Sync Now, New Period, Onboard New Slide Type, Resolve
   Unmatched Fields.
4. Click each and confirm it invokes the matching flow (Sync Now/New Period
   will report "no paired workbook yet" on a fresh deck, which is the correct,
   expected response before any onboarding has happened).
5. Unload the add-in (`.Loaded = False` or via the same Add-ins dialog) and
   confirm the toolbar disappears (`Auto_Close`).
