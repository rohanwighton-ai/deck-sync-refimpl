# VBA implementation: `AdoptFlow.bas`

The ribbon-facing entry point for `specs/deck-adoption.md`'s bulk retroactive
linking engine (`vba/DeckAdoption.bas`) -- deferred by that spec's own
Non-goals ("the UI itself... a future pass"), built once `CommandBarUI.bas`'s
toolbar existed to hang it on (2026-07-26). Mirrors `OnboardFlow.bas`'s
relationship to `Onboarding.bas` exactly: this module is pure interactive glue
(selection validation, the phase-gate review, instance-key prompts) --
`DeckAdoption.bas`'s `PlanAdoption`/`CommitAdoption` split does every real
decision, no new matching/adoption logic lives here.

**Executed against real Office 2026-07-26.** 4 tests pass. Full run: 65/65
across PowerPoint and Excel.

## Real bugs found and fixed this pass

1. **Whole-array `ByRef` assignment doesn't work the way this project
   otherwise assumes.** The first version of `ValidateAdoptionSelection`
   built a separate local array (`unsorted()`), sorted it, then did
   `outSlides = unsorted` at the end to hand it back through the `ByRef
   outSlides() As Object` parameter. Real-Office run: the caller's array came
   back genuinely unallocated (`UBound` raised "Subscript out of range"),
   not merely a style inconsistency. Fixed by `ReDim`-ing `outSlides`
   directly and populating it element-by-element in place, the same
   convention `PlanAdoption`'s own `harvestedValues()` out-param already
   uses. No other module here had tried the bulk-assignment shape before,
   so this is a new, confirmed addition to `AGENTS.md`'s Known Patterns
   territory, not a repeat of a logged gotcha.
2. **`Slides.Range(...).Select` needs Slide Sorter view to register as a real
   slide-type selection under COM automation.** In Normal view (where
   `NewBlankSlide()`'s own `View.GotoSlide` call leaves the test window),
   `Selection.Type` stayed `ppSelectionNone` after selecting a `SlideRange`
   -- the call itself raised no error, it just silently didn't take. Fixed in
   the test only (switch to `ppViewSlideSorter`, select, read `Selection`,
   switch back to Normal so later `NewBlankSlide()`-dependent tests are
   unaffected) -- confirmed this is an automation-only quirk, not a real-user
   limitation: a human Ctrl/Shift-clicking slide thumbnails in Normal view's
   own thumbnail pane does not have this problem, only headless `.Select()`
   with no genuine pane focus does.

## Design

`ValidateAdoptionSelection` sorts the selection into ascending `SlideIndex`
order -- `DeckAdoption.PlanAdoption`'s own documented precondition, since a
Ctrl-click multi-select is not guaranteed to already be in deck order.
`ExcludeTemplateSlide` removes the template (matched by `SlideID`, not object
identity) from the batch if the user's selection happened to include it --
adopting a whole child deck in one Select All will naturally include the
template slide itself. `BuildAdoptionReviewSummary` is the phase-gate display:
counts by disposition plus every non-`already_linked` slide's label and
reason, so a human can see what's being skipped and why before confirming.

Per `DeckAdoption.bas`'s own documented precondition, this flow requires an
**already-registered type** -- the greenfield "pick a template from scratch"
path is explicitly out of scope for adoption (use Onboard New Slide Type
first). `PromptAdoptExistingSlides` reuses `RibbonUI.BuildTypePickerPrompt`/
`ResolveTypeAnswer` for the type picker rather than duplicating that logic.

Instance-key prompts (one per "ready" slide) default to a blank `InputBox` --
a blank answer maps directly to `CommitAdoption`'s own `confirmedInstanceKeys(i)
= ""` semantics ("not yet confirmed, skip this slide"), so the prompt loop
needs no separate skip-handling of its own.

## Divergence from the spec

None of substance -- `specs/deck-adoption.md` already fully specified the
engine layer this wraps; this module is straight ribbon wiring per that
spec's own deferred UI bullet.

## Manual verification recipe

1. Onboard a type first (one example slide, via Onboard New Slide Type).
2. Add several more slides of the same visual layout, with real (varying)
   harvested content, none of them tagged.
3. Select all of them (the template slide is fine to include -- it gets
   filtered out automatically) and click "Adopt Existing Slides" on the
   toolbar (or run `Application.Run "AdoptFlow.AdoptExistingSlides"`).
4. Pick the type from the picker.
5. Confirm the phase-gate review summary matches what you'd expect (ready vs.
   needs-confirmation vs. unclassified counts).
6. Enter an instance key for each "ready" slide (or leave blank to skip).
7. Confirm the final report: linked count, already-linked count,
   excluded/unclassified/unconfirmed count, and (hopefully zero)
   failed-verification count.
8. Confirm in Excel: one new Data-sheet row per newly linked slide, correctly
   populated from harvested values.
