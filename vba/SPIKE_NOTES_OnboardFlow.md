# VBA implementation: `OnboardFlow.bas`

Implements `specs/ribbon-ui.md`'s "Onboard New Slide Type" flow, following
`onboard-slide-type.md`'s six steps exactly. The only genuinely new logic is
the phase-gate review this module builds -- `Discovery`, `Onboarding.
ConfirmFieldMatch`, `ExcelOutput`'s `CreateSheet`/`UpsertRow`, and
`InjectPrimitive` (the verify-the-link step) are all existing, already-tested
calls, unchanged.

**Executed against real Office 2026-07-26.** 6 tests pass, including a full
`OnboardFlow_CommitAndVerifyOnboardingRoundTrip` that creates a real slide with
a `ph_`-named text shape, plans onboarding, commits against a real throwaway
Excel workbook created via COM, and confirms `VerifyOnboarding` reports every
field on the no-op path.

## Design calls this module makes that no spec pins down

**The example slide becomes both the type's template AND its first live
instance** (carries `slide_type` + `instance_key`, not just field-level `role`
tags). Reasoning: Step 5 requires a seed Data-sheet row, and `ExcelOutput.
ReadSheet` deliberately excludes blank-Instance-ID rows (`DeckAdoption.bas`'s
own finding) -- a seed row with no `instance_key` would sit permanently
invisible to `RunRoutineSync`'s normal case-3 dispatch, needing new "keyless
row" handling `DeckAdoption.bas` already had to invent for a different reason.
Giving the template its own `instance_key` avoids inventing that a second time:
it becomes a completely ordinary tagged instance like any other, and
`RunSync.GatherInstances`/`DuplicateAndTag` need no special-casing (a template
that also happens to be instance #1 is not a contradiction anywhere else in
this codebase).

**Seed instance key derivation**: the period-key field's harvested value
(sanitized, spaces to dashes) if one was marked, else the fixed key
`"evergreen"` (an evergreen type by definition only ever has one instance, so a
fixed key is sufficient and stable).

**InputBox-based, not a UserForm ListBox** -- same posture as `ResolveFields.bas`,
for the same reason: no established, real-Office-verified precedent in this
project for authoring a `.frm`/`.frx` pair. A real Office session was available
this pass (unlike when `ResolveFields.bas` was built), so this constraint could
have been lifted, but building and proving a first-ever UserForm was
deliberately kept out of an already-large pass -- flagged as a real, deferred
upgrade, not an oversight.

## Divergence from the spec

Supports exactly 1 example slide, not the spec's "1-2." The 2-slide case (cross-
checking two examples of the same type) is a real scope reduction, not
silently dropped -- flagged here as deferred, not decided against.

## Manual verification recipe

1. Build a slide with 1-2 `ph_`-named text shapes carrying real content.
2. Select it, click "Onboard New Slide Type" (or run
   `Application.Run "RibbonUI.OnboardNewType"` directly).
3. Walk through the InputBox chain: type name, per-field rename/exclude/accept,
   period-key selection, final Yes/No confirm, workbook path if this is the
   deck's first onboarding.
4. Expect a result report: field count, seed instance key, and a per-field
   verify-the-link report -- every field should show `OK`, not `FAIL`.
5. Confirm in Excel: a new sheet (or existing one) now has a header row plus
   one seed data row under the reported instance key.
6. Confirm in PowerPoint: the original slide selection is untouched (Step 1's
   "never write to the original" invariant) -- only the duplicate carries the
   new tags.
