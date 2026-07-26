# VBA implementation: `BatchOnboardFlow.bas`

Onboards a new slide type AND bulk-links a whole batch of its already-existing
instances in one pass, driven by an editable Excel review grid instead of
InputBox chains. Built 2026-07-26 specifically because a real, richly-designed
deck (`test-fixtures/crc-real-deck-redacted.pptx`, and the unredacted
`test.pptx` used for the actual live run) discovers 60-90 raw candidate shapes
per slide — walking that many one InputBox at a time (`OnboardFlow.bas`'s
existing flow) is unusable. Also closes `specs/onboarding.md`'s own
documented-but-never-built "boilerplate-vs-varying pre-filter" gap, in a
different shape than that spec sketched: rather than a silent geometric
clustering filter, this surfaces a computed guess (diffing harvested text
across the whole batch) and lets a human freely override it in a real,
editable table (Rohan's own framing, 2026-07-26).

**Executed against real Office 2026-07-26. 70/70 tests pass** (up from 61 at
the start of this module's work). Getting there took an unusually long,
genuinely difficult debugging session — full account below, because the root
causes are exactly the kind of thing worth not re-discovering blind next time.

## The debugging saga — three separate real bugs, not one

The symptom throughout was consistent and misleading: **"Compile error:
User-defined type not defined"**, with VBE's error navigation landing on
`BuildBatchPlan`'s own signature line (or occasionally a caller's `Dim plan As
BatchOnboardPlan` line) almost every time — which kept pointing investigation
at the wrong place. Three genuinely distinct bugs were hiding behind that one
symptom:

### Bug 1 (real, found by a Fable-model agent): UDT array stored in a Dictionary

`BuildBatchPlan` originally cached each other-slide's discovered candidates in
`Scripting.Dictionary` objects, one cached per slide index, to avoid
re-scanning a slide once per field:
```vb
Dim oc() As Candidate
oc = Discovery.DiscoverSlideWithShapes(otherSlides(s), os_)
Set otherCandidatesBySlide(s) = oc   ' <- illegal
```
A `Candidate()` array (`Candidate` is a `Public Type` in `Discovery.bas`, a
standard module) can **never** be stored as a `Dictionary`/`Variant` value —
VBA's compiler rejects it outright: *"Only user-defined types defined in
public object modules can be coerced to or from a variant or passed to
late-bound functions."* This is documented in `AGENTS.md`'s Known Patterns
already (found 2026-07-25 in `Onboarding.bas` and `SyncOperations.bas`) — this
module hit it a third time, having never checked that section before writing
new Dictionary-caching code.

**Fix**: flattened all other-slides' `Candidate()`/`Object()`/`Boolean()` data
into single plain, properly-typed arrays (`allOtherCandidates()`,
`allOtherShapes()`, `allOtherAvailable()`) instead of caching per-slide in a
Dictionary, with `otherSlideCandStart()`/`otherSlideCandCount()` (plain
`Long()` arrays — primitives, no restriction) recording where each slide's own
candidates sit within the flattened set. Each slide is still discovered only
once.

This was a real, necessary fix — but fixing it did **not** make the "type not
defined" error go away. That turned out to be a second, unrelated bug.

### Bug 2 (the actual cause of the persistent symptom): declaration order

Real, confirmed VBA compiler quirk, found by bisecting a long series of
minimal, isolated repros (not a guess — each step below was independently
verified live against real Office): **if a `Public Function`/`Sub` appears
textually *before* a `Public Type` in the same standard module, that Type
fails to resolve when referenced from a *different* module** — "Compile
error: User-defined type not defined" — even though:
- the function containing the "bad" reference compiles and runs fine when
  called directly with real arguments,
- the Type itself, alone with a trivial function, works cross-module,
- the Type with all its real members (6 `Object` Dictionaries) works
  cross-module,
- and same-module usage of the Type always works regardless of order.

Only the combination of (Type declared after another Function) + (referenced
from a *different* module) reproduces it. `BatchOnboardFlow.bas` originally
had `AllValuesIdentical`/`SuggestBatchFieldName` declared before `Public Type
BatchOnboardPlan`, and `WriteReviewGrid`/`ReadReviewGrid`/etc. before `Public
Type BatchCommitResult` — both real instances of this ordering.

A **second, apparently-related** symptom of the same underlying rule: the six
`Private Const COL_*` declarations (originally sitting between `BuildBatchPlan`
and `WriteReviewGrid`) produced a *different* error once Bug 2's Type-ordering
fix was applied — **"Only comments may appear after End Sub, End Function, or
End Property"** — also resolved by moving them above every Function/Sub.

**Fix**: every module-level declaration (`Const`, `Type`) in this module now
lives at the very top, above all `Function`/`Sub` declarations — not a style
preference, a hard requirement this module discovered the expensive way.
**Recommendation for every other module in this project going forward**: keep
all `Const`/`Type` declarations first, before any `Function`/`Sub`, even
though the rest of this codebase (written before this was known) doesn't
consistently follow that — nothing has broken elsewhere yet, but this is now
a known trap, not a hypothetical one.

### Bug 3: a reserved word and a test-data bug (ordinary, once bugs 1-2 were found)

`vba/tests/TestRunner.bas` had `Dim single As New Collection` — `Single` is a
VBA reserved word (the floating-point type), producing a plain "Syntax error."
Renamed to `onlyOne`.

Separately, `Test_BatchOnboardFlow_CommitBatchTagsLinksAndVerifies` used
identical harvested text ("Overall Status") on both the template and the one
other slide — `BuildBatchPlan` correctly classified this as decoration
(`FieldSuggestIdentical = True`) and defaulted `FieldInclude` to `False`,
so `CommitBatch` correctly tagged nothing. The test's own assertions expected
tagging to happen, which was a test bug, not a `CommitBatch` bug — fixed by
explicitly setting `plan.FieldInclude(1) = True` before committing (simulating
a human overriding the suggestion in the review grid), with new assertions
added confirming the *suggestion* itself was correct first.

## Diagnostic technique that actually worked

Manual live probing (import, run, read the stuck VBE dialog) got the search
going but produced inconsistent, misleading signal for hours — the error's
reported location moved around depending on which unrelated test scaffolding
happened to be in play, which is exactly what Bug 2's "wrong module gets
blamed" behavior produces. What actually closed it:
1. A Fable-model agent, briefed with everything already ruled out, found Bug 1
   via its own from-scratch minimal repro (see its report for the exact
   isolation technique).
2. UI Automation (`System.Windows.Automation`) reading the VBE's exact
   selected/highlighted text directly, rather than relying on a human
   describing a screenshot — this is what made rapid, precise bisection of
   Bug 2 possible (dozens of tiny `TypeDefModule`/`CallerModule` pairs,
   isolating exactly which combination of same-module-vs-cross-module and
   before-vs-after-the-Type broke it).
3. Systematic reduction: minimal reproductions built up in small, deliberate
   increments (bare Type → Type+trivial-function → Type+real-body →
   Type+real-body+other-functions), each tested in complete isolation before
   moving to the next increment, rather than continuing to prod the full
   ~20-module production project.

## Manual verification recipe

1. Select 2+ slides of a genuinely repeated layout (a template plus at least
   one real instance) — the unredacted `test.pptx` (kept local, never
   committed) is the actual target deck this was built for.
2. Click "Bulk Onboard Type" on the toolbar (or run
   `Application.Run "BatchOnboardFlow.BatchOnboardType"`).
3. Name the type. Confirm the "Field Review" Excel sheet appears with one row
   per discovered field, a computed Suggested classification, and sample
   values from other slides.
4. Edit the sheet: rename fields, flip Include Y/N, exclude anything wrong.
   Confirm.
5. Provide an instance key for the template (required) and each other slide
   (blank = skip).
6. Confirm the final report: linked count, skipped count, and (should be
   zero) failed-verification count.
7. Confirm in Excel: one Data-sheet row per linked slide, correctly populated
   from harvested values, only for included fields.
8. Confirm in PowerPoint: only included fields got role tags; every selected
   slide (not just the template) carries `slide_type`/`instance_key` if it
   got a real instance key.
