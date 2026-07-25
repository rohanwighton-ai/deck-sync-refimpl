# VBA implementation: `SlideDuplication.bas` (specs/slide-duplication-trigger.md)

No Python equivalent exists -- `sync_operations.py`'s own Non-goals
explicitly excluded physical slide duplication ("this isn't needed at all
in the real VBA target, where duplication is one native call
[`Slide.Duplicate`]"). `specs/slide-duplication-trigger.md` is the
governing spec for what this module implements; this is its first (and so
far only) implementation, written and executed against real Office the
same day (2026-07-25), via `vba/tests/TestRunner.bas`'s
`Test_SlideDuplication_*` functions.

## What was built

`SlideDuplication.DuplicateAndTag(sourceSld, slideType, newInstanceKey,
values, existingInstances)`:

1. **Instance-key collision guard** -- refuses (returns `Ok=False`, no
   slide created) if any of `existingInstances` already carries
   `newInstanceKey`.
2. **`sourceSld.Duplicate()`** -- placement correction (row-order per the
   spec) is deliberately not applied here; see `RunSync.bas`.
3. **Mandatory structural/z-order verification** (`Verification.
   VerifyStructure`/`VerifyZOrder`) before any tag write.
4. **Tags `slide_type`/`instance_key` unconditionally**, only after
   verification passes.
5. **Injects `values` into the duplicate's tagged fields**, per the
   template's own field set (`Onboarding.BuildTemplateFieldShapes` against
   `sourceSld`) -- a field the template defines but `values` doesn't supply
   is flagged in `MissingFields()`, never silently left blank with no
   signal, and never withholds creating the slide.

## A real judgment call the spec doesn't fully pin down

**A failed verification deletes the malformed duplicate** (`newSld.Delete`)
rather than leaving it sitting untagged in the deck. The spec says "a
malformed duplicate must never receive an instance_key" but doesn't say
what happens to the duplicate itself. Leaving it in place, untagged, would
be confusing debris a human has to notice and clean up by hand; deleting it
keeps the deck in the same state as if `DuplicateAndTag` had never run at
all for that call. A different, equally defensible choice (leave it,
visibly flagged some other way, for a human to inspect) was not made.
Flagging this explicitly since it's exactly the kind of decision this
project's own practice is to surface, not bury in a diff.

## Deliberate divergences / design notes

- **Collision-checking instances are caller-supplied**, not gathered here
  -- matches the non-goal every other module in this port draws for itself
  (`sync_operations.py`'s "gathering instances is the caller's job").
  `RunSync.bas` is that caller.
- **The type's field set comes from re-running `Onboarding.
  BuildTemplateFieldShapes` against the source**, not from a separately
  passed-in list -- avoids a second, possibly-inconsistent way of answering
  "what fields does this type have."

## What was deliberately left out of scope

- Row-order placement -- a whole-type, driver-level concern per the spec
  ("Order is an enforced invariant... not just placing newly-created slides
  correctly"), owned by `RunSync.ResequenceByRowOrder`, not this module.
- Case 5/7, the "add new quarterly record" invocation mechanism, picture-field
  injection semantics, and all multi-deck concerns -- all already Non-goals
  in `specs/slide-duplication-trigger.md` itself.

## Verification

Automated: `Test_SlideDuplication_CreatesTaggedInjectedSlide`,
`Test_SlideDuplication_RefusesInstanceKeyCollision`,
`Test_SlideDuplication_PartialRowStillCreatesSlideButFlagsMissing`, all
passing for real via `run_vba_tests.ps1`. No separate manual recipe written.
