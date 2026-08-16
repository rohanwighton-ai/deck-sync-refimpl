# VBA port: `verification` module (structural/z-order checks)

Not part of `specs/vba-port.md`'s original 6-module port order -- that order
only ever covered `inject_primitive` (via `InjectPrimitive.bas`), never
`src/verification.py`'s `verify_structure()`/`verify_z_order()`. Those two
checks became necessary once `specs/slide-duplication-trigger.md` made them
mandatory before tagging any duplicate ("A malformed duplicate must never
receive an instance_key"). Ports `verify_structure()`/`verify_z_order()`
(see `specs/verification.md`) to native VBA.

**Executed against real Office the same day it was written (2026-07-25)**,
via `vba/tests/TestRunner.bas`'s `Test_Verification_*` functions, driven by
`vba/tests/run_vba_tests.ps1`. All 3 tests pass for real: a clean
`Slide.Duplicate` reports `Ok=True` on both checks, a shape deleted from
the duplicate is caught as a structural mismatch, and a z-order swap on the
duplicate (same shapes/tags/types, different stacking) is caught by
`VerifyZOrder` while `VerifyStructure` correctly reports the structure
itself as unaffected -- confirming the two checks are genuinely independent
claims, not one inferring the other.

## What was ported

- `Verification.VerifyStructure(sourceSld, duplicateSld)`: shape count,
  type, and identity-tag correspondence between two live slides, tagged
  shapes paired by role tag, untagged shapes falling back to positional
  pairing within just the untagged subsequence -- field-for-field against
  `verify_structure()`.
- `Verification.VerifyZOrder(sourceSld, duplicateSld)`: every pair of
  commonly-tagged shapes checked for relative stacking-order agreement
  between source and duplicate (not just adjacent pairs) -- field-for-field
  against `verify_z_order()`.

## A real gap this module had to close first

`Discovery.Candidate.IdentityTag` is always `""` straight out of
`DiscoverSlide`/`DiscoverSlideWithShapes` (discovery does not read tags,
per `discovery.md`'s non-goals -- true of the Python original's `discover()`
too). But tag-based pairing is the entire reason these two checks exist (a
pure z-order swap must never also register as a structural defect, which
positional pairing alone cannot distinguish from a real mismatch). Both
functions do their own `DiscoverSlideWithShapes` + live `Shape.Tags("role")`
read internally -- mirroring what `Onboarding.BuildTemplateFieldShapes`
already does for the identical underlying reason -- rather than trusting a
pre-populated `IdentityTag` field that never actually gets set anywhere in
this port.

## Deliberate divergences from the Python semantics

1. **Takes live `Slide` objects, not pre-discovered `Candidate` arrays.**
   `verify_structure()`/`verify_z_order()` take `Sequence[Candidate]`
   directly in Python because the caller already has them from a prior
   `discover()` call over a file. Here, since tag-reading has to happen
   internally anyway (see above), it was simpler and safer for both
   functions to own their own discovery pass than to require callers to
   pre-build tagged `Candidate` arrays correctly.
2. **`Dictionary(tag As String -> index As Long)` replaces Python's
   `dict[str, Candidate]`.** `Long` is a primitive, Variant-safe value; a
   `Candidate` (UDT) cannot go into a Dictionary at all -- see
   `SPIKE_NOTES_Onboarding.md`'s original finding and `AGENTS.md`'s Known
   Patterns. Indices into the already-discovered `Candidate()` array are
   stored instead of the candidates themselves.
3. **`StructuralVerification`/`ZOrderVerification`'s `Mismatches()` arrays
   carry an explicit `MismatchCount` field alongside them, not just relying
   on `LBound`/`UBound`.** Learned from the project-wide `ReDim(1 To 0)`
   finding (`AGENTS.md`'s Known Patterns): an empty result means the array
   stays genuinely unallocated, and every new UDT in this pass that can be
   "empty" carries its own count field so callers never need an
   error-guarded bounds check just to know whether anything happened.

## What was deliberately left out of scope

- Deciding what to do about a failed verification -- `specs/verification.md`'s
  own non-goal; that's `SlideDuplication.bas`'s job (delete the malformed
  duplicate, refuse to tag it).

## Verification

No separate manual recipe written -- the automated tests in
`vba/tests/TestRunner.bas` (`Test_Verification_StructureMatchesAfterDuplicate`,
`Test_Verification_DetectsShapeCountMismatch`,
`Test_Verification_DetectsZOrderSwap`), run via `run_vba_tests.ps1`, already
constitute real, repeatable, stronger evidence than a hand-walked recipe
would be. Re-run with `powershell.exe -File <run_vba_tests.ps1, via
wslpath -w>`.
