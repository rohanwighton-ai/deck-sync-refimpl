# Test fixtures for crc-vba-deck-sync

Pulled 2026-07-19 from `scanny/python-pptx` (MIT license), which is python-pptx's own
test suite — not a realistic business deck, narrow structural test files chosen to
exercise specific mechanisms designed that day.

- `shp-groupshape.pptx` — grouped/nested shapes. Exercises discovery_scope's
  group-recursion rule and the sibling_ambiguity / z-order checks in
  `skills/crc-vba-deck-sync/references/shape-identity-and-matching.md`.
- `mst-slide-layouts.pptx` — multiple slide layouts and placeholders. Exercises
  placeholder-index matching (untagged_fallback tier 1 signal) and multi-type discovery.

Source: https://github.com/scanny/python-pptx/tree/master/features/steps/test_files

## `crc-real-deck-redacted.pptx`

Added 2026-07-25. A real 46-slide CRC project-reporting deck (Rohan's actual SAAFE
work), redacted before commit: every unique text run was replaced with a deterministic
placeholder (`REDACTED_NNNNN`, same original string -> same placeholder everywhere, so
which shapes repeat identically across "instances" vs. genuinely vary is preserved
exactly), every media file replaced with a tiny placeholder image, and doc-props
author/creator fields scrubbed. Verified before commit: `discover_from_pptx()` returns
identical candidate counts per slide on the original vs. redacted file, and the
identical-vs-varying text-run pattern between two same-type slides (92/96 runs
identical between slide 2 and slide 3) survives redaction exactly. No real partner
names, project content, or "(withheld)" status text present in the committed file —
checked via grep across the unzipped parts before commit.

**Why this fixture matters, and the gap it surfaced:** this is the first genuinely
real, organically-designed deck run through `discover()` — every other fixture here is
a narrow structural test file. Result: 60-90 "candidate fields" per slide (vs. a
handful in the synthetic fixtures), because a richly-designed slide has dozens of
text-bearing shapes that are static design chrome (section headers, icon labels,
legend text), not real dynamic fields — `discovery.md`'s "has text or is a picture"
rule is correct but doesn't scale to real content without a further filter. This is
the concrete, real-deck evidence behind `DECISIONS.md`'s 2026-07-19 "not validated
against an actual organically-drifted CRC deck" uncertainty, and motivates adding a
cross-instance boilerplate-vs-varying filter to `onboarding.md`/`deck-adoption.md`'s
phase-gate review before a human is ever shown a candidate list.
