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
