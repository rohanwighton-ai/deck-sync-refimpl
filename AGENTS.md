# Operational Learnings

This file contains project-specific guidance that Ralph has learned through observation.

Start minimal. Add entries only when Ralph exhibits repeated failures or needs specific
guidance.

## Build/Test Commands

- Tests: `python3 -m pytest tests/ -v`
- Type check: `python3 -m mypy src/`
- No pip packages beyond pytest/mypy are installed in this image. Prefer stdlib
  (`zipfile`, `xml.etree.ElementTree`) over adding new dependencies unless a spec
  genuinely can't be satisfied without one — this project deliberately stays
  dependency-light so it's trivially runnable.

## Known Patterns

- OOXML (`.pptx`) shape trees: groups are `<p:grpSp>`, not automatically flattened —
  must be walked recursively. See `src/discovery.py` for the reference walk.
- `xml.etree.ElementTree.Element.__bool__` is based on child count, not None-ness, and is
  deprecated for exactly this reason. A chain like `el.find(a) or el.find(b)` will
  silently pick the wrong result if the first match has no children (e.g. `<p:cNvPr
  name="..."/>` is a valid, real match with zero children). Always use explicit
  `is not None` checks per candidate, never `or`-chain `Element.find()` calls. (Hit and
  fixed during initial design, 2026-07-19 — see `src/discovery.py`'s `_shape_name`.)

## Constraints

- This is a reference/test implementation, not the production sync engine — the real
  target is VBA. Python code here exists to harden the discovery/matching/verification
  *logic* against a growing fixture corpus (see `test-fixtures/`), not to become a
  shipped tool. Don't add production concerns (CLI, packaging, distribution) unless a
  spec asks for it.
- Keep this repo separate from both `claude-brain` and any CRC system — see the initial
  commit message for why.
