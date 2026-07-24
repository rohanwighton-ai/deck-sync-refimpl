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

## Cost / Tool Selection

- This repo is small (specs/ and src/ are each under a dozen short files). A direct
  `Read` is strictly cheaper than a subagent spawn for a file this size — every subagent
  is a fresh, uncached API call, and pays a full context-establishment cost to read
  something a direct Read gets for a fraction of that. Reserve parallel subagents for
  genuinely large fan-out (dozens-to-hundreds of files, or slow independent searches),
  not as a default "study the codebase" step. (Root-caused 2026-07-25: an earlier
  version of `PROMPT_build.md`/`PROMPT_plan.md` authorized "up to 500 parallel Sonnet
  subagents" for exactly this small a repo, which was almost certainly the dominant
  cost driver during real iterations, independent of the retry-storm bug logged below.)
- If a task's obvious approach would be expensive relative to what it accomplishes, say
  so — in the commit message or a plan note — rather than silently paying the cost.
  Flagging "this would cost too much, here's a cheaper way" is a valid, wanted outcome
  of an iteration, not a failure to complete the task.

## Constraints

- This is a reference/test implementation, not the production sync engine — the real
  target is VBA. Python code here exists to harden the discovery/matching/verification
  *logic* against a growing fixture corpus (see `test-fixtures/`), not to become a
  shipped tool. Don't add production concerns (CLI, packaging, distribution) unless a
  spec asks for it.
- Keep this repo separate from both `claude-brain` and any CRC system — see the initial
  commit message for why.
- `vba/InjectPrimitive.bas` was a single de-risking spike, not yet governed by a spec
  the way every `src/*.py` file is. Further VBA porting (discovery, matching,
  sync-dispatch logic) is planned and now has a governing spec, `specs/vba-port.md` —
  written 2026-07-24, mirroring how every other module here began from a spec rather
  than growing VBA scope ad hoc after the fact. Follow its port order and
  manual-verification-recipe requirement for every subsequent VBA module.
