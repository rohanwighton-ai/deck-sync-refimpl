# VBA Port

Port the reference implementation's logic (`src/discovery.py`, `matching.py`,
`verification.py`, `identity_tags.py`, `resolve.py`, `sync_operations.py`,
`onboarding.py`) to native VBA, so the decision logic proven against fixtures in Python
actually runs against a real open PowerPoint deck and Excel workbook. This is the
production engine; the Python reference implementation stays exactly what it is
(logic-hardening only, see `AGENTS.md`) and does not get retired.

## Requirements

- **Prefer native object model over OOXML surgery wherever VBA has a native path.**
  `vba/InjectPrimitive.bas` already established the pattern for identity tags: Python
  reverse-engineers the Tags Part XML by hand (no library support), but VBA has
  `Shape.Tags`/`Slide.Tags` natively — use those directly, don't port the Python XML
  approach. Apply the same judgment per module: if `Shapes`/`GroupShapes`/`Range` gives
  a direct native equivalent, use it; only fall back to raw OOXML (via `Shell`-invoked
  Python, a zip library, or manual XML) where VBA's object model genuinely has no path,
  and flag that fallback explicitly when it happens.
- **Port order, each a separate module with its own manual verification notes** (mirror
  `vba/SPIKE_NOTES.md`'s format — scope, divergences from the Python spec, a step-by-step
  recipe a human runs in real Office):
  1. `discovery` — walk a slide's `Shapes` collection recursively into `GroupShapes`,
     matching `src/discovery.py`'s candidate rules (non-empty text or picture, never
     filtered by shape type, full metadata capture) exactly.
  2. `identity_tags` — already done (`InjectPrimitive.bas`'s `Shape.Tags`/`Slide.Tags`
     reads). Confirm the existing spike's read/write primitives cover
     `identity-tags.md`'s upsert semantics fully; extend only if a gap is found.
  3. `matching` — tier-1 tag trust and tier-2 geometry/text scoring per `matching.md`,
     operating on native `Shape.Left/Top/Width/Height` instead of Python's raw
     `a:off`/`a:ext` XML reads.
  4. `resolve` + `sync_operations` (cases 1/3/4/6 only — case 2 is
     `plan_period_rollover`, separately invoked, never inferred; cases 5/7 are not yet
     specified, see Non-goals) — compose discovery + identity_tags + matching into the
     same dispatch decision tree as `src/resolve.py`/`sync_operations.py`, writing via
     `InjectPrimitive`'s existing primitive.
  5. `onboarding` — port `match_slide_against_template`/`onboard_new_instance`/
     `confirm_field_match` per `onboarding.md`. `confirm_field_match` is the one module
     boundary that a real UI will call into (see Non-goals) — build the VBA function
     itself, not the UI.
  6. Excel-side reads/writes (`excel_output.py`'s Python equivalent) — this is the one
     module where VBA is strictly simpler than the Python port it's replacing: Python
     rebuilds `.xlsx` zip parts by hand because it has no host application; VBA runs
     *inside* Excel (or drives it via COM from the PowerPoint side), so this is plain
     `Range`/`Cells` reads and writes, no XML. Don't port `excel_output.py`'s
     zip-rebuilding approach — it only existed to work around Python lacking a live
     Excel instance.
- **No test harness exists for VBA the way `pytest`/`mypy` do for Python** — this
  environment has no Office/COM to execute against (confirmed: the existing spike is
  "not executed... needs manual verification on Windows," per `vba/SPIKE_NOTES.md`).
  Each ported module must ship with its own `SPIKE_NOTES.md`-style manual recipe;
  building a module without one is not complete, whatever the code looks like.
- **Fixture parity**: every module's manual verification recipe should exercise it
  against the same `test-fixtures/*.pptx` files the Python side already validated
  against, so a human running the recipe is confirming equivalent behavior, not just
  "it doesn't crash."

## Non-goals (out of scope for this spec)

- Cases 5 (`record_retired`) and 7 (`deck_side_conflict`) — no agreed convention/store
  exists yet for either (see `sync-operations.md`'s existing Non-goals). Do not invent
  one here; these need a dedicated spec once the underlying product decision is made.
- Physical slide duplication — confirmed moving to VBA's native `Slide.Duplicate`
  per `vba/SPIKE_NOTES.md`, but the trigger semantics (when duplication fires, from
  what input) aren't decided. Separate spec, pending decision.
- Packaging, distribution, installer format, signing/trust — separate spec, pending
  decision on how this is actually shipped and run.
- The shape-selection UI a human would use to resolve medium/low-confidence matches or
  onboard a new slide type — this spec builds the VBA functions such a UI would call
  (`confirm_field_match`, `onboard_new_instance`), not the UI itself. Separate spec.

## Reference

- Each numbered module above traces to its own existing spec (`discovery.md`,
  `identity-tags.md`, `matching.md`, `sync-operations.md`, `onboarding.md`) — this spec
  governs *how* to port that already-decided logic to VBA, not what the logic is.
- `vba/InjectPrimitive.bas` + `vba/SPIKE_NOTES.md` is the reference pattern and style for
  every module this spec covers: native object model first, explicit scope/divergence
  notes, a manual recipe a human can actually run.
