"""Integration layer: resolve a real, already-tagged slide into the
SlideInstance shape sync_operations.py needs.

Every module up to this point (discovery, matching, verification,
excel_output, sync_operations, identity_tags) was built and tested in
isolation -- individually correct, but nothing composed them against a real
deck. This module is that composition, for the one case sync_operations.py
actually needs: an *already onboarded* slide (tags already written via
identity_tags.upsert_slide_tags/upsert_shape_tags). It intentionally does
not implement onboarding itself -- matching a subsequent slide's untagged
candidates against a reference (matching.py's tier-2 path) is
onboarding.py's job. The "per-type reference configuration" that scoring
needs isn't a separate format this module would have to invent: it's just
another already-onboarded slide of that type, resolved through
resolve_slide_instance() itself, and passed to
onboarding.match_slide_against_template() as its `template` argument.
"""

from __future__ import annotations

from discovery import discover_from_pptx_part
from identity_tags import read_shape_tags, read_slide_tags
from sync_operations import SlideInstance


def resolve_slide_instance(path: str, slide_part: str) -> SlideInstance:
    """Resolve `slide_part` (an already-tagged slide) into a SlideInstance:
    read its slide-level tags (slide_type, instance_key) and, for each
    discovered shape, its shape-level tag (role) -- tier-1 trust per
    matching.md, since a shape reaching this function is expected to already
    carry a tag from onboarding. An untagged shape is simply not added to
    field_shapes (sync_operations.py already treats a Data-sheet field with
    no corresponding shape as "nothing to inject", not an error).
    """
    slide_tags = read_slide_tags(path, slide_part)
    candidates = discover_from_pptx_part(path, slide_part)

    field_shapes = {}
    for candidate in candidates:
        shape_tags = read_shape_tags(path, slide_part, candidate)
        role = shape_tags.get("role")
        if role is not None:
            field_shapes[role] = candidate

    return SlideInstance(
        part_path=slide_part,
        instance_key=slide_tags.get("instance_key"),
        type_tag=slide_tags.get("slide_type"),
        field_shapes=field_shapes,
    )
