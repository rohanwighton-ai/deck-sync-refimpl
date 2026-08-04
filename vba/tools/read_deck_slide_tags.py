#!/usr/bin/env python3
"""Read every slide's tags straight from the file, no Office.

Companion to read_deck_props.py, same reasoning: tags live in the .pptx's own
XML (ppt/tags/tagN.xml, one part per slide, referenced from that slide's own
_rels file by a relationship of type ".../relationships/tags") and are not
subject to the custom-document-property save bug DeckRegistry.bas documents --
but a rename that touches every slide's tags is exactly the kind of change
that should be checked against the file's own bytes, not trusted because the
write reported success in-process.

Tag NAMES are read case-insensitively: PowerPoint persists
`Tags.Add "slide_type", x` as `<p:tag name="SLIDE_TYPE" .../>` -- uppercased on
save, confirmed against a real slide -- so a caller asking for "slide_type"
would silently match nothing without this.

    python3 read_deck_slide_tags.py <deck.pptx> <tag-name>

Prints "slide-index: value" for every slide that carries the tag, then a
one-line summary count. Exit 0 always; an empty result is a legitimate answer,
not an error -- the caller decides what "0 slides" means.
"""

from __future__ import annotations

import re
import sys
from zipfile import ZipFile

RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
TAGS_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags"

REL_RE = re.compile(r'<Relationship[^>]*Type="([^"]+)"[^>]*Target="([^"]+)"[^>]*/?>')
TAG_RE = re.compile(r'<p:tag\s+name="([^"]*)"\s+val="([^"]*)"')


def slide_indices(z: ZipFile) -> list[int]:
    out = []
    for n in z.namelist():
        m = re.fullmatch(r"ppt/slides/slide(\d+)\.xml", n)
        if m:
            out.append(int(m.group(1)))
    return sorted(out)


def tag_parts_for_slide(z: ZipFile, slide_idx: int) -> list[str]:
    # A slide has ONE tags relationship PER TAG -- confirmed against a real
    # slide carrying six (SLIDE_TYPE, INSTANCE_KEY and others each get their
    # own tagN.xml part and their own Relationship entry). Taking only the
    # first, as an earlier version of this did, silently missed whichever tag
    # didn't happen to be first in file order.
    rels_path = f"ppt/slides/_rels/slide{slide_idx}.xml.rels"
    if rels_path not in z.namelist():
        return []
    rels_xml = z.read(rels_path).decode("utf-8")
    parts = []
    for rel_type, target in REL_RE.findall(rels_xml):
        if rel_type == TAGS_REL_TYPE:
            # Target is relative to ppt/slides/ -- e.g. "../tags/tag3.xml"
            part = "ppt/slides/" + target if not target.startswith("../") else \
                "ppt/" + target.split("../", 1)[1]
            parts.append(part)
    return parts


def tags_for_slide(z: ZipFile, slide_idx: int) -> dict[str, str]:
    merged: dict[str, str] = {}
    for part in tag_parts_for_slide(z, slide_idx):
        if part not in z.namelist():
            continue
        xml = z.read(part).decode("utf-8")
        for name, val in TAG_RE.findall(xml):
            merged[name.upper()] = val
    return merged


def main() -> int:
    if len(sys.argv) < 3:
        print("usage: read_deck_slide_tags.py <deck.pptx> <tag-name>", file=sys.stderr)
        return 1
    deck, wanted_name = sys.argv[1], sys.argv[2].upper()

    with ZipFile(deck) as z:
        matches = 0
        total = 0
        for idx in slide_indices(z):
            total += 1
            tags = tags_for_slide(z, idx)
            if wanted_name in tags:
                matches += 1
                print(f"slide {idx}: {tags[wanted_name]}")

    print(f"--- {matches} of {total} slide(s) carry tag '{wanted_name}' ---")
    return 0


if __name__ == "__main__":
    sys.exit(main())
