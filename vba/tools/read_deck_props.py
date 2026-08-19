#!/usr/bin/env python3
"""Read a deck's registry properties straight from the file, no Office.

THE AUTHORITATIVE CHECK, and the reason it exists is specific.

`E2EField.SetPeriod` verifies a write by closing the presentation and reopening
it -- but it reopens inside the SAME PowerPoint instance, which can serve a
cached copy rather than reading disk. So a save that failed can verify as
successful, and a save that worked can verify as failed. Neither direction is
trustworthy, which makes the check worthless exactly when it matters.

On 2026-07-31 three consecutive period writes gave: fail, succeed, fail, with no
pattern anyone could explain. Reading the file's own bytes settled it in one
command: the deck genuinely held the value from the successful run, so the save
really does fail sometimes and the in-process check was right that time and
possibly wrong the others.

FIXED 2026-08-20: this tool read ONLY `docProps/custom.xml`, which
`DeckRegistry.bas`'s own header describes as "THE OLD MECHANISM, read-only
fallback now" since 2026-08-16 -- a cloud-hosted deck's first
CustomDocumentProperties write of a session sticks and every write after it,
to ANY of these properties, is stuck reporting that first value forever, so
the real values now live as plain text on a hidden slide named
`DeckSyncRegistry`, keyed by shape NAME (`DeckRegistry.ReadStringProperty`:
registry slide first, `docProps/custom.xml` only as a fallback for decks
never touched since the migration). This tool never followed that move, so
every check made with it read stale or wrong values without any error --
same failure shape as `WorkbookBridge.DescribeSheet`/`LifespanOf` drifting
out of step with the sheet roster the same day. Registry-slide values now
win; `docProps/custom.xml` is consulted only for a key absent from the
registry slide, mirroring `ReadStringProperty` exactly.

Deliberately NOT importable from VBA: the point is to check the writer using
something that shares no process, no cache and no code with it.

    python3 read_deck_props.py <deck.pptx> [property-name]

Exit 0 if found (or if listing), 1 if a named property is absent.
"""

from __future__ import annotations

import re
import sys
import zipfile
import xml.etree.ElementTree as ET

PART = "docProps/custom.xml"
# <property ... name="X"><vt:lpwstr>VALUE</vt:lpwstr></property>
PROP_RE = re.compile(r'name="([^"]+)"[^>]*>\s*<[^>]+>([^<]*)<')

REGISTRY_SLIDE_NAME = "DeckSyncRegistry"
P_NS = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def read_registry_slide(z: zipfile.ZipFile) -> dict[str, str]:
    """Shape NAME -> shape text, on the hidden slide named DeckSyncRegistry.
    Mirrors DeckRegistry.bas's REGISTRY_SLIDE_NAME mechanism exactly."""
    for name in z.namelist():
        m = re.fullmatch(r"ppt/slides/slide(\d+)\.xml", name)
        if not m:
            continue
        root = ET.fromstring(z.read(name))
        cSld = root.find(f"{P_NS}cSld")
        if cSld is None or cSld.get("name") != REGISTRY_SLIDE_NAME:
            continue
        shapes: dict[str, str] = {}
        for sp in root.iter(P_NS + "sp"):
            nvpr = sp.find(f"{P_NS}nvSpPr/{P_NS}cNvPr")
            shpname = nvpr.get("name") if nvpr is not None else None
            if shpname:
                shapes[shpname] = "".join(t.text or "" for t in sp.iter(A_NS + "t"))
        return shapes
    return {}


def read_props(path: str) -> dict[str, str]:
    with zipfile.ZipFile(path) as z:
        registry = read_registry_slide(z)
        legacy: dict[str, str] = {}
        if PART in z.namelist():
            xml = z.read(PART).decode("utf-8", "replace")
            legacy = {m.group(1): m.group(2) for m in PROP_RE.finditer(xml)}
    # Registry slide wins; legacy custom.xml fills in only what the registry
    # slide doesn't have -- same priority as DeckRegistry.ReadStringProperty.
    return {**legacy, **registry}


def main() -> int:
    if len(sys.argv) < 2:
        print(__doc__.strip().splitlines()[-3].strip(), file=sys.stderr)
        return 2

    deck = sys.argv[1]
    props = read_props(deck)

    if len(sys.argv) == 2:
        if not props:
            print("no custom document properties on this deck")
            return 0
        for k, v in props.items():
            print(f"{k} = [{v}]")
        return 0

    wanted = sys.argv[2]
    if wanted in props:
        print(props[wanted])
        return 0

    # An absent property must not print an empty line that reads as an empty
    # value -- that is the ambiguous-zero mistake in a different costume.
    print(f"<{wanted} is NOT PRESENT on this deck>", file=sys.stderr)
    return 1


if __name__ == "__main__":
    sys.exit(main())
