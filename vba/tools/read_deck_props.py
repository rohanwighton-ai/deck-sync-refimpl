#!/usr/bin/env python3
"""Read a deck's custom document properties straight from the file, no Office.

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

A .pptx is a zip. Custom document properties live in `docProps/custom.xml` as
plain values -- not slide text, so the usual warning about XML disagreeing with
the object model does not apply here. This is what is on disk, full stop.

Deliberately NOT importable from VBA: the point is to check the writer using
something that shares no process, no cache and no code with it.

    python3 read_deck_props.py <deck.pptx> [property-name]

Exit 0 if found (or if listing), 1 if a named property is absent.
"""

from __future__ import annotations

import re
import sys
import zipfile

PART = "docProps/custom.xml"
# <property ... name="X"><vt:lpwstr>VALUE</vt:lpwstr></property>
PROP_RE = re.compile(r'name="([^"]+)"[^>]*>\s*<[^>]+>([^<]*)<')


def read_props(path: str) -> dict[str, str]:
    with zipfile.ZipFile(path) as z:
        if PART not in z.namelist():
            return {}
        xml = z.read(PART).decode("utf-8", "replace")
    return {m.group(1): m.group(2) for m in PROP_RE.finditer(xml)}


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
