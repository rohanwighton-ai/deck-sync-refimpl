"""Shared OOXML plumbing used by more than one module.

verification.py and identity_tags.py both need to re-walk a slide's shape
tree to the exact element discover() numbered, and both need to safely
rewrite some parts of a .pptx zip while leaving everything else byte-for-
byte untouched. Promoted here once a second module needed each (see
IMPLEMENTATION_PLAN.md's "Notes for next planning pass") rather than
writing a third/fourth slightly-different copy.
"""

from __future__ import annotations

import os
import tempfile
import xml.etree.ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

_XML_DECLARATION = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'


def find_shape_element_by_z_order(spTree: ET.Element, z_order: int) -> ET.Element:
    """Re-walk a slide's shape tree in the exact order discover() numbers
    shapes in (see discovery.py's walk()) and return the raw <p:sp>/<p:pic>
    element a previously-discovered Candidate.z_order refers to."""
    z = [0]
    found: list[ET.Element] = []

    def walk(el: ET.Element) -> None:
        for child in el:
            tag = child.tag.split("}")[-1]
            if tag == "grpSp":
                walk(child)
            elif tag in ("sp", "pic"):
                z[0] += 1
                if z[0] == z_order:
                    found.append(child)

    walk(spTree)
    if not found:
        raise ValueError(f"no shape with z_order={z_order} found")
    return found[0]


def write_zip_parts(path: str, updates: dict[str, bytes]) -> None:
    """Rewrite `path`, replacing or adding each part in `updates`, copying
    every other existing entry byte-for-byte untouched. Written to a
    sibling temp file first and swapped into place so a failure mid-write
    never leaves `path` corrupt."""
    directory = os.path.dirname(os.path.abspath(path)) or "."
    fd, tmp_path = tempfile.mkstemp(suffix=".pptx", dir=directory)
    os.close(fd)
    try:
        with ZipFile(path) as src, ZipFile(tmp_path, "w", ZIP_DEFLATED) as dst:
            written = set()
            for item in src.infolist():
                data = updates.get(item.filename)
                dst.writestr(item, data if data is not None else src.read(item.filename))
                written.add(item.filename)
            for name, data in updates.items():
                if name not in written:
                    dst.writestr(name, data)
        os.replace(tmp_path, path)
    except BaseException:
        os.remove(tmp_path)
        raise
