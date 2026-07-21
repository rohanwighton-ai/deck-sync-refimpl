"""Value verification: given a tagged field shape and its linked data-source
value, prove the link resolves correctly rather than assuming a tag-and-seed
pairing that merely looks consistent is actually wired up right.

See specs/verification.md for the requirements this implements. Stdlib-only
(zipfile + xml.etree), matching discovery.py's dependency-light approach.
"""

from __future__ import annotations

import hashlib
import os
import tempfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from zipfile import ZIP_DEFLATED, ZipFile

from discovery import NS, Candidate

_XML_DECLARATION = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

# Preserve the original p:/a: prefixes on write-back -- without this,
# ET.tostring() invents ns0/ns1 prefixes instead of reusing the ones already
# declared in the part, which is needlessly different from the source.
for _prefix, _uri in NS.items():
    ET.register_namespace(_prefix, _uri)


@dataclass(frozen=True)
class InjectResult:
    written: bool  # False if the current value already matched the source (no-op)
    initial_hash: str  # hash of the shape's value before this operation
    source_hash: str  # hash of the value the shape is linked to
    final_hash: str  # hash of the shape's value after this operation (== initial_hash if not written)
    verified: bool  # final_hash == source_hash, checked explicitly, never assumed


def _hash(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


def _get_text(sp: ET.Element) -> str:
    return "".join(t.text or "" for t in sp.findall(".//a:t", NS))


def _set_text(sp: ET.Element, value: str) -> None:
    """Write `value` into sp's first text run, per specs/verification.md's
    inject_primitive. Any additional runs are cleared rather than left with
    stale text, so a later _get_text() reflects exactly `value`."""
    runs = sp.findall(".//a:t", NS)
    if not runs:
        raise ValueError("shape has no text runs to write a value into")
    runs[0].text = value
    for extra in runs[1:]:
        extra.text = ""


def _find_shape_by_z_order(spTree: ET.Element, z_order: int) -> ET.Element:
    """Re-walk in the exact order discover() numbers shapes in (see
    discovery.py's walk()), to locate the element a previously-discovered
    Candidate.z_order refers to."""
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


def _write_part(path: str, part_name: str, xml_bytes: bytes) -> None:
    """Zip write-back: replace part_name's bytes, leaving every other zip
    entry byte-for-byte untouched. Writes to a sibling temp file first and
    swaps it into place so a failure mid-write never leaves `path` corrupt."""
    directory = os.path.dirname(os.path.abspath(path)) or "."
    fd, tmp_path = tempfile.mkstemp(suffix=".pptx", dir=directory)
    os.close(fd)
    try:
        with ZipFile(path) as src, ZipFile(tmp_path, "w", ZIP_DEFLATED) as dst:
            for item in src.infolist():
                data = xml_bytes if item.filename == part_name else src.read(item.filename)
                dst.writestr(item, data)
        os.replace(tmp_path, path)
    except BaseException:
        os.remove(tmp_path)
        raise


def _read_part(path: str, part_name: str) -> ET.Element:
    with ZipFile(path) as z, z.open(part_name) as f:
        return ET.parse(f).getroot()


def inject_primitive(path: str, part_name: str, shape: Candidate, source_value: str) -> InjectResult:
    """Core verification operation per specs/verification.md: hash `shape`'s
    current value and `source_value`; no-op (write nothing) if they already
    match; otherwise write `source_value` into the shape and re-hash the
    written-back value to confirm the write actually took, rather than
    assuming success from the write call alone.
    """
    root = _read_part(path, part_name)
    spTree = root.find(".//p:spTree", NS)
    if spTree is None:
        raise ValueError(f"no p:spTree found in {part_name}")
    sp = _find_shape_by_z_order(spTree, shape.z_order)

    initial_hash = _hash(_get_text(sp))
    source_hash = _hash(source_value)

    if initial_hash == source_hash:
        return InjectResult(
            written=False, initial_hash=initial_hash, source_hash=source_hash, final_hash=initial_hash, verified=True
        )

    _set_text(sp, source_value)
    _write_part(path, part_name, _XML_DECLARATION + ET.tostring(root, encoding="unicode").encode("utf-8"))

    reread_root = _read_part(path, part_name)
    reread_spTree = reread_root.find(".//p:spTree", NS)
    if reread_spTree is None:
        raise ValueError(f"no p:spTree found in {part_name} after write-back")
    reread_sp = _find_shape_by_z_order(reread_spTree, shape.z_order)
    final_hash = _hash(_get_text(reread_sp))

    return InjectResult(
        written=True,
        initial_hash=initial_hash,
        source_hash=source_hash,
        final_hash=final_hash,
        verified=final_hash == source_hash,
    )
