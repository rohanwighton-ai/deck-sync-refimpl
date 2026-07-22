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
from dataclasses import dataclass, field
from typing import Sequence
from zipfile import ZIP_DEFLATED, ZipFile

from discovery import NS, Candidate, discover_from_pptx_part

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


@dataclass(frozen=True)
class StructuralMismatch:
    # Position in the shape sequence (paired by discover()'s z_order, the
    # same canonical ordering _find_shape_by_z_order already relies on).
    # -1 for whole-sequence issues (e.g. the count mismatch itself) that
    # aren't about one specific position.
    index: int
    kind: str  # "shape_count" | "type" | "missing_in_duplicate" | "extra_in_duplicate"
    detail: str


@dataclass(frozen=True)
class StructuralVerification:
    source_count: int
    duplicate_count: int
    mismatches: list[StructuralMismatch] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not self.mismatches


def verify_structure(source: Sequence[Candidate], duplicate: Sequence[Candidate]) -> StructuralVerification:
    """Structural verification after duplication, per specs/verification.md:
    shape count, type, and identity-tag correspondence between a duplicate
    and its source, checked explicitly rather than assumed from the
    duplication API succeeding.

    `source` and `duplicate` are the discover() output for the source part
    and its duplicate, respectively. Tagged shapes (identity_tag is not
    None) are paired by tag, not by list position: a tag is the persistent
    identity, so a pure reorder -- verify_z_order's whole concern -- must
    not also look like a structural defect here. Pairing tagged shapes
    positionally would make that impossible, since a reorder moves a tag's
    list position by definition. Untagged shapes (always true straight out
    of today's discover(), which never sets identity_tag) have no such
    signal and fall back to positional pairing within just the untagged
    subsequence. Any count mismatch is reported explicitly rather than
    silently truncating the comparison to whichever list is shorter.
    """
    mismatches: list[StructuralMismatch] = []

    if len(source) != len(duplicate):
        mismatches.append(
            StructuralMismatch(
                index=-1,
                kind="shape_count",
                detail=f"source has {len(source)} shape(s), duplicate has {len(duplicate)}",
            )
        )

    source_tagged = {c.identity_tag: c for c in source if c.identity_tag is not None}
    duplicate_tagged = {c.identity_tag: c for c in duplicate if c.identity_tag is not None}

    for tag in sorted(set(source_tagged) & set(duplicate_tagged)):
        s, d = source_tagged[tag], duplicate_tagged[tag]
        if s.shape_type != d.shape_type:
            mismatches.append(
                StructuralMismatch(
                    index=s.z_order,
                    kind="type",
                    detail=f"tagged shape {tag!r}: source is {s.shape_type!r}, duplicate is {d.shape_type!r}",
                )
            )
    for tag in sorted(set(source_tagged) - set(duplicate_tagged)):
        mismatches.append(
            StructuralMismatch(
                index=source_tagged[tag].z_order,
                kind="missing_in_duplicate",
                detail=f"tagged shape {tag!r} has no counterpart in duplicate",
            )
        )
    for tag in sorted(set(duplicate_tagged) - set(source_tagged)):
        mismatches.append(
            StructuralMismatch(
                index=duplicate_tagged[tag].z_order,
                kind="extra_in_duplicate",
                detail=f"duplicate has tagged shape {tag!r} with no source counterpart",
            )
        )

    source_untagged = [c for c in source if c.identity_tag is None]
    duplicate_untagged = [c for c in duplicate if c.identity_tag is None]

    for i, (s, d) in enumerate(zip(source_untagged, duplicate_untagged)):
        if s.shape_type != d.shape_type:
            mismatches.append(
                StructuralMismatch(
                    index=i,
                    kind="type",
                    detail=f"source shape {i} is {s.shape_type!r}, duplicate is {d.shape_type!r}",
                )
            )

    for i in range(len(duplicate_untagged), len(source_untagged)):
        mismatches.append(
            StructuralMismatch(
                index=i,
                kind="missing_in_duplicate",
                detail=f"source shape {i} ({source_untagged[i].name!r}) has no counterpart",
            )
        )
    for i in range(len(source_untagged), len(duplicate_untagged)):
        mismatches.append(
            StructuralMismatch(
                index=i,
                kind="extra_in_duplicate",
                detail=f"duplicate shape {i} ({duplicate_untagged[i].name!r}) has no source counterpart",
            )
        )

    return StructuralVerification(source_count=len(source), duplicate_count=len(duplicate), mismatches=mismatches)


def verify_structure_from_pptx(
    path: str, source_part: str, duplicate_part: str
) -> StructuralVerification:
    """Convenience entry point: run verify_structure() on two shape-tree-bearing
    parts of a .pptx file (e.g. a source slide and its duplicate)."""
    source = discover_from_pptx_part(path, source_part)
    duplicate = discover_from_pptx_part(path, duplicate_part)
    return verify_structure(source, duplicate)


@dataclass(frozen=True)
class ZOrderMismatch:
    tag_a: str
    tag_b: str
    detail: str


@dataclass(frozen=True)
class ZOrderVerification:
    pairs_checked: int
    mismatches: list[ZOrderMismatch] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not self.mismatches


def verify_z_order(source: Sequence[Candidate], duplicate: Sequence[Candidate]) -> ZOrderVerification:
    """Z-order (stacking order) verification per specs/verification.md, kept
    distinct from verify_structure's count/type/tag correspondence: a
    duplicate can have exactly the right shapes, tags, and values while a
    stacking-order regression still makes an overlaid field invisible (e.g. a
    transparent text box ending up behind its background shape). Structural
    correctness and stacking correctness are different claims, so neither is
    inferred from the other.

    Pairing here is by identity_tag, same as verify_structure's tagged-shape
    path -- both need tag-based, order-independent correspondence for the
    same reason: position-based pairing can never observe a reorder, since
    discover()'s z_order *is* each shape's position in its own list, so
    position-to-position comparison always finds z_order == z_order. This
    function only checks *relative* stacking order between tagged shapes
    (verify_structure already confirms the tagged shapes themselves
    correspond); untagged shapes (identity_tag is None) are excluded here
    for the same reason verify_structure falls back to position for them --
    there's no reliable way to say which duplicate shape an untagged source
    shape corresponds to.

    Every pair of commonly-tagged shapes is compared (not just adjacent
    ones), so a single swap deep in the stack is caught regardless of how
    many other shapes sit between the two that moved.
    """
    source_by_tag = {c.identity_tag: c for c in source if c.identity_tag is not None}
    duplicate_by_tag = {c.identity_tag: c for c in duplicate if c.identity_tag is not None}
    common_tags = sorted(set(source_by_tag) & set(duplicate_by_tag))

    mismatches: list[ZOrderMismatch] = []
    for i, tag_a in enumerate(common_tags):
        for tag_b in common_tags[i + 1 :]:
            source_below = source_by_tag[tag_a].z_order < source_by_tag[tag_b].z_order
            duplicate_below = duplicate_by_tag[tag_a].z_order < duplicate_by_tag[tag_b].z_order
            if source_below != duplicate_below:
                mismatches.append(
                    ZOrderMismatch(
                        tag_a=tag_a,
                        tag_b=tag_b,
                        detail=(
                            f"{tag_a!r} is {'below' if source_below else 'above'} {tag_b!r} in the source "
                            f"(z_order {source_by_tag[tag_a].z_order} vs {source_by_tag[tag_b].z_order}), but "
                            f"{'below' if duplicate_below else 'above'} it in the duplicate "
                            f"(z_order {duplicate_by_tag[tag_a].z_order} vs {duplicate_by_tag[tag_b].z_order})"
                        ),
                    )
                )

    n = len(common_tags)
    return ZOrderVerification(pairs_checked=n * (n - 1) // 2, mismatches=mismatches)


def verify_z_order_from_pptx(path: str, source_part: str, duplicate_part: str) -> ZOrderVerification:
    """Convenience entry point: run verify_z_order() on two shape-tree-bearing
    parts of a .pptx file (e.g. a source slide and its duplicate)."""
    source = discover_from_pptx_part(path, source_part)
    duplicate = discover_from_pptx_part(path, duplicate_part)
    return verify_z_order(source, duplicate)
