"""Field discovery for a single slide's shape tree.

See specs/discovery.md for the requirements this implements. Stdlib-only
(zipfile + xml.etree) deliberately -- no python-pptx dependency, so this
stays trivially runnable without an environment setup step.
"""

from __future__ import annotations

import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from zipfile import ZipFile

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}


@dataclass
class Candidate:
    name: str
    group_path: tuple[str, ...]
    z_order: int
    shape_type: str  # "autoshape_or_textbox" | "picture"
    placeholder_type: str | None  # e.g. "title", "body"; None if not a placeholder
    placeholder_idx: int | None  # layout placeholder index; None if not a placeholder
    has_text: bool
    position: tuple[int, int] | None = None  # (off_x, off_y) in EMU; None if no a:xfrm
    size: tuple[int, int] | None = None  # (cx, cy) in EMU; None if no a:xfrm
    # Set by a previous matching/tagging pass, per specs/matching.md's two-tier
    # rule. discover() never populates this (per specs/discovery.md's non-goals,
    # discovery only finds and describes candidates -- it doesn't read or write
    # identity tags), so this is always None coming out of discover(). It exists
    # on Candidate so matching.py has somewhere to look for an already-trusted tag.
    identity_tag: str | None = None

    @property
    def has_placeholder(self) -> bool:
        return self.placeholder_type is not None

    @property
    def is_candidate_field(self) -> bool:
        return self.shape_type == "picture" or self.has_text


def _shape_name(el: ET.Element) -> str:
    for path in ("./p:nvSpPr/p:cNvPr", "./p:nvGrpSpPr/p:cNvPr", "./p:nvPicPr/p:cNvPr"):
        cNvPr = el.find(path, NS)
        if cNvPr is not None:
            return cNvPr.get("name", "?")
    return "?"


def _has_text(sp: ET.Element) -> bool:
    return any((t.text or "").strip() for t in sp.findall(".//a:t", NS))


def _placeholder_info(el: ET.Element) -> tuple[str, int] | None:
    """Return (type, idx) if el's nvPr declares a placeholder, else None.

    Per OOXML, both attributes are optional on <p:ph> itself: an omitted
    `type` defaults to "obj" and an omitted `idx` defaults to 0 -- the
    element's mere presence is what marks a shape as a placeholder.
    """
    ph = el.find(".//p:nvPr/p:ph", NS)
    if ph is None:
        return None
    return ph.get("type", "obj"), int(ph.get("idx", "0"))


def _geometry(el: ET.Element) -> tuple[tuple[int, int], tuple[int, int]] | None:
    """Return ((off_x, off_y), (cx, cy)) in EMU from el's own <p:spPr/a:xfrm>.

    Reads the shape's own local off/ext verbatim -- it does not walk up
    through any parent group's transform, so a nested shape's position is
    only absolute when its group's chOff/chExt equals the group's own
    off/ext (true for every group in this project's current fixtures). A
    general child-coordinate-system transform is future work if a fixture
    ever needs it.
    """
    xfrm = el.find("./p:spPr/a:xfrm", NS)
    if xfrm is None:
        return None
    off = xfrm.find("./a:off", NS)
    ext = xfrm.find("./a:ext", NS)
    if off is None or ext is None:
        return None
    return (int(off.get("x", "0")), int(off.get("y", "0"))), (
        int(ext.get("cx", "0")),
        int(ext.get("cy", "0")),
    )


def discover(slide_xml_root: ET.Element) -> list[Candidate]:
    """Walk a slide's shape tree per specs/discovery.md: type-agnostic,
    recurses into groups, tags leaves not containers."""
    spTree = slide_xml_root.find(".//p:spTree", NS)
    if spTree is None:
        raise ValueError("no p:spTree found in slide XML root")
    results: list[Candidate] = []
    z = [0]

    def walk(el: ET.Element, group_path: tuple[str, ...]) -> None:
        for child in el:
            tag = child.tag.split("}")[-1]
            if tag == "grpSp":
                walk(child, group_path + (_shape_name(child),))
            elif tag in ("sp", "pic"):
                z[0] += 1
                is_pic = tag == "pic"
                ph_info = _placeholder_info(child)
                geometry = _geometry(child)
                results.append(
                    Candidate(
                        name=_shape_name(child),
                        group_path=group_path,
                        z_order=z[0],
                        shape_type="picture" if is_pic else "autoshape_or_textbox",
                        placeholder_type=ph_info[0] if ph_info else None,
                        placeholder_idx=ph_info[1] if ph_info else None,
                        has_text=_has_text(child) if not is_pic else False,
                        position=geometry[0] if geometry else None,
                        size=geometry[1] if geometry else None,
                    )
                )

    walk(spTree, ())
    return results


def discover_from_pptx(path: str, slide_index: int = 1) -> list[Candidate]:
    """Convenience entry point: discover candidates on slideN.xml of a .pptx file."""
    return discover_from_pptx_part(path, f"ppt/slides/slide{slide_index}.xml")


def discover_from_pptx_layout(path: str, layout_index: int = 1) -> list[Candidate]:
    """Convenience entry point: discover candidates on slideLayoutN.xml of a .pptx
    file. Some decks (e.g. a layouts-only master export) have no ppt/slides/* entries
    at all -- discover() itself is root-agnostic (only looks for p:spTree), so the
    same walk applies unchanged to a slideLayout root."""
    return discover_from_pptx_part(path, f"ppt/slideLayouts/slideLayout{layout_index}.xml")


def discover_from_pptx_part(path: str, part_name: str) -> list[Candidate]:
    """Discover candidates on an arbitrary shape-tree-bearing part (e.g. a slide or
    slideLayout XML entry) inside a .pptx file."""
    with ZipFile(path) as z:
        with z.open(part_name) as f:
            root = ET.parse(f).getroot()
    return discover(root)
