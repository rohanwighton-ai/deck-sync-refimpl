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
    has_placeholder: bool
    has_text: bool

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


def _has_placeholder(el: ET.Element) -> bool:
    return el.find(".//p:nvPr/p:ph", NS) is not None


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
                results.append(
                    Candidate(
                        name=_shape_name(child),
                        group_path=group_path,
                        z_order=z[0],
                        shape_type="picture" if is_pic else "autoshape_or_textbox",
                        has_placeholder=_has_placeholder(child),
                        has_text=_has_text(child) if not is_pic else False,
                    )
                )

    walk(spTree, ())
    return results


def discover_from_pptx(path: str, slide_index: int = 1) -> list[Candidate]:
    """Convenience entry point: discover candidates on slideN.xml of a .pptx file."""
    with ZipFile(path) as z:
        with z.open(f"ppt/slides/slide{slide_index}.xml") as f:
            root = ET.parse(f).getroot()
    return discover(root)
