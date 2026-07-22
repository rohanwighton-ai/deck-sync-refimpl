"""Identity tag storage: read/write PowerPoint's hidden Shape.Tags/Slide.Tags
mechanism at the OOXML level.

See specs/identity-tags.md for the requirements this implements, including
citations for the underlying XML mechanism -- verified against ECMA-376's
User Defined Tags Part definition and a real-world example, not guessed,
since python-pptx has no built-in support for this and no fixture on disk
carried any tags to reverse-engineer from.
"""

from __future__ import annotations

import posixpath
import xml.etree.ElementTree as ET
from zipfile import ZipFile

from discovery import NS, Candidate
from lib.ooxml import find_shape_element_by_z_order, write_zip_parts

PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
TAGS_REL_TYPE = f"{DOC_REL_NS}/tags"
TAGS_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.tags+xml"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
CONTENT_TYPES_PART = "[Content_Types].xml"

_XML_DECLARATION = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

for _prefix, _uri in NS.items():
    ET.register_namespace(_prefix, _uri)
ET.register_namespace("", PKG_REL_NS)  # only used when serializing a standalone .rels tree
ET.register_namespace("r", DOC_REL_NS)


def _rels_path_for(part_name: str) -> str:
    directory = posixpath.dirname(part_name)
    filename = posixpath.basename(part_name)
    prefix = f"{directory}/" if directory else ""
    return f"{prefix}_rels/{filename}.rels"


def _resolve_relative_target(owning_part: str, target: str) -> str:
    return posixpath.normpath(posixpath.join(posixpath.dirname(owning_part), target))


def _read_optional(z: ZipFile, part_name: str) -> bytes | None:
    try:
        return z.read(part_name)
    except KeyError:
        return None


def _parse_rels(rels_bytes: bytes | None) -> list[tuple[str, str, str]]:
    if rels_bytes is None:
        return []
    root = ET.fromstring(rels_bytes)
    return [
        (el.get("Id", ""), el.get("Type", ""), el.get("Target", ""))
        for el in root.findall(f"{{{PKG_REL_NS}}}Relationship")
    ]


def _serialize_rels(relationships: list[tuple[str, str, str]]) -> bytes:
    root = ET.Element(f"{{{PKG_REL_NS}}}Relationships")
    for rel_id, rel_type, target in relationships:
        ET.SubElement(root, f"{{{PKG_REL_NS}}}Relationship", {"Id": rel_id, "Type": rel_type, "Target": target})
    return _XML_DECLARATION + ET.tostring(root, encoding="unicode").encode("utf-8")


def _next_relationship_id(relationships: list[tuple[str, str, str]]) -> str:
    max_n = 0
    for rel_id, _, _ in relationships:
        if rel_id.startswith("rId") and rel_id[3:].isdigit():
            max_n = max(max_n, int(rel_id[3:]))
    return f"rId{max_n + 1}"


def _next_tags_part_name(existing_names: set[str]) -> str:
    max_n = 0
    for name in existing_names:
        if name.startswith("ppt/tags/tag") and name.endswith(".xml"):
            num = name[len("ppt/tags/tag") : -len(".xml")]
            if num.isdigit():
                max_n = max(max_n, int(num))
    return f"ppt/tags/tag{max_n + 1}.xml"


def _parse_tag_list(tags_bytes: bytes | None) -> dict[str, str]:
    if tags_bytes is None:
        return {}
    root = ET.fromstring(tags_bytes)
    return {el.get("name", ""): el.get("val", "") for el in root.findall("p:tag", NS)}


def _serialize_tag_list(tags: dict[str, str]) -> bytes:
    root = ET.Element(f"{{{NS['p']}}}tagLst")
    for name, val in sorted(tags.items()):
        ET.SubElement(root, f"{{{NS['p']}}}tag", {"name": name, "val": val})
    return _XML_DECLARATION + ET.tostring(root, encoding="unicode").encode("utf-8")


def _add_content_type_override(ct_bytes: bytes, part_name: str, content_type: str) -> bytes:
    root = ET.fromstring(ct_bytes)
    for el in root.findall(f"{{{CONTENT_TYPES_NS}}}Override"):
        if el.get("PartName") == f"/{part_name}":
            return ct_bytes  # already registered -- idempotent no-op
    ET.SubElement(root, f"{{{CONTENT_TYPES_NS}}}Override", {"PartName": f"/{part_name}", "ContentType": content_type})
    return _XML_DECLARATION + ET.tostring(root, encoding="unicode").encode("utf-8")


def _find_nvpr(spTree: ET.Element, z_order: int) -> ET.Element:
    """Locate the shape at `z_order` (via the shared shape-tree walk) and
    return its <p:nvPr>. Only leaf sp/pic shapes carry a tag-bearing nvPr --
    a group is never itself tagged (shape-identity-and-matching.md's
    discovery_scope: "tag the leaf, not the container"), so groups aren't
    candidates here.
    """
    shape = find_shape_element_by_z_order(spTree, z_order)
    tag = shape.tag.split("}")[-1]
    nvpr_path = "./p:nvSpPr/p:nvPr" if tag == "sp" else "./p:nvPicPr/p:nvPr"
    nvpr = shape.find(nvpr_path, NS)
    if nvpr is None:
        raise ValueError(f"shape with z_order={z_order} has no nvPr")
    return nvpr


def _insert_after_ph(nvpr: ET.Element, new_el: ET.Element) -> None:
    """Insert `new_el` respecting CT_ApplicationNonVisualDrawingProps' schema
    order (ph, media, custDataLst, extLst): before extLst if present,
    otherwise appended (correct either way, since ph/media -- if present --
    are already earlier in document order from the original file)."""
    ext_lst = nvpr.find("./p:extLst", NS)
    if ext_lst is not None:
        nvpr.insert(list(nvpr).index(ext_lst), new_el)
    else:
        nvpr.append(new_el)


def read_slide_tags(path: str, slide_part: str) -> dict[str, str]:
    """Read a slide's Slide.Tags-equivalent tags (slide_type, instance_key,
    period_key) via its direct Tags Part relationship. No such relationship
    yet -> no tags yet, returns {}."""
    rels_part = _rels_path_for(slide_part)
    with ZipFile(path) as z:
        relationships = _parse_rels(_read_optional(z, rels_part))
        target = next((t for _, ty, t in relationships if ty == TAGS_REL_TYPE), None)
        if target is None:
            return {}
        tags_bytes = _read_optional(z, _resolve_relative_target(slide_part, target))
    return _parse_tag_list(tags_bytes)


def upsert_slide_tags(path: str, slide_part: str, tags: dict[str, str]) -> None:
    """Read-merge-write a slide's tags: only the given keys are added/
    updated, any other existing tag on this slide is left untouched. Creates
    the Tags Part + relationship + content-type override on first use;
    reuses them (no duplicate part/relationship) on every call after."""
    rels_part = _rels_path_for(slide_part)
    with ZipFile(path) as z:
        names = set(z.namelist())
        relationships = _parse_rels(_read_optional(z, rels_part))
        ct_bytes = z.read(CONTENT_TYPES_PART)
        target = next((t for _, ty, t in relationships if ty == TAGS_REL_TYPE), None)

        creating_new = target is None
        if target is not None:
            tags_part = _resolve_relative_target(slide_part, target)
            existing_tags = _parse_tag_list(_read_optional(z, tags_part))
        else:
            tags_part = _next_tags_part_name(names)
            existing_tags = {}
            new_rel_id = _next_relationship_id(relationships)
            relationships.append(
                (new_rel_id, TAGS_REL_TYPE, posixpath.relpath(tags_part, posixpath.dirname(slide_part)))
            )

    merged = {**existing_tags, **tags}
    updates: dict[str, bytes] = {tags_part: _serialize_tag_list(merged)}
    if creating_new:
        updates[rels_part] = _serialize_rels(relationships)
        updates[CONTENT_TYPES_PART] = _add_content_type_override(ct_bytes, tags_part, TAGS_CONTENT_TYPE)

    write_zip_parts(path, updates)


def read_shape_tags(path: str, slide_part: str, shape: Candidate) -> dict[str, str]:
    """Read a shape's Shape.Tags-equivalent tags (role) via its indirect
    <p:custDataLst><p:tags r:id=".."/></p:custDataLst> reference, resolved
    through the owning slide's .rels. No custDataLst/tags present -> no tags
    yet, returns {}."""
    rels_part = _rels_path_for(slide_part)
    with ZipFile(path) as z:
        slide_root = ET.fromstring(z.read(slide_part))
        spTree = slide_root.find(".//p:spTree", NS)
        if spTree is None:
            raise ValueError(f"no p:spTree found in {slide_part}")
        nvpr = _find_nvpr(spTree, shape.z_order)
        tags_el = nvpr.find("./p:custDataLst/p:tags", NS)
        if tags_el is None:
            return {}
        r_id = tags_el.get(f"{{{DOC_REL_NS}}}id")

        relationships = _parse_rels(_read_optional(z, rels_part))
        target = next((t for rid, _, t in relationships if rid == r_id), None)
        if target is None:
            return {}
        tags_bytes = _read_optional(z, _resolve_relative_target(slide_part, target))
    return _parse_tag_list(tags_bytes)


def upsert_shape_tags(path: str, slide_part: str, shape: Candidate, tags: dict[str, str]) -> None:
    """Read-merge-write a shape's tags: only the given keys are added/
    updated, any other existing tag on this shape is left untouched. Creates
    the Tags Part + relationship + content-type override + the shape's
    <p:custDataLst><p:tags r:id=".."/></p:custDataLst> reference on first
    use; reuses them (no duplicate part/relationship, custDataLst inserted
    only once) on every call after. Two shapes on the same slide tagged
    independently each get their own Tags Part and relationship -- neither
    call reads or reuses the other's.
    """
    rels_part = _rels_path_for(slide_part)
    with ZipFile(path) as z:
        names = set(z.namelist())
        slide_root = ET.fromstring(z.read(slide_part))
        spTree = slide_root.find(".//p:spTree", NS)
        if spTree is None:
            raise ValueError(f"no p:spTree found in {slide_part}")
        nvpr = _find_nvpr(spTree, shape.z_order)
        tags_el = nvpr.find("./p:custDataLst/p:tags", NS)

        relationships = _parse_rels(_read_optional(z, rels_part))
        ct_bytes = z.read(CONTENT_TYPES_PART)

        creating_new = tags_el is None
        if tags_el is not None:
            r_id = tags_el.get(f"{{{DOC_REL_NS}}}id")
            target = next((t for rid, _, t in relationships if rid == r_id), None)
            if target is None:
                raise ValueError(f"shape references relationship {r_id!r} not found in {rels_part}")
            tags_part = _resolve_relative_target(slide_part, target)
            existing_tags = _parse_tag_list(_read_optional(z, tags_part))
        else:
            tags_part = _next_tags_part_name(names)
            existing_tags = {}
            new_rel_id = _next_relationship_id(relationships)
            relationships.append(
                (new_rel_id, TAGS_REL_TYPE, posixpath.relpath(tags_part, posixpath.dirname(slide_part)))
            )
            custdatalst = nvpr.find("./p:custDataLst", NS)
            if custdatalst is None:
                custdatalst = ET.Element(f"{{{NS['p']}}}custDataLst")
                _insert_after_ph(nvpr, custdatalst)
            ET.SubElement(custdatalst, f"{{{NS['p']}}}tags", {f"{{{DOC_REL_NS}}}id": new_rel_id})

    merged = {**existing_tags, **tags}
    updates: dict[str, bytes] = {tags_part: _serialize_tag_list(merged)}
    if creating_new:
        updates[slide_part] = _XML_DECLARATION + ET.tostring(slide_root, encoding="unicode").encode("utf-8")
        updates[rels_part] = _serialize_rels(relationships)
        updates[CONTENT_TYPES_PART] = _add_content_type_override(ct_bytes, tags_part, TAGS_CONTENT_TYPE)

    write_zip_parts(path, updates)
