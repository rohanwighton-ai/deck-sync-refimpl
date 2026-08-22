#!/usr/bin/env python3
"""From-disk verification of MS<n>_CALDATE shapes in a saved .pptx.

Gist: opens the saved PowerPoint file as a zip and reads the raw slide XML
to count the hidden calendar-date textboxes, so success is judged from the
file's own bytes rather than from what the script that wrote them claimed.

This is the external check add_caldate_shapes.vbs's header points at: after
that script reports success for a slide, the wrapper loop runs this against
the SAVED file. It never trusts the writer's in-process report -- the whole
lesson of 2026-08-08 ("a verifier sharing state with the writer answers for
the cache") and of the duplicate MS4_CALDATE found 2026-08-22 only by
reading the saved XML.

For each MS<n>_CALDATE (n=1..7) on the requested slide(s) it reports:
  - how many occurrences exist in the slide's XML,
  - whether each is GROUPED (nested inside a <p:grpSp>) or STRAY
    (a direct child of <p:spTree> -- the duplicate/interrupted-run
    signature), and
  - the shape's stored width in EMU (catches the AutoSize collapse bug
    that once shipped cx="65" shapes: anything under ~1000 EMU is flagged).

Slide numbers are resolved through presentation.xml's sldIdLst -> rels
mapping, not by assuming slideN.xml == slide N.

Usage:
  python3 verify_caldate.py <path-to-pptx> <slideIndex> [<slideIndex> ...]
  python3 verify_caldate.py <path-to-pptx> all

Exit code 0 = every requested slide has exactly one GROUPED occurrence of
each expected name and no strays; 1 = any slide deviates (missing shapes
count as a deviation only when the slide has a milestone device, i.e. an
MS_BAR shape); 2 = usage/IO error.
"""
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
CALDATE_RE = re.compile(r"^MS([1-7])_CALDATE$")
MIN_SANE_CX = 1000  # EMU; the AutoSize-collapse bug shipped cx="65"


def slide_order(zf):
    """Return the list of slide part names in presentation order."""
    pres = ET.fromstring(zf.read("ppt/presentation.xml"))
    rels = ET.fromstring(zf.read("ppt/_rels/presentation.xml.rels"))
    rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    rid_to_target = {
        rel.get("Id"): rel.get("Target")
        for rel in rels.findall(f"{{{rel_ns}}}Relationship")
    }
    order = []
    for sld in pres.findall("p:sldIdLst/p:sldId", NS):
        rid = sld.get(f"{{{NS['r']}}}embed") or sld.get(f"{{{NS['r']}}}id")
        target = rid_to_target[rid]
        order.append("ppt/" + target.lstrip("/").replace("../", ""))
    return order


def audit_slide(zf, part):
    """Return (has_device, findings) for one slide part.

    findings: {name: [(where, cx), ...]} where 'where' is GROUPED|STRAY.
    has_device: True if the slide contains a shape named MS_BAR.
    """
    root = ET.fromstring(zf.read(part))
    sp_tree = root.find("p:cSld/p:spTree", NS)
    findings = {}
    has_device = False

    def sp_name(el):
        nv = el.find("./p:nvSpPr/p:cNvPr", NS)
        if nv is None:
            nv = el.find("./p:nvGrpSpPr/p:cNvPr", NS)
        return nv.get("name") if nv is not None else None

    def sp_cx(el):
        ext = el.find("./p:spPr/a:xfrm/a:ext", NS)
        return int(ext.get("cx")) if ext is not None else None

    def walk(el, inside_group):
        for child in el:
            tag = child.tag.split("}")[1]
            if tag == "sp":
                name = sp_name(child)
                if name == "MS_BAR":
                    nonlocal has_device
                    has_device = True
                if name and CALDATE_RE.match(name):
                    where = "GROUPED" if inside_group else "STRAY"
                    findings.setdefault(name, []).append((where, sp_cx(child)))
            elif tag == "grpSp":
                walk(child, True)

    walk(sp_tree, False)
    return has_device, findings


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        return 2
    path = sys.argv[1]
    try:
        zf = zipfile.ZipFile(path)
    except OSError as e:
        print(f"cannot open {path}: {e}")
        return 2
    with zf:
        order = slide_order(zf)
        if sys.argv[2].lower() == "all":
            indices = range(1, len(order) + 1)
        else:
            indices = [int(a) for a in sys.argv[2:]]
        ok = True
        for idx in indices:
            if not 1 <= idx <= len(order):
                print(f"slide {idx}: out of range (deck has {len(order)})")
                ok = False
                continue
            has_device, findings = audit_slide(zf, order[idx - 1])
            problems = []
            for n in range(1, 8):
                name = f"MS{n}_CALDATE"
                occ = findings.get(name, [])
                grouped = [o for o in occ if o[0] == "GROUPED"]
                strays = [o for o in occ if o[0] == "STRAY"]
                tiny = [o for o in occ if o[1] is not None and o[1] < MIN_SANE_CX]
                if len(occ) > 1:
                    problems.append(
                        f"{name}: DUPLICATE x{len(occ)} "
                        f"(grouped={len(grouped)}, stray={len(strays)})"
                    )
                elif strays:
                    problems.append(f"{name}: STRAY (ungrouped)")
                elif not occ and has_device:
                    problems.append(f"{name}: MISSING")
                if tiny:
                    problems.append(f"{name}: collapsed width cx={tiny[0][1]}")
            if not has_device and not findings:
                print(f"slide {idx}: no milestone device -- nothing expected, OK")
            elif problems:
                ok = False
                print(f"slide {idx}: FAIL")
                for p in problems:
                    print(f"  {p}")
            else:
                print(f"slide {idx}: OK -- 7/7 grouped, no strays, no duplicates")
        return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
