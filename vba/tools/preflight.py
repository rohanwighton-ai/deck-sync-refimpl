#!/usr/bin/env python3
"""Will this deck and its workbook actually connect? Answer before spending an evening.

READS THE FILES' OWN BYTES. No Office, no COM, nothing opened. That is the point:
every in-process check this project has relied on has been fooled at least once by
PowerPoint's cache, and this has to be trustworthy on a machine nobody has verified.
It is also why it can be run on a deck that is closed, or on a colleague's copy.

WHAT IT IS FOR. Rohan, 2026-08-05: "I'll get copilot to get them in a drafting sheet
tomorrow if you can guarantee it will connect to the decks it needs to." Nobody can
guarantee that for a deck they have never seen. This makes it checkable in a minute
instead, so a bad pairing costs a minute rather than an evening's drafting.

    python3 preflight.py <deck.pptx> [--period Q3F26]

Exit 0 when the loop would connect, 1 when it would not. Every finding says what
breaks, not just that something is wrong.
"""

from __future__ import annotations

import re
import sys
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZipFile

RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
TAGS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags"
XL = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
DOC_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
P_NS = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
REGISTRY_SLIDE_NAME = "DeckSyncRegistry"

INSTANCE_TAG = "INSTANCE_KEY"     # PowerPoint uppercases tag names on save
TEMPLATE_TAG = "IS_TEMPLATE"
INSTANCE_HDR = "Instance ID"
QUARTER_HDR = "Quarter"

# Any quarter-ish literal a slide might carry, so a deck full of last quarter's
# text can be spotted. Deliberately loose: the point is to notice, not to parse.
PERIOD_RE = re.compile(r"\b(?:FY\s?\d{2}\s?Q[1-4]|Q[1-4]\s?FY?\s?\d{2})\b", re.I)

findings: list[tuple[str, str]] = []


def fail(msg: str) -> None:
    findings.append(("FAIL", msg))


def warn(msg: str) -> None:
    findings.append(("WARN", msg))


def ok(msg: str) -> None:
    findings.append(("ok", msg))


# --- deck ------------------------------------------------------------------

def registry_slide_props(z: ZipFile) -> dict[str, str]:
    """Shape NAME -> shape text, on the hidden slide named DeckSyncRegistry --
    the REAL storage since 2026-08-16 (see DeckRegistry.bas's REGISTRY_SLIDE_NAME
    comment). docProps/custom.xml is a read-only fallback now, for decks never
    touched since the migration."""
    for name in z.namelist():
        if not re.fullmatch(r"ppt/slides/slide\d+\.xml", name):
            continue
        root = ET.fromstring(z.read(name))
        cSld = root.find(f"{P_NS}cSld")
        if cSld is None or cSld.get("name") != REGISTRY_SLIDE_NAME:
            continue
        shapes = {}
        for sp in root.iter(P_NS + "sp"):
            nvpr = sp.find(f"{P_NS}nvSpPr/{P_NS}cNvPr")
            shpname = nvpr.get("name") if nvpr is not None else None
            if shpname:
                shapes[shpname] = "".join(t.text or "" for t in sp.iter(A_NS + "t"))
        return shapes
    return {}


def deck_props(z: ZipFile) -> dict[str, str]:
    legacy = {}
    if "docProps/custom.xml" in z.namelist():
        root = ET.fromstring(z.read("docProps/custom.xml"))
        for p in root:
            name = p.get("name")
            val = "".join(c.text or "" for c in p)
            if name:
                legacy[name] = val
    # Registry slide wins; legacy custom.xml fills in only what's absent from
    # it -- same priority as DeckRegistry.ReadStringProperty.
    return {**legacy, **registry_slide_props(z)}


def slide_indices(z: ZipFile) -> list[int]:
    idx = []
    for n in z.namelist():
        m = re.fullmatch(r"ppt/slides/slide(\d+)\.xml", n)
        if m:
            idx.append(int(m.group(1)))
    return sorted(idx)


def tags_for_slide(z: ZipFile, i: int) -> dict[str, str]:
    rels = f"ppt/slides/_rels/slide{i}.xml.rels"
    if rels not in z.namelist():
        return {}
    out = {}
    for rel in ET.fromstring(z.read(rels)):
        if rel.get("Type") != TAGS_REL:
            continue
        target = rel.get("Target", "").replace("../", "ppt/")
        if target not in z.namelist():
            continue
        for tag in ET.fromstring(z.read(target)):
            nm = tag.get("name")
            if nm:
                out[nm.upper()] = tag.get("val", "")
    return out


def slide_text(z: ZipFile, i: int) -> str:
    part = f"ppt/slides/slide{i}.xml"
    xml = z.read(part).decode("utf8", errors="replace")
    return " ".join(re.findall(r"<a:t>(.*?)</a:t>", xml, re.S))


# --- workbook --------------------------------------------------------------

def read_sheet(path: Path, sheet_name: str):
    """-> (headers dict col->name, rows list of dict col->text) or None if absent."""
    z = ZipFile(path)
    shared = []
    if "xl/sharedStrings.xml" in z.namelist():
        for si in ET.fromstring(z.read("xl/sharedStrings.xml")):
            shared.append("".join(t.text or "" for t in si.iter(XL + "t")))
    wbx = ET.fromstring(z.read("xl/workbook.xml"))
    rels = {r.get("Id"): r.get("Target") for r in ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))}
    names = {}
    for sh in wbx.iter(XL + "sheet"):
        names[sh.get("name")] = rels[sh.get(DOC_REL + "id")]
    if sheet_name not in names:
        return None, sorted(names)
    part = "xl/" + names[sheet_name].lstrip("/").removeprefix("xl/")

    def cell(c):
        v = c.find(XL + "v")
        if c.get("t") == "inlineStr":
            return "".join(t.text or "" for t in c.iter(XL + "t"))
        if v is None or v.text is None:
            return ""
        return shared[int(v.text)] if c.get("t") == "s" else v.text

    grid = {}
    for row in ET.fromstring(z.read(part)).iter(XL + "row"):
        grid[int(row.get("r"))] = {
            re.match(r"([A-Z]+)", c.get("r")).group(1): cell(c) for c in row.iter(XL + "c")
        }
    hdr = {c: v for c, v in grid.get(1, {}).items() if v.strip()}
    rows = [cells for r, cells in sorted(grid.items()) if r >= 2]
    return (hdr, rows), sorted(names)


# --- the check -------------------------------------------------------------

def main() -> int:
    if len(sys.argv) < 2:
        print("usage: preflight.py <deck.pptx> [--period Q3F26]", file=sys.stderr)
        return 2
    deck = Path(sys.argv[1])
    want_period = None
    if "--period" in sys.argv:
        want_period = sys.argv[sys.argv.index("--period") + 1]

    if not deck.exists():
        print(f"no such deck: {deck}", file=sys.stderr)
        return 2

    z = ZipFile(deck)
    props = deck_props(z)
    print(f"DECK  {deck}")

    # 1. period
    period = props.get("DeckSyncPeriod", "").strip()
    if not period:
        fail("the deck declares NO period. Drafting and sync both filter on it, and "
             "onboarding refuses without it. Run 'Start a Quarter' first.")
    else:
        ok(f"declares period '{period}'")
        if want_period and period.upper() != want_period.upper():
            fail(f"deck says '{period}' but you expected '{want_period}'. Periods are free "
                 f"text matched exactly -- two spellings of one quarter match NOTHING and "
                 f"report as a clean run of zero rows.")

    # 2. registered slide types
    types = {k[len("DeckSyncType:"):]: v for k, v in props.items() if k.startswith("DeckSyncType:")}
    if not types:
        fail("no slide type registered. Onboard a slide type before drafting.")
    elif len(types) > 1:
        fail(f"{len(types)} slide types registered ({', '.join(sorted(types))}). A deck that "
             f"gets drafted carries exactly one -- usually this means it was onboarded twice.")
    else:
        tname, traw = next(iter(types.items()))
        wsname = traw.split("|", 1)[1] if "|" in traw else ""
        ok(f"slide type '{tname}' -> worksheet '{wsname}'")

    # 3. paired workbook
    wbpath_raw = props.get("DeckSyncWorkbookPath", "").strip()
    wb_local = None
    if not wbpath_raw:
        fail("no paired workbook. Onboarding establishes the pairing.")
    else:
        wb_local = Path(wbpath_raw.replace("\\", "/").replace("C:", "/mnt/c"))
        if not wb_local.exists():
            fail(f"paired workbook does NOT exist: {wbpath_raw}")
            wb_local = None
        else:
            ok(f"paired workbook exists: {wbpath_raw}")

    # 4. the register sheet itself
    reg_rows, reg_hdr = None, {}
    if wb_local is not None and len(types) == 1 and wsname:
        result, all_sheets = read_sheet(wb_local, wsname)
        if result is None:
            fail(f"the workbook has NO sheet named '{wsname}' (it has: {', '.join(all_sheets)}). "
                 f"Publish and sync would address different tabs.")
        else:
            reg_hdr, reg_rows = result
            cols = {v: k for k, v in reg_hdr.items()}
            if INSTANCE_HDR not in cols:
                fail(f"sheet '{wsname}' has no '{INSTANCE_HDR}' column -- it is not a register.")
            elif QUARTER_HDR not in cols:
                warn(f"sheet '{wsname}' has no '{QUARTER_HDR}' column: an old one-row-per-slide "
                     f"sheet. It is readable but holds no periods.")
            else:
                ci, cq = cols[INSTANCE_HDR], cols[QUARTER_HDR]
                live = [r for r in reg_rows if r.get(ci, "").strip()]
                periods = {}
                for r in live:
                    periods[r.get(cq, "").strip()] = periods.get(r.get(cq, "").strip(), 0) + 1
                ok(f"register '{wsname}': {len(live)} rows, periods " +
                   ", ".join(f"{k or '(blank)'}={v}" for k, v in sorted(periods.items())))
                n_here = periods.get(period, 0)
                if period and n_here == 0:
                    fail(f"NO register rows for the deck's period '{period}'. Drafting sheets "
                         f"would be empty and sync would write nothing, reporting success.")
                fields = [v for k, v in sorted(reg_hdr.items()) if v not in (INSTANCE_HDR, QUARTER_HDR)]
                ok(f"fields on the register: {', '.join(fields) if fields else '(none)'}")

    # 5. slides vs rows
    idx = slide_indices(z)
    tagged, templates = {}, 0
    for i in idx:
        t = tags_for_slide(z, i)
        if t.get(TEMPLATE_TAG, "").strip():
            templates += 1
            continue
        key = t.get(INSTANCE_TAG, "").strip()
        if key:
            tagged[key] = i
    ok(f"deck has {len(idx)} slides: {len(tagged)} tagged with an instance key, "
       f"{templates} template, {len(idx) - len(tagged) - templates} untagged")
    if not tagged:
        fail("no slide carries an instance key. Nothing links slides to register rows; "
             "onboard the slides first.")

    if reg_rows is not None and tagged and INSTANCE_HDR in {v: k for k, v in reg_hdr.items()}:
        cols = {v: k for k, v in reg_hdr.items()}
        ci = cols[INSTANCE_HDR]
        cq = cols.get(QUARTER_HDR)
        keys_here = {r[ci].strip() for r in reg_rows
                     if r.get(ci, "").strip() and (cq is None or r.get(cq, "").strip() == period)}
        no_row = sorted(set(tagged) - keys_here)
        no_slide = sorted(keys_here - set(tagged))
        if no_row:
            warn(f"{len(no_row)} slide(s) have NO register row for '{period}' -- they will be "
                 f"skipped silently: {', '.join(no_row[:6])}{' ...' if len(no_row) > 6 else ''}")
        if no_slide:
            warn(f"{len(no_slide)} register row(s) have NO slide: "
                 f"{', '.join(no_slide[:6])}{' ...' if len(no_slide) > 6 else ''}")
        if not no_row and not no_slide:
            ok(f"every tagged slide has a row for '{period}', and vice versa ({len(keys_here)})")

    # 6. THE TRAP: does the deck's text still describe a different quarter?
    if period:
        others: dict[str, int] = {}
        for i in idx:
            for m in set(PERIOD_RE.findall(slide_text(z, i))):
                lit = m.strip()
                if lit.upper().replace(" ", "") != period.upper().replace(" ", ""):
                    others[lit] = others.get(lit, 0) + 1
        if others:
            top = ", ".join(f"{k} on {v} slide(s)" for k, v in
                            sorted(others.items(), key=lambda x: -x[1])[:4])
            warn(f"the deck declares '{period}' but its SLIDES still mention: {top}. "
                 f"If this deck holds last quarter's content, onboard it as THAT period -- "
                 f"onboarding stamps every harvested row with whatever the deck declares "
                 f"now, and labelling last quarter's text as this quarter's cannot be undone.")

    print()
    for level, msg in findings:
        mark = {"ok": "  ok  ", "WARN": " WARN ", "FAIL": " FAIL "}[level]
        print(f"[{mark}] {msg}")

    fails = sum(1 for l, _ in findings if l == "FAIL")
    warns = sum(1 for l, _ in findings if l == "WARN")
    print()
    if fails:
        print(f"NOT READY -- {fails} blocking problem(s), {warns} warning(s).")
        return 1
    print(f"READY -- 0 blocking problems, {warns} warning(s)." +
          ("  Read the warnings before drafting." if warns else ""))
    return 0


if __name__ == "__main__":
    sys.exit(main())
