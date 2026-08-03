#!/usr/bin/env python3
"""Pivot a long register into the wide, one-row-per-slide-per-period sheet.

WHY THIS EXISTS

The long register stores one row per project x field x quarter -- 220 rows on
the e2e rig, 1,848 on the real deck's shape. Rohan's model (2026-08-03) is that
a row is a SLIDE and carries its own Quarter, so FY26Q4 and FY27Q1 for one
project sit side by side and a deck picks up the rows for the period it
declares. `ExcelOutput.ReadSheetForPeriod` reads that shape. Nothing had ever
produced one, so the model was theory. This converts the rig.

WHAT HAPPENS TO `Quarter = ALL`

It stops existing. An ALL row meant "this value carries across every quarter" --
a sentinel Rohan had never been told about. In the wide sheet the value is
simply COPIED into each of the entity's period rows, which is the same outcome
with no concept attached, and is what `RollForwardPeriod` does going forward.

WHAT HAPPENS TO `Status`

It stops existing too. Only approved text ever reaches the sheet, so presence IS
approval. Both Seed (harvested off the slide) and Approved (drafted and ticked)
rows migrate: each is the text the slide should currently carry.

SAFETY

Never writes the input. Output must be a different path, and the OneDrive guard
from set_deck_period.py applies here for the same reason -- the original is not
a legitimate target for a tool that rewrites a sheet.

    python3 migrate_register_to_wide.py <in.xlsx> <out.xlsx> [--sheet Register]

Stdlib only (zipfile + ElementTree), matching src/excel_output.py.
"""

from __future__ import annotations

import argparse
import re
import shutil
import sys
import xml.etree.ElementTree as ET
from collections import OrderedDict
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
RELNS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

ALL_SENTINEL = "ALL"

# Mirrors ExcelOutput.bas's two structural column names. Anything else on the
# sheet is a field.
INSTANCE_ID_HEADER = "Instance ID"
QUARTER_HEADER = "Quarter"

# Reused from the source sheet's own style table so the migrated sheet looks
# like the one it replaces: 9/10 header (plain/wrapped), 7/11 body.
S_HEAD, S_HEAD_WRAP = "9", "10"
S_BODY, S_BODY_WRAP = "7", "11"

# Field order in the output, chosen for reading left to right: what the project
# is called, then its state, then the prose. Purely cosmetic -- the reader
# matches columns by header NAME, never by position.
PREFERRED_FIELD_ORDER = [
    "PROJECT_CODE",
    "PROJECT_NAME",
    "PROJECT_STATUS",
    "ABOUT_BODY",
    "KEY_EVENTS_BODY",
]
WRAPPED_FIELDS = {"ABOUT_BODY", "KEY_EVENTS_BODY"}


class Refused(Exception):
    """A condition the migration will not guess its way past."""


# ---------------------------------------------------------------------
# Read
# ---------------------------------------------------------------------


def _shared_strings(z: ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    root = ET.fromstring(z.read("xl/sharedStrings.xml"))
    return ["".join(t.text or "" for t in si.iter(f"{NS}t")) for si in root.findall(f"{NS}si")]


def _sheet_paths(z: ZipFile) -> "OrderedDict[str, str]":
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    target = {r.get("Id"): r.get("Target") for r in rels}
    out: OrderedDict[str, str] = OrderedDict()
    for s in wb.find(f"{NS}sheets"):
        t = target[s.get(f"{RELNS}id")]
        out[s.get("name")] = "xl/" + t.lstrip("/").removeprefix("xl/")
    return out


def _col_index(ref: str) -> int:
    n = 0
    for ch in re.match(r"([A-Z]+)", ref).group(1):
        n = n * 26 + (ord(ch) - 64)
    return n


def _col_letter(n: int) -> str:
    out = ""
    while n:
        n, rem = divmod(n - 1, 26)
        out = chr(65 + rem) + out
    return out


def _rows(z: ZipFile, path: str, sst: list[str]) -> list[list[str]]:
    root = ET.fromstring(z.read(path))
    data = root.find(f"{NS}sheetData")
    out: list[list[str]] = []
    for r in data.findall(f"{NS}row"):
        cells: dict[int, str] = {}
        for c in r.findall(f"{NS}c"):
            v = c.find(f"{NS}v")
            if c.get("t") == "inlineStr":
                is_ = c.find(f"{NS}is")
                text = "".join(t.text or "" for t in is_.iter(f"{NS}t")) if is_ is not None else ""
            elif c.get("t") == "s" and v is not None:
                text = sst[int(v.text)]
            elif v is not None:
                text = v.text or ""
            else:
                text = ""
            cells[_col_index(c.get("r"))] = text
        out.append([cells.get(i, "") for i in range(1, max(cells) + 1)] if cells else [])
    return out


# ---------------------------------------------------------------------
# Pivot
# ---------------------------------------------------------------------


class Pivot:
    def __init__(self) -> None:
        self.fields: list[str] = []
        self.rows: "OrderedDict[tuple[str, str], dict[str, str]]" = OrderedDict()
        self.carried = 0        # values copied from an ALL row into a period row
        self.overridden = 0     # ALL value shadowed by a real period value
        self.orphan_entities: list[str] = []  # ALL rows only, no period to live in


def pivot(long_rows: list[list[str]], fallback_period: str | None) -> Pivot:
    header = long_rows[0]
    idx = {h: i for i, h in enumerate(header)}
    for required in ("Quarter", "EntityCode", "SlideType", "FieldID", "Value"):
        if required not in idx:
            raise Refused(f"the long register has no {required!r} column; got {header}")

    def cell(row: list[str], name: str) -> str:
        i = idx[name]
        return row[i].strip() if i < len(row) else ""

    body = [r for r in long_rows[1:] if r and cell(r, "EntityCode")]

    slide_types = {cell(r, "SlideType") for r in body}
    if len(slide_types) > 1:
        raise Refused(
            "a wide sheet holds ONE slide type -- this register mixes "
            f"{sorted(slide_types)}, which needs one output sheet each"
        )

    p = Pivot()

    # Field order: preferred names first (only those actually present), then any
    # others in first-appearance order, so an unknown field is never dropped.
    seen_fields = list(OrderedDict.fromkeys(cell(r, "FieldID") for r in body))
    p.fields = [f for f in PREFERRED_FIELD_ORDER if f in seen_fields]
    p.fields += [f for f in seen_fields if f not in p.fields]

    entities = list(OrderedDict.fromkeys(cell(r, "EntityCode") for r in body))

    periods: "OrderedDict[str, list[str]]" = OrderedDict()
    all_values: dict[str, dict[str, str]] = {}
    period_values: dict[tuple[str, str], dict[str, str]] = {}

    for r in body:
        ent, q, fid, val = cell(r, "EntityCode"), cell(r, "Quarter"), cell(r, "FieldID"), cell(r, "Value")
        if q == ALL_SENTINEL:
            bucket = all_values.setdefault(ent, {})
        else:
            periods.setdefault(ent, [])
            if q not in periods[ent]:
                periods[ent].append(q)
            bucket = period_values.setdefault((ent, q), {})
        if fid in bucket:
            raise Refused(
                f"{ent} / {q} / {fid} appears twice in the long register -- "
                "whichever sat lower would silently win, so this refuses instead"
            )
        bucket[fid] = val

    for ent in entities:
        ent_periods = periods.get(ent, [])
        if not ent_periods:
            # ALL rows and nothing else: no period row for the value to live in.
            # Dropping it silently is exactly the failure this project keeps
            # paying for, so it is named and, without a fallback, refused.
            p.orphan_entities.append(ent)
            if fallback_period is None:
                continue
            ent_periods = [fallback_period]

        for q in ent_periods:
            values = dict(all_values.get(ent, {}))
            p.carried += len(values)
            for fid, val in period_values.get((ent, q), {}).items():
                if fid in values:
                    p.overridden += 1
                    p.carried -= 1
                values[fid] = val
            p.rows[(ent, q)] = values

    if p.orphan_entities and fallback_period is None:
        raise Refused(
            f"{len(p.orphan_entities)} entit(ies) carry only {ALL_SENTINEL} rows and no period "
            f"row: {', '.join(p.orphan_entities[:5])}"
            f"{' ...' if len(p.orphan_entities) > 5 else ''}\n"
            "In the wide model a value lives on a period row, so these have nowhere to go.\n"
            "Re-run with --fallback-period FY26Q4 to place them, having decided that is right."
        )

    return p


# ---------------------------------------------------------------------
# Write
# ---------------------------------------------------------------------


def _esc(s: str) -> str:
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def _cell_xml(ref: str, style: str, value: str) -> str:
    # Inline strings: no sharedStrings surgery, and every value round-trips as
    # the exact harvested string (the invariant UpsertRow's header argues for).
    if value == "":
        return f'<c r="{ref}" s="{style}"/>'
    space = ' xml:space="preserve"' if value != value.strip() else ""
    return f'<c r="{ref}" s="{style}" t="inlineStr"><is><t{space}>{_esc(value)}</t></is></c>'


def build_sheet_xml(original_xml: str, p: Pivot) -> str:
    headers = [INSTANCE_ID_HEADER, QUARTER_HEADER] + p.fields
    n_cols, n_rows = len(headers), len(p.rows) + 1

    cols = []
    for i, h in enumerate(headers, start=1):
        width = "70.6328125" if h in WRAPPED_FIELDS else "18.6328125"
        style = S_BODY_WRAP if h in WRAPPED_FIELDS else S_BODY
        cols.append(f'<col min="{i}" max="{i}" width="{width}" style="{style}" customWidth="1"/>')
    cols.append(f'<col min="{n_cols + 1}" max="16384" width="8.7265625" style="2"/>')

    body = [
        '<row r="1" spans="1:%d" s="5" customFormat="1" ht="26" customHeight="1">%s</row>'
        % (
            n_cols,
            "".join(
                _cell_xml(f"{_col_letter(i)}1", S_HEAD_WRAP if h in WRAPPED_FIELDS else S_HEAD, h)
                for i, h in enumerate(headers, start=1)
            ),
        )
    ]
    for n, ((ent, q), values) in enumerate(p.rows.items(), start=2):
        line = [
            _cell_xml(f"A{n}", S_BODY, ent),
            _cell_xml(f"B{n}", S_BODY, q),
        ]
        for i, f in enumerate(p.fields, start=3):
            line.append(
                _cell_xml(
                    f"{_col_letter(i)}{n}",
                    S_BODY_WRAP if f in WRAPPED_FIELDS else S_BODY,
                    values.get(f, ""),
                )
            )
        body.append(f'<row r="{n}" spans="1:{n_cols}" ht="40" customHeight="1">{"".join(line)}</row>')

    head = original_xml.split("<dimension")[0]
    validation = ""
    if "PROJECT_STATUS" in p.fields:
        # The long sheet validated a scatter of individual Value cells; a whole
        # column is the same rule, stated once.
        letter = _col_letter(headers.index("PROJECT_STATUS") + 1)
        validation = (
            '<dataValidations count="1"><dataValidation type="list" allowBlank="1" '
            'showInputMessage="1" showErrorMessage="1" '
            f'sqref="{letter}2:{letter}{n_rows}">'
            '<formula1>"In Progress,Not Started,Project Closed"</formula1>'
            "</dataValidation></dataValidations>"
        )

    return (
        head
        + f'<dimension ref="A1:{_col_letter(n_cols)}{n_rows}"/>'
        + '<sheetViews><sheetView workbookViewId="0">'
        '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
        '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>'
        + '<sheetFormatPr defaultRowHeight="10.5"/>'
        + f"<cols>{''.join(cols)}</cols>"
        + f"<sheetData>{''.join(body)}</sheetData>"
        + validation
        + '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
        + "</worksheet>"
    )


def write_copy(src: Path, dst: Path, sheet_path: str, new_xml: str) -> None:
    with ZipFile(src) as zin, ZipFile(dst, "w", ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = new_xml.encode("utf-8") if item.filename == sheet_path else zin.read(item.filename)
            zout.writestr(item, data)


# ---------------------------------------------------------------------


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("source")
    ap.add_argument("dest")
    ap.add_argument("--sheet", default="Register")
    ap.add_argument("--fallback-period", default=None,
                    help="period to place entities that carry only ALL rows")
    a = ap.parse_args()

    src, dst = Path(a.source), Path(a.dest)
    if not src.exists():
        print(f"no such workbook: {src}", file=sys.stderr)
        return 1
    if src.resolve() == dst.resolve():
        print("REFUSED: this never rewrites its input. Give a different output path.", file=sys.stderr)
        return 1
    if "OneDrive" in str(dst):
        print("REFUSED: not writing into OneDrive. Work on a copy.", file=sys.stderr)
        return 1

    with ZipFile(src) as z:
        sheets = _sheet_paths(z)
        if a.sheet not in sheets:
            print(f"no sheet named {a.sheet!r}; workbook has {list(sheets)}", file=sys.stderr)
            return 1
        sheet_path = sheets[a.sheet]
        long_rows = _rows(z, sheet_path, _shared_strings(z))
        original_xml = z.read(sheet_path).decode("utf-8")

    try:
        p = pivot(long_rows, a.fallback_period)
    except Refused as e:
        print(f"REFUSED: {e}", file=sys.stderr)
        return 1

    write_copy(src, dst, sheet_path, build_sheet_xml(original_xml, p))

    by_period: "OrderedDict[str, int]" = OrderedDict()
    for _, q in p.rows:
        by_period[q] = by_period.get(q, 0) + 1

    print(f"source:  {src}  ({len(long_rows) - 1} long rows)")
    print(f"dest:    {dst}")
    print(f"sheet:   {a.sheet}  ->  {len(p.rows)} rows x {len(p.fields) + 2} columns")
    print(f"fields:  {', '.join(p.fields)}")
    print(f"periods: {', '.join(f'{q} ({n} rows)' for q, n in by_period.items())}")
    print(f"carried from {ALL_SENTINEL}: {p.carried} values"
          + (f"; {p.overridden} shadowed by a period value" if p.overridden else ""))
    if p.orphan_entities:
        print(f"placed in {a.fallback_period} (had only {ALL_SENTINEL} rows): "
              f"{len(p.orphan_entities)} entities")
    return 0


if __name__ == "__main__":
    sys.exit(main())
