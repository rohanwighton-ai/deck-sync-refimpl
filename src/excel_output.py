"""Excel (.xlsx) output: produce and incrementally update a linked data sheet.

See specs/excel-output.md. Stdlib-only (zipfile + xml.etree), matching
discovery.py's and verification.py's dependency-light approach: `.xlsx` is
itself a zipped OOXML package (SpreadsheetML), so no dependency (openpyxl et
al.) is needed for the minimal single-sheet structure this spec requires. No
`.xlsx` fixtures existed in test-fixtures/ for this spec (nothing was pulled
from upstream for it) -- this module is both the writer and the reader, so
its own tests are necessarily round-trip/self-consistency tests rather than
tests against an externally-produced file, same limitation the plan called
out before this was implemented.

Layout convention: column A is always the reserved instance-identity column
(header label "Instance ID"); columns B.. hold confirmed fields, one per
column, in first-seen (append) order. Field *identity* is the column's
header text, looked up by name on every read -- never assumed from position,
since a field's column index can grow as new fields are added. Column A's
position, by contrast, is a structural convention of this module, not a
field identity subject to that rule.
"""

from __future__ import annotations

import os
import re
import tempfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from zipfile import ZIP_DEFLATED, ZipFile

SSML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
CUSTOM_PROPS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

for _prefix, _uri in {"": SSML_NS}.items():
    ET.register_namespace(_prefix, _uri)

_XML_DECLARATION = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

INSTANCE_ID_HEADER = "Instance ID"
_DECK_REFERENCE_PROPERTY_NAME = "DeckReference"

SHEET_PART = "xl/worksheets/sheet1.xml"
CUSTOM_PROPS_PART = "docProps/custom.xml"

# Static boilerplate parts: never vary by data, so hand-written once rather
# than rebuilt via ElementTree on every write.
_CONTENT_TYPES = _XML_DECLARATION + (
    b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    b'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    b'<Default Extension="xml" ContentType="application/xml"/>'
    b'<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    b'<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    b'<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    b'<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>'
    b"</Types>"
)

_ROOT_RELS = _XML_DECLARATION + (
    b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    b'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
    b'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>'
    b"</Relationships>"
)

_WORKBOOK = _XML_DECLARATION + (
    b'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    b'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    b'<sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>'
    b"</workbook>"
)

_WORKBOOK_RELS = _XML_DECLARATION + (
    b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    b'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
    b'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    b"</Relationships>"
)

_STYLES = _XML_DECLARATION + (
    b'<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    b'<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
    b'<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
    b'<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
    b'<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    b'<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
    b"</styleSheet>"
)


@dataclass
class Sheet:
    # Stable reference back to the specific deck this sheet is paired with --
    # per specs/excel-output.md, never inferred solely from column-name
    # matching against multiple candidate decks. Stored as a real OOXML
    # custom document property (docProps/custom.xml), a dedicated slot for
    # exactly this kind of metadata, not a row/column that could collide
    # with real field/instance data.
    deck_reference: str
    # Confirmed fields, in first-seen order. Order is preserved on every
    # write (new fields append; existing ones never move), since order is
    # incidental to a spreadsheet's column layout but field *identity* (the
    # header name) is the only thing correctness depends on.
    fields: list[str] = field(default_factory=list)
    # Instance ids in first-seen (row) order -- append-only, never reordered
    # or re-keyed by position, per specs/excel-output.md's persistent-identity
    # requirement.
    instance_order: list[str] = field(default_factory=list)
    # instance_id -> {field_name: value}. A field absent from a given
    # instance's dict means "no value harvested for this field on this
    # instance yet", not "blanked out" -- rendered as an omitted cell, never
    # a forced empty string.
    rows: dict[str, dict[str, str]] = field(default_factory=dict)


def _col_letter(n: int) -> str:
    """1-based column index -> Excel column letters (1 -> 'A', 27 -> 'AA')."""
    letters = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


_CELL_REF_RE = re.compile(r"^([A-Z]+)(\d+)$")


def _col_index_from_ref(cell_ref: str) -> int:
    """'B2' -> 2 (1-based column index), inverse of _col_letter."""
    match = _CELL_REF_RE.match(cell_ref)
    if match is None:
        raise ValueError(f"not a valid cell reference: {cell_ref!r}")
    letters = match.group(1)
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n


def _inline_str_cell(parent: ET.Element, cell_ref: str, value: str) -> None:
    c = ET.SubElement(parent, f"{{{SSML_NS}}}c", {"r": cell_ref, "t": "inlineStr"})
    is_el = ET.SubElement(c, f"{{{SSML_NS}}}is")
    t_el = ET.SubElement(is_el, f"{{{SSML_NS}}}t")
    t_el.text = value


def _build_sheet_xml(sheet: Sheet) -> bytes:
    worksheet = ET.Element(f"{{{SSML_NS}}}worksheet")
    sheet_data = ET.SubElement(worksheet, f"{{{SSML_NS}}}sheetData")

    header_row = ET.SubElement(sheet_data, f"{{{SSML_NS}}}row", {"r": "1"})
    _inline_str_cell(header_row, "A1", INSTANCE_ID_HEADER)
    for i, field_name in enumerate(sheet.fields):
        _inline_str_cell(header_row, f"{_col_letter(i + 2)}1", field_name)

    for row_offset, instance_id in enumerate(sheet.instance_order):
        r = row_offset + 2  # row 1 is the header
        row_el = ET.SubElement(sheet_data, f"{{{SSML_NS}}}row", {"r": str(r)})
        _inline_str_cell(row_el, f"A{r}", instance_id)
        values = sheet.rows.get(instance_id, {})
        for i, field_name in enumerate(sheet.fields):
            if field_name in values:
                _inline_str_cell(row_el, f"{_col_letter(i + 2)}{r}", values[field_name])

    return _XML_DECLARATION + ET.tostring(worksheet, encoding="unicode").encode("utf-8")


def _build_custom_props_xml(deck_reference: str) -> bytes:
    props = ET.Element(f"{{{CUSTOM_PROPS_NS}}}Properties")
    ET.register_namespace("vt", VT_NS)
    prop = ET.SubElement(
        props,
        f"{{{CUSTOM_PROPS_NS}}}property",
        {"fmtid": "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}", "pid": "2", "name": _DECK_REFERENCE_PROPERTY_NAME},
    )
    value_el = ET.SubElement(prop, f"{{{VT_NS}}}lpwstr")
    value_el.text = deck_reference
    return _XML_DECLARATION + ET.tostring(props, encoding="unicode").encode("utf-8")


def _write_xlsx(path: str, sheet: Sheet) -> None:
    """Regenerate the whole .xlsx from `sheet`. Safe against a crash mid-write:
    written to a sibling temp file first, then swapped into place, same
    pattern verification.py's _write_part uses for its own writes.
    """
    directory = os.path.dirname(os.path.abspath(path)) or "."
    fd, tmp_path = tempfile.mkstemp(suffix=".xlsx", dir=directory)
    os.close(fd)
    try:
        with ZipFile(tmp_path, "w", ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", _CONTENT_TYPES)
            z.writestr("_rels/.rels", _ROOT_RELS)
            z.writestr("xl/workbook.xml", _WORKBOOK)
            z.writestr("xl/_rels/workbook.xml.rels", _WORKBOOK_RELS)
            z.writestr("xl/styles.xml", _STYLES)
            z.writestr(SHEET_PART, _build_sheet_xml(sheet))
            z.writestr(CUSTOM_PROPS_PART, _build_custom_props_xml(sheet.deck_reference))
        os.replace(tmp_path, path)
    except BaseException:
        os.remove(tmp_path)
        raise


def create_sheet(path: str, deck_reference: str) -> None:
    """Create a fresh data sheet with no fields/rows yet, bound to
    `deck_reference`. Refuses to overwrite an existing file -- a second
    "create" is almost certainly a mistake (it would silently discard
    whatever the file already holds), consistent with this project's
    never-assume, never-silently-drop-data approach elsewhere.
    """
    if os.path.exists(path):
        raise FileExistsError(f"refusing to overwrite existing sheet: {path}")
    _write_xlsx(path, Sheet(deck_reference=deck_reference))


def read_sheet(path: str) -> Sheet:
    """Read a data sheet back into a Sheet. Fields are recovered from the
    header row (row 1), looked up by column *position* only to pair a
    header cell with the data cells below it in the same column -- the
    field's *identity* is still its header text, not that position, and a
    field's column index is free to differ across files.
    """
    with ZipFile(path) as z:
        with z.open(SHEET_PART) as f:
            worksheet = ET.parse(f).getroot()
        with z.open(CUSTOM_PROPS_PART) as f:
            custom_props = ET.parse(f).getroot()

    deck_reference = ""
    for prop in custom_props.findall(f"{{{CUSTOM_PROPS_NS}}}property"):
        if prop.get("name") == _DECK_REFERENCE_PROPERTY_NAME:
            value_el = prop.find(f"{{{VT_NS}}}lpwstr")
            deck_reference = value_el.text or "" if value_el is not None else ""

    rows_el = worksheet.findall(f".//{{{SSML_NS}}}row")
    fields: list[str] = []
    instance_order: list[str] = []
    rows: dict[str, dict[str, str]] = {}

    field_by_col: dict[int, str] = {}
    for row_index, row_el in enumerate(rows_el):
        cells: dict[str, str] = {}
        for c in row_el.findall(f"{{{SSML_NS}}}c"):
            ref = c.get("r")
            if ref is None:
                continue  # every real cell carries an r attribute; nothing to key this by otherwise
            is_el = c.find(f"{{{SSML_NS}}}is")
            t_el = is_el.find(f"{{{SSML_NS}}}t") if is_el is not None else None
            cells[ref] = t_el.text or "" if t_el is not None else ""

        if row_index == 0:
            for ref, value in cells.items():
                if ref == "A1":
                    continue
                field_by_col[_col_index_from_ref(ref)] = value
            fields = [field_by_col[col] for col in sorted(field_by_col)]
            continue

        instance_id = cells.get(f"A{row_el.get('r')}")
        if instance_id is None:
            continue
        instance_order.append(instance_id)
        values = {}
        for ref, value in cells.items():
            col = _col_index_from_ref(ref)
            if col == 1:
                continue
            field_name = field_by_col.get(col)
            if field_name is not None:
                values[field_name] = value
        rows[instance_id] = values

    return Sheet(deck_reference=deck_reference, fields=fields, instance_order=instance_order, rows=rows)


def upsert_row(path: str, instance_id: str, values: dict[str, str]) -> None:
    """Read-merge-write per specs/excel-output.md's no-data-loss requirement:
    add any of `values`'s keys not already a known field as a new column
    (appended, never replacing/reordering existing columns), then create or
    update `instance_id`'s row.

    A new instance is appended as a new row, seeded entirely from `values`
    (the harvested values), never left blank for the user to backfill. An
    existing instance is merged: only the given keys are updated, so a
    partial call (e.g. re-syncing one changed field) can never blank out
    that instance's other, previously-populated fields.
    """
    sheet = read_sheet(path)

    for field_name in values:
        if field_name not in sheet.fields:
            sheet.fields.append(field_name)

    if instance_id not in sheet.rows:
        sheet.instance_order.append(instance_id)
        sheet.rows[instance_id] = dict(values)
    else:
        sheet.rows[instance_id].update(values)

    _write_xlsx(path, sheet)
