"""Tests for src/excel_output.py against specs/excel-output.md.

No .xlsx fixtures exist in test-fixtures/ for this spec (nothing was pulled
from upstream for it, per IMPLEMENTATION_PLAN.md's notes) -- excel_output.py
is both the writer and reader, so these are round-trip/self-consistency
tests, not tests against an externally-produced file.
"""

import os
import sys
import tempfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from excel_output import INSTANCE_ID_HEADER, create_sheet, read_sheet, upsert_row  # noqa: E402


def _temp_path():
    fd, path = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    os.remove(path)  # create_sheet must be the thing that creates it
    return path


def test_create_sheet_seeds_deck_reference_and_no_fields_or_rows():
    path = _temp_path()
    try:
        create_sheet(path, "deck-v1")
        sheet = read_sheet(path)

        assert sheet.deck_reference == "deck-v1"
        assert sheet.fields == []
        assert sheet.instance_order == []
        assert sheet.rows == {}
    finally:
        os.remove(path)


def test_create_sheet_refuses_to_overwrite_an_existing_file():
    path = _temp_path()
    try:
        create_sheet(path, "deck-v1")
        try:
            create_sheet(path, "deck-v2")
            assert False, "expected FileExistsError"
        except FileExistsError:
            pass
        # Original deck reference must survive the refused overwrite attempt.
        assert read_sheet(path).deck_reference == "deck-v1"
    finally:
        os.remove(path)


def test_upsert_row_seeds_a_new_instance_from_harvested_values_not_blank():
    path = _temp_path()
    try:
        create_sheet(path, "deck-v1")
        upsert_row(path, "slide-1", {"Title": "Q3 Revenue", "Date": "2026-07"})

        sheet = read_sheet(path)
        assert sheet.fields == ["Title", "Date"]
        assert sheet.instance_order == ["slide-1"]
        assert sheet.rows["slide-1"] == {"Title": "Q3 Revenue", "Date": "2026-07"}
    finally:
        os.remove(path)


def test_upsert_row_new_field_appends_a_column_without_touching_existing_data():
    path = _temp_path()
    try:
        create_sheet(path, "deck-v1")
        upsert_row(path, "slide-1", {"Title": "Q3 Revenue", "Date": "2026-07"})

        # A later sync introduces a field this instance didn't carry before.
        upsert_row(path, "slide-1", {"Region": "APAC"})

        sheet = read_sheet(path)
        assert sheet.fields == ["Title", "Date", "Region"]  # appended, not reordered
        assert sheet.rows["slide-1"] == {"Title": "Q3 Revenue", "Date": "2026-07", "Region": "APAC"}
    finally:
        os.remove(path)


def test_upsert_row_partial_update_merges_rather_than_replacing():
    path = _temp_path()
    try:
        create_sheet(path, "deck-v1")
        upsert_row(path, "slide-1", {"Title": "Q3 Revenue", "Date": "2026-07"})

        # Only Title changes; Date must survive this call untouched.
        upsert_row(path, "slide-1", {"Title": "Q3 Revenue (revised)"})

        sheet = read_sheet(path)
        assert sheet.rows["slide-1"] == {"Title": "Q3 Revenue (revised)", "Date": "2026-07"}
    finally:
        os.remove(path)


def test_upsert_row_new_instance_does_not_disturb_existing_rows():
    path = _temp_path()
    try:
        create_sheet(path, "deck-v1")
        upsert_row(path, "slide-1", {"Title": "Q3 Revenue", "Date": "2026-07"})
        upsert_row(path, "slide-2", {"Title": "Q4 Revenue"})

        sheet = read_sheet(path)
        assert sheet.instance_order == ["slide-1", "slide-2"]  # append order, never reordered
        assert sheet.rows["slide-1"] == {"Title": "Q3 Revenue", "Date": "2026-07"}  # untouched
        # slide-2 genuinely has no Date yet -- absent, not forced to an empty string.
        assert "Date" not in sheet.rows["slide-2"]
        assert sheet.rows["slide-2"]["Title"] == "Q4 Revenue"
    finally:
        os.remove(path)


def test_read_sheet_preserves_field_and_instance_order_across_many_writes():
    path = _temp_path()
    try:
        create_sheet(path, "deck-v1")
        upsert_row(path, "slide-3", {"Zeta": "z"})
        upsert_row(path, "slide-1", {"Alpha": "a"})
        upsert_row(path, "slide-2", {"Zeta": "z2", "Alpha": "a2"})

        sheet = read_sheet(path)
        # First-seen order, not alphabetical or re-sorted.
        assert sheet.fields == ["Zeta", "Alpha"]
        assert sheet.instance_order == ["slide-3", "slide-1", "slide-2"]
    finally:
        os.remove(path)


def test_header_row_reserves_column_a_for_instance_id():
    path = _temp_path()
    try:
        create_sheet(path, "deck-v1")
        upsert_row(path, "slide-1", {"Title": "Q3 Revenue"})

        with __import__("zipfile").ZipFile(path) as z:
            import xml.etree.ElementTree as ET

            with z.open("xl/worksheets/sheet1.xml") as f:
                root = ET.parse(f).getroot()
        ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
        header_cells = root.findall(f".//{ns}row[@r='1']/{ns}c")
        a1 = next(c for c in header_cells if c.get("r") == "A1")
        assert a1.find(f"{ns}is/{ns}t").text == INSTANCE_ID_HEADER
    finally:
        os.remove(path)
