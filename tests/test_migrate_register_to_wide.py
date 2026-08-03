"""Tests for vba/tools/migrate_register_to_wide.py.

These exercise the pivot on synthetic long rows -- no Excel, no COM. The three
refusals each get a test that MAKES THEM FIRE, because a guard nobody has
watched fail is not evidence of anything (see CLAUDE.md; this project has been
bitten by a silently-empty read twice).
"""

import os
import sys

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "vba", "tools"))

from migrate_register_to_wide import ALL_SENTINEL, Refused, pivot  # noqa: E402

HEADER = ["Quarter", "EntityCode", "SlideType", "FieldID", "FieldType", "Value",
          "CharCount", "Status", "UpdatedDate"]


def row(quarter, entity, field, value, slide_type="q"):
    return [quarter, entity, slide_type, field, "Text", value, str(len(value)), "Seed", "2026-08-03"]


# --- the shape it exists to produce -----------------------------------------


def test_all_row_is_copied_into_every_period_row():
    p = pivot([HEADER,
               row(ALL_SENTINEL, "P1", "PROJECT_NAME", "Widget"),
               row("FY26Q4", "P1", "PROJECT_STATUS", "In Progress"),
               row("FY27Q1", "P1", "PROJECT_STATUS", "Project Closed")],
              fallback_period=None)

    assert list(p.rows) == [("P1", "FY26Q4"), ("P1", "FY27Q1")]
    assert p.rows[("P1", "FY26Q4")] == {"PROJECT_NAME": "Widget", "PROJECT_STATUS": "In Progress"}
    assert p.rows[("P1", "FY27Q1")] == {"PROJECT_NAME": "Widget", "PROJECT_STATUS": "Project Closed"}
    assert p.carried == 2


def test_a_period_value_shadows_the_all_value_and_is_counted():
    p = pivot([HEADER,
               row(ALL_SENTINEL, "P1", "ABOUT_BODY", "carried"),
               row("FY26Q4", "P1", "ABOUT_BODY", "this quarter")],
              fallback_period=None)

    assert p.rows[("P1", "FY26Q4")] == {"ABOUT_BODY": "this quarter"}
    assert p.overridden == 1
    assert p.carried == 0


def test_a_field_missing_for_one_period_is_simply_absent():
    # KEY_EVENTS_BODY has no FY27Q1 row on the rig -- next quarter's events are
    # not written yet. Absent, not blank-and-carried-forward.
    p = pivot([HEADER,
               row("FY26Q4", "P1", "KEY_EVENTS_BODY", "shipped"),
               row("FY26Q4", "P1", "PROJECT_STATUS", "In Progress"),
               row("FY27Q1", "P1", "PROJECT_STATUS", "In Progress")],
              fallback_period=None)

    assert "KEY_EVENTS_BODY" not in p.rows[("P1", "FY27Q1")]


def test_unknown_fields_keep_their_order_after_the_known_ones():
    p = pivot([HEADER,
               row("FY26Q4", "P1", "ZZ_CUSTOM", "x"),
               row("FY26Q4", "P1", "PROJECT_CODE", "P1")],
              fallback_period=None)

    assert p.fields == ["PROJECT_CODE", "ZZ_CUSTOM"]


# --- the refusals, each made to fire ----------------------------------------


def test_refuses_an_entity_that_has_only_all_rows():
    with pytest.raises(Refused, match="nowhere to go"):
        pivot([HEADER, row(ALL_SENTINEL, "P9", "PROJECT_NAME", "orphan")],
              fallback_period=None)


def test_fallback_period_places_an_all_only_entity_and_names_it():
    p = pivot([HEADER, row(ALL_SENTINEL, "P9", "PROJECT_NAME", "orphan")],
              fallback_period="FY26Q4")

    assert p.rows[("P9", "FY26Q4")] == {"PROJECT_NAME": "orphan"}
    assert p.orphan_entities == ["P9"]


def test_refuses_two_rows_for_the_same_entity_period_and_field():
    with pytest.raises(Refused, match="appears twice"):
        pivot([HEADER,
               row("FY26Q4", "P1", "PROJECT_STATUS", "In Progress"),
               row("FY26Q4", "P1", "PROJECT_STATUS", "Project Closed")],
              fallback_period=None)


def test_refuses_a_duplicate_all_row_too():
    with pytest.raises(Refused, match="appears twice"):
        pivot([HEADER,
               row(ALL_SENTINEL, "P1", "ABOUT_BODY", "one"),
               row(ALL_SENTINEL, "P1", "ABOUT_BODY", "two")],
              fallback_period=None)


def test_refuses_a_register_mixing_slide_types():
    with pytest.raises(Refused, match="ONE slide type"):
        pivot([HEADER,
               row("FY26Q4", "P1", "PROJECT_STATUS", "In Progress", slide_type="q"),
               row("FY26Q4", "P2", "PROJECT_STATUS", "In Progress", slide_type="milestone")],
              fallback_period=None)


def test_refuses_a_register_missing_a_structural_column():
    with pytest.raises(Refused, match="Quarter"):
        pivot([["EntityCode", "SlideType", "FieldID", "FieldType", "Value"],
               ["P1", "q", "PROJECT_STATUS", "Text", "In Progress"]],
              fallback_period=None)
