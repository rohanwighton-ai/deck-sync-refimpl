"""Sync dispatch: given a type's known slide instances and its linked
Data-sheet rows, decide which sync-event case applies to each row/instance.

See specs/sync-operations.md for the requirements this implements, including
what's deliberately out of scope (physical slide duplication, cases 5 and 7).
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Sequence, Union

from discovery import Candidate
from excel_output import Sheet
from verification import InjectResult, inject_primitive


@dataclass(frozen=True)
class SlideInstance:
    """An already-resolved slide instance, as this module needs it. How
    instance_key/type_tag/field-to-shape correspondence actually gets read
    off a real deck is a separate concern -- specs/matching.md's match() plus
    a still-undecided identity-tag physical storage format (see
    IMPLEMENTATION_PLAN.md's notes) -- this module starts from an
    already-resolved SlideInstance rather than solving that here.

    instance_key/type_tag are None when a slide couldn't be classified at
    all (case 6 candidate): no recognized type tag, or no persistent
    instance key found on it.
    """

    part_path: str  # the pptx part this slide's shapes live in (for inject_primitive)
    instance_key: str | None
    type_tag: str | None
    field_shapes: dict[str, Candidate] = field(default_factory=dict)


@dataclass(frozen=True)
class NoChange:
    """Case 1: the instance's current values already match the Data-sheet row."""

    instance_key: str


@dataclass(frozen=True)
class InPlaceCorrection:
    """Case 4: one or more fields differed and were written + verified in
    place. No archival copy -- the period didn't change, only the value did.
    """

    instance_key: str
    changed_fields: dict[str, InjectResult]


@dataclass(frozen=True)
class NewRecord:
    """Case 3: no known instance carries this Data-sheet row's instance key --
    a new slide instance is needed, duplicated from the type's template.
    Physically performing that duplication is out of scope (see
    specs/sync-operations.md's Non-goals); this is the decision alone.
    """

    row_instance_key: str
    values: dict[str, str]
    reason: str


@dataclass(frozen=True)
class PeriodRollover:
    """Case 2: an explicit "add new quarterly record" command against a
    *specific* existing instance. The duplicate becomes the new period's
    slide; the original instance is left untouched as the permanent
    archival record. Only reachable via plan_period_rollover(), never from
    plan_routine_sync() -- see specs/sync-operations.md: never inferred from
    a value simply looking different.
    """

    source_instance_key: str
    new_values: dict[str, str]
    reason: str


@dataclass(frozen=True)
class Flagged:
    """Anything that doesn't cleanly resolve to cases 1/3/4 -- stopped and
    surfaced rather than forced into a case it doesn't satisfy. Today this
    is only case 6 (unclassified_slide); cases 5 and 7 are non-goals (see
    specs/sync-operations.md) and never produced here.
    """

    subject: str
    kind: str
    detail: str


SyncAction = Union[NoChange, InPlaceCorrection, NewRecord, Flagged]


def plan_routine_sync(path: str, instances: Sequence[SlideInstance], data_sheet: Sheet) -> list[SyncAction]:
    """Dispatch per specs/sync-operations.md's routine-sync rules: cases
    1, 3, and 4 only. Case 2 (period rollover) is a distinct, explicitly
    invoked operation (see plan_period_rollover()) and is never reachable
    from here, regardless of how different a value looks.

    Iterates the Data-sheet's rows (the same frame run-sync.md's own
    dispatch process uses) plus every given instance for case-6 detection --
    an instance with no recognized type/instance key can't even be checked
    against the Data-sheet, so it's flagged independently of that iteration.
    """
    actions: list[SyncAction] = []

    for instance in instances:
        if instance.type_tag is None or instance.instance_key is None:
            actions.append(
                Flagged(
                    subject=instance.part_path,
                    kind="unclassified_slide",
                    detail="no recognized type tag / persistent instance key -- flagged for reclassification, not guessed",
                )
            )

    known_by_key = {i.instance_key: i for i in instances if i.instance_key is not None}

    for instance_id in data_sheet.instance_order:
        row_values = data_sheet.rows.get(instance_id, {})
        known_instance = known_by_key.get(instance_id)

        if known_instance is None:
            actions.append(
                NewRecord(
                    row_instance_key=instance_id,
                    values=dict(row_values),
                    reason="no known slide instance carries this row's instance key",
                )
            )
            continue

        changed: dict[str, InjectResult] = {}
        for field_name, source_value in row_values.items():
            shape = known_instance.field_shapes.get(field_name)
            if shape is None:
                # This field isn't present on this instance's slide. Not one
                # of the named cases (1/3/4/6) -- structural drift like this
                # is adjacent to case 7 (deck-side conflict) territory, which
                # is a non-goal here. Skipped rather than guessed at.
                continue
            result = inject_primitive(path, known_instance.part_path, shape, source_value)
            if result.written:
                changed[field_name] = result

        if changed:
            actions.append(InPlaceCorrection(instance_key=instance_id, changed_fields=changed))
        else:
            actions.append(NoChange(instance_key=instance_id))

    return actions


def plan_period_rollover(instance: SlideInstance, new_values: dict[str, str]) -> PeriodRollover:
    """Case 2, per specs/sync-operations.md: only ever called explicitly,
    against one specific known instance, never dispatched automatically from
    plan_routine_sync(). Decides that `instance`'s current slide should be
    duplicated (left untouched as history) with `new_values` injected onto
    the duplicate under a fresh instance key -- physically performing that
    duplication is out of scope (see specs/sync-operations.md's Non-goals).
    """
    if instance.instance_key is None:
        raise ValueError("cannot roll over a period for an unclassified instance (no instance_key)")
    return PeriodRollover(
        source_instance_key=instance.instance_key,
        new_values=dict(new_values),
        reason="explicit period-rollover command",
    )
