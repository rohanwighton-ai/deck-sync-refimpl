#!/usr/bin/env python3
"""Fast static checks over vba/*.bas, for the compile errors that are cheap to
find here and expensive to find any other way.

Why this exists. A VBA compile error is not reported where it is written --
Application.Run surfaces it as "Sub or function not defined" / "Variable not
defined" / "User-defined type not defined" attributed to whichever module
REFERENCED the broken thing, and the import log says every module imported
fine. So the ~8-minute suite run dies with a message pointing at the wrong
file. That happened twice on 2026-07-30, from the same root cause, and the
second time was after the pattern had already been documented in AGENTS.md --
because the check was a thing to remember rather than a thing to run.

Exit code 1 on any finding, so this can gate the suite run.

Deliberately narrow. These are pattern-matchable certainties, not style
opinions -- a checker that reports maybes trains you to ignore it.
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

PROC_RE = re.compile(r"^(?:Public |Private |Friend )?(?:Sub|Function|Property(?:\s+(?:Get|Let|Set))?)\s+(\w+)")
DECL_RE = re.compile(r"^(?:Public |Private )?(Type|Const|Enum)\s+(\w+)")
END_PROC_RE = re.compile(r"^End (?:Sub|Function|Property)\b")
DIM_RE = re.compile(r"^Dim\s+(.*)$")
NAME_AS_RE = re.compile(r"\b(\w+)\s*(?:\([^)]*\))?\s+As\s+\w")


def check_declaration_order(path: Path, lines: list[str]) -> list[str]:
    """Module-level Type/Const/Enum must precede the first procedure.

    VBA rejects a module-level declaration that appears after any Sub/Function,
    and reports it from a DIFFERENT module. Procedure-local Const is legal and
    is excluded by only matching column-0 declarations.
    """
    findings: list[str] = []
    first_proc: tuple[int, str] | None = None
    for i, raw in enumerate(lines, 1):
        if raw[:1].isspace():
            continue  # indented => inside a procedure body
        if first_proc is None:
            m = PROC_RE.match(raw)
            if m:
                first_proc = (i, m.group(1))
                continue
        m = DECL_RE.match(raw)
        if m and first_proc:
            findings.append(
                f"{path}:{i}: module-level {m.group(1)} '{m.group(2)}' declared AFTER "
                f"procedure '{first_proc[1]}' (line {first_proc[0]}). Move all "
                f"Type/Const/Enum above the first procedure -- VBA reports this "
                f"error in whichever OTHER module references it."
            )
    return findings


def check_duplicate_dims(path: Path, lines: list[str]) -> list[str]:
    """The same name Dim'd twice in one procedure is a compile error.

    VBA scopes Dim to the whole procedure regardless of the block it sits in,
    so two `Dim i As Long` inside separate If branches do not coexist.
    """
    findings: list[str] = []
    proc: str | None = None
    seen: dict[str, int] = {}
    for i, raw in enumerate(lines, 1):
        s = raw.strip()
        m = PROC_RE.match(s)
        if m:
            proc, seen = m.group(1), {}
            continue
        if END_PROC_RE.match(s):
            proc = None
            continue
        if proc is None:
            continue
        d = DIM_RE.match(s)
        if not d:
            continue
        for name in NAME_AS_RE.findall(d.group(1)):
            if name in seen:
                findings.append(
                    f"{path}:{i}: '{name}' Dim'd again in procedure '{proc}' "
                    f"(already at line {seen[name]}) -- Dim is procedure-scoped in VBA."
                )
            else:
                seen[name] = i
    return findings


def check_empty_redim(path: Path, lines: list[str]) -> list[str]:
    """`ReDim arr(1 To 0)` raises Err 9 at runtime -- see AGENTS.md.

    Runtime, not compile, so the suite catches it; but it catches it as one
    failed test rather than as the systematic mistake it tends to be.
    """
    findings: list[str] = []
    for i, raw in enumerate(lines, 1):
        if re.search(r"ReDim\s+(?:Preserve\s+)?\w+\s*\(\s*1\s+To\s+0\s*\)", raw):
            findings.append(
                f"{path}:{i}: ReDim to (1 To 0) raises 'Subscript out of range'. "
                f"Leave the array unallocated and let the first ReDim Preserve size it."
            )
    return findings


def main() -> int:
    root = Path(__file__).resolve().parent.parent
    files = sorted(root.glob("*.bas")) + sorted((root / "tests").glob("*.bas"))
    if not files:
        print(f"no .bas files found under {root}", file=sys.stderr)
        return 1

    findings: list[str] = []
    for path in files:
        lines = path.read_text(errors="replace").splitlines()
        rel = path.relative_to(root.parent)
        findings += check_declaration_order(rel, lines)
        findings += check_duplicate_dims(rel, lines)
        findings += check_empty_redim(rel, lines)

    if findings:
        print(f"=== {len(findings)} static finding(s) ===")
        for f in findings:
            print(f"  {f}")
        return 1

    print(f"static checks clean across {len(files)} module(s)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
