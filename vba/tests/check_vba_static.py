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


def code_only(line: str) -> str:
    """Reduce a line to executable code: comment removed, string contents blanked.

    Every check here is about CODE. Scanning comment text too makes the checker
    fire on its own documentation -- ReviewQueue.bas carries the comment "Never
    pre-ReDim to (1 To 0)", which check_empty_redim matched (`\\w+` bound to the
    word "to"), reporting a correct line as a defect on 2026-07-31.

    That failure mode is worse than it looks. This file's own docstring says a
    checker that reports maybes trains you to ignore it -- and a checker that
    flags the comment WARNING ABOUT a mistake, on a line that does not make it,
    trains you to ignore it faster than a mere maybe would.

    String CONTENTS are blanked for the same reason, one step further out: a
    literal is data, not code, and `MsgBox "use ReDim arr(1 To 0) to..."` is no
    more a defect than the comment is. Blanking rather than deleting keeps every
    column where it was, so the quote-toggle stays honest on the rest of the line.

    Quote-aware because `'` is legal inside a string literal (a delimiter, an
    apostrophe in a message). VBA escapes a double quote by doubling it, which
    needs no special handling: the doubled pair toggles the flag off and straight
    back on, and both quote characters are themselves preserved.
    """
    out: list[str] = []
    in_string = False
    for ch in line:
        if ch == '"':
            in_string = not in_string
            out.append(ch)
        elif ch == "'" and not in_string:
            break
        else:
            out.append(" " if in_string else ch)
    return "".join(out)


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


# VBA keywords that are legal-looking but illegal as a variable name. Not the
# full reserved list -- only ones a person plausibly reaches for as an
# identifier. `Empty` cost an 8-minute suite run on 2026-07-31: `Dim empty As
# ReviewQueueSet` compiles nowhere and fails as a bare "Syntax error" modal that
# blocks the whole headless run, so the failure arrives as a HANG with no output
# rather than as a message naming the line.
RESERVED_NAMES = {
    "empty", "error", "date", "time", "name", "string", "single", "double",
    "integer", "long", "currency", "variant", "boolean", "byte", "object",
    "type", "stop", "next", "loop", "input", "output", "print", "write",
    "get", "put", "close", "open", "line", "circle", "scale", "width",
    "property", "option", "resume", "select", "then", "to", "step", "is",
    "like", "mod", "not", "and", "or", "xor", "new", "set", "let", "call",
    "end", "exit", "for", "each", "in", "do", "while", "until", "if", "else",
}

# The trailing `As` is what keeps this off DECLARATION forms. Without it,
# `Public Type ReviewItem` binds the group to the word "Type" and the check
# fires on 27 lines of working code across every module -- which is the
# cry-wolf failure this file's docstring warns about, self-inflicted.
# Known gap, accepted: a bare `Dim empty` with no `As` clause is missed.
DIM_NAME_RE = re.compile(r"^\s*(?:Dim|Static|Public|Private|Const)\s+(\w+)\s+As\s", re.IGNORECASE)


def check_reserved_names(path: Path, lines: list[str]) -> list[str]:
    """A VBA keyword used as a variable name is a compile-time Syntax error.

    Worth catching here specifically because of HOW it fails: the headless
    driver gets no output at all and the process sits on a modal dialog, so it
    reads as a hang rather than an error. The 8 minutes are spent before you
    learn anything.
    """
    findings: list[str] = []
    for i, raw in enumerate(lines, 1):
        m = DIM_NAME_RE.match(raw)
        if m and m.group(1).lower() in RESERVED_NAMES:
            findings.append(
                f"{path}:{i}: '{m.group(1)}' is a VBA reserved word and cannot be a "
                f"variable name -- this is a compile-time Syntax error, and headless "
                f"it surfaces as a hang with no output."
            )
    return findings


FUNC_DECL_RE = re.compile(r"^(?:Public |Private |Friend )?Function\s+(\w+)", re.IGNORECASE)


def check_function_returns(path: Path, lines: list[str]) -> list[str]:
    """A Function that never assigns to its own name is almost always a bad rename.

    VBA returns a value by assigning to the function's name, so renaming the
    declaration and missing the assignment leaves `OldName = result` referring to
    nothing. Under Option Explicit that is "Variable not defined" -- another
    compile error, so another headless hang with no output.

    Cost 8 minutes on 2026-07-31, immediately after the reserved-word one cost 8
    minutes, both while adding the same feature. Two consecutive runs lost to
    compile errors that a text scan could have found in under a second.

    Reports only functions with a body -- a declaration-only stub legitimately
    assigns nothing.
    """
    findings: list[str] = []
    current: tuple[int, str] | None = None
    assigned = False
    body = 0
    for i, raw in enumerate(lines, 1):
        s = raw.strip()
        m = FUNC_DECL_RE.match(s)
        if m:
            current, assigned, body = (i, m.group(1)), False, 0
            continue
        if END_PROC_RE.match(s):
            if current and body > 0 and not assigned:
                findings.append(
                    f"{path}:{current[0]}: Function '{current[1]}' never assigns to its own "
                    f"name -- it returns nothing, and a leftover `OldName = ...` from a rename "
                    f"is 'Variable not defined' at compile time (a headless hang)."
                )
            current = None
            continue
        if current:
            if s and not s.startswith("'"):
                body += 1
            # `Set Name = obj` is the object-returning form and is just as valid
            # an assignment. Omitting it reported 37 working functions as broken
            # on the first attempt -- the second time in one session that this
            # checker was shipped without being run against the real corpus.
            # Searched ANYWHERE in the line, not anchored at its start: VBA
            # allows the assignment after a `Case x:` label or a single-line
            # `If ... Then`, and anchoring reported 11 more working functions as
            # broken. `=` is also VBA's comparison operator, so a line like
            # `If Foo = "x" Then` counts as an assignment here and hides a real
            # defect -- accepted deliberately, because a missed defect costs one
            # compile error while a false alarm costs the checker its credibility.
            if re.search(rf"\b{re.escape(current[1])}\s*=", s, re.IGNORECASE):
                assigned = True
    return findings


def check_structural_sanity(path: Path, lines: list[str]) -> list[str]:
    """Every module has at least one procedure, and starts/ends balance.

    Added after this checker reported a CORRUPTED file as clean. A bad edit had
    split the word `Private` across an insertion point, leaving `rivate
    Function ...` — so no line matched PROC_RE, so `check_declaration_order`
    found no "first procedure" and could not fire, and the file passed.

    A checker that cannot see the thing it checks reports success. That is the
    same failure as an always-true guard, and it is worse here because the
    output looks like evidence. So: assert the file has the shape the other
    checks assume, before trusting what they say about it.
    """
    findings: list[str] = []
    starts = sum(1 for l in lines if PROC_RE.match(l))
    ends = sum(1 for l in lines if END_PROC_RE.match(l.strip()))

    if starts == 0:
        findings.append(
            f"{path}: no procedures found at all. Either the module is empty, or a "
            f"procedure keyword is mangled (a bad edit splitting `Private`/`Public` "
            f"will do this). The other checks cannot be trusted on this file."
        )
    elif starts != ends:
        findings.append(
            f"{path}: {starts} procedure start(s) but {ends} End Sub/Function — "
            f"unbalanced, so a procedure is unterminated or a keyword is mangled."
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
        raw_lines = path.read_text(errors="replace").splitlines()
        # Comments stripped once, here, rather than in each check -- indices are
        # preserved so reported line numbers still point at the real file.
        lines = [code_only(l) for l in raw_lines]
        rel = path.relative_to(root.parent)
        findings += check_declaration_order(rel, lines)
        findings += check_duplicate_dims(rel, lines)
        findings += check_empty_redim(rel, lines)
        findings += check_reserved_names(rel, lines)
        findings += check_function_returns(rel, lines)
        findings += check_structural_sanity(rel, lines)

    if findings:
        print(f"=== {len(findings)} static finding(s) ===")
        for f in findings:
            print(f"  {f}")
        return 1

    print(f"static checks clean across {len(files)} module(s)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
