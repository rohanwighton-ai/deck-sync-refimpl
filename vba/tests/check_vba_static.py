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
    # Built-in FUNCTIONS are just as illegal as identifiers, and were missing.
    # "fix" cost a compile-gate failure on 2026-08-09 -- Readiness.bas had a
    # ReadyLine member and a parameter both called Fix, which VBE reports only
    # as "Syntax error" on a line that looks fine. Added as a class rather than
    # the one word, since every name here fails the same way.
    "fix", "int", "abs", "sgn", "sqr", "log", "exp", "rnd", "val", "hex",
    "oct", "asc", "chr", "len", "mid", "left", "right", "trim", "ltrim",
    "rtrim", "space", "format", "join", "split", "filter", "replace",
    "array", "choose", "iif", "switch", "cint", "clng", "cdbl", "cstr",
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


PROC_PARAMS_RE = re.compile(
    r"^(?:Public |Private |Friend )?(?:Sub|Function|Property(?:\s+(?:Get|Let|Set))?)\s+\w+\s*\((.*)$",
    re.IGNORECASE)
PARAM_NAME_RE = re.compile(r"(?:ByVal\s+|ByRef\s+|Optional\s+|ParamArray\s+)*(\w+)", re.IGNORECASE)


def check_reserved_params(path: Path, lines: list[str]) -> list[str]:
    """A reserved word as a PARAMETER name is a Syntax error too.

    check_reserved_names only inspects declarations (`Dim x As ...`). It does
    not look inside a procedure's parameter list, so
    `Function Foo(variant As String)` sailed through -- and `variant` is a VBA
    data type, so that is a compile-time Syntax error. Cost a full agent run on
    2026-08-01, having already cost a run as `Dim empty` the day before: the
    same mistake in the one place the checker was not looking.

    Handles a continued signature (` _` line breaks) by joining forward.
    """
    findings: list[str] = []
    for i, raw in enumerate(lines, 1):
        m = PROC_PARAMS_RE.match(raw.strip())
        if not m:
            continue
        sig = m.group(1)
        j = i
        while sig.rstrip().endswith("_") and j < len(lines):
            sig = sig.rstrip()[:-1] + lines[j].strip()
            j += 1
        sig = sig.split(")")[0]
        for part in sig.split(","):
            pm = PARAM_NAME_RE.match(part.strip())
            if pm and pm.group(1).lower() in RESERVED_NAMES:
                findings.append(
                    f"{path}:{i}: parameter '{pm.group(1)}' is a VBA reserved word -- "
                    f"a compile-time Syntax error, and headless it is a hang with no output."
                )
    return findings


# A PROCEDURE THAT CALLS ITSELF WITH ITS OWN ARGUMENT LIST, VERBATIM.
#
# Twice in one session (2026-08-09) a scripted find-and-replace rewrote a line
# INSIDE the procedure it was creating:
#
#   RemedyText = RemedyText(RM_SAVE_DECK_THEN_REBUILD)   ' was a literal string
#   Say text, style, caption                             ' was MsgBox text, ...
#
# Both are infinite recursion. NEITHER was caught: VBA compiles them happily,
# and the second shipped in an add-in and surfaced as "Run-time error 28: Out of
# stack space" the first time it ran. 163 tests could not see it because the
# path only runs outside a chain, which nothing tests.
#
# Deliberately NARROW. Real recursion passes arguments that CHANGE -- that is
# what makes it terminate. A call passing exactly the parameter names in exactly
# the declared order can never terminate, so this cannot cry wolf on a genuine
# recursive function.
def check_self_call_with_own_params(path: Path, lines: list[str]) -> list[str]:
    problems: list[str] = []
    decl = re.compile(r"^\s*(?:Public|Private)?\s*(?:Sub|Function)\s+(\w+)\s*\((.*?)\)", re.I)
    name = None
    params: list[str] = []
    for i, raw in enumerate(lines, 1):
        line = code_only(raw)
        m = decl.match(line)
        if m:
            name = m.group(1)
            params = [
                p.strip().split()[-1] if " " in p.strip() else p.strip()
                for p in m.group(2).split(",")
                if p.strip()
            ]
            params = [
                re.sub(r"^(?:ByRef|ByVal|Optional)\s+", "", p, flags=re.I).split(" As ")[0].strip()
                for p in m.group(2).split(",")
                if p.strip()
            ]
            continue
        if not name or not params:
            continue
        if re.match(r"^\s*End (Sub|Function)", line, re.I):
            name, params = None, []
            continue
        body = line.strip()
        joined = r"\s*,\s*".join(re.escape(p) for p in params)
        # "Name a, b" or "Name = Name(a, b)"
        as_sub = re.fullmatch(rf"{re.escape(name)}\s+{joined}", body, re.I)
        as_fn = re.fullmatch(rf"{re.escape(name)}\s*=\s*{re.escape(name)}\s*\(\s*{joined}\s*\)", body, re.I)
        if as_sub or as_fn:
            problems.append(
                f"{path.as_posix()}:{i}: '{name}' calls itself passing its own parameters "
                f"unchanged -- infinite recursion. Almost always a find-and-replace that "
                f"rewrote a line inside the procedure it was creating."
            )
    return problems

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



# Modules whose Public procedures are USER-FACING CAPABILITIES -- things a person
# does, as opposed to plumbing another module calls. A capability nothing can
# reach is the failure this scan exists for.
#
# WHY: CreateMissingSlides was built, tested, and had no toolbar button, so on the
# work machine -- where the toolbar is the whole interface -- there was no way to
# add a slide for a new project. RollForwardPeriod was the same story a week
# earlier: StartQuarter told Rohan to hand-copy 43 rows in Excel because the
# function that did it could not be pressed.
#
# The suite could not catch either. TestRunner asserts every BUTTON resolves to a
# real Sub; nothing asserts the reverse. Tests call functions directly, a person
# presses buttons, and nothing covered the gap.
# AdoptFlow.bas ADDED 2026-08-14. Its omission is why this check reported clean
# while AdoptFlow.AdoptExistingSlides sat with no button and no caller for the
# whole life of the three-button toolbar -- the checker built to catch "built and
# unreachable by a person" could not see the module it was happening in. Proven
# by adding it and re-running before the fix: two notes appeared immediately.
UI_MODULES = {"RibbonUI.bas", "DraftingUI.bas", "DiscoverUI.bas", "BatchOnboardFlow.bas",
              "AdoptFlow.bas"}

# Reached some other way, on purpose. Each needs a reason, so the list cannot
# quietly become where inconvenient findings go to be silenced.
REACHABLE_OTHERWISE = {
    "ShowSyncResult": "shared reporting helper, called by every action",
    "CapReport": "shared truncation helper",
    "UnexpectedErrorText": "error text builder used by every wrapper",
    "ShowToolbar": "called by the add-in's Auto_Open",
    "HideToolbar": "called by ShowToolbar and Auto_Close",
    "ToolbarName": "read by tests and by ShowToolbar",
    "ToolbarNames": "read by tests and by HideToolbar",
    "ResolveRegisterSheet": "resolver used by publish and drafting",
    "DraftingPromptFor": "prompt builder used by the drafting sheet",
}


def check_unreachable_capabilities(files: list[Path], root: Path) -> list[str]:
    """Public procedures in UI modules that no toolbar button fires and no other
    module calls -- built, possibly tested, and unreachable by a person."""
    # THE WIRING IS READ FROM RAW TEXT, NOT code_only OUTPUT.
    #
    # code_only blanks string contents, and every button is wired with a STRING:
    #   AddButton bar, "4. Sync Now", "RibbonUI.SyncNow", 1004, "Use to ..."
    # so scanning the processed text found no actions at all and reported all
    # sixteen live buttons as unreachable. A check whose input cannot contain the
    # evidence is the same shape as verifying a save against the writer's cache --
    # it produces a confident answer about nothing.
    #
    # Parsed as the third argument of AddButton specifically, rather than by
    # grepping for the name anywhere: a mention in a comment must not count as
    # wired, or the scan goes quiet exactly where documentation is thickest.
    #
    # THE CAPTION IS NO LONGER A LITERAL. 2026-08-09: captions moved to
    # Public Const CAP_* in CommandBarUI so a rename is one edit rather than 88,
    # and this pattern -- which required a quoted second argument -- stopped
    # matching anything. It reported that rather than passing on zero buttons,
    # which is the only reason the change did not silently blind the check. The
    # second argument now accepts either form; the ACTION is still required to be
    # a literal, because that is the thing being verified.
    wired: set[str] = set()
    bodies: dict[str, str] = {}
    public_procs: list[tuple[str, str, int]] = []

    pub_re = re.compile(r"^Public\s+(?:Sub|Function)\s+(\w+)")
    addbtn_re = re.compile(r'AddButton\s+\w+\s*,\s*(?:"[^"]*"|\w+)\s*,\s*"([^"]+)"')

    for path in files:
        name = path.name
        raw = path.read_text(errors="replace")
        bodies[name] = "\n".join(code_only(l) for l in raw.splitlines())
        if name == "CommandBarUI.bas":
            for action in addbtn_re.findall(raw):
                wired.add(action.rsplit(".", 1)[-1])
        if name in UI_MODULES:
            for i, line in enumerate(bodies[name].splitlines(), 1):
                m = pub_re.match(line)
                if m:
                    public_procs.append((name, m.group(1), i))

    if not wired:
        return [
            "vba/tests/check_vba_static.py: parsed ZERO toolbar buttons from "
            "CommandBarUI.bas -- the AddButton pattern no longer matches, so this "
            "check would pass by seeing nothing. Fix the pattern before trusting it."
        ]

    findings = []
    for mod, proc, lineno in public_procs:
        if proc in REACHABLE_OTHERWISE or proc in wired:
            continue
        called_elsewhere = any(
            other != mod and re.search(rf"\b{re.escape(proc)}\b", text)
            for other, text in bodies.items()
        )
        if called_elsewhere:
            continue

        # Distinguished deliberately. "Unreachable" and "over-exposed" need
        # different fixes, and reporting the wrong one is how a checker stops
        # being read: a helper used ten lines below its own declaration is not a
        # missing button, it is a missing Private.
        own = bodies[mod]
        uses_in_own = len(re.findall(rf"\b{re.escape(proc)}\b", own))
        if uses_in_own > 1:
            findings.append(
                f"vba/{mod}:{lineno}: Public {proc} is used only inside its own module "
                f"-- make it Private, or wire it to a button if it is meant to be a "
                f"capability."
            )
        else:
            findings.append(
                f"vba/{mod}:{lineno}: Public {proc} is a UI capability with no toolbar "
                f"button and no caller anywhere -- built and unreachable by a person. "
                f"Wire a button, fold it into one that exists, or add it to "
                f"REACHABLE_OTHERWISE with the reason."
            )
    return findings


def main() -> int:
    root = Path(__file__).resolve().parent.parent
    # .cls ADDED 2026-08-16 for AppEvents.cls, the first class module this
    # project has ever had (LOBBY-DESIGN.md). The checks below are line-based
    # over ordinary VBA code and do not depend on the .bas-specific
    # `Attribute VB_Name` header position, so a class module's extra
    # VERSION/BEGIN...END preamble does not need special-casing.
    files = (sorted(root.glob("*.bas")) + sorted((root / "tests").glob("*.bas")) +
             sorted(root.glob("*.cls")))
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
        findings += check_reserved_params(rel, lines)
        findings += check_function_returns(rel, lines)
        findings += check_structural_sanity(rel, lines)
        findings += check_self_call_with_own_params(rel, lines)


    # REACHABILITY IS REPORTED, NOT ENFORCED, AND THE ASYMMETRY IS DELIBERATE.
    #
    # Everything above is a pattern-matchable certainty -- a reserved word as a
    # parameter IS a compile error. Whether a Public proc SHOULD be a capability
    # is a judgement, and this file's own docstring says a checker that reports
    # maybes trains you to ignore it. Blocking the suite on a judgement nobody
    # will resolve today produces a gate people learn to bypass, which costs more
    # than the finding is worth.
    #
    # The enforcing half lives in the suite instead:
    # TestRunner.Test_CommandBarUI_EveryDeclaredCapabilityHasAButton fails when a
    # capability we have DECLARED loses its button. Declared list blocks; scan
    # discovers. Neither alone would have caught CreateMissingSlides being built,
    # tested and unreachable.
    reach = check_unreachable_capabilities(files, root)

    exit_code = 0
    if findings:
        print(f"=== {len(findings)} static finding(s) ===")
        for f in findings:
            print(f"  {f}")
        exit_code = 1
    else:
        print(f"static checks clean across {len(files)} module(s)")

    if reach:
        print(f"--- {len(reach)} reachability note(s) (reported, not blocking) ---")
        for f in reach:
            print(f"  {f}")

    return exit_code


if __name__ == "__main__":
    sys.exit(main())
