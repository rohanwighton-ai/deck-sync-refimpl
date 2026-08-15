#!/usr/bin/env python3
"""Invisible Failure: an error suppressed and then never tested for, by ANY means.

The naive scan -- "On Error Resume Next with no later Err.Number" -- reports ~68
sites and cannot tell a swallowed error from a correctly handled one, because
plenty of VBA suppresses the error and then tests the RESULT instead:

    On Error Resume Next
    Set ws = wb.Worksheets(name)
    On Error GoTo 0
    If ws Is Nothing Then ...          <- handled, just not via Err

A correctly handled probe and a swallowed failure produce identical output from
that scan, so it has no criterion and its findings are candidates, not defects.

The criterion here: between `On Error Resume Next` and the point error handling
is restored (or the procedure ends), plus a short tail afterwards, there must be
SOME test of whether the thing worked -- either the error object, or the result.
If there is neither, nothing downstream can know it failed.
"""
import re, sys, glob, os

# Evidence that failure was considered. Deliberately generous: this check should
# UNDER-report rather than accuse working code. A false accusation costs more
# than a miss here, because it trains people to dismiss the whole check.
HANDLED = re.compile(
    r"Err\.Number|Err\.Description|Err\.Raise"      # asked the error object
    r"|Is Nothing|Not\s+\w+\s+Is\s+Nothing"          # tested the object
    r'|=\s*""|<>\s*""'                               # tested for empty
    r"|\.Count|\.Exists|FileExists|FolderExists"     # tested presence
    r"|=\s*0\b|>\s*0\b|<>\s*0\b"                     # tested a count/flag
    r"|IsEmpty|IsNull|IsError|IsArray|IsObject"
    r"|problem\s*<>|problem\s*=|outcome\s*<>"        # this codebase's own idiom
    , re.IGNORECASE)

RESTORE = re.compile(r"On Error GoTo\b", re.IGNORECASE)
SUPPRESS = re.compile(r"On Error Resume Next", re.IGNORECASE)
PROC_END = re.compile(r"^\s*End (Sub|Function)\b", re.IGNORECASE)
PROC_START = re.compile(r"^\s*(Public|Private|Friend)?\s*(Sub|Function)\s+(\w+)", re.IGNORECASE)

TAIL = 6   # lines to keep looking after handling is restored


def scan(path):
    lines = open(path, encoding="utf-8", errors="replace").read().splitlines()
    findings = []
    proc = "(module level)"
    for i, ln in enumerate(lines):
        m = PROC_START.match(ln)
        if m:
            proc = m.group(3)
        if not SUPPRESS.search(ln):
            continue

        # SCOPE IS THE REST OF THE PROCEDURE, and that is the correction that
        # matters. The first version looked only inside the suppressed region
        # plus six lines, and reported 40 sites of which the two sampled were
        # both handled -- by an explicit POSTCONDITION further down. This
        # codebase's idiom is "suppress the call, then assert what should now be
        # true", which is stronger than checking Err.Number because it tests the
        # world rather than the error object. A check that does not understand
        # the idiom of the code it audits reports style, not defects.
        region, j = [], i + 1
        while j < len(lines):
            if PROC_END.match(lines[j]):
                break
            region.append(lines[j])
            j += 1

        body = "\n".join(region)
        # Strip comments: a comment SAYING "Err" is not a test of it.
        body = "\n".join(re.sub(r"'.*$", "", r_) for r_ in body.splitlines())

        if not HANDLED.search(body):
            findings.append((i + 1, proc, lines[i].strip()))
    return findings


def main():
    root = sys.argv[1] if len(sys.argv) > 1 else "vba"
    total = 0
    for path in sorted(glob.glob(os.path.join(root, "*.bas"))):
        f = scan(path)
        if not f:
            continue
        print(f"\n=== {os.path.basename(path)} : {len(f)} ===")
        for line, proc, _ in f:
            print(f"  line {line:>5}  in {proc}")
        total += len(f)
    print(f"\nTOTAL suppressed-and-never-tested: {total}")


if __name__ == "__main__":
    main()
