#!/usr/bin/env python3
"""Set a deck's period reliably, on top of a write that is not reliable.

WHY THIS EXISTS

Writing a PowerPoint CustomDocumentProperty does not dependably reach disk.
Measured 2026-08-01 on the ~49MB e2e deck, Office fully closed between runs and
every result judged by reading the .pptx's own bytes:

    Save          lost      (and 1 success in 5 in earlier testing)
    SaveAs        4 of 5    -- markedly better, still not reliable
    hidden window lost
    double Save   lost

The working theory, and it is only that: on a large package `Save` performs an
incremental rewrite that does not always regenerate `docProps/custom.xml`, while
`SaveAs` forces a full rewrite. That explains the direction but not why SaveAs
still misses one in five.

SO THIS DOES NOT TRY TO MAKE THE WRITE RELIABLE. It accepts an unreliable
primitive and builds a reliable operation out of it: write, verify against the
file itself, retry, and fail loudly rather than return a wrong answer.

The verification is the load-bearing part and it must stay OUT of process. An
in-process reopen shares PowerPoint's cache with the writer and cannot be
trusted in either direction -- it can pass a failed save and fail a good one,
which makes it worthless precisely when it matters. Reading the zip shares no
process, no cache and no code with the writer.

    python3 set_deck_period.py <deck.pptx> <period> [--attempts N]

Exit 0 only when the value is confirmed on disk. Exit 1 otherwise, having said
so plainly -- never exit 0 on an unverified write.
"""

from __future__ import annotations

import argparse
import os
import subprocess
import sys
import time
from pathlib import Path

HERE = Path(__file__).resolve().parent
REPO = HERE.parent.parent
DRIVER = HERE / "field_e2e.ps1"
WINTMP = Path("/mnt/c/Users/rohan/AppData/Local/Temp")


def win_path(p: Path) -> str:
    return subprocess.run(["wslpath", "-w", str(p)], capture_output=True, text=True,
                          check=True).stdout.strip()


def read_period(deck: Path) -> str | None:
    r = subprocess.run([sys.executable, str(HERE / "read_deck_props.py"), str(deck),
                        "DeckSyncPeriod"], capture_output=True, text=True)
    return r.stdout.strip() if r.returncode == 0 else None


def close_office() -> None:
    # A wedged Office instance has produced false results here more than once,
    # so every attempt starts from a known state rather than assuming one.
    subprocess.run(["powershell.exe", "-NoProfile", "-Command",
                    "Get-Process POWERPNT,EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force"],
                   capture_output=True)
    time.sleep(3)


def attempt_write(deck: Path, period: str) -> None:
    staged = WINTMP / DRIVER.name
    staged.write_bytes(DRIVER.read_bytes())
    subprocess.run(["powershell.exe", "-NoProfile", "-File", win_path(staged),
                    "-RepoRoot", win_path(REPO), "-Mode", "setperiodvariant",
                    "-Period", period, "-Variant", "saveas", "-DeckPath", str(deck).replace("/mnt/c", "C:").replace("/", "\\")],
                   capture_output=True)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("deck")
    ap.add_argument("period")
    ap.add_argument("--attempts", type=int, default=4)
    a = ap.parse_args()

    deck = Path(a.deck)
    if not deck.exists():
        print(f"no such deck: {deck}", file=sys.stderr)
        return 1

    # Guard, not decoration. The real deck is never a legitimate target for a
    # tool that writes in a loop and is known to fail intermittently.
    if "OneDrive" in str(deck):
        print("REFUSED: this writes repeatedly to a deck known to lose writes.\n"
              "Work on a copy, never on the original.", file=sys.stderr)
        return 1

    before = read_period(deck)
    print(f"deck:   {deck}")
    print(f"before: {before if before else '<none>'}")
    print(f"want:   {a.period}")

    for n in range(1, a.attempts + 1):
        close_office()
        attempt_write(deck, a.period)
        got = read_period(deck)
        if got == a.period:
            print(f"attempt {n}: CONFIRMED ON DISK")
            return 0
        print(f"attempt {n}: not on disk yet (reads '{got}') -- retrying")

    final = read_period(deck)
    print(f"\n*** FAILED after {a.attempts} attempts ***")
    print(f"    wanted:  {a.period}")
    print(f"    on disk: {final if final else '<none>'}")
    print("    The deck's period is NOT what you asked for. Do not trust any sync")
    print("    run against this deck until it is.")
    return 1


if __name__ == "__main__":
    sys.exit(main())
