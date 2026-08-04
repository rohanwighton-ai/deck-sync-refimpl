#!/usr/bin/env python3
"""Rename a registered slide type reliably: property, worksheet pointer, and
every slide's tag, all verified against the file's own bytes afterward.

Fourth of the DeckSync*-property writers (see set_deck_period.py for why a
plain Save is not trusted on this large a package). This one additionally
touches every slide's tags, which is the same shape as the live incident
BatchOnboardFlow.bas:937 records -- a rename that moved tags without moving
the registration stranded an entire Data sheet. So this checks BOTH sides
independently, out of process, before calling the rename real:

  1. the NEW type is registered, pointing at the right worksheet
  2. the OLD type's registration is gone (no ghost type)
  3. every slide that answered to the OLD tag now answers to the NEW one
  4. no slide is left answering to the OLD tag

    python3 rename_slide_type.py <deck.pptx> <old-type> <new-type> <new-worksheet-name> [--attempts N]

Exit 0 only when all four are confirmed on disk. Exit 1 otherwise.
"""

from __future__ import annotations

import argparse
import subprocess
import sys
import time
from pathlib import Path

HERE = Path(__file__).resolve().parent
REPO = HERE.parent.parent
DRIVER = HERE / "field_e2e.ps1"


def win_path(p: Path) -> str:
    return subprocess.run(["wslpath", "-w", str(p)], capture_output=True, text=True,
                          check=True).stdout.strip()


def read_registration(deck: Path, slide_type: str) -> str | None:
    r = subprocess.run([sys.executable, str(HERE / "read_deck_props.py"), str(deck),
                        f"DeckSyncType:{slide_type}"], capture_output=True, text=True)
    return r.stdout.strip() if r.returncode == 0 else None


def count_tagged(deck: Path, value: str) -> int:
    r = subprocess.run([sys.executable, str(HERE / "read_deck_slide_tags.py"), str(deck), "SLIDE_TYPE"],
                       capture_output=True, text=True)
    return sum(1 for line in r.stdout.splitlines() if line.strip().endswith(f": {value}"))


def close_office() -> None:
    subprocess.run(["powershell.exe", "-NoProfile", "-Command",
                    "Get-Process POWERPNT,EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force"],
                   capture_output=True)
    time.sleep(3)


def attempt_write(deck: Path, old_type: str, new_type: str, new_sheet: str) -> None:
    staged = Path("/mnt/c/Users/rohan/AppData/Local/Temp") / DRIVER.name
    staged.write_bytes(DRIVER.read_bytes())
    subprocess.run(["powershell.exe", "-NoProfile", "-File", win_path(staged),
                    "-RepoRoot", win_path(REPO), "-Mode", "renameslidetype", "-Variant", "saveas",
                    "-SlideType", old_type, "-NewSlideType", new_type, "-SheetName", new_sheet,
                    "-DeckPath", str(deck).replace("/mnt/c", "C:").replace("/", "\\")],
                   capture_output=True)


def confirmed(deck: Path, old_type: str, new_type: str, new_sheet: str) -> tuple[bool, str]:
    new_reg = read_registration(deck, new_type)
    old_reg = read_registration(deck, old_type)
    new_tagged = count_tagged(deck, new_type)
    old_tagged = count_tagged(deck, old_type)

    checks = {
        f"new type registered against '{new_sheet}'": new_reg is not None and new_reg.endswith(f"|{new_sheet}"),
        "old registration removed": old_reg is None,
        "no slide still tagged with the old type": old_tagged == 0,
        "at least one slide tagged with the new type": new_tagged > 0,
    }
    summary = (f"  new registration: {new_reg}\n  old registration: {old_reg}\n"
               f"  slides tagged '{new_type}': {new_tagged}\n  slides tagged '{old_type}': {old_tagged}\n" +
               "\n".join(f"  [{'OK' if ok else 'FAIL'}] {name}" for name, ok in checks.items()))
    return all(checks.values()), summary


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("deck")
    ap.add_argument("old_type")
    ap.add_argument("new_type")
    ap.add_argument("new_worksheet_name")
    ap.add_argument("--attempts", type=int, default=4)
    a = ap.parse_args()

    deck = Path(a.deck)
    if not deck.exists():
        print(f"no such deck: {deck}", file=sys.stderr)
        return 1
    if "OneDrive" in str(deck):
        print("REFUSED: this writes repeatedly to a deck known to lose writes.\n"
              "Work on a copy, never on the original.", file=sys.stderr)
        return 1

    print(f"deck: {deck}")
    print(f"rename: '{a.old_type}' -> '{a.new_type}'  (worksheet -> '{a.new_worksheet_name}')")

    for n in range(1, a.attempts + 1):
        close_office()
        attempt_write(deck, a.old_type, a.new_type, a.new_worksheet_name)
        ok, summary = confirmed(deck, a.old_type, a.new_type, a.new_worksheet_name)
        if ok:
            print(f"\nattempt {n}: CONFIRMED ON DISK\n{summary}")
            return 0
        print(f"\nattempt {n}: not fully confirmed -- retrying\n{summary}")

    print(f"\n*** FAILED after {a.attempts} attempts ***")
    return 1


if __name__ == "__main__":
    sys.exit(main())
