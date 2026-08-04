#!/usr/bin/env python3
"""Re-point a deck's paired workbook reliably, on top of a write that is not.

Same reasoning as set_deck_period.py, applied to DeckSyncWorkbookPath instead
of DeckSyncPeriod -- both are CustomDocumentProperty writes on the same large
package, and a plain Save has been measured to lose them. See that file's
header for the numbers. Verification is out-of-process for the same reason:
an in-process reopen shares PowerPoint's cache with the writer.

    python3 set_deck_workbook_path.py <deck.pptx> <new-workbook-path> [--attempts N]

Exit 0 only when the value is confirmed on disk. Exit 1 otherwise.
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


def read_workbook_path(deck: Path) -> str | None:
    r = subprocess.run([sys.executable, str(HERE / "read_deck_props.py"), str(deck),
                        "DeckSyncWorkbookPath"], capture_output=True, text=True)
    return r.stdout.strip() if r.returncode == 0 else None


def close_office() -> None:
    subprocess.run(["powershell.exe", "-NoProfile", "-Command",
                    "Get-Process POWERPNT,EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force"],
                   capture_output=True)
    time.sleep(3)


def attempt_write(deck: Path, new_workbook: str) -> None:
    staged = Path("/mnt/c/Users/rohan/AppData/Local/Temp") / DRIVER.name
    staged.write_bytes(DRIVER.read_bytes())
    subprocess.run(["powershell.exe", "-NoProfile", "-File", win_path(staged),
                    "-RepoRoot", win_path(REPO), "-Mode", "repointworkbook",
                    "-Variant", "saveas", "-NewWorkbookPath", new_workbook,
                    "-DeckPath", str(deck).replace("/mnt/c", "C:").replace("/", "\\")],
                   capture_output=True)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("deck")
    ap.add_argument("new_workbook_path", help="Windows path, e.g. C:\\Users\\rohan\\deck-sync-e2e\\register-wide.xlsx")
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

    before = read_workbook_path(deck)
    print(f"deck:   {deck}")
    print(f"before: {before if before else '<none>'}")
    print(f"want:   {a.new_workbook_path}")

    for n in range(1, a.attempts + 1):
        close_office()
        attempt_write(deck, a.new_workbook_path)
        got = read_workbook_path(deck)
        if got == a.new_workbook_path:
            print(f"attempt {n}: CONFIRMED ON DISK")
            return 0
        print(f"attempt {n}: not on disk yet (reads '{got}') -- retrying")

    final = read_workbook_path(deck)
    print(f"\n*** FAILED after {a.attempts} attempts ***")
    print(f"    wanted:  {a.new_workbook_path}")
    print(f"    on disk: {final if final else '<none>'}")
    return 1


if __name__ == "__main__":
    sys.exit(main())
