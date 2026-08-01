#!/usr/bin/env python3
"""Do the three hand-maintained module lists cover what the code references?

Three separate lists name which .bas files get imported:
  vba/tests/run_vba_tests.ps1   the test suite
  vba/tests/build_ppam.ps1      the shipped add-in
  vba/tools/field_e2e.ps1       the driver harness

Nothing checked them against the source. On 2026-08-01 two were stale at once:
build_ppam.ps1 had never included ReviewQueue.bas while RibbonUI.SyncNow called
it in nine places (so the built add-in could not compile, silently breaking
Sync Now / Review Changes / Apply Approved), and run_vba_tests.ps1 was missing
Sources.bas (so the "135 tests pass" claim had not been true all day).

An undefined module reference is a COMPILE error in VBA, and a compile error
takes out the entire project -- so this fails as everything-at-once, which is
exactly the failure mode that gets attributed to the wrong cause.

Run it before trusting any of those three.
"""
import re, sys, pathlib

ROOT = pathlib.Path(__file__).resolve().parents[2]
VBA = ROOT / "vba"

LISTS = {
    "test suite  (run_vba_tests.ps1)": VBA / "tests" / "run_vba_tests.ps1",
    "add-in      (build_ppam.ps1)":    VBA / "tests" / "build_ppam.ps1",
    "harness     (field_e2e.ps1)":     VBA / "tools" / "field_e2e.ps1",
}

# every module that exists, by its VB_Name
modules = {}
for f in list(VBA.glob("*.bas")) + list((VBA / "tools").glob("*.bas")) + list((VBA / "tests").glob("*.bas")):
    m = re.search(r'Attribute VB_Name = "([^"]+)"', f.read_text(encoding="utf-8", errors="replace"))
    if m:
        modules[m.group(1)] = f

failed = False
for label, script in LISTS.items():
    text = script.read_text(encoding="utf-8", errors="replace")
    imported = set(re.findall(r'"([A-Za-z0-9_\\]+\.bas)"', text))
    imported = {pathlib.Path(i.replace("\\", "/")).stem for i in imported}

    # what do the imported modules actually reference?
    referenced = set()
    for name in imported:
        if name not in modules:
            continue
        src = modules[name].read_text(encoding="utf-8", errors="replace")
        src = "\n".join(l for l in src.splitlines() if not l.strip().startswith("'"))
        for other in modules:
            if other != name and re.search(r'\b' + re.escape(other) + r'\s*\.', src):
                referenced.add(other)

    missing = sorted(referenced - imported)
    if missing:
        failed = True
        print(f"FAIL  {label}")
        for m in missing:
            users = [n for n in imported if n in modules and
                     re.search(r'\b' + re.escape(m) + r'\s*\.',
                               modules[n].read_text(encoding='utf-8', errors='replace'))]
            print(f"        missing {m}.bas  -- referenced by {', '.join(sorted(users)[:4])}")
    else:
        print(f"PASS  {label}  ({len(imported)} modules, all references satisfied)")

sys.exit(1 if failed else 0)
