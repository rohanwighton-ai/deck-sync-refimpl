#!/usr/bin/env python3
"""Document control: fail when a .md states a fact the code contradicts.

The standing rule is that a machine-knowable fact must be DERIVED, never written
into prose, because a sentence cannot fail a test and so it drifts silently. This
script is the test the sentences never had.

Every check below exists because that exact drift was found in this repo on
2026-08-14, in docs that had been read in full the same morning:

  - WORKFLOW.md gave drafting columns as G and I -- layout 3, two layouts stale,
    pointing a person at the SOURCES and AI DRAFT columns
  - TOOLBAR.md described three toolbar buttons; there are two
  - FIX-LIST.md said the register workbook "the tool rebuilds and clears" -- it
    does not; Register, Field Spec and Sources are PERMANENT
  - HANDOVER said three projects' text was "not recoverable from any backup";
    it was on the slides all along
  - docs cited vba/tools/check_vba_static.py, which lives in vba/tests/

Run before committing docs, and as part of any review pass.
Exit 0 clean, 1 with findings, 2 if the checker could not run.
"""
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ROOT = os.path.dirname(ROOT) if os.path.basename(ROOT) == 'vba' else ROOT
VBA = os.path.join(ROOT, 'vba')

SKIP_DIRS = {'.git', '__pycache__', '.pytest_cache', 'node_modules'}
# Historical by design: exchange rounds and dated snapshots are records of what
# was believed at the time and must NOT be rewritten to match today's code.
HISTORICAL = re.compile(
    r'(^archive/)|(^specs/)|(NEXT-SESSION-\d{4}-\d\d-\d\d\.md$)|(CYCLE-FINDINGS-)|'
    r'(FIRST-REAL-RUN\.md$)|(IMPLEMENTATION_PLAN\.md$)|(SPIKE_NOTES)|'
    r'(test-fixtures/)|(PROMPT_)|(AGENTS\.md$)')
# ^archive/ added 2026-08-16 when specs/ and vba/SPIKE_NOTES_*.md moved under
# archive/ during the documentation sweep -- ^specs/ alone stopped matching the
# moment the directory changed. Everything under archive/ is historical by
# construction now, so this also makes future archival additions exempt
# automatically instead of needing a new pattern each time. The other patterns
# above are unanchored (match by filename, not directory) and kept working
# through the move without needing this fix.


def read(p):
    with open(p, encoding='utf8', errors='replace') as f:
        return f.read()


def vba_source():
    out = {}
    for dirpath, dirnames, files in os.walk(VBA):
        dirnames[:] = [d for d in dirnames if d not in SKIP_DIRS]
        for f in files:
            if f.endswith('.bas'):
                out[f[:-4]] = read(os.path.join(dirpath, f))
    return out


def docs():
    out = []
    for dirpath, dirnames, files in os.walk(ROOT):
        dirnames[:] = [d for d in dirnames if d not in SKIP_DIRS]
        for f in files:
            if not f.endswith('.md'):
                continue
            full = os.path.join(dirpath, f)
            rel = os.path.relpath(full, ROOT).replace(os.sep, '/')
            if HISTORICAL.search(rel):
                continue
            out.append((rel, read(full)))
    return sorted(out)


def strip_quoted_history(text):
    """Blockquoted lines are corrections and archived banners -- a doc is allowed
    to QUOTE a wrong string in the act of correcting it. Checking them would make
    every fix look like a fresh defect."""
    return '\n'.join(l for l in text.split('\n') if not l.lstrip().startswith('>'))


def main():
    src = vba_source()
    if not src:
        print('CANNOT RUN: no .bas files found under', VBA)
        return 2
    all_src = '\n'.join(src.values())
    findings = []

    # ---- 1. toolbar captions -------------------------------------------------
    live = set(re.findall(r'Public Const CAP_[A-Z_]+ As String = "([^"]+)"',
                          src.get('CommandBarUI', '')))
    # Caption-shaped strings this project has used: "1. Sync Now", "Setup B: ..."
    caption_shape = re.compile(r'`(\d[a-z]?\.\s+[A-Z][^`]{2,30}|Setup [A-Z]\d?:[^`]{2,30})`')

    # ---- 2. drafting column letters that CONTRADICT the constants -------------
    #
    # Deliberately not "any letter in prose". A first version flagged every
    # mention and produced 77 findings including correct ones -- and this repo
    # has already paid for that lesson: a checker that cries wolf is worse than
    # no checker (AGENTS.md, 2026-07-31, two checks that reported 27 and 37 false
    # positives the moment they met the real corpus).
    #
    # So: only where a letter is named NEAR the role it claims to hold, and the
    # letter disagrees with the constant. That is the actual defect -- WORKFLOW.md
    # telling a person to tick column I, which is the character-count column.
    cols = {}
    for m in re.finditer(r'Public Const (COL_D_[A-Z]+) As Long = (\d+)',
                         src.get('Drafting', '')):
        cols[m.group(1)] = chr(64 + int(m.group(2)))
    # Most specific token first: "sources" must win over "draft" for a phrase
    # like "drafting sheet column D (sources)", which is correct and was being
    # reported as a defect.
    ROLE = [
        (r'sources?\b', 'COL_D_SOURCES'),
        (r'ai draft|copilot writes', 'COL_D_DRAFT'),
        (r'tick|approv', 'COL_D_APPROVED'),
        (r'submit|what gets sent|this publishes', 'COL_D_SUBMIT'),
        # 'original'/'what the slide says' removed 2026-08-20 along with
        # COL_D_CURRENT itself -- the drafting sheet no longer has an
        # ORIGINAL column (see Drafting.bas's DRAFT_LAYOUT_VERSION header).
        (r'reported last time|what was reported', 'COL_D_PREV'),
    ]
    col_claim = re.compile(r'column\s+([A-L])\b', re.I)

    # ---- 3. paths and VBA symbols cited in docs ------------------------------
    path_ref = re.compile(r'`((?:vba|src|specs|test-fixtures)/[\w./-]+\.(?:py|bas|ps1|md))`')
    sym_ref = re.compile(r'`([A-Z][A-Za-z]+)\.([A-Za-z_][A-Za-z0-9_]*)`')

    for rel, raw in docs():
        head = '\n'.join(raw.split('\n')[:40])

        # EVERY DOC MUST DECLARE ITS STATUS. That is the whole of document
        # control: a reader must never have to guess whether a page describes
        # the tool as it is or as it was.
        if not re.search(r'(stale|historical|superseded|status:|written|current)',
                         head, re.I):
            findings.append((rel, 'NO STATUS MARKER',
                             'first 40 lines say nothing about how current this is'))
            continue

        # A doc that SAYS it is historical is allowed to describe the old world
        # -- that is what it is for. Rewriting it to match today's code would
        # destroy the record. It is exempt from the content checks below, and
        # the banner is the thing holding it honest.
        if re.search(r'(historical|superseded|stale)', head, re.I):
            continue

        text = strip_quoted_history(raw)

        # A doc naming a removed button AS REMOVED is correct documentation --
        # that is what a fix list and a correction note are for. The defect this
        # catches is a caption used as though a person could press it today.
        dead_marker = re.compile(
            r'(no longer|removed|dead|deleted|do(?:es)? not exist|did not exist|'
            r'superseded|was written|absorbed|stale|old toolbar|reported)', re.I)

        for m in caption_shape.finditer(text):
            cap = m.group(1).strip()
            around = text[max(0, m.start() - 160): m.end() + 160]
            if dead_marker.search(around):
                continue
            if cap not in live and not any(cap in l for l in live):
                findings.append((rel, 'DEAD BUTTON CAPTION',
                                 '"%s" is not a live toolbar caption (live: %s)'
                                 % (cap, ', '.join(sorted(live)))))

        for m in col_claim.finditer(text):
            letter = m.group(1).upper()
            window = text[max(0, m.start() - 30): m.end() + 30].lower()
            # AMBIGUOUS WINDOW -- more than one column named, so the role word
            # cannot be attributed to this letter. "column C readable, F/G
            # obvious" is a correct sentence and was reported as a defect.
            if len(set(re.findall(r'\b([a-l])\b(?=[/,\s])|column\s+([a-l])\b',
                                  window))) > 1:
                continue
            if len(re.findall(r'column\s+[a-l]\b', window)) > 1:
                continue
            for pattern, const in ROLE:
                if not re.search(pattern, window):
                    continue
                right = cols.get(const)
                if right and letter != right:
                    findings.append((rel, 'WRONG COLUMN LETTER',
                                     '"column %s" described as %s -- %s is column %s'
                                     % (letter, const, const, right)))
                break

        for m in path_ref.finditer(text):
            if not os.path.exists(os.path.join(ROOT, m.group(1))):
                findings.append((rel, 'PATH DOES NOT EXIST', m.group(1)))

        for m in sym_ref.finditer(text):
            mod, sym = m.group(1), m.group(2)
            if mod in src and not re.search(r'\b%s\b' % re.escape(sym), src[mod]):
                if not re.search(r'\b%s\b' % re.escape(sym), all_src):
                    findings.append((rel, 'SYMBOL NOT IN SOURCE',
                                     '%s.%s' % (mod, sym)))

    if not findings:
        print('document control clean across %d live doc(s)' % len(docs()))
        return 0

    by_doc = {}
    for rel, kind, detail in findings:
        by_doc.setdefault(rel, []).append((kind, detail))
    for rel in sorted(by_doc):
        print('\n%s' % rel)
        for kind, detail in by_doc[rel]:
            print('   %-24s %s' % (kind, detail))
    print('\n%d finding(s) across %d document(s)' % (len(findings), len(by_doc)))
    return 1


if __name__ == '__main__':
    sys.exit(main())
