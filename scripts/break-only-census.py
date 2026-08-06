#!/usr/bin/env python3
"""Census of BREAK-ONLY paragraphs across the fixture corpus, split by sectPr.

A break-only paragraph is one whose entire run content is one or more hard page
breaks: empty w:t is allowed, and anything else (text, tabs, drawings, fields,
note references, line/column breaks) disqualifies it. This mirrors
isPageBreakOnlyParagraph in the engine, so the two agree on what they count.

The split that matters is whether such a paragraph ALSO carries a w:sectPr in
its pPr, because Word treats the two shapes differently. Word keeps a plain
break-only paragraph on the current page however little room is left, but it
applies the ordinary fit test to one that ends a section -- and when that one
spills, its own page plus the section's next-page start leave a page with no
body ink at all (wild2-legal-ca-agreement page 2).

Usage: python3 scripts/break-only-census.py [name-or-path ...]
       python3 scripts/break-only-census.py          # the whole fixture corpus
"""
import glob
import os
import re
import sys
import zipfile

FIXTURES = os.path.join(os.path.dirname(__file__), "..", "apps", "demo", "public", "fixtures")

PARA_OPEN = re.compile(r"<w:p(?:\s[^>]*)?>")
PARA_CLOSE = re.compile(r"</w:p>")
RUN = re.compile(r"<w:r(?:\s[^>]*)?>(.*?)</w:r>", re.S)
RPR = re.compile(r"<w:rPr>.*?</w:rPr>", re.S)
ELEMENT = re.compile(r"<(w:[A-Za-z]+)(\s[^>]*?)?(/?)>")


def paragraphs(document):
    """(start, end) of every w:p, closing tag included.

    The scan takes the first </w:p> that ends STRICTLY after the opening tag:
    the previous paragraph's close can sit at exactly the opening offset, and
    accepting it collapses the slice to nothing. A self-closing <w:p .../> is
    its own whole paragraph and has no closing tag to look for.
    """
    closes = [m.end() for m in PARA_CLOSE.finditer(document)]
    spans = []
    for opening in PARA_OPEN.finditer(document):
        if opening.group(0).endswith("/>"):
            spans.append((opening.start(), opening.end()))
            continue
        after = [c for c in closes if c > opening.start()]
        if after:
            spans.append((opening.start(), min(after)))
    return spans


def is_break_only(block):
    """Does this paragraph's run content consist only of hard page breaks?"""
    saw_break = False
    for run in RUN.findall(block):
        for element in ELEMENT.finditer(RPR.sub("", run)):
            name, attrs = element.group(1), element.group(2)
            if name == "w:t":
                # Only an EMPTY w:t is allowed; a self-closing one is empty.
                body = re.search(r"<w:t(?:\s[^>]*)?>(.*?)</w:t>", run, re.S)
                if body and body.group(1):
                    return False
            elif name == "w:br":
                if 'w:type="page"' not in (attrs or ""):
                    return False  # line and column breaks disqualify
                saw_break = True
            else:
                return False
    return saw_break


def census(path):
    name = os.path.splitext(os.path.basename(path))[0]
    try:
        document = zipfile.ZipFile(path).read("word/document.xml").decode("utf-8")
    except Exception as error:  # unreadable package: report, do not abort the sweep
        print(f"{name:36s} UNREADABLE: {error}")
        return 0, 0
    found = []
    for index, (start, end) in enumerate(paragraphs(document)):
        block = document[start:end]
        if is_break_only(block):
            found.append((index, "<w:sectPr" in block))
    sectioned = sum(1 for _, s in found if s)
    if found:
        detail = ", ".join(f"p{i}{'+SECTPR' if s else ''}" for i, s in found[:10])
        if len(found) > 10:
            detail += f", +{len(found) - 10} more"
        print(f"{name:36s} break-only={len(found):3d}  with sectPr={sectioned:3d}  {detail}")
    return len(found), sectioned


def main(argv):
    targets = []
    for arg in argv:
        targets.append(arg if os.path.exists(arg) else os.path.join(FIXTURES, f"{arg}.docx"))
    if not targets:
        targets = sorted(glob.glob(os.path.join(FIXTURES, "*.docx")))
    total = sectioned = 0
    for path in targets:
        t, s = census(path)
        total += t
        sectioned += s
    print(f"\n{len(targets)} fixtures: {total} break-only paragraphs, {sectioned} of them carrying a sectPr")


if __name__ == "__main__":
    main(sys.argv[1:])
