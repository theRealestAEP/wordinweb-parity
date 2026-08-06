#!/usr/bin/env python3
"""Set w:fldLock on every DATE and TIME field, so Word keeps the cached result.

Word recomputes DATE and TIME when it opens a document, so a reference PDF
exported from a fixture carrying one records the day it was exported and can
never be reproduced by anyone else. Two corpus fixtures are built that way, and
both are permanent false positives in any drift screen.

w:fldLock is the mechanism OOXML provides for exactly this (ECMA-376 17.16.18,
"Field Shall Not Be Recalculated"). It leaves the field code in place, so the
fixture still exercises the field renderer, and makes Word display the result
already cached in the package -- which is the behaviour the fixtures claim to
test in the first place.

Only DATE and TIME are locked. CREATEDATE, PAGE, NUMPAGES, SEQ, STYLEREF, REF,
QUOTE and AUTHOR are either stable across exports or are meant to recompute.

The rewrite is byte-level: only the matched begin-fldChar elements change, and
every other byte and every other part is copied through unchanged, so the
locked copy stays a checkable function of its source.

Usage: python3 scripts/lock-volatile-fields.py <source.docx> <destination.docx>
"""
import re
import sys
import zipfile

BEGIN = re.compile(rb'<w:fldChar\s+w:fldCharType="begin"\s*/>')
INSTR = re.compile(rb"<w:instrText[^>]*>([^<]*)</w:instrText>")
VOLATILE = re.compile(rb"^\s*(DATE|TIME)\b", re.IGNORECASE)


def lock(data):
    """Lock each begin whose own field code is DATE or TIME. Fields do not nest
    here, so the instrText that follows a begin is that field's."""
    begins = [m for m in BEGIN.finditer(data)]
    instrs = [m for m in INSTR.finditer(data)]
    edits = []
    for b in begins:
        nxt = next((i for i in instrs if i.start() > b.end()), None)
        if nxt and VOLATILE.match(nxt.group(1)):
            edits.append(b)
    out = bytearray(data)
    for b in reversed(edits):
        out[b.start():b.end()] = b'<w:fldChar w:fldCharType="begin" w:fldLock="true"/>'
    return bytes(out), len(edits)


def main(argv):
    if len(argv) != 2:
        print(__doc__.strip().splitlines()[-1])
        return 1
    source, destination = argv
    with zipfile.ZipFile(source) as z:
        parts = [(info.filename, z.read(info.filename)) for info in z.infolist()]
    total = 0
    rewritten = []
    for name, data in parts:
        if name.endswith(".xml") and b"fldCharType" in data:
            data, count = lock(data)
            total += count
            if count:
                print(f"  {name}: {count} DATE/TIME fields locked")
        rewritten.append((name, data))
    if not total:
        print(f"{source} carries no unlocked DATE or TIME field")
    with zipfile.ZipFile(destination, "w", zipfile.ZIP_DEFLATED) as out:
        for name, data in rewritten:
            out.writestr(name, data)
    print(f"Wrote {destination}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
