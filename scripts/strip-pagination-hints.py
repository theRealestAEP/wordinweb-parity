#!/usr/bin/env python3
"""Remove every w:lastRenderedPageBreak, so Word must compute the pagination.

Word stores where IT last broke the pages as w:lastRenderedPageBreak elements
and replays them on open, until some edit disturbs the file. A reference PDF
exported from a package that still carries them can therefore show a layout
Word would no longer compute -- it is a recording, not ground truth. Stripping
the hints forces a real layout pass, which is what a reference has to be.

The rewrite is byte-level: only the hint elements leave the part, and every
other byte and every other part is copied through unchanged. That keeps the
stripped copy provably a function of its source, so its lineage can be checked
later by stripping the source again and comparing.

Usage: python3 scripts/strip-pagination-hints.py <source.docx> <destination.docx>
"""
import shutil
import sys
import zipfile

HINT = b"<w:lastRenderedPageBreak/>"


def main(argv):
    if len(argv) != 2:
        print(__doc__.strip().splitlines()[-1])
        return 1
    source, destination = argv
    with zipfile.ZipFile(source) as z:
        parts = [(info.filename, z.read(info.filename)) for info in z.infolist()]
    removed = {name: data.count(HINT) for name, data in parts if HINT in data}
    if not removed:
        shutil.copyfile(source, destination)
        print(f"{source} carries no pagination hints; copied byte for byte")
        return 0
    with zipfile.ZipFile(destination, "w", zipfile.ZIP_DEFLATED) as out:
        for name, data in parts:
            out.writestr(name, data.replace(HINT, b""))
    for name, count in removed.items():
        print(f"  {name}: {count} hints removed")
    print(f"Wrote {destination}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
