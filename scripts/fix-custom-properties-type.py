#!/usr/bin/env python3
"""Give docProps/custom.xml its content type, so Word will open the package.

A package that CARRIES docProps/custom.xml and REFERENCES it from _rels/.rels
with the custom-properties relationship must also declare its content type with
an Override. Without one the part falls through the Default for .xml and
arrives as application/xml, and Word refuses to open the document at all: no
repair prompt it will answer over AppleScript, just a modal that never clears,
so the export harness times out with "Word did not finish opening".

Measured as a controlled pair on wild-wirfp: the package rezipped unchanged is
refused, and the same bytes plus this one Override open and export. An untyped
custom.xml that NOTHING references is harmless, because Word never loads it --
the reference is what makes the missing type fatal.

Nothing else in the package is touched, and a package that already declares the
Override is left exactly as it was.

Usage: python3 scripts/fix-custom-properties-type.py <name-or-path> [...]
       python3 scripts/fix-custom-properties-type.py --check     # report only
       python3 scripts/fix-custom-properties-type.py             # fix the corpus
"""
import glob
import os
import re
import shutil
import sys
import zipfile

FIXTURES = os.path.join(os.path.dirname(__file__), "..", "apps", "demo", "public", "fixtures")
PART = "docProps/custom.xml"
CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.custom-properties+xml"
OVERRIDE = f'<Override PartName="/{PART}" ContentType="{CONTENT_TYPE}"/>'


def needs_fix(path):
    """True when the package carries and references custom.xml but never types it."""
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        if PART not in names:
            return False
        content_types = z.read("[Content_Types].xml").decode("utf-8")
        if f'PartName="/{PART}"' in content_types:
            return False
        rels = z.read("_rels/.rels").decode("utf-8")
        return bool(re.search(r"<Relationship[^>]*custom-properties[^>]*/>", rels))


def fix(path):
    """Rewrite the package in place with the Override added."""
    with zipfile.ZipFile(path) as z:
        parts = [(info.filename, z.read(info.filename)) for info in z.infolist()]
    temporary = path + ".fixing"
    with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as out:
        for name, data in parts:
            if name == "[Content_Types].xml":
                text = data.decode("utf-8")
                data = text.replace("</Types>", OVERRIDE + "</Types>").encode("utf-8")
            out.writestr(name, data)
    shutil.move(temporary, path)


def main(argv):
    check_only = "--check" in argv
    argv = [a for a in argv if a != "--check"]
    targets = [a if os.path.exists(a) else os.path.join(FIXTURES, f"{a}.docx") for a in argv]
    if not targets:
        targets = sorted(glob.glob(os.path.join(FIXTURES, "*.docx")))
    affected = []
    for path in targets:
        try:
            if needs_fix(path):
                affected.append(path)
        except Exception as error:
            print(f"  SKIP {os.path.basename(path)}: {error}")
    for path in affected:
        name = os.path.basename(path)
        if check_only:
            print(f"  WOULD FIX {name}")
        else:
            fix(path)
            print(f"  FIXED {name}")
    verb = "need" if check_only else "were given"
    print(f"\n{len(affected)} of {len(targets)} packages {verb} the custom-properties Override")


if __name__ == "__main__":
    main(sys.argv[1:])
