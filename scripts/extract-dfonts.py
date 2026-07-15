#!/usr/bin/env python3
"""Extract Word's embedded Office fonts (DFonts) to standalone .ttf files for
DEV-ONLY use by the demo. The output directory apps/demo/public/fonts-local/ is
gitignored — these fonts are licensed and must NEVER be committed. This script
regenerates them from a local Microsoft Word install.

Usage:  python3 scripts/extract-dfonts.py
Requires: fontTools  (pip3 install fonttools)

TrueType Collections (.ttc) can't be loaded by a browser @font-face src, so we
save the single needed face out of each collection as a plain .ttf. Standalone
.ttf files are copied as-is.
"""
import os
import shutil
import sys

try:
    from fontTools.ttLib import TTFont
except ImportError:
    sys.exit("fontTools missing: pip3 install fonttools")

DFONTS = "/Applications/Microsoft Word.app/Contents/Resources/DFonts"
OUT = os.path.join(os.path.dirname(__file__), "..", "apps", "demo", "public", "fonts-local")
OUT = os.path.abspath(OUT)

# (source file, face index in the collection [0 for plain .ttf], output name)
JOBS = [
    # --- Cambria + Cambria Math (the math pipeline's real face) ---
    ("Cambria.ttc", 0, "Cambria.ttf"),
    ("Cambria.ttc", 1, "CambriaMath.ttf"),
    ("Cambriab.ttf", 0, "Cambria-Bold.ttf"),
    ("Cambriai.ttf", 0, "Cambria-Italic.ttf"),
    ("Cambriaz.ttf", 0, "Cambria-BoldItalic.ttf"),
    # --- Times New Roman (Windows) ---
    ("times.ttf", 0, "TimesNewRoman.ttf"),
    ("timesbd.ttf", 0, "TimesNewRoman-Bold.ttf"),
    ("timesi.ttf", 0, "TimesNewRoman-Italic.ttf"),
    ("timesbi.ttf", 0, "TimesNewRoman-BoldItalic.ttf"),
    # --- Calibri (real; currently substituted by Carlito) ---
    ("Calibri.ttf", 0, "Calibri.ttf"),
    ("Calibrib.ttf", 0, "Calibri-Bold.ttf"),
    ("Calibrii.ttf", 0, "Calibri-Italic.ttf"),
    ("Calibriz.ttf", 0, "Calibri-BoldItalic.ttf"),
    ("calibril.ttf", 0, "CalibriLight.ttf"),
    ("calibrili.ttf", 0, "CalibriLight-Italic.ttf"),
    # --- Arial ---
    ("arial.ttf", 0, "Arial.ttf"),
    ("arialbd.ttf", 0, "Arial-Bold.ttf"),
    ("ariali.ttf", 0, "Arial-Italic.ttf"),
    ("arialbi.ttf", 0, "Arial-BoldItalic.ttf"),
    # --- CJK: Japanese ---
    ("msmincho.ttc", 0, "MSMincho.ttf"),
    ("msgothic.ttc", 0, "MSGothic.ttf"),
    ("meiryo.ttc", 0, "Meiryo.ttf"),
    ("meiryob.ttc", 0, "Meiryo-Bold.ttf"),
    ("YuGothR.ttc", 0, "YuGothic.ttf"),
    ("YuGothB.ttc", 0, "YuGothic-Bold.ttf"),
    ("yumin.ttf", 0, "YuMincho.ttf"),
    ("yumindb.ttf", 0, "YuMincho-Bold.ttf"),
    # --- CJK: Chinese ---
    ("Simsun.ttc", 0, "SimSun.ttf"),
    ("SimHei.ttf", 0, "SimHei.ttf"),
    ("MSJH.ttf", 0, "MSJhengHei.ttf"),
    ("MSJHBD.ttf", 0, "MSJhengHei-Bold.ttf"),
    ("msyh.ttc", 0, "MicrosoftYaHei.ttf"),
    # --- Tahoma (wild2-lit-yiddish-rtl's dominant cs face) ---
    ("tahoma.ttf", 0, "Tahoma.ttf"),
    ("tahomabd.ttf", 0, "Tahoma-Bold.ttf"),
    # --- CJK: Korean ---
    ("malgun.ttf", 0, "MalgunGothic.ttf"),
    ("malgunbd.ttf", 0, "MalgunGothic-Bold.ttf"),
    # --- Franklin Gothic Medium (wild-doerfp's H.x sub-headings) ---
    ("Franklin Gothic Medium.ttf", 0, "FranklinGothicMedium.ttf"),
    ("Franklin Gothic Medium Italic.ttf", 0, "FranklinGothicMedium-Italic.ttf"),
    # --- Indic: Mangal (Devanagari) + Latha (Tamil). Word substitutes a
    # Nirmala UI run to these DFonts on export (probe3-indic's ref embeds
    # Mangal for the Devanagari paragraph; Latha is the closest available
    # Tamil face — the ref's Vijaya is not in DFonts on this machine). ---
    ("mangal.ttf", 0, "Mangal.ttf"),
    ("MangalB.ttf", 0, "Mangal-Bold.ttf"),
    ("latha.ttf", 0, "Latha.ttf"),
    ("Lathab.ttf", 0, "Latha-Bold.ttf"),
]


def main():
    if not os.path.isdir(DFONTS):
        sys.exit(f"DFonts not found: {DFONTS} (is Microsoft Word installed?)")
    os.makedirs(OUT, exist_ok=True)
    done = 0
    for src, face, dst in JOBS:
        sp = os.path.join(DFONTS, src)
        op = os.path.join(OUT, dst)
        if not os.path.exists(sp):
            print(f"  SKIP (missing source): {src}")
            continue
        try:
            if src.lower().endswith(".ttc") or face != 0:
                f = TTFont(sp, fontNumber=face)
                f.save(op)
            else:
                shutil.copyfile(sp, op)
            done += 1
            print(f"  {dst}")
        except Exception as e:  # noqa: BLE001
            print(f"  FAIL {dst}: {e}")
    print(f"Extracted {done}/{len(JOBS)} faces to {OUT}")


if __name__ == "__main__":
    main()
