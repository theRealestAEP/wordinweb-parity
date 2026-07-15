#!/bin/zsh
# Ground-truth parity check: export a .docx to PDF via local Microsoft Word.
# Usage: scripts/word-parity.sh path/to/doc.docx [out.pdf]
# Compare the PDF pages against the DocxInWeb render of the same file.
set -e
SRC="$(cd "$(dirname "$1")" && pwd)/$(basename "$1")"
OUT="${2:-$(dirname "$0")/../parity/$(basename "${1%.docx}")-word.pdf}"
mkdir -p "$(dirname "$OUT")"
OUT="$(cd "$(dirname "$OUT")" && pwd)/$(basename "$OUT")"
osascript <<EOF
tell application "Microsoft Word"
  set d to open file name "$SRC"
  save as d file name "$OUT" file format format PDF
  close d saving no
end tell
EOF
echo "Wrote $OUT"
