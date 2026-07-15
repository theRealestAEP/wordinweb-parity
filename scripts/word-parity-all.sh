#!/bin/zsh
# Export Word ground-truth PDFs for every parity fixture, then compare:
#   scripts/word-parity-all.sh && npm run dev &   # server up
#   npm run parity
# Requires Microsoft Word (grants an automation permission prompt on first run).
set -e
cd "$(dirname "$0")/.."
for f in apps/demo/public/fixtures/parity-*.docx apps/demo/public/fixtures/sample.docx; do
  echo "Exporting $(basename "$f") ..."
  ./scripts/word-parity.sh "$f"
done
echo "Done. Now: npm run dev (in one shell) and npm run parity (in another)."
