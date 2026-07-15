# Scripts

The public scripts support three workflows.

## Visual parity

- `parity-parallel.mjs` runs the complete Word comparison in parallel.
- `parity-compare.mjs` runs selected fixtures or a serial comparison.
- `parity-render-report.mjs` rebuilds the HTML report from saved results.
- `parity-report.mjs` contains the shared report generator.
- `word-parity.sh` exports one Word document to the reference PDF.
- `word-parity-all.sh` exports the standard reference set.

## Fixture safety

- `validate-docx.py` rejects malformed generated fixtures before Word opens them.
- `sanitize-docx.py` anonymizes a document while retaining its layout structure.
- `audit-fixtures.py` scans fixtures for identifying or sensitive information.

## Local fonts

- `extract-dfonts.py` extracts licensed Office fonts for local, git-ignored use.
- `extract-font-metrics.py` regenerates the checked-in font metrics table.

One-off probe generators, forensic readers, and local export experiments are
kept outside the public repository under `internal/scripts/`.
