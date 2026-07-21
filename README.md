# WordInWeb parity

This repository contains the demo, browser tests, DOCX fixtures, fonts, and
Microsoft Word reference PDFs for
[WordInWeb](https://github.com/theRealestAEP/wordinweb).

## Setup

Clone and install:

```bash
git clone https://github.com/theRealestAEP/wordinweb-parity.git
cd wordinweb-parity
npm install
```

## Commands

```bash
npm run dev              # demo and parity dashboard at http://localhost:5173
npm run build            # production demo
npm test                 # Playwright editor tests
npm run test:e2e         # Playwright editor tests
npm run test:interop     # saved-DOCX table/tab smoke checks in LibreOffice
npm run test:interop:google # same checks through Google Docs import/export
node scripts/parity-parallel.mjs
node scripts/parity-render-report.mjs
```

Append `?perf=1` to the demo URL to show the per-edit performance HUD.

## Repository layout

| Path | Contents |
| --- | --- |
| `apps/demo/` | Vite demo, fixture corpus, and demo fonts. |
| `e2e/` | Browser editing and performance tests. |
| `fixtures-staging/` | Fixture provenance, staging notes, and additional layout probes. |
| `parity/` | Microsoft Word reference PDFs. |
| `scripts/` | Fixture and visual-parity tooling. |

## Cross-editor compatibility

`npm run test:interop` saves four table-heavy fixtures and the tab-stop parity
fixture through WordInWeb, opens
the results with headless LibreOffice, exports PDF and DOCX copies, and checks
table/row/cell retention, tab-stop retention, text retention, page-count bounds,
and nonblank page content. This is a structural smoke gate rather than a
pixel-parity score.

`npm run test:interop:google` runs the same saved files through native Google
Docs import and export. Set either `GOOGLE_INTEROP_ACCESS_TOKEN` or
`GOOGLE_INTEROP_SERVICE_ACCOUNT_JSON`; the latter accepts inline JSON or a JSON
file path. Set `GOOGLE_INTEROP_FOLDER_ID` when the service account writes into a
shared Drive folder. Imported smoke-test documents are deleted after export.
Successful runs refresh the Google Docs and LibreOffice tabs on the parity
report with page previews for every compatibility fixture.

LibreOffice and Poppler (`pdfinfo`, `pdftotext`, and `pdftoppm`) are required for
both commands. Add `--keep` to the script command to retain its temporary
artifacts for inspection.

## Fonts

The Microsoft fonts in `apps/demo/public/fonts-local` are included only to
reproduce this experimental demo's parity results. Do not copy or use them
outside this demo.
## Updating WordInWeb

```bash
npm install wordinweb@latest -w demo
```

The lockfile records the exact published package version used by the demo and browser tests.

## License

MIT. See `LICENSE`.
