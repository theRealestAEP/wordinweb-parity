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
node scripts/parity-parallel.mjs
node scripts/parity-render-report.mjs
```

Append `?perf=1` to the demo URL to show the per-edit performance HUD.

## Repository layout

| Path | Contents |
| --- | --- |
| `apps/demo/` | Vite demo, fixture corpus, and demo fonts. |
| `e2e/` | Browser editing and performance tests. |
| `fixtures-staging/` | Additional layout and robustness probes. |
| `parity/` | Microsoft Word reference PDFs. |
| `scripts/` | Fixture and visual-parity tooling. |

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
