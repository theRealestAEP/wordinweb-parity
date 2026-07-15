# WordInWeb parity

This repository contains the demo, browser tests, DOCX fixtures, fonts, and
Microsoft Word reference PDFs for
[WordInWeb](https://github.com/theRealestAEP/wordinweb). The engine source is a
Git submodule at `wordinweb/`; it is not copied into this repository.

## Setup

Clone with the engine submodule:

```bash
git clone --recurse-submodules https://github.com/theRealestAEP/wordinweb-parity.git
cd wordinweb-parity
npm install
```

If you already cloned the repository, initialize the submodule with:

```bash
git submodule update --init
```

## Commands

```bash
npm run dev              # demo and parity dashboard at http://localhost:5173
npm run build            # engine packages and production demo
npm test                 # engine unit tests
npm run test:integration # fixture-backed incremental layout tests
npm run test:e2e         # Playwright editor tests
node scripts/parity-parallel.mjs
node scripts/parity-render-report.mjs
```

Append `?perf=1` to the demo URL to show the per-edit performance HUD.

## Repository layout

| Path | Contents |
| --- | --- |
| `wordinweb/` | Git submodule containing the engine and React package. |
| `apps/demo/` | Vite demo, fixture corpus, and demo fonts. |
| `e2e/` | Browser editing and performance tests. |
| `fixtures-staging/` | Additional layout and robustness probes. |
| `parity/` | Microsoft Word reference PDFs. |
| `scripts/` | Fixture and visual-parity tooling. |

## Fonts

The Microsoft fonts in `apps/demo/public/fonts-local` are included only to
reproduce this experimental demo's parity results. Do not copy or use them
outside this demo. Applications embedding WordInWeb must source and license
their own fonts. Without the same fonts, glyphs and line breaks may differ from
Microsoft Word.

## Updating the engine

```bash
git submodule update --remote wordinweb
git add wordinweb
git commit -m "Update wordinweb"
```

The parity commit records the exact engine commit used for its results.

## License

PolyForm Noncommercial 1.0.0. See `LICENSE`.
