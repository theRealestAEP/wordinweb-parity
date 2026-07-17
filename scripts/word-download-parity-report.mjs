#!/usr/bin/env node

import { copyFileSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const escapeHtml = (value) => String(value)
  .replaceAll("&", "&amp;")
  .replaceAll("<", "&lt;")
  .replaceAll(">", "&gt;")
  .replaceAll('"', "&quot;")
  .replaceAll("'", "&#39;");

const pct = (value) => `${value.toFixed(6)}%`;

export function writeWordDownloadParityReport(manifest, reportDir) {
  if (!manifest.summary || !Array.isArray(manifest.fixtures)) {
    throw new Error("A completed word-download parity manifest is required");
  }

  rmSync(reportDir, { recursive: true, force: true });
  mkdirSync(reportDir, { recursive: true });

  const pages = manifest.fixtures.flatMap((fixture) => fixture.pages.map((page) => ({
    ...page,
    fixture: fixture.fixture,
  })));
  const changedPages = pages
    .filter((page) => page.mismatchedPixels > 0)
    .sort((a, b) => b.mismatchPct - a.mismatchPct);
  const worst = changedPages[0] ?? pages[0];
  const meanPassed = manifest.summary.pageMeanPct < manifest.summary.thresholds.pageMeanPct;
  const worstPassed = manifest.summary.worstPct < manifest.summary.thresholds.worstPct;
  const fixtureRows = manifest.fixtures
    .map((fixture) => {
      const mean = fixture.pages.reduce((sum, page) => sum + page.mismatchPct, 0) / fixture.pages.length;
      const fixtureWorst = fixture.pages.reduce(
        (current, page) => page.mismatchPct > current.mismatchPct ? page : current,
        fixture.pages[0],
      );
      return {
        fixture: fixture.fixture,
        pages: fixture.pages.length,
        mean,
        worst: fixtureWorst,
        changed: fixture.pages.filter((page) => page.mismatchedPixels > 0).length,
      };
    })
    .sort((a, b) => b.worst.mismatchPct - a.worst.mismatchPct || a.fixture.localeCompare(b.fixture));

  for (const page of changedPages) {
    const stem = `${page.fixture}-p${page.page}`;
    copyFileSync(page.referencePng, join(reportDir, `${stem}-reference.png`));
    copyFileSync(page.candidatePng, join(reportDir, `${stem}-candidate.png`));
  }

  const changedPageHtml = changedPages.map((page) => {
    const stem = `${page.fixture}-p${page.page}`;
    return `<details>
      <summary><span>${escapeHtml(page.fixture)} p${page.page}</span><strong>${pct(page.mismatchPct)}</strong></summary>
      <div class="pair">
        <figure><figcaption>Cached Microsoft Word reference</figcaption><a href="${stem}-reference.png" target="_blank"><img src="${stem}-reference.png" loading="lazy" alt="${escapeHtml(page.fixture)} page ${page.page} reference"></a></figure>
        <figure><figcaption>Downloaded DOCX exported by Microsoft Word</figcaption><a href="${stem}-candidate.png" target="_blank"><img src="${stem}-candidate.png" loading="lazy" alt="${escapeHtml(page.fixture)} page ${page.page} candidate"></a></figure>
      </div>
    </details>`;
  }).join("\n");

  const tableRows = fixtureRows.map((fixture) => `<tr>
    <td>${escapeHtml(fixture.fixture)}</td>
    <td>${fixture.pages}</td>
    <td>${pct(fixture.mean)}</td>
    <td>${pct(fixture.worst.mismatchPct)}${fixture.worst.mismatchPct > 0 ? ` · p${fixture.worst.page}` : ""}</td>
    <td>${fixture.changed}</td>
  </tr>`).join("\n");

  const html = `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>WordInWeb Word-export parity</title>
<style>
  :root { color-scheme: dark; --bg:#111; --panel:#181818; --line:#343434; --muted:#b7b7b7; --good:#32d583; --bad:#f05252; --blue:#55a5ff; }
  * { box-sizing:border-box; }
  body { margin:0; background:var(--bg); color:#fff; font:15px/1.45 system-ui,-apple-system,sans-serif; }
  main { max-width:1280px; margin:auto; padding:36px 28px 64px; }
  h1 { margin:0; font-size:32px; }
  h2 { margin:38px 0 8px; }
  .subtitle,.intro,.note { color:var(--muted); }
  .subtitle { margin-top:4px; font-size:13px; }
  .intro { max-width:900px; font-size:17px; }
  .pipeline { display:flex; flex-wrap:wrap; gap:8px; margin:18px 0; }
  .pipeline span { border:1px solid var(--line); border-radius:999px; padding:6px 11px; color:#d8eaff; }
  .kpis { display:grid; grid-template-columns:repeat(3,minmax(0,1fr)); gap:16px; margin:24px 0; }
  .kpi { min-height:130px; padding:20px; border:1px solid var(--line); border-radius:14px; background:var(--panel); }
  .label { color:var(--muted); text-transform:uppercase; letter-spacing:.08em; font-size:12px; }
  .value { margin-top:8px; font-size:36px; font-weight:750; }
  .good { color:var(--good); }
  .bad { color:var(--bad); }
  table { width:100%; border-collapse:collapse; background:var(--panel); border:1px solid var(--line); }
  th,td { padding:10px 12px; border-bottom:1px solid var(--line); text-align:left; }
  th { color:var(--muted); font-size:12px; text-transform:uppercase; letter-spacing:.06em; }
  td:not(:first-child),th:not(:first-child) { text-align:right; font-variant-numeric:tabular-nums; }
  .table-wrap { overflow:auto; border-radius:12px; }
  details { margin:10px 0; border:1px solid var(--line); border-radius:12px; background:var(--panel); }
  summary { display:flex; justify-content:space-between; gap:20px; padding:14px 16px; cursor:pointer; }
  summary strong { color:var(--blue); font-variant-numeric:tabular-nums; }
  .pair { display:grid; grid-template-columns:1fr 1fr; gap:14px; padding:0 14px 14px; }
  figure { margin:0; }
  figcaption { color:var(--muted); margin:0 0 7px; }
  img { display:block; width:100%; height:auto; background:#fff; border-radius:6px; }
  code { color:#d8eaff; }
  @media (max-width:800px) { .kpis,.pair { grid-template-columns:1fr; } }
</style>
</head>
<body><main>
  <header>
    <h1>WordInWeb Word-export parity</h1>
    <div class="subtitle">${escapeHtml(manifest.generatedAt)} · ${manifest.summary.fixtures} fixtures · ${manifest.summary.pages} pages</div>
  </header>
  <p class="intro">Every candidate follows the product's built-in Download path, then desktop Microsoft Word exports that DOCX to PDF. The candidate PDF and the cached Microsoft Word reference PDF are both rasterized with <code>pdftoppm -r 192</code> and compared pixel-by-pixel.</p>
  <div class="pipeline"><span>Built-in Download</span><span>DOCX</span><span>Desktop Microsoft Word PDF</span><span>192 DPI PNG</span><span>Exact pixel comparison</span></div>
  <section class="kpis">
    <div class="kpi"><div class="label">Mean page mismatch · target &lt;${manifest.summary.thresholds.pageMeanPct}%</div><div class="value ${meanPassed ? "good" : "bad"}">${pct(manifest.summary.pageMeanPct)}</div></div>
    <div class="kpi"><div class="label">Worst page · target &lt;${manifest.summary.thresholds.worstPct}%</div><div class="value ${worstPassed ? "good" : "bad"}">${pct(manifest.summary.worstPct)}</div><div class="note">${escapeHtml(worst.fixture)} p${worst.page}</div></div>
    <div class="kpi"><div class="label">Pages passing</div><div class="value good">${pages.filter((page) => page.mismatchPct < manifest.summary.thresholds.worstPct).length} / ${pages.length}</div></div>
    <div class="kpi"><div class="label">Fixtures</div><div class="value">${manifest.summary.fixtures}</div></div>
    <div class="kpi"><div class="label">Pixel-weighted mean</div><div class="value">${pct(manifest.summary.pixelWeightedMeanPct)}</div></div>
    <div class="kpi"><div class="label">Release gate</div><div class="value ${manifest.summary.passed ? "good" : "bad"}">${manifest.summary.passed ? "PASS" : "FAIL"}</div></div>
  </section>
  <h2>Fixtures</h2>
  <p class="note">Sorted by worst page. “Changed pages” counts pages containing any above-threshold RGB pixel delta.</p>
  <div class="table-wrap"><table><thead><tr><th>Fixture</th><th>Pages</th><th>Mean</th><th>Worst</th><th>Changed pages</th></tr></thead><tbody>${tableRows}</tbody></table></div>
  <h2>Changed page pairs</h2>
  <p class="note">${changedPages.length} pages contain a measurable difference. These images come only from the two Microsoft Word PDFs used by the gate.</p>
  ${changedPageHtml || '<p class="note">Every page is identical.</p>'}
</main></body></html>`;

  writeFileSync(join(reportDir, "report.html"), html);
  writeFileSync(join(reportDir, "results.json"), `${JSON.stringify(manifest, null, 2)}\n`);
}

if (process.argv[1] && resolve(process.argv[1]) === resolve(fileURLToPath(import.meta.url))) {
  const resultsPath = resolve(process.argv[2] ?? "");
  const reportDir = resolve(process.argv[3] ?? "");
  if (!process.argv[2] || !process.argv[3]) {
    throw new Error("Usage: node scripts/word-download-parity-report.mjs RESULTS_JSON REPORT_DIR");
  }
  writeWordDownloadParityReport(JSON.parse(readFileSync(resultsPath, "utf8")), reportDir);
  console.log(`Report: ${join(reportDir, "report.html")}`);
}
