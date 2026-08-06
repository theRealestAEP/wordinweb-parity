#!/usr/bin/env node

/**
 * Saved-DOCX parity gate.
 *
 * For every selected fixture this runner:
 *   1. loads the fixture in the demo and clicks its built-in Download button;
 *   2. exports only that downloaded DOCX to PDF with desktop Microsoft Word;
 *   3. rasterizes the candidate PDF with pdftoppm at 192 DPI;
 *   4. compares it with the persistent raster cache for parity/<name>-word.pdf.
 *
 * Reference DOCX/PDF files are never regenerated here. Browser screenshots,
 * browser-generated PDFs, report PNGs, and comment-card chrome never enter the
 * metric.
 *
 * Usage:
 *   node scripts/word-download-parity.mjs parity-text benchmark
 *   node scripts/word-download-parity.mjs --base http://127.0.0.1:5299
 *
 * Microsoft Word needs Full Disk Access on macOS. Inputs and outputs are staged
 * in Word's own container so the run does not prompt for each file.
 */

import { execFileSync } from "node:child_process";
import {
  existsSync,
  mkdirSync,
  readFileSync,
  readdirSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import { basename, dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { chromium } from "@playwright/test";
import { writeWordDownloadParityReport } from "./word-download-parity-report.mjs";
import {
  comparePngs,
  ensureRasters,
  exportWithWord,
  packageSha256,
  pdfInfo,
  pngs,
  sha256,
  wordIoDir,
} from "./word-export.mjs";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const parityDir = join(root, "parity");
const fixtureDir = join(root, "apps/demo/public/fixtures");
const args = process.argv.slice(2);

function option(name, fallback) {
  const index = args.indexOf(name);
  return index >= 0 ? args[index + 1] : fallback;
}

if (args.includes("--help")) {
  console.log("Usage: node scripts/word-download-parity.mjs [fixture ...] [--base URL] [--out DIR]");
  process.exit(0);
}

const base = option("--base", "http://127.0.0.1:5299");
const outDir = resolve(option(
  "--out",
  join(tmpdir(), `wordinweb-word-download-parity-${new Date().toISOString().replace(/[:.]/g, "-")}`),
));
const optionValues = new Set(["--base", "--out"].flatMap((name) => {
  const index = args.indexOf(name);
  return index >= 0 ? [name, args[index + 1]] : [];
}));
const requested = args.filter((arg) => !optionValues.has(arg) && !arg.startsWith("--"));
const references = readdirSync(parityDir)
  .filter((name) => name.endsWith("-word.pdf"))
  .map((name) => name.slice(0, -"-word.pdf".length))
  .filter((name) => existsSync(join(fixtureDir, `${name}.docx`)))
  .filter((name) => requested.length === 0 || requested.includes(name))
  .sort();

if (references.length === 0) throw new Error("No selected fixtures have cached Word reference PDFs");
for (const name of requested) {
  if (!references.includes(name)) throw new Error(`Missing fixture or cached Word reference: ${name}`);
}

const referenceManifestPath = join(parityDir, "word-reference-manifest.json");
if (!existsSync(referenceManifestPath)) throw new Error(`Missing Word reference manifest: ${referenceManifestPath}`);
const referenceManifest = JSON.parse(readFileSync(referenceManifestPath, "utf8"));
for (const name of references) {
  const entry = referenceManifest.fixtures?.[name];
  if (!entry) throw new Error(`Missing Word reference manifest entry: ${name}`);
  const sourceDocx = join(fixtureDir, `${name}.docx`);
  const referenceDocx = resolve(root, entry.referenceDocx);
  const referencePdf = join(parityDir, `${name}-word.pdf`);
  if (packageSha256(sourceDocx) !== entry.sourcePackageSha256) {
    throw new Error(`${name}: source DOCX changed; intentionally refresh its cached Word reference`);
  }
  if (!existsSync(referenceDocx) || packageSha256(referenceDocx) !== entry.referenceDocxPackageSha256) {
    throw new Error(`${name}: cached reference DOCX is missing or changed`);
  }
  if (sha256(referencePdf) !== entry.referencePdfSha256) {
    throw new Error(`${name}: cached Word reference PDF changed without a manifest refresh`);
  }
}

/**
 * Pages where desktop Word does not compute the same layout every time.
 *
 * wild-doerfp page 35 is the known case (#71): four exports of one package on
 * one build produce two different page 35s, and one of them is byte-identical
 * to the cached reference. Whichever side a run lands on is Word's coin flip,
 * not the document's and not the engine's, so letting it into the aggregate
 * makes the gate's headline number move for no reason anybody can act on.
 *
 * A declared page is still measured, still recorded, and still shown on the
 * report. What changes is that it is kept out of the mean and the worst-page
 * statistic, and held instead to `bistableCeilingPct` — a ceiling set from the
 * measured amplitude of the flip, so a real regression on that page still
 * fails. Tolerating the known wobble is not the same as not looking.
 */
const bistablePagesOf = (name) => new Set(referenceManifest.fixtures?.[name]?.bistablePages ?? []);
const bistableCeilingOf = (name) => referenceManifest.fixtures?.[name]?.bistableCeilingPct ?? 0;

const downloadsDir = join(outDir, "downloads");
const candidatePdfDir = join(outDir, "candidate-pdf");
const candidatePngDir = join(outDir, "candidate-png");
const rasterCacheRoot = join(parityDir, ".raster-cache");
const candidatePdfCacheDir = join(wordIoDir, "candidate-pdf-cache");
const candidateRasterCacheRoot = join(wordIoDir, "candidate-raster-cache");
for (const dir of [outDir, downloadsDir, candidatePdfDir, candidatePngDir, rasterCacheRoot, wordIoDir, candidatePdfCacheDir, candidateRasterCacheRoot]) {
  mkdirSync(dir, { recursive: true });
}

const manifest = {
  schemaVersion: 1,
  generatedAt: new Date().toISOString(),
  base,
  pipeline: {
    candidate: "built-in website Download -> DOCX -> desktop Microsoft Word PDF -> pdftoppm -r 192 PNG",
    reference: "cached parity/<fixture>-word.pdf -> persistent pdftoppm -r 192 PNG",
    excluded: ["website screenshots", "browser PDFs", "report PNGs", "off-page comment UI"],
    mismatchRule: "abs(Rdiff)+abs(Gdiff)+abs(Bdiff) > 90",
  },
  fixtures: [],
  summary: null,
};
const manifestPath = join(outDir, "results.json");
const browser = await chromium.launch();
const page = await browser.newPage();
const metricPage = await browser.newPage();
await metricPage.setContent("<!doctype html><title>Word parity pixel metric</title>");

try {
  await page.goto(base, { waitUntil: "domcontentloaded" });
  for (const [index, name] of references.entries()) {
    console.log(`[${index + 1}/${references.length}] ${name}: built-in Download`);
    await page.goto(`${base}/?doc=/fixtures/${encodeURIComponent(name)}.docx`, { waitUntil: "domcontentloaded" });
    await page.waitForSelector(".dxw-page", { timeout: 120_000 });
    await page.locator("[data-dxw-loading]").waitFor({ state: "detached", timeout: 120_000 }).catch(() => {});
    const pending = page.waitForEvent("download", { timeout: 120_000 });
    await page.getByText("Download", { exact: true }).click();
    const download = await pending;
    const downloadedDocx = join(downloadsDir, `${name}-website.docx`);
    await download.saveAs(downloadedDocx);
    execFileSync("unzip", ["-t", downloadedDocx], { stdio: "ignore" });

    const downloadedDocxSha256 = sha256(downloadedDocx);
    const downloadedPackageSha256 = packageSha256(downloadedDocx);
    console.log(`[${index + 1}/${references.length}] ${name}: Microsoft Word candidate PDF`);
    const candidatePdf = join(candidatePdfDir, `${name}-website-word.pdf`);
    exportWithWord({
      name,
      docx: downloadedDocx,
      destination: candidatePdf,
      packageHash: downloadedPackageSha256,
      cacheDir: candidatePdfCacheDir,
    });
    const referencePdf = join(parityDir, `${name}-word.pdf`);
    const referenceInfo = pdfInfo(referencePdf);
    const candidateInfo = pdfInfo(candidatePdf);
    if (referenceInfo.pages !== referenceManifest.fixtures[name].pages) {
      throw new Error(`${name}: cached Word reference page count changed`);
    }
    if (referenceInfo.pages !== candidateInfo.pages) {
      throw new Error(`${name}: page-count mismatch ${referenceInfo.pages} reference vs ${candidateInfo.pages} candidate`);
    }

    const referenceRasterDir = ensureRasters(
      referencePdf,
      join(rasterCacheRoot, `${basename(referencePdf, ".pdf")}-${sha256(referencePdf)}-r192`),
      "word",
      referenceInfo.pages,
    );
    const fixtureCandidatePngDir = ensureRasters(
      candidatePdf,
      join(candidateRasterCacheRoot, `${sha256(candidatePdf)}-r192`),
      "candidate",
      candidateInfo.pages,
    );
    const referencePngs = pngs(referenceRasterDir, "word");
    const candidatePngs = pngs(fixtureCandidatePngDir, "candidate");
    if (referencePngs.length !== referenceInfo.pages || candidatePngs.length !== candidateInfo.pages) {
      throw new Error(`${name}: incomplete raster set`);
    }

    const bistable = bistablePagesOf(name);
    const pages = [];
    for (let pageIndex = 0; pageIndex < referencePngs.length; pageIndex++) {
      const metric = await comparePngs(metricPage, referencePngs[pageIndex], candidatePngs[pageIndex]);
      pages.push({
        page: pageIndex + 1,
        ...metric,
        mismatchPct: metric.mismatchedPixels * 100 / metric.pixels,
        bistable: bistable.has(pageIndex + 1),
        referencePng: referencePngs[pageIndex],
        candidatePng: candidatePngs[pageIndex],
      });
    }
    manifest.fixtures.push({
      fixture: name,
      sourceDocx: join(fixtureDir, `${name}.docx`),
      sourceDocxSha256: sha256(join(fixtureDir, `${name}.docx`)),
      downloadedDocx,
      downloadedDocxSha256,
      downloadedPackageSha256,
      referencePdf,
      referencePdfSha256: sha256(referencePdf),
      candidatePdf,
      candidatePdfSha256: sha256(candidatePdf),
      referenceCreator: referenceInfo.creator,
      candidateCreator: candidateInfo.creator,
      pages,
    });
    writeFileSync(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`);
  }
} finally {
  await browser.close();
}

const allPages = manifest.fixtures.flatMap((fixture) =>
  fixture.pages.map((page) => ({ ...page, fixture: fixture.fixture })),
);
const bistablePages = allPages.filter((page) => page.bistable);
const pages = allPages.filter((page) => !page.bistable);
const pageMeanPct = pages.reduce((sum, page) => sum + page.mismatchPct, 0) / pages.length;
const totalPixels = pages.reduce((sum, page) => sum + page.pixels, 0);
const mismatchedPixels = pages.reduce((sum, page) => sum + page.mismatchedPixels, 0);
const pixelWeightedMeanPct = mismatchedPixels * 100 / totalPixels;
const worst = pages.reduce((current, page) => page.mismatchPct > current.mismatchPct ? page : current, pages[0]);
// A declared-bistable page is still held to a ceiling, so a real regression on
// it still fails; only the known wobble is tolerated.
const overCeiling = bistablePages.filter((page) => page.mismatchPct >= bistableCeilingOf(page.fixture));
manifest.summary = {
  fixtures: manifest.fixtures.length,
  pages: pages.length,
  totalPixels,
  mismatchedPixels,
  pageMeanPct,
  pixelWeightedMeanPct,
  worstPct: worst.mismatchPct,
  bistable: bistablePages.map((page) => ({
    fixture: page.fixture,
    page: page.page,
    mismatchPct: page.mismatchPct,
    ceilingPct: bistableCeilingOf(page.fixture),
  })),
  thresholds: { pageMeanPct: 0.05, worstPct: 2 },
  passed: pageMeanPct < 0.05 && worst.mismatchPct < 2 && overCeiling.length === 0,
};
for (const page of overCeiling) {
  console.error(
    `${page.fixture} page ${page.page} is declared bistable but reads ${page.mismatchPct.toFixed(6)}%, ` +
      `at or above its ${bistableCeilingOf(page.fixture)}% ceiling — that is more than the flip`,
  );
}
writeFileSync(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`);
if (requested.length === 0) {
  writeWordDownloadParityReport(manifest, join(parityDir, "out"));
}

console.log(`\n${pages.length} pages — mean ${pageMeanPct.toFixed(6)}%, weighted ${pixelWeightedMeanPct.toFixed(6)}%, worst ${worst.mismatchPct.toFixed(6)}%`);
console.log(`Results: ${manifestPath}`);
if (requested.length === 0) console.log(`Report: ${join(parityDir, "out/report.html")}`);
if (!manifest.summary.passed) process.exitCode = 1;
