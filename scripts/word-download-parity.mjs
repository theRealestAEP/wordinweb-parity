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
import { createHash } from "node:crypto";
import {
  copyFileSync,
  existsSync,
  mkdirSync,
  readFileSync,
  readdirSync,
  renameSync,
  rmSync,
  statSync,
  writeFileSync,
} from "node:fs";
import { homedir, tmpdir } from "node:os";
import { basename, dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { chromium } from "@playwright/test";
import { unzipSync } from "fflate";
import { writeWordDownloadParityReport } from "./word-download-parity-report.mjs";

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

const downloadsDir = join(outDir, "downloads");
const candidatePdfDir = join(outDir, "candidate-pdf");
const candidatePngDir = join(outDir, "candidate-png");
const rasterCacheRoot = join(parityDir, ".raster-cache");
const wordIoDir = join(homedir(), "Library/Containers/com.microsoft.Word/Data/Documents/WordInWebParity");
const candidatePdfCacheDir = join(wordIoDir, "candidate-pdf-cache");
const candidateRasterCacheRoot = join(wordIoDir, "candidate-raster-cache");
for (const dir of [outDir, downloadsDir, candidatePdfDir, candidatePngDir, rasterCacheRoot, wordIoDir, candidatePdfCacheDir, candidateRasterCacheRoot]) {
  mkdirSync(dir, { recursive: true });
}

function sha256(path) {
  return execFileSync("shasum", ["-a", "256", path], { encoding: "utf8" }).trim().split(/\s+/)[0];
}

function packageSha256(path) {
  const hash = createHash("sha256");
  const files = unzipSync(readFileSync(path));
  for (const name of Object.keys(files).sort()) {
    const bytes = files[name];
    hash.update(name);
    hash.update("\0");
    hash.update(String(bytes.length));
    hash.update("\0");
    hash.update(bytes);
  }
  return hash.digest("hex");
}

function pdfInfo(path) {
  const text = execFileSync("pdfinfo", [path], { encoding: "utf8" });
  const creator = text.match(/^Creator:\s*(.*)$/m)?.[1]?.trim() ?? "";
  const pages = Number(text.match(/^Pages:\s*(\d+)$/m)?.[1] ?? 0);
  // Word occasionally preserves an empty Creator field from the source
  // document. The AppleScript export in this process establishes candidate
  // provenance; reject any non-empty metadata that names another producer.
  if (creator && creator !== "Microsoft Word") {
    throw new Error(`${path} is not a Microsoft Word PDF (Creator=${creator})`);
  }
  if (!Number.isInteger(pages) || pages < 1) throw new Error(`Could not read page count from ${path}`);
  return { creator, pages };
}

function pngs(dir, prefix) {
  return readdirSync(dir)
    .filter((name) => name.startsWith(`${prefix}-`) && name.endsWith(".png"))
    .sort((a, b) => Number(a.match(/-(\d+)\.png$/)?.[1]) - Number(b.match(/-(\d+)\.png$/)?.[1]))
    .map((name) => join(dir, name));
}

function referenceRasterDir(pdf) {
  return join(rasterCacheRoot, `${basename(pdf, ".pdf")}-${sha256(pdf)}-r192`);
}

function ensureReferenceRasters(pdf, expectedPages) {
  const cacheDir = referenceRasterDir(pdf);
  const complete = join(cacheDir, ".complete");
  if (!existsSync(complete) || pngs(cacheDir, "word").length !== expectedPages) {
    const temp = `${cacheDir}.tmp-${process.pid}`;
    rmSync(temp, { recursive: true, force: true });
    mkdirSync(temp, { recursive: true });
    execFileSync("pdftoppm", ["-r", "192", "-png", pdf, join(temp, "word")], { stdio: "inherit" });
    writeFileSync(join(temp, ".complete"), "");
    rmSync(cacheDir, { recursive: true, force: true });
    renameSync(temp, cacheDir);
  }
  return cacheDir;
}

function ensureCandidateRasters(pdf, expectedPages) {
  const cacheDir = join(candidateRasterCacheRoot, `${sha256(pdf)}-r192`);
  const complete = join(cacheDir, ".complete");
  if (!existsSync(complete) || pngs(cacheDir, "candidate").length !== expectedPages) {
    const temp = `${cacheDir}.tmp-${process.pid}`;
    rmSync(temp, { recursive: true, force: true });
    mkdirSync(temp, { recursive: true });
    execFileSync("pdftoppm", ["-r", "192", "-png", pdf, join(temp, "candidate")], { stdio: "inherit" });
    writeFileSync(join(temp, ".complete"), "");
    rmSync(cacheDir, { recursive: true, force: true });
    renameSync(temp, cacheDir);
  }
  return cacheDir;
}

function exportCandidateWithWord(name, docx, destination, packageHash) {
  const cachedPdf = join(candidatePdfCacheDir, `${packageHash}.pdf`);
  if (existsSync(cachedPdf) && statSync(cachedPdf).size > 0) {
    pdfInfo(cachedPdf);
    copyFileSync(cachedPdf, destination);
    console.log(`Reused ${cachedPdf}`);
    return;
  }
  const stagedDocx = join(wordIoDir, `${name}-website.docx`);
  const stagedPdf = join(wordIoDir, `${name}-website-word.pdf`);
  rmSync(stagedDocx, { force: true });
  rmSync(stagedPdf, { force: true });
  copyFileSync(docx, stagedDocx);
  const escapeAppleScript = (value) => value.replaceAll("\\", "\\\\").replaceAll('"', '\\"');
  const script = `with timeout of 600 seconds\n` +
    `tell application "Microsoft Word"\n` +
    `  try\n` +
    `    close document "${escapeAppleScript(basename(stagedDocx))}" saving no\n` +
    `  end try\n` +
    `  open file name "${escapeAppleScript(stagedDocx)}"\n` +
    `  repeat with attempt from 1 to 120\n` +
    `    if exists document "${escapeAppleScript(basename(stagedDocx))}" then exit repeat\n` +
    `    delay 1\n` +
    `  end repeat\n` +
    `  if not (exists document "${escapeAppleScript(basename(stagedDocx))}") then error "Word did not finish opening ${escapeAppleScript(basename(stagedDocx))}"\n` +
    `  set candidateDocument to document "${escapeAppleScript(basename(stagedDocx))}"\n` +
    `  delay 5\n` +
    `  save as candidateDocument file name "${escapeAppleScript(stagedPdf)}" file format format PDF\n` +
    `  close candidateDocument saving no\n` +
    `end tell\n` +
    `end timeout`;
  execFileSync("osascript", ["-e", script], { timeout: 610_000, stdio: "inherit" });
  if (!existsSync(stagedPdf) || statSync(stagedPdf).size === 0) throw new Error(`Word export failed for ${name}`);
  pdfInfo(stagedPdf);
  copyFileSync(stagedPdf, destination);
  copyFileSync(stagedPdf, cachedPdf);
  console.log(`Wrote ${stagedPdf}`);
}

async function comparePngs(page, reference, candidate) {
  const [referenceBytes, candidateBytes] = [readFileSync(reference), readFileSync(candidate)];
  return page.evaluate(async ({ referenceData, candidateData }) => {
    const load = (data) => new Promise((resolve, reject) => {
      const image = new Image();
      image.onload = () => resolve(image);
      image.onerror = () => reject(new Error("PNG decode failed"));
      image.src = `data:image/png;base64,${data}`;
    });
    const [a, b] = await Promise.all([load(referenceData), load(candidateData)]);
    const width = Math.max(a.width, b.width);
    const height = Math.max(a.height, b.height);
    const pixels = (image) => {
      const canvas = document.createElement("canvas");
      canvas.width = width;
      canvas.height = height;
      const context = canvas.getContext("2d", { willReadFrequently: true });
      context.fillStyle = "white";
      context.fillRect(0, 0, width, height);
      context.drawImage(image, 0, 0);
      return context.getImageData(0, 0, width, height).data;
    };
    const aa = pixels(a);
    const bb = pixels(b);
    let mismatchedPixels = 0;
    for (let offset = 0; offset < aa.length; offset += 4) {
      const delta = Math.abs(aa[offset] - bb[offset])
        + Math.abs(aa[offset + 1] - bb[offset + 1])
        + Math.abs(aa[offset + 2] - bb[offset + 2]);
      if (delta > 90) mismatchedPixels++;
    }
    return { width, height, pixels: width * height, mismatchedPixels };
  }, {
    referenceData: referenceBytes.toString("base64"),
    candidateData: candidateBytes.toString("base64"),
  });
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
    exportCandidateWithWord(name, downloadedDocx, candidatePdf, downloadedPackageSha256);
    const referencePdf = join(parityDir, `${name}-word.pdf`);
    const referenceInfo = pdfInfo(referencePdf);
    const candidateInfo = pdfInfo(candidatePdf);
    if (referenceInfo.pages !== referenceManifest.fixtures[name].pages) {
      throw new Error(`${name}: cached Word reference page count changed`);
    }
    if (referenceInfo.pages !== candidateInfo.pages) {
      throw new Error(`${name}: page-count mismatch ${referenceInfo.pages} reference vs ${candidateInfo.pages} candidate`);
    }

    const referenceRasterDir = ensureReferenceRasters(referencePdf, referenceInfo.pages);
    const fixtureCandidatePngDir = ensureCandidateRasters(candidatePdf, candidateInfo.pages);
    const referencePngs = pngs(referenceRasterDir, "word");
    const candidatePngs = pngs(fixtureCandidatePngDir, "candidate");
    if (referencePngs.length !== referenceInfo.pages || candidatePngs.length !== candidateInfo.pages) {
      throw new Error(`${name}: incomplete raster set`);
    }

    const pages = [];
    for (let pageIndex = 0; pageIndex < referencePngs.length; pageIndex++) {
      const metric = await comparePngs(metricPage, referencePngs[pageIndex], candidatePngs[pageIndex]);
      pages.push({
        page: pageIndex + 1,
        ...metric,
        mismatchPct: metric.mismatchedPixels * 100 / metric.pixels,
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

const pages = manifest.fixtures.flatMap((fixture) => fixture.pages);
const pageMeanPct = pages.reduce((sum, page) => sum + page.mismatchPct, 0) / pages.length;
const totalPixels = pages.reduce((sum, page) => sum + page.pixels, 0);
const mismatchedPixels = pages.reduce((sum, page) => sum + page.mismatchedPixels, 0);
const pixelWeightedMeanPct = mismatchedPixels * 100 / totalPixels;
const worst = pages.reduce((current, page) => page.mismatchPct > current.mismatchPct ? page : current, pages[0]);
manifest.summary = {
  fixtures: manifest.fixtures.length,
  pages: pages.length,
  totalPixels,
  mismatchedPixels,
  pageMeanPct,
  pixelWeightedMeanPct,
  worstPct: worst.mismatchPct,
  thresholds: { pageMeanPct: 0.05, worstPct: 2 },
  passed: pageMeanPct < 0.05 && worst.mismatchPct < 2,
};
writeFileSync(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`);
if (requested.length === 0) {
  writeWordDownloadParityReport(manifest, join(parityDir, "out"));
}

console.log(`\n${pages.length} pages — mean ${pageMeanPct.toFixed(6)}%, weighted ${pixelWeightedMeanPct.toFixed(6)}%, worst ${worst.mismatchPct.toFixed(6)}%`);
console.log(`Results: ${manifestPath}`);
if (requested.length === 0) console.log(`Report: ${join(parityDir, "out/report.html")}`);
if (!manifest.summary.passed) process.exitCode = 1;
