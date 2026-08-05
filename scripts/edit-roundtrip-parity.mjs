#!/usr/bin/env node

/**
 * Edit round-trip parity gate (W0).
 *
 * The saved-DOCX gate (scripts/word-download-parity.mjs) proves the website
 * re-serializes a document Word already agreed with. This gate proves the
 * harder half: that a document the website EDITED still means the same thing to
 * Word. For every scenario it
 *
 *   1. loads a fixture in the demo (editable), applies a scripted edit sequence
 *      through the editor api, and clicks the built-in Download;
 *   2. exports that edited DOCX to PDF with desktop Microsoft Word, and treats
 *      a failed open (Word's repair prompt never answers AppleScript) as a
 *      scenario failure;
 *   3. renders the same edited DOCX in the web demo and compares the two
 *      rasters page by page at 192 DPI;
 *   4. re-opens the edited DOCX in the demo and saves it again, requiring
 *      byte-identical output.
 *
 * Results append to parity/edit-roundtrip-history.jsonl, one JSON line per run.
 *
 * Usage:
 *   node scripts/edit-roundtrip-parity.mjs                     # every scenario
 *   node scripts/edit-roundtrip-parity.mjs --scenario typing   # one (repeatable)
 *   node scripts/edit-roundtrip-parity.mjs --base http://localhost:5299
 *
 * The demo dev server must already be running:
 *   npm run dev -w demo -- --port 5299 --strictPort
 */

import { execFileSync } from "node:child_process";
import { appendFileSync, existsSync, lstatSync, mkdirSync, readFileSync, realpathSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { chromium } from "@playwright/test";
import {
  SCALE,
  comparePngs,
  ensureRasters,
  exportWithWord,
  packageSha256,
  pdfInfo,
  pngs,
  sha256,
  wordIoDir,
} from "./word-export.mjs";
import { scenarios } from "./edit-roundtrip-scenarios.mjs";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const parityDir = join(root, "parity");
const fixtureDir = join(root, "apps/demo/public/fixtures");
const historyPath = join(parityDir, "edit-roundtrip-history.jsonl");

const THRESHOLDS = { pageMeanPct: 0.05, worstPct: 2 };

const args = process.argv.slice(2);
if (args.includes("--help")) {
  console.log("Usage: node scripts/edit-roundtrip-parity.mjs [--scenario NAME ...] [--base URL] [--out DIR]");
  console.log(`Scenarios: ${scenarios.map((s) => s.name).join(", ")}`);
  process.exit(0);
}

function option(name, fallback) {
  const index = args.indexOf(name);
  return index >= 0 ? args[index + 1] : fallback;
}

const base = option("--base", "http://localhost:5299");
const outDir = resolve(option(
  "--out",
  join(tmpdir(), `wordinweb-edit-roundtrip-${new Date().toISOString().replace(/[:.]/g, "-")}`),
));
const requested = args.flatMap((arg, index) => (arg === "--scenario" ? [args[index + 1]] : []));
for (const name of requested) {
  if (!scenarios.some((scenario) => scenario.name === name)) {
    throw new Error(`Unknown scenario: ${name}. Known: ${scenarios.map((s) => s.name).join(", ")}`);
  }
}
const selected = requested.length === 0 ? scenarios : scenarios.filter((s) => requested.includes(s.name));

const editedDir = join(outDir, "edited-docx");
const wordPdfDir = join(outDir, "word-pdf");
const webPngDir = join(outDir, "web-png");
const wordPdfCacheDir = join(wordIoDir, "edit-roundtrip-pdf-cache");
const wordRasterCacheDir = join(wordIoDir, "edit-roundtrip-raster-cache");
for (const dir of [outDir, editedDir, wordPdfDir, webPngDir, wordIoDir, wordPdfCacheDir, wordRasterCacheDir]) {
  mkdirSync(dir, { recursive: true });
}

function git(cwd, ...gitArgs) {
  try {
    return execFileSync("git", ["-C", cwd, ...gitArgs], { encoding: "utf8" }).trim();
  } catch {
    return null;
  }
}

/** Which wordinweb build this run measured — the link setup is never touched,
 * only recorded, so a result can be traced back to the engine that produced it. */
function wordinwebBuild() {
  const packageDir = join(root, "node_modules/wordinweb");
  const link = lstatSync(packageDir);
  const symlink = link.isSymbolicLink();
  const target = symlink ? realpathSync(packageDir) : null;
  const manifest = JSON.parse(readFileSync(join(packageDir, "package.json"), "utf8"));
  return {
    version: manifest.version,
    symlink,
    target,
    targetGitSha: target ? git(target, "rev-parse", "HEAD") : null,
    targetGitBranch: target ? git(target, "rev-parse", "--abbrev-ref", "HEAD") : null,
    targetGitDirty: target ? git(target, "status", "--porcelain") !== "" : null,
  };
}

// ---------------------------------------------------------------------------
// The `ed` driver handed to each scenario's edit().
// ---------------------------------------------------------------------------

function editDriver(page) {
  const call = (method, ...callArgs) => page.evaluate(([name, values]) => {
    const api = window.__dxwApi;
    if (!api) throw new Error("editor api hook missing — load the demo with ?apihook=1");
    return api[name](...values);
  }, [method, callArgs]);

  const settle = async () => {
    await page.waitForTimeout(250);
    await page.locator("[data-dxw-layout-status]").waitFor({ state: "hidden", timeout: 60_000 }).catch(() => {});
    await page.waitForTimeout(150);
  };

  // The renderer emits one span per word, so a phrase never lives in a single
  // span: locate it in the concatenated page text and map back to the span the
  // match starts in.
  const locate = (text) => page.evaluate((needle) => {
    const spans = [...document.querySelectorAll(".dxw-page span")].filter((s) => s.childElementCount === 0);
    const starts = [];
    let flat = "";
    for (const span of spans) {
      starts.push(flat.length);
      flat += span.textContent ?? "";
    }
    const at = flat.indexOf(needle);
    if (at < 0) return null;
    let index = 0;
    while (index + 1 < starts.length && starts[index + 1] <= at) index++;
    const span = spans[index];
    span.scrollIntoView({ block: "center" });
    const rect = span.getBoundingClientRect();
    return { x: rect.x, y: rect.y, width: rect.width, height: rect.height };
  }, text);

  return {
    /** Call a DocxViewApi method by name and return its result. */
    async call(method, ...callArgs) {
      const result = await call(method, ...callArgs);
      await settle();
      return result;
    },
    /** Run arbitrary page code (for arguments the bridge cannot serialize, e.g. a Blob). */
    async evaluate(fn, arg) {
      const result = await page.evaluate(fn, arg);
      await settle();
      return result;
    },
    /** Click the first rendered span containing `text` — this is what focuses
     * the editor's hidden input sink, so every scenario starts with one. */
    async clickText(text) {
      const box = await locate(text);
      if (!box) throw new Error(`The rendered pages do not contain ${JSON.stringify(text)}`);
      await page.mouse.click(box.x + Math.min(4, box.width / 2), box.y + box.height / 2);
      await settle();
    },
    /** Select a phrase by content (api.find), so scenarios never depend on pixels. */
    async select(query) {
      const matches = await call("find", query);
      if (!matches) throw new Error(`find() matched nothing for ${JSON.stringify(query)}`);
      await settle();
      return matches;
    },
    async type(text) {
      await page.keyboard.type(text);
      await settle();
    },
    async press(key) {
      await page.keyboard.press(key);
      await settle();
    },
    settle,
    /** getTableCellFill is undefined outside a table — the caret-in-table probe. */
    inTable: () => page.evaluate(() => window.__dxwApi.getTableCellFill() !== undefined),
    async expectRendered(text) {
      if (!(await locate(text))) throw new Error(`Expected the pages to render ${JSON.stringify(text)}`);
    },
    assert(condition, message) {
      if (!condition) throw new Error(message);
    },
  };
}

// ---------------------------------------------------------------------------

/**
 * Load the edited bytes into the demo through its own file picker and wait for
 * the editor that replaces the current one. The demo raises its loading status
 * as soon as it accepts the file and drops it when the document is laid out, so
 * that transition — not the page text — is the signal; a scenario whose edit
 * changes no text (formatting) would otherwise be indistinguishable from the
 * document still on screen.
 *
 * `expectApi` only holds in editable mode: DocxView builds the api behind an
 * `editable` guard, so a viewing-mode page never publishes the hook.
 */
async function loadEditedDocx(page, docx, { expectApi }) {
  if (expectApi) await page.evaluate(() => { delete window.__dxwApi; });
  const loading = page.locator("[data-dxw-loading]");
  await page.locator("#docx-upload").setInputFiles(docx);
  // A fast parse can clear the status before this observes it, so only the
  // disappearance is required.
  await loading.waitFor({ state: "attached", timeout: 15_000 }).catch(() => {});
  await loading.waitFor({ state: "detached", timeout: 120_000 });
  await page.waitForSelector(".dxw-page span", { state: "attached", timeout: 120_000 });
  if (expectApi) await page.waitForFunction(() => Boolean(window.__dxwApi), null, { timeout: 120_000 });
}

async function downloadDocx(page, destination) {
  const pending = page.waitForEvent("download", { timeout: 120_000 });
  // Keep the rejection handled: if the click below fails, an unawaited download
  // promise rejects on page close and buries the error that actually mattered.
  pending.catch(() => {});
  await page.getByText("Download", { exact: true }).click();
  await (await pending).saveAs(destination);
  // A package Word cannot open is a round-trip failure, not a Word failure;
  // catch it here where the error still names the zip entry.
  execFileSync("unzip", ["-t", destination], { stdio: "ignore" });
  return destination;
}

/** Screenshot every rendered page of the edited document in viewing mode.
 * deviceScaleFactor 2 puts the web raster at 192 DPI — the same grid pdftoppm
 * writes for Word's PDF, so pages compare pixel for pixel. */
async function renderWebPages(browser, scenario, docx, directory) {
  const page = await browser.newPage({ viewport: { width: 1700, height: 1200 }, deviceScaleFactor: SCALE });
  try {
    await page.goto(`${base}/?doc=/fixtures/${scenario.fixture}.docx&editable=0&comments=0`, {
      waitUntil: "domcontentloaded",
    });
    await page.waitForSelector(".dxw-page span", { state: "attached", timeout: 120_000 });
    await loadEditedDocx(page, docx, { expectApi: false });
    await page.waitForFunction(
      () => !document.querySelector(".dxw-body-mode, .dxw-hf-mode, .dxw-comment-hl, .dxw-comment-card, .dxw-sel"),
      null,
      { timeout: 60_000 },
    );
    await page.waitForTimeout(1200); // fonts + images settle
    const count = await page.locator(".dxw-page").count();
    const files = [];
    for (let index = 0; index < count; index++) {
      const element = page.locator(".dxw-page").nth(index);
      await element.scrollIntoViewIfNeeded();
      await page.waitForTimeout(100);
      const file = join(directory, `web-${index + 1}.png`);
      writeFileSync(file, await element.screenshot());
      files.push(file);
    }
    return files;
  } finally {
    await page.close();
  }
}

async function runScenario(browser, metricPage, scenario) {
  const startedAt = Date.now();
  const result = {
    scenario: scenario.name,
    fixture: scenario.fixture,
    description: scenario.description,
    failures: [],
    editedDocxSha256: null,
    editedPackageSha256: null,
    resavedDocxSha256: null,
    byteStableResave: null,
    wordOpened: null,
    wordError: null,
    wordPdfSha256: null,
    wordPages: null,
    webPages: null,
    pageMeanPct: null,
    worstPct: null,
    pages: [],
  };
  if (!existsSync(join(fixtureDir, `${scenario.fixture}.docx`))) {
    result.failures.push(`Missing fixture ${scenario.fixture}.docx`);
    return { ...result, passed: false, durationMs: Date.now() - startedAt };
  }

  const editPage = await browser.newPage({ viewport: { width: 1700, height: 1200 } });
  const editedDocx = join(editedDir, `${scenario.name}.docx`);
  try {
    await editPage.goto(`${base}/?doc=/fixtures/${scenario.fixture}.docx&apihook=1`, { waitUntil: "domcontentloaded" });
    await editPage.waitForSelector(".dxw-page span", { state: "attached", timeout: 120_000 });
    await editPage.waitForFunction(() => Boolean(window.__dxwApi), null, { timeout: 120_000 });
    await editPage.waitForTimeout(500);

    const driver = editDriver(editPage);
    await scenario.edit(driver);
    await driver.settle();
    await downloadDocx(editPage, editedDocx);
    result.editedDocxSha256 = sha256(editedDocx);
    result.editedPackageSha256 = packageSha256(editedDocx);

    // Byte-stable re-save: open what we just wrote and write it again.
    const resavedDocx = join(editedDir, `${scenario.name}-resaved.docx`);
    await loadEditedDocx(editPage, editedDocx, { expectApi: true });
    await editPage.waitForTimeout(600);
    await downloadDocx(editPage, resavedDocx);
    result.resavedDocxSha256 = sha256(resavedDocx);
    result.byteStableResave = result.resavedDocxSha256 === result.editedDocxSha256;
    if (!result.byteStableResave) {
      result.failures.push(
        `Re-saving the edited DOCX changed its bytes (${result.editedDocxSha256.slice(0, 12)} -> ${result.resavedDocxSha256.slice(0, 12)})`,
      );
    }
  } catch (error) {
    result.failures.push(`Edit sequence failed: ${error.message}`);
    return { ...result, passed: false, durationMs: Date.now() - startedAt };
  } finally {
    if (!editPage.isClosed()) await editPage.close();
  }

  const wordPdf = join(wordPdfDir, `${scenario.name}-word.pdf`);
  try {
    exportWithWord({
      name: `edit-${scenario.name}`,
      docx: editedDocx,
      destination: wordPdf,
      packageHash: result.editedPackageSha256,
      cacheDir: wordPdfCacheDir,
    });
    result.wordOpened = true;
  } catch (error) {
    // Word's repair prompt is modal and never answers AppleScript, so a damaged
    // package arrives here as an open error or the 600s timeout.
    result.wordOpened = false;
    result.wordError = error.message;
    result.failures.push(`Word could not open or export the edited DOCX: ${error.message}`);
    return { ...result, passed: false, durationMs: Date.now() - startedAt };
  }

  result.wordPdfSha256 = sha256(wordPdf);
  const info = pdfInfo(wordPdf);
  result.wordPages = info.pages;
  const scenarioWebDir = join(webPngDir, scenario.name);
  mkdirSync(scenarioWebDir, { recursive: true });
  const webPngs = await renderWebPages(browser, scenario, editedDocx, scenarioWebDir);
  result.webPages = webPngs.length;
  if (result.wordPages !== result.webPages) {
    result.failures.push(`Page-count mismatch: Word ${result.wordPages} vs web ${result.webPages}`);
  }

  const wordPngs = pngs(
    ensureRasters(wordPdf, join(wordRasterCacheDir, `${result.wordPdfSha256}-r192`), "word", info.pages),
    "word",
  );
  for (let index = 0; index < Math.min(wordPngs.length, webPngs.length); index++) {
    const metric = await comparePngs(metricPage, wordPngs[index], webPngs[index]);
    result.pages.push({
      page: index + 1,
      ...metric,
      mismatchPct: metric.mismatchedPixels * 100 / metric.pixels,
      wordPng: wordPngs[index],
      webPng: webPngs[index],
    });
  }
  if (result.pages.length === 0) {
    result.failures.push("No comparable pages");
    return { ...result, passed: false, durationMs: Date.now() - startedAt };
  }
  result.pageMeanPct = result.pages.reduce((sum, page) => sum + page.mismatchPct, 0) / result.pages.length;
  result.worstPct = result.pages.reduce((worst, page) => Math.max(worst, page.mismatchPct), 0);
  if (result.pageMeanPct >= THRESHOLDS.pageMeanPct) {
    result.failures.push(`Mean mismatch ${result.pageMeanPct.toFixed(4)}% >= ${THRESHOLDS.pageMeanPct}%`);
  }
  if (result.worstPct >= THRESHOLDS.worstPct) {
    result.failures.push(`Worst page ${result.worstPct.toFixed(4)}% >= ${THRESHOLDS.worstPct}%`);
  }
  return { ...result, passed: result.failures.length === 0, durationMs: Date.now() - startedAt };
}

// ---------------------------------------------------------------------------

try {
  await fetch(base);
} catch {
  console.error(`Demo server not reachable at ${base} — start it with \`npm run dev -w demo -- --port 5299 --strictPort\`.`);
  process.exit(1);
}

const browser = await chromium.launch();
const metricPage = await browser.newPage();
await metricPage.setContent("<!doctype html><title>Edit round-trip pixel metric</title>");

const results = [];
try {
  for (const [index, scenario] of selected.entries()) {
    console.log(`\n[${index + 1}/${selected.length}] ${scenario.name} (${scenario.fixture})`);
    const result = await runScenario(browser, metricPage, scenario);
    results.push(result);
    const detail = result.pages.length
      ? `mean ${result.pageMeanPct.toFixed(4)}%, worst ${result.worstPct.toFixed(4)}%, ${result.pages.length} page(s)`
      : "no pages compared";
    console.log(`  ${result.passed ? "PASS" : "FAIL"} — ${detail}`);
    for (const failure of result.failures) console.log(`    - ${failure}`);
  }
} finally {
  await browser.close();
}

const record = {
  schemaVersion: 1,
  ts: new Date().toISOString(),
  base,
  outDir,
  thresholds: THRESHOLDS,
  repo: {
    gitSha: git(root, "rev-parse", "HEAD"),
    gitBranch: git(root, "rev-parse", "--abbrev-ref", "HEAD"),
    gitDirty: git(root, "status", "--porcelain") !== "",
  },
  wordinweb: wordinwebBuild(),
  pipeline: {
    edited: "fixture -> demo editor api edit sequence -> built-in Download -> DOCX",
    word: "edited DOCX -> desktop Microsoft Word PDF -> pdftoppm -r 192 PNG",
    web: "edited DOCX -> demo viewing mode -> .dxw-page screenshot at deviceScaleFactor 2 (192 DPI)",
    mismatchRule: "abs(Rdiff)+abs(Gdiff)+abs(Bdiff) > 90",
  },
  scenarios: results,
  summary: {
    scenarios: results.length,
    passed: results.filter((r) => r.passed).length,
    failed: results.filter((r) => !r.passed).length,
  },
  passed: results.every((r) => r.passed),
};
writeFileSync(join(outDir, "results.json"), `${JSON.stringify(record, null, 2)}\n`);
appendFileSync(historyPath, `${JSON.stringify(record)}\n`);

console.log(`\n${record.summary.passed}/${record.summary.scenarios} scenarios passed`);
console.log(`Results: ${join(outDir, "results.json")}`);
console.log(`History: ${historyPath}`);
if (!record.passed) process.exitCode = 1;
