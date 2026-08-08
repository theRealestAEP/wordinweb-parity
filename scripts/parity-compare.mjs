#!/usr/bin/env node
/**
 * Visual parity check: Word ground truth vs the WordInWeb render.
 *
 * For each parity/<name>-word.pdf (produced by scripts/word-parity.sh) with a
 * matching apps/demo/public/fixtures/<name>.docx, renders both sides page by
 * page and writes parity/out/<name>-p<N>.png — Word | web | diff overlay —
 * plus a mismatch percentage per page.
 *
 * Usage:
 *   npm run dev              # demo server (default http://localhost:5299)
 *   node scripts/parity-compare.mjs [name ...] [--base http://localhost:5299]
 *
 * Use scripts/parity-parallel.mjs for every full-corpus or large parity run.
 * Direct runs of this file are intended for selected fixtures and as workers
 * launched by the parallel runner.
 *
 * Requires poppler (`brew install poppler`) for pdftoppm.
 */
import { execFileSync } from "node:child_process";
import { appendFileSync, copyFileSync, existsSync, linkSync, mkdirSync, readFileSync, readdirSync, renameSync, rmSync, statSync, writeFileSync } from "node:fs";
import { basename, dirname, join } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { chromium } from "@playwright/test";
import {
  APPEARANCE_METRIC_VERSION,
  buildReport,
  METRIC_VERSION,
} from "./parity-report.mjs";
import { pageMetric } from "./parity-metric.mjs";
import { describeBuild, wordinwebBuild } from "./engine-provenance.mjs";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const parityDir = join(root, "parity");
const outDir = join(parityDir, "out");
const interopReportFile = join(root, "apps", "demo", "public", "interop", "results.json");
const interop = existsSync(interopReportFile) ? JSON.parse(readFileSync(interopReportFile, "utf8")) : null;
const scratchDir = join(outDir, ".tmp");
const args = process.argv.slice(2);
const baseIdx = args.indexOf("--base");
const base = baseIdx >= 0 ? args[baseIdx + 1] : "http://localhost:5299";
const acceptRun = args.includes("--accept");
const labelIdx = args.indexOf("--label");
const runLabel = labelIdx >= 0 ? args[labelIdx + 1] : null;
const runSlug = runLabel?.replace(/[^a-z0-9._-]+/gi, "-") ?? null;
// Only argv[baseIdx + 1] is the base value — guard so the first fixture name
// isn't swallowed when --base is absent (baseIdx === -1 makes baseIdx + 1 === 0).
const only = args.filter(
  (a, i) =>
    !a.startsWith("--") &&
    !(baseIdx >= 0 && i === baseIdx + 1) &&
    !(labelIdx >= 0 && i === labelIdx + 1),
);
const isFullRun = only.length === 0;
// Shard workers (parity-parallel.mjs) always write plain accepted-style page
// PNGs - the orchestrator owns the accepted results.json for the whole run.
const outcome = isFullRun || acceptRun || process.env.DXW_PARITY_SHARD_OUT ? "accepted" : "candidate";

// Optional candidate-page subset for bounded calibration runs. Entries are
// `fixture:page`, comma-separated. A fixture argument is still required, so a
// page subset can never accidentally become an accepted full-corpus run.
const selectedPages = process.env.DXW_PARITY_PAGES ? new Map() : null;
for (const entry of process.env.DXW_PARITY_PAGES?.split(",") ?? []) {
  const [fixture, rawPage] = entry.split(":");
  const pageNumber = Number(rawPage);
  if (!fixture || !Number.isInteger(pageNumber) || pageNumber < 1) {
    throw new Error(`Invalid DXW_PARITY_PAGES entry: ${entry}`);
  }
  const pages = selectedPages.get(fixture) ?? new Set();
  pages.add(pageNumber);
  selectedPages.set(fixture, pages);
}
if (selectedPages && isFullRun) {
  throw new Error("DXW_PARITY_PAGES requires explicit fixture arguments");
}

const SCALE = 2; // device pixels per CSS px; pdftoppm dpi = 96 * SCALE
const refs = readdirSync(parityDir)
  .filter((f) => f.endsWith("-word.pdf"))
  .map((f) => f.replace(/-word\.pdf$/, ""))
  .filter((n) => existsSync(join(root, "apps/demo/public/fixtures", `${n}.docx`)))
  .filter((n) => only.length === 0 || only.includes(n));

if (refs.length === 0) {
  console.error("No matching parity references. Export one with scripts/word-parity.sh first.");
  process.exit(1);
}

try {
  await fetch(base);
} catch {
  console.error(`Demo server not reachable at ${base} — start it with \`npm run dev\` first.`);
  process.exit(1);
}

// Which engine these numbers describe. Recorded into results.json and every
// history entry, so a corpus number can always be traced back to the build that
// produced it — the same stamp scripts/edit-roundtrip-parity.mjs carries.
const wordinweb = wordinwebBuild();
console.log(describeBuild(wordinweb));
if (wordinweb.shadowed) {
  // Refuse rather than warn, for the same reason the edit round-trip gate does:
  // a shadowed checkout measures an engine nobody selected, and use-engine.mjs
  // sets both locations together, so there is no case where continuing is right.
  console.error(
    `\nERROR: apps/demo/node_modules/wordinweb shadows the root link, so the demo does NOT\n` +
    `load the engine the root link names. Results would describe an engine nobody chose.\n\n` +
    `  demo loads : ${wordinweb.version} at ${wordinweb.target}\n` +
    `  root link  : ${wordinweb.shadowedRootLink.version} at ${wordinweb.shadowedRootLink.target}\n\n` +
    `Select one engine for both, then restart the dev server:\n` +
    `  node scripts/use-engine.mjs <path-to-engine-react-pkg>\n` +
    `  node scripts/use-engine.mjs npm:${wordinweb.version}\n`,
  );
  process.exit(1);
}

mkdirSync(outDir, { recursive: true });
mkdirSync(scratchDir, { recursive: true });
process.env.TMPDIR = scratchDir;
const browser = await chromium.launch();

const results = []; // { fixture, page, mismatchPct, pngRel }

const pageCountFromPdf = (pdf) => {
  const info = execFileSync("pdfinfo", [pdf]).toString();
  const match = info.match(/^Pages:\s+(\d+)/m);
  if (!match) throw new Error(`Could not read page count from ${pdf}`);
  return Number(match[1]);
};

const semanticLayerSpecs = {
  text: ["-dFILTERIMAGE", "-dFILTERVECTOR"],
  image: ["-dFILTERTEXT", "-dFILTERVECTOR"],
  vector: ["-dFILTERTEXT", "-dFILTERIMAGE"],
};

/** Reference rasters are deterministic (same PDF, same dpi, same poppler),
 * so they are rendered once into a persistent cache keyed by the PDF's
 * mtime+size and hardlinked/copied into the work dir on every later run.
 * Delete parity/.raster-cache to force re-rasterization (e.g. after a
 * poppler upgrade). */
const rasterCacheRoot = join(parityDir, ".raster-cache");

function cachedReferenceDir(pdf) {
  const st = statSync(pdf);
  const key = `${basename(pdf, ".pdf")}-${st.size}-${Math.round(st.mtimeMs)}-r${96 * SCALE}`;
  return join(rasterCacheRoot, key);
}

function linkOrCopy(src, dest) {
  try {
    linkSync(src, dest);
  } catch {
    copyFileSync(src, dest);
  }
}

function renderFullReference(pdf, work, pages) {
  const cacheDir = cachedReferenceDir(pdf);
  const stamp = join(cacheDir, ".complete");
  // Parallel shards of one fixture race to fill the same cache entry: one
  // filler wins the lock and rasterizes into a temp dir renamed into place;
  // the rest wait on the stamp. sleep(1) via execFileSync keeps this dep-free.
  const lock = `${cacheDir}.lock`;
  // The lock is a non-recursive mkdir; with no cache root (fresh worktree) it
  // throws EVERY iteration and the wait loop spins forever.
  mkdirSync(rasterCacheRoot, { recursive: true });
  while (!existsSync(stamp)) {
    let locked = false;
    try {
      mkdirSync(lock, { recursive: false });
      locked = true;
    } catch {
      execFileSync("sleep", ["1"]);
      continue;
    }
    try {
      if (existsSync(stamp)) break;
      const tmp = `${cacheDir}.tmp-${process.pid}`;
      rmSync(tmp, { recursive: true, force: true });
      rmSync(cacheDir, { recursive: true, force: true });
      mkdirSync(tmp, { recursive: true });
      execFileSync("pdftoppm", ["-r", String(96 * SCALE), "-png", pdf, join(tmp, "word")]);
      writeFileSync(join(tmp, ".complete"), "");
      renameSync(tmp, cacheDir);
    } finally {
      if (locked) rmSync(lock, { recursive: true, force: true });
    }
  }
  const cached = readdirSync(cacheDir).filter((f) => f.endsWith(".png"));
  // pdftoppm zero-pads page numbers by document width (word-007.png); the
  // comparison loop expects unpadded names for page subsets and padded names
  // for full runs, so link BOTH spellings.
  for (const f of cached) {
    const m = f.match(/^word-0*(\d+)\.png$/);
    if (!m) continue;
    const pageNumber = Number(m[1]);
    if (pages && !pages.includes(pageNumber)) continue;
    linkOrCopy(join(cacheDir, f), join(work, f));
    const bare = `word-${pageNumber}.png`;
    if (bare !== f) linkOrCopy(join(cacheDir, f), join(work, bare));
  }
}

function renderSemanticReferences(pdf, work, pages, needs) {
  // DXW_PARITY_FAST=1 skips the ghostscript semantic-layer extraction (the
  // appearance columns become n/a; the structural severity metric is
  // unaffected). Large references (yiddish 215pp, nih 419pp) take tens of
  // minutes per layer in gs — fast mode keeps full-suite runs tractable.
  if (process.env.DXW_PARITY_FAST) return;
  for (const [layer, filters] of Object.entries(semanticLayerSpecs)) {
    if (!needs[layer]) continue;
    if (pages) {
      for (const pageNumber of pages) {
        const layerPdf = join(work, `${layer}-${pageNumber}.pdf`);
        execFileSync("gs", [
          "-q",
          "-dSAFER",
          "-dBATCH",
          "-dNOPAUSE",
          "-sDEVICE=pdfwrite",
          ...filters,
          `-dFirstPage=${pageNumber}`,
          `-dLastPage=${pageNumber}`,
          `-sOutputFile=${layerPdf}`,
          pdf,
        ]);
        execFileSync("pdftocairo", [
          "-png",
          "-transp",
          "-singlefile",
          "-r", String(96 * SCALE),
          layerPdf,
          join(work, `${layer}-${pageNumber}`),
        ]);
      }
      continue;
    }

    const layerPdf = join(work, `${layer}.pdf`);
    execFileSync("gs", [
      "-q",
      "-dSAFER",
      "-dBATCH",
      "-dNOPAUSE",
      "-sDEVICE=pdfwrite",
      ...filters,
      `-sOutputFile=${layerPdf}`,
      pdf,
    ]);
    execFileSync("pdftocairo", [
      "-png",
      "-transp",
      "-r", String(96 * SCALE),
      layerPdf,
      join(work, layer),
    ]);
  }
}

async function captureSemanticPage(page, pageIndex) {
  const pageEl = page.locator(".dxw-page").nth(pageIndex);
  const metadata = await pageEl.evaluate((sourcePage) => {
    document.getElementById("dxw-semantic-capture")?.remove();
    const pageRect = sourcePage.getBoundingClientRect();
    const rectFor = (element) => {
      const rect = element.getBoundingClientRect();
      return {
        x: rect.left - pageRect.left,
        y: rect.top - pageRect.top,
        width: rect.width,
        height: rect.height,
      };
    };
    const elements = (selector) => Array.from(sourcePage.querySelectorAll(selector));
    const text = elements('[data-dxw-item-kind="text"]').map((element) => ({
      ...rectFor(element),
      family: element.dataset.dxwFontFamily ?? "",
      size: Number(element.dataset.dxwFontSize ?? 0),
      weight: element.dataset.dxwFontWeight ?? "400",
      style: element.dataset.dxwFontStyle ?? "normal",
    }));
    const images = elements('[data-dxw-item-kind="image"]').map((element) => ({
      ...rectFor(element),
      format: element.dataset.dxwImageFormat ?? "",
    }));
    const fills = elements('[data-dxw-role="table-fill"]').map(rectFor);
    const rules = elements('[data-dxw-role="table-rule"]').map((element) => {
      const rect = rectFor(element);
      return { ...rect, axis: rect.width >= rect.height ? "horizontal" : "vertical" };
    });

    const wrapper = document.createElement("div");
    wrapper.id = "dxw-semantic-capture";
    wrapper.style.cssText =
      `position:absolute;left:0;top:${document.documentElement.scrollHeight + 100}px;` +
      `display:flex;gap:0;width:${pageRect.width * 4}px;height:${pageRect.height}px;` +
      "margin:0;padding:0;background:transparent;isolation:isolate;z-index:2147483647;";
    const panels = ["text", "image", "table-fill", "table-rule"];
    for (const panel of panels) {
      const clone = sourcePage.cloneNode(true);
      clone.removeAttribute("id");
      clone.style.margin = "0";
      clone.style.boxShadow = "none";
      clone.style.background = "transparent";
      clone.style.flex = "none";
      for (const element of clone.querySelectorAll("[data-dxw-item-kind]")) {
        const visible =
          (panel === "text" && element.dataset.dxwItemKind === "text") ||
          (panel === "image" && element.dataset.dxwItemKind === "image") ||
          element.dataset.dxwRole === panel;
        if (!visible) element.style.visibility = "hidden";
        if (visible && panel === "text") element.style.textDecoration = "none";
      }
      wrapper.appendChild(clone);
    }
    document.body.appendChild(wrapper);
    return { text, images, fills, rules, width: pageRect.width, height: pageRect.height };
  });

  const capture = page.locator("#dxw-semantic-capture");
  const png = await capture.screenshot({ animations: "disabled", omitBackground: true });
  await capture.evaluate((element) => element.remove());
  return { png, metadata };
}

for (const name of refs) {
  const pdf = join(parityDir, `${name}-word.pdf`);
  // Suffixed with pid: parallel shards of one fixture each get their own
  // work dir (a shared dir gets rmSync'd by a sibling mid-compare).
  const work = join(scratchDir, `dxw-parity-${name}-${process.pid}`);
  const pageSubset = selectedPages?.get(name);
  if (selectedPages && !pageSubset) {
    throw new Error(`DXW_PARITY_PAGES has no page selection for fixture ${name}`);
  }
  const requestedPages = pageSubset ? [...pageSubset].sort((a, b) => a - b) : null;
  const pdfPageCount = pageCountFromPdf(pdf);
  rmSync(work, { recursive: true, force: true });
  mkdirSync(work, { recursive: true });
  renderFullReference(pdf, work, requestedPages);

  const page = await browser.newPage({
    // 1400, not 1200: the demo's control bar ends 100.6px down, and an A4
    // page (1122.5px) needs more room than 1200 leaves below it -
    // scrollIntoViewIfNeeded then parks the page top UNDER the bar and
    // el.screenshot() records the bar's chrome as page ink. On a sparse A4
    // page that strip dominates the score (wild-hamburg p14, a heading and
    // a footer, read 30.45% structural from it; Letter pages always fit
    // clear, so their rasters are unchanged by this).
    viewport: { width: 1700, height: 1400 },
    deviceScaleFactor: SCALE,
  });
  // DATE/TIME fields: Word bakes the EXPORT moment into the reference PDF,
  // while our render reads the live clock — the comparison drifts a little
  // further every day. Freeze the page clock to the reference's own
  // CreationDate so both sides evaluate the field at the same instant.
  const created = /CreationDate:\s+(.+)/.exec(execFileSync("pdfinfo", [pdf]).toString())?.[1];
  const refClock = created ? Date.parse(created) : NaN;
  if (Number.isFinite(refClock)) await page.clock.install({ time: refClock });
  await page.goto(`${base}/?doc=/fixtures/${name}.docx&editable=0&comments=0`);
  // attached, not visible: the first span in DOM order can be an empty
  // caret-anchor span (wild-athabasca), which never becomes "visible".
  // Shard workers share one dev server; cold loads under 8-way contention
  // can exceed the interactive default.
  const selectorTimeout = process.env.DXW_PARITY_SHARD_OUT ? 90000 : 20000;
  await page.waitForSelector(".dxw-page span", { timeout: selectorTimeout, state: "attached" });
  await page.waitForFunction(
    () =>
      !document.querySelector(
        ".dxw-body-mode, .dxw-hf-mode, .dxw-comment-hl, .dxw-comment-card, .dxw-sel",
      ),
  );
  await page.waitForTimeout(1200); // fonts + images settle
  const pageCount = await page.locator(".dxw-page").count();
  const needs = await page.evaluate(() => {
    const images = Array.from(document.querySelectorAll('[data-dxw-item-kind="image"]'));
    const hasMetafile = images.some((element) =>
      /^(emf|wmf|svg)$/i.test(element.dataset.dxwImageFormat ?? ""),
    );
    return {
      text: true,
      image: images.length > 0,
      vector: hasMetafile || document.querySelector("[data-dxw-role]") !== null,
    };
  });
  renderSemanticReferences(pdf, work, requestedPages, needs);
  const layerPages = (prefix) => new Map(
    readdirSync(work)
      .map((file) => ({ file, match: file.match(new RegExp(`^${prefix}-(\\d+)\\.png$`)) }))
      .filter(({ match }) => match)
      .map(({ file, match }) => [Number(match[1]), join(work, file)]),
  );
  const wordLayerPages = layerPages("word");
  const textLayerPages = layerPages("text");
  const imageLayerPages = layerPages("image");
  const vectorLayerPages = layerPages("vector");

  const comparedPages = requestedPages ??
    Array.from({ length: Math.max(pageCount, pdfPageCount) }, (_, index) => index + 1);
  for (const pageNumber of comparedPages) {
    const i = pageNumber - 1;
    let webShot = null;
    let semanticShot = null;
    let semanticMetadata = { text: [], images: [], fills: [], rules: [], width: 0, height: 0 };
    if (i < pageCount) {
      const el = page.locator(".dxw-page").nth(i);
      await el.scrollIntoViewIfNeeded();
      // Snap the page to a whole-pixel viewport position. The demo parks
      // .dxw-page at a fractional y (100.609375 at this viewport), and
      // Chromium pixel-snaps composited boxes — the renderer's 1px-high
      // scaled table rules — while glyphs paint unsnapped, so rules land up
      // to a CSS pixel below their layout position relative to the text.
      // el.screenshot() then also rounds the fractional crop origin. Word's
      // raster starts at the exact PDF page origin, so a fractional page
      // origin reads as a rules-vs-text asymmetry the engine does not have
      // (uspto-follow-on p1: 4.02% structural entirely from this).
      await el.evaluate((node) => {
        // Scrolling cannot remove the fraction (scrollTop is integer-quantized
        // here), so shift the pages container by the fractional part instead —
        // a layout change, so everything repaints at the new integral position
        // with no resampling. Page pitch (page height + gap) is integral, so
        // one adjustment snaps every page at once.
        const wrap = node.parentElement;
        for (let pass = 0; pass < 3; pass++) {
          const r = node.getBoundingClientRect();
          const fx = r.x - Math.round(r.x);
          const fy = r.y - Math.round(r.y);
          if (Math.abs(fx) < 0.01 && Math.abs(fy) < 0.01) return;
          wrap.style.marginTop = `${parseFloat(wrap.style.marginTop || "0") - fy}px`;
          wrap.style.marginLeft = `${parseFloat(wrap.style.marginLeft || "0") - fx}px`;
        }
      });
      await page.waitForTimeout(100);
      webShot = await el.screenshot();
      const semantic = await captureSemanticPage(page, i);
      semanticShot = semantic.png;
      semanticMetadata = semantic.metadata;
    }
    const readLayerPage = (pages) => {
      const file = pages.get(pageNumber);
      return file ? readFileSync(file) : null;
    };
    const wordPng = pageNumber <= pdfPageCount ? readLayerPage(wordLayerPages) : null;
    const wordTextPng = needs.text && pageNumber <= pdfPageCount
      ? readLayerPage(textLayerPages)
      : null;
    const wordImagePng = needs.image && pageNumber <= pdfPageCount
      ? readLayerPage(imageLayerPages)
      : null;
    const wordVectorPng = needs.vector && pageNumber <= pdfPageCount
      ? readLayerPage(vectorLayerPages)
      : null;
    const pageStatus = wordPng && webShot ? "matched" : wordPng ? "missing-web" : "extra-web";

    // Compose Word | web | diff in a scratch page (no image deps needed).
    const compare = await browser.newPage({ viewport: { width: 100, height: 100 } });
    const result = await compare.evaluate(
      pageMetric,
      [
        wordPng?.toString("base64") ?? null,
        webShot?.toString("base64") ?? null,
        wordTextPng?.toString("base64") ?? null,
        wordImagePng?.toString("base64") ?? null,
        wordVectorPng?.toString("base64") ?? null,
        semanticShot?.toString("base64") ?? null,
        pageStatus,
        semanticMetadata,
      ],
    );
    await compare.close();

    const candidatePrefix = outcome === "candidate" ? `candidate-${runSlug ? `${runSlug}-` : ""}` : "";
    const pngRel = `${candidatePrefix}${name}-p${pageNumber}.png`;
    const outFile = join(outDir, pngRel);
    if (process.env.DXW_LINE_DEBUG && result.lineDebugRows?.length) {
      for (const r of result.lineDebugRows) console.log("[LDBG]", r);
    }
    writeFileSync(outFile, Buffer.from(result.png, "base64"));
    results.push({
      fixture: name,
      page: pageNumber,
      mismatchPct: Number(result.mismatchPct),
      severityPct: Number(result.severityPct),
      lineShiftPct: Number(result.lineShiftPct),
      alignPx: Number(result.alignPx),
      alignP95: Number(result.alignP95),
      misalignedPct: Number(result.misalignedPct),
      appearanceWeightRatio:
        result.appearanceWeightRatio == null ? null : Number(result.appearanceWeightRatio),
      appearanceWeightErrorPct:
        result.appearanceWeightErrorPct == null ? null : Number(result.appearanceWeightErrorPct),
      appearanceColorDeltaE:
        result.appearanceColorDeltaE == null ? null : Number(result.appearanceColorDeltaE),
      textWeightRatio: result.textWeightRatio,
      textWeightErrorPct: result.textWeightErrorPct,
      textWeightMass: result.textWeightMass,
      textColorDeltaE: result.textColorDeltaE,
      textColorMass: result.textColorMass,
      textColorCoveragePct: result.textColorCoveragePct,
      imageWeightRatio: result.imageWeightRatio,
      imageWeightErrorPct: result.imageWeightErrorPct,
      imageWeightMass: result.imageWeightMass,
      imageItemCount: result.imageItemCount,
      tableFillWeightRatio: result.tableFillWeightRatio,
      tableFillWeightErrorPct: result.tableFillWeightErrorPct,
      tableFillWeightMass: result.tableFillWeightMass,
      tableRuleWeightRatio: result.tableRuleWeightRatio,
      tableRuleWeightErrorPct: result.tableRuleWeightErrorPct,
      tableRuleWeightMass: result.tableRuleWeightMass,
      tableRuleCount: result.tableRuleCount,
      categoryMetricStatus: result.categoryMetricStatus,
      categoryMetricMs: Number(result.categoryMetricMs),
      appearanceMetricMs: Number(result.appearanceMetricMs),
      driftClass: result.driftClass,
      pageStatus: result.pageStatus,
      inkTiles: result.inkTiles,
      pngRel,
    });
    console.log(
      `${name} page ${pageNumber}: ${result.severityPct}% structural (${result.pageStatus}, ${result.driftClass}, ` +
        `line ${result.lineShiftPct}%, align ${result.alignPx}px, ` +
        `weight ${result.appearanceWeightErrorPct == null ? "n/a" : result.appearanceWeightErrorPct + "%"}, ` +
        `colour ${result.appearanceColorDeltaE == null ? "n/a" : result.appearanceColorDeltaE + " ΔE00"}, ` +
        `semantic text ${result.textWeightErrorPct ?? "n/a"}%/${result.textColorDeltaE ?? "n/a"} ΔE00, ` +
        `image ${result.imageWeightErrorPct ?? "n/a"}%, ` +
        `table fill/rule ${result.tableFillWeightErrorPct ?? "n/a"}%/${result.tableRuleWeightErrorPct ?? "n/a"}%, ` +
        `semantic status ${result.categoryMetricStatus}, ` +
        `appearance ${result.appearanceMetricMs}ms, ` +
        `${result.mismatchPct}% raw) -> ${outFile}`,
    );
  }
  await page.close();
  rmSync(work, { recursive: true, force: true });
}

// --- Report generation ---------------------------------------------------
// A run that filtered by fixture name is partial. Every run is persisted, but
// only full runs feed the corpus trend and the carry-forward baseline.
let gitSha = null;
try {
  gitSha = execFileSync("git", ["rev-parse", "--short", "HEAD"], { cwd: root }).toString().trim();
} catch {
  gitSha = null;
}

// Authoring provenance: parity truth is ALWAYS Word's render of the file,
// but fixtures authored by other suites (LibreOffice, eventually Google
// Docs) exercise their idiosyncratic markup. Tagging keeps their drift a
// separate column so the Word-authored numbers stay clean - Word-authored
// parity is the priority axis.
const provenanceOf = (fixture) =>
  /^probe3-lo-/.test(fixture) ? "libreoffice" : "word";

const historyFile = join(parityDir, "history.jsonl");
const readHistory = () => {
  if (!existsSync(historyFile)) return [];
  return readFileSync(historyFile, "utf8")
    .split("\n")
    .filter((l) => l.trim())
    .map((l) => {
      try {
        return JSON.parse(l);
      } catch {
        return null; // skip corrupt lines rather than fail the run
      }
    })
    // Also drop schema-partial entries (valid JSON but no results array) so the
    // report math never sees an undefined results list.
    .filter((e) => e && Array.isArray(e.results));
};

const generatedAt = new Date().toISOString();
const resultMeta = {
  generatedAt,
  gitSha,
  base,
  wordinweb,
  metricVersion: METRIC_VERSION,
  appearanceMetricVersion: APPEARANCE_METRIC_VERSION,
  isFullRun,
  outcome,
  label: runLabel,
  refreshed: isFullRun ? null : refs,
};

// Report generation is best-effort: a failure here must neither fail the parity
// run nor leak the browser, so everything runs inside try/finally.
const reportFile = join(outDir, "report.html");
const reportPng = join(outDir, "report.png");
const candidateSuffix = runSlug ? `-${runSlug}` : "";
const candidateFile = join(outDir, `candidate-results${candidateSuffix}.json`);
const candidateReportFile = join(outDir, `candidate-report${candidateSuffix}.html`);
const candidateReportPng = join(outDir, `candidate-report${candidateSuffix}.png`);
for (const r of results) r.provenance = provenanceOf(r.fixture);

// Shard mode (scripts/parity-parallel.mjs): write this worker's raw results
// to the given path and stop - the orchestrator owns merging, results.json,
// and the history entry.
if (process.env.DXW_PARITY_SHARD_OUT) {
  writeFileSync(
    process.env.DXW_PARITY_SHARD_OUT,
    JSON.stringify({ ...resultMeta, results }, null, 2),
  );
  await browser.close();
  process.exit(0);
}

try {
  const history = readHistory();
  const entry = {
    ts: generatedAt,
    gitSha,
    wordinweb,
    metricVersion: METRIC_VERSION,
    appearanceMetricVersion: APPEARANCE_METRIC_VERSION,
    isFullRun,
    outcome,
    label: runLabel,
    refreshed: isFullRun ? null : refs,
    results: results.map((r) => ({
      fixture: r.fixture,
      page: r.page,
      provenance: provenanceOf(r.fixture),
      mismatchPct: r.mismatchPct,
      severityPct: r.severityPct,
      lineShiftPct: r.lineShiftPct,
      alignPx: r.alignPx,
      alignP95: r.alignP95,
      misalignedPct: r.misalignedPct,
      appearanceWeightRatio: r.appearanceWeightRatio,
      appearanceWeightErrorPct: r.appearanceWeightErrorPct,
      appearanceColorDeltaE: r.appearanceColorDeltaE,
      textWeightRatio: r.textWeightRatio,
      textWeightErrorPct: r.textWeightErrorPct,
      textWeightMass: r.textWeightMass,
      textColorDeltaE: r.textColorDeltaE,
      textColorMass: r.textColorMass,
      textColorCoveragePct: r.textColorCoveragePct,
      imageWeightRatio: r.imageWeightRatio,
      imageWeightErrorPct: r.imageWeightErrorPct,
      imageWeightMass: r.imageWeightMass,
      imageItemCount: r.imageItemCount,
      tableFillWeightRatio: r.tableFillWeightRatio,
      tableFillWeightErrorPct: r.tableFillWeightErrorPct,
      tableFillWeightMass: r.tableFillWeightMass,
      tableRuleWeightRatio: r.tableRuleWeightRatio,
      tableRuleWeightErrorPct: r.tableRuleWeightErrorPct,
      tableRuleWeightMass: r.tableRuleWeightMass,
      tableRuleCount: r.tableRuleCount,
      categoryMetricStatus: r.categoryMetricStatus,
      categoryMetricMs: r.categoryMetricMs,
      appearanceMetricMs: r.appearanceMetricMs,
      driftClass: r.driftClass,
      pageStatus: r.pageStatus,
      pngRel: r.pngRel,
    })),
  };
  appendFileSync(historyFile, JSON.stringify(entry) + "\n");
  history.push(entry);
  // TODO: a second results source (compare-linebreaks.mjs) could merge in here.

  if (outcome === "candidate") {
    writeFileSync(
      candidateFile,
      JSON.stringify({ ...resultMeta, results }, null, 2),
    );
    writeFileSync(
      candidateReportFile,
      buildReport(results, history, {
        generatedAt,
        gitSha,
        base,
        wordinweb,
        isFullRun,
        outcome,
        label: runLabel,
        refreshed: refs,
        appearanceMetricVersion: APPEARANCE_METRIC_VERSION,
        interop,
      }),
    );
    const shot = await browser.newPage({
      viewport: { width: 1200, height: 900 },
      colorScheme: "light",
    });
    await shot.goto(pathToFileURL(candidateReportFile).href);
    await shot.waitForTimeout(300);
    await shot.screenshot({ path: candidateReportPng, fullPage: true });
    await shot.close();
    console.log(`\nCandidate results: ${candidateFile}`);
    console.log(`Candidate report: ${candidateReportFile}`);
    console.log("Accepted results.json/report.html/report.png left unchanged.");
  } else {

    // Build the accepted dashboard from the last compatible full run plus each
    // later explicitly accepted partial. Candidate experiments are persisted to
    // candidate-results.json and history without changing the accepted dashboard.
    const compatible = history.filter((run) => run.metricVersion === METRIC_VERSION);
    let fullIndex = -1;
    for (let i = compatible.length - 1; i >= 0; i--) {
      if (compatible[i].isFullRun !== false) {
        fullIndex = i;
        break;
      }
    }
    let reportResults = fullIndex >= 0 ? [...compatible[fullIndex].results] : [...results];
    for (let i = fullIndex + 1; i < compatible.length; i++) {
      const run = compatible[i];
      if (run.isFullRun !== false) {
        reportResults = [...run.results];
        continue;
      }
      if (run.outcome !== "accepted") continue;
      const rerun = new Set(run.results.map((r) => r.fixture));
      reportResults = [
        ...reportResults.filter((r) => !rerun.has(r.fixture)),
        ...run.results,
      ];
    }
    reportResults = reportResults.map((r) => ({
      ...r,
      pngRel: r.pngRel ?? `${r.fixture}-p${r.page}.png`,
    }));

  // results.json mirrors the complete dashboard state. For a partial run the
  // freshly measured subset is also retained separately, while the main array
  // carries forward untouched fixtures from the last compatible full run.
    writeFileSync(
      join(outDir, "results.json"),
      JSON.stringify(
        {
          ...resultMeta,
          results: reportResults,
          ...(isFullRun ? {} : { runResults: results }),
        },
        null,
        2,
      ),
    );

    const html = buildReport(reportResults, history, {
      generatedAt,
      gitSha,
      base,
      wordinweb,
      isFullRun,
      outcome,
      label: runLabel,
      refreshed: isFullRun ? null : refs,
      appearanceMetricVersion: APPEARANCE_METRIC_VERSION,
      interop,
    });
    writeFileSync(reportFile, html);
    console.log(`\nReport: ${reportFile}`);

    try {
      const shot = await browser.newPage({
        viewport: { width: 1200, height: 900 },
        colorScheme: "light",
      });
      await shot.goto(pathToFileURL(reportFile).href);
      await shot.waitForTimeout(300);
      await shot.screenshot({ path: reportPng, fullPage: true });
      await shot.close();
      console.log(`Report PNG: ${reportPng}`);
    } catch (err) {
      console.warn(`Report screenshot skipped: ${err.message}`);
    }
  }
} catch (err) {
  console.warn(`Report generation skipped: ${err.message}`);
} finally {
  await browser.close();
}

if (outcome === "accepted") {
  execFileSync(process.execPath, [join(root, "scripts", "snapshot-report.mjs")], {
    cwd: root,
    stdio: "inherit",
  });
}
