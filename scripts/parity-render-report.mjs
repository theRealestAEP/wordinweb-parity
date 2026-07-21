#!/usr/bin/env node
/**
 * Regenerate parity/out/report.html (+ report.png) from the CURRENT
 * parity/out/results.json and parity/history.jsonl, without re-running any
 * comparison. The serial parity-compare.mjs builds the HTML inline at the end
 * of a run, but parity-parallel.mjs writes results.json only — so after a
 * parallel run the dashboard is stale. This reuses the same buildReport() so
 * the two paths produce identical HTML.
 *
 *   node scripts/parity-render-report.mjs
 */
import { existsSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { chromium } from "@playwright/test";
import { APPEARANCE_METRIC_VERSION, buildReport } from "./parity-report.mjs";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const outDir = join(root, "parity", "out");
const resultsFile = join(outDir, "results.json");
const historyFile = join(root, "parity", "history.jsonl");
const interopFile = join(root, "apps", "demo", "public", "interop", "results.json");
const reportFile = join(outDir, "report.html");
const reportPng = join(outDir, "report.png");

if (!existsSync(resultsFile)) {
  console.error(`No ${resultsFile} — run a parity suite first.`);
  process.exit(1);
}
const doc = JSON.parse(readFileSync(resultsFile, "utf8"));
const results = (doc.results ?? []).map((r) => ({
  ...r,
  pngRel: r.pngRel ?? `${r.fixture}-p${r.page}.png`,
}));
const history = existsSync(historyFile)
  ? readFileSync(historyFile, "utf8")
      .split("\n")
      .filter((l) => l.trim())
      .map((l) => {
        try {
          return JSON.parse(l);
        } catch {
          return null;
        }
      })
      .filter((e) => e && Array.isArray(e.results))
  : [];
const interop = existsSync(interopFile) ? JSON.parse(readFileSync(interopFile, "utf8")) : null;

const html = buildReport(results, history, {
  generatedAt: doc.generatedAt ?? new Date().toISOString(),
  gitSha: doc.gitSha ?? null,
  base: doc.base ?? null,
  isFullRun: doc.isFullRun ?? true,
  outcome: doc.outcome ?? "accepted",
  label: doc.label ?? null,
  refreshed: doc.refreshed ?? null,
  appearanceMetricVersion: doc.appearanceMetricVersion ?? APPEARANCE_METRIC_VERSION,
  interop,
});
writeFileSync(reportFile, html);
console.log(`Report: ${reportFile} (${results.length} pages)`);

const browser = await chromium.launch();
try {
  const shot = await browser.newPage({ viewport: { width: 1200, height: 900 }, colorScheme: "light" });
  await shot.goto(pathToFileURL(reportFile).href);
  await shot.waitForTimeout(300);
  await shot.screenshot({ path: reportPng, fullPage: true });
  console.log(`Report PNG: ${reportPng}`);
} catch (err) {
  console.warn(`Report screenshot skipped: ${err.message}`);
} finally {
  await browser.close();
}
