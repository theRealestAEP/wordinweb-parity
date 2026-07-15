#!/usr/bin/env node
/**
 * Parallel parity runner: shards the fixture list across N worker processes
 * (each a normal scripts/parity-compare.mjs run in DXW_PARITY_SHARD_OUT mode),
 * then merges the shards into the accepted results.json and appends ONE
 * full-run history entry — the same artifacts a serial full run produces,
 * minus report.html (run the serial script when the dashboard is wanted).
 *
 * Shards are balanced by reference page count (LPT), so the NIH contract
 * anchors one worker while the small fixtures pack the others. Reference
 * rasters come from parity/.raster-cache after the first run, so workers
 * spend their time on rendering + comparison only.
 *
 *   node scripts/parity-parallel.mjs            # all fixtures
 *   DXW_PARITY_JOBS=8 node scripts/parity-parallel.mjs
 *
 * DXW_PARITY_FAST is passed through. Scores were validated shard-vs-serial
 * (identical per-page severities) before this became the default path.
 */
import { execFileSync, spawn } from "node:child_process";
import { existsSync, appendFileSync, mkdirSync, readFileSync, readdirSync, writeFileSync } from "node:fs";
import { cpus } from "node:os";
import { basename, dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const root = dirname(dirname(fileURLToPath(import.meta.url)));
const parityDir = join(root, "parity");
const fixturesDir = join(root, "apps/demo/public/fixtures");
const outDir = join(parityDir, "out");
mkdirSync(outDir, { recursive: true });

const requested = process.argv.slice(2);
const refs = readdirSync(parityDir)
  .filter((f) => f.endsWith("-word.pdf"))
  .map((f) => basename(f, "-word.pdf"))
  .filter((name) => existsSync(join(fixturesDir, `${name}.docx`)))
  .filter((name) => requested.length === 0 || requested.includes(name))
  .sort();
if (refs.length === 0) {
  console.error("No matching parity references.");
  process.exit(1);
}

const pageCount = (name) => {
  const info = execFileSync("pdfinfo", [join(parityDir, `${name}-word.pdf`)]).toString();
  return Number(/Pages:\s+(\d+)/.exec(info)?.[1] ?? 1);
};

const jobs = Math.max(1, Number(process.env.DXW_PARITY_JOBS) || Math.min(6, cpus().length - 2));
// Big fixtures split into page-range chunks so one 400-page doc doesn't
// serialize the whole run; each chunk re-lays the doc but only compares its
// range. Chunk target keeps per-worker layout overhead ~10% of compare time.
const CHUNK = Math.max(40, Number(process.env.DXW_PARITY_CHUNK) || 80);
const sized = refs.map((name) => ({ name, pages: pageCount(name) }));
const units = [];
for (const f of sized) {
  if (f.pages <= CHUNK * 1.5) {
    units.push({ name: f.name, pages: f.pages, range: null });
  } else {
    const parts = Math.ceil(f.pages / CHUNK);
    const per = Math.ceil(f.pages / parts);
    for (let start = 1; start <= f.pages; start += per) {
      const end = Math.min(f.pages, start + per - 1);
      units.push({ name: f.name, pages: end - start + 1, range: [start, end] });
    }
  }
}
units.sort((a, b) => b.pages - a.pages);
// LPT bin packing: biggest units first onto the least-loaded shard.
const shards = Array.from({ length: Math.min(jobs, units.length) }, () => ({ pages: 0, units: [] }));
for (const u of units) {
  shards.sort((a, b) => a.pages - b.pages);
  shards[0].units.push(u);
  shards[0].pages += u.pages;
}
const totalPages = sized.reduce((a, f) => a + f.pages, 0);
console.log(`${refs.length} fixtures / ${totalPages} pages / ${units.length} units across ${shards.length} workers:`);
for (const [i, sh] of shards.entries())
  console.log(`  worker ${i}: ${sh.pages}p  ${sh.units.map((u) => u.range ? `${u.name}[${u.range[0]}-${u.range[1]}]` : u.name).join(" ")}`);

const started = Date.now();
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const runShard = (sh, i, attempt) => {
  const outFile = join(outDir, `.shard-${i}.json`);
  const names = [...new Set(sh.units.map((u) => u.name))];
  const pageSel = sh.units
    .flatMap((u) => {
      if (!u.range) return [];
      const out = [];
      for (let p = u.range[0]; p <= u.range[1]; p++) out.push(`${u.name}:${p}`);
      return out;
    })
    .join(",");
  // A shard mixing whole fixtures and ranged chunks must enumerate the whole
  // fixtures' pages too - DXW_PARITY_PAGES is all-or-nothing per run.
  const env = { ...process.env, DXW_PARITY_SHARD_OUT: outFile };
  if (pageSel) {
    const wholeSel = sh.units
      .filter((u) => !u.range)
      .flatMap((u) => Array.from({ length: u.pages }, (_, k) => `${u.name}:${k + 1}`))
      .join(",");
    env.DXW_PARITY_PAGES = wholeSel ? `${pageSel},${wholeSel}` : pageSel;
  }
  const child = spawn("node", [join(root, "scripts/parity-compare.mjs"), ...names], {
    cwd: root,
    env,
    stdio: ["ignore", "pipe", "pipe"],
  });
  const prefix = (data) => {
    for (const line of data.toString().split("\n")) {
      if (line.trim()) console.log(`[w${i}] ${line}`);
    }
  };
  child.stdout.on("data", prefix);
  child.stderr.on("data", prefix);
  return new Promise((resolve, reject) => {
    child.on("exit", (code) =>
      code === 0 ? resolve(outFile) : reject(new Error(`worker ${i} exited ${code}`)),
    );
  });
};
const runs = shards.map(async (sh, i) => {
  await sleep(i * 1500); // stagger cold loads on the shared dev server
  try {
    return await runShard(sh, i, 1);
  } catch (err) {
    console.log(`[w${i}] failed (${err.message}) - retrying once`);
    return runShard(sh, i, 2);
  }
});

const shardFiles = await Promise.all(runs);
const shardData = shardFiles.map((f) => JSON.parse(readFileSync(f, "utf8")));
const results = shardData
  .flatMap((d) => d.results)
  .sort((a, b) => (a.fixture === b.fixture ? a.page - b.page : a.fixture < b.fixture ? -1 : 1))
  .map((r) => ({ ...r, pngRel: r.pngRel ?? `${r.fixture}-p${r.page}.png` }));
const meta = shardData[0];
const resultMeta = {
  generatedAt: new Date().toISOString(),
  gitSha: meta.gitSha,
  base: meta.base,
  metricVersion: meta.metricVersion,
  appearanceMetricVersion: meta.appearanceMetricVersion,
  isFullRun: requested.length === 0,
  outcome: "accepted",
  label: meta.label ?? null,
  refreshed: requested.length === 0 ? null : refs,
};
// Partial runs merge into the existing accepted dashboard instead of
// clobbering it (the serial script achieves this via history replay).
let dashboardResults = results;
const resultsPath = join(outDir, "results.json");
if (!resultMeta.isFullRun && existsSync(resultsPath)) {
  try {
    const prev = JSON.parse(readFileSync(resultsPath, "utf8"));
    const rerun = new Set(results.map((r) => r.fixture));
    dashboardResults = [
      ...(prev.results ?? []).filter((r) => !rerun.has(r.fixture)),
      ...results,
    ].sort((a, b) => (a.fixture === b.fixture ? a.page - b.page : a.fixture < b.fixture ? -1 : 1));
  } catch {
    // unreadable previous dashboard - fall back to just this run
  }
}
writeFileSync(resultsPath, JSON.stringify({ ...resultMeta, results: dashboardResults }, null, 2));
appendFileSync(
  join(parityDir, "history.jsonl"),
  JSON.stringify({
    ts: resultMeta.generatedAt,
    gitSha: resultMeta.gitSha,
    metricVersion: resultMeta.metricVersion,
    appearanceMetricVersion: resultMeta.appearanceMetricVersion,
    isFullRun: resultMeta.isFullRun,
    outcome: "accepted",
    label: resultMeta.label,
    refreshed: resultMeta.refreshed,
    results: results.map((r) => ({
      fixture: r.fixture,
      page: r.page,
      provenance: r.provenance,
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
  }) + "\n",
);
// Regenerate the HTML dashboard from the just-written results (parallel mode
// otherwise leaves report.html stale — the serial script builds it inline).
try {
  execFileSync("node", [join(root, "scripts/parity-render-report.mjs")], { cwd: root, stdio: "inherit" });
} catch (err) {
  console.warn(`report.html regen skipped: ${err.message}`);
}

const byProv = {};
for (const r of results) (byProv[r.provenance ?? "word"] ??= []).push(Number(r.severityPct) || 0);
if (Object.keys(byProv).length > 1) {
  for (const [prov, sevs] of Object.entries(byProv)) {
    console.log(
      `  [${prov}] ${sevs.length} pages, mean ${(sevs.reduce((a, b) => a + b, 0) / sevs.length).toFixed(3)}%, worst ${Math.max(...sevs).toFixed(2)}%`,
    );
  }
}
const secs = ((Date.now() - started) / 1000).toFixed(0);
const sev = results.map((r) => Number(r.severityPct) || 0);
const hot = results.filter((r) => (Number(r.severityPct) || 0) >= 1);
console.log(`\n${results.length} pages in ${secs}s — mean ${(sev.reduce((a, b) => a + b, 0) / sev.length).toFixed(3)}%, pages ≥1%: ${hot.length}`);
for (const h of hot) console.log(`  ${h.fixture} p${h.page}: ${h.severityPct}`);
console.log(`results.json + history updated (no report.html in parallel mode).`);
