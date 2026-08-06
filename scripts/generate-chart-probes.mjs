#!/usr/bin/env node

/**
 * Generate fixtures-staging/probe-charts-basic.docx — the five common chart
 * kinds at Word defaults, one per page — and fixtures-staging/
 * probe-charts-autoscale.docx, which varies only the data maximum.
 *
 * The corpus has no chart fixture at all, so three things the renderer does are
 * documented guesses rather than measurements:
 *
 *   1. the automatic value-axis scale (where the axis starts and ends, and the
 *      major unit between labels);
 *   2. the default text sizes (14pt title, 9pt for axis labels, legend entries
 *      and data labels);
 *   3. where labels sit relative to the plot.
 *
 * Every chart here carries the SAME data, so a difference between two pages is
 * a difference between chart kinds and nothing else. The values are deliberately
 * not round — a 9.6 maximum makes the axis heuristic visible, where a 10 would
 * hide it.
 *
 * Charts are authored through the engine's own insertChartAt, so the probe
 * measures the ChartML the product writes rather than ChartML written for the
 * probe. The body is grafted onto parity-text.docx: Word silently refuses a
 * minimal hand-built package, and starting from one it already opens keeps the
 * probe about charts.
 *
 *   node scripts/generate-chart-probes.mjs
 */

import { existsSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { wordinwebBuild } from "./engine-provenance.mjs";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
// Probes live in fixtures-staging, like every other probe with a Word
// reference, so the full-corpus run does not adopt them as tracked fixtures.
// To re-measure: copy one into apps/demo/public/fixtures, run
// `node scripts/parity-compare.mjs <name>`, then remove it again.
const fixtures = join(root, "fixtures-staging");
const corpus = join(root, "apps/demo/public/fixtures");

// wordinweb re-exports DocxDocument but not the chart editors, and it does not
// publish a core subpath — so reach the core build of the SELECTED engine, and
// say so plainly when the selection is a published package with no sibling core.
const build = wordinwebBuild();
const core = join(build.target ?? build.packageDir, "../core/dist");
if (!existsSync(join(core, "edit/charts.js"))) {
  throw new Error(`No core build beside the selected engine at ${core} — select a local build with scripts/use-engine.mjs`);
}
const { DocxDocument } = await import(pathToFileURL(join(core, "index.js")).href);
const { insertChartAt } = await import(pathToFileURL(join(core, "edit/charts.js")).href);

const CATEGORIES = ["Q1", "Q2", "Q3", "Q4"];
const SERIES = [
  { name: "Alpha", values: [4.2, 7.8, 3.1, 9.6] },
  { name: "Beta", values: [6.5, 2.4, 8.9, 5.3] },
];

const CHARTS = [
  { type: "column", grouping: "clustered", title: "Clustered column" },
  { type: "line", title: "Line" },
  // A pie plots one series, so giving it two would only test which one wins.
  { type: "pie", title: "Pie", series: SERIES.slice(0, 1) },
  { type: "bar", grouping: "clustered", title: "Bar" },
  { type: "area", title: "Area" },
];

const W = "w:";
const el = (name, attrs = {}, children = [], text = "") => ({ name, attrs, children, text });
const caption = (text) =>
  el(`${W}p`, {}, [el(`${W}r`, {}, [el(`${W}t`, { "xml:space": "preserve" }, [], text)])]);
const pageBreak = () => el(`${W}p`, {}, [el(`${W}r`, {}, [el(`${W}br`, { "w:type": "page" })])]);

const doc = DocxDocument.load(new Uint8Array(readFileSync(join(corpus, "parity-text.docx"))));
const body = doc.docRoot.children.find((c) => c.name.endsWith("body"));
const sectPr = body.children.find((c) => c.name.endsWith("sectPr"));

// One page per chart: caption, chart, then a break before the next caption.
const captions = [];
body.children = [];
CHARTS.forEach((chart, index) => {
  if (index > 0) body.children.push(pageBreak());
  const p = caption(`${index + 1}. ${chart.title}`);
  captions.push(p.children[0].children[0]);
  body.children.push(p);
});
if (sectPr) body.children.push(sectPr);
doc.refresh();

CHARTS.forEach((chart, index) => {
  const { title, series, ...rest } = chart;
  const inserted = insertChartAt(doc, captions[index], {
    ...rest,
    title,
    categories: CATEGORIES,
    series: series ?? SERIES,
  });
  if (!inserted) throw new Error(`insertChartAt refused chart ${index + 1} (${title})`);
});

const out = join(fixtures, "probe-charts-basic.docx");
writeFileSync(out, doc.save());
console.log(`Wrote ${out} — ${CHARTS.length} charts: ${CHARTS.map((c) => c.type).join(", ")}`);

// ---------------------------------------------------------------------------
// probe-charts-autoscale.docx — what the value axis does with the data it has.
//
// probe-charts-basic shows our axis stopping one major unit below Word's for a
// 9.6 maximum. One dataset cannot say whether Word always adds headroom, only
// rounds up differently, or scales the headroom with the range, so this varies
// the maximum alone: one series, four categories, column charts throughout.
// Read the tick labels Word prints on each page against ours.
// ---------------------------------------------------------------------------

const MAXIMA = [
  { title: "max 9.6", values: [4.2, 7.8, 3.1, 9.6] },
  { title: "max 10 exactly", values: [4.2, 7.8, 3.1, 10] },
  { title: "max 10.4", values: [4.2, 7.8, 3.1, 10.4] },
  { title: "max 3.7", values: [1.2, 2.8, 1.1, 3.7] },
  { title: "max 47", values: [12, 31, 19, 47] },
  { title: "max 0.85", values: [0.2, 0.6, 0.31, 0.85] },
];

const axes = DocxDocument.load(new Uint8Array(readFileSync(join(corpus, "parity-text.docx"))));
const axesBody = axes.docRoot.children.find((c) => c.name.endsWith("body"));
const axesSect = axesBody.children.find((c) => c.name.endsWith("sectPr"));
const axesCaptions = [];
axesBody.children = [];
MAXIMA.forEach((probe, index) => {
  if (index > 0) axesBody.children.push(pageBreak());
  const p = caption(`${index + 1}. ${probe.title}`);
  axesCaptions.push(p.children[0].children[0]);
  axesBody.children.push(p);
});
if (axesSect) axesBody.children.push(axesSect);
axes.refresh();

MAXIMA.forEach((probe, index) => {
  const inserted = insertChartAt(axes, axesCaptions[index], {
    type: "column",
    grouping: "clustered",
    title: probe.title,
    categories: CATEGORIES,
    series: [{ name: "Alpha", values: probe.values }],
  });
  if (!inserted) throw new Error(`insertChartAt refused autoscale chart ${index + 1}`);
});

const axesOut = join(fixtures, "probe-charts-autoscale.docx");
writeFileSync(axesOut, axes.save());
console.log(`Wrote ${axesOut} — ${MAXIMA.length} column charts: ${MAXIMA.map((m) => m.title).join(", ")}`);
