#!/usr/bin/env node

/**
 * Generate fixtures-staging/probe-legendedge.docx — the per-type right-edge
 * sweep #84 asks for.
 *
 * #81 measured probe-charts-basic and found the plot-right-to-legend-key gap
 * is a different number on every page (column 21.33, bar 28.46, area 29.77,
 * line 38.05 chart-local px), so `LEGEND_GAP = 16` generalises one page and
 * the rule that actually sets a side-legend chart's plot right edge per chart
 * type is unestablished. One chart box cannot separate the candidate shapes —
 * a fixed inset from the chart's right edge, a fixed gap left of the legend,
 * and a fraction of the box width all predict the same single measurement.
 *
 * Two box sizes separate them: a fixed quantity stays put in pt while a
 * fractional one scales with the box. So: bar, line and pie, each at the
 * default 360x216 pt box and at 240x144 pt, one chart per page, same data and
 * the same right-side legend throughout (stacks, not sandwiches — within a
 * type only the box changes; across types at one size only the type changes).
 *
 * Charts are authored through the engine's own insertChartAt like
 * probe-charts-basic, then the wp:extent of the small cases is rewritten —
 * the ChartML itself never states a size, so the frame extent is the whole
 * variable. Read each page with fitz get_drawings(): plot rect from the white
 * plot fill and its gridlines, legend key and text from the swatch rects and
 * text spans, all in chart-local px.
 *
 *   node scripts/generate-legendedge-probe.mjs
 */

import { existsSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { wordinwebBuild } from "./engine-provenance.mjs";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const fixtures = join(root, "fixtures-staging");
const corpus = join(root, "apps/demo/public/fixtures");

const build = wordinwebBuild();
const core = join(build.target ?? build.packageDir, "../core/dist");
if (!existsSync(join(core, "edit/charts.js"))) {
  throw new Error(`No core build beside the selected engine at ${core} — select a local build with scripts/use-engine.mjs`);
}
const { DocxDocument } = await import(pathToFileURL(join(core, "index.js")).href);
const { insertChartAt } = await import(pathToFileURL(join(core, "edit/charts.js")).href);

const EMU_PER_PX = 9525;
const CATEGORIES = ["Q1", "Q2", "Q3", "Q4"];
const SERIES = [
  { name: "Alpha", values: [4.2, 7.8, 3.1, 9.6] },
  { name: "Beta", values: [6.5, 2.4, 8.9, 5.3] },
];

// L is insertChartAt's own 480x288 px (360x216 pt); S is 2/3 of it. The 2/3
// keeps the aspect, so a fractional inset scales by exactly 2/3 in pt while a
// fixed one does not move.
const SIZES = { L: [480, 288], S: [320, 192] };

const CASES = [];
for (const type of ["bar", "line", "pie"]) {
  for (const size of ["L", "S"]) {
    CASES.push({
      type,
      size,
      title: `${type} ${size}`,
      grouping: type === "bar" ? "clustered" : undefined,
      series: type === "pie" ? SERIES.slice(0, 1) : SERIES,
    });
  }
}

const el = (name, attrs = {}, children = [], text = "") => ({ name, attrs, children, text });
const caption = (text) =>
  el("w:p", {}, [el("w:r", {}, [el("w:t", { "xml:space": "preserve" }, [], text)])]);
const pageBreak = () => el("w:p", {}, [el("w:r", {}, [el("w:br", { "w:type": "page" })])]);

const doc = DocxDocument.load(new Uint8Array(readFileSync(join(corpus, "parity-text.docx"))));
const body = doc.docRoot.children.find((c) => c.name.endsWith("body"));
const sectPr = body.children.find((c) => c.name.endsWith("sectPr"));

const captions = [];
body.children = [];
CASES.forEach((probe, index) => {
  if (index > 0) body.children.push(pageBreak());
  const p = caption(`${index + 1}. ${probe.title}`);
  captions.push(p.children[0].children[0]);
  body.children.push(p);
});
if (sectPr) body.children.push(sectPr);
doc.refresh();

CASES.forEach((probe, index) => {
  const drawing = insertChartAt(doc, captions[index], {
    type: probe.type,
    ...(probe.grouping ? { grouping: probe.grouping } : {}),
    title: probe.title,
    categories: CATEGORIES,
    series: probe.series,
  });
  if (!drawing) throw new Error(`insertChartAt refused chart ${index + 1} (${probe.title})`);
  const [px, py] = SIZES[probe.size];
  const extent = drawing.children[0].children.find((c) => c.name === "wp:extent");
  extent.attrs.cx = String(Math.round(px * EMU_PER_PX));
  extent.attrs.cy = String(Math.round(py * EMU_PER_PX));
});

const out = join(fixtures, "probe-legendedge.docx");
writeFileSync(out, doc.save());
console.log(`Wrote ${out} — ${CASES.length} charts: ${CASES.map((c) => c.title).join(", ")}`);
