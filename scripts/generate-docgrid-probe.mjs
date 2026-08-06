#!/usr/bin/env node

/**
 * Generate the `w:docGrid` sweep that #62 asked for.
 *
 * wild2-math-eq-as-images stops filling page 7 at 938.36 CSS px where Word runs
 * to 1022.33, and pushes a two-line paragraph to page 8 that Word keeps. The
 * standing suspect was that section's `<w:docGrid w:type="lines"
 * w:linePitch="312"/>`, on the theory that we ignore a line pitch Word snaps
 * to.
 *
 * THAT THEORY IS ALREADY DEAD, and this probe is built to test what is left of
 * it. Measured against the reference PDF, our line pitch on that page is
 * IDENTICAL to Word's — 26.00 px through the body text and 41.67 px through the
 * paragraph that spills — so the grid is not moving our lines. What differs is
 * only WHERE THE PAGE ENDS: the paragraph's second line would land at
 * 1000.72..1021.39, inside the nominal body bottom of 1026.53, and we reject it
 * anyway. Engine commit 355be56 recorded the same shape on
 * wild2-legal-ca-agreement and deliberately left it alone: "the ordinary test's
 * effective bottom for a line can sit ~14px above the nominal 960".
 *
 * So the question this sweeps is not the line pitch but the PAGE BOTTOM: does a
 * docGrid change the last line a page will accept? Six sections, each starting
 * `nextPage`, each holding the same 60 single-line paragraphs, differing in one
 * authored thing:
 *
 *   N     no docGrid element at all           (control)
 *   D312  <w:docGrid w:linePitch="312"/>      (w:type omitted, so "default")
 *   L312  w:type="lines"        pitch 312     (the fixture's own shape)
 *   L240  w:type="lines"        pitch 240     two settings of the pitch, to say
 *   L480  w:type="lines"        pitch 480     whether the effect scales with it
 *   C312  w:type="linesAndChars" pitch 312
 *
 * Read PER CASE: the number of the last paragraph on the section's first page,
 * that line's bottom, and the line pitch. The control fixes what this geometry
 * does with no grid; a case that ends earlier than the control loses room to
 * its grid, and L240 against L480 says whether the loss tracks the pitch.
 *
 * Every paragraph is labelled `<case>L<nn>`, so a case reads the same way from
 * a Word PDF, from a raster, and from the browser.
 *
 *   node scripts/generate-docgrid-probe.mjs
 */

import { readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { strToU8, unzipSync, zipSync } from "fflate";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
// Base package chosen so the styles, fonts and theme are the ones the fixture
// under investigation actually uses; only document.xml is replaced.
const source = join(root, "apps/demo/public/fixtures/wild2-math-eq-as-images.docx");
const parts = unzipSync(new Uint8Array(readFileSync(source)));

/** The fixture's own A4 geometry, minus its header and footer, so the body is
 *  exactly 96.00 .. 1026.53 CSS px and nothing else can reserve into it. */
const pgProps =
  '<w:pgSz w:w="11906" w:h="16838"/>' +
  '<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/>' +
  '<w:cols w:space="720"/>';

const cases = [
  { id: "N", grid: "" },
  { id: "D312", grid: '<w:docGrid w:linePitch="312"/>' },
  { id: "L312", grid: '<w:docGrid w:type="lines" w:linePitch="312"/>' },
  { id: "L240", grid: '<w:docGrid w:type="lines" w:linePitch="240"/>' },
  { id: "L480", grid: '<w:docGrid w:type="lines" w:linePitch="480"/>' },
  { id: "C312", grid: '<w:docGrid w:type="linesAndChars" w:linePitch="312"/>' },
];

/** Enough single lines to overflow one page under every grid in the sweep. */
const LINES = 60;

const sectPr = (grid) => `<w:sectPr><w:type w:val="nextPage"/>${pgProps}${grid}</w:sectPr>`;

/** One line, no authored spacing, so the only vertical quantity in play is the
 *  line height itself and whatever the grid does to it. */
const para = (label, tail = "") =>
  `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>` +
  `<w:jc w:val="left"/>${tail}</w:pPr>` +
  `<w:r><w:t xml:space="preserve">${label} pahiwe nuqaji bepon rukefev kicuziri</w:t></w:r></w:p>`;

const body = cases
  .map(({ id, grid }, ci) => {
    const paras = [];
    for (let i = 1; i <= LINES; i++) {
      const label = `${id}L${String(i).padStart(2, "0")}`;
      // A section's own properties ride the LAST paragraph of that section, so
      // this case's grid goes on this case's last paragraph. The final section
      // has no such paragraph and takes the body-level sectPr instead.
      const last = i === LINES && ci < cases.length - 1;
      paras.push(para(label, last ? sectPr(grid) : ""));
    }
    return paras.join("");
  })
  .join("");

const documentXml =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  `<w:body>${body}<w:sectPr><w:type w:val="nextPage"/>${pgProps}${cases.at(-1).grid}</w:sectPr></w:body>` +
  "</w:document>";

// Only document.xml is replaced. The other parts stay in the package even
// though the new body references none of them, so the relationships and
// content types the base declares all still resolve.
parts["word/document.xml"] = strToU8(documentXml);

const out = join(root, "apps/demo/public/fixtures/probe-docgrid.docx");
writeFileSync(out, Buffer.from(zipSync(parts)));
console.log(`wrote ${out}`);
console.log(`cases: ${cases.map((c) => c.id).join(", ")}  (${LINES} lines each)`);
