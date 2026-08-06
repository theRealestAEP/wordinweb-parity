#!/usr/bin/env node

/**
 * Generate the probe that isolates #38's SECOND step — the one the quarter-line
 * paragraphs were wrongly blamed for.
 *
 * On wild2-legal-ca-agreement page 1 we place `[TIV gece]` 2.33 CSS px below
 * where Word places it, measured from `TIV'W VIQIMUSIM` — both 10 pt runs, so
 * the ascent offset cancels and the number is the flow itself. Everything
 * between them lives in the signature table, whose rows carry
 * `<w:trHeight w:hRule="exact"/>`: 495, 110 and 1089 twips. An exact row is
 * exactly that tall whatever it holds, which is why the two empty quarter-line
 * paragraphs in its first row cannot move anything, and why the residual has to
 * come from the row stack rather than from a paragraph.
 *
 * Two candidates remain, and they differ in one authored byte each:
 *
 *   exact row height   do we honour `hRule="exact"`, or grow a row to its
 *                      content? Row 1 is 110 twips (5.5 pt) and holds a 10 pt
 *                      empty paragraph that wants ~11.5 pt.
 *   shared borders     row 0's bottom and row 1's top are both `sz="12"`
 *                      (1.5 pt). Word resolves an adjacent pair into ONE rule;
 *                      charging both would add 1.5 pt = 2.0 CSS px, which is
 *                      the right size.
 *
 * The probe reproduces the fixture's three rows with a marked 10 pt line in
 * each, then varies ONE thing at a time:
 *
 *   R01  the fixture's shape: exact 495/110/1089, all borders sz=12/8
 *   R02  the same rows with hRule="atLeast"      (exact-height handling alone)
 *   R03  the same rows with NO table borders     (shared-border cost alone)
 *   R04  no borders AND atLeast                  (neither)
 *   R05  exact rows, borders, but row 1 removed  (is the cost per boundary?)
 *   R06  R01 plus the fixture's own per-cell <w:tcPr> — borders and shading
 *   R07  R06 with the shading stripped        (the cell BORDERS alone)
 *   R08  R06 with the borders stripped        (the SHADING alone)
 *
 * Read `top(row 2 mark) - top(row 0 mark)` in each. R01 minus R03 is what the
 * borders cost us; R01 minus R02 is what `exact` costs us; Word's own numbers
 * say what each should cost.
 *
 *   node scripts/generate-exactrow-probe.mjs
 */

import { readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const source = join(root, "apps/demo/public/fixtures/wild2-legal-ca-agreement.docx");
const parts = unzipSync(new Uint8Array(readFileSync(source)));

const SECT =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '<w:cols w:space="720"/></w:sectPr>';

const BORDERS =
  "<w:tblBorders>" +
  ["top", "left", "bottom", "right", "insideH", "insideV"]
    .map((e) => `<w:${e} w:val="single" w:sz="12" w:space="0" w:color="auto"/>`)
    .join("") +
  "</w:tblBorders>";

const tblPr = (borders) =>
  '<w:tblPr><w:tblW w:w="10080" w:type="dxa"/><w:tblInd w:w="-615" w:type="dxa"/>' +
  (borders ? BORDERS : "") +
  '<w:tblLayout w:type="fixed"/>' +
  '<w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>';

/** A 10 pt marked line, as the fixture's row text is. */
const line = (text) =>
  `<w:p><w:pPr><w:rPr><w:b/><w:sz w:val="20"/></w:rPr></w:pPr>` +
  `<w:r><w:rPr><w:b/><w:sz w:val="20"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;
/** The fixture's own empty quarter-line paragraph, mark at 13 pt. */
const quarter =
  '<w:p><w:pPr><w:spacing w:line="60" w:lineRule="auto"/><w:jc w:val="center"/>' +
  '<w:rPr><w:b/><w:sz w:val="26"/></w:rPr></w:pPr></w:p>';
const empty = (sz) => `<w:p><w:pPr><w:rPr><w:sz w:val="${sz}"/></w:rPr></w:pPr></w:p>`;

const cell = (w, body, tcPr) => `<w:tc>${tcPr ?? `<w:tcPr><w:tcW w:w="${w}" w:type="dxa"/></w:tcPr>`}${body}</w:tc>`;
const row = (twips, rule, bodies, tcPrs) =>
  `<w:tr><w:trPr><w:trHeight w:hRule="${rule}" w:val="${twips}"/></w:trPr>` +
  cell(5130, bodies[0], tcPrs?.[0]) +
  cell(4950, bodies[1], tcPrs?.[1]) +
  "</w:tr>";

/**
 * The fixture's own <w:tcPr> for the three rows, lifted verbatim from
 * wild2-legal-ca-agreement's signature table. The probe's own cells carry only
 * a width, so R06 is the variant that adds the per-cell borders and shading
 * back — the one thing the authored variants above leave out.
 */
const fixtureTcPr = (() => {
  const body = strFromU8(parts["word/document.xml"]);
  const start = body.indexOf("<w:tbl>", body.indexOf("<w:body>"));
  const tbl = body.slice(start, body.indexOf("</w:tbl>", start));
  // Table 2 is the signature table; find it by its exact first row height.
  const sig = body.slice(body.indexOf('<w:trHeight w:hRule="exact" w:val="495"/>') - 400);
  const rows = sig.match(/<w:tr[ >][\s\S]*?<\/w:tr>/g) ?? [];
  return rows.slice(0, 3).map((r) => (r.match(/<w:tcPr>[\s\S]*?<\/w:tcPr>/g) ?? []).slice(0, 2));
})();

/** The fixture's three rows, with a mark placed where its text sits. */
function table(id, { borders = true, rule = "exact", middleRow = true, tcPr = false, strip } = {}) {
  const at = (i) => {
    if (!tcPr) return undefined;
    const cells = fixtureTcPr[i];
    if (!strip) return cells;
    return cells.map((x) => x.replace(strip, ""));
  };
  const rows = [
    row(495, rule, [quarter + line(`${id}TOP`) + quarter, quarter + line(`${id}top2`)], at(0)),
    middleRow ? row(110, rule, [empty(20), empty(13)], at(1)) : "",
    row(1089, rule, [empty(14) + line(`${id}BOT`), empty(14) + line(`${id}bot2`)], at(2)),
  ].join("");
  return `<w:tbl>${tblPr(borders)}${rows}</w:tbl>`;
}

const PAGE_BREAK = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';

const variants = [
  ["R01", {}],
  ["R02", { rule: "atLeast" }],
  ["R03", { borders: false }],
  ["R04", { borders: false, rule: "atLeast" }],
  ["R05", { middleRow: false }],
  ["R06", { tcPr: true }],
  ["R07", { tcPr: true, strip: /<w:shd[^/]*\/>/g }],
  ["R08", { tcPr: true, strip: /<w:tcBorders>[\s\S]*?<\/w:tcBorders>/g }],
];

const body = variants
  .map(([id, opts], idx) => (idx ? PAGE_BREAK : "") + table(id, opts) + line(`${id}END`))
  .join("");

const documentXml =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
  '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
  `<w:body>${body}${SECT}</w:body></w:document>`;

const path = join(root, "fixtures-staging", "probe-exactrow.docx");
writeFileSync(path, zipSync({ ...parts, "word/document.xml": strToU8(documentXml) }));
console.log(`Wrote ${path} — ${variants.length} variants, one per page`);
