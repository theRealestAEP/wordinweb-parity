#!/usr/bin/env node

/**
 * Generate the sided-ness variant left open by #51 and the exact-row work.
 *
 * `generate-exactrow-probe.mjs` established that a cell's OWN borders come out
 * of an `hRule="exact"` row's content box where a table's borders do not: Word
 * draws the cell border inside the row and the content loses its width. What it
 * could not say is WHICH SIDE pays, because every case it measured carried the
 * fixture's borders on both edges, where all three models predict the same
 * total. Three models survive on that evidence:
 *
 *   half-share  each edge takes half its own width from the content box, so a
 *               one-sided border shifts the content down by w/2
 *   own-top     the content box loses the full width at the edge the border is
 *               on, so a TOP border shifts the content down by w and a BOTTOM
 *               border shifts it not at all
 *   own-bottom  the content box always loses at the bottom, so neither shifts
 *
 * An exact row's HEIGHT is fixed by definition, so total height cannot separate
 * them — only where the content sits INSIDE the row can. Each case puts a marker
 * paragraph immediately before a single exact row and another marker inside it,
 * and the measurement is `top(MK) - top(REF)`. The reference line is the same in
 * every case, so any difference between cases is exactly the border charged at
 * the TOP of the row.
 *
 *   N   no cell borders                      (baseline)
 *   T   cell TOP border only, sz=12 (1.5pt = 2.00 CSS px)
 *   B   cell BOTTOM border only, sz=12
 *   TB  both, sz=12                          (the case already measured)
 *
 * Predictions, as offsets from N:
 *
 *   model       T       B      TB
 *   half-share  +1.00   +1.00  +2.00
 *   own-top     +2.00    0.00  +2.00
 *   own-bottom   0.00    0.00   0.00
 *
 * T against B is the whole discriminator, and TB is the control that ties the
 * result back to what the exact-row probe already measured.
 *
 *   node scripts/generate-sidedness-probe.mjs
 */

import { readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { strToU8, unzipSync, zipSync } from "fflate";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const source = join(root, "apps/demo/public/fixtures/wild2-legal-ca-agreement.docx");
const parts = unzipSync(new Uint8Array(readFileSync(source)));

const SECT =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '<w:cols w:space="720"/></w:sectPr>';

/** A 10 pt line of exactly known height, used for both markers. */
const mark = (text) =>
  '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/>' +
  '<w:rPr><w:sz w:val="20"/></w:rPr></w:pPr>' +
  `<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;

/**
 * No `w:tblBorders`: a table rule would be charged by a different rule than the
 * cell border under test, and the exact-row probe already separated the two.
 */
const table = (id, tcBorders) =>
  '<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/><w:tblLayout w:type="fixed"/>' +
  '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/>' +
  '<w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid>' +
  '<w:tr><w:trPr><w:trHeight w:hRule="exact" w:val="495"/></w:trPr>' +
  `<w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/>${tcBorders}</w:tcPr>` +
  mark(`${id}MK`) +
  "</w:tc></w:tr></w:tbl>";

const edge = (side) => `<w:${side} w:val="single" w:sz="12" w:space="0" w:color="000000"/>`;
const CASES = [
  { id: "N", b: "" },
  { id: "T", b: `<w:tcBorders>${edge("top")}</w:tcBorders>` },
  { id: "B", b: `<w:tcBorders>${edge("bottom")}</w:tcBorders>` },
  { id: "TB", b: `<w:tcBorders>${edge("top")}${edge("bottom")}</w:tcBorders>` },
];

// A spacer paragraph after each table keeps consecutive tables from merging.
const body = CASES.map((c) => mark(`${c.id}REF`) + table(c.id, c.b) + mark(`${c.id}END`)).join("");

parts["word/document.xml"] = strToU8(
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    `<w:body>${body}${SECT}</w:body></w:document>`,
);

// Probes live in fixtures-staging: parity-parallel.mjs adopts a reference only
// when a DOCX of the same name sits in apps/demo/public/fixtures.
const out = join(root, "fixtures-staging/probe-sidedness.docx");
writeFileSync(out, Buffer.from(zipSync(parts)));
console.log(`wrote ${out} — ${CASES.map((c) => c.id).join(", ")}`);
