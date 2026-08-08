#!/usr/bin/env node

/**
 * Round 3 of probe-repeathdr: what discriminates the repeated exact row's
 * bottom-margin charge. probe-exactpad (read off the us-courts fixture)
 * measured the REPEATED instance of the exact-144 header row charging HALF
 * its bottom cell margin (p1 gap 22.81pt, p2/p7 22.08); probe-repeathdr2
 * measured the continuation stack EQUAL to the first-page stack. The one
 * authored difference: the fixture's first data row carries tcMar top 58tw,
 * the probe's carried 0. This probe varies that pair directly.
 *
 * Each case: the fixture's rows 0-1 verbatim as the tblHeader stack (row1's
 * tcMar bottom varied), then 75 one-line data rows whose tcMar top varies,
 * in the fixture's own package (compat 11) with its tblPr:
 *
 *   F0   follower top 0,  exact bottom 29   (= repeathdr2 W0 control)
 *   F29  follower top 29, exact bottom 29
 *   F58  follower top 58, exact bottom 29   (= the fixture construct)
 *   B58  follower top 58, exact bottom 58   (does the reduction scale?)
 *   B0   follower top 58, exact bottom 0    (nothing to halve)
 *
 * Observable per case: the Vop-top -> first-data-mark gap on the case's
 * first page vs its continuation page.
 *
 *   node scripts/generate-repeathdr3-probe.mjs
 */

import { readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const source = join(root, "apps/demo/public/fixtures/wild3-template-us-courts-answer.docx");

const parts = unzipSync(new Uint8Array(readFileSync(source)));
const srcDoc = strFromU8(parts["word/document.xml"]);
const rows = srcDoc.match(/<w:tr(?: [^>]*)?>[\s\S]*?<\/w:tr>/g);
const row0 = rows[0];
const row1 = rows[1];
if (!row0.includes("tblHeader") || !row1.includes('w:val="144"')) throw new Error("header rows not found");
const tblPr = srcDoc.match(/<w:tblPr>[\s\S]*?<\/w:tblPr>/)[0];
const grid = srcDoc.match(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/)[0];

const SECT =
  '<w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '<w:cols w:space="720"/>';

const mark = (text) =>
  '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/>' +
  '<w:rPr><w:sz w:val="20"/></w:rPr></w:pPr>' +
  `<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;

const dataRow = (text, topTw) =>
  "<w:tr>" +
  `<w:tc><w:tcPr><w:tcW w:w="10800" w:type="dxa"/><w:gridSpan w:val="20"/>` +
  '<w:tcBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/></w:tcBorders>' +
  `<w:tcMar><w:top w:w="${topTw}" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:left w:w="58" w:type="dxa"/><w:right w:w="58" w:type="dxa"/></w:tcMar></w:tcPr>` +
  mark(text) +
  "</w:tc></w:tr>";

const setRow1Bottom = (r, tw) =>
  r.replace(/<w:bottom w:w="\d+" w:type="dxa"\/><w:right/, `<w:bottom w:w="${tw}" w:type="dxa"/><w:right`);
if (setRow1Bottom(row1, 999) === row1) throw new Error("row1 bottom margin not found");

const CASES = [
  { id: "F0", top: 0, bottom: 29 },
  { id: "F29", top: 29, bottom: 29 },
  { id: "F58", top: 58, bottom: 29 },
  { id: "B58", top: 58, bottom: 58 },
  { id: "B0", top: 58, bottom: 0 },
];

const body =
  CASES.map((c, i) => {
    let data = "";
    for (let d = 1; d <= 75; d++) data += dataRow(`${c.id}D${String(d).padStart(2, "0")}`, c.top);
    const tbl = "<w:tbl>" + tblPr + grid + row0 + setRow1Bottom(row1, c.bottom) + data + "</w:tbl>";
    const brk = i < CASES.length - 1 ? `<w:p><w:pPr><w:sectPr>${SECT}</w:sectPr></w:pPr></w:p>` : "";
    return mark(`${c.id}REF`) + tbl + brk;
  }).join("") + `<w:sectPr>${SECT}</w:sectPr>`;

const docOpen = srcDoc.slice(srcDoc.indexOf("<w:document"), srcDoc.indexOf(">", srcDoc.indexOf("<w:document")) + 1);
parts["word/document.xml"] = strToU8(
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + docOpen + `<w:body>${body}</w:body></w:document>`,
);
const out = join(root, "fixtures-staging/probe-repeathdr3-11.docx");
writeFileSync(out, Buffer.from(zipSync(parts)));
console.log(`wrote ${out} — ${CASES.map((c) => c.id).join(", ")}`);
