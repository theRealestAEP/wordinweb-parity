#!/usr/bin/env node

/**
 * Round 2 of probe-repeathdr: the us-courts-answer HEADER STACK, verbatim.
 * probe-repeathdr pinned a one-row repeated header's continuation top at the
 * true outer-top arithmetic, but the fixture repeats TWO rows — a content
 * caption row whose cell bottom is a DOUBLE sz-8 border, then an exact-144
 * all-nil row — and our continuation pages open 2.7px short of Word's
 * against that stack (uniform across pages 2-7, absorbed per page by the
 * metric's global offset but deciding every continuation-page row fit).
 *
 * Each case: the fixture's rows 0-1 verbatim as the tblHeader stack, then
 * 75 one-line data rows, in the fixture's own package (compat 11) with its
 * tblPr. One authored thing varies:
 *
 *   W0  verbatim
 *   W1  row0's double bottom -> nil        the double boundary's charge
 *   W2  row1's exact-144 removed           the exact row's part
 *   W3  row0/row1 tcMar top/bottom -> 0    the margins' part
 *
 * Observable per case: on the case's SECOND page, the repeated row0 text
 * top and the first data mark top.
 *
 *   node scripts/generate-repeathdr2-probe.mjs
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

const dataRow = (text) =>
  "<w:tr>" +
  `<w:tc><w:tcPr><w:tcW w:w="10800" w:type="dxa"/><w:gridSpan w:val="20"/>` +
  '<w:tcBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/></w:tcBorders>' +
  '<w:tcMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:left w:w="58" w:type="dxa"/><w:right w:w="58" w:type="dxa"/></w:tcMar></w:tcPr>' +
  mark(text) +
  "</w:tc></w:tr>";

const zeroMar = (r) =>
  r.replace(/<w:tcMar><w:top w:w="\d+" w:type="dxa"\/>/g, '<w:tcMar><w:top w:w="0" w:type="dxa"/>')
   .replace(/<w:bottom w:w="\d+" w:type="dxa"\/><w:right/g, '<w:bottom w:w="0" w:type="dxa"/><w:right');

const CASES = [
  { id: "W0", r0: row0, r1: row1 },
  { id: "W1", r0: row0.replace('<w:bottom w:val="double" w:sz="8" w:space="0" w:color="000000"/>', '<w:bottom w:val="nil"/>'), r1: row1 },
  { id: "W2", r0: row0, r1: row1.replace('<w:trHeight w:hRule="exact" w:val="144"/>', "") },
  { id: "W3", r0: zeroMar(row0), r1: zeroMar(row1) },
];
for (const c of CASES.slice(1)) {
  if (c.r0 === row0 && c.r1 === row1) throw new Error(`case ${c.id} stripped nothing`);
}

const body =
  CASES.map((c, i) => {
    let data = "";
    for (let d = 1; d <= 75; d++) data += dataRow(`${c.id}D${String(d).padStart(2, "0")}`);
    const tbl = "<w:tbl>" + tblPr + grid + c.r0 + c.r1 + data + "</w:tbl>";
    const brk = i < CASES.length - 1 ? `<w:p><w:pPr><w:sectPr>${SECT}</w:sectPr></w:pPr></w:p>` : "";
    return mark(`${c.id}REF`) + tbl + brk;
  }).join("") + `<w:sectPr>${SECT}</w:sectPr>`;

const docOpen = srcDoc.slice(srcDoc.indexOf("<w:document"), srcDoc.indexOf(">", srcDoc.indexOf("<w:document")) + 1);
parts["word/document.xml"] = strToU8(
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + docOpen + `<w:body>${body}</w:body></w:document>`,
);
const out = join(root, "fixtures-staging/probe-repeathdr2-11.docx");
writeFileSync(out, Buffer.from(zipSync(parts)));
console.log(`wrote ${out} — ${CASES.map((c) => c.id).join(", ")}`);
