#!/usr/bin/env node

/**
 * Round 2 of probe-rowfit11. Round 1 pinned a cantSplit row of EXACT-rule
 * lines at zero overhang allowance (move at room 30, stay at 32, for a
 * 32.00px row). us-courts' rows are NATURAL 11pt Times lines, whose trailing
 * leading Word's fit may exempt ("fit uses the font box"), so the allowance
 * for the fixture's rows is still unmeasured. This round sweeps it:
 *
 *   N24..N40  cantSplit row, two NATURAL 11pt lines, room swept 24..40 step 2
 *   M14..M24  cantSplit row, ONE natural 11pt line, room swept 14..24 step 2
 *
 * The one-line sweep separates per-line leading from per-row bottom slack.
 * Word's natural 11pt Times line measures 16.9px (probe-uscourtsblock2 CN),
 * so N's row is ~33.8px and M's ~16.9px.
 *
 * Also, unfloored this time (content rows, no trHeight), the boundary pair
 * that us-courts' spacer bottoms author — a DECLARED cell border facing a
 * declared nil — at a width the raster can see:
 *
 *   D0  two natural rows, no borders anywhere        control
 *   DW  upper cell bottom sz-12, lower cell top nil  charge of (declared, nil)
 *   DI  DW plus tblBorders insideH sz-8              does insideH re-enter?
 *
 *   node scripts/generate-rowfit2-probe.mjs
 */

import { readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const source = join(root, "apps/demo/public/fixtures/wild3-template-us-courts-answer.docx");

const SECT =
  "<w:sectPr>" +
  '<w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '<w:cols w:space="720"/></w:sectPr>';

const exactLine = (text, twips = 240) =>
  `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="${twips}" w:lineRule="exact"/>` +
  '<w:rPr><w:sz w:val="20"/></w:rPr></w:pPr>' +
  `<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;

/** Natural 11pt line — the fixture's own text shape. */
const natLine = (text) =>
  '<w:p><w:pPr><w:spacing w:before="0" w:after="0"/><w:rPr><w:sz w:val="22"/></w:rPr></w:pPr>' +
  `<w:r><w:rPr><w:sz w:val="22"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;

const TBLPR = (borders) =>
  "<w:tblPr>" +
  '<w:tblW w:w="9360" w:type="dxa"/>' +
  borders +
  '<w:tblLayout w:type="fixed"/>' +
  '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/>' +
  '<w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid>';

const fitTable = (id, nLines) =>
  "<w:tbl>" +
  TBLPR("") +
  "<w:tr><w:trPr><w:cantSplit/></w:trPr>" +
  '<w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/></w:tcPr>' +
  Array.from({ length: nLines }, (_, i) => natLine(`${id}T${i + 1}`)).join("") +
  "</w:tc></w:tr></w:tbl>";

const pairTable = (id, upperTc, lowerTc, borders) =>
  "<w:tbl>" +
  TBLPR(borders) +
  `<w:tr><w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/>${upperTc}</w:tcPr>` +
  natLine(`${id}UP`) +
  "</w:tc></w:tr>" +
  `<w:tr><w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/>${lowerTc}</w:tcPr>` +
  natLine(`${id}MK`) +
  "</w:tc></w:tr></w:tbl>";

const FILLERS = 40; // 640 px

const cases = [];
for (const room of [24, 26, 28, 30, 32, 34, 36, 38, 40]) cases.push({ id: `N${room}`, room, lines: 2 });
for (const room of [14, 16, 18, 20, 22, 24]) cases.push({ id: `M${room}`, room, lines: 1 });

let body = cases
  .map((c) => {
    const shimTw = (224 - c.room) * 15;
    let s = "";
    for (let f = 1; f <= FILLERS; f++) s += exactLine(`${c.id}F${f}`);
    s += exactLine(`${c.id}S`, shimTw);
    s += fitTable(c.id, c.lines);
    s += exactLine(`${c.id}MK`);
    s += `<w:p><w:pPr>${SECT}</w:pPr></w:p>`;
    return s;
  })
  .join("");

const NIL_TOP = '<w:tcBorders><w:top w:val="nil"/></w:tcBorders>';
const BOT12 = '<w:tcBorders><w:bottom w:val="single" w:sz="12" w:space="0" w:color="000000"/></w:tcBorders>';
const INSIDE8 = '<w:tblBorders><w:insideH w:val="single" w:sz="8" w:space="0" w:color="000000"/></w:tblBorders>';
body +=
  exactLine("D0REF") + pairTable("D0", "", "", "") + exactLine("D0END") +
  exactLine("DWREF") + pairTable("DW", BOT12, NIL_TOP, "") + exactLine("DWEND") +
  exactLine("DIREF") + pairTable("DI", BOT12, NIL_TOP, INSIDE8) + exactLine("DIEND") +
  SECT;

const parts = unzipSync(new Uint8Array(readFileSync(source)));
const srcDoc = strFromU8(parts["word/document.xml"]);
const docOpen = srcDoc.slice(srcDoc.indexOf("<w:document"), srcDoc.indexOf(">", srcDoc.indexOf("<w:document")) + 1);
parts["word/document.xml"] = strToU8(
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + docOpen + `<w:body>${body}</w:body></w:document>`,
);
const out = join(root, "fixtures-staging/probe-rowfit2-11.docx");
writeFileSync(out, Buffer.from(zipSync(parts)));
console.log(`wrote ${out} — ${cases.length + 3} cases`);
