#!/usr/bin/env node

/**
 * probe-exactmar localized the us-courts deficit to CELL MARGINS around the
 * exact-115 spacer row: with the fixture's tcMar (top 58, bottom 29) Word
 * spaces 'for the' -> 'Rewugofi of' 5.75pt wider than the no-margin variant,
 * and we only add 4.35pt — the content rows' own 29-bottom + 58-top. The
 * missing ~1.4pt is one of the EXACT row's own pads; this probe varies that
 * row's tcMar alone to pin which:
 *
 *   M0  (58, 29)  verbatim control — deficit reproduces?
 *   M1  (29, 58)  swapped — deficit follows the bottom value?
 *   M2  (58, 0)   top only
 *   M3  (0, 29)   bottom only
 *   M4  (0, 0)    neither
 *
 * Rows 3-5 of the caption table otherwise verbatim (fixture package, nil
 * borders, cantSplit, content-row margins untouched). Observable: the
 * 'for the' -> 'Rewugofi of' glyph-top distance, plus REF/END flow.
 *
 *   node scripts/generate-exactpad-probe.mjs
 */

import { readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const source = join(root, "apps/demo/public/fixtures/wild3-template-us-courts-answer.docx");
const parts = unzipSync(new Uint8Array(readFileSync(source)));

const doc = strFromU8(parts["word/document.xml"]);
const tblStart = doc.indexOf("<w:tbl>");
const tblXml = doc.slice(tblStart, doc.indexOf("</w:tbl>", tblStart) + "</w:tbl>".length);
const tblPr = tblXml.match(/<w:tblPr>[\s\S]*?<\/w:tblPr>/)[0];
const tblGrid = tblXml.match(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/)[0];
const rowXmls = [];
{
  const re = /<w:tr[ >]/g;
  let m;
  while ((m = re.exec(tblXml))) {
    const e = tblXml.indexOf("</w:tr>", m.index);
    rowXmls.push(tblXml.slice(m.index, e + "</w:tr>".length));
    re.lastIndex = e;
  }
}
const [row3, row4, row5] = rowXmls.slice(3, 6);
if (!/w:val="115"/.test(row4)) throw new Error("row 4 is not the exact-115 spacer");

const padRow4 = (top, bottom) =>
  row4.replace(
    /<w:tcMar>[\s\S]*?<\/w:tcMar>/,
    "<w:tcMar>" +
      `<w:top w:w="${top}" w:type="dxa"/><w:left w:w="58" w:type="dxa"/>` +
      `<w:bottom w:w="${bottom}" w:type="dxa"/><w:right w:w="58" w:type="dxa"/>` +
      "</w:tcMar>",
  );

const SECT =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '<w:cols w:space="720"/></w:sectPr>';

const mark = (text) =>
  '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/>' +
  '<w:rPr><w:sz w:val="20"/></w:rPr></w:pPr>' +
  `<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;

const variants = [
  { id: "M0", top: 58, bottom: 29 },
  { id: "M1", top: 29, bottom: 58 },
  { id: "M2", top: 58, bottom: 0 },
  { id: "M3", top: 0, bottom: 29 },
  { id: "M4", top: 0, bottom: 0 },
];

const body = variants
  .map(
    (v) =>
      mark(`${v.id}REF`) +
      "<w:tbl>" +
      tblPr +
      tblGrid +
      row3 +
      padRow4(v.top, v.bottom) +
      row5 +
      "</w:tbl>" +
      mark(`${v.id}END`),
  )
  .join("");

parts["word/document.xml"] = strToU8(
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    `<w:body>${body}${SECT}</w:body></w:document>`,
);

const out = join(root, "fixtures-staging/probe-exactpad.docx");
writeFileSync(out, Buffer.from(zipSync(parts)));
console.log(`wrote ${out} — ${variants.map((v) => v.id).join(", ")}`);
