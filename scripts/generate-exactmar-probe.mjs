#!/usr/bin/env node

/**
 * Third round on the us-courts caption table. probe-exactnil11/-11p/-exactclip
 * proved Word zeroes a both-nil boundary at exact rows in compat 11 — zero
 * cell margins, any row height, full borders crossing — yet the fixture's
 * Word layout spaces those same rows exactly as if the boundary were charged.
 * The one construct difference left: the fixture's cells carry w:tcMar
 * (top 58, bottom 29 twips) and w:cantSplit, and every probe so far zeroed
 * margins.
 *
 * So: rows 3-5 of the caption table VERBATIM (content 'for the', exact-115
 * spacer, content 'Rewugofi of'), under the fixture's own tblPr and tblGrid,
 * in the fixture's own package. Then one variable stripped per case:
 *
 *   V0  verbatim                       reproduces the fixture spacing?
 *   V1  tcMar removed                  margins are the charge?
 *   V2  tcBorders (the nils) removed   the nil is irrelevant here?
 *   V3  tcMar AND tcBorders removed    both
 *
 * Observable: END-REF across the 3-row table (a 12pt marker line above and
 * below). Fixture ground truth in situ: 'for the' to 'Rewugofi of' spans
 * 24.02pt in the Word PDF where authored heights + margins predict ~23pt.
 *
 *   node scripts/generate-exactmar-probe.mjs
 */

import { readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const source = join(root, "apps/demo/public/fixtures/wild3-template-us-courts-answer.docx");
const parts = unzipSync(new Uint8Array(readFileSync(source)));

const doc = strFromU8(parts["word/document.xml"]);

// The caption table is the body's first <w:tbl>; lift its tblPr, tblGrid and
// rows verbatim so nothing is re-authored.
const tblStart = doc.indexOf("<w:tbl>");
const tblEnd = doc.indexOf("</w:tbl>", tblStart);
const tblXml = doc.slice(tblStart, tblEnd + "</w:tbl>".length);
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
const rows345 = rowXmls.slice(3, 6).join("");
if (!/for the/.test(rows345) || !/w:val="115"/.test(rows345)) throw new Error("rows 3-5 are not the expected slice");

const SECT =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '<w:cols w:space="720"/></w:sectPr>';

const mark = (text) =>
  '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/>' +
  '<w:rPr><w:sz w:val="20"/></w:rPr></w:pPr>' +
  `<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;

const variants = [
  { id: "V0", xml: rows345 },
  { id: "V1", xml: rows345.replace(/<w:tcMar>[\s\S]*?<\/w:tcMar>/g, "") },
  { id: "V2", xml: rows345.replace(/<w:tcBorders>[\s\S]*?<\/w:tcBorders>/g, "") },
  {
    id: "V3",
    xml: rows345
      .replace(/<w:tcMar>[\s\S]*?<\/w:tcMar>/g, "")
      .replace(/<w:tcBorders>[\s\S]*?<\/w:tcBorders>/g, ""),
  },
];

const body = variants
  .map((v) => mark(`${v.id}REF`) + "<w:tbl>" + tblPr + tblGrid + v.xml + "</w:tbl>" + mark(`${v.id}END`))
  .join("");

parts["word/document.xml"] = strToU8(
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    `<w:body>${body}${SECT}</w:body></w:document>`,
);

const out = join(root, "fixtures-staging/probe-exactmar.docx");
writeFileSync(out, Buffer.from(zipSync(parts)));
console.log(`wrote ${out} — ${variants.map((v) => v.id).join(", ")}`);
