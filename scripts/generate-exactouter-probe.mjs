#!/usr/bin/env node

/**
 * Follow-up to probe-exactmar's V1 case (#100). V1 lifted the us-courts
 * caption rows out of their table, which turned two interior boundaries into
 * the table's OUTER top and bottom edges — and exposed a pre-existing
 * overcharge there: with the fixture's all-nil tcBorders our END-REF reads
 * +2.0 pt over Word (one sz-8 rule per outer edge), while the nil-stripped
 * V2/V3 agree to 0.02 pt. Every earlier nil probe measured SHARED boundaries
 * only, so the outer-edge variable was never isolated.
 *
 * This probe isolates it. Two packages from the same base
 * (wild2-legal-ca-agreement, the exactnil base), the single package variable
 * being settings.xml compatibilityMode: probe-exactouter15.docx as authored,
 * probe-exactouter11.docx rewritten 15 -> 11. Each holds two-row tables with
 * the target row at the table's FIRST row (outer top edge, F cases) or LAST
 * row (outer bottom edge, L cases), varying one authored thing per case:
 *
 *   0  no tblBorders at all               baseline
 *   R  outer rule single sz=12 (2.00 px)  what does a live outer rule cost?
 *   N  same rule + target cell nil there  does the nil zero it?
 *
 * Target rows: exact 115 tw (X, shorter than its line), exact 495 tw (Y, the
 * scaling control), and a content row (C, the V1 construct). The companion
 * row is always a content row. Cell margins are zero throughout so the
 * pre-15 bottom-margin charge cannot confound the border question.
 *
 * Observable per case: top(END) - top(REF). Every row's own mark line is
 * 10 pt text on an exact 240 line (16.00 px).
 *
 *   node scripts/generate-exactouter-probe.mjs
 */

import { readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const source = join(root, "apps/demo/public/fixtures/wild2-legal-ca-agreement.docx");

const SECT =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '<w:cols w:space="720"/></w:sectPr>';

/** A 10 pt line of exactly known height (16.00 px), used for every marker. */
const mark = (text) =>
  '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/>' +
  '<w:rPr><w:sz w:val="20"/></w:rPr></w:pPr>' +
  `<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;

/** Zero cell margins, so the only charge in play is the border under test. */
const TBL_PR = (borders) =>
  '<w:tblPr><w:tblW w:w="9360" w:type="dxa"/>' +
  borders +
  '<w:tblLayout w:type="fixed"/>' +
  '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/>' +
  '<w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>';

const RULE = (side) =>
  `<w:tblBorders><w:${side} w:val="single" w:sz="12" w:space="0" w:color="000000"/></w:tblBorders>`;
const NIL = (side) => `<w:tcBorders><w:${side} w:val="nil"/></w:tcBorders>`;
const EXACT = (h) => `<w:trPr><w:trHeight w:hRule="exact" w:val="${h}"/></w:trPr>`;

const row = (trPr, tcBorders, text) =>
  `<w:tr>${trPr}` +
  `<w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/>${tcBorders}</w:tcPr>` +
  mark(text) +
  "</w:tc></w:tr>";

/** F: target row first (outer TOP edge under test); L: target row last. */
const table = (id, borders, targetTrPr, targetTc, first) => {
  const target = row(targetTrPr, targetTc, `${id}A`);
  const companion = row("", "", `${id}B`);
  const rows = first ? target + companion : companion + target;
  return "<w:tbl>" + TBL_PR(borders) + '<w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid>' + rows + "</w:tbl>";
};

const CASES = [];
for (const [tag, trPr] of [
  ["X", EXACT(115)],
  ["Y", EXACT(495)],
  ["C", ""],
]) {
  for (const [pos, first, side] of [
    ["F", true, "top"],
    ["L", false, "bottom"],
  ]) {
    CASES.push({ id: `${tag}${pos}0`, borders: "", trPr, tc: "", first });
    CASES.push({ id: `${tag}${pos}R`, borders: RULE(side), trPr, tc: "", first });
    CASES.push({ id: `${tag}${pos}N`, borders: RULE(side), trPr, tc: NIL(side), first });
  }
}

const body = CASES.map(
  (c) => mark(`${c.id}REF`) + table(c.id, c.borders, c.trPr, c.tc, c.first) + mark(`${c.id}END`),
).join("");

const documentXml =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  `<w:body>${body}${SECT}</w:body></w:document>`;

for (const compat of ["15", "11"]) {
  const parts = unzipSync(new Uint8Array(readFileSync(source)));
  const settings = strFromU8(parts["word/settings.xml"]);
  if (!/compatSetting w:name="compatibilityMode"[^>]*w:val="15"/.test(settings))
    throw new Error("compatibilityMode 15 not found in base settings.xml");
  if (compat !== "15")
    parts["word/settings.xml"] = strToU8(
      settings.replace(/(compatSetting w:name="compatibilityMode"[^>]*w:val=")15(")/, `$1${compat}$2`),
    );
  parts["word/document.xml"] = strToU8(documentXml);
  const out = join(root, `fixtures-staging/probe-exactouter${compat}.docx`);
  writeFileSync(out, Buffer.from(zipSync(parts)));
  console.log(`wrote ${out} — ${CASES.map((c) => c.id).join(", ")}`);
}
