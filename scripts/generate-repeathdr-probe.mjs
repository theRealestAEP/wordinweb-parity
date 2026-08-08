#!/usr/bin/env node

/**
 * #108a: what does Word charge at a repeated-header CONTINUATION segment top?
 *
 * probe-exactouter pinned the table's TRUE outer top: a live tblBorders top
 * rule charges the flow its full width above a content first row, nothing
 * above an exact first row, and a nil on row 0's cell edge zeroes it. No
 * probe measured the same edge where a repeated tblHeader stack re-paints
 * row 0 at the top of a continuation page, so the engine keeps a
 * fixture-calibrated charge there (nilSuppressedOuterTop, calibrated on
 * us-courts-answer pages 2-7 only).
 *
 * This probe measures that edge directly. Each case is its own section
 * holding one single-column table: row 0 is a tblHeader row, followed by 75
 * one-line content rows, so the table always continues onto the section's
 * second page and Word repeats the header there. One authored thing varies
 * per case:
 *
 *   0   no tblBorders                       baseline
 *   R   tblBorders top single sz=12         live outer rule
 *   N   same rule + row-0 cells nil top     the us-courts construct
 *   W   tblBorders top single sz=24         width scaling (X height only)
 *   M   sz=24 rule + row-0 cells nil top    width scaling of the nil case
 *
 * Header heights: exact 115 tw (X, shorter than its 16.00 px line — the
 * us-courts shape), exact 495 tw (Y, the scaling control), content (C).
 * Only w:top is authored in tblBorders so no insideH confounds the data-row
 * boundaries; cell margins are zero throughout.
 *
 * Observables per case, on the case's FIRST page and its CONTINUATION page:
 * the header mark's top (content inset channel) and the first data mark's
 * top (flow channel). Two packages differing only in compatibilityMode
 * (15 as authored, 11 rewritten), us-courts being compat 11.
 *
 *   node scripts/generate-repeathdr-probe.mjs
 */

import { readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const source = join(root, "apps/demo/public/fixtures/wild2-legal-ca-agreement.docx");

const SECT =
  '<w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '<w:cols w:space="720"/>';

/** A 10 pt line of exactly known height (16.00 px), used for every marker. */
const mark = (text) =>
  '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/>' +
  '<w:rPr><w:sz w:val="20"/></w:rPr></w:pPr>' +
  `<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;

const TBL_PR = (borders) =>
  '<w:tblPr><w:tblW w:w="9360" w:type="dxa"/>' +
  borders +
  '<w:tblLayout w:type="fixed"/>' +
  '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/>' +
  '<w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>';

const RULE = (sz) =>
  `<w:tblBorders><w:top w:val="single" w:sz="${sz}" w:space="0" w:color="000000"/></w:tblBorders>`;
const NIL_TOP = '<w:tcBorders><w:top w:val="nil"/></w:tcBorders>';
const EXACT = (h) => `<w:trHeight w:hRule="exact" w:val="${h}"/>`;

const row = (trPr, tcBorders, text) =>
  `<w:tr>${trPr}` +
  `<w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/>${tcBorders}</w:tcPr>` +
  mark(text) +
  "</w:tc></w:tr>";

const DATA_ROWS = 75;

const table = (id, borders, headerTrHeight, headerTc) => {
  const header = row(`<w:trPr>${headerTrHeight}<w:tblHeader/></w:trPr>`, headerTc, `${id}H`);
  let data = "";
  for (let i = 1; i <= DATA_ROWS; i++)
    data += row("", "", `${id}D${String(i).padStart(2, "0")}`);
  return "<w:tbl>" + TBL_PR(borders) + '<w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid>' + header + data + "</w:tbl>";
};

const CASES = [];
for (const [tag, trHeight] of [
  ["X", EXACT(115)],
  ["Y", EXACT(495)],
  ["C", ""],
]) {
  CASES.push({ id: `${tag}0`, borders: "", trHeight, tc: "" });
  CASES.push({ id: `${tag}R`, borders: RULE(12), trHeight, tc: "" });
  CASES.push({ id: `${tag}N`, borders: RULE(12), trHeight, tc: NIL_TOP });
}
CASES.push({ id: "XW", borders: RULE(24), trHeight: EXACT(115), tc: "" });
CASES.push({ id: "XM", borders: RULE(24), trHeight: EXACT(115), tc: NIL_TOP });

const body =
  CASES.map((c, i) => {
    const tbl = table(c.id, c.borders, c.trHeight, c.tc);
    const sectBreak =
      i < CASES.length - 1
        ? `<w:p><w:pPr><w:sectPr>${SECT}</w:sectPr></w:pPr></w:p>`
        : "";
    return mark(`${c.id}REF`) + tbl + sectBreak;
  }).join("") + `<w:sectPr>${SECT}</w:sectPr>`;

const documentXml =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  `<w:body>${body}</w:body></w:document>`;

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
  const out = join(root, `fixtures-staging/probe-repeathdr${compat}.docx`);
  writeFileSync(out, Buffer.from(zipSync(parts)));
  console.log(`wrote ${out} — ${CASES.map((c) => c.id).join(", ")}`);
}
