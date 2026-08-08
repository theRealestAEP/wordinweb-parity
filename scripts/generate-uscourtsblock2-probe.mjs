#!/usr/bin/env node

/**
 * Round 2 of probe-uscourtsblock: the first round showed Word's block pitch
 * is INVARIANT to the tblBorders insideH (V2), the sz-1 cell rule (V1), the
 * tblPrEx (V3), the run mix (V5) and the field (V6), and moves only with
 * the leading empty paragraph (V4) and the trHeight floor (V7). Our render
 * carries three offsetting errors the mixed stack cannot separate: the
 * boundary charge, the atLeast floor arithmetic, and the natural row
 * heights. This round isolates each with single-kind stacks:
 *
 *   S0  spacer row x6 verbatim            floored spacer + spacer/spacer boundary
 *   SN  spacer x6, trHeight removed       natural spacer + boundary
 *   SB  spacer x6, insideH removed        does insideH enter when every cell
 *                                         declares its own edge? (per-pair rule)
 *   SW  spacer x6, cell-1 bottom sz-12    the declared cell rule charged in
 *                                         full against the nil below?
 *   C6  content row x6 verbatim           natural content row + both-nil boundary
 *   CN  content x6, empty para removed    the empty 11pt line
 *
 * Same base packages (us-courts-answer, compat 11 as authored / 15
 * rewritten), one case per page, six repeats per table.
 *
 *   node scripts/generate-uscourtsblock2-probe.mjs
 */

import { readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const source = join(root, "apps/demo/public/fixtures/wild3-template-us-courts-answer.docx");

const srcDoc = strFromU8(unzipSync(new Uint8Array(readFileSync(source)))["word/document.xml"]);
const rows = srcDoc.match(/<w:tr(?: [^>]*)?>[\s\S]*?<\/w:tr>/g);
const contentRow = rows.find((r) => r.includes("Zujowibi"));
const spacerRow = rows[rows.indexOf(contentRow) + 1];
if (!contentRow || !spacerRow.includes('w:sz="1"')) throw new Error("block rows not found");

const tblPr = srcDoc.match(/<w:tblPr>[\s\S]*?<\/w:tblPr>/)[0];
const grid = srcDoc.match(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/)[0];

const SECT =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '<w:cols w:space="720"/></w:sectPr>';

const mark = (text) =>
  '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/>' +
  '<w:rPr><w:sz w:val="20"/></w:rPr></w:pPr>' +
  `<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;

const CASES = [
  { id: "S0", row: spacerRow, tp: tblPr },
  { id: "SN", row: spacerRow.replace(/<w:trHeight w:val="624"\/>/, ""), tp: tblPr },
  { id: "SB", row: spacerRow, tp: tblPr.replace(/<w:insideH [^/]*\/>/, "") },
  {
    id: "SW",
    row: spacerRow.replace(
      '<w:bottom w:val="single" w:sz="1" w:space="0" w:color="000000"/>',
      '<w:bottom w:val="single" w:sz="12" w:space="0" w:color="000000"/>',
    ),
    tp: tblPr,
  },
  { id: "C6", row: contentRow, tp: tblPr },
  {
    id: "CN",
    row: contentRow.replace(
      /<w:p [^>]*><w:pPr><w:widowControl w:val="0"\/><w:rPr><w:sz w:val="22"\/><\/w:rPr><\/w:pPr><\/w:p>/,
      "",
    ),
    tp: tblPr,
  },
];
for (const c of CASES) {
  if (c.id !== "S0" && c.id !== "C6" && c.row === spacerRow && c.tp === tblPr)
    throw new Error(`case ${c.id} stripped nothing`);
}

const body =
  CASES.map((c, i) => {
    const tbl = "<w:tbl>" + c.tp + grid + c.row.repeat(6) + "</w:tbl>";
    const brk = i < CASES.length - 1 ? '<w:p><w:r><w:br w:type="page"/></w:r></w:p>' : "";
    return mark(`${c.id}REF`) + tbl + mark(`${c.id}END`) + brk;
  }).join("") + SECT;

const documentXml =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  srcDoc.slice(srcDoc.indexOf("<w:document"), srcDoc.indexOf(">", srcDoc.indexOf("<w:document")) + 1) +
  `<w:body>${body}</w:body></w:document>`;

for (const compat of ["11", "15"]) {
  const parts = unzipSync(new Uint8Array(readFileSync(source)));
  const settings = strFromU8(parts["word/settings.xml"]);
  const m = settings.match(/compatSetting w:name="compatibilityMode"[^>]*w:val="(\d+)"/);
  if (!m) throw new Error("compatibilityMode not found");
  if (m[1] !== compat)
    parts["word/settings.xml"] = strToU8(
      settings.replace(/(compatSetting w:name="compatibilityMode"[^>]*w:val=")\d+(")/, `$1${compat}$2`),
    );
  parts["word/document.xml"] = strToU8(documentXml);
  const out = join(root, `fixtures-staging/probe-uscourtsblock2-${compat}.docx`);
  writeFileSync(out, Buffer.from(zipSync(parts)));
  console.log(`wrote ${out} — ${CASES.map((c) => c.id).join(", ")}`);
}
