#!/usr/bin/env node

/**
 * #108a follow-up. probe-repeathdr pinned the repeated-header continuation
 * top at ZERO charge (nil honored, exactly like the true outer top), but
 * removing the fixture-calibrated nilSuppressedOuterTop charge exposes what
 * it was masking: us-courts-answer p4's ten uniform "answer" blocks run
 * 0.69 px (0.52 pt) per block short of Word in our render — a rule-to-rule
 * pitch deficit measured from the sz-1 spacer-bottom rules, present in
 * every build and independent of the segment top.
 *
 * Each block is [content row + spacer row], both carrying a
 * `tblPrEx tblCellMar 0/0`, per-cell tcMar top 58 / bottom 14, all-nil
 * tcBorders except the spacer's second cell, whose bottom is single sz-1;
 * the content row opens with an EMPTY 11 pt paragraph and its text line
 * mixes 11 pt and italic 9 pt runs; the spacer holds a FORMTEXT field of
 * en-spaces under a trHeight 624 (atLeast) floor; the table declares sz-8
 * borders everywhere including insideH.
 *
 * This probe rebuilds the block VERBATIM from the fixture's own package
 * (styles, fonts, settings — compat 11 as authored; a second package
 * rewritten to 15) and repeats it six times per table, one table per page,
 * stripping ONE authored thing per case:
 *
 *   V0  verbatim                        the baseline; must reproduce 85.03 px
 *   V1  spacer cell-1 bottom sz-1->nil  is the pitch in the per-cell rule?
 *   V2  tblBorders insideH removed      ... or in the table-wide rule?
 *   V3  tblPrEx removed                 ... or the row-level margin zero?
 *   V4  leading empty paragraph removed ... or the empty 11pt line?
 *   V5  italic sz-18 run -> sz-22       ... or the mixed-size line height?
 *   V6  FORMTEXT -> plain runs          ... or the field?
 *   V7  trHeight 624 removed            ... or the atLeast floor?
 *
 * Observable per case: the intra-table block pitch — sz-1 rule to sz-1 rule
 * in the Word PDF (text tops where V1 removes the rules), five gaps per
 * table.
 *
 *   node scripts/generate-uscourtsblock-probe.mjs
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
const spacerIdx = rows.indexOf(contentRow) + 1;
const spacerRow = rows[spacerIdx];
if (!contentRow || !spacerRow.includes('w:sz="1"')) throw new Error("block rows not found");

// The fixture's own tblPr, minus its 124-row table's width quirks kept as is.
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

/** One case: strip transforms applied to the two verbatim rows / tblPr. */
const CASES = [
  { id: "V0", tf: (c, s, tp) => [c, s, tp] },
  { id: "V1", tf: (c, s, tp) => [c, s.replace('<w:bottom w:val="single" w:sz="1" w:space="0" w:color="000000"/>', '<w:bottom w:val="nil"/>'), tp] },
  { id: "V2", tf: (c, s, tp) => [c, s, tp.replace(/<w:insideH [^/]*\/>/, "")] },
  { id: "V3", tf: (c, s, tp) => [c.replace(/<w:tblPrEx>[\s\S]*?<\/w:tblPrEx>/, ""), s.replace(/<w:tblPrEx>[\s\S]*?<\/w:tblPrEx>/, ""), tp] },
  { id: "V4", tf: (c, s, tp) => [c.replace(/<w:p [^>]*><w:pPr><w:widowControl w:val="0"\/><w:rPr><w:sz w:val="22"\/><\/w:rPr><\/w:pPr><\/w:p>/, ""), s, tp] },
  { id: "V5", tf: (c, s, tp) => [c.replace('<w:rPr><w:i/><w:sz w:val="18"/></w:rPr>', '<w:rPr><w:sz w:val="22"/></w:rPr>'), s, tp] },
  {
    id: "V6",
    tf: (c, s, tp) => [
      c,
      s
        .replace(/<w:r><w:rPr><w:sz w:val="22"\/><\/w:rPr><w:fldChar[\s\S]*?<w:fldChar w:fldCharType="separate"\/><\/w:r>/, "")
        .replace(/<w:r><w:rPr><w:sz w:val="22"\/><\/w:rPr><w:fldChar w:fldCharType="end"\/><\/w:r>/, "")
        .replace(/<w:bookmarkStart[^/]*\/>/, "")
        .replace(/<w:bookmarkEnd[^/]*\/>/, ""),
      tp,
    ],
  },
  { id: "V7", tf: (c, s, tp) => [c, s.replace(/<w:trHeight w:val="624"\/>/, ""), tp] },
];

const body =
  CASES.map((cs, i) => {
    const [c, s, tp] = cs.tf(contentRow, spacerRow, tblPr);
    const tbl = "<w:tbl>" + tp + grid + (c + s).repeat(6) + "</w:tbl>";
    const brk =
      i < CASES.length - 1
        ? '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
        : "";
    return mark(`${cs.id}REF`) + tbl + mark(`${cs.id}END`) + brk;
  }).join("") + SECT;

const documentXml =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  srcDoc.slice(srcDoc.indexOf("<w:document"), srcDoc.indexOf(">", srcDoc.indexOf("<w:document")) + 1) +
  `<w:body>${body}</w:body></w:document>`;

// Sanity: each strip case must actually change its target.
for (const cs of CASES.slice(1)) {
  const [c, s, tp] = cs.tf(contentRow, spacerRow, tblPr);
  if (c === contentRow && s === spacerRow && tp === tblPr)
    throw new Error(`case ${cs.id} stripped nothing`);
}

for (const compat of ["11", "15"]) {
  const parts = unzipSync(new Uint8Array(readFileSync(source)));
  const settings = strFromU8(parts["word/settings.xml"]);
  const m = settings.match(/compatSetting w:name="compatibilityMode"[^>]*w:val="(\d+)"/);
  if (!m) throw new Error("compatibilityMode not found in base settings.xml");
  if (m[1] !== compat)
    parts["word/settings.xml"] = strToU8(
      settings.replace(/(compatSetting w:name="compatibilityMode"[^>]*w:val=")\d+(")/, `$1${compat}$2`),
    );
  parts["word/document.xml"] = strToU8(documentXml);
  const out = join(root, `fixtures-staging/probe-uscourtsblock${compat}.docx`);
  writeFileSync(out, Buffer.from(zipSync(parts)));
  console.log(`wrote ${out} — ${CASES.map((c) => c.id).join(", ")}`);
}
