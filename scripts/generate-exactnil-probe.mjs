#!/usr/bin/env node

/**
 * The one variant e186c1e's restriction left open: a both-nil boundary INSIDE
 * an `hRule="exact"` row.
 *
 * #51 settled the height question on `atLeast` rows — a `w:val="nil"` on BOTH
 * cells sharing a boundary returns the row to its no-rule height, and a nil on
 * only one side suppresses nothing — and #78 landed it. #51b settled the
 * sided-ness question on exact rows: a cell border comes out of the content box
 * at the edge the border is ON, in full.
 *
 * Neither reaches this case. An exact row's HEIGHT cannot answer anything (it
 * is the authored value whatever the borders say), so the only observable is
 * whether Word INSETS the content — and the both-nil rule has never been tested
 * where the inset, rather than the height, is what it would change.
 *
 * Same shape as the sidedness probe, so the numbers compare directly: a marker
 * immediately before an `hRule="exact"` row and another inside it, reporting
 * `top(MK) - top(REF)` in CSS px. Here the table DOES carry `w:tblBorders`,
 * because a nil has to have a rule to decline; the sidedness probe deliberately
 * had none.
 *
 *   N      no tblBorders, no tcBorders            baseline: no rule anywhere
 *   R      tblBorders insideH sz=12, no nil       the rule is charged (#51b: +2.00)
 *   RN     insideH sz=12, BOTH cells nil          both-nil: 0.00 or +2.00?
 *   RO     insideH sz=12, only the LOWER cell nil one-sided: #51 says charged
 *   RU     insideH sz=12, only the UPPER cell nil the other one-sided ordering
 *
 * Every case is TWO exact rows, because a shared interior boundary needs a row
 * on each side of it; MK sits in the second row and REF above the table, so the
 * measured distance carries row 1's whole height plus whatever the boundary
 * costs the second row's content.
 *
 * The prediction under test: if Word honours a both-nil boundary here the way
 * it does an atLeast row's height, RN equals N + row1 and R exceeds it by the
 * 1.5pt rule = 2.00 px. If RN equals R, the both-nil rule is about the row
 * HEIGHT only and does not reach the inset — which would mean #78's fix is
 * correctly scoped and nothing more is owed.
 *
 *   node scripts/generate-exactnil-probe.mjs
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

/** A 10 pt line of exactly known height, used for every marker. */
const mark = (text) =>
  '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/>' +
  '<w:rPr><w:sz w:val="20"/></w:rPr></w:pPr>' +
  `<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;

/** Zero cell margins, so the only thing between the row box and its content is
 *  the border under test. */
// CT_TblPrBase is a SEQUENCE: tblW, tblBorders, tblLayout, tblCellMar. Word
// repairs the wrong order silently, which would leave the probe measuring a
// package Word rewrote rather than the one authored here.
const TBL_PR = (borders) =>
  '<w:tblPr><w:tblW w:w="9360" w:type="dxa"/>' +
  borders +
  '<w:tblLayout w:type="fixed"/>' +
  '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/>' +
  '<w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>';

const INSIDE_H = '<w:tblBorders><w:insideH w:val="single" w:sz="12" w:space="0" w:color="000000"/></w:tblBorders>';
const NIL = (side) => `<w:tcBorders><w:${side} w:val="nil"/></w:tcBorders>`;

/** Two exact rows of the same height; `upper`/`lower` carry each row's own
 *  tcBorders, which is where a nil declines the shared insideH rule. */
const table = (id, borders, upper, lower) =>
  "<w:tbl>" +
  TBL_PR(borders) +
  '<w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid>' +
  '<w:tr><w:trPr><w:trHeight w:hRule="exact" w:val="495"/></w:trPr>' +
  `<w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/>${upper}</w:tcPr>` +
  mark(`${id}UP`) +
  "</w:tc></w:tr>" +
  '<w:tr><w:trPr><w:trHeight w:hRule="exact" w:val="495"/></w:trPr>' +
  `<w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/>${lower}</w:tcPr>` +
  mark(`${id}MK`) +
  "</w:tc></w:tr></w:tbl>";

const CASES = [
  { id: "N", borders: "", upper: "", lower: "" },
  { id: "R", borders: INSIDE_H, upper: "", lower: "" },
  { id: "RN", borders: INSIDE_H, upper: NIL("bottom"), lower: NIL("top") },
  { id: "RO", borders: INSIDE_H, upper: "", lower: NIL("top") },
  { id: "RU", borders: INSIDE_H, upper: NIL("bottom"), lower: "" },
];

// A marker after each table keeps consecutive tables from merging.
const body = CASES.map((c) => mark(`${c.id}REF`) + table(c.id, c.borders, c.upper, c.lower) + mark(`${c.id}END`)).join(
  "",
);

parts["word/document.xml"] = strToU8(
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    `<w:body>${body}${SECT}</w:body></w:document>`,
);

// Probes live in fixtures-staging: parity-parallel.mjs adopts a reference only
// when a DOCX of the same name sits in apps/demo/public/fixtures.
const out = join(root, "fixtures-staging/probe-exactnil.docx");
writeFileSync(out, Buffer.from(zipSync(parts)));
console.log(`wrote ${out} — ${CASES.map((c) => c.id).join(", ")}`);
