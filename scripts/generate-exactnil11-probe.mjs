#!/usr/bin/env node

/**
 * probe-exactnil measured the both-nil boundary at hRule="exact" rows in a
 * COMPAT-15 package, and #86 rebuilt the exact-row arithmetic on it. But the
 * guard #86 removed (rowBorderWidths' contentSized check) was not compat
 * gated, and wild3-template-us-courts-answer — compatibilityMode 11, a
 * 124-row caption table whose boundaries are nearly all both-nil around
 * exact rows — moved p1 from 0.08% to 24.94% at exactly that commit. The
 * pre-15 case is unmeasured; this probe measures it.
 *
 * Same shape and same marker as generate-exactnil-probe.mjs, and the SAME
 * base package (wild2-legal-ca-agreement), so the single variable against
 * probe-exactnil is settings.xml's compatibilityMode, rewritten 15 -> 11
 * here. Three row-sizing pairings, each with a no-rule baseline, a live
 * insideH sz-12 rule, and the nil variants:
 *
 *   exact/exact      A0 none | A1 rule | A2 both-nil | A3 lower-nil | A4 upper-nil
 *   exact/content    B0 none | B1 rule | B2 both-nil
 *   content/content  C0 none | C1 rule | C2 both-nil
 *
 * Every case reports top(MK) - top(REF): REF sits immediately above the
 * table, MK inside the second row. If pre-15 Word charges a both-nil
 * boundary at an exact row, A2 reads like A1; if the compat-15 zero holds,
 * A2 reads like A0.
 *
 *   node scripts/generate-exactnil11-probe.mjs
 */

import { readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const source = join(root, "apps/demo/public/fixtures/wild2-legal-ca-agreement.docx");
const parts = unzipSync(new Uint8Array(readFileSync(source)));

// The one variable: the base package's compatibilityMode 15 becomes 11.
const settings = strFromU8(parts["word/settings.xml"]);
const rewritten = settings.replace(
  /(compatSetting w:name="compatibilityMode"[^>]*w:val=")15(")/,
  "$111$2",
);
if (rewritten === settings) throw new Error("compatibilityMode 15 not found in base settings.xml");
parts["word/settings.xml"] = strToU8(rewritten);

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
 *  the border under test. Property order per CT_TblPrBase's sequence. */
const TBL_PR = (borders) =>
  '<w:tblPr><w:tblW w:w="9360" w:type="dxa"/>' +
  borders +
  '<w:tblLayout w:type="fixed"/>' +
  '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/>' +
  '<w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>';

const INSIDE_H = '<w:tblBorders><w:insideH w:val="single" w:sz="12" w:space="0" w:color="000000"/></w:tblBorders>';
// Rules on every edge, so left/right verticals cross the horizontal boundary
// a nil pair declines (the us-courts caption table's construct).
const FULL =
  "<w:tblBorders>" +
  ["top", "left", "bottom", "right", "insideH", "insideV"]
    .map((s) => `<w:${s} w:val="single" w:sz="12" w:space="0" w:color="000000"/>`)
    .join("") +
  "</w:tblBorders>";
const NIL = (side) => `<w:tcBorders><w:${side} w:val="nil"/></w:tcBorders>`;

const EXACT = '<w:trPr><w:trHeight w:hRule="exact" w:val="495"/></w:trPr>';

/** Two rows; `rule1`/`rule2` choose exact-495 or content sizing, and
 *  `upper`/`lower` carry each row's own tcBorders. */
const table = (id, borders, upperRow, lowerRow, upper, lower) =>
  "<w:tbl>" +
  TBL_PR(borders) +
  '<w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid>' +
  `<w:tr>${upperRow}` +
  `<w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/>${upper}</w:tcPr>` +
  mark(`${id}UP`) +
  "</w:tc></w:tr>" +
  `<w:tr>${lowerRow}` +
  `<w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/>${lower}</w:tcPr>` +
  mark(`${id}MK`) +
  "</w:tc></w:tr></w:tbl>";

const CASES = [
  { id: "A0", borders: "", up: EXACT, lo: EXACT, upper: "", lower: "" },
  { id: "A1", borders: INSIDE_H, up: EXACT, lo: EXACT, upper: "", lower: "" },
  { id: "A2", borders: INSIDE_H, up: EXACT, lo: EXACT, upper: NIL("bottom"), lower: NIL("top") },
  { id: "A3", borders: INSIDE_H, up: EXACT, lo: EXACT, upper: "", lower: NIL("top") },
  { id: "A4", borders: INSIDE_H, up: EXACT, lo: EXACT, upper: NIL("bottom"), lower: "" },
  { id: "B0", borders: "", up: EXACT, lo: "", upper: "", lower: "" },
  { id: "B1", borders: INSIDE_H, up: EXACT, lo: "", upper: "", lower: "" },
  { id: "B2", borders: INSIDE_H, up: EXACT, lo: "", upper: NIL("bottom"), lower: NIL("top") },
  { id: "C0", borders: "", up: "", lo: "", upper: "", lower: "" },
  { id: "C1", borders: INSIDE_H, up: "", lo: "", upper: "", lower: "" },
  { id: "C2", borders: INSIDE_H, up: "", lo: "", upper: NIL("bottom"), lower: NIL("top") },
  { id: "D1", borders: FULL, up: EXACT, lo: EXACT, upper: "", lower: "" },
  { id: "D2", borders: FULL, up: EXACT, lo: EXACT, upper: NIL("bottom"), lower: NIL("top") },
];

// A marker after each table keeps consecutive tables from merging.
const body = CASES.map((c) =>
  mark(`${c.id}REF`) + table(c.id, c.borders, c.up, c.lo, c.upper, c.lower) + mark(`${c.id}END`),
).join("");

parts["word/document.xml"] = strToU8(
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    `<w:body>${body}${SECT}</w:body></w:document>`,
);

const out = join(root, "fixtures-staging/probe-exactnil11.docx");
writeFileSync(out, Buffer.from(zipSync(parts)));
console.log(`wrote ${out} — ${CASES.map((c) => c.id).join(", ")}`);
