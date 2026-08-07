#!/usr/bin/env node

/**
 * probe-exactnil11 answered the compat question cleanly — a ca-agreement
 * package rewritten to compatibilityMode 11 zeroes a both-nil boundary at
 * hRule="exact" rows exactly as compat 15 does (A2 = A0 to 0.03px, twice
 * exported) — and yet wild3-template-us-courts-answer, whose caption table is
 * exactly that construct, sits 1.2-1.5pt PER BOUNDARY below Word without the
 * charge. Some other property of the pleading package flips Word's answer.
 *
 * This probe is the same eleven cases in the PLEADING's own package: its
 * settings.xml (compat 11 plus the Word-95-era flag pile: alignTablesRowByRow,
 * layoutTableRowsApart, layoutRawTableWidth, ...), its styles, its fonts —
 * only document.xml is replaced. One variable against probe-exactnil11: the
 * base package. If A2 reads like A1 here, the flag pile (or another package
 * property) is what charges the boundary, and a follow-up can bisect it.
 *
 * Cases (same ids as probe-exactnil11):
 *
 *   exact/exact      A0 none | A1 rule | A2 both-nil | A3 lower-nil | A4 upper-nil
 *   exact/content    B0 none | B1 rule | B2 both-nil
 *   content/content  C0 none | C1 rule | C2 both-nil
 *
 *   node scripts/generate-exactnil11p-probe.mjs
 */

import { readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const source = join(root, "apps/demo/public/fixtures/wild3-template-us-courts-answer.docx");
const parts = unzipSync(new Uint8Array(readFileSync(source)));

// The section carries no header/footer references, so the pleading's header
// parts stay inert; drop their sectPr-independent relationships alone.
const SECT =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '<w:cols w:space="720"/></w:sectPr>';

/** A 10 pt line of exactly known height, used for every marker. */
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

const INSIDE_H = '<w:tblBorders><w:insideH w:val="single" w:sz="12" w:space="0" w:color="000000"/></w:tblBorders>';
// The pleading's construct: rules on every edge, so left/right verticals
// cross the horizontal boundary a nil pair declines.
const FULL =
  "<w:tblBorders>" +
  ["top", "left", "bottom", "right", "insideH", "insideV"]
    .map((s) => `<w:${s} w:val="single" w:sz="12" w:space="0" w:color="000000"/>`)
    .join("") +
  "</w:tblBorders>";
const NIL = (side) => `<w:tcBorders><w:${side} w:val="nil"/></w:tcBorders>`;

const EXACT = '<w:trPr><w:trHeight w:hRule="exact" w:val="495"/></w:trPr>';

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

const body = CASES.map((c) =>
  mark(`${c.id}REF`) + table(c.id, c.borders, c.up, c.lo, c.upper, c.lower) + mark(`${c.id}END`),
).join("");

parts["word/document.xml"] = strToU8(
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    `<w:body>${body}${SECT}</w:body></w:document>`,
);

// Verify the base really is compat 11 (the whole point of this variant).
const settings = strFromU8(parts["word/settings.xml"]);
if (!/compatibilityMode[^>]*w:val="11"/.test(settings)) throw new Error("pleading base is not compat 11");

const out = join(root, "fixtures-staging/probe-exactnil11p.docx");
writeFileSync(out, Buffer.from(zipSync(parts)));
console.log(`wrote ${out} — ${CASES.map((c) => c.id).join(", ")}`);
