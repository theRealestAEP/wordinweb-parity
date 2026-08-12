#!/usr/bin/env node

/**
 * Line-numbering conformance probe (w:lnNumType). Settles what Word counts as
 * a "line" for margin numbering and how countBy/restart/start/distance behave,
 * so the engine's emitLineNumber (packages/core/src/layout/engine.ts) can be
 * checked against ground truth instead of assumption.
 *
 * Each case is its own section (own w:sectPr, own w:lnNumType), 1in margins on
 * Letter (pgSz 12240x15840, pgMar 1440 all sides -> 9in / 12960tw text height).
 * Every real text line is marked with a unique run so the reader can pair a
 * margin number's y-position with the y-position of the line it names.
 *
 *   A-SINGLE    default (auto) single spacing, countBy=1 — how many real
 *               lines actually fit a page at plain 12pt spacing (the AI's
 *               naive "just turn on line numbers" case).
 *   B-EXACT24   spacing exact 480tw (24pt), spacing before/after=0 — the
 *               pleading-paper double-spaced grid; ~27 lines should fit.
 *   C-COUNTBY5  same as B, countBy=5 — only every 5th line prints, but does
 *               the counter still advance on every real line underneath?
 *   D-CONTINUOUS same as B, restart=continuous, enough lines for 2 pages —
 *               does the number continue past the page break?
 *   E-EMPTY     same as B, alternating text/empty paragraphs — does a blank
 *               paragraph consume a line number?
 *   F-WRAP      same as B, one long paragraph that wraps into several visual
 *               lines flanked by short marked lines — are wrapped
 *               continuation lines numbered individually?
 *   G-TABLE     same as B, a marked paragraph, a 3-row table with marked cell
 *               text, then a marked paragraph — does table content consume
 *               line numbers, and does the counter resume correctly after?
 *   H-START10   same as B, start=10 — first printed number.
 *
 * node scripts/generate-linenum-probe.mjs
 */

import { readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
// wild2-legal-ca-agreement: compat 15 (the engine's authoring default), Times
// New Roman docDefaults, no "Line Number" style override to confound the
// number's font with the pleading fixture's own styling.
const source = join(root, "apps/demo/public/fixtures/wild2-legal-ca-agreement.docx");
const parts = unzipSync(new Uint8Array(readFileSync(source)));

const pgProps =
  '<w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '<w:cols w:space="720"/>';

const lnNumAttrs = ({ countBy = 1, restart = "newPage", start = 1 } = {}) =>
  `w:countBy="${countBy}"` +
  (restart !== "newPage" ? ` w:restart="${restart}"` : "") +
  (start !== 1 ? ` w:start="${start}"` : "");

const sectPr = (opts) => `<w:sectPr><w:type w:val="nextPage"/>${pgProps}<w:lnNumType ${lnNumAttrs(opts)}/></w:sectPr>`;

// A marked single-line paragraph: "<id><n> " then filler so it does not wrap.
const line = (mark, exact) =>
  `<w:p><w:pPr>${exact ? '<w:spacing w:before="0" w:after="0" w:line="480" w:lineRule="exact"/>' : ""}</w:pPr>` +
  `<w:r><w:t xml:space="preserve">${mark}</w:t></w:r></w:p>`;
const empty = (exact) =>
  `<w:p><w:pPr>${exact ? '<w:spacing w:before="0" w:after="0" w:line="480" w:lineRule="exact"/>' : ""}</w:pPr></w:p>`;

// ---- A-SINGLE: default spacing, 40 short marked lines ----
const caseA = Array.from({ length: 40 }, (_, i) => line(`ASGL${i + 1}`, false)).join("");

// ---- B-EXACT24: exact double spacing, 40 short marked lines (spans 2 pages) ----
const caseB = Array.from({ length: 40 }, (_, i) => line(`BEXT${i + 1}`, true)).join("");

// ---- C-COUNTBY5: exact double spacing, countBy=5, 40 lines ----
const caseC = Array.from({ length: 40 }, (_, i) => line(`CCBY${i + 1}`, true)).join("");

// ---- D-CONTINUOUS: exact double spacing, restart=continuous, 60 lines (2+ pages) ----
const caseD = Array.from({ length: 60 }, (_, i) => line(`DCNT${i + 1}`, true)).join("");

// ---- E-EMPTY: exact double spacing, alternating text/empty, 15 text + 15 empty ----
const caseE = Array.from({ length: 15 }, (_, i) => line(`EEMP${i + 1}`, true) + empty(true)).join("");

// ---- F-WRAP: exact double spacing; short marked line, one long wrapping
// paragraph (marked at its start), short marked line ----
const longText =
  "FWRAPSTART " +
  Array.from({ length: 40 }, () => "wordword").join(" ") +
  " FWRAPEND";
const caseF =
  line("FPRE1", true) +
  line("FPRE2", true) +
  `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="480" w:lineRule="exact"/></w:pPr>` +
  `<w:r><w:t xml:space="preserve">${longText}</w:t></w:r></w:p>` +
  line("FPOST1", true) +
  line("FPOST2", true);

// ---- G-TABLE: exact double spacing; marked line, 3-row table with marked
// cell text, marked line ----
const tblRow = (mark) =>
  `<w:tr><w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>${line(mark, true)}</w:tc></w:tr>`;
const caseG =
  line("GPRE1", true) +
  line("GPRE2", true) +
  `<w:tbl><w:tblPr><w:tblW w:w="4500" w:type="dxa"/><w:tblBorders>` +
  ["top", "left", "bottom", "right", "insideH", "insideV"]
    .map((e) => `<w:${e} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`)
    .join("") +
  `</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="4500"/></w:tblGrid>` +
  tblRow("GROW1") +
  tblRow("GROW2") +
  tblRow("GROW3") +
  `</w:tbl>` +
  line("GPOST1", true) +
  line("GPOST2", true);

// ---- H-START10: exact double spacing, start=10, 10 short marked lines ----
const caseH = Array.from({ length: 10 }, (_, i) => line(`HSTA${i + 1}`, true)).join("");

const cases = [
  { id: "A", body: caseA, opts: {} },
  { id: "B", body: caseB, opts: {} },
  { id: "C", body: caseC, opts: { countBy: 5 } },
  { id: "D", body: caseD, opts: { restart: "continuous" } },
  { id: "E", body: caseE, opts: {} },
  { id: "F", body: caseF, opts: {} },
  { id: "G", body: caseG, opts: {} },
  { id: "H", body: caseH, opts: { start: 10 } },
];

let body = "";
cases.forEach((c, i) => {
  body += c.body;
  if (i < cases.length - 1) {
    // Section break carries the outgoing section's sectPr on an empty final
    // paragraph, per the docgrid probes' convention.
    body += `<w:p><w:pPr>${sectPr(c.opts)}</w:pPr></w:p>`;
  } else {
    body += sectPr(c.opts);
  }
});

const documentXml =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
  '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
  `<w:body>${body}</w:body></w:document>`;

const out = join(root, "fixtures-staging/probe-linenum.docx");
writeFileSync(out, zipSync({ ...parts, "word/document.xml": strToU8(documentXml) }));
console.log("wrote", out, "-", cases.length, "cases");
