#!/usr/bin/env node

/**
 * Generate the probe for #51: does a `nil` CELL border still cost a row its
 * TABLE rule's width?
 *
 * Observed during the exact-row fix (engine b0b8e2f): `rowBorderWidths` takes
 * `max(tblBorders.insideH, cellBottom, cellTop)` per boundary. A cell border of
 * `w:val="nil"` is a real declaration that OVERRIDES the table rule, and
 * `paintCellEdges` already honours it and draws nothing — but the max above
 * reads the nil as "no opinion" rather than as zero, so the rule's width is
 * still charged to the row height and the content inset. Wherever a nil cell
 * border overrides a table rule we would then be one rule-width too tall while
 * painting no rule at all.
 *
 * Six variants of one three-row table, differing only in what the row 0 / row 1
 * boundary declares. Every row is `atLeast` with a small value so the CONTENT
 * governs the height and a charged border shows up as extra height rather than
 * being swallowed by an exact value. The measurement is the same mark-to-mark
 * distance `generate-exactrow-probe.mjs` uses, `top(row 2 mark) - top(row 0
 * mark)`, which is a pure layout quantity and never looks at paint.
 *
 *   A-none     tblBorders only; cells carry a width and nothing else
 *   B-nilboth  BOTH cells at the boundary declare nil (row 0 bottom, row 1 top)
 *   C-nilone   only row 0's cells declare bottom nil; row 1 says nothing
 *   D-norule   no tblBorders at all                       (zero-rule control)
 *   E-own12    row 0's cells declare bottom single sz=12  (same as the rule)
 *   F-own24    row 0's cells declare bottom single sz=24  (scaling check)
 *
 * D fixes what the geometry costs with no rule anywhere, so A - D is what one
 * table rule costs. E - A says whether a cell border that merely RESTATES the
 * rule adds anything. F - E is the scaling check: if what is charged is a border
 * width, doubling `w:sz` doubles the difference, and if it is a constant it does
 * not. B and C are the question itself, and C separately answers whether Word
 * needs BOTH sides of a shared boundary to say nil before it suppresses.
 *
 * Read every case against D, not against A: A already contains whatever a rule
 * costs, and reading B against A alone cannot tell "the nil was honoured" from
 * "the rule was never charged here in the first place".
 *
 *   node scripts/generate-nilborder-probe.mjs
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

const EDGES = ["top", "left", "bottom", "right", "insideH", "insideV"];
const BORDERS =
  "<w:tblBorders>" +
  EDGES.map((e) => `<w:${e} w:val="single" w:sz="12" w:space="0" w:color="auto"/>`).join("") +
  "</w:tblBorders>";

const tblPr = (borders) =>
  '<w:tblPr><w:tblW w:w="10080" w:type="dxa"/>' +
  (borders ? BORDERS : "") +
  '<w:tblLayout w:type="fixed"/>' +
  '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/>' +
  '<w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>';

/** One 10 pt line, single spaced, so a row's content height is one known line. */
const line = (text) =>
  '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>' +
  `<w:rPr><w:b/><w:sz w:val="20"/></w:rPr></w:pPr>` +
  `<w:r><w:rPr><w:b/><w:sz w:val="20"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;

const tcPr = (w, borders) =>
  `<w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${borders ?? ""}</w:tcPr>`;
const cell = (w, body, borders) => `<w:tc>${tcPr(w, borders)}${body}</w:tc>`;

/** atLeast with a small value: the content governs, so a charged border shows. */
const row = (bodies, borders) =>
  '<w:tr><w:trPr><w:trHeight w:hRule="atLeast" w:val="100"/></w:trPr>' +
  cell(5040, bodies[0], borders?.[0]) +
  cell(5040, bodies[1], borders?.[1]) +
  "</w:tr>";

const bd = (edge, val, sz) =>
  `<w:tcBorders><w:${edge} w:val="${val}"${val === "nil" ? "" : ` w:sz="${sz}" w:space="0" w:color="auto"`}/></w:tcBorders>`;
const both = (x) => [x, x];

const cases = [
  { id: "A-none", borders: true, r0: undefined, r1: undefined },
  { id: "B-nilboth", borders: true, r0: both(bd("bottom", "nil")), r1: both(bd("top", "nil")) },
  { id: "C-nilone", borders: true, r0: both(bd("bottom", "nil")), r1: undefined },
  { id: "D-norule", borders: false, r0: undefined, r1: undefined },
  { id: "E-own12", borders: true, r0: both(bd("bottom", "single", 12)), r1: undefined },
  { id: "F-own24", borders: true, r0: both(bd("bottom", "single", 24)), r1: undefined },
];

const table = ({ id, borders, r0, r1 }) =>
  `<w:tbl>${tblPr(borders)}` +
  row([line(`${id}TOP`), line(`${id}top2`)], r0) +
  row([line(`${id}MID`), line(`${id}mid2`)], r1) +
  row([line(`${id}BOT`), line(`${id}bot2`)], undefined) +
  "</w:tbl>";

/** A labelled paragraph between tables keeps consecutive tables from merging. */
const spacer = (id) =>
  `<w:p><w:pPr><w:spacing w:before="0" w:after="0"/><w:rPr><w:sz w:val="20"/></w:rPr></w:pPr>` +
  `<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>gap ${id}</w:t></w:r></w:p>`;

const body = cases.map((c, i) => table(c) + spacer(i)).join("");

const documentXml =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  `<w:body>${body}${SECT}</w:body></w:document>`;

parts["word/document.xml"] = strToU8(documentXml);

// Probes live in fixtures-staging, like every other probe with a Word
// reference. parity-parallel.mjs adopts a reference only when a DOCX of
// the same name sits in apps/demo/public/fixtures, so keeping it out of
// there is what stops the full corpus run from picking the probe up.
// To read it in the browser: copy it into apps/demo/public/fixtures,
// measure, then remove it again.
const out = join(root, "fixtures-staging/probe-nilborder.docx");
writeFileSync(out, Buffer.from(zipSync(parts)));
console.log(`wrote ${out}`);
console.log(`cases: ${cases.map((c) => c.id).join(", ")}`);
