#!/usr/bin/env node

/**
 * probe-exactnil11 / probe-exactnil11p settled that Word zeroes a both-nil
 * boundary and charges a live rule to the content INSET only — exact-row FLOW
 * equals the authored height, in compat 11, in the pleading's own package,
 * with full borders crossing (END-REF identical across A0/A1/A2/D1/D2). Yet
 * wild3-template-us-courts-answer's caption rows sit ~0.6-0.75pt APART-er in
 * Word than authored heights predict, per boundary. Every probe row so far
 * was 495tw = 33px, taller than its 16px line; the pleading's problem rows
 * are 115tw (7.67px) and 144tw (9.6px) — SHORTER than the line they clip.
 *
 * This probe varies exactly that: exact/exact pairs at 115, 144, 361 and
 * 495tw in the pleading package, each with no rule / live insideH sz-8 /
 * both-nil sz-8, plus full-borders variants at 115tw. The observable is
 * END-REF — the whole table's flow — because a clipped row hides its own
 * mark. If short exact rows charge their rules into FLOW, S1 > S0; if the
 * both-nil boundary still zeroes there, S2 = S0; if Word floors a clipped
 * row at something other than its authored height, S0 itself says so.
 *
 *   node scripts/generate-exactclip-probe.mjs
 */

import { readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const source = join(root, "apps/demo/public/fixtures/wild3-template-us-courts-answer.docx");
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

const TBL_PR = (borders) =>
  '<w:tblPr><w:tblW w:w="9360" w:type="dxa"/>' +
  borders +
  '<w:tblLayout w:type="fixed"/>' +
  '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/>' +
  '<w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>';

const INSIDE_H8 = '<w:tblBorders><w:insideH w:val="single" w:sz="8" w:space="0" w:color="000000"/></w:tblBorders>';
const FULL8 =
  "<w:tblBorders>" +
  ["top", "left", "bottom", "right", "insideH", "insideV"]
    .map((s) => `<w:${s} w:val="single" w:sz="8" w:space="0" w:color="000000"/>`)
    .join("") +
  "</w:tblBorders>";
const NIL = (side) => `<w:tcBorders><w:${side} w:val="nil"/></w:tcBorders>`;

const EXACT = (tw) => `<w:trPr><w:trHeight w:hRule="exact" w:val="${tw}"/></w:trPr>`;

const table = (id, borders, tw, upper, lower) =>
  "<w:tbl>" +
  TBL_PR(borders) +
  '<w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid>' +
  `<w:tr>${EXACT(tw)}` +
  `<w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/>${upper}</w:tcPr>` +
  mark(`${id}UP`) +
  "</w:tc></w:tr>" +
  `<w:tr>${EXACT(tw)}` +
  `<w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/>${lower}</w:tcPr>` +
  mark(`${id}MK`) +
  "</w:tc></w:tr></w:tbl>";

const CASES = [];
for (const [pfx, tw] of [["S", 115], ["T", 144], ["U", 361], ["W", 495]]) {
  CASES.push({ id: `${pfx}0`, borders: "", tw, upper: "", lower: "" });
  CASES.push({ id: `${pfx}1`, borders: INSIDE_H8, tw, upper: "", lower: "" });
  CASES.push({ id: `${pfx}2`, borders: INSIDE_H8, tw, upper: NIL("bottom"), lower: NIL("top") });
}
CASES.push({ id: "F1", borders: FULL8, tw: 115, upper: "", lower: "" });
CASES.push({ id: "F2", borders: FULL8, tw: 115, upper: NIL("bottom"), lower: NIL("top") });

const body = CASES.map((c) => mark(`${c.id}REF`) + table(c.id, c.borders, c.tw, c.upper, c.lower) + mark(`${c.id}END`)).join(
  "",
);

parts["word/document.xml"] = strToU8(
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    `<w:body>${body}${SECT}</w:body></w:document>`,
);

const settings = strFromU8(parts["word/settings.xml"]);
if (!/compatibilityMode[^>]*w:val="11"/.test(settings)) throw new Error("pleading base is not compat 11");

const out = join(root, "fixtures-staging/probe-exactclip.docx");
writeFileSync(out, Buffer.from(zipSync(parts)));
console.log(`wrote ${out} — ${CASES.map((c) => c.id).join(", ")}`);
