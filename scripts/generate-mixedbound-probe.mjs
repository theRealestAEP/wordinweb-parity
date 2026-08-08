#!/usr/bin/env node

/**
 * #108b: the MIXED exact/content interior boundary. probe-exactnil11's B1
 * showed Word charging the full insideH sz-12 rule to the CONTENT row below
 * an exact row (MK-UP 26.25, flow +1.52pt) where our engine charges half —
 * `exactInsetRow`'s comment calls the mixed case unmeasured, and B was the
 * only mixed shape probed (exact above, both-nil or live only).
 *
 * This probe varies the boundary one side at a time, both orders:
 *
 *   EC*  exact 495 above, content below     (B's shape, now with one-sided nils)
 *   CE*  content above, exact 495 below     (the inverse, never measured)
 *   CC0/CCR and EE0/EER  same-kind controls, re-run in-probe
 *
 * Per order: 0 no rule | R insideH sz-12 | N both-nil | U upper cell nil
 * only | L lower cell nil only | W insideH sz-24 (width scaling, mixed
 * orders only). Zero cell margins; only insideH is authored so no outer
 * edge confounds. Two packages differing only in compatibilityMode (15 as
 * authored, 11 rewritten).
 *
 * Observables per case: top(UP), top(MK), top(REF), top(END). MK-UP reads
 * the boundary's charge into the lower row's content position; END-REF
 * reads the total flow.
 *
 *   node scripts/generate-mixedbound-probe.mjs
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

const TBL_PR = (borders) =>
  '<w:tblPr><w:tblW w:w="9360" w:type="dxa"/>' +
  borders +
  '<w:tblLayout w:type="fixed"/>' +
  '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/>' +
  '<w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>';

const INSIDE_H = (sz) =>
  `<w:tblBorders><w:insideH w:val="single" w:sz="${sz}" w:space="0" w:color="000000"/></w:tblBorders>`;
const NIL = (side) => `<w:tcBorders><w:${side} w:val="nil"/></w:tcBorders>`;
const EXACT = '<w:trPr><w:trHeight w:hRule="exact" w:val="495"/></w:trPr>';

const table = (id, borders, upperTr, lowerTr, upperTc, lowerTc) =>
  "<w:tbl>" +
  TBL_PR(borders) +
  '<w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid>' +
  `<w:tr>${upperTr}` +
  `<w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/>${upperTc}</w:tcPr>` +
  mark(`${id}UP`) +
  "</w:tc></w:tr>" +
  `<w:tr>${lowerTr}` +
  `<w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/>${lowerTc}</w:tcPr>` +
  mark(`${id}MK`) +
  "</w:tc></w:tr></w:tbl>";

const ORDERS = {
  EC: [EXACT, ""],
  CE: ["", EXACT],
  CC: ["", ""],
  EE: [EXACT, EXACT],
};

const CASES = [];
for (const [ord, [up, lo]] of Object.entries(ORDERS)) {
  CASES.push({ id: `${ord}0`, borders: "", up, lo, upper: "", lower: "" });
  CASES.push({ id: `${ord}R`, borders: INSIDE_H(12), up, lo, upper: "", lower: "" });
  if (ord === "EC" || ord === "CE") {
    CASES.push({ id: `${ord}N`, borders: INSIDE_H(12), up, lo, upper: NIL("bottom"), lower: NIL("top") });
    CASES.push({ id: `${ord}U`, borders: INSIDE_H(12), up, lo, upper: NIL("bottom"), lower: "" });
    CASES.push({ id: `${ord}L`, borders: INSIDE_H(12), up, lo, upper: "", lower: NIL("top") });
    CASES.push({ id: `${ord}W`, borders: INSIDE_H(24), up, lo, upper: "", lower: "" });
  }
}

const body = CASES.map((c) =>
  mark(`${c.id}REF`) + table(c.id, c.borders, c.up, c.lo, c.upper, c.lower) + mark(`${c.id}END`),
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
  const out = join(root, `fixtures-staging/probe-mixedbound${compat}.docx`);
  writeFileSync(out, Buffer.from(zipSync(parts)));
  console.log(`wrote ${out} — ${CASES.map((c) => c.id).join(", ")}`);
}
