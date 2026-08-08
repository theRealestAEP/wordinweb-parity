#!/usr/bin/env node

/**
 * Round 2 of probe-docgrid15: bracket Word's MS Mincho NATURAL line height,
 * the quantity the compat-15 lines-grid snap tests. probe-docgrid15 pinned
 * E360 (ja 11pt, pitch 360tw = 24px -> 1 row) and F360 (ja 16pt -> 2 rows),
 * which only brackets the em in [1.125, 1.636]. Our substituted Hiragino
 * raw profile reads 1.643em (ja 11pt raw box 24.10px), which is why the
 * textSnap EA carve-out exists; the natural-pitch contexts use a measured
 * 1.296em, and probe2-ruby-vertical's vertical columns read 1.4em. This
 * sweep separates those candidates: each case is one 3-line ja paragraph
 * (MS Mincho, two <w:br/>) under its own lines-grid pitch, and whether the
 * advance reads 1 or 2 grid rows brackets em against pitchPx/sizePx:
 *
 *   case      size  pitch   px    em threshold   1.296 says   1.4 says
 *   J11P240   11pt  240tw  16.0      1.091          2 rows      2 rows
 *   J16P360   16pt  360tw  24.0      1.125          2 rows      2 rows
 *   J10P240   10pt  240tw  16.0      1.200          2 rows      2 rows
 *   J12P300   12pt  300tw  20.0      1.250          2 rows      2 rows
 *   J16P420   16pt  420tw  28.0      1.3125         1 row       2 rows
 *   J11P300   11pt  300tw  20.0      1.364          1 row       2 rows
 *   J10P300   10pt  300tw  20.0      1.500          1 row       1 row
 *   J12P360   12pt  360tw  24.0      1.500          1 row       1 row
 *   J16P480   16pt  480tw  32.0      1.500          1 row       1 row
 *   J11P360   11pt  360tw  24.0      1.636          1 row       1 row
 *
 *   node scripts/generate-docgrid15b-probe.mjs
 */

import { readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { unzipSync, zipSync } from "fflate";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
// staging-eastasian's own package: compat 15 settings, the styles under test.
const source = join(root, "fixtures-staging/staging-eastasian.docx");
const parts = unzipSync(new Uint8Array(readFileSync(source)));

const pgProps =
  '<w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '<w:cols w:space="720"/>';

const sectPr = (pitch) =>
  `<w:sectPr><w:type w:val="nextPage"/>${pgProps}<w:docGrid w:type="lines" w:linePitch="${pitch}"/></w:sectPr>`;

const latinRun = (text, family, halfPts) =>
  `<w:r><w:rPr><w:rFonts w:ascii="${family}" w:hAnsi="${family}"/><w:sz w:val="${halfPts}"/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;
const jaRun = (text, halfPts) =>
  `<w:r><w:rPr><w:rFonts w:eastAsia="MS Mincho" w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="${halfPts}"/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;
const jaBr = (halfPts) =>
  `<w:r><w:rPr><w:rFonts w:eastAsia="MS Mincho" w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="${halfPts}"/></w:rPr><w:br/></w:r>`;

const jaCase = (id, halfPts) =>
  `<w:p><w:pPr><w:kinsoku/></w:pPr>` +
  latinRun(`${id} `, "Calibri", 22) +
  jaRun("水は方円の器に隨う。", halfPts) +
  jaBr(halfPts) +
  jaRun("山は静かにして大なり。", halfPts) +
  jaBr(halfPts) +
  jaRun("海は深くして広し。", halfPts) +
  `</w:p>` +
  `<w:p>${latinRun(`${id}F follower`, "Calibri", 22)}</w:p>`;

const sections = [
  { body: jaCase("J11P240", 22), pitch: 240 },
  { body: jaCase("J16P360", 32), pitch: 360 },
  { body: jaCase("J10P240", 20), pitch: 240 },
  { body: jaCase("J12P300", 24), pitch: 300 },
  { body: jaCase("J16P420", 32), pitch: 420 },
  { body: jaCase("J11P300", 22), pitch: 300 },
  { body: jaCase("J10P300", 20), pitch: 300 },
  { body: jaCase("J12P360", 24), pitch: 360 },
  { body: jaCase("J16P480", 32), pitch: 480 },
  { body: jaCase("J11P360", 22), pitch: 360 },
];

let body = "";
sections.forEach((s, i) => {
  body += s.body;
  if (i < sections.length - 1) {
    body += `<w:p><w:pPr>${sectPr(s.pitch)}</w:pPr></w:p>`;
  } else {
    body += sectPr(s.pitch);
  }
});

const documentXml =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
  `<w:body>${body}</w:body></w:document>`;

parts["word/document.xml"] = new TextEncoder().encode(documentXml);
const out = join(root, "fixtures-staging/probe-docgrid15b.docx");
writeFileSync(out, zipSync(parts));
console.log("wrote", out);
