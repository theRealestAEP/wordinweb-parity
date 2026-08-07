#!/usr/bin/env node

/**
 * How does Word charge a run of EMPTY paragraphs whose line rule is EXACT?
 *
 * wild3-template-caed-pleading p1 residual (#67's ~12.6px leftover): our body
 * sits +12.5..13.0px low from the Court heading down, while the header rail
 * and the body top (probe-negmargin re-measured: ours = Word + 0.4) are
 * right. Word's own PDF closes the whole page arithmetic - Court's exact-32px
 * line is bottom-anchored (baseline 25.66 into the line), spacing-after
 * 660tw = 44.0px is charged in full before the caption table, the caption
 * cell's exact-16 lines anchor at baseline 12.6 with the cell bottom rule at
 * 730.56 - EXCEPT one quantity: the twelve empty BodyText/BodyTextContinued
 * paragraphs (line 480 exact = 32.00px each) span 371.32px, not 384. Word
 * discounts 12.68px SOMEWHERE in that run of empty exact-line paragraphs; we
 * charge 12 x 32 exactly.
 *
 * Sections (own page each, fixture geometry w:top=-1325 w:left=2088 unless
 * noted; MK marker = Court-style line 480 exact, before/after 0):
 *
 *   E0   marker alone                       calibrates bodyTop + baseline
 *   E6   6 empty exact-480 + marker         slope of the empty charge
 *   E12  12 empty exact-480 + marker        the fixture's own count
 *   C12  12 exact-480 with "." content      content control - expect 32 each
 *   Q6   6 empty exact-240 + marker         second exact setting (16px rows)
 *   P12  12 empty exact-480, w:top=1440     margin-sign control
 *
 * Read from the Word PDF: the MK baseline per page. X = per-empty charge
 * falls out of (E12-E0)/12 and (E6-E0)/6; Q6 says whether a discount scales
 * with the exact height; C12 says whether emptiness is the trigger; P12 says
 * whether the negative margin is.
 *
 *   node scripts/generate-emptyexact-probe.mjs
 */

import { readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { unzipSync, zipSync } from "fflate";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const source = join(root, "apps/demo/public/fixtures/wild3-template-caed-pleading.docx");
const parts = unzipSync(new Uint8Array(readFileSync(source)));

const pgProps = (top) =>
  `<w:pgSz w:w="12240" w:h="15840" w:code="1"/>` +
  `<w:pgMar w:top="${top}" w:right="720" w:bottom="-1267" w:left="2088" w:header="432" w:footer="360" w:gutter="0"/>` +
  `<w:cols w:space="720"/>`;

// No header/footer references: header height is probed inert for the negative
// top margin (probe-negmargin), and leaving the rail out keeps the page clean.
const sectPr = (top) => `<w:sectPr>${pgProps(top)}</w:sectPr>`;

const empty = (line) =>
  `<w:p><w:pPr><w:spacing w:line="${line}" w:lineRule="exact"/></w:pPr></w:p>`;
const content = (line) =>
  `<w:p><w:pPr><w:spacing w:line="${line}" w:lineRule="exact"/></w:pPr><w:r><w:t>.</w:t></w:r></w:p>`;
const marker = (id, sz = 24, line = 480) =>
  `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="${line}" w:lineRule="exact"/>` +
  `<w:rPr><w:sz w:val="${sz}"/></w:rPr></w:pPr>` +
  `<w:r><w:rPr><w:sz w:val="${sz}"/></w:rPr><w:t xml:space="preserve">MK${id}X</w:t></w:r></w:p>`;

const rep = (n, s) => s.repeat(n);

const cases = [
  { id: "E0", body: marker("E0"), top: -1325 },
  { id: "E6", body: rep(6, empty(480)) + marker("E6"), top: -1325 },
  { id: "E12", body: rep(12, empty(480)) + marker("E12"), top: -1325 },
  { id: "C12", body: rep(12, content(480)) + marker("C12"), top: -1325 },
  { id: "Q6", body: rep(6, empty(240)) + marker("Q6"), top: -1325 },
  { id: "P12", body: rep(12, empty(480)) + marker("P12"), top: 1440 },
  // The 0.8 x font-size anchor claim needs a second and third size: a 9pt and
  // an 18pt marker in the same exact-480 line (predictions 9.6 / 19.2px from
  // the band top if the anchor scales with size; ~12.9 flat if it does not).
  { id: "M9", body: marker("M9", 18), top: -1325 },
  { id: "M18", body: marker("M18", 36), top: -1325 },
  // The anchor's functional form: 12pt text in exact 360/720 (24/48px bands).
  // A 0.4 x height anchor predicts 9.6 / 19.2px; a flat ~12.8px predicts
  // 12.8 / 12.8.
  { id: "X24", body: marker("X24", 24, 360), top: -1325 },
  { id: "X48", body: marker("X48", 24, 720), top: -1325 },
];

let body = "";
cases.forEach((c, i) => {
  body += c.body;
  if (i < cases.length - 1) body += `<w:p><w:pPr>${sectPr(c.top)}</w:pPr></w:p>`;
  else body += sectPr(c.top);
});

const documentXml =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
  `<w:body>${body}</w:body></w:document>`;

parts["word/document.xml"] = new TextEncoder().encode(documentXml);
// Drop header/footer parts and their relationships: document.xml no longer
// references them.
for (const name of Object.keys(parts)) {
  if (/word\/(header|footer)\d*\.xml/.test(name)) delete parts[name];
}
{
  const rels = new TextDecoder().decode(parts["word/_rels/document.xml.rels"]);
  parts["word/_rels/document.xml.rels"] = new TextEncoder().encode(
    rels.replace(/<Relationship [^>]*Target="(header|footer)\d*\.xml"[^>]*\/>/g, ""),
  );
  const ct = new TextDecoder().decode(parts["[Content_Types].xml"]);
  parts["[Content_Types].xml"] = new TextEncoder().encode(
    ct.replace(/<Override [^>]*\/word\/(header|footer)\d*\.xml[^>]*\/>/g, ""),
  );
}
const out = join(root, "fixtures-staging/probe-emptyexact.docx");
writeFileSync(out, zipSync(parts));
console.log("wrote", out);
