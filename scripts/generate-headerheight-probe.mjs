#!/usr/bin/env node

/**
 * Generate the header-height sweep for #72.
 *
 * `generate-negmargin-probe.mjs` measured, as a by-product of #67, that above a
 * zero top margin we put the body about 30.5 px lower than Word: with its short
 * header Word's body top is 236.87 px and ours 267.59, and at `w:header="1440"`
 * the same gap reappears (304.24 against 334.59). That probe could not say WHICH
 * component of `measureHeaderFooter` overcharges, because it never varied the
 * header's own composition — it only ever used ONE header shape, a single-row
 * table with an exact `trHeight`.
 *
 * That shape matters. `measureHeaderFooter` ends with
 *
 *     Math.max(height, contentBottom) +
 *       (!hasOnlyUnwrappedAnchors && complexHeader ? ptToPx(22.5) : 0)
 *
 * and `complexHeader` is true when the header holds a TABLE or a positioned
 * shape. `ptToPx(22.5)` is exactly 30.0 px. So the negmargin header takes that
 * clearance by construction, and the measured overcharge is 30.4-30.7 — which
 * makes the clearance a candidate the earlier probe could not have seen, and it
 * has to be separated from the named suspect (the header's trailing
 * space-after) rather than assumed.
 *
 * Six sections, each starting `nextPage`, each with `w:top="0"` so the HEADER
 * governs the body top, each with its own header part, differing in one thing:
 *
 *   P1   one plain paragraph, w:after=0                    (baseline)
 *   P2   two plain paragraphs, w:after=0        line count, 2 settings with P1/P3
 *   P3   three plain paragraphs, w:after=0
 *   S2   one plain paragraph, w:after=200 tw (10 pt)       trailing space-after,
 *   S4   one plain paragraph, w:after=400 tw (20 pt)       2 settings with P1
 *   TB   one table row, exact trHeight 240 tw              table vs paragraph,
 *                                                          content height held equal
 *
 * Every header line is `w:line="240" w:lineRule="exact"` — exactly 12 pt, 16.00
 * CSS px — and TB's row is an exact 240 tw, so TB and P1 hold the SAME content
 * height and differ only in being a table. Read the body top per case:
 *
 *   - if the overcharge tracks the line count, P1/P2/P3 diverge progressively;
 *   - if it is the trailing space-after, P1/S2/S4 diverge and by 13.33 then
 *     26.67 px, which is what those two settings are for;
 *   - if it is the 22.5 pt complex-header clearance, ONLY TB diverges, by 30.0.
 *
 * The three are mutually exclusive on this design, so one run decides it.
 *
 *   node scripts/generate-headerheight-probe.mjs
 */

import { readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const source = join(root, "apps/demo/public/fixtures/wild3-template-caed-pleading.docx");
const parts = unzipSync(new Uint8Array(readFileSync(source)));

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

/** One header line of exactly 12 pt, so a header's height is countable. */
const hline = (tag, after = 0) =>
  `<w:p><w:pPr><w:spacing w:before="0" w:after="${after}" w:line="240" w:lineRule="exact"/>` +
  '<w:rPr><w:sz w:val="20"/></w:rPr></w:pPr>' +
  `<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>${tag}</w:t></w:r></w:p>`;

/** A table whose row is an exact 240 tw: the same content height as one hline. */
const htable = (tag) =>
  '<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/><w:tblLayout w:type="fixed"/>' +
  '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/>' +
  '<w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid>' +
  '<w:tr><w:trPr><w:trHeight w:hRule="exact" w:val="240"/></w:trPr>' +
  '<w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/></w:tcPr>' +
  hline(tag) +
  "</w:tc></w:tr></w:tbl>";

const HEADERS = [
  { id: "P1", body: (t) => hline(`${t}a`) },
  { id: "P2", body: (t) => hline(`${t}a`) + hline(`${t}b`) },
  { id: "P3", body: (t) => hline(`${t}a`) + hline(`${t}b`) + hline(`${t}c`) },
  { id: "S2", body: (t) => hline(`${t}a`, 200) },
  { id: "S4", body: (t) => hline(`${t}a`, 400) },
  { id: "TB", body: (t) => htable(`${t}a`) },
];

// Fresh header parts, relationships and content-type overrides. The base
// package's own header1-3 are left in place and simply go unreferenced.
let rels = strFromU8(parts["word/_rels/document.xml.rels"]);
let types = strFromU8(parts["[Content_Types].xml"]);
HEADERS.forEach((h, i) => {
  const name = `headerP72_${i}.xml`;
  h.rid = `rIdH72${i}`;
  parts[`word/${name}`] = strToU8(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
      `<w:hdr ${NS}>${h.body(`HD${h.id}`)}</w:hdr>`,
  );
  rels = rels.replace(
    "</Relationships>",
    `<Relationship Id="${h.rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="${name}"/></Relationships>`,
  );
  types = types.replace(
    "</Types>",
    `<Override PartName="/word/${name}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>`,
  );
});
parts["word/_rels/document.xml.rels"] = strToU8(rels);
parts["[Content_Types].xml"] = strToU8(types);

/** A marker of exactly known height, so its top IS the body top. */
const marker = (id) =>
  '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/>' +
  '<w:rPr><w:sz w:val="20"/></w:rPr></w:pPr>' +
  `<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>${id}</w:t></w:r></w:p>`;

/** w:top="0" so the header, not the margin, decides where the body starts. */
const sectPr = (h) =>
  `<w:sectPr><w:headerReference w:type="default" r:id="${h.rid}"/>` +
  '<w:type w:val="nextPage"/><w:pgSz w:w="12240" w:h="15840" w:code="1"/>' +
  '<w:pgMar w:top="0" w:right="720" w:bottom="1440" w:left="2088" ' +
  'w:header="720" w:footer="360" w:gutter="0"/>' +
  '<w:cols w:space="720"/></w:sectPr>';

const body = HEADERS.map((h, i) => {
  const mark = marker(`MK${h.id}X`);
  return i === HEADERS.length - 1 ? mark : `${mark}<w:p><w:pPr>${sectPr(h)}</w:pPr></w:p>`;
}).join("");

parts["word/document.xml"] = strToU8(
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
    `<w:document ${NS}><w:body>${body}${sectPr(HEADERS.at(-1))}</w:body></w:document>`,
);

// Probes live in fixtures-staging: parity-parallel.mjs adopts a reference only
// when a DOCX of the same name sits in apps/demo/public/fixtures.
const out = join(root, "fixtures-staging", "probe-headerheight.docx");
writeFileSync(out, Buffer.from(zipSync(parts)));
console.log(`wrote ${out} — ${HEADERS.map((h) => h.id).join(", ")}`);
