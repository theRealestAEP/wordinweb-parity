#!/usr/bin/env node

/**
 * #108a follow-up 2: how much may a cantSplit table row overhang the body
 * bottom before Word moves it?
 *
 * Our engine grants a compat-11 cantSplit row (heightRule unset or a real
 * atLeast) an overhang allowance of the FOOTER HEIGHT — a July-window
 * fixture calibration with no probe behind it. us-courts-answer p6/p7 now
 * sit on that knife edge: Word moves a signature row whose overhang would
 * be ~19px, and whether we move it too is decided by that fudge to within
 * a pixel. This probe measures the allowance directly.
 *
 * Each case is its own section on the us-courts package (compat 11, Times
 * fonts): 40 exact 240tw filler lines (16.00 px each), a shim paragraph
 * whose exact line height tunes the ROOM left above the body bottom, then a
 * one-row table (two exact 16 px lines in one cell, no borders, zero
 * margins — the row needs 32.00 px), then a marker. If the row's first
 * text lands on the filler's page the row STAYED; on the next page it
 * MOVED. Sweeps:
 *
 *   K16..K40  cantSplit row, no footer,   room 16..40 px step 2
 *   G16..G40  cantSplit row, WITH footer, room 16..40 px step 2
 *   P26..P36  plain row (no cantSplit), no footer, room 26..36 px step 2
 *
 * A stay-threshold at room 32 means NO allowance; every 2 px below 32 the
 * row still stays measures 2 px of allowance. The footer sweep separates
 * "allowance = footer height" from "allowance is footer-independent".
 *
 *   node scripts/generate-rowfit-probe.mjs
 */

import { readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const source = join(root, "apps/demo/public/fixtures/wild3-template-us-courts-answer.docx");

// Body: pgSz 15840 - 2x1440 margins = 12960 tw = 864 px.
const SECT = (footer) =>
  "<w:sectPr>" +
  (footer ? '<w:footerReference w:type="default" r:id="rId9"/>' : "") +
  '<w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '<w:cols w:space="720"/></w:sectPr>';

const line = (text, twips = 240) =>
  `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="${twips}" w:lineRule="exact"/>` +
  '<w:rPr><w:sz w:val="20"/></w:rPr></w:pPr>' +
  `<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;

const table = (id, cantSplit) =>
  "<w:tbl><w:tblPr>" +
  '<w:tblW w:w="9360" w:type="dxa"/><w:tblLayout w:type="fixed"/>' +
  '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/>' +
  '<w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid>' +
  `<w:tr>${cantSplit ? "<w:trPr><w:cantSplit/></w:trPr>" : ""}` +
  '<w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/></w:tcPr>' +
  line(`${id}T1`) +
  line(`${id}T2`) +
  "</w:tc></w:tr></w:tbl>";

const FILLERS = 40; // 640 px

const cases = [];
for (const room of [16, 18, 20, 22, 24, 26, 28, 30, 32, 34, 36, 38, 40]) {
  cases.push({ id: `K${room}`, room, cantSplit: true, footer: false });
  cases.push({ id: `G${room}`, room, cantSplit: true, footer: true });
}
for (const room of [26, 28, 30, 32, 34, 36]) {
  cases.push({ id: `P${room}`, room, cantSplit: false, footer: false });
}

const body =
  cases
    .map((c, i) => {
      // room = 864 - 640 - shim  ->  shim = 224 - room (px); 1 px = 15 tw.
      const shimTw = (224 - c.room) * 15;
      let s = "";
      for (let f = 1; f <= FILLERS; f++) s += line(`${c.id}F${f}`);
      s += line(`${c.id}S`, shimTw);
      s += table(c.id, c.cantSplit);
      s += line(`${c.id}MK`);
      if (i < cases.length - 1) s += `<w:p><w:pPr>${SECT(c.footer)}</w:pPr></w:p>`;
      else s += SECT(c.footer);
      return s;
    })
    .join("");

const parts0 = unzipSync(new Uint8Array(readFileSync(source)));
const srcDoc = strFromU8(parts0["word/document.xml"]);
const docOpen = srcDoc.slice(srcDoc.indexOf("<w:document"), srcDoc.indexOf(">", srcDoc.indexOf("<w:document")) + 1);

const documentXml =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + docOpen + `<w:body>${body}</w:body></w:document>`;

const parts = unzipSync(new Uint8Array(readFileSync(source)));
parts["word/document.xml"] = strToU8(documentXml);
const out = join(root, "fixtures-staging/probe-rowfit11.docx");
writeFileSync(out, Buffer.from(zipSync(parts)));
console.log(`wrote ${out} — ${cases.length} cases`);
