#!/usr/bin/env node

/**
 * Generate the probe that isolates wild2-med-phase23-protocol's extra page.
 *
 * Word lays that document in 69 pages and we lay 70, and pages 1 to 68 hold the
 * same content. The divergence is one fit decision at the foot of page 68.
 * After the section 10.4 table our render and Word's agree line for line — the
 * table ends at the same place and the empty paragraph after it sits at the same
 * place — and then Word puts ONE MORE empty paragraph on the page and we do not.
 * That paragraph carries `<w:br w:type="page"/>`, so spilling it costs a whole
 * page: its break then starts the heading a page later than Word starts it.
 *
 * The paragraph authors no `w:spacing` of its own, so it inherits this
 * document's `w:pPrDefault`, which is `before="200" after="200" line="276"`. At
 * 11 pt that is a 19.4 px line with 13.3 px above it and 13.3 px below. We had
 * 22.1 px of room left. Which of those three the fit test has to accommodate
 * decides the page count, and that is the one thing this probe measures.
 *
 * Each variant fills a page with exact-height paragraphs, then a shim tuned so
 * an exact amount of room is left, then the target paragraph — the same shape
 * as the fixture's, spacing inherited. The room sweeps 18 to 45 CSS px:
 *
 *   line only            needs 19.4 px  -> fits from room 21
 *   space-before + line  needs 32.7 px  -> fits from room 33
 *   before + line + after needs 46.1 px -> fits at no room in this sweep
 *
 * So the room at which the target stops spilling names the rule, separately for
 * Word and for us. Built on the phase23 package so the docDefaults that make the
 * question meaningful come with it.
 *
 *   node scripts/generate-pagefit-probe.mjs
 */

import { readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { strToU8, unzipSync, zipSync } from "fflate";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const source = join(root, "apps/demo/public/fixtures/wild2-med-phase23-protocol.docx");
const parts = unzipSync(new Uint8Array(readFileSync(source)));

/** No header or footer, so the body is exactly 96..960 CSS px. */
const SECT =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '<w:cols w:space="720"/></w:sectPr>';

const RPR = '<w:rPr><w:sz w:val="22"/></w:rPr>';

/** A paragraph of an exactly known height: no space either side, line exact. */
const exact = (twips, text, breakBefore = false) =>
  "<w:p><w:pPr>" +
  (breakBefore ? "<w:pageBreakBefore/>" : "") +
  `<w:spacing w:before="0" w:after="0" w:line="${twips}" w:lineRule="exact"/>${RPR}</w:pPr>` +
  (text ? `<w:r>${RPR}<w:t>${text}</w:t></w:r>` : "") +
  "</w:p>";

/**
 * The paragraph under test: exactly the fixture's shape. No w:spacing at all,
 * so before, after and line all come from w:pPrDefault.
 */
const target = (text, pageBreak = false) =>
  `<w:p><w:pPr>${RPR}</w:pPr><w:r>${RPR}<w:t>${text}</w:t></w:r>` +
  (pageBreak ? `<w:r>${RPR}<w:br w:type="page"/></w:r>` : "") +
  "</w:p>";

const BODY_PX = 960 - 96;
const FILLER_TWIPS = 480; // 24 pt = 32 CSS px
const FILLER_PX = 32;
const FILLERS = 25; // 800 px, leaving 64 px for the shim plus the room

const rooms = [18, 21, 24, 27, 30, 33, 36, 45];

function sweep(tag, pageBreak) {
  return rooms
  .map((room, idx) => {
    const id = `${tag}${String(idx + 1).padStart(2, "0")}`;
    const shimPx = BODY_PX - FILLERS * FILLER_PX - room;
    if (shimPx <= 0) throw new Error(`room ${room} leaves no shim`);
    let out = "";
    for (let n = 0; n < FILLERS; n++) {
      out += exact(FILLER_TWIPS, n === FILLERS - 1 ? `${id}LAST` : "", n === 0);
    }
    out += exact(Math.round(shimPx * 15), `${id}SHIM`);
    out += target(`${id}TGT`, pageBreak);
    return out;
  })
  .join("");
}

/**
 * The second sweep is the first with ONE byte added: the target paragraph also
 * carries `<w:br w:type="page"/>`, exactly as phase23's block 1395 does. If the
 * threshold moves between the two sweeps, the break is what changes the fit.
 */
/**
 * The third sweep replaces the target with phase23's block 1395 verbatim: an
 * EMPTY paragraph whose only run is the page break, carrying that paragraph's
 * own run properties. It has no text, so it is read indirectly — a marker
 * paragraph follows it, and the marker lands one page after the filled page
 * when the target fits on it and two pages after when the target spills.
 */
const FIXTURE_RPR =
  "<w:rPr><w:b/><w:bCs/><w:caps/><w:color w:val=\"FFFFFF\" w:themeColor=\"background1\"/>" +
  '<w:spacing w:val="15"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>';

const emptyBreak = () =>
  `<w:p><w:pPr>${FIXTURE_RPR}</w:pPr><w:r>${FIXTURE_RPR}<w:br w:type="page"/></w:r></w:p>`;

function sweepEmpty(tag) {
  return rooms
    .map((room, idx) => {
      const id = `${tag}${String(idx + 1).padStart(2, "0")}`;
      const shimPx = BODY_PX - FILLERS * FILLER_PX - room;
      let out = "";
      for (let n = 0; n < FILLERS; n++) {
        out += exact(FILLER_TWIPS, n === FILLERS - 1 ? `${id}LAST` : "", n === 0);
      }
      out += exact(Math.round(shimPx * 15), `${id}SHIM`);
      out += emptyBreak();
      out += exact(FILLER_TWIPS, `${id}AFTER`);
      return out;
    })
    .join("");
}

/**
 * The fourth sweep drops the break and keeps the emptiness, so the two are not
 * confounded. With no break the marker lands on the next page either way, but
 * its TOP says which: at the body top when the empty paragraph stayed behind,
 * and one empty paragraph lower when it came along.
 */
function sweepEmptyNoBreak(tag) {
  return rooms
    .map((room, idx) => {
      const id = `${tag}${String(idx + 1).padStart(2, "0")}`;
      const shimPx = BODY_PX - FILLERS * FILLER_PX - room;
      let out = "";
      for (let n = 0; n < FILLERS; n++) {
        out += exact(FILLER_TWIPS, n === FILLERS - 1 ? `${id}LAST` : "", n === 0);
      }
      out += exact(Math.round(shimPx * 15), `${id}SHIM`);
      out += `<w:p><w:pPr>${FIXTURE_RPR}</w:pPr></w:p>`;
      out += exact(FILLER_TWIPS, `${id}AFTER`);
      return out;
    })
    .join("");
}

const body = sweep("P", false) + sweep("Q", true) + sweepEmpty("S") + sweepEmptyNoBreak("T");

const documentXml =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
  '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
  `<w:body>${body}${SECT}</w:body></w:document>`;

const path = join(root, "fixtures-staging", "probe-pagefit.docx");
writeFileSync(path, zipSync({ ...parts, "word/document.xml": strToU8(documentXml) }));
console.log(`Wrote ${path} — rooms ${rooms.join(", ")} CSS px`);
