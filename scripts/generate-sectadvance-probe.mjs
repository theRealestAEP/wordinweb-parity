#!/usr/bin/env node

/**
 * Generate the probe that counts PAGE ADVANCES at a break-only paragraph that
 * also carries a `w:sectPr`.
 *
 * Two documents hold that exact construct — an empty paragraph whose only run
 * is `<w:br w:type="page"/>`, with a `w:sectPr` in its `w:pPr` — and current
 * Word paginates them differently against us. `wild2-legal-ca-agreement`
 * paragraph 24 and `wild2-med-nccih-protocol` paragraph 17 are byte-for-byte
 * the same shape, so the difference has to be somewhere else in the document.
 *
 * The candidate this probe was built to test is the START TYPE OF THE SECTION
 * THAT FOLLOWS. Per ECMA-376 17.6.22 a `w:type` states how the section
 * CONTAINING it begins, so the start type of the section after the break-only
 * paragraph is authored on the NEXT `w:sectPr`, not on the paragraph's own.
 * Read that way the two fixtures do differ — ca-agreement's next section starts
 * `nextPage` and nccih's starts `continuous` — and a `nextPage` start is itself
 * a page advance, so ca-agreement's construct appears to offer two advances
 * where nccih's offers one.
 *
 * IT IS NOT THE VARIABLE. Word takes ONE advance for a page break followed by a
 * `nextPage` section start, our layout already does the same, and the two
 * fixtures diverge on the FIT test instead: they sit at 28 px and 543 px of
 * room, so only one of them is anywhere near a threshold. The probe is kept
 * because a swept variable that turns out to be inert is worth being able to
 * re-measure, and because the same sweep pins the fit demand that IS the rule.
 * See "One break, one advance" in scripts/README.md.
 *
 * Each case fills a page with exact-height paragraphs, tunes a shim so an exact
 * room is left, places the target paragraph carrying a `w:sectPr`, and then a
 * MARKER. The marker's page minus the last filler's page IS the advance count,
 * and it is only unambiguous when the target certainly fitted — which is why
 * the sweep carries a large room as well as a tight one. Four target shapes
 * separate the break from the emptiness; four following-section start types
 * sweep the variable.
 *
 * Case ids are `<shape><type><room>`, e.g. `BKNP200`. Every case is readable
 * from a Word PDF with pdftotext and from the browser by text search.
 *
 *   node scripts/generate-sectadvance-probe.mjs
 */

import { readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { strToU8, unzipSync, zipSync } from "fflate";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const source = join(root, "apps/demo/public/fixtures/wild2-med-phase23-protocol.docx");
const parts = unzipSync(new Uint8Array(readFileSync(source)));

/** No header or footer, so the body is exactly 96..960 CSS px. */
const pgProps =
  '<w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '<w:cols w:space="720"/>';

/**
 * `w:type` states how the section carrying it BEGINS. The filler section always
 * begins `nextPage` so each case starts on a fresh page and cases cannot
 * interact; the following section's start type is the swept variable and is
 * authored on the sectPr that terminates THAT section.
 */
const sectPr = (type) => `<w:sectPr><w:type w:val="${type}"/>${pgProps}</w:sectPr>`;

/** ca-agreement p24's own run properties, so the target matches the fixture. */
const rpr = (halfPoints) =>
  '<w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorHAnsi"/>' +
  '<w:b/><w:bCs/><w:color w:val="000000" w:themeColor="text1"/><w:kern w:val="28"/>' +
  `<w:sz w:val="${halfPoints}"/></w:rPr>`;

const FILLER_RPR = '<w:rPr><w:sz w:val="22"/></w:rPr>';

/** A paragraph of an exactly known height: no space either side, line exact. */
const exact = (twips, text) =>
  "<w:p><w:pPr>" +
  `<w:spacing w:before="0" w:after="0" w:line="${twips}" w:lineRule="exact"/>${FILLER_RPR}</w:pPr>` +
  (text ? `<w:r>${FILLER_RPR}<w:t>${text}</w:t></w:r>` : "") +
  "</w:p>";

/**
 * The target: ca-agreement p24's shape, with the section break on it. `text`
 * and `pageBreak` are the two things separated, so emptiness and the break are
 * never confounded.
 */
const target = (text, pageBreak, sz) => {
  const RPR = rpr(sz);
  return (
    `<w:p><w:pPr>${RPR}${sectPr("nextPage")}</w:pPr>` +
    (text ? `<w:r>${RPR}<w:t>${text}</w:t></w:r>` : "") +
    (pageBreak ? `<w:r>${RPR}<w:br w:type="page"/></w:r>` : "") +
    "</w:p>"
  );
};

const BODY_PX = 960 - 96;
const FILLER_TWIPS = 480; // 24 pt = 32 CSS px
const FILLER_PX = 32;
const FILLERS = 20; // 640 px, leaving 224 px for the shim plus the room

/**
 * `sz` is in half-points. 20 is ca-agreement p24's own size; `B2` repeats the
 * fixture construct at 40 and is the SCALING control — a demand that is a line
 * height doubles with the size, a demand that is a constant does not, and one
 * setting cannot tell those apart.
 */
const shapes = [
  { tag: "BK", text: "", pageBreak: true, sz: 20 }, // the fixture construct: empty, break only
  { tag: "TB", text: "TEXT", pageBreak: true, sz: 20 }, // text and a break
  { tag: "TO", text: "TEXT", pageBreak: false, sz: 20 }, // text, no break: the section advance alone
  { tag: "EO", text: "", pageBreak: false, sz: 20 }, // empty, no break
  { tag: "B2", text: "", pageBreak: true, sz: 40 }, // break only at twice the size
];

const types = [
  { tag: "NP", val: "nextPage" },
  { tag: "CN", val: "continuous" },
  { tag: "EV", val: "evenPage" },
  { tag: "OD", val: "oddPage" },
];

/**
 * Rooms in CSS px left under the target.
 *
 * 200 is the load-bearing setting and every case carries it: the target
 * certainly fits, so marker page minus filler page IS the advance count with
 * nothing else in it. The tighter rooms sweep the FIT threshold, where a spill
 * adds an advance of its own — read an advance count there without the room-200
 * row beside it and the two questions confound, which is the trap sweep S fell
 * into.
 *
 * `nextPage` and `continuous` get the whole sweep. `evenPage` and `oddPage` get
 * room 200 only: their advance count also depends on the parity of the page the
 * case happens to land on, so they corroborate at one room rather than sweep.
 */
const ROOMS = [200, 33, 27, 22, 21, 20, 19, 18, 17, 16, 8, 4];

/** The 20 pt construct's threshold is twice as far up, so it needs its own bracket. */
const ROOMS_B2 = [200, 40, 38, 36, 35, 34, 33, 32, 31, 30, 20, 4];

const roomsFor = (shape, type) => {
  if (type !== "nextPage" && type !== "continuous") return [200];
  return shape.tag === "B2" ? ROOMS_B2 : ROOMS;
};

/** Small, so a tight room stays readable: the marker itself must not spill. */
const MARKER_TWIPS = 120; // 6 pt = 8 CSS px

const cases = [];
for (const shape of shapes) {
  for (const type of types) {
    for (const room of roomsFor(shape, type.val)) {
      const id = `${shape.tag}${type.tag}${room}`;
      const shimPx = BODY_PX - FILLERS * FILLER_PX - room;
      if (shimPx <= 0) throw new Error(`room ${room} leaves no shim`);
      let out = "";
      for (let n = 0; n < FILLERS; n += 1) {
        out += exact(FILLER_TWIPS, n === FILLERS - 1 ? `${id}LAST` : "");
      }
      out += exact(Math.round(shimPx * 15), `${id}SHIM`);
      out += target(shape.text ? `${id}TGT` : "", shape.pageBreak, shape.sz);
      out += exact(MARKER_TWIPS, `${id}MARK`);
      // Terminates the FOLLOWING section, and so states how it began.
      out += `<w:p><w:pPr>${sectPr(type.val)}</w:pPr></w:p>`;
      cases.push({ id, shape: shape.tag, type: type.val, room, xml: out });
    }
  }
}

/**
 * The companion document is the NO-SECTION control, and it is the one that
 * decides whether a `w:sectPr` belongs in the rule at all. Every target above
 * carries one, so the document above can only ever describe the with-sectPr
 * branch — measure a gated engine there and you measure the gate, not the rule.
 *
 * Here the targets carry no `w:sectPr` and the document has no sections: each
 * case starts on a fresh page with `w:pageBreakBefore` instead. Rooms and sizes
 * are the same, so the two documents' thresholds compare directly. If they
 * agree for Word, a `sectPr` gate has nothing to gate.
 */
const bareTarget = (pageBreak, sz) => {
  const RPR = rpr(sz);
  return `<w:p><w:pPr>${RPR}</w:pPr>` + (pageBreak ? `<w:r>${RPR}<w:br w:type="page"/></w:r>` : "") + "</w:p>";
};

const bareShapes = [
  { tag: "NS", sz: 20, rooms: ROOMS }, // empty, break only, no sectPr, 10 pt
  { tag: "N2", sz: 40, rooms: ROOMS_B2 }, // the same at 20 pt
];

const bareCases = [];
for (const shape of bareShapes) {
  for (const room of shape.rooms) {
    const id = `${shape.tag}XX${room}`;
    const shimPx = BODY_PX - FILLERS * FILLER_PX - room;
    if (shimPx <= 0) throw new Error(`room ${room} leaves no shim`);
    let out = "";
    for (let n = 0; n < FILLERS; n += 1) {
      out +=
        n === 0
          ? `<w:p><w:pPr><w:pageBreakBefore/><w:spacing w:before="0" w:after="0" w:line="${FILLER_TWIPS}" ` +
            `w:lineRule="exact"/>${FILLER_RPR}</w:pPr></w:p>`
          : exact(FILLER_TWIPS, n === FILLERS - 1 ? `${id}LAST` : "");
    }
    out += exact(Math.round(shimPx * 15), `${id}SHIM`);
    out += bareTarget(true, shape.sz);
    out += exact(MARKER_TWIPS, `${id}MARK`);
    bareCases.push(out);
  }
}

const wrap = (body) =>
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
  '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
  `<w:body>${body}<w:sectPr><w:type w:val="nextPage"/>${pgProps}</w:sectPr></w:body></w:document>`;

const path = join(root, "fixtures-staging", "probe-sectadvance.docx");
writeFileSync(path, zipSync({ ...parts, "word/document.xml": strToU8(wrap(cases.map((c) => c.xml).join(""))) }));
console.log(`Wrote ${path} — ${cases.length} cases, rooms ${ROOMS.join(", ")} CSS px`);

const barePath = join(root, "fixtures-staging", "probe-sectadvance-nosect.docx");
writeFileSync(barePath, zipSync({ ...parts, "word/document.xml": strToU8(wrap(bareCases.join(""))) }));
console.log(`Wrote ${barePath} — ${bareCases.length} cases, no sectPr on any target`);
