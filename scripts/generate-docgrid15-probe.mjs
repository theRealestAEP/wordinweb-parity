#!/usr/bin/env node

/**
 * Compat-15 `w:docGrid type="lines"` sweep: what does Word do with a line
 * whose natural height EXCEEDS the grid pitch?
 *
 * staging-eastasian (compat 15, pitch 360tw = 24px) motivates this. Measured
 * by ink bands against its clean reference, every Normal<->Normal paragraph
 * boundary agrees to 0.5px while all four boundaries that touch a Heading1
 * line (Calibri Light 16pt, natural 26.04px > pitch 24) disagree: Word gives
 * the heading ~19.9px more flow than our natural x multiplier lays, split
 * above and below the glyphs - exactly ceil(26.04/24) x 24 = 48px with the
 * glyph box centered. The engine's textSnap branch (inline.ts) implements
 * that rule but gates it on compatibilityMode < 12..14, citing this very
 * fixture: "staging-eastasian (compat 15) lays multiplier x natural for faces
 * well above its pitch". That was calibrated on the CHINESE FALLBACK lines,
 * where snap (2 x 24 = 48) and multiplier x natural (44.48 x 1.0792 = 48.0)
 * are NUMERICALLY IDENTICAL - the case discriminates nothing. probe-docgrid
 * and probe-gridopen are both compat 12, so the compat-15 side of the gate
 * has never been probed with a case that separates the two readings.
 *
 * Sections, each `nextPage`, all compat 15, Normal = Calibri 11
 * (after=160 line=259 auto), the staging-eastasian styles.xml:
 *
 *   A360  Calibri 11pt        pitch 360   control: natural (17.9) < pitch (24)
 *   B360  Calibri Light 16pt  pitch 360   natural 26.04 > 24: snap->48, mult->28.1
 *   D360  Calibri Light 18pt  pitch 360   natural 29.3 > 24:  snap->48, mult->31.6
 *   E360  MS Mincho 11pt ja   pitch 360   EA control (Word measured 26.00)
 *   F360  MS Mincho 16pt ja   pitch 360   EA oversized: does the EA carve-out exist?
 *   H360  Heading1 2 lines    pitch 360   the fixture's construct + follower gap
 *   A240  Calibri 11pt        pitch 240   natural 17.9 > 16: snap->32, mult->19.3
 *   B240  Calibri Light 16pt  pitch 240   natural 26.04 > 16: snap->32, mult->28.1
 *
 * Every case is one 3-line paragraph (two <w:br/>) so the in-paragraph line
 * advance reads directly, followed by one Normal 11pt paragraph (label xF)
 * so the paragraph-boundary gap reads too. Two pitches make the two-setting
 * sweep for any "rows of the pitch" claim.
 *
 *   node scripts/generate-docgrid15-probe.mjs
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
const latinBr = (family, halfPts) =>
  `<w:r><w:rPr><w:rFonts w:ascii="${family}" w:hAnsi="${family}"/><w:sz w:val="${halfPts}"/></w:rPr><w:br/></w:r>`;
const jaRun = (text, halfPts) =>
  `<w:r><w:rPr><w:rFonts w:eastAsia="MS Mincho" w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="${halfPts}"/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;
const jaBr = (halfPts) =>
  `<w:r><w:rPr><w:rFonts w:eastAsia="MS Mincho" w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="${halfPts}"/></w:rPr><w:br/></w:r>`;

/** 3-line paragraph via two explicit line breaks; no authored spacing so
 * Normal's after=160 line=259 govern (the fixture's own environment). */
const latinCase = (id, family, halfPts) =>
  `<w:p><w:pPr><w:kinsoku/></w:pPr>` +
  latinRun(`${id} line one`, family, halfPts) +
  latinBr(family, halfPts) +
  latinRun(`${id} line two`, family, halfPts) +
  latinBr(family, halfPts) +
  latinRun(`${id} line three`, family, halfPts) +
  `</w:p>` +
  `<w:p>${latinRun(`${id}F follower`, "Calibri", 22)}</w:p>`;

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

const headingCase = (id) =>
  `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>` +
  `<w:r><w:t xml:space="preserve">${id} heading line one</w:t></w:r>` +
  `<w:r><w:br/></w:r>` +
  `<w:r><w:t xml:space="preserve">${id} heading line two</w:t></w:r>` +
  `</w:p>` +
  `<w:p>${latinRun(`${id}F follower`, "Calibri", 22)}</w:p>`;

const sections = [
  { body: latinCase("A360", "Calibri", 22), pitch: 360 },
  { body: latinCase("B360", "Calibri Light", 32), pitch: 360 },
  { body: latinCase("D360", "Calibri Light", 36), pitch: 360 },
  { body: jaCase("E360", 22), pitch: 360 },
  { body: jaCase("F360", 32), pitch: 360 },
  { body: headingCase("H360"), pitch: 360 },
  { body: latinCase("A240", "Calibri", 22), pitch: 240 },
  { body: latinCase("B240", "Calibri Light", 32), pitch: 240 },
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
const out = join(root, "fixtures-staging/probe-docgrid15.docx");
writeFileSync(out, zipSync(parts));
console.log("wrote", out);
