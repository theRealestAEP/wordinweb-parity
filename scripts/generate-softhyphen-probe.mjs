#!/usr/bin/env node

/**
 * Generate the probe that pins Word's `<w:softHyphen/>` line-break behavior.
 *
 * The engine's layout currently ignores soft hyphens entirely (the element
 * parses to a plain U+00AD atom that is not in the in-word break set), and the
 * tool-depth matrix filed the three coupled changes an implementation needs:
 * zero-width measurement, a breakAfter split, and hyphen paint + fit
 * reservation at a soft break. Each of those is a claim about WORD, so each is
 * measured here rather than assumed:
 *
 *  - does Word break a line at a soft hyphen at all?
 *  - which fit demand governs the break: the prefix alone, or the prefix PLUS
 *    the hyphen glyph it will paint?
 *  - what does a soft hyphen measure MID-line (expected: zero)?
 *  - does compatibilityMode change any of it?
 *
 * Construction: each case is a label paragraph and a single-word target
 * paragraph whose available width is set by `w:ind w:right`. The target word
 * carries two soft hyphens, so which text lands on line 1 reads the rule
 * directly: `hydro-` (break at SH1), `hydromatic-` (break at SH2), the whole
 * word (fits), or a character split (no opportunity fits). The available
 * widths bracket every knife edge in 1pt steps, from exact Calibri/Carlito
 * advances at 11pt:
 *
 *   hydro 26.17   hydro- 29.54   hydromatic 51.09   hydromatic- 54.46
 *   hydromaticgraphs 81.23       (and the U family at its own widths)
 *
 * Two words sweep the same rule at two prefix widths. `M1` reads the mid-line
 * width of a soft hyphen off the painted x-extents at full width. `R50`/`R58`
 * hold the OTHER construct — a raw U+00AD character typed into w:t — which
 * Word paints as an always-visible, never-breaking hyphen (the parse already
 * maps it to U+2011; these cases keep that rule measured).
 *
 * Two packages differing only in compatibilityMode (15 / 11).
 *
 *   node scripts/generate-softhyphen-probe.mjs
 */

import { readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { strToU8, strFromU8, unzipSync, zipSync } from "fflate";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const source = join(root, "apps/demo/public/fixtures/probe-wrapclear.docx");
const parts = unzipSync(new Uint8Array(readFileSync(source)));

/** Letter, 1in margins: text width 9360 tw = 468 pt. */
const TEXT_PT = 468;
const pgProps =
  '<w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '<w:cols w:space="720"/>';

const SPACING = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>';

const label = (id) =>
  `<w:p><w:pPr>${SPACING}</w:pPr><w:r><w:t xml:space="preserve">CASE ${id}</w:t></w:r></w:p>`;

/** A target paragraph: one run, `segs` joined by <w:softHyphen/>, available
 * width `availPt` set by the right indent. */
const shTarget = (availPt, segs) => {
  const indR = Math.round((TEXT_PT - availPt) * 20);
  const runChildren = segs
    .map((s) => `<w:t xml:space="preserve">${s}</w:t>`)
    .join("<w:softHyphen/>");
  return `<w:p><w:pPr>${SPACING}<w:ind w:right="${indR}"/></w:pPr><w:r>${runChildren}</w:r></w:p>`;
};

/** The raw-character control: U+00AD typed into w:t. */
const rawTarget = (availPt, text) => {
  const indR = Math.round((TEXT_PT - availPt) * 20);
  return `<w:p><w:pPr>${SPACING}<w:ind w:right="${indR}"/></w:pPr><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
};

const H = ["hydro", "matic", "graphs"];
const U = ["under", "standing", "ly"];

const cases = [];
const sweep = (tag, segs, avails) => {
  for (const a of avails) cases.push(label(`${tag}${a}`) + shTarget(a, segs));
};

const range = (from, to) => Array.from({ length: to - from + 1 }, (_, i) => from + i);

// SH1 knife (prefix "hydro" 26.17, "hydro-" 29.54): 24..33.
sweep("H", H, range(24, 33));
// SH2 knife ("hydromatic" 51.09, "hydromatic-" 54.46): 49..58.
sweep("H", H, range(49, 58));
// whole-word knife (81.23): 78..84.
sweep("H", H, range(78, 84));
// The U family repeats every knife at its own widths.
sweep("U", U, range(25, 32)); // under 26.65 / under- 30.01
sweep("U", U, range(62, 70)); // understanding 64.94 / understanding- 68.31
sweep("U", U, range(71, 76)); // understandingly 72.45

// Mid-line width readout at full width: x-extents around the soft hyphens.
cases.push(label("M1") + shTarget(TEXT_PT, ["mid hydro", "matic", "graphs end"]));

// Raw U+00AD characters (the always-visible, never-breaking construct).
cases.push(label("R50") + rawTarget(50, "hydro­matic"));
cases.push(label("R58") + rawTarget(58, "hydro­matic"));

const wrap = (body) =>
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  `<w:body>${body}<w:sectPr>${pgProps}</w:sectPr></w:body></w:document>`;

const documentXml = wrap(cases.join(""));

const out15 = join(root, "fixtures-staging/probe-softhyphen.docx");
writeFileSync(out15, zipSync({ ...parts, "word/document.xml": strToU8(documentXml) }));
console.log("wrote", out15, `(${cases.length / 1} case parts)`);

const settings11 = strFromU8(parts["word/settings.xml"]).replace(
  'w:val="15" w:uri="http://schemas.microsoft.com/office/word" w:name="compatibilityMode"',
  'w:val="11" w:uri="http://schemas.microsoft.com/office/word" w:name="compatibilityMode"',
);
if (!settings11.includes('w:val="11"')) throw new Error("compat rewrite failed");
const out11 = join(root, "fixtures-staging/probe-softhyphen11.docx");
writeFileSync(out11, zipSync({ ...parts, "word/document.xml": strToU8(documentXml), "word/settings.xml": strToU8(settings11) }));
console.log("wrote", out11);
