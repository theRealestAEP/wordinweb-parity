#!/usr/bin/env node

/**
 * Generate the rule paint-width sweep for #19.
 *
 * `renderEdge` (core/src/render/dom.ts) snaps a rule's painted width to a whole
 * DEVICE pixel: `Math.max(1/dpr, Math.round(declaredWidth * dpr) / dpr)`. At the
 * parity renderer's 2x that is a half-CSS-pixel grid, and it is exact for most
 * weights but not for the two commonest ones:
 *
 *   w:sz   authored   in CSS px   we paint   error
 *      4     0.50 pt      0.667      0.500   -25.0%
 *      8     1.00 pt      1.333      1.500   +12.5%
 *      6     0.75 pt      1.000      1.000     0
 *     12     1.50 pt      2.000      2.000     0
 *     18     2.25 pt      3.000      3.000     0
 *
 * The snap was deliberate — a 0.5 pt rule lands on ONE physical pixel at 2x and
 * looks crisp, where the unsnapped 1.333 device px antialiases across two — but
 * it is 25% light on the weight Word uses by default, and that is what dominates
 * `tableRuleWeightErrorPct` (27.5% on staging-tblextreme with positions already
 * matching).
 *
 * WHAT NEEDS MEASURING IS WORD, NOT US. Our painted width is exactly computable
 * from the formula above, and total ink mass is preserved under antialiasing, so
 * once Word's own painted width per `w:sz` is known the whole error curve is
 * arithmetic and needs no raster. The open question is whether Word paints
 * `sz/8` points faithfully at every weight or quantizes too — if Word also
 * snaps, the snap is not the defect and the calibration target moves.
 *
 * One paragraph per weight, each carrying only a bottom border, spaced so no two
 * rules can touch. Every weight Word's UI offers plus the two hairlines:
 * sz 2, 4, 6, 8, 12, 18, 24, 36, 48 — that is 0.25 to 6.00 pt.
 *
 * Read the Word export's stroke widths straight out of the content stream
 * (PyMuPDF `get_drawings()` reports a `width` per path in points), and compare
 * with `sz/8`. A rule painted as a filled RECTANGLE rather than a stroke reports
 * as a rect, and its height is the width — check both shapes before concluding
 * Word quantizes anything.
 *
 *   node scripts/generate-rulewidth-probe.mjs
 */

import { readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { strToU8, unzipSync, zipSync } from "fflate";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const source = join(root, "apps/demo/public/fixtures/wild2-legal-ca-agreement.docx");
const parts = unzipSync(new Uint8Array(readFileSync(source)));

const SECT =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '<w:cols w:space="720"/></w:sectPr>';

/** Every weight Word's border UI offers, plus the sub-hairline 0.25 pt. */
const SIZES = [2, 4, 6, 8, 12, 18, 24, 36, 48];

/**
 * `w:space="0"` keeps the rule on the paragraph's own box edge, which is the
 * case the renderer snaps; a spaced border takes the other placement branch and
 * would confound the width question with a position one.
 */
const rule = (sz) =>
  // CT_PPr is a SEQUENCE: w:pBdr precedes w:spacing. Word repairs the wrong
  // order silently, which would leave the probe measuring a package Word
  // rewrote rather than the one authored here.
  `<w:p><w:pPr>` +
  `<w:pBdr><w:bottom w:val="single" w:sz="${sz}" w:space="0" w:color="000000"/></w:pBdr>` +
  `<w:spacing w:before="360" w:after="360" w:line="240" w:lineRule="auto"/>` +
  `<w:rPr><w:sz w:val="20"/></w:rPr></w:pPr>` +
  `<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>SZ${String(sz).padStart(2, "0")}</w:t></w:r></w:p>`;

const documentXml =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  `<w:body>${SIZES.map(rule).join("")}${SECT}</w:body></w:document>`;

parts["word/document.xml"] = strToU8(documentXml);

// Probes live in fixtures-staging: parity-parallel.mjs adopts a reference only
// when a DOCX of the same name sits in apps/demo/public/fixtures, so keeping it
// out of there is what stops the full corpus run from picking the probe up.
// To read it in the browser: copy it in, measure, then remove it again.
const out = join(root, "fixtures-staging/probe-rulewidth.docx");
writeFileSync(out, Buffer.from(zipSync(parts)));
console.log(`wrote ${out}`);
console.log(`sizes: ${SIZES.map((s) => `sz${s} (${s / 8}pt)`).join(", ")}`);
