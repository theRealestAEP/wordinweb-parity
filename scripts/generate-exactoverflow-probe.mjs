#!/usr/bin/env node

/**
 * Generate the probe for what an `hRule="exact"` row does when its content does
 * not fit.
 *
 * Two observations disagree. `generate-exactrow-probe.mjs` put 8 plain
 * paragraphs in an exact row and BOTH engines clipped them identically, which
 * is why `hRule="exact"` was recorded as handled correctly. Then a 14-paragraph
 * TOC placed in a 260 tw row vanished entirely in our render while Word grew
 * the row and painted about 90 lines out of it. Both cannot be the rule, so
 * something separates the cases, and guessing which is how #56 got filed under
 * a title that turned out to be false.
 *
 * Four candidates, each moved on its own:
 *
 *   AMOUNT     paragraphs in the row, swept 1..160. The two observations differ
 *              by an order of magnitude in overflow, and no measurement has ever
 *              varied it. If Word clips a little and grows a lot, this alone
 *              explains both and nothing else needs to be true.
 *   FIELD      the same overflow authored as a TOC field rather than as plain
 *              paragraphs, at one size in the clipping range and one in the
 *              growing range.
 *   CANTSPLIT  `w:cantSplit` on the row, same two sizes.
 *   POSITION   the row at the top of a page against the row at its foot, same
 *              two sizes, because a row that would overflow the PAGE is a
 *              different question from one that only overflows its own height.
 *
 * Every case is read the same way: the row's paragraphs are numbered, so the
 * highest one still painted says how much survived, and a MARK paragraph after
 * the table says how tall the row actually came out. Case ids are `<tag><n>`.
 *
 *   node scripts/generate-exactoverflow-probe.mjs
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

const RPR = '<w:rPr><w:sz w:val="20"/></w:rPr>';

/** A paragraph of an exactly known height: 240 tw exact = 16 CSS px. */
const line = (text, breakBefore = false) =>
  "<w:p><w:pPr>" +
  (breakBefore ? "<w:pageBreakBefore/>" : "") +
  `<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/>${RPR}</w:pPr>` +
  (text ? `<w:r>${RPR}<w:t>${text}</w:t></w:r>` : "") +
  "</w:p>";

/** The fixture's own row height: 260 tw, 13 pt, one 16 px line does not fit it. */
const ROW_TWIPS = 260;

/**
 * The row's content, numbered so the highest surviving number is readable. A
 * TOC field is authored with a real begin/instrText/separate/end and cached
 * results, exactly as an inserted TOC is written.
 */
const rowParagraphs = (id, n) =>
  Array.from({ length: n }, (_, k) => line(`${id}L${String(k + 1).padStart(3, "0")}`)).join("");

const rowToc = (id, n) => {
  const results = Array.from(
    { length: n },
    (_, k) =>
      `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/>${RPR}</w:pPr>` +
      `<w:r>${RPR}<w:t>${id}L${String(k + 1).padStart(3, "0")}</w:t></w:r></w:p>`,
  ).join("");
  return (
    `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/>${RPR}</w:pPr>` +
    `<w:r>${RPR}<w:fldChar w:fldCharType="begin"/></w:r>` +
    `<w:r>${RPR}<w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>` +
    `<w:r>${RPR}<w:fldChar w:fldCharType="separate"/></w:r></w:p>` +
    results +
    `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/>${RPR}</w:pPr>` +
    `<w:r>${RPR}<w:fldChar w:fldCharType="end"/></w:r></w:p>`
  );
};

const table = (id, n, { toc = false, cantSplit = false } = {}) =>
  '<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/><w:tblLayout w:type="fixed"/>' +
  '<w:tblBorders><w:top w:val="single" w:sz="6" w:space="0" w:color="auto"/>' +
  '<w:bottom w:val="single" w:sz="6" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid>' +
  `<w:tr><w:trPr><w:trHeight w:hRule="exact" w:val="${ROW_TWIPS}"/>` +
  (cantSplit ? "<w:cantSplit/>" : "") +
  "</w:trPr>" +
  '<w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/></w:tcPr>' +
  (toc ? rowToc(id, n) : rowParagraphs(id, n)) +
  "</w:tc></w:tr></w:tbl>";

const BODY_PX = 960 - 96;
const LINE_PX = 16;

/**
 * `fillPx` of filler above the table pushes the row down the page, so the same
 * overflow can be asked at the top of a page and at its foot.
 */
const buildCase = (id, n, { toc = false, cantSplit = false, fillPx = 0 } = {}) => {
  let out = line(`${id}TOP`, true);
  const fillers = Math.round(fillPx / LINE_PX);
  for (let k = 0; k < fillers; k += 1) out += line("");
  out += table(id, n, { toc, cantSplit });
  out += line(`${id}MARK`);
  return out;
};

const cases = [];
const add = (id, n, opts) => cases.push({ id, n, ...opts, xml: buildCase(id, n, opts) });

// AMOUNT: the one variable no measurement has moved. 1 fits the row's own
// 13 pt; everything above it overflows, by a little and then by a lot.
for (const n of [1, 2, 4, 8, 16, 32, 64, 90, 128, 160]) add(`A${n}`, n, {});

// FIELD: the same overflow as a TOC field, one small and one large.
for (const n of [8, 90]) add(`F${n}`, n, { toc: true });

// CANTSPLIT: a row that may not split across pages.
for (const n of [8, 90]) add(`C${n}`, n, { cantSplit: true });

// POSITION: the row at the foot of a page rather than the top of one.
for (const n of [8, 90]) add(`P${n}`, n, { fillPx: BODY_PX - 120 });

const documentXml =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
  '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
  `<w:body>${cases.map((c) => c.xml).join("")}${SECT}</w:body></w:document>`;

const path = join(root, "fixtures-staging", "probe-exactoverflow.docx");
writeFileSync(path, zipSync({ ...parts, "word/document.xml": strToU8(documentXml) }));
console.log(`Wrote ${path} — ${cases.length} cases, row ${ROW_TWIPS} tw exact`);
