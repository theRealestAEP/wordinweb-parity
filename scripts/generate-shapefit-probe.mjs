#!/usr/bin/env node

/**
 * Generate the probe that pins Word's `a:normAutofit` shrink-text behavior.
 *
 * The engine is about to gain `setDrawingTextFit {mode: "shrinkText"}`, which
 * writes `a:normAutofit fontScale / lnSpcReduction`. Every part of that is a
 * claim about WORD, so every part is measured here rather than assumed:
 *
 *  - given a shape with overfull text and a bare `<a:normAutofit/>` (no cached
 *    scale at all), does Word COMPUTE one, and what value does it write back?
 *  - is that value quantized, and onto what ladder? Six text lengths at two
 *    box heights give twelve independent demands on the same rule, so the
 *    steps read straight off the resaved XML.
 *  - does Word RENDER the shrink — is the painted text smaller than the run's
 *    authored size, and by how much?
 *  - does Word honor a fontScale IT did not compute (row 7's authored caches)?
 *
 * The existing corpus fixture probe3-shape-autofit already carries a single
 * authored `fontScale="62500" lnSpcReduction="20000"` box, and Word's PDF
 * paints it at the full 11pt. That is one case at one box size, which is why
 * the engine's parse comment calls the mode a clip. This probe varies the
 * demand on both sides of that: the bare-normAutofit column asks Word to do
 * the arithmetic itself, and the authored column re-tests the ignore rule at a
 * second box height.
 *
 * Layout: three columns of 2.0in boxes at page x = 0.4 / 3.0 / 5.6 in, one
 * exact-height paragraph per row so the rows cannot interact.
 *
 *   A = bare <a:normAutofit/>, box 2.0 x 1.4 in
 *   B = bare <a:normAutofit/>, box 2.0 x 0.7 in
 *   C = <a:noAutofit/>,        box 2.0 x 1.4 in   (full-size clip control)
 *
 * Rows 1-6 hold 6 / 12 / 20 / 30 / 45 / 70 words of lorem, from "fits" to
 * "three times overfull". Row 7 swaps in the authored caches:
 *
 *   A7 = <a:normAutofit fontScale="90000"/>
 *   B7 = <a:normAutofit fontScale="62500" lnSpcReduction="20000"/>
 *   C7 = <a:spAutoFit/>                            (grow-the-box control)
 *
 *   node scripts/generate-shapefit-probe.mjs
 */

import { readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { strToU8, strFromU8, unzipSync, zipSync } from "fflate";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const source = join(root, "apps/demo/public/fixtures/probe-wrapclear.docx");
const parts = unzipSync(new Uint8Array(readFileSync(source)));
/** Reuse the base package's w:document open tag: it already declares wp / a /
 * wps / mc, which a shape needs. */
const documentOpen = strFromU8(parts["word/document.xml"]).match(/<w:document [^>]*>/)[0];

const EMU = 914400;
const inches = (value) => Math.round(value * EMU);

const LOREM = (
  "Lorem ipsum dolor sit amet consectetur adipiscing elit sed do eiusmod tempor " +
  "incididunt ut labore et dolore magna aliqua Ut enim ad minim veniam quis nostrud " +
  "exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat Duis aute " +
  "irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla " +
  "pariatur Excepteur sint occaecat cupidatat non proident sunt in culpa qui officia " +
  "deserunt mollit anim id est laborum"
).split(" ");

const esc = (text) => text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

let nextDrawingId = 900;

/** One anchored rectangle with a bold case label and `words` lorem words. */
function shape(id, xIn, widthIn, heightIn, autofit, words) {
  const cx = inches(widthIn);
  const cy = inches(heightIn);
  const drawingId = nextDrawingId++;
  const text = LOREM.slice(0, words).join(" ");
  const txbx =
    "<wps:txbx><w:txbxContent>" +
    '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>' +
    `<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">${id}: </w:t></w:r>` +
    `<w:r><w:t xml:space="preserve">${esc(text)}</w:t></w:r>` +
    "</w:p></w:txbxContent></wps:txbx>";
  return (
    "<w:r><w:drawing>" +
    '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" ' +
    `relativeHeight="${2516000 + drawingId}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    '<wp:simplePos x="0" y="0"/>' +
    `<wp:positionH relativeFrom="page"><wp:posOffset>${inches(xIn)}</wp:posOffset></wp:positionH>` +
    '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
    `<wp:extent cx="${cx}" cy="${cy}"/>` +
    '<wp:effectExtent l="0" t="0" r="0" b="0"/>' +
    "<wp:wrapNone/>" +
    `<wp:docPr id="${drawingId}" name="ShapeFit ${id}"/>` +
    "<wp:cNvGraphicFramePr/>" +
    "<a:graphic>" +
    '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
    "<wps:wsp><wps:cNvSpPr/><wps:spPr>" +
    `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
    '<a:solidFill><a:srgbClr val="FFF2CC"/></a:solidFill>' +
    '<a:ln w="12700"><a:solidFill><a:srgbClr val="C55A11"/></a:solidFill></a:ln>' +
    "</wps:spPr>" +
    txbx +
    `<wps:bodyPr rot="0" anchor="t">${autofit}</wps:bodyPr>` +
    "</wps:wsp></a:graphicData></a:graphic>" +
    "</wp:anchor></w:drawing></w:r>"
  );
}

const BARE = "<a:normAutofit/>";
const NONE = "<a:noAutofit/>";
const WORDS = [6, 12, 20, 30, 45, 70];

/**
 * One row per page: the three shapes hang off a single short paragraph at the
 * top of the body, and a page break ends the row.
 *
 * The first cut gave each row an exact-height 1.65in paragraph instead, and
 * that put six rows on Word's page 1 against our five — an exact-line-height
 * pagination question that has nothing to do with autofit but swamped every
 * page score. A page per row removes the flow from the probe entirely.
 */
const row = (runs, last) =>
  `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>${runs}</w:p>` +
  (last ? "" : '<w:p><w:r><w:br w:type="page"/></w:r></w:p>');

const rows = WORDS.map((words, index) => {
  const n = index + 1;
  return row(
    shape(`A${n}`, 0.4, 2.0, 1.4, BARE, words) +
      shape(`B${n}`, 3.0, 2.0, 0.7, BARE, words) +
      shape(`C${n}`, 5.6, 2.0, 1.4, NONE, words),
    false,
  );
});

// Row 7: caches Word did not compute, at the same box the bare column uses.
rows.push(
  row(
    shape("A7", 0.4, 2.0, 1.4, '<a:normAutofit fontScale="90000"/>', 30) +
      shape("B7", 3.0, 2.0, 1.4, '<a:normAutofit fontScale="62500" lnSpcReduction="20000"/>', 30) +
      shape("C7", 5.6, 2.0, 1.4, "<a:spAutoFit/>", 30),
    true,
  ),
);

const sectPr =
  "<w:sectPr>" +
  '<w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '<w:cols w:space="720"/>' +
  "</w:sectPr>";

const documentXml =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  `${documentOpen}<w:body>${rows.join("")}${sectPr}</w:body></w:document>`;

const out = join(root, "fixtures-staging/probe-shapefit.docx");
writeFileSync(out, zipSync({ ...parts, "word/document.xml": strToU8(documentXml) }));
console.log("wrote", out, `(${rows.length * 3} shapes)`);
