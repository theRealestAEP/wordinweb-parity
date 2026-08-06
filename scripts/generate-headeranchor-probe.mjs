#!/usr/bin/env node

/**
 * Generate the anchored-shape variant that scopes #80.
 *
 * `generate-headerheight-probe.mjs` showed Word reserves NOTHING above a header
 * whose content is a table, where we add `ptToPx(22.5)` — 30.00 CSS px. But the
 * term that adds it is guarded by two conditions, and only one of them was
 * exercised:
 *
 *     const hasOnlyUnwrappedAnchors =
 *       anchors.length > 0 && anchors.every((s) => s.wrap === undefined || s.wrap === "none");
 *     const complexHeader = anchors.length > 0 || blocks.some((b) => b.type === "table");
 *     ... + (!hasOnlyUnwrappedAnchors && complexHeader ? ptToPx(22.5) : 0)
 *
 * A table header has `anchors.length === 0`, so it reaches the clearance through
 * `complexHeader`'s SECOND disjunct with `hasOnlyUnwrappedAnchors` false by
 * vacuity. The anchored-shape path is untouched by that measurement, and the
 * `hasOnlyUnwrappedAnchors` carve-out exists because of #67's pleading-paper
 * rails — so 22.5 pt may be exactly right for a WRAPPED anchored shape, and
 * removing the term outright would regress the case it was built for.
 *
 * Three anchor cases against a plain-paragraph control, one wrap setting each:
 *
 *   PB  plain paragraph, no anchor              (control, fixes the baseline)
 *   AN  anchor, <wp:wrapNone/>                  hasOnlyUnwrappedAnchors TRUE  -> no clearance
 *   AS  anchor, <wp:wrapSquare/>                hasOnlyUnwrappedAnchors FALSE -> clearance
 *   AT  anchor, <wp:wrapTopAndBottom/>          hasOnlyUnwrappedAnchors FALSE -> clearance
 *
 * The shape is 8 pt tall and anchored at the paragraph's own top, so it ends
 * ABOVE the 12 pt line's bottom and cannot raise `contentBottom`. That keeps the
 * only quantity in play the clearance itself: if Word's body top for AS and AT
 * equals PB's, Word reserves nothing for a wrapped anchor either and the whole
 * term goes; if AS and AT sit 30 px lower in BOTH engines, the term is correct
 * there and #80 must be scoped to the table disjunct alone.
 *
 * Same geometry as the header-height probe — `w:top="0"` so the header governs
 * the body top, one 12 pt header line, marker of exactly known height — so the
 * two probes' numbers are directly comparable.
 *
 *   node scripts/generate-headeranchor-probe.mjs
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
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"';

/** 8 pt tall (101600 EMU), so it ends above a 12 pt line's bottom. */
const CX = 914400; // 1 inch wide
const CY = 101600; // 8 pt tall

/**
 * A positioned shape anchored at the paragraph's own top. `wps` needs no image
 * part, so the anchor is the only thing that varies between cases.
 */
const anchor = (wrap, id) =>
  "<w:r><w:drawing>" +
  '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" ' +
  'behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/>' +
  '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>' +
  '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
  `<wp:extent cx="${CX}" cy="${CY}"/>` +
  '<wp:effectExtent l="0" t="0" r="0" b="0"/>' +
  wrap +
  `<wp:docPr id="${id}" name="Shape${id}"/>` +
  '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
  "<wps:wsp><wps:cNvSpPr/><wps:spPr>" +
  `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${CX}" cy="${CY}"/></a:xfrm>` +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
  '<a:solidFill><a:srgbClr val="000000"/></a:solidFill>' +
  "</wps:spPr><wps:bodyPr/></wps:wsp>" +
  "</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>";

/** One header line of exactly 12 pt, optionally carrying an anchored shape. */
const hline = (tag, extra = "") =>
  '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/>' +
  '<w:rPr><w:sz w:val="20"/></w:rPr></w:pPr>' +
  extra +
  `<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>${tag}</w:t></w:r></w:p>`;

const HEADERS = [
  { id: "PB", body: (t) => hline(t) },
  { id: "AN", body: (t) => hline(t, anchor("<wp:wrapNone/>", 101)) },
  { id: "AS", body: (t) => hline(t, anchor('<wp:wrapSquare wrapText="bothSides"/>', 102)) },
  { id: "AT", body: (t) => hline(t, anchor("<wp:wrapTopAndBottom/>", 103)) },
];

let rels = strFromU8(parts["word/_rels/document.xml.rels"]);
let types = strFromU8(parts["[Content_Types].xml"]);
HEADERS.forEach((h, i) => {
  const name = `headerA80_${i}.xml`;
  h.rid = `rIdA80${i}`;
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

const marker = (id) =>
  '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/>' +
  '<w:rPr><w:sz w:val="20"/></w:rPr></w:pPr>' +
  `<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>${id}</w:t></w:r></w:p>`;

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

const out = join(root, "fixtures-staging", "probe-headeranchor.docx");
writeFileSync(out, Buffer.from(zipSync(parts)));
console.log(`wrote ${out} — ${HEADERS.map((h) => h.id).join(", ")}`);
