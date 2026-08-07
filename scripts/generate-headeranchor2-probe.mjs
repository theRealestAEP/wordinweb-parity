#!/usr/bin/env node

/**
 * #97: what does a header's anchored shape that extends BELOW the header's
 * last line do to the body top?
 *
 * parity-hftemplates p3/p4 (#95 residual): Word's body top is 119.03px where
 * we stop at the 96px margin (p3, Banded - a wrapSquare bar), and 181.39px
 * where we reach 151.69 (p4, Ion Dark - wrapTopAndBottom). The #80 probe
 * (probe-headeranchor) cannot see this: its 8pt shape ends ABOVE the 12pt
 * header line's bottom, and Word reserves nothing there at any wrap setting.
 * The hftemplates bars end at page-relative 56.9pt (wrapSquare) and 96.0pt
 * (wrapTopAndBottom) - well below the header text.
 *
 * Same geometry as probe-headeranchor (w:top="0" so the header governs,
 * header distance 720tw = 48px, one exact-12pt header line, exact-10pt
 * marker first in the body so its top IS the body top). One 1-inch-wide bar
 * anchored at the header paragraph's top, varying ONLY wrap mode x extent:
 *
 *   S8            wrapSquare,       8pt bar  (ends above the line - #80 control)
 *   N40, N72      wrapNone,        40 / 72pt
 *   S40, S72      wrapSquare,      40 / 72pt
 *   T40, T72      wrapTopAndBottom 40 / 72pt
 *
 * Two extents per wrap make the two-setting sweep: if the body top tracks the
 * shape's bottom, the 72pt cases sit 32pt lower than the 40pt ones; a flat
 * clearance does not scale.
 *
 *   node scripts/generate-headeranchor2-probe.mjs
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

const CX = 914400; // 1 inch wide
const PT = 12700; // EMU per point

const anchor = (wrap, cyPt, id) =>
  "<w:r><w:drawing>" +
  '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" ' +
  'behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/>' +
  '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>' +
  '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
  `<wp:extent cx="${CX}" cy="${cyPt * PT}"/>` +
  '<wp:effectExtent l="0" t="0" r="0" b="0"/>' +
  wrap +
  `<wp:docPr id="${id}" name="Shape${id}"/>` +
  '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
  "<wps:wsp><wps:cNvSpPr/><wps:spPr>" +
  `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${CX}" cy="${cyPt * PT}"/></a:xfrm>` +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
  '<a:solidFill><a:srgbClr val="000000"/></a:solidFill>' +
  "</wps:spPr><wps:bodyPr/></wps:wsp>" +
  "</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>";

const hline = (tag, extra = "") =>
  '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/>' +
  '<w:rPr><w:sz w:val="20"/></w:rPr></w:pPr>' +
  extra +
  `<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>${tag}</w:t></w:r></w:p>`;

const WRAPS = {
  N: "<wp:wrapNone/>",
  S: '<wp:wrapSquare wrapText="bothSides"/>',
  T: "<wp:wrapTopAndBottom/>",
};

const HEADERS = [
  { id: "S8", wrap: "S", cy: 8 },
  { id: "N40", wrap: "N", cy: 40 },
  { id: "N72", wrap: "N", cy: 72 },
  { id: "S40", wrap: "S", cy: 40 },
  { id: "S72", wrap: "S", cy: 72 },
  { id: "T40", wrap: "T", cy: 40 },
  { id: "T72", wrap: "T", cy: 72 },
];

let rels = strFromU8(parts["word/_rels/document.xml.rels"]);
let types = strFromU8(parts["[Content_Types].xml"]);
HEADERS.forEach((h, i) => {
  const name = `headerA97_${i}.xml`;
  h.rid = `rIdA97${i}`;
  parts[`word/${name}`] = strToU8(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
      `<w:hdr ${NS}>${hline(`HD${h.id}`, anchor(WRAPS[h.wrap], h.cy, 200 + i))}</w:hdr>`,
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

const out = join(root, "fixtures-staging", "probe-headeranchor2.docx");
writeFileSync(out, Buffer.from(zipSync(parts)));
console.log(`wrote ${out} — ${HEADERS.map((h) => h.id).join(", ")}`);
