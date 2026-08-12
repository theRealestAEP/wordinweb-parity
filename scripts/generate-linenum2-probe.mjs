#!/usr/bin/env node

/**
 * Follow-up to probe-linenum: isolate w:lnNumType's start attribute with NO
 * preceding section (rules out a table/section-boundary artifact from
 * probe-linenum's case G immediately before case H). Two clean one-section
 * documents, each the ONLY content, exact 24pt spacing, restart=newPage:
 *
 *   probe-linenum2a.docx  w:start="10"  (matches probe-linenum's case H)
 *   probe-linenum2b.docx  w:start="1"   (explicit, vs. H's implicit default)
 *
 * node scripts/generate-linenum2-probe.mjs
 */

import { readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { strToU8, unzipSync, zipSync } from "fflate";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const source = join(root, "apps/demo/public/fixtures/wild2-legal-ca-agreement.docx");
const parts = unzipSync(new Uint8Array(readFileSync(source)));

const pgProps =
  '<w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '<w:cols w:space="720"/>';

const line = (mark) =>
  `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="480" w:lineRule="exact"/></w:pPr>` +
  `<w:r><w:t xml:space="preserve">${mark}</w:t></w:r></w:p>`;

function build(startVal, prefix) {
  const lines = Array.from({ length: 5 }, (_, i) => line(`${prefix}${i + 1}`)).join("");
  const sectPr = `<w:sectPr>${pgProps}<w:lnNumType w:countBy="1" w:start="${startVal}"/></w:sectPr>`;
  const documentXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    `<w:body>${lines}${sectPr}</w:body></w:document>`;
  return documentXml;
}

const outA = join(root, "fixtures-staging/probe-linenum2a.docx");
writeFileSync(outA, zipSync({ ...parts, "word/document.xml": strToU8(build(10, "S10L")) }));
console.log("wrote", outA);

const outB = join(root, "fixtures-staging/probe-linenum2b.docx");
writeFileSync(outB, zipSync({ ...parts, "word/document.xml": strToU8(build(1, "S1EL")) }));
console.log("wrote", outB);
