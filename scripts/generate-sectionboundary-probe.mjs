#!/usr/bin/env node

/**
 * Generate apps/demo/public/fixtures/probe-sectionboundary.docx — what sets the
 * body origin on a section's first page.
 *
 * On wild2-med-nccih-protocol we place page 4's first heading 34 CSS px
 * (510 twips) below where Word puts it. An earlier probe
 * (generate-sectionstart-probe.mjs) ruled out titlePg and an empty-Heading2
 * sectPr carrier: with no footer parts declared, Word and this engine agreed on
 * every combination. What that probe could NOT vary was the first-page FOOTER,
 * because it declared no footer parts at all — and nccih's section carries both
 * a default and a first-page footer reference while carrying no header
 * reference.
 *
 * This one declares real footers and sweeps in two blocks, so each variable
 * moves alone:
 *
 *   Block 1  titlePg x first-page-footer-reference, carrier held constant
 *            (a non-empty paragraph)
 *   Block 2  the sectPr carrier — empty Heading2 / empty Normal / non-empty —
 *            with titlePg and both footer references held at the nccih shape
 *
 * The carrier of a section's sectPr is the LAST paragraph of that section, so
 * it can only influence the page that FOLLOWS. Block 2 therefore reads as: the
 * carrier of case k affects the measured heading of case k+1.
 *
 *   node scripts/generate-sectionboundary-probe.mjs
 */

import { writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { strToU8, zipSync } from "fflate";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const out = join(root, "apps/demo/public/fixtures/probe-sectionboundary.docx");

const PAGE = '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="576" w:gutter="0"/>';

const CASES = [
  // Block 1: titlePg x first-page footer, carrier constant.
  { id: "1", titlePg: false, firstFooter: false, carrier: "text", note: "no titlePg, no first footer" },
  { id: "2", titlePg: true,  firstFooter: false, carrier: "text", note: "titlePg only" },
  { id: "3", titlePg: false, firstFooter: true,  carrier: "text", note: "first footer only" },
  { id: "4", titlePg: true,  firstFooter: true,  carrier: "text", note: "titlePg + first footer" },
  // Block 2: carrier variants at the nccih shape.
  { id: "5", titlePg: true, firstFooter: true, carrier: "emptyHeading", note: "carrier = empty Heading2" },
  { id: "6", titlePg: true, firstFooter: true, carrier: "emptyNormal", note: "carrier = empty Normal" },
  { id: "7", titlePg: true, firstFooter: true, carrier: "text", note: "carrier = non-empty" },
];

const filler = (n) => Array.from({ length: n },
  (_, i) => `<w:p><w:r><w:t xml:space="preserve">filler ${i + 1}</w:t></w:r></w:p>`).join("");

function sectPr({ titlePg, firstFooter }) {
  return "<w:sectPr>" +
    '<w:footerReference w:type="default" r:id="rId10"/>' +
    (firstFooter ? '<w:footerReference w:type="first" r:id="rId11"/>' : "") +
    (titlePg ? "<w:titlePg/>" : "") +
    PAGE + "</w:sectPr>";
}

function terminator(item) {
  const sp = sectPr(item);
  if (item.carrier === "emptyHeading") return `<w:p><w:pPr><w:pStyle w:val="Heading2"/>${sp}</w:pPr></w:p>`;
  if (item.carrier === "emptyNormal") return `<w:p><w:pPr>${sp}</w:pPr></w:p>`;
  return `<w:p><w:pPr>${sp}</w:pPr><w:r><w:t xml:space="preserve">end of section</w:t></w:r></w:p>`;
}

const body = CASES.map((item) =>
  `<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t xml:space="preserve">CASE ${item.id} HEADING</w:t></w:r></w:p>` +
  filler(3) + terminator(item)).join("") +
  `<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t xml:space="preserve">CASE 8 HEADING</w:t></w:r></w:p>` +
  `<w:sectPr><w:footerReference w:type="default" r:id="rId10"/>${PAGE}</w:sectPr>`;

const footer = (text) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:ftr>`;

writeFileSync(out, zipSync({
  "[Content_Types].xml": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
  <Override PartName="/word/footer2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
</Types>`),
  "_rels/.rels": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`),
  "word/_rels/document.xml.rels": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
  <Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
  <Relationship Id="rId11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer2.xml"/>
</Relationships>`),
  "word/document.xml": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${body}</w:body></w:document>`),
  "word/footer1.xml": strToU8(footer("default footer")),
  "word/footer2.xml": strToU8(footer("first page footer")),
  "word/styles.xml": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr></w:rPrDefault>
  <w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/>
    <w:pPr><w:spacing w:before="240" w:after="120" w:line="240" w:lineRule="atLeast"/><w:outlineLvl w:val="1"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="26"/></w:rPr></w:style>
</w:styles>`),
  "word/settings.xml": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:compat/></w:settings>`),
}));
console.log(`Wrote ${out}`);
for (const c of CASES) console.log(`  CASE ${c.id}: ${c.note} (carrier ${c.carrier})`);
