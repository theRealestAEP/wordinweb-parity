#!/usr/bin/env node

/**
 * Generate apps/demo/public/fixtures/probe-sectionstart.docx — where a section's
 * first page puts its first paragraph.
 *
 * On wild2-med-nccih-protocol we place page 4's heading 34 CSS px (510 twips)
 * lower than Word does. Page 4 is the first page of a new section whose
 * PRECEDING paragraph is an empty Heading2 carrying the <w:sectPr>, and that
 * sectPr sets <w:titlePg/>. Heading2's own space-before is 240 twips, which is
 * only 16 CSS px — under half the gap — so something else contributes.
 *
 * This varies the two structural candidates independently against a plain
 * control, holding the styles and the following heading constant:
 *
 *   A  section ends on a NORMAL paragraph, no titlePg      (control)
 *   B  section ends on an EMPTY Heading2, no titlePg       (carrier)
 *   C  section ends on a NORMAL paragraph, titlePg         (titlePg)
 *   D  section ends on an EMPTY Heading2, titlePg          (both, = nccih)
 *
 * Measure the first heading's top offset on each section's first page. Whichever
 * variable moves it is the one that matters; if D moves further than B and C
 * separately, they interact.
 *
 *   node scripts/generate-sectionstart-probe.mjs
 */

import { writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { strToU8, zipSync } from "fflate";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const out = join(root, "apps/demo/public/fixtures/probe-sectionstart.docx");

const PAGE = '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="576" w:gutter="0"/>';

const CASES = [
  { id: "A", titlePg: false, emptyHeadingCarrier: false },
  { id: "B", titlePg: false, emptyHeadingCarrier: true },
  { id: "C", titlePg: true, emptyHeadingCarrier: false },
  { id: "D", titlePg: true, emptyHeadingCarrier: true },
];

const filler = (n) => Array.from({ length: n },
  (_, i) => `<w:p><w:r><w:t xml:space="preserve">filler ${i + 1}</w:t></w:r></w:p>`).join("");

/** The paragraph that carries the sectPr and therefore ends the section. */
function terminator({ titlePg, emptyHeadingCarrier }) {
  const sectPr = `<w:sectPr>${titlePg ? "<w:titlePg/>" : ""}${PAGE}</w:sectPr>`;
  return emptyHeadingCarrier
    ? `<w:p><w:pPr><w:pStyle w:val="Heading2"/>${sectPr}</w:pPr></w:p>`
    : `<w:p><w:pPr>${sectPr}</w:pPr><w:r><w:t xml:space="preserve">end of section</w:t></w:r></w:p>`;
}

// Each case: a heading whose offset is measured, some filler, then the
// terminator that opens the next case's section.
const body = CASES.map((item) => (
  `<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t xml:space="preserve">CASE ${item.id} HEADING</w:t></w:r></w:p>` +
  filler(3) +
  terminator(item)
)).join("") +
  `<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t xml:space="preserve">CASE E HEADING</w:t></w:r></w:p>` +
  `<w:sectPr>${PAGE}</w:sectPr>`;

const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`;

// Heading2 exactly as wild2-med-nccih-protocol declares it.
const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr></w:rPrDefault>
  <w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/>
    <w:pPr><w:spacing w:before="240" w:after="120" w:line="240" w:lineRule="atLeast"/><w:outlineLvl w:val="1"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="26"/></w:rPr></w:style>
</w:styles>`;

writeFileSync(out, zipSync({
  "[Content_Types].xml": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>`),
  "_rels/.rels": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`),
  "word/_rels/document.xml.rels": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>`),
  "word/document.xml": strToU8(documentXml),
  "word/styles.xml": strToU8(stylesXml),
  // nccih carries a settings.xml with NO compatibilityMode; match that.
  "word/settings.xml": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:compat/></w:settings>`),
}));
console.log(`Wrote ${out}`);
console.log("Cases A-D end a section; E is the final page. Measure each page's first heading offset.");
