#!/usr/bin/env node

/**
 * Generate the compatibility-mode / page-size probes for Word's percentage
 * table width rule.
 *
 * The first probe (generate-pctwidth-probe.mjs) found Word ADDING the cell
 * margins to the table box at every percentage. The corpus fixture
 * parity-tables shows Word EXCLUDING them at the same percentage. The two
 * packages differ in one obvious way: parity-tables declares
 * compatibilityMode 15, the probe declared no settings.xml at all.
 *
 * These three documents are identical except for that single part, so whichever
 * way they fall is attributable to compatibility mode and nothing else:
 *
 *   probe-compat15.docx   settings.xml, compatibilityMode 15
 *   probe-compat12.docx   settings.xml, compatibilityMode 12
 *   probe-nocompat.docx   no settings.xml (what the first probe measured)
 *
 * Each carries two sections, Letter then A4, so page size is measured in the
 * same export rather than across runs. Each section holds the SAME pair of
 * tables: 90% width with 10pt left/right cell margins, and 90% width with none.
 * The difference between the pair IS the margin allowance, self-normalised
 * against that section's own content width.
 *
 *   node scripts/generate-pctwidth-compat-probe.mjs
 */

import { writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { strToU8, zipSync } from "fflate";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const fixtures = join(root, "apps/demo/public/fixtures");

const pct = (percent) => Math.round(percent * 50);
const pt = (points) => Math.round(points * 20);

const PAGES = [
  { name: "Letter", w: 12240, h: 15840 },
  { name: "A4", w: 11906, h: 16838 },
];
const MARGIN = 1440;
const PERCENT = 90;
const CELL_MARGIN_PT = 10;

const borders = "<w:tblBorders>" +
  ["top", "left", "bottom", "right", "insideH", "insideV"]
    .map((edge) => `<w:${edge} w:val="single" w:sz="8" w:color="000000"/>`).join("") +
  "</w:tblBorders>";

function table(label, marginPt) {
  const cellMargins = marginPt === null
    ? ""
    : `<w:tblCellMar><w:left w:w="${pt(marginPt)}" w:type="dxa"/><w:right w:w="${pt(marginPt)}" w:type="dxa"/></w:tblCellMar>`;
  return `<w:p><w:r><w:t xml:space="preserve">${label}</w:t></w:r></w:p>` +
    `<w:tbl><w:tblPr><w:tblW w:w="${pct(PERCENT)}" w:type="pct"/><w:tblLayout w:type="fixed"/>${borders}${cellMargins}</w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid>` +
    `<w:tr><w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>c1</w:t></w:r></w:p></w:tc>` +
    `<w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>c2</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>`;
}

const sectPr = (page) =>
  `<w:pgSz w:w="${page.w}" w:h="${page.h}"/>` +
  `<w:pgMar w:top="${MARGIN}" w:right="${MARGIN}" w:bottom="${MARGIN}" w:left="${MARGIN}" w:header="720" w:footer="720" w:gutter="0"/>`;

function section(page, last) {
  const body =
    table(`${page.name} / ${PERCENT}% / ${CELL_MARGIN_PT}pt cell margins`, CELL_MARGIN_PT) +
    table(`${page.name} / ${PERCENT}% / no cell margins`, null);
  // A non-final section carries its sectPr inside a trailing paragraph.
  return last ? body : `${body}<w:p><w:pPr><w:sectPr>${sectPr(page)}</w:sectPr></w:pPr></w:p>`;
}

const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
  PAGES.map((page, index) => section(page, index === PAGES.length - 1)).join("") +
  `<w:sectPr>${sectPr(PAGES[PAGES.length - 1])}</w:sectPr></w:body></w:document>`;

const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>
  <w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>`;

const settingsXml = (mode) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:compat><w:compatSetting w:val="${mode}" w:uri="http://schemas.microsoft.com/office/word" w:name="compatibilityMode"/></w:compat>
</w:settings>`;

function build(fileName, compatMode) {
  const withSettings = compatMode !== null;
  const parts = {
    "[Content_Types].xml": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>${
    withSettings
      ? `\n  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>`
      : ""
  }
</Types>`),
    "_rels/.rels": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`),
    "word/_rels/document.xml.rels": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>${
    withSettings
      ? `\n  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>`
      : ""
  }
</Relationships>`),
    "word/document.xml": strToU8(documentXml),
    "word/styles.xml": strToU8(stylesXml),
  };
  if (withSettings) parts["word/settings.xml"] = strToU8(settingsXml(compatMode));
  const out = join(fixtures, fileName);
  writeFileSync(out, zipSync(parts));
  console.log(`Wrote ${out}`);
}

build("probe-compat15.docx", 15);
build("probe-compat12.docx", 12);
build("probe-nocompat.docx", null);

console.log(`\nEach document: two sections (Letter, A4); each section has a ${PERCENT}% table`);
console.log(`with ${CELL_MARGIN_PT}pt left/right cell margins and one with none.`);
for (const page of PAGES) {
  const content = page.w - 2 * MARGIN;
  const base = content * PERCENT / 100 / 1440;
  const allowance = 2 * pt(CELL_MARGIN_PT) * PERCENT / 100 / 1440;
  console.log(
    `  ${page.name.padEnd(7)} content ${(content / 1440).toFixed(3)}in — expect ${base.toFixed(3)}in ` +
    `if margins excluded, ${(base + allowance).toFixed(3)}in if included (allowance ${allowance.toFixed(3)}in)`,
  );
}
