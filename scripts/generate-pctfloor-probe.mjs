#!/usr/bin/env node

/**
 * Generate apps/demo/public/fixtures/probe-pctfloor.docx — pins the floor Word
 * applies when a gridCol is narrower than the cell margins themselves.
 *
 * The column-distribution rule (probe-pctcolumn) is:
 *
 *   painted[i] = margins + (grid[i] - margins) / SUM(grid[j] - margins)
 *                          * (tableWidth - n * margins)
 *
 * which matches Word within 1px except when `grid[i] - margins` goes negative.
 * There Word does not use zero: the extreme case painted 56px where the model
 * said 53px. This sweeps one column's grid value from well above the margin
 * width down to far below it, against a fixed wide column.
 *
 * The margins here are 10pt per side, so 400 twips PER COLUMN in total, and the
 * sweep brackets that. If the floor is a constant, every grid value at or below
 * it paints the same width — the plateau is the answer, and where the plateau
 * starts confirms what the subtraction is measured against.
 *
 *   node scripts/generate-pctfloor-probe.mjs
 */

import { writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { strToU8, zipSync } from "fflate";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const out = join(root, "apps/demo/public/fixtures/probe-pctfloor.docx");

const PAGE = { w: 12240, h: 15840 };
const MARGIN = 1440;
const CONTENT = PAGE.w - 2 * MARGIN; // 9360 twips
const PERCENT = 90;
const WIDE = 9000;

/**
 * The sweep runs at two margin settings. One sweep gives the floor's value; two
 * say whether it is an absolute constant or scales with the margins — which is
 * the difference between "floor = 420tw" and "floor = margins + 1pt".
 */
const MARGIN_SETS_PT = [10, 20];

/** Grid values for the swept narrow column, in twips, bracketing both margin widths. */
const SWEEP = [1600, 1000, 900, 850, 820, 800, 700, 600, 500, 450, 420, 400, 300, 200, 100, 20];

const borders = "<w:tblBorders>" +
  ["top", "left", "bottom", "right", "insideH", "insideV"]
    .map((edge) => `<w:${edge} w:val="single" w:sz="8" w:color="000000"/>`).join("") +
  "</w:tblBorders>";

function table(narrow, marginPt) {
  const cellMargins =
    `<w:tblCellMar><w:left w:w="${marginPt * 20}" w:type="dxa"/><w:right w:w="${marginPt * 20}" w:type="dxa"/></w:tblCellMar>`;
  const cell = (width, text) =>
    `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr><w:p>${
      text ? `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>` : ""
    }</w:p></w:tc>`;
  return `<w:p><w:r><w:t xml:space="preserve">margins ${marginPt}pt / grid ${narrow}tw</w:t></w:r></w:p>` +
    `<w:tbl><w:tblPr><w:tblW w:w="${PERCENT * 50}" w:type="pct"/><w:tblLayout w:type="fixed"/>${borders}${cellMargins}</w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="${narrow}"/><w:gridCol w:w="${WIDE}"/></w:tblGrid>` +
    `<w:tr>${cell(narrow, "x")}${cell(WIDE, "A wide column that takes whatever is left over")}</w:tr>` +
    `<w:tr>${cell(narrow, "")}${cell(WIDE, "")}</w:tr>` +
    `</w:tbl><w:p/>`;
}

const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
  MARGIN_SETS_PT.flatMap((marginPt) => SWEEP.map((narrow) => table(narrow, marginPt))).join("") +
  `<w:sectPr><w:pgSz w:w="${PAGE.w}" w:h="${PAGE.h}"/>` +
  `<w:pgMar w:top="${MARGIN}" w:right="${MARGIN}" w:bottom="${MARGIN}" w:left="${MARGIN}" w:header="720" w:footer="720" w:gutter="0"/>` +
  `</w:sectPr></w:body></w:document>`;

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
  "word/styles.xml": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>
  <w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>`),
  "word/settings.xml": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:compat><w:compatSetting w:val="15" w:uri="http://schemas.microsoft.com/office/word" w:name="compatibilityMode"/></w:compat>
</w:settings>`),
}));

console.log(`Wrote ${out}`);
const width = CONTENT * PERCENT / 100;
for (const marginPt of MARGIN_SETS_PT) {
  const margins = marginPt * 20 * 2;
  const avail = width - 2 * margins;
  console.log(`margins ${marginPt}pt (${margins}tw/column): table ${width}tw, distributable ${avail}tw`);
}
