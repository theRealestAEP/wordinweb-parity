#!/usr/bin/env node

/**
 * Generate apps/demo/public/fixtures/probe-pctwidth.docx.
 *
 * Settles one question: when a fixed-layout table has a PERCENTAGE width and
 * non-default cell margins, does Word paint the table box at the percentage of
 * the content width, or at the percentage plus the cell margins?
 *
 * The engine currently allows for the margins only at exactly 100% and excludes
 * them below, and each branch of that rule rests on a single measurement. This
 * probe puts four percentages (90/95/99/100) next to each other with identical
 * 10pt left/right cell margins, plus two controls at the same percentages with
 * default margins. Comparing a percentage against its own control isolates the
 * margin allowance from everything else, and comparing across percentages shows
 * whether the switch really happens at 100%.
 *
 * Letter, 1in margins: the content width is 9360 twips (6.5in) exactly, so the
 * expected widths are round numbers.
 *
 *   node scripts/generate-pctwidth-probe.mjs
 */

import { writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { strToU8, zipSync } from "fflate";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const out = join(root, "apps/demo/public/fixtures/probe-pctwidth.docx");

/** w:tblW pct is in fiftieths of a percent, so 5000 is 100%. */
const pct = (percent) => Math.round(percent * 50);
/** Word measures in twentieths of a point. */
const pt = (points) => Math.round(points * 20);

const CONTENT_TWIPS = 9360; // 12240 page - 2 x 1440 margin

const cases = [
  { label: "A: 90% width, 10pt cell margins", percent: 90, marginPt: 10 },
  { label: "B: 95% width, 10pt cell margins", percent: 95, marginPt: 10 },
  { label: "C: 99% width, 10pt cell margins", percent: 99, marginPt: 10 },
  { label: "D: 100% width, 10pt cell margins", percent: 100, marginPt: 10 },
  // Controls: same percentages, no cell margins (this package declares no
  // styles.xml, so absent tblCellMar means zero). Any width difference against
  // A and D is the margin allowance and nothing else.
  { label: "E: 100% width, no cell margins", percent: 100, marginPt: null },
  { label: "F: 90% width, no cell margins", percent: 90, marginPt: null },
  // The grid in cases A-F sums to exactly the content width. These two ask
  // whether that is what decides the margin allowance: the gate's parity-tables
  // measurement, which showed Word EXCLUDING the margins at 90%, had a grid
  // summing to well over the content width.
  { label: "G: 90% width, 10pt margins, oversized grid", percent: 90, marginPt: 10, grid: [2160, 813, 7724] },
  { label: "H: 90% width, 10pt margins, undersized grid", percent: 90, marginPt: 10, grid: [1200, 1200, 1200] },
];

const borders = `<w:tblBorders>
  <w:top w:val="single" w:sz="8" w:color="000000"/>
  <w:left w:val="single" w:sz="8" w:color="000000"/>
  <w:bottom w:val="single" w:sz="8" w:color="000000"/>
  <w:right w:val="single" w:sz="8" w:color="000000"/>
  <w:insideH w:val="single" w:sz="8" w:color="000000"/>
  <w:insideV w:val="single" w:sz="8" w:color="000000"/>
</w:tblBorders>`;

function table({ label, percent, marginPt, grid = [4680, 4680] }) {
  const cellMargins = marginPt === null
    ? ""
    : `<w:tblCellMar><w:left w:w="${pt(marginPt)}" w:type="dxa"/><w:right w:w="${pt(marginPt)}" w:type="dxa"/></w:tblCellMar>`;
  const cell = (width, text) =>
    `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:tc>`;
  return `<w:p><w:r><w:t xml:space="preserve">${label}</w:t></w:r></w:p>
<w:tbl>
  <w:tblPr>
    <w:tblW w:w="${pct(percent)}" w:type="pct"/>
    <w:tblLayout w:type="fixed"/>
    ${borders}
    ${cellMargins}
  </w:tblPr>
  <w:tblGrid>${grid.map((width) => `<w:gridCol w:w="${width}"/>`).join("")}</w:tblGrid>
  <w:tr>${grid.map((width, index) => cell(width, `c${index + 1}`)).join("")}</w:tr>
</w:tbl>
<w:p/>`;
}

const document = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t xml:space="preserve">Fixed-layout percentage table widths with and without custom cell margins.</w:t></w:r></w:p>
    ${cases.map(table).join("\n")}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`;

writeFileSync(out, zipSync({
  "[Content_Types].xml": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`),
  "_rels/.rels": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`),
  "word/document.xml": strToU8(document),
}));

console.log(`Wrote ${out}`);
console.log(`Content width ${CONTENT_TWIPS} twips (6.5in). Expected painted widths, in inches:`);
for (const item of cases) {
  const inside = CONTENT_TWIPS * item.percent / 100;
  const plusMargins = (CONTENT_TWIPS + 2 * pt(item.marginPt ?? 0)) * item.percent / 100;
  console.log(
    `  ${item.label.padEnd(38)} margins inside ${(inside / 1440).toFixed(3)}` +
    `   margins added ${(plusMargins / 1440).toFixed(3)}`,
  );
}
