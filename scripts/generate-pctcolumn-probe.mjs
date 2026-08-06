#!/usr/bin/env node

/**
 * Generate apps/demo/public/fixtures/probe-pctcolumn.docx — Word ground truth
 * for how a fixed-layout PERCENTAGE table divides its width between columns.
 *
 * The total width is settled (compatibilityMode decides the cell-margin
 * allowance; see generate-pctwidth-compat-probe.mjs). What is not settled is
 * the split: on parity-tables we scale the saved grid strictly proportionally
 * and Word does not — it gives the two narrow columns MORE than their share and
 * the wide column less. The working hypothesis is that Word scales the grid but
 * clamps each column to a minimum driven by its content (longest unbreakable
 * word plus cell margins) and takes the shortfall out of the columns that have
 * room.
 *
 * Every table here is compatibilityMode 15, fixed layout, one content row and
 * one EMPTY row. The empty row is what gets measured: its vertical rules are
 * the column boundaries with no glyphs in the way. One dimension varies per
 * case, from a baseline shaped like parity-tables:
 *
 *   grid proportions, percentage, content lengths, column count, cell margins.
 *
 *   node scripts/generate-pctcolumn-probe.mjs
 */

import { writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { strToU8, zipSync } from "fflate";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const out = join(root, "apps/demo/public/fixtures/probe-pctcolumn.docx");

const pct = (percent) => Math.round(percent * 50);
const pt = (points) => Math.round(points * 20);

const PAGE = { w: 12240, h: 15840 };
const MARGIN = 1440;
const CONTENT = PAGE.w - 2 * MARGIN; // 9360 twips = 6.5in

const SHORT_A = "Key";
const SHORT_B = "Status";
const LONG = "A much longer description cell that should dominate the width";

/** Baseline is the shape parity-tables has after the table-widths scenario. */
const BASE = { grid: [2160, 813, 7724], percent: 90, marginPt: 10, cells: [SHORT_A, SHORT_B, LONG] };

const cases = [
  { id: "A1", note: "baseline: parity-tables grid", ...BASE },
  { id: "A2", note: "even grid", ...BASE, grid: [1000, 1000, 1000] },
  { id: "A3", note: "extreme grid", ...BASE, grid: [100, 100, 10000] },
  { id: "B2", note: "50% width", ...BASE, percent: 50 },
  { id: "C1", note: "all-short content", ...BASE, cells: ["Key", "St", "Desc"] },
  { id: "C2", note: "middle cell empty", ...BASE, cells: [SHORT_A, "", LONG] },
  { id: "C3", note: "middle barely wrapping", ...BASE, cells: [SHORT_A, "Statuses", LONG] },
  { id: "D1", note: "two columns", ...BASE, grid: [2160, 7724], cells: [SHORT_A, LONG] },
  { id: "D2", note: "five columns", ...BASE, grid: [1000, 1000, 1000, 1000, 6000], cells: ["Key", "St", "Qty", "Rate", LONG] },
  { id: "E1", note: "no cell margins", ...BASE, marginPt: null },
];

const borders = "<w:tblBorders>" +
  ["top", "left", "bottom", "right", "insideH", "insideV"]
    .map((edge) => `<w:${edge} w:val="single" w:sz="8" w:color="000000"/>`).join("") +
  "</w:tblBorders>";

function table({ id, note, grid, percent, marginPt, cells }) {
  const cellMargins = marginPt === null
    ? ""
    : `<w:tblCellMar><w:left w:w="${pt(marginPt)}" w:type="dxa"/><w:right w:w="${pt(marginPt)}" w:type="dxa"/></w:tblCellMar>`;
  const cell = (width, text) =>
    `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr><w:p>${
      text ? `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>` : ""
    }</w:p></w:tc>`;
  const contentRow = `<w:tr>${grid.map((width, index) => cell(width, cells[index] ?? "")).join("")}</w:tr>`;
  // The measured row: empty everywhere, so its vertical rules are the column
  // boundaries and nothing else.
  const emptyRow = `<w:tr>${grid.map((width) => cell(width, "")).join("")}</w:tr>`;
  return `<w:p><w:r><w:t xml:space="preserve">${id}: ${note}</w:t></w:r></w:p>` +
    `<w:tbl><w:tblPr><w:tblW w:w="${pct(percent)}" w:type="pct"/><w:tblLayout w:type="fixed"/>${borders}${cellMargins}</w:tblPr>` +
    `<w:tblGrid>${grid.map((width) => `<w:gridCol w:w="${width}"/>`).join("")}</w:tblGrid>` +
    `${contentRow}${emptyRow}</w:tbl><w:p/>`;
}

const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
  cases.map(table).join("") +
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
// The expectations the measurement is held against: strict proportional
// scaling of the saved grid onto the table's own painted width.
const rows = cases.map((item) => {
  const sum = item.grid.reduce((total, width) => total + width, 0);
  const available = CONTENT * item.percent / 100;
  return {
    id: item.id,
    note: item.note,
    grid: item.grid.join("/"),
    gridSum: sum,
    availableIn: available / 1440,
    proportionalPx: item.grid.map((width) => Math.round(width / sum * (available / 1440) * 192)).join("/"),
  };
});
console.table(rows);
