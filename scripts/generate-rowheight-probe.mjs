#!/usr/bin/env node

/**
 * Generate apps/demo/public/fixtures/probe-rowheight.docx — how a table row
 * whose height rule is "exact" treats content taller than the rule.
 *
 * This is the construct behind the toc-insert pagination difference (#29). The
 * TOC was inserted into the letterhead table of wild2-legal-ca-agreement, whose
 * row declares <w:trHeight w:hRule="exact" w:val="260"/> — 260 twips, 13pt. A
 * dozen TOC paragraphs then live in a cell whose row is fixed at 13pt tall.
 *
 * Word clips an exact row to its declared height. If the engine grows the row
 * to fit instead, the surplus has to go somewhere, and on that document it
 * became a spurious empty page.
 *
 * Three rows, same content, differing only in height rule: exact, atLeast, and
 * absent (auto). The marker paragraph after the table is what shows where the
 * table really ended.
 *
 *   node scripts/generate-rowheight-probe.mjs
 */

import { writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { strToU8, zipSync } from "fflate";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const out = join(root, "apps/demo/public/fixtures/probe-rowheight.docx");

const LINES = 8;
const EXACT = 260; // twips, as in the wild2-legal-ca-agreement letterhead row

const borders = "<w:tblBorders>" +
  ["top", "left", "bottom", "right", "insideH", "insideV"]
    .map((edge) => `<w:${edge} w:val="single" w:sz="8" w:color="000000"/>`).join("") +
  "</w:tblBorders>";

const paragraphs = (tag) => Array.from({ length: LINES },
  (_, i) => `<w:p><w:r><w:t xml:space="preserve">${tag} line ${i + 1}</w:t></w:r></w:p>`).join("");

function row(label, trPr) {
  return `<w:tr>${trPr}<w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/></w:tcPr>${paragraphs(label)}</w:tc></w:tr>`;
}

const table = `<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/><w:tblLayout w:type="fixed"/>${borders}</w:tblPr>` +
  `<w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid>` +
  row("EXACT", `<w:trPr><w:cantSplit/><w:trHeight w:hRule="exact" w:val="${EXACT}"/></w:trPr>`) +
  row("ATLEAST", `<w:trPr><w:cantSplit/><w:trHeight w:hRule="atLeast" w:val="${EXACT}"/></w:trPr>`) +
  row("AUTO", "") +
  `</w:tbl>`;

const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
<w:p><w:r><w:t xml:space="preserve">MARKER TOP</w:t></w:r></w:p>
${table}
<w:p><w:r><w:t xml:space="preserve">MARKER AFTER TABLE</w:t></w:r></w:p>
<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>
</w:body></w:document>`;

writeFileSync(out, zipSync({
  "[Content_Types].xml": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>`),
  "_rels/.rels": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`),
  "word/_rels/document.xml.rels": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>`),
  "word/document.xml": strToU8(documentXml),
  "word/settings.xml": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:compat><w:compatSetting w:val="15" w:uri="http://schemas.microsoft.com/office/word" w:name="compatibilityMode"/></w:compat>
</w:settings>`),
}));
console.log(`Wrote ${out}`);
console.log(`Each row holds ${LINES} paragraphs; the exact/atLeast rows declare ${EXACT} twips (${EXACT / 20}pt).`);
