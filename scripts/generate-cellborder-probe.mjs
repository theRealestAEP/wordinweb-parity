#!/usr/bin/env node

/**
 * Generate the row-pitch probe for cell (tcBorders) against table
 * (tblBorders) rules.
 *
 * probe3-lo-provenance's rows run ~0.9px tall each in the engine: its double
 * sz-2 tcBorders rules leave Word's row pitch unchanged inside 3pt cell
 * margins, where the engine charges half the rule on each side. Yet
 * parity2-nestedtables (single sz-4 tblBorders, 2pt margins) does charge its
 * rules. The two differ in border source, border style, cell margin and
 * compatibility mode (12 against 15), so each is varied here on its own.
 *
 * Eight one-column tables of six single-line rows each, labelled TA..TH:
 *
 *   TA  tcBorders  single sz4   margins 60tw
 *   TB  tcBorders  double sz2   margins 60tw   (probe3-lo-provenance)
 *   TC  tblBorders single sz4   margins 60tw
 *   TD  tblBorders double sz2   margins 60tw
 *   TE  tcBorders  single sz24  margins 60tw   (rule as wide as the margin)
 *   TF  tcBorders  single sz4   margins 0
 *   TG  no borders              margins 60tw   (baseline)
 *   TH  no borders              margins 0      (baseline)
 *
 * And four more that add what probe3-lo-provenance's cells carry, one thing at
 * a time: 283tw of spacing after (TI, TJ without rules), then vAlign center
 * on top of that (TK, TL without rules).
 *
 *   TI  tcBorders  double sz2   margins 60tw   after 283
 *   TJ  no borders              margins 60tw   after 283  (baseline for TI)
 *   TK  tcBorders  double sz2   margins 60tw   after 283, vAlign center
 *   TL  no borders              margins 60tw   after 283, vAlign center
 *
 * Then LibreOffice's pattern, which probe3-lo-provenance turned out to hinge
 * on: body-row cells carry left, bottom and right rules but NO top, so each
 * inner rule is only the bottom of the row above. The first row also has a
 * top. And its mirror image, top rules only (the last row adds a bottom).
 *
 *   TM  tcBorders  single sz4   bottom-only inner rules
 *   TN  tcBorders  double sz2   bottom-only inner rules
 *   TO  tcBorders  single sz4   top-only inner rules
 *   TP  tcBorders  double sz2   top-only inner rules
 *
 * wild-doerfp also draws bottom-only rules, with ZERO cell margins, and there
 * the rules are charged. So: bottom-only rules against the cell margins
 * (top/bottom, in twips), to find how much margin absorbs how much rule and
 * which side's margin it is.
 *
 *   TQ  bottom-only single sz4    margins 0/0
 *   TR  bottom-only single sz24   margins 20/20
 *   TS  bottom-only single sz48   margins 60/60
 *   TT  bottom-only single sz24   margins 60/60
 *   TU  no borders                margins 20/20   (baseline for TR)
 *   TV  bottom-only single sz24   margins 0/60
 *   TW  bottom-only single sz24   margins 60/0
 *   TX  no borders                margins 0/60    (baseline for TV, TW)
 *
 * Three settings variants: compatibilityMode 15, 12, and 12 with
 * w:adjustLineHeightInTable (probe3-lo-provenance carries both). Read the
 * row pitch from the baselines of each table's cell text in Word's PDF and
 * subtract the matching baseline table's pitch: what is left is what the rule
 * charges a row.
 *
 *   node scripts/generate-cellborder-probe.mjs
 */

import { readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const source = join(root, "apps/demo/public/fixtures/wild2-legal-ca-agreement.docx");

const SECT =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="720" w:right="1440" w:bottom="720" w:left="1440" w:header="360" w:footer="360" w:gutter="0"/>' +
  '<w:cols w:space="720"/></w:sectPr>';

const para = (text, after = 0) =>
  `<w:p><w:pPr><w:spacing w:before="0" w:after="${after}" w:line="240" w:lineRule="auto"/><w:rPr><w:sz w:val="20"/></w:rPr></w:pPr>` +
  `<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;

const edges = (val, sz, sides) => sides.map((s) => `<w:${s} w:val="${val}" w:sz="${sz}" w:space="0" w:color="000000"/>`).join("");

const TABLES = [
  ["TA", { source: "tc", val: "single", sz: 4, mar: 60 }],
  ["TB", { source: "tc", val: "double", sz: 2, mar: 60 }],
  ["TC", { source: "tbl", val: "single", sz: 4, mar: 60 }],
  ["TD", { source: "tbl", val: "double", sz: 2, mar: 60 }],
  ["TE", { source: "tc", val: "single", sz: 24, mar: 60 }],
  ["TF", { source: "tc", val: "single", sz: 4, mar: 0 }],
  ["TG", { source: null, mar: 60 }],
  ["TH", { source: null, mar: 0 }],
  ["TI", { source: "tc", val: "double", sz: 2, mar: 60, after: 283 }],
  ["TJ", { source: null, mar: 60, after: 283 }],
  ["TK", { source: "tc", val: "double", sz: 2, mar: 60, after: 283, center: true }],
  ["TL", { source: null, mar: 60, after: 283, center: true }],
  ["TM", { source: "tc", val: "single", sz: 4, mar: 60, inner: "bottom" }],
  ["TN", { source: "tc", val: "double", sz: 2, mar: 60, inner: "bottom" }],
  ["TO", { source: "tc", val: "single", sz: 4, mar: 60, inner: "top" }],
  ["TP", { source: "tc", val: "double", sz: 2, mar: 60, inner: "top" }],
  ["TQ", { source: "tc", val: "single", sz: 4, mar: 0, inner: "bottom" }],
  ["TR", { source: "tc", val: "single", sz: 24, mar: 20, inner: "bottom" }],
  ["TS", { source: "tc", val: "single", sz: 48, mar: 60, inner: "bottom" }],
  ["TT", { source: "tc", val: "single", sz: 24, mar: 60, inner: "bottom" }],
  ["TU", { source: null, mar: 20 }],
  ["TV", { source: "tc", val: "single", sz: 24, mar: [0, 60], inner: "bottom" }],
  ["TW", { source: "tc", val: "single", sz: 24, mar: [60, 0], inner: "bottom" }],
  ["TX", { source: null, mar: [0, 60] }],
];

function table(label, { source, val, sz, mar, after = 0, center = false, inner }) {
  const tblBorders =
    source === "tbl"
      ? `<w:tblBorders>${edges(val, sz, ["top", "left", "bottom", "right", "insideH", "insideV"])}</w:tblBorders>`
      : `<w:tblBorders>${["top", "left", "bottom", "right", "insideH", "insideV"].map((s) => `<w:${s} w:val="nil"/>`).join("")}</w:tblBorders>`;
  // `inner` keeps only one of the two rules at each row boundary.
  const tcBorders = (i) => {
    if (source !== "tc") return "";
    const sides = ["top", "left", "bottom", "right"].filter(
      (s) => !inner || (s === "top" ? inner === "top" || i === 0 : s === "bottom" ? inner === "bottom" || i === 5 : true),
    );
    return `<w:tcBorders>${edges(val, sz, sides)}</w:tcBorders>`;
  };
  const [top, bottom] = Array.isArray(mar) ? mar : [mar, mar];
  const cellMar = `<w:tblCellMar><w:top w:w="${top}" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="${bottom}" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>`;
  const rows = Array.from(
    { length: 6 },
    (_, i) =>
      `<w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/>${tcBorders(i)}${center ? `<w:vAlign w:val="center"/>` : ""}</w:tcPr>` +
      `${para(`${label}R${i + 1}`, after)}</w:tc></w:tr>`,
  ).join("");
  return (
    para(`${label} ${source ?? "none"} ${val ?? ""} ${sz ?? ""} mar${mar}`) +
    `<w:tbl><w:tblPr><w:tblW w:w="3000" w:type="dxa"/>${tblBorders}<w:tblLayout w:type="fixed"/>${cellMar}</w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>${rows}</w:tbl>` +
    para("")
  );
}

const documentXml =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  `<w:body>${TABLES.map(([l, o]) => table(l, o)).join("")}${SECT}</w:body></w:document>`;

const VARIANTS = [
  ["probe-cellborder15", `<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>`],
  ["probe-cellborder12", `<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="12"/>`],
  ["probe-cellborder12a", `<w:adjustLineHeightInTable/><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="12"/>`],
];

for (const [name, compat] of VARIANTS) {
  const parts = unzipSync(new Uint8Array(readFileSync(source)));
  parts["word/document.xml"] = strToU8(documentXml);
  // CT_Settings is a sequence: w:compat goes ahead of w:rsids (or whichever
  // of its followers comes first).
  const settings = strFromU8(parts["word/settings.xml"]).replace(/<w:compat>[\s\S]*?<\/w:compat>|<w:compat\/>/, "");
  parts["word/settings.xml"] = strToU8(settings.replace(/(<w:rsids>|<m:mathPr>|<w:themeFontLang)/, `<w:compat>${compat}</w:compat>$1`));
  const out = join(root, `fixtures-staging/${name}.docx`);
  writeFileSync(out, Buffer.from(zipSync(parts)));
  console.log(`wrote ${out}`);
}
