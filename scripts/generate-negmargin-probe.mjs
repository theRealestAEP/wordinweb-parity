#!/usr/bin/env node

/**
 * Generate the probe that locates the BODY TOP under a negative `w:top` margin.
 *
 * `wild3-template-caed-pleading` is the corpus's only fixture with negative page
 * margins (`w:top="-1325"`, `w:bottom="-1267"`), and against a reference exported
 * from the committed fixture by current Word our body sits 160.25 CSS px too
 * high — a clean uniform translation with identical horizontal positions and no
 * reflow. The header is placed correctly to a rounding pixel; only the body
 * origin moves, so one number explains the whole 36.81%.
 *
 * The fixture's header is a pleading line-number column: a single-row table with
 * `<w:trHeight w:hRule="exact" w:val="14880"/>`, 744 pt tall, on a 792 pt page.
 * A header that large is exactly the case where Word's ordinary "push the body
 * below the header" rule and the authored top margin disagree, and the negative
 * margin is presumably the author overriding it. Which of the two Word obeys,
 * and how a negative value participates, is what this measures.
 *
 * Each case is its own section, so it gets its own page and its own `w:pgMar`.
 * The section's first body paragraph is a marker with an exact line height, so
 * its top in the PDF IS the body top with nothing inferred. Three variables move
 * one at a time:
 *
 *   w:top      swept through the negative range, past zero, into the positive
 *   header height  tall (the fixture's 14880 tw) against short (2880 tw)
 *   w:header   the header's own distance from the page top
 *
 * If the body origin tracks `|w:top|` from the page edge it is flat in the
 * header variables; if it tracks the header it is flat in `w:top`. Two settings
 * of each is what separates those, and one setting of either cannot.
 *
 *   node scripts/generate-negmargin-probe.mjs
 */

import { readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { strToU8, unzipSync, zipSync } from "fflate";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const source = join(root, "apps/demo/public/fixtures/wild3-template-caed-pleading.docx");
const parts = unzipSync(new Uint8Array(readFileSync(source)));

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

/**
 * A header of an exactly known height: one table row with an exact `trHeight`,
 * carrying a single mark so the header's own placement stays readable.
 */
const headerXml = (twips, tag) =>
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
  `<w:hdr ${NS}><w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/><w:tblLayout w:type="fixed"/>` +
  '<w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid>' +
  `<w:tr><w:trPr><w:trHeight w:hRule="exact" w:val="${twips}"/></w:trPr>` +
  '<w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/></w:tcPr>' +
  '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/>' +
  `<w:rPr><w:sz w:val="20"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>${tag}</w:t></w:r></w:p>` +
  "</w:tc></w:tr></w:tbl>" +
  '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/></w:pPr></w:p></w:hdr>';

const TALL = 14880; // the fixture's own row: 744 pt
const SHORT = 2880; // 144 pt

parts["word/header1.xml"] = strToU8(headerXml(TALL, "HDRTALL"));
parts["word/header2.xml"] = strToU8(headerXml(SHORT, "HDRSHORT"));

/** rId7 is the even header, rId8 the default. Both are rewritten above. */
const TALL_ID = "rId7";
const SHORT_ID = "rId8";

/** A marker line of exactly known height, so its top is the body top. */
const marker = (id) =>
  '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/>' +
  '<w:rPr><w:sz w:val="20"/></w:rPr></w:pPr>' +
  `<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>${id}</w:t></w:r></w:p>`;

const sectPr = ({ headerId, top, header }) =>
  `<w:sectPr><w:headerReference w:type="default" r:id="${headerId}"/>` +
  '<w:type w:val="nextPage"/><w:pgSz w:w="12240" w:h="15840" w:code="1"/>' +
  `<w:pgMar w:top="${top}" w:right="720" w:bottom="1440" w:left="2088" ` +
  `w:header="${header}" w:footer="360" w:gutter="0"/>` +
  '<w:cols w:space="720"/></w:sectPr>';

const TOPS = [-2880, -2160, -1440, -1325, -720, -360, 0, 720, 1440];

const cases = [];
const add = (id, opts) => cases.push({ id, ...opts });

// A: the fixture's header height and header distance, sweeping w:top.
for (const top of TOPS) add(`A${TOPS.indexOf(top)}`, { headerId: TALL_ID, top, header: 432 });
// B: the same sweep with a short header. Flat difference means the header is inert.
for (const top of TOPS) add(`B${TOPS.indexOf(top)}`, { headerId: SHORT_ID, top, header: 432 });
// C and D: the header's own distance, at the fixture's top margin and at zero.
add("C0", { headerId: TALL_ID, top: -1325, header: 1440 });
add("C1", { headerId: TALL_ID, top: 0, header: 1440 });
add("D0", { headerId: SHORT_ID, top: -1325, header: 1440 });
add("D1", { headerId: SHORT_ID, top: 0, header: 1440 });

const body = cases
  .map((c, i) => {
    const last = i === cases.length - 1;
    const mark = marker(`MK${c.id}X`);
    // Every section but the last carries its sectPr on a trailing paragraph.
    return last ? mark : `${mark}<w:p><w:pPr>${sectPr(c)}</w:pPr></w:p>`;
  })
  .join("");

const documentXml =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
  `<w:document ${NS}><w:body>${body}${sectPr(cases.at(-1))}</w:body></w:document>`;

parts["word/document.xml"] = strToU8(documentXml);

const path = join(root, "fixtures-staging", "probe-negmargin.docx");
writeFileSync(path, zipSync(parts));
console.log(`Wrote ${path} — ${cases.length} cases, tops ${TOPS.join(", ")} twips`);
