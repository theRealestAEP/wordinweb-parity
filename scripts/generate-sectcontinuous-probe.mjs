#!/usr/bin/env node

/**
 * Generate the two probes that isolate #30 — where a page created by a section
 * break puts its first paragraph.
 *
 * On wild2-med-nccih-protocol we place page 4's heading 34 CSS px lower than
 * Word's computed layout does. Page 4 begins where an empty Heading2 paragraph
 * carries BOTH a <w:sectPr> and an explicit <w:br w:type="page"/>, and the
 * section that follows declares <w:type w:val="continuous"/>.
 *
 *   probe-sectcontinuous.docx   varies the break kind, holding styles constant:
 *     A  section break, following section continuous   (= the nccih construct)
 *     B  section break, following section nextPage
 *     C  explicit page break only, no section break    (control)
 *
 *   probe-sectcontinuous-spacing.docx  varies the spacing either side of the
 *   break, so the residual offset can be attributed to a rule rather than a
 *   constant. PT1/PT2/PT3 are Heading2 clones differing only in w:spacing:
 *     V1  terminator PT1 (after 120), heading PT1 (before 240)
 *     V2  terminator PT1 (after 120), heading PT2 (before 480)
 *     V3  terminator PT3 (after 480), heading PT1 (before 240)
 *     V4  terminator PT2 (after 120), heading PT1 (before 240)
 *
 * Both are built by swapping word/document.xml into the repaired nccih package
 * rather than authoring a package from scratch. Word silently refuses to open a
 * minimal hand-built package here, and reusing the fixture's own styles.xml and
 * settings.xml is what makes the measured numbers transfer to the fixture.
 *
 *   node scripts/generate-sectcontinuous-probe.mjs
 */

import { readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const source = join(root, "parity/word-reference-docx/wild2-med-nccih-protocol.docx");
const parts = unzipSync(new Uint8Array(readFileSync(source)));

const SECT =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840" w:code="1"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="576" w:gutter="0"/>' +
  '<w:cols w:space="720"/><w:docGrid w:linePitch="360"/></w:sectPr>';
const CONTINUOUS = SECT.replace("<w:sectPr>", '<w:sectPr><w:type w:val="continuous"/>');
const PAGE_BREAK = '<w:r><w:br w:type="page"/></w:r>';

/** An empty paragraph that ends a section and forces a page, as the fixture has. */
const terminator = (style) =>
  `<w:p><w:pPr><w:pStyle w:val="${style}"/>${SECT}</w:pPr>${PAGE_BREAK}</w:p>`;
const heading = (style, text) =>
  `<w:p><w:pPr><w:pStyle w:val="${style}"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
/** An empty paragraph whose only job is to carry the section's own sectPr. */
const sectionEnd = (sect) => `<w:p><w:pPr>${sect}</w:pPr></w:p>`;

const documentXml = (body) =>
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
  '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
  `<w:body>${body}</w:body></w:document>`;

function write(name, body, styles) {
  const out = { ...parts, "word/document.xml": strToU8(documentXml(body)) };
  if (styles) out["word/styles.xml"] = strToU8(styles);
  const path = join(root, "fixtures-staging", name);
  writeFileSync(path, zipSync(out));
  console.log(`Wrote ${path}`);
}

write(
  "probe-sectcontinuous.docx",
  '<w:p><w:r><w:t>probe body line</w:t></w:r></w:p>' +
    terminator("Heading2") + heading("Heading2", "MARK A continuous section after br") + sectionEnd(CONTINUOUS) +
    terminator("Heading2") + heading("Heading2", "MARK B nextPage section after br") + sectionEnd(SECT) +
    `<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr>${PAGE_BREAK}</w:p>` +
    heading("Heading2", "MARK C br only no section break") +
    SECT,
);

const baseStyles = strFromU8(parts["word/styles.xml"]);
const heading2 = baseStyles.match(/<w:style w:type="paragraph" w:styleId="Heading2".*?<\/w:style>/s)[0];
const clone = (id, before, after) =>
  heading2
    .replace('w:styleId="Heading2"', `w:styleId="${id}"`)
    .replace(/<w:name w:val="[^"]*"\/>/, `<w:name w:val="${id}"/>`)
    .replace('<w:aliases w:val="H2-Sec. Head"/>', "")
    .replace("<w:qFormat/>", "")
    .replace(/<w:spacing [^>]*\/>/, `<w:spacing w:before="${before}" w:after="${after}" w:line="240" w:lineRule="atLeast"/>`)
    .replace(/<w:outlineLvl [^>]*\/>/, "");

write(
  "probe-sectcontinuous-spacing.docx",
  '<w:p><w:r><w:t>probe body line</w:t></w:r></w:p>' +
    terminator("PT1") + heading("PT1", "V1 term PT1 after120, mark PT1 before240") +
    terminator("PT1") + heading("PT2", "V2 term PT1 after120, mark PT2 before480") +
    terminator("PT3") + heading("PT1", "V3 term PT3 after480, mark PT1 before240") +
    terminator("PT2") + heading("PT1", "V4 term PT2 after120, mark PT1 before240") +
    SECT,
  baseStyles.replace("</w:styles>", clone("PT1", 240, 120) + clone("PT2", 480, 120) + clone("PT3", 240, 480) + "</w:styles>"),
);
