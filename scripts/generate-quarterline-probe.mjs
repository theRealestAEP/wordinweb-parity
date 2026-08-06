#!/usr/bin/env node

/**
 * Generate the probe that isolates #38's second rule — what a paragraph whose
 * `w:lineRule="auto"` multiple is a FRACTION of a line costs in height.
 *
 * On wild2-legal-ca-agreement page 1 we sit 1.7 px below Word after the
 * bordered-run step is accounted for, and the only unusual construct left in
 * that stretch is two EMPTY paragraphs carrying
 * `<w:spacing w:line="60" w:lineRule="auto"/>` — a quarter line — at 13 pt.
 * Empty paragraphs have no text to measure between, so the fixture itself
 * cannot answer what either one costs. This probe puts text either side of one.
 *
 * Each case is alone on a page, so no case can disturb another and no case can
 * split across a break:
 *
 *   ABOVE   a plain 13/26 pt bold line, spacing before=0 after=0 line=240 auto
 *   TEST    the paragraph under test, same before/after, varying w:line
 *   BELOW   a second plain line, identical to ABOVE
 *
 * `top(BELOW) - top(ABOVE)` is then `height(ABOVE line) + height(TEST)`, and
 * case 01 — which omits TEST entirely — measures `height(ABOVE line)` on its
 * own. Subtracting gives TEST's height with nothing else in it.
 *
 * The sweep runs at TWO font sizes, 13 pt (the fixture's) and 26 pt, because a
 * rule that reads "a fraction of the line" and a rule that reads "a fixed
 * floor" agree at one size and disagree at two.
 *
 *   01  no TEST paragraph                       (control: the ABOVE line alone)
 *   02  TEST empty, no w:spacing                (control: a default empty line)
 *   03  TEST empty, w:line=60   auto            (the fixture's construct)
 *   04  TEST empty, w:line=120  auto
 *   05  TEST empty, w:line=240  auto
 *   06  TEST empty, w:line=480  auto
 *   07  TEST holds "x", w:line=60  auto
 *   08  TEST holds "x", w:line=120 auto
 *   09  TEST holds "x", w:line=240 auto
 *   10  TEST holds "x", w:line=480 auto
 *   11  TWO empty w:line=60 paragraphs          (does the cost add?)
 *
 * Marks read `S13C03A` / `S13C03B`, so one pdftotext pass locates every line in
 * Word's own render and the same strings locate them in the browser.
 *
 * Built by swapping word/document.xml into the agreement's own package, as
 * scripts/generate-sectcontinuous-probe.mjs does: reusing the fixture's
 * styles.xml and settings.xml is what makes the numbers transfer back to it.
 *
 *   node scripts/generate-quarterline-probe.mjs
 */

import { readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { strToU8, unzipSync, zipSync } from "fflate";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const source = join(root, "apps/demo/public/fixtures/wild2-legal-ca-agreement.docx");
const parts = unzipSync(new Uint8Array(readFileSync(source)));

/** A bare section: the fixture's page geometry, without its headers or footers. */
const SECT =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  "<w:cols w:space=\"720\"/></w:sectPr>";

const rpr = (sz) => `<w:rPr><w:b/><w:sz w:val="${sz}"/></w:rPr>`;
const SPACING = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>';

/** A plain marked line. Left aligned, so pdftotext reports a stable x. */
const mark = (sz, text) =>
  `<w:p><w:pPr>${SPACING}${rpr(sz)}</w:pPr><w:r>${rpr(sz)}<w:t>${text}</w:t></w:r></w:p>`;

/**
 * The paragraph under test. `line` null means "author no w:spacing at all".
 * Centred and bold at `sz`, exactly as the fixture's two empty paragraphs are.
 */
const test = (sz, line, text) => {
  const spacing =
    line === null ? "" : `<w:spacing w:before="0" w:after="0" w:line="${line}" w:lineRule="auto"/>`;
  const run = text ? `<w:r>${rpr(sz)}<w:t>${text}</w:t></w:r>` : "";
  return `<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="3600"/></w:tabs>${spacing}<w:jc w:val="center"/>${rpr(sz)}</w:pPr>${run}</w:p>`;
};

const PAGE_BREAK = '<w:p><w:pPr>' + SPACING + '</w:pPr><w:r><w:br w:type="page"/></w:r></w:p>';

/** [case number, how the TEST paragraph is built] */
const cases = [
  ["01", () => ""],
  ["02", (sz) => test(sz, null, "")],
  ["03", (sz) => test(sz, 60, "")],
  ["04", (sz) => test(sz, 120, "")],
  ["05", (sz) => test(sz, 240, "")],
  ["06", (sz) => test(sz, 480, "")],
  ["07", (sz) => test(sz, 60, "x")],
  ["08", (sz) => test(sz, 120, "x")],
  ["09", (sz) => test(sz, 240, "x")],
  ["10", (sz) => test(sz, 480, "x")],
  ["11", (sz) => test(sz, 60, "") + test(sz, 60, "")],
];

const groups = [
  ["S13", 26],
  ["S26", 52],
];

const body = groups
  .flatMap(([tag, sz], groupIdx) =>
    cases.map(([num, build], caseIdx) => {
      const id = `${tag}C${num}`;
      const page = groupIdx === 0 && caseIdx === 0 ? "" : PAGE_BREAK;
      return page + mark(sz, `${id}A`) + build(sz) + mark(sz, `${id}B`);
    }),
  )
  .join("");

const documentXml =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
  '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
  `<w:body>${body}${SECT}</w:body></w:document>`;

function write(name, docBody, note) {
  const xml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    `<w:body>${docBody}${SECT}</w:body></w:document>`;
  const path = join(root, "fixtures-staging", name);
  writeFileSync(path, zipSync({ ...parts, "word/document.xml": strToU8(xml) }));
  console.log(`Wrote ${path} — ${note}`);
}

write("probe-quarterline.docx", body, `${groups.length * cases.length} cases, one per page`);

/**
 * The companion probe. The sweep above showed both renderers producing the SAME
 * two line heights at 13 pt — 20.000 and 19.667 CSS px — and handing them to
 * opposite paragraphs: Word gives 20.000 to the marker lines and 19.667 to the
 * paragraph under test, and we do the reverse. The marker and the test
 * paragraph carry the same font, size and w:spacing, so the difference has to
 * come from one of the three things that still differ between them, and this
 * varies them ONE AT A TIME over a paragraph that is otherwise a marker:
 *
 *   D01  marker shape exactly, text "y"          (control: expect the 19.672)
 *   D02  + <w:jc w:val="center"/>
 *   D03  + <w:tabs>
 *   D04  + both                                  (= the test shape, short text)
 *   D05  marker shape, long text                 (content length, alone)
 *   D06  both, long text                         (= the test shape, long text)
 *   D07  no paragraph at all                     (control)
 */
const TABS = '<w:tabs><w:tab w:val="left" w:pos="3600"/></w:tabs>';
const JC = '<w:jc w:val="center"/>';

/** [case, tabs?, centred?, text] — null text means "author no paragraph at all". */
const shapes = [
  ["01", false, false, "y"],
  ["02", false, true, "y"],
  ["03", true, false, "y"],
  ["04", true, true, "y"],
  ["05", false, false, "SHAPE05TEXT"],
  ["06", true, true, "SHAPE06TEXT"],
  ["07", false, false, null],
];

const shapeBody = shapes
  .map(([num, tabs, centred, text], idx) => {
    const id = `S13D${num}`;
    // w:tabs precedes w:spacing and w:jc follows it — CT_PPr is a sequence.
    const mid =
      text === null
        ? ""
        : `<w:p><w:pPr>${tabs ? TABS : ""}${SPACING}${centred ? JC : ""}${rpr(26)}</w:pPr>` +
          `<w:r>${rpr(26)}<w:t>${text}</w:t></w:r></w:p>`;
    return (idx === 0 ? "" : PAGE_BREAK) + mark(26, `${id}A`) + mid + mark(26, `${id}B`);
  })
  .join("");

write("probe-quarterline-shape.docx", shapeBody, `${shapes.length} shapes, one per page`);

/**
 * The probe that answers the question the fixture actually asks.
 *
 * Neither of the two probes above can be read by subtracting a control,
 * because Word's baseline-to-baseline advance turns out to depend on WHERE on
 * the page the line sits: at 13 pt the first advance is 15.000 pt and the rest
 * are 14.750, and at 26 pt the order reverses. So a per-paragraph "cost"
 * derived from a two-line control is not a quantity Word has.
 *
 * What decides whether page 1 of wild2-legal-ca-agreement fits is not a per
 * paragraph cost at all — it is the ABSOLUTE distance from the body top to the
 * last line. This probe measures exactly that, over a stack long enough for a
 * per-line difference to show as drift, with the quarter-line paragraphs placed
 * as the fixture places them:
 *
 *   E01  12 plain 13 pt lines                     (the drift control)
 *   E02  the same, one quarter-line empty after line 6
 *   E03  the same, two of them, after lines 4 and 8 — NOT adjacent, as the
 *        fixture has them
 *   E04  the same, two of them adjacent after line 6
 *
 * Read mark N's position against mark 1's in each render and compare the two
 * renders' accumulations. A per-line disagreement grows with N; a
 * per-construct one appears as a step at the construct and then stays flat.
 */
const stacks = [
  ["01", []],
  ["02", [6]],
  ["03", [4, 8]],
  ["04", [6, 6]],
];

const LINES = 12;
const quarter = `<w:p><w:pPr>${SPACING.replace('w:line="240"', 'w:line="60"')}${rpr(26)}</w:pPr></w:p>`;

const stackBody = stacks
  .map(([num, after], idx) => {
    let out = idx === 0 ? "" : PAGE_BREAK;
    for (let n = 1; n <= LINES; n++) {
      out += mark(26, `S13E${num}L${String(n).padStart(2, "0")}`);
      for (const at of after) if (at === n) out += quarter;
    }
    return out;
  })
  .join("");

write("probe-quarterline-stack.docx", stackBody, `${stacks.length} stacks of ${LINES} lines`);
