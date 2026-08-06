#!/usr/bin/env node

/**
 * Generate the opening-paragraph sweep for #82.
 *
 * `engine.ts`'s `docGridDropBefore` block decides how far the FIRST paragraph
 * of a lines-grid section sits below the body top, and after #76 deleted the
 * four-row section reserve it has two surviving branches that no Word
 * measurement stands behind:
 *
 *   - an explicit `w:snapToGrid="0"` with NO spacing-before drops TWO grid
 *     rows. Its provenance is a synthetic test, and it is semantically odd:
 *     opting OUT of the grid should not push a paragraph DOWN, let alone by
 *     two rows of the grid it opted out of.
 *   - a `Heading1` opener starts one grid row below the body top plus 1.5pt,
 *     the "grid leading". Also unmeasured.
 *
 * Six sections, each starting `nextPage`, each on the same geometry with the
 * same `<w:docGrid w:type="lines" w:linePitch="312"/>`, differing only in the
 * OPENING paragraph. Every case then runs the same plain body paragraphs, so
 * the whole section's position is readable from any of them and a case that
 * opens low stays low:
 *
 *   P   plain paragraph, no spacing-before          (control: opens AT bodyTop)
 *   S   plain paragraph, w:before="240" (12pt)      (control: a before is kept)
 *   G   w:snapToGrid="0", no spacing-before         the two-row drop
 *   GS  w:snapToGrid="0", w:before="240"            the branch that keeps it
 *   H   Heading1, no spacing-before                 the row + 1.5pt carry-over
 *   HG  Heading1 + w:snapToGrid="0"                 which branch wins
 *
 * Read PER CASE as the TOP of the opening line against the section's body top
 * (96.00 CSS px on this geometry). P fixes what an ordinary opener does; every
 * other case is that number plus whatever its own branch adds. The predictions
 * under test are G - P = 2 x 20.8 = 41.6 px and H - P = 20.8 + 2.0 = 22.8 px.
 *
 * G and GS TOGETHER are what make the snapToGrid branch falsifiable: the code
 * drops two rows only when there is no spacing-before, so if Word puts G and GS
 * in the same place the drop is not conditional on the before, and if Word puts
 * BOTH at the body top there is no drop at all.
 *
 * Every paragraph is labelled `<case>L<nn>`, so a case reads the same way from
 * a Word PDF, from a raster, and from the browser.
 *
 *   node scripts/generate-gridopen-probe.mjs
 */

import { readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { strToU8, unzipSync, zipSync } from "fflate";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
// Same base package as the docGrid sweep, so the two probes' numbers compare
// directly: same styles, same fonts, same theme, same geometry.
const source = join(root, "apps/demo/public/fixtures/wild2-math-eq-as-images.docx");
const parts = unzipSync(new Uint8Array(readFileSync(source)));

/** A4 minus header and footer: the body is exactly 96.00 .. 1026.53 CSS px, so
 *  the opening line's top IS its distance below the body top. */
const pgProps =
  '<w:pgSz w:w="11906" w:h="16838"/>' +
  '<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/>' +
  '<w:cols w:space="720"/>';

/** The fixture's own grid, so this sweep and #62's read the same pitch: 312 tw
 *  = 20.8 CSS px, comfortably taller than the 16.32 px natural line. */
const GRID = '<w:docGrid w:type="lines" w:linePitch="312"/>';

const sectPr = () => `<w:sectPr><w:type w:val="nextPage"/>${pgProps}${GRID}</w:sectPr>`;

const cases = [
  { id: "P", style: "", snap: false, before: 0 },
  { id: "S", style: "", snap: false, before: 240 },
  { id: "G", style: "", snap: true, before: 0 },
  { id: "GS", style: "", snap: true, before: 240 },
  { id: "H", style: "Heading1", snap: false, before: 0 },
  { id: "HG", style: "Heading1", snap: true, before: 0 },
];

/** Enough lines that the section's first page is full either way, so a case
 *  that opens low is visible as a lost line at the foot as well as a moved top. */
const LINES = 50;

/** The opening paragraph carries the case's variable; every following one is
 *  plain, so any difference downstream is the opener's doing and nothing else.
 *  w:line="240" w:lineRule="auto" is single spacing - the grid, not the
 *  paragraph, decides the advance. */
const para = (label, { style = "", snap = false, before = 0 } = {}, tail = "") =>
  `<w:p><w:pPr>` +
  (style ? `<w:pStyle w:val="${style}"/>` : "") +
  (snap ? `<w:snapToGrid w:val="0"/>` : "") +
  `<w:spacing w:before="${before}" w:after="0" w:line="240" w:lineRule="auto"/>` +
  `<w:jc w:val="left"/>${tail}</w:pPr>` +
  `<w:r><w:t xml:space="preserve">${label} pahiwe nuqaji bepon rukefev kicuziri</w:t></w:r></w:p>`;

const body = cases
  .map((c, ci) => {
    const paras = [];
    for (let i = 1; i <= LINES; i++) {
      const label = `${c.id}L${String(i).padStart(2, "0")}`;
      // A section's properties ride the LAST paragraph of that section, so this
      // case's sectPr goes on this case's last paragraph. The final section has
      // no such paragraph and takes the body-level sectPr instead.
      const last = i === LINES && ci < cases.length - 1;
      // Only paragraph 1 carries the case's variable.
      paras.push(para(label, i === 1 ? c : {}, last ? sectPr() : ""));
    }
    return paras.join("");
  })
  .join("");

const documentXml =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  `<w:body>${body}<w:sectPr><w:type w:val="nextPage"/>${pgProps}${GRID}</w:sectPr></w:body>` +
  "</w:document>";

parts["word/document.xml"] = strToU8(documentXml);

// Probes live in fixtures-staging: parity-parallel.mjs adopts a reference only
// when a DOCX of the same name sits in apps/demo/public/fixtures, so keeping it
// out of there is what stops the full corpus run from picking the probe up.
// To read it in the browser: copy it in, measure, then remove it again.
const out = join(root, "fixtures-staging/probe-gridopen.docx");
writeFileSync(out, Buffer.from(zipSync(parts)));
console.log(`wrote ${out}`);
console.log(`cases: ${cases.map((c) => c.id).join(", ")}  (${LINES} lines each, pitch 312tw = 20.8px)`);
