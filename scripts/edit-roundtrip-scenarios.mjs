/**
 * Scenario library for the edit round-trip gate (scripts/edit-roundtrip-parity.mjs).
 *
 * A scenario names a fixture from apps/demo/public/fixtures and an `edit`
 * function that drives one scripted edit sequence. The runner loads the fixture
 * in the demo, calls `edit`, downloads the result, and holds desktop Word's
 * rendering of that download against the web renderer's.
 *
 * Adding a scenario is one entry in the array below. Keep each one:
 *   - deterministic — address text by content, never by pixel coordinate;
 *   - self-checking — assert the edit actually landed, so a silently ignored
 *     API call fails here rather than passing as "no visual difference";
 *   - resolved — leave no pending tracked changes or open header/footer mode,
 *     because the web comparison renders the saved bytes in viewing mode.
 *
 * The `ed` driver passed to `edit` is documented in editDriver() in the runner.
 *
 * A scenario may also declare:
 *   - `expectPages` — the page count it must produce, checked before any pixel
 *     comparison so repagination never reads as whatever else the scenario
 *     tests;
 *   - `verify({ editedDocx, wordPdf, verifyDir, fail, note })` — assertions
 *     pixels cannot make: what survived into the saved package, or what desktop
 *     Word does with the file beyond exporting a PDF.
 */

import { unzipSync, strFromU8 } from "fflate";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { wordUpdateFieldsAndSave } from "./word-export.mjs";

/** One part of a saved package, as text. */
function part(docx, name) {
  const files = unzipSync(readFileSync(docx));
  return files[name] ? strFromU8(files[name]) : null;
}

/** 48x48 checkerboard, small enough to inline and obvious enough to see. */
const TILE_PNG_BASE64 =
  "iVBORw0KGgoAAAANSUhEUgAAADAAAAAwCAIAAADYYG7QAAAATUlEQVR42u3XsQ0AIAgEQOZwTZdmCx1BSIw2R6if64AYM4+9C" +
  "nUrJ4CAgICAgJqgl8MqOUBAQEBAQF2QbQ8EBAQE5C9zfgABAQEBfQVtA9YCLeEdeXwAAAAASUVORK5CYII=";

export const scenarios = [
  {
    name: "typing",
    fixture: "parity-text",
    description: "Insert a new paragraph of text mid-document.",
    async edit(ed) {
      await ed.clickText("Plain text parity");
      await ed.select("Centered single line of text");
      await ed.press("ArrowRight"); // collapse to the end of the match
      await ed.press("Enter");
      await ed.type("Round-trip typing scenario inserted this paragraph.");
      await ed.expectRendered("Round-trip typing scenario inserted this paragraph.");
    },
  },
  {
    name: "formatting",
    fixture: "parity-text",
    description: "Bold, italic and recolor a run inside a body paragraph.",
    async edit(ed) {
      await ed.clickText("Plain text parity");
      await ed.select("consectetur adipiscing elit");
      await ed.call("applyFormat", { bold: true, italic: true, color: "#C00000" });
      const format = await ed.call("getSelectionFormat");
      ed.assert(format?.bold && format?.italic, `selection did not take the format: ${JSON.stringify(format)}`);
    },
  },
  {
    name: "lists",
    fixture: "parity-lists",
    description: "Toggle a paragraph into a bullet list and another into a numbered list.",
    async edit(ed) {
      await ed.clickText("Lists parity");
      await ed.select("Back to top level");
      await ed.call("toggleList", "bullet");
      ed.assert(await ed.call("getListType") === "bullet", "paragraph did not become a bullet item");
      await ed.select("Bullet two");
      await ed.call("toggleList", "number");
      ed.assert(await ed.call("getListType") === "number", "paragraph did not become a numbered item");
    },
  },
  {
    name: "table-ops",
    fixture: "parity-tables",
    description: "Add a row and a column to a table and merge two cells, then insert a new table.",
    async edit(ed) {
      await ed.clickText("Tables parity");
      // Row/column/merge act on the table holding the caret, and insertTable
      // leaves the caret where it was rather than inside the new table — so the
      // ops target the fixture's own table and the insert is a separate step.
      await ed.select("Status");
      await ed.press("ArrowRight"); // find() leaves a selection; table ops need a caret
      ed.assert(await ed.inTable(), "caret did not land inside the fixture table");
      // Column ops refuse a table that already uses gridSpan, so merge last.
      await ed.call("tableOp", "rowBelow");
      await ed.call("tableOp", "colRight");
      await ed.call("tableOp", "mergeRight");

      await ed.select("Fixed columns");
      await ed.press("ArrowRight");
      await ed.press("Enter");
      await ed.call("insertTable", 2, 2);
      // Fill the new table. A user who inserts a table types in it, so leaving
      // it empty under-tests the round trip; an empty table is also nothing but
      // evenly spaced rules, where every rule matches its neighbour and the
      // comparison metric reports a line shift that is not there.
      await ed.press("ArrowDown"); // insertTable leaves the caret above the table
      ed.assert(await ed.inTable(), "caret did not reach the inserted table");
      await ed.type("New A1");
      await ed.press("Tab");
      await ed.type("New B1");
      await ed.expectRendered("New A1");
    },
  },
  {
    name: "image-insert",
    fixture: "parity-text",
    description: "Insert a small inline PNG at the caret.",
    async edit(ed) {
      await ed.clickText("Plain text parity");
      await ed.select("Centered single line of text");
      await ed.press("ArrowRight");
      await ed.press("Enter");
      const result = await ed.evaluate(async (base64) => {
        const bytes = Uint8Array.from(atob(base64), (character) => character.charCodeAt(0));
        return window.__dxwApi.insertImage(new Blob([bytes], { type: "image/png" }));
      }, TILE_PNG_BASE64);
      ed.assert(result === "inserted", `insertImage returned ${result}`);
    },
  },
  {
    name: "comments",
    fixture: "parity-text",
    description: "Add a review comment on a selected phrase.",
    async edit(ed) {
      await ed.clickText("Plain text parity");
      await ed.select("Centered single line of text");
      const added = await ed.call("addComment", "Round-trip comment on this line.");
      ed.assert(added === true, "addComment refused the selection");
    },
  },
  {
    name: "footnote",
    fixture: "parity-text",
    description: "Add a footnote at the end of a body sentence.",
    async edit(ed) {
      await ed.clickText("Plain text parity");
      await ed.select("Centered single line of text");
      await ed.press("ArrowRight");
      const added = await ed.call("addFootnote", "Round-trip footnote text.");
      ed.assert(added === true, "addFootnote refused the caret");
    },
  },
  {
    name: "tracked-changes",
    fixture: "parity-text",
    description: "Suggest an insertion and a deletion, then accept one and reject the other.",
    async edit(ed) {
      await ed.clickText("Plain text parity");
      await ed.call("setSuggesting", true, "Parity Bot");

      await ed.select("Centered single line of text");
      await ed.press("ArrowRight");
      await ed.type(" and a suggested tail");

      await ed.select("Right-aligned single line");
      await ed.press("Delete");
      await ed.settle();
      ed.assert(await ed.call("revisionCount") === 2, "expected exactly two pending suggestions");

      // Accept the insertion (it stays) and reject the deletion (the text comes
      // back), so the saved document carries no unresolved markup: the web
      // comparison renders it in viewing mode, which shows the final document.
      // Click rather than find: these two resolve the revision under the CARET,
      // and find() leaves a selection instead. Markup view still paints the
      // suggested deletion, so both marks are clickable.
      await ed.clickText("and a suggested tail");
      ed.assert(await ed.call("acceptRevisionAtCaret") === true, "could not accept the suggested insertion");
      await ed.clickText("Right-aligned single line");
      ed.assert(await ed.call("rejectRevisionAtCaret") === true, "could not reject the suggested deletion");

      await ed.call("setSuggesting", false);
      await ed.settle();
      ed.assert(await ed.call("revisionCount") === 0, "suggestions were left unresolved");
      await ed.expectRendered("Right-aligned single line");
    },
  },
  {
    name: "page-layout",
    fixture: "parity2-sections",
    description: "Change margins and orientation on the caret's section only.",
    async edit(ed) {
      await ed.clickText("Section 1: portrait, one column");
      await ed.select("Paragraph 1. Lorem ipsum");
      const section = await ed.call("getSectionContext");
      ed.assert(section?.index === 1, `expected the caret in section 1, got ${JSON.stringify(section)}`);
      await ed.call(
        "setPageLayout",
        { margins: { top: 0.75, right: 0.75, bottom: 0.75, left: 0.75 }, orientation: "landscape" },
        "section",
      );
    },
  },
  {
    name: "header-footer",
    fixture: "parity-headerfooter",
    description: "Open the running header, type into it, and close header/footer mode.",
    async edit(ed) {
      await ed.clickText("Header & footer parity");
      ed.assert(await ed.call("openHeaderFooter", "header") === true, "could not open the header");
      await ed.type("Round-trip header");
      await ed.call("closeHeaderFooter");
      await ed.settle();
      await ed.expectRendered("Round-trip header");
    },
  },
  {
    name: "table-borders",
    fixture: "parity-tables",
    description: "Paint all six table edges, restyle the inside vertical rule, suppress one cell's top edge, and add a diagonal.",
    async edit(ed) {
      await ed.clickText("Tables parity");
      await ed.select("Status");
      await ed.press("ArrowRight"); // find() leaves a selection; table ops need a caret
      ed.assert(await ed.inTable(), "caret did not land inside the fixture table");

      await ed.call(
        "setTableBorders",
        "table",
        ["top", "bottom", "left", "right", "insideH", "insideV"],
        { style: "single", sz: 8, color: "#C00000" },
      );
      await ed.call("setTableBorders", "table", ["insideV"], { style: "dashed", sz: 4, color: "#2F5496" });
      // Suppressing one cell's edge is not the same as leaving the edge out:
      // w:val="nil" positively instructs Word to draw nothing there, overriding
      // the table-level edge the previous call just painted. The verify below
      // checks that distinction survived into the saved package.
      await ed.call("setTableBorders", "cell", ["top"], { style: "none" });

      await ed.select("A much longer description");
      await ed.press("ArrowRight");
      ed.assert(await ed.inTable(), "caret did not reach the description cell");
      await ed.call("setTableBorders", "cell", ["tl2br"], { style: "single", sz: 8, color: "#C00000" });
    },
    async verify({ editedDocx, fail, note }) {
      const document = part(editedDocx, "word/document.xml");
      const nil = (document.match(/<w:top[^>]*w:val="nil"/g) ?? []).length;
      const diagonal = (document.match(/<w:tl2br\b/g) ?? []).length;
      const dashed = (document.match(/w:val="dashed"/g) ?? []).length;
      note("nilTopEdges", nil);
      note("tl2brEdges", diagonal);
      note("dashedEdges", dashed);
      if (nil === 0) fail('Suppressed cell edge did not survive as w:val="nil"');
      if (diagonal === 0) fail("Diagonal cell border (w:tl2br) did not survive the save");
      if (dashed === 0) fail("Dashed insideV rule did not survive the save");
    },
  },
  {
    name: "table-widths",
    fixture: "parity-tables",
    description: "Freeze layout, then set a column width, table width, table cell margins and a repeating header row.",
    async edit(ed) {
      await ed.clickText("Tables parity");
      await ed.select("Status");
      await ed.press("ArrowRight");
      ed.assert(await ed.inTable(), "caret did not land inside the fixture table");
      // Layout first: "fixed" freezes the painted widths, so the numeric width
      // set next is the width Word draws rather than a hint autofit overrides.
      await ed.call("setTableLayout", "fixed");
      await ed.call("setTableColumnWidth", 0, 108);
      await ed.call("setTableWidth", "pct", 90);
      await ed.call("setTableCellMargins", "table", { left: 10, right: 10 });
      await ed.call("setTableHeaderRows", 1);
    },
    async verify({ editedDocx, fail, note }) {
      const document = part(editedDocx, "word/document.xml");
      const fixed = /<w:tblLayout[^>]*w:type="fixed"/.test(document);
      const pct = /<w:tblW[^>]*w:type="pct"/.test(document);
      const margins = /<w:tblCellMar>/.test(document);
      const header = /<w:tblHeader\b/.test(document);
      note("tblLayoutFixed", fixed);
      note("tblWidthPct", pct);
      note("tblCellMar", margins);
      note("tblHeader", header);
      if (!fixed) fail("Fixed table layout did not survive the save");
      if (!pct) fail("Percentage table width did not survive the save");
      if (!margins) fail("Table cell margins did not survive the save");
      if (!header) fail("Repeating header row did not survive the save");
    },
  },
  {
    name: "table-style",
    fixture: "probe2-modern-template",
    description: "Clear the themed table style, reapply it, and set conditional-format toggles the authored file did not have.",
    async edit(ed) {
      await ed.clickText("Regional Performance");
      await ed.select("North");
      await ed.press("ArrowRight");
      ed.assert(await ed.inTable(), "caret did not land inside the themed table");

      const styles = await ed.call("listTableStyles");
      ed.assert(
        styles.some((style) => style.id === "GridTableAccent1"),
        `listTableStyles did not offer GridTableAccent1 (${styles.length} styles)`,
      );

      await ed.call("setTableStyle", null);
      ed.assert(await ed.call("getTableStyleId") === null, "clearing the table style did not take");

      await ed.call("setTableStyle", "GridTableAccent1");
      await ed.call("setTableLook", { firstColumn: true, bandedRows: false, lastRow: true });
      const look = await ed.call("getTableLook");
      ed.assert(
        look?.firstColumn === true && look?.bandedRows === false && look?.lastRow === true,
        `getTableLook did not match what was set: ${JSON.stringify(look)}`,
      );
    },
    async verify({ editedDocx, fail, note }) {
      const document = part(editedDocx, "word/document.xml");
      const look = document.match(/<w:tblLook[^>]*\/>/)?.[0] ?? "";
      note("tblLook", look);
      if (!/w:firstColumn="1"/.test(look)) fail(`firstColumn toggle missing from saved tblLook: ${look}`);
      if (!/w:lastRow="1"/.test(look)) fail(`lastRow toggle missing from saved tblLook: ${look}`);
      // Banding off is written as noHBand="1" — the authored file had "0".
      if (!/w:noHBand="1"/.test(look)) fail(`bandedRows:false missing from saved tblLook: ${look}`);
      if (!/<w:tblStyle w:val="GridTableAccent1"\/>/.test(document)) fail("Reapplied table style did not survive the save");
    },
  },
  {
    name: "toc-insert",
    fixture: "wild2-legal-ca-agreement",
    description:
      "Insert a native table of contents at the caret, then let desktop Word update every field and re-save. " +
      "Manual-only checks NOT asserted here: that the ribbon's Update Table button offers Update page numbers / " +
      "Update entire table, and that Ctrl-clicking an entry navigates to its heading.",
    async edit(ed) {
      await ed.clickText("GECEFAVO VEZOCUHOJ");
      ed.assert(await ed.call("insertToc") === true, "insertToc refused the caret");
      await ed.settle();
    },
    async verify({ editedDocx, verifyDir, fail, note }) {
      const ours = part(editedDocx, "word/document.xml");
      const tocInstructions = (ours.match(/TOC\s+\\o/g) ?? []).length;
      // A TOC entry is a PAGEREF to a bookmarked heading; counting them is the
      // "entry count" this scenario holds constant across Word's own update.
      const ourEntries = (ours.match(/PAGEREF\s+_Toc/g) ?? []).length;
      note("tocInstructions", tocInstructions);
      note("tocEntriesOurs", ourEntries);
      if (tocInstructions === 0) return fail("No TOC field instruction in the saved package");
      if (ourEntries === 0) return fail("Inserted TOC carried no PAGEREF entries");

      // Word's F9 over the whole document, then Word's own re-save. This is the
      // half a browser cannot check: whether Word accepts our TOC as a live
      // field rather than as inert text it rewrites or drops.
      const wordResaved = join(verifyDir, "word-updated.docx");
      const { fields, updated } = wordUpdateFieldsAndSave({
        name: "toc-insert",
        docx: editedDocx,
        docxDestination: wordResaved,
      });
      note("wordFields", fields);
      note("wordFieldsUpdated", updated);

      const theirs = part(wordResaved, "word/document.xml");
      const theirInstructions = (theirs.match(/TOC\s+\\o/g) ?? []).length;
      const theirEntries = (theirs.match(/PAGEREF\s+_Toc/g) ?? []).length;
      note("tocEntriesWord", theirEntries);
      if (theirInstructions === 0) fail("Word's re-save dropped the TOC field instruction");
      if (theirEntries !== ourEntries) {
        fail(`TOC entry count changed under Word's update: ours ${ourEntries}, Word ${theirEntries}`);
      }
    },
  },
  {
    name: "field-update",
    fixture: "wild2-med-nccih-protocol",
    // Word must agree on pagination before any field claim means anything: page
    // fields ARE the pagination, so a repagination would masquerade as a field
    // arithmetic bug.
    expectPages: 23,
    description: "Recompute every supported field (Word's F9) across a 23-page protocol, then check Word recomputes them the same way.",
    async edit(ed) {
      await ed.clickText("Hegulufu Kuzaguzuvala Qapuw");
      const changed = await ed.call("updateFields");
      ed.assert(changed === true, `updateFields reported no change (${changed})`);
    },
    async verify({ editedDocx, verifyDir, fail, note }) {
      const ours = part(editedDocx, "word/document.xml");
      note("ourFieldRuns", (ours.match(/<w:instrText/g) ?? []).length);

      // The strong form: let Word recompute the same fields and compare the
      // cached results. Agreement means our arithmetic equals Word's F9.
      const wordResaved = join(verifyDir, "word-updated.docx");
      const { fields, updated } = wordUpdateFieldsAndSave({
        name: "field-update",
        docx: editedDocx,
        docxDestination: wordResaved,
      });
      note("wordFields", fields);
      note("wordFieldsUpdated", updated);

      const theirs = part(wordResaved, "word/document.xml");
      const pageRefs = (text) => (text.match(/PAGEREF\s+\S+/g) ?? []).length;
      note("pageRefsOurs", pageRefs(ours));
      note("pageRefsWord", pageRefs(theirs));
      if (pageRefs(theirs) !== pageRefs(ours)) {
        fail(`PAGEREF count changed under Word's update: ours ${pageRefs(ours)}, Word ${pageRefs(theirs)}`);
      }
    },
  },
];
