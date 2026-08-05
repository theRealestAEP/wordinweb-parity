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
 */

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
      await ed.expectRendered("Tables parity");
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
];
