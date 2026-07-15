import { describe, it, expect } from "vitest";
import { readFileSync, existsSync } from "node:fs";
import { join } from "node:path";
import { DocxDocument } from "../wordinweb/packages/core/src/docx.js";
import { layoutDocument, layoutDocumentAsync, relayoutHeadersFooters, __incrStats } from "../wordinweb/packages/core/src/layout/engine.js";
import { createMeasurer } from "../wordinweb/packages/core/src/layout/measure.js";
import { XmlElement } from "../wordinweb/packages/core/src/xml.js";
import { LaidOutPage } from "../wordinweb/packages/core/src/layout/types.js";
import { topLevelBlockOf } from "../wordinweb/packages/core/src/edit/blocks.js";

// Full-vs-incremental equivalence harness (task #60). For each fixture, lay the
// document out fully (the "previous" result), apply an edit at several
// positions, then run BOTH a full layout and an incremental layout of the
// edited document. The incremental path MUST produce output identical to the
// full path for every page — ignoring only the editor-only `src` back-reference,
// which points at freshly-created model objects each layout and never affects
// rendering (renderToDom ignores it too).

const FIX_DIR = join(__dirname, "../apps/demo/public/fixtures");

/** True for a parsed XML element (stable across refresh — compared by identity,
 * like renderToDom's item diff, since the element tree is mutated in place). */
function isXmlish(v: unknown): boolean {
  return typeof v === "object" && v !== null && "attrs" in (v as object) && "children" in (v as object) && "name" in (v as object);
}

/** Deep equality ignoring the `src` key; XML-element fields by identity. */
function eq(a: unknown, b: unknown, depth = 0): boolean {
  if (a === b) return true;
  if (depth > 20) return false;
  if (typeof a !== "object" || typeof b !== "object" || a === null || b === null) return false;
  if (isXmlish(a) || isXmlish(b)) return a === b;
  const aArr = Array.isArray(a);
  const bArr = Array.isArray(b);
  if (aArr || bArr) {
    if (!aArr || !bArr || a.length !== b.length) return false;
    for (let i = 0; i < a.length; i++) if (!eq(a[i], b[i], depth + 1)) return false;
    return true;
  }
  const ao = a as Record<string, unknown>;
  const bo = b as Record<string, unknown>;
  for (const k in ao) {
    if (k === "src") continue;
    if (!(k in bo) || !eq(ao[k], bo[k], depth + 1)) return false;
  }
  for (const k in bo) {
    if (k === "src") continue;
    if (!(k in ao)) return false;
  }
  return true;
}

/** First page-level field or item index where two page lists differ, or null. */
function firstPageDiff(A: LaidOutPage[], B: LaidOutPage[]): string | null {
  if (A.length !== B.length) return `page count ${A.length} vs ${B.length}`;
  for (let i = 0; i < A.length; i++) {
    const a = A[i];
    const b = B[i];
    for (const f of ["width", "height", "index", "number", "bodyTop", "bodyBottom", "hfStart"] as const) {
      if (a[f] !== b[f]) return `page ${i} field ${f}: ${a[f]} vs ${b[f]}`;
    }
    if (a.items.length !== b.items.length) return `page ${i} item count ${a.items.length} vs ${b.items.length}`;
    for (let j = 0; j < a.items.length; j++) {
      if (!eq(a.items[j], b.items[j])) {
        return `page ${i} item ${j} kind=${(a.items[j] as { kind?: string }).kind}`;
      }
    }
  }
  return null;
}

/** Collect w:t elements with non-empty text, in document order. */
function textNodes(root: XmlElement, out: XmlElement[] = []): XmlElement[] {
  if (root.name.endsWith(":t") || root.name === "t") {
    if (root.text && root.text.trim().length > 0) out.push(root);
  }
  for (const c of root.children) textNodes(c, out);
  return out;
}

/** The top-level block element (block.src) whose subtree contains `t` — the
 * dirty-block hint the editor would pass for an edit at `t`. */
function blockSrcContaining(doc: DocxDocument, t: XmlElement): XmlElement | undefined {
  return topLevelBlockOf(doc, t) ?? undefined;
}

const FIXTURES = [
  "sample",
  "dense-skewtest",
  "wild2-med-nccih-protocol",
  "wild2-med-phase23-protocol",
  "wild2-math-omml-dense",
  "wild2-legal-nih-contract",
];

describe("full-vs-incremental layout equivalence", () => {
  for (const name of FIXTURES) {
    const path = join(FIX_DIR, `${name}.docx`);
    const present = existsSync(path);
    it.skipIf(!present)(`${name}: incremental == full across edit positions`, { timeout: 120000 }, () => {
      const bytes = new Uint8Array(readFileSync(path));
      const measurer = createMeasurer();
      const doc = DocxDocument.load(bytes);
      let prev = layoutDocument(doc, { measurer });

      const nodes = textNodes(doc.docRoot);
      expect(nodes.length).toBeGreaterThan(0);
      // Sample edit positions spread across the document.
      const positions = [0, 0.25, 0.5, 0.75, 0.98].map((f) => Math.min(nodes.length - 1, Math.floor(f * nodes.length)));

      // Count how often the correct hint actually engaged the fast path, so a
      // silently-dead optimisation (e.g. hint never matching a block) fails.
      // `incrEligible` tracks whether the incremental scan runs at all for this
      // fixture (some — multi-section, footnotes, columns — always fall back to
      // a full layout, so the fast path can never fire and is not required).
      let fastPathFirings = 0;
      let incrEligible = false;

      for (const pos of positions) {
        const t = nodes[pos];
        const original = t.text;
        // Simulate a keystroke: insert a character mid-text.
        const at = Math.floor(original.length / 2);
        const correctHint = blockSrcContaining(doc, t);
        // A different top-level block: a stale/wrong hint that must fall back.
        const wrongHint = doc.sections[0].blocks
          .map((b) => b.src)
          .find((s): s is XmlElement => !!s && s !== correctHint);

        t.text = original.slice(0, at) + "X" + original.slice(at);
        doc.refresh();

        const full = layoutDocument(doc, { measurer });

        // The dirty hint is purely an optimisation: absent, correct, and
        // wrong/stale hints must all produce output identical to a full layout.
        const modes: ReadonlyArray<readonly ["absent" | "correct" | "wrong", XmlElement | undefined]> = [
          ["absent", undefined],
          ["correct", correctHint],
          ["wrong", wrongHint],
        ];
        for (const [mode, hint] of modes) {
          __incrStats.hintFastPath = false;
          __incrStats.blocksHashed = -1;
          const incr = layoutDocument(doc, { measurer, prev, dirtyHint: hint });
          const diff = firstPageDiff(full.pages, incr.pages);
          expect(diff, `${name} pos ${pos} hint=${mode}: ${diff}`).toBeNull();

          // blocksHashed leaves its -1 reset only when the incremental scan
          // actually ran (i.e. the fixture is in the incremental envelope).
          if (__incrStats.blocksHashed >= 0) incrEligible = true;
          if (__incrStats.hintFastPath) {
            // When the fast path engages it hashes only the hinted block plus
            // its (up to two) neighbours — never the whole document.
            expect(__incrStats.blocksHashed, `${name} pos ${pos} hint=${mode}`).toBeLessThanOrEqual(3);
          }
          if (mode === "correct" && __incrStats.hintFastPath) fastPathFirings++;
          if (mode === "wrong") {
            // An unchanged (wrong) block can never satisfy the "signature
            // changed" guard, so it always falls back to the full scan.
            expect(__incrStats.hintFastPath, `${name} pos ${pos}: wrong hint must fall back`).toBe(false);
          }
        }

        // Restore for the next isolated position.
        t.text = original;
        doc.refresh();
        prev = layoutDocument(doc, { measurer });
      }

      // For fixtures the incremental engine handles, the fast path must engage
      // for at least one sampled edit; otherwise the hint is wired up but never
      // actually saving work. (Fixtures that always fall back to a full layout
      // legitimately never fire it.)
      if (incrEligible) {
        expect(fastPathFirings, `${name}: dirty-hint fast path never fired`).toBeGreaterThan(0);
      }
    });
  }
});

describe("cooperative full layout", () => {
  const path = join(FIX_DIR, "wild2-legal-nih-contract.docx");
  it.skipIf(!existsSync(path))("matches synchronous book layout exactly", { timeout: 120000 }, async () => {
    const bytes = new Uint8Array(readFileSync(path));
    const doc = DocxDocument.load(bytes);
    const sync = layoutDocument(doc, { measurer: createMeasurer() });
    const cooperative = await layoutDocumentAsync(doc, { measurer: createMeasurer(), sliceMs: 4 });
    const diff = firstPageDiff(sync.pages, cooperative.pages);
    expect(diff, diff ?? undefined).toBeNull();
  });

  it("matches synchronous multi-section layout exactly", async () => {
    const bytes = new Uint8Array(readFileSync(join(FIX_DIR, "parity2-sections.docx")));
    const doc = DocxDocument.load(bytes);
    const sync = layoutDocument(doc, { measurer: createMeasurer() });
    const cooperative = await layoutDocumentAsync(doc, { measurer: createMeasurer(), sliceMs: 4 });
    const diff = firstPageDiff(sync.pages, cooperative.pages);
    expect(diff, diff ?? undefined).toBeNull();
  });

  it("honors cancellation before starting", async () => {
    const controller = new AbortController();
    controller.abort();
    const doc = DocxDocument.load(new Uint8Array(readFileSync(join(FIX_DIR, "sample.docx"))));
    await expect(layoutDocumentAsync(doc, { signal: controller.signal })).rejects.toMatchObject({ name: "AbortError" });
  });
});

describe("header/footer-only layout", () => {
  const path = join(FIX_DIR, "wild2-legal-nih-contract.docx");
  it.skipIf(!existsSync(path))("matches full layout and rejects geometry changes", { timeout: 120000 }, () => {
    const doc = DocxDocument.load(new Uint8Array(readFileSync(path)));
    const measurer = createMeasurer();
    const prev = layoutDocument(doc, { measurer });

    let source: XmlElement | undefined;
    let model: { text: string } | undefined;
    const parts = [...doc.headers.values(), ...doc.footers.values()];
    for (const part of parts) {
      for (const block of part.blocks) {
        if (block.type !== "paragraph") continue;
        for (const child of block.children) {
          const runs = child.type === "hyperlink" ? child.runs : [child];
          for (const run of runs) {
            for (const content of run.content) {
              if (content.kind === "text" && content.srcT && content.text.trim().length > 3) {
                source = content.srcT;
                model = content;
                break;
              }
            }
            if (source) break;
          }
          if (source) break;
        }
        if (source) break;
      }
      if (source) break;
    }
    expect(source).toBeDefined();
    expect(model).toBeDefined();

    source!.text += "x";
    model!.text = source!.text;
    const fast = relayoutHeadersFooters(doc, prev, measurer);
    expect(fast).not.toBeNull();
    const full = layoutDocument(doc, { measurer });
    const diff = firstPageDiff(full.pages, fast!.pages);
    expect(diff, diff ?? undefined).toBeNull();

    source!.text += " W".repeat(1000);
    model!.text = source!.text;
    expect(relayoutHeadersFooters(doc, fast!, measurer) === null).toBe(true);
  });
});

describe("editor dirty-block resolution", () => {
  const path = join(FIX_DIR, "wild2-legal-nih-contract.docx");
  it.skipIf(!existsSync(path))("maps table-cell text to its top-level table", () => {
    const doc = DocxDocument.load(new Uint8Array(readFileSync(path)));
    const table = doc.sections[0].blocks.find((block) =>
      block.type === "table" && block.rows.some((row) =>
        row.cells.some((cell) => cell.blocks.some((nested) => nested.type === "paragraph" && !!nested.src)),
      ),
    );
    expect(table?.type).toBe("table");
    if (!table || table.type !== "table") return;
    const paragraph = table.rows.flatMap((row) => row.cells)
      .flatMap((cell) => cell.blocks)
      .find((block) => block.type === "paragraph" && block.src);
    expect(paragraph?.src).toBeDefined();
    if (!paragraph?.src) return;
    const text = textNodes(paragraph.src)[0];
    expect(text).toBeDefined();
    expect(topLevelBlockOf(doc, text)).toBe(table.src);
  });
});
