import { test, expect, Page, Locator } from "@playwright/test";

/**
 * Behavior tests for the editing/viewing interactions, driven with real
 * mouse/keyboard input. These exist because interaction regressions
 * (selection, grips, header gating) don't show up in the core unit tests.
 */

const DOC = "/?doc=/fixtures/sample.docx";
const MOD = process.platform === "darwin" ? "Meta" : "Control";

async function load(page: Page, url = DOC): Promise<void> {
  await page.goto(url);
  await page.waitForSelector(".dxw-page span");
  await page.waitForTimeout(300);
}

function span(page: Page, text: string): Locator {
  return page.locator(".dxw-page span", { hasText: text }).filter({ hasText: new RegExp(`^${text}$`) }).first();
}

async function clickText(page: Page, text: string, position: "start" | "end" = "end"): Promise<void> {
  const el = span(page, text);
  const box = (await el.boundingBox())!;
  await page.mouse.click(position === "end" ? box.x + box.width - 1 : box.x + 1, box.y + box.height / 2);
  await page.waitForTimeout(120);
}

async function caretVisible(page: Page): Promise<boolean> {
  return page.evaluate(() =>
    [...document.querySelectorAll("div")].some(
      (d) => d.style.width === "1.5px" && d.style.pointerEvents === "none" && d.style.display === "block",
    ),
  );
}

test.describe("rendering", () => {
  test("renders 4 pages with header and computed footer", async ({ page }) => {
    await load(page);
    await expect(page.locator(".dxw-page")).toHaveCount(4);
    await expect(span(page, "Fidelity")).toBeVisible();
    const pageText = (await page.locator(".dxw-page").first().innerText()).replace(/\s+/g, " ");
    expect(pageText).toContain("Page 1 of 4");
  });

  test("exposes semantic metadata without extra paint wrappers", async ({ page }) => {
    await load(page);
    const text = await span(page, "Fidelity").evaluate((el) => ({
      kind: (el as HTMLElement).dataset.dxwItemKind,
      family: (el as HTMLElement).dataset.dxwFontFamily,
      size: Number((el as HTMLElement).dataset.dxwFontSize),
      weight: (el as HTMLElement).dataset.dxwFontWeight,
      style: (el as HTMLElement).dataset.dxwFontStyle,
    }));
    expect(text.kind).toBe("text");
    expect(text.family).toBeTruthy();
    expect(text.size).toBeGreaterThan(0);
    expect(text.weight).toMatch(/^(400|700)$/);
    expect(text.style).toMatch(/^(normal|italic)$/);

    const fills = page.locator('[data-dxw-item-kind="rect"][data-dxw-role="table-fill"]');
    const rules = page.locator('[data-dxw-item-kind="edge"][data-dxw-role="table-rule"]');
    expect(await fills.count()).toBeGreaterThan(0);
    expect(await rules.count()).toBeGreaterThan(0);
    expect(
      await page.locator('[data-dxw-role="table-fill"]:not([data-dxw-item-kind="rect"])').count(),
    ).toBe(0);
    expect(
      await page.locator('[data-dxw-role="table-rule"]:not([data-dxw-item-kind="edge"])').count(),
    ).toBe(0);
  });

  test("view-only mode has no editing affordances", async ({ page }) => {
    await load(page);
    // Switch to Viewing via the mode dropdown (turns editing off).
    await page.locator("[data-dxw-mode]").click();
    await page.locator('[data-dxw-mode-option="viewing"]').click();
    await page.waitForTimeout(600);
    await expect(page.locator(".dxw-page")).toHaveCount(4);
    expect(await page.locator("[data-dxw-grip]").count()).toBe(0);
    expect(await page.locator("button", { hasText: "Table" }).count()).toBe(0);
    await page.locator(".dxw-page span").first().click();
    expect(await caretVisible(page)).toBe(false);
  });

  test("renders an embedded EMF as a PNG", async ({ page }) => {
    await load(page, "/?doc=/fixtures/wild2-sci-chem-omml.docx");
    const image = page.locator(".dxw-page").nth(9).locator("img");
    await expect.poll(() => image.evaluate((img) => (img as HTMLImageElement).naturalWidth)).toBeGreaterThan(200);

    const rendered = await image.evaluate((img) => {
      const source = img as HTMLImageElement;
      const canvas = document.createElement("canvas");
      canvas.width = 20;
      canvas.height = 20;
      const context = canvas.getContext("2d")!;
      context.drawImage(source, 0, 0, 20, 20);
      const pixels = context.getImageData(0, 0, 20, 20).data;
      const colors = new Set<string>();
      for (let i = 0; i < pixels.length; i += 4) {
        colors.add(`${pixels[i]},${pixels[i + 1]},${pixels[i + 2]}`);
      }
      return {
        src: source.src,
        width: source.naturalWidth,
        height: source.naturalHeight,
        colors: colors.size,
      };
    });
    expect(rendered.src).toMatch(/^data:image\/png;base64,/);
    expect(rendered.width).toBe(680);
    expect(rendered.height).toBe(665);
    expect(rendered.colors).toBeGreaterThan(20);
    await expect(image).toHaveAttribute("data-dxw-item-kind", "image");
    await expect(image).toHaveAttribute("data-dxw-image-format", "emf");
  });
});

test.describe("typing", () => {
  test("click, type, undo burst", async ({ page }) => {
    await load(page);
    await clickText(page, "document");
    await page.keyboard.type("XYZ", { delay: 40 });
    await expect(span(page, "documentXYZ")).toBeVisible();
    await page.keyboard.press(`${MOD}+z`);
    await page.waitForTimeout(300);
    expect(await span(page, "documentXYZ").count()).toBe(0);
  });

  test("Enter splits, Backspace at start merges", async ({ page }) => {
    await load(page);
    await clickText(page, "Justified");
    await page.keyboard.press("Enter");
    await page.waitForTimeout(300);
    const t1 = await span(page, "Justified").boundingBox();
    const t2 = await span(page, "paragraph").first().boundingBox();
    expect(t2!.y).toBeGreaterThan(t1!.y + 5);
    // caret is at start of "paragraph" — Backspace merges back
    await page.keyboard.press("Backspace");
    await page.waitForTimeout(300);
    const t3 = await span(page, "paragraph").first().boundingBox();
    const t4 = await span(page, "Justified").boundingBox();
    expect(Math.abs(t3!.y - t4!.y)).toBeLessThan(3);
  });

  test("click below text places caret at end of last line", async ({ page }) => {
    await load(page);
    const pageBox = (await page.locator(".dxw-page").first().boundingBox())!;
    // blank area below the content but inside the viewport
    await page.mouse.click(pageBox.x + 300, pageBox.y + 700);
    await page.waitForTimeout(150);
    expect(await caretVisible(page)).toBe(true);
  });
});

test.describe("selection", () => {
  test("drag-select paints owned highlight; formatting persists selection and composes", async ({ page }) => {
    await load(page);
    const a = (await span(page, "Lorem").boundingBox())!;
    const b = (await span(page, "consectetur").boundingBox())!;
    await page.mouse.move(a.x + 2, a.y + 5);
    await page.mouse.down();
    await page.mouse.move(b.x + b.width - 2, b.y + 5, { steps: 5 });
    await page.mouse.up();
    await page.waitForTimeout(200);
    // Rects are merged per visual line; the highlight must span the range.
    const rects = await page.locator(".dxw-sel").count();
    expect(rects).toBeGreaterThan(0);
    const selBox = (await page.locator(".dxw-sel").first().boundingBox())!;
    expect(selBox.x).toBeLessThanOrEqual(a.x + 2);
    expect(selBox.x + selBox.width).toBeGreaterThanOrEqual(b.x + b.width - 8);

    await page.locator("button[title*='Bold']").click();
    await page.waitForTimeout(400);
    expect(await page.locator(".dxw-sel").count()).toBeGreaterThan(0); // persisted
    const loremFont = await span(page, "Lorem").evaluate((el) => el.style.font);
    expect(loremFont).toContain("700");

    await page.locator("button[title='Italic']").click();
    await page.waitForTimeout(400);
    const loremFont2 = await span(page, "Lorem").evaluate((el) => el.style.font);
    expect(loremFont2).toContain("italic");

    // click elsewhere clears
    await clickText(page, "Bullet");
    expect(await page.locator(".dxw-sel").count()).toBe(0);
  });

  test("double-click selects a word", async ({ page }) => {
    await load(page);
    const el = span(page, "exercises");
    const box = (await el.boundingBox())!;
    await page.mouse.dblclick(box.x + box.width / 2, box.y + box.height / 2);
    await page.waitForTimeout(200);
    expect(await page.locator(".dxw-sel").count()).toBeGreaterThan(0);
  });

  test("copy puts selection text on the clipboard", async ({ page }) => {
    await load(page);
    const el = span(page, "exercises");
    const box = (await el.boundingBox())!;
    await page.mouse.dblclick(box.x + box.width / 2, box.y + box.height / 2);
    await page.waitForTimeout(200);
    await page.keyboard.press(`${MOD}+c`);
    await page.waitForTimeout(200);
    const clip = await page.evaluate(() => navigator.clipboard.readText());
    expect(clip).toContain("exercises");
  });
});

test.describe("tables", () => {
  test("column grip drag moves the boundary by the drag distance", async ({ page }) => {
    await load(page);
    const grips = page.locator("[data-dxw-grip]");
    let grip = grips.first();
    for (let i = 0; i < (await grips.count()); i++) {
      if ((await grips.nth(i).evaluate((el) => el.style.cursor)) === "col-resize") {
        grip = grips.nth(i);
        break;
      }
    }
    const gb = (await grip.boundingBox())!;
    const before = (await span(page, "Status").boundingBox())!.x;
    // Drag the Feature/Status boundary LEFT: autofit keeps these columns
    // near their content width, so growing into Status would starve it.
    await page.mouse.move(gb.x + 3, gb.y + 20);
    await page.mouse.down();
    await page.mouse.move(gb.x + 3 - 40, gb.y + 20, { steps: 4 });
    await page.mouse.up();
    await page.waitForTimeout(400);
    const after = (await span(page, "Status").boundingBox())!.x;
    expect(Math.abs(before - after - 40)).toBeLessThanOrEqual(3);
  });

  test("row grip drag grows the row", async ({ page }) => {
    await load(page);
    const rowGrip = page.locator("[data-dxw-grip]").filter({ has: page.locator(":scope") }).nth(0);
    // pick the first row-resize grip specifically
    const grips = page.locator("[data-dxw-grip]");
    const count = await grips.count();
    let target = null;
    for (let i = 0; i < count; i++) {
      const cursor = await grips.nth(i).evaluate((el) => el.style.cursor);
      if (cursor === "row-resize") {
        target = grips.nth(i);
        break;
      }
    }
    expect(target).not.toBeNull();
    const gb = (await target!.boundingBox())!;
    const before = (await span(page, "Pagination").boundingBox())!.y;
    await page.mouse.move(gb.x + 40, gb.y + 3);
    await page.mouse.down();
    await page.mouse.move(gb.x + 40, gb.y + 33, { steps: 4 });
    await page.mouse.up();
    await page.waitForTimeout(400);
    const after = (await span(page, "Pagination").boundingBox())!.y;
    expect(after - before).toBeGreaterThan(20);
    void rowGrip;
  });

  test("grid picker inserts a table at the caret", async ({ page }) => {
    await load(page);
    await clickText(page, "laborum.");
    const edgesBefore = await page.evaluate(
      () => document.querySelectorAll(".dxw-page [data-dxw-edge]").length,
    );
    await page.locator('button[data-tab="insert"]').click();
    await page.waitForTimeout(100);
    await page.locator("button[title='Table']").click();
    await page.locator("div", { hasText: /^Insert table$/ }).first().waitFor();
    // click the 2x3 cell (row 2, col 3) in the 10-col grid
    const cells = page.locator("div").filter({ hasText: /^$/ });
    // simpler: use the grid cells by size
    const gridCells = page.locator("div[style*='width: 16px'][style*='height: 16px']");
    await gridCells.nth(12).click();
    await page.waitForTimeout(400);
    const edgesAfter = await page.evaluate(
      () => document.querySelectorAll(".dxw-page [data-dxw-edge]").length,
    );
    expect(edgesAfter).toBeGreaterThan(edgesBefore);
    void cells;
  });
});

test.describe("headers and footers", () => {
  test("single click is gated; double-click enters with chrome; body double-click exits", async ({ page }) => {
    await load(page);
    const hdr = span(page, "Fidelity");
    const hb = (await hdr.boundingBox())!;
    await page.mouse.click(hb.x + 4, hb.y + 5);
    await page.waitForTimeout(150);
    expect(await caretVisible(page)).toBe(false);

    await page.mouse.dblclick(hb.x + 4, hb.y + 5);
    await page.waitForTimeout(250);
    expect(await caretVisible(page)).toBe(true);
    expect(await page.locator(".dxw-hf-marker").count()).toBeGreaterThan(0);
    const bodyOpacity = await span(page, "Lorem").evaluate((el) => getComputedStyle(el).opacity);
    expect(parseFloat(bodyOpacity)).toBeLessThan(1);

    // The dimmed body is inert: single clicks stay in header/footer mode.
    await clickText(page, "Lorem", "start");
    expect(await page.locator(".dxw-hf-marker").count()).toBeGreaterThan(0);

    // Double-click returns to body editing.
    const lb = (await span(page, "Lorem").boundingBox())!;
    await page.mouse.dblclick(lb.x + 1, lb.y + lb.height / 2);
    await page.waitForTimeout(250);
    expect(await page.locator(".dxw-hf-marker").count()).toBe(0);
  });

  test("body cannot be edited while header/footer mode is active", async ({ page }) => {
    await load(page);
    const hdr = span(page, "Fidelity");
    const hb = (await hdr.boundingBox())!;
    await page.mouse.dblclick(hb.x + 4, hb.y + 5);
    await page.waitForTimeout(250);

    // Drag-select over body text must not create a selection or edit it.
    const body = span(page, "Lorem");
    const bb = (await body.boundingBox())!;
    await page.mouse.move(bb.x + 2, bb.y + bb.height / 2);
    await page.mouse.down();
    await page.mouse.move(bb.x + 80, bb.y + bb.height / 2, { steps: 8 });
    await page.mouse.up();
    await page.waitForTimeout(150);
    expect(await page.locator(".dxw-sel").count()).toBe(0);
    await page.keyboard.type("ZAP");
    await page.waitForTimeout(300);
    expect(await span(page, "Lorem").count()).toBe(1);
    expect(await page.locator(".dxw-hf-marker").count()).toBeGreaterThan(0);
  });

  test("caret lands on the page whose header was double-clicked", async ({ page }) => {
    await load(page);
    const hdr2 = page.locator('.dxw-page[data-page="2"] span[data-dxw-hf]').first();
    await hdr2.scrollIntoViewIfNeeded();
    await page.waitForTimeout(200);
    const hb = (await hdr2.boundingBox())!;
    await page.mouse.dblclick(hb.x + 4, hb.y + hb.height / 2);
    await page.waitForTimeout(250);
    const caretPage = await page.evaluate(() => {
      const d = [...document.querySelectorAll("div")].find(
        (d) => d.style.width === "1.5px" && d.style.pointerEvents === "none" && d.style.display === "block",
      );
      return d?.closest(".dxw-page")?.getAttribute("data-page") ?? null;
    });
    expect(caretPage).toBe("2");
  });
});

test.describe("comments", () => {
  test("renders margin balloons with author/text, highlights ranges, hover links both ways", async ({ page }) => {
    await load(page);
    await expect(page.locator(".dxw-comment-card")).toHaveCount(2);
    const first = page.locator(".dxw-comment-card").first();
    await expect(first).toContainText("Ada Reviewer");
    await expect(first).toContainText("brand blue");
    expect(await page.locator("span[data-dxw-comment]").count()).toBeGreaterThan(0);

    // The balloon sits in the rail right of the page, near its anchor line.
    const pageBox = (await page.locator(".dxw-page").first().boundingBox())!;
    const cardBox = (await first.boundingBox())!;
    expect(cardBox.x).toBeGreaterThan(pageBox.x + pageBox.width);
    const anchor = (await page.locator("span[data-dxw-comment]").first().boundingBox())!;
    expect(Math.abs(cardBox.y - anchor.y)).toBeLessThan(60);

    // Hovering the commented text lights up its balloon.
    await page.locator("span[data-dxw-comment]").first().hover();
    await page.waitForTimeout(150);
    expect(await first.evaluate((el) => el.classList.contains("dxw-hot"))).toBe(true);
  });

  test("balloons stay anchored to the page when the window resizes", async ({ page }) => {
    await load(page);
    const gap = async () => {
      const pg = (await page.locator(".dxw-page").first().boundingBox())!;
      const cd = (await page.locator(".dxw-comment-card").first().boundingBox())!;
      return cd.x - (pg.x + pg.width);
    };
    const before = await gap();
    await page.setViewportSize({ width: 1150, height: 1000 });
    await page.waitForTimeout(300);
    expect(Math.abs((await gap()) - before)).toBeLessThan(2);
    await page.setViewportSize({ width: 1700, height: 1000 });
    await page.waitForTimeout(300);
    expect(Math.abs((await gap()) - before)).toBeLessThan(2);
  });

  test("replies nest in the parent balloon and round-trip through undo", async ({ page }) => {
    await load(page);
    const input = page.locator(".dxw-comment-card").first().locator(".dxw-comment-reply-input");
    await input.click();
    await input.fill("Agreed, let's do it.");
    await input.press("Enter");
    await page.waitForTimeout(600);
    await expect(page.locator(".dxw-comment-card")).toHaveCount(2); // reply nests, no new balloon
    await expect(page.locator(".dxw-comment-reply")).toHaveCount(1);
    await expect(page.locator(".dxw-comment-reply .dxw-comment-text")).toHaveText("Agreed, let's do it.");
    // Typing in the reply box must not edit the document body.
    expect(await page.locator(".dxw-page").first().innerText()).not.toContain("Agreed");

    await page.locator(".dxw-page span").first().click();
    await page.keyboard.press(`${MOD}+z`);
    await page.waitForTimeout(500);
    await expect(page.locator(".dxw-comment-reply")).toHaveCount(0);
  });

  test("comments are deletable from the balloon, with undo", async ({ page }) => {
    await load(page);
    await expect(page.locator(".dxw-comment-card")).toHaveCount(2);
    const first = page.locator(".dxw-comment-card").first();
    await first.hover();
    await first.locator(".dxw-comment-delete").click();
    await page.waitForTimeout(500);
    await expect(page.locator(".dxw-comment-card")).toHaveCount(1);
    expect(await page.locator(".dxw-comment-hl").count()).toBe(1);

    await page.locator(".dxw-page span").first().click();
    await page.keyboard.press(`${MOD}+z`);
    await page.waitForTimeout(500);
    await expect(page.locator(".dxw-comment-card")).toHaveCount(2);
    expect(await page.locator(".dxw-comment-hl").count()).toBe(2);
  });

  test("junk-grid table autofits columns to content like Word", async ({ page }) => {
    await load(page);
    const colX = async (t: string) => {
      const el = span(page, t);
      return (await el.boundingBox())!.x;
    };
    const f = await colX("Feature");
    const s = await colX("Status");
    const n = await colX("Notes");
    // An even three-way split would make each column ~208px; content
    // autofit keeps Feature/Status narrow and gives Notes the rest.
    expect(s - f).toBeLessThan(160);
    expect(n - s).toBeLessThan(120);
  });
});

test.describe("images in exact line spacing (pleading paper)", () => {
  const overlapCount = (page: Page) =>
    page.evaluate(() => {
      const img = document.querySelector(".dxw-page img");
      if (!img) return -1;
      const ir = img.getBoundingClientRect();
      let n = 0;
      for (const s of document.querySelectorAll(".dxw-page span:not([data-dxw-hf])")) {
        const r = s.getBoundingClientRect();
        const ox = Math.max(0, Math.min(ir.right, r.right) - Math.max(ir.left, r.left));
        const oy = Math.max(0, Math.min(ir.bottom, r.bottom) - Math.max(ir.top, r.top));
        if (ox > 4 && oy > 4) n++;
      }
      return n;
    });

  test("insert floats with square wrap; drag and resize keep text wrapping", async ({ page }) => {
    // Tall viewport: the cross-page drag at the end needs both pages visible
    // (mouse drags do not auto-scroll).
    await page.setViewportSize({ width: 1400, height: 2500 });
    await load(page, "/?doc=/fixtures/exact.docx");
    const line5 = span(page, "5:"); // the "5:" token of "Exact line 5: …"
    const lb = (await line5.boundingBox())!;
    await page.mouse.click(lb.x + 10, lb.y + lb.height / 2);
    await page.waitForTimeout(150);
    // 120x80 red png — taller than the fixed 32px line, so it must float.
    const png = Buffer.from(
      await page.evaluate(async () => {
        const c = document.createElement("canvas");
        c.width = 120;
        c.height = 80;
        c.getContext("2d")!.fillRect(0, 0, 120, 80);
        const blob: Blob = await new Promise((res) => c.toBlob((b) => res(b!), "image/png"));
        return Array.from(new Uint8Array(await blob.arrayBuffer()));
      }),
    );
    await page.locator("input[type=file][accept*='image']").setInputFiles({
      name: "t.png",
      mimeType: "image/png",
      buffer: png,
    });
    await page.waitForTimeout(700);
    expect(await overlapCount(page)).toBe(0);

    // Drag the floating image down into the middle of the text.
    const img = page.locator(".dxw-page img").first();
    const ib = (await img.boundingBox())!;
    await page.mouse.move(ib.x + ib.width / 2, ib.y + ib.height / 2);
    await page.mouse.down();
    await page.mouse.move(ib.x + ib.width / 2 + 100, ib.y + ib.height / 2 + 150, { steps: 10 });
    await page.mouse.up();
    await page.waitForTimeout(700);
    expect(await overlapCount(page)).toBe(0);

    // Resize larger via the corner handle — text must re-wrap around it.
    const b2 = (await page.locator(".dxw-page img").first().boundingBox())!;
    await page.mouse.click(b2.x + b2.width / 2, b2.y + b2.height / 2);
    await page.waitForTimeout(300);
    const corner = await page.evaluate(() => {
      const d = document.querySelector('[data-dxw-img-handle="se"]');
      if (!d) return null;
      const r = d.getBoundingClientRect();
      return { x: r.x + r.width / 2, y: r.y + r.height / 2 };
    });
    expect(corner).not.toBeNull();
    await page.mouse.move(corner!.x, corner!.y);
    await page.mouse.down();
    await page.mouse.move(corner!.x + 90, corner!.y + 60, { steps: 8 });
    await page.mouse.up();
    await page.waitForTimeout(700);
    const box = (await page.locator(".dxw-page img").first().boundingBox())!;
    expect(box.width).toBeGreaterThan(180);
    expect(await overlapCount(page)).toBe(0);

    // Drag the floating image onto page 2 — it must re-anchor there and
    // keep wrapping (floats follow the anchor paragraph across pages).
    const p2line = span(page, "33:");
    const pb = (await p2line.boundingBox())!;
    const cur = (await page.locator(".dxw-page img").first().boundingBox())!;
    await page.mouse.move(cur.x + cur.width / 2, cur.y + cur.height / 2);
    await page.mouse.down();
    await page.mouse.move(pb.x + 200, pb.y + 10, { steps: 12 });
    await page.mouse.up();
    await page.waitForTimeout(700);
    const onPage = await page.evaluate(
      () => document.querySelector(".dxw-page img")?.closest(".dxw-page")?.getAttribute("data-page") ?? null,
    );
    expect(onPage).toBe("2");
    expect(await overlapCount(page)).toBe(0);
  });
});

test.describe("images", () => {
  test("insert via toolbar, resize via corner handle, undo", async ({ page }) => {
    await load(page);
    await clickText(page, "laborum.");
    // 40x20 red png
    const png = Buffer.from(
      await page.evaluate(async () => {
        const c = document.createElement("canvas");
        c.width = 40;
        c.height = 20;
        c.getContext("2d")!.fillRect(0, 0, 40, 20);
        const blob: Blob = await new Promise((res) => c.toBlob((b) => res(b!), "image/png"));
        const buf = await blob.arrayBuffer();
        return Array.from(new Uint8Array(buf));
      }),
    );
    await page.locator("input[type=file][accept*='image']").setInputFiles({
      name: "t.png",
      mimeType: "image/png",
      buffer: png,
    });
    await page.waitForTimeout(700);
    const img = page.locator(".dxw-page img").first();
    await expect(img).toBeVisible();

    // select → corner handle appears → drag to resize
    await img.click();
    await page.waitForTimeout(200);
    const handle = page.locator('[data-dxw-img-handle="se"]');
    await expect(handle).toBeVisible();
    const hb = (await handle.boundingBox())!;
    await page.mouse.move(hb.x + 5, hb.y + 5);
    await page.mouse.down();
    await page.mouse.move(hb.x + 45, hb.y + 25, { steps: 4 });
    await page.mouse.up();
    await page.waitForTimeout(500);
    const w = await page.locator(".dxw-page img").first().evaluate((el) => parseFloat((el as HTMLElement).style.width));
    expect(w).toBeGreaterThan(60);

    await page.keyboard.press(`${MOD}+z`);
    await page.keyboard.press(`${MOD}+z`);
    await page.waitForTimeout(400);
    expect(await page.locator(".dxw-page img").count()).toBe(0);
  });

  test("Backspace deletes a selected image", async ({ page }) => {
    await load(page);
    await clickText(page, "laborum.");
    const png = Buffer.from(
      await page.evaluate(async () => {
        const c = document.createElement("canvas");
        c.width = 30;
        c.height = 30;
        c.getContext("2d")!.fillRect(0, 0, 30, 30);
        const blob: Blob = await new Promise((res) => c.toBlob((b) => res(b!), "image/png"));
        return Array.from(new Uint8Array(await blob.arrayBuffer()));
      }),
    );
    await page.locator("input[type=file][accept*='image']").setInputFiles({
      name: "d.png",
      mimeType: "image/png",
      buffer: png,
    });
    await page.waitForTimeout(700);
    await expect(page.locator(".dxw-page img")).toHaveCount(1);
    await page.locator(".dxw-page img").first().click();
    await page.waitForTimeout(200);
    await page.keyboard.press("Backspace");
    await page.waitForTimeout(400);
    expect(await page.locator(".dxw-page img").count()).toBe(0);
    // and it's undoable
    await page.keyboard.press(`${MOD}+z`);
    await page.waitForTimeout(400);
    expect(await page.locator(".dxw-page img").count()).toBe(1);
  });
});

test.describe("keyboard control", () => {
  test("Cmd+A selects all; Cmd+B bolds; arrows collapse selection", async ({ page }) => {
    await load(page);
    await clickText(page, "Lorem", "start");
    await page.keyboard.press(`${MOD}+a`);
    await page.waitForTimeout(250);
    const rects = await page.locator(".dxw-sel").count();
    expect(rects).toBeGreaterThan(20);
    await page.keyboard.press(`${MOD}+b`);
    await page.waitForTimeout(500);
    const font = await span(page, "Lorem").evaluate((el) => el.style.font);
    expect(font).toContain("700");
    // ArrowRight collapses to selection end and clears rects
    await page.keyboard.press("ArrowRight");
    await page.waitForTimeout(200);
    expect(await page.locator(".dxw-sel").count()).toBe(0);
    // undo the bold
    await page.keyboard.press(`${MOD}+z`);
    await page.waitForTimeout(300);
  });

  test("Shift+ArrowRight extends selection character by character", async ({ page }) => {
    await load(page);
    await clickText(page, "Lorem", "start");
    for (let i = 0; i < 5; i++) await page.keyboard.press("Shift+ArrowRight");
    await page.waitForTimeout(250);
    expect(await page.locator(".dxw-sel").count()).toBeGreaterThan(0);
    await page.keyboard.press(`${MOD}+c`);
    await page.waitForTimeout(200);
    const clip = await page.evaluate(() => navigator.clipboard.readText());
    expect(clip).toBe("Lorem");
  });

  test("Shift+click extends selection to the click point", async ({ page }) => {
    await load(page);
    await clickText(page, "Lorem", "start");
    const dest = span(page, "dolor");
    const db = (await dest.boundingBox())!;
    await page.keyboard.down("Shift");
    await page.mouse.click(db.x + db.width - 1, db.y + db.height / 2);
    await page.keyboard.up("Shift");
    await page.waitForTimeout(250);
    await page.keyboard.press(`${MOD}+c`);
    await page.waitForTimeout(200);
    const clip = await page.evaluate(() => navigator.clipboard.readText());
    expect(clip.replace(/\s+/g, " ")).toBe("Lorem ipsum dolor");
  });

  test("Home/End move to line edges", async ({ page }) => {
    await load(page);
    await clickText(page, "ipsum");
    await page.keyboard.press("Home");
    await page.waitForTimeout(150);
    await page.keyboard.press("Shift+End");
    await page.waitForTimeout(250);
    expect(await page.locator(".dxw-sel").count()).toBeGreaterThan(0);
    await page.keyboard.press(`${MOD}+c`);
    await page.waitForTimeout(200);
    const clip = await page.evaluate(() => navigator.clipboard.readText());
    expect(clip.startsWith("Lorem ipsum")).toBe(true);
  });
});

test.describe("styles", () => {
  test("style dropdown applies Heading 1 to the caret paragraph", async ({ page }) => {
    await load(page);
    await clickText(page, "ipsum");
    const px = (font: string) => parseFloat(/([\d.]+)px/.exec(font)?.[1] ?? "0");
    const sizeBefore = await span(page, "Lorem").evaluate((el) => el.style.font).then(px);
    const styleSel = page.locator("select[title='Paragraph style']");
    const value = await styleSel.evaluate((el) => {
      const opts = [...(el as HTMLSelectElement).options];
      return opts.find((o) => /heading\s*1/i.test(o.textContent ?? ""))?.value ?? "";
    });
    expect(value).not.toBe("");
    await styleSel.selectOption(value);
    await page.waitForTimeout(500);
    const sizeAfter = await span(page, "Lorem").evaluate((el) => el.style.font).then(px);
    expect(sizeAfter).toBeGreaterThan(sizeBefore);
    await page.keyboard.press(`${MOD}+z`);
  });

  test("style dropdown reflects the caret paragraph; heading hotkeys apply styles", async ({ page }) => {
    await load(page);
    const styleSel = page.locator("select[title='Paragraph style']");
    // Caret in the Heading 2 line → dropdown shows it.
    await clickText(page, "Justified", "start");
    expect(await styleSel.inputValue()).toBe("Heading2");
    // Caret in body text → Normal; Cmd/Ctrl+Alt+1 makes it Heading 1.
    await clickText(page, "Lorem", "start");
    expect(await styleSel.inputValue()).toBe("__normal");
    const px = (font: string) => parseFloat(/([\d.]+)px/.exec(font)?.[1] ?? "0");
    const before = await span(page, "Lorem").evaluate((el) => el.style.font).then(px);
    await page.keyboard.press(`${MOD}+Alt+1`);
    await page.waitForTimeout(500);
    expect(await styleSel.inputValue()).toBe("Heading1");
    const after = await span(page, "Lorem").evaluate((el) => el.style.font).then(px);
    expect(after).toBeGreaterThan(before);
    // Cmd/Ctrl+Alt+0 returns to Normal.
    await page.keyboard.press(`${MOD}+Alt+0`);
    await page.waitForTimeout(500);
    expect(await styleSel.inputValue()).toBe("__normal");
  });
});

test.describe("ime", () => {
  test("composition inserts the final text at the caret and skips raw keydowns", async ({ page }) => {
    await page.goto("/?doc=/fixtures/parity-text.docx");
    await page.waitForSelector(".dxw-page span");
    await page.waitForTimeout(300);
    await page.locator('.dxw-page span:text-is("Plain")').first().click();
    await page.waitForTimeout(200);
    // Simulate an IME session on the hidden textarea: composition events
    // with interleaved 229-style keydowns that must be ignored.
    await page.evaluate(() => {
      const ta = document.querySelector(".dxw-docx textarea, textarea")!;
      const fire = (type: string, data: string) =>
        ta.dispatchEvent(new CompositionEvent(type, { bubbles: true, data }));
      fire("compositionstart", "");
      fire("compositionupdate", "に");
      fire("compositionupdate", "にほ");
      fire("compositionupdate", "日本");
      fire("compositionend", "日本");
    });
    await page.waitForTimeout(400);
    const text = await page.evaluate(() =>
      [...document.querySelectorAll(".dxw-page span")].map((s) => s.textContent).join(""),
    );
    expect(text).toContain("日本");
    expect(text).not.toContain("にほ");
  });
});

test.describe("selection gating", () => {
  test("body drag selection never captures footer text", async ({ page }) => {
    await page.goto("/?doc=/fixtures/parity-headerfooter.docx");
    await page.waitForSelector(".dxw-page span");
    await page.waitForTimeout(300);
    const pg = (await page.locator(".dxw-page").first().boundingBox())!;
    // drag from mid-body down past the footer band
    await page.mouse.move(pg.x + pg.width / 2, pg.y + pg.height * 0.5);
    await page.mouse.down();
    await page.mouse.move(pg.x + pg.width / 2, pg.y + pg.height - 20, { steps: 8 });
    await page.mouse.up();
    await page.waitForTimeout(200);
    await page.keyboard.press("Backspace");
    await page.waitForTimeout(400);
    // footer text survives
    const text = await page.evaluate(() => [...document.querySelectorAll(".dxw-page span")].map((s) => s.textContent).join(" "));
    expect(text).toMatch(/Page\s+\d/);
  });

  test("clicking whitespace places a caret at the nearest body text", async ({ page }) => {
    await page.goto("/?doc=/fixtures/parity-text.docx");
    await page.waitForSelector(".dxw-page span");
    await page.waitForTimeout(300);
    const pg = (await page.locator(".dxw-page").first().boundingBox())!;
    // click deep in the blank lower half of the page
    await page.mouse.click(pg.x + pg.width / 2, pg.y + pg.height * 0.7);
    await page.waitForTimeout(250);
    const caretVisible = await page.evaluate(() => {
      const carets = [...document.querySelectorAll("div")].filter((d) => d.style.width === "1.5px" && d.style.display === "block");
      return carets.length;
    });
    expect(caretVisible).toBeGreaterThan(0);
    await page.keyboard.type("Q");
    await page.waitForTimeout(300);
    const text = await page.evaluate(() => [...document.querySelectorAll(".dxw-page span")].map((s) => s.textContent).join(""));
    expect(text).toContain("Q");
  });
});

test.describe("page break hotkey", () => {
  test("Cmd/Ctrl+Enter inserts a page break at the caret", async ({ page }) => {
    await page.goto("/?doc=/fixtures/parity-text.docx");
    await page.waitForSelector(".dxw-page span");
    await page.waitForTimeout(300);
    const before = await page.locator(".dxw-page").count();
    await page.locator('.dxw-page span:text-is("Plain")').first().click();
    await page.waitForTimeout(200);
    const mod = process.platform === "darwin" ? "Meta" : "Control";
    await page.keyboard.press(`${mod}+Enter`);
    await page.waitForTimeout(500);
    expect(await page.locator(".dxw-page").count()).toBeGreaterThan(before);
  });
});

test.describe("vertical caret navigation", () => {
  test("ArrowDown exits a trailing table row and crosses the page gap", async ({ page }) => {
    await page.goto("/?doc=/fixtures/sample.docx");
    await page.waitForSelector(".dxw-page span");
    await page.waitForTimeout(300);
    await page.locator('.dxw-page span:text-is("time")').first().click();
    await page.waitForTimeout(200);
    const caretPage = async () =>
      page.evaluate(() => {
        const c = [...document.querySelectorAll("div")].find((d) => d.style.width === "1.5px" && d.style.display === "block");
        return c ? [...document.querySelectorAll(".dxw-page")].indexOf(c.closest(".dxw-page")!) : -1;
      });
    const caretTop = async () =>
      page.evaluate(() => {
        const c = [...document.querySelectorAll("div")].find((d) => d.style.width === "1.5px" && d.style.display === "block");
        return c ? Math.round(parseFloat((c as HTMLElement).style.top)) : -1;
      });
    expect(await caretPage()).toBe(0);
    const startTop = await caretTop();
    // ArrowDown must leave the table row (there is now an empty paragraph
    // between the table and the page break, then page 2) - keep going until
    // the caret reaches page 2, proving it isn't stuck in the trailing row.
    let reachedP2 = false;
    for (let i = 0; i < 4; i++) {
      await page.keyboard.press("ArrowDown");
      await page.waitForTimeout(150);
      if ((await caretPage()) === 1) { reachedP2 = true; break; }
      // otherwise the caret must at least have moved down off the table row
      expect(await caretTop()).toBeGreaterThan(startTop);
    }
    expect(reachedP2).toBe(true);
    await page.keyboard.press("ArrowUp");
    await page.waitForTimeout(200);
    expect(await caretPage()).toBe(0);
  });
});

test.describe("backspace over page break", () => {
  test("Backspace at heading start after a page break bumps it up, keeping style", async ({ page }) => {
    await page.goto("/?doc=/fixtures/sample.docx");
    await page.waitForSelector(".dxw-page span");
    await page.waitForTimeout(300);
    const t = page.locator('.dxw-page span:text-is("Page")').nth(0);
    const box = (await t.boundingBox())!;
    await page.mouse.click(box.x + 1, box.y + box.height / 2);
    await page.waitForTimeout(200);
    await page.keyboard.press("Home");
    await page.keyboard.press("Backspace");
    await page.waitForTimeout(500);
    const onP1 = await page.evaluate(() => {
      const p1 = document.querySelectorAll(".dxw-page")[0];
      return [...p1.querySelectorAll("span")].some((s) => s.textContent === "Page" && parseFloat(s.style.top) > 130);
    });
    expect(onP1).toBe(true);
  });
});

test.describe("caret in empty paragraphs", () => {
  test("clicking below a table places a caret and types on the same page", async ({ page }) => {
    await page.goto("/?doc=/fixtures/sample.docx");
    await page.waitForSelector(".dxw-page span");
    await page.waitForTimeout(300);
    const cell = (await page.locator('.dxw-page span:text-is("time")').first().boundingBox())!;
    await page.mouse.click(cell.x, cell.y + 30);
    await page.waitForTimeout(200);
    await page.keyboard.type("BELOWTABLE");
    await page.waitForTimeout(400);
    const where = await page.evaluate(() => {
      const s = [...document.querySelectorAll(".dxw-page span")].find((x) => x.textContent!.includes("BELOWTABLE"));
      if (!s) return null;
      return [...document.querySelectorAll(".dxw-page")].indexOf(s.closest(".dxw-page")!);
    });
    expect(where).toBe(0); // typed on page 1, below the table (not lost)
  });
});





