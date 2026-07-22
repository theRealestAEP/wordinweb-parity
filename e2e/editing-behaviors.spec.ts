import { expect, test } from "@playwright/test";
import { Editor } from "./editing.js";

/**
 * Declarative editing-behavior expectations, expressed with the Editor DSL.
 * Each test states a user action and the layout/caret outcome it must
 * produce. This is the growing "expected behaviors" suite — add a case here
 * whenever a bug is fixed so it can never silently regress.
 */

test.describe("editing behaviors", () => {
  test("clicking below a table places a caret and types on the same page", async ({ page }) => {
    const ed = await Editor.open(page, "sample");
    await ed.clickNear("time", 0, 30); // just below the last table row on page 1
    await ed.expectCaretVisible();
    await ed.type("BELOWTABLE");
    await ed.expectTextOnPage("BELOWTABLE", 0);
  });

  test("clicking whitespace far below text still gives a caret", async ({ page }) => {
    const ed = await Editor.open(page, "parity-text");
    await ed.clickPageFraction(0.5, 0.7);
    await ed.expectCaretVisible();
    await ed.type("WS");
    await ed.expectHasText("WS");
  });

  test("clicking below a bulleted list creates plain whitespace lines", async ({ page }) => {
    const ed = await Editor.open(page, "parity-text");
    const before = await page.locator(".dxw-page span").filter({ hasText: /^[•●]$/ }).count();
    await ed.clickText("Plain");
    await page.locator('[title="Bulleted list"], [data-tip="Bulleted list"]').first().click();
    await ed.clickPageFraction(0.5, 0.72);
    await ed.type("AFTERLIST");
    await ed.expectHasText("AFTERLIST");
    const bulletCount = await page.locator(".dxw-page span").filter({ hasText: /^[•●]$/ }).count();
    expect(bulletCount).toBe(before + 1);
  });

  test("Enter on an empty bullet exits the list", async ({ page }) => {
    const ed = await Editor.open(page, "parity-text");
    const before = await page.locator(".dxw-page span").filter({ hasText: /^[•●]$/ }).count();
    await ed.clickText("Plain");
    await page.locator('[title="Bulleted list"], [data-tip="Bulleted list"]').first().click();
    await ed.press("End");
    await ed.press("Enter");
    await ed.press("Enter");
    await ed.type("PLAINAFTERBULLET");
    await ed.expectHasText("PLAINAFTERBULLET");
    const bulletCount = await page.locator(".dxw-page span").filter({ hasText: /^[•●]$/ }).count();
    expect(bulletCount).toBe(before + 1);
  });

  test("Backspace at a heading after a page break bumps it onto the previous page", async ({ page }) => {
    const ed = await Editor.open(page, "sample");
    // "Page" heading is the first line on page 2.
    await ed.clickText("Page");
    await ed.press("Home");
    await ed.press("Backspace");
    await ed.expectTextOnPage("Page", 0);
  });

  test("Cmd/Ctrl+Enter inserts a page break and moves typing to the new page", async ({ page }) => {
    const ed = await Editor.open(page, "parity-text");
    const before = await ed.pageCount();
    await ed.clickText("Plain");
    const mod = process.platform === "darwin" ? "Meta" : "Control";
    await ed.press(`${mod}+Enter`);
    if ((await ed.pageCount()) <= before) throw new Error("page break did not add a page");
    await ed.type("AFTERBREAK");
    await ed.expectTextOnPage("AFTERBREAK", before);
  });

  test("typing enough Enters overflows to a new page", async ({ page }) => {
    const ed = await Editor.open(page, "parity-text");
    const before = await ed.pageCount();
    await ed.clickText("Plain");
    await ed.press("End");
    await ed.pressMany("Enter", 60);
    await ed.expectPageCountAtLeast(before + 1);
  });
});

test.describe("WordArt / watermark editing", () => {
  // Helpers: the CONFIDENTIAL watermark in parity2-watermark is a behind-text
  // VML WordArt repeated on every page.
  const inkBox = (page: import("@playwright/test").Page) =>
    page.evaluate(() => {
      const ink = document.querySelector("[data-dxw-wordart-ink]") as HTMLElement | null;
      if (!ink) return null;
      const r = ink.getBoundingClientRect();
      return { l: r.left, t: r.top, w: r.width, h: r.height };
    });
  const wmToolbarShown = (page: import("@playwright/test").Page) =>
    page.evaluate(() => !!document.querySelector("[data-dxw-wm-btn]"));

  test("clicking body text over a watermark places a caret, not a watermark selection", async ({ page }) => {
    await page.goto("/?doc=/fixtures/parity2-watermark.docx");
    await page.waitForSelector(".dxw-page span");
    await page.waitForTimeout(400);
    // Click directly on a glyph that overlaps the watermark ink.
    const glyph = page.locator('.dxw-page span:has-text("consectetur")').first();
    await glyph.click();
    await page.waitForTimeout(200);
    expect(await wmToolbarShown(page)).toBe(false); // caret, not watermark
  });

  test("clicking the watermark ink (off-glyph) selects it and shows the toolbar", async ({ page }) => {
    await page.goto("/?doc=/fixtures/parity2-watermark.docx");
    await page.waitForSelector(".dxw-page span");
    await page.waitForTimeout(400);
    const box = (await inkBox(page))!;
    expect(box).toBeTruthy();
    // Top-right corner of the ink box: over the art, unlikely to hit a glyph.
    await page.mouse.click(box.l + box.w - 14, box.t + 12);
    await page.waitForTimeout(200);
    expect(await wmToolbarShown(page)).toBe(true);
    await expect(page.locator('[data-dxw-wm-btn="Edit watermark text"]')).toBeVisible();
    await expect(page.locator('[data-dxw-wm-btn="Delete watermark"]')).toBeVisible();
  });

  test("editing the watermark text updates every instance", async ({ page }) => {
    await page.goto("/?doc=/fixtures/parity2-watermark.docx");
    await page.waitForSelector(".dxw-page span");
    await page.waitForTimeout(400);
    const box = (await inkBox(page))!;
    await page.mouse.click(box.l + box.w - 14, box.t + 12);
    await page.waitForTimeout(150);
    await page.locator('[data-dxw-wm-btn="Edit watermark text"]').click();
    await page.getByRole("textbox", { name: "Watermark text", exact: true }).fill("DRAFT COPY");
    await page.getByRole("button", { name: "Apply", exact: true }).click();
    await page.waitForTimeout(400);
    const texts = await page.evaluate(() =>
      [...document.querySelectorAll('[data-dxw-item-kind="wordart"]')].map((e) => e.textContent),
    );
    expect(texts.length).toBeGreaterThan(0);
    expect(texts.every((t) => t === "DRAFT COPY")).toBe(true);
  });

  test("opacity buttons change the watermark ink opacity", async ({ page }) => {
    await page.goto("/?doc=/fixtures/parity2-watermark.docx");
    await page.waitForSelector(".dxw-page span");
    await page.waitForTimeout(400);
    const box = (await inkBox(page))!;
    await page.mouse.click(box.l + box.w - 14, box.t + 12);
    await page.waitForTimeout(150);
    const opacityOf = () =>
      page.evaluate(() =>
        getComputedStyle(document.querySelector("[data-dxw-wordart-ink]") as HTMLElement).opacity,
      );
    const before = parseFloat(await opacityOf());
    await page.locator('[data-dxw-wm-btn="Less opaque"]').click();
    await page.waitForTimeout(300);
    const after = parseFloat(await opacityOf());
    expect(after).toBeLessThan(before);
  });

  test("deleting the watermark removes it from every page", async ({ page }) => {
    await page.goto("/?doc=/fixtures/parity2-watermark.docx");
    await page.waitForSelector(".dxw-page span");
    await page.waitForTimeout(400);
    const box = (await inkBox(page))!;
    await page.mouse.click(box.l + box.w - 14, box.t + 12);
    await page.waitForTimeout(150);
    await page.locator('[data-dxw-wm-btn="Delete watermark"]').click();
    await page.waitForTimeout(400);
    const count = await page.evaluate(
      () => document.querySelectorAll('[data-dxw-item-kind="wordart"]').length,
    );
    expect(count).toBe(0);
  });
});

test.describe("in-front image drag around the footer", () => {
  test("image stays visible, footer survives, drops clamp to the page", async ({ page }) => {
    // User report: dragging an in-front image around the footer made things
    // disappear. Three underlying defects: in-front images painted UNDER the
    // text layer (z-order — clicks landed on spans, drags were dead), a drop
    // resolving to footer text moved the run into the footer part where its
    // r:embed rel dangles (image destroyed), and off-page drops stranded the
    // image outside the sheet.
    await page.goto("/?doc=/fixtures/benchmark.docx");
    await page.waitForSelector(".dxw-page span", { state: "attached" });
    await page.waitForTimeout(500);

    const red = () => page.locator('.dxw-page [data-dxw-item-kind="image"][style*="width: 140px"]').first();
    const footerCount = () =>
      page.evaluate(() => {
        const pg = document.querySelector(".dxw-page")!;
        const pr = pg.getBoundingClientRect();
        return [...pg.querySelectorAll('[data-dxw-hf="1"]')].filter(
          (e) => parseFloat((e as HTMLElement).style.top || "0") > pr.height * 0.6 && e.textContent!.trim(),
        ).length;
      });
    const f0 = await footerCount();
    expect(f0).toBeGreaterThan(0);

    await red().click();
    await page.waitForTimeout(200);
    await page.locator('button:text-is("In front")').click();
    await page.waitForTimeout(400);
    // In-front images paint above the text layer.
    expect(parseInt(await red().evaluate((element) => getComputedStyle(element).zIndex), 10)).toBeGreaterThan(0);

    // Drop directly on the footer text: image must survive, land near the
    // drop, and the footer must keep its spans.
    const target = await page.evaluate(() => {
      const pg = document.querySelector(".dxw-page")!;
      const pr = pg.getBoundingClientRect();
      const hf = [...pg.querySelectorAll<HTMLElement>('[data-dxw-hf="1"]')].filter(
        (e) => parseFloat(e.style.top || "0") > pr.height * 0.6 && e.textContent!.trim(),
      );
      const r = hf[0].getBoundingClientRect();
      return { x: r.x + r.width / 2, y: r.y + r.height / 2 };
    });
    let b = (await red().boundingBox())!;
    await page.mouse.move(b.x + b.width / 2, b.y + b.height / 2);
    await page.mouse.down();
    await page.mouse.move(target.x, target.y, { steps: 8 });
    await page.mouse.up();
    await page.waitForTimeout(500);
    expect(await red().count()).toBe(1);
    expect(await footerCount()).toBe(f0);
    const pos = await red().evaluate((el) => {
      const pg = el.closest(".dxw-page")!;
      const pr = pg.getBoundingClientRect();
      const r = el.getBoundingClientRect();
      return { y: r.y - pr.y, pageH: pr.height };
    });
    expect(pos.y).toBeGreaterThan(pos.pageH * 0.75); // landed in the footer band

    // Drop into the page gap: clamps onto a page, never stranded off-sheet.
    await red().scrollIntoViewIfNeeded();
    await page.waitForTimeout(150);
    b = (await red().boundingBox())!;
    const pg1 = (await page.locator(".dxw-page").first().boundingBox())!;
    await page.mouse.move(b.x + b.width / 2, b.y + b.height / 2);
    await page.mouse.down();
    await page.mouse.move(pg1.x + pg1.width / 2, b.y + b.height / 2 + 60, { steps: 8 });
    await page.mouse.up();
    await page.waitForTimeout(500);
    expect(await red().count()).toBe(1);
    const pos2 = await red().evaluate((el) => {
      const pg = el.closest(".dxw-page")!;
      const pr = pg.getBoundingClientRect();
      const r = el.getBoundingClientRect();
      return { top: r.y - pr.y, bottom: r.bottom - pr.y, pageH: pr.height };
    });
    expect(pos2.top).toBeGreaterThanOrEqual(-1);
    expect(pos2.bottom).toBeLessThanOrEqual(pos2.pageH + 90); // on the sheet (small engine bias tolerated)
  });
});
