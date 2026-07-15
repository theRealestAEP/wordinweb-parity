import { test, expect, Page } from "@playwright/test";

// probe2-picture-watermark carries a floating picture watermark in its header
// (one instance painted per page, data-dxw-hf). These specs prove header
// images are only interactive while editing the header/footer — a body-mode
// click must not select them, and there is no body-mode path to drag (and lose)
// one. In hf-edit mode they select and drag, and a drag never deletes them
// (a cross-part move would dangle the part-scoped r:embed rel).

async function open(page: Page): Promise<void> {
  await page.goto(`/?doc=/fixtures/probe2-picture-watermark.docx`);
  await page.waitForSelector(".dxw-page", { timeout: 8000 });
  await page.waitForTimeout(400);
}

/** First header image geometry + the current header-image count. */
async function hfImage(page: Page) {
  return page.evaluate(() => {
    const el = document.querySelector<HTMLElement>("img[data-dxw-hf]");
    if (!el) return null;
    const r = el.getBoundingClientRect();
    const p = document.querySelector<HTMLElement>(".dxw-page")!.getBoundingClientRect();
    return {
      count: document.querySelectorAll("img[data-dxw-hf]").length,
      x: r.x, y: r.y, w: r.width, h: r.height,
      cornerX: r.x + r.width * 0.85, cornerY: r.y + r.height * 0.15, // over the pic, clear of body glyphs
      pageX: p.x, pageTop: p.y, pageW: p.width,
    };
  });
}

const handleCount = (page: Page) => page.evaluate(() => document.querySelectorAll("[data-dxw-img-handle]").length);
const hfMode = (page: Page) => page.evaluate(() => document.querySelector(".dxw-hf-mode") != null);

async function enterHf(page: Page, img: { pageX: number; pageTop: number; pageW: number }): Promise<void> {
  // Word UX: double-click the top margin band enters header editing.
  await page.mouse.dblclick(img.pageX + img.pageW / 2, img.pageTop + 25);
  await page.waitForTimeout(250);
}

test.describe("header image interaction is gated to header/footer edit mode", () => {
  test("body mode: clicking a header image does not select it", async ({ page }) => {
    await open(page);
    const img = (await hfImage(page))!;
    expect(img).toBeTruthy();
    await page.mouse.click(img.cornerX, img.cornerY);
    await page.waitForTimeout(150);
    expect(await handleCount(page)).toBe(0); // no resize chrome → not selected
  });

  test("body mode: a header image cannot be dragged (stays put, never disappears)", async ({ page }) => {
    await open(page);
    const before = (await hfImage(page))!;
    await page.mouse.move(before.cornerX, before.cornerY);
    await page.mouse.down();
    await page.mouse.move(before.cornerX + 30, before.cornerY + 300, { steps: 12 });
    await page.mouse.up();
    await page.waitForTimeout(300);
    const after = (await hfImage(page))!;
    expect(after.count).toBe(before.count); // none lost
    expect(Math.abs(after.y - before.y)).toBeLessThan(2); // did not move
  });

  test("hf mode: clicking a header image selects it", async ({ page }) => {
    await open(page);
    const img = (await hfImage(page))!;
    await enterHf(page, img);
    expect(await hfMode(page)).toBe(true);
    await page.mouse.click(img.cornerX, img.cornerY);
    await page.waitForTimeout(150);
    expect(await handleCount(page)).toBeGreaterThan(0); // selected
  });

  test("hf mode: dragging a header image never deletes it", async ({ page }) => {
    await open(page);
    const before = (await hfImage(page))!;
    await enterHf(page, before);
    await page.mouse.click(before.cornerX, before.cornerY);
    await page.waitForTimeout(120);
    await page.mouse.move(before.cornerX, before.cornerY);
    await page.mouse.down();
    await page.mouse.move(before.cornerX - 60, before.cornerY + 40, { steps: 10 });
    await page.mouse.up();
    await page.waitForTimeout(300);
    const after = (await hfImage(page))!;
    expect(after.count).toBe(before.count); // still present on every page
  });

  // A header image paints one instance per page from the SAME source drawing.
  // The post-drop residual correction must measure the instance the user
  // dragged — measured against page 1's copy, a drag on page 3 reads a
  // ~2-page "error" and flings the image off every sheet.
  test("hf mode: dragging a LATER page's instance keeps every instance on its page", async ({ page }) => {
    await open(page);
    const inst = await page.evaluate(() => {
      const imgs = [...document.querySelectorAll<HTMLElement>("img[data-dxw-hf]")];
      const el = imgs[Math.min(2, imgs.length - 1)]; // page-3 instance
      el.scrollIntoView({ block: "center" });
      const r = el.getBoundingClientRect();
      const pg = el.closest(".dxw-page")!.getBoundingClientRect();
      return {
        count: imgs.length,
        cx: r.x + r.width * 0.85, cy: r.y + r.height * 0.15,
        pageX: pg.x, pageTop: pg.y, pageW: pg.width,
      };
    });
    expect(inst.count).toBeGreaterThan(2);
    await enterHf(page, inst);
    await page.mouse.move(inst.cx, inst.cy);
    await page.mouse.down();
    await page.mouse.move(inst.cx + 40, inst.cy + 25, { steps: 10 });
    await page.mouse.up();
    await page.waitForTimeout(400);
    const after = await page.evaluate(() =>
      [...document.querySelectorAll<HTMLElement>("img[data-dxw-hf]")].map((el) => {
        const r = el.getBoundingClientRect();
        const pg = el.closest(".dxw-page")!.getBoundingClientRect();
        return r.bottom > pg.top && r.top < pg.bottom && r.right > pg.left && r.left < pg.right;
      }),
    );
    expect(after.length).toBe(inst.count); // none lost
    for (const visible of after) expect(visible).toBe(true); // none flung off-sheet
  });
});
