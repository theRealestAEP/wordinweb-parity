import { test, expect } from "@playwright/test";

test("wordart-warps: colored fill selects the shape (no click-through)", async ({ page }) => {
  await page.goto(`/?doc=/fixtures/probe3-wordart-warps.docx`);
  await page.waitForTimeout(1500);
  const fill = await page.evaluate(() => {
    const h = document.querySelector('[data-dxw-drawing][style*="z-index: 1"]') as HTMLElement | null;
    if (!h) return null;
    const r = h.getBoundingClientRect();
    return { x: Math.round(r.x + 8), y: Math.round(r.y + 8) };
  });
  expect(fill).not.toBeNull();
  await page.mouse.click(fill!.x, fill!.y);
  await page.waitForTimeout(150);
  await expect(page.locator("[data-dxw-img-handle]")).toHaveCount(8);
});

test("phase23 flowchart shape: fill selects, its text glyph still edits", async ({ page }) => {
  await page.goto(`/?doc=/fixtures/wild2-med-phase23-protocol.docx`);
  await page.waitForTimeout(1500);
  // Scroll a flowchart shape (real textbox text) into view and read a bare-fill
  // point plus a glyph point inside it.
  let pts: { fill: { x: number; y: number }; glyph: { x: number; y: number } } | null = null;
  const pages = page.locator(".dxw-page");
  for (let i = 0; i < await pages.count() && !pts; i++) {
    await pages.nth(i).scrollIntoViewIfNeeded();
    await page.waitForTimeout(50);
    pts = await page.evaluate(() => {
      const hits = Array.from(document.querySelectorAll('[data-dxw-drawing][style*="z-index: 1"]')) as HTMLElement[];
      // Pick a hit that actually contains an editable text span.
      for (const h of hits) {
        h.scrollIntoView({ block: "center" });
        const r = h.getBoundingClientRect();
        const spans = Array.from(document.querySelectorAll(".dxw-page span")) as HTMLElement[];
        for (const s of spans) {
          const sr = s.getBoundingClientRect();
          const cx = sr.x + sr.width / 2;
          const cy = sr.y + sr.height / 2;
          if (sr.width > 2 && s.textContent && cx > r.x + 3 && cx < r.right - 3 && cy > r.y + 3 && cy < r.bottom - 3) {
            return { fill: { x: Math.round(r.x + 5), y: Math.round(r.y + 4) }, glyph: { x: Math.round(cx), y: Math.round(cy) } };
          }
        }
      }
      return null;
    });
  }
  expect(pts).not.toBeNull();

  await page.mouse.click(pts!.fill.x, pts!.fill.y);
  await page.waitForTimeout(150);
  await expect(page.locator("[data-dxw-img-handle]")).toHaveCount(8);

  // Clicking the shape's own text deselects the shape and enters text editing.
  await page.mouse.click(pts!.glyph.x, pts!.glyph.y);
  await page.waitForTimeout(150);
  await expect(page.locator("[data-dxw-img-handle]")).toHaveCount(0);
});
