import { expect, test } from "@playwright/test";

const SVG = '<svg xmlns="http://www.w3.org/2000/svg" width="96" height="96" viewBox="0 0 96 96"><path fill="#2e74b5" d="M48 5l11 28h30L65 51l9 30-26-18-26 18 9-30L7 33h30z"/></svg>';

test("advanced Insert renders a selected SVG icon", async ({ page }) => {
  await page.goto("/?doc=/fixtures/benchmark.docx");
  await page.waitForSelector(".dxw-page span");
  await page.getByRole("button", { name: "insert", exact: true }).click();
  const target = page.locator(".dxw-page span", { hasText: "classic." }).first();
  await target.click();
  await page.keyboard.press("End");
  await page.getByTitle("Insert SVG icon").click();
  await page.locator('input[accept=".svg"]').setInputFiles({
    name: "star.svg",
    mimeType: "image/svg+xml",
    buffer: Buffer.from(SVG),
  });
  const icon = page.locator('.dxw-page img[style*="width: 96px"][style*="height: 96px"]');
  await icon.click();
  await expect(page.locator("[data-dxw-img-handle]")).toHaveCount(8);
  await page.evaluate(() => window.scrollTo(0, 0));
  await expect(page.locator(".dxw-page").first()).toHaveScreenshot("inserted-svg-icon.png", {
    animations: "disabled",
    caret: "hide",
  });
});
