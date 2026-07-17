import { expect, test } from "@playwright/test";

test("advanced Insert content has a stable two-page visual", async ({ page }) => {
  await page.setViewportSize({ width: 1400, height: 1400 });
  await page.goto("/?doc=/fixtures/insert-editing-parity.docx&editable=0&comments=0");
  const pages = page.locator(".dxw-pages");
  await expect(page.locator(".dxw-page")).toHaveCount(2);
  await expect(pages).toContainText("Quarterly Revenue");
  await expect(pages).toContainText("Symbols: Ω");
  await expect(pages).toContainText("Revenue target is on page 1");
  await expect(pages).toContainText("Rounded rectangle");
  await expect(pages).toContainText("Editable text box");
  for (let index = 0; index < 2; index++) {
    const onePage = page.locator(".dxw-page").nth(index);
    await onePage.scrollIntoViewIfNeeded();
    await expect(onePage).toHaveScreenshot(`insert-editing-parity-p${index + 1}.png`, {
      animations: "disabled",
      caret: "hide",
    });
  }
});
