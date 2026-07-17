import { expect, test } from "@playwright/test";

test("inserted drop cap has a stable Word-style visual", async ({ page }) => {
  await page.goto("/?doc=/fixtures/parity-text.docx&comments=0");
  await page.waitForSelector(".dxw-page span");
  const body = page.locator(".dxw-page span", { hasText: /^Lorem$/ }).first();
  await body.click();
  await page.keyboard.press("Home");
  await page.getByRole("button", { name: "insert", exact: true }).click();
  await page.getByTitle("Drop cap").selectOption("drop");
  await expect(page.locator(".dxw-page span", { hasText: /^L$/ }).first()).toBeVisible();
  await expect(page.locator(".dxw-page").first()).toHaveScreenshot("inserted-drop-cap.png", {
    animations: "disabled",
    caret: "hide",
  });
});
