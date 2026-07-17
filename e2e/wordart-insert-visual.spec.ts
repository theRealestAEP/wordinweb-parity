import { expect, test } from "@playwright/test";

test("advanced Insert renders native wave WordArt with selection chrome", async ({ page }) => {
  await page.goto("/?doc=/fixtures/benchmark.docx");
  await page.waitForSelector(".dxw-page span");
  await page.getByRole("button", { name: "insert", exact: true }).click();
  const target = page.locator(".dxw-page span", { hasText: "classic." }).first();
  await target.click();
  await page.keyboard.press("End");
  await page.getByTitle("Insert WordArt").click();
  await page.getByLabel("WordArt text").fill("WORDINWEB");
  await page.getByTitle("Insert WordArt Wave").click();

  const drawing = page.locator("[data-dxw-drawing]").last();
  await drawing.click();
  await page.getByRole("button", { name: "layout", exact: true }).click();
  await page.getByTitle("Align selected object to page").selectOption("alignCenter");
  await expect(page.locator("[data-dxw-img-handle]")).toHaveCount(8);
  await expect(page.locator("[data-dxw-object-selection]")).toHaveScreenshot("inserted-wordart.png", {
    animations: "disabled",
    caret: "hide",
  });
});
