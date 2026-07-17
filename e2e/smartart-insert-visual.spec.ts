import { expect, test } from "@playwright/test";

test("advanced Insert renders selected native SmartArt", async ({ page }) => {
  await page.goto("/?doc=/fixtures/benchmark.docx");
  await page.waitForSelector(".dxw-page span");
  await page.getByRole("button", { name: "insert", exact: true }).click();
  const target = page.locator(".dxw-page span", { hasText: "classic." }).first();
  await target.click();
  await page.keyboard.press("End");
  await page.getByRole("button", { name: "SmartArt", exact: true }).click();
  await page.getByRole("button", { name: "Insert or update SmartArt" }).click();
  const smartArt = page.locator("[data-dxw-smart-art]").last();
  const smartArtBox = await smartArt.boundingBox();
  expect(smartArtBox).not.toBeNull();
  await page.mouse.click(smartArtBox!.x + 6, smartArtBox!.y + 6);
  await expect(page.locator("[data-dxw-img-handle]")).toHaveCount(8);
  await expect(page.locator(".dxw-page").first()).toHaveScreenshot("inserted-smartart.png", {
    animations: "disabled",
    caret: "hide",
  });
});
