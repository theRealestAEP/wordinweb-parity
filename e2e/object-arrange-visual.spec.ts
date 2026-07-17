import { expect, test } from "@playwright/test";

test("Layout Arrange renders a rotated centered shape with selection chrome", async ({ page }) => {
  await page.goto("/?doc=/fixtures/benchmark.docx");
  await page.waitForSelector(".dxw-page span");
  await page.getByRole("button", { name: "insert", exact: true }).click();
  const target = page.locator(".dxw-page span", { hasText: "classic." }).first();
  await target.click();
  await page.keyboard.press("End");
  await page.getByTitle("Insert shape").click();
  await page.getByLabel("Shape text").fill("Layout object");
  await page.getByTitle("Insert Rounded rectangle").click();
  const shape = page.locator("[data-dxw-drawing]").last();
  const box = await shape.boundingBox();
  await page.mouse.click(box!.x + box!.width - 12, box!.y + box!.height - 12);
  await page.getByRole("button", { name: "layout", exact: true }).click();
  await page.getByTitle("Align selected object to page").selectOption("alignCenter");
  await page.getByTitle("Align selected object to page").selectOption("alignMiddle");
  await page.getByTitle("Rotate selected object").selectOption("rotateRight");
  await expect(page.locator("[data-dxw-img-handle]")).toHaveCount(8);
  await expect(page.locator("[data-dxw-object-selection]")).toHaveScreenshot("layout-arranged-shape.png", {
    animations: "disabled",
    caret: "hide",
  });
});
