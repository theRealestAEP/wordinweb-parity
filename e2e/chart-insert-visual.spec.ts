import { expect, test } from "@playwright/test";

test("advanced Insert renders a native chart with selection chrome", async ({ page }) => {
  await page.goto("/?doc=/fixtures/benchmark.docx");
  await page.waitForSelector(".dxw-page span");
  await page.getByRole("button", { name: "insert", exact: true }).click();
  const target = page.locator(".dxw-page span", { hasText: "classic." }).first();
  await target.click();
  await page.keyboard.press("End");
  await page.getByRole("button", { name: "Chart", exact: true }).click();
  await page.getByRole("button", { name: "Insert or update chart" }).click();

  const chart = page.locator("[data-dxw-chart]").last();
  const box = await chart.boundingBox();
  expect(box).not.toBeNull();
  await page.mouse.click(box!.x + box!.width / 2, box!.y + box!.height / 2);
  await expect(page.locator("[data-dxw-img-handle]")).toHaveCount(8);
  await expect(page.locator(".dxw-page").first()).toHaveScreenshot("inserted-chart.png", {
    animations: "disabled",
    caret: "hide",
  });
});
