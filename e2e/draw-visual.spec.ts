import { expect, Page, test } from "@playwright/test";

function tool(page: Page, title: string) {
  return page.locator(`[title=${JSON.stringify(title)}], [data-tip=${JSON.stringify(title)}]`).first();
}

test("Draw ribbon, ink, and lasso selection have a stable visual", async ({ page }) => {
  await page.setViewportSize({ width: 1400, height: 1200 });
  await page.goto("/?doc=/fixtures/insert-editing-parity.docx&comments=0");
  await page.waitForSelector(".dxw-page span");
  await page.getByRole("button", { name: "draw", exact: true }).click();
  await tool(page, "Draw with pen").click();

  const toolbar = page.locator('[data-dxw-toolbar-mode="advanced"]');
  await expect(toolbar).toHaveScreenshot("draw-ribbon.png", { animations: "disabled", caret: "hide" });

  const pageOne = page.locator(".dxw-page").first();
  const box = await pageOne.boundingBox();
  expect(box).not.toBeNull();
  await page.mouse.move(box!.x + 400, box!.y + 500);
  await page.mouse.down();
  await page.mouse.move(box!.x + 475, box!.y + 535, { steps: 8 });
  await page.mouse.up();

  await tool(page, "Draw with highlighter").click();
  await page.mouse.move(box!.x + 395, box!.y + 555);
  await page.mouse.down();
  await page.mouse.move(box!.x + 485, box!.y + 555, { steps: 8 });
  await page.mouse.up();

  await tool(page, "Lasso ink").click();
  await page.mouse.move(box!.x + 375, box!.y + 475);
  await page.mouse.down();
  for (const point of [
    [box!.x + 510, box!.y + 475],
    [box!.x + 510, box!.y + 585],
    [box!.x + 375, box!.y + 585],
    [box!.x + 375, box!.y + 475],
  ] as const) await page.mouse.move(...point, { steps: 5 });
  await page.mouse.up();
  await expect(page.locator("[data-dxw-ink-selection]")).toBeVisible();
  await expect(pageOne).toHaveScreenshot("draw-ink-selection.png", { animations: "disabled", caret: "hide" });
});
