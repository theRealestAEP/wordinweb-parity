import { expect, test } from "@playwright/test";
import { readFileSync } from "node:fs";
import { strFromU8, unzipSync } from "fflate";

test("Layout uses a clean tab row and a single full-width control row", async ({ page }) => {
  await page.goto("/?doc=/fixtures/preset-publication.docx");
  await expect(page.locator(".dxw-page")).toHaveCount(1);
  const layoutTab = page.getByRole("button", { name: "layout", exact: true });
  await layoutTab.click();

  const tabBox = (await layoutTab.locator("..").boundingBox())!;
  const ribbon = page.locator("[data-dxw-layout-ribbon]");
  const ribbonBox = (await ribbon.boundingBox())!;
  expect(ribbonBox.y).toBeGreaterThanOrEqual(tabBox.y + tabBox.height);

  const controlTops = await ribbon.evaluate((element) =>
    [...element.querySelectorAll<HTMLElement>("[data-dxw-menu-select-trigger], [data-dxw-layout-menu-trigger]")]
      .map((control) => Math.round(control.getBoundingClientRect().top)),
  );
  expect(Math.max(...controlTops) - Math.min(...controlTops)).toBeLessThanOrEqual(2);
  await expect(page.getByTitle("Align selected object to page")).toHaveCount(0);
  await expect(page.getByTitle("Rotate selected object")).toHaveCount(0);
  await expect(page.getByTitle("Change selected object stacking order")).toHaveCount(0);
});

test("Magazine layout identifies the caret section and changes only that section's columns", async ({ page }) => {
  await page.goto("/?doc=/fixtures/preset-publication.docx");
  const text = (value: string) => page.locator(".dxw-page span").filter({ hasText: value }).first();
  await expect(text("SPECIAL")).toBeVisible();

  await text("SPECIAL").click();
  await page.getByRole("button", { name: "layout", exact: true }).click();
  const scope = page.locator('select[title="Apply layout changes to"]');
  await scope.selectOption("section");
  const trigger = scope.locator("..").getByRole("button");
  await expect(trigger).toContainText("This section · 2 of 3");

  await page.locator('[data-dxw-layout-menu-trigger="columns"]').click();
  await page.locator('[data-dxw-layout-option="1"]').click();

  const pending = page.waitForEvent("download");
  await page.getByText("Download", { exact: true }).click();
  const path = await (await pending).path();
  if (!path) throw new Error("download path unavailable");
  const documentXml = strFromU8(unzipSync(new Uint8Array(readFileSync(path)))["word/document.xml"]);
  expect(documentXml).not.toContain("<w:cols");

  await text("◆").click();
  await expect(trigger).toContainText("This section · 3 of 3");
  await text("SUNDAY").click();
  await expect(trigger).toContainText("This section · 1 of 3");
});
