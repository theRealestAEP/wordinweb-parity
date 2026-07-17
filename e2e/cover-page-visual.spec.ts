import { expect, Page, test } from "@playwright/test";

function tool(page: Page, title: string) {
  return page.locator(`[title=${JSON.stringify(title)}], [data-tip=${JSON.stringify(title)}]`).first();
}

test("inserted cover page has a stable Word-style visual", async ({ page }) => {
  await page.setViewportSize({ width: 1400, height: 1200 });
  await page.goto("/?doc=/fixtures/parity-text.docx&comments=0");
  await page.waitForSelector(".dxw-page span");
  await page.getByRole("button", { name: "insert", exact: true }).click();
  await tool(page, "Insert cover page").click();
  await page.getByLabel("Cover title").fill("Project Atlas");
  await page.getByLabel("Cover subtitle").fill("Launch readiness report");
  await page.getByLabel("Cover author").fill("Ada Lovelace");
  await page.getByRole("button", { name: "Insert cover", exact: true }).click();
  const cover = page.locator(".dxw-page").first();
  await expect(cover).toContainText("Project Atlas");
  await expect(cover).toHaveScreenshot("inserted-cover-page.png", { animations: "disabled", caret: "hide" });
});
