import { expect, test } from "@playwright/test";
import { readFileSync } from "node:fs";
import { strFromU8, unzipSync } from "fflate";

const SVG = '<svg xmlns="http://www.w3.org/2000/svg" width="96" height="96" viewBox="0 0 96 96"><path fill="#2e74b5" d="M48 5l11 28h30L65 51l9 30-26-18-26 18 9-30L7 33h30z"/></svg>';
const MOD = process.platform === "darwin" ? "Meta" : "Control";

test("advanced Insert adds a native SVG icon with undo and save", async ({ page }) => {
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
  await expect(icon).toHaveCount(1);
  await expect(icon).toBeVisible();
  await icon.click();
  await expect(page.locator("[data-dxw-img-handle]")).toHaveCount(8);
  await page.keyboard.press(`${MOD}+z`);
  await expect(icon).toHaveCount(0);
  await page.keyboard.press(`${MOD}+Shift+z`);
  await expect(icon).toHaveCount(1);

  const pending = page.waitForEvent("download");
  await page.getByText("Download", { exact: true }).click();
  const path = await (await pending).path();
  expect(path).not.toBeNull();
  const files = unzipSync(new Uint8Array(readFileSync(path!)));
  expect(strFromU8(files["[Content_Types].xml"])).toContain('Extension="svg" ContentType="image/svg+xml"');
  expect(strFromU8(files["word/media/image1.svg"])).toContain('<path fill="#2e74b5"');
});
