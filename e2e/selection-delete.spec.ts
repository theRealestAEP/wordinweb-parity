import { expect, test, Page } from "@playwright/test";
import { readFileSync } from "node:fs";
import { strFromU8, unzipSync } from "fflate";

const MOD = process.platform === "darwin" ? "Meta" : "Control";

async function documentXml(page: Page): Promise<string> {
  const downloadPromise = page.waitForEvent("download");
  await page.getByText("Download", { exact: true }).click();
  const download = await downloadPromise;
  const path = await download.path();
  expect(path).not.toBeNull();
  return strFromU8(unzipSync(new Uint8Array(readFileSync(path!)))["word/document.xml"]);
}

test("Cmd+A then Delete removes list and table structure", async ({ page }) => {
  await page.goto("/?doc=/fixtures/sample.docx");
  await page.locator(".dxw-page span", { hasText: /^Lorem$/ }).first().click();

  await page.keyboard.press(`${MOD}+a`);
  await page.keyboard.press("Delete");

  await expect(page.locator(".dxw-page")).toHaveCount(1);
  await expect(page.locator('[data-dxw-role="table-rule"]')).toHaveCount(0);
  const xml = await documentXml(page);
  expect(xml).not.toContain("<w:tbl");
  expect(xml).not.toContain("<w:numPr");
});

test("deleting fully selected tables removes their structure", async ({ page }) => {
  await page.goto("/?doc=/fixtures/parity-tables.docx");
  await page.waitForSelector(".dxw-page span");

  const first = page.locator(".dxw-page span", { hasText: /^Key$/ }).first();
  const last = page.locator(".dxw-page span", { hasText: /^Plain$/ }).last();
  const firstBox = (await first.boundingBox())!;
  const lastBox = (await last.boundingBox())!;
  await page.mouse.move(firstBox.x + 1, firstBox.y + firstBox.height / 2);
  await page.mouse.down();
  await page.mouse.move(lastBox.x + lastBox.width - 1, lastBox.y + lastBox.height / 2, { steps: 20 });
  await page.mouse.up();
  await page.keyboard.press("Delete");

  await expect(page.locator('[data-dxw-role="table-rule"]')).toHaveCount(0);
  expect(await documentXml(page)).not.toContain("<w:tbl");
});

test("deleting a fully selected list item clears its list formatting", async ({ page }) => {
  await page.goto("/?doc=/fixtures/parity-lists.docx");
  const item = page.locator(".dxw-page span", { hasText: /^First$/ }).first();
  await item.click({ clickCount: 3 });
  await page.keyboard.press("Delete");
  await page.keyboard.type("Replacement paragraph");

  const xml = await documentXml(page);
  const paragraph = xml
    .match(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g)
    ?.find((candidate) => candidate.includes("Replacement paragraph"));
  expect(paragraph).toBeTruthy();
  expect(paragraph).not.toContain("<w:numPr");
});

test("deleting a fully selected style-based list item clears its list style", async ({ page }) => {
  await page.goto("/?doc=/fixtures/real.docx");
  const item = page.locator(".dxw-page span", { hasText: /^Lebatujofuv:$/ }).first();
  await item.click({ clickCount: 3 });
  await page.keyboard.press("Delete");
  await page.keyboard.type("Replacement styled paragraph");

  const xml = await documentXml(page);
  const paragraph = xml
    .match(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g)
    ?.find((candidate) => candidate.includes("Replacement styled paragraph"));
  expect(paragraph).toBeTruthy();
  expect(paragraph).not.toContain("<w:numPr");
  expect(paragraph).not.toContain("<w:pStyle");
});
