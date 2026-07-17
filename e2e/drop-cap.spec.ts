import { expect, Page, test } from "@playwright/test";
import { readFileSync } from "node:fs";
import { strFromU8, unzipSync } from "fflate";

function dropCapMenu(page: Page) {
  return page.locator('[title="Drop cap"], [data-tip="Drop cap"]').first();
}

async function openAtBody(page: Page) {
  await page.goto("/?doc=/fixtures/parity-text.docx&comments=0");
  await page.waitForSelector(".dxw-page span");
  const body = page.locator(".dxw-page span", { hasText: /^Lorem$/ }).first();
  await body.click();
  await page.keyboard.press("Home");
  await page.getByRole("button", { name: "insert", exact: true }).click();
}

test("Drop Cap applies, changes, removes, undoes, and saves native OOXML", async ({ page }) => {
  await openAtBody(page);
  await dropCapMenu(page).selectOption("drop");
  const cap = page.locator(".dxw-page span", { hasText: "L" }).filter({ hasText: /^L$/ }).first();
  await expect(cap).toBeVisible();
  expect(parseFloat(await cap.evaluate((element) => getComputedStyle(element).fontSize))).toBeGreaterThan(35);

  await page.keyboard.press(process.platform === "darwin" ? "Meta+z" : "Control+z");
  await expect(page.locator(".dxw-page span", { hasText: /^Lorem$/ }).first()).toBeVisible();
  await page.keyboard.press(process.platform === "darwin" ? "Meta+Shift+z" : "Control+y");
  await expect(cap).toBeVisible();

  const body = page.locator(".dxw-page span", { hasText: /^orem$/ }).first();
  await body.click();
  await page.getByRole("button", { name: "insert", exact: true }).click();
  await dropCapMenu(page).selectOption("margin");
  const pending = page.waitForEvent("download");
  await page.getByText("Download", { exact: true }).click();
  const path = await (await pending).path();
  const xml = strFromU8(unzipSync(new Uint8Array(readFileSync(path!)))["word/document.xml"]);
  expect(xml).toContain('w:dropCap="margin"');
  expect(xml).toContain('w:lines="3"');

  await body.click();
  await page.getByRole("button", { name: "insert", exact: true }).click();
  await dropCapMenu(page).selectOption("none");
  await expect(page.locator(".dxw-page span", { hasText: /^Lorem$/ }).first()).toBeVisible();
});
