import { expect, Page, test } from "@playwright/test";
import { readFileSync } from "node:fs";
import { strFromU8, unzipSync } from "fflate";
import { Editor } from "./editing.js";

const MOD = process.platform === "darwin" ? "Meta" : "Control";

function tool(page: Page, title: string) {
  return page.locator(`[title=${JSON.stringify(title)}], [data-tip=${JSON.stringify(title)}]`).first();
}

test("Cover Page prepends editable styled content with undo, redo, and save-back", async ({ page }) => {
  const editor = await Editor.open(page, "parity-text");
  const before = await editor.pageCount();
  await page.getByRole("button", { name: "insert", exact: true }).click();
  await tool(page, "Insert cover page").click();
  await page.getByLabel("Cover title").fill("Project Atlas");
  await page.getByLabel("Cover subtitle").fill("Launch plan");
  await page.getByLabel("Cover author").fill("Ada Lovelace");
  await page.getByRole("button", { name: "Insert cover", exact: true }).click();
  await expect(page.locator(".dxw-page")).toHaveCount(before + 1);
  await expect(page.locator(".dxw-page").first()).toContainText("Project Atlas");
  expect(await editor.pageOf("Plain")).toBe(1);

  await editor.clickText("Plain");
  await page.keyboard.press(`${MOD}+z`);
  await expect(page.locator(".dxw-pages")).not.toContainText("Project Atlas");
  await expect(page.locator(".dxw-page")).toHaveCount(before);
  await page.keyboard.press(process.platform === "darwin" ? "Meta+Shift+z" : "Control+y");
  await expect(page.locator(".dxw-page")).toHaveCount(before + 1);
  await expect(page.locator(".dxw-page").first()).toContainText("Project Atlas");

  const pending = page.waitForEvent("download");
  await page.getByText("Download", { exact: true }).click();
  const path = await (await pending).path();
  const xml = strFromU8(unzipSync(new Uint8Array(readFileSync(path!)))["word/document.xml"]);
  expect(xml).toContain("Project Atlas");
  expect(xml).toContain("Launch plan");
  expect(xml).toContain("Ada Lovelace");
  expect(xml).toContain('w:pStyle w:val="Title"');
});

test("Blank Page inserts two page breaks with undo, redo, and save-back", async ({ page }) => {
  const editor = await Editor.open(page, "parity-text");
  const before = await editor.pageCount();
  await editor.clickText("Plain");
  await page.getByRole("button", { name: "insert", exact: true }).click();
  await tool(page, "Insert blank page").click();
  await expect(page.locator(".dxw-page")).toHaveCount(before + 2);
  await expect(page.locator(".page-count")).toHaveText(`${before + 2} pages`);

  await page.keyboard.press(`${MOD}+z`);
  await expect(page.locator(".dxw-page")).toHaveCount(before);
  await expect(page.locator(".page-count")).toHaveText(`${before} page${before === 1 ? "" : "s"}`);
  await page.keyboard.press(process.platform === "darwin" ? "Meta+Shift+z" : "Control+y");
  await expect(page.locator(".dxw-page")).toHaveCount(before + 2);
  await expect(page.locator(".page-count")).toHaveText(`${before + 2} pages`);

  const pending = page.waitForEvent("download");
  await page.getByText("Download", { exact: true }).click();
  const path = await (await pending).path();
  const xml = strFromU8(unzipSync(new Uint8Array(readFileSync(path!)))["word/document.xml"]);
  expect((xml.match(/<w:br w:type="page"\/>/g) ?? [])).toHaveLength(2);
});

test("a newly inserted blank page accepts a clicked caret and text", async ({ page }) => {
  const editor = await Editor.open(page, "parity-text");
  await editor.clickText("Plain");
  await page.getByRole("button", { name: "insert", exact: true }).click();
  await tool(page, "Insert blank page").click();

  const blank = page.locator(".dxw-page").nth(1);
  await blank.scrollIntoViewIfNeeded();
  const box = (await blank.boundingBox())!;
  await page.mouse.click(box.x + 120, box.y + 180);
  await page.keyboard.type("BLANK PAGE TEXT");
  await expect(blank).toContainText("BLANK PAGE TEXT");

  const pending = page.waitForEvent("download");
  await page.getByText("Download", { exact: true }).click();
  const path = await (await pending).path();
  const xml = strFromU8(unzipSync(new Uint8Array(readFileSync(path!)))["word/document.xml"]);
  expect(xml).toContain("BLANK PAGE TEXT");
});

test("an inserted current date has editable positions beside and below it", async ({ page }) => {
  const editor = await Editor.open(page, "parity-text");
  await editor.clickText("Plain");
  await page.keyboard.press("End");
  await page.keyboard.press("Enter");
  await page.getByRole("button", { name: "insert", exact: true }).click();
  await tool(page, "Insert an automatically updating date or time").selectOption("date:short");

  const date = page.locator(".dxw-page span").filter({ hasText: /^\d{1,2}\/\d{1,2}\/\d{4}$/ }).last();
  const dateBox = (await date.boundingBox())!;
  await page.mouse.click(dateBox.x + dateBox.width + 4, dateBox.y + dateBox.height / 2);
  await page.keyboard.type("AFTERDATE");
  await expect(page.locator(".dxw-pages")).toContainText("AFTERDATE");
  await page.keyboard.press("Enter");

  await page.mouse.click(dateBox.x, dateBox.y + dateBox.height + 22);
  await page.keyboard.type("BELOWDATE");
  const below = (await page.locator(".dxw-page span", { hasText: "BELOWDATE" }).boundingBox())!;
  expect(below.y).toBeGreaterThan(dateBox.y + dateBox.height - 1);
});

test("Insert Header and Footer enter editing directly, Escape closes, and both save", async ({ page }) => {
  const editor = await Editor.open(page, "parity-text");
  await editor.clickText("Plain");
  await page.getByRole("button", { name: "insert", exact: true }).click();

  await tool(page, "Edit header").click();
  await expect(page.locator("[data-dxw-hf-hotbar]")).toBeVisible();
  await page.keyboard.type("DIRECT HEADER");
  await expect(page.locator(".dxw-pages")).toContainText("DIRECT HEADER");
  await page.keyboard.press("Escape");
  await expect(page.locator("[data-dxw-hf-hotbar]")).toHaveCount(0);

  await tool(page, "Edit footer").click();
  await expect(page.locator("[data-dxw-hf-hotbar]")).toBeVisible();
  await page.keyboard.type("DIRECT FOOTER");
  await expect(page.locator(".dxw-pages")).toContainText("DIRECT FOOTER");
  await page.keyboard.press("Escape");
  await expect(page.locator("[data-dxw-hf-hotbar]")).toHaveCount(0);

  const pending = page.waitForEvent("download");
  await page.getByText("Download", { exact: true }).click();
  const path = await (await pending).path();
  const parts = unzipSync(new Uint8Array(readFileSync(path!)));
  const headers = Object.entries(parts).filter(([name]) => /^word\/header\d+\.xml$/.test(name)).map(([, bytes]) => strFromU8(bytes)).join("\n");
  const footers = Object.entries(parts).filter(([name]) => /^word\/footer\d+\.xml$/.test(name)).map(([, bytes]) => strFromU8(bytes)).join("\n");
  expect(headers).toContain("DIRECT HEADER");
  expect(footers).toContain("DIRECT FOOTER");
});
