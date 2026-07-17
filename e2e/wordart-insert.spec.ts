import { expect, Page, test } from "@playwright/test";
import { readFileSync } from "node:fs";
import { strFromU8, unzipSync } from "fflate";

const MOD = process.platform === "darwin" ? "Meta" : "Control";

async function openInsert(page: Page): Promise<void> {
  await page.goto("/?doc=/fixtures/benchmark.docx");
  await page.waitForSelector(".dxw-page span");
  await page.getByRole("button", { name: "insert", exact: true }).click();
  const target = page.locator(".dxw-page span", { hasText: "classic." }).first();
  await target.click();
  await page.keyboard.press("End");
}

async function downloadXml(page: Page): Promise<string> {
  const pending = page.waitForEvent("download");
  await page.getByText("Download", { exact: true }).click();
  const path = await (await pending).path();
  expect(path).not.toBeNull();
  return strFromU8(unzipSync(new Uint8Array(readFileSync(path!)))["word/document.xml"]);
}

test("advanced Insert creates editable native WordArt with undo, redo, and save", async ({ page }) => {
  await openInsert(page);
  await page.getByTitle("Insert WordArt").click();
  await page.getByLabel("WordArt text").fill("Quarterly momentum");
  await page.getByTitle("Insert WordArt Arch up").click();

  const drawing = page.locator("[data-dxw-drawing]").last();
  await expect(drawing).toBeVisible();
  await expect(page.locator("svg text", { hasText: "Quarterly momentum" })).toBeVisible();
  await drawing.click();
  await expect(page.locator("[data-dxw-img-handle]")).toHaveCount(8);

  await page.keyboard.press(`${MOD}+z`);
  await expect(page.locator("svg text", { hasText: "Quarterly momentum" })).toHaveCount(0);
  await page.keyboard.press(`${MOD}+Shift+z`);
  await expect(page.locator("svg text", { hasText: "Quarterly momentum" })).toBeVisible();

  const xml = await downloadXml(page);
  expect(xml).toContain('name="WordArt ');
  expect(xml).toContain('<a:prstTxWarp prst="textArchUp">');
  expect(xml).toContain("Quarterly momentum");
  expect(xml).toContain('<w:color w:val="2E74B5"');
});

test("plain WordArt has tight selection bounds and edits inline through save/reopen", async ({ page }) => {
  await openInsert(page);
  await page.getByTitle("Insert WordArt").click();
  await page.getByLabel("WordArt text").fill("Direct edit");
  await page.getByTitle("Insert WordArt Plain").click();

  const drawing = page.locator("[data-dxw-drawing]").last();
  const ink = page.locator('[data-dxw-item-kind="text"]').filter({ hasText: /^Direct$/ }).last();
  await expect(drawing).toBeVisible();
  await expect(ink).toBeVisible();
  const inkBox = await ink.boundingBox();
  expect(inkBox).not.toBeNull();
  await page.mouse.click(inkBox!.x + inkBox!.width / 2, inkBox!.y + inkBox!.height / 2);
  const selection = page.locator("[data-dxw-object-selection]");
  await expect(selection).toBeVisible();
  const selectionBox = await selection.boundingBox();
  expect(selectionBox).not.toBeNull();
  expect(selectionBox!.height).toBeLessThanOrEqual(inkBox!.height + 18);
  expect(Math.abs(selectionBox!.y - inkBox!.y)).toBeLessThanOrEqual(12);

  await page.mouse.dblclick(inkBox!.x + inkBox!.width / 2, inkBox!.y + inkBox!.height / 2);
  const editor = page.getByLabel("Edit WordArt text");
  await expect(editor).toBeVisible();
  const editorBox = await editor.boundingBox();
  expect(editorBox).not.toBeNull();
  expect(editorBox!.height).toBeLessThanOrEqual(inkBox!.height + 18);
  await editor.fill("Edited inline");
  await editor.press("Enter");
  await expect(page.locator('[data-dxw-item-kind="text"]').filter({ hasText: /^Edited$/ }).last()).toBeVisible();

  const pending = page.waitForEvent("download");
  await page.getByText("Download", { exact: true }).click();
  const path = await (await pending).path();
  expect(path).not.toBeNull();
  const saved = strFromU8(unzipSync(new Uint8Array(readFileSync(path!)))["word/document.xml"]);
  expect(saved).toContain("Edited inline");
  expect(saved).not.toContain(">Direct edit<");

  await page.locator("#docx-upload").setInputFiles(path!);
  await page.waitForSelector(".dxw-page span");
  const editedInk = page.locator('[data-dxw-item-kind="text"]').filter({ hasText: /^Edited$/ }).last();
  await expect(editedInk).toBeVisible();
  const editedInkBox = await editedInk.boundingBox();
  expect(editedInkBox).not.toBeNull();
  await page.mouse.dblclick(editedInkBox!.x + editedInkBox!.width / 2, editedInkBox!.y + editedInkBox!.height / 2);
  await expect(page.getByLabel("Edit WordArt text")).toHaveValue("Edited inline");
});
