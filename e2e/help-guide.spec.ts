import { expect, Page, test } from "@playwright/test";
import { readFileSync } from "node:fs";
import { strFromU8, unzipSync } from "fflate";

const MOD = process.platform === "darwin" ? "Meta" : "Control";

async function load(page: Page): Promise<void> {
  await page.goto("/?doc=/fixtures/parity-text.docx");
  await expect(page.getByText("Plain", { exact: true })).toBeVisible();
}

test("Help is searchable and covers tools, shortcuts, and special document recipes", async ({ page }) => {
  await load(page);
  const trigger = page.getByRole("button", { name: "Help", exact: true });
  await trigger.click();

  const dialog = page.getByRole("dialog", { name: "WordInWeb help" });
  await expect(dialog).toBeVisible();
  for (const group of ["Home tools", "Insert tools", "Draw tools", "Layout tools", "Object controls", "Special-case recipes"]) {
    await expect(dialog.getByRole("heading", { name: group, exact: true })).toBeVisible();
  }
  await expect(dialog.getByRole("heading", { name: "California pleading paper", exact: true })).toBeVisible();
  await expect(dialog.getByRole("heading", { name: "Magazine columns with a middle rule", exact: true })).toBeVisible();

  const search = dialog.getByRole("searchbox", { name: "Search help" });
  await search.fill("pleading");
  await expect(dialog.getByRole("heading", { name: "California pleading paper", exact: true })).toBeVisible();
  await expect(dialog.getByText(/Create the fixed 1–28 pleading grid in the repeating header layer/)).toBeVisible();
  await expect(dialog.getByText(/set the exact line height to 24 pt/)).toBeVisible();
  await expect(dialog.getByText(/Home → Line & paragraph spacing → Exactly 24 pt/)).toBeVisible();
  await expect(dialog.getByRole("heading", { name: "Home tools", exact: true })).toHaveCount(0);

  await search.fill("column break");
  await dialog.getByRole("tab", { name: "Shortcuts", exact: true }).click();
  await expect(dialog.getByText("Column break", { exact: true })).toBeVisible();
  await expect(dialog.locator("kbd")).toContainText(process.platform === "darwin" ? "⇧⌘Enter" : "Ctrl+Shift+Enter");

  await page.keyboard.press("Escape");
  await expect(dialog).toHaveCount(0);
  await expect(trigger).toBeFocused();

  await page.keyboard.press(`${MOD}+/`);
  await expect(page.getByRole("dialog", { name: "WordInWeb help" })).toBeVisible();
});

test("the documented modified-Enter shortcut inserts a native column break", async ({ page }) => {
  await load(page);
  await page.getByText("Plain", { exact: true }).click();
  await page.keyboard.press("End");
  await page.keyboard.press(`${MOD}+Shift+Enter`);
  await page.keyboard.type("AFTER-COLUMN-BREAK");

  const pending = page.waitForEvent("download");
  await page.getByText("Download", { exact: true }).click();
  const path = await (await pending).path();
  if (!path) throw new Error("download path unavailable");
  const parts = unzipSync(new Uint8Array(readFileSync(path)));
  const documentXml = strFromU8(parts["word/document.xml"]);
  expect(documentXml).toContain('<w:br w:type="column"');
  expect(documentXml).toContain("AFTER-COLUMN-BREAK");
});

test("Help remains usable in a narrow editor", async ({ page }) => {
  await page.setViewportSize({ width: 640, height: 720 });
  await load(page);
  await page.keyboard.press(`${MOD}+/`);

  const dialog = page.getByRole("dialog", { name: "WordInWeb help" });
  await expect(dialog).toBeVisible();
  await expect(dialog.getByRole("searchbox", { name: "Search help" })).toBeVisible();
  await expect(dialog.getByRole("tab", { name: "Shortcuts", exact: true })).toBeVisible();
  const box = await dialog.boundingBox();
  expect(box).not.toBeNull();
  expect(box!.x).toBeGreaterThanOrEqual(0);
  expect(box!.y).toBeGreaterThanOrEqual(0);
  expect(box!.x + box!.width).toBeLessThanOrEqual(640);
  expect(box!.y + box!.height).toBeLessThanOrEqual(720);
});
