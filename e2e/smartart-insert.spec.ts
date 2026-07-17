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

async function submitSmartArt(page: Page, layout: "process" | "cycle" | "hierarchy" | "list", items: string): Promise<void> {
  await page.getByRole("button", { name: "SmartArt", exact: true }).click();
  await page.getByLabel("SmartArt layout").selectOption(layout);
  await page.getByLabel("SmartArt items").fill(items);
  await page.getByRole("button", { name: "Insert or update SmartArt" }).click();
}

test("advanced Insert creates and edits native SmartArt with undo, redo, and save", async ({ page }) => {
  await openInsert(page);
  await submitSmartArt(page, "process", "Discover\nDesign\nDeliver");
  await expect(page.getByText("Discover", { exact: true })).toBeVisible();
  await expect(page.getByText("Deliver", { exact: true })).toBeVisible();

  const smartArt = page.locator("[data-dxw-smart-art]").last();
  const smartArtBox = await smartArt.boundingBox();
  expect(smartArtBox).not.toBeNull();
  await page.mouse.click(smartArtBox!.x + 6, smartArtBox!.y + 6);
  await expect(page.locator("[data-dxw-img-handle]")).toHaveCount(8);
  await submitSmartArt(page, "hierarchy", "Lead\nPlan\nBuild\nTest");
  await expect(page.getByText("Lead", { exact: true })).toBeVisible();
  await expect(page.getByText("Test", { exact: true })).toBeVisible();

  await page.keyboard.press(`${MOD}+z`);
  await expect(page.getByText("Discover", { exact: true })).toBeVisible();
  await page.keyboard.press(`${MOD}+Shift+z`);
  await expect(page.getByText("Lead", { exact: true })).toBeVisible();

  const pending = page.waitForEvent("download");
  await page.getByText("Download", { exact: true }).click();
  const path = await (await pending).path();
  expect(path).not.toBeNull();
  const files = unzipSync(new Uint8Array(readFileSync(path!)));
  expect(strFromU8(files["word/document.xml"])).toContain("<dgm:relIds");
  expect(strFromU8(files["word/diagrams/data1.xml"])).toContain("Lead");
  expect(strFromU8(files["word/diagrams/layout1.xml"])).toContain("smartart:hierarchy");
  expect(strFromU8(files["word/diagrams/quickStyle1.xml"])).toContain("<dgm:styleDef");
  expect(strFromU8(files["word/diagrams/colors1.xml"])).toContain("<dgm:colorsDef");
  expect(strFromU8(files["word/diagrams/drawing1.xml"])).toContain("<dsp:drawing");
  expect(strFromU8(files["word/_rels/document.xml.rels"])).toContain("relationships/diagramData");
  expect(strFromU8(files["word/_rels/document.xml.rels"])).toContain("relationships/diagramDrawing");
});
