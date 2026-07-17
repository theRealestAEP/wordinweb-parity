import { expect, Page, test } from "@playwright/test";
import { readFileSync } from "node:fs";
import { strFromU8, unzipSync } from "fflate";

function tool(page: Page, title: string) {
  return page.locator(`[title=${JSON.stringify(title)}], [data-tip=${JSON.stringify(title)}]`).first();
}

async function insertSelectedShape(page: Page, text: string) {
  await page.goto("/?doc=/fixtures/benchmark.docx");
  await page.waitForSelector(".dxw-page span");
  await page.getByRole("button", { name: "insert", exact: true }).click();
  const target = page.locator(".dxw-page span", { hasText: "classic." }).first();
  await target.click();
  await page.keyboard.press("End");
  await tool(page, "Insert shape").click();
  await page.getByLabel("Shape text").fill(text);
  await page.getByTitle("Insert Rounded rectangle").click();
  const hit = page.locator("[data-dxw-drawing]").last();
  const box = await hit.boundingBox();
  expect(box).not.toBeNull();
  await page.mouse.click(box!.x + box!.width - 12, box!.y + box!.height - 12);
  await expect(page.locator("[data-dxw-img-handle]")).toHaveCount(8);
  return hit;
}

test("selected objects nudge, align, rotate, reorder, undo, and save back", async ({ page }) => {
  const hit = await insertSelectedShape(page, "Arrange me");
  const before = await hit.boundingBox();
  expect(before).not.toBeNull();

  await page.keyboard.press("ArrowRight");
  await page.keyboard.press("Shift+ArrowDown");
  const nudged = await page.locator("[data-dxw-drawing]").last().boundingBox();
  expect(nudged!.x - before!.x).toBeCloseTo(1, 0);
  expect(nudged!.y - before!.y).toBeCloseTo(10, 0);

  await page.keyboard.press(process.platform === "darwin" ? "Meta+z" : "Control+z");
  const currentShape = page.locator("[data-dxw-drawing]").last();
  const undoDown = await currentShape.boundingBox();
  expect(undoDown!.y).toBeCloseTo(before!.y, 0);
  await page.mouse.click(undoDown!.x + undoDown!.width - 12, undoDown!.y + undoDown!.height - 12);
  await expect(page.locator("[data-dxw-img-handle]")).toHaveCount(8);

  await page.getByRole("button", { name: "layout", exact: true }).click();
  await tool(page, "Align selected object to page").selectOption("alignRight");
  const aligned = await page.locator("[data-dxw-drawing]").last().boundingBox();
  const pageBox = await page.locator(".dxw-page").first().boundingBox();
  expect(aligned!.x + aligned!.width).toBeCloseTo(pageBox!.x + pageBox!.width, 0);

  const widthBeforeRotate = aligned!.width;
  const heightBeforeRotate = aligned!.height;
  await tool(page, "Rotate selected object").selectOption("rotateRight");
  const rotated = await page.locator("[data-dxw-drawing]").last().boundingBox();
  expect(rotated!.width).toBeCloseTo(heightBeforeRotate, 0);
  expect(rotated!.height).toBeCloseTo(widthBeforeRotate, 0);

  await tool(page, "Change selected object stacking order").selectOption("bringToFront");
  const pending = page.waitForEvent("download");
  await page.getByText("Download", { exact: true }).click();
  const path = await (await pending).path();
  const xml = strFromU8(unzipSync(new Uint8Array(readFileSync(path!)))["word/document.xml"]);
  expect(xml).toContain('relativeFrom="page"');
  expect(xml).toContain('rot="5400000"');
  expect(xml).toMatch(/relativeHeight="\d+"/);
});

test("lassoed ink nudges as one object selection", async ({ page }) => {
  await page.goto("/?doc=/fixtures/benchmark.docx");
  await page.waitForSelector(".dxw-page span");
  await page.getByRole("button", { name: "draw", exact: true }).click();
  await tool(page, "Draw with pen").click();
  const surface = page.locator(".dxw-page").first().locator(":scope > *").first();
  const pageBox = await surface.boundingBox();
  expect(pageBox).not.toBeNull();
  for (const offset of [0, 80]) {
    await page.mouse.move(pageBox!.x + 360 + offset, pageBox!.y + 420);
    await page.mouse.down();
    await page.mouse.move(pageBox!.x + 410 + offset, pageBox!.y + 445, { steps: 5 });
    await page.mouse.up();
  }
  const inks = page.locator("[data-dxw-ink]");
  const count = await inks.count();
  const first = await inks.nth(count - 2).boundingBox();
  const second = await inks.last().boundingBox();
  const left = Math.min(first!.x, second!.x) - 10;
  const top = Math.min(first!.y, second!.y) - 10;
  const right = Math.max(first!.x + first!.width, second!.x + second!.width) + 10;
  const bottom = Math.max(first!.y + first!.height, second!.y + second!.height) + 10;
  await tool(page, "Lasso ink").click();
  await page.mouse.move(left, top);
  await page.mouse.down();
  for (const point of [[right, top], [right, bottom], [left, bottom], [left, top]] as const) await page.mouse.move(...point, { steps: 4 });
  await page.mouse.up();
  await expect(page.locator("[data-dxw-ink-selection]")).toBeVisible();

  await page.keyboard.press("Shift+ArrowRight");
  const movedFirst = await inks.nth(count - 2).boundingBox();
  const movedSecond = await inks.last().boundingBox();
  expect(movedFirst!.x - first!.x).toBeCloseTo(10, 0);
  expect(movedSecond!.x - second!.x).toBeCloseTo(10, 0);
});
