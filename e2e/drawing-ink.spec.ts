import { expect, Page, test } from "@playwright/test";
import { readFileSync } from "node:fs";
import { strFromU8, unzipSync } from "fflate";

function tool(page: Page, title: string) {
  return page.locator(`[title=${JSON.stringify(title)}], [data-tip=${JSON.stringify(title)}]`).first();
}

test("freehand pen creates selectable DrawingML ink with undo, redo, and save-back", async ({ page }) => {
  await page.goto("/?doc=/fixtures/benchmark.docx");
  await page.waitForSelector(".dxw-page span");
  await page.getByRole("button", { name: "draw", exact: true }).click();
  await tool(page, "Draw with pen").click();

  const viewer = page.locator("[data-dxw-drawing-tool=pen]");
  await expect(viewer).toBeVisible();
  const before = await page.locator("[data-dxw-drawing]").count();
  const surface = page.locator(".dxw-page").first().locator(":scope > *").first();
  const box = await surface.boundingBox();
  expect(box).not.toBeNull();
  const points = [
    [box!.x + 430, box!.y + 300],
    [box!.x + 455, box!.y + 325],
    [box!.x + 480, box!.y + 295],
    [box!.x + 505, box!.y + 330],
  ] as const;
  await page.mouse.move(...points[0]);
  await page.mouse.down();
  for (const point of points.slice(1)) await page.mouse.move(...point, { steps: 4 });
  await page.mouse.up();
  await expect(page.locator("[data-dxw-drawing]")).toHaveCount(before + 1);

  await tool(page, "Select objects").click();
  await expect(viewer).toHaveCount(0);
  const ink = page.locator("[data-dxw-drawing]").last();
  const inkBox = await ink.boundingBox();
  expect(inkBox).not.toBeNull();
  await page.mouse.click(inkBox!.x + inkBox!.width / 2, inkBox!.y + inkBox!.height / 2);
  await expect(page.locator("[data-dxw-img-handle]")).toHaveCount(8);

  await page.keyboard.press(process.platform === "darwin" ? "Meta+z" : "Control+z");
  await expect(page.locator("[data-dxw-drawing]")).toHaveCount(before);
  await page.keyboard.press(process.platform === "darwin" ? "Meta+Shift+z" : "Control+y");
  await expect(page.locator("[data-dxw-drawing]")).toHaveCount(before + 1);

  const pending = page.waitForEvent("download");
  await page.getByText("Download", { exact: true }).click();
  const path = await (await pending).path();
  const xml = strFromU8(unzipSync(new Uint8Array(readFileSync(path!)))["word/document.xml"]);
  expect(xml).toContain("<a:custGeom>");
  expect(xml).toContain("<a:pathLst>");
  expect(xml).toContain("<a:moveTo>");
  expect(xml).toContain("<a:lnTo>");
});

test("freehand ink drawn in white space below a table stays at the gesture", async ({ page }) => {
  await page.goto("/?doc=/fixtures/preset-tables.docx");
  await page.waitForSelector(".dxw-page span");
  await page.getByRole("button", { name: "draw", exact: true }).click();
  await tool(page, "Draw with pen").click();

  const surface = page.locator(".dxw-page").first().locator(":scope > *").first();
  await page.locator(".dxw-pages").locator("..").evaluate((element) => {
    element.scrollTop = element.scrollHeight - element.clientHeight;
  });
  const box = (await surface.boundingBox())!;
  const start = { x: box.x + box.width * 0.58, y: box.y + box.height - 150 };
  const end = { x: start.x + 72, y: start.y + 24 };
  const before = await page.locator("[data-dxw-drawing]").count();
  await page.mouse.move(start.x, start.y);
  await page.mouse.down();
  await page.mouse.move(end.x, end.y, { steps: 8 });
  await page.mouse.up();

  await expect(page.locator("[data-dxw-drawing]")).toHaveCount(before + 1);
  const ink = (await page.locator("[data-dxw-ink]").last().boundingBox())!;
  expect(ink.x).toBeCloseTo(start.x, -1);
  expect(ink.y).toBeCloseTo(start.y, -1);
  expect(ink.y + ink.height).toBeLessThan(box.y + box.height);
});

test("stroke eraser removes ink with undo, redo, and save-back", async ({ page }) => {
  await page.goto("/?doc=/fixtures/benchmark.docx");
  await page.waitForSelector(".dxw-page span");
  await page.getByRole("button", { name: "draw", exact: true }).click();
  await tool(page, "Draw with pen").click();

  const drawings = page.locator("[data-dxw-drawing]");
  const before = await drawings.count();
  const surface = page.locator(".dxw-page").first().locator(":scope > *").first();
  const pageBox = await surface.boundingBox();
  expect(pageBox).not.toBeNull();
  await page.mouse.move(pageBox!.x + 430, pageBox!.y + 380);
  await page.mouse.down();
  await page.mouse.move(pageBox!.x + 500, pageBox!.y + 410, { steps: 8 });
  await page.mouse.up();
  await expect(drawings).toHaveCount(before + 1);

  const ink = page.locator("[data-dxw-ink]").last();
  const inkBox = await ink.boundingBox();
  expect(inkBox).not.toBeNull();
  await tool(page, "Stroke eraser").click();
  await expect(page.locator("[data-dxw-drawing-tool=eraser]")).toBeVisible();
  await page.mouse.move(inkBox!.x + inkBox!.width / 2, inkBox!.y + inkBox!.height / 2);
  await page.mouse.down();
  await page.mouse.move(inkBox!.x + inkBox!.width, inkBox!.y + inkBox!.height / 2, { steps: 4 });
  await page.mouse.up();
  await expect(drawings).toHaveCount(before);

  await page.keyboard.press(process.platform === "darwin" ? "Meta+z" : "Control+z");
  await expect(drawings).toHaveCount(before + 1);
  await page.keyboard.press(process.platform === "darwin" ? "Meta+Shift+z" : "Control+y");
  await expect(drawings).toHaveCount(before);

  const pending = page.waitForEvent("download");
  await page.getByText("Download", { exact: true }).click();
  const path = await (await pending).path();
  const xml = strFromU8(unzipSync(new Uint8Array(readFileSync(path!)))["word/document.xml"]);
  expect(xml).not.toContain("WordInWeb ink");
});

test("lasso selects, moves, and deletes multiple ink strokes", async ({ page }) => {
  await page.goto("/?doc=/fixtures/benchmark.docx");
  await page.waitForSelector(".dxw-page span");
  await page.getByRole("button", { name: "draw", exact: true }).click();
  await tool(page, "Draw with pen").click();

  const drawings = page.locator("[data-dxw-drawing]");
  const before = await drawings.count();
  const surface = page.locator(".dxw-page").first().locator(":scope > *").first();
  const pageBox = await surface.boundingBox();
  expect(pageBox).not.toBeNull();
  for (const offset of [0, 90]) {
    await page.mouse.move(pageBox!.x + 380 + offset, pageBox!.y + 460);
    await page.mouse.down();
    await page.mouse.move(pageBox!.x + 440 + offset, pageBox!.y + 485, { steps: 6 });
    await page.mouse.up();
  }
  await expect(drawings).toHaveCount(before + 2);

  const inks = page.locator("[data-dxw-ink]");
  const first = await inks.nth((await inks.count()) - 2).boundingBox();
  const second = await inks.last().boundingBox();
  expect(first).not.toBeNull();
  expect(second).not.toBeNull();
  const left = Math.min(first!.x, second!.x) - 12;
  const top = Math.min(first!.y, second!.y) - 12;
  const right = Math.max(first!.x + first!.width, second!.x + second!.width) + 12;
  const bottom = Math.max(first!.y + first!.height, second!.y + second!.height) + 12;

  await tool(page, "Lasso ink").click();
  await expect(page.locator("[data-dxw-drawing-tool=lasso]")).toBeVisible();
  await page.mouse.move(left, top);
  await page.mouse.down();
  for (const point of [[right, top], [right, bottom], [left, bottom], [left, top]] as const) {
    await page.mouse.move(...point, { steps: 5 });
  }
  await page.mouse.up();

  const selection = page.locator("[data-dxw-ink-selection]");
  await expect(selection).toBeVisible();
  const selectionBox = await selection.boundingBox();
  expect(selectionBox).not.toBeNull();
  await page.mouse.move(selectionBox!.x + selectionBox!.width / 2, selectionBox!.y + selectionBox!.height / 2);
  await page.mouse.down();
  await page.mouse.move(selectionBox!.x + selectionBox!.width / 2 + 40, selectionBox!.y + selectionBox!.height / 2 + 20, { steps: 5 });
  await page.mouse.up();
  const movedFirst = await inks.nth((await inks.count()) - 2).boundingBox();
  const movedSecond = await inks.last().boundingBox();
  expect(movedFirst!.x - first!.x).toBeCloseTo(40, 0);
  expect(movedSecond!.x - second!.x).toBeCloseTo(40, 0);

  await page.keyboard.press("Delete");
  await expect(drawings).toHaveCount(before);
  await page.keyboard.press(process.platform === "darwin" ? "Meta+z" : "Control+z");
  await expect(drawings).toHaveCount(before + 2);
  await page.keyboard.press(process.platform === "darwin" ? "Meta+Shift+z" : "Control+y");
  await expect(drawings).toHaveCount(before);

  const pending = page.waitForEvent("download");
  await page.getByText("Download", { exact: true }).click();
  const path = await (await pending).path();
  const xml = strFromU8(unzipSync(new Uint8Array(readFileSync(path!)))["word/document.xml"]);
  expect(xml).not.toContain("WordInWeb ink");
});
