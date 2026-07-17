import { expect, Page, test } from "@playwright/test";
import { readFileSync } from "node:fs";
import { strFromU8, unzipSync } from "fflate";

function tool(page: Page, title: string) {
  return page.locator(`[title=${JSON.stringify(title)}], [data-tip=${JSON.stringify(title)}]`).first();
}

test("1026px Insert ribbon keeps every tool reachable without horizontal bleed", async ({ page }) => {
  await page.setViewportSize({ width: 1026, height: 900 });
  await page.goto("/?doc=/fixtures/benchmark.docx");
  await page.waitForSelector(".dxw-page span");
  const toolbar = page.locator('[data-dxw-toolbar-mode="advanced"]');
  await toolbar.getByRole("button", { name: "insert", exact: true }).click();

  await expect(toolbar.getByTitle("Insert image")).toBeVisible();
  await expect(toolbar.getByTitle("Capture and insert a screen, window, or tab")).toBeVisible();
  await toolbar.getByTitle("More tools").click();
  const overflow = toolbar.locator("[data-dxw-overflow-menu]");
  await expect(overflow).toBeVisible();
  for (const title of [
    "Insert a GLB 3D model",
    "Insert or edit SmartArt",
    "Insert or edit chart",
    "Insert online video",
    "Insert shape",
    "Insert text box",
    "Insert WordArt",
    "Insert link",
    "Add comment (select text first)",
    "Insert footnote (at the caret)",
    "Insert bookmark",
    "Insert cross-reference",
    "Edit header",
    "Edit footer",
    "Insert a dynamic page number at the caret",
    "Insert a page, column or section break at the caret",
    "Insert an automatically updating date or time",
    "Insert a Word field",
    "Insert equation",
    "Insert advanced symbol",
    "Drop cap",
    "Embed a file in this document",
  ]) await expect(tool(page, title)).toBeVisible();

  await tool(page, "Insert or edit chart").click();
  const chart = page.getByLabel("Chart type").locator("..");
  await expect(chart).toBeVisible();
  const chartBox = await chart.boundingBox();
  expect(chartBox).not.toBeNull();
  expect(chartBox!.x).toBeGreaterThanOrEqual(0);
  expect(chartBox!.x + chartBox!.width).toBeLessThanOrEqual(1026);

  await tool(page, "Insert or edit chart").click({ force: true });
  await expect(chart).toBeHidden();
  await tool(page, "Insert WordArt").click();
  const wordArt = page.getByLabel("WordArt text").locator("..");
  await expect(wordArt).toBeVisible();
  const wordArtBox = await wordArt.boundingBox();
  expect(wordArtBox).not.toBeNull();
  expect(wordArtBox!.x).toBeGreaterThanOrEqual(0);
  expect(wordArtBox!.x + wordArtBox!.width).toBeLessThanOrEqual(1026);

  expect(await page.evaluate(() => ({ scroll: document.documentElement.scrollWidth, viewport: innerWidth }))).toEqual({ scroll: 1026, viewport: 1026 });
});

test("Draw keeps semantic tools and cursors distinct, saves repeated sparse strokes, and resets on tab exit", async ({ page }) => {
  await page.goto("/?doc=/fixtures/benchmark.docx");
  await page.waitForSelector(".dxw-page span");
  const toolbar = page.locator('[data-dxw-toolbar-mode="advanced"]');
  await toolbar.getByRole("button", { name: "draw", exact: true }).click();
  const surface = page.locator(".dxw-page").first().locator(":scope > *").first();

  const cursors: string[] = [];
  for (const [title, kind] of [
    ["Draw with pen", "pen"],
    ["Draw with highlighter", "highlighter"],
    ["Stroke eraser", "eraser"],
    ["Lasso ink", "lasso"],
  ] as const) {
    await tool(page, title).click();
    const viewer = page.locator(`[data-dxw-drawing-tool=${kind}]`);
    await expect(viewer).toBeVisible();
    cursors.push(await surface.evaluate((element) => getComputedStyle(element).cursor));
  }
  expect(new Set(cursors).size).toBe(4);

  await tool(page, "Draw with highlighter").click();
  const highlighterBackground = await tool(page, "Draw with highlighter").evaluate((element) => getComputedStyle(element).backgroundColor);
  const penBackground = await tool(page, "Draw with pen").evaluate((element) => getComputedStyle(element).backgroundColor);
  expect(highlighterBackground).not.toBe(penBackground);

  await tool(page, "Draw with pen").click();
  const before = await page.locator("[data-dxw-ink]").count();
  const box = await surface.boundingBox();
  expect(box).not.toBeNull();
  for (let index = 0; index < 4; index++) {
    const x = box!.x + 500 + index * 18;
    const y = box!.y + 610 + index * 22;
    await page.mouse.move(x, y);
    await page.mouse.down();
    await page.mouse.move(x + 42, y + 15);
    await page.mouse.up();
    await expect(page.locator("[data-dxw-ink]")).toHaveCount(before + index + 1);
  }

  await tool(page, "Draw with highlighter").click();
  for (let index = 0; index < 4; index++) {
    const x = box!.x + 180 + index * 22;
    const y = box!.y + 480 + index * 20;
    await page.mouse.move(x, y);
    await page.mouse.down();
    await page.mouse.move(x + 48, y + 10);
    await page.mouse.up();
    await expect(page.locator("[data-dxw-ink]")).toHaveCount(before + 5 + index);
  }
  await expect(page.locator('path[stroke="#f9d949"][stroke-opacity="0.45"]')).toHaveCount(4);

  const pending = page.waitForEvent("download");
  await page.getByText("Download", { exact: true }).click();
  const path = await (await pending).path();
  expect(path).not.toBeNull();
  const xml = strFromU8(unzipSync(new Uint8Array(readFileSync(path!)))["word/document.xml"]);
  expect((xml.match(/WordInWeb ink/g) ?? []).length).toBeGreaterThanOrEqual(8);
  expect((xml.match(/<a:srgbClr val="F9D949"/g) ?? []).length).toBeGreaterThanOrEqual(4);
  expect((xml.match(/<a:alpha val="45000"\/>/g) ?? []).length).toBeGreaterThanOrEqual(4);

  await toolbar.getByRole("button", { name: "home", exact: true }).click();
  await expect(page.locator("[data-dxw-drawing-tool]")).toHaveCount(0);
});

test("upload paints a dedicated busy overlay before a delayed large file read completes", async ({ page }) => {
  test.setTimeout(120_000);
  await page.addInitScript(() => {
    const original = File.prototype.arrayBuffer;
    File.prototype.arrayBuffer = async function () {
      await new Promise((resolve) => setTimeout(resolve, 700));
      return original.call(this);
    };
  });
  await page.goto("/?doc=/fixtures/benchmark.docx");
  await page.waitForSelector(".dxw-page span");
  const buffer = readFileSync("apps/demo/public/fixtures/wild2-legal-nih-contract.docx");
  await page.locator("#docx-upload").setInputFiles({
    name: "wild2-legal-nih-contract.docx",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    buffer,
  });
  const loading = page.locator("[data-dxw-loading]");
  await expect(loading).toBeVisible();
  await expect(loading).toContainText("Loading wild2-legal-nih-contract.docx");
  await expect(loading).toHaveAttribute("aria-busy", "true");
  await expect(loading).toHaveCount(0, { timeout: 60_000 });
  await expect(page.locator(".dxw-page")).toHaveCount(419, { timeout: 60_000 });
  await expect(page.locator(".dxw-page span").first()).toBeVisible();
});
