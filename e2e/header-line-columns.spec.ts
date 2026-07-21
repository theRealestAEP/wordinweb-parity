import { expect, Page, test } from "@playwright/test";
import { readFileSync } from "node:fs";
import { strFromU8, unzipSync } from "fflate";

function tool(page: Page, title: string) {
  return page.locator(`[title=${JSON.stringify(title)}], [data-tip=${JSON.stringify(title)}]`).first();
}

async function load(page: Page, fixture = "parity-text"): Promise<void> {
  await page.setViewportSize({ width: 1400, height: 900 });
  await page.goto(`/?doc=/fixtures/${fixture}.docx`);
  await expect(page.locator(".dxw-page").first()).toBeVisible();
}

async function insertLineFromToolbar(page: Page): Promise<void> {
  const hotbar = page.locator("[data-dxw-hf-hotbar]");
  await expect(hotbar.getByRole("button", { name: "Line", exact: true })).toHaveCount(0);
  await tool(page, "Insert shape").click();
  await page.getByTitle("Insert Line").click();
}

async function insertRotatedHeaderLine(page: Page): Promise<void> {
  await page.getByRole("button", { name: "insert", exact: true }).click();
  await tool(page, "Edit header").click();
  const hotbar = page.locator("[data-dxw-hf-hotbar]");
  await expect(hotbar).toBeVisible();
  await insertLineFromToolbar(page);
  await page.getByRole("button", { name: "Rotate", exact: true }).click();
  await page.getByRole("textbox", { name: "Degrees clockwise" }).fill("90");
  await page.getByRole("button", { name: "Apply", exact: true }).click();
  await expect(page.locator('[data-dxw-drawing][data-dxw-hf="1"]')).toHaveCount(1);
}

test("a rotated header line repeats after a page split while its hotbar stays above the page", async ({ page }) => {
  await load(page);
  await expect(page.locator("[data-dxw-editor-context-row]")).toHaveCount(0);
  await insertRotatedHeaderLine(page);

  const hotbar = page.locator("[data-dxw-hf-hotbar]");
  const hotbarBox = (await hotbar.boundingBox())!;
  const firstPageBox = (await page.locator(".dxw-page").first().boundingBox())!;
  expect(hotbarBox.y + hotbarBox.height).toBeLessThanOrEqual(firstPageBox.y + 1);

  await hotbar.getByRole("button", { name: "Close", exact: true }).click();
  await page.getByText("Plain", { exact: true }).click();
  await page.getByRole("button", { name: "insert", exact: true }).click();
  await tool(page, "Insert blank page").click();
  const pages = page.locator(".dxw-page");
  await expect(pages).toHaveCount(3);
  await expect(page.locator('[data-dxw-drawing][data-dxw-hf="1"]')).toHaveCount(3);
});

test("a pleading-paper header line can be dragged horizontally and rotated without covering the page", async ({ page }) => {
  await load(page, "pleading-anon");
  await page.getByRole("button", { name: "insert", exact: true }).click();
  await tool(page, "Edit header").click();

  const contextRow = page.locator("[data-dxw-editor-context-row]");
  const hotbar = page.locator("[data-dxw-hf-hotbar]");
  await expect(contextRow).toBeVisible();
  await expect(hotbar).toBeVisible();
  const hotbarBox = (await hotbar.boundingBox())!;
  const firstPageBox = (await page.locator(".dxw-page").first().boundingBox())!;
  expect(hotbarBox.y + hotbarBox.height).toBeLessThanOrEqual(firstPageBox.y);

  await insertLineFromToolbar(page);
  await expect(page.getByRole("button", { name: "Rotate", exact: true })).toBeVisible();
  const line = page.locator("[data-dxw-object-selection]");
  const before = (await line.boundingBox())!;
  const move = page.locator("[data-dxw-object-move]");
  await expect(move).toBeVisible();
  const moveBox = (await move.boundingBox())!;
  await page.mouse.move(moveBox.x + moveBox.width / 2, moveBox.y + moveBox.height / 2);
  await page.mouse.down();
  await page.mouse.move(moveBox.x + moveBox.width / 2 + 80, moveBox.y + moveBox.height / 2, { steps: 5 });
  await page.mouse.up();
  const moved = (await line.boundingBox())!;
  expect(moved.x).toBeGreaterThan(before.x + 60);

  await page.getByRole("button", { name: "Outline", exact: true }).click();
  await page.getByRole("textbox", { name: "Color and width in pixels" }).fill("#FF0000, 4");
  await page.getByRole("button", { name: "Apply", exact: true }).click();
  await page.getByRole("button", { name: "Rotate", exact: true }).click();
  await page.getByRole("textbox", { name: "Degrees clockwise" }).fill("90");
  await page.getByRole("button", { name: "Apply", exact: true }).click();
  const rotated = (await line.boundingBox())!;
  expect(rotated.height).toBeGreaterThan(rotated.width);

  const pending = page.waitForEvent("download");
  await page.getByText("Download", { exact: true }).click();
  const path = await (await pending).path();
  if (!path) throw new Error("download path unavailable");
  const parts = unzipSync(new Uint8Array(readFileSync(path)));
  const headerXml = Object.entries(parts)
    .filter(([name]) => /^word\/header\d+\.xml$/.test(name))
    .map(([, bytes]) => strFromU8(bytes))
    .join("\n");
  expect(headerXml).toContain('<a:ln w="38100"');
  expect(headerXml).toContain('<a:srgbClr val="FF0000"');
});

test("page two pleading-paper whitespace accepts a clicked caret and typed text", async ({ page }) => {
  await load(page, "pleading-anon");
  const secondPage = page.locator(".dxw-page").nth(1);
  await secondPage.scrollIntoViewIfNeeded();
  const box = (await secondPage.boundingBox())!;
  const bodyTop = Number(await secondPage.getAttribute("data-body-top"));
  const bodyBottom = Number(await secondPage.getAttribute("data-body-bottom"));
  const y = box.y + bodyTop + (bodyBottom - bodyTop) * 0.6;
  await page.mouse.click(box.x + box.width * 0.6, y);
  await expect(page.locator("[data-dxw-caret]")).toBeVisible();
  await page.keyboard.type("PLEADING PAGE TWO");
  await expect(secondPage).toContainText("PLEADING PAGE TWO");
  const typedCaret = (await page.locator("[data-dxw-caret]").boundingBox())!;
  await page.keyboard.press("ArrowUp");
  const previousLineCaret = (await page.locator("[data-dxw-caret]").boundingBox())!;
  expect(typedCaret.y - previousLineCaret.y).toBeGreaterThan(10);
  expect(typedCaret.y - previousLineCaret.y).toBeLessThan(50);
  await page.keyboard.type("PREVIOUS LINE");
  await expect(secondPage).toContainText("PREVIOUS LINE");
});

test("a footer line repeats, stays isolated from body typing, and survives reopen", async ({ page }) => {
  await load(page);
  await page.getByText("Plain", { exact: true }).click();
  await page.getByRole("button", { name: "insert", exact: true }).click();
  await tool(page, "Edit footer").click();
  const hotbar = page.locator("[data-dxw-hf-hotbar]");
  await insertLineFromToolbar(page);
  await page.getByRole("button", { name: "Rotate", exact: true }).click();
  await page.getByRole("textbox", { name: "Degrees clockwise" }).fill("90");
  await page.getByRole("button", { name: "Apply", exact: true }).click();
  await hotbar.getByRole("button", { name: "Close", exact: true }).click();

  await page.getByText("Plain", { exact: true }).click();
  await page.keyboard.press("End");
  await page.keyboard.type(" BODY AFTER FOOTER");
  await expect(page.locator(".dxw-pages")).toContainText("BODY AFTER FOOTER");
  await tool(page, "Insert blank page").click();
  await expect(page.locator(".dxw-page")).toHaveCount(3);
  await expect(page.locator('[data-dxw-drawing][data-dxw-hf="1"]')).toHaveCount(3);

  const pending = page.waitForEvent("download");
  await page.getByText("Download", { exact: true }).click();
  const path = await (await pending).path();
  if (!path) throw new Error("download path unavailable");
  const parts = unzipSync(new Uint8Array(readFileSync(path)));
  const documentXml = strFromU8(parts["word/document.xml"]);
  const footerXml = Object.entries(parts)
    .filter(([name]) => /^word\/footer\d+\.xml$/.test(name))
    .map(([, bytes]) => strFromU8(bytes))
    .join("\n");
  expect(documentXml).toContain("BODY AFTER FOOTER");
  expect(footerXml).toContain('rot="5400000"');

  await page.locator("#docx-upload").setInputFiles(path);
  await expect(page.locator(".dxw-page")).toHaveCount(3);
  await expect(page.locator('[data-dxw-drawing][data-dxw-hf="1"]')).toHaveCount(3);
  await expect(page.locator(".dxw-pages")).toContainText("BODY AFTER FOOTER");
});

test("clicking right of a rotated header line types in the requested native column", async ({ page }) => {
  await load(page);
  await insertRotatedHeaderLine(page);
  await page.locator("[data-dxw-hf-hotbar]").getByRole("button", { name: "Close", exact: true }).click();

  await page.getByRole("button", { name: "layout", exact: true }).click();
  await expect(page.getByRole("button", { name: "Columns & divider", exact: true })).toBeVisible();
  await page.locator('[data-dxw-layout-menu-trigger="columns"]').click();
  await page.locator('[data-dxw-layout-option="2-divider"]').click();

  const firstPage = page.locator(".dxw-page").first();
  const pageBox = (await firstPage.boundingBox())!;
  const bodyTop = await firstPage.evaluate((element) => Number((element as HTMLElement).dataset.bodyTop));
  await page.mouse.click(pageBox.x + pageBox.width * 0.72, pageBox.y + bodyTop + 90);
  await page.keyboard.type("RIGHTCOL");

  const marker = page.locator(".dxw-page span").filter({ hasText: "RIGHTCOL" }).last();
  await expect(marker).toBeVisible();
  const markerBox = (await marker.boundingBox())!;
  expect(markerBox.x).toBeGreaterThan(pageBox.x + pageBox.width / 2);
  await expect(page.locator("[data-dxw-img-handle]")).toHaveCount(0);

  const pending = page.waitForEvent("download");
  await page.getByText("Download", { exact: true }).click();
  const path = await (await pending).path();
  if (!path) throw new Error("download path unavailable");
  const parts = unzipSync(new Uint8Array(readFileSync(path)));
  const documentXml = strFromU8(parts["word/document.xml"]);
  const headerXml = Object.entries(parts)
    .filter(([name]) => /^word\/header\d+\.xml$/.test(name))
    .map(([, bytes]) => strFromU8(bytes))
    .join("\n");
  expect(documentXml).toMatch(/<w:cols\b[^>]*w:num="2"[^>]*w:sep="1"/);
  expect(documentXml).toContain('<w:br w:type="column"');
  expect(headerXml).toContain('rot="5400000"');
});
