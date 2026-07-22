import { expect, Page, test } from "@playwright/test";
import { readFileSync } from "node:fs";
import { strFromU8, unzipSync } from "fflate";

function tool(page: Page, title: string) {
  return page.locator(`[title=${JSON.stringify(title)}], [data-tip=${JSON.stringify(title)}]`).first();
}

test("inserted DrawingML shapes have editable text, move/resize handles, undo, and save-back", async ({ page }) => {
  await page.goto("/?doc=/fixtures/benchmark.docx");
  await page.waitForSelector(".dxw-page span");
  await page.getByRole("button", { name: "insert", exact: true }).click();
  const target = page.locator(".dxw-page span", { hasText: "classic." }).first();
  await target.click();
  await page.keyboard.press("End");

  await tool(page, "Insert shape").click();
  await page.getByLabel("Shape text").fill("Drawn shape");
  await page.getByTitle("Insert Rounded rectangle").click();
  await expect(page.locator(".dxw-pages")).toContainText("Drawn shape");

  const hit = page.locator("[data-dxw-drawing]").last();
  let before = await hit.boundingBox();
  expect(before).not.toBeNull();
  await page.mouse.click(before!.x + before!.width - 12, before!.y + before!.height - 12);
  await expect(page.locator("[data-dxw-img-handle]")).toHaveCount(8);

  await expect(page.getByRole("button", { name: "Shape Format", exact: true })).toBeVisible();
  const format = page.locator("[data-dxw-object-format]");
  await expect(format.getByRole("button", { name: "Edit text", exact: true })).toBeVisible();
  await format.getByRole("button", { name: "Fill", exact: true }).click();
  await expect(page.getByLabel("Fill color picker")).toHaveAttribute("type", "color");
  await page.getByRole("textbox", { name: "Fill color", exact: true }).fill("#FF0000");
  await page.getByRole("button", { name: "Apply", exact: true }).click();
  await format.getByRole("button", { name: "Outline", exact: true }).click();
  await page.getByRole("textbox", { name: "Outline color", exact: true }).fill("#00FF00");
  await page.getByRole("spinbutton", { name: "Outline width in pixels" }).fill("3");
  await page.getByRole("button", { name: "Apply", exact: true }).click();
  await format.getByRole("button", { name: "Size", exact: true }).click();
  await page.getByRole("spinbutton", { name: "Width (pixels)" }).fill("240");
  await page.getByRole("spinbutton", { name: "Height (pixels)" }).fill("120");
  await page.getByRole("button", { name: "Apply", exact: true }).click();
  before = await hit.boundingBox();
  expect(before!.width).toBeCloseTo(240, 0);
  expect(before!.height).toBeCloseTo(120, 0);

  await page.getByRole("button", { name: "home", exact: true }).click();
  await page.getByRole("button", { name: "Shape Format", exact: true }).click();
  await page.keyboard.press("Escape");
  await expect(page.getByRole("button", { name: "Shape Format", exact: true })).toHaveCount(0);
  await expect(page.getByRole("button", { name: "home", exact: true })).toHaveAttribute("aria-pressed", "true");
  const shapeAfterDeselect = await hit.boundingBox();
  await page.mouse.click(shapeAfterDeselect!.x + shapeAfterDeselect!.width - 12, shapeAfterDeselect!.y + shapeAfterDeselect!.height - 12);

  await page.mouse.move(before!.x + before!.width - 16, before!.y + before!.height - 16);
  await page.mouse.down();
  await page.mouse.move(before!.x + before!.width + 34, before!.y + before!.height + 24, { steps: 6 });
  await page.mouse.up();
  const moved = await page.locator("[data-dxw-drawing]").last().boundingBox();
  expect(moved).not.toBeNull();
  expect(Math.hypot(moved!.x - before!.x, moved!.y - before!.y)).toBeGreaterThan(15);

  const se = page.locator('[data-dxw-img-handle="se"]');
  const handle = await se.boundingBox();
  expect(handle).not.toBeNull();
  await page.mouse.move(handle!.x + handle!.width / 2, handle!.y + handle!.height / 2);
  await page.mouse.down();
  await page.mouse.move(handle!.x + 45, handle!.y + 25, { steps: 5 });
  await page.mouse.up();
  const resized = await page.locator("[data-dxw-drawing]").last().boundingBox();
  expect(resized!.width).toBeGreaterThan(moved!.width + 20);

  await page.keyboard.press(process.platform === "darwin" ? "Meta+z" : "Control+z");
  const undone = await page.locator("[data-dxw-drawing]").last().boundingBox();
  expect(undone!.width).toBeCloseTo(moved!.width, 0);

  const pending = page.waitForEvent("download");
  await page.getByText("Download", { exact: true }).click();
  const path = await (await pending).path();
  const xml = strFromU8(unzipSync(new Uint8Array(readFileSync(path!)))["word/document.xml"]);
  expect(xml).toContain("<wp:anchor");
  expect(xml).toContain("<wps:wsp");
  expect(xml).toContain('a:prstGeom prst="roundRect"');
  expect(xml).toContain("Drawn shape");
  expect(xml).toContain('<a:srgbClr val="FF0000"');
  expect(xml).toContain('<a:srgbClr val="00FF00"');
  expect(xml).toContain('<a:ln w="28575"');
});

test("line selection exposes only line formatting and preserves position when wrapping", async ({ page }) => {
  await page.goto("/?doc=/fixtures/benchmark.docx");
  await page.waitForSelector(".dxw-page span");
  await page.getByRole("button", { name: "insert", exact: true }).click();
  const target = page.locator(".dxw-page span", { hasText: "classic." }).first();
  await target.click();
  await page.keyboard.press("End");
  await tool(page, "Insert shape").click();
  await page.getByTitle("Insert Line").click();

  await expect(page.getByRole("button", { name: "Line Format", exact: true })).toBeVisible();
  const format = page.locator("[data-dxw-object-format]");
  await expect(format.getByRole("button", { name: "Line style", exact: true })).toBeVisible();
  await expect(format.getByRole("button", { name: "Fill", exact: true })).toHaveCount(0);
  await expect(format.getByRole("button", { name: "Outline", exact: true })).toHaveCount(0);
  await expect(format.getByRole("button", { name: "Edit text", exact: true })).toHaveCount(0);

  await format.getByRole("button", { name: "Line style", exact: true }).click();
  const lineDialog = page.getByRole("dialog", { name: "Line style" });
  await lineDialog.getByLabel("Line style").selectOption("dotted");
  await lineDialog.getByRole("button", { name: "Apply", exact: true }).click();
  await expect(lineDialog).toHaveCount(0);
  const before = await page.locator("[data-dxw-drawing]").last().boundingBox();
  expect(before).not.toBeNull();
  await format.getByRole("button", { name: "Wrap", exact: true }).click();
  await page.getByRole("option", { name: "Square", exact: true }).click();
  const after = await page.locator("[data-dxw-drawing]").last().boundingBox();
  expect(after).not.toBeNull();
  expect(after!.x).toBeCloseTo(before!.x, 0);
  await expect(page.getByRole("button", { name: "Line Format", exact: true })).toBeVisible();

  const pending = page.waitForEvent("download");
  await page.getByText("Download", { exact: true }).click();
  const path = await (await pending).path();
  if (!path) throw new Error("download path unavailable");
  const xml = strFromU8(unzipSync(new Uint8Array(readFileSync(path)))["word/document.xml"]);
  expect(xml).toContain('<a:prstDash val="dot"');
});

test("a floating shape stays visible when moved repeatedly onto a blank page and reopened", async ({ page }) => {
  await page.setViewportSize({ width: 1400, height: 1800 });
  await page.goto("/?doc=/fixtures/benchmark.docx");
  await page.waitForSelector(".dxw-page span");
  await page.getByLabel("Document zoom").selectOption("0.5");
  await page.getByRole("button", { name: "insert", exact: true }).click();
  const target = page.locator(".dxw-page span", { hasText: "classic." }).first();
  await target.click();
  await page.keyboard.press("End");

  await tool(page, "Insert blank page").click();
  await tool(page, "Insert shape").click();
  await page.getByLabel("Shape text").fill("Cross page shape");
  await page.getByTitle("Insert Rounded rectangle").click();

  const shape = page.locator("[data-dxw-drawing]").last();
  const blankPage = page.locator(".dxw-page").nth(1);
  await expect(shape).toBeVisible();
  await expect(blankPage).toBeVisible();

  const initial = (await shape.boundingBox())!;
  const blank = (await blankPage.boundingBox())!;
  const startX = initial.x + initial.width - 12;
  const startY = initial.y + initial.height - 12;
  await page.mouse.move(startX, startY);
  await page.mouse.down();
  await page.mouse.move(
    startX + blank.x + 80 - initial.x,
    startY + blank.y + 110 - initial.y,
    { steps: 12 },
  );
  await page.mouse.up();

  await expect(shape).toBeVisible();
  expect(await shape.evaluate((element) =>
    [...document.querySelectorAll(".dxw-page")].indexOf(element.closest(".dxw-page")!),
  )).toBe(1);
  const firstMove = (await shape.boundingBox())!;

  await page.mouse.move(firstMove.x + firstMove.width - 12, firstMove.y + firstMove.height - 12);
  await page.mouse.down();
  await page.mouse.move(firstMove.x + firstMove.width + 28, firstMove.y + firstMove.height + 18, { steps: 6 });
  await page.mouse.up();
  await expect(shape).toBeVisible();
  const secondMove = (await shape.boundingBox())!;
  expect(secondMove.x).toBeGreaterThan(firstMove.x + 20);
  expect(secondMove.y).toBeGreaterThan(firstMove.y + 10);

  const destinationPage = (await blankPage.boundingBox())!;
  const savedPosition = {
    x: secondMove.x - destinationPage.x,
    y: secondMove.y - destinationPage.y,
  };
  const pending = page.waitForEvent("download");
  await page.getByText("Download", { exact: true }).click();
  const path = await (await pending).path();
  expect(path).not.toBeNull();
  const xml = strFromU8(unzipSync(new Uint8Array(readFileSync(path!)))["word/document.xml"]);
  expect((xml.match(/Cross page shape/g) ?? [])).toHaveLength(1);
  expect(xml).toContain('<wp:positionH relativeFrom="margin">');
  expect(xml).toContain('<wp:positionV relativeFrom="paragraph">');

  await page.locator("#docx-upload").setInputFiles(path!);
  await page.waitForSelector(".dxw-page span");
  const reopened = page.locator("[data-dxw-drawing]").last();
  await expect(reopened).toBeVisible();
  expect(await reopened.evaluate((element) =>
    [...document.querySelectorAll(".dxw-page")].indexOf(element.closest(".dxw-page")!),
  )).toBe(1);
  const reopenedBox = (await reopened.boundingBox())!;
  const reopenedPage = (await page.locator(".dxw-page").nth(1).boundingBox())!;
  expect(reopenedBox.x - reopenedPage.x).toBeCloseTo(savedPosition.x, 0);
  expect(reopenedBox.y - reopenedPage.y).toBeCloseTo(savedPosition.y, 0);
});

test("shape wrap and layering controls remain clickable through mode changes, delete, undo, and reopen", async ({ page }) => {
  await page.goto("/?doc=/fixtures/benchmark.docx");
  await page.waitForSelector(".dxw-page span");
  await page.getByRole("button", { name: "insert", exact: true }).click();
  const target = page.locator(".dxw-page span", { hasText: "classic." }).first();
  await target.click();
  await page.keyboard.press("End");
  await tool(page, "Insert shape").click();
  await page.getByLabel("Shape text").fill("Wrap modes");
  await page.getByTitle("Insert Rectangle").click();

  const shape = page.locator("[data-dxw-drawing]").last();
  const box = (await shape.boundingBox())!;
  await page.mouse.click(box.x + box.width - 12, box.y + box.height - 12);
  const objectBar = page.getByRole("main");
  for (const mode of ["Behind", "In front", "Top+Bottom", "Wrap", "Inline", "Behind"]) {
    await objectBar.getByRole("button", { name: mode, exact: true }).click();
    await expect(objectBar.getByRole("button", { name: mode, exact: true })).toHaveAttribute("aria-pressed", "true");
  }

  await page.locator("[data-dxw-object-format]").getByRole("button", { name: "Delete", exact: true }).click();
  await expect(page.locator(".dxw-pages")).not.toContainText("Wrap modes");
  await page.keyboard.press(process.platform === "darwin" ? "Meta+z" : "Control+z");
  await expect(page.locator(".dxw-pages")).toContainText("Wrap modes");

  const pending = page.waitForEvent("download");
  await page.getByText("Download", { exact: true }).click();
  const path = await (await pending).path();
  if (!path) throw new Error("download path unavailable");
  const xml = strFromU8(unzipSync(new Uint8Array(readFileSync(path)))["word/document.xml"]);
  expect(xml).toContain('behindDoc="1"');
  expect(xml).toContain("<wp:wrapNone");

  await page.locator("#docx-upload").setInputFiles(path);
  await expect(page.locator(".dxw-pages")).toContainText("Wrap modes");
  await expect(page.locator("[data-dxw-drawing]").last()).toBeVisible();
});
