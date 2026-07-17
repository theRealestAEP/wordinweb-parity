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
  const before = await hit.boundingBox();
  expect(before).not.toBeNull();
  await page.mouse.click(before!.x + before!.width - 12, before!.y + before!.height - 12);
  await expect(page.locator("[data-dxw-img-handle]")).toHaveCount(8);

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
