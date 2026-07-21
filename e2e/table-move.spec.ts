import { readFileSync } from "node:fs";
import { expect, Page, test } from "@playwright/test";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

const MOD = process.platform === "darwin" ? "Meta" : "Control";

async function open(page: Page): Promise<void> {
  await page.goto("/?doc=/fixtures/preset-tables.docx");
  await page.waitForSelector(".dxw-page span");
  await page.waitForTimeout(300);
}

function exact(page: Page, text: string) {
  return page.locator(`.dxw-page span:text-is(${JSON.stringify(text)})`).first();
}

function tool(page: Page, title: string) {
  return page.locator(`[title=${JSON.stringify(title)}], [data-tip=${JSON.stringify(title)}]`).first();
}

type Box = { x: number; y: number; width: number; height: number };

const firstTableTexts = ["Region", "Status"];
const secondTableTexts = ["Team", "Due"];

function intersects(a: Box, b: Box): boolean {
  return a.x < b.x + b.width && a.x + a.width > b.x &&
    a.y < b.y + b.height && a.y + a.height > b.y;
}

async function tableBounds(page: Page, texts: string[]): Promise<Box> {
  const boxes = await Promise.all(texts.map(async (text) => {
    const box = await exact(page, text).boundingBox();
    if (!box) throw new Error(`Missing table text ${text}`);
    return box;
  }));
  const left = Math.min(...boxes.map((box) => box.x));
  const top = Math.min(...boxes.map((box) => box.y));
  const right = Math.max(...boxes.map((box) => box.x + box.width));
  const bottom = Math.max(...boxes.map((box) => box.y + box.height));
  return { x: left, y: top, width: right - left, height: bottom - top };
}

async function dragTableBy(page: Page, cellText: string, dx: number, dy: number): Promise<void> {
  await exact(page, cellText).hover();
  const handles = page.locator("[data-dxw-table-move]");
  const visibleIndex = await handles.evaluateAll((elements) =>
    elements.findIndex((element) => (element as HTMLElement).style.opacity === "1"),
  );
  expect(visibleIndex).toBeGreaterThanOrEqual(0);
  const box = await handles.nth(visibleIndex).boundingBox();
  if (!box) throw new Error(`Move handle missing for ${cellText}`);
  const x = box.x + box.width / 2;
  const y = box.y + box.height / 2;
  await page.mouse.move(x, y);
  await page.mouse.down();
  await page.mouse.move(x + dx, y + dy, { steps: 8 });
  await page.mouse.up();
  await page.waitForTimeout(350);
}

async function floatingTableBounds(page: Page, cellText: string): Promise<Box> {
  await exact(page, cellText).hover();
  const handles = page.locator("[data-dxw-table-move]");
  const visibleIndex = await handles.evaluateAll((elements) =>
    elements.findIndex((element) => (element as HTMLElement).style.opacity === "1"),
  );
  expect(visibleIndex).toBeGreaterThanOrEqual(0);
  const handle = await handles.nth(visibleIndex).boundingBox();
  if (!handle) throw new Error(`Move handle missing for ${cellText}`);
  const left = handle.x + 24;
  const top = handle.y + 24;
  const rows = await page.locator('[data-dxw-item-kind="grip"][style*="row-resize"]').evaluateAll(
    (elements, expectedLeft) => elements.map((element) => {
      const rect = element.getBoundingClientRect();
      return { x: rect.x, y: rect.y, width: rect.width, height: rect.height };
    }).filter((rect) => Math.abs(rect.x - expectedLeft) <= 1),
    left,
  );
  if (rows.length === 0) throw new Error(`Row grips missing for ${cellText}`);
  const width = Math.max(...rows.map((row) => row.width));
  const bottom = Math.max(...rows.map((row) => row.y + row.height / 2));
  return { x: left, y: top, width, height: bottom - top };
}

async function downloadAndReopen(page: Page): Promise<string> {
  const pending = page.waitForEvent("download");
  await page.getByText("Download", { exact: true }).click();
  const path = await (await pending).path();
  if (!path) throw new Error("Downloaded document is missing");
  await page.locator("#docx-upload").setInputFiles(path);
  await page.waitForSelector(".dxw-page span");
  await page.waitForTimeout(350);
  return path;
}

async function openWithSecondTablePolicy(page: Page, policy: "allow" | "never"): Promise<void> {
  if (policy === "allow") {
    await open(page);
    return;
  }
  const files = unzipSync(new Uint8Array(readFileSync("apps/demo/public/fixtures/preset-tables.docx")));
  let table = 0;
  const xml = strFromU8(files["word/document.xml"]).replace(/<w:tblPr>/g, (tag) => {
    table += 1;
    return table === 2 ? `${tag}<w:tblOverlap w:val="never"/>` : tag;
  });
  files["word/document.xml"] = strToU8(xml);
  await page.goto("/");
  await page.locator("#docx-upload").setInputFiles({
    name: "tables-overlap-never.docx",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    buffer: Buffer.from(zipSync(files)),
  });
  await page.waitForSelector(".dxw-page span");
  await page.waitForTimeout(300);
}

async function moveSecondTableOntoFirst(page: Page): Promise<void> {
  await dragTableBy(page, "Region", 24, -48);
  const first = await tableBounds(page, firstTableTexts);
  const second = await tableBounds(page, secondTableTexts);
  await dragTableBy(page, "Team", first.x - second.x, first.y - second.y);
}

async function expectRequestedObjectsWithinPages(page: Page): Promise<void> {
  const drawings = page.locator("[data-dxw-drawing]");
  const drawingCount = await drawings.count();
  expect(drawingCount, "shape and WordArt must be present").toBeGreaterThanOrEqual(2);
  const objects = [
    ["table", page.locator('[data-dxw-role="table-fill"], [data-dxw-role="table-rule"]')],
    ["equation", page.locator("[data-dxw-math]")],
    ["image", page.locator('[data-dxw-item-kind="image"]')],
    ["shape", drawings.nth(drawingCount - 2)],
    ["WordArt", drawings.nth(drawingCount - 1)],
  ] as const;
  for (const [name, locator] of objects) {
    expect(await locator.count(), `${name} must be present`).toBeGreaterThan(0);
    const outside = await locator.evaluateAll((elements) => elements.flatMap((element) => {
      const sheet = element.closest<HTMLElement>(".dxw-page")?.getBoundingClientRect();
      const rect = element.getBoundingClientRect();
      if (!sheet || rect.left < sheet.left - 1 || rect.top < sheet.top - 1 ||
          rect.right > sheet.right + 1 || rect.bottom > sheet.bottom + 1) {
        return [{ left: rect.left, top: rect.top, right: rect.right, bottom: rect.bottom }];
      }
      return [];
    }));
    expect(outside, `${name} crossed its page boundary`).toEqual([]);
  }
}

async function expectTypedBelowTable(
  page: Page,
  marker: string,
  move: () => Promise<void>,
): Promise<void> {
  const lastCell = (await exact(page, "16").boundingBox())!;
  await move();
  await page.keyboard.type(marker);
  await page.waitForTimeout(250);
  const typed = await exact(page, marker).boundingBox();
  expect(typed).not.toBeNull();
  expect(typed!.y).toBeGreaterThan(lastCell.y + lastCell.height);
}

test.describe("Word-style table movement", () => {
  test("the hover handle drags an inline table and saves its page position", async ({ page }) => {
    await open(page);
    const region = exact(page, "Region");
    const before = (await region.boundingBox())!;

    await region.hover();
    const handles = page.locator("[data-dxw-table-move]");
    await expect(handles).toHaveCount(2);
    const visibleIndex = await handles.evaluateAll((elements) =>
      elements.findIndex((element) => (element as HTMLElement).style.opacity === "1"),
    );
    expect(visibleIndex).toBeGreaterThanOrEqual(0);
    const handle = handles.nth(visibleIndex);
    const box = (await handle.boundingBox())!;

    // Crossing from table content toward the corner handle can briefly leave
    // both hit zones. Keep the affordance armed long enough to acquire it.
    await page.mouse.move(before.x + before.width / 2, box.y + box.height / 2);
    await page.waitForTimeout(350);
    await expect(handle).toHaveCSS("opacity", "1");

    await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + box.width / 2 - 25, box.y + box.height / 2 + 70, { steps: 6 });
    await page.mouse.up();
    await page.waitForTimeout(400);

    const after = (await region.boundingBox())!;
    expect(after.x - before.x).toBeCloseTo(-25, 0);
    expect(after.y - before.y).toBeCloseTo(70, 0);

    const downloadPromise = page.waitForEvent("download");
    await page.getByText("Download", { exact: true }).click();
    const path = await (await downloadPromise).path();
    expect(path).not.toBeNull();
    const xml = strFromU8(unzipSync(new Uint8Array(readFileSync(path!)))["word/document.xml"]);
    expect(xml).toContain("<w:tblpPr");
    expect(xml).toContain('w:horzAnchor="page"');
    expect(xml).toContain('w:vertAnchor="page"');
  });

  test("dragging a table from the top of a later page keeps it on that page", async ({ page }) => {
    await page.goto("/?doc=/fixtures/pleading-paper.docx");
    await page.waitForSelector(".dxw-page span");
    await page.waitForTimeout(300);

    const firstPage = page.locator(".dxw-page").first();
    const secondPage = page.locator(".dxw-page").nth(1);
    await secondPage.scrollIntoViewIfNeeded();
    const tableText = secondPage.locator("span").filter({ hasText: "MEWINE" }).first();
    await tableText.hover();
    const handle = secondPage.locator("[data-dxw-table-move]");
    await expect(handle).toHaveCount(1);
    const box = (await handle.boundingBox())!;

    await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + box.width / 2 + 30, box.y + box.height / 2 + 30, { steps: 6 });
    await page.mouse.up();
    await page.waitForTimeout(400);

    await expect(firstPage).not.toContainText("MEWINE");
    await expect(secondPage).toContainText("MEWINE");
    await expect(page.locator(".dxw-page")).toHaveCount(7);

    await downloadAndReopen(page);
    await expect(page.locator(".dxw-page").first()).not.toContainText("MEWINE");
    await expect(page.locator(".dxw-page").nth(1)).toContainText("MEWINE");
    await expect(page.locator(".dxw-page")).toHaveCount(7);
  });

  test("typing below a moved terminal table stays below and outside its cells", async ({ page }) => {
    await open(page);
    const lastCell = exact(page, "16");
    const beforeCell = (await lastCell.boundingBox())!;
    await lastCell.hover();
    const handles = page.locator("[data-dxw-table-move]");
    const visibleIndex = await handles.evaluateAll((elements) =>
      elements.findIndex((element) => (element as HTMLElement).style.opacity === "1"),
    );
    const handle = handles.nth(visibleIndex);
    const beforeHandle = (await handle.boundingBox())!;
    await page.mouse.move(beforeHandle.x + beforeHandle.width / 2, beforeHandle.y + beforeHandle.height / 2);
    await page.mouse.down();
    await page.mouse.move(beforeHandle.x + beforeHandle.width / 2, beforeHandle.y + beforeHandle.height / 2 + 90, { steps: 6 });
    await page.mouse.up();
    await page.waitForTimeout(350);

    const movedCell = (await lastCell.boundingBox())!;
    await page.mouse.click(movedCell.x + movedCell.width / 2, movedCell.y + movedCell.height + 12);
    await page.keyboard.type("~AFTERMOVE");
    await page.waitForTimeout(250);

    await expect(lastCell).toHaveText("16");
    const marker = (await exact(page, "~AFTERMOVE").boundingBox())!;
    expect(marker.y).toBeGreaterThan(movedCell.y + movedCell.height);

    await page.keyboard.press(`${MOD}+z`);
    await expect(exact(page, "~AFTERMOVE")).toHaveCount(0);
    await expect(page.locator(".dxw-page").first()).not.toContainText("~");
    const afterUndoCell = (await lastCell.boundingBox())!;
    expect(afterUndoCell.y).toBeCloseTo(movedCell.y, 0);
    const emptyTops = await page.locator(".dxw-page span").evaluateAll((elements) =>
      elements
        .filter((element) => element.textContent === "")
        .map((element) => element.getBoundingClientRect().top),
    );
    expect(emptyTops.some((top) => top < afterUndoCell.y)).toBe(true);

    await page.keyboard.press(`${MOD}+z`);
    const afterSecondUndoCell = (await lastCell.boundingBox())!;
    expect(afterSecondUndoCell.y).toBeCloseTo(beforeCell.y, 0);
  });

  test("click-and-type below a moved table works outside its horizontal bounds", async ({ page }) => {
    await open(page);
    const lastCell = exact(page, "16");
    await lastCell.hover();
    const handles = page.locator("[data-dxw-table-move]");
    const visibleIndex = await handles.evaluateAll((elements) =>
      elements.findIndex((element) => (element as HTMLElement).style.opacity === "1"),
    );
    const handle = handles.nth(visibleIndex);
    const beforeHandle = (await handle.boundingBox())!;
    await page.mouse.move(beforeHandle.x + beforeHandle.width / 2, beforeHandle.y + beforeHandle.height / 2);
    await page.mouse.down();
    await page.mouse.move(beforeHandle.x + beforeHandle.width / 2, beforeHandle.y + beforeHandle.height / 2 + 90, { steps: 6 });
    await page.mouse.up();
    await page.waitForTimeout(350);

    const movedCell = (await lastCell.boundingBox())!;
    const pageBox = (await page.locator(".dxw-page").first().boundingBox())!;
    await page.mouse.click(pageBox.x + 12, movedCell.y + movedCell.height + 12);
    await page.keyboard.type("OUTSIDE");
    await page.waitForTimeout(250);

    await expect(lastCell).toHaveText("16");
    const marker = (await exact(page, "OUTSIDE").boundingBox())!;
    expect(marker.y).toBeGreaterThan(movedCell.y + movedCell.height);
  });

  test("later floating tables overlap on top and keep that ordering after reopen", async ({ page }) => {
    await openWithSecondTablePolicy(page, "allow");
    await moveSecondTableOntoFirst(page);

    const first = await tableBounds(page, firstTableTexts);
    const second = await tableBounds(page, secondTableTexts);
    expect(intersects(first, second)).toBe(true);
    const team = (await exact(page, "Team").boundingBox())!;
    const topTableText = await page.evaluate(({ x, y }) =>
      document.elementsFromPoint(x, y)
        .find((element) => element.textContent === "Team" || element.textContent === "Region")
        ?.textContent,
    { x: team.x + team.width / 2, y: team.y + team.height / 2 });
    expect(topTableText).toBe("Team");

    const path = await downloadAndReopen(page);
    const xml = strFromU8(unzipSync(new Uint8Array(readFileSync(path)))["word/document.xml"]);
    expect(xml.match(/<w:tblpPr\b/g)).toHaveLength(2);
    expect(xml).not.toContain("<w:tblOverlap");
    const reopenedFirst = await tableBounds(page, firstTableTexts);
    const reopenedSecond = await tableBounds(page, secondTableTexts);
    expect(intersects(reopenedFirst, reopenedSecond)).toBe(true);
    const reopenedTeam = (await exact(page, "Team").boundingBox())!;
    const reopenedTop = await page.evaluate(({ x, y }) =>
      document.elementsFromPoint(x, y)
        .find((element) => element.textContent === "Team" || element.textContent === "Region")
        ?.textContent,
    { x: reopenedTeam.x + reopenedTeam.width / 2, y: reopenedTeam.y + reopenedTeam.height / 2 });
    expect(reopenedTop).toBe("Team");
  });

  test("tblOverlap never keeps an attempted overlap separated after reopen", async ({ page }) => {
    await openWithSecondTablePolicy(page, "never");
    await moveSecondTableOntoFirst(page);

    const first = await tableBounds(page, firstTableTexts);
    const second = await tableBounds(page, secondTableTexts);
    expect(intersects(first, second)).toBe(false);
    expect(second.y).toBeGreaterThan(first.y + first.height);

    const path = await downloadAndReopen(page);
    const xml = strFromU8(unzipSync(new Uint8Array(readFileSync(path)))["word/document.xml"]);
    expect(xml.match(/<w:tblpPr\b/g)).toHaveLength(2);
    expect(xml).toContain('<w:tblOverlap w:val="never"');
    const reopenedFirst = await tableBounds(page, firstTableTexts);
    const reopenedSecond = await tableBounds(page, secondTableTexts);
    expect(intersects(reopenedFirst, reopenedSecond)).toBe(false);
    expect(reopenedSecond.y).toBeGreaterThan(reopenedFirst.y + reopenedFirst.height);
  });

  test("typing around a repeatedly moved table stays in document flow", async ({ page }) => {
    await open(page);
    await dragTableBy(page, "Team", 24, 54);
    await dragTableBy(page, "Team", -18, 34);
    await dragTableBy(page, "Team", 12, -16);

    const pageBox = (await page.locator(".dxw-page").first().boundingBox())!;
    const clicks = [
      ["~ABOVE~", (box: Box) => ({ x: box.x + 20, y: box.y - 12 })],
      ["~LEFT~", (box: Box) => ({ x: pageBox.x + 10, y: box.y + box.height / 2 })],
      ["~RIGHT~", (box: Box) => ({ x: pageBox.x + pageBox.width - 10, y: box.y + box.height / 2 })],
      ["~BELOW~", (box: Box) => ({ x: box.x + 20, y: box.y + box.height + 12 })],
    ] as const;
    for (const [marker, point] of clicks) {
      const table = await floatingTableBounds(page, "Team");
      const click = point(table);
      await page.mouse.click(click.x, click.y);
      await page.keyboard.type(marker);
      await page.waitForTimeout(200);
      await expect(page.locator(".dxw-page span").filter({ hasText: marker }).first()).toBeVisible();
    }

    const table = await floatingTableBounds(page, "Team");
    for (const [marker] of clicks) {
      const box = (await page.locator(".dxw-page span").filter({ hasText: marker }).first().boundingBox())!;
      expect(intersects(table, box), `${marker} landed inside a table cell`).toBe(false);
    }
    const below = (await page.locator(".dxw-page span").filter({ hasText: "~BELOW~" }).first().boundingBox())!;
    expect(below.y).toBeGreaterThanOrEqual(table.y + table.height - 1);

    const path = await downloadAndReopen(page);
    const xml = strFromU8(unzipSync(new Uint8Array(readFileSync(path)))["word/document.xml"]);
    const cells = xml.match(/<w:tc\b[\s\S]*?<\/w:tc>/g) ?? [];
    for (const [marker] of clicks) {
      expect(xml).toContain(marker);
      expect(cells.some((cell) => cell.includes(marker)), `${marker} saved inside w:tc`).toBe(false);
      await expect(page.locator(".dxw-page span").filter({ hasText: marker }).first()).toBeVisible();
    }
    const reopenedTable = await floatingTableBounds(page, "Team");
    for (const [marker] of clicks) {
      const box = (await page.locator(".dxw-page span").filter({ hasText: marker }).first().boundingBox())!;
      expect(intersects(reopenedTable, box), `${marker} reopened inside a table cell`).toBe(false);
    }
  });
});

test("tables, equations, images, shapes, and WordArt stay on-sheet after movement and reopen", async ({ page }) => {
  test.setTimeout(120_000);
  await page.goto("/?doc=/fixtures/showcase.docx");
  await page.waitForSelector(".dxw-page span");
  const title = exact(page, "Sample");
  await title.click();
  await page.keyboard.press("End");
  await page.getByRole("button", { name: "insert", exact: true }).click();
  await tool(page, "Insert shape").click();
  await page.getByLabel("Shape text").fill("Boundary shape");
  await page.getByTitle("Insert Rounded rectangle").click();

  await title.click();
  await page.keyboard.press("End");
  await tool(page, "Insert WordArt").click();
  await page.getByLabel("WordArt text").fill("BOUNDARY ART");
  await page.getByTitle("Insert WordArt Plain").click();

  await dragTableBy(page, "Model", 700, 70);
  const drawings = page.locator("[data-dxw-drawing]");
  const drawingCount = await drawings.count();
  expect(drawingCount).toBeGreaterThanOrEqual(2);
  for (const [name, drawing] of [
    ["Boundary shape", drawings.nth(drawingCount - 2)],
    ["BOUNDARY ART", drawings.nth(drawingCount - 1)],
  ] as const) {
    const box = await drawing.boundingBox();
    if (!box) throw new Error(`Missing ${name}`);
    await page.mouse.move(box.x + box.width - 12, box.y + box.height - 12);
    await page.mouse.down();
    await page.mouse.move(2000, box.y + box.height + 35, { steps: 8 });
    await page.mouse.up();
    await page.waitForTimeout(300);
  }
  await expectRequestedObjectsWithinPages(page);

  const path = await downloadAndReopen(page);
  const xml = strFromU8(unzipSync(new Uint8Array(readFileSync(path)))["word/document.xml"]);
  expect(xml).toContain("<w:tbl");
  expect(xml).toContain("<m:oMath");
  expect(xml).toContain("<pic:pic");
  expect(xml).toContain("<wps:wsp");
  expect(xml).toContain("<a:prstTxWarp");
  await expectRequestedObjectsWithinPages(page);
});

test.describe("caret after a terminal table", () => {
  test("clicking below the table enters the trailing paragraph", async ({ page }) => {
    await open(page);
    const lastCell = (await exact(page, "16").boundingBox())!;
    await expectTypedBelowTable(page, "CLICKBELOW", () =>
      page.mouse.click(lastCell.x, lastCell.y + lastCell.height + 8),
    );

    const pending = page.waitForEvent("download");
    await page.getByText("Download", { exact: true }).click();
    const saved = await (await pending).path();
    expect(saved).not.toBeNull();
    await page.locator("#docx-upload").setInputFiles(saved!);
    await page.waitForSelector(".dxw-page span");
    await page.waitForTimeout(300);
    const reopenedCell = (await exact(page, "16").boundingBox())!;
    const reopenedMarker = (await exact(page, "CLICKBELOW").boundingBox())!;
    expect(reopenedMarker.y).toBeGreaterThan(reopenedCell.y + reopenedCell.height);
  });

  test("ArrowDown exits the table into the trailing paragraph", async ({ page }) => {
    await open(page);
    await exact(page, "16").click();
    await expectTypedBelowTable(page, "ARROWDOWN", () => page.keyboard.press("ArrowDown"));
  });

  test("Cmd/Ctrl+ArrowDown exits the table into the trailing paragraph", async ({ page }) => {
    await open(page);
    await exact(page, "16").click();
    await expectTypedBelowTable(page, "CMDDOWN", () => page.keyboard.press(`${MOD}+ArrowDown`));
  });

  test("a far-below click aligned with the last cell stays outside the table", async ({ page }) => {
    await open(page);
    const lastCell = exact(page, "16");
    await page.locator(".dxw-pages").locator("..").evaluate((element) => {
      element.scrollTop = element.scrollHeight - element.clientHeight;
    });
    const cellBox = (await lastCell.boundingBox())!;
    const pageBox = (await page.locator(".dxw-page").first().boundingBox())!;
    await page.mouse.click(cellBox.x + cellBox.width / 2, pageBox.y + pageBox.height - 80);
    await page.keyboard.type("FARWHITE");
    await page.waitForTimeout(250);

    await expect(lastCell).toHaveText("16");
    const marker = (await exact(page, "FARWHITE").boundingBox())!;
    expect(marker.y).toBeGreaterThan(cellBox.y + cellBox.height);
  });
});
