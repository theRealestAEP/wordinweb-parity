import { readFileSync } from "node:fs";
import { expect, Page, test } from "@playwright/test";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

const MOD = process.platform === "darwin" ? "Meta" : "Control";

async function waitForDocument(page: Page): Promise<void> {
  await page.waitForSelector(".dxw-page", { state: "attached" });
  await page.waitForTimeout(250);
}

async function mathBottom(page: Page): Promise<number> {
  return page.locator("[data-dxw-math]").evaluateAll((elements) =>
    Math.max(...elements.map((element) => element.getBoundingClientRect().bottom)),
  );
}

async function download(page: Page): Promise<string> {
  const pending = page.waitForEvent("download");
  await page.getByText("Download", { exact: true }).click();
  const path = await (await pending).path();
  if (!path) throw new Error("download did not produce a path");
  return path;
}

test.describe("terminal object caret invariant", () => {
  test("clicking below terminal equations types below them and survives save/reopen", async ({ page }) => {
    await page.goto("/?doc=/fixtures/preset-equations.docx");
    await waitForDocument(page);
    const bottom = await mathBottom(page);
    const last = (await page.locator("[data-dxw-math]").last().boundingBox())!;
    await page.mouse.click(last.x, bottom + 80);
    await page.keyboard.type("CLICKED BELOW MATH");
    await page.waitForTimeout(250);
    await expect(page.locator(".dxw-page").first()).toContainText("CLICKED BELOW MATH");
    const marker = (await page.locator(".dxw-page span").filter({ hasText: /^CLICKED$/ }).boundingBox())!;
    expect(marker.y).toBeGreaterThan(bottom);

    const saved = await download(page);
    await page.locator("#docx-upload").setInputFiles(saved);
    await waitForDocument(page);
    await expect(page.locator(".dxw-page").first()).toContainText("CLICKED BELOW MATH");
    const reopened = (await page.locator(".dxw-page span").filter({ hasText: /^CLICKED$/ }).boundingBox())!;
    expect(reopened.y).toBeGreaterThan(await mathBottom(page));
  });

  for (const key of ["ArrowDown", `${MOD}+ArrowDown`]) {
    test(`${key} reaches the trailing paragraph after terminal equations`, async ({ page }) => {
      await page.goto("/?doc=/fixtures/preset-equations.docx");
      await waitForDocument(page);
      const bottom = await mathBottom(page);
      await page.locator(".dxw-page span").filter({ hasText: /^Integral$/ }).first().click();
      await page.keyboard.press(key);
      const markerText = key === "ArrowDown" ? "ARROW BELOW MATH" : "COMMAND BELOW MATH";
      await page.keyboard.type(markerText);
      await page.waitForTimeout(250);
      await expect(page.locator(".dxw-page").first()).toContainText(markerText);
      const firstWord = markerText.slice(0, markerText.indexOf(" "));
      const marker = (await page.locator(".dxw-page span").filter({ hasText: new RegExp(`^${firstWord}$`) }).boundingBox())!;
      expect(marker.y).toBeGreaterThan(bottom);
    });
  }

  test("a terminal image-only paragraph also has a click target below it", async ({ page }) => {
    const png = Uint8Array.from(Buffer.from(
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=",
      "base64",
    ));
    const document = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body><w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="terminal"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rId2"/></pic:blipFill><pic:spPr><a:xfrm><a:ext cx="914400" cy="914400"/></a:xfrm></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`;
    const bytes = zipSync({
      "[Content_Types].xml": strToU8(`<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`),
      "_rels/.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
      "word/document.xml": strToU8(document),
      "word/_rels/document.xml.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>`),
      "word/media/image1.png": png,
    });
    await page.goto("/");
    await page.locator("#docx-upload").setInputFiles({ name: "terminal-image.docx", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", buffer: Buffer.from(bytes) });
    await waitForDocument(page);
    await page.waitForSelector(".dxw-page img", { state: "visible" });
    const image = (await page.locator(".dxw-page img").boundingBox())!;
    await page.mouse.click(image.x, image.y + image.height + 60);
    await page.keyboard.type("BELOW IMAGE");
    await page.waitForTimeout(250);
    await expect(page.locator(".dxw-page").first()).toContainText("BELOW IMAGE");
    const marker = (await page.locator(".dxw-page span").filter({ hasText: /^BELOW$/ }).boundingBox())!;
    expect(marker.y).toBeGreaterThan(image.y + image.height);

    const saved = await download(page);
    await page.locator("#docx-upload").setInputFiles(saved);
    await waitForDocument(page);
    await expect(page.locator(".dxw-page").first()).toContainText("BELOW IMAGE");
    const reopenedImage = (await page.locator(".dxw-page img").boundingBox())!;
    const reopenedMarker = (await page.locator(".dxw-page span").filter({ hasText: /^BELOW$/ }).boundingBox())!;
    expect(reopenedMarker.y).toBeGreaterThan(reopenedImage.y + reopenedImage.height);
  });
});

test("real NIH numId 7 blank d/e items remain clickable and editable", async ({ page }) => {
  test.setTimeout(120_000);
  await page.goto("/?doc=/fixtures/wild2-legal-nih-contract.docx");
  await waitForDocument(page);
  const contractPage = page.locator(".dxw-page").nth(10);
  await contractPage.scrollIntoViewIfNeeded();
  await page.waitForTimeout(300);
  const ga = contractPage.locator("span").filter({ hasText: /^Ga$/ }).first();
  const initialGa = (await ga.boundingBox())!;
  await page.mouse.click(initialGa.x + 1, initialGa.y + initialGa.height / 2);
  await page.keyboard.press("Enter");
  await expect.poll(async () => {
    const gaBox = await ga.boundingBox();
    const labels = await contractPage.locator("span").filter({ hasText: /^e\.$/ }).evaluateAll((spans) =>
      spans.map((span) => span.getBoundingClientRect().y),
    );
    return !!gaBox && labels.some((y) => Math.abs(y - gaBox.y) < 2);
  }, { timeout: 30_000 }).toBe(true);
  await expect(page.locator('[aria-busy="true"]')).toHaveCount(0);
  await page.keyboard.press("Enter");
  await expect.poll(async () => {
    const gaBox = await ga.boundingBox();
    const labels = await contractPage.locator("span").filter({ hasText: /^f\.$/ }).evaluateAll((spans) =>
      spans.map((span) => span.getBoundingClientRect().y),
    );
    return !!gaBox && labels.some((y) => Math.abs(y - gaBox.y) < 2);
  }, { timeout: 30_000 }).toBe(true);
  await expect(page.locator('[aria-busy="true"]')).toHaveCount(0);
  const gaBox = (await ga.boundingBox())!;
  const d = (await contractPage.locator("span").filter({ hasText: /^d\.$/ }).first().boundingBox())!;
  await page.mouse.click(gaBox.x + 1, d.y + d.height / 2);
  await page.keyboard.type("DTEXT");
  await page.waitForTimeout(250);
  const dMarker = (await page.locator(".dxw-page span").filter({ hasText: /^DTEXT$/ }).boundingBox())!;
  expect(dMarker.y).toBeCloseTo(d.y, 0);
  expect(dMarker.x).toBeCloseTo(gaBox.x, 0);
  const e = (await contractPage.locator("span").filter({ hasText: /^e\.$/ }).first().boundingBox())!;
  await page.mouse.click(gaBox.x + 1, e.y + e.height / 2);
  await page.keyboard.type("ETEXT");
  await page.waitForTimeout(250);
  const eMarker = (await page.locator(".dxw-page span").filter({ hasText: /^ETEXT$/ }).boundingBox())!;
  expect(eMarker.y).toBeCloseTo(e.y, 0);
  expect(eMarker.x).toBeCloseTo(gaBox.x, 0);

  const saved = await download(page);
  const xml = strFromU8(unzipSync(new Uint8Array(readFileSync(saved)))["word/document.xml"]);
  const dParagraph = xml.match(/<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*DTEXT(?:(?!<\/w:p>)[\s\S])*<\/w:p>/)?.[0] ?? "";
  const eParagraph = xml.match(/<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*ETEXT(?:(?!<\/w:p>)[\s\S])*<\/w:p>/)?.[0] ?? "";
  expect(dParagraph).toContain('<w:numId w:val="7"');
  expect(eParagraph).toContain('<w:numId w:val="7"');
  expect(xml.indexOf("DTEXT")).toBeLessThan(xml.indexOf("ETEXT"));
  expect(xml.indexOf("ETEXT")).toBeLessThan(xml.indexOf("Ga the Neken Gosi"));

  await page.locator("#docx-upload").setInputFiles(saved);
  await waitForDocument(page);
  const reopenedPage = page.locator(".dxw-page").nth(10);
  await reopenedPage.scrollIntoViewIfNeeded();
  await page.waitForTimeout(250);
  await expect(reopenedPage.locator("span").filter({ hasText: /^DTEXT$/ })).toHaveCount(1);
  await expect(reopenedPage.locator("span").filter({ hasText: /^ETEXT$/ })).toHaveCount(1);
});

test.describe("table movement and floating flow", () => {
  test("the moved benchmark table excludes the following inline table", async ({ page }) => {
    const files = unzipSync(new Uint8Array(readFileSync("apps/demo/public/fixtures/benchmark.docx")));
    let xml = strFromU8(files["word/document.xml"]);
    xml = xml.replace(
      "<w:tblPr>",
      `<w:tblPr><w:tblpPr w:leftFromText="0" w:rightFromText="0" w:topFromText="0" w:bottomFromText="0" w:horzAnchor="page" w:vertAnchor="page" w:tblpX="2325" w:tblpY="9480"/>`,
    );
    files["word/document.xml"] = strToU8(xml);
    const edited = zipSync(files);
    await page.goto("/");
    await page.locator("#docx-upload").setInputFiles({ name: "benchmark-moved.docx", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", buffer: Buffer.from(edited) });
    await waitForDocument(page);
    const lastFloatRow = (await page.getByText("Justify", { exact: true }).boundingBox())!;
    const inline = (await page.locator(".dxw-page span").filter({ hasText: /^Fixed$/ }).first().boundingBox())!;
    expect(inline.y).toBeGreaterThan(lastFloatRow.y + lastFloatRow.height);
  });

  test("the table handle survives a stepped pointer path and saved moves stay on-sheet", async ({ page }) => {
    await page.goto("/?doc=/fixtures/preset-tables.docx");
    await waitForDocument(page);
    const tableText = page.getByText("Region", { exact: true });
    const table = (await tableText.boundingBox())!;
    const beforeX = table.x;
    await page.mouse.move(table.x + table.width / 2, table.y + table.height / 2);
    const handles = page.locator("[data-dxw-table-move]");
    const visibleIndex = await handles.evaluateAll((elements) =>
      elements.findIndex((element) => (element as HTMLElement).style.opacity === "1"),
    );
    expect(visibleIndex).toBeGreaterThanOrEqual(0);
    const visible = handles.nth(visibleIndex);
    const handle = (await visible.boundingBox())!;
    // Leave through the top edge well away from the corner corridor. The
    // handle stays armed during the short, deliberate pointer travel.
    await page.mouse.move(table.x + table.width / 2, table.y - 10, { steps: 3 });
    await page.waitForTimeout(120);
    await page.mouse.move(handle.x + handle.width / 2, handle.y + handle.height / 2, { steps: 5 });
    await page.waitForTimeout(150);
    await expect(visible).toHaveCSS("opacity", "1");
    await page.mouse.down();
    await page.mouse.move(2000, handle.y + 90, { steps: 8 });
    await page.mouse.up();
    await page.waitForTimeout(300);
    const movedX = ((await page.getByText("Region", { exact: true }).boundingBox())!).x;
    expect(movedX).toBeGreaterThan(beforeX + 40);

    const saved = await download(page);
    await page.locator("#docx-upload").setInputFiles(saved);
    await waitForDocument(page);
    const reopenedRegion = (await page.getByText("Region", { exact: true }).boundingBox())!;
    expect(reopenedRegion.x).toBeCloseTo(movedX, 0);
    const outside = await page.locator(".dxw-page").evaluateAll((pages) => pages.flatMap((page) => {
      const sheet = page.getBoundingClientRect();
      return [...page.querySelectorAll<HTMLElement>("[data-dxw-item-kind]")]
        .filter((item) => item.dataset.dxwItemKind !== "grip")
        .map((item) => item.getBoundingClientRect())
        .filter((rect) => rect.left < sheet.left - 1 || rect.right > sheet.right + 1)
        .map((rect) => ({ left: rect.left, right: rect.right, sheetLeft: sheet.left, sheetRight: sheet.right }));
    }));
    expect(outside).toEqual([]);
  });
});
