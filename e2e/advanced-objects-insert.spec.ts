import { expect, Locator, Page, test } from "@playwright/test";
import { readFileSync } from "node:fs";
import { strFromU8, unzipSync } from "fflate";

async function openInsert(page: Page): Promise<Locator> {
  await page.goto("/?doc=/fixtures/benchmark.docx");
  await page.waitForSelector(".dxw-page span");
  await page.getByRole("button", { name: "insert", exact: true }).click();
  const target = page.locator(".dxw-page span", { hasText: "classic." }).first();
  await target.click({ force: true });
  await page.keyboard.press("End");
  return target;
}

function minimalGlb(): Buffer {
  const json = Buffer.from(JSON.stringify({ asset: { version: "2.0" }, scenes: [{}], scene: 0 }).padEnd(52, " "));
  const glb = Buffer.alloc(20 + json.length);
  glb.writeUInt32LE(0x46546c67, 0);
  glb.writeUInt32LE(2, 4);
  glb.writeUInt32LE(glb.length, 8);
  glb.writeUInt32LE(json.length, 12);
  glb.writeUInt32LE(0x4e4f534a, 16);
  json.copy(glb, 20);
  return glb;
}

test("advanced Insert creates editable native 3D, media, and embedded objects", async ({ page }) => {
  await openInsert(page);

  await page.locator('input[accept*="gltf"]').setInputFiles({
    name: "cube.glb",
    mimeType: "model/gltf-binary",
    buffer: minimalGlb(),
  });
  const model = page.locator("[data-dxw-model3d]").last();
  await expect(model).toBeVisible();
  const caretTarget = page.locator(".dxw-page span", { hasText: "classic." }).first();
  await caretTarget.click();
  await page.keyboard.press("End");
  await page.getByRole("button", { name: "Media", exact: true }).click();
  await page.getByLabel("Online video URL").fill("https://youtu.be/dQw4w9WgXcQ");
  await page.getByRole("button", { name: "Insert online video", exact: true }).click();
  await expect(page.locator("[data-dxw-web-video]").last()).toBeVisible();

  await page.keyboard.press("Escape");
  await caretTarget.click();
  await page.keyboard.press("End");
  await page.locator('input[aria-label="Embedded object file"]').setInputFiles({
    name: "report.txt",
    mimeType: "text/plain",
    buffer: Buffer.from("embedded report data"),
  });
  const object = page.locator("[data-dxw-embedded-object]").last();
  await expect(object).toBeVisible();

  const before = await model.boundingBox();
  expect(before).not.toBeNull();
  await model.click();
  await expect(page.locator("[data-dxw-img-handle]")).toHaveCount(8);
  await expect(page.getByRole("button", { name: "3D Format", exact: true })).toBeVisible();
  await expect(page.locator("[data-dxw-object-format]").getByRole("button", { name: "Reset 3D", exact: true })).toBeVisible();
  await expect(page.locator("[data-dxw-object-format]").getByRole("button", { name: "Fill", exact: true })).toHaveCount(0);
  const corner = await page.locator('[data-dxw-img-handle="se"]').boundingBox();
  expect(corner).not.toBeNull();
  await page.mouse.move(corner!.x + corner!.width / 2, corner!.y + corner!.height / 2);
  await page.mouse.down();
  await page.mouse.move(corner!.x + 55, corner!.y + 35, { steps: 5 });
  await page.mouse.up();
  const resized = await page.locator("[data-dxw-model3d]").last().boundingBox();
  expect(resized!.width).toBeGreaterThan(before!.width + 20);

  const viewer = page.locator("[data-dxw-model3d-viewer]").last();
  const viewerBox = (await viewer.boundingBox())!;
  await page.mouse.move(viewerBox.x + viewerBox.width / 2, viewerBox.y + viewerBox.height / 2);
  await page.mouse.down();
  await page.mouse.move(viewerBox.x + viewerBox.width / 2 + 40, viewerBox.y + viewerBox.height / 2 + 20, { steps: 6 });
  await page.mouse.up();
  await expect(viewer).toHaveAttribute("orientation", /0deg 10deg 20deg/);

  const rotatedBox = (await model.boundingBox())!;
  const moveGrip = page.locator('[data-dxw-object-move][title="Drag to move 3D object"]');
  const moveBox = (await moveGrip.boundingBox())!;
  await page.mouse.move(moveBox.x + moveBox.width / 2, moveBox.y + moveBox.height / 2);
  await page.mouse.down();
  await page.mouse.move(moveBox.x + moveBox.width / 2 + 50, moveBox.y + moveBox.height / 2, { steps: 6 });
  await page.mouse.up();
  const movedBox = (await model.boundingBox())!;
  expect(movedBox.x).toBeGreaterThan(rotatedBox.x + 35);

  await page.locator("[data-dxw-object-format]").getByRole("button", { name: "Reset 3D", exact: true }).click();
  await expect(viewer).toHaveAttribute("orientation", "0deg 0deg 0deg");
  await page.keyboard.press(process.platform === "darwin" ? "Meta+z" : "Control+z");
  await expect(viewer).toHaveAttribute("orientation", /0deg 10deg 20deg/);

  await page.keyboard.press("Escape");
  await object.click();
  await expect(page.locator("[data-dxw-img-handle]")).toHaveCount(8);

  const pending = page.waitForEvent("download");
  await page.getByText("Download", { exact: true }).click();
  const path = await (await pending).path();
  expect(path).not.toBeNull();
  const files = unzipSync(new Uint8Array(readFileSync(path!)));
  const documentXml = strFromU8(files["word/document.xml"]);
  const rels = strFromU8(files["word/_rels/document.xml.rels"]);
  const contentTypes = strFromU8(files["[Content_Types].xml"]);
  expect(documentXml).toContain("<am3d:model3d");
  expect(documentXml).toMatch(/<am3d:rot\b[^>]*ax="600000"[^>]*ay="1200000"/);
  expect(documentXml).toContain("<wp15:webVideoPr");
  expect(documentXml).toContain("<o:OLEObject");
  expect(documentXml).toContain("https://www.youtube.com/embed/dQw4w9WgXcQ");
  expect(rels).toContain("relationships/model3d");
  expect(rels).toContain("relationships/oleObject");
  expect(contentTypes).toContain("model/gltf-binary");
  expect(Array.from(files["word/media/model3d1.glb"].slice(0, 4))).toEqual([0x67, 0x6c, 0x54, 0x46]);
  expect(Array.from(files["word/embeddings/oleObject1.bin"].slice(0, 8))).toEqual([208, 207, 17, 224, 161, 177, 26, 225]);
});

test("double-clicking an embedded Word document downloads the original valid DOCX", async ({ page }) => {
  await page.goto("/?doc=/fixtures/word-interop-embedded-only.docx");
  const object = page.locator("[data-dxw-embedded-object]").first();
  await expect(object).toBeVisible();

  const pending = page.waitForEvent("download");
  await object.dblclick();
  const download = await pending;
  expect(download.suggestedFilename()).toMatch(/\.docx$/i);
  const path = await download.path();
  expect(path).not.toBeNull();

  const files = unzipSync(new Uint8Array(readFileSync(path!)));
  expect(files["[Content_Types].xml"]).toBeDefined();
  expect(strFromU8(files["word/document.xml"])).toContain("Word interoperability validation");
});
