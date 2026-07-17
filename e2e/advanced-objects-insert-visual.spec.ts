import { expect, test } from "@playwright/test";

test("advanced Insert renders a selected native 3D model poster", async ({ page }) => {
  await page.goto("/?doc=/fixtures/benchmark.docx");
  await page.waitForSelector(".dxw-page span");
  await page.getByRole("button", { name: "insert", exact: true }).click();
  const target = page.locator(".dxw-page span", { hasText: "classic." }).first();
  await target.click();
  await page.keyboard.press("End");
  const json = Buffer.from(JSON.stringify({ asset: { version: "2.0" }, scenes: [{}], scene: 0 }).padEnd(52, " "));
  const glb = Buffer.alloc(20 + json.length);
  glb.writeUInt32LE(0x46546c67, 0);
  glb.writeUInt32LE(2, 4);
  glb.writeUInt32LE(glb.length, 8);
  glb.writeUInt32LE(json.length, 12);
  glb.writeUInt32LE(0x4e4f534a, 16);
  json.copy(glb, 20);
  await page.locator('input[accept*="gltf"]').setInputFiles({ name: "cube.glb", mimeType: "model/gltf-binary", buffer: glb });
  const model = page.locator("[data-dxw-model3d]").last();
  await expect(model).toBeVisible();
  await model.click();
  await expect(page.locator("[data-dxw-img-handle]")).toHaveCount(8);
  await page.evaluate(() => window.scrollTo(0, 0));
  await expect(page.locator(".dxw-page").first()).toHaveScreenshot("inserted-3d-model.png", {
    animations: "disabled",
    caret: "hide",
  });
});
