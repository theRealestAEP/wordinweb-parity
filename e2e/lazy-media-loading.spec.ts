import { expect, test } from "@playwright/test";

test("loads the 3D runtime only for documents that contain a 3D model", async ({ page }) => {
  const runtimeRequests: string[] = [];
  page.on("request", (request) => {
    if (/model-viewer|three/i.test(request.url())) runtimeRequests.push(request.url());
  });

  await page.goto("/?doc=/fixtures/benchmark.docx");
  await expect(page.locator(".dxw-page").first()).toBeVisible();
  expect(await page.evaluate(() => Boolean(customElements.get("model-viewer")))).toBe(false);
  expect(runtimeRequests).toEqual([]);

  await page.goto("/?doc=/fixtures/model3d-cube.docx");
  await expect(page.locator("[data-dxw-model3d-viewer]").first()).toBeVisible();
  await expect.poll(() => page.evaluate(() => Boolean(customElements.get("model-viewer")))).toBe(true);
  expect(runtimeRequests.length).toBeGreaterThan(0);
});

test("renders WMF images after loading the decoder on demand", async ({ page }) => {
  await page.goto("/?doc=/fixtures/wild2-sci-ieee-2col.docx");
  const image = page.locator('img[data-dxw-image-format="wmf"]').first();
  await image.scrollIntoViewIfNeeded();
  await expect.poll(() => image.evaluate((node) => node.naturalWidth)).toBeGreaterThan(0);
  await expect(image).toHaveAttribute("src", /^data:image\/png;base64,/);
});
