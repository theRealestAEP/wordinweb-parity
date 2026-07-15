import { test, expect } from "@playwright/test";
import { readFileSync } from "node:fs";
import { strFromU8, unzipSync } from "fflate";

test("long documents keep early and post-undo edits visible and saved", async ({ page }) => {
  test.setTimeout(120_000);
  await page.addInitScript(() => {
    (globalThis as { __dxwPerf?: unknown }).__dxwPerf = { samples: [] };
  });
  await page.goto("/?doc=/fixtures/wild2-legal-nih-contract.docx");

  const pages = page.locator(".dxw-page");
  await pages.first().waitFor({ state: "attached", timeout: 60_000 });
  await expect(pages).toHaveCount(419, { timeout: 60_000 });

  const firstSpan = pages.first().locator("span", { hasText: /\w{3,}/ }).first();
  await firstSpan.click();
  await page.keyboard.press("End");
  await page.keyboard.type("FIRSTEDIT");

  const firstPageSamples = await page.evaluate(() =>
    (globalThis as { __dxwPerf?: { samples?: { total: number; pagesReused: number }[] } }).__dxwPerf?.samples ?? [],
  );
  expect(firstPageSamples).toHaveLength(9);
  expect(firstPageSamples.every((sample) => sample.pagesReused === 418)).toBe(true);
  expect(Math.max(...firstPageSamples.map((sample) => sample.total))).toBeLessThan(250);

  await page.evaluate(() => {
    const state = globalThis as typeof globalThis & {
      __dxwLayoutTicks?: number;
      __dxwLayoutTimer?: number;
      __dxwLayoutLastTick?: number;
      __dxwLayoutMaxGap?: number;
    };
    state.__dxwLayoutTicks = 0;
    state.__dxwLayoutLastTick = performance.now();
    state.__dxwLayoutMaxGap = 0;
    state.__dxwLayoutTimer = window.setInterval(() => {
      const now = performance.now();
      state.__dxwLayoutTicks!++;
      state.__dxwLayoutMaxGap = Math.max(state.__dxwLayoutMaxGap!, now - state.__dxwLayoutLastTick!);
      state.__dxwLayoutLastTick = now;
    }, 10);
  });
  await page.keyboard.press("ControlOrMeta+z");
  const layoutStatus = page.locator("[data-dxw-layout-status]");
  await expect(layoutStatus).toBeVisible();
  await expect(page.locator("[data-dxw-layout-busy]")).toHaveAttribute("aria-busy", "true");
  // Supersede the in-flight undo layout twice. Only the final undo may paint.
  await page.getByTitle("Redo (⇧⌘Z)").click();
  await page.getByTitle("Undo (⌘Z)").click();
  await expect(layoutStatus).toBeHidden({ timeout: 30_000 });
  const responsiveness = await page.evaluate(() => {
    const state = globalThis as typeof globalThis & {
      __dxwLayoutTicks?: number;
      __dxwLayoutTimer?: number;
      __dxwLayoutMaxGap?: number;
    };
    clearInterval(state.__dxwLayoutTimer);
    return { ticks: state.__dxwLayoutTicks ?? 0, maxGap: state.__dxwLayoutMaxGap ?? Infinity };
  });
  console.log(`[book-layout] timer ticks=${responsiveness.ticks} maxGap=${responsiveness.maxGap.toFixed(1)}ms`);
  expect(responsiveness.ticks).toBeGreaterThan(5);
  expect(responsiveness.maxGap).toBeLessThan(500);
  await expect(pages.first()).not.toContainText("FIRSTEDIT");

  const targetPage = pages.nth(20);
  await targetPage.scrollIntoViewIfNeeded();
  const targetSpan = targetPage.locator("span", { hasText: /\w{3,}/ }).first();
  await targetSpan.waitFor({ state: "visible" });
  await targetSpan.click();
  await page.keyboard.press("End");

  const marker = "ZQXVJZ";
  await page.keyboard.type(marker);
  await expect(page.locator(".dxw-page", { hasText: marker })).toHaveCount(1);
  await expect(page.locator("[data-dxw-caret]")).toBeVisible();

  const downloadPromise = page.waitForEvent("download");
  await page.getByText("Download", { exact: true }).click();
  const download = await downloadPromise;
  const path = await download.path();
  expect(path).not.toBeNull();
  const parts = unzipSync(new Uint8Array(readFileSync(path!)));
  expect(strFromU8(parts["word/document.xml"])).toContain(marker);
});

test("long-document header edits preview immediately, refresh globally, and save", async ({ page }) => {
  test.setTimeout(120_000);
  await page.addInitScript(() => {
    (globalThis as { __dxwPerf?: unknown }).__dxwPerf = { samples: [] };
  });
  await page.goto("/?doc=/fixtures/wild2-med-phase23-protocol.docx");

  const pages = page.locator(".dxw-page");
  await pages.first().waitFor({ state: "attached", timeout: 60_000 });
  await expect(pages).toHaveCount(70, { timeout: 60_000 });

  const headerText = pages.first().locator("span[data-dxw-hf]").filter({ hasText: /^Hegulufu$/ }).first();
  const box = (await headerText.boundingBox())!;
  await headerText.dblclick({ position: { x: box.width - 2, y: box.height / 2 } });

  const marker = "HDRQXVJ";
  await page.keyboard.type(marker, { delay: 10 });
  // The debounce window keeps the entire burst editable and previews it before
  // the expensive all-page layout starts.
  await expect(pages.first()).toContainText(marker);
  await expect(page.locator("[data-dxw-layout-status]")).toHaveCount(0);

  const laterPage = pages.nth(4);
  await laterPage.scrollIntoViewIfNeeded();
  await expect(laterPage.locator("[data-dxw-hf]", { hasText: marker })).toHaveCount(1);
  // The text stayed on one line, so the geometry gate refreshed only the
  // repeated header layer and never entered the blocking repagination state.
  await expect(page.locator("[data-dxw-layout-status]")).toHaveCount(0);
  const timing = await page.evaluate(() =>
    (globalThis as { __dxwPerf?: { last?: { layout?: number; render?: number } } }).__dxwPerf?.last,
  );
  console.log(`[header-fast] layout=${timing?.layout?.toFixed(1)}ms render=${timing?.render?.toFixed(1)}ms`);
  expect(timing?.layout).toBeLessThan(500);

  const downloadPromise = page.waitForEvent("download");
  await page.getByText("Download", { exact: true }).click();
  const download = await downloadPromise;
  const path = await download.path();
  expect(path).not.toBeNull();
  const parts = unzipSync(new Uint8Array(readFileSync(path!)));
  const headers = Object.entries(parts)
    .filter(([name]) => /^word\/header\d+\.xml$/.test(name))
    .map(([, bytes]) => strFromU8(bytes))
    .join("\n");
  expect(headers).toContain(marker);
});
