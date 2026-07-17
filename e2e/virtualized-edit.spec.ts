import { test, expect } from "@playwright/test";
import { readFileSync } from "node:fs";
import { strFromU8, unzipSync } from "fflate";

test("Cmd+A does not materialize every page of a long document", async ({ page }) => {
  test.setTimeout(120_000);
  await page.goto("/?doc=/fixtures/wild2-med-phase23-protocol.docx");

  const pages = page.locator(".dxw-page");
  await pages.first().waitFor({ state: "attached", timeout: 60_000 });
  await expect(pages).toHaveCount(70, { timeout: 60_000 });
  const firstSpan = pages.first().locator("span", { hasText: /\w{3,}/ }).first();
  await firstSpan.click();
  const mountedBefore = await page.evaluate(
    () => new Set([...document.querySelectorAll(".dxw-page span")].map((span) => span.closest(".dxw-page"))).size,
  );

  const started = performance.now();
  await page.keyboard.press("ControlOrMeta+a");
  const elapsed = performance.now() - started;
  const mountedAfter = await page.evaluate(
    () => new Set([...document.querySelectorAll(".dxw-page span")].map((span) => span.closest(".dxw-page"))).size,
  );

  expect(elapsed).toBeLessThan(1_000);
  expect(mountedAfter).toBeLessThanOrEqual(mountedBefore + 4);
  await expect(page.locator(".dxw-sel")).not.toHaveCount(0);
});

test("NIH Select All deletes and undoes the complete virtualized body", async ({ page }) => {
  test.setTimeout(120_000);
  await page.goto("/?doc=/fixtures/wild2-legal-nih-contract.docx");

  const pages = page.locator(".dxw-page");
  await pages.first().waitFor({ state: "attached", timeout: 60_000 });
  await expect(pages).toHaveCount(419, { timeout: 60_000 });
  const targetPage = pages.nth(209);
  await targetPage.scrollIntoViewIfNeeded();
  const bodyText = targetPage.locator("span:not([data-dxw-hf])", { hasText: /[A-Za-z]{4,}/ }).nth(20);
  await bodyText.waitFor({ state: "visible" });
  await bodyText.click();

  const mountedBefore = await page.evaluate(
    () => new Set([...document.querySelectorAll(".dxw-page span")].map((span) => span.closest(".dxw-page"))).size,
  );
  await page.keyboard.press("ControlOrMeta+a");
  await expect(page.locator(".dxw-sel")).not.toHaveCount(0);
  const mountedAfter = await page.evaluate(
    () => new Set([...document.querySelectorAll(".dxw-page span")].map((span) => span.closest(".dxw-page"))).size,
  );
  expect(mountedAfter).toBeLessThanOrEqual(mountedBefore + 4);

  await page.keyboard.press("Delete");
  const layoutStatus = page.locator("[data-dxw-layout-status]");
  await expect(layoutStatus).toBeVisible({ timeout: 10_000 });
  await expect(layoutStatus).toBeHidden({ timeout: 60_000 });
  await expect(pages).toHaveCount(1);
  await expect(page.locator("[data-dxw-caret]")).toBeVisible();

  await page.keyboard.press("ControlOrMeta+z");
  await expect(pages).toHaveCount(419, { timeout: 60_000 });
  await expect(layoutStatus).toHaveCount(0);
  await pages.nth(209).scrollIntoViewIfNeeded();
  await expect(pages.nth(209).locator("span:not([data-dxw-hf])", { hasText: /[A-Za-z]{4,}/ }).nth(20))
    .toBeVisible();
});

test("virtualized Select All survives multiple formats, save/reopen, replacement, and undo", async ({ page }) => {
  test.setTimeout(120_000);
  await page.goto("/?doc=/fixtures/wild2-med-phase23-protocol.docx");

  const pages = page.locator(".dxw-page");
  await pages.first().waitFor({ state: "attached", timeout: 60_000 });
  await expect(pages).toHaveCount(70, { timeout: 60_000 });
  const bodyText = (index: number) =>
    pages.nth(index).locator("span:not([data-dxw-hf])", { hasText: /[A-Za-z]{4,}/ }).first();
  const representativeTexts: string[] = [];
  for (const index of [0, 34, 69]) {
    await pages.nth(index).scrollIntoViewIfNeeded();
    const span = bodyText(index);
    await span.waitFor({ state: "visible" });
    representativeTexts.push((await span.textContent())!);
  }
  const early = bodyText(0);
  await pages.first().scrollIntoViewIfNeeded();
  await early.waitFor({ state: "visible" });
  await early.click();
  await page.keyboard.press("ControlOrMeta+a");
  await expect(page.locator(".dxw-sel")).not.toHaveCount(0);

  await page.keyboard.press("ControlOrMeta+b");
  await expect(page.locator("[data-dxw-layout-status]")).toBeHidden({ timeout: 30_000 });
  await expect(page.locator(".dxw-sel")).not.toHaveCount(0);
  await page.locator('button[title="Italic"], button[data-tip="Italic"]').first().click();
  await expect(page.locator("[data-dxw-layout-status]")).toBeHidden({ timeout: 30_000 });
  await expect(page.locator(".dxw-sel")).not.toHaveCount(0);
  await page.keyboard.press("ControlOrMeta+u");
  await expect(page.locator("[data-dxw-layout-status]")).toBeHidden({ timeout: 30_000 });
  await expect(page.locator(".dxw-sel")).not.toHaveCount(0);

  for (const index of [0, 34, 69]) {
    await pages.nth(index).scrollIntoViewIfNeeded();
    const span = bodyText(index);
    await span.waitFor({ state: "visible" });
    const style = await span.evaluate((element) => {
      const computed = getComputedStyle(element);
      return { weight: computed.fontWeight, style: computed.fontStyle, decoration: computed.textDecorationLine };
    });
    expect(Number(style.weight)).toBeGreaterThanOrEqual(700);
    expect(style.style).toBe("italic");
    expect(style.decoration).toContain("underline");
  }

  const formattedDownload = page.waitForEvent("download");
  await page.getByText("Download", { exact: true }).click();
  const formattedPath = await (await formattedDownload).path();
  expect(formattedPath).not.toBeNull();
  const formattedParts = unzipSync(new Uint8Array(readFileSync(formattedPath!)));
  const documentXml = strFromU8(formattedParts["word/document.xml"]);
  const bodyXml = documentXml.match(/<w:body>[\s\S]*<\/w:body>/)?.[0] ?? "";
  const textRuns = bodyXml.match(/<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/g)?.filter((run) => /<w:t(?:\s|>)/.test(run)) ?? [];
  expect(textRuns.length).toBeGreaterThan(1_000);
  // Hidden field instructions and independent drawing/textbox stories are not
  // part of the body story selection; representative visible runs below prove
  // the early/middle/late semantic scope directly.
  expect(textRuns.filter((run) => /<w:b(?:\s|\/|>)/.test(run)).length / textRuns.length).toBeGreaterThan(0.85);
  expect(textRuns.filter((run) => /<w:i(?:\s|\/|>)/.test(run)).length / textRuns.length).toBeGreaterThan(0.85);
  expect(textRuns.filter((run) => /<w:u(?:\s|\/|>)/.test(run)).length / textRuns.length).toBeGreaterThan(0.85);
  const xmlEscape = (value: string) => value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
  for (const text of representativeTexts) {
    const sourceRun = textRuns.find((run) => run.includes(xmlEscape(text)));
    expect(sourceRun, `saved run for representative text ${JSON.stringify(text)}`).toBeDefined();
    expect(sourceRun).toMatch(/<w:b(?:\s|\/|>)/);
    expect(sourceRun).toMatch(/<w:i(?:\s|\/|>)/);
    expect(sourceRun).toMatch(/<w:u(?:\s|\/|>)/);
  }

  await page.locator("#docx-upload").setInputFiles(formattedPath!);
  await expect(pages).toHaveCount(70, { timeout: 60_000 });
  await pages.nth(69).scrollIntoViewIfNeeded();
  const reopenedLate = bodyText(69);
  await reopenedLate.waitFor({ state: "visible" });
  expect(await reopenedLate.evaluate((element) => getComputedStyle(element).fontStyle)).toBe("italic");

  const reopenedEarly = bodyText(0);
  await pages.first().scrollIntoViewIfNeeded();
  await reopenedEarly.click();
  await page.keyboard.press("ControlOrMeta+a");
  await page.keyboard.type("REPLACED", { delay: 5 });
  await expect(page.locator("[data-dxw-layout-status]")).toBeHidden({ timeout: 30_000 });
  await expect(pages).toHaveCount(1, { timeout: 30_000 });
  await expect(pages.first().locator("span:not([data-dxw-hf])", { hasText: "REPLACED" })).toHaveCount(1);

  const replacementDownload = page.waitForEvent("download");
  await page.getByText("Download", { exact: true }).click();
  const replacementPath = await (await replacementDownload).path();
  expect(replacementPath).not.toBeNull();
  const replacementParts = unzipSync(new Uint8Array(readFileSync(replacementPath!)));
  const replacementXml = strFromU8(replacementParts["word/document.xml"]);
  expect(replacementXml).toContain(">REPLACED</w:t>");
  const replacementBody = replacementXml.match(/<w:body>[\s\S]*<\/w:body>/)?.[0] ?? "";
  const replacementTexts = [...replacementBody.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)]
    .map((match) => match[1]);
  expect(replacementTexts.join("")).toBe("REPLACED");

  await page.keyboard.press("ControlOrMeta+z");
  await expect(page.locator("[data-dxw-layout-status]")).toBeHidden({ timeout: 30_000 });
  await expect(pages).toHaveCount(70, { timeout: 60_000 });
  await pages.nth(69).scrollIntoViewIfNeeded();
  const restoredLate = bodyText(69);
  await restoredLate.waitFor({ state: "visible" });
  expect(await restoredLate.evaluate((element) => getComputedStyle(element).fontStyle)).toBe("italic");

  await page.locator("#docx-upload").setInputFiles(replacementPath!);
  await expect(pages).toHaveCount(1, { timeout: 60_000 });
  await expect(pages.first().locator("span:not([data-dxw-hf])", { hasText: "REPLACED" })).toHaveCount(1);
});

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
  const undoStarted = performance.now();
  await page.keyboard.press("ControlOrMeta+z");
  const layoutStatus = page.locator("[data-dxw-layout-status]");
  await expect(pages.first()).not.toContainText("FIRSTEDIT");
  await expect(page.locator("[data-dxw-caret]")).toBeVisible();
  await expect(pages).toHaveCount(419);
  await expect(layoutStatus).toHaveCount(0);
  const undoMs = performance.now() - undoStarted;
  const responsiveness = await page.evaluate(() => {
    const state = globalThis as typeof globalThis & {
      __dxwLayoutTicks?: number;
      __dxwLayoutTimer?: number;
      __dxwLayoutMaxGap?: number;
    };
    clearInterval(state.__dxwLayoutTimer);
    return { ticks: state.__dxwLayoutTicks ?? 0, maxGap: state.__dxwLayoutMaxGap ?? Infinity };
  });
  console.log(`[text-undo] settled=${undoMs.toFixed(1)}ms timer ticks=${responsiveness.ticks} maxGap=${responsiveness.maxGap.toFixed(1)}ms`);
  expect(undoMs).toBeLessThan(250);
  expect(responsiveness.maxGap).toBeLessThan(100);

  // A direct body paragraph split stays on the incremental path. Record both
  // the editor's synchronous sample and whether the background busy UI ever
  // appeared; the old full refresh took roughly 2.5 seconds here.
  const enterTarget = pages.nth(209);
  await enterTarget.scrollIntoViewIfNeeded();
  const enterCandidates = enterTarget.locator("span:not([data-dxw-hf])", { hasText: /[A-Za-z]{4,}/ });
  const enterSpan = enterCandidates.nth(20);
  await enterSpan.waitFor({ state: "visible" });
  await enterSpan.click();
  await expect(page.locator("[data-dxw-caret]").locator("xpath=ancestor::*[@data-page][1]"))
    .toHaveAttribute("data-page", "210");
  await page.keyboard.press("End");
  await page.evaluate(() => {
    const state = globalThis as typeof globalThis & {
      __dxwPerf?: { samples?: unknown[] };
      __dxwSawLayoutBusy?: boolean;
      __dxwBusyTimer?: number;
    };
    if (state.__dxwPerf) state.__dxwPerf.samples = [];
    state.__dxwSawLayoutBusy = false;
    state.__dxwBusyTimer = window.setInterval(() => {
      state.__dxwSawLayoutBusy ||= !!document.querySelector("[data-dxw-layout-busy]");
    }, 1);
  });
  const enterStarted = performance.now();
  await page.keyboard.press("Enter");
  const enterWallMs = performance.now() - enterStarted;
  await expect(page.locator("[data-dxw-caret]")).toBeVisible();
  await expect(pages).toHaveCount(419);
  await page.waitForTimeout(100);
  const enter = await page.evaluate(() => {
    const state = globalThis as typeof globalThis & {
      __dxwPerf?: {
        samples?: { total: number; layout: number; pagesReused: number; totalPages: number }[];
        incr?: {
          hintFastPath: boolean;
          blocksHashed: number;
          firstDirty: number;
          resumeBlock: number;
          resumePage: number;
          convergedBlock: number;
          convergedPage: number;
          blocksLaid: number;
        };
      };
      __dxwSawLayoutBusy?: boolean;
      __dxwBusyTimer?: number;
    };
    clearInterval(state.__dxwBusyTimer);
    return {
      sample: state.__dxwPerf?.samples?.at(-1),
      incremental: state.__dxwPerf?.incr,
      sawBusy: state.__dxwSawLayoutBusy ?? false,
    };
  });
  console.log(`[nih-enter] wall=${enterWallMs.toFixed(1)}ms sample=${JSON.stringify(enter.sample)} incremental=${JSON.stringify(enter.incremental)} busy=${enter.sawBusy}`);
  expect(enterWallMs).toBeLessThan(250);
  expect(enter.sample?.total).toBeLessThan(80);
  expect(enter.sample?.layout).toBeLessThan(80);
  expect(enter.sample?.totalPages).toBe(419);
  expect(enter.sample?.pagesReused).toBeGreaterThanOrEqual(400);
  expect(enter.incremental?.hintFastPath).toBe(true);
  expect(enter.incremental?.blocksHashed).toBeLessThanOrEqual(4);
  expect(enter.incremental?.firstDirty).toBeGreaterThan(2_000);
  expect(enter.incremental?.resumePage).toBeGreaterThan(200);
  expect(enter.incremental?.convergedBlock).toBeGreaterThan(enter.incremental?.firstDirty ?? Infinity);
  expect(enter.incremental?.convergedPage).toBeGreaterThan(enter.incremental?.resumePage ?? Infinity);
  expect(enter.incremental?.blocksLaid).toBeGreaterThan(100);
  expect(enter.incremental?.blocksLaid).toBeLessThanOrEqual(256);
  expect(enter.sawBusy).toBe(false);
  await expect(layoutStatus).toHaveCount(0);

  // Undo restores the pre-split XML snapshot, which intentionally keeps the
  // existing cooperative full-refresh fallback.
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
  const structuralUndoStarted = performance.now();
  await page.keyboard.press("ControlOrMeta+z");
  await expect(layoutStatus).toBeVisible();
  await expect(page.locator("[data-dxw-layout-busy]")).toHaveAttribute("aria-busy", "true");
  await expect(layoutStatus).toBeHidden({ timeout: 30_000 });
  const structuralUndoMs = performance.now() - structuralUndoStarted;
  const structuralResponsiveness = await page.evaluate(() => {
    const state = globalThis as typeof globalThis & {
      __dxwLayoutTicks?: number;
      __dxwLayoutTimer?: number;
      __dxwLayoutMaxGap?: number;
    };
    clearInterval(state.__dxwLayoutTimer);
    return { ticks: state.__dxwLayoutTicks ?? 0, maxGap: state.__dxwLayoutMaxGap ?? Infinity };
  });
  console.log(`[book-layout] structuralUndo=${structuralUndoMs.toFixed(1)}ms timer ticks=${structuralResponsiveness.ticks} maxGap=${structuralResponsiveness.maxGap.toFixed(1)}ms`);
  expect(structuralUndoMs).toBeLessThan(4_000);
  expect(structuralResponsiveness.ticks).toBeGreaterThan(5);
  expect(structuralResponsiveness.maxGap).toBeLessThan(500);
  await expect(pages).toHaveCount(419);

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

  const caret = page.locator("[data-dxw-caret]");
  const container = page.locator(".dxw-pages").locator("..");
  const caretIsInViewport = async () => {
    const [cr, vr] = await Promise.all([caret.boundingBox(), container.boundingBox()]);
    return !!cr && !!vr && cr.y >= vr.y && cr.y + cr.height <= vr.y + vr.height;
  };
  await page.keyboard.press("ControlOrMeta+ArrowUp");
  const upPage = Number(await caret.evaluate((el) => el.closest(".dxw-page")?.getAttribute("data-page")));
  expect(upPage).toBeGreaterThan(1);
  expect(upPage).toBeLessThan(419);
  expect(await caretIsInViewport()).toBe(true);
  await page.keyboard.press("ControlOrMeta+ArrowDown");
  const downPage = Number(await caret.evaluate((el) => el.closest(".dxw-page")?.getAttribute("data-page")));
  expect(downPage).toBeGreaterThan(1);
  expect(downPage).toBeLessThan(419);
  expect(await caretIsInViewport()).toBe(true);

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
