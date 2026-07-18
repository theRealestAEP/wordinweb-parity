import { test, expect, Page } from "@playwright/test";
import { existsSync } from "node:fs";
import { join } from "node:path";

const FIX_DIR = join(__dirname, "../apps/demo/public/fixtures");

/**
 * Per-keystroke performance budgets. Drives a real mid-document
 * typing burst in the headless browser and reads the __dxwPerf.samples the core
 * records (editor.commit + renderToDom). Budgets are deliberately generous —
 * ~2x the quiet-machine targets — so they pin the SCENARIO against regressions
 * (a commented doc rebuilding every page, a stress fixture breaking reuse)
 * without flaking on a busy CI box. The per-stage breakdown is logged so a
 * failure shows WHERE the time went, the way a visual-diff failure shows where.
 */

type PerfSample = {
  total: number; layout: number; render: number; destroy: number;
  refresh: number; chromeCaret: number; totalPages: number; pagesReused: number;
  wall: number;
  blocksLaid?: number;
  resumePage?: number;
  convergedPage?: number;
  fallbackReason?: string;
};

function median(xs: number[]): number {
  const s = [...xs].sort((a, b) => a - b);
  return s[Math.floor(s.length / 2)];
}

function percentile(xs: number[], p: number): number {
  const s = [...xs].sort((a, b) => a - b);
  return s[Math.min(s.length - 1, Math.ceil(s.length * p) - 1)];
}

/** Open a fixture with perf recording armed before any app script runs, place
 * the caret near the vertical middle of a mid-document page, and type `n`
 * single characters. Returns the recorded samples (one per keystroke commit). */
async function measure(page: Page, fixture: string, n: number, targetOrdinal = 0): Promise<PerfSample[]> {
  await page.addInitScript(() => {
    (globalThis as { __dxwPerf?: unknown }).__dxwPerf = { samples: [] };
  });
  await page.goto(`/?doc=/fixtures/${fixture}.docx`);
  // Wait for pages to attach (a heavy doc's spans can render outside Playwright's
  // "visible" heuristic under a page transform, so wait on the page element, not
  // a visible span), then let layout settle.
  await page.locator(".dxw-page").first().waitFor({ state: "attached", timeout: 60000 });
  await expect.poll(() => page.locator(".dxw-page").count(), { timeout: 60000 }).toBeGreaterThan(0);
  await page.waitForTimeout(600);

  // Place the caret on real text in a mid-document page so the edit is neither
  // the first nor last page — the realistic "typing in the middle of a long
  // doc" case. Clicking an actual span (not a blank fraction) guarantees a
  // caret; typing into whitespace can no-op.
  const pages = page.locator(".dxw-page");
  const count = await pages.count();
  const midIdx = Math.floor(count / 2);
  const targetPage = pages.nth(midIdx);
  await targetPage.scrollIntoViewIfNeeded();
  const span = targetPage.locator("span:not([data-dxw-hf])", { hasText: /[A-Za-z]{4,}/ }).nth(targetOrdinal);
  await span.waitFor({ state: "visible" });
  await expect(span).toHaveText(/[A-Za-z]{4,}/);
  await span.click();
  await expect(page.locator("[data-dxw-caret]").locator("xpath=ancestor::*[@data-page][1]"))
    .toHaveAttribute("data-page", await targetPage.getAttribute("data-page") ?? "");
  await page.waitForTimeout(150);
  await page.keyboard.press("End");
  await page.waitForTimeout(50);

  // Reset samples after the click (the click itself may not record, but any
  // warm-up render should not skew the median).
  await page.evaluate(() => {
    const p = (globalThis as { __dxwPerf?: { samples?: unknown[] } }).__dxwPerf;
    if (p) p.samples = [];
  });

  const samples: PerfSample[] = [];
  for (let i = 0; i < n; i++) {
    const started = performance.now();
    await page.keyboard.type("x");
    const wall = performance.now() - started;
    const sample = await page.evaluate(() => {
      const p = (globalThis as {
        __dxwPerf?: {
          samples?: PerfSample[];
          incr?: { blocksLaid?: number; resumePage?: number; convergedPage?: number; fallbackReason?: string };
        };
      }).__dxwPerf;
      const latest = p?.samples?.at(-1);
      return latest ? { ...latest, ...p?.incr } : null;
    });
    if (sample) samples.push({ ...sample, wall });
    await page.waitForTimeout(30);
  }
  await page.waitForTimeout(150);
  return samples;
}

function report(name: string, s: PerfSample[]): string {
  const med = (k: keyof PerfSample) => median(s.map((x) => x[k])).toFixed(1);
  const totals = s.map((sample) => sample.total);
  return `[perf] ${name}: n=${s.length} total first=${totals[0].toFixed(1)} median=${med("total")} `
    + `p95=${percentile(totals, 0.95).toFixed(1)} max=${Math.max(...totals).toFixed(1)} `
    + `wall p95=${percentile(s.map((sample) => sample.wall), 0.95).toFixed(1)} `
    + `layout=${med("layout")} `
    + `render=${med("render")} destroy=${med("destroy")} refresh=${med("refresh")} `
    + `chrome=${med("chromeCaret")} reused=${median(s.map((x) => x.pagesReused))}/${s[0]?.totalPages}`;
}

const CASES: {
  fixture: string;
  budgetMs: number;
}[] = [
  // Quiet-machine targets ~ dense 50, image-stress 75, comments 40.
  // dense-comments guards the comments-reuse regression class: a commented doc
  // used to rebuild every page per keystroke (render ~80ms); it now adopts
  // pages and runs the comment overlay per keystroke (a primary editing case,
  // the user's real lag doc was a commented legal letter).
  { fixture: "dense-skewtest", budgetMs: 120 },
  { fixture: "dense-imagestress", budgetMs: 150 },
  { fixture: "dense-comments", budgetMs: 120 },
];

for (const { fixture, budgetMs } of CASES) {
  test(`keystroke budget: ${fixture} < ${budgetMs}ms median`, async ({ page }) => {
    test.setTimeout(120_000); // cold vite + a heavy doc's first render can be slow
    // Some stress fixtures are generated locally and are not part of the
    // checked-in set; skip rather than fail where absent.
    test.skip(!existsSync(join(FIX_DIR, `${fixture}.docx`)), `${fixture}.docx not present`);
    const samples = await measure(page, fixture, 15);
    expect(samples.length, "should record keystroke samples").toBeGreaterThan(8);
    const medTotal = median(samples.map((s) => s.total));
    // eslint-disable-next-line no-console
    console.log(report(fixture, samples));
    expect(medTotal, `median keystroke total for ${fixture}`).toBeLessThan(budgetMs);
  });
}

test("NIH body typing keeps wrapping and non-wrapping keystrokes below 80ms", async ({ page }) => {
  test.setTimeout(120_000);
  const fixture = "wild2-legal-nih-contract";
  const samples = await measure(page, fixture, 15, 20);
  expect(samples).toHaveLength(15);

  // eslint-disable-next-line no-console
  console.log(report(fixture, samples));
  // eslint-disable-next-line no-console
  console.log(`[perf-samples] ${JSON.stringify(samples.map((sample, key) => ({
    key: key + 1,
    wall: sample.wall,
    total: sample.total,
    layout: sample.layout,
    reused: sample.pagesReused,
    blocks: sample.blocksLaid,
    resumePage: sample.resumePage,
    convergedPage: sample.convergedPage,
    fallback: sample.fallbackReason,
  })))}`);

  for (const [index, sample] of samples.entries()) {
    expect(sample.totalPages, `page count after key ${index + 1}`).toBe(419);
    expect(sample.total, `editor total after key ${index + 1}`).toBeLessThan(80);
    expect(sample.layout, `layout after key ${index + 1}`).toBeLessThan(80);
    expect(sample.wall, `wall time after key ${index + 1}`).toBeLessThan(80);
    expect(sample.fallbackReason, `incremental fallback after key ${index + 1}`).toBe("");
  }

  const wrapping = samples.filter((sample) => sample.pagesReused < 418);
  const nonWrapping = samples.filter((sample) => sample.pagesReused === 418);
  expect(wrapping.length, "burst must cross a real line/page wrap").toBeGreaterThan(0);
  expect(nonWrapping.length, "burst must include ordinary non-wrapping keys").toBeGreaterThan(8);
  expect(nonWrapping.every((sample) => (sample.blocksLaid ?? Infinity) <= 32)).toBe(true);
  expect(wrapping.every((sample) => sample.pagesReused >= 400)).toBe(true);
  expect(wrapping.every((sample) => (sample.blocksLaid ?? 0) > 32 && (sample.blocksLaid ?? Infinity) <= 256)).toBe(true);
  expect(wrapping.every((sample) => (sample.convergedPage ?? -1) > (sample.resumePage ?? Infinity))).toBe(true);
});

test("NIH page 50 repeated paragraph splits stay responsive", async ({ page }) => {
  test.setTimeout(180_000);
  await page.addInitScript(() => {
    (globalThis as { __dxwPerf?: unknown }).__dxwPerf = { samples: [] };
  });
  await page.goto("/?doc=/fixtures/wild2-legal-nih-contract.docx");

  const pages = page.locator(".dxw-page");
  await pages.first().waitFor({ state: "attached", timeout: 60_000 });
  await expect(pages).toHaveCount(419, { timeout: 60_000 });
  const targetPage = pages.nth(49);
  await targetPage.scrollIntoViewIfNeeded();
  const target = targetPage.locator("span:not([data-dxw-hf])", { hasText: /[A-Za-z]{4,}/ }).nth(20);
  await target.waitFor({ state: "visible" });
  await target.click();
  await page.keyboard.press("End");
  await page.evaluate(() => {
    const perf = (globalThis as { __dxwPerf?: { samples?: unknown[] } }).__dxwPerf;
    if (perf) perf.samples = [];
  });

  const timings: Array<{
    wall: number;
    sample?: PerfSample;
    sampleCount: number;
    pages: number;
    caretPage?: string;
    activeElement?: string;
  }> = [];
  for (let index = 0; index < 6; index++) {
    const started = performance.now();
    await page.keyboard.press("Enter");
    const wall = performance.now() - started;
    await expect(page.locator("[data-dxw-caret]")).toBeVisible();
    const state = await page.evaluate(() => {
      const samples = (globalThis as { __dxwPerf?: { samples?: PerfSample[] } }).__dxwPerf?.samples ?? [];
      const caret = document.querySelector<HTMLElement>("[data-dxw-caret]");
      return {
        sample: samples.at(-1),
        sampleCount: samples.length,
        caretPage: caret?.closest<HTMLElement>(".dxw-page")?.dataset.page,
        activeElement: document.activeElement?.tagName,
      };
    });
    timings.push({ wall, ...state, pages: await pages.count() });
  }

  // eslint-disable-next-line no-console
  console.log(`[nih-page-50-enter] ${JSON.stringify(timings)}`);
  for (const [index, timing] of timings.entries()) {
    expect(timing.wall, `Enter ${index + 1} wall time`).toBeLessThan(250);
    expect(timing.sampleCount, `Enter ${index + 1} must commit`).toBe(index + 1);
    expect(timing.sample?.total, `Enter ${index + 1} editor time`).toBeLessThan(100);
    expect(timing.sample?.layout, `Enter ${index + 1} layout time`).toBeLessThan(100);
    expect(timing.pages, `Enter ${index + 1} page count`).toBeGreaterThanOrEqual(419);
    expect(timing.pages, `Enter ${index + 1} page count`).toBeLessThanOrEqual(421);
  }
});
