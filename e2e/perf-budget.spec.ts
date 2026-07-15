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
};

function median(xs: number[]): number {
  const s = [...xs].sort((a, b) => a - b);
  return s[Math.floor(s.length / 2)];
}

/** Open a fixture with perf recording armed before any app script runs, place
 * the caret near the vertical middle of a mid-document page, and type `n`
 * single characters. Returns the recorded samples (one per keystroke commit). */
async function measure(page: Page, fixture: string, n: number): Promise<PerfSample[]> {
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
  const span = targetPage.locator("span", { hasText: /\w{3,}/ }).first();
  await span.waitFor({ state: "visible" });
  await span.click();
  await page.waitForTimeout(150);
  await page.keyboard.press("End");
  await page.waitForTimeout(50);

  // Reset samples after the click (the click itself may not record, but any
  // warm-up render should not skew the median).
  await page.evaluate(() => {
    const p = (globalThis as { __dxwPerf?: { samples?: unknown[] } }).__dxwPerf;
    if (p) p.samples = [];
  });

  for (let i = 0; i < n; i++) {
    await page.keyboard.type("x");
    await page.waitForTimeout(30);
  }
  await page.waitForTimeout(150);

  return page.evaluate(() => {
    const p = (globalThis as { __dxwPerf?: { samples?: PerfSample[] } }).__dxwPerf;
    return p?.samples ?? [];
  });
}

function report(name: string, s: PerfSample[]): string {
  const med = (k: keyof PerfSample) => median(s.map((x) => x[k])).toFixed(1);
  return `[perf] ${name}: n=${s.length} total=${med("total")} layout=${med("layout")} `
    + `render=${med("render")} destroy=${med("destroy")} refresh=${med("refresh")} `
    + `chrome=${med("chromeCaret")} reused=${median(s.map((x) => x.pagesReused))}/${s[0]?.totalPages}`;
}

const CASES: { fixture: string; budgetMs: number; firstBudgetMs?: number }[] = [
  // Quiet-machine targets ~ dense 50, image-stress 75, comments 40, NIH 254 ->
  // ~2x budgets so each scenario is pinned without flaking on a busy box.
  // dense-comments guards the comments-reuse regression class: a commented doc
  // used to rebuild every page per keystroke (render ~80ms); it now adopts
  // pages and runs the comment overlay per keystroke (a primary editing case,
  // the user's real lag doc was a commented legal letter).
  { fixture: "dense-skewtest", budgetMs: 120 },
  { fixture: "dense-imagestress", budgetMs: 150 },
  { fixture: "dense-comments", budgetMs: 120 },
  { fixture: "wild2-legal-nih-contract", budgetMs: 400, firstBudgetMs: 400 },
];

for (const { fixture, budgetMs, firstBudgetMs } of CASES) {
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
    if (firstBudgetMs !== undefined) {
      expect(samples[0].total, `first keystroke total for ${fixture}`).toBeLessThan(firstBudgetMs);
    }
  });
}
