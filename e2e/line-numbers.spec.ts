import { expect, test, Page } from "@playwright/test";

/**
 * Margin line-number editing (Layout > Line numbers). Toggling the control
 * mutates w:sectPr/w:lnNumType and relayouts, so numbers appear/disappear in
 * the left margin. Driven through the real toolbar on the sample fixture
 * (which ships without line numbers).
 */

async function open(page: Page): Promise<void> {
  await page.goto(`/?doc=/fixtures/sample.docx`);
  await page.waitForSelector(".dxw-page span");
  await page.waitForTimeout(500);
  // Switch to the Layout ribbon tab.
  await page.getByRole("button", { name: /layout/i }).first().click();
  await page.waitForTimeout(150);
}

/** Count numeric spans sitting in the left margin (line numbers). */
async function marginNumbers(page: Page): Promise<number> {
  return page.evaluate(() => {
    const pr = document.querySelector(".dxw-page")!.getBoundingClientRect();
    let n = 0;
    for (const el of document.querySelectorAll(".dxw-page span")) {
      if (el.children.length) continue;
      const t = (el.textContent || "").trim();
      const r = el.getBoundingClientRect();
      if (/^\d+$/.test(t) && r.x - pr.x < 60) n++;
    }
    return n;
  });
}

test.describe("line-number editing", () => {
  test("toggling line numbers on then off adds and removes margin numbers", async ({ page }) => {
    await open(page);
    expect(await marginNumbers(page)).toBe(0);

    await page.selectOption('select[title="Line numbers"]', "continuous");
    await page.waitForTimeout(400);
    const on = await marginNumbers(page);
    expect(on, "line numbers appear in the left margin").toBeGreaterThan(10);

    await page.selectOption('select[title="Line numbers"]', "off");
    await page.waitForTimeout(400);
    expect(await marginNumbers(page), "line numbers removed").toBe(0);
  });

  test("count-by-5 numbers only every fifth line", async ({ page }) => {
    await open(page);
    await page.selectOption('select[title="Line numbers"]', "continuous");
    await page.waitForTimeout(400);
    const all = await marginNumbers(page);
    await page.selectOption('select[title="Line numbers"]', "by5");
    await page.waitForTimeout(400);
    const by5 = await marginNumbers(page);
    // Every-5th prints far fewer numbers than every line.
    expect(by5).toBeGreaterThan(0);
    expect(by5).toBeLessThan(all / 2);
  });
});
