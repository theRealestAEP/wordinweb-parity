import { expect, test, Page } from "@playwright/test";

/**
 * Tab-to-next-cell navigation inside tables (Word-familiar). Tab advances to
 * the next cell, Shift+Tab to the previous; Tab never inserts a literal tab
 * character into a cell. Runs against the parity-tables fixture, whose first
 * body table has a "Key | Status | Description ..." header row.
 */

async function open(page: Page, fixture: string): Promise<void> {
  await page.goto(`/?doc=/fixtures/${fixture}.docx`);
  await page.waitForSelector(".dxw-page span");
  await page.waitForTimeout(500);
}

async function caretX(page: Page): Promise<number | null> {
  return page.evaluate(() => {
    const c = [...document.querySelectorAll("div")].find(
      (d) => (d as HTMLElement).style.width === "1.5px" && (d as HTMLElement).style.position === "absolute",
    ) as HTMLElement | undefined;
    if (!c || c.style.display === "none") return null;
    return c.getBoundingClientRect().x;
  });
}

async function clickSpan(page: Page, text: string): Promise<number> {
  const box = await page.locator(`.dxw-page span:text-is(${JSON.stringify(text)})`).first().boundingBox();
  expect(box, `span "${text}" present`).not.toBeNull();
  await page.mouse.click(box!.x + 3, box!.y + box!.height / 2);
  await page.waitForTimeout(120);
  return box!.x;
}

test.describe("table Tab navigation", () => {
  test("Tab advances to the next cell (caret jumps to the next column)", async ({ page }) => {
    await open(page, "parity-tables");
    await clickSpan(page, "Key");
    const start = await caretX(page);
    expect(start).not.toBeNull();
    await page.keyboard.press("Tab");
    await page.waitForTimeout(150);
    const afterTab = await caretX(page);
    expect(afterTab, "caret still visible after Tab").not.toBeNull();
    // Next cell ("Status") sits to the right of "Key".
    expect(afterTab!).toBeGreaterThan(start! + 15);
  });

  test("Tab does not insert a tab character into the cell", async ({ page }) => {
    await open(page, "parity-tables");
    await clickSpan(page, "Key");
    await page.keyboard.press("Tab");
    await page.waitForTimeout(150);
    // The "Key" cell text is unchanged — no tab/whitespace was inserted.
    const keyText = await page.evaluate(
      () => [...document.querySelectorAll(".dxw-page span")].find((s) => s.textContent === "Key")?.textContent ?? null,
    );
    expect(keyText).toBe("Key");
  });

  test("Shift+Tab retreats to the previous cell", async ({ page }) => {
    await open(page, "parity-tables");
    const statusX = await clickSpan(page, "Status");
    await page.keyboard.press("Shift+Tab");
    await page.waitForTimeout(150);
    const afterShiftTab = await caretX(page);
    expect(afterShiftTab, "caret visible after Shift+Tab").not.toBeNull();
    // Previous cell ("Key") sits to the left of "Status".
    expect(afterShiftTab!).toBeLessThan(statusX - 5);
  });
});
