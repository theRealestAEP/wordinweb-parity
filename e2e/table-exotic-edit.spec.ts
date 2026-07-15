import { test, expect, Page } from "@playwright/test";

// probe3-table-exotics has two page-anchored floating tables (Float A/B, each
// 2x2) and an old-style separated-border table (w:tblCellSpacing, "cs" cells).
// These specs prove the exotic tables are editable the same way inline ones
// are: caret + typing, and the Table menu row/column ops.

async function open(page: Page): Promise<void> {
  await page.goto(`/?doc=/fixtures/probe3-table-exotics.docx`);
  await page.waitForSelector(".dxw-page span");
  await page.waitForTimeout(300);
}

/** Bounding box (viewport px) of the Nth span whose text === `exact`. */
async function spanBox(page: Page, exact: string, nth = 0) {
  return page.evaluate(
    ({ exact, nth }) => {
      const els = [...document.querySelectorAll(".dxw-page span")].filter(
        (s) => (s.textContent ?? "") === exact,
      );
      const s = els[nth] as HTMLElement | undefined;
      if (!s) return null;
      const r = s.getBoundingClientRect();
      return { x: r.x, y: r.y, w: r.width, h: r.height, cx: r.x + r.width / 2, cy: r.y + r.height / 2 };
    },
    { exact, nth },
  );
}

async function countSpans(page: Page, exact: string): Promise<number> {
  return page.evaluate(
    (exact) =>
      [...document.querySelectorAll(".dxw-page span")].filter((s) => (s.textContent ?? "") === exact).length,
    exact,
  );
}

/** Interactive resize grips, split by axis via their cursor style. */
async function gripCounts(page: Page): Promise<{ col: number; row: number }> {
  return page.evaluate(() => {
    const grips = [...document.querySelectorAll<HTMLElement>("[data-dxw-grip]")];
    return {
      col: grips.filter((g) => g.style.cursor === "col-resize").length,
      row: grips.filter((g) => g.style.cursor === "row-resize").length,
    };
  });
}

async function openTableMenuOp(page: Page, label: string): Promise<void> {
  await page.locator('button[data-tab="insert"]').click();
  await page.waitForTimeout(80);
  await page.locator("button[title='Table']").click();
  await page.locator("div", { hasText: new RegExp(`^${label}$`) }).first().click();
  await page.waitForTimeout(300);
}

test.describe("editing exotic tables in the browser", () => {
  test("typing into a floating cell persists and the table does not jump", async ({ page }) => {
    await open(page);
    // Anchor: Float A's top-left cell ("r1c1"). Record its position, click at
    // the END of the word, and append text — the table must stay put.
    const before = await spanBox(page, "r1c1", 0);
    expect(before, "Float A r1c1 span not found").not.toBeNull();
    await page.mouse.click(before!.x + before!.w - 3, before!.cy);
    await page.waitForTimeout(120);
    await page.keyboard.type("ZZ");
    await page.waitForTimeout(200);

    const hasTyped = await page.evaluate(() =>
      [...document.querySelectorAll(".dxw-page span")].some((s) => (s.textContent ?? "").includes("r1c1ZZ")),
    );
    expect(hasTyped, "typed text should be inside the floating cell").toBe(true);

    // The floating anchor stays put: the cell's top is unchanged (short text
    // does not wrap, so the row keeps its height and the table its position).
    const after = await spanBox(page, "r1c1ZZ", 0);
    expect(after).not.toBeNull();
    expect(Math.abs(after!.y - before!.y)).toBeLessThan(2);
    expect(Math.abs(after!.x - before!.x)).toBeLessThan(2);
  });

  test("adding a row to the cell-spacing table adds a row boundary", async ({ page }) => {
    await open(page);
    const before = await gripCounts(page);
    const floatsBefore = await countSpans(page, "r1c1");
    const cs = await spanBox(page, "cs", 0);
    await page.mouse.click(cs!.cx, cs!.cy);
    await page.waitForTimeout(120);
    await openTableMenuOp(page, "Insert row below");
    const after = await gripCounts(page);
    // Row grips are emitted once per table row (independent of how flow tables
    // paginate), so the new row shows up as exactly one extra row-resize grip.
    expect(after.row).toBe(before.row + 1);
    // The floating tables are untouched by an edit elsewhere in the flow.
    expect(await countSpans(page, "r1c1")).toBe(floatsBefore);
  });

  test("deleting a column in a floating table drops that column's cells", async ({ page }) => {
    await open(page);
    // Both floats carry an "r1c1" first-column cell (2 spans total).
    const r1c1Before = await countSpans(page, "r1c1");
    const cell = await spanBox(page, "r1c1", 0);
    await page.mouse.click(cell!.cx, cell!.cy);
    await page.waitForTimeout(120);
    await openTableMenuOp(page, "Delete column");
    // Float A's first column is gone: one fewer "r1c1" span. Its neighbour
    // column ("r1c2") and the second float are untouched.
    expect(await countSpans(page, "r1c1")).toBe(r1c1Before - 1);
    expect(await countSpans(page, "r1c2")).toBeGreaterThan(0);
  });
});
