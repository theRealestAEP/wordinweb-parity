import { expect, test, Page } from "@playwright/test";
import { Editor } from "./editing.js";

/**
 * Typing a space at the end of a table cell's text. The trailing space hangs
 * invisibly and carries caretClampX (the cell content edge) so the caret pins
 * there — but cell items are laid out cell-locally and translated to page
 * space by offsetItem, which shifted x and not caretClampX, so the clamp
 * stayed cell-local and min(pageX, clamp) flung the caret to the left of the
 * table on every space.
 */

async function caretInfo(page: Page): Promise<{ x: number; y: number } | null> {
  return page.evaluate(() => {
    const caret = document.querySelector('[data-dxw-caret="1"]') as HTMLElement | null;
    if (!caret || caret.style.display === "none") return null;
    const pr = document.querySelector(".dxw-page")!.getBoundingClientRect();
    const r = caret.getBoundingClientRect();
    return { x: r.x - pr.x, y: r.y - pr.y };
  });
}

test.describe("space at the end of a table cell", () => {
  test("caret advances by one space width, on the same line", async ({ page }) => {
    const ed = await Editor.open(page, "parity-tables");
    const b = (await ed.span("Status").boundingBox())!;
    await page.mouse.click(b.x + b.width - 2, b.y + b.height / 2);
    await page.waitForTimeout(150);
    const before = await caretInfo(page);
    expect(before).not.toBeNull();
    await page.keyboard.type(" ");
    await page.waitForTimeout(200);
    const after = await caretInfo(page);
    expect(after).not.toBeNull();
    expect(Math.abs(after!.y - before!.y)).toBeLessThan(3);
    const dx = after!.x - before!.x;
    expect(dx, "caret flew away from the typed space").toBeGreaterThan(0);
    expect(dx).toBeLessThan(15);
  });

  test("many trailing spaces pin the caret at the cell edge, never the neighbor cell", async ({ page }) => {
    const ed = await Editor.open(page, "parity-tables");
    const cell = (await ed.span("Status").boundingBox())!;
    // The next column's text gives us the neighbor cell's left edge.
    const pr = (await page.locator(".dxw-page").first().boundingBox())!;
    await page.mouse.click(cell.x + cell.width - 2, cell.y + cell.height / 2);
    await page.waitForTimeout(150);
    const start = await caretInfo(page);
    expect(start).not.toBeNull();
    let prev = start!;
    for (let i = 0; i < 8; i++) {
      await page.keyboard.type(" ");
      await page.waitForTimeout(80);
      const c = await caretInfo(page);
      expect(c, `caret missing after space ${i + 1}`).not.toBeNull();
      // Same line, monotonically non-decreasing x, never jumping left.
      expect(Math.abs(c!.y - prev.y)).toBeLessThan(3);
      expect(c!.x).toBeGreaterThanOrEqual(prev.x - 0.5);
      prev = c!;
    }
    // Clamped: after 8 spaces the caret must not have crossed into the
    // neighbor column (the "Description" header sits to the right).
    const neighbor = (await ed.span("Description").boundingBox())!;
    expect(prev.x + pr.x).toBeLessThan(neighbor.x + 1);
  });
});

test.describe("typing growth in an autofit table", () => {
  // A long unbroken word typed into a cell: the column autofits wider (Word
  // semantics) but the TABLE must pin at the page content width and the word
  // must character-wrap downward — it once grew past the page edge without
  // bound (min-content clamp escaping the available-width cap), with the
  // caret abandoned at the start of a phantom 1-char second line.
  test("a long unbroken word never grows the table past the page edge and the caret rides it", async ({ page }) => {
    const ed = await Editor.open(page, "parity-tables");
    const b = (await ed.span("ok").boundingBox())!;
    await page.mouse.click(b.x + b.width - 1, b.y + b.height / 2);
    await page.waitForTimeout(150);
    let prevCaretX = -1;
    for (let i = 0; i < 10; i++) {
      await page.keyboard.type("mmmmmmmmmm", { delay: 5 });
      await page.waitForTimeout(150);
      const g = await page.evaluate(() => {
        const pg = document.querySelector(".dxw-page")!;
        const pr = pg.getBoundingClientRect();
        let maxX = -1e9;
        for (const el of pg.querySelectorAll("span")) {
          if (!(el.textContent ?? "").trim()) continue;
          const r = el.getBoundingClientRect();
          maxX = Math.max(maxX, r.x + r.width - pr.x);
        }
        const caret = document.querySelector('[data-dxw-caret="1"]') as HTMLElement | null;
        const cr = caret?.getBoundingClientRect();
        return { maxX, pageW: pr.width, caretX: cr ? cr.x - pr.x : null };
      });
      expect(g.maxX, `text ran past the page edge after ${(i + 1) * 10} chars`).toBeLessThan(g.pageW);
      expect(g.caretX).not.toBeNull();
      // The caret must move between bursts: either riding the word right, or
      // wrapping to a fresh line — never frozen at a stale spot.
      expect(g.caretX, `caret frozen after ${(i + 1) * 10} chars`).not.toBe(prevCaretX);
      prevCaretX = g.caretX!;
    }
  });
});
