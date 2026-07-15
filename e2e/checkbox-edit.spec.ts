import { test, expect, Page } from "@playwright/test";

// probe2-form-checkboxes has legacy FORMCHECKBOX form fields and modern
// w14:checkbox content controls, both outside and inside a table. All render
// as a ballot glyph (☐ / ☒); clicking one toggles its state.

const MOD = process.platform === "darwin" ? "Meta" : "Control";

async function open(page: Page): Promise<void> {
  await page.goto(`/?doc=/fixtures/probe2-form-checkboxes.docx`);
  await page.waitForSelector(".dxw-page span");
  await page.waitForTimeout(300);
}

/** All checkbox glyph spans, top-to-bottom, with their glyph + pointer style. */
async function boxes(page: Page) {
  return page.evaluate(() => {
    const els = [...document.querySelectorAll<HTMLElement>("span[data-dxw-checkbox]")];
    return els
      .map((s) => {
        const r = s.getBoundingClientRect();
        return { text: s.textContent ?? "", cursor: s.style.cursor, cx: r.x + r.width / 2, cy: r.y + r.height / 2 };
      })
      .sort((a, b) => a.cy - b.cy);
  });
}

async function glyphCounts(page: Page): Promise<{ checked: number; unchecked: number }> {
  const b = await boxes(page);
  return { checked: b.filter((x) => x.text === "☒").length, unchecked: b.filter((x) => x.text === "☐").length };
}

test.describe("checkbox editing in the browser", () => {
  test("legacy + modern checkboxes render as pointer-cursor glyphs, inside and outside tables", async ({ page }) => {
    await open(page);
    const b = await boxes(page);
    // 5 legacy (3 outside + 2 in table) + 4 modern (2 outside + 2 in table).
    expect(b.length).toBe(9);
    expect(b.every((x) => x.cursor === "pointer")).toBe(true);
    // The "Form fields inside table cells" heading sits above the table; some
    // checkboxes render below it (inside the table).
    const headingBox = (await page.locator(".dxw-page span:text-is('Item')").first().boundingBox())!;
    expect(b.some((x) => x.cy > headingBox.y)).toBe(true);
  });

  test("clicking a checkbox toggles its glyph, and undo restores it", async ({ page }) => {
    await open(page);
    const before = await glyphCounts(page);
    // Click the first checked box (☒) → becomes unchecked (☐).
    const first = (await boxes(page)).find((x) => x.text === "☒")!;
    await page.mouse.click(first.cx, first.cy);
    await page.waitForTimeout(200);
    let now = await glyphCounts(page);
    expect(now.checked).toBe(before.checked - 1);
    expect(now.unchecked).toBe(before.unchecked + 1);
    // Undo restores.
    await page.keyboard.press(`${MOD}+z`);
    await page.waitForTimeout(250);
    now = await glyphCounts(page);
    expect(now).toEqual(before);
  });

  test("toggling a checkbox inside a table cell is stable (table does not move)", async ({ page }) => {
    await open(page);
    // Anchor: the table row label "Approved" — its position must not shift.
    const approvedBefore = (await page.locator(".dxw-page span:text-is('Approved')").first().boundingBox())!;
    // Pick a checkbox below the table header (inside the table body).
    const headingBox = (await page.locator(".dxw-page span:text-is('Item')").first().boundingBox())!;
    const tableBox = (await boxes(page)).find((x) => x.cy > headingBox.y)!;
    const wasChecked = tableBox.text === "☒";
    await page.mouse.click(tableBox.cx, tableBox.cy);
    await page.waitForTimeout(200);
    // A checkbox glyph still sits at ~that cell, flipped.
    const after = (await boxes(page)).find((x) => Math.abs(x.cy - tableBox.cy) < 6 && Math.abs(x.cx - tableBox.cx) < 6);
    expect(after).toBeTruthy();
    expect(after!.text).toBe(wasChecked ? "☐" : "☒");
    // Table row label did not jump.
    const approvedAfter = (await page.locator(".dxw-page span:text-is('Approved')").first().boundingBox())!;
    expect(Math.abs(approvedAfter.y - approvedBefore.y)).toBeLessThan(2);
    expect(Math.abs(approvedAfter.x - approvedBefore.x)).toBeLessThan(2);
  });

  test("a modern glyph resists text edits: arrow into it, Backspace/typing are no-ops", async ({ page }) => {
    await open(page);
    const before = await page.evaluate(() =>
      [...document.querySelectorAll<HTMLElement>("span[data-dxw-checkbox]")].map((x) => x.textContent).join(""),
    );
    // Find the modern checkbox glyph that sits immediately left of the
    // " Option A" label, then place a caret just to its right and step left
    // onto the glyph run.
    const glyph = await page.evaluate(() => {
      const spans = [...document.querySelectorAll<HTMLElement>(".dxw-page span")];
      const label = spans.find((s) => (s.textContent ?? "").includes("Option"))!;
      const lr = label.getBoundingClientRect();
      const box = [...document.querySelectorAll<HTMLElement>("span[data-dxw-checkbox]")]
        .map((s) => ({ s, r: s.getBoundingClientRect() }))
        .filter(({ r }) => Math.abs(r.y - lr.y) < 6 && r.right <= lr.left + 4)
        .sort((a, b) => b.r.right - a.r.right)[0];
      const r = box.r;
      return { rx: r.right, cy: r.y + r.height / 2 };
    });
    await page.mouse.click(glyph.rx + 3, glyph.cy);
    await page.waitForTimeout(120);
    await page.keyboard.press("ArrowLeft");
    await page.waitForTimeout(60);
    await page.keyboard.press("Backspace");
    await page.keyboard.type("Z");
    await page.waitForTimeout(200);
    const after = await page.evaluate(() =>
      [...document.querySelectorAll<HTMLElement>("span[data-dxw-checkbox]")].map((x) => x.textContent).join(""),
    );
    expect(after).toBe(before); // glyphs uncorrupted (no deletion, no "Z" merged in)
    // No stray checkbox glyph contains extra characters.
    const clean = await page.evaluate(() =>
      [...document.querySelectorAll<HTMLElement>("span[data-dxw-checkbox]")].every((x) => (x.textContent ?? "").length === 1),
    );
    expect(clean).toBe(true);
  });

  test("clicking a checkbox does not drop a text caret into it", async ({ page }) => {
    await open(page);
    const first = (await boxes(page))[0];
    await page.mouse.click(first.cx, first.cy);
    await page.waitForTimeout(150);
    // The blinking caret element should not be showing (a toggle is not a
    // text-editing click).
    const caretVisible = await page.evaluate(() => {
      const c = document.querySelector<HTMLElement>("[data-dxw-caret]");
      return !!c && c.style.display !== "none" && c.style.visibility !== "hidden";
    });
    expect(caretVisible).toBe(false);
  });
});
