import { expect, test, Page } from "@playwright/test";

/**
 * Footnote editing (user: "how do we edit footnotes?"). Footnote content was
 * render-only in v1 — its text carried no editable source, so a click fell
 * through to the body. These lock in: clicking footnote text places the caret
 * IN the footnote and typing edits it, and double-clicking a reference mark
 * jumps to the note at the page bottom. Runs against parity2-notes.
 */

async function open(page: Page): Promise<void> {
  await page.goto(`/?doc=/fixtures/parity2-notes.docx`);
  await page.waitForSelector(".dxw-page span");
  await page.waitForTimeout(500);
}

/** Caret top in page-space px, or null. */
async function caretY(page: Page): Promise<number | null> {
  return page.evaluate(() => {
    const c = [...document.querySelectorAll("div")].find(
      (d) => (d as HTMLElement).style.width === "1.5px" && (d as HTMLElement).style.position === "absolute" && (d as HTMLElement).style.display === "block",
    ) as HTMLElement | undefined;
    if (!c) return null;
    const pr = document.querySelector(".dxw-page")!.getBoundingClientRect();
    return Math.round(c.getBoundingClientRect().y - pr.top);
  });
}

/** The first footnote body word ("Lorem" on the note line) — client point + page-y. */
async function footnoteWord(page: Page) {
  return page.evaluate(() => {
    const pg = document.querySelector(".dxw-page")!;
    const pr = pg.getBoundingClientRect();
    const spans = [...pg.querySelectorAll("span")].filter((s) => !s.children.length);
    const labels = spans.filter((s) => s.textContent === "Footnote");
    if (!labels.length) return null;
    labels.sort((a, b) => b.getBoundingClientRect().y - a.getBoundingClientRect().y);
    const fy = labels[0].getBoundingClientRect().y;
    const lorem = spans.find((s) => s.textContent === "Lorem" && Math.abs(s.getBoundingClientRect().y - fy) < 3);
    const s = lorem ?? labels[0];
    s.scrollIntoView({ block: "center" });
    const r = s.getBoundingClientRect();
    const pr2 = pg.getBoundingClientRect();
    return { cx: r.x + 3, cy: r.y + r.height / 2, py: Math.round(r.y - pr2.top) };
  });
}

test.describe("footnote editing", () => {
  test("clicking footnote text places the caret in the footnote and typing edits it", async ({ page }) => {
    await open(page);
    const w = await footnoteWord(page);
    expect(w, "footnote body text present").not.toBeNull();
    await page.mouse.click(w!.cx, w!.cy);
    await page.waitForTimeout(150);
    const cy = await caretY(page);
    expect(cy, "caret visible after clicking footnote").not.toBeNull();
    // The caret is on the footnote line (bottom of page), not up in the body.
    expect(Math.abs(cy! - w!.py)).toBeLessThan(18);
    await page.keyboard.type("ZZ");
    await page.waitForTimeout(250);
    const edited = await page.evaluate(() =>
      [...document.querySelectorAll(".dxw-page span")].some((s) => (s.textContent || "").startsWith("ZZ")),
    );
    expect(edited, "typed text appears in the footnote").toBe(true);
  });

  test("double-clicking a reference mark jumps to its note", async ({ page }) => {
    await open(page);
    const ref = await page.evaluate(() => {
      const el = document.querySelector(".dxw-page [data-note-ref]") as HTMLElement | null;
      if (!el) return null;
      const pr = document.querySelector(".dxw-page")!.getBoundingClientRect();
      const r = el.getBoundingClientRect();
      return { cx: r.x + r.width / 2, cy: r.y + r.height / 2, py: Math.round(r.y - pr.top) };
    });
    expect(ref, "a footnote reference mark is present").not.toBeNull();
    await page.mouse.dblclick(ref!.cx, ref!.cy);
    await page.waitForTimeout(250);
    const cy = await caretY(page);
    expect(cy, "caret visible after jump").not.toBeNull();
    // The note sits well below the reference mark.
    expect(cy!).toBeGreaterThan(ref!.py + 100);
  });
});
