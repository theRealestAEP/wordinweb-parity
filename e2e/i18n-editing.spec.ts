import { expect, test, Page } from "@playwright/test";

/**
 * International-script editing invariants (i18n audit). These lock in the
 * fixes for editing in RTL and complex scripts:
 *  - RTL margin/whitespace clicks resolve to the correct LOGICAL end (the
 *    visual right edge of an RTL line is its logical START).
 *  - Arrow keys move LOGICALLY: ArrowRight advances the offset, which in an
 *    RTL run paints the caret leftward (Word's logical caret order).
 *  - Backspace/Delete and shift-selection step by whole grapheme clusters, so
 *    a Devanagari conjunct or Arabic base+harakāt is never half-deleted.
 *  - CJK IME composition previews and commits at the caret.
 * Run against the i18n fixtures so they hold on any machine.
 */

async function open(page: Page, fixture: string): Promise<void> {
  await page.goto(`/?doc=/fixtures/${fixture}.docx`);
  await page.waitForSelector(".dxw-page span");
  await page.waitForTimeout(500);
}

/** Caret left/top in page-space px, or null when hidden. */
async function caret(page: Page): Promise<{ x: number; y: number } | null> {
  return page.evaluate(() => {
    const c = [...document.querySelectorAll("div")].find(
      (d) => (d as HTMLElement).style.width === "1.5px" && (d as HTMLElement).style.position === "absolute",
    ) as HTMLElement | undefined;
    if (!c || c.style.display === "none") return null;
    const pr = document.querySelector(".dxw-page")!.getBoundingClientRect();
    const r = c.getBoundingClientRect();
    return { x: Math.round((r.x - pr.x) * 10) / 10, y: Math.round((r.y - pr.y) * 10) / 10 };
  });
}

/** Client rect of the first span whose exact text equals `text`. */
async function spanRect(page: Page, text: string) {
  return page.evaluate((t) => {
    const pr = document.querySelector(".dxw-page")!.getBoundingClientRect();
    for (const el of document.querySelectorAll(".dxw-page span")) {
      if (el.children.length) continue;
      if ((el.textContent || "") === t) {
        const r = el.getBoundingClientRect();
        return { cx: r.x, cy: r.y + r.height / 2, w: r.width, left: (r.x - pr.x), right: (r.right - pr.x) };
      }
    }
    return null;
  }, text);
}

test.describe("i18n editing", () => {
  test("RTL: right-margin click snaps to the line's logical start, not the wrong word end", async ({ page }) => {
    await open(page, "wild2-lit-yiddish-rtl");
    // "שלום" is the rightmost (logical-first) word on its line.
    const sr = await spanRect(page, "שלום");
    expect(sr, "shalom span present").not.toBeNull();
    // Click far into the right margin, past all text on that line.
    await page.mouse.click(sr!.cx + sr!.w + 250, sr!.cy);
    await page.waitForTimeout(120);
    const c = await caret(page);
    expect(c, "caret visible after margin click").not.toBeNull();
    // Logical start of the line = visual RIGHT edge of the first RTL word.
    expect(Math.abs(c!.x - sr!.right)).toBeLessThan(6);
  });

  test("RTL: ArrowRight moves the caret logically forward (leftward on screen)", async ({ page }) => {
    await open(page, "wild2-lit-yiddish-rtl");
    const sr = await spanRect(page, "שלום");
    // Start at the logical start (visual right edge) of the RTL word.
    await page.mouse.click(sr!.cx + sr!.w - 3, sr!.cy);
    await page.waitForTimeout(120);
    const start = await caret(page);
    await page.keyboard.press("ArrowRight");
    await page.waitForTimeout(100);
    const after = await caret(page);
    expect(start && after).toBeTruthy();
    // Logical-forward in RTL paints LEFTWARD.
    expect(after!.x).toBeLessThan(start!.x);
  });

  test("Indic: Backspace deletes a whole grapheme cluster (conjunct + matra)", async ({ page }) => {
    await open(page, "probe3-indic");
    // Locate a Devanagari word of >= 3 code units (has a conjunct/matra).
    const info = await page.evaluate(() => {
      for (const el of document.querySelectorAll(".dxw-page span")) {
        if (el.children.length) continue;
        const s = el.textContent || "";
        if (/[ऀ-ॿ]/.test(s) && s.length >= 3) {
          const r = el.getBoundingClientRect();
          return { text: s, len: s.length, rightClient: r.right, cy: r.y + r.height / 2 };
        }
      }
      return null;
    });
    expect(info, "a Devanagari word is present").not.toBeNull();
    const clusters = [...new Intl.Segmenter("hi", { granularity: "grapheme" }).segment(info!.text)];
    // The last cluster must be more than one code unit for this test to be meaningful.
    const lastLen = clusters[clusters.length - 1].segment.length;
    expect(lastLen).toBeGreaterThan(1);
    // Caret at the word's logical end (LTR: visual right edge), then Backspace.
    await page.mouse.click(info!.rightClient - 2, info!.cy);
    await page.waitForTimeout(120);
    await page.keyboard.press("Backspace");
    await page.waitForTimeout(150);
    const remaining = await page.evaluate((orig) => {
      for (const el of document.querySelectorAll(".dxw-page span")) {
        if (el.children.length) continue;
        const s = el.textContent || "";
        if (s && /[ऀ-ॿ]/.test(s) && orig.startsWith(s)) return s;
      }
      return null;
    }, info!.text);
    // Exactly one whole cluster removed (not a single dangling combining mark).
    expect(remaining).toBe(info!.text.slice(0, info!.len - lastLen));
  });

  test("CJK: IME composition previews then commits at the caret", async ({ page }) => {
    await open(page, "staging-eastasian");
    const target = await page.evaluate(() => {
      for (const el of document.querySelectorAll(".dxw-page span")) {
        if (el.children.length) continue;
        const t = el.textContent || "";
        if (/[぀-ヿ一-鿿]/.test(t)) {
          const r = el.getBoundingClientRect();
          return { cx: r.x + 4, cy: r.y + r.height / 2 };
        }
      }
      return null;
    });
    expect(target).not.toBeNull();
    await page.mouse.click(target!.cx, target!.cy);
    await page.waitForTimeout(120);
    const result = await page.evaluate(async () => {
      const ta = document.activeElement as HTMLElement | null;
      if (!ta || ta.tagName !== "TEXTAREA") return { ok: false, why: "no IME textarea focused" };
      const fire = (type: string, data: string) =>
        ta.dispatchEvent(new CompositionEvent(type, { data, bubbles: true }));
      fire("compositionstart", "");
      (ta as HTMLTextAreaElement).value = "にほん";
      fire("compositionupdate", "にほん");
      await new Promise((r) => setTimeout(r, 40));
      const preview = [...document.querySelectorAll(".dxw-page span")].some(
        (s) => (s as HTMLElement).style.textDecoration === "underline" && s.textContent === "にほん",
      );
      fire("compositionend", "日本語");
      return { ok: true, preview };
    });
    expect(result.ok, result.why).toBe(true);
    expect(result.preview, "composition preview shown").toBe(true);
    await page.waitForTimeout(150);
    const committed = await page.evaluate(() => document.body.textContent!.includes("日本語"));
    expect(committed, "composed text committed").toBe(true);
  });
});
