import { Page, Locator, expect } from "@playwright/test";

/**
 * High-level editing-behavior DSL for expectation tests: open a fixture,
 * drive the caret/keyboard the way a user would, and assert on the resulting
 * layout (page counts, where text landed, caret position). Keeps behavior
 * specs declarative — see editing-behaviors.spec.ts.
 */
export class Editor {
  constructor(private page: Page) {}

  static async open(page: Page, fixture: string, query = ""): Promise<Editor> {
    await page.goto(`/?doc=/fixtures/${fixture}.docx${query}`);
    await page.waitForSelector(".dxw-page span");
    await page.waitForTimeout(300);
    return new Editor(page);
  }

  /** Computed style of the first span whose text contains `text`. */
  async spanStyle(text: string): Promise<{ color: string; decoration: string } | null> {
    return this.page.evaluate((t) => {
      const s = [...document.querySelectorAll(".dxw-page span")].find((x) =>
        (x.textContent ?? "").includes(t),
      ) as HTMLElement | undefined;
      if (!s) return null;
      const cs = getComputedStyle(s);
      return { color: cs.color, decoration: cs.textDecorationLine || cs.textDecoration };
    }, text);
  }

  /** Click the Accept/Reject button in the suggestion popover. */
  async reviewClick(which: "Accept" | "Reject"): Promise<void> {
    await this.page.locator(`button[title$="suggestion"]:has-text("${which}")`).click();
    await this.page.waitForTimeout(150);
  }

  /** The exact-text span (first match). */
  span(text: string): Locator {
    return this.page.locator(`.dxw-page span:text-is(${JSON.stringify(text)})`).first();
  }

  /** Click a span by its exact text. */
  async clickText(text: string): Promise<void> {
    await this.span(text).click();
    await this.page.waitForTimeout(120);
  }

  /** Click `dy` px below (or above, if negative) a span — for whitespace /
   * below-table / above-text clicks. */
  async clickNear(text: string, dx: number, dy: number): Promise<void> {
    const b = (await this.span(text).boundingBox())!;
    await this.page.mouse.click(b.x + dx, b.y + b.height / 2 + dy);
    await this.page.waitForTimeout(120);
  }

  /** Click at an absolute fraction of the first page (0..1). */
  async clickPageFraction(fx: number, fy: number): Promise<void> {
    const page = this.page.locator(".dxw-page").first();
    const b = (await page.boundingBox())!;
    await page.click({ position: { x: b.width * fx, y: b.height * fy } });
    await this.page.waitForTimeout(120);
  }

  async type(text: string): Promise<void> {
    await this.page.keyboard.type(text);
    await this.page.waitForTimeout(200);
  }

  async press(key: string): Promise<void> {
    await this.page.keyboard.press(key);
    await this.page.waitForTimeout(150);
  }

  async pressMany(key: string, n: number): Promise<void> {
    for (let i = 0; i < n; i++) await this.page.keyboard.press(key);
    await this.page.waitForTimeout(250);
  }

  // ---- assertions ----

  async pageCount(): Promise<number> {
    return this.page.locator(".dxw-page").count();
  }

  async expectPageCount(n: number): Promise<void> {
    expect(await this.pageCount()).toBe(n);
  }

  async expectPageCountAtLeast(n: number): Promise<void> {
    expect(await this.pageCount()).toBeGreaterThanOrEqual(n);
  }

  /** Index of the page containing a span with `text` (substring), or -1. */
  async pageOf(text: string): Promise<number> {
    return this.page.evaluate((t) => {
      const s = [...document.querySelectorAll(".dxw-page span")].find((x) => x.textContent!.includes(t));
      return s ? [...document.querySelectorAll(".dxw-page")].indexOf(s.closest(".dxw-page")!) : -1;
    }, text);
  }

  async expectTextOnPage(text: string, pageIndex: number): Promise<void> {
    expect(await this.pageOf(text)).toBe(pageIndex);
  }

  async expectHasText(text: string): Promise<void> {
    const found = await this.page.evaluate(
      (t) => [...document.querySelectorAll(".dxw-page span")].some((s) => s.textContent!.includes(t)),
      text,
    );
    expect(found, `expected rendered text to contain "${text}"`).toBe(true);
  }

  /** Caret box (pt, page-space) or null when hidden. */
  async caret(): Promise<{ page: number; topPt: number; leftPt: number } | null> {
    return this.page.evaluate(() => {
      const c = [...document.querySelectorAll("div")].find(
        (d) => {
          const style = (d as HTMLElement).style;
          return (style.width === "1.5px" || style.height === "1.5px") && style.display === "block";
        },
      ) as HTMLElement | undefined;
      if (!c) return null;
      const pg = c.closest(".dxw-page");
      return {
        page: pg ? [...document.querySelectorAll(".dxw-page")].indexOf(pg) : -1,
        topPt: Math.round(parseFloat(c.style.top) * 0.75),
        leftPt: Math.round(parseFloat(c.style.left) * 0.75),
      };
    });
  }

  async expectCaretVisible(): Promise<void> {
    expect(await this.caret(), "expected a visible caret").not.toBeNull();
  }

  async expectCaretOnPage(pageIndex: number): Promise<void> {
    const c = await this.caret();
    expect(c, "expected a visible caret").not.toBeNull();
    expect(c!.page).toBe(pageIndex);
  }
}
