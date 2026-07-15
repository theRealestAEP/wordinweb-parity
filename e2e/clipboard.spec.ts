import { expect, test, Page } from "@playwright/test";
import { readFileSync } from "node:fs";
import { strFromU8, unzipSync } from "fflate";
import { Editor } from "./editing.js";

/**
 * Clipboard end-to-end: copy/cut/paste through the real browser clipboard.
 * The load-bearing case is RTL — selection segments arrive in visual (bidi-
 * reordered) paint order, so a naive copy mirror-reverses Hebrew/Arabic. Copy
 * must emit LOGICAL (source/reading) order. A Latin control proves the path is
 * not broken in general, and a suggesting-mode paste must record w:ins.
 */

const MOD = process.platform === "darwin" ? "Meta" : "Control";

async function open(page: Page, fixture: string, query = ""): Promise<void> {
  await page.goto(`/?doc=/fixtures/${fixture}.docx${query}`);
  await page.waitForSelector(".dxw-page span");
  await page.waitForTimeout(400);
}

async function readClip(page: Page): Promise<string> {
  return page.evaluate(() => navigator.clipboard.readText());
}

async function pasteData(page: Page, plain: string, html = ""): Promise<void> {
  await page.evaluate(({ plainText, htmlText }) => {
    const data = new DataTransfer();
    data.setData("text/plain", plainText);
    if (htmlText) data.setData("text/html", htmlText);
    document.activeElement?.dispatchEvent(new ClipboardEvent("paste", { clipboardData: data, bubbles: true }));
  }, { plainText: plain, htmlText: html });
}

async function downloadedDocumentXml(page: Page): Promise<string> {
  const downloadPromise = page.waitForEvent("download");
  await page.getByText("Download", { exact: true }).click();
  const download = await downloadPromise;
  const path = await download.path();
  expect(path).not.toBeNull();
  return strFromU8(unzipSync(new Uint8Array(readFileSync(path!)))["word/document.xml"]);
}

/** Click the leaf span whose exact text equals `text` (its geometric center). */
async function clickExact(page: Page, text: string): Promise<void> {
  const box = await page.evaluate((t) => {
    for (const el of document.querySelectorAll(".dxw-page span")) {
      if (el.children.length) continue;
      if ((el.textContent || "") === t) {
        const r = el.getBoundingClientRect();
        return { x: r.x + r.width / 2, y: r.y + r.height / 2 };
      }
    }
    return null;
  }, text);
  expect(box, `span "${text}" present`).not.toBeNull();
  await page.mouse.click(box!.x, box!.y);
  await page.waitForTimeout(120);
}

/** Select the whole visual line the caret sits on (logical Home..End). */
async function selectLine(page: Page): Promise<void> {
  await page.keyboard.press("Home");
  await page.waitForTimeout(60);
  await page.keyboard.press("Shift+End");
  await page.waitForTimeout(100);
}

test.describe("clipboard", () => {
  test("RTL copy emits logical (source) order, not visual-reversed", async ({ page }) => {
    await open(page, "wild2-lit-yiddish-rtl");
    // Paragraph "שלום עליכם" is authored (logical) as שלום then עליכם, but paints
    // RTL so עליכם sits visually left of שלום across three runs.
    await clickExact(page, "שלום");
    await selectLine(page);
    await page.keyboard.press(`${MOD}+c`);
    await page.waitForTimeout(200);
    const clip = (await readClip(page)).trim();
    expect(clip).toBe("שלום עליכם");
    // Guard against a regression that dumps visual order.
    expect(clip).not.toBe("עליכם שלום");
    expect(clip.indexOf("שלום")).toBeLessThan(clip.indexOf("עליכם"));
  });

  test("Latin control: word copy and line copy round-trip through the clipboard", async ({ page }) => {
    await open(page, "benchmark");
    await page.locator('.dxw-page span:text-is("Kitchen")').first().dblclick();
    await page.waitForTimeout(120);
    await page.keyboard.press(`${MOD}+c`);
    await page.waitForTimeout(200);
    expect((await readClip(page)).trim()).toBe("Kitchen");

    // Whole-line copy: LTR logical order equals visual, so the title comes back intact.
    await clickExact(page, "Kitchen");
    await selectLine(page);
    await page.keyboard.press(`${MOD}+c`);
    await page.waitForTimeout(200);
    expect((await readClip(page)).trim()).toBe("The Kitchen Sink Benchmark");
  });

  test("cut removes the selected text and preserves it on the clipboard", async ({ page }) => {
    await open(page, "benchmark");
    await page.locator('.dxw-page span:text-is("Kitchen")').first().dblclick();
    await page.waitForTimeout(120);
    await page.keyboard.press(`${MOD}+x`);
    await page.waitForTimeout(250);
    expect((await readClip(page)).trim()).toBe("Kitchen");
    const stillThere = await page.evaluate(() =>
      [...document.querySelectorAll(".dxw-page span")].some((s) => (s.textContent || "") === "Kitchen"),
    );
    expect(stillThere, "cut word no longer rendered").toBe(false);
  });

  test("paste inserts clipboard text at the caret", async ({ page }) => {
    await open(page, "benchmark");
    await page.evaluate(() => navigator.clipboard.writeText("ZQXPASTE"));
    await clickExact(page, "Typography");
    await page.keyboard.press(`${MOD}+v`);
    await page.waitForTimeout(300);
    const ed = new Editor(page);
    await ed.expectHasText("ZQXPASTE");
  });

  test("multi-line paste splits into paragraphs", async ({ page }) => {
    await open(page, "benchmark");
    await page.evaluate(() => navigator.clipboard.writeText("ZQXONE\nZQXTWO"));
    await clickExact(page, "Typography");
    await page.keyboard.press(`${MOD}+v`);
    await page.waitForTimeout(300);
    const ed = new Editor(page);
    await ed.expectHasText("ZQXONE");
    await ed.expectHasText("ZQXTWO");
  });

  test("RTL paste keeps logical order after copy", async ({ page }) => {
    await open(page, "wild2-lit-yiddish-rtl");
    await clickExact(page, "שלום");
    await selectLine(page);
    await page.keyboard.press(`${MOD}+c`);
    await page.waitForTimeout(200);
    const copied = (await readClip(page)).trim();
    expect(copied).toBe("שלום עליכם");
    // Paste it back at the end of the same line; the inserted text stays logical.
    await page.keyboard.press("End");
    await page.keyboard.press(`${MOD}+v`);
    await page.waitForTimeout(300);
    // Re-select the (now longer) line and re-copy: it must still read logically.
    await clickExact(page, "שלום");
    await selectLine(page);
    await page.keyboard.press(`${MOD}+c`);
    await page.waitForTimeout(200);
    const again = (await readClip(page)).trim();
    expect(again.startsWith("שלום עליכם")).toBe(true);
    expect(again).not.toContain("עליכם שלום");
  });

  test("paste in suggesting mode records an insertion (w:ins, underlined + colored)", async ({ page }) => {
    const ed = await Editor.open(page, "sample", "&suggest=1");
    await page.evaluate(() => navigator.clipboard.writeText("ZQXSUG"));
    await ed.clickText("time");
    await page.keyboard.press(`${MOD}+v`);
    await page.waitForTimeout(300);
    await ed.expectHasText("ZQXSUG");
    const style = await ed.spanStyle("ZQXSUG");
    expect(style, "pasted span present").not.toBeNull();
    expect(style!.color).toBe("rgb(192, 0, 0)"); // insertion color #C00000
    expect(style!.decoration).toContain("underline");
  });

  test("copy and paste retains character formatting", async ({ page }) => {
    await open(page, "benchmark");
    await page.locator('.dxw-page span:text-is("bold,")').first().dblclick();
    await expect(page.locator(".dxw-sel")).not.toHaveCount(0);
    await page.keyboard.press(`${MOD}+c`);
    await page.waitForTimeout(100);
    const types = await page.evaluate(async () => (await navigator.clipboard.read()).flatMap((item) => item.types));
    expect(types).toContain("text/html");
    await clickExact(page, "classic.");
    await page.keyboard.press("End");
    await page.keyboard.press(`${MOD}+v`);
    await page.waitForTimeout(300);

    const boldCopies = await page.evaluate(() =>
      [...document.querySelectorAll(".dxw-page span")].filter((span) =>
        span.textContent === "bold," && Number.parseInt(getComputedStyle(span).fontWeight, 10) >= 600,
      ).length,
    );
    expect(boldCopies).toBe(2);
  });

  test("HTML paste retains table structure and inline formatting", async ({ page }) => {
    await open(page, "benchmark");
    await clickExact(page, "Typography");
    await pasteData(
      page,
      "RICHLEFT\tRICHBOLD\nRICHBOTTOM\tRICHRIGHT",
      `<table><tr><td>RICHLEFT</td><td><strong style="color:#c00000">RICHBOLD</strong></td></tr>` +
        `<tr><td>RICHBOTTOM</td><td>RICHRIGHT</td></tr></table>`,
    );
    await page.waitForTimeout(400);
    const ed = new Editor(page);
    await ed.expectHasText("RICHLEFT");
    await ed.expectHasText("RICHBOLD");

    const xml = await downloadedDocumentXml(page);
    expect(xml).toMatch(/<w:tbl>[\s\S]*RICHLEFT[\s\S]*RICHBOLD[\s\S]*RICHBOTTOM[\s\S]*RICHRIGHT[\s\S]*<\/w:tbl>/);
    expect(xml).toMatch(/<w:b\/>[\s\S]*<w:color w:val="C00000"\/>[\s\S]*RICHBOLD/);
  });

  test("large multi-paragraph paste stays responsive and saves every paragraph", async ({ page }) => {
    test.setTimeout(60_000);
    await open(page, "sample");
    await page.locator(".dxw-page").nth(1).scrollIntoViewIfNeeded();
    await page.waitForTimeout(100);
    await clickExact(page, "starts");
    const text = Array.from({ length: 3_000 }, (_, i) => `LARGECOPY-${i.toString().padStart(4, "0")}-${"x".repeat(70)}`).join("\n");
    await page.evaluate(() => {
      const state = globalThis as typeof globalThis & { ticks?: number; maxGap?: number; lastTick?: number; timer?: number };
      state.ticks = 0;
      state.maxGap = 0;
      state.lastTick = performance.now();
      state.timer = window.setInterval(() => {
        const now = performance.now();
        state.ticks!++;
        state.maxGap = Math.max(state.maxGap!, now - state.lastTick!);
        state.lastTick = now;
      }, 10);
    });
    await pasteData(page, text);
    await expect(page.locator("[data-dxw-layout-status]")).toBeHidden({ timeout: 45_000 });
    const responsiveness = await page.evaluate(() => {
      const state = globalThis as typeof globalThis & { ticks?: number; maxGap?: number; timer?: number };
      clearInterval(state.timer);
      return { ticks: state.ticks ?? 0, maxGap: state.maxGap ?? Infinity };
    });
    expect(responsiveness.ticks).toBeGreaterThan(0);
    expect(responsiveness.maxGap).toBeLessThan(500);

    const xml = await downloadedDocumentXml(page);
    expect(xml).toContain("LARGECOPY-0000");
    expect(xml).toContain("LARGECOPY-2999");
  });
});
