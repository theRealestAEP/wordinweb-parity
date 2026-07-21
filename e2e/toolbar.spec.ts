import { test, expect, Page, Locator } from "@playwright/test";
import { readFileSync } from "node:fs";
import { strFromU8, unzipSync } from "fflate";

/**
 * Every toolbar control, driven end-to-end with real input. Exists because
 * a command can pass its unit tests while the button wiring above it is
 * broken (footnote insert regressed exactly this way).
 */

const MOD = process.platform === "darwin" ? "Meta" : "Control";

async function load(page: Page, doc = "sample"): Promise<void> {
  await page.goto(`/?doc=/fixtures/${doc}.docx`);
  await page.waitForSelector(".dxw-page span");
  await page.waitForTimeout(300);
}

async function openTab(page: Page, tab: "home" | "insert" | "layout"): Promise<void> {
  await page.locator(`button[data-tab="${tab}"]`).click();
  await page.waitForTimeout(100);
}

function exact(page: Page, text: string): Locator {
  return page.locator(`.dxw-page span:text-is("${text}")`).first();
}

async function selectWord(page: Page, text: string): Promise<void> {
  await exact(page, text).dblclick();
  await page.waitForTimeout(150);
}

async function styleOf(page: Page, text: string): Promise<{ font: string; textDecoration: string; color: string }> {
  return exact(page, text).evaluate((el) => ({
    font: el.style.font,
    textDecoration: el.style.textDecoration,
    color: el.style.color,
  }));
}

function btn(page: Page, title: string): Locator {
  return page.locator(`button[title="${title}"], button[data-tip="${title}"]`).first();
}

test.describe("formatting controls", () => {
  test("bold, italic, underline, strike, super/subscript, clear", async ({ page }) => {
    await load(page);
    await selectWord(page, "Lorem");
    await btn(page, "Bold (⌘B)").click();
    await page.waitForTimeout(200);
    expect((await styleOf(page, "Lorem")).font).toContain("700");

    await selectWord(page, "Lorem");
    await btn(page, "Italic").click();
    await page.waitForTimeout(200);
    expect((await styleOf(page, "Lorem")).font).toContain("italic");

    await selectWord(page, "ipsum");
    await btn(page, "Underline").click();
    await page.waitForTimeout(200);
    expect((await styleOf(page, "ipsum")).textDecoration).toContain("underline");

    // strike renders as an engine-painted rule at Word's height, not CSS
    await selectWord(page, "dolor");
    await btn(page, "Strikethrough").click();
    await page.waitForTimeout(250);
    const strikeRule = await page.evaluate(() => {
      const s = [...document.querySelectorAll(".dxw-page span")].find((x) => x.textContent === "dolor");
      const r = s!.getBoundingClientRect();
      return [...document.querySelectorAll(".dxw-page div")].some((d) => {
        const b = d.getBoundingClientRect();
        return b.height <= 2 && Math.abs(b.left - r.left) < 3 && b.top > r.top && b.bottom < r.bottom;
      });
    });
    expect(strikeRule).toBe(true);

    await selectWord(page, "amet,");
    await btn(page, "Superscript").click();
    await page.waitForTimeout(200);
    const supH = await exact(page, "amet,").evaluate((el) => el.getBoundingClientRect().height);
    expect(supH).toBeLessThan(14); // scaled glyph box

    await selectWord(page, "amet,");
    await btn(page, "Clear formatting").click();
    await page.waitForTimeout(200);
    const clearedH = await exact(page, "amet,").evaluate((el) => el.getBoundingClientRect().height);
    expect(clearedH).toBeGreaterThan(14);
  });

  test("change case menu", async ({ page }) => {
    await load(page);
    await selectWord(page, "consectetur");
    await page.locator('select[title="Change case"]').selectOption("upper");
    await page.waitForTimeout(300);
    await expect(exact(page, "CONSECTETUR")).toHaveCount(1);
  });

  test("font family and size selects", async ({ page }) => {
    await load(page);
    await selectWord(page, "Lorem");
    await page.locator('select[title="Font size"]').selectOption("18");
    await page.waitForTimeout(300);
    const f = await styleOf(page, "Lorem");
    expect(f.font).toContain("24px"); // 18pt = 24px
  });
});

test.describe("paragraph controls", () => {
  test("indent and outdent shift the paragraph by half an inch", async ({ page }) => {
    await load(page, "parity-text");
    const before = (await exact(page, "Plain").boundingBox())!.x;
    await exact(page, "Plain").click();
    await btn(page, "Increase indent").click();
    await page.waitForTimeout(300);
    expect((await exact(page, "Plain").boundingBox())!.x).toBeCloseTo(before + 48, 0);
    await btn(page, "Decrease indent").click();
    await page.waitForTimeout(300);
    expect((await exact(page, "Plain").boundingBox())!.x).toBeCloseTo(before, 0);
  });

  test("line spacing menu grows the paragraph", async ({ page }) => {
    await load(page, "parity-text");
    const gap = async () => {
      const a = (await exact(page, "Plain").boundingBox())!.y;
      const rows = await page.evaluate(() => {
        const plain = [...document.querySelectorAll(".dxw-page span")].find((s) => s.textContent === "Plain");
        const py = plain!.getBoundingClientRect().y;
        const next = [...document.querySelectorAll(".dxw-page span")]
          .map((s) => s.getBoundingClientRect().y)
          .filter((y) => y > py + 4)
          .sort((x, y) => x - y);
        return next[0] - py;
      });
      void a;
      return rows;
    };
    const single = await gap();
    await exact(page, "Plain").click();
    await page.locator('select[title="Line & paragraph spacing"]').selectOption("l:2");
    await page.waitForTimeout(300);
    expect(await gap()).toBeGreaterThan(single + 8);
  });

  test("exact line height presets and custom points update the paragraph", async ({ page }) => {
    await load(page, "parity-text");
    const paragraph = exact(page, "Plain");
    const downloadParagraphXml = async () => {
      const pending = page.waitForEvent("download");
      await btn(page, "Save edited .docx").click();
      const path = await (await pending).path();
      if (!path) throw new Error("download path unavailable");
      const xml = strFromU8(unzipSync(new Uint8Array(readFileSync(path)))["word/document.xml"]);
      const textAt = xml.indexOf("Plain text parity");
      return xml.slice(xml.lastIndexOf("<w:p", textAt), xml.indexOf("</w:p>", textAt) + 6);
    };
    await paragraph.click();
    const spacing = page.locator('select[title="Line & paragraph spacing"]');

    await spacing.selectOption("e:24");
    expect(await downloadParagraphXml()).toMatch(/<w:spacing\b[^>]*w:line="480"[^>]*w:lineRule="exact"/);

    await spacing.selectOption("e:custom");
    const dialog = page.getByRole("dialog", { name: "Exact line height" });
    await expect(dialog).toBeVisible();
    await dialog.getByRole("textbox", { name: "Line height (points)" }).fill("30");
    await dialog.getByRole("button", { name: "Apply" }).click();
    expect(await downloadParagraphXml()).toMatch(/<w:spacing\b[^>]*w:line="600"[^>]*w:lineRule="exact"/);
  });

  test("lists toggle and Tab changes level", async ({ page }) => {
    await load(page, "parity-text");
    await exact(page, "Plain").click();
    // The bullet glyph depends on which numbering definition gets reused
    // (docx-lib fixtures carry ●/○/■; our on-demand defs use •/○/■).
    const lvl0 = page.locator(".dxw-page span").filter({ hasText: /^[•●]$/ });
    const lvl1 = page.locator(".dxw-page span").filter({ hasText: /^[○o]$/ });
    await btn(page, "Bulleted list").click();
    await page.waitForTimeout(300);
    await expect(lvl0).toHaveCount(1);
    await page.keyboard.press("Tab");
    await page.waitForTimeout(300);
    await expect(lvl1).toHaveCount(1);
    await page.keyboard.press("Shift+Tab");
    await page.waitForTimeout(300);
    await expect(lvl0).toHaveCount(1);
    await btn(page, "Bulleted list").click();
    await page.waitForTimeout(300);
    await expect(lvl0).toHaveCount(0);
  });

  test("alignment buttons move the line", async ({ page }) => {
    await load(page, "parity-text");
    await exact(page, "Plain").click();
    const before = (await exact(page, "Plain").boundingBox())!.x;
    await btn(page, "Center").click();
    await page.waitForTimeout(300);
    expect((await exact(page, "Plain").boundingBox())!.x).toBeGreaterThan(before + 40);
  });
});

test.describe("insert controls", () => {
  test("hyperlink: apply and remove", async ({ page }) => {
    await load(page);
    await selectWord(page, "veniam,");
    await openTab(page, "insert");
    await btn(page, "Insert link").click();
    await page.fill('input[placeholder="Paste or type a link"]', "example.org/x");
    await page.keyboard.press("Enter");
    await page.waitForTimeout(400);
    const a = page.locator('.dxw-page a:text-is("veniam,")');
    await expect(a).toHaveCount(1);
    await expect(a).toHaveAttribute("href", "https://example.org/x");
  });

  test("hyperlinks require a modifier while editing and show their full URL", async ({ page, context }) => {
    await load(page);
    await selectWord(page, "veniam,");
    await openTab(page, "insert");
    await btn(page, "Insert link").click();
    const target = new URL("/?opened=hyperlink", page.url()).href;
    await page.fill('input[placeholder="Paste or type a link"]', target);
    await page.keyboard.press("Enter");
    await page.waitForTimeout(400);

    const link = page.locator('.dxw-page a:text-is("veniam,")');
    await expect(link).toHaveAttribute("title", target);

    const pageCount = context.pages().length;
    await link.click();
    await page.waitForTimeout(200);
    expect(context.pages()).toHaveLength(pageCount);
    await expect(page.locator("[data-dxw-caret]")).toBeVisible();

    const popupPromise = page.waitForEvent("popup");
    await link.click({ modifiers: [MOD] });
    const popup = await popupPromise;
    await popup.waitForLoadState("domcontentloaded");
    expect(new URL(popup.url()).searchParams.get("opened")).toBe("hyperlink");
    await popup.close();
  });

  test("comment on a selection shows a balloon", async ({ page }) => {
    await load(page);
    await selectWord(page, "nostrud");
    await openTab(page, "insert");
    await btn(page, "Add comment (select text first)").click();
    await page.fill('textarea[placeholder="Comment on the selection…"]', "toolbar spec comment");
    await page.locator('button:text-is("Comment")').click();
    await page.waitForTimeout(500);
    await expect(page.locator(".dxw-comment-card", { hasText: "toolbar spec comment" })).toHaveCount(1);
  });

  test("footnote inserts at the caret and hints without one", async ({ page }) => {
    await load(page);
    await openTab(page, "insert");
    // no caret: popover explains instead of silently failing
    await btn(page, "Insert footnote (at the caret)").click();
    await page.fill('textarea[placeholder="Footnote text…"]', "note body text");
    await page.locator('button:text-is("Insert")').click();
    await page.waitForTimeout(300);
    await expect(page.locator("text=Click into the text first")).toHaveCount(1);
    // with caret: mark + note area appear
    await exact(page, "laborum.").click();
    await btn(page, "Insert footnote (at the caret)").click();
    await page.fill('textarea[placeholder="Footnote text…"]', "note body text");
    await page.locator('button:text-is("Insert")').click();
    await page.waitForTimeout(600);
    await expect(page.locator('.dxw-page span:text-is("body")')).toHaveCount(1);
  });

  test("table grid picker inserts a table; cell fill applies", async ({ page }) => {
    await load(page, "parity-text");
    await exact(page, "Plain").click();
    await openTab(page, "insert");
    await btn(page, "Table").click();
    // pick 2x2 in the grid (grid cells are 16px squares in a 6-wide flex grid)
    const grid = page.locator("div").filter({ has: page.locator('div[style*="width: 16px"]') }).first();
    void grid;
    const cells = page.locator('div[style*="width: 16px"][style*="height: 16px"]');
    await cells.nth(7).click(); // row 2, col 2
    await page.waitForTimeout(500);
    const edges = await page.evaluate(
      () => document.querySelectorAll(".dxw-page [data-dxw-edge]").length,
    );
    expect(edges).toBeGreaterThan(5);
  });

  test("image inserts from a file", async ({ page }) => {
    await load(page, "parity-text");
    await exact(page, "Plain").click();
    await openTab(page, "insert");
    const png = Buffer.from(
      "iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAIAAAACUFjqAAAAFUlEQVR42mNkaPjPgBswMeAFI1UaANkuAZOtAgEcAAAAAElFTkSuQmCC",
      "base64",
    );
    await page.locator('input[type="file"][accept*="image/png"]').setInputFiles({ name: "dot.png", mimeType: "image/png", buffer: png });
    await page.waitForTimeout(700);
    await expect(page.locator(".dxw-page img")).toHaveCount(1);
  });
});

test.describe("document controls", () => {
  test("find and replace all", async ({ page }) => {
    await load(page, "parity-text");
    await page.keyboard.press(`${MOD}+f`);
    await page.fill('input[placeholder="Find in document"]', "parity");
    await page.waitForTimeout(400);
    await expect(page.locator("text=/1 of \\d/")).toHaveCount(1);
    await page.fill('input[placeholder="Replace with"]', "PARITY");
    await page.locator('button:text-is("Replace all")').click();
    await page.waitForTimeout(500);
    await expect(page.locator('.dxw-page span:text-is("PARITY")').first()).toBeVisible();
  });

  test("layout tab: margins narrow moves text left", async ({ page }) => {
    await load(page, "parity-text");
    const before = (await exact(page, "Plain").boundingBox())!.x;
    await openTab(page, "layout");
    await page.locator('[data-dxw-layout-menu-trigger="margins"]').click();
    await page.locator('[data-dxw-layout-option="m:narrow"]').click();
    await page.waitForTimeout(400);
    expect((await exact(page, "Plain").boundingBox())!.x).toBeLessThan(before - 20);
  });

  test("layout tab: page border option draws the box", async ({ page }) => {
    await load(page, "parity-text");
    await openTab(page, "layout");
    await page.locator('[data-dxw-layout-menu-trigger="page-border"]').click();
    await page.locator('[data-dxw-layout-option="thin"]').click();
    await page.waitForTimeout(400);
    const edges = await page.evaluate(
      () => document.querySelectorAll(".dxw-page [data-dxw-edge]").length,
    );
    expect(edges).toBeGreaterThanOrEqual(4);
    await page.locator('[data-dxw-layout-menu-trigger="page-border"]').click();
    await page.locator('[data-dxw-layout-option="none"]').click();
    await page.waitForTimeout(400);
    const edges2 = await page.evaluate(
      () => document.querySelectorAll(".dxw-page [data-dxw-edge]").length,
    );
    expect(edges2).toBeLessThan(edges);
  });

  test("insert tab: dynamic page number at the caret", async ({ page }) => {
    await load(page, "parity-text");
    const target = exact(page, "Plain");
    await target.click();
    await page.waitForTimeout(200);
    await openTab(page, "insert");
    await page.locator('select[title="Insert a dynamic page number at the caret"]').selectOption("pn:pageof");
    await page.waitForTimeout(400);
    await expect(page.locator('.dxw-page span:text-is("1")').first()).toBeVisible();
    const text = await page.evaluate(() =>
      [...document.querySelectorAll(".dxw-page span")].map((s) => s.textContent).join(" "),
    );
    expect(text).toContain("Page ");
    expect(text).toContain(" of ");
  });

  test("insert tab: section break + per-section landscape", async ({ page }) => {
    await load(page, "parity-text");
    await exact(page, "Plain").click();
    await page.waitForTimeout(200);
    await openTab(page, "insert");
    await page.locator('select[title="Insert a page, column or section break at the caret"]').selectOption("br:next");
    await page.waitForTimeout(500);
    const pages1 = await page.locator(".dxw-page").count();
    expect(pages1).toBeGreaterThanOrEqual(2);
    // caret is still in section 1; landscape only that section
    await exact(page, "Plain").click();
    await page.waitForTimeout(200);
    await openTab(page, "layout");
    await page.locator('select[title="Apply layout changes to"]').selectOption("section");
    await page.locator('[data-dxw-layout-menu-trigger="orientation"]').click();
    await page.locator('[data-dxw-layout-option="landscape"]').click();
    await page.waitForTimeout(500);
    const sizes = await page.evaluate(() =>
      [...document.querySelectorAll(".dxw-page")].map((p) => {
        const r = p.getBoundingClientRect();
        return r.width > r.height ? "landscape" : "portrait";
      }),
    );
    expect(sizes[0]).toBe("landscape");
    expect(sizes[sizes.length - 1]).toBe("portrait");
  });

  test("header/footer hotbar appears in hf mode and inserts a page number", async ({ page }) => {
    await load(page, "parity-text");
    // dblclick in the top margin band enters header editing
    const pageBox = (await page.locator(".dxw-page").first().boundingBox())!;
    await page.mouse.dblclick(pageBox.x + pageBox.width / 2, pageBox.y + 30);
    await page.waitForTimeout(400);
    const hotbar = page.locator("[data-dxw-hf-hotbar]");
    await expect(hotbar).toBeVisible();
    await hotbar.locator('button:text-is("Page number")').click();
    await page.waitForTimeout(400);
    await expect(page.locator('.dxw-page span:text-is("1")').first()).toBeVisible();
    await hotbar.locator('button:text-is("Close")').click();
    await page.waitForTimeout(300);
    await expect(hotbar).toHaveCount(0);
  });

  test("styles dropdown applies a heading", async ({ page }) => {
    await load(page, "parity-text");
    await exact(page, "Plain").click();
    await page.locator('select[title="Paragraph style"]').selectOption("Heading1");
    await page.waitForTimeout(400);
    const f = await styleOf(page, "Plain");
    expect(f.font).not.toContain("14.6"); // heading size differs from body
  });
});

test.describe("math controls", () => {
  test("click opens the linear editor; edit round-trips; drag moves the equation", async ({ page }) => {
    await load(page, "parity-math");
    await page.click("[data-dxw-math]");
    const input = page.locator("input").last();
    await expect(input).toHaveValue(/\^/);
    await input.fill("y^3");
    await input.press("Enter");
    await page.waitForTimeout(500);
    await expect(page.locator('.dxw-page span:text-is("𝑦")')).toHaveCount(1);
    // drag the equation to another word
    const from = (await page.locator("[data-dxw-math]").first().boundingBox())!;
    const to = (await exact(page, "resumes.").boundingBox())!;
    await page.mouse.move(from.x + from.width / 2, from.y + from.height / 2);
    await page.mouse.down();
    await page.mouse.move(to.x + 2, to.y + to.height / 2, { steps: 8 });
    await page.mouse.up();
    await page.waitForTimeout(500);
    const mathX = (await page.locator("[data-dxw-math]").first().boundingBox())!.x;
    const andX = (await exact(page, "and").boundingBox())!.x;
    expect(mathX).toBeGreaterThan(andX);
  });
});

test.describe("header/footer entry", () => {
  test("margin dblclick creates and edits a header even with page borders", async ({ page }) => {
    await load(page, "parity-pageborders");
    const pageBox = (await page.locator(".dxw-page").first().boundingBox())!;
    await page.mouse.dblclick(pageBox.x + 300, pageBox.y + 30);
    await page.waitForTimeout(500);
    await page.keyboard.type("Draft Header");
    await page.waitForTimeout(500);
    await expect(page.locator('.dxw-page span:text-is("Draft")')).toHaveCount(1);
  });
});
