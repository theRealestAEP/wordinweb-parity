import { expect, Page, test } from "@playwright/test";
import { readFileSync } from "node:fs";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

const textItem = (page: Page, text: string) => page.getByText(text, { exact: true });

const SECTION =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';

function fixture(): Buffer {
  const body =
    '<w:p><w:r><w:t>Odd page</w:t></w:r></w:p>' +
    '<w:p><w:r><w:br w:type="page"/></w:r></w:p>' +
    '<w:p><w:r><w:t>Even page</w:t></w:r></w:p>' +
    `<w:p><w:pPr>${SECTION}</w:pPr><w:r><w:t>First section end</w:t></w:r></w:p>` +
    '<w:p><w:r><w:t>Second section</w:t></w:r></w:p>' +
    SECTION;
  const files = {
    "[Content_Types].xml": `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>`,
    "_rels/.rels": `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
    "word/_rels/document.xml.rels": `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>`,
    "word/settings.xml": `<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:zoom w:percent="100"/><w:proofState w:spelling="clean"/><w:defaultTabStop w:val="720"/><w:compat/></w:settings>`,
    "word/document.xml": `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`,
  };
  return Buffer.from(zipSync(Object.fromEntries(Object.entries(files).map(([name, xml]) => [name, strToU8(xml)]))));
}

async function load(page: Page): Promise<void> {
  await page.goto("/");
  await page.locator("#docx-upload").setInputFiles({
    name: "layout-page-parity.docx",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    buffer: fixture(),
  });
  await expect(textItem(page, "Odd page")).toBeVisible();
}

async function openLayout(page: Page): Promise<void> {
  await page.locator('button[data-tab="layout"]').click();
  await expect(page.locator('[data-dxw-layout-menu-trigger="margins"]')).toBeVisible();
}

async function pickLayout(page: Page, menu: string, value: string): Promise<void> {
  await page.locator(`[data-dxw-layout-menu-trigger="${menu}"]`).click();
  await page.locator(`[data-dxw-layout-option="${value}"]`).click();
}

async function downloadParts(page: Page): Promise<Record<string, Uint8Array>> {
  const pending = page.waitForEvent("download");
  await page.getByText("Download", { exact: true }).click();
  const path = await (await pending).path();
  if (!path) throw new Error("download path unavailable");
  return unzipSync(new Uint8Array(readFileSync(path)));
}

function xml(parts: Record<string, Uint8Array>, name: string): string {
  const bytes = parts[name];
  if (!bytes) throw new Error(`missing ${name}`);
  return strFromU8(bytes);
}

function tags(source: string, name: "pgMar" | "pgSz"): string[] {
  return [...source.matchAll(new RegExp(`<w:${name}\\b[^>]*`, "g"))].map((match) => match[0]);
}

function value(tag: string, attribute: string): number {
  const match = new RegExp(`w:${attribute}="(\\d+)"`).exec(tag);
  if (!match) throw new Error(`missing ${attribute} in ${tag}`);
  return Number(match[1]);
}

async function relativeTextX(page: Page, text: string): Promise<number> {
  return page.evaluate((wanted) => {
    const item = [...document.querySelectorAll<HTMLElement>(".dxw-page span")]
      .find((element) => element.textContent === wanted);
    const sheet = item?.closest<HTMLElement>(".dxw-page");
    if (!item || !sheet) throw new Error(`missing geometry for ${wanted}`);
    return item.getBoundingClientRect().x - sheet.getBoundingClientRect().x;
  }, text);
}

test("Word margin presets, custom dialog, mirror geometry, and history", async ({ page }) => {
  await load(page);
  await openLayout(page);
  const margins = page.locator('[data-dxw-layout-menu-trigger="margins"]');
  await margins.click();
  const options = await page.locator('[data-dxw-layout-menu="margins"] [data-dxw-layout-option]').allTextContents();
  expect(options).toEqual(expect.arrayContaining([
    expect.stringContaining("Normal"), expect.stringContaining("Narrow"),
    expect.stringContaining("Moderate"), expect.stringContaining("Wide"),
    expect.stringContaining("Facing pages"),
    expect.stringContaining("Custom Margins"),
  ]));
  await page.keyboard.press("Escape");

  const presets = [
    ["m:normal", [1440, 1440, 1440, 1440], false],
    ["m:narrow", [720, 720, 720, 720], false],
    ["m:moderate", [1440, 1080, 1440, 1080], false],
    ["m:wide", [1440, 2880, 1440, 2880], false],
    ["m:mirrored", [1440, 1440, 1440, 1800], true],
  ] as const;
  for (const [preset, expected, mirrored] of presets) {
    await pickLayout(page, "margins", preset);
    const parts = await downloadParts(page);
    const documentXml = xml(parts, "word/document.xml");
    for (const tag of tags(documentXml, "pgMar")) {
      expect([value(tag, "top"), value(tag, "right"), value(tag, "bottom"), value(tag, "left")]).toEqual(expected);
    }
    const settings = xml(parts, "word/settings.xml");
    expect(settings.includes("mirrorMargins")).toBe(mirrored);
    if (mirrored) {
      expect(settings.indexOf("<w:zoom")).toBeLessThan(settings.indexOf("<w:mirrorMargins"));
      expect(settings.indexOf("<w:mirrorMargins")).toBeLessThan(settings.indexOf("<w:proofState"));
      expect(settings.indexOf("<w:mirrorMargins")).toBeLessThan(settings.indexOf("<w:compat"));
    }
  }

  expect(await relativeTextX(page, "Odd")).toBeCloseTo(120, 0);
  expect(await relativeTextX(page, "Even")).toBeCloseTo(96, 0);
  await page.locator('button[data-tab="home"]').click();
  await page.getByTitle("Undo (⌘Z)").click();
  expect(await relativeTextX(page, "Odd")).toBeCloseTo(192, 0);
  expect(await relativeTextX(page, "Even")).toBeCloseTo(192, 0);
  await page.getByTitle("Redo (⇧⌘Z)").click();
  expect(await relativeTextX(page, "Odd")).toBeCloseTo(120, 0);
  expect(await relativeTextX(page, "Even")).toBeCloseTo(96, 0);
  await openLayout(page);

  await pickLayout(page, "margins", "m:custom");
  const dialog = page.getByRole("dialog", { name: "Custom Margins" });
  await expect(dialog).toBeVisible();
  await expect(page.getByLabel("Top margin (inches)")).toBeFocused();
  await page.getByLabel("Top margin (inches)").fill("");
  await expect(dialog.getByRole("button", { name: "Apply" })).toBeDisabled();
  await page.getByLabel("Top margin (inches)").press("Escape");
  await expect(dialog).toBeHidden();

  await pickLayout(page, "margins", "m:custom");
  await page.getByLabel("Top margin (inches)").fill("0.2");
  await dialog.getByRole("button", { name: "Cancel" }).click();
  await expect(dialog).toBeHidden();
  const cancelledParts = await downloadParts(page);
  expect(xml(cancelledParts, "word/settings.xml")).toContain("mirrorMargins");
  for (const tag of tags(xml(cancelledParts, "word/document.xml"), "pgMar")) {
    expect([value(tag, "top"), value(tag, "right"), value(tag, "bottom"), value(tag, "left")]).toEqual([1440, 1440, 1440, 1800]);
  }

  await pickLayout(page, "margins", "m:custom");
  await page.getByLabel("Top margin (inches)").fill("0.6");
  await page.getByLabel("Bottom margin (inches)").fill("0.7");
  await page.getByLabel("Left margin (inches)").fill("0.8");
  await page.getByLabel("Right margin (inches)").fill("0.9");
  await dialog.getByRole("button", { name: "Apply" }).click();
  const customParts = await downloadParts(page);
  for (const tag of tags(xml(customParts, "word/document.xml"), "pgMar")) {
    expect([value(tag, "top"), value(tag, "right"), value(tag, "bottom"), value(tag, "left")]).toEqual([864, 1296, 1008, 1152]);
  }
  expect(xml(customParts, "word/settings.xml")).not.toContain("mirrorMargins");
});

test("all page sizes emit exact pgSz twips and borderless adds no metadata", async ({ page }) => {
  test.setTimeout(90_000);
  await load(page);
  await openLayout(page);
  const size = page.locator('[data-dxw-layout-menu-trigger="size"]');
  const groups = [
    [["letter"], 8.5, 11], [["legal"], 8.5, 14],
    [["3.5x5", "3.5x5-borderless"], 3.5, 5],
    [["4x6", "4x6-borderless"], 4, 6],
    [["5x7", "5x7-borderless"], 5, 7],
    [["8x10", "8x10-borderless"], 8, 10],
    [["a4", "a4-borderless"], 8.27, 11.69],
    [["a6"], 4.13, 5.83], [["envelope10"], 4.13, 9.5],
  ] as const;
  await size.click();
  expect(await page.locator('[data-dxw-layout-menu="size"] [data-dxw-layout-option]').allTextContents()).toEqual([
    'Letter8.5" × 11"', 'Legal8.5" × 14"',
    '3.5 × 53.5" × 5"', '3.5 × 5 Borderless3.5" × 5"',
    '4 × 64" × 6"', '4 × 6 Borderless4" × 6"',
    '5 × 75" × 7"', '5 × 7 Borderless5" × 7"',
    '8 × 108" × 10"', '8 × 10 Borderless8" × 10"',
    'A48.27" × 11.69"', 'A4 Borderless8.27" × 11.69"',
    'A64.13" × 5.83"', 'Envelope #104.13" × 9.5"',
    'Custom Paper Size…Set width and height in inches',
  ]);
  await page.keyboard.press("Escape");
  for (const [choices, width, height] of groups) {
    let rendered: { width: number; height: number } | null = null;
    for (const choice of choices) {
      await pickLayout(page, "size", choice);
      const box = await page.locator(".dxw-page").first().boundingBox();
      expect(box).not.toBeNull();
      if (rendered) {
        expect(box!.width).toBeCloseTo(rendered.width, 0);
        expect(box!.height).toBeCloseTo(rendered.height, 0);
      } else rendered = { width: box!.width, height: box!.height };
    }
    const parts = await downloadParts(page);
    const documentXml = xml(parts, "word/document.xml");
    for (const tag of tags(documentXml, "pgSz")) {
      expect([value(tag, "w"), value(tag, "h")]).toEqual([
        Math.round(width * 1440), Math.round(height * 1440),
      ]);
    }
    if (choices.some((choice) => choice.includes("borderless"))) {
      expect(documentXml.toLowerCase()).not.toContain("borderless");
    }
  }

  await pickLayout(page, "size", "custom");
  const custom = page.getByRole("dialog", { name: "Custom Paper Size" });
  await expect(custom).toBeVisible();
  await expect(page.getByLabel("Page width (inches)")).toBeFocused();
  await page.getByLabel("Page width (inches)").fill("7.25");
  await page.getByLabel("Page height (inches)").fill("10.5");
  await custom.getByRole("button", { name: "Apply" }).click();
  const customParts = await downloadParts(page);
  for (const tag of tags(xml(customParts, "word/document.xml"), "pgSz")) {
    expect([value(tag, "w"), value(tag, "h")]).toEqual([10440, 15120]);
  }

  await page.setViewportSize({ width: 1026, height: 800 });
  await pickLayout(page, "size", "letter");
  const noBleed = await page.evaluate(() => {
    const toolbar = document.querySelector<HTMLElement>("[data-dxw-toolbar-mode]")!;
    const pages = [...document.querySelectorAll<HTMLElement>(".dxw-page")];
    return {
      document: document.documentElement.scrollWidth <= window.innerWidth,
      toolbar: toolbar.getBoundingClientRect().right <= window.innerWidth + 1,
      pages: pages.every((sheet) => sheet.getBoundingClientRect().right <= window.innerWidth + 1),
    };
  });
  expect(noBleed).toEqual({ document: true, toolbar: true, pages: true });
});

test("layout menus are representative, exclusive, selection-safe, and viewport-clamped", async ({ page }) => {
  await load(page);
  await openLayout(page);
  const selectedRun = page.locator(".dxw-page span").filter({ hasText: /^Odd$/ });
  await selectedRun.click();
  const editorHasFocus = () => page.evaluate(() =>
    document.activeElement?.tagName === "TEXTAREA" && document.activeElement?.getAttribute("tabindex") === "-1",
  );
  expect(await editorHasFocus()).toBe(true);
  await page.locator("[data-dxw-toolbar-mode]").evaluate((toolbar) => {
    (toolbar as HTMLElement).style.setProperty("--dxw-popover-bg", "rgb(241, 247, 255)");
    (toolbar as HTMLElement).style.setProperty("--dxw-layout-preview-bg", "rgb(250, 252, 255)");
    (toolbar as HTMLElement).style.setProperty("--dxw-toolbar-z-index", "999");
  });

  const margins = page.locator('[data-dxw-layout-menu-trigger="margins"]');
  await margins.click();
  const marginMenu = page.locator('[data-dxw-layout-menu="margins"]');
  await expect(marginMenu).toHaveAttribute("role", "menu");
  await expect(marginMenu).toHaveCSS("background-color", "rgb(241, 247, 255)");
  await expect(marginMenu).toHaveCSS("z-index", "999");
  await expect(margins).toHaveAttribute("aria-expanded", "true");
  const marginOptions = marginMenu.locator('[data-dxw-layout-option]');
  expect(await marginOptions.count()).toBe(6);
  expect(await marginMenu.locator('[data-dxw-layout-preview="margins"]').count()).toBe(6);
  await expect(marginMenu.locator('[data-dxw-layout-option="m:moderate"]')).toContainText('1" top/bottom, 0.75" left/right');
  expect(await editorHasFocus()).toBe(true);
  await page.keyboard.press("Escape");
  await expect(marginMenu).toHaveCount(0);

  await margins.focus();
  await margins.press("Enter");
  const normal = page.locator('[data-dxw-layout-option="m:normal"]');
  const narrow = page.locator('[data-dxw-layout-option="m:narrow"]');
  const custom = page.locator('[data-dxw-layout-option="m:custom"]');
  await expect(normal).toBeFocused();
  await normal.press("ArrowDown");
  await expect(narrow).toBeFocused();
  await narrow.press("End");
  await expect(custom).toBeFocused();
  await custom.press("Home");
  await expect(normal).toBeFocused();
  await normal.press("Space");
  await expect(marginMenu).toHaveCount(0);
  await expect(margins).toBeFocused();

  await margins.press("Enter");
  await expect(normal).toBeFocused();
  await normal.press("Escape");
  await expect(marginMenu).toHaveCount(0);
  await expect(margins).toBeFocused();

  await margins.press("Enter");
  await expect(normal).toBeFocused();
  await normal.press("Tab");
  await expect(marginMenu).toHaveCount(0);
  await expect(margins).toBeFocused();

  await page.locator('[data-dxw-layout-menu-trigger="orientation"]').click();
  await expect(marginMenu).toHaveCount(0);
  await expect(page.locator('[data-dxw-layout-menu="orientation"]')).toBeVisible();
  await expect(page.locator('[data-dxw-layout-menu="orientation"] [data-dxw-layout-preview="orientation"]')).toHaveCount(2);
  await page.keyboard.press("Escape");
  await expect(page.locator("[data-dxw-layout-menu]")).toHaveCount(0);

  await page.locator('[data-dxw-layout-menu-trigger="columns"]').click();
  await expect(page.locator('[data-dxw-layout-menu="columns"]')).toBeVisible();
  await page.locator(".dxw-page").first().click({ position: { x: 300, y: 300 } });
  await expect(page.locator("[data-dxw-layout-menu]")).toHaveCount(0);

  await page.setViewportSize({ width: 520, height: 420 });
  await page.locator('[data-dxw-layout-menu-trigger="size"]').click();
  const menuBounds = await page.locator('[data-dxw-layout-menu="size"]').evaluate((menu) => {
    const rect = menu.getBoundingClientRect();
    return {
      left: rect.left,
      right: rect.right,
      top: rect.top,
      bottom: rect.bottom,
      scrollable: menu.scrollHeight > menu.clientHeight,
    };
  });
  expect(menuBounds.left).toBeGreaterThanOrEqual(7);
  expect(menuBounds.right).toBeLessThanOrEqual(513);
  expect(menuBounds.top).toBeGreaterThanOrEqual(7);
  expect(menuBounds.bottom).toBeLessThanOrEqual(413);
  expect(menuBounds.scrollable).toBe(true);

  const ribbonLayout = await page.locator("[data-dxw-layout-ribbon]").evaluate((ribbon) => {
    const controls = [...ribbon.querySelectorAll<HTMLElement>('select, [data-dxw-layout-menu-trigger]')]
      .filter((item) => item.offsetParent !== null)
      .map((item) => item.getBoundingClientRect());
    let overlap = false;
    for (let i = 0; i < controls.length; i++) {
      for (let j = i + 1; j < controls.length; j++) {
        const width = Math.min(controls[i].right, controls[j].right) - Math.max(controls[i].left, controls[j].left);
        const height = Math.min(controls[i].bottom, controls[j].bottom) - Math.max(controls[i].top, controls[j].top);
        if (width > 0.5 && height > 0.5) overlap = true;
      }
    }
    return { overlap, fits: ribbon.scrollWidth <= ribbon.clientWidth + 1 };
  });
  expect(ribbonLayout).toEqual({ overlap: false, fits: true });
});

test("orientation, columns, border, and line numbers round-trip through Download and reopen", async ({ page }) => {
  await load(page);
  await openLayout(page);
  await pickLayout(page, "size", "legal");
  await pickLayout(page, "orientation", "landscape");
  await pickLayout(page, "columns", "2");
  await pickLayout(page, "page-border", "accent");
  await pickLayout(page, "line-numbers", "continuous");

  const first = await downloadParts(page);
  const firstXml = xml(first, "word/document.xml");
  expect(firstXml).toMatch(/<w:pgSz\b[^>]*w:w="20160"[^>]*w:h="12240"[^>]*w:orient="landscape"/);
  expect(firstXml).toMatch(/<w:cols\b[^>]*w:num="2"/);
  expect(firstXml).toMatch(/<w:pgBorders\b/);
  expect(firstXml).toMatch(/<w:(top|left|bottom|right)\b[^>]*w:color="4472C4"/);
  expect(firstXml).toMatch(/<w:lnNumType\b[^>]*w:countBy="1"[^>]*w:restart="continuous"/);

  await page.locator("#docx-upload").setInputFiles({
    name: "layout-round-trip.docx",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    buffer: Buffer.from(zipSync(first)),
  });
  await expect(page.locator(".file-name")).toHaveText("layout-round-trip.docx");
  await expect(page.locator("main")).toContainText("Odd page");
  const reopened = xml(await downloadParts(page), "word/document.xml");
  expect(reopened).toMatch(/<w:pgSz\b[^>]*w:w="20160"[^>]*w:h="12240"[^>]*w:orient="landscape"/);
  expect(reopened).toMatch(/<w:cols\b[^>]*w:num="2"/);
  expect(reopened).toContain("4472C4");
  expect(reopened).toMatch(/<w:lnNumType\b[^>]*w:countBy="1"[^>]*w:restart="continuous"/);
});

test("section scope changes selected pgMar while mirror mode remains global", async ({ page }) => {
  await load(page);
  await page.locator(".dxw-page span").filter({ hasText: /^Second$/ }).click();
  await openLayout(page);
  await page.locator('select[title="Apply layout changes to"]').selectOption("section");
  await pickLayout(page, "margins", "m:mirrored");

  const mirroredParts = await downloadParts(page);
  const mar = tags(xml(mirroredParts, "word/document.xml"), "pgMar");
  expect(mar).toHaveLength(2);
  expect([value(mar[0], "right"), value(mar[0], "left")]).toEqual([1440, 1440]);
  expect([value(mar[1], "right"), value(mar[1], "left")]).toEqual([1440, 1800]);
  expect(xml(mirroredParts, "word/settings.xml")).toContain("mirrorMargins");

  await pickLayout(page, "margins", "m:custom");
  const dialog = page.getByRole("dialog", { name: "Custom Margins" });
  await expect(dialog).toContainText("Applies to this section");
  await page.getByLabel("Left margin (inches)").fill("0.75");
  await dialog.getByRole("button", { name: "Apply" }).click();
  const customParts = await downloadParts(page);
  const customMar = tags(xml(customParts, "word/document.xml"), "pgMar");
  expect(value(customMar[0], "left")).toBe(1440);
  expect(value(customMar[1], "left")).toBe(1080);
  expect(xml(customParts, "word/settings.xml")).not.toContain("mirrorMargins");
});
