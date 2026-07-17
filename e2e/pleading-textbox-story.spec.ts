import { expect, Page, test } from "@playwright/test";
import { readFileSync } from "node:fs";
import { strFromU8, unzipSync } from "fflate";

const MOD = process.platform === "darwin" ? "Meta" : "Control";

async function open(page: Page): Promise<void> {
  await page.goto("/?doc=/fixtures/pleading-paper.docx");
  await page.waitForSelector("[data-dxw-textbox-story-object]", { timeout: 10_000 });
  await page.waitForTimeout(300);
}

const hfMode = (page: Page) => page.locator(".dxw-hf-mode");
const storyText = (page: Page) => page.locator("[data-dxw-textbox-story]");
const storyObject = (page: Page) => page.locator("[data-dxw-textbox-story-object]");

function headerStories(files: Record<string, Uint8Array>): [string, string][] {
  return Object.entries(files)
    .filter(([name]) => /^word\/header\d+\.xml$/.test(name))
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([name, bytes]) => {
      const xml = strFromU8(bytes);
      const story = xml.match(/<w:txbxContent>[\s\S]*?<\/w:txbxContent>/)?.[0] ?? "";
      return [name, [...story.matchAll(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g)].map((match) => match[1]).join("|")];
    });
}

async function downloadDocx(page: Page): Promise<string> {
  const download = await Promise.all([
    page.waitForEvent("download"),
    page.getByRole("button", { name: "Download" }).click(),
  ]).then(([result]) => result);
  const path = await download.path();
  if (!path) throw new Error("download path missing");
  return path;
}

test.describe("header-owned pleading textbox story", () => {
  test("selects and edits the gutter independently, scopes Cmd+A, and exits on body click", async ({ page }) => {
    await open(page);
    const bodyBefore = await page.locator(".dxw-page span:not([data-dxw-hf])").allTextContents();
    const firstObject = storyObject(page).first();
    const objectBox = await firstObject.boundingBox();
    if (!objectBox) throw new Error("pleading gutter missing");

    await firstObject.click({ position: { x: 2, y: objectBox.height / 2 } });
    await expect(page.locator("[data-dxw-object-selection]")).toBeVisible();
    await expect(hfMode(page)).toHaveCount(0);

    await storyText(page).first().dblclick();
    await expect(hfMode(page)).toHaveCount(0);
    await expect(page.locator("[data-dxw-caret]")).toBeVisible();
    await page.keyboard.press(`${MOD}+a`);
    await page.keyboard.type("GUTTER-STORY");
    await expect.poll(async () => (await storyText(page).allTextContents()).join(""))
      .toContain("GUTTER-STORY");
    expect(await page.locator(".dxw-page span:not([data-dxw-hf])").allTextContents()).toEqual(bodyBefore);
    const editedStory = (await storyText(page).allTextContents()).join("");

    const path = await downloadDocx(page);
    const files = unzipSync(readFileSync(path));
    const editedHeader = Object.entries(files)
      .filter(([name]) => /^word\/header\d+\.xml$/.test(name))
      .map(([, bytes]) => strFromU8(bytes))
      .find((xml) => xml.includes("GUTTER-STORY"));
    expect(editedHeader).toBeTruthy();
    expect(editedHeader).toContain('id="LineNumbers"');
    expect(editedHeader).toContain("margin-left:-47.15pt");
    expect(editedHeader!.match(/<w:txbxContent>[\s\S]*?<\/w:txbxContent>/)?.[0])
      .toContain("GUTTER-STORY");
    expect(strFromU8(files["word/document.xml"])).toContain("Cofazeta W. Ruzezegom");

    // Prove reopen uses the downloaded bytes, rather than the still-mounted
    // editor state: temporarily change the live story, then upload the save.
    await page.keyboard.press(`${MOD}+a`);
    await page.keyboard.type("TEMP-STORY");
    await expect.poll(async () => (await storyText(page).allTextContents()).join(""))
      .toContain("TEMP-STORY");
    await page.locator("#docx-upload").setInputFiles(path);
    await expect.poll(async () => (await storyText(page).allTextContents()).join(""), { timeout: 10_000 })
      .toContain("GUTTER-STORY");
    expect((await storyText(page).allTextContents()).join("")).not.toContain("TEMP-STORY");
    await expect(hfMode(page)).toHaveCount(0);

    // A blank body-margin click exits story mode even though the full-height
    // gutter remains the vertically nearest header text.
    const firstPage = page.locator(".dxw-page").first();
    const pageBox = await firstPage.boundingBox();
    if (!pageBox) throw new Error("page missing");
    await page.mouse.click(pageBox.x + pageBox.width - 18, pageBox.y + pageBox.height / 2);
    await expect(page.locator("[data-dxw-caret]")).toBeVisible();
    await page.keyboard.type("BODY-TOKEN");
    expect((await storyText(page).allTextContents()).join("")).toBe(editedStory);

    await storyText(page).first().click();
    await expect(page.locator("[data-dxw-object-selection]")).toBeVisible();
    await expect(hfMode(page)).toHaveCount(0);

  });

  test("body Select All leaves every header textbox story unchanged", async ({ page }) => {
    await open(page);
    const response = await page.request.get("/fixtures/pleading-paper.docx");
    const before = headerStories(unzipSync(await response.body()));
    const body = page.locator(".dxw-page span:not([data-dxw-hf])").first();
    await body.click();
    await page.keyboard.press(`${MOD}+a`);
    await page.keyboard.type("BODY-ONLY");
    const after = headerStories(unzipSync(readFileSync(await downloadDocx(page))));
    expect(after).toEqual(before);
  });

  test("Escape returns to the selected object on the active repeated page", async ({ page }) => {
    await open(page);
    const objects = storyObject(page);
    expect(await objects.count()).toBeGreaterThan(1);
    const index = Math.min(2, (await objects.count()) - 1);
    const object = objects.nth(index);
    const objectPageEl = object.locator("xpath=ancestor::*[contains(concat(' ', normalize-space(@class), ' '), ' dxw-page ')]");
    const text = objectPageEl.locator("[data-dxw-textbox-story]").first();
    await text.scrollIntoViewIfNeeded();
    await text.dblclick();
    await expect(page.locator("[data-dxw-caret]")).toBeVisible();
    await expect(hfMode(page)).toHaveCount(0);
    await page.keyboard.press("Escape");
    const selectedPage = await page.locator("[data-dxw-object-selection]").evaluate((element) =>
      Number((element.closest(".dxw-page") as HTMLElement | null)?.dataset.page),
    );
    const objectPage = await object.evaluate((element) =>
      Number((element.closest(".dxw-page") as HTMLElement | null)?.dataset.page),
    );
    expect(selectedPage).toBe(objectPage);
    await expect(hfMode(page)).toHaveCount(0);
  });

  test("ordinary top-margin editing stays outside txbxContent", async ({ page }) => {
    await open(page);
    const pageEl = page.locator(".dxw-page").first();
    const geometry = await pageEl.evaluate((element) => {
      const rect = element.getBoundingClientRect();
      return {
        x: rect.x,
        y: rect.y,
        width: rect.width,
        bodyTop: Number((element as HTMLElement).dataset.bodyTop || 96),
      };
    });
    await page.mouse.dblclick(
      geometry.x + geometry.width / 2,
      geometry.y + Math.max(8, geometry.bodyTop / 2),
    );
    await expect(hfMode(page)).toHaveCount(1);
    await page.keyboard.type("OUTER-HEADER");
    await page.keyboard.press(`${MOD}+a`);
    await page.keyboard.type("HEADER-ONLY");
    await expect.poll(async () => (await storyText(page).allTextContents()).join(""))
      .toContain("123456789");

    const path = await downloadDocx(page);
    const files = unzipSync(readFileSync(path));
    const headers = Object.entries(files)
      .filter(([name]) => /^word\/header\d+\.xml$/.test(name))
      .map(([, bytes]) => strFromU8(bytes));
    const edited = headers.find((xml) => xml.includes("HEADER-ONLY"));
    expect(edited).toBeTruthy();
    expect(edited).toContain('id="LineNumbers"');
    expect(edited).toContain("margin-left:-47.15pt");
    const outsideStory = edited!.replace(/<w:txbxContent>[\s\S]*?<\/w:txbxContent>/g, "");
    expect(outsideStory).toContain("HEADER-ONLY");

  });
});
