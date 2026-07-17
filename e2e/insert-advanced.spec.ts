import { expect, Page, test } from "@playwright/test";
import { readFileSync } from "node:fs";
import { strFromU8, unzipSync } from "fflate";

const MOD = process.platform === "darwin" ? "Meta" : "Control";

async function open(page: Page): Promise<void> {
  await page.goto("/?doc=/fixtures/benchmark.docx");
  await page.waitForSelector(".dxw-page span");
  await page.getByRole("button", { name: "insert", exact: true }).click();
}

async function downloadXml(page: Page): Promise<string> {
  const pending = page.waitForEvent("download");
  await page.getByText("Download", { exact: true }).click();
  const path = await (await pending).path();
  expect(path).not.toBeNull();
  return strFromU8(unzipSync(new Uint8Array(readFileSync(path!)))["word/document.xml"]);
}

function tool(page: Page, title: string) {
  return page.locator(`[title=${JSON.stringify(title)}], [data-tip=${JSON.stringify(title)}]`).first();
}

function classic(page: Page) {
  return page.locator(".dxw-page span", { hasText: "classic." }).first();
}

test.describe("advanced Insert ribbon", () => {
  test("bookmark and live cross-references insert, undo, redo, and save", async ({ page }) => {
    await open(page);
    const title = page.locator('.dxw-page span:text-is("Typography")').first();
    await title.dblclick();
    await tool(page, "Insert bookmark").click();
    await page.getByPlaceholder("Quarterly_Revenue").fill("TypographyTarget");
    await page.getByRole("button", { name: "Add", exact: true }).click();

    await classic(page).click();
    await page.keyboard.press("End");
    await tool(page, "Insert cross-reference").click();
    await page.getByRole("button", { name: "Bookmark text", exact: true }).click();
    await expect(page.locator('.dxw-page span:text-is("Typography")')).toHaveCount(2);

    await classic(page).click();
    await page.keyboard.press(`${MOD}+z`);
    await expect(page.locator('.dxw-page span:text-is("Typography")')).toHaveCount(1);
    await page.keyboard.press(`${MOD}+Shift+z`);
    await expect(page.locator('.dxw-page span:text-is("Typography")')).toHaveCount(2);

    await tool(page, "Insert cross-reference").click();
    await page.getByRole("button", { name: "Page number", exact: true }).click();
    const xml = await downloadXml(page);
    expect(xml).toContain('w:name="TypographyTarget"');
    expect(xml).toContain("REF TypographyTarget");
    expect(xml).toContain("PAGEREF TypographyTarget");
  });

  test("date/time and generic fields render and save as live Word fields", async ({ page }) => {
    await open(page);
    await classic(page).click();
    await page.keyboard.press("End");
    await tool(page, "Insert an automatically updating date or time").selectOption("date:long");
    await tool(page, "Insert a Word field").selectOption("NUMPAGES");

    await expect(page.locator(".dxw-page").filter({ hasText: /[A-Z][a-z]+ \d{1,2}, \d{4}/ }).first()).toBeVisible();
    const xml = await downloadXml(page);
    expect(xml).toContain('DATE \\@ &quot;MMMM d, yyyy&quot; \\* MERGEFORMAT');
    expect(xml).toContain("NUMPAGES \\* MERGEFORMAT");
  });

  test("equations and advanced symbols use native editable content and round-trip", async ({ page }) => {
    await open(page);
    await classic(page).click();
    await page.keyboard.press("End");

    await tool(page, "Insert equation").click();
    await page.getByLabel("Linear equation").fill("{x+1}/{2y}");
    await page.getByRole("button", { name: "Insert", exact: true }).click();

    await classic(page).click();
    await page.keyboard.press("End");
    await tool(page, "Insert advanced symbol").click();
    await page.getByTitle("Insert Ω").click();
    await expect(page.locator(".dxw-page").filter({ hasText: "Ω" }).first()).toBeVisible();

    await classic(page).click();
    await page.keyboard.press(`${MOD}+z`);
    await expect(page.locator(".dxw-page").filter({ hasText: "Ω" })).toHaveCount(0);
    await page.keyboard.press(`${MOD}+Shift+z`);
    await expect(page.locator(".dxw-page").filter({ hasText: "Ω" }).first()).toBeVisible();

    const xml = await downloadXml(page);
    expect(xml).toContain("<m:oMath>");
    expect(xml).toContain("<m:f>");
    expect(xml).toContain("Ω");
  });

  test("advanced symbol accepts arbitrary Unicode characters", async ({ page }) => {
    await open(page);
    await classic(page).click();
    await page.keyboard.press("End");
    await tool(page, "Insert advanced symbol").click();
    await page.getByLabel("Advanced symbol characters").fill("ℏ⊕");
    await page.getByRole("button", { name: "Insert", exact: true }).click();

    await expect(page.locator(".dxw-page").filter({ hasText: "ℏ⊕" }).first()).toBeVisible();
    const xml = await downloadXml(page);
    expect(xml).toContain("ℏ⊕");
  });

  test("dedicated Text Box inserts an editable native text-box story and saves", async ({ page }) => {
    await open(page);
    await classic(page).click();
    await page.keyboard.press("End");
    await tool(page, "Insert text box").click();
    await page.getByLabel("Text box text").fill("Editable text box");
    await page.getByRole("button", { name: "Insert", exact: true }).click();

    const story = page.locator('[data-dxw-textbox-story]').filter({ hasText: "Editable" }).last();
    await expect(story).toBeVisible();
    await story.dblclick();
    await page.keyboard.press("End");
    await page.keyboard.type(" revised");
    await expect(page.locator('[data-dxw-textbox-story]').filter({ hasText: "revised" }).last()).toBeVisible();
    const xml = await downloadXml(page);
    expect(xml).toContain('name="Text Box ');
    expect(xml).toContain("Editable text box");
    expect(xml).toContain("revised");
    expect(xml).toContain("<wps:txbx>");
  });
});
