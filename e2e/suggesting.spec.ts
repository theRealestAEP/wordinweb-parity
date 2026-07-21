import { expect, test } from "@playwright/test";
import { Editor } from "./editing.js";

/**
 * Suggesting mode ("my own tracked changes"): typing records as w:ins (rendered
 * underlined + author-colored), deleting records as w:del (struck, text kept),
 * and clicking a change offers accept/reject. Toggling the mode off returns to
 * direct editing. XML shape is covered by the suggest.test.ts unit suite; these
 * specs assert the live rendering and the review flows in the real browser.
 */

const INS = "rgb(192, 0, 0)"; // #C00000
const DEL = "rgb(176, 38, 28)"; // #B0261C

test.describe("suggesting mode", () => {
  test("typing records an insertion, rendered underlined and colored", async ({ page }) => {
    const ed = await Editor.open(page, "sample", "&suggest=1");
    await ed.clickText("time");
    await ed.type("ZQXINS");
    await ed.expectHasText("ZQXINS");
    const style = await ed.spanStyle("ZQXINS");
    expect(style).not.toBeNull();
    expect(style!.color).toBe(INS);
    expect(style!.decoration).toContain("underline");
  });

  test("deleting a selected word records a deletion, keeping the struck text", async ({ page }) => {
    const ed = await Editor.open(page, "sample", "&suggest=1");
    await ed.span("time").dblclick(); // select the word
    await page.waitForTimeout(120);
    await ed.press("Delete");
    // The text is NOT removed — it stays, painted in the deletion color (the
    // strike itself is an engine-drawn rule, not a CSS text-decoration).
    await ed.expectHasText("time");
    const style = await ed.spanStyle("time");
    expect(style).not.toBeNull();
    expect(style!.color).toBe(DEL);
  });

  test("clicking a suggestion then Accept keeps the text as normal", async ({ page }) => {
    const ed = await Editor.open(page, "sample", "&suggest=1");
    await ed.clickText("time");
    await ed.type("ZQXACC");
    await ed.span("ZQXACC").click(); // opens the accept/reject popover
    await page.waitForTimeout(120);
    await expect(page.locator('button[title$="suggestion"]')).toHaveCount(2);
    await ed.reviewClick("Accept");
    await ed.expectHasText("ZQXACC");
    const style = await ed.spanStyle("ZQXACC");
    expect(style!.color).not.toBe(INS); // no longer an insertion color
    expect(style!.decoration).not.toContain("underline");
  });

  test("clicking a suggestion then Reject removes the inserted text", async ({ page }) => {
    const ed = await Editor.open(page, "sample", "&suggest=1");
    await ed.clickText("time");
    await ed.type("ZQXREJ");
    await ed.span("ZQXREJ").click();
    await page.waitForTimeout(120);
    await ed.reviewClick("Reject");
    const found = await page.evaluate(() =>
      [...document.querySelectorAll(".dxw-page span")].some((s) => (s.textContent ?? "").includes("ZQXREJ")),
    );
    expect(found).toBe(false);
  });

  test("switching the mode dropdown to Editing returns to direct editing", async ({ page }) => {
    const ed = await Editor.open(page, "sample", "&suggest=1");
    // Pick "Editing" from the pencil dropdown (replaces the old toggle).
    await page.locator("[data-dxw-mode]").click();
    await page.locator('[data-dxw-mode-option="editing"]').click();
    await page.waitForTimeout(150);
    await ed.clickText("time");
    await ed.type("ZQXPLAIN");
    await ed.expectHasText("ZQXPLAIN");
    const style = await ed.spanStyle("ZQXPLAIN");
    expect(style!.color).not.toBe(INS);
    expect(style!.decoration).not.toContain("underline");
  });

  test("the mode dropdown reflects and switches editing / suggesting / viewing", async ({ page }) => {
    await Editor.open(page, "sample"); // opens in editing mode (no ?suggest=1)
    const trigger = page.locator("[data-dxw-mode]");
    await expect(trigger).toHaveAttribute("data-dxw-mode", "editing");

    // Enter suggesting: the trigger's state attribute flips and the author chip appears.
    await trigger.click();
    await page.locator('[data-dxw-mode-option="suggesting"]').click();
    await expect(trigger).toHaveAttribute("data-dxw-mode", "suggesting");
    await expect(page.getByText("You", { exact: true })).toBeVisible(); // author chip

    // Viewing disables editing (no toolbar rendered).
    await trigger.click();
    await page.locator('[data-dxw-mode-option="viewing"]').click();
    await expect(trigger).toHaveAttribute("data-dxw-mode", "viewing");

    // Back to editing.
    await trigger.click();
    await page.locator('[data-dxw-mode-option="editing"]').click();
    await expect(trigger).toHaveAttribute("data-dxw-mode", "editing");
  });

  test("suggestions survive switching to Viewing and back (no reparse data loss)", async ({ page }) => {
    const ed = await Editor.open(page, "sample", "&suggest=1");
    await ed.clickText("time");
    await ed.type("ZQXKEEP");
    await ed.expectHasText("ZQXKEEP");

    // Viewing: mode switches used to re-parse the original bytes, silently
    // dropping every unsaved edit. The suggestion must still be there (it
    // renders as final text while viewing) and the review pill still counts it.
    const trigger = page.locator("[data-dxw-mode]");
    await trigger.click();
    await page.locator('[data-dxw-mode-option="viewing"]').click();
    await page.waitForTimeout(400);
    await ed.expectHasText("ZQXKEEP");
    await expect(page.locator("[data-dxw-review-bar]")).toContainText("1 suggestion");

    // Back to suggesting: the same pending insertion renders as markup again.
    await trigger.click();
    await page.locator('[data-dxw-mode-option="suggesting"]').click();
    await page.waitForTimeout(400);
    await ed.expectHasText("ZQXKEEP");
    const style = await ed.spanStyle("ZQXKEEP");
    expect(style!.color).toBe(INS);
  });

  test("the review pill counts suggestions and Accept all applies them", async ({ page }) => {
    const ed = await Editor.open(page, "sample", "&suggest=1");
    await ed.clickText("time");
    await ed.type("ZQXONE");
    await ed.clickText("dolore");
    await ed.type("ZQXTWO");
    const pill = page.locator("[data-dxw-review-bar]");
    await expect(pill).toContainText("2 suggestions");

    await page.locator("[data-dxw-accept-all]").click();
    await page.waitForTimeout(300);
    await expect(pill).toHaveCount(0); // nothing pending
    await ed.expectHasText("ZQXONE");
    await ed.expectHasText("ZQXTWO");
    const style = await ed.spanStyle("ZQXONE");
    expect(style!.color).not.toBe(INS); // permanent text now
  });

  test("Reject all removes pending insertions", async ({ page }) => {
    const ed = await Editor.open(page, "sample", "&suggest=1");
    await ed.clickText("time");
    await ed.type("ZQXDROP");
    await expect(page.locator("[data-dxw-review-bar]")).toContainText("1 suggestion");
    await page.locator("[data-dxw-reject-all]").click();
    await page.waitForTimeout(300);
    await expect(page.locator("[data-dxw-review-bar]")).toHaveCount(0);
    const found = await page.evaluate(() =>
      [...document.querySelectorAll(".dxw-page span")].some((s) => (s.textContent ?? "").includes("ZQXDROP")),
    );
    expect(found).toBe(false);
  });

  test("a table-cell suggestion accepts cleanly and survives save/reopen", async ({ page }) => {
    const ed = await Editor.open(page, "parity-tables", "&suggest=1");
    const status = (await ed.span("Status").boundingBox())!;
    const descriptionBefore = (await ed.span("Description").boundingBox())!;
    await page.mouse.click(status.x + status.width - 1, status.y + status.height / 2);
    await page.keyboard.type("ZQXTABLE");
    await ed.expectHasText("ZQXTABLE");
    const style = await ed.spanStyle("ZQXTABLE");
    expect(style?.color).toBe(INS);
    await expect(page.locator("[data-dxw-review-bar]")).toContainText("1 suggestion");

    await page.locator("[data-dxw-accept-all]").click();
    await expect(page.locator("[data-dxw-review-bar]")).toHaveCount(0);
    const accepted = await ed.spanStyle("ZQXTABLE");
    expect(accepted?.color).not.toBe(INS);
    const inserted = (await ed.span("ZQXTABLE").boundingBox())!;
    const descriptionAfter = (await ed.span("Description").boundingBox())!;
    expect(descriptionAfter.x).toBeGreaterThan(inserted.x + inserted.width);
    expect(descriptionAfter.y).toBeCloseTo(descriptionBefore.y, 0);

    const pending = page.waitForEvent("download");
    await page.getByText("Download", { exact: true }).click();
    const path = await (await pending).path();
    if (!path) throw new Error("download path unavailable");
    await page.locator("#docx-upload").setInputFiles(path);
    await expect(page.locator(".dxw-pages")).toContainText("ZQXTABLE");
  });
});
