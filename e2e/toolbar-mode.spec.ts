import { expect, test } from "@playwright/test";

test("simple editing exposes the basic Home toolbar", async ({ page }) => {
  await page.goto("/?toolbar=simple");
  const toolbar = page.locator('[data-dxw-toolbar-mode="simple"]');
  await expect(toolbar).toBeVisible();
  await expect(toolbar.getByTitle("Bold (⌘B)")).toBeVisible();
  await expect(toolbar.locator("[data-tab]")).toHaveCount(0);
});

test("advanced editing exposes the full ribbon", async ({ page }) => {
  await page.goto("/");
  const toolbar = page.locator('[data-dxw-toolbar-mode="advanced"]');
  await expect(toolbar).toBeVisible();
  await expect(toolbar.locator("[data-tab]")).toHaveCount(4);
  await toolbar.getByRole("button", { name: "draw", exact: true }).click();
  await expect(toolbar.getByTitle("Draw with pen")).toBeVisible();
});

test("advanced stays the default and feature overrides hide their ribbon", async ({ page }) => {
  await page.goto("/?layout=off");
  const toolbar = page.locator('[data-dxw-toolbar-mode="advanced"]');
  await expect(toolbar).toBeVisible();
  await expect(toolbar.getByRole("button", { name: "home", exact: true })).toBeVisible();
  await expect(toolbar.getByRole("button", { name: "layout", exact: true })).toHaveCount(0);
});
