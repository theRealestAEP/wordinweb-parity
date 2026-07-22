import { expect, Page, test } from "@playwright/test";
import { readFileSync } from "node:fs";
import { strFromU8, unzipSync } from "fflate";

const MOD = process.platform === "darwin" ? "Meta" : "Control";

async function openInsert(page: Page): Promise<void> {
  await page.goto("/?doc=/fixtures/benchmark.docx");
  await page.waitForSelector(".dxw-page span");
  await page.getByRole("button", { name: "insert", exact: true }).click();
  const target = page.locator(".dxw-page span", { hasText: "classic." }).first();
  await target.click();
  await page.keyboard.press("End");
}

async function submitSmartArt(page: Page, layout: "process" | "cycle" | "hierarchy" | "list", items: string): Promise<void> {
  const edit = page.getByRole("button", { name: "Edit SmartArt", exact: true });
  await (await edit.isVisible() ? edit : page.getByRole("button", { name: "SmartArt", exact: true })).click();
  const label = `${layout[0].toUpperCase()}${layout.slice(1)} SmartArt`;
  const layoutButton = page.getByRole("button", { name: label, exact: true });
  if (!(await layoutButton.isVisible())) {
    await page.getByRole("button", { name: "Back", exact: true }).click();
  }
  await page.getByRole("button", { name: label, exact: true }).click();
  const values = items.split("\n");
  let count = await page.getByLabel(/^SmartArt item \d+$/).count();
  while (count < values.length) {
    await page.getByRole("button", { name: "Add item" }).click();
    count++;
  }
  while (count > values.length) {
    await page.getByLabel(`Remove SmartArt item ${count}`).click();
    count--;
  }
  for (let index = 0; index < values.length; index++) {
    await page.getByRole("textbox", { name: `SmartArt item ${index + 1}`, exact: true }).fill(values[index]);
  }
  await page.getByRole("dialog").getByRole("button", { name: /^(Insert|Update)$/ }).click();
}

test("advanced Insert creates and edits native SmartArt with undo, redo, and save", async ({ page }) => {
  await openInsert(page);
  await page.getByRole("button", { name: "SmartArt", exact: true }).click();
  await expect(page.getByRole("dialog", { name: "Insert SmartArt" })).toBeVisible();
  await expect(page.getByRole("button", { name: "Process SmartArt", exact: true })).toBeVisible();
  await expect(page.getByRole("button", { name: "Cycle SmartArt", exact: true })).toBeVisible();
  await expect(page.getByRole("button", { name: "Hierarchy SmartArt", exact: true })).toBeVisible();
  await expect(page.getByRole("button", { name: "List SmartArt", exact: true })).toBeVisible();
  await expect(page.getByLabel(/^SmartArt item \d+$/)).toHaveCount(0);
  await page.getByRole("button", { name: "Cancel", exact: true }).click();
  await submitSmartArt(page, "process", "Discover\nDesign\nDeliver");
  await expect(page.getByText("Discover", { exact: true })).toBeVisible();
  await expect(page.getByText("Deliver", { exact: true })).toBeVisible();

  const smartArt = page.locator("[data-dxw-smart-art]").last();
  const smartArtBox = await smartArt.boundingBox();
  expect(smartArtBox).not.toBeNull();
  await page.mouse.click(smartArtBox!.x + 6, smartArtBox!.y + 6);
  await expect(page.locator("[data-dxw-img-handle]")).toHaveCount(8);
  await expect(page.getByRole("button", { name: "SmartArt Format", exact: true })).toBeVisible();
  await expect(page.locator("[data-dxw-object-format]").getByRole("button", { name: "Outline", exact: true })).toHaveCount(0);
  await expect(page.locator("[data-dxw-object-format]").getByRole("button", { name: "Edit text", exact: true })).toHaveCount(0);
  await page.getByRole("button", { name: "Edit SmartArt", exact: true }).click();
  await expect(page.getByRole("dialog", { name: "Edit SmartArt" })).toBeVisible();
  await expect(page.getByRole("textbox", { name: "SmartArt item 1", exact: true })).toHaveValue("Discover");
  await expect(page.getByRole("textbox", { name: "SmartArt item 2", exact: true })).toHaveValue("Design");
  await expect(page.getByRole("textbox", { name: "SmartArt item 3", exact: true })).toHaveValue("Deliver");
  await page.getByRole("button", { name: "Cancel", exact: true }).click();
  await submitSmartArt(page, "hierarchy", "Lead\nPlan\nBuild\nTest");
  await expect(page.getByText("Lead", { exact: true })).toBeVisible();
  await expect(page.getByText("Test", { exact: true })).toBeVisible();

  await page.keyboard.press(`${MOD}+z`);
  await expect(page.getByText("Discover", { exact: true })).toBeVisible();
  await page.keyboard.press(`${MOD}+Shift+z`);
  await expect(page.getByText("Lead", { exact: true })).toBeVisible();

  const secondNode = page.locator('[data-dxw-smart-art-node="1"]').last();
  const secondNodeBox = await secondNode.boundingBox();
  expect(secondNodeBox).not.toBeNull();
  await page.mouse.click(secondNodeBox!.x + 6, secondNodeBox!.y + 6);
  await expect(page.locator('[data-dxw-smart-art-node-selection="1"]')).toBeVisible();
  await page.keyboard.type("Revised plan");
  const nodeEditor = page.getByRole("textbox", { name: "Edit SmartArt node 2", exact: true });
  await expect(nodeEditor).toHaveValue("Revised plan");
  await page.keyboard.press("Enter");
  await expect(page.locator(".dxw-page").first()).toContainText("Revised plan");
  await expect(page.locator('[data-dxw-smart-art-node-selection="1"]')).toBeVisible();
  await page.locator("[data-dxw-object-format]").getByRole("button", { name: "Node fill", exact: true }).click();
  await page.getByLabel("Fill color", { exact: true }).fill("#AA22CC");
  await page.getByRole("button", { name: "Apply", exact: true }).click();
  await expect.poll(() => page.locator("svg path").evaluateAll((paths) =>
    paths.filter((path) => path.getAttribute("fill")?.toLowerCase() === "#aa22cc").length,
  )).toBe(1);

  const recoloredNodeBox = await page.locator('[data-dxw-smart-art-node="1"]').last().boundingBox();
  expect(recoloredNodeBox).not.toBeNull();
  await page.mouse.click(recoloredNodeBox!.x + 6, recoloredNodeBox!.y + 6);
  await page.locator("[data-dxw-object-format]").getByRole("button", { name: "Node fill", exact: true }).click();
  await expect(page.getByLabel("No fill")).not.toBeChecked();
  await expect(page.getByLabel("Fill color", { exact: true })).toHaveValue("#AA22CC");
  await page.getByRole("button", { name: "Cancel", exact: true }).click();

  const formatBar = page.locator("[data-dxw-object-format]");
  await formatBar.locator("[data-dxw-menu-select-trigger]").nth(0).click();
  await page.getByRole("option", { name: "Arial", exact: true }).click();
  await formatBar.locator("[data-dxw-menu-select-trigger]").nth(1).click();
  await page.getByRole("option", { name: "18", exact: true }).click();
  await formatBar.getByRole("button", { name: "B", exact: true }).click();
  await formatBar.locator("[data-dxw-color-trigger]").click();
  await page.getByLabel("Choose #000000", { exact: true }).click();
  const revisedWord = page.getByText("Revised", { exact: true });
  await expect.poll(() => revisedWord.evaluate((element) => getComputedStyle(element).fontFamily)).toContain("Arial");
  await expect.poll(() => revisedWord.evaluate((element) => getComputedStyle(element).fontSize)).toBe("24px");
  await expect.poll(() => revisedWord.evaluate((element) => getComputedStyle(element).fontWeight)).toBe("400");
  await expect.poll(() => revisedWord.evaluate((element) => getComputedStyle(element).color)).toBe("rgb(0, 0, 0)");

  const pending = page.waitForEvent("download");
  await page.getByText("Download", { exact: true }).click();
  const path = await (await pending).path();
  expect(path).not.toBeNull();
  const files = unzipSync(new Uint8Array(readFileSync(path!)));
  expect(strFromU8(files["word/document.xml"])).toContain("<dgm:relIds");
  expect(strFromU8(files["word/diagrams/data1.xml"])).toContain("Lead");
  expect(strFromU8(files["word/diagrams/data1.xml"])).toContain("Revised plan");
  expect(strFromU8(files["word/diagrams/layout1.xml"])).toContain("smartart:hierarchy");
  expect(strFromU8(files["word/diagrams/quickStyle1.xml"])).toContain("<dgm:styleDef");
  expect(strFromU8(files["word/diagrams/colors1.xml"])).toContain("<dgm:colorsDef");
  expect(strFromU8(files["word/diagrams/drawing1.xml"])).toContain("<dsp:drawing");
  expect(strFromU8(files["word/diagrams/drawing1.xml"])).toContain("Revised plan");
  expect(strFromU8(files["word/diagrams/drawing1.xml"])).toContain('<a:srgbClr val="AA22CC"/>');
  expect(strFromU8(files["word/diagrams/drawing1.xml"])).toContain('sz="1800" b="0" i="0"');
  expect(strFromU8(files["word/diagrams/drawing1.xml"])).toContain('<a:latin typeface="Arial"/>');
  expect(strFromU8(files["word/diagrams/drawing1.xml"])).toContain('<a:srgbClr val="000000"/>');
  expect(strFromU8(files["word/_rels/document.xml.rels"])).toContain("relationships/diagramData");
  expect(strFromU8(files["word/_rels/document.xml.rels"])).toContain("relationships/diagramDrawing");
});
