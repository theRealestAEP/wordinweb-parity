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

async function submitChart(page: Page, title: string, type: "column" | "bar" | "line" | "pie", series: string): Promise<void> {
  const edit = page.getByRole("button", { name: "Edit data", exact: true });
  await (await edit.isVisible() ? edit : page.getByRole("button", { name: "Chart", exact: true })).click();
  await page.getByLabel("Chart type").selectOption(type);
  await page.getByLabel("Chart title").fill(title);
  const entries = series.split("\n").map((line) => {
    const [name, values] = line.split(":");
    return { name: name.trim(), values: values.split(",").map((value) => value.trim()) };
  });
  const categoryCount = Math.max(...entries.map((entry) => entry.values.length));
  let currentCategories = await page.getByLabel(/^Chart category \d+$/).count();
  while (currentCategories < categoryCount) {
    await page.getByRole("button", { name: "Add category", exact: true }).click();
    currentCategories++;
  }
  for (let index = 0; index < categoryCount; index++) {
    await page.getByRole("textbox", { name: `Chart category ${index + 1}`, exact: true }).fill(`Q${index + 1}`);
  }
  let currentSeries = await page.getByLabel(/^Chart series \d+ name$/).count();
  while (currentSeries < entries.length) {
    await page.getByRole("button", { name: "Add series", exact: true }).click();
    currentSeries++;
  }
  for (let seriesIndex = 0; seriesIndex < entries.length; seriesIndex++) {
    await page.getByLabel(`Chart series ${seriesIndex + 1} name`).fill(entries[seriesIndex].name);
    for (let valueIndex = 0; valueIndex < entries[seriesIndex].values.length; valueIndex++) {
      await page.getByLabel(`Chart series ${seriesIndex + 1} value ${valueIndex + 1}`).fill(entries[seriesIndex].values[valueIndex]);
    }
  }
  await page.getByRole("button", { name: "Insert or update chart" }).click();
}

test("advanced Insert creates and edits a native chart with undo, redo, and save", async ({ page }) => {
  await openInsert(page);
  await page.getByRole("button", { name: "Chart", exact: true }).click();
  await expect(page.getByLabel("Chart title")).toHaveValue("");
  await expect(page.getByRole("textbox", { name: "Chart category 1", exact: true })).toHaveValue("");
  await expect(page.getByLabel("Chart series 1 name")).toHaveValue("");
  await page.getByRole("button", { name: "Chart", exact: true }).click();
  await submitChart(page, "Quarterly sales", "column", "Revenue: 12, 19, 15, 24\nCosts: 8, 11, 10, 14");

  const chart = page.locator("[data-dxw-chart]").last();
  await expect(chart).toBeVisible();
  await expect(chart.getByText("Quarterly sales", { exact: true })).toBeVisible();
  await expect(chart.locator("rect")).not.toHaveCount(1);

  const box = await chart.boundingBox();
  expect(box).not.toBeNull();
  await page.mouse.click(box!.x + box!.width / 2, box!.y + box!.height / 2);
  await expect(page.locator("[data-dxw-img-handle]")).toHaveCount(8);
  await expect(page.getByRole("button", { name: "Chart Format", exact: true })).toBeVisible();
  await expect(page.locator("[data-dxw-object-format]").getByRole("button", { name: "Fill", exact: true })).toHaveCount(0);
  await expect(page.locator("[data-dxw-object-format]").getByRole("button", { name: "Edit text", exact: true })).toHaveCount(0);
  await page.getByRole("button", { name: "Edit data", exact: true }).click();
  await expect(page.getByLabel("Chart title")).toHaveValue("Quarterly sales");
  await expect(page.getByRole("textbox", { name: "Chart category 4", exact: true })).toHaveValue("Q4");
  await expect(page.getByLabel("Chart series 2 name")).toHaveValue("Costs");
  await expect(page.getByLabel("Chart series 2 value 4")).toHaveValue("14");
  await page.getByRole("button", { name: "Edit data", exact: true }).click();

  await submitChart(page, "Updated trend", "line", "Revenue: 4, 8, 6, 11\nCosts: 2, 3, 4, 5");
  await expect(page.locator("[data-dxw-chart]", { hasText: "Updated trend" })).toBeVisible();
  await expect(page.locator("[data-dxw-chart] polyline")).toHaveCount(2);

  await page.keyboard.press(`${MOD}+z`);
  await expect(page.locator("[data-dxw-chart]", { hasText: "Quarterly sales" })).toBeVisible();
  await page.keyboard.press(`${MOD}+Shift+z`);
  await expect(page.locator("[data-dxw-chart]", { hasText: "Updated trend" })).toBeVisible();

  const pending = page.waitForEvent("download");
  await page.getByText("Download", { exact: true }).click();
  const path = await (await pending).path();
  expect(path).not.toBeNull();
  const files = unzipSync(new Uint8Array(readFileSync(path!)));
  const chartXml = strFromU8(files["word/charts/chart1.xml"]);
  expect(strFromU8(files["word/document.xml"])).toContain("<c:chart");
  expect(chartXml).toContain("<c:lineChart>");
  expect(chartXml).toContain("Updated trend");
  expect(strFromU8(files["word/charts/_rels/chart1.xml.rels"])).toContain("relationships/package");
  const workbook = unzipSync(files["word/embeddings/Microsoft_Excel_Worksheet1.xlsx"]);
  const sheet = strFromU8(workbook["xl/worksheets/sheet1.xml"]);
  expect(sheet).toContain("Revenue");
  expect(sheet).toContain("<v>11</v>");
});
