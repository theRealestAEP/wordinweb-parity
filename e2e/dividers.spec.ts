import { expect, Page, test } from "@playwright/test";
import { readFileSync } from "node:fs";
import { strFromU8, unzipSync } from "fflate";

function tool(page: Page, title: string) {
  return page.locator(`[title=${JSON.stringify(title)}], [data-tip=${JSON.stringify(title)}]`).first();
}

async function downloadDocumentXml(page: Page): Promise<string> {
  const pending = page.waitForEvent("download");
  await page.getByText("Download", { exact: true }).click();
  const path = await (await pending).path();
  if (!path) throw new Error("download path unavailable");
  return strFromU8(unzipSync(new Uint8Array(readFileSync(path)))["word/document.xml"]);
}

test("the imported résumé divider can be edited and another divider can be created", async ({ page }) => {
  await page.setViewportSize({ width: 1400, height: 900 });
  await page.goto("/?doc=/fixtures/wild3-resume.docx");
  await page.waitForSelector(".dxw-page span");

  await page.locator(".dxw-page span").filter({ hasText: "Baltimore" }).first().click();
  await page.getByRole("button", { name: "insert", exact: true }).click();
  await tool(page, "Insert or edit divider").click();
  await expect(page.getByLabel("Divider style")).toHaveValue("thinThickSmallGap");
  await expect(page.getByLabel("Divider width in points")).toHaveValue("3");
  await expect(page.getByLabel("Divider gap in points")).toHaveValue("1");

  await page.getByLabel("Divider style").selectOption("dashed");
  await page.getByLabel("Divider color").fill("#2e74b5");
  await page.getByLabel("Divider width in points").fill("2");
  await page.getByLabel("Divider gap in points").fill("2");
  await page.getByRole("button", { name: "Apply divider", exact: true }).click();

  await page.locator(".dxw-page span").filter({ hasText: "Professional" }).first().click();
  await tool(page, "Insert or edit divider").click();
  await page.getByLabel("Divider style").selectOption("double");
  await page.getByLabel("Divider color").fill("#c62828");
  await page.getByLabel("Divider width in points").fill("1");
  await page.getByLabel("Divider gap in points").fill("1");
  await page.getByRole("button", { name: "Apply divider", exact: true }).click();

  const xml = await downloadDocumentXml(page);
  expect(xml).toContain('<w:bottom w:val="dashed" w:sz="16" w:space="2" w:color="2E74B5"/>');
  expect(xml).toContain('<w:bottom w:val="double" w:sz="8" w:space="1" w:color="C62828"/>');
});

test("the résumé divider remains visible when printing with background graphics disabled", async ({ page }) => {
  await page.goto("/?doc=/fixtures/wild3-resume.docx");
  await page.waitForSelector(".dxw-page [data-dxw-edge]");

  await page.evaluate(() => {
    const state = window as Window & {
      __dividerPrintProbe?: { background: string; colorAdjust: string };
    };
    const observer = new MutationObserver(() => {
      const frame = [...document.querySelectorAll("iframe")].find((candidate) =>
        candidate.contentDocument?.querySelector(".dxw-pages"),
      );
      if (!frame?.contentWindow || !frame.contentDocument) return;
      observer.disconnect();
      frame.contentWindow.print = () => {
        const edge = frame.contentDocument?.querySelector<HTMLElement>("[data-dxw-edge]");
        const style = edge ? frame.contentWindow?.getComputedStyle(edge) : null;
        state.__dividerPrintProbe = {
          background: style?.backgroundImage ?? "",
          colorAdjust: style?.getPropertyValue("print-color-adjust") ?? "",
        };
      };
    });
    observer.observe(document.body, { childList: true });
  });

  await page.getByRole("button", { name: "Print", exact: true }).click();
  await expect.poll(() => page.evaluate(() =>
    (window as Window & { __dividerPrintProbe?: unknown }).__dividerPrintProbe,
  )).toEqual({
    background: expect.stringContaining("linear-gradient"),
    colorAdjust: "exact",
  });
});
