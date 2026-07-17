import { expect, test } from "@playwright/test";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { unzipSync } from "fflate";

const MOD = process.platform === "darwin" ? "Meta" : "Control";
const FIXTURE = join(process.cwd(), "apps/demo/public/fixtures/benchmark.docx");

test("advanced Insert captures a screen frame as an editable PNG picture", async ({ page }) => {
  await page.addInitScript(() => {
    let stopped = false;
    Object.defineProperty(navigator, "mediaDevices", {
      configurable: true,
      value: {
        getDisplayMedia: async () => {
          const canvas = document.createElement("canvas");
          canvas.width = 160;
          canvas.height = 90;
          const context = canvas.getContext("2d")!;
          context.fillStyle = "#2e74b5";
          context.fillRect(0, 0, canvas.width, canvas.height);
          const stream = canvas.captureStream(1);
          for (const track of stream.getTracks()) {
            const stop = track.stop.bind(track);
            track.stop = () => {
              stopped = true;
              stop();
            };
          }
          return stream;
        },
      },
    });
    Object.defineProperty(window, "__captureStopped", { get: () => stopped });
  });

  await page.goto("/?doc=/fixtures/benchmark.docx");
  await page.waitForSelector(".dxw-page span");
  await page.getByRole("button", { name: "insert", exact: true }).click();
  const target = page.locator(".dxw-page span", { hasText: "classic." }).first();
  await target.click();
  await page.keyboard.press("End");
  await page.getByTitle("Capture and insert a screen, window, or tab").click();

  const screenshot = page.locator('.dxw-page img[style*="width: 160px"][style*="height: 90px"]');
  await expect(screenshot).toHaveCount(1);
  await expect(page.locator("[data-dxw-screenshot-status]")).toHaveText("Screenshot inserted.");
  await expect.poll(() => page.evaluate(() => (window as unknown as { __captureStopped: boolean }).__captureStopped)).toBe(true);
  await screenshot.click();
  await expect(page.locator("[data-dxw-img-handle]")).toHaveCount(8);
  await page.keyboard.press(`${MOD}+z`);
  await expect(screenshot).toHaveCount(0);
  await page.keyboard.press(`${MOD}+Shift+z`);
  await expect(screenshot).toHaveCount(1);

  const before = Object.keys(unzipSync(new Uint8Array(readFileSync(FIXTURE)))).filter((name) => /^word\/media\/.*\.png$/i.test(name)).length;
  const pending = page.waitForEvent("download");
  await page.getByText("Download", { exact: true }).click();
  const path = await (await pending).path();
  expect(path).not.toBeNull();
  const after = Object.keys(unzipSync(new Uint8Array(readFileSync(path!)))).filter((name) => /^word\/media\/.*\.png$/i.test(name)).length;
  expect(after).toBe(before + 1);
});

test("Screenshot reports unsupported capture instead of silently doing nothing", async ({ page }) => {
  await page.addInitScript(() => {
    Object.defineProperty(navigator, "mediaDevices", { configurable: true, value: {} });
  });
  await page.goto("/?doc=/fixtures/benchmark.docx");
  await page.waitForSelector(".dxw-page span");
  await page.getByRole("button", { name: "insert", exact: true }).click();
  await page.locator(".dxw-page span").first().click();
  const before = await page.locator("[data-dxw-image-format]").count();
  await page.getByTitle("Capture and insert a screen, window, or tab").click();
  await expect(page.locator("[data-dxw-screenshot-status]")).toHaveText("Screen capture is not supported in this browser.");
  await expect(page.locator("[data-dxw-screenshot-status]")).toHaveAttribute("role", "alert");
  await expect(page.locator("[data-dxw-image-format]")).toHaveCount(before);
});

test("Screenshot asks for a document caret when an object is selected", async ({ page }) => {
  await page.addInitScript(() => {
    Object.defineProperty(navigator, "mediaDevices", {
      configurable: true,
      value: { getDisplayMedia: async () => { throw new Error("should not request capture without a caret"); } },
    });
  });
  await page.goto("/?doc=/fixtures/benchmark.docx");
  await page.waitForSelector(".dxw-page span");
  await page.getByRole("button", { name: "insert", exact: true }).click();
  await page.locator("[data-dxw-image-format]").first().click();
  await page.getByTitle("Capture and insert a screen, window, or tab").click();
  await expect(page.locator("[data-dxw-screenshot-status]")).toHaveText("Click in the document before inserting a screenshot.");
});

test("Screenshot reports a cancelled picker instead of silently doing nothing", async ({ page }) => {
  await page.addInitScript(() => {
    Object.defineProperty(navigator, "mediaDevices", {
      configurable: true,
      value: { getDisplayMedia: async () => { throw new DOMException("Cancelled", "NotAllowedError"); } },
    });
  });
  await page.goto("/?doc=/fixtures/benchmark.docx");
  await page.waitForSelector(".dxw-page span");
  await page.getByRole("button", { name: "insert", exact: true }).click();
  await page.locator(".dxw-page span").first().click();
  const before = await page.locator("[data-dxw-image-format]").count();
  await page.getByTitle("Capture and insert a screen, window, or tab").click();
  await expect(page.locator("[data-dxw-screenshot-status]")).toHaveText("Screen capture was cancelled or denied.");
  await expect(page.locator("[data-dxw-image-format]")).toHaveCount(before);
});

test("Screenshot stops capture tracks when frame readiness fails", async ({ page }) => {
  await page.addInitScript(() => {
    let stopped = false;
    Object.defineProperty(navigator, "mediaDevices", {
      configurable: true,
      value: {
        getDisplayMedia: async () => {
          const canvas = document.createElement("canvas");
          canvas.width = 160;
          canvas.height = 90;
          const stream = canvas.captureStream(1);
          for (const track of stream.getTracks()) {
            const stop = track.stop.bind(track);
            track.stop = () => { stopped = true; stop(); };
          }
          return stream;
        },
      },
    });
    HTMLMediaElement.prototype.play = async () => { throw new DOMException("Playback failed", "NotSupportedError"); };
    Object.defineProperty(window, "__captureStopped", { get: () => stopped });
  });
  await page.goto("/?doc=/fixtures/benchmark.docx");
  await page.waitForSelector(".dxw-page span");
  await page.getByRole("button", { name: "insert", exact: true }).click();
  await page.locator(".dxw-page span").first().click();
  await page.getByTitle("Capture and insert a screen, window, or tab").click();
  await expect(page.locator("[data-dxw-screenshot-status]")).toHaveText("Screenshot failed. Please try again.");
  await expect.poll(() => page.evaluate(() => (window as unknown as { __captureStopped: boolean }).__captureStopped)).toBe(true);
});
