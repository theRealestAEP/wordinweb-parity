// Does the gate's viewing-mode upload actually replace the rendered document?
// Reproduces renderWebPages exactly: load the FIXTURE by URL in viewing mode,
// then setInputFiles the EDITED docx, then look for the inserted TOC.
import { chromium } from "/Users/alexpickett/Desktop/Projects/wordinweb-parity/node_modules/playwright/index.mjs";

const base = "http://localhost:5299";
const fixture = "wild2-legal-ca-agreement";
const edited = "/Users/alexpickett/Desktop/Projects/wordinweb-parity/apps/demo/public/fixtures/probe-tocinsert-edited.docx";

const browser = await chromium.launch();
const page = await browser.newPage({ viewport: { width: 1700, height: 1200 }, deviceScaleFactor: 2 });
await page.goto(`${base}/?doc=/fixtures/${fixture}.docx&editable=0&comments=0`, { waitUntil: "domcontentloaded" });
await page.waitForSelector(".dxw-page span", { state: "attached", timeout: 120_000 });
await page.waitForTimeout(8000);

const snapshot = () => page.evaluate(() => {
  const pages = document.querySelectorAll(".dxw-page");
  const first = pages[0];
  const text = first ? (first.textContent ?? "") : "";
  return {
    pages: pages.length,
    // A dot-leader run is the TOC's signature; the authored title page has none.
    hasLeaders: /\.{12,}/.test(text),
    page1Chars: text.replace(/\s+/g, " ").trim().length,
  };
});

console.log("BEFORE upload:", JSON.stringify(await snapshot()));

const input = page.locator("#docx-upload");
console.log("#docx-upload count:", await input.count());
const box = await page.evaluate(() => {
  const el = document.querySelector("#docx-upload");
  if (!el) return null;
  const cs = getComputedStyle(el);
  return { disabled: el.disabled, display: cs.display, visibility: cs.visibility, type: el.type };
});
console.log("#docx-upload state:", JSON.stringify(box));

const loading = page.locator("[data-dxw-loading]");
await input.setInputFiles(edited);
await loading.waitFor({ state: "attached", timeout: 15_000 }).then(
  () => console.log("loading indicator ATTACHED (a real re-parse started)"),
  () => console.log("loading indicator NEVER ATTACHED (no re-parse observed)"),
);
await loading.waitFor({ state: "detached", timeout: 120_000 });
await page.waitForSelector(".dxw-page span", { state: "attached", timeout: 120_000 });
await page.waitForTimeout(8000);

console.log("AFTER  upload:", JSON.stringify(await snapshot()));
await browser.close();
