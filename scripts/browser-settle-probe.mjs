// Page count as a function of settle time, then the room once it is stable.
//
// The demo lays out once with an approximate measurer and re-lays when the real
// font metrics arrive, so a page count read too early is a DIFFERENT number
// from the settled one. Poll until the count repeats, and print the timeline so
// a disagreement about "the browser says N" can be read rather than argued.
import { chromium } from "/Users/alexpickett/Desktop/Projects/wordinweb-parity/node_modules/playwright/index.mjs";

const doc = process.argv[2];
const browser = await chromium.launch();
const page = await browser.newPage({ viewport: { width: 1500, height: 1300 } });
const t0 = Date.now();
await page.goto(`http://localhost:5299/?doc=/fixtures/${doc}.docx&editable=0&comments=0`, { waitUntil: "domcontentloaded" });
await page.waitForSelector(".dxw-page", { timeout: 180000 });

const counts = [];
let stable = 0;
let last = -1;
for (let i = 0; i < 45; i++) {
  const n = await page.evaluate(() => document.querySelectorAll(".dxw-page").length);
  counts.push([((Date.now() - t0) / 1000).toFixed(1), n]);
  stable = n === last ? stable + 1 : 0;
  last = n;
  if (stable >= 8) break; // eight identical reads, 2s apart
  await page.waitForTimeout(2000);
}
console.log(`# ${doc}`);
console.log("timeline (s, pages):", counts.map(([t, n]) => `${t}:${n}`).join(" "));
console.log(`SETTLED PAGE COUNT = ${last}`);

// Now the body extent of the first three pages, at the settled layout.
for (const idx of [0, 1, 2]) {
  await page.evaluate((i) => document.querySelectorAll(".dxw-page")[i]?.scrollIntoView({ block: "center" }), idx);
  await page.waitForTimeout(700);
  const r = await page.evaluate((i) => {
    const pg = document.querySelectorAll(".dxw-page")[i];
    if (!pg) return null;
    const pr = pg.getBoundingClientRect();
    const body = [];
    pg.querySelectorAll("span, img, canvas, svg").forEach((el) => {
      if (el.tagName.toLowerCase() === "span" && el.childElementCount) return;
      const rr = el.getBoundingClientRect();
      if (!rr.height) return;
      const top = +(rr.y - pr.y).toFixed(2);
      const bottom = +(top + rr.height).toFixed(2);
      if (top >= 90 && bottom <= 985) body.push({ top, bottom, t: (el.textContent ?? "").trim() });
    });
    body.sort((a, b) => a.bottom - b.bottom);
    return {
      pageH: +pr.height.toFixed(2),
      bodyItems: body.length,
      chars: body.reduce((a, b) => a + b.t.length, 0),
      lastBottom: body.length ? body.at(-1).bottom : null,
      firstText: body.find((b) => b.t)?.t.slice(0, 40) ?? "",
    };
  }, idx);
  if (r) {
    console.log(`page ${idx + 1}: bodyItems=${r.bodyItems} bodyChars=${r.chars} lastBodyBottom=${r.lastBottom} first=${JSON.stringify(r.firstText)}`);
  }
}
await browser.close();
