// Remaining room at the foot of a page's BODY, in page-relative CSS px.
//
// The break-only paragraph that carries the section's first sectPr is EMPTY, so
// it paints nothing and cannot be found by text. What can be measured is the
// room it was offered: the gap between the last body item on the page it sits
// at the foot of and the body bottom. bodyBottom is derived from the document
// itself -- the deepest body item seen on any page -- rather than assumed from
// the margins, so a footer reserve is included whatever its size.
import { chromium } from "/Users/alexpickett/Desktop/Projects/wordinweb-parity/node_modules/playwright/index.mjs";

const [doc, pagesArg] = process.argv.slice(2);
const want = (pagesArg ?? "1").split(",").map(Number);

const browser = await chromium.launch();
const page = await browser.newPage({ viewport: { width: 1500, height: 1300 } });
await page.goto(`http://localhost:5299/?doc=/fixtures/${doc}.docx&editable=0&comments=0`, { waitUntil: "domcontentloaded" });
await page.waitForSelector(".dxw-page", { timeout: 180000 });
await page.waitForTimeout(15000);

const count = await page.evaluate(() => document.querySelectorAll(".dxw-page").length);

// A body item is anything painted between the top and bottom margins. The
// header and footer bands are excluded by y so a running header cannot be
// mistaken for the last line of the body.
const scan = async (i) => {
  await page.evaluate((idx) => document.querySelectorAll(".dxw-page")[idx]?.scrollIntoView({ block: "center" }), i);
  await page.waitForTimeout(450);
  return page.evaluate((idx) => {
    const pg = document.querySelectorAll(".dxw-page")[idx];
    const pr = pg.getBoundingClientRect();
    const items = [];
    pg.querySelectorAll("span, img, canvas, svg, td, th").forEach((el) => {
      if (el.tagName.toLowerCase() === "span" && el.childElementCount) return;
      const r = el.getBoundingClientRect();
      if (!r.height) return;
      const top = +(r.y - pr.y).toFixed(2);
      const bottom = +(top + r.height).toFixed(2);
      items.push({ top, bottom, text: (el.textContent ?? "").trim().slice(0, 34) });
    });
    const body = items.filter((it) => it.top >= 90 && it.bottom <= 985);
    body.sort((a, b) => a.bottom - b.bottom);
    return {
      page: idx + 1,
      pageH: +pr.height.toFixed(2),
      lastBodyBottom: body.length ? body.at(-1).bottom : null,
      tail: body.slice(-3),
    };
  }, i);
};

// Calibrate the body bottom from the deepest body item anywhere in the document.
let deepest = 0;
const probe = [];
for (let i = 0; i < count; i++) {
  const r = await scan(i);
  probe.push(r);
  if (r.lastBodyBottom && r.lastBodyBottom > deepest) deepest = r.lastBodyBottom;
}

console.log(`# ${doc}: ${count} pages, page height ${probe[0].pageH}`);
console.log(`# deepest body item anywhere = ${deepest.toFixed(2)}  (this is the bodyBottom estimate)`);
for (const p of probe) {
  if (!want.includes(p.page)) continue;
  console.log(`\npage ${p.page}: last body bottom = ${p.lastBodyBottom ?? "(no body items)"}`);
  if (p.lastBodyBottom !== null) {
    console.log(`  ROOM BELOW IT = ${(deepest - p.lastBodyBottom).toFixed(2)} px`);
  }
  for (const t of p.tail) console.log(`    top=${t.top.toFixed(2).padStart(8)} bottom=${t.bottom.toFixed(2).padStart(8)}  ${JSON.stringify(t.text)}`);
}
await browser.close();
