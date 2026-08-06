// Image boxes and text-line tops for one page, in page-relative CSS px.
// Image boxes are directly comparable with the placed boxes pdf-page.py reads
// from the Word content stream; line TOPS are not comparable with Word's
// baselines and are here only to say where the text between the images sits.
import { chromium } from "/Users/alexpickett/Desktop/Projects/wordinweb-parity/node_modules/playwright/index.mjs";

const [doc, pn] = process.argv.slice(2);
const browser = await chromium.launch();
const page = await browser.newPage({ viewport: { width: 1500, height: 1300 } });
await page.goto(`http://localhost:5299/?doc=/fixtures/${doc}.docx&editable=0&comments=0`, { waitUntil: "domcontentloaded" });
await page.waitForSelector(".dxw-page", { timeout: 180000 });
await page.waitForTimeout(15000);

const count = await page.evaluate(() => document.querySelectorAll(".dxw-page").length);
await page.evaluate((i) => document.querySelectorAll(".dxw-page")[i]?.scrollIntoView({ block: "center" }), Number(pn) - 1);
await page.waitForTimeout(900);

const info = await page.evaluate((i) => {
  const pg = document.querySelectorAll(".dxw-page")[i];
  const pr = pg.getBoundingClientRect();
  const imgs = [];
  pg.querySelectorAll("img, canvas, svg").forEach((el) => {
    const r = el.getBoundingClientRect();
    if (!r.height) return;
    imgs.push({
      tag: el.tagName.toLowerCase(),
      top: +(r.y - pr.y).toFixed(2),
      h: +r.height.toFixed(2),
      x: +(r.x - pr.x).toFixed(2),
      w: +r.width.toFixed(2),
    });
  });
  const m = new Map();
  pg.querySelectorAll("span").forEach((s) => {
    if (s.childElementCount) return;
    const r = s.getBoundingClientRect();
    if (!r.height) return;
    const top = +(r.y - pr.y).toFixed(2);
    const cur = m.get(top) ?? { top, text: "", h: +r.height.toFixed(2) };
    cur.text += s.textContent ?? "";
    m.set(top, cur);
  });
  const lines = [...m.values()].filter((l) => l.text.trim()).sort((a, b) => a.top - b.top);
  return { pageH: +pr.height.toFixed(2), imgs: imgs.sort((a, b) => a.top - b.top), lines };
}, Number(pn) - 1);

console.log(`# ours ${doc} page ${pn} of ${count}, page height ${info.pageH}`);
for (const g of info.imgs) {
  console.log(`  IMG  top=${g.top.toFixed(2).padStart(8)} h=${g.h.toFixed(2).padStart(7)} bottom=${(g.top + g.h).toFixed(2).padStart(8)} x=${g.x.toFixed(2).padStart(7)} w=${g.w.toFixed(2).padStart(7)}  <${g.tag}>`);
}
for (const l of info.lines) {
  console.log(`  TXT  top=${l.top.toFixed(2).padStart(8)} h=${l.h.toFixed(2).padStart(7)}                            ${JSON.stringify(l.text.trim().slice(0, 46))}`);
}
await browser.close();
