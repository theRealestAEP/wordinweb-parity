import { expect, test, Page } from "@playwright/test";
import { Editor } from "./editing.js";

/**
 * Per-keystroke typing & caret invariants (task #16). These lock in the
 * editing-stability fixes: the caret is visible and sane after EVERY
 * keystroke (incl. space-after-period, which used to fly to the page edge),
 * lines above the caret's paragraph never move while typing, and words don't
 * oscillate between lines beyond a small bound. Run against generated
 * fixtures so they hold on any machine.
 */

/** Snapshot visible word spans (text + page position) around the caret. */
async function spanSnapshot(page: Page): Promise<{ cy: number; spans: [string, number, number][] }> {
  return page.evaluate(() => {
    const pg = document.querySelector(".dxw-page")!;
    const pr = pg.getBoundingClientRect();
    const caret = [...document.querySelectorAll("div")].find(
      (d) => (d as HTMLElement).style.width === "1.5px" && (d as HTMLElement).style.position === "absolute",
    ) as HTMLElement | undefined;
    const cr = caret?.getBoundingClientRect();
    const spans: [string, number, number][] = [];
    for (const el of pg.querySelectorAll("span")) {
      const t = (el.textContent || "").trim();
      if (!t || el.children.length) continue;
      const r = el.getBoundingClientRect();
      spans.push([t, Math.round((r.x - pr.x) * 2) / 2, Math.round((r.y - pr.y) * 2) / 2]);
    }
    return { cy: cr ? cr.y - pr.y : -1, spans };
  });
}

async function caretInfo(page: Page): Promise<{ x: number; y: number; visible: boolean } | null> {
  return page.evaluate(() => {
    const caret = [...document.querySelectorAll("div")].find(
      (d) => (d as HTMLElement).style.width === "1.5px" && (d as HTMLElement).style.position === "absolute",
    ) as HTMLElement | undefined;
    if (!caret || caret.style.display === "none") return null;
    const pr = document.querySelector(".dxw-page")!.getBoundingClientRect();
    const r = caret.getBoundingClientRect();
    return { x: r.x - pr.x, y: r.y - pr.y, visible: r.height > 2 };
  });
}

test.describe("typing & caret invariants", () => {
  test("caret stays visible and in the text area through a typing burst with punctuation", async ({ page }) => {
    const ed = await Editor.open(page, "sample");
    await ed.clickText("dolore");
    await ed.expectCaretVisible();
    for (const ch of "end. Next word. More text follows here") {
      await page.keyboard.type(ch);
      await page.waitForTimeout(30);
      const c = await caretInfo(page);
      expect(c, `caret missing after '${ch}'`).not.toBeNull();
      expect(c!.visible, `caret collapsed after '${ch}'`).toBe(true);
      // Body text never starts left of the margin; a caret at the page edge
      // is the zero-width-space-span bug.
      expect(c!.x, `caret flew to page edge after '${ch}'`).toBeGreaterThan(50);
    }
  });

  test("lines above the caret's paragraph never move while typing", async ({ page }) => {
    const ed = await Editor.open(page, "sample");
    await ed.clickText("dolore");
    const before = await spanSnapshot(page);
    await page.keyboard.type("wwwww wwww www wwwww ");
    await page.waitForTimeout(200);
    const after = await spanSnapshot(page);
    // Every span that sat at least 40px above the caret before typing must
    // still be at exactly the same spot.
    const key = (s: [string, number, number]) => `${s[0]}@${s[1]},${s[2]}`;
    const afterSet = new Set(after.spans.map(key));
    const missing = before.spans.filter((s) => s[2] < before.cy - 40 && !afterSet.has(key(s)));
    expect(missing, `spans above the caret moved: ${missing.slice(0, 3).map(key).join(" | ")}`).toHaveLength(0);
  });

  test("words do not oscillate between lines during a burst", async ({ page }) => {
    const ed = await Editor.open(page, "sample");
    await ed.clickText("dolore");
    const hist = new Map<string, number[]>();
    const record = async () => {
      const snap = await spanSnapshot(page);
      const seen = new Map<string, number[]>();
      for (const [t, , y] of snap.spans) {
        if (!seen.has(t)) seen.set(t, []);
        seen.get(t)!.push(y);
      }
      for (const [t, ys] of seen) {
        if (ys.length !== 1) continue; // only uniquely-identifiable words
        const h = hist.get(t) ?? [];
        if (h.length === 0 || h[h.length - 1] !== ys[0]) h.push(ys[0]);
        hist.set(t, h);
      }
    };
    await record();
    for (const ch of "adding several words to push line boundaries around here") {
      await page.keyboard.type(ch);
      await page.waitForTimeout(25);
      await record();
    }
    let oscillators = 0;
    for (const ys of hist.values()) {
      for (let i = 2; i < ys.length; i++) {
        if (ys.slice(0, i - 1).includes(ys[i])) {
          oscillators++;
          break;
        }
      }
    }
    // A word returning to a line it already left reads as flicker; a couple
    // of legitimate reflow events are tolerated, flapping is not.
    expect(oscillators).toBeLessThanOrEqual(3);
  });

  test("typed text lands at the caret and stays contiguous", async ({ page }) => {
    const ed = await Editor.open(page, "sample");
    await ed.clickText("dolore");
    await page.keyboard.type("MARKER123");
    await page.waitForTimeout(150);
    await ed.expectHasText("MARKER123");
  });

  test("spaces at a wrapped boundary hang on the upper line, keep the caret visible, and undo exactly", async ({ page }) => {
    const ed = await Editor.open(page, "sample");
    const labore = ed.span("labore");
    const box = (await labore.boundingBox())!;
    await page.mouse.click(box.x + box.width - 0.5, box.y + box.height / 2);
    await page.keyboard.press("ArrowRight");

    const state = () => page.evaluate(() => {
      const spans = [...document.querySelectorAll<HTMLElement>(".dxw-page span")];
      const labore = spans.find((span) => span.textContent === "labore")!;
      const et = spans.find((span) => span.textContent === "et")!;
      const caret = [...document.querySelectorAll<HTMLElement>("div")].find(
        (div) => div.style.width === "1.5px" && div.style.position === "absolute",
      );
      // Word hangs wrap-boundary spaces past the end of the UPPER line: count
      // space spans on labore's line to its right.
      const hanging = [...labore.parentElement!.querySelectorAll<HTMLElement>("span")].filter(
        (span) =>
          span.textContent === " " &&
          span.style.top === labore.style.top &&
          parseFloat(span.style.left) >= parseFloat(labore.style.left),
      );
      return {
        laboreTop: parseFloat(labore.style.top),
        etX: parseFloat(et.style.left),
        etTop: parseFloat(et.style.top),
        caretX: caret?.isConnected && caret.style.display === "block"
          ? parseFloat(caret.style.left)
          : null,
        caretTop: caret?.isConnected && caret.style.display === "block"
          ? parseFloat(caret.style.top)
          : null,
        hanging: hanging.map((space) => parseFloat(space.style.left)).sort((a, b) => a - b),
      };
    });

    const before = await state();
    expect(before.caretX).not.toBeNull();
    // The wrap separator space itself hangs on the upper line.
    expect(before.hanging).toHaveLength(1);

    await page.keyboard.press("Space");
    await page.waitForTimeout(100);
    const once = await state();
    // The typed space joins the hanging run; "et" does not move.
    expect(once.hanging).toHaveLength(2);
    expect(once.etX).toBeCloseTo(before.etX, 1);
    // Caret stays at the end of the UPPER line and never retreats. Its exact X
    // is min(logical position, caretClampX): hanging spaces may extend PAST the
    // line's available-width edge, and the caret deliberately pins at that edge
    // rather than following them — so under wide space metrics the caret can
    // sit left of the second space's start. Assert the invariant (on the upper
    // line, monotone, past the first hanging space), not a font-specific X.
    expect(once.caretTop).toBeCloseTo(before.laboreTop, 0);
    expect(once.caretX!).toBeGreaterThanOrEqual(before.caretX!);
    // >=, not >: when the clamp edge coincides with the first hanging space's
    // left (font-metric dependent), the pinned caret sits exactly there.
    expect(once.caretX!).toBeGreaterThanOrEqual(once.hanging[0]);

    await page.keyboard.press("ArrowLeft");
    const inSpace = await state();
    // Between the two hanging spaces: still on the upper line. Visual X never
    // moves RIGHT on ArrowLeft, but it may not move at all when both logical
    // positions pin at the clamp edge.
    expect(inSpace.caretTop).toBeCloseTo(before.laboreTop, 0);
    expect(inSpace.caretX!).toBeLessThanOrEqual(once.caretX!);

    await page.keyboard.press("Space");
    await page.waitForTimeout(100);
    const twice = await state();
    expect(twice.hanging).toHaveLength(3);
    expect(twice.etX).toBeCloseTo(before.etX, 1);
    expect(twice.caretX).not.toBeNull();
    expect(twice.caretTop).toBeCloseTo(before.laboreTop, 0);

    const mod = process.platform === "darwin" ? "Meta" : "Control";
    await page.keyboard.press(`${mod}+z`);
    await page.waitForTimeout(150);
    const undone = await state();
    expect(undone.hanging).toHaveLength(1);
    expect(undone.etX).toBeCloseTo(before.etX, 3);
  });

  test("spaces typed mid-line in a justified paragraph push text right and wrap, never creep backwards", async ({ page }) => {
    // User report round 2: in the justified paragraph, holding space mid-line
    // compressed every other space on the line (the pack rule counted the
    // typed run as budget), so the gap "grew backwards" and the line never
    // re-wrapped. Invariants: the line START never moves, the caret advances
    // monotonically, and within 8 spaces the line's tail word wraps down.
    const ed = await Editor.open(page, "sample");
    const state = () => page.evaluate(() => {
      const pg = document.querySelector(".dxw-page")!;
      const spans = [...pg.querySelectorAll<HTMLElement>("span")];
      const et = spans.find((s) => s.textContent === "et" && Math.abs(parseFloat(s.style.top) - 197) < 10);
      const line = spans.filter(
        (s) => et && Math.abs(parseFloat(s.style.top) - parseFloat(et.style.top)) < 8 && s.textContent!.trim(),
      );
      const caret = [...document.querySelectorAll<HTMLElement>("div")].find(
        (d) => d.style.width === "1.5px" && d.style.position === "absolute" && d.style.display !== "none",
      );
      return {
        lineStartX: et ? parseFloat(et.style.left) : NaN,
        lineWords: line.map((s) => s.textContent),
        caretX: caret ? parseFloat(caret.style.left) : NaN,
      };
    });
    const veniam = ed.span("veniam,");
    const box = (await veniam.boundingBox())!;
    await page.mouse.click(box.x + box.width - 8, box.y + box.height / 2);
    await page.waitForTimeout(150);
    const before = await state();
    expect(before.lineStartX).toBeCloseTo(96, 0);
    let prevCaret = before.caretX;
    for (let i = 1; i <= 8; i++) {
      await page.keyboard.press("Space");
      await page.waitForTimeout(80);
      const s = await state();
      expect(s.caretX, `caret went backwards after space #${i}`).toBeGreaterThan(prevCaret - 0.25);
      expect(s.lineStartX, `line start moved after space #${i}`).toBeCloseTo(96, 0);
      prevCaret = s.caretX;
    }
    const after = await state();
    // The typed run displaced the line's tail word to the next line. (Word
    // counts shift by ±1 as "veniam," splits into two spans post-edit, so
    // compare the trailing word, not the count.)
    expect(after.lineWords[after.lineWords.length - 1]).not.toBe(before.lineWords[before.lineWords.length - 1]);
    // And the caret advanced a sane total (~8 natural spaces, not compressed
    // crumbs — pre-fix total advance was ~12px for 8 spaces).
    expect(after.caretX - before.caretX).toBeGreaterThan(20);
  });

  test("spaces typed at a line end advance the caret forward, never backwards", async ({ page }) => {
    // User report: "typing causes backwards space" — a space typed at the end
    // of a wrapped line landed at the START of the next line and the caret
    // leapt backwards with it.
    const ed = await Editor.open(page, "sample");
    await ed.clickText("document");
    await page.keyboard.press("End");
    await page.waitForTimeout(100);
    let prev = await caretInfo(page);
    expect(prev).not.toBeNull();
    for (let i = 0; i < 3; i++) {
      await page.keyboard.press("Space");
      await page.waitForTimeout(80);
      const c = await caretInfo(page);
      expect(c, `caret missing after space #${i + 1}`).not.toBeNull();
      expect(c!.y, `caret changed lines after space #${i + 1}`).toBeCloseTo(prev!.y, 0);
      expect(c!.x, `caret moved backwards after space #${i + 1}`).toBeGreaterThan(prev!.x);
      prev = c;
    }
  });

  test("a space typed after a whitespace click keeps a visible caret", async ({ page }) => {
    // User report: "the caret breaks in white space" — clicking the blank
    // area right of a short line, then hitting Space, hid the caret entirely
    // (the trailing space had no layout span to bind to).
    const ed = await Editor.open(page, "sample");
    const short = ed.span("superscript");
    const box = (await short.boundingBox())!;
    await page.mouse.click(box.x + box.width + 150, box.y + box.height / 2);
    await page.waitForTimeout(120);
    const before = await caretInfo(page);
    expect(before).not.toBeNull();
    await page.keyboard.press("Space");
    await page.waitForTimeout(120);
    const after = await caretInfo(page);
    expect(after, "caret vanished after typing a space").not.toBeNull();
    expect(after!.visible).toBe(true);
    expect(after!.x).toBeGreaterThan(before!.x);
    await page.keyboard.type("QQ");
    await page.waitForTimeout(120);
    await ed.expectHasText("QQ");
  });

  test("caret can enter an empty paragraph and typing works there", async ({ page }) => {
    const ed = await Editor.open(page, "parity-text");
    await ed.clickText("Plain");
    await page.keyboard.press("End");
    await page.keyboard.press("Enter");
    await page.keyboard.press("Enter");
    await page.keyboard.press("ArrowUp"); // caret onto the empty line
    await page.waitForTimeout(150);
    const c = await caretInfo(page);
    expect(c, "no caret on the empty paragraph").not.toBeNull();
    await page.keyboard.type("EMPTYLINE");
    await page.waitForTimeout(150);
    await ed.expectHasText("EMPTYLINE");
  });

  test("bottom-to-top table text keeps a transformed caret and remains editable", async ({ page }) => {
    const ed = await Editor.open(page, "parity2-nestedtables");
    const label = ed.span("VERTICAL");
    const labelBox = (await label.boundingBox())!;
    await page.mouse.click(
      labelBox.x + labelBox.width / 2,
      labelBox.y + labelBox.height / 2,
    );
    await page.waitForTimeout(150);

    const geometry = () => page.evaluate(() => {
      const label = [...document.querySelectorAll<HTMLElement>(".dxw-page span")].find(
        (span) => span.textContent?.includes("VERT") && getComputedStyle(span).transform !== "none",
      );
      const caret = [...document.querySelectorAll<HTMLElement>("div")].find(
        (div) =>
          div.style.position === "absolute" &&
          div.style.pointerEvents === "none" &&
          div.style.display === "block" &&
          (div.style.width === "1.5px" || div.style.height === "1.5px"),
      );
      const labelRect = label?.getBoundingClientRect();
      const caretRect = caret?.getBoundingClientRect();
      return {
        text: label?.textContent,
        label: labelRect && {
          left: labelRect.left,
          right: labelRect.right,
          top: labelRect.top,
          bottom: labelRect.bottom,
        },
        caret: caretRect && {
          left: caretRect.left,
          right: caretRect.right,
          top: caretRect.top,
          bottom: caretRect.bottom,
          width: caretRect.width,
          height: caretRect.height,
        },
      };
    });
    const before = await geometry();
    expect(before.caret).toBeTruthy();
    expect(before.caret!.width).toBeGreaterThan(before.caret!.height * 4);
    expect(before.caret!.right).toBeGreaterThan(before.label!.left);
    expect(before.caret!.left).toBeLessThan(before.label!.right);
    expect(before.caret!.top).toBeGreaterThanOrEqual(before.label!.top - 1);
    expect(before.caret!.bottom).toBeLessThanOrEqual(before.label!.bottom + 1);

    await page.keyboard.type("X");
    await page.waitForTimeout(200);
    const after = await geometry();
    expect(after.text?.replace("X", "")).toBe("VERTICAL");
    expect(after.caret!.width).toBeGreaterThan(after.caret!.height * 4);
    expect(after.caret!.top).toBeGreaterThanOrEqual(after.label!.top - 1);
    expect(after.caret!.bottom).toBeLessThanOrEqual(after.label!.bottom + 1);
  });

  test("undo after a burst restores the exact layout", async ({ page }) => {
    const ed = await Editor.open(page, "sample");
    await ed.clickText("dolore");
    const before = await spanSnapshot(page);
    await page.keyboard.type("scramble the layout with plenty of words ");
    await page.waitForTimeout(150);
    const mod = process.platform === "darwin" ? "Meta" : "Control";
    await page.keyboard.press(`${mod}+z`);
    await page.waitForTimeout(300);
    const after = await spanSnapshot(page);
    const key = (s: [string, number, number]) => `${s[0]}@${s[1]},${s[2]}`;
    const beforeSet = new Set(before.spans.map(key));
    const stray = after.spans.filter((s) => !beforeSet.has(key(s)));
    expect(stray, `layout differs after undo: ${stray.slice(0, 3).map(key).join(" | ")}`).toHaveLength(0);
  });
});

test("spaces typed at a table cell's end pin the caret at the cell edge", async ({ page }) => {
  await page.goto("/?doc=/fixtures/parity-tables.docx");
  await page.waitForSelector(".dxw-page span", { state: "attached" });
  await page.waitForTimeout(800);
  const ok = page.locator('.dxw-page span', { hasText: /^ok$/ }).first();
  const b = await ok.boundingBox();
  await page.mouse.click(b!.x + b!.width - 2, b!.y + b!.height / 2);
  await page.keyboard.press("End");
  await page.waitForTimeout(200);
  const caretX = () =>
    page.evaluate(() => {
      const divs = [...document.querySelectorAll(".dxw-page div")].filter(
        (d) => (d as HTMLElement).style.width === "1.5px" && (d as HTMLElement).style.display !== "none",
      );
      return divs.length ? divs[0].getBoundingClientRect().x : null;
    });
  const xs: number[] = [];
  for (let i = 0; i < 20; i++) {
    await page.keyboard.press("Space");
    await page.waitForTimeout(60);
    const x = await caretX();
    expect(x).not.toBeNull();
    xs.push(x!);
  }
  // The caret advances at first, then pins at the cell's content edge —
  // it must never keep walking into the neighboring cell (Word pins).
  const lastFive = xs.slice(-5);
  expect(Math.max(...lastFive) - Math.min(...lastFive)).toBeLessThan(0.5);
  // And it never travelled beyond ~one cell width.
  expect(Math.max(...xs) - xs[0]).toBeLessThan(200);
});
