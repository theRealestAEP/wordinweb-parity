/**
 * The canonical per-page parity metric.
 *
 * This is the exact function scripts/parity-compare.mjs has always handed to
 * page.evaluate, moved here unchanged so scripts/edit-roundtrip-parity.mjs can
 * score its pages on the same scale instead of inventing a second one. It is a
 * pure function of its arguments — nothing closes over Node scope — because
 * Playwright serializes it into the browser.
 *
 * Raw mismatch percentage cannot separate a visually excellent page from a
 * broken one: Word-PDF and Chrome disagree on sub-pixel glyph placement
 * everywhere, which floors any per-pixel threshold around 1%. `severityPct` is
 * the alignment-aware answer — ink with no counterpart within a small spatial
 * tolerance, plus corroborated line reflow — and reads 0 on a page that merely
 * rasterizes differently.
 *
 * The four semantic layer arguments (word text/image/vector and the web
 * semantic composite) drive the appearance metrics only. Pass null for them and
 * the structural fields, `severityPct` included, are still computed in full.
 *
 * Version this with METRIC_VERSION in parity-report.mjs: changing the maths
 * makes historical numbers incomparable.
 */
export const pageMetric =
async ([
  wordB64,
  webB64,
  wordTextB64,
  wordImageB64,
  wordVectorB64,
  semanticB64,
  pageStatus,
  semanticMetadata,
]) => {
  const load = (b64) =>
    new Promise((res, rej) => {
      const img = new Image();
      img.onload = () => res(img);
      img.onerror = rej;
      img.src = "data:image/png;base64," + b64;
    });
  const [a, b, wordText, wordImage, wordVector, semantic] = await Promise.all([
    wordB64 ? load(wordB64) : null,
    webB64 ? load(webB64) : null,
    wordTextB64 ? load(wordTextB64) : null,
    wordImageB64 ? load(wordImageB64) : null,
    wordVectorB64 ? load(wordVectorB64) : null,
    semanticB64 ? load(semanticB64) : null,
  ]);
  const reference = a || b;
  const dimensionMismatch =
    a && b && (Math.abs(a.width - b.width) > 2 || Math.abs(a.height - b.height) > 2);
  const w = a && b ? Math.max(a.width, b.width) : reference.width;
  const h = a && b ? Math.max(a.height, b.height) : reference.height;
  const px = (img) => {
    const c = document.createElement("canvas");
    c.width = w;
    c.height = h;
    const g = c.getContext("2d");
    g.fillStyle = "#fff";
    g.fillRect(0, 0, w, h);
    if (img) g.drawImage(img, 0, 0);
    return g.getImageData(0, 0, w, h);
  };
  const da = px(a);
  const db = px(b);
  const diff = new ImageData(w, h);
  let mismatch = 0;
  for (let p = 0; p < w * h * 4; p += 4) {
    const d =
      Math.abs(da.data[p] - db.data[p]) +
      Math.abs(da.data[p + 1] - db.data[p + 1]) +
      Math.abs(da.data[p + 2] - db.data[p + 2]);
    const bad = d > 90;
    if (bad) mismatch++;
    diff.data[p] = bad ? 255 : da.data[p];
    diff.data[p + 1] = bad ? 0 : da.data[p + 1];
    diff.data[p + 2] = bad ? 0 : da.data[p + 2];
    diff.data[p + 3] = bad ? 255 : 60;
  }
  // --- Classified drift metrics (alignment / weight / color / structural)
  // Raw page-area % can't separate a visually-excellent page from a broken
  // one: Word-PDF vs Chrome rasterization has pervasive sub-pixel glyph
  // placement differences that dominate any per-pixel or block-mean
  // threshold. The fix is alignment-aware analysis. Downsample device px to
  // CSS px (SCALE=2 -> average 2x2), register one global page offset, then
  // compare binary ink support with a small spatial tolerance. Rasterizers
  // may draw the same glyph differently inside that tolerance; missing,
  // extra, reflowed, or genuinely displaced content remains unmatched. We
  // then classify the residual, shift magnitude, ink-mass ratio, and ink
  // color into four deterministic buckets.
  const S = 2; // device px per CSS px (matches SCALE)
  const cw = Math.floor(w / S);
  const ch = Math.floor(h / S);
  // CSS-resolution luminance + color, averaged over each SxS device block.
  const wGray = new Float32Array(cw * ch);
  const bGray = new Float32Array(cw * ch);
  const wR = new Float32Array(cw * ch);
  const wGc = new Float32Array(cw * ch);
  const wB = new Float32Array(cw * ch);
  const bR = new Float32Array(cw * ch);
  const bGc = new Float32Array(cw * ch);
  const bB = new Float32Array(cw * ch);
  const inv = 1 / (S * S);
  for (let cy = 0; cy < ch; cy++) {
    for (let cx = 0; cx < cw; cx++) {
      let ar = 0, ag = 0, ab = 0, br = 0, bg = 0, bb = 0;
      for (let sy = 0; sy < S; sy++) {
        for (let sx = 0; sx < S; sx++) {
          const p = ((cy * S + sy) * w + (cx * S + sx)) * 4;
          ar += da.data[p]; ag += da.data[p + 1]; ab += da.data[p + 2];
          br += db.data[p]; bg += db.data[p + 1]; bb += db.data[p + 2];
        }
      }
      const ci = cy * cw + cx;
      wR[ci] = ar * inv; wGc[ci] = ag * inv; wB[ci] = ab * inv;
      bR[ci] = br * inv; bGc[ci] = bg * inv; bB[ci] = bb * inv;
      wGray[ci] = (wR[ci] + wGc[ci] + wB[ci]) / 3;
      bGray[ci] = (bR[ci] + bGc[ci] + bB[ci]) / 3;
    }
  }

  // Tunable constants (device-independent, in CSS px / luminance units).
  const T = 32; // tile edge, CSS px
  const SEARCH = 4; // max |offset| searched per axis, CSS px
  const INK_MIN = T * T * 255 * 0.008; // >=~0.8% coverage => tile has ink
  const STRUCT_INK_CUTOFF = 240; // ignore only near-white anti-aliasing
  const STRUCT_EDGE = 2; // exclude the screenshot's page-frame pixels
  // Radius 6 is the smallest calibration that puts the two aligned RTL
  // pages below 10%; their first visibly reflowed page remains above 15%.
  const STRUCT_TOL = 6;
  const MATCH_F = 0.35; // residual fraction below which a tile is "matched"
  const MISALIGN_PX = 1.0; // |offset| above this = "misaligned" content
  const LINE_SEARCH = 28; // wider search reveals matches saturated at the 24px scoring edge
  const LINE_MAX_SHIFT = 24; // farther matches are ambiguous repeated rows, not line-order evidence
  const LINE_STEP = 4; // coarse Y step; refine the winning band to 1px
  const LINE_SAMPLE = 2; // sample every other tile pixel in both axes
  const LINE_MIN_SHIFT = 8; // ignore ordinary glyph/baseline jitter
  const LINE_IMPROVE = 0.25; // far match must reduce tile SAD by at least 25%
  // A tile may vote "reflowed" only if it actually MISMATCHES at its
  // registered position. A genuinely moved line leaves near-total SAD
  // at base; a pixel-exact line's base SAD is antialiasing noise (a few
  // % of mass), and among self-similar ink (equation stacks, repeated
  // digits) a spurious far match easily clears the 25% RELATIVE bar -
  // parity-math p1 scored 5.55 lineShift with 0.22% raw pixel mismatch,
  // which is self-contradictory (real reflow of 5.5% of ink would leave
  // ~10x that raw mismatch). Requiring 20% base mismatch keeps every
  // true reflow (verified suite-wide: only pixel-exact pages moved).
  const LINE_NOISE_FLOOR = 5; // first 5% of flagged ink is calibration noise

  const txN = Math.floor(cw / T);
  const tyN = Math.floor(ch / T);
  // Sum of Absolute Differences between the Word tile at (ox,oy) and the web
  // image shifted by (dx,dy). Web sampled outside its bounds reads white.
  const sadAt = (ox, oy, dx, dy) => {
    let sad = 0;
    for (let ly = 0; ly < T; ly++) {
      const wy = oy + ly;
      const by2 = wy + dy;
      for (let lx = 0; lx < T; lx++) {
        const wx = ox + lx;
        const wv = wGray[wy * cw + wx];
        const bx2 = wx + dx;
        const bv =
          bx2 < 0 || by2 < 0 || bx2 >= cw || by2 >= ch ? 255 : bGray[by2 * cw + bx2];
        sad += Math.abs(wv - bv);
      }
    }
    return sad;
  };
  const lineSadAt = (ox, oy, dx, dy) => {
    let sad = 0;
    for (let ly = 0; ly < T; ly += LINE_SAMPLE) {
      const wy = oy + ly;
      const by2 = wy + dy;
      for (let lx = 0; lx < T; lx += LINE_SAMPLE) {
        const wx = ox + lx;
        const wv = wGray[wy * cw + wx];
        const bx2 = wx + dx;
        const bv =
          bx2 < 0 || by2 < 0 || bx2 >= cw || by2 >= ch ? 255 : bGray[by2 * cw + bx2];
        sad += Math.abs(wv - bv);
      }
    }
    return sad;
  };

  // Pass 1: for every inked tile find the local shift (+/-SEARCH) that best
  // registers Word to web. These per-tile offsets seed a single GLOBAL page
  // offset (their median) so a uniform whole-page shift - a benign, common
  // Word-vs-Chrome difference - is not mistaken for drift.
  const tiles = [];
  // Registration voters: only tiles whose search actually discriminates
  // an offset (see "discriminated" below). Instrumented on probe-wrapclear
  // (scripts/README.md #117): a uniform-fill region (e.g. a solid box
  // interior) returns the IDENTICAL SAD at every candidate offset - not
  // approximately equal, exactly equal, because every offset still lands
  // fill-on-fill. Such a tile's "best" offset is whichever candidate the
  // scan order happened to try first, not a measurement. Feeding that into
  // the registration median drags the whole page's global offset off the
  // real content. A tile with any texture (an edge, a rule, a glyph) always
  // has a strictly-better offset than its runner-up.
  const bxs = [];
  const bys = [];
  for (let ty = 0; ty < tyN; ty++) {
    for (let tx = 0; tx < txN; tx++) {
      const ox = tx * T;
      const oy = ty * T;
      // Static ink masses (zero offset) decide whether the tile matters and
      // how much it weighs; max() so web-only extra content still counts.
      let wordInk = 0;
      let webInk0 = 0;
      for (let ly = 0; ly < T; ly++) {
        for (let lx = 0; lx < T; lx++) {
          const ci = (oy + ly) * cw + (ox + lx);
          wordInk += 255 - wGray[ci];
          webInk0 += 255 - bGray[ci];
        }
      }
      if (wordInk < INK_MIN && webInk0 < INK_MIN) continue;

      // Coarse-to-fine SAD search: 5x5 grid at step 2, then refine +/-1
      // around the winner. Deterministic; ~34 candidates vs 81 exhaustive.
      // Track the runner-up SAD alongside the winner so a flat landscape
      // (best === second) can be told apart from a real match below.
      let bx = 0, by = 0, best = Infinity, second = Infinity;
      for (let dy = -SEARCH; dy <= SEARCH; dy += 2) {
        for (let dx = -SEARCH; dx <= SEARCH; dx += 2) {
          const s = sadAt(ox, oy, dx, dy);
          if (s < best) { second = best; best = s; bx = dx; by = dy; }
          else if (s < second) { second = s; }
        }
      }
      const cbx = bx, cby = by;
      for (let dy = cby - 1; dy <= cby + 1; dy++) {
        for (let dx = cbx - 1; dx <= cbx + 1; dx++) {
          if (dx < -SEARCH || dx > SEARCH || dy < -SEARCH || dy > SEARCH) continue;
          if (dx === cbx && dy === cby) continue;
          const s = sadAt(ox, oy, dx, dy);
          if (s < best) { second = best; best = s; bx = dx; by = dy; }
          else if (s < second) { second = s; }
        }
      }
      // Discriminated: the winning offset actually beats its runner-up. A
      // tile that ties every candidate (best === second, always exactly
      // equal for a uniform fill - never approximately) has no vote to cast.
      const discriminated = second > best;
      tiles.push({ ox, oy, wordInk, webInk0, tileMass: (wordInk + webInk0) / 2, bx, by, discriminated });
      if (discriminated) {
        bxs.push(bx);
        bys.push(by);
      }
    }
  }
  // A page with no discriminated tile at all (nothing but flat fill) has no
  // better signal than the raw population - fall back to it rather than
  // registering on an empty vote.
  const regBxs = bxs.length ? bxs : tiles.map((t) => t.bx);
  const regBys = bys.length ? bys : tiles.map((t) => t.by);

  const median = (arr) => {
    if (!arr.length) return 0;
    const s = [...arr].sort((p, q) => p - q);
    const m = s.length >> 1;
    return s.length % 2 ? s[m] : Math.round((s[m - 1] + s[m]) / 2);
  };
  const gDx = median(regBxs);
  const gDy = median(regBys);

  // Full-page appearance metrics use only the one global registration above.
  // They deliberately include all page-interior content: no matched-tile
  // filter and no local shift can hide missing ink or let text/images/rules
  // cancel selectively.
  const appearanceStarted = performance.now();
  const darkness = (r, g, b) => 1 - (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255;
  const symmetricWeightError = (wordMass, webMass) => {
    const total = wordMass + webMass;
    return total ? (200 * Math.abs(webMass - wordMass)) / total : 0;
  };
  const rgbToLab = (r8, g8, b8) => {
    const linear = (value) => {
      const c = value / 255;
      return c <= 0.04045 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4;
    };
    const r = linear(r8), g = linear(g8), b = linear(b8);
    const x = (0.4124564 * r + 0.3575761 * g + 0.1804375 * b) / 0.95047;
    const y = 0.2126729 * r + 0.7151522 * g + 0.072175 * b;
    const z = (0.0193339 * r + 0.119192 * g + 0.9503041 * b) / 1.08883;
    const f = (value) =>
      value > 216 / 24389 ? Math.cbrt(value) : (24389 / 27 * value + 16) / 116;
    const fx = f(x), fy = f(y), fz = f(z);
    return [116 * fy - 16, 500 * (fx - fy), 200 * (fy - fz)];
  };
  const deltaE00 = (lab1, lab2) => {
    const [l1, a1, b1] = lab1;
    const [l2, a2, b2] = lab2;
    const c1 = Math.hypot(a1, b1);
    const c2 = Math.hypot(a2, b2);
    const cBar = (c1 + c2) / 2;
    const cBar7 = cBar ** 7;
    const g = 0.5 * (1 - Math.sqrt(cBar7 / (cBar7 + 25 ** 7)));
    const a1p = (1 + g) * a1;
    const a2p = (1 + g) * a2;
    const c1p = Math.hypot(a1p, b1);
    const c2p = Math.hypot(a2p, b2);
    const hue = (a, b) => {
      if (a === 0 && b === 0) return 0;
      const degrees = Math.atan2(b, a) * 180 / Math.PI;
      return degrees >= 0 ? degrees : degrees + 360;
    };
    const h1p = hue(a1p, b1);
    const h2p = hue(a2p, b2);
    const dl = l2 - l1;
    const dc = c2p - c1p;
    let dh = h2p - h1p;
    if (c1p * c2p === 0) dh = 0;
    else if (dh > 180) dh -= 360;
    else if (dh < -180) dh += 360;
    const dH = 2 * Math.sqrt(c1p * c2p) * Math.sin((dh / 2) * Math.PI / 180);
    const lBar = (l1 + l2) / 2;
    const cpBar = (c1p + c2p) / 2;
    let hpBar;
    if (c1p * c2p === 0) hpBar = h1p + h2p;
    else if (Math.abs(h1p - h2p) <= 180) hpBar = (h1p + h2p) / 2;
    else if (h1p + h2p < 360) hpBar = (h1p + h2p + 360) / 2;
    else hpBar = (h1p + h2p - 360) / 2;
    const rad = Math.PI / 180;
    const t =
      1 - 0.17 * Math.cos((hpBar - 30) * rad) +
      0.24 * Math.cos(2 * hpBar * rad) +
      0.32 * Math.cos((3 * hpBar + 6) * rad) -
      0.20 * Math.cos((4 * hpBar - 63) * rad);
    const sl = 1 + 0.015 * (lBar - 50) ** 2 / Math.sqrt(20 + (lBar - 50) ** 2);
    const sc = 1 + 0.045 * cpBar;
    const sh = 1 + 0.015 * cpBar * t;
    const cpBar7 = cpBar ** 7;
    const rt =
      -2 * Math.sqrt(cpBar7 / (cpBar7 + 25 ** 7)) *
      Math.sin(60 * Math.exp(-(((hpBar - 275) / 25) ** 2)) * rad);
    const lTerm = dl / sl;
    const cTerm = dc / sc;
    const hTerm = dH / sh;
    return Math.sqrt(lTerm ** 2 + cTerm ** 2 + hTerm ** 2 + rt * cTerm * hTerm);
  };
  // Recover source ink chromaticity from pixels composited over white.
  // Opacity, antialiasing, and stroke weight scale all three subtractive
  // channels together, so normalizing that scale keeps them in the page-
  // weight metric instead of falsely reporting black text as wrong colour.
  const normalizedInkRgb = (r, g, b) => {
    const ar = 255 - r, ag = 255 - g, ab = 255 - b;
    const scale = Math.max(ar, ag, ab);
    if (scale <= 0) return [255, 255, 255];
    return [255 - 255 * ar / scale, 255 - 255 * ag / scale, 255 - 255 * ab / scale];
  };
  const assertNear = (actual, expected, tolerance, label) => {
    if (Math.abs(actual - expected) > tolerance) {
      throw new Error(`${label} self-check failed: ${actual} vs ${expected}`);
    }
  };
  assertNear(symmetricWeightError(2, 1), 200 / 3, 1e-9, "weight symmetry");
  assertNear(symmetricWeightError(1, 2), 200 / 3, 1e-9, "weight reverse symmetry");
  const redLab = rgbToLab(255, 0, 0);
  assertNear(redLab[0], 53.2408, 0.001, "sRGB red L*");
  assertNear(redLab[1], 80.0925, 0.001, "sRGB red a*");
  assertNear(redLab[2], 67.2032, 0.001, "sRGB red b*");
  assertNear(
    deltaE00([50, 2.6772, -79.7751], [50, 0, -82.7485]),
    2.0425,
    0.0001,
    "CIEDE2000 reference",
  );
  assertNear(
    deltaE00(rgbToLab(...normalizedInkRgb(0, 0, 0)), rgbToLab(...normalizedInkRgb(160, 160, 160))),
    0,
    1e-9,
    "black/gray chromaticity",
  );
  const orange = [255, 102, 0];
  const fadedOrange = orange.map((channel) => 255 - 0.55 * (255 - channel));
  assertNear(
    deltaE00(
      rgbToLab(...normalizedInkRgb(...orange)),
      rgbToLab(...normalizedInkRgb(...fadedOrange)),
    ),
    0,
    1e-9,
    "opacity-invariant chromaticity",
  );
  if (
    deltaE00(
      rgbToLab(...normalizedInkRgb(...orange)),
      rgbToLab(...normalizedInkRgb(0, 102, 255)),
    ) < 40
  ) {
    throw new Error("orange/blue chromaticity self-check failed");
  }
  if (deltaE00(rgbToLab(0, 0, 0), rgbToLab(160, 160, 160)) < 25) {
    throw new Error("semantic black/gray colour self-check failed");
  }

  let wordAppearanceMass = 0;
  let webAppearanceMass = 0;
  for (let y = STRUCT_EDGE; y < ch - STRUCT_EDGE; y++) {
    for (let x = STRUCT_EDGE; x < cw - STRUCT_EDGE; x++) {
      const wi = y * cw + x;
      wordAppearanceMass += darkness(wR[wi], wGc[wi], wB[wi]);
      const bx = x + gDx, by = y + gDy;
      if (bx >= 0 && by >= 0 && bx < cw && by < ch) {
        const bi = by * cw + bx;
        webAppearanceMass += darkness(bR[bi], bGc[bi], bB[bi]);
      }
    }
  }
  const appearanceWeightRatio =
    wordAppearanceMass > 0
      ? webAppearanceMass / wordAppearanceMass
      : webAppearanceMass === 0 ? 1 : null;
  const appearanceWeightErrorPct = symmetricWeightError(
    wordAppearanceMass,
    webAppearanceMass,
  );

  // Sigma 0.5 CSS px: a fixed separable three-tap Gaussian suppresses only
  // antialias phase noise before the globally aligned colour comparison.
  const blur = (input) => {
    const side = 0.1065069789;
    const center = 0.7869860422;
    const tmp = new Float32Array(input.length);
    const out = new Float32Array(input.length);
    for (let y = 0; y < ch; y++) {
      const row = y * cw;
      for (let x = 0; x < cw; x++) {
        tmp[row + x] =
          side * input[row + Math.max(0, x - 1)] +
          center * input[row + x] +
          side * input[row + Math.min(cw - 1, x + 1)];
      }
    }
    for (let y = 0; y < ch; y++) {
      const up = Math.max(0, y - 1) * cw;
      const row = y * cw;
      const down = Math.min(ch - 1, y + 1) * cw;
      for (let x = 0; x < cw; x++) {
        out[row + x] = side * tmp[up + x] + center * tmp[row + x] + side * tmp[down + x];
      }
    }
    return out;
  };
  const wRBlur = blur(wR), wGBlur = blur(wGc), wBBlur = blur(wB);
  const bRBlur = blur(bR), bGBlur = blur(bGc), bBBlur = blur(bB);
  let appearanceColorNum = 0;
  let appearanceColorDen = 0;
  for (let y = STRUCT_EDGE; y < ch - STRUCT_EDGE; y++) {
    for (let x = STRUCT_EDGE; x < cw - STRUCT_EDGE; x++) {
      const bx = x + gDx, by = y + gDy;
      const wi = y * cw + x;
      const inWeb = bx >= 0 && by >= 0 && bx < cw && by < ch;
      const bi = inWeb ? by * cw + bx : -1;
      const wr = wRBlur[wi], wg = wGBlur[wi], wb = wBBlur[wi];
      const br = inWeb ? bRBlur[bi] : 255;
      const bg = inWeb ? bGBlur[bi] : 255;
      const bb = inWeb ? bBBlur[bi] : 255;
      const wordDarkness = darkness(wr, wg, wb);
      const webDarkness = darkness(br, bg, bb);
      const inkWeight = Math.min(wordDarkness, webDarkness);
      if (inkWeight <= 1 / 255) continue;
      appearanceColorNum +=
        inkWeight * Math.min(
          deltaE00(
            rgbToLab(...normalizedInkRgb(wr, wg, wb)),
            rgbToLab(...normalizedInkRgb(br, bg, bb)),
          ),
          100,
        );
      appearanceColorDen += inkWeight;
    }
  }
  const appearanceColorDeltaE =
    appearanceColorDen ? appearanceColorNum / appearanceColorDen : 0;

  // Category layers are independently rendered on transparent backgrounds:
  // Word through filtered PDFs, web through detached clones of the same page.
  // Font weight is alpha coverage, while image/fill weight is visible darkness
  // after compositing onto white. No category-local registration is allowed.
  const categoryStarted = performance.now();
  const hasImages = (semanticMetadata.images?.length ?? 0) > 0;
  const hasTableFills = (semanticMetadata.fills?.length ?? 0) > 0;
  const hasTableRules = (semanticMetadata.rules?.length ?? 0) > 0;
  const needsVectorLayer = hasTableFills || hasTableRules ||
    semanticMetadata.images?.some((region) => /^(emf|wmf|svg)$/i.test(region.format));
  const layerDimensionsValid =
    semantic && b &&
    Math.abs(semantic.width / 4 - b.width) <= 2 &&
    Math.abs(semantic.height - b.height) <= 2 &&
    wordText && a &&
    Math.abs(wordText.width - a.width) <= 2 &&
    Math.abs(wordText.height - a.height) <= 2 &&
    (!hasImages || wordImage) &&
    (!wordImage || (Math.abs(wordImage.width - a.width) <= 2 && Math.abs(wordImage.height - a.height) <= 2)) &&
    (!needsVectorLayer || wordVector) &&
    (!wordVector || (Math.abs(wordVector.width - a.width) <= 2 && Math.abs(wordVector.height - a.height) <= 2));

  const rgbaLayer = (img, panel = null) => {
    if (!img) return null;
    const canvas = document.createElement("canvas");
    canvas.width = w;
    canvas.height = h;
    const context = canvas.getContext("2d");
    if (panel == null) {
      context.drawImage(img, 0, 0);
    } else {
      const panelWidth = img.width / 4;
      context.drawImage(
        img,
        panel * panelWidth,
        0,
        panelWidth,
        img.height,
        0,
        0,
        b.width,
        b.height,
      );
    }
    const pixels = context.getImageData(0, 0, w, h).data;
    const A = new Float32Array(cw * ch);
    const R = new Float32Array(cw * ch);
    const G = new Float32Array(cw * ch);
    const B = new Float32Array(cw * ch);
    for (let y = 0; y < ch; y++) {
      for (let x = 0; x < cw; x++) {
        let alphaSum = 0;
        let redPremul = 0;
        let greenPremul = 0;
        let bluePremul = 0;
        for (let sy = 0; sy < S; sy++) {
          for (let sx = 0; sx < S; sx++) {
            const offset = ((y * S + sy) * w + x * S + sx) * 4;
            const alpha = pixels[offset + 3] / 255;
            alphaSum += alpha;
            redPremul += pixels[offset] * alpha;
            greenPremul += pixels[offset + 1] * alpha;
            bluePremul += pixels[offset + 2] * alpha;
          }
        }
        const index = y * cw + x;
        A[index] = alphaSum * inv;
        if (alphaSum > 0) {
          R[index] = redPremul / alphaSum;
          G[index] = greenPremul / alphaSum;
          B[index] = bluePremul / alphaSum;
        } else {
          R[index] = G[index] = B[index] = 255;
        }
      }
    }
    return { A, R, G, B };
  };

  const wordTextLayer = layerDimensionsValid ? rgbaLayer(wordText) : null;
  const wordImageLayer = layerDimensionsValid && hasImages ? rgbaLayer(wordImage) : null;
  const wordVectorLayer = layerDimensionsValid && needsVectorLayer ? rgbaLayer(wordVector) : null;
  const webTextLayer = layerDimensionsValid ? rgbaLayer(semantic, 0) : null;
  const webImageLayer = layerDimensionsValid && hasImages ? rgbaLayer(semantic, 1) : null;
  const webFillLayer = layerDimensionsValid && (hasTableFills || hasTableRules)
    ? rgbaLayer(semantic, 2)
    : null;
  const webRuleLayer = layerDimensionsValid && hasTableRules ? rgbaLayer(semantic, 3) : null;
  const semanticLayersTransparent =
    webTextLayer == null || webTextLayer.A[0] <= 1 / 255;
  const categoryMetricStatus = !layerDimensionsValid
    ? "layer-dimension-mismatch"
    : !semanticLayersTransparent
      ? "web-layer-background-not-transparent"
      : "ok";
  const semanticValid = categoryMetricStatus === "ok";

  const buildOwner = (regions, padding = 0) => {
    const owner = new Int32Array(cw * ch);
    owner.fill(-1);
    regions.forEach((rect, index) => {
      const x1 = Math.max(STRUCT_EDGE, Math.floor(rect.x - padding));
      const y1 = Math.max(STRUCT_EDGE, Math.floor(rect.y - padding));
      const x2 = Math.min(cw - STRUCT_EDGE, Math.ceil(rect.x + rect.width + padding));
      const y2 = Math.min(ch - STRUCT_EDGE, Math.ceil(rect.y + rect.height + padding));
      for (let y = y1; y < y2; y++) {
        for (let x = x1; x < x2; x++) owner[y * cw + x] = index;
      }
    });
    return owner;
  };
  const ratio = (wordMass, webMass) =>
    wordMass > 0 ? webMass / wordMass : webMass === 0 ? 1 : null;
  const roundMetric = (value, digits = 2) =>
    value == null || !Number.isFinite(value) ? null : Number(value.toFixed(digits));
  const visibleDarkness = (layer, index) => {
    if (!layer) return 0;
    return layer.A[index] * darkness(layer.R[index], layer.G[index], layer.B[index]);
  };
  assertNear(
    200 * (Math.abs(2 - 1) + Math.abs(1 - 2)) / (2 + 1 + 1 + 2),
    200 / 3,
    1e-9,
    "per-region non-cancellation",
  );
  assertNear(
    visibleDarkness({ A: [0.5], R: [0], G: [0], B: [0] }, 0),
    0.5,
    1e-12,
    "alpha-weighted visible darkness",
  );

  let textWeightRatio = null;
  let textWeightErrorPct = null;
  let textWeightMass = 0;
  let textColorDeltaE = null;
  let textColorMass = 0;
  let textColorCoveragePct = null;
  if (semanticValid && wordTextLayer && webTextLayer) {
    const regions = semanticMetadata.text ?? [];
    const owner = buildOwner(regions, 1);
    const wordMasses = new Float64Array(regions.length);
    const webMasses = new Float64Array(regions.length);
    const wordRgb = new Float64Array(regions.length * 3);
    const webRgb = new Float64Array(regions.length * 3);
    let unmatchedWord = 0;
    let unmatchedWeb = 0;
    for (let by = STRUCT_EDGE; by < ch - STRUCT_EDGE; by++) {
      for (let bx = STRUCT_EDGE; bx < cw - STRUCT_EDGE; bx++) {
        const bi = by * cw + bx;
        const webAlpha = webTextLayer.A[bi];
        const webOwner = owner[bi];
        if (webOwner >= 0) {
          webMasses[webOwner] += webAlpha;
          const rgbOffset = webOwner * 3;
          webRgb[rgbOffset] += webTextLayer.R[bi] * webAlpha;
          webRgb[rgbOffset + 1] += webTextLayer.G[bi] * webAlpha;
          webRgb[rgbOffset + 2] += webTextLayer.B[bi] * webAlpha;
        } else {
          unmatchedWeb += webAlpha;
        }

        const wx = bx - gDx;
        const wy = by - gDy;
        if (wx < STRUCT_EDGE || wy < STRUCT_EDGE || wx >= cw - STRUCT_EDGE || wy >= ch - STRUCT_EDGE) {
          continue;
        }
        const wi = wy * cw + wx;
        const wordAlpha = wordTextLayer.A[wi];
        if (webOwner >= 0) {
          wordMasses[webOwner] += wordAlpha;
          const rgbOffset = webOwner * 3;
          wordRgb[rgbOffset] += wordTextLayer.R[wi] * wordAlpha;
          wordRgb[rgbOffset + 1] += wordTextLayer.G[wi] * wordAlpha;
          wordRgb[rgbOffset + 2] += wordTextLayer.B[wi] * wordAlpha;
        } else {
          unmatchedWord += wordAlpha;
        }
      }
    }

    let wordTotal = unmatchedWord;
    let webTotal = unmatchedWeb;
    let weightErrorNum = unmatchedWord + unmatchedWeb;
    let colorNum = 0;
    for (let index = 0; index < regions.length; index++) {
      const wordMass = wordMasses[index];
      const webMass = webMasses[index];
      wordTotal += wordMass;
      webTotal += webMass;
      weightErrorNum += Math.abs(webMass - wordMass);
      const matchedMass = Math.min(wordMass, webMass);
      if (matchedMass <= 1 / 255) continue;
      const rgbOffset = index * 3;
      const wordColor = [
        wordRgb[rgbOffset] / wordMass,
        wordRgb[rgbOffset + 1] / wordMass,
        wordRgb[rgbOffset + 2] / wordMass,
      ];
      const webColor = [
        webRgb[rgbOffset] / webMass,
        webRgb[rgbOffset + 1] / webMass,
        webRgb[rgbOffset + 2] / webMass,
      ];
      colorNum += matchedMass * Math.min(
        deltaE00(rgbToLab(...wordColor), rgbToLab(...webColor)),
        100,
      );
      textColorMass += matchedMass;
    }
    textWeightMass = wordTotal + webTotal;
    textWeightRatio = ratio(wordTotal, webTotal);
    textWeightErrorPct = textWeightMass
      ? 200 * weightErrorNum / textWeightMass
      : null;
    textColorDeltaE = textColorMass ? colorNum / textColorMass : null;
    textColorCoveragePct = textWeightMass
      ? 200 * textColorMass / textWeightMass
      : null;
  }

  const measureDarkRegions = (regions, webLayer, referenceAt, inset) => {
    let wordTotal = 0;
    let webTotal = 0;
    let errorNum = 0;
    for (const rect of regions) {
      let wordMass = 0;
      let webMass = 0;
      const x1 = Math.max(STRUCT_EDGE, Math.ceil(rect.x + inset));
      const y1 = Math.max(STRUCT_EDGE, Math.ceil(rect.y + inset));
      const x2 = Math.min(cw - STRUCT_EDGE, Math.floor(rect.x + rect.width - inset));
      const y2 = Math.min(ch - STRUCT_EDGE, Math.floor(rect.y + rect.height - inset));
      for (let by = y1; by < y2; by++) {
        for (let bx = x1; bx < x2; bx++) {
          const bi = by * cw + bx;
          webMass += visibleDarkness(webLayer, bi);
          const wx = bx - gDx;
          const wy = by - gDy;
          if (wx < 0 || wy < 0 || wx >= cw || wy >= ch) continue;
          wordMass += referenceAt(wy * cw + wx, rect);
        }
      }
      wordTotal += wordMass;
      webTotal += webMass;
      errorNum += Math.abs(webMass - wordMass);
    }
    const mass = wordTotal + webTotal;
    return {
      ratio: ratio(wordTotal, webTotal),
      error: mass ? 200 * errorNum / mass : null,
      mass,
    };
  };

  let imageWeightRatio = null;
  let imageWeightErrorPct = null;
  let imageWeightMass = 0;
  const imageRegions = semanticMetadata.images ?? [];
  if (semanticValid && webImageLayer && imageRegions.length) {
    const measured = measureDarkRegions(
      imageRegions,
      webImageLayer,
      (index, rect) => {
        const raster = visibleDarkness(wordImageLayer, index);
        return /^(emf|wmf|svg)$/i.test(rect.format)
          ? Math.max(raster, visibleDarkness(wordVectorLayer, index))
          : raster;
      },
      1,
    );
    imageWeightRatio = measured.ratio;
    imageWeightErrorPct = measured.error;
    imageWeightMass = measured.mass;
  }

  let tableFillWeightRatio = null;
  let tableFillWeightErrorPct = null;
  let tableFillWeightMass = 0;
  const fillRegions = semanticMetadata.fills ?? [];
  if (semanticValid && webFillLayer && wordVectorLayer && fillRegions.length) {
    const measured = measureDarkRegions(
      fillRegions,
      webFillLayer,
      (index) => visibleDarkness(wordVectorLayer, index),
      2,
    );
    tableFillWeightRatio = measured.ratio;
    tableFillWeightErrorPct = measured.error;
    tableFillWeightMass = measured.mass;
  }

  const compositeWhite = (layer, index) => {
    if (!layer) return [255, 255, 255];
    const alpha = layer.A[index];
    return [
      layer.R[index] * alpha + 255 * (1 - alpha),
      layer.G[index] * alpha + 255 * (1 - alpha),
      layer.B[index] * alpha + 255 * (1 - alpha),
    ];
  };
  const compositeLayersOnWhite = (top, bottom, index) => {
    const topAlpha = top?.A[index] ?? 0;
    const bottomAlpha = bottom?.A[index] ?? 0;
    const bottomShare = bottomAlpha * (1 - topAlpha);
    const whiteShare = (1 - topAlpha) * (1 - bottomAlpha);
    return [0, 1, 2].map((channel) =>
      (top?.[channel === 0 ? "R" : channel === 1 ? "G" : "B"]?.[index] ?? 255) * topAlpha +
      (bottom?.[channel === 0 ? "R" : channel === 1 ? "G" : "B"]?.[index] ?? 255) * bottomShare +
      255 * whiteShare,
    );
  };
  const rgbContrast = (a, b) =>
    Math.hypot(a[0] - b[0], a[1] - b[1], a[2] - b[2]) / (255 * Math.sqrt(3));
  assertNear(rgbContrast([0, 0, 0], [255, 255, 255]), 1, 1e-12, "rule contrast");
  assertNear(rgbContrast([80, 90, 100], [80, 90, 100]), 0, 1e-12, "uniform fill contrast");

  let tableRuleWeightRatio = null;
  let tableRuleWeightErrorPct = null;
  let tableRuleWeightMass = 0;
  const ruleRegions = semanticMetadata.rules ?? [];
  if (semanticValid && webRuleLayer && wordVectorLayer && ruleRegions.length) {
    const axis = new Uint8Array(cw * ch);
    const centers = new Float32Array(cw * ch);
    for (const rect of ruleRegions) {
      const horizontal = rect.axis === "horizontal";
      const value = horizontal ? 1 : 2;
      const center = horizontal
        ? rect.y + rect.height / 2
        : rect.x + rect.width / 2;
      const x1 = Math.max(STRUCT_EDGE, Math.floor(horizontal ? rect.x : center - 2));
      const x2 = Math.min(cw - STRUCT_EDGE, Math.ceil(horizontal ? rect.x + rect.width : center + 2));
      const y1 = Math.max(STRUCT_EDGE, Math.floor(horizontal ? center - 2 : rect.y));
      const y2 = Math.min(ch - STRUCT_EDGE, Math.ceil(horizontal ? center + 2 : rect.y + rect.height));
      for (let y = y1; y < y2; y++) {
        for (let x = x1; x < x2; x++) {
          const index = y * cw + x;
          if (axis[index] && axis[index] !== value) axis[index] = 3;
          else axis[index] = value;
          centers[index] = center;
        }
      }
    }

    const wordMass = [0, 0];
    const webMass = [0, 0];
    for (let by = STRUCT_EDGE + 4; by < ch - STRUCT_EDGE - 4; by++) {
      for (let bx = STRUCT_EDGE + 4; bx < cw - STRUCT_EDGE - 4; bx++) {
        const bi = by * cw + bx;
        const kind = axis[bi];
        if (kind !== 1 && kind !== 2) continue;
        const webPaint = compositeLayersOnWhite(webRuleLayer, webFillLayer, bi);
        const webBackground = compositeWhite(webFillLayer, bi);
        webMass[kind - 1] += rgbContrast(webPaint, webBackground);

        const wx = bx - gDx;
        const wy = by - gDy;
        if (wx < 4 || wy < 4 || wx >= cw - 4 || wy >= ch - 4) continue;
        const wi = wy * cw + wx;
        const center = centers[bi];
        const offset = kind === 1 ? by + 0.5 - center : bx + 0.5 - center;
        const beforeIndex = kind === 1 ? (wy - 4) * cw + wx : wy * cw + wx - 4;
        const afterIndex = kind === 1 ? (wy + 4) * cw + wx : wy * cw + wx + 4;
        const before = compositeWhite(wordVectorLayer, beforeIndex);
        const after = compositeWhite(wordVectorLayer, afterIndex);
        const background = offset < 0
          ? before
          : offset > 0
            ? after
            : before.map((value, channel) => (value + after[channel]) / 2);
        wordMass[kind - 1] += rgbContrast(compositeWhite(wordVectorLayer, wi), background);
      }
    }
    const wordTotal = wordMass[0] + wordMass[1];
    const webTotal = webMass[0] + webMass[1];
    tableRuleWeightMass = wordTotal + webTotal;
    tableRuleWeightRatio = ratio(wordTotal, webTotal);
    tableRuleWeightErrorPct = tableRuleWeightMass
      ? 200 * (Math.abs(webMass[0] - wordMass[0]) + Math.abs(webMass[1] - wordMass[1])) /
        tableRuleWeightMass
      : null;
  }

  textWeightRatio = roundMetric(textWeightRatio, 4);
  textWeightErrorPct = roundMetric(textWeightErrorPct);
  textWeightMass = roundMetric(textWeightMass, 2) ?? 0;
  textColorDeltaE = roundMetric(textColorDeltaE);
  textColorMass = roundMetric(textColorMass, 2) ?? 0;
  textColorCoveragePct = roundMetric(textColorCoveragePct);
  imageWeightRatio = roundMetric(imageWeightRatio, 4);
  imageWeightErrorPct = roundMetric(imageWeightErrorPct);
  imageWeightMass = roundMetric(imageWeightMass, 2) ?? 0;
  tableFillWeightRatio = roundMetric(tableFillWeightRatio, 4);
  tableFillWeightErrorPct = roundMetric(tableFillWeightErrorPct);
  tableFillWeightMass = roundMetric(tableFillWeightMass, 2) ?? 0;
  tableRuleWeightRatio = roundMetric(tableRuleWeightRatio, 4);
  tableRuleWeightErrorPct = roundMetric(tableRuleWeightErrorPct);
  tableRuleWeightMass = roundMetric(tableRuleWeightMass, 2) ?? 0;
  const categoryMetricMs = performance.now() - categoryStarted;
  const appearanceMetricMs = performance.now() - appearanceStarted;

  // A one-line vertical reflow can evade the local structural matcher by
  // landing on the neighbouring text row. Search farther in Y after the
  // page's global registration and count only ink whose tile match gets
  // materially better at least 8px away. Non-overlapping tiles let their
  // existing ink mass provide a stable page-level fraction.
  let lineShiftPositiveMass = 0;
  let lineShiftNegativeMass = 0;
  let lineTotalMass = 0;
  const lineDebugRows = [];
  for (const t of tiles) {
    const { ox, oy, tileMass } = t;
    const base = lineSadAt(ox, oy, gDx, gDy);
    let best = base;
    let bestDelta = 0;
    for (let delta = -LINE_SEARCH; delta <= LINE_SEARCH; delta += LINE_STEP) {
      if (delta === 0) continue;
      const s = lineSadAt(ox, oy, gDx, gDy + delta);
      if (s < best) { best = s; bestDelta = delta; }
    }
    const coarseDelta = bestDelta;
    for (let delta = coarseDelta - LINE_STEP / 2; delta <= coarseDelta + LINE_STEP / 2; delta++) {
      if (delta < -LINE_SEARCH || delta > LINE_SEARCH || delta === coarseDelta) continue;
      const s = lineSadAt(ox, oy, gDx, gDy + delta);
      if (s < best) { best = s; bestDelta = delta; }
    }
    lineTotalMass += tileMass;
    const improvement = base > 0 ? (base - best) / base : 0;
    const baseMismatched = base >= tileMass * 0.2;
    if (Math.abs(bestDelta) >= LINE_MIN_SHIFT && improvement >= LINE_IMPROVE) {
      lineDebugRows.push(`tile@${ox},${oy} mass=${tileMass.toFixed(0)} base/m=${(base / tileMass).toFixed(2)} best/m=${(best / tileMass).toFixed(2)} imp=${improvement.toFixed(2)} d=${bestDelta}`);
    }
    if (
      baseMismatched &&
      Math.abs(bestDelta) >= LINE_MIN_SHIFT &&
      Math.abs(bestDelta) <= LINE_MAX_SHIFT &&
      improvement >= LINE_IMPROVE
    ) {
      if (bestDelta > 0) lineShiftPositiveMass += tileMass;
      else lineShiftNegativeMass += tileMass;
    }
  }
  // Real line reflow has a coherent direction. Repeated table rows can
  // match equally well above and below; counting only the dominant sign
  // prevents those ambiguous neighbours from doubling the signal.
  const lineShiftMass = Math.max(lineShiftPositiveMass, lineShiftNegativeMass);
  const lineShiftRawPct = lineTotalMass ? (lineShiftMass / lineTotalMass) * 100 : 0;
  const lineShiftPct = Math.max(0, lineShiftRawPct - LINE_NOISE_FLOOR);

  // Align web ink to Word coordinates, then make a square-dilated support
  // mask. The symmetric unmatched-ink fraction below is a Dice distance
  // with tolerance: stroke weight and glyph outline differences disappear,
  // while ink with no counterpart anywhere nearby still counts in full.
  const wordMask = new Uint8Array(cw * ch);
  const webMask = new Uint8Array(cw * ch);
  for (let y = 0; y < ch; y++) {
    for (let x = 0; x < cw; x++) {
      const i = y * cw + x;
      wordMask[i] =
        x >= STRUCT_EDGE && y >= STRUCT_EDGE && x < cw - STRUCT_EDGE && y < ch - STRUCT_EDGE &&
        wGray[i] < STRUCT_INK_CUTOFF ? 1 : 0;
      const bx = x + gDx;
      const by = y + gDy;
      webMask[i] =
        bx >= STRUCT_EDGE && by >= STRUCT_EDGE &&
        bx < cw - STRUCT_EDGE && by < ch - STRUCT_EDGE &&
        bGray[by * cw + bx] < STRUCT_INK_CUTOFF ? 1 : 0;
    }
  }
  const dilate = (mask) => {
    const horizontal = new Uint8Array(mask.length);
    const out = new Uint8Array(mask.length);
    for (let y = 0; y < ch; y++) {
      const row = y * cw;
      let count = 0;
      for (let x = 0; x <= Math.min(STRUCT_TOL, cw - 1); x++) count += mask[row + x];
      for (let x = 0; x < cw; x++) {
        horizontal[row + x] = count > 0 ? 1 : 0;
        const leaving = x - STRUCT_TOL;
        const entering = x + STRUCT_TOL + 1;
        if (leaving >= 0) count -= mask[row + leaving];
        if (entering < cw) count += mask[row + entering];
      }
    }
    for (let x = 0; x < cw; x++) {
      let count = 0;
      for (let y = 0; y <= Math.min(STRUCT_TOL, ch - 1); y++) count += horizontal[y * cw + x];
      for (let y = 0; y < ch; y++) {
        out[y * cw + x] = count > 0 ? 1 : 0;
        const leaving = y - STRUCT_TOL;
        const entering = y + STRUCT_TOL + 1;
        if (leaving >= 0) count -= horizontal[leaving * cw + x];
        if (entering < ch) count += horizontal[entering * cw + x];
      }
    }
    return out;
  };
  const wordDilated = dilate(wordMask);
  const webDilated = dilate(webMask);

  let structNum = 0;
  let structDen = 0;
  for (let i = 0; i < wordMask.length; i++) {
    if (wordMask[i]) {
      structDen++;
      if (!webDilated[i]) structNum++;
    }
    if (webMask[i]) {
      structDen++;
      if (!wordDilated[i]) structNum++;
    }
  }

  const inkTiles = tiles.length;
  const offsets = []; // |per-tile offset - global| for matched tiles
  let misaligned = 0;

  // Pass 2 uses the same tolerant residual per tile to measure local
  // displacement beyond the single page-global registration.
  for (const t of tiles) {
    const { ox, oy, bx, by, discriminated } = t;
    let tileStructNum = 0;
    let tileStructDen = 0;
    for (let ly = 0; ly < T; ly++) {
      for (let lx = 0; lx < T; lx++) {
        const i = (oy + ly) * cw + (ox + lx);
        if (wordMask[i]) {
          tileStructDen++;
          if (!webDilated[i]) tileStructNum++;
        }
        if (webMask[i]) {
          tileStructDen++;
          if (!wordDilated[i]) tileStructNum++;
        }
      }
    }
    const resF = tileStructDen ? tileStructNum / tileStructDen : 0;

    // Alignment/weight/color describe HOW matched content differs, so they
    // are only meaningful where content actually matches (resF small). The
    // tile's own best offset (pass 1) is its true local shift; measured
    // against the global offset it becomes displacement beyond the uniform.
    // A tile that isn't discriminated (pass 1) has no real "own offset" -
    // its bx/by is a scan-order artifact, not evidence of misalignment - so
    // it cannot corroborate a reflow either.
    if (resF >= MATCH_F || !discriminated) continue;

    const mag = Math.sqrt((bx - gDx) * (bx - gDx) + (by - gDy) * (by - gDy));
    offsets.push(mag);
    if (mag > MISALIGN_PX) misaligned++;
  }

  offsets.sort((x, y) => x - y);
  const pctl = (arr, q) => (arr.length ? arr[Math.min(arr.length - 1, Math.floor(q * arr.length))] : 0);
  const alignPx = offsets.length ? pctl(offsets, 0.5) : 0;
  const alignP95 = offsets.length ? pctl(offsets, 0.95) : 0;
  const misalignedPct = offsets.length ? (misaligned / offsets.length) * 100 : 0;
  const structuralPct = structDen ? (structNum / structDen) * 100 : 0;
  // The line-shift channel needs CORROBORATION before it can drive
  // severity: a real reflow always leaves either widespread local
  // misalignment (every historical true flip in parity/history.jsonl
  // measures misalignedPct 56.7-100, including single-heading moves
  // like phase23 p61 at 77.6) or unmatched structural ink. Without
  // either, a "better match N px away" among self-similar ink (equation
  // fraction stacks, repeated digits) is noise - parity-math p1 scored
  // 5.55 lineShift on a page whose raw pixel mismatch is 0.22% and
  // whose glyph positions match the PDF to 0.1pt. Historical spurious
  // cases sit at misalignedPct 21-39; the gate at 48 splits the two
  // populations with margin on both sides.
  const lineShiftCorroborated = misalignedPct >= 48 || structuralPct >= 3;
  const severityPct = Math.max(structuralPct, lineShiftCorroborated ? lineShiftPct : 0);

  // Dominant classification - a deterministic priority chain over robust
  // page-level signals (not the noisier per-tile %s, which stay as reported
  // context). Rationale for each gate:
  //  - A lot of unregisterable content (>= STRUCT_HI) is badly broken:
  //    structural, regardless of anything else.
  //  - A coherent median shift (>= ALIGN_MED CSS px of the *matched* tiles,
  //    beyond the global offset) is drift the eye reads as position, not
  //    missing content: alignment. Median is robust to the ~1px offset-
  //    search jitter that inflates misalignedPct, so it is the gate here.
  //  - Remaining moderate residual (>= STRUCT_LO) is localized/partial
  //    mismatch (e.g. an accumulating advance drift on part of a line):
  //    structural.
  //  - Remaining pages use the explicit appearance targets: full-page and
  //    every available semantic category error must stay below 3.
  const STRUCT_HI = 15; // severity % => structural outright
  const ALIGN_MED = 2; // median matched-tile offset (CSS px) => alignment
  const STRUCT_LO = 10; // residual severity % => structural (localized)
  const APPEARANCE_TARGET = 3;
  const semanticWeightErrors = [
    textWeightErrorPct,
    imageWeightErrorPct,
    tableFillWeightErrorPct,
    tableRuleWeightErrorPct,
  ].filter((value) => value != null);
  const worstSemanticWeight = semanticWeightErrors.length
    ? Math.max(...semanticWeightErrors)
    : 0;
  let driftClass = "clean";
  if (severityPct >= STRUCT_HI) driftClass = "structural";
  else if (alignPx >= ALIGN_MED) driftClass = "alignment";
  else if (severityPct >= STRUCT_LO) driftClass = "structural";
  else if (
    appearanceColorDeltaE >= APPEARANCE_TARGET ||
    (textColorDeltaE ?? 0) >= APPEARANCE_TARGET
  ) driftClass = "color";
  else if (
    appearanceWeightErrorPct >= APPEARANCE_TARGET ||
    worstSemanticWeight >= APPEARANCE_TARGET
  ) driftClass = "weight";

  const gap = 12;
  const out = document.createElement("canvas");
  out.width = w * 3 + gap * 2;
  out.height = h;
  const g = out.getContext("2d");
  g.fillStyle = "#666";
  g.fillRect(0, 0, out.width, out.height);
  g.fillStyle = "#fff";
  g.fillRect(0, 0, w, h);
  g.fillRect(w + gap, 0, w, h);
  if (a) g.drawImage(a, 0, 0);
  if (b) g.drawImage(b, w + gap, 0);
  g.fillStyle = "#fff";
  g.fillRect((w + gap) * 2, 0, w, h);
  const dc = document.createElement("canvas");
  dc.width = w;
  dc.height = h;
  dc.getContext("2d").putImageData(diff, 0, 0);
  g.drawImage(dc, (w + gap) * 2, 0);
  const effectivePageStatus =
    pageStatus === "matched" && dimensionMismatch ? "dimension-mismatch" : pageStatus;
  const unmatched = effectivePageStatus !== "matched";
  return {
    png: out.toDataURL("image/png").split(",")[1],
    mismatchPct: ((mismatch / (w * h)) * 100).toFixed(2),
    lineDebugRows,
    severityPct: unmatched ? "100.00" : severityPct.toFixed(2),
    lineShiftPct: unmatched ? "0.00" : lineShiftPct.toFixed(2),
    alignPx: unmatched ? "0.00" : alignPx.toFixed(2),
    alignP95: unmatched ? "0.00" : alignP95.toFixed(2),
    misalignedPct: unmatched ? "0.00" : misalignedPct.toFixed(2),
    appearanceWeightRatio:
      unmatched || appearanceWeightRatio == null ? null : appearanceWeightRatio.toFixed(4),
    appearanceWeightErrorPct:
      unmatched ? null : appearanceWeightErrorPct.toFixed(2),
    appearanceColorDeltaE: unmatched ? null : appearanceColorDeltaE.toFixed(2),
    textWeightRatio: unmatched ? null : textWeightRatio,
    textWeightErrorPct: unmatched ? null : textWeightErrorPct,
    textWeightMass: unmatched ? 0 : textWeightMass,
    textColorDeltaE: unmatched ? null : textColorDeltaE,
    textColorMass: unmatched ? 0 : textColorMass,
    textColorCoveragePct: unmatched ? null : textColorCoveragePct,
    imageWeightRatio: unmatched ? null : imageWeightRatio,
    imageWeightErrorPct: unmatched ? null : imageWeightErrorPct,
    imageWeightMass: unmatched ? 0 : imageWeightMass,
    imageItemCount: unmatched ? 0 : imageRegions.length,
    tableFillWeightRatio: unmatched ? null : tableFillWeightRatio,
    tableFillWeightErrorPct: unmatched ? null : tableFillWeightErrorPct,
    tableFillWeightMass: unmatched ? 0 : tableFillWeightMass,
    tableRuleWeightRatio: unmatched ? null : tableRuleWeightRatio,
    tableRuleWeightErrorPct: unmatched ? null : tableRuleWeightErrorPct,
    tableRuleWeightMass: unmatched ? 0 : tableRuleWeightMass,
    tableRuleCount: unmatched ? 0 : ruleRegions.length,
    categoryMetricStatus: unmatched ? "unmatched-page" : categoryMetricStatus,
    categoryMetricMs: categoryMetricMs.toFixed(1),
    appearanceMetricMs: appearanceMetricMs.toFixed(1),
    driftClass: unmatched ? "structural" : driftClass,
    pageStatus: effectivePageStatus,
    inkTiles,
  };
};
