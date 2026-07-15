/**
 * buildReport(results, history, meta) -> self-contained HTML string.
 *
 * Pure function: no I/O, no clock reads, no randomness. Everything that varies
 * per run (timestamp, sha) arrives via `meta` so the layout stays deterministic.
 *
 *   results  [{ fixture, page, mismatchPct, severityPct, driftClass, pageStatus,
 *              lineShiftPct, alignPx, alignP95, misalignedPct,
 *              appearanceWeightRatio, appearanceWeightErrorPct,
 *              appearanceColorDeltaE,
 *              textWeightRatio, textWeightErrorPct, textWeightMass,
 *              textColorDeltaE, textColorMass, textColorCoveragePct,
 *              imageWeightRatio, imageWeightErrorPct, imageWeightMass,
 *              imageItemCount, tableFillWeightRatio,
 *              tableFillWeightErrorPct, tableFillWeightMass,
 *              tableRuleWeightRatio, tableRuleWeightErrorPct,
 *              tableRuleWeightMass, tableRuleCount,
 *              categoryMetricMs }]                              this run
 *   history  [{ ts, gitSha, metricVersion, appearanceMetricVersion,
 *               isFullRun, outcome, refreshed,
 *               results:[{...same subset...}] }]
 *            all persisted runs, INCLUDING this run.
 *   meta     { generatedAt, gitSha, base, isFullRun, outcome, label, refreshed,
 *              appearanceMetricVersion }
 *
 * severityPct is the PRIMARY metric everywhere - sorting, bars, KPIs, deltas,
 * trend. It is the STRUCTURAL residual: after registering a single global page
 * offset, the fraction of binary ink that still has no counterpart within a
 * small local tolerance, or ink that matches only after a one-line vertical
 * shift beyond the calibrated noise floor. driftClass names the dominant difference (clean / alignment / weight
 * / color / structural); the other metrics quantify each axis. mismatchPct (raw
 * page-area %) and the sub-metrics are secondary context in the chip, tooltip and
 * table. History entries predating this metric set stay visible as a separate
 * legacy trend segment, but are excluded from deltas/ticks (see comparable)
 * rather than mixing incompatible severities.
 */

const escapeHtml = (s) =>
  String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");

const fmtPct = (n) => `${n.toFixed(2)}%`;
const keyOf = (r) => `${r.fixture}|${r.page}`;

// Page darkness ratio -> signed percent heavier, e.g. 1.235 -> "+24%", null -> "—".
const fmtWeight = (ratio) =>
  ratio == null ? "—" : `${ratio - 1 >= 0 ? "+" : "−"}${Math.abs(Math.round((ratio - 1) * 100))}%`;

const meanWeightError = (results) => {
  const values = results
    .filter((r) => r.appearanceWeightErrorPct != null)
    .map((r) => r.appearanceWeightErrorPct);
  return values.length ? values.reduce((sum, value) => sum + value, 0) / values.length : null;
};

const meanColorError = (results) => {
  const values = results
    .filter((r) => r.appearanceColorDeltaE != null)
    .map((r) => r.appearanceColorDeltaE);
  return values.length ? values.reduce((sum, value) => sum + value, 0) / values.length : null;
};

const weightedMean = (results, valueKey, massKey) => {
  let weighted = 0;
  let mass = 0;
  for (const result of results) {
    const value = result[valueKey];
    const itemMass = result[massKey];
    if (!Number.isFinite(value) || !(itemMass > 0)) continue;
    weighted += value * itemMass;
    mass += itemMass;
  }
  return mass > 0 ? weighted / mass : null;
};

const textColorCoverage = (results) => {
  let colorMass = 0;
  let textWeightMass = 0;
  for (const result of results) {
    if (!(result.textWeightMass > 0) || !Number.isFinite(result.textColorMass)) continue;
    colorMass += result.textColorMass;
    textWeightMass += result.textWeightMass;
  }
  return textWeightMass > 0 ? (200 * colorMass) / textWeightMass : null;
};

const meanCategoryMetricMs = (results) => {
  const values = results
    .filter((r) => Number.isFinite(r.categoryMetricMs))
    .map((r) => r.categoryMetricMs);
  return values.length ? values.reduce((sum, value) => sum + value, 0) / values.length : null;
};

const semanticMeans = (results) => ({
  textWeight: weightedMean(results, "textWeightErrorPct", "textWeightMass"),
  textColor: weightedMean(results, "textColorDeltaE", "textColorMass"),
  textCoverage: textColorCoverage(results),
  imageWeight: weightedMean(results, "imageWeightErrorPct", "imageWeightMass"),
  tableFillWeight: weightedMean(results, "tableFillWeightErrorPct", "tableFillWeightMass"),
  tableRuleWeight: weightedMean(results, "tableRuleWeightErrorPct", "tableRuleWeightMass"),
  categoryMetricMs: meanCategoryMetricMs(results),
});

// Smallest "nice" number (1/2/2.5/5 x 10^k) >= raw, for readable axis steps.
function niceStep(raw) {
  if (!(raw > 0)) return 1;
  const mag = Math.pow(10, Math.floor(Math.log10(raw)));
  const norm = raw / mag;
  const nice = norm <= 1 ? 1 : norm <= 2 ? 2 : norm <= 2.5 ? 2.5 : norm <= 5 ? 5 : 10;
  return nice * mag;
}

const DRIFT_CLASSES = ["clean", "alignment", "weight", "color", "structural"];
export const METRIC_VERSION = "ink-dilate-line-v5";
export const APPEARANCE_METRIC_VERSION = "semantic-rgba-v1";

// ─── Category taxonomy ───────────────────────────────────────────────────────
// Curated fixture → category map. Every fixture in parity/out/results.json is
// classified by name + domain knowledge so the eval reads as capability areas
// (what WordInWeb can render well) instead of an undifferentiated 1000-row list.
// One line per fixture. Any fixture NOT listed here falls into the visible
// "uncategorized" bucket via categoryOf() — never silently dropped — so a new
// fixture surfaces immediately and prompts a classification.
const FIXTURE_CATEGORY = {
  // General word processing — body text, notes, fields, revisions, TOC, forms
  benchmark: "general",
  chronology: "general",
  msa: "general",
  "parity-comments": "general",
  "parity-revisions": "general",
  "parity-text": "general",
  "parity2-fields": "general",
  "parity2-notes": "general",
  "parity2-toc": "general",
  pickett: "general",
  pleading: "general",
  "probe2-content-controls": "general",
  "probe2-form-checkboxes": "general",
  "probe3-field-switches": "general",
  "probe3-index-xrefs": "general",
  "probe3-lo-provenance": "general",
  "probe3-tracked-changes": "general",
  sample: "general",
  "staging-fields2": "general",
  // Complex tables — grids, nesting, row-splitting, extreme/long tables
  "parity-rowsplit": "tables",
  "parity-tables": "tables",
  "parity2-nestedtables": "tables",
  "probe-nih-rowheight": "tables",
  "probe3-table-exotics": "tables",
  "staging-longtable": "tables",
  "staging-tblextreme": "tables",
  // Math & equations — OMML, matrices, inline/display equations
  "parity-math": "math",
  "parity-math2": "math",
  "parity2-equations": "math",
  "probe2-math-matrices": "math",
  // Other languages & scripts — RTL, Indic, CJK, Thai, bidi, ruby, vertical
  "probe2-arabic-rtl": "languages",
  "probe2-ruby-vertical": "languages",
  "probe3-chargrid": "languages",
  "probe3-indic": "languages",
  "probe3-kashida": "languages",
  "probe3-thai": "languages",
  "staging-bidi": "languages",
  "staging-eastasian": "languages",
  // Formatting & layout — columns, sections, breaks, tabs, borders, styles, frames
  "parity-colbalance": "formatting",
  "parity-columns": "formatting",
  "parity-dividers": "formatting",
  "parity-firstpage": "formatting",
  "parity-headerfooter": "formatting",
  "parity-lists": "formatting",
  "parity-pageborders": "formatting",
  "parity-wrapmodes": "formatting",
  "parity2-charstyles": "formatting",
  "parity2-coverpage": "formatting",
  "parity2-dropcap": "formatting",
  "parity2-lists": "formatting",
  "parity2-sections": "formatting",
  "parity2-tabs": "formatting",
  "parity2-textboxes": "formatting",
  "probe2-dropcaps-frames": "formatting",
  "probe2-hyphenation": "formatting",
  "probe2-mixed-orientation": "formatting",
  "probe2-modern-template": "formatting",
  "probe2-run-borders": "formatting",
  "probe2-styleref-headers": "formatting",
  "probe3-columns-unequal": "formatting",
  "probe3-linked-textboxes": "formatting",
  "probe3-mirror-book": "formatting",
  "probe3-text-effects": "formatting",
  "staging-anchors2": "formatting",
  "staging-breaks": "formatting",
  "staging-frames": "formatting",
  "staging-grid4": "formatting",
  "staging-hf2": "formatting",
  "staging-styles": "formatting",
  "staging-typography": "formatting",
  // Graphics & media — pictures, watermarks, WordArt, shapes, emoji glyphs
  "parity-pictures": "graphics",
  "parity2-watermark": "graphics",
  "probe2-picture-watermark": "graphics",
  "probe3-emoji": "graphics",
  "probe3-shape-autofit": "graphics",
  "probe3-wordart-warps": "graphics",
  // Real-world documents — the wild-* / wild2-* corpus of authored documents
  "wild-athabasca": "realworld",
  "wild-doerfp": "realworld",
  "wild-gatech": "realworld",
  "wild-hamburg": "realworld",
  "wild-multicolumn": "realworld",
  "wild-wirfp": "realworld",
  "wild2-legal-ca-agreement": "realworld",
  "wild2-legal-nih-contract": "realworld",
  "wild2-lit-yiddish-rtl": "realworld",
  "wild2-math-eq-as-images": "realworld",
  "wild2-math-omml-dense": "realworld",
  "wild2-med-nccih-protocol": "realworld",
  "wild2-med-phase23-protocol": "realworld",
  "wild2-sci-chem-omml": "realworld",
  "wild2-sci-elsevier-template": "realworld",
  "wild2-sci-ieee-2col": "realworld",
};

// Category metadata. `accent` drives the per-category colour so the report reads
// as one designed system (no emoji, theme-safe). `blurb` is the write-up line.
// Order here is the fallback display order; cards/sections re-sort worst-first.
const CATEGORIES = [
  {
    id: "general",
    label: "General word processing",
    accent: "#2a78d6",
    blurb:
      "Everyday document content — body text, paragraphs, footnotes and endnotes, fields, tables of contents, tracked changes, comments and form controls.",
  },
  {
    id: "tables",
    label: "Complex tables",
    accent: "#1baf7a",
    blurb:
      "Table structure and rendering — nested tables, row splitting across pages, computed row heights, and extreme or very long tables.",
  },
  {
    id: "math",
    label: "Math & equations",
    accent: "#7c5ce0",
    blurb:
      "OfficeMath (OMML) — inline and display equations, matrices, fractions and stretchy delimiters.",
  },
  {
    id: "languages",
    label: "Other languages & scripts",
    accent: "#e0803c",
    blurb:
      "Non-Latin and complex scripts — Arabic and kashida justification, Hebrew and bidi, Indic shaping, Thai, CJK grids, ruby and vertical text.",
  },
  {
    id: "formatting",
    label: "Formatting & layout",
    accent: "#2ba3c7",
    blurb:
      "Page and character layout — columns, sections and breaks, tabs, drop caps, borders, styles, frames, orientation, mirror margins and typography.",
  },
  {
    id: "graphics",
    label: "Graphics & media",
    accent: "#d05a9e",
    blurb:
      "Embedded and drawn objects — pictures, watermarks, WordArt and text warps, autofit shapes and colour emoji glyphs.",
  },
  {
    id: "realworld",
    label: "Real-world documents",
    accent: "#5a6b8c",
    blurb:
      "Whole documents pulled from the wild — legal agreements, government RFPs, scientific and medical templates, and multi-column journal layouts — exercising every feature at once.",
  },
  {
    id: "uncategorized",
    label: "Uncategorized",
    accent: "#8a8a86",
    blurb:
      "Fixtures not yet mapped to a category. New fixtures land here until they are classified in FIXTURE_CATEGORY.",
  },
];

const CATEGORY_BY_ID = new Map(CATEGORIES.map((c) => [c.id, c]));
const categoryOf = (fixture) => {
  const id = FIXTURE_CATEGORY[fixture];
  return id && CATEGORY_BY_ID.has(id) ? id : "uncategorized";
};

// Colour band for a mean-severity value (percent). Tuned to the corpus: most
// categories sit well under 0.5%; a category over ~1.5% mean is dominated by a
// known floor (e.g. complex-script rasterization) and reads amber/red.
const sevClass = (v) => (v == null ? "" : v < 0.5 ? "sev-good" : v < 1.5 ? "sev-warn" : "sev-bad");

// Per-category subscores over one provenance's results. Empty categories are
// dropped; the rest are returned worst-first by mean severity so problem areas
// surface at the top of the card grid and section list.
function categoryStats(results, prev) {
  const buckets = new Map();
  for (const r of results) {
    const id = categoryOf(r.fixture);
    if (!buckets.has(id)) buckets.set(id, []);
    buckets.get(id).push(r);
  }
  const prevMean = new Map();
  const prevRows = comparable(prev);
  if (prevRows.length) {
    const pb = new Map();
    for (const r of prevRows) {
      const id = categoryOf(r.fixture);
      if (!pb.has(id)) pb.set(id, []);
      pb.get(id).push(r.severityPct);
    }
    for (const [id, vals] of pb) prevMean.set(id, vals.reduce((a, b) => a + b, 0) / vals.length);
  }
  const out = [];
  for (const cat of CATEGORIES) {
    const rows = buckets.get(cat.id);
    if (!rows || rows.length === 0) continue;
    const vals = rows.map((r) => r.severityPct);
    const mean = vals.reduce((a, b) => a + b, 0) / vals.length;
    const worst = rows.reduce((a, b) => (b.severityPct > a.severityPct ? b : a), rows[0]);
    out.push({
      ...cat,
      rows,
      pages: rows.length,
      fixtures: new Set(rows.map((r) => r.fixture)).size,
      mean,
      worst,
      over1: rows.filter((r) => r.severityPct >= 1).length,
      clean: rows.filter((r) => r.driftClass === "clean").length,
      prevMean: prevMean.has(cat.id) ? prevMean.get(cat.id) : null,
    });
  }
  out.sort((a, b) => b.mean - a.mean);
  return out;
}

// Chip label: the class word plus the one number that names it (median offset for
// alignment, signed weight for weight). The colored dot is decorative - the word
// carries the meaning, so the chip is never color-only.
function chipLabel(r) {
  const cls = DRIFT_CLASSES.includes(r.driftClass) ? r.driftClass : "clean";
  if (r.pageStatus === "missing-web") return "missing web page";
  if (r.pageStatus === "extra-web") return "extra web page";
  if (r.pageStatus === "dimension-mismatch") return "page size mismatch";
  if (cls === "alignment") return `alignment ${(r.alignPx ?? 0).toFixed(1)}px`;
  if (cls === "weight") return `weight ${fmtWeight(r.appearanceWeightRatio)}`;
  return cls;
}
const chipClass = (r) => (DRIFT_CLASSES.includes(r.driftClass) ? r.driftClass : "clean");
/** Comparable results from a history entry: only those carrying THIS metric set
 * (severityPct is the structural residual and driftClass is present), so runs
 * recorded under an older/incompatible severity definition never feed deltas
 * or previous-run ticks. The trend displays them as a separate segment. */
const comparable = (entry) =>
  entry && entry.metricVersion === METRIC_VERSION && Array.isArray(entry.results)
    ? entry.results.filter((r) => r.severityPct != null && r.driftClass != null)
    : [];

const comparableAppearance = (entry) =>
  entry?.appearanceMetricVersion === APPEARANCE_METRIC_VERSION && Array.isArray(entry.results)
    ? entry.results
    : [];

const historyResults = (entry) =>
  entry && Array.isArray(entry.results)
    ? entry.results.filter((r) => r.severityPct != null && r.driftClass != null)
    : [];

const isFullHistoryRun = (entry) => entry?.isFullRun !== false;
const metricName = (entry) => entry?.metricVersion ?? "legacy (unversioned)";

// The previous full run is what deltas and the "previous run" ticks compare
// against: the run before this one on a full run, else the latest full run.
// Only entries with a real results array count, so a directly-passed history
// carrying schema-partial lines can't feed undefined into the delta math.
function previousRun(history, isFullRun) {
  const valid = history.filter(
    (h) => isFullHistoryRun(h) && comparable(h).length > 0,
  );
  const idx = isFullRun ? valid.length - 2 : valid.length - 1;
  return valid[idx] || null;
}

function previousAppearanceRun(history, isFullRun) {
  const valid = history.filter(
    (h) => isFullHistoryRun(h) && comparableAppearance(h).length > 0,
  );
  const idx = isFullRun ? valid.length - 2 : valid.length - 1;
  return valid[idx] || null;
}

function deltaLine(diff, { goodWhenNegative, unit, decimals }) {
  if (Math.abs(diff) < 0.5 * Math.pow(10, -decimals)) {
    return `<div class="delta neutral">no change</div>`;
  }
  const negative = diff < 0;
  const good = goodWhenNegative ? negative : !negative;
  const arrow = negative ? "▼" : "▲";
  const cls = good ? "good" : "critical";
  const suffix = unit ? ` ${unit}` : "";
  return `<div class="delta ${cls}">${arrow} ${Math.abs(diff).toFixed(decimals)}${suffix}</div>`;
}

function buildKpis(results, prev) {
  const tile = (label, value, delta) =>
    `<div class="kpi"><div class="kpi-label">${label}</div>` +
    `<div class="kpi-value">${value}</div>${delta || ""}</div>`;

  if (results.length === 0) {
    return (
      `<section class="kpis">` +
      tile("Mean structural", "—", "") +
      tile("Worst page", "—", "") +
      tile("Pages clean", "0 / 0", "") +
      tile("Fixtures", "0", "") +
      tile("Mean page-weight error (<3%)", "—", "") +
      tile("Mean source-colour ΔE00 (<3)", "—", "") +
      `</section>`
    );
  }

  const vals = results.map((r) => r.severityPct);
  const n = vals.length;
  const mean = vals.reduce((a, b) => a + b, 0) / n;
  const worst = results.reduce((a, b) => (b.severityPct > a.severityPct ? b : a), results[0]);
  // Provenance-aware worst: fixtures authored by other suites (LibreOffice)
  // are tracked as a separate drift axis and must not bury the Word-authored
  // headline (the user-facing parity target is Word's render of Word files).
  const wordResults = results.filter((r) => (r.provenance ?? "word") === "word");
  const worstWord = wordResults.length
    ? wordResults.reduce((a, b) => (b.severityPct > a.severityPct ? b : a), wordResults[0])
    : null;
  const nonWord = results.filter((r) => (r.provenance ?? "word") !== "word");
  const worstNonWord = nonWord.length
    ? nonWord.reduce((a, b) => (b.severityPct > a.severityPct ? b : a), nonWord[0])
    : null;
  const cleanCount = results.filter((r) => r.driftClass === "clean").length;
  const fixtures = new Set(results.map((r) => r.fixture)).size;
  const weightMean = meanWeightError(results);
  const colorMean = meanColorError(results);

  let pMean = null;
  let pWorst = null;
  let pClean = null;
  let pWeightMean = null;
  let pColorMean = null;
  const pRes = comparable(prev);
  if (pRes.length) {
    const pv = pRes.map((r) => r.severityPct);
    pMean = pv.reduce((a, b) => a + b, 0) / pv.length;
    pWorst = Math.max(...pv);
    pClean = pRes.filter((r) => r.driftClass === "clean").length;
    pWeightMean = meanWeightError(pRes);
    pColorMean = meanColorError(pRes);
  }

  return (
    `<section class="kpis">` +
    tile(
      "Mean structural",
      fmtPct(mean),
      pMean == null ? "" : deltaLine(mean - pMean, { goodWhenNegative: true, unit: "pp", decimals: 2 }),
    ) +
    tile(
      "Worst page (Word-authored)",
      worstWord
        ? `<span class="worst-name">${escapeHtml(worstWord.fixture)} p${worstWord.page}</span> — ${fmtPct(worstWord.severityPct)}`
        : "n/a",
      pWorst == null ? "" : deltaLine(worst.severityPct - pWorst, { goodWhenNegative: true, unit: "pp", decimals: 2 }),
    ) +
    (worstNonWord
      ? tile(
          `Worst page (${escapeHtml(worstNonWord.provenance ?? "other")})`,
          `<span class="worst-name">${escapeHtml(worstNonWord.fixture)} p${worstNonWord.page}</span> — ${fmtPct(worstNonWord.severityPct)}`,
          "deferred axis",
        )
      : "") +
    tile(
      "Pages clean",
      `${cleanCount} / ${n}`,
      pClean == null ? "" : deltaLine(cleanCount - pClean, { goodWhenNegative: false, unit: "", decimals: 0 }),
    ) +
    tile("Fixtures", String(fixtures), "") +
    tile(
      "Mean page-weight error (<3%)",
      weightMean == null ? "—" : fmtPct(weightMean),
      weightMean == null || pWeightMean == null
        ? ""
        : deltaLine(weightMean - pWeightMean, { goodWhenNegative: true, unit: "pp", decimals: 2 }),
    ) +
    tile(
      "Mean source-colour ΔE00 (<3)",
      colorMean == null ? "—" : colorMean.toFixed(2),
      colorMean == null || pColorMean == null
        ? ""
        : deltaLine(colorMean - pColorMean, { goodWhenNegative: true, unit: "ΔE", decimals: 2 }),
    ) +
    `</section>`
  );
}

function buildSemanticKpis(results, prev, appearanceMetricVersion) {
  const compatible = appearanceMetricVersion === APPEARANCE_METRIC_VERSION;
  const measuredPages = results.filter(
    (result) => Number.isFinite(result.categoryMetricMs) && result.categoryMetricStatus === "ok",
  ).length;
  const invalidPages = results.filter((result) => result.pageStatus !== "matched").length;
  const invalidSemanticPages = results.filter(
    (result) => result.categoryMetricStatus !== "ok",
  ).length;
  const complete =
    compatible && results.length > 0 && measuredPages === results.length && invalidPages === 0;
  const values = compatible ? semanticMeans(results) : semanticMeans([]);
  const previous = semanticMeans(comparableAppearance(prev));
  const tile = (label, value, target, previousValue, formatter, higherIsBetter = false) => {
    let renderedValue;
    let delta = "";
    if (!compatible) {
      renderedValue = `<span class="unavailable">unavailable</span>`;
    } else if (value == null) {
      renderedValue = `<span class="unavailable">not present</span>`;
    } else {
      const passes = higherIsBetter ? value >= target : value < target;
      renderedValue =
        `${formatter(value)} ` +
        (complete
          ? `<span class="gate ${passes ? "good" : "critical"}">${passes ? "pass" : "fail"}</span>`
          : `<span class="gate unavailable">incomplete</span>`);
      if (complete && previousValue != null) {
        delta = deltaLine(value - previousValue, {
          goodWhenNegative: !higherIsBetter,
          unit: higherIsBetter ? "pp" : label.includes("colour") ? "ΔE" : "pp",
          decimals: 2,
        });
      }
    }
    return (
      `<div class="kpi semantic-kpi"><div class="kpi-label">${label}</div>` +
      `<div class="kpi-value">${renderedValue}</div>${delta}</div>`
    );
  };
  const pct = (value) => fmtPct(value);
  const deltaE = (value) => value.toFixed(2);
  const note = compatible
    ? `Corpus means are weighted by semantic category mass. ${measuredPages} / ${results.length} pages have valid semantic measurements${invalidPages ? `; ${invalidPages} unmatched page${invalidPages === 1 ? "" : "s"}` : ""}${invalidSemanticPages ? `; ${invalidSemanticPages} invalid semantic page${invalidSemanticPages === 1 ? "" : "s"}` : ""}${invalidPages || invalidSemanticPages ? " prevent a pass" : ""}. All error gates require &lt;3; text-colour coverage requires ≥95%.`
    : `Semantic appearance is unavailable because this run was not recorded as ${escapeHtml(APPEARANCE_METRIC_VERSION)}.`;

  return (
    `<section class="semantic-block"><div class="chart-head"><h2>Semantic appearance gates</h2>` +
    `<div class="metric-version">${compatible ? escapeHtml(APPEARANCE_METRIC_VERSION) : "unavailable"}</div></div>` +
    `<p class="metric-note">${note}</p><div class="kpis semantic-kpis">` +
    tile("Text weight (<3%)", values.textWeight, 3, previous.textWeight, pct) +
    tile("Text colour ΔE00 (<3)", values.textColor, 3, previous.textColor, deltaE) +
    tile("Text colour coverage (≥95%)", values.textCoverage, 95, previous.textCoverage, pct, true) +
    tile("Image weight (<3%)", values.imageWeight, 3, previous.imageWeight, pct) +
    tile("Table-fill weight (<3%)", values.tableFillWeight, 3, previous.tableFillWeight, pct) +
    tile("Table-rule weight (<3%)", values.tableRuleWeight, 3, previous.tableRuleWeight, pct) +
    `</div></section>`
  );
}

// Order: fixtures by their worst page (worst-first); pages within a fixture by
// mismatch (worst-first). Keeps same-fixture pages adjacent so the grouping gaps
// are meaningful while the overall read is still worst-first.
function orderRows(results) {
  const byFixture = new Map();
  for (const r of results) {
    if (!byFixture.has(r.fixture)) byFixture.set(r.fixture, []);
    byFixture.get(r.fixture).push(r);
  }
  const groups = [...byFixture.entries()].map(([fixture, rows]) => {
    rows.sort((a, b) => b.severityPct - a.severityPct);
    return { fixture, rows, worst: rows[0].severityPct };
  });
  groups.sort((a, b) => b.worst - a.worst);
  return groups;
}

// Renders just the horizontal-bar chart (gridlines + rows + axis) for a set of
// results, scoped to whatever slice is passed (whole corpus or one category).
// Returns { chart, hasPrev }. buildBars wraps it with the section chrome; the
// per-category sections drop it in directly so the drill-down (hover tooltip +
// click-to-open-diff) is identical everywhere.
function renderBarChart(results, prev) {
  const groups = orderRows(results);
  const prevMap = new Map();
  for (const r of comparable(prev)) prevMap.set(keyOf(r), r.severityPct);

  const dataMax = results.reduce((m, r) => Math.max(m, r.severityPct), 0);
  const xmax = Math.max(5, Math.ceil(dataMax));
  // "Nice" step targeting ~8 gridlines so the axis stays legible whether the
  // worst page is 5% or 98% (a fixed 2.5% step would print ~40 labels at 98%).
  const step = niceStep(xmax / 8);
  const ticks = [];
  for (let t = 0; t <= xmax + 1e-9; t += step) ticks.push(t);

  const gridlines = ticks
    .map((t) => `<div class="gridline" style="left:${(t / xmax) * 100}%"></div>`)
    .join("");
  const axis = ticks
    .map((t) => `<div class="axis-tick" style="left:${(t / xmax) * 100}%">${+t.toFixed(2)}%</div>`)
    .join("");

  let rowsHtml = "";
  groups.forEach((g, gi) => {
    g.rows.forEach((r, ri) => {
      const gap = ri === 0 ? (gi === 0 ? 0 : 12) : 4;
      const pct = (Math.min(r.severityPct, xmax) / xmax) * 100;
      const hasPrev = prevMap.has(keyOf(r));
      const tick = hasPrev
        ? `<div class="prev-tick" style="left:${(Math.min(prevMap.get(keyOf(r)), xmax) / xmax) * 100}%"></div>`
        : "";
      const png = r.pngRel ?? `${r.fixture}-p${r.page}.png`;
      const cls = chipClass(r);
      const attr = (k, v) => (v == null ? "" : `data-${k}="${v}" `);
      rowsHtml +=
        `<div class="row" style="margin-top:${gap}px" ` +
        `data-fixture="${escapeHtml(r.fixture)}" data-page="${r.page}" ` +
        `data-val="${r.severityPct.toFixed(2)}" ` +
        `data-raw="${r.mismatchPct.toFixed(2)}" ` +
        `data-class="${cls}" ` +
        attr("line", r.lineShiftPct != null ? r.lineShiftPct.toFixed(1) : null) +
        attr("align", r.alignPx != null ? r.alignPx.toFixed(1) : null) +
        attr("alignp95", r.alignP95 != null ? r.alignP95.toFixed(1) : null) +
        attr("mis", r.misalignedPct != null ? r.misalignedPct.toFixed(0) : null) +
        attr("weight", r.appearanceWeightRatio != null ? r.appearanceWeightRatio.toFixed(4) : null) +
        attr("weighterror", r.appearanceWeightErrorPct != null ? r.appearanceWeightErrorPct.toFixed(2) : null) +
        attr("color", r.appearanceColorDeltaE != null ? r.appearanceColorDeltaE.toFixed(2) : null) +
        `data-prev="${hasPrev ? prevMap.get(keyOf(r)).toFixed(2) : ""}" ` +
        `data-png="${escapeHtml(png)}">` +
        `<div class="rlabel"><span class="chip ${cls}">${escapeHtml(chipLabel(r))}</span>` +
        `<span class="fxwrap"><span class="fx">${escapeHtml(r.fixture)}</span> · p${r.page}</span></div>` +
        `<div class="track">${tick}` +
        `<div class="bar" style="width:${pct}%"></div>` +
        `<div class="vlabel" style="left:${pct}%">${fmtPct(r.severityPct)}</div>` +
        `</div></div>`;
    });
  });

  const chart =
    `<div class="chart"><div class="rows">${gridlines}${rowsHtml}</div>` +
    `<div class="axis">${axis}</div></div>`;
  return { chart, hasPrev: prevMap.size > 0 };
}

const driftKeyHtml = () =>
  DRIFT_CLASSES.map(
    (c) => `<span class="lg"><span class="chip-dot ${c}"></span>${c}</span>`,
  ).join("");

// Grid of category subscore cards — the visual centrepiece. Each card links to
// its detail section and its worst page's diff image. Sorted worst-first.
function buildCategoryCards(results, prev) {
  const stats = categoryStats(results, prev);
  if (stats.length === 0) return "";
  const cards = stats
    .map((c) => {
      const worstPng = c.worst.pngRel ?? `${c.worst.fixture}-p${c.worst.page}.png`;
      let delta = "";
      if (c.prevMean != null) {
        delta = deltaLine(c.mean - c.prevMean, { goodWhenNegative: true, unit: "pp", decimals: 2 });
      }
      return (
        `<div class="cat-card" style="--accent:${c.accent}">` +
        `<a class="cat-card-hd" href="#cat-${c.id}">` +
        `<span class="cat-dot"></span><span class="cat-name">${escapeHtml(c.label)}</span></a>` +
        `<div class="cat-meta">${c.pages} page${c.pages === 1 ? "" : "s"} · ${c.fixtures} fixture${c.fixtures === 1 ? "" : "s"}</div>` +
        `<div class="cat-mean"><span class="cat-mean-val ${sevClass(c.mean)}">${fmtPct(c.mean)}</span>` +
        `<span class="cat-mean-lbl">mean structural</span>${delta}</div>` +
        `<div class="cat-stats">` +
        `<span><b>${c.over1}</b> page${c.over1 === 1 ? "" : "s"} ≥ 1%</span>` +
        `<span><b>${c.clean}</b>/${c.pages} clean</span></div>` +
        `<div class="cat-worst">worst <a href="${escapeHtml(worstPng)}" target="_blank" rel="noopener">` +
        `${escapeHtml(c.worst.fixture)} p${c.worst.page}</a> — ${fmtPct(c.worst.severityPct)}</div>` +
        `<p class="cat-blurb">${escapeHtml(c.blurb)}</p>` +
        `</div>`
      );
    })
    .join("");
  return (
    `<section class="cat-cards-block">` +
    `<div class="chart-head"><h2>Scores by category</h2></div>` +
    `<p class="metric-note">Every page is classified into a capability area. Each card shows the mean ` +
    `structural score, worst page (linked to its diff image), and how many pages exceed 1%. ` +
    `Cards are ordered worst-first. Expand a category below to see every page.</p>` +
    `<div class="cat-cards">${cards}</div></section>`
  );
}

// Per-category detail sections: collapsible, each holding the same bar chart
// (with hover tooltip + click-to-diff drill-down) and a compact page table with
// diff links. The worst category is open by default so the drill-down is visible.
function buildCategorySections(results, prev) {
  const stats = categoryStats(results, prev);
  if (stats.length === 0) return "";
  const sections = stats
    .map((c, i) => {
      const { chart } = renderBarChart(c.rows, prev);
      const table = buildCategoryTable(c.rows, prev);
      return (
        `<details class="cat-section" id="cat-${c.id}" style="--accent:${c.accent}"${i === 0 ? " open" : ""}>` +
        `<summary><span class="cat-dot"></span><span class="cat-name">${escapeHtml(c.label)}</span>` +
        `<span class="cat-sum-meta">${c.pages} pages · mean <b class="${sevClass(c.mean)}">${fmtPct(c.mean)}</b> · worst ${fmtPct(c.worst.severityPct)}</span></summary>` +
        `<p class="cat-blurb">${escapeHtml(c.blurb)}</p>` +
        `<div class="drift-key">${driftKeyHtml()}</div>` +
        chart +
        table +
        `</details>`
      );
    })
    .join("");
  return (
    `<section class="cat-sections-block">` +
    `<div class="chart-head"><h2>Pages by category</h2></div>` +
    `<p class="metric-note">Severity = <b>structural residual</b>: unmatched binary ink after global page ` +
    `registration, or ink that only matches after an 8–24px vertical shift (5% noise floor removed). ` +
    `Alignment-, weight- and colour-only differences are excluded and named by the chip instead. ` +
    `Hover a bar for the full metric breakdown; click it to open the page's diff image.</p>` +
    sections +
    `</section>`
  );
}

// Compact per-category page table: fixture, page, structural, drift, Δ, diff link.
function buildCategoryTable(results, prev) {
  const prevMap = new Map();
  for (const r of comparable(prev)) prevMap.set(keyOf(r), r.severityPct);
  const rows = [...results].sort((a, b) => b.severityPct - a.severityPct);
  const body = rows
    .map((r) => {
      const png = r.pngRel ?? `${r.fixture}-p${r.page}.png`;
      let dcell = "—";
      if (prevMap.has(keyOf(r))) {
        const d = r.severityPct - prevMap.get(keyOf(r));
        if (Math.abs(d) < 0.005) dcell = "no change";
        else {
          const arrow = d < 0 ? "▼" : "▲";
          const cls = d < 0 ? "good" : "critical";
          dcell = `<span class="${cls}">${arrow} ${Math.abs(d).toFixed(2)} pp</span>`;
        }
      }
      return (
        `<tr><td>${escapeHtml(r.fixture)}</td><td>p${r.page}</td>` +
        `<td class="num">${fmtPct(r.severityPct)}</td>` +
        `<td><span class="chip ${chipClass(r)}">${escapeHtml(r.driftClass || "clean")}</span></td>` +
        `<td>${dcell}</td>` +
        `<td><a href="${escapeHtml(png)}" target="_blank" rel="noopener">diff</a></td></tr>`
      );
    })
    .join("");
  return (
    `<details class="cat-table"><summary>${rows.length} pages as a table</summary>` +
    `<div class="table-scroll"><table class="ptable"><thead><tr><th>Fixture</th><th>Page</th>` +
    `<th class="num">Structural</th><th>Drift</th><th>Δ vs previous</th><th>Diff</th></tr></thead>` +
    `<tbody>${body}</tbody></table></div></details>`
  );
}


function buildTrend(history, provenanceLabel) {
  // Tabbed reports: the trend respects the tab's provenance so the deferred
  // LibreOffice fixture's 57% doesn't ride the Word tab's worst-line
  // (results predating the provenance column default to word).
  const provFilter = (res) =>
    provenanceLabel ? res.filter((r) => (r.provenance ?? "word") === provenanceLabel) : res;
  const runs = history
    .filter(isFullHistoryRun)
    .map((h) => ({ h, res: provFilter(historyResults(h)) }))
    .filter(({ res }) => res.length > 0)
    .slice(-30)
    .map(({ h, res }, index) => {
      const vals = res.map((r) => r.severityPct);
      return {
        index,
        sha: h.gitSha || "?",
        metric: metricName(h),
        mean: vals.reduce((a, b) => a + b, 0) / vals.length,
        worst: Math.max(...vals),
      };
    });
  if (runs.length === 0) return "";

  const W = 1000;
  const H = 180;
  const padL = 40;
  const padR = 64;
  const padT = 16;
  const padB = 28;
  const ymax = Math.max(5, Math.ceil(Math.max(...runs.map((r) => r.worst))));
  const x = (i) => padL + (i / Math.max(1, runs.length - 1)) * (W - padL - padR);
  const y = (v) => padT + (1 - v / ymax) * (H - padT - padB);

  const groups = new Map();
  for (const run of runs) {
    const group = groups.get(run.metric) ?? [];
    group.push(run);
    groups.set(run.metric, group);
  }
  const path = (group, sel) =>
    group
      .map((r, i) => `${i === 0 ? "M" : "L"}${x(r.index).toFixed(1)},${y(sel(r)).toFixed(1)}`)
      .join(" ");
  const paths = (sel, cls) =>
    [...groups.entries()]
      .filter(([, group]) => group.length > 1)
      .map(([metric, group]) =>
        `<path d="${path(group, sel)}" class="t-line ${cls}${metric === METRIC_VERSION ? "" : " legacy"}"/>`,
      )
      .join("");

  const yTicks = [];
  const ystep = ymax <= 6 ? 1 : Math.ceil(ymax / 5);
  for (let t = 0; t <= ymax + 1e-9; t += ystep) yTicks.push(t);
  const yGrid = yTicks
    .map(
      (t) =>
        `<line x1="${padL}" y1="${y(t)}" x2="${W - padR}" y2="${y(t)}" class="t-grid"/>` +
        `<text x="${padL - 6}" y="${y(t) + 3}" class="t-ylabel">${t}%</text>`,
    )
    .join("");
  const xLabels = runs
    .map((r) => `<text x="${x(r.index)}" y="${H - 8}" class="t-xlabel">${escapeHtml(r.sha)}</text>`)
    .join("");

  const last = runs[runs.length - 1];
  const endLabels =
    `<text x="${x(last.index) + 6}" y="${y(last.mean) + 3}" class="t-end s1">${fmtPct(last.mean)}</text>` +
    `<text x="${x(last.index) + 6}" y="${y(last.worst) + 3}" class="t-end s2">${fmtPct(last.worst)}</text>`;

  const dots = (sel, cls) =>
    runs
      .map((r) =>
        `<circle cx="${x(r.index)}" cy="${y(sel(r))}" r="3" class="${cls}${r.metric === METRIC_VERSION ? "" : " legacy"}"/>`,
      )
      .join("");

  const trendData = JSON.stringify(
    runs.map((r) => ({
      x: x(r.index),
      sha: r.sha,
      metric: r.metric,
      mean: r.mean,
      worst: r.worst,
      my: y(r.mean),
      wy: y(r.worst),
    })),
  );

  return (
    `<section class="trend-block">` +
    `<div class="chart-head"><h2>Trend across runs</h2><div class="legend">` +
    `<span class="lg"><span class="lg-line s1"></span>mean</span>` +
    `<span class="lg"><span class="lg-line s2"></span>worst page</span>` +
    `<span class="lg"><span class="lg-line legacy"></span>legacy metric</span></div></div>` +
    `<p class="metric-note">Metric changes are shown as disconnected segments. Legacy runs remain visible, ` +
    `but only ${escapeHtml(METRIC_VERSION)} runs are used for current deltas.</p>` +
    `<div class="trend-wrap"><svg viewBox="0 0 ${W} ${H}" class="trend" data-runs='${trendData}' data-w="${W}" data-padt="${padT}" data-h="${H}" data-padb="${padB}">` +
    yGrid +
    paths((r) => r.worst, "s2") +
    paths((r) => r.mean, "s1") +
    dots((r) => r.worst, "t-dot s2") +
    dots((r) => r.mean, "t-dot s1") +
    xLabels +
    endLabels +
    `<line class="t-cross" x1="0" y1="${padT}" x2="0" y2="${H - padB}" style="display:none"/>` +
    `</svg><div class="trend-tip" style="display:none"></div></div>` +
    `</section>`
  );
}

function buildRunHistory(history) {
  const runs = history
    .filter((h) => historyResults(h).length > 0)
    .slice(-50)
    .reverse();
  if (runs.length === 0) return "";

  const rows = runs
    .map((run) => {
      const results = historyResults(run);
      const values = results.map((r) => r.severityPct);
      const mean = values.reduce((sum, value) => sum + value, 0) / values.length;
      const appearanceResults = comparableAppearance(run);
      const appearanceAvailable = appearanceResults.length > 0;
      const semantic = semanticMeans(appearanceResults);
      const pageWeight = appearanceAvailable ? meanWeightError(appearanceResults) : null;
      const pageColor = appearanceAvailable ? meanColorError(appearanceResults) : null;
      const imageItems = appearanceResults.reduce(
        (sum, result) => sum + (Number.isFinite(result.imageItemCount) ? result.imageItemCount : 0),
        0,
      );
      const tableRules = appearanceResults.reduce(
        (sum, result) => sum + (Number.isFinite(result.tableRuleCount) ? result.tableRuleCount : 0),
        0,
      );
      const compact = (value, unit = "%") => (value == null ? "—" : `${value.toFixed(2)}${unit}`);
      const scope = isFullHistoryRun(run)
        ? "full corpus"
        : `partial: ${(run.refreshed ?? [...new Set(results.map((r) => r.fixture))]).join(", ")}`;
      const outcome = run.outcome ?? (isFullHistoryRun(run) ? "accepted" : "legacy");
      const pageAppearance = appearanceAvailable
        ? `weight ${compact(pageWeight)} · colour ${compact(pageColor, " ΔE00")}`
        : "unavailable";
      const textAppearance = appearanceAvailable
        ? `weight ${compact(semantic.textWeight)} · colour ${compact(semantic.textColor, " ΔE00")} · coverage ${compact(semantic.textCoverage)}`
        : "unavailable";
      const objectAppearance = appearanceAvailable
        ? `image ${compact(semantic.imageWeight)} (${imageItems}) · fill ${compact(semantic.tableFillWeight)} · rules ${compact(semantic.tableRuleWeight)} (${tableRules})`
        : "unavailable";
      return (
        `<tr><td>${escapeHtml(run.ts ?? "unknown")}</td>` +
        `<td>${escapeHtml(run.gitSha ?? "?")}</td>` +
        `<td>${escapeHtml(run.label ?? run.experiment ?? "—")}</td>` +
        `<td>${escapeHtml(metricName(run))}</td>` +
        `<td>${appearanceAvailable ? escapeHtml(APPEARANCE_METRIC_VERSION) : "unavailable"}</td>` +
        `<td>${escapeHtml(scope)}</td>` +
        `<td>${escapeHtml(outcome)}</td>` +
        `<td class="num">${results.length}</td>` +
        `<td class="num">${fmtPct(mean)}</td>` +
        `<td class="num">${fmtPct(Math.max(...values))}</td>` +
        `<td class="compact">${pageAppearance}</td>` +
        `<td class="compact">${textAppearance}</td>` +
        `<td class="compact">${objectAppearance}</td>` +
        `<td class="num">${semantic.categoryMetricMs == null ? "—" : `${semantic.categoryMetricMs.toFixed(1)}ms`}</td></tr>`
      );
    })
    .join("");

  return (
    `<section class="run-history"><h2>Persisted run history</h2>` +
    `<p class="metric-note">Semantic appearance is comparable only for ${escapeHtml(APPEARANCE_METRIC_VERSION)} runs; earlier appearance data is shown as unavailable. Structural history remains independently comparable by ${escapeHtml(METRIC_VERSION)}.</p>` +
    `<div class="table-scroll"><table class="ptable"><thead><tr>` +
    `<th>Timestamp</th><th>SHA</th><th>Label</th><th>Structural metric</th><th>Appearance metric</th><th>Scope</th><th>Outcome</th>` +
    `<th class="num">Pages</th><th class="num">Structural mean</th><th class="num">Worst</th>` +
    `<th>Page controls</th><th>Text</th><th>Images / tables</th><th class="num">Category time</th>` +
    `</tr></thead><tbody>${rows}</tbody></table></div></section>`
  );
}

function buildTable(results, prev) {
  const prevMap = new Map();
  for (const r of comparable(prev)) prevMap.set(keyOf(r), r.severityPct);
  const rows = [...results].sort((a, b) => b.severityPct - a.severityPct);
  const body = rows
    .map((r) => {
      const png = r.pngRel ?? `${r.fixture}-p${r.page}.png`;
      let dcell = "—";
      if (prevMap.has(keyOf(r))) {
        const d = r.severityPct - prevMap.get(keyOf(r));
        if (Math.abs(d) < 0.005) dcell = "no change";
        else {
          const arrow = d < 0 ? "▼" : "▲";
          const cls = d < 0 ? "good" : "critical";
          dcell = `<span class="${cls}">${arrow} ${Math.abs(d).toFixed(2)} pp</span>`;
        }
      }
      const align = r.alignPx != null ? `${r.alignPx.toFixed(1)}px` : "—";
      const lineShift = r.lineShiftPct != null ? fmtPct(r.lineShiftPct) : "—";
      const weight =
        r.appearanceWeightErrorPct != null && r.appearanceWeightRatio != null
          ? `${fmtPct(r.appearanceWeightErrorPct)} · ${r.appearanceWeightRatio.toFixed(3)}×`
          : "—";
      const color =
        r.appearanceColorDeltaE != null ? `${r.appearanceColorDeltaE.toFixed(2)} ΔE00` : "—";
      const weightMetric = (error, ratio, mass) => {
        if (error == null) return "—";
        const signed = ratio == null ? "" : ` · ${ratio.toFixed(3)}×`;
        const title = mass == null ? "" : ` title="category mass ${mass.toFixed(1)}"`;
        return `<span${title}>${fmtPct(error)}${signed}</span>`;
      };
      const textWeight = weightMetric(r.textWeightErrorPct, r.textWeightRatio, r.textWeightMass);
      const textColor =
        r.textColorDeltaE == null
          ? "—"
          : `<span${r.textColorMass == null ? "" : ` title="colour mass ${r.textColorMass.toFixed(1)}"`}>${r.textColorDeltaE.toFixed(2)} ΔE00</span>`;
      const textCoverage =
        r.textColorCoveragePct == null ? "—" : fmtPct(r.textColorCoveragePct);
      const textAppearance = `weight ${textWeight}<br>colour ${textColor} · coverage ${textCoverage}`;
      const imageWeight = weightMetric(r.imageWeightErrorPct, r.imageWeightRatio, r.imageWeightMass);
      const imageCount = Number.isFinite(r.imageItemCount) ? r.imageItemCount : 0;
      const imageAppearance = r.imageWeightErrorPct == null ? "—" : `${imageWeight} · ${imageCount} item${imageCount === 1 ? "" : "s"}`;
      const tableFill = weightMetric(
        r.tableFillWeightErrorPct,
        r.tableFillWeightRatio,
        r.tableFillWeightMass,
      );
      const tableRule = weightMetric(
        r.tableRuleWeightErrorPct,
        r.tableRuleWeightRatio,
        r.tableRuleWeightMass,
      );
      const ruleCount = Number.isFinite(r.tableRuleCount) ? r.tableRuleCount : 0;
      const tableAppearance =
        r.tableFillWeightErrorPct == null && r.tableRuleWeightErrorPct == null
          ? "—"
          : `fill ${tableFill}<br>rules ${tableRule} · ${ruleCount}`;
      const categoryTime = r.categoryMetricMs == null
        ? escapeHtml(r.categoryMetricStatus ?? "—")
        : `${r.categoryMetricMs.toFixed(1)}ms · ${escapeHtml(r.categoryMetricStatus ?? "unknown")}`;
      return (
        `<tr><td>${escapeHtml(r.fixture)}</td><td>p${r.page}</td>` +
        `<td class="num">${fmtPct(r.severityPct)}</td>` +
        `<td><span class="chip ${chipClass(r)}">${escapeHtml(r.driftClass || "clean")}</span></td>` +
        `<td class="num">${lineShift}</td><td class="num">${align}</td><td class="num">${weight}<br>${color}</td>` +
        `<td class="compact">${textAppearance}</td><td class="compact">${imageAppearance}</td>` +
        `<td class="compact">${tableAppearance}</td><td class="num">${categoryTime}</td>` +
        `<td class="num">${fmtPct(r.mismatchPct)}</td><td>${dcell}</td>` +
        `<td><a href="${escapeHtml(png)}" target="_blank" rel="noopener">diff</a></td></tr>`
      );
    })
    .join("");
  return (
    `<details class="table-block"><summary>All pages as text</summary>` +
    `<table class="ptable"><thead><tr><th>Fixture</th><th>Page</th>` +
    `<th class="num">Structural</th><th>Drift</th><th class="num">Line order</th><th class="num">Offset</th>` +
    `<th class="num">Page controls</th><th>Text</th><th>Images</th><th>Tables</th><th class="num">Category time</th><th class="num">Page area</th>` +
    `<th>Δ vs previous</th><th>Diff</th></tr></thead>` +
    `<tbody>${body}</tbody></table></details>`
  );
}

export function buildReport(results, history, meta) {
  // Provenance tabs: the main page evaluates Word-authored fixtures only
  // (the parity target). Other provenances (libreoffice, googledocs, ...)
  // each get their own tab so cross-suite drift is visible without
  // polluting the Word axis. History deltas stay word-only for continuity.
  const byProv = new Map();
  for (const r of results) {
    const prov = r.provenance ?? "word";
    if (!byProv.has(prov)) byProv.set(prov, []);
    byProv.get(prov).push(r);
  }
  if (byProv.size > 1) {
    const tabs = [...byProv.keys()].sort((a, b) => (a === "word" ? -1 : b === "word" ? 1 : a < b ? -1 : 1));
    const sections = tabs.map((prov) => {
      const inner = buildReportSingle(byProv.get(prov), prov === "word" ? history : [], {
        ...meta,
        provenanceLabel: prov,
      });
      const body = inner.replace(/^[\s\S]*?<body[^>]*>/, "").replace(/<\/body>[\s\S]*$/, "");
      return `<div class="prov-tab" id="tab-${prov}" style="display:${prov === "word" ? "block" : "none"}">${body}</div>`;
    });
    const nav = tabs
      .map(
        (prov) =>
          `<button class="prov-btn${prov === "word" ? " active" : ""}" onclick="for(const t of document.querySelectorAll('.prov-tab'))t.style.display='none';document.getElementById('tab-${prov}').style.display='block';for(const b of document.querySelectorAll('.prov-btn'))b.classList.remove('active');this.classList.add('active')">${prov}${prov === "word" ? "" : " (deferred)"}</button>`,
      )
      .join("");
    const shell = buildReportSingle(byProv.get("word") ?? results, history, meta);
    const head = shell.match(/^[\s\S]*?<body[^>]*>/)[0];
    const tail = shell.match(/<\/body>[\s\S]*$/)[0];
    const navCss = `<style>.prov-nav{display:flex;gap:8px;margin:0 0 16px}.prov-btn{padding:6px 14px;border:1px solid var(--border,#ccc);border-radius:6px;background:transparent;cursor:pointer;font:inherit}.prov-btn.active{background:var(--text-secondary,#333);color:#fff}</style>`;
    return head + navCss + `<nav class="prov-nav">${nav}</nav>` + sections.join("") + tail;
  }
  return buildReportSingle(results, history, meta);
}

// One-line framing; the full methodology lives in BLOG.md.
function buildIntro() {
  return (
    `<p class="intro">WordInWeb's in-browser render of every fixture, compared pixel-by-pixel ` +
    `against desktop Microsoft Word's PDF of the same file. Scores are the % of ink that ` +
    `doesn't match (antialiasing noise excluded); lower is better, 0.00 is pixel-parity.</p>`
  );
}

function buildReportSingle(results, history, meta) {
  const prev = previousRun(history, meta.isFullRun);
  const appearancePrev = previousAppearanceRun(history, meta.isFullRun);
  const subtitle = [
    meta.gitSha ? `sha ${escapeHtml(meta.gitSha)}` : "sha unknown",
    escapeHtml(meta.generatedAt),
    `${results.length} pages`,
    `base ${escapeHtml(meta.base)}`,
    // Partial runs carry the other fixtures over from the last full run.
    ...(meta.refreshed ? [`partial run — refreshed ${escapeHtml(meta.refreshed.join(", "))}`] : []),
    ...(meta.outcome ? [`outcome ${escapeHtml(meta.outcome)}`] : []),
    ...(meta.label ? [`label ${escapeHtml(meta.label)}`] : []),
    `appearance ${
      meta.appearanceMetricVersion === APPEARANCE_METRIC_VERSION
        ? escapeHtml(APPEARANCE_METRIC_VERSION)
        : "unavailable"
    }`,
  ].join(" · ");

  return `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>WordInWeb pixel parity</title>
<style>
${STYLE}
</style>
</head>
<body>
<div class="viz-root">
  <header class="report-head">
    <h1>WordInWeb pixel parity</h1>
    <div class="subtitle">${subtitle}</div>
  </header>
  ${buildIntro()}
  ${buildKpis(results, prev)}
  ${buildCategoryCards(results, prev)}
  ${buildCategorySections(results, prev)}
  ${buildSemanticKpis(results, appearancePrev, meta.appearanceMetricVersion)}
  ${buildTrend(history, meta.provenanceLabel)}
  ${buildRunHistory(history)}
  ${buildTable(results, prev)}
  <div class="tooltip" style="display:none"></div>
</div>
<script>
${SCRIPT}
</script>
</body>
</html>`;
}

const STYLE = `
.viz-root {
  --surface-1:#fcfcfb; --text-primary:#0b0b0b; --text-secondary:#52514e;
  --grid:#e8e7e3; --series-1:#2a78d6; --series-2:#1baf7a;
  --good:#0ca30c; --critical:#d03b3b;
}
@media (prefers-color-scheme: dark) {
  .viz-root {
    --surface-1:#1a1a19; --text-primary:#ffffff; --text-secondary:#c3c2b7;
    --grid:#33332f; --series-1:#3987e5; --series-2:#199e70;
    --good:#0ca30c; --critical:#d03b3b;
  }
}
* { box-sizing: border-box; }
body { margin: 0; background: var(--surface-1); }
.viz-root {
  font-family: system-ui, -apple-system, sans-serif;
  color: var(--text-primary);
  background: var(--surface-1);
  max-width: 1100px;
  margin: 0 auto;
  padding: 24px;
}
.report-head h1 { font-size: 26px; margin: 0 0 4px; }
.subtitle { color: var(--text-secondary); font-size: 13px; }
.metric-note { color: var(--text-secondary); font-size: 12px; margin: 2px 0 10px; }
.empty-note { color: var(--text-secondary); font-size: 14px; margin: 8px 0 0; }
.kpis { display: grid; grid-template-columns: repeat(3, 1fr); gap: 16px; margin: 24px 0; }
.kpi {
  border: 1px solid var(--grid); border-radius: 10px; padding: 16px;
  display: flex; flex-direction: column; gap: 6px; min-width: 0;
}
.kpi-label { font-size: 11px; text-transform: uppercase; letter-spacing: .05em; color: var(--text-secondary); }
.kpi-value { font-size: 40px; font-weight: 650; line-height: 1.05; }
.kpi:nth-child(2) .kpi-value { font-size: 22px; font-weight: 600; }
.semantic-block .chart-head { margin-bottom: 4px; }
.semantic-kpis { margin-top: 12px; }
.semantic-kpi .kpi-value { font-size: 25px; font-weight: 600; }
.metric-version { color: var(--text-secondary); font-size: 12px; }
.gate { font-size: 11px; font-weight: 700; text-transform: uppercase; }
.gate.good { color: var(--good); }
.gate.critical { color: var(--critical); }
.unavailable { color: var(--text-secondary); font-size: 18px; font-weight: 500; }
.gate.unavailable { font-size: 11px; font-weight: 700; }
.worst-name { color: var(--text-secondary); font-weight: 500; }
.delta { font-size: 13px; font-weight: 600; }
.delta.good { color: var(--good); }
.delta.critical { color: var(--critical); }
.delta.neutral { color: var(--text-secondary); font-weight: 500; }
.chart-head { display: flex; align-items: baseline; justify-content: space-between; gap: 16px; margin: 28px 0 12px; }
.chart-head h2 { font-size: 16px; margin: 0; }
.legend { display: flex; gap: 16px; font-size: 12px; color: var(--text-secondary); }
.lg { display: inline-flex; align-items: center; gap: 6px; }
.lg-bar { width: 14px; height: 10px; border-radius: 0 3px 3px 0; background: var(--series-1); }
.lg-tick { width: 2px; height: 14px; background: var(--text-secondary); opacity: .45; }
.lg-line { width: 16px; height: 2px; }
.lg-line.s1 { background: var(--series-1); }
.lg-line.s2 { background: var(--series-2); }
.chart { position: relative; padding-left: 250px; padding-right: 56px; }
.rows { position: relative; }
.gridline { position: absolute; top: 0; bottom: 0; width: 1px; background: var(--grid); z-index: 0; }
.row { position: relative; height: 18px; z-index: 1; cursor: pointer; }
.row:hover { background: color-mix(in srgb, var(--series-1) 8%, transparent); }
.rlabel {
  position: absolute; left: -250px; width: 242px; height: 18px;
  display: flex; align-items: center; gap: 8px;
  font-size: 11px; color: var(--text-secondary);
}
.rlabel .fxwrap {
  flex: 1 1 auto; text-align: right; line-height: 18px;
  white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
}
.rlabel .fx { color: var(--text-primary); font-weight: 550; }
/* Drift-class chip: coloured dot + word (never colour alone). */
.chip {
  flex: 0 0 auto; display: inline-flex; align-items: center; gap: 4px;
  font-size: 10px; color: var(--text-secondary); white-space: nowrap;
}
.chip::before, .chip-dot {
  content: ""; width: 8px; height: 8px; border-radius: 50%;
  background: var(--text-secondary); flex: 0 0 auto; display: inline-block;
}
.chip.clean::before, .chip-dot.clean { background: var(--good); }
.chip.alignment::before, .chip-dot.alignment { background: var(--series-1); }
.chip.color::before, .chip-dot.color { background: var(--series-2); }
.chip.weight::before, .chip-dot.weight { background: var(--text-secondary); }
.chip.structural::before, .chip-dot.structural { background: var(--critical); }
.drift-key { display: flex; flex-wrap: wrap; gap: 14px; margin: 0 0 12px; font-size: 12px; color: var(--text-secondary); }
.track { position: relative; height: 18px; }
.bar {
  position: absolute; top: 2px; height: 14px; min-width: 2px;
  background: var(--series-1); border-radius: 0 4px 4px 0;
}
.prev-tick {
  position: absolute; top: 0; height: 18px; width: 2px;
  background: var(--text-secondary); opacity: .45; z-index: 2;
}
.vlabel {
  position: absolute; top: 0; height: 18px; line-height: 18px; margin-left: 6px;
  font-size: 11px; color: var(--text-primary); white-space: nowrap;
}
.axis { position: relative; height: 18px; margin-top: 8px; }
.axis-tick {
  position: absolute; top: 0; font-size: 10px; color: var(--text-secondary);
  transform: translateX(-50%);
}
.trend-block { margin-top: 8px; }
.trend-wrap { position: relative; }
.trend { width: 100%; height: auto; display: block; overflow: visible; }
.t-grid { stroke: var(--grid); stroke-width: 1; }
.t-ylabel { fill: var(--text-secondary); font-size: 10px; text-anchor: end; }
.t-xlabel { fill: var(--text-secondary); font-size: 10px; text-anchor: middle; }
.t-line { fill: none; stroke-width: 2; }
.t-line.s1 { stroke: var(--series-1); }
.t-line.s2 { stroke: var(--series-2); }
.t-line.legacy { stroke-dasharray: 6 5; opacity: .55; }
.t-dot.s1 { fill: var(--series-1); }
.t-dot.s2 { fill: var(--series-2); }
.t-dot.legacy { opacity: .55; }
.t-end { font-size: 11px; font-weight: 600; }
.t-end.s1 { fill: var(--series-1); }
.t-end.s2 { fill: var(--series-2); }
.t-cross { stroke: var(--text-secondary); stroke-width: 1; opacity: .4; }
.trend-tip, .tooltip {
  position: fixed; z-index: 10; pointer-events: none;
  background: var(--surface-1); border: 1px solid var(--grid); border-radius: 8px;
  padding: 8px 10px; font-size: 12px; color: var(--text-primary);
  box-shadow: 0 4px 16px rgba(0,0,0,.18); max-width: 240px;
}
.tooltip .tt-title { font-weight: 650; margin-bottom: 2px; }
.tooltip .tt-row { color: var(--text-secondary); }
.tooltip .tt-hint { margin-top: 4px; color: var(--series-1); }
.table-block { margin-top: 28px; }
.run-history { margin-top: 28px; }
.run-history h2 { font-size: 15px; margin: 0 0 8px; }
.table-scroll { overflow-x: auto; }
.table-block summary { cursor: pointer; font-size: 14px; font-weight: 600; padding: 6px 0; }
.ptable { border-collapse: collapse; width: 100%; font-size: 13px; margin-top: 8px; }
.ptable th, .ptable td { text-align: left; padding: 6px 10px; border-bottom: 1px solid var(--grid); }
.ptable th.num, .ptable td.num { text-align: right; font-variant-numeric: tabular-nums; }
.ptable td.compact { min-width: 150px; font-size: 12px; white-space: nowrap; }
.ptable .good { color: var(--good); }
.ptable .critical { color: var(--critical); }
.ptable a { color: var(--series-1); }

/* Intro write-up */
.intro {
  font-size: 13.5px; line-height: 1.55; color: var(--text-secondary);
  margin: 10px 0 4px; max-width: 88ch;
}

/* Severity colour bands */
.sev-good { color: var(--good); }
.sev-warn { color: #d08a1e; }
.sev-bad { color: var(--critical); }

/* Category card grid */
.cat-cards-block { margin-top: 20px; }
.cat-cards {
  display: grid; grid-template-columns: repeat(auto-fill, minmax(280px, 1fr));
  gap: 14px; margin-top: 12px;
}
.cat-card {
  border: 1px solid var(--grid); border-radius: 12px; padding: 16px 16px 14px;
  display: flex; flex-direction: column; gap: 7px; min-width: 0;
  background: var(--surface-1);
  /* Accent strip as an inset pseudo-element: a border-top on a rounded box
     smears around the corner radius and dies into the grey side border. The
     strip is clipped by the card's own radius instead. */
  position: relative; overflow: hidden;
}
.cat-card::before {
  content: ""; position: absolute; inset: 0 0 auto 0; height: 3px;
  background: var(--accent);
}
.cat-card-hd { display: flex; align-items: center; gap: 8px; text-decoration: none; color: inherit; }
.cat-card-hd:hover .cat-name { text-decoration: underline; }
.cat-dot { width: 10px; height: 10px; border-radius: 50%; background: var(--accent); flex: 0 0 auto; }
.cat-name { font-size: 14.5px; font-weight: 650; }
.cat-meta { font-size: 11.5px; color: var(--text-secondary); }
.cat-mean { display: flex; align-items: baseline; gap: 8px; flex-wrap: wrap; margin-top: 2px; }
.cat-mean-val { font-size: 30px; font-weight: 680; line-height: 1; font-variant-numeric: tabular-nums; }
.cat-mean-lbl { font-size: 11px; text-transform: uppercase; letter-spacing: .04em; color: var(--text-secondary); }
.cat-mean .delta { font-size: 12px; }
.cat-stats { display: flex; gap: 16px; font-size: 12px; color: var(--text-secondary); }
.cat-stats b { color: var(--text-primary); font-weight: 650; }
.cat-worst { font-size: 12px; color: var(--text-secondary); }
.cat-worst a { color: var(--series-1); }
.cat-blurb { font-size: 12px; line-height: 1.5; color: var(--text-secondary); margin: 4px 0 0; }

/* Category detail sections */
.cat-sections-block { margin-top: 28px; }
.cat-section {
  border: 1px solid var(--grid);
  border-radius: 10px; padding: 4px 16px; margin: 10px 0;
  position: relative; overflow: hidden;
}
.cat-section::before {
  content: ""; position: absolute; inset: 0 auto 0 0; width: 3px;
  background: var(--accent);
}
.cat-section > summary {
  cursor: pointer; padding: 12px 0; list-style: none;
  display: flex; align-items: center; gap: 8px; flex-wrap: wrap;
}
.cat-section > summary::-webkit-details-marker { display: none; }
.cat-section > summary::after { content: "▸"; margin-left: auto; color: var(--text-secondary); }
.cat-section[open] > summary::after { content: "▾"; }
.cat-sum-meta { font-size: 12px; color: var(--text-secondary); font-weight: 400; }
.cat-sum-meta b { font-weight: 650; }
.cat-section .chart { margin-top: 6px; }
.cat-section .cat-blurb { margin: 2px 0 10px; }
.cat-table { margin: 14px 0 6px; }
.cat-table > summary { cursor: pointer; font-size: 13px; font-weight: 600; color: var(--text-secondary); padding: 4px 0; }

@media (max-width: 720px) {
  .kpis { grid-template-columns: repeat(2, 1fr); }
  .cat-cards { grid-template-columns: 1fr; }
}
`;

const SCRIPT = `
(function () {
  var root = document.querySelector(".viz-root");
  var tip = root.querySelector(".tooltip");
  // textContent (never innerHTML) for data-derived text: attribute values are
  // decoded back to raw on read, so innerHTML would re-interpret any markup.
  function mkRow(cls, text) {
    var d = document.createElement("div");
    d.className = cls;
    d.textContent = text;
    return d;
  }
  function place(el, x, y) {
    var w = el.offsetWidth, h = el.offsetHeight, m = 14;
    var px = x + m, py = y + m;
    if (px + w > window.innerWidth - 4) px = x - w - m;
    if (py + h > window.innerHeight - 4) py = y - h - m;
    el.style.left = Math.max(4, px) + "px";
    el.style.top = Math.max(4, py) + "px";
  }
  root.querySelectorAll(".row").forEach(function (row) {
    row.addEventListener("mousemove", function (e) {
      var prev = row.dataset.prev;
      var d = row.dataset;
      tip.textContent = "";
      tip.appendChild(mkRow("tt-title", d.fixture + " \\u00b7 p" + d.page));
      tip.appendChild(mkRow("tt-row", "drift class: " + d.class));
      tip.appendChild(mkRow("tt-row", "structural: " + d.val + "%"));
      if (d.line !== undefined)
        tip.appendChild(mkRow("tt-row", "line-order shift: " + d.line + "% beyond floor"));
      if (d.raw) tip.appendChild(mkRow("tt-row", "page area: " + d.raw + "%"));
      if (d.align !== undefined)
        tip.appendChild(mkRow("tt-row", "offset: " + d.align + "px median, " + (d.alignp95 || "0") + "px p95 (" + (d.mis || "0") + "% shifted)"));
      if (d.weight !== undefined) {
        var wp = Math.round((parseFloat(d.weight) - 1) * 100);
        tip.appendChild(mkRow("tt-row", "page weight: " + d.weighterror + "% error, " + parseFloat(d.weight).toFixed(3) + "x (" + (wp >= 0 ? "+" : "\\u2212") + Math.abs(wp) + "% signed)"));
      }
      if (d.color !== undefined)
        tip.appendChild(mkRow("tt-row", "source colour: " + d.color + " mean \\u0394E00"));
      if (prev !== "") {
        var diff = parseFloat(d.val) - parseFloat(prev);
        var arrow = diff < 0 ? "\\u25bc" : (diff > 0 ? "\\u25b2" : "");
        tip.appendChild(mkRow("tt-row", "\\u0394 vs previous: " + arrow + " " + Math.abs(diff).toFixed(2) + " pp"));
      }
      tip.appendChild(mkRow("tt-hint", "click to open diff"));
      tip.style.display = "block";
      place(tip, e.clientX, e.clientY);
    });
    row.addEventListener("mouseleave", function () { tip.style.display = "none"; });
    row.addEventListener("click", function () { window.open(row.dataset.png, "_blank"); });
  });

  var svg = root.querySelector(".trend");
  if (svg) {
    var runs = JSON.parse(svg.getAttribute("data-runs"));
    var W = +svg.getAttribute("data-w");
    var padT = +svg.getAttribute("data-padt");
    var H = +svg.getAttribute("data-h");
    var padB = +svg.getAttribute("data-padb");
    var cross = svg.querySelector(".t-cross");
    var ttip = root.querySelector(".trend-tip");
    svg.addEventListener("mousemove", function (e) {
      var rect = svg.getBoundingClientRect();
      var vx = ((e.clientX - rect.left) / rect.width) * W;
      var best = 0, bd = Infinity;
      for (var i = 0; i < runs.length; i++) {
        var dd = Math.abs(runs[i].x - vx);
        if (dd < bd) { bd = dd; best = i; }
      }
      var r = runs[best];
      cross.setAttribute("x1", r.x); cross.setAttribute("x2", r.x); cross.style.display = "block";
      cross.setAttribute("y1", padT); cross.setAttribute("y2", H - padB);
      ttip.textContent = "";
      ttip.appendChild(mkRow("tt-title", r.sha));
      ttip.appendChild(mkRow("tt-row", "metric: " + r.metric));
      ttip.appendChild(mkRow("tt-row", "mean: " + r.mean.toFixed(2) + "%"));
      ttip.appendChild(mkRow("tt-row", "worst: " + r.worst.toFixed(2) + "%"));
      ttip.style.display = "block";
      place(ttip, e.clientX, e.clientY);
    });
    svg.addEventListener("mouseleave", function () {
      cross.style.display = "none"; ttip.style.display = "none";
    });
  }
})();
`;
