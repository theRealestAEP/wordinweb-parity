import React, { useCallback, useEffect, useRef, useState } from "react";
// Metric-compatible substitutes for Office fonts (Calibri->Carlito,
// Cambria->Caladea) so measurement matches Word's glyph advances.
import "@fontsource/carlito/400.css";
import "@fontsource/carlito/400-italic.css";
import "@fontsource/carlito/700.css";
import "@fontsource/carlito/700-italic.css";
import "@fontsource/caladea/400.css";
import "@fontsource/caladea/700.css";
// OFL substitute for Word's DokChampa (Lao script): same looped style;
// DokChampa itself is licensed and cannot be bundled.
import "@fontsource/noto-sans-lao-looped/400.css";
import "@fontsource/noto-sans-lao-looped/700.css";
// Dev-only: register the REAL Office fonts (Cambria Math, real Calibri/Times/
// Arial, CJK families) from /fonts-local/ when present. Git-ignored; falls back
// to the substitutes above when the files are absent. See fonts-local.css.
import "./fonts-local.css";
import "./app.css";
import { createRoot } from "react-dom/client";
import { gzip, gunzip } from "fflate";
import { DocxView, DocxToolbar, DocxViewApi, ToolbarMenuSelect, printPages } from "wordinweb";

// Tracked-change insertion ink (see packages/core/src/parse/document.ts) — the
// color your suggestions render in, echoed by the mode control + author chip.
const MODE_INK = "#C00000";

type PerfSample = {
  total: number; layout: number; render: number; destroy: number;
  refresh: number; chromeCaret: number; rerenderCall: number;
  totalPages: number; pagesReused: number;
};
type PerfGlobal = { samples?: PerfSample[]; last?: Record<string, number>; lastReused?: number };

// ?perf=1 turns on the per-keystroke performance HUD. Setting the __dxwPerf
// global (which the core editor/renderer check before recording) here — at
// module load, before any editing — activates the sample recording the overlay
// reads. Off by default so normal sessions pay nothing.
const PERF_ON = new URLSearchParams(location.search).get("perf") === "1";
if (PERF_ON) {
  (globalThis as { __dxwPerf?: PerfGlobal }).__dxwPerf = { samples: [] };
}

function median(xs: number[]): number {
  if (xs.length === 0) return 0;
  const s = [...xs].sort((a, b) => a - b);
  return s[Math.floor(s.length / 2)];
}

/** Fixed overlay showing the last keystroke's per-stage cost and a rolling
 * median over the recent samples, reading the __dxwPerf globals the core sets.
 * The numbers render as selectable text (plus a Copy button) so the user can
 * paste ground-truth timings from their own session. */
function PerfHud() {
  const [, tick] = useState(0);
  const perf = (globalThis as { __dxwPerf?: PerfGlobal }).__dxwPerf;
  useEffect(() => {
    const id = setInterval(() => tick((n) => n + 1), 200);
    return () => clearInterval(id);
  }, []);
  const samples = perf?.samples ?? [];
  const last = samples[samples.length - 1];
  const recent = samples.slice(-15);
  const med = (k: keyof PerfSample) => median(recent.map((s) => s[k]));
  const f = (n: number | undefined) => (n === undefined ? "—" : n.toFixed(1));
  const text = last
    ? [
        `keystrokes:  ${samples.length}`,
        `total:       ${f(last.total)} ms   (median15 ${f(med("total"))})`,
        `  layout:    ${f(last.layout)} ms   (median15 ${f(med("layout"))})`,
        `  render:    ${f(last.render)} ms   (median15 ${f(med("render"))})`,
        `  destroy:   ${f(last.destroy)} ms`,
        `  refresh:   ${f(last.refresh)} ms`,
        `  chrome/car:${f(last.chromeCaret)} ms`,
        `pages:       ${last.pagesReused}/${last.totalPages} reused`,
      ].join("\n")
    : "Type in the document to record keystroke timings…";
  const copy = () => navigator.clipboard?.writeText(text);
  const reset = () => {
    if (perf) perf.samples = [];
    tick((n) => n + 1);
  };
  return (
    <div
      style={{
        position: "fixed", right: 12, bottom: 12, zIndex: 9999,
        background: "rgba(17,17,17,0.92)", color: "#e6e6e6",
        font: "12px/1.45 ui-monospace, SFMono-Regular, Menlo, monospace",
        padding: "10px 12px", borderRadius: 8, boxShadow: "0 4px 16px rgba(0,0,0,0.4)",
        maxWidth: 320, whiteSpace: "pre", userSelect: "text",
      }}
    >
      <div style={{ display: "flex", gap: 8, marginBottom: 6, alignItems: "center" }}>
        <strong style={{ color: "#8fd3ff" }}>perf</strong>
        <button onClick={copy} style={{ marginLeft: "auto", fontSize: 11 }}>Copy</button>
        <button onClick={reset} style={{ fontSize: 11 }}>Reset</button>
      </div>
      {text}
    </div>
  );
}

type Mode = "editing" | "suggesting" | "viewing";

const PencilIcon = ({ color }: { color: string }) => (
  <svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke={color} strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round">
    <path d="M11.5 2.5l2 2L6 12l-2.6.6.6-2.6z" />
    <path d="M10.2 3.8l2 2" />
  </svg>
);
const SuggestIcon = ({ color }: { color: string }) => (
  <svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke={color} strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round">
    <path d="M10.5 2.5l2 2L6 11l-2.6.6.6-2.6z" />
    <path d="M11.5 9.5v4M9.5 11.5h4" />
  </svg>
);
const EyeIcon = ({ color }: { color: string }) => (
  <svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke={color} strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round">
    <path d="M1.5 8S4 3.5 8 3.5 14.5 8 14.5 8 12 12.5 8 12.5 1.5 8 1.5 8z" />
    <circle cx="8" cy="8" r="1.8" />
  </svg>
);
const ChevIcon = ({ color }: { color: string }) => (
  <svg width="12" height="12" viewBox="0 0 12 12" fill="none" stroke={color} strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round">
    <path d="M3 4.5L6 7.5 9 4.5" />
  </svg>
);

const MODES: { id: Mode; label: string; desc: string; ink: string; bg: string; Icon: typeof PencilIcon }[] = [
  { id: "editing", label: "Editing", desc: "Edit the document directly", ink: "#1a73e8", bg: "#e8f0fe", Icon: PencilIcon },
  { id: "suggesting", label: "Suggesting", desc: "Edits become tracked suggestions", ink: MODE_INK, bg: "#fce8e6", Icon: SuggestIcon },
  { id: "viewing", label: "Viewing", desc: "Read or print — no editing", ink: "#188038", bg: "#e6f4ea", Icon: EyeIcon },
];

// Résumé and pleading paper are REAL documents, not generated: the résumé is
// CareerOneStop's (US DOL) published sample re-personed to Jane Doe with
// scrubbed metadata; the pleading paper is a copy of the anonymized wild
// fixture (authentic 28-line layout) with its digit-ciphered literal line
// numbers restored to 1–28 — pleading-anon.docx itself stays byte-identical
// to the copy its Word-reference parity PDF was exported from.
const PRESETS = [
  { id: "resume", label: "Résumé", path: "/fixtures/wild3-resume.docx" },
  { id: "pleading", label: "California pleading paper", path: "/fixtures/pleading-paper.docx" },
  { id: "equations", label: "Math equations", path: "/fixtures/preset-equations.docx" },
  { id: "tables", label: "Tables & reports", path: "/fixtures/preset-tables.docx" },
  { id: "publication", label: "Magazine / newspaper", path: "/fixtures/preset-publication.docx" },
  { id: "chapter", label: "Chapter book", path: "/fixtures/preset-chapter-book.docx" },
  { id: "model3d", label: "Native 3D model", path: "/fixtures/model3d-cube.docx" },
] as const;

const WORKSPACE_DB = "wordinweb-demo";
const WORKSPACE_STORE = "workspace";
const WORKSPACE_KEY = "current";
const AUTOSAVE_MS = 60_000;

type SavedWorkspace = {
  id: typeof WORKSPACE_KEY;
  version: 1;
  fileName: string;
  preset: string;
  savedAt: number;
  compression: "gzip" | "none";
  bytes: ArrayBuffer;
};

let workspaceDb: Promise<IDBDatabase> | null = null;

function openWorkspaceDb(): Promise<IDBDatabase> {
  if (workspaceDb) return workspaceDb;
  workspaceDb = new Promise((resolve, reject) => {
    const request = indexedDB.open(WORKSPACE_DB, 1);
    request.onupgradeneeded = () => {
      if (!request.result.objectStoreNames.contains(WORKSPACE_STORE)) {
        request.result.createObjectStore(WORKSPACE_STORE, { keyPath: "id" });
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error ?? new Error("Could not open local storage"));
  });
  return workspaceDb;
}

async function readSavedWorkspace(): Promise<SavedWorkspace | null> {
  const db = await openWorkspaceDb();
  return new Promise((resolve, reject) => {
    const request = db.transaction(WORKSPACE_STORE, "readonly").objectStore(WORKSPACE_STORE).get(WORKSPACE_KEY);
    request.onsuccess = () => resolve((request.result as SavedWorkspace | undefined) ?? null);
    request.onerror = () => reject(request.error ?? new Error("Could not read local work"));
  });
}

async function writeSavedWorkspace(workspace: SavedWorkspace): Promise<void> {
  const db = await openWorkspaceDb();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(WORKSPACE_STORE, "readwrite");
    transaction.objectStore(WORKSPACE_STORE).put(workspace);
    transaction.oncomplete = () => resolve();
    transaction.onerror = () => reject(transaction.error ?? new Error("Could not save local work"));
    transaction.onabort = () => reject(transaction.error ?? new Error("Local save was cancelled"));
  });
}

function gzipBytes(bytes: Uint8Array): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    gzip(bytes, { level: 6 }, (error, result) => error ? reject(error) : resolve(result));
  });
}

function gunzipBytes(bytes: Uint8Array): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    gunzip(bytes, (error, result) => error ? reject(error) : resolve(result));
  });
}

function copyBuffer(bytes: Uint8Array): ArrayBuffer {
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
}

/**
 * Google-Docs mode picker: one pencil dropdown replacing the tangle of
 * Edit/Suggesting toggles. Editing edits directly, Suggesting records tracked
 * changes (button tints in the revision ink + shows the author, a persistent
 * "your edits are tracked" cue), Viewing turns editing off entirely.
 */
function ModeControl({ mode, author, onChange }: { mode: Mode; author: string; onChange: (m: Mode) => void }) {
  const [open, setOpen] = useState(false);
  const rootRef = useRef<HTMLDivElement | null>(null);
  useEffect(() => {
    if (!open) return;
    const close = (e: MouseEvent) => {
      if (!rootRef.current?.contains(e.target as Node)) setOpen(false);
    };
    document.addEventListener("mousedown", close);
    return () => document.removeEventListener("mousedown", close);
  }, [open]);
  const cur = MODES.find((m) => m.id === mode)!;
  return (
    <div ref={rootRef} style={{ display: "inline-flex", alignItems: "center", gap: 8, position: "relative" }}>
      <button
        type="button"
        data-dxw-mode={mode}
        aria-haspopup="menu"
        aria-expanded={open}
        title={`${cur.label} — click to change mode`}
        onMouseDown={(e) => e.preventDefault()}
        onClick={() => setOpen(!open)}
        style={{
          display: "inline-flex", alignItems: "center", gap: 7,
          padding: "5px 10px",
          background: cur.bg, color: cur.ink,
          border: `1px solid ${mode === "editing" ? "#c6dafc" : mode === "suggesting" ? "#f3c9c4" : "#c3e6cd"}`,
          borderRadius: 8, cursor: "pointer",
          // Integer line box (18px, not 1.4×12.5=17.5): this button is the
          // tallest header child, so a fractional height leaks a half-pixel into
          // the header height and the DocxView scroll container's top offset. A
          // page then lands on a half CSS pixel, and Playwright's element
          // screenshot rounds the fractional box outward — capturing the page 2
          // device-px taller than Word's reference and spiking parity line-shift.
          font: "600 12.5px system-ui, sans-serif", lineHeight: "18px",
        }}
      >
        <cur.Icon color={cur.ink} />
        {cur.label}
        <ChevIcon color={cur.ink} />
      </button>
      {mode === "suggesting" && (
        <span
          title={`Your suggestions are attributed to ${author} and shown in this color`}
          style={{
            display: "inline-flex", alignItems: "center", gap: 6,
            padding: "3px 9px 3px 7px", borderRadius: 12,
            background: "#fce8e6", border: "1px solid #f3c9c4",
            font: "600 11.5px system-ui, sans-serif", color: MODE_INK, whiteSpace: "nowrap",
          }}
        >
          <span style={{ width: 8, height: 8, borderRadius: "50%", background: MODE_INK, flexShrink: 0 }} />
          {author}
        </span>
      )}
      {open && (
        <div
          role="menu"
          data-dxw-mode-menu=""
          style={{
            position: "absolute", top: 38, left: 0, zIndex: 100, width: 246,
            background: "#fff", border: "1px solid #dadce0", borderRadius: 10,
            boxShadow: "0 4px 20px rgba(0,0,0,.18)", padding: 6,
          }}
        >
          {MODES.map((m) => {
            const active = m.id === mode;
            return (
              <button
                key={m.id}
                role="menuitemradio"
                aria-checked={active}
                data-dxw-mode-option={m.id}
                onMouseDown={(e) => e.preventDefault()}
                onClick={() => { onChange(m.id); setOpen(false); }}
                style={{
                  display: "flex", alignItems: "center", gap: 10, width: "100%",
                  padding: "8px 10px", border: "none", borderRadius: 7, cursor: "pointer",
                  background: active ? m.bg : "transparent", textAlign: "left",
                }}
                onMouseEnter={(e) => { if (!active) (e.currentTarget as HTMLElement).style.background = "#f1f3f4"; }}
                onMouseLeave={(e) => { if (!active) (e.currentTarget as HTMLElement).style.background = "transparent"; }}
              >
                <span style={{ display: "inline-flex", width: 20, justifyContent: "center", flexShrink: 0 }}>
                  <m.Icon color={active ? m.ink : "#5f6368"} />
                </span>
                <span style={{ minWidth: 0, flex: 1 }}>
                  <span style={{ display: "block", font: `600 12.5px system-ui, sans-serif`, color: active ? m.ink : "#3c4043" }}>
                    {m.label}
                  </span>
                  <span style={{ display: "block", font: "11.5px system-ui, sans-serif", color: "#5f6368" }}>
                    {m.desc}
                  </span>
                </span>
                {active && <span style={{ color: m.ink, fontWeight: 700, flexShrink: 0 }}>✓</span>}
              </button>
            );
          })}
        </div>
      )}
    </div>
  );
}

function App() {
  const query = new URLSearchParams(location.search);
  const toolbarMode = query.get("toolbar") === "simple" ? "simple" : "advanced";
  const toolbarFeatures = query.get("layout") === "off" ? { layout: false } : undefined;
  const initial = query.get("doc") ?? "/fixtures/showcase.docx";
  const persistenceEnabled = !query.has("doc");
  const [source, setSource] = useState<ArrayBuffer | string | null>(persistenceEnabled ? null : initial);
  const [preset, setPreset] = useState(PRESETS.find((item) => item.path === initial)?.id ?? "");
  const [zoom, setZoom] = useState(1);
  const [editable, setEditable] = useState(query.get("editable") !== "0");
  const [showComments, setShowComments] = useState(query.get("comments") !== "0");
  // Author stamped on comments, replies and tracked-change suggestions.
  // Persisted so it behaves like an identity, not a per-session setting.
  const [author, setAuthor] = useState(
    query.get("author") ?? localStorage.getItem("dxw-author") ?? "You",
  );
  const changeAuthor = (name: string) => {
    const v = name.trim() || "You";
    setAuthor(v);
    localStorage.setItem("dxw-author", v);
  };
  const [suggesting, setSuggesting] = useState(query.get("suggest") === "1");
  const [pageCount, setPageCount] = useState<number | null>(null);
  const [status, setStatus] = useState<string>(persistenceEnabled ? "Loading saved work…" : "");
  const [fileName, setFileName] = useState(initial.split("/").pop() ?? "document.docx");
  const [api, setApi] = useState<DocxViewApi | null>(null);
  const apiRef = useRef<DocxViewApi | null>(null);
  const workspaceMetaRef = useRef({ fileName, preset });
  const saveInFlight = useRef<Promise<void> | null>(null);
  const [saveState, setSaveState] = useState<"idle" | "saving" | "saved" | "error">("idle");
  const [lastSaved, setLastSaved] = useState<number | null>(null);
  const [findOpen, setFindOpen] = useState(false);
  const [findQ, setFindQ] = useState("");
  const [replQ, setReplQ] = useState("");
  const [findStat, setFindStat] = useState<{ index: number; total: number } | null>(null);
  const findInput = useRef<HTMLInputElement | null>(null);
  const [missingFonts, setMissingFonts] = useState<{ family: string }[]>([]);
  const [fontWarnDismissed, setFontWarnDismissed] = useState(false);
  workspaceMetaRef.current = { fileName, preset };

  useEffect(() => {
    if (!persistenceEnabled) return;
    let active = true;
    void readSavedWorkspace()
      .then(async (saved) => {
        if (!active) return;
        if (!saved || saved.version !== 1) {
          setSource(initial);
          return;
        }
        const stored = new Uint8Array(saved.bytes);
        const restored = saved.compression === "gzip" ? await gunzipBytes(stored) : stored;
        if (!active) return;
        setFileName(saved.fileName);
        setPreset(saved.preset);
        setLastSaved(saved.savedAt);
        setSaveState("saved");
        setSource(copyBuffer(restored));
      })
      .catch(() => {
        if (!active) return;
        setSaveState("error");
        setSource(initial);
      });
    return () => { active = false; };
  }, []);
  // Pending tracked changes (suggestions) — drives the review pill. Updated
  // on selection events (clicks, accept/reject) and polled lightly while
  // suggesting so the count follows live typing.
  const [pendingRevisions, setPendingRevisions] = useState(0);
  useEffect(() => {
    if (!api) return;
    const upd = () => setPendingRevisions(api.revisionCount());
    upd();
    document.addEventListener("dxw-selection", upd);
    const iv = suggesting ? window.setInterval(upd, 1200) : 0;
    return () => {
      document.removeEventListener("dxw-selection", upd);
      if (iv) clearInterval(iv);
    };
  }, [api, suggesting]);

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === "f") {
        e.preventDefault();
        setFindOpen(true);
        setTimeout(() => findInput.current?.focus(), 0);
      } else if (e.key === "Escape") {
        setFindOpen(false);
      }
    };
    document.addEventListener("keydown", onKey);
    return () => document.removeEventListener("keydown", onKey);
  }, []);

  // Re-apply suggesting mode whenever a fresh api arrives (initial ?suggest=1,
  // or after a reload that recreates the editor).
  useEffect(() => {
    if (api && suggesting) api.setSuggesting(true);
  }, [api, suggesting]);

  const runFind = useCallback((q: string) => {
    setFindQ(q);
    if (!api || !q) { setFindStat(null); return; }
    const total = api.find(q);
    setFindStat({ index: total > 0 ? 1 : 0, total });
  }, [api]);

  const saveLocally = useCallback((currentApi = apiRef.current): Promise<void> => {
    if (!persistenceEnabled || !currentApi) return Promise.resolve();
    let raw: Uint8Array;
    try {
      raw = currentApi.save();
    } catch {
      setSaveState("error");
      return Promise.resolve();
    }
    setSaveState("saving");
    const previous = saveInFlight.current;
    const task = (async () => {
      if (previous) await previous;
      const compressed = await gzipBytes(raw);
      const useCompression = compressed.byteLength < raw.byteLength;
      const savedAt = Date.now();
      await writeSavedWorkspace({
        id: WORKSPACE_KEY,
        version: 1,
        ...workspaceMetaRef.current,
        savedAt,
        compression: useCompression ? "gzip" : "none",
        bytes: copyBuffer(useCompression ? compressed : raw),
      });
      setLastSaved(savedAt);
      setSaveState("saved");
    })().catch(() => {
      setSaveState("error");
    }).finally(() => {
      if (saveInFlight.current === task) saveInFlight.current = null;
    });
    saveInFlight.current = task;
    return task;
  }, [persistenceEnabled]);

  useEffect(() => {
    if (!api || !persistenceEnabled) return;
    const interval = window.setInterval(() => void saveLocally(api), AUTOSAVE_MS);
    return () => clearInterval(interval);
  }, [api, persistenceEnabled, saveLocally]);

  useEffect(() => {
    if (!persistenceEnabled) return;
    const saveShortcut = (event: KeyboardEvent) => {
      if (!(event.metaKey || event.ctrlKey) || event.key.toLowerCase() !== "s") return;
      event.preventDefault();
      void saveLocally();
    };
    document.addEventListener("keydown", saveShortcut);
    return () => document.removeEventListener("keydown", saveShortcut);
  }, [persistenceEnabled, saveLocally]);

  const onFile = useCallback(async (file: File) => {
    setStatus(`Loading ${file.name}…`);
    setPageCount(null);
    setApi(null);
    apiRef.current = null;
    try {
      const buf = await file.arrayBuffer();
      setSource(buf);
      setPreset("");
      setFileName(file.name);
    } catch (error) {
      setStatus(`Error: ${error instanceof Error ? error.message : "Could not read file"}`);
    }
  }, []);

  const loadPreset = (id: string) => {
    const next = PRESETS.find((item) => item.id === id);
    if (!next) return;
    setPreset(next.id);
    setSource(next.path);
    setFileName(next.path.split("/").pop()!);
    setPageCount(null);
    setApi(null);
    apiRef.current = null;
    setMissingFonts([]);
    setStatus(`Loading ${next.label}…`);
  };

  const download = (bytes: Uint8Array) => {
    const buf = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
    const blob = new Blob([buf], {
      type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = fileName.replace(/\.docx$/i, "") + "-edited.docx";
    a.click();
    URL.revokeObjectURL(url);
  };

  const onEditorReady = useCallback((nextApi: DocxViewApi) => {
    apiRef.current = nextApi;
    setApi(nextApi);
    void saveLocally(nextApi);
  }, [saveLocally]);

  const savedTime = lastSaved
    ? new Date(lastSaved).toLocaleTimeString([], { hour: "numeric", minute: "2-digit" })
    : null;

  return (
    <div className="app-shell">
      <header className="app-header">
        <div className="brand-lockup">
          <span className="brand-mark" aria-hidden="true">W</span>
          <span className="brand-copy">
            <strong>WordInWeb</strong>
            <span>DOCX rendering and editing, in the browser</span>
          </span>
        </div>
        <nav className="app-links" aria-label="Project links">
          <a href="https://github.com/theRealestAEP/wordinweb" target="_blank" rel="noreferrer">GitHub</a>
          <a href="https://www.aepick.me/blog" target="_blank" rel="noreferrer">Blog</a>
          <a href="/report" target="_blank" rel="noreferrer">Parity report</a>
        </nav>
      </header>
      <div className="control-bar">
        <div className="preset-control">
          <span className="control-label">Try a template</span>
          <ToolbarMenuSelect
            className="app-preset-select"
            value={preset}
            ariaLabel="Choose a document template"
            placeholder="Choose a document…"
            width={218}
            menuWidth={260}
            options={PRESETS.map((item) => ({ value: item.id, label: item.label }))}
            onChange={(value) => loadPreset(value)}
          />
        </div>
        <input
          id="docx-upload"
          className="visually-hidden"
          type="file"
          accept=".docx"
          onChange={(e) => {
            const f = e.target.files?.[0];
            if (f) void onFile(f);
          }}
        />
        <label className="upload-button" htmlFor="docx-upload">Upload .docx</label>
        <span className="control-divider" aria-hidden="true" />
        <label className="compact-control">
          <span>Zoom</span>
          <ToolbarMenuSelect
            className="app-zoom-select"
            ariaLabel="Document zoom"
            value={String(zoom)}
            width={86}
            menuWidth={112}
            options={[
              { value: "0.5", label: "50%" },
              { value: "0.75", label: "75%" },
              { value: "1", label: "100%" },
              { value: "1.25", label: "125%" },
              { value: "1.5", label: "150%" },
            ]}
            onChange={(value) => setZoom(parseFloat(value))}
          />
        </label>
        <ModeControl
          mode={!editable ? "viewing" : suggesting ? "suggesting" : "editing"}
          author={author}
          onChange={(m) => {
            if (m === "viewing") {
              setSuggesting(false);
              api?.setSuggesting(false);
              setEditable(false);
            } else {
              const wantSuggest = m === "suggesting";
              setEditable(true);
              setSuggesting(wantSuggest);
              api?.setSuggesting(wantSuggest);
            }
          }}
        />
        {pendingRevisions > 0 && (
          <span
            data-dxw-review-bar
            title="Pending suggestions (tracked changes). Click any colored suggestion in the text to accept or reject it individually."
            style={{
              display: "inline-flex",
              alignItems: "center",
              gap: 6,
              fontSize: 12.5,
              padding: "3px 6px 3px 10px",
              borderRadius: 14,
              border: "1px solid #f3c9c4",
              background: "#fce8e6",
              color: MODE_INK,
              fontWeight: 600,
              whiteSpace: "nowrap",
            }}
          >
            {pendingRevisions} suggestion{pendingRevisions === 1 ? "" : "s"}
            <button
              data-dxw-accept-all
              disabled={!editable}
              title={editable ? "Accept all suggestions (undoable)" : "Switch to Editing or Suggesting to review"}
              onClick={() => api?.acceptAllRevisions()}
              style={{
                border: "1px solid #1a73e8", background: "#1a73e8", color: "#fff", borderRadius: 12,
                padding: "2px 10px", cursor: editable ? "pointer" : "default", font: "600 12px system-ui,sans-serif",
                opacity: editable ? 1 : 0.5,
              }}
            >
              ✓ Accept all
            </button>
            <button
              data-dxw-reject-all
              disabled={!editable}
              title={editable ? "Reject all suggestions (undoable)" : "Switch to Editing or Suggesting to review"}
              onClick={() => api?.rejectAllRevisions()}
              style={{
                border: "1px solid #dadce0", background: "#fff", color: "#3c4043", borderRadius: 12,
                padding: "2px 10px", cursor: editable ? "pointer" : "default", font: "600 12px system-ui,sans-serif",
                opacity: editable ? 1 : 0.5,
              }}
            >
              ✗ Reject all
            </button>
          </span>
        )}
        <label className="compact-control author-control" title="Name stamped on your comments and suggestions">
          <span>Author</span>
          <input
            data-dxw-author
            defaultValue={author}
            onBlur={(e) => changeAuthor(e.target.value)}
            onKeyDown={(e) => { if (e.key === "Enter") (e.target as HTMLInputElement).blur(); }}
          />
        </label>
        <label className="comments-control" title="Show review comments">
          <input type="checkbox" checked={showComments} onChange={(e) => setShowComments(e.target.checked)} />
          <span>Comments</span>
        </label>
        <div className="document-meta" aria-live="polite">
          <span className={`status-dot${status.startsWith("Error") ? " error" : ""}`} aria-hidden="true" />
          <span className="file-name">{fileName}</span>
          <span className="page-count">
            {status || (pageCount !== null ? `${pageCount} page${pageCount === 1 ? "" : "s"}` : "Loading…")}
          </span>
        </div>
        {persistenceEnabled && (
          <button
            className={`save-button${saveState === "error" ? " error" : ""}`}
            disabled={!api || saveState === "saving"}
            title={saveState === "error"
              ? "Local save failed. Download a copy to keep your work."
              : `Save to this browser. Autosaves every minute${savedTime ? `; last saved ${savedTime}` : ""}.`}
            onClick={() => void saveLocally()}
          >
            {saveState === "saving" ? "Saving…" : saveState === "error" ? "Save failed" : savedTime ? `Saved ${savedTime}` : "Save"}
          </button>
        )}
        <button
          className="print-button"
          title="Print / save as PDF"
          onClick={() => {
            const root = document.querySelector(".dxw-pages") as HTMLElement | null;
            const page = root?.querySelector(".dxw-page") as HTMLElement | null;
            if (root && page) printPages(root, parseFloat(page.style.width) || 816, parseFloat(page.style.height) || 1056);
          }}
        >
          Print
        </button>
      </div>
      {editable && <div className="document-toolbar"><DocxToolbar api={api} mode={toolbarMode} features={toolbarFeatures} onSave={download} /></div>}
      {findOpen && (
        <div className="find-bar">
          <input
            ref={findInput}
            value={findQ}
            placeholder="Find in document"
            onChange={(e) => runFind(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === "Enter" && api && findStat && findStat.total > 0) {
                setFindStat({ index: api.findStep(e.shiftKey ? -1 : 1), total: findStat.total });
              }
            }}
            style={{ border: "1px solid #dadce0", borderRadius: 6, padding: "4px 8px", width: 220 }}
          />
          <span style={{ color: "#5f6368", minWidth: 54 }}>
            {findStat ? `${findStat.index} of ${findStat.total}` : ""}
          </span>
          <button onClick={() => api && findStat && setFindStat({ index: api.findStep(-1), total: findStat.total })}>↑</button>
          <button onClick={() => api && findStat && setFindStat({ index: api.findStep(1), total: findStat.total })}>↓</button>
          {editable && (
            <>
              <input
                value={replQ}
                placeholder="Replace with"
                onChange={(e) => setReplQ(e.target.value)}
                style={{ border: "1px solid #dadce0", borderRadius: 6, padding: "4px 8px", width: 180 }}
              />
              <button
                onClick={() => {
                  if (!api || !findStat) return;
                  const left = api.replaceCurrent(replQ);
                  setFindStat({ index: left > 0 ? Math.min(findStat.index, left) : 0, total: left });
                }}
              >
                Replace
              </button>
              <button
                onClick={() => {
                  if (!api) return;
                  const n = api.replaceAll(findQ, replQ);
                  setStatus(`Replaced ${n} occurrence${n === 1 ? "" : "s"}`);
                  setFindStat(null);
                }}
              >
                Replace all
              </button>
            </>
          )}
          <button onClick={() => setFindOpen(false)} style={{ marginLeft: "auto" }}>✕</button>
        </div>
      )}
      {missingFonts.length > 0 && !fontWarnDismissed && (
        <div
          data-dxw-font-warning
          style={{
            display: "flex", alignItems: "center", gap: 8, padding: "6px 12px",
            background: "#fff7e0", borderBottom: "1px solid #e8d9a0",
            font: "12.5px system-ui, sans-serif", color: "#6b5518",
          }}
        >
          <span>
            Some fonts this document asks for aren't available here, so substitutes are shown —
            layout may differ from Word: <b>{missingFonts.map((f) => f.family).join(", ")}</b>
          </span>
          <button
            onClick={() => setFontWarnDismissed(true)}
            style={{ marginLeft: "auto", border: "none", background: "none", cursor: "pointer", color: "inherit" }}
            title="Dismiss"
          >
            ✕
          </button>
        </div>
      )}
      <main className="editor-stage">
        {status.startsWith("Loading ") && (
          <div className="document-loading" data-dxw-loading="" role="status" aria-live="assertive" aria-busy="true">
            <span className="document-loading-spinner" aria-hidden="true" />
            <strong>{status}</strong>
            <span>Preparing pages for editing…</span>
          </div>
        )}
        {source && (
          <DocxView
            source={source}
            zoom={zoom}
            editable={editable}
            showComments={showComments}
            // Google-Docs display model, no toggle: Editing/Suggesting show
            // tracked changes as marks; Viewing previews the final document
            // (as if every suggestion were accepted). The review pill keeps
            // the pending count visible in every mode.
            revisions={editable ? "markup" : "final"}
            commentAuthor={author}
            style={{ height: "100%" }}
            onLoad={({ pageCount }) => {
              setPageCount(pageCount);
              setStatus("");
            }}
            onMissingFonts={(m) => {
              setMissingFonts(m);
              setFontWarnDismissed(false);
            }}
            onReady={onEditorReady}
            onError={(e) => setStatus(`Error: ${e.message}`)}
          />
        )}
      </main>
      {PERF_ON && <PerfHud />}
    </div>
  );
}

createRoot(document.getElementById("root")!).render(<App />);
