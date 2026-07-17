import { defineConfig, searchForWorkspaceRoot, type Plugin, type ViteDevServer } from "vite";
import react from "@vitejs/plugin-react";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import { readFileSync, realpathSync } from "node:fs";

// In a git worktree, node_modules is a symlink into the main checkout; vite
// resolves assets (e.g. @fontsource woff2) to their REAL path and serves them
// via /@fs/, which its fs allowlist rejects unless the real location is
// permitted. Without this the Carlito/Caladea metric fonts 403 and canvas
// text measurement silently falls back to sans-serif (~8% wide).
const realNodeModules = (() => {
  try {
    return [realpathSync(fileURLToPath(new URL("../../node_modules", import.meta.url)))];
  } catch {
    return [];
  }
})();

// Serve the tracked parity dashboard and its diff images at /report/.
function parityReport(): Plugin {
  const reportPath = fileURLToPath(new URL("./public/report.html", import.meta.url));
  const placeholder = `<!doctype html>
<html lang="en"><head><meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>WordInWeb parity — no results yet</title>
<style>
  :root { color-scheme: light dark; }
  body { margin:0; min-height:100vh; display:grid; place-items:center;
    font:15px/1.6 system-ui,sans-serif; background:#f6f7f9; color:#1f2328; }
  @media (prefers-color-scheme: dark){ body{ background:#0d1117; color:#e6edf3; } .card{ background:#161b22 !important; border-color:#30363d !important; } code{ background:#21262d !important; } }
  .card { max-width:560px; padding:32px 36px; background:#fff; border:1px solid #e2e4e8;
    border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,.06); }
  h1 { margin:0 0 6px; font-size:20px; }
  p { margin:10px 0; color:inherit; opacity:.85; }
  code { display:block; padding:10px 12px; margin:10px 0; border-radius:8px;
    background:#f1f3f5; font:13px ui-monospace,SFMono-Regular,Menlo,monospace; white-space:pre-wrap; }
  a { color:#2a78d6; }
</style></head>
<body><div class="card">
  <h1>No parity results yet</h1>
  <p>The pixel-parity dashboard is generated from a full comparison run against
  the Word reference PDFs. Nothing has been run in this checkout yet.</p>
  <p>Generate it, then reload this page:</p>
  <code>node scripts/parity-parallel.mjs
node scripts/parity-render-report.mjs</code>
  <p><a href="/">← back to the viewer</a></p>
</div></body></html>`;

  const handler = (req: { url?: string }, res: import("node:http").ServerResponse, next: () => void): void => {
    const path = (req.url ?? "").split("?")[0];
    // The report's diff-image links are relative siblings (fixture-pN.png,
    // candidate-*.png) living next to report.html in parity/out — serve them
    // under /report/ so the links resolve. /report redirects to /report/ so
    // relative URLs base correctly.
    if (path === "/report" || path === "/report.html") {
      res.statusCode = 302;
      res.setHeader("Location", "/report/");
      return res.end();
    }
    if (!path.startsWith("/report/")) return next();
    const name = decodeURIComponent(path.slice("/report/".length)) || "report.html";
    // Sibling files only — no path traversal.
    if (name.includes("/") || name.includes("..")) return next();
    const file = name === "" || name === "report.html" ? reportPath : join(dirname(reportPath), name);
    try {
      const body = readFileSync(file);
      res.setHeader(
        "Content-Type",
        file.endsWith(".png") ? "image/png" : file.endsWith(".json") ? "application/json" : "text/html; charset=utf-8",
      );
      res.end(body);
    } catch {
      if (name === "" || name === "report.html") {
        res.statusCode = 200;
        res.setHeader("Content-Type", "text/html; charset=utf-8");
        res.end(placeholder);
      } else {
        next();
      }
    }
  };

  return {
    name: "dxw-parity-report",
    configureServer(server: ViteDevServer) {
      server.middlewares.use(handler);
    },
    configurePreviewServer(server) {
      server.middlewares.use(handler);
    },
  };
}

export default defineConfig({
  plugins: [react(), parityReport()],
  server: {
    fs: { allow: [searchForWorkspaceRoot(process.cwd()), ...realNodeModules] },
  },
});
