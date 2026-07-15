import { defineConfig } from "@playwright/test";

// Port is overridable so a git worktree can run e2e against its own dev server
// (each worktree edits its own source) instead of the main checkout's :5299.
const PORT = process.env.DXW_E2E_PORT ?? "5299";

export default defineConfig({
  testDir: "e2e",
  timeout: 30_000,
  retries: 1,
  use: {
    baseURL: `http://localhost:${PORT}`,
    viewport: { width: 1400, height: 1000 },
    permissions: ["clipboard-read", "clipboard-write"],
  },
  webServer: {
    command: `npm run dev -w demo -- --port ${PORT} --strictPort`,
    url: `http://localhost:${PORT}`,
    reuseExistingServer: true,
    timeout: 30_000,
  },
});
