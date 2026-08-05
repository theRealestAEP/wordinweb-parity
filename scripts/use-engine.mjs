#!/usr/bin/env node

/**
 * Select the wordinweb engine this checkout measures.
 *
 * The demo resolves wordinweb from apps/demo/node_modules before the root one,
 * so setting only the root link leaves the demo loading whatever else is
 * installed. That state is silent: vite serves a pre-bundled copy and only
 * notices the package changed when it re-optimizes, so a dev-server restart can
 * swap the engine under a run with nothing in the output to say so. This script
 * sets BOTH locations together, so there is one answer to "which engine".
 *
 * Usage:
 *   node scripts/use-engine.mjs ../wordinweb-likeoffice/packages/react
 *   node scripts/use-engine.mjs npm:0.1.22
 *
 * Clearing vite's dep cache is part of the switch: without it the next dev
 * server keeps serving the previous engine's pre-bundle.
 */

import { execFileSync } from "node:child_process";
import {
  cpSync,
  existsSync,
  lstatSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  readdirSync,
  realpathSync,
  rmSync,
  symlinkSync,
} from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { describeBuild, repoRoot, wordinwebBuild } from "./engine-provenance.mjs";

const linkDirs = [join(repoRoot, "node_modules/wordinweb"), join(repoRoot, "apps/demo/node_modules/wordinweb")];
const viteCache = join(repoRoot, "apps/demo/node_modules/.vite");

const spec = process.argv[2];
if (!spec || spec === "--help") {
  console.log("Usage: node scripts/use-engine.mjs <path-to-engine-react-pkg | npm:x.y.z>");
  console.log("  node scripts/use-engine.mjs ../wordinweb-likeoffice/packages/react");
  console.log("  node scripts/use-engine.mjs npm:0.1.22");
  console.log(`\nCurrently:\n${describeBuild(wordinwebBuild())}`);
  process.exit(spec ? 0 : 1);
}

/** Replace whatever is at `path` — link, directory or nothing. */
function clear(path) {
  rmSync(path, { recursive: true, force: true });
  mkdirSync(dirname(path), { recursive: true });
}

if (spec.startsWith("npm:")) {
  const version = spec.slice("npm:".length);
  // Fetch the tarball and unpack it by hand rather than `npm install --prefix`.
  // An install at the repo prefix resolves the WHOLE tree, which quietly
  // upgraded @playwright/test here and left the gate unable to launch a browser.
  // Selecting an engine must not touch any other package.
  const stage = mkdtempSync(join(tmpdir(), "use-engine-"));
  try {
    console.log(`Fetching wordinweb@${version}`);
    execFileSync("npm", ["pack", `wordinweb@${version}`, "--pack-destination", stage], { stdio: "inherit" });
    const tarball = readdirSync(stage).find((name) => name.endsWith(".tgz"));
    if (!tarball) throw new Error(`npm pack produced no tarball for wordinweb@${version}`);
    execFileSync("tar", ["-xzf", join(stage, tarball), "-C", stage]);
    for (const dir of linkDirs) {
      clear(dir);
      cpSync(join(stage, "package"), dir, { recursive: true });
      console.log(`Installed wordinweb@${version} at ${dir}`);
    }
  } finally {
    rmSync(stage, { recursive: true, force: true });
  }
} else {
  const enginePath = realpathSync(resolve(spec));
  const manifestPath = join(enginePath, "package.json");
  if (!existsSync(manifestPath)) throw new Error(`No package.json at ${enginePath}`);
  const manifest = JSON.parse(readFileSync(manifestPath, "utf8"));
  if (manifest.name !== "wordinweb") throw new Error(`${enginePath} is ${manifest.name}, not wordinweb`);
  // A stale or missing build is the failure this script exists to make loud:
  // linking a package whose dist predates its source measures the old code.
  if (!existsSync(join(enginePath, "dist/index.js"))) {
    throw new Error(`${enginePath} has no dist/index.js — build it first (npm run build)`);
  }
  for (const dir of linkDirs) {
    clear(dir);
    symlinkSync(enginePath, dir);
    console.log(`Linked ${dir} -> ${enginePath}`);
  }
}

rmSync(viteCache, { recursive: true, force: true });
console.log(`Cleared ${viteCache}`);

const build = wordinwebBuild();
console.log(`\n${describeBuild(build)}`);
if (build.shadowed) {
  throw new Error(
    `Still shadowed after switching: demo loads ${build.target}, root link points at ${build.shadowedRootLink.target}`,
  );
}
for (const dir of linkDirs) {
  const kind = lstatSync(dir).isSymbolicLink() ? "symlink" : "directory";
  console.log(`  ${kind.padEnd(9)} ${dir}`);
}
console.log("\nRestart the demo dev server so vite re-optimizes against this engine.");
