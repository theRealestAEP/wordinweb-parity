/**
 * Which wordinweb build this checkout actually loads.
 *
 * The demo imports wordinweb from apps/demo/src/main.tsx, so Node and vite
 * resolve apps/demo/node_modules/wordinweb before the root one. A real package
 * installed there shadows the root link, and a run then measures an engine
 * nobody selected — silently, because vite only re-reads the package when it
 * re-optimizes dependencies.
 *
 * scripts/use-engine.mjs selects an engine and scripts/edit-roundtrip-parity.mjs
 * records the one it measured. They share this module so they cannot disagree
 * about what "the engine" means.
 */

import { execFileSync } from "node:child_process";
import { existsSync, lstatSync, readFileSync, realpathSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

export const repoRoot = join(dirname(fileURLToPath(import.meta.url)), "..");

export function git(cwd, ...args) {
  try {
    return execFileSync("git", ["-C", cwd, ...args], { encoding: "utf8" }).trim();
  } catch {
    return null;
  }
}

/** First `node_modules/<name>` walking up from `startDir` — how a bare specifier resolves. */
export function resolvePackageDir(name, startDir) {
  let dir = startDir;
  for (;;) {
    const candidate = join(dir, "node_modules", name);
    if (existsSync(candidate)) return candidate;
    const parent = dirname(dir);
    if (parent === dir) return null;
    dir = parent;
  }
}

export function packageProvenance(packageDir) {
  if (!packageDir || !existsSync(join(packageDir, "package.json"))) return null;
  const symlink = lstatSync(packageDir).isSymbolicLink();
  const target = realpathSync(packageDir);
  const manifest = JSON.parse(readFileSync(join(packageDir, "package.json"), "utf8"));
  // Git provenance only means something when the target is a checkout OF THE
  // ENGINE. A published tarball unpacked under this repo answers git with THIS
  // repo's SHA, which would name a commit unrelated to the engine.
  const toplevel = git(target, "rev-parse", "--show-toplevel");
  const ownCheckout = toplevel !== null && toplevel !== realpathSync(repoRoot);
  const status = ownCheckout ? git(target, "status", "--porcelain") : null;
  return {
    packageDir,
    version: manifest.version,
    symlink,
    target,
    installedCopy: !ownCheckout,
    targetGitSha: ownCheckout ? git(target, "rev-parse", "HEAD") : null,
    targetGitBranch: ownCheckout ? git(target, "rev-parse", "--abbrev-ref", "HEAD") : null,
    targetGitDirty: status === null ? null : status !== "",
  };
}

/**
 * Do two locations hold the same engine?
 *
 * The same path is the obvious yes. Two SEPARATE installs of one published
 * version are also yes: a published version is immutable, so which copy loads
 * cannot change behavior. Only a difference that can change behavior — a
 * different version, or a checkout versus anything else — counts as a shadow.
 */
function sameEngine(a, b) {
  if (a.target === b.target) return true;
  return a.installedCopy && b.installedCopy && a.version === b.version;
}

/**
 * The engine the demo resolves, plus the root link when the two differ.
 * `shadowed` means the demo does NOT load what the root link names.
 */
export function wordinwebBuild() {
  const resolved = packageProvenance(resolvePackageDir("wordinweb", join(repoRoot, "apps/demo")));
  const rootLink = packageProvenance(join(repoRoot, "node_modules/wordinweb"));
  const effective = resolved ?? rootLink;
  if (!effective) return { version: null, shadowed: false, resolutionFailed: true };
  const shadowed = Boolean(resolved && rootLink && !sameEngine(resolved, rootLink));
  return { ...effective, shadowed, shadowedRootLink: shadowed ? rootLink : null };
}

export function describeBuild(build) {
  if (build.resolutionFailed) return "wordinweb: NOT RESOLVED";
  const where = build.installedCopy ? "installed copy" : `${build.targetGitBranch} @ ${build.targetGitSha?.slice(0, 12)}`;
  const dirty = build.targetGitDirty ? " (dirty)" : "";
  return `wordinweb ${build.version} — ${build.target}\n  ${where}${dirty}`;
}
