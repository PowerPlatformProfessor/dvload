/**
 * The CLI tests spawn the compiled entry point (`dist/index.js`) as a real
 * child process — which is the point: they assert on exit codes, stderr, and
 * argument parsing exactly as a user experiences them.
 *
 * That makes them depend on a build step, and a fresh `git clone && npm test`
 * would otherwise fail with `Cannot find module .../dist/index.js` — a
 * confusing first impression for a new contributor and a support burden for
 * whoever triages the issue.
 *
 * So: build if missing or stale, once, before the CLI project runs.
 */

import { execFileSync } from "node:child_process";
import { statSync, readdirSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const cliRoot = path.resolve(here, "..");
const repoRoot = path.resolve(cliRoot, "..", "..");

/** Newest mtime under a directory tree, or 0 if it doesn't exist. */
function newestMtime(dir: string): number {
  let entries;
  try {
    entries = readdirSync(dir, { withFileTypes: true, encoding: "utf8" });
  } catch {
    return 0; // not built yet
  }
  let newest = 0;
  for (const e of entries) {
    const full = path.join(dir, e.name);
    newest = Math.max(newest, e.isDirectory() ? newestMtime(full) : statSync(full).mtimeMs);
  }
  return newest;
}

function build(workspace: string): void {
  execFileSync("npm", ["run", "build", "--workspace", workspace], {
    cwd: repoRoot,
    stdio: "inherit",
    shell: process.platform === "win32",
  });
}

export async function setup(): Promise<void> {
  // core first: the CLI imports @dvload/core's compiled output.
  const coreSrc = newestMtime(path.join(repoRoot, "packages", "core", "src"));
  const coreDist = newestMtime(path.join(repoRoot, "packages", "core", "dist"));
  if (coreDist === 0 || coreSrc > coreDist) build("@dvload/core");

  const cliSrc = newestMtime(path.join(cliRoot, "src"));
  const cliDist = newestMtime(path.join(cliRoot, "dist"));
  if (cliDist === 0 || cliSrc > cliDist) build("dvload");
}
