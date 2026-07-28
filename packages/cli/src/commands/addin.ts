// Sideload helpers for the Excel add-in.
//
// Note the division of labour with `dvload serve`: this command registers the
// manifest with Excel, `serve` supplies the pane itself and its tokens. The
// manifest's SourceLocation points at the serve port, so sideloading alone is
// not enough — hence the reminder printed below rather than a silent blank
// pane.

import { access } from "node:fs/promises";
import path from "node:path";
import { spawn } from "node:child_process";
import kleur from "kleur";

import { DEFAULT_PORT } from "./serve.js";

interface AddinOpts {
  withDevServer?: boolean;
}

export async function addinCommand(action: string, opts: AddinOpts): Promise<void> {
  const repoRoot = await findRepoRoot(process.cwd());
  if (!repoRoot) {
    throw new Error(
      "Could not find the dvload repo root (expected packages/addin/manifest.xml). Run this command from the repository."
    );
  }

  if (action === "start") {
    if (opts.withDevServer) {
      await runNpm(repoRoot, ["run", "dev:addin"]);
      return;
    }
    noteServeRequirement();
    await runNpm(repoRoot, ["--workspace=@dvload/addin", "run", "start"]);
    return;
  }
  if (action === "stop") {
    await runNpm(repoRoot, ["--workspace=@dvload/addin", "run", "stop"]);
    return;
  }
  if (action === "dev") {
    // Rebuild-on-change only. The pane is served by `dvload serve` out of
    // packages/addin/dist, so a watch build is the whole dev loop — the
    // webpack dev server would serve the UI without an /api to talk to.
    noteServeRequirement();
    await runNpm(repoRoot, ["run", "watch:addin"]);
    return;
  }

  throw new Error(`Unknown addin action "${action}". Use: start, stop, or dev.`);
}

function noteServeRequirement(): void {
  console.log(
    kleur.yellow(
      `The pane loads from https://localhost:${DEFAULT_PORT}. Run \`dvload serve\` in another ` +
        `terminal if it isn't already running, or the task pane will come up blank.`
    )
  );
}

async function findRepoRoot(start: string): Promise<string | null> {
  let dir = path.resolve(start);
  while (true) {
    const probe = path.join(dir, "packages", "addin", "manifest.xml");
    try {
      await access(probe);
      return dir;
    } catch {
      // keep walking up
    }
    const parent = path.dirname(dir);
    if (parent === dir) return null;
    dir = parent;
  }
}

function runNpm(cwd: string, args: string[]): Promise<void> {
  return new Promise((resolve, reject) => {
    const npmExec = process.env.npm_execpath;
    const command = npmExec ? process.execPath : process.platform === "win32" ? "npm.cmd" : "npm";
    const commandArgs = npmExec ? [npmExec, ...args] : args;
    const child = spawn(command, commandArgs, {
      cwd,
      stdio: "inherit",
    });
    child.on("error", reject);
    child.on("exit", (code) => {
      if (code === 0) resolve();
      else reject(new Error(`Command failed: npm ${args.join(" ")}`));
    });
  });
}
