import { access } from "node:fs/promises";
import path from "node:path";
import { spawn } from "node:child_process";

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
    await runNpm(repoRoot, ["--workspace=@dvload/addin", "run", "start"]);
    return;
  }
  if (action === "stop") {
    await runNpm(repoRoot, ["--workspace=@dvload/addin", "run", "stop"]);
    return;
  }
  if (action === "dev") {
    await runNpm(repoRoot, ["run", "dev:addin"]);
    return;
  }

  throw new Error(`Unknown addin action "${action}". Use: start, stop, or dev.`);
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
    const npmCmd = process.platform === "win32" ? "npm.cmd" : "npm";
    const child = spawn(npmCmd, args, {
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
