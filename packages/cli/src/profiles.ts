// Named environment profiles stored in ~/.dvload/profiles.json.
// A profile is just a friendly name → environment URL alias, so you can
// write `dvload login --profile prod` instead of `dvload login --env https://…`.

import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";

export const PROFILES_FILE = path.join(os.homedir(), ".dvload", "profiles.json");

/** Map of profile name → environment URL. */
export type Profiles = Record<string, string>;

export async function readProfiles(): Promise<Profiles> {
  try {
    const raw = await fs.readFile(PROFILES_FILE, "utf8");
    return JSON.parse(raw) as Profiles;
  } catch {
    return {};
  }
}

export async function writeProfiles(profiles: Profiles): Promise<void> {
  await fs.mkdir(path.dirname(PROFILES_FILE), { recursive: true });
  await fs.writeFile(PROFILES_FILE, JSON.stringify(profiles, null, 2) + "\n", "utf8");
}

/**
 * Resolve the environment URL from either `--env <url>` or `--profile <name>`.
 * Throws a clear error if neither is provided or the profile doesn't exist.
 */
export async function resolveEnv(opts: { env?: string; profile?: string }): Promise<string> {
  if (opts.env) return opts.env;
  if (opts.profile) {
    const profiles = await readProfiles();
    const url = profiles[opts.profile];
    if (!url) {
      const names = Object.keys(profiles);
      const hint = names.length
        ? `Available: ${names.join(", ")}`
        : "No profiles saved yet — run `dvload profile add <name> <url>`.";
      throw new Error(`Profile "${opts.profile}" not found. ${hint}`);
    }
    return url;
  }
  throw new Error("Provide --env <url> or --profile <name>.");
}
