// Secret storage without native modules.
//
// Replaces keytar (archived/unmaintained, and its native .node binary was
// both a supply-chain risk and the main `npm install` failure mode on
// locked-down machines).
//
// Layout: one JSON map { account -> secret } persisted at:
//   - Windows: ~/.dvload/secrets.dat — DPAPI-protected (CurrentUser scope)
//     via PowerShell's [Security.Cryptography.ProtectedData]. Same
//     protection class as Windows Credential Manager: only the same user
//     on the same machine can decrypt.
//   - POSIX:   ~/.dvload/secrets.json — plaintext, mode 0600.
//
// The map is loaded lazily once per process and rewritten on mutation.

import { spawn } from "node:child_process";
import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";

const DIR = path.join(os.homedir(), ".dvload");
const WIN_FILE = path.join(DIR, "secrets.dat");
const POSIX_FILE = path.join(DIR, "secrets.json");

const isWindows = process.platform === "win32";

/**
 * Absolute path, not "powershell.exe": this child receives every secret over
 * stdin, so it must not be resolvable through a user-writable PATH entry.
 */
const POWERSHELL_EXE = path.join(
  process.env.SystemRoot ?? "C:\\Windows",
  "System32",
  "WindowsPowerShell",
  "v1.0",
  "powershell.exe"
);

/* -------------------------------------------------------------------------- */
/* DPAPI via PowerShell                                                        */
/* -------------------------------------------------------------------------- */

/**
 * Round-trip bytes through Windows DPAPI (CurrentUser scope) using a
 * PowerShell child process. Base64 over stdin/stdout avoids all quoting
 * issues. ~200ms per call, but we only call it once per process (load)
 * plus once per mutation (save) — never in a hot path.
 */
function dpapi(mode: "protect" | "unprotect", data: Buffer): Promise<Buffer> {
  const op = mode === "protect" ? "Protect" : "Unprotect";
  const script =
    `$ErrorActionPreference='Stop';` +
    `Add-Type -AssemblyName System.Security;` +
    `$in=[Convert]::FromBase64String([Console]::In.ReadToEnd());` +
    `$out=[Security.Cryptography.ProtectedData]::${op}($in,$null,'CurrentUser');` +
    `[Console]::Out.Write([Convert]::ToBase64String($out))`;
  return new Promise((resolve, reject) => {
    const child = spawn(
      POWERSHELL_EXE,
      ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", script],
      { windowsHide: true }
    );
    let stdout = "";
    let stderr = "";
    child.stdout.on("data", (d: Buffer) => (stdout += d.toString()));
    child.stderr.on("data", (d: Buffer) => (stderr += d.toString()));
    child.on("error", reject);
    child.on("exit", (code) => {
      if (code === 0) resolve(Buffer.from(stdout.trim(), "base64"));
      else reject(new Error(`DPAPI ${mode} failed (exit ${code}): ${stderr.trim()}`));
    });
    child.stdin.write(data.toString("base64"));
    child.stdin.end();
  });
}

/** Protect bytes for at-rest storage (DPAPI on Windows, identity elsewhere). */
export async function protectBytes(data: Buffer): Promise<Buffer> {
  return isWindows ? dpapi("protect", data) : data;
}

/** Inverse of protectBytes. */
export async function unprotectBytes(data: Buffer): Promise<Buffer> {
  return isWindows ? dpapi("unprotect", data) : data;
}

/* -------------------------------------------------------------------------- */
/* Store                                                                       */
/* -------------------------------------------------------------------------- */

type SecretMap = Record<string, string>;

let cache: SecretMap | null = null;

async function load(): Promise<SecretMap> {
  if (cache) return cache;
  const file = isWindows ? WIN_FILE : POSIX_FILE;
  try {
    const raw = await fs.readFile(file);
    const plain = await unprotectBytes(raw);
    cache = JSON.parse(plain.toString("utf8")) as SecretMap;
  } catch (e: unknown) {
    if ((e as NodeJS.ErrnoException).code === "ENOENT") {
      cache = {};
    } else {
      throw new Error(
        `Could not read the dvload secret store (${file}). ` +
          `If it is corrupt, delete it and run \`dvload login\` / \`dvload app-login\` again. ` +
          `Underlying error: ${e instanceof Error ? e.message : String(e)}`,
        { cause: e }
      );
    }
  }
  return cache;
}

async function save(map: SecretMap): Promise<void> {
  await fs.mkdir(DIR, { recursive: true });
  const plain = Buffer.from(JSON.stringify(map), "utf8");
  const file = isWindows ? WIN_FILE : POSIX_FILE;
  const data = await protectBytes(plain);
  // Write-then-rename: this file is the only copy of every session and
  // credential, so a crash mid-write must not be able to corrupt it.
  const tmp = `${file}.tmp`;
  await fs.writeFile(tmp, data, { mode: 0o600 });
  await fs.rename(tmp, file);
  cache = map;
}

export async function getSecret(account: string): Promise<string | null> {
  const map = await load();
  return map[account] ?? null;
}

export async function setSecret(account: string, value: string): Promise<void> {
  const map = { ...(await load()) };
  map[account] = value;
  await save(map);
}

export async function deleteSecret(account: string): Promise<void> {
  const map = { ...(await load()) };
  if (!(account in map)) return;
  delete map[account];
  await save(map);
}

/** All account keys currently stored (optionally filtered by prefix). */
export async function listAccounts(prefix?: string): Promise<string[]> {
  const map = await load();
  const keys = Object.keys(map);
  return prefix ? keys.filter((k) => k.startsWith(prefix)) : keys;
}

/** Test hook: drop the in-process cache so the next call re-reads disk. */
export function __resetSecretStoreCache(): void {
  cache = null;
}
