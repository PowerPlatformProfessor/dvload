#!/usr/bin/env node
// Stamp a released version + artifact hash into the winget and Scoop
// manifests.
//
// Both package managers need the sha256 of the exact published file, which
// only exists after release.yml has run, so the manifests necessarily carry
// placeholders in git. Doing that update by hand across four files is how
// you end up shipping a manifest whose hash belongs to the previous version
// — winget rejects it, Scoop installs it and then fails verification.
//
// Usage:
//   node packaging/stamp-release.mjs 0.2.0
//   node packaging/stamp-release.mjs 0.2.0 <sha256>
//
// With no hash, the checksum is fetched from the release's dvload.exe.sha256
// asset, which release.yml uploads alongside the exe.

import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const OWNER = "PowerPlatformProfessor";
const REPO = "dvload";

const [, , rawVersion, rawHash] = process.argv;

if (!rawVersion) {
  console.error("usage: node packaging/stamp-release.mjs <version> [sha256]");
  process.exit(2);
}

// Accept "v0.2.0" as well — that is what the tag looks like, and being
// strict about it here just invites a typo'd manifest.
const version = rawVersion.replace(/^v/, "");
if (!/^\d+\.\d+\.\d+(-[\w.]+)?$/.test(version)) {
  console.error(`Not a semver version: "${rawVersion}"`);
  process.exit(2);
}

const hash = (rawHash ?? (await fetchHash(version))).trim().toLowerCase();
if (!/^[0-9a-f]{64}$/.test(hash)) {
  console.error(`Not a sha256: "${hash}"`);
  process.exit(2);
}

async function fetchHash(v) {
  const url = `https://github.com/${OWNER}/${REPO}/releases/download/v${v}/dvload.exe.sha256`;
  console.log(`  fetching ${url}`);
  const res = await fetch(url, { redirect: "follow" });
  if (!res.ok) {
    console.error(
      `Could not fetch the checksum (HTTP ${res.status}).\n` +
        `Is the v${v} release published? Otherwise pass the hash as the second argument.`
    );
    process.exit(1);
  }
  // Format is "<hash>  dvload.exe".
  return (await res.text()).split(/\s+/)[0];
}

async function edit(rel, fn) {
  const file = path.join(here, rel);
  const before = await readFile(file, "utf8");
  const after = fn(before);
  if (before === after) {
    console.warn(`  unchanged   ${rel}  (already at ${version}?)`);
    return;
  }
  await writeFile(file, after);
  console.log(`  stamped     ${rel}`);
}

const SHA_PLACEHOLDER = /\b[0-9a-f]{64}\b/g;
const ANY_VERSION = /\d+\.\d+\.\d+(?:-[\w.]+)?/g;

await edit("scoop/dvload.json", (s) => {
  const j = JSON.parse(s);
  j.version = version;
  j.architecture["64bit"].url = j.architecture["64bit"].url.replace(ANY_VERSION, version);
  j.architecture["64bit"].hash = hash;
  return JSON.stringify(j, null, 2) + "\n";
});

// YAML edited as text rather than parsed: winget manifests are hand-reviewed
// in the winget-pkgs PR, and a round-trip through a YAML library would
// reflow the comments and block scalars that make them readable.
for (const rel of [
  "winget/PowerPlatformProfessor.dvload.yaml",
  "winget/PowerPlatformProfessor.dvload.locale.en-US.yaml",
]) {
  await edit(rel, (s) => s.replace(/^PackageVersion: .*$/m, `PackageVersion: ${version}`));
}

await edit("winget/PowerPlatformProfessor.dvload.installer.yaml", (s) =>
  s
    .replace(/^PackageVersion: .*$/m, `PackageVersion: ${version}`)
    .replace(/^ReleaseDate: .*$/m, `ReleaseDate: ${new Date().toISOString().slice(0, 10)}`)
    .replace(/(InstallerUrl: .*)/, (m) => m.replace(ANY_VERSION, version))
    .replace(SHA_PLACEHOLDER, hash)
);

console.log(`\nStamped v${version} (${hash.slice(0, 12)}…). Next:`);
console.log("  winget validate --manifest packaging/winget");
console.log(`  wingetcreate submit packaging/winget  # or open the winget-pkgs PR by hand`);
console.log("  copy packaging/scoop/dvload.json into your Scoop bucket");
