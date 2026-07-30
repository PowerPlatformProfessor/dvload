// Bundle the compiled CLI (dist/) into a single CommonJS file suitable for
// Node's Single Executable Application (SEA) pipeline. All dependencies are
// pure JS since keytar was replaced with the DPAPI secure store, so the
// whole CLI fits in one file with no native addons.
//
// Also stages the built add-in UI into build/web, because `dvload serve`
// serves it. The UI is no longer deployed to a public origin — it is served
// from loopback by this CLI, which is what lets it use the CLI's auth
// instead of needing its own Entra app registration. So the two ship
// together or not at all.
//
// Finally, writes build/sea-config.json listing every staged UI file as a
// SEA asset. This is generated rather than checked in because the asset map
// has to name each file individually and the add-in's output filenames
// change with its build; a hand-maintained config would drift the moment
// webpack emitted a new chunk, and the failure mode is a released exe whose
// task pane 404s.
//
// Usage:  npm run build && npm run bundle [--require-web]
//         --require-web turns a missing UI from a warning into an error.
// Output: build/dvload.cjs + build/web/ + build/sea-config.json
//         (then see release.yml for the .exe steps)

import { build } from "esbuild";
import { cp, access, rm, readdir, writeFile, stat } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const cliRoot = path.resolve(here, "..");
const addinDist = path.resolve(cliRoot, "..", "addin", "dist");
const buildDir = path.join(cliRoot, "build");
const webOut = path.join(buildDir, "web");
const seaConfigOut = path.join(buildDir, "sea-config.json");
const requireWeb = process.argv.includes("--require-web");

/**
 * TypeScript declarations land in the add-in's dist alongside the browser
 * bundle. They are useless to a browser and would otherwise be embedded in
 * the exe and served over loopback, so they are dropped on the way in.
 */
function isServable(rel) {
  return !rel.endsWith(".d.ts") && !rel.endsWith(".d.ts.map");
}

async function listFiles(dir, prefix = "") {
  const out = [];
  for (const entry of await readdir(dir, { withFileTypes: true })) {
    const rel = prefix ? `${prefix}/${entry.name}` : entry.name;
    if (entry.isDirectory()) out.push(...(await listFiles(path.join(dir, entry.name), rel)));
    else out.push(rel);
  }
  return out;
}

await build({
  entryPoints: ["dist/index.js"],
  bundle: true,
  platform: "node",
  target: "node20",
  format: "cjs",
  outfile: "build/dvload.cjs",
  // dist/ is ESM; import.meta.url doesn't exist in CJS. Recreate it from
  // __filename so schedule.ts's entry-point resolution keeps working.
  define: { "import.meta.url": "__importMetaUrl" },
  banner: {
    js: "const __importMetaUrl = require('node:url').pathToFileURL(__filename).href;",
  },
  // Optional at runtime: only reached when `dvload serve` needs to mint a
  // localhost certificate and none is already installed. Bundling it would
  // pull a dev-only dependency (and its native-ish cert tooling) into the
  // single-file build for a path most users never hit.
  external: ["office-addin-dev-certs"],
  logLevel: "info",
});

let webFiles = [];
try {
  await access(path.join(addinDist, "taskpane.html"));
  await rm(webOut, { recursive: true, force: true });
  await cp(addinDist, webOut, { recursive: true, filter: (src) => isServable(src) });
  webFiles = (await listFiles(webOut)).filter(isServable).sort();
  console.log(`  bundled UI  ${path.relative(cliRoot, webOut)}  (${webFiles.length} files)`);
} catch (err) {
  // A CLI-only bundle is legitimate (CI builds the workspaces separately),
  // but a release that quietly ships without a UI would strand every add-in
  // user, so this has to be loud rather than silent.
  const message =
    "packages/addin/dist not found — build/web will be missing.\n" +
    "  `dvload serve` and `dvload gui` will fail with a 'could not find the\n" +
    "  built add-in UI' error. Run `npm run build --workspace=@dvload/addin`\n" +
    "  before bundling a release.";
  if (requireWeb) throw new Error(`${message}\n  (--require-web was set, so this is fatal.)\n\n${err}`);
  console.warn(`\n  WARNING: ${message}\n`);
}

// Asset keys are `web/<posix path>`, matching resolveWebSource() in
// commands/serve.ts. Paths are relative to the working directory the SEA
// config is invoked from (packages/cli), which is how Node resolves them.
const assets = Object.fromEntries(webFiles.map((rel) => [`web/${rel}`, `build/web/${rel}`]));

await writeFile(
  seaConfigOut,
  JSON.stringify(
    {
      main: "build/dvload.cjs",
      output: "build/sea-prep.blob",
      disableExperimentalSEAWarning: true,
      // Snapshots and assets are mutually exclusive in the SEA pipeline, and
      // the UI matters more than the few ms a snapshot would save on start.
      useSnapshot: false,
      assets,
    },
    null,
    2
  ) + "\n"
);

const totalBytes = (
  await Promise.all(webFiles.map(async (rel) => (await stat(path.join(webOut, rel))).size))
).reduce((a, b) => a + b, 0);

console.log(
  `  sea config  ${path.relative(cliRoot, seaConfigOut)}  ` +
    `(${webFiles.length} embedded assets, ${(totalBytes / 1024).toFixed(0)} KB)`
);
