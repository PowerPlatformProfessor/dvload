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
// Usage:  npm run build && npm run bundle
// Output: build/dvload.cjs + build/web/  (then see release.yml for the .exe steps)

import { build } from "esbuild";
import { cp, access, rm } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const cliRoot = path.resolve(here, "..");
const addinDist = path.resolve(cliRoot, "..", "addin", "dist");
const webOut = path.join(cliRoot, "build", "web");

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

try {
  await access(path.join(addinDist, "taskpane.html"));
  await rm(webOut, { recursive: true, force: true });
  await cp(addinDist, webOut, { recursive: true });
  console.log(`  bundled UI  ${path.relative(cliRoot, webOut)}`);
} catch {
  // A CLI-only bundle is legitimate (CI builds the workspaces separately),
  // but a release that quietly ships without a UI would strand every add-in
  // user, so this has to be loud rather than silent.
  console.warn(
    "\n  WARNING: packages/addin/dist not found — build/web will be missing.\n" +
      "  `dvload serve` and `dvload gui` will fail with a 'could not find the\n" +
      "  built add-in UI' error. Run `npm run build --workspace=@dvload/addin`\n" +
      "  before bundling a release.\n"
  );
}
