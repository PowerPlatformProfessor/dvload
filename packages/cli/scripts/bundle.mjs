// Bundle the compiled CLI (dist/) into a single CommonJS file suitable for
// Node's Single Executable Application (SEA) pipeline. All dependencies are
// pure JS since keytar was replaced with the DPAPI secure store, so the
// whole CLI fits in one file with no native addons.
//
// Usage:  npm run build && npm run bundle
// Output: build/dvload.cjs  (then see release.yml for the .exe steps)

import { build } from "esbuild";

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
  logLevel: "info",
});
