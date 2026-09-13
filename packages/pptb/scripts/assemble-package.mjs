// Assemble the publishable ToolBox package.
//
// The registry's structure_validation expects a package that CONTAINS the
// built output — `package.json` at the root, `dist/index.html` beside it —
// which is the same shape Load Local Tool wants from a project folder.
// Publishing the contents of dist/ as the package root (what this repo used
// to do) is rejected with "dist folder is required but not found".
//
// So the publishable artifact is assembled here rather than being dist/
// itself:
//
//   publish/
//     package.json   <- tool.package.json, the PPTB manifest
//     README.md      <- shown on the tool's marketplace page
//     dist/          <- webpack output: index.html, bundle, icons/
//
// `main` ("index.html") and `icon` ("icons/icon.svg") stay relative to dist/,
// which is where the ToolBox resolves them from.
//
// Usage: npm run package --workspace=@dvload/pptb   (after a build)

import { cp, mkdir, rm, access } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const pkgRoot = path.resolve(here, "..");
const dist = path.join(pkgRoot, "dist");
const out = path.join(pkgRoot, "publish");

try {
  await access(path.join(dist, "index.html"));
} catch {
  throw new Error(
    `No dist/index.html in ${path.relative(process.cwd(), dist)}.\n` +
      `  Run \`npm run build --workspace=@dvload/pptb\` first.`
  );
}

await rm(out, { recursive: true, force: true });
await mkdir(out, { recursive: true });
await cp(dist, path.join(out, "dist"), { recursive: true });
await cp(path.join(pkgRoot, "tool.package.json"), path.join(out, "package.json"));
await cp(path.join(pkgRoot, "README.md"), path.join(out, "README.md"));

console.log(`  packaged   ${path.relative(process.cwd(), out)}`);
console.log(`  publish with:  cd packages/pptb/publish && npm publish --access public`);
