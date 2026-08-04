// Sandbox-only config: the npm workspace junction node_modules/@dvload/core
// does not survive the mount, so alias it to the built output directly and
// skip the repo's global setup (which rebuilds the workspace).
import { defineConfig } from "vitest/config";
import path from "node:path";

const repo = "/sessions/hopeful-awesome-lamport/mnt/dvload";

export default defineConfig({
  resolve: {
    alias: {
      "@dvload/core": path.join(repo, "packages/core/dist/index.js"),
    },
  },
  test: {
    include: [
      path.join(repo, "packages/cli/src/auth.test.ts"),
      path.join(repo, "packages/cli/src/commands/login.test.ts"),
      path.join(repo, "packages/cli/src/commands/serve.test.ts"),
    ],
    environment: "node",
  },
});
