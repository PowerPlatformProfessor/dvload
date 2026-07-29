// ESLint flat config.
//
// The repo previously had a `lint` script but no config file at all, so
// `npm run lint` failed in CI. This restores it, and adds the rules that
// matter for a tool that writes to production data with contributions from
// people who don't know the codebase.
//
// Philosophy: lint only for things a reviewer shouldn't have to notice.
// Formatting is Prettier's job, so no stylistic rules are enabled here.
//
// Rules that would require touching a lot of existing code are set to "warn"
// rather than "error" — see the note above the type-safety block. A lint
// config that fails 400 times on day one gets bypassed, and a bypassed gate
// protects nothing.

import js from "@eslint/js";
import tseslint from "typescript-eslint";
import vitest from "eslint-plugin-vitest";

/** Everything that is real, checked-in source. */
const SOURCE = ["packages/*/src/**/*.ts", "packages/*/test/**/*.ts"];

export default tseslint.config(
  {
    ignores: [
      "**/dist/**",
      "**/build/**",
      "**/node_modules/**",
      "**/coverage/**",
      "**/reports/**",
      ".stryker-tmp/**",
      "packages/addin/webpack.config.js",
      "packages/addin/dist/**",
      "tests/dummy-data/**",
      "**/*.d.ts",
    ],
  },

  js.configs.recommended,

  /* --- type-aware linting, scoped to files that are in a tsconfig --------- */
  {
    files: SOURCE,
    extends: [...tseslint.configs.recommendedTypeChecked],
    languageOptions: {
      parserOptions: {
        // The *.test.json projects include both src and test, so type-aware
        // rules apply to test files too. A type error in a test is still a bug.
        project: [
          "packages/core/tsconfig.test.json",
          "packages/cli/tsconfig.test.json",
          "packages/addin/tsconfig.test.json",
        ],
        tsconfigRootDir: import.meta.dirname,
      },
    },
    rules: {
      /* --- correctness: these block ------------------------------------- */

      // A floating promise in the loader means a batch nobody waits for: the
      // run reports success while writes are still in flight.
      "@typescript-eslint/no-floating-promises": "error",
      "@typescript-eslint/no-misused-promises": "error",
      "@typescript-eslint/await-thenable": "error",

      // `catch {}` around a Dataverse call silently turns a failed write into
      // a reported success. Force an explicit decision.
      "no-empty": ["error", { allowEmptyCatch: false }],

      eqeqeq: ["error", "always", { null: "ignore" }],
      "prefer-const": "error",
      "no-var": "error",
      "no-console": "off", // this is a CLI

      "@typescript-eslint/no-unused-vars": [
        "error",
        { argsIgnorePattern: "^_", varsIgnorePattern: "^_", caughtErrors: "none" },
      ],

      /* --- type safety: advisory for now -------------------------------- */
      // Data from Excel and from Dataverse is genuinely `unknown`, and the
      // existing code casts through it in a lot of places. These are the
      // rules to tighten to "error" once the current warnings are worked
      // through — file by file, not in one sweep.
      "@typescript-eslint/no-explicit-any": "warn",
      "@typescript-eslint/no-unsafe-assignment": "warn",
      "@typescript-eslint/no-unsafe-member-access": "warn",
      "@typescript-eslint/no-unsafe-call": "warn",
      "@typescript-eslint/no-unsafe-return": "warn",
      "@typescript-eslint/no-unsafe-argument": "warn",
      "@typescript-eslint/require-await": "warn",
      "@typescript-eslint/restrict-template-expressions": "warn",
      "@typescript-eslint/no-base-to-string": "warn",
      "@typescript-eslint/no-redundant-type-constituents": "warn",
      "@typescript-eslint/no-unnecessary-type-assertion": "warn",
    },
  },

  /* --- tests -------------------------------------------------------------- */
  {
    files: ["**/*.test.ts", "packages/*/test/**/*.ts"],
    plugins: { vitest },
    rules: {
      // A test that asserts nothing passes forever. This is the single most
      // valuable lint rule in a repo that gates merges on coverage.
      "vitest/expect-expect": [
        "error",
        // node:assert is still used throughout the migrated suites.
        { assertFunctionNames: ["expect", "assert", "assert.*"] },
      ],
      "vitest/no-focused-tests": "error", // a stray .only silently skips the rest
      "vitest/no-identical-title": "error",
      "vitest/valid-expect": "error",
      "vitest/no-disabled-tests": "warn",

      // Fixtures legitimately fake partial shapes.
      "@typescript-eslint/no-unsafe-assignment": "off",
      "@typescript-eslint/no-unsafe-member-access": "off",
      "@typescript-eslint/no-unsafe-call": "off",
      "@typescript-eslint/no-unsafe-argument": "off",
      "@typescript-eslint/no-unsafe-return": "off",
      "@typescript-eslint/no-explicit-any": "off",
    },
  },

  /* --- add-in task pane: known debt, deliberately not fixed here ---------- */
  {
    files: ["packages/addin/src/taskpane/**/*.ts", "packages/addin/src/commands/**/*.ts"],
    rules: {
      // ~10 occurrences of `element.addEventListener("click", asyncHandler)`.
      // The rule is right — if such a handler rejects, the user sees nothing
      // and the failure is an unhandled rejection in the Office host. But
      // fixing it means editing 1,700 lines of untested UI, which does not
      // belong in a change to the testing infrastructure.
      //
      // TODO: wrap these handlers so rejections surface in the task pane,
      // then delete this override. Tracked as debt in TESTING.md.
      "@typescript-eslint/no-misused-promises": "warn",
      "@typescript-eslint/no-floating-promises": "warn",
    },
  },

  /* --- tooling outside any tsconfig --------------------------------------- */
  // Root config files and plain-JS scripts. Type-aware rules need a program
  // to consult; these files aren't in one, so they get syntactic rules only
  // rather than being excluded entirely.
  {
    files: ["*.config.ts", "*.config.mjs", "**/*.mjs", "**/*.js", "scripts/**", "tests/*.mjs"],
    extends: [tseslint.configs.disableTypeChecked],
    languageOptions: {
      globals: {
        process: "readonly",
        console: "readonly",
        URL: "readonly",
        URLSearchParams: "readonly",
        fetch: "readonly",
        Buffer: "readonly",
        structuredClone: "readonly",
        __dirname: "readonly",
      },
    },
    rules: {
      // The base rule, not the TS one: this block has no TypeScript program
      // and therefore no @typescript-eslint plugin registered.
      "no-unused-vars": ["error", { argsIgnorePattern: "^_", varsIgnorePattern: "^_", caughtErrors: "none" }],
    },
  }
);
