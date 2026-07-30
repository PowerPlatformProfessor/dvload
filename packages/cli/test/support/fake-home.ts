/**
 * A disposable HOME directory.
 *
 * profiles.ts, secure-store.ts and telemetry.ts all compute their file paths
 * from `os.homedir()` at MODULE LOAD time:
 *
 *   const DIR = path.join(os.homedir(), ".dvload");
 *
 * That means setting $HOME inside a test is too late, and it also means these
 * modules would otherwise read and write the real `~/.dvload` of whoever runs
 * the suite — clobbering a developer's actual saved profiles and, worse, their
 * stored refresh tokens.
 *
 * So the mock has to be installed before the import, and the import has to be
 * dynamic. `withFakeHome()` packages that up.
 */

import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { vi } from "vitest";

export interface FakeHome {
  /** The temporary home directory. */
  home: string;
  /** `<home>/.dvload`, where every module under test writes. */
  dvloadDir: string;
  /** Remove the directory tree. */
  cleanup(): void;
}

/**
 * Create a temp home, point `node:os` at it, reset the module registry, and
 * hand back both the paths and a freshly imported copy of the module.
 *
 * Only `homedir` is replaced; `platform`, `tmpdir` and the rest keep their
 * real behaviour so platform-specific branches are still exercised for real.
 *
 * The module is supplied as a callback rather than a path string, because a
 * bare `import(someString)` here would resolve relative to THIS file rather
 * than to the caller's:
 *
 *   const { mod } = await withFakeHome(() => import("../src/profiles.js"));
 */
export async function withFakeHome<T>(importModule: () => Promise<T>): Promise<{ mod: T } & FakeHome> {
  const home = mkdtempSync(path.join(tmpdir(), "dvload-home-"));

  vi.doMock("node:os", async () => {
    const actual = await vi.importActual<typeof import("node:os")>("node:os");
    return { ...actual, default: { ...actual, homedir: () => home }, homedir: () => home };
  });

  // Force a fresh evaluation so the module-level path constants pick up the
  // mock rather than a value captured by an earlier test.
  vi.resetModules();
  const mod = await importModule();

  return {
    mod,
    home,
    dvloadDir: path.join(home, ".dvload"),
    cleanup: () => {
      vi.doUnmock("node:os");
      rmSync(home, { recursive: true, force: true });
    },
  };
}
