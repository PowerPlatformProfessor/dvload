/**
 * `openInDefaultApp` hands a generated workbook to the OS. The contract that
 * matters is not "Excel appeared" — it's that we:
 *
 *   - pass the path as a single argv entry, never through a shell, so spaces
 *     and `&` in a filename can't be reinterpreted as command syntax;
 *   - detach and unref, because Excel outlives the CLI and a referenced child
 *     would hold the event loop open until the user closed the spreadsheet;
 *   - reject rather than hang when the launcher binary is missing.
 *
 * `node:child_process` is mocked so these run on any platform.
 */

import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";
import { EventEmitter } from "node:events";

class FakeChild extends EventEmitter {
  unref = vi.fn();
}

const spawn = vi.hoisted(() => vi.fn());

vi.mock("node:child_process", async (importOriginal) => {
  const actual = await importOriginal<typeof import("node:child_process")>();
  return { ...actual, default: { ...actual, spawn }, spawn };
});

/** Swap out process.platform, which the module reads at call time. */
function setPlatform(value: NodeJS.Platform): void {
  vi.spyOn(process, "platform", "get").mockReturnValue(value);
}

/** A child that reports a successful handoff on the next tick. */
function childSpawning(): FakeChild {
  const child = new FakeChild();
  spawn.mockReturnValueOnce(child);
  queueMicrotask(() => child.emit("spawn"));
  return child;
}

function lastSpawn(): { cmd: string; args: string[]; opts: Record<string, unknown> } {
  const [cmd, args, opts] = spawn.mock.calls[spawn.mock.calls.length - 1] as [
    string,
    string[],
    Record<string, unknown>,
  ];
  return { cmd, args, opts };
}

beforeEach(() => {
  spawn.mockReset();
});

afterEach(() => {
  vi.restoreAllMocks();
});

describe("openInDefaultApp", () => {
  it("uses rundll32 on Windows so the path needs no shell quoting", async () => {
    setPlatform("win32");
    const { openInDefaultApp } = await import("../src/open-file.js");
    childSpawning();

    // A path with a space and an ampersand: fatal through `cmd /c start`.
    await openInDefaultApp(String.raw`C:\Users\a b\R & D\out.xlsx`);

    const { cmd, args, opts } = lastSpawn();
    expect(cmd).toBe("rundll32");
    expect(args).toEqual(["url.dll,FileProtocolHandler", String.raw`C:\Users\a b\R & D\out.xlsx`]);
    // No shell: the path stays one argv entry and is never re-parsed.
    expect(opts.shell).toBeUndefined();
    expect(opts.windowsHide).toBe(true);
  });

  it("uses open on macOS and xdg-open on Linux", async () => {
    const { openInDefaultApp } = await import("../src/open-file.js");

    setPlatform("darwin");
    childSpawning();
    await openInDefaultApp("/tmp/out.xlsx");
    expect(lastSpawn()).toMatchObject({ cmd: "open", args: ["/tmp/out.xlsx"] });

    setPlatform("linux");
    childSpawning();
    await openInDefaultApp("/tmp/out.xlsx");
    expect(lastSpawn()).toMatchObject({ cmd: "xdg-open", args: ["/tmp/out.xlsx"] });
  });

  it("detaches and unrefs so the CLI can exit while Excel stays open", async () => {
    setPlatform("linux");
    const { openInDefaultApp } = await import("../src/open-file.js");
    const child = childSpawning();

    await openInDefaultApp("/tmp/out.xlsx");

    expect(lastSpawn().opts).toMatchObject({ detached: true, stdio: "ignore" });
    expect(child.unref).toHaveBeenCalledOnce();
  });

  it("rejects when the launcher binary is missing rather than hanging", async () => {
    setPlatform("linux");
    const { openInDefaultApp } = await import("../src/open-file.js");
    const child = new FakeChild();
    spawn.mockReturnValueOnce(child);
    queueMicrotask(() => child.emit("error", new Error("spawn xdg-open ENOENT")));

    await expect(openInDefaultApp("/tmp/out.xlsx")).rejects.toThrow(/Couldn't launch xdg-open.*ENOENT/);
  });

  it("refuses platforms it has no launcher for, without spawning anything", async () => {
    setPlatform("aix");
    const { openInDefaultApp } = await import("../src/open-file.js");

    await expect(openInDefaultApp("/tmp/out.xlsx")).rejects.toThrow(/Don't know how to open/);
    expect(spawn).not.toHaveBeenCalled();
  });
});
