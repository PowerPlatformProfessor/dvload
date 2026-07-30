/**
 * Excel COM refresh (Windows-only, shells out to PowerShell).
 *
 * Actually refreshing a workbook needs Excel installed, so that stays in the
 * manual protocol. What CAN be tested — and matters more — is how the child
 * process is invoked:
 *
 *   - The workbook path travels via an ENVIRONMENT VARIABLE, never
 *     interpolated into the script. A path is user input; a workbook called
 *     `x'; Remove-Item C:\ -Recurse; '.xlsx` must not become PowerShell code.
 *   - The timeout lives inside the script, so the `finally` block always runs
 *     and Excel always quits. Killing PowerShell from Node skips it and
 *     orphans a headless EXCEL.EXE — fatal on a box running nightly imports.
 *   - Failures surface with stderr attached, not as a bare exit code.
 *
 * `node:child_process` is mocked so these run on any platform.
 */

import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";
import { EventEmitter } from "node:events";
import path from "node:path";

/** A spawn() stand-in whose exit code and stderr the test controls. */
class FakeChild extends EventEmitter {
  stderr = new EventEmitter();
  stdout = new EventEmitter();
  kill = vi.fn();
}

const spawn = vi.hoisted(() => vi.fn());

vi.mock("node:child_process", async (importOriginal) => {
  const actual = await importOriginal<typeof import("node:child_process")>();
  return { ...actual, default: { ...actual, spawn }, spawn };
});

const { refreshWorkbook } = await import("../src/refresh.js");

/** Arrange spawn() to return a child that exits with `code` after a tick. */
function childExiting(code: number, stderr = ""): FakeChild {
  const child = new FakeChild();
  spawn.mockReturnValueOnce(child);
  queueMicrotask(() => {
    if (stderr) child.stderr.emit("data", Buffer.from(stderr));
    child.emit("exit", code);
  });
  return child;
}

/** The arguments spawn() was called with, destructured for readability. */
function lastSpawn(): { cmd: string; args: string[]; opts: { env: Record<string, string> } } {
  const [cmd, args, opts] = spawn.mock.calls[spawn.mock.calls.length - 1] as [
    string,
    string[],
    { env: Record<string, string> },
  ];
  return { cmd, args, opts };
}

beforeEach(() => {
  spawn.mockReset();
  vi.useFakeTimers();
});

afterEach(() => {
  vi.useRealTimers();
});

describe("invocation", () => {
  it("resolves once PowerShell exits cleanly", async () => {
    childExiting(0);
    await expect(refreshWorkbook({ workbook: "book.xlsx" })).resolves.toBeUndefined();
  });

  it("runs PowerShell non-interactively with no profile and a hidden window", async () => {
    // A scheduled task runs with no desktop session. A profile script or an
    // interactive prompt here hangs the nightly import with nothing on screen.
    childExiting(0);
    await refreshWorkbook({ workbook: "book.xlsx" });

    const { cmd, args, opts } = lastSpawn();
    expect(cmd).toBe("powershell.exe");
    expect(args).toContain("-NoProfile");
    expect(args).toContain("-NonInteractive");
    expect(opts).toMatchObject({ windowsHide: true });
  });

  it("passes the workbook path through the environment, never inside the script", async () => {
    // The injection guard. If the path were interpolated, this filename
    // would close the string and run a second statement.
    const hostile = `evil'; Remove-Item C:\\ -Recurse -Force; '.xlsx`;
    childExiting(0);

    await refreshWorkbook({ workbook: hostile });

    const { args, opts } = lastSpawn();
    const script = args[args.length - 1];

    expect(opts.env.DVLOAD_WB).toBe(path.resolve(hostile));
    expect(script).not.toContain("Remove-Item");
    expect(script).toContain("$env:DVLOAD_WB");
  });

  it("resolves the workbook to an absolute path", async () => {
    // Excel's COM Open() resolves relative paths against its own working
    // directory, not the CLI's — which under Task Scheduler is C:\Windows\System32.
    childExiting(0);
    await refreshWorkbook({ workbook: "data/book.xlsx" });

    expect(lastSpawn().opts.env.DVLOAD_WB).toBe(path.resolve("data/book.xlsx"));
    expect(path.isAbsolute(lastSpawn().opts.env.DVLOAD_WB)).toBe(true);
  });

  it("keeps Excel hidden by default and shows it only on request", async () => {
    childExiting(0);
    await refreshWorkbook({ workbook: "book.xlsx" });
    expect(lastSpawn().args[lastSpawn().args.length - 1]).toContain("$excel.Visible = $false");

    childExiting(0);
    await refreshWorkbook({ workbook: "book.xlsx", visible: true });
    expect(lastSpawn().args[lastSpawn().args.length - 1]).toContain("$excel.Visible = $true");
  });
});

describe("timeout", () => {
  it("gives the script its own deadline so Excel is always quit", async () => {
    // The script must own the timeout. Node killing PowerShell skips the
    // `finally` that calls $excel.Quit(), leaving an orphaned EXCEL.EXE.
    childExiting(0);
    await refreshWorkbook({ workbook: "book.xlsx", timeoutMs: 90_000 });

    const { args, opts } = lastSpawn();
    const script = args[args.length - 1];

    expect(opts.env.DVLOAD_TIMEOUT_MS).toBe("90000");
    expect(script).toContain("$env:DVLOAD_TIMEOUT_MS");
    expect(script).toContain("finally");
    expect(script).toContain("$excel.Quit()");
  });

  it("defaults to five minutes", async () => {
    childExiting(0);
    await refreshWorkbook({ workbook: "book.xlsx" });
    expect(lastSpawn().opts.env.DVLOAD_TIMEOUT_MS).toBe(String(5 * 60 * 1000));
  });

  it("kills PowerShell only 30s past the script's own deadline, and says so", async () => {
    const child = new FakeChild();
    spawn.mockReturnValueOnce(child); // never exits

    // .catch() up front attaches a handler immediately, so advancing the
    // timers below can't trip an unhandled-rejection warning.
    const settled = refreshWorkbook({ workbook: "book.xlsx", timeoutMs: 1000 }).catch((e: Error) => e);

    // Not yet: the script is still inside its own budget.
    await vi.advanceTimersByTimeAsync(1000 + 29_000);
    expect(child.kill).not.toHaveBeenCalled();

    await vi.advanceTimersByTimeAsync(2000);
    expect(child.kill).toHaveBeenCalledWith("SIGKILL");

    const err = await settled;
    expect(err).toBeInstanceOf(Error);
    expect((err as Error).message).toMatch(/killed PowerShell/);
  });

  it("warns about a possible orphaned Excel when it has to kill", async () => {
    // The user needs to know to check Task Manager; a silent kill leaves a
    // process holding the workbook open and the next run fails confusingly.
    const child = new FakeChild();
    spawn.mockReturnValueOnce(child);

    const settled = refreshWorkbook({ workbook: "book.xlsx", timeoutMs: 1000 }).catch((e: Error) => e);
    await vi.advanceTimersByTimeAsync(31_001);

    const err = await settled;
    expect(err).toBeInstanceOf(Error);
    expect((err as Error).message).toMatch(/EXCEL\.EXE/);
  });

  it("clears the kill timer on success so the process can exit", async () => {
    // A dangling timer keeps the Node event loop alive and the CLI never
    // returns to the shell — very visible under Task Scheduler.
    childExiting(0);
    await refreshWorkbook({ workbook: "book.xlsx" });

    expect(vi.getTimerCount()).toBe(0);
  });
});

describe("failure reporting", () => {
  it("includes stderr in the error, not just the exit code", async () => {
    childExiting(1, "Cannot open workbook: file is locked by another user");

    await expect(refreshWorkbook({ workbook: "book.xlsx" })).rejects.toThrow(/locked by another user/);
  });

  it("reports the exit code too", async () => {
    childExiting(2, "boom");
    await expect(refreshWorkbook({ workbook: "book.xlsx" })).rejects.toThrow(/exit 2/);
  });

  it("propagates a spawn failure, e.g. PowerShell missing", async () => {
    // What a non-Windows machine, or a locked-down box without PowerShell,
    // actually produces.
    const child = new FakeChild();
    spawn.mockReturnValueOnce(child);
    queueMicrotask(() =>
      child.emit("error", Object.assign(new Error("spawn powershell.exe ENOENT"), { code: "ENOENT" }))
    );

    await expect(refreshWorkbook({ workbook: "book.xlsx" })).rejects.toThrow(/ENOENT/);
    expect(vi.getTimerCount()).toBe(0);
  });
});
