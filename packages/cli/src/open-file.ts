// Hand a file to the OS default application.
//
// Used by `pqt-to-xlsx --open` so the generated workbook lands in Excel
// without a second command. Deliberately fire-and-forget: we detach the child
// and unref it, because Excel outlives the CLI process and we must not keep
// the event loop alive waiting for the user to close a spreadsheet.
//
// Windows note: the obvious `cmd /c start "" "<path>"` is a quoting minefield
// under Node's spawn (cmd.exe re-parses the command line, and Node's own
// argument quoting fights with `start`'s treatment of the first quoted token
// as a window title). `rundll32 url.dll,FileProtocolHandler` takes the path as
// a single ordinary argv entry with no shell involved, so paths containing
// spaces, `&`, or `^` work unmodified.

import { spawn } from "node:child_process";

/**
 * The launcher for the current platform, or null if we don't know one.
 *
 * Reads `process.platform` per call rather than destructuring it at import
 * time: a module-level `import { platform } from "node:process"` is captured
 * once at evaluation and can't be varied, which would make the per-platform
 * behaviour here untestable on a single machine.
 */
function launcher(filePath: string): { cmd: string; args: string[] } | null {
  switch (process.platform) {
    case "win32":
      return { cmd: "rundll32", args: ["url.dll,FileProtocolHandler", filePath] };
    case "darwin":
      return { cmd: "open", args: [filePath] };
    case "linux":
      return { cmd: "xdg-open", args: [filePath] };
    default:
      return null;
  }
}

/**
 * Open `filePath` in the OS default application.
 *
 * Resolves once the child has been handed off — not when the application
 * closes. Rejects only if the launcher itself could not be started (missing
 * `xdg-open` on a headless Linux box, unrecognized platform), which callers
 * should treat as "tell the user the path" rather than as a hard failure.
 */
export async function openInDefaultApp(filePath: string): Promise<void> {
  const spec = launcher(filePath);
  if (!spec) {
    throw new Error(`Don't know how to open files on platform "${process.platform}".`);
  }

  await new Promise<void>((resolve, reject) => {
    const child = spawn(spec.cmd, spec.args, {
      detached: true,
      stdio: "ignore",
      windowsHide: true,
    });
    // "error" fires asynchronously when the binary is missing or not
    // executable; anything after handoff is the OS's problem, not ours.
    child.once("error", (err: Error) => {
      reject(new Error(`Couldn't launch ${spec.cmd}: ${err.message}`));
    });
    child.once("spawn", () => {
      child.unref();
      resolve();
    });
  });
}
