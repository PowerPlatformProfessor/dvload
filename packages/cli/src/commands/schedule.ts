// Registers a daily Windows Scheduled Task that:
//   1. Refreshes Power Query in the workbook (via Excel COM)
//   2. Runs `dvload run <mapping> -w <workbook>`
//
// We shell out to schtasks.exe (built in to Windows). Lives in the user's
// scope, no admin rights required.
//
// Instead of putting the whole command line into /TR (which is executed
// via cmd.exe and subject to its quoting/metacharacter rules), we write a
// small .cmd wrapper into ~/.dvload/tasks/ and point /TR at that file.
// The wrapper pins the absolute node executable and CLI entry script, so
// the task keeps working even when `dvload` isn't on the task's PATH.

import { spawn } from "node:child_process";
import { access, mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import os from "node:os";
import { fileURLToPath } from "node:url";
import kleur from "kleur";
import { parseMapping } from "@dvload/core";
import { detectAuthMode } from "../auth.js";

interface ScheduleOpts {
  workbook: string;
  time: string;
  name?: string;
}

export async function scheduleCommand(mappingPath: string, opts: ScheduleOpts): Promise<void> {
  if (process.platform !== "win32") {
    throw new Error("`schedule` is Windows-only in v1. Use cron + the `run --refresh` command instead.");
  }
  if (!/^\d{2}:\d{2}$/.test(opts.time)) {
    throw new Error("--time must be HH:MM in 24h format, e.g. 03:30");
  }

  const mappingFile = path.resolve(mappingPath);
  const workbookFile = path.resolve(opts.workbook);
  const mapping = parseMapping(JSON.parse(await readFile(mappingFile, "utf8")));
  const taskName = opts.name ?? `dvload: ${mapping.name}`;

  // Quotes and newlines cannot be represented safely in a .cmd argument or
  // a schtasks task name. Everything else (spaces, &, %, …) is fine now
  // that the command lives in a wrapper script.
  for (const [label, value] of [
    ["mapping path", mappingFile],
    ["workbook path", workbookFile],
    ["task name", taskName],
  ] as const) {
    if (/["\r\n]/.test(value)) {
      throw new Error(
        `The ${label} contains a double quote or newline, which cannot be scheduled safely: ${value}\n` +
          `Rename or move it, or pass a different --name.`
      );
    }
  }

  // Strongly nudge toward app-only for scheduled runs. Delegated tokens
  // expire after 90 days of inactivity (or sooner under conditional access),
  // and a scheduled task can't pop a device-code prompt to recover.
  // cmd.exe expands %VAR% inside batch files; escape literal % as %%.
  const q = (s: string): string => `"${s.replace(/%/g, "%%")}"`;

  const mode = await detectAuthMode(mapping.environmentUrl);
  if (mode !== "appOnly") {
    console.log(
      kleur.yellow(
        "Warning: this environment is using " + mode + " auth.\n" +
          "Scheduled runs will eventually fail when the refresh token expires.\n" +
          "Recommended: run `dvload app-login --env " + mapping.environmentUrl +
          " --client-id <id> --tenant-id <id>` first (a certificate via --cert never expires by rotation policy)."
      )
    );
  }

  // Resolve the CLI entry point relative to this compiled file
  // (dist/commands/schedule.js → dist/index.js). Pinning node + script
  // avoids depending on PATH inside the Task Scheduler environment.
  // In a single-exe (SEA) build there is no script file on disk — the exe
  // itself IS the CLI, so invoke it directly.
  const cliEntry = fileURLToPath(new URL("../index.js", import.meta.url));
  let launcher: string;
  try {
    await access(cliEntry);
    launcher = `${q(process.execPath)} ${q(cliEntry)}`;
  } catch {
    launcher = q(process.execPath);
  }

  const wrapper =
    "@echo off\r\n" +
    `${launcher} run ${q(mappingFile)} -w ${q(workbookFile)} --refresh --non-interactive\r\n`;

  const taskDir = path.join(os.homedir(), ".dvload", "tasks");
  await mkdir(taskDir, { recursive: true });
  const scriptFile = path.join(
    taskDir,
    `${taskName.replace(/[^\w.-]+/g, "_")}.cmd`
  );
  await writeFile(scriptFile, wrapper, "utf8");

  const args = [
    "/Create",
    "/TN",
    taskName,
    "/TR",
    `"${scriptFile}"`,
    "/SC",
    "DAILY",
    "/ST",
    opts.time,
    "/F", // overwrite existing
    "/RL",
    "LIMITED",
  ];

  await new Promise<void>((resolve, reject) => {
    const child = spawn("schtasks.exe", args, { windowsHide: true });
    let stderr = "";
    child.stderr.on("data", (d: Buffer) => (stderr += d.toString()));
    child.on("error", reject);
    child.on("exit", (code) => {
      if (code === 0) resolve();
      else reject(new Error(`schtasks failed (exit ${code}): ${stderr}`));
    });
  });

  console.log(kleur.green(`Registered Scheduled Task "${taskName}" for ${opts.time} daily.`));
  console.log(kleur.gray(`  Wrapper script: ${scriptFile}`));
  console.log(kleur.gray("  To remove: schtasks /Delete /TN \"" + taskName + "\" /F"));
  console.log(kleur.gray("  To run now: schtasks /Run  /TN \"" + taskName + "\""));
}
