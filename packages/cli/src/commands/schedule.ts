// Registers a daily Windows Scheduled Task that:
//   1. Refreshes Power Query in the workbook (via Excel COM)
//   2. Runs `dvload run <mapping> -w <workbook>`
//
// We shell out to schtasks.exe (built in to Windows). Lives in the user's
// scope, no admin rights required.

import { spawn } from "node:child_process";
import path from "node:path";
import { readFile } from "node:fs/promises";
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

  // The /TR value is executed via `cmd /c`, so cmd metacharacters in a path
  // or task name would be interpreted as commands. Refuse rather than trying
  // to escape cmd's quoting rules.
  for (const [label, value] of [
    ["mapping path", mappingFile],
    ["workbook path", workbookFile],
    ["task name", taskName],
  ] as const) {
    if (/[&|<>^%"!\r\n]/.test(value)) {
      throw new Error(
        `The ${label} contains characters not allowed in a scheduled command ` +
          `(& | < > ^ % " !): ${value}\nRename or move it, or pass a different --name.`
      );
    }
  }

  // Strongly nudge toward app-only for scheduled runs. Delegated tokens
  // expire after 90 days of inactivity (or sooner under conditional access),
  // and a scheduled task can't pop a device-code prompt to recover.
  const mode = await detectAuthMode(mapping.environmentUrl);
  if (mode !== "appOnly") {
    console.log(
      kleur.yellow(
        "Warning: this environment is using " + mode + " auth.\n" +
          "Scheduled runs will eventually fail when the refresh token expires.\n" +
          "Recommended: run `dvload app-login --env " + mapping.environmentUrl +
          " --client-id <id> --tenant-id <id>` first."
      )
    );
  }

  // Locate the dvload CLI: prefer the global npm install on PATH;
  // fall back to "node <this script>" so the task works in dev too.
  const command = `dvload run "${mappingFile}" -w "${workbookFile}" --refresh`;

  const args = [
    "/Create",
    "/TN",
    taskName,
    "/TR",
    `cmd /c ${command}`,
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
  console.log(kleur.gray("  To remove: schtasks /Delete /TN \"" + taskName + "\" /F"));
  console.log(kleur.gray("  To run now: schtasks /Run  /TN \"" + taskName + "\""));
}
