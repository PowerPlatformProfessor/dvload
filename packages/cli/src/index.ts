#!/usr/bin/env node
// CLI entry. Subcommands live in ./commands/*.ts.

import { Command } from "commander";
import { runCommand } from "./commands/run.js";
import { runAllCommand } from "./commands/run-all.js";
import {
  loginCommand,
  logoutCommand,
  appLoginCommand,
  appLogoutCommand,
  whoamiCommand,
} from "./commands/login.js";
import { validateCommand } from "./commands/validate.js";
import { scheduleCommand } from "./commands/schedule.js";
import { extractPqtCommand, importPqtCommand, pqtToXlsxCommand } from "./commands/pqt.js";
import { profileAddCommand, profileRemoveCommand, profileListCommand } from "./commands/profile.js";
import { addinCommand } from "./commands/addin.js";
import { telemetryCommand } from "./telemetry.js";

const program = new Command();

program
  .name("dvload")
  .description("Map an Excel table to Dataverse and load via the OData Web API.")
  .version("0.1.0");

program
  .command("run")
  .description("Execute a saved .dvmap.json against an .xlsx (or .csv/.tsv) source file.")
  .argument("<mapping>", "Path to a .dvmap.json file")
  .requiredOption("-w, --workbook <path>", "Path to the .xlsx workbook (or .csv/.tsv file)")
  .option("--refresh", "Refresh Power Query in Excel before reading (Windows-only).")
  .option("--dry-run", "Coerce + plan the import but don't call Dataverse.")
  .option("--user", "Force delegated (interactive) auth even if app-only credentials are configured.")
  .option("--max-errors <n>", "Override mapping.maxErrors", parseIntStrict)
  .option("--concurrency <n>", "Parallel $batch requests (1-8); overrides mapping.concurrency", parseIntStrict)
  .option(
    "--max-attempts <n>",
    "Attempts per request, including the first (default 5). Raise it to ride out longer network outages.",
    parseIntStrict
  )
  .option("--resume", "Resume an interrupted run from its checkpoint (same workbook only).")
  .option("--notify-url <url>", "POST a {text} summary to this webhook after the run.")
  .option("--no-failed-rows", "Don't write the failed-rows .xlsx re-run file.")
  .option("--no-color", "Disable color output.")
  .option("--non-interactive", "Fail fast instead of prompting for device-code login (default when piped).")
  .option("--json", "Print the run result as JSON on stdout; suppress progress output.")
  .action(runCommand);

program
  .command("run-all")
  .description("Run several mappings from a run plan (stages/dependencies; legacy runs[] manifests still supported).")
  .argument("<manifest>", "Path to a .dvplan.json file (or legacy manifest .json with runs[])")
  .option("--dry-run", "Coerce + plan all imports but don't call Dataverse.")
  .option("--user", "Force delegated auth for all runs.")
  .option("--notify-url <url>", "POST a {text} summary per run.")
  .option(
    "--max-attempts <n>",
    "Attempts per request, including the first (default 5). Applied to every step.",
    parseIntStrict
  )
  .option("--no-failed-rows", "Don't write failed-rows .xlsx files.")
  .action(runAllCommand);

program
  .command("validate")
  .description("Validate a .dvmap.json against the schema and against Dataverse metadata.")
  .argument("<mapping>", "Path to a .dvmap.json file")
  .option("--no-remote", "Skip Dataverse-side metadata checks.")
  .option("--user", "Use delegated auth for the metadata probe.")
  .action(validateCommand);

program
  .command("login")
  .description("Sign in interactively (delegated flow). Caches a refresh token.")
  .option("--env <url>", "Dataverse environment URL")
  .option("-p, --profile <name>", "Use a saved environment profile instead of --env")
  .option("--tenant <id>", "Azure AD tenant id (default: 'organizations')")
  .option("--client-id <id>", "Override the public-client app id")
  .option("--interactive", "Sign in via the system browser (default when one is available)")
  .option(
    "--device-code",
    "Sign in by entering a code on another device. Needed on headless machines; " +
      "often blocked by Conditional Access."
  )
  .action(loginCommand);

program
  .command("logout")
  .description("Clear cached delegated tokens.")
  .option("--env <url>", "Only clear this environment (default: clear all).")
  .option("-p, --profile <name>", "Use a saved environment profile instead of --env")
  .action(logoutCommand);

// App-only (client credentials) auth ----------------------------------------
program
  .command("app-login")
  .description("Configure app-only (client credentials) auth for unattended runs.")
  .option("--env <url>", "Dataverse environment URL")
  .option("-p, --profile <name>", "Use a saved environment profile instead of --env")
  .requiredOption("--client-id <id>", "Azure AD app (confidential client) id")
  .requiredOption("--tenant-id <id>", "Azure AD tenant id (the Dataverse tenant)")
  .option("--secret-env <name>", "Read the secret from this env var instead of prompting")
  .option("--cert <pem-path>", "Use a certificate (PEM with cert + private key) instead of a secret")
  .action(appLoginCommand);

program
  .command("app-logout")
  .description("Clear stored app-only credentials.")
  .option("--env <url>", "Dataverse environment URL")
  .option("-p, --profile <name>", "Use a saved environment profile instead of --env")
  .action(appLogoutCommand);

program
  .command("whoami")
  .description("Show which auth mode is configured for an environment and probe a token.")
  .option("--env <url>", "Dataverse environment URL")
  .option("-p, --profile <name>", "Use a saved environment profile instead of --env")
  .action(whoamiCommand);

// Environment profiles -------------------------------------------------------
const profile = program
  .command("profile")
  .description("Manage named environment profiles (~/.dvload/profiles.json).");

profile
  .command("add <name> <url>")
  .description("Save an environment URL under a short name.")
  .action(profileAddCommand);

profile
  .command("remove <name>")
  .description("Delete a saved profile.")
  .action(profileRemoveCommand);

profile
  .command("list")
  .description("List all saved profiles.")
  .action(profileListCommand);

// Power Query Template (.pqt) ----------------------------------------------
program
  .command("extract-pqt")
  .description("Extract Power Query M from an Excel workbook into a .pqt file (Dataverse Dataflows format).")
  .argument("<workbook>", "Path to the .xlsx workbook")
  .option("-o, --out <path>", "Output .pqt path (default: <workbook>.pqt next to the source)")
  .option("--name <name>", "Metadata.json Name field")
  .option("--description <text>", "Metadata.json Description field")
  .option("--version <ver>", "Metadata.json Version field (default 1.0.0.0)")
  .option("--locale <code>", "Document locale (default en-US)")
  .option("--mapping <path>", "Inject FieldsMetadata from a .dvmap.json so the .pqt imports with mappings ready.")
  .action(extractPqtCommand);

program
  .command("import-pqt")
  .description("Synthesize a .dvmap.json from a .pqt's FieldsMetadata. Migration path from Dataverse Dataflows.")
  .argument("<pqt>", "Path to a .pqt file")
  .option("--env <url>", "Dataverse environment URL to write into the mapping")
  .option("-p, --profile <name>", "Use a saved environment profile instead of --env")
  .option("--query <name>", "Which query in the .pqt to read (defaults to the LoadEnabled one)")
  .option("--entity <set>", "Override the target entity set name")
  .option("-o, --out <path>", "Output .dvmap.json path (default: <query>.dvmap.json next to the .pqt)")
  .option("--emit-m", "Also write the MashupDocument.pq alongside the mapping.")
  .option("--all-queries", "Emit one .dvmap.json per query (-o becomes the output directory).")
  .action(importPqtCommand);

program
  .command("pqt-to-xlsx")
  .description("EXPERIMENTAL: build an .xlsx with the .pqt's queries embedded in Power Query.")
  .argument("<pqt>", "Path to a .pqt file (e.g. exported from Dataverse Dataflows)")
  .option("-o, --out <path>", "Output .xlsx path (default: <pqt-name>.xlsx next to the .pqt)")
  .action(pqtToXlsxCommand);

program
  .command("telemetry")
  .description("Show or change anonymous usage telemetry (see TELEMETRY.md).")
  .argument("[action]", '"on", "off", or omit for status')
  .action(telemetryCommand);

program
  .command("addin")
  .description("Launch or stop the Excel add-in workflow from the CLI (local/dev installs).")
  .argument("<action>", "start | stop | dev")
  .option(
    "--with-dev-server",
    "With action=start, run the webpack dev server (equivalent to npm run dev:addin)."
  )
  .action(addinCommand);

// Scheduling ---------------------------------------------------------------
program
  .command("schedule")
  .description("Register a Windows Task Scheduler job that refreshes Excel and runs an import.")
  .argument("<mapping>", "Path to a .dvmap.json file")
  .requiredOption("-w, --workbook <path>", "Path to the .xlsx workbook")
  .requiredOption("--time <HH:MM>", "Daily run time, 24h format")
  .option("--name <name>", "Scheduled task name (default: derived from mapping name).")
  .action(scheduleCommand);

program.parseAsync().catch((err: unknown) => {
  console.error(err instanceof Error ? err.message : String(err));
  process.exit(1);
});

function parseIntStrict(v: string): number {
  const n = Number(v);
  if (!Number.isInteger(n) || n < 0) throw new Error(`Expected non-negative integer, got "${v}"`);
  return n;
}
