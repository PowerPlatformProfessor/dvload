#!/usr/bin/env node
// CLI entry. Subcommands live in ./commands/*.ts.

import { Command } from "commander";
import { runCommand } from "./commands/run.js";
import {
  loginCommand,
  logoutCommand,
  appLoginCommand,
  appLogoutCommand,
  whoamiCommand,
} from "./commands/login.js";
import { validateCommand } from "./commands/validate.js";
import { scheduleCommand } from "./commands/schedule.js";
import { extractPqtCommand, importPqtCommand } from "./commands/pqt.js";

const program = new Command();

program
  .name("dvload")
  .description("Map an Excel table to Dataverse and load via the OData Web API.")
  .version("0.1.0");

program
  .command("run")
  .description("Execute a saved .dvmap.json against an .xlsx workbook.")
  .argument("<mapping>", "Path to a .dvmap.json file")
  .requiredOption("-w, --workbook <path>", "Path to the .xlsx workbook")
  .option("--refresh", "Refresh Power Query in Excel before reading (Windows-only).")
  .option("--dry-run", "Coerce + plan the import but don't call Dataverse.")
  .option("--user", "Force delegated (interactive) auth even if app-only credentials are configured.")
  .option("--max-errors <n>", "Override mapping.maxErrors", parseIntStrict)
  .option("--no-color", "Disable color output.")
  .action(runCommand);

program
  .command("validate")
  .description("Validate a .dvmap.json against the schema and against Dataverse metadata.")
  .argument("<mapping>", "Path to a .dvmap.json file")
  .option("--no-remote", "Skip Dataverse-side metadata checks.")
  .option("--user", "Use delegated auth for the metadata probe.")
  .action(validateCommand);

// Delegated (user) auth -----------------------------------------------------
program
  .command("login")
  .description("Sign in interactively (delegated flow). Caches a refresh token.")
  .requiredOption("--env <url>", "Dataverse environment URL")
  .option("--tenant <id>", "Azure AD tenant id (default: 'organizations')")
  .option("--client-id <id>", "Override the public-client app id")
  .action(loginCommand);

program
  .command("logout")
  .description("Clear cached delegated tokens.")
  .option("--env <url>", "Only clear this environment (default: clear all).")
  .action(logoutCommand);

// App-only (client credentials) auth ----------------------------------------
program
  .command("app-login")
  .description("Configure app-only (client credentials) auth for unattended runs.")
  .requiredOption("--env <url>", "Dataverse environment URL")
  .requiredOption("--client-id <id>", "Azure AD app (confidential client) id")
  .requiredOption("--tenant-id <id>", "Azure AD tenant id (the Dataverse tenant)")
  .option("--secret-env <name>", "Read the secret from this env var instead of prompting")
  .action(appLoginCommand);

program
  .command("app-logout")
  .description("Clear stored app-only credentials.")
  .requiredOption("--env <url>", "Dataverse environment URL")
  .action(appLogoutCommand);

program
  .command("whoami")
  .description("Show which auth mode is configured for an environment and probe a token.")
  .requiredOption("--env <url>", "Dataverse environment URL")
  .action(whoamiCommand);

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
  .requiredOption("--env <url>", "Dataverse environment URL to write into the mapping")
  .option("--query <name>", "Which query in the .pqt to read (defaults to the LoadEnabled one)")
  .option("--entity <set>", "Override the target entity set name")
  .option("-o, --out <path>", "Output .dvmap.json path (default: <query>.dvmap.json next to the .pqt)")
  .option("--emit-m", "Also write the MashupDocument.pq alongside the mapping.")
  .action(importPqtCommand);

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
