// `dvload mcp` — a Model Context Protocol server on stdio, so AI agents
// (Claude, GitHub Copilot, Cursor, ...) can drive dvload through typed tools
// instead of parsing --help.
//
// Three rules shape everything here:
//
//   - stdout is the protocol channel. Nothing in this process may print to
//     it, so the tools call the same functions the commands do, in their
//     silent form, and mcpCommand() points console.log at stderr as a net
//     for anything that slips through.
//   - Nobody is at a prompt. Auth is always silentOnly: an expired or missing
//     session comes back as a tool error telling the agent to have the user
//     run `dvload login`, never as a browser window or a device code.
//   - Writes are opt-in twice. `run_mapping` defaults to a dry run, and
//     `dvload mcp --read-only` refuses real runs outright.
//
// Local only by design: tools take file paths on this machine and use the
// credentials in ~/.dvload. See docs/AGENTS.md.

import path from "node:path";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { z } from "zod";
import {
  DataverseClient,
  listDataflows,
  listTablesFromFile,
  readTableFromFile,
} from "@dvload/core";
import {
  detectAuthMode,
  getSignedInAccount,
  getTokenProvider,
  loadAppOnlyCredentials,
} from "../auth.js";
import { readProfiles, resolveEnv } from "../profiles.js";
import { TOOL_VERSION } from "../telemetry.js";
import { executeRun } from "./run.js";
import { checkMapping } from "./validate.js";

export interface McpOpts {
  /** Refuse run_mapping with dryRun=false. */
  readOnly?: boolean;
}

/** Row errors returned inline; the rest are in the run log and failed-rows file. */
const MAX_ERRORS_RETURNED = 20;
const MAX_SAMPLE_ROWS = 50;

const INSTRUCTIONS =
  "dvload loads Excel/CSV tables into Microsoft Dataverse, driven by a .dvmap.json mapping file. " +
  "Typical flow: inspect_source to see the columns, validate_mapping, run_mapping as a dry run, " +
  "then run_mapping with dryRun=false only once the user has agreed to write. " +
  "File paths are on the machine running this server; pass absolute paths. " +
  "Sign-in is never started from here: if a tool reports that login is required, ask the user " +
  "to run `dvload login --env <url>` in a terminal and retry.";

const envShape = {
  env: z.string().optional().describe("Dataverse environment URL, e.g. https://contoso.crm.dynamics.com"),
  profile: z.string().optional().describe("A saved profile name (see list_profiles), instead of env"),
};

export function buildMcpServer(opts: McpOpts = {}): McpServer {
  const server = new McpServer({ name: "dvload", version: TOOL_VERSION }, { instructions: INSTRUCTIONS });

  server.registerTool(
    "inspect_source",
    {
      title: "Inspect a source file",
      description:
        "List the tables in an .xlsx workbook (name, sheet, columns, row count), or the columns of a " +
        ".csv/.tsv file. Pass `table` (or any .csv/.tsv) to also get the first rows as a sample. " +
        "Use this to find the sourceTable and source column names for a .dvmap.json.",
      inputSchema: {
        file: z.string().describe("Absolute path to an .xlsx, .csv or .tsv file"),
        table: z.string().optional().describe("Excel table (or sheet) name to sample rows from"),
        sampleRows: z.number().int().min(0).max(MAX_SAMPLE_ROWS).optional().describe("Rows to sample (default 5)"),
      },
      annotations: { readOnlyHint: true, openWorldHint: false },
    },
    tool(async ({ file, table, sampleRows }) => {
      const source = path.resolve(file);
      const delimited = /\.(csv|tsv)$/i.test(source);
      const tables = delimited ? null : await listTablesFromFile(source);
      if (!delimited && !table) return { file: source, tables };

      // A `sheet` entry has no table name to read back by — only a sheet.
      const info = tables?.find((t) => t.name === table);
      const { headers, rows } = await readTableFromFile(
        source,
        info?.kind === "sheet" ? { sheetName: info.sheetName } : { tableName: table }
      );
      return {
        file: source,
        ...(tables ? { tables } : {}),
        columns: headers,
        rowCount: rows.length,
        sample: rows.slice(0, sampleRows ?? 5),
      };
    })
  );

  server.registerTool(
    "validate_mapping",
    {
      title: "Validate a mapping",
      description:
        "Check a .dvmap.json against the mapping schema and, unless remote=false, against the target " +
        "table's Dataverse metadata (needs a signed-in session). Returns schemaErrors, warnings and any " +
        "target attributes that don't exist. Writes nothing.",
      inputSchema: {
        mapping: z.string().describe("Absolute path to a .dvmap.json file"),
        remote: z.boolean().optional().describe("Also check target attributes in Dataverse (default true)"),
        user: z.boolean().optional().describe("Use the signed-in user even if app-only credentials exist"),
      },
      annotations: { readOnlyHint: true },
    },
    tool(async ({ mapping, remote, user }) => {
      const check = await checkMapping(mapping, { remote: remote ?? true, user, silentOnly: true });
      const ok = check.schemaErrors.length === 0 && (check.remote?.missingAttributes.length ?? 0) === 0;
      return { ok, ...check };
    })
  );

  server.registerTool(
    "run_mapping",
    {
      title: "Run a mapping",
      description:
        "Load the rows of a source file into Dataverse as described by a .dvmap.json. " +
        "DRY RUN BY DEFAULT: rows are read, coerced and planned but nothing is sent. " +
        "Pass dryRun=false to actually create/update (and, in sync mode, deactivate or delete) records — " +
        "do that only after a clean dry run and the user's go-ahead. " +
        "Returns counts (created/updated/skipped/failed), the first row errors, and the path of the " +
        "failed-rows file when rows failed." +
        (opts.readOnly ? " This server is read-only: dryRun=false is refused." : ""),
      inputSchema: {
        mapping: z.string().describe("Absolute path to a .dvmap.json file"),
        workbook: z.string().describe("Absolute path to the .xlsx, .csv or .tsv source file"),
        dryRun: z.boolean().optional().describe("Plan without calling Dataverse (default true)"),
        refresh: z.boolean().optional().describe("Refresh Power Query in Excel first (Windows, needs Excel)"),
        resume: z.boolean().optional().describe("Resume an interrupted run from its checkpoint"),
        maxErrors: z.number().int().min(0).optional().describe("Override mapping.maxErrors"),
        concurrency: z.number().int().min(1).max(8).optional().describe("Parallel $batch requests"),
        maxAttempts: z.number().int().min(1).optional().describe("Attempts per request (default 5)"),
        user: z.boolean().optional().describe("Use the signed-in user even if app-only credentials exist"),
      },
      annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false },
    },
    tool(async (args, extra) => {
      const dryRun = args.dryRun ?? true;
      if (!dryRun && opts.readOnly) {
        throw new Error(
          "This dvload MCP server was started with --read-only, so only dry runs are allowed. " +
            "The user can run `dvload run` themselves, or restart the server without --read-only."
        );
      }

      // executeRun reports schema errors on stderr and returns null; an agent
      // needs them in the result, so check first.
      const check = await checkMapping(args.mapping, { remote: false });
      if (check.schemaErrors.length > 0) {
        throw new Error("Mapping has schema errors:\n" + check.schemaErrors.map((e) => `  - ${e}`).join("\n"));
      }

      const progressToken = extra._meta?.progressToken;
      const report = await executeRun(args.mapping, {
        workbook: args.workbook,
        dryRun,
        refresh: args.refresh,
        resume: args.resume,
        maxErrors: args.maxErrors,
        concurrency: args.concurrency,
        maxAttempts: args.maxAttempts,
        user: args.user,
        nonInteractive: true,
        silent: true,
        signal: extra.signal,
        onProgress:
          progressToken === undefined
            ? undefined
            : (progress, total) => {
                // Best-effort: a client that went away must not fail the load.
                void extra
                  .sendNotification({
                    method: "notifications/progress",
                    params: { progressToken, progress, total },
                  })
                  .catch(() => {});
              },
      });
      if (!report) throw new Error("Mapping failed schema validation.");

      const { errors, ...counts } = report;
      return {
        dryRun,
        ...counts,
        warnings: check.warnings,
        errors: errors.slice(0, MAX_ERRORS_RETURNED),
        errorsOmitted: Math.max(0, errors.length - MAX_ERRORS_RETURNED),
      };
    })
  );

  server.registerTool(
    "auth_status",
    {
      title: "Check sign-in",
      description:
        "Report how dvload would authenticate to an environment (appOnly, delegated, or none) and " +
        "whether a token can be acquired right now without prompting. Never returns secrets or tokens.",
      inputSchema: envShape,
      annotations: { readOnlyHint: true },
    },
    tool(async (args) => {
      const environmentUrl = await resolveEnv(args);
      const mode = await detectAuthMode(environmentUrl);
      if (mode === "none") {
        return {
          environmentUrl,
          mode,
          tokenOk: false,
          hint: `Not configured. The user should run \`dvload login --env ${environmentUrl}\`.`,
        };
      }

      const creds = mode === "appOnly" ? await loadAppOnlyCredentials(environmentUrl) : null;
      const account = mode === "delegated" ? await getSignedInAccount(environmentUrl) : null;
      const identity = creds
        ? { clientId: creds.clientId, tenantId: creds.tenantId }
        : { username: account?.username ?? null };

      try {
        const getToken = await getTokenProvider({ environmentUrl, silentOnly: true });
        await getToken();
        return { environmentUrl, mode, ...identity, tokenOk: true };
      } catch (e) {
        return { environmentUrl, mode, ...identity, tokenOk: false, hint: (e as Error).message };
      }
    })
  );

  server.registerTool(
    "list_profiles",
    {
      title: "List environment profiles",
      description: "List saved environment profiles (name → Dataverse environment URL).",
      inputSchema: {},
      annotations: { readOnlyHint: true, openWorldHint: false },
    },
    tool(async () => ({ profiles: await readProfiles() }))
  );

  server.registerTool(
    "list_dataflows",
    {
      title: "List dataflows",
      description:
        "List the Power Platform dataflows in a Dataverse environment, with the queries and target " +
        "tables each one loads. Needs a signed-in session.",
      inputSchema: {
        ...envShape,
        drafts: z.boolean().optional().describe("Also include editing drafts"),
      },
      annotations: { readOnlyHint: true },
    },
    tool(async ({ drafts, ...envArgs }) => {
      const environmentUrl = await resolveEnv(envArgs);
      const getToken = await getTokenProvider({ environmentUrl, silentOnly: true });
      const client = new DataverseClient({ environmentUrl, getToken });
      return { environmentUrl, dataflows: await listDataflows(client, { includeDrafts: drafts }) };
    })
  );

  return server;
}

/**
 * Wrap a handler so its return value becomes JSON text and a throw becomes a
 * tool error the agent can read and act on, rather than a protocol failure.
 */
function tool<A, E>(fn: (args: A, extra: E) => Promise<unknown>) {
  return async (args: A, extra: E): Promise<CallToolResult> => {
    try {
      const value = await fn(args, extra);
      return { content: [{ type: "text", text: JSON.stringify(value, null, 2) }] };
    } catch (e) {
      return { content: [{ type: "text", text: e instanceof Error ? e.message : String(e) }], isError: true };
    } finally {
      // Commands signal failure through process.exitCode; a server outlives
      // the call and must not inherit one tool's failure as its exit status.
      process.exitCode = undefined;
    }
  };
}

export async function mcpCommand(opts: McpOpts): Promise<void> {
  console.log = console.error;
  console.info = console.error;

  const server = buildMcpServer(opts);
  await server.connect(new StdioServerTransport());
  console.error(`dvload MCP server ${TOOL_VERSION} listening on stdio${opts.readOnly ? " (read-only)" : ""}.`);
}
