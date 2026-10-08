# Using dvload from AI agents

dvload can be driven by an AI agent — Claude Code, Claude Desktop, GitHub
Copilot in VS Code, Cursor, or anything else that can run a local command or
speak the [Model Context Protocol](https://modelcontextprotocol.io). There are
two ways in, and both use the same engine, the same `.dvmap.json` files and
the same credentials in `~/.dvload/` as the CLI you run by hand.

| | How the agent calls dvload | Use it when |
|---|---|---|
| **Shell** | runs `dvload …` like you would | the agent already has a terminal (Claude Code, Copilot agent mode, Codex) |
| **MCP** | typed tools from `dvload mcp` | you want guard rails (dry run by default, `--read-only`), or the agent has no shell (Claude Desktop) |

Both are **local**: dvload has to be installed on the machine the agent runs
on, and files are read from that machine's disk. A cloud-hosted agent such as
Copilot Studio cannot start a local program, so neither route reaches it —
see [What this does not cover](#what-this-does-not-cover).

## Before either: sign in yourself

An agent cannot complete a browser sign-in, and dvload never starts one on an
agent's behalf. Sign in once, in a terminal:

```bash
dvload login --env https://yourorg.crm.dynamics.com
```

After that the cached session is used silently. When it expires, the agent
gets an error naming the `dvload login` command to run, and you run it. For
fully unattended setups use `dvload app-login` instead
([SCHEDULED-RUNS.md](./SCHEDULED-RUNS.md)).

## MCP server

```bash
dvload mcp              # stdio server; agents may run real imports
dvload mcp --read-only  # inspect, validate and dry-run only
```

You don't run this yourself — the agent's host starts it. Register it once:

**Claude Code**

```bash
claude mcp add dvload -- dvload mcp
```

**Claude Desktop** (`claude_desktop_config.json`) and **Cursor** (`.cursor/mcp.json`)

```json
{ "mcpServers": { "dvload": { "command": "dvload", "args": ["mcp"] } } }
```

**VS Code / GitHub Copilot** (`.vscode/mcp.json`)

```json
{ "servers": { "dvload": { "type": "stdio", "command": "dvload", "args": ["mcp"] } } }
```

Add `"--read-only"` to `args` to stop the agent from writing to Dataverse at
all.

### Tools

| Tool | Writes? | What it does |
|---|---|---|
| `inspect_source` | no | Tables in an `.xlsx` (name, sheet, columns, row count) or the columns of a `.csv`/`.tsv`; optionally the first rows as a sample (max 50). |
| `validate_mapping` | no | Schema errors, warnings, and target attributes missing from Dataverse. `remote: false` skips the Dataverse check. |
| `run_mapping` | **only with `dryRun: false`** | Runs a mapping against a source file. Returns the same counts as `dvload run --json`, the first 20 row errors, and the failed-rows file path. |
| `auth_status` | no | Auth mode for an environment (`appOnly` / `delegated` / `none`) and whether a token can be acquired silently. Never returns secrets or tokens. |
| `list_profiles` | no | Saved environment profiles. |
| `list_dataflows` | no | Power Platform dataflows in an environment. |

Paths are resolved on the machine running the server; agents should pass
absolute paths.

### Safety model

- **`run_mapping` is a dry run unless the agent passes `dryRun: false`.** A
  dry run reads, coerces and plans every row but sends nothing to Dataverse.
- **`--read-only`** makes the server refuse `dryRun: false` outright.
- **No interactive auth.** A missing or expired session is a tool error, not
  a browser window. A real run checks for a token *before* loading, so it
  fails once rather than failing every row.
- **Same side effects as the CLI.** Every run, dry or not, writes its `.jsonl` log
  next to the mapping (`logDir`); a real run with failures writes the
  failed-rows `.xlsx`; `mapping.notifyUrl` is honoured on real runs. A sync
  mapping deactivates or deletes target records exactly as `dvload run`
  would.
- Your MCP host's own per-tool approval prompt still applies. `run_mapping`
  is annotated as destructive so hosts that honour tool annotations ask.

Long runs report progress (`notifications/progress`) when the client asks for
it, and honour cancellation: the load stops between batches and the result
comes back with `cancelled: true`.

Not exposed over MCP: `run-all`, `schedule`, the login commands, and the
`.pqt`/dataflow import commands. Use the shell for those.

## Shell

For an agent with a terminal, these are the commands and flags that make
dvload safe to script. Paste this section into your agent's instructions file
(`CLAUDE.md`, `AGENTS.md`, `.github/copilot-instructions.md`) if you want it
to know them up front.

```bash
# What's configured? Exit 1 if no token can be acquired.
dvload whoami --env <url>

# Schema-only check (no network), then the Dataverse metadata check.
dvload validate <mapping> --no-remote
dvload validate <mapping>

# Always dry-run first. --json prints one JSON object on stdout.
dvload run <mapping> -w <workbook> --dry-run --json --non-interactive

# The real run.
dvload run <mapping> -w <workbook> --json --non-interactive
```

- **`--json`** suppresses progress output and prints the result as a single
  JSON object: `total`, `succeeded`, `created`, `updated`, `skipped`,
  `unchanged`, `removed`, `failed`, `startedAt`, `finishedAt`, `errors[]`
  (`rowIndex`, `message`, `code`, `httpStatus`, `sourceRow`),
  `failedRowsFile`, `retries`. See
  [DATA-FORMATS.md](./DATA-FORMATS.md#--json-output).
- **`--non-interactive`** fails fast instead of waiting on a sign-in prompt.
  It is already the default when stdout is not a terminal.
- **Exit codes for `run`:** `0` all rows loaded, `1` some rows failed (or any
  other error), `2` the mapping failed schema validation.
- **Mapping format:** [`docs/schema/dvmap.schema.json`](./schema/dvmap.schema.json)
  is the JSON Schema; [RECIPES.md](./RECIPES.md) has worked examples for
  lookups, upserts, choices and run plans.
- `dvload dataflows --json` lists dataflows as JSON. Other commands print
  text meant for people.

## What this does not cover

**Copilot Studio and other cloud-hosted agents.** They can only call MCP
servers reachable over HTTP, and they cannot read files on your PC. Serving
them would mean hosting dvload's engine as a web service with server-side
credentials and a way to hand it files (SharePoint, blob storage) — a
different product from a local CLI, and not something `dvload mcp` does.
