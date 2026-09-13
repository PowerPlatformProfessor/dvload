// Anonymous usage telemetry (Application Insights). Design rules:
//
//   - OFF unless a connection string is baked in — a fork or local build
//     sends nothing.
//   - Announced: a one-time notice prints before the first event is sent.
//   - Opt-out: `dvload telemetry off`, DVLOAD_TELEMETRY=0, or the
//     cross-tool DO_NOT_TRACK=1 convention.
//   - Fire-and-forget: a 2s cap at exit; telemetry can never slow or fail
//     an import, and offline machines just drop events.
//   - No customer data. Events carry ONLY what TELEMETRY.md documents:
//     tool version, OS, command options (mode, flags), bucketed row
//     counts, and error CODES. Never environment URLs, mapping contents,
//     column names, cell values, or messages.
//   - No persistent identifier. Launches are counted in aggregate — there
//     is no per-install id, so events cannot be linked back to a machine
//     or user (GDPR: nothing here is an online identifier).
//
// Full event reference: TELEMETRY.md in the repo root.

import { readFile, writeFile, mkdir } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import kleur from "kleur";

/** Baked in by scripts/bundle.mjs (esbuild define) when the release is
 *  built with DVLOAD_AI_CONNECTION_STRING set; undefined in a plain tsc
 *  build, where only the runtime env var applies. */
declare const __DVLOAD_AI_DEFAULT__: string | undefined;

/** Runtime env var wins so a baked build can still be pointed elsewhere
 *  (or silenced with an empty value). Empty string = telemetry fully
 *  disabled (nothing is stored or sent). */
const CONNECTION_STRING =
  process.env.DVLOAD_AI_CONNECTION_STRING ??
  (typeof __DVLOAD_AI_DEFAULT__ === "string" ? __DVLOAD_AI_DEFAULT__ : "");

const TOOL_VERSION = "0.2.0"; // keep in sync with package.json

const CONFIG_PATH = path.join(os.homedir(), ".dvload", "telemetry.json");

interface TelemetryConfig {
  enabled: boolean;
  /** First-run notice already shown. */
  notified: boolean;
}

let config: TelemetryConfig | null = null;
const inflight: Promise<unknown>[] = [];

/**
 * GDPR posture: the run where the notice first appears sends NOTHING —
 * the user gets a full opportunity to opt out before any data leaves the
 * machine. Collection starts from the next run.
 */
let noticeJustShown = false;

function parseConnectionString(cs: string): { iKey: string; endpoint: string } | null {
  const iKey = /InstrumentationKey=([^;]+)/i.exec(cs)?.[1];
  if (!iKey) return null;
  const endpoint =
    /IngestionEndpoint=([^;]+)/i.exec(cs)?.[1]?.replace(/\/$/, "") ??
    "https://dc.services.visualstudio.com";
  return { iKey, endpoint };
}

async function loadConfig(): Promise<TelemetryConfig> {
  if (config) return config;
  try {
    config = JSON.parse(await readFile(CONFIG_PATH, "utf8")) as TelemetryConfig;
  } catch {
    config = { enabled: true, notified: false };
    await saveConfig().catch(() => {});
  }
  return config;
}

async function saveConfig(): Promise<void> {
  if (!config) return;
  await mkdir(path.dirname(CONFIG_PATH), { recursive: true });
  await writeFile(CONFIG_PATH, JSON.stringify(config, null, 2), "utf8");
}

function envOptedOut(): boolean {
  return process.env.DVLOAD_TELEMETRY === "0" || !!process.env.DO_NOT_TRACK;
}

export async function telemetryEnabled(): Promise<boolean> {
  if (!CONNECTION_STRING || envOptedOut()) return false;
  return (await loadConfig()).enabled;
}

/**
 * One-time notice, printed through the caller's logger so --json runs stay
 * clean. Prints only when telemetry would actually send.
 */
export async function telemetryNotice(log: (msg: string) => void): Promise<void> {
  if (!(await telemetryEnabled())) return;
  const cfg = await loadConfig();
  if (cfg.notified) return;
  cfg.notified = true;
  noticeJustShown = true;
  await saveConfig().catch(() => {});
  log(
    kleur.gray(
      "dvload collects anonymous usage data (command, options, bucketed row counts — never\n" +
        "your data, URLs, or mapping contents) to improve the tool. Details: TELEMETRY.md.\n" +
        "Disable any time: `dvload telemetry off` or DVLOAD_TELEMETRY=0."
    )
  );
}

/** Round a count into a coarse bucket so no exact volumes leave the machine. */
export function bucket(n: number): string {
  if (n <= 0) return "0";
  if (n <= 100) return "1-100";
  if (n <= 10000) return "101-10k";
  if (n <= 100000) return "10k-100k";
  return ">100k";
}

/** Fire-and-forget event. Never throws; never awaited by business logic. */
export function track(name: string, properties: Record<string, string> = {}): void {
  const p = (async () => {
    if (noticeJustShown) return; // first run is notice-only
    if (!(await telemetryEnabled())) return;
    const conn = parseConnectionString(CONNECTION_STRING);
    if (!conn) return;
    const envelope = {
      name: "Microsoft.ApplicationInsights.Event",
      time: new Date().toISOString(),
      iKey: conn.iKey,
      tags: { "ai.cloud.roleInstance": "cli" },
      data: {
        baseType: "EventData",
        baseData: {
          ver: 2,
          name,
          properties: {
            ...properties,
            toolVersion: TOOL_VERSION,
            os: process.platform,
          },
        },
      },
    };
    await fetch(`${conn.endpoint}/v2/track`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify([envelope]),
      signal: AbortSignal.timeout(2000),
    });
  })().catch(() => {});
  inflight.push(p);
}

/** Give pending events up to 2s to leave, then move on. Call before exit. */
export async function flushTelemetry(): Promise<void> {
  if (inflight.length === 0) return;
  await Promise.race([
    Promise.allSettled(inflight),
    new Promise((r) => setTimeout(r, 2000)),
  ]);
}

/** `dvload telemetry [on|off|status]` */
export async function telemetryCommand(action?: string): Promise<void> {
  const cfg = await loadConfig();
  switch (action) {
    case "on":
      cfg.enabled = true;
      await saveConfig();
      console.log(kleur.green("Telemetry enabled."));
      return;
    case "off":
      cfg.enabled = false;
      await saveConfig();
      console.log(kleur.green("Telemetry disabled. Nothing will be sent."));
      return;
    default: {
      const active = await telemetryEnabled();
      console.log(`Telemetry: ${active ? kleur.green("on") : kleur.yellow("off")}`);
      if (!CONNECTION_STRING) console.log(kleur.gray("  (no connection string baked into this build — nothing can be sent)"));
      if (envOptedOut()) console.log(kleur.gray("  (disabled via DVLOAD_TELEMETRY=0 / DO_NOT_TRACK)"));
      console.log(kleur.gray(`  Config: ${CONFIG_PATH}`));
      console.log(kleur.gray("  No install id: launches are counted in aggregate, not linked to you."));
      console.log(kleur.gray("  Events and fields are documented in TELEMETRY.md."));
      console.log(kleur.gray("  Example event: {\"name\":\"cli_run\",\"mode\":\"insert\",\"rows\":\"101-10k\",\"outcome\":\"ok\"}"));
    }
  }
}
