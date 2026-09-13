// Anonymous usage telemetry for the add-in (Application Insights). Same
// rules as the CLI module: off unless a connection string is baked in at
// build time, a visible checkbox controls it, events carry only what
// TELEMETRY.md documents (never environment URLs, mapping contents,
// column names, or cell values), and sending is fire-and-forget.
//
// No persistent identifier is stored: launches are counted in aggregate,
// so no localStorage id or cookie is written and events can't be linked
// across sessions (GDPR: nothing here is an online identifier).

declare const ADDIN_AI_CONNECTION: string; // injected by webpack DefinePlugin

const ENABLED_KEY = "dvload:telemetry-enabled";
const NOTIFIED_KEY = "dvload:telemetry-notified";
const TOOL_VERSION = "0.2.0"; // keep in sync with package.json

// GDPR posture: the very first pane session sends NOTHING — the user sees
// the checkbox and can untick it before any data leaves the machine.
// Collection starts from the next session.
const firstSession = ((): boolean => {
  try {
    if (localStorage.getItem(NOTIFIED_KEY)) return false;
    localStorage.setItem(NOTIFIED_KEY, "1");
    return true;
  } catch {
    return true; // storage unavailable → err on the silent side
  }
})();

function connection(): { iKey: string; endpoint: string } | null {
  const cs = typeof ADDIN_AI_CONNECTION === "string" ? ADDIN_AI_CONNECTION : "";
  const iKey = /InstrumentationKey=([^;]+)/i.exec(cs)?.[1];
  if (!iKey) return null;
  const endpoint =
    /IngestionEndpoint=([^;]+)/i.exec(cs)?.[1]?.replace(/\/$/, "") ??
    "https://dc.services.visualstudio.com";
  return { iKey, endpoint };
}

export function telemetryAvailable(): boolean {
  return connection() !== null;
}

export function telemetryEnabled(): boolean {
  if (!telemetryAvailable()) return false;
  return localStorage.getItem(ENABLED_KEY) !== "0";
}

export function setTelemetryEnabled(on: boolean): void {
  localStorage.setItem(ENABLED_KEY, on ? "1" : "0");
}

/** Fire-and-forget event. Never throws, never blocks the UI. */
export function track(name: string, properties: Record<string, string> = {}): void {
  try {
    if (firstSession) return; // first session is notice-only
    if (!telemetryEnabled()) return;
    const conn = connection();
    if (!conn) return;
    const envelope = {
      name: "Microsoft.ApplicationInsights.Event",
      time: new Date().toISOString(),
      iKey: conn.iKey,
      tags: { "ai.cloud.roleInstance": "addin" },
      data: {
        baseType: "EventData",
        baseData: {
          ver: 2,
          name,
          properties: { ...properties, toolVersion: TOOL_VERSION },
        },
      },
    };
    void fetch(`${conn.endpoint}/v2/track`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify([envelope]),
      keepalive: true, // survives pane teardown for end-of-run events
    }).catch(() => {});
  } catch {
    // telemetry must never break the pane
  }
}

/** Coarse buckets so exact volumes never leave the machine. */
export function bucket(n: number): string {
  if (n <= 0) return "0";
  if (n <= 100) return "1-100";
  if (n <= 10000) return "101-10k";
  if (n <= 100000) return "10k-100k";
  return ">100k";
}
