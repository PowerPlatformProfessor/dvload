/**
 * Telemetry.
 *
 * The risk here is not "does it work" — nobody is harmed by a dropped event.
 * The risk is that it sends something it promised not to. TELEMETRY.md makes
 * specific commitments to users, and this file turns each of them into a test
 * that fails if the commitment is broken:
 *
 *   1. Off entirely unless a connection string is baked into the build.
 *   2. Honours `dvload telemetry off`, DVLOAD_TELEMETRY=0, and DO_NOT_TRACK.
 *   3. The run that first shows the notice sends NOTHING, so the user has a
 *      real opportunity to opt out before any data leaves the machine.
 *   4. Row counts are bucketed, never exact.
 *   5. No environment URL, tenant id, mapping content, column name, cell
 *      value or error message ever appears in a payload.
 *
 * Every test captures the actual HTTP body rather than trusting the shape of
 * the code — (5) in particular is only meaningful if asserted on the bytes.
 */

import { describe, it, expect, afterEach, vi } from "vitest";
import { mkdtempSync, rmSync, writeFileSync, mkdirSync, readFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

type TelemetryModule = typeof import("../src/telemetry.js");

interface Ctx {
  mod: TelemetryModule;
  home: string;
  dvloadDir: string;
  /** Bodies POSTed to the ingestion endpoint, parsed. */
  sent: unknown[];
  cleanup(): void;
}

let ctx: Ctx | undefined;

const CONNECTION_STRING =
  "InstrumentationKey=00000000-0000-0000-0000-000000000001;IngestionEndpoint=https://telemetry.invalid/";

/**
 * Load telemetry.ts with a controlled environment.
 *
 * CONNECTION_STRING and CONFIG_PATH are both module-level constants, so the
 * env and the homedir mock have to be in place before the import.
 */
async function load(env: Record<string, string | undefined> = {}): Promise<Ctx> {
  const home = mkdtempSync(path.join(tmpdir(), "dvload-tele-"));

  for (const key of ["DVLOAD_AI_CONNECTION_STRING", "DVLOAD_TELEMETRY", "DO_NOT_TRACK"]) {
    if (key in env) {
      const v = env[key];
      if (v === undefined) vi.stubEnv(key, "");
      else vi.stubEnv(key, v);
    } else {
      vi.stubEnv(key, "");
    }
  }

  vi.doMock("node:os", async () => {
    const actual = await vi.importActual<typeof import("node:os")>("node:os");
    return { ...actual, default: { ...actual, homedir: () => home }, homedir: () => home };
  });

  const sent: unknown[] = [];
  vi.stubGlobal(
    "fetch",
    vi.fn((_url: string, init?: { body?: string }) => {
      if (init?.body) sent.push(JSON.parse(init.body));
      return Promise.resolve(new Response("{}", { status: 200 }));
    })
  );

  vi.resetModules();
  const mod = await import("../src/telemetry.js");

  ctx = {
    mod,
    home,
    dvloadDir: path.join(home, ".dvload"),
    sent,
    cleanup: () => {
      vi.doUnmock("node:os");
      vi.unstubAllGlobals();
      vi.unstubAllEnvs();
      rmSync(home, { recursive: true, force: true });
    },
  };
  return ctx;
}

/** Pre-seed the on-disk config so a test can skip the first-run notice. */
function writeConfig(dvloadDir: string, cfg: { enabled: boolean; notified: boolean }): void {
  mkdirSync(dvloadDir, { recursive: true });
  writeFileSync(path.join(dvloadDir, "telemetry.json"), JSON.stringify(cfg), "utf8");
}

afterEach(() => {
  ctx?.cleanup();
  ctx = undefined;
});

/* -------------------------------------------------------------------------- */
/* 1. Off by default                                                           */
/* -------------------------------------------------------------------------- */

describe("disabled builds", () => {
  it("is off when no connection string is compiled in", async () => {
    // A fork, a local build, or the current state of this repo. Nothing is
    // stored and nothing can be sent.
    const { mod } = await load({ DVLOAD_AI_CONNECTION_STRING: "" });
    await expect(mod.telemetryEnabled()).resolves.toBe(false);
  });

  it("sends nothing even when track() is called explicitly", async () => {
    const { mod, sent } = await load({ DVLOAD_AI_CONNECTION_STRING: "" });

    mod.track("cli_run", { mode: "insert" });
    await mod.flushTelemetry();

    expect(sent).toHaveLength(0);
  });

  it("prints no first-run notice", async () => {
    const { mod } = await load({ DVLOAD_AI_CONNECTION_STRING: "" });
    const lines: string[] = [];

    await mod.telemetryNotice((m) => lines.push(m));

    expect(lines).toHaveLength(0);
  });
});

/* -------------------------------------------------------------------------- */
/* 2. Opt-out                                                                  */
/* -------------------------------------------------------------------------- */

describe("opt-out", () => {
  it("honours DVLOAD_TELEMETRY=0", async () => {
    const { mod, sent } = await load({
      DVLOAD_AI_CONNECTION_STRING: CONNECTION_STRING,
      DVLOAD_TELEMETRY: "0",
    });

    await expect(mod.telemetryEnabled()).resolves.toBe(false);
    mod.track("cli_run");
    await mod.flushTelemetry();
    expect(sent).toHaveLength(0);
  });

  it("honours the cross-tool DO_NOT_TRACK convention", async () => {
    const { mod, sent } = await load({
      DVLOAD_AI_CONNECTION_STRING: CONNECTION_STRING,
      DO_NOT_TRACK: "1",
    });

    await expect(mod.telemetryEnabled()).resolves.toBe(false);
    mod.track("cli_run");
    await mod.flushTelemetry();
    expect(sent).toHaveLength(0);
  });

  it("honours a stored `dvload telemetry off`", async () => {
    const { mod, dvloadDir, sent } = await load({ DVLOAD_AI_CONNECTION_STRING: CONNECTION_STRING });
    writeConfig(dvloadDir, { enabled: false, notified: true });

    await expect(mod.telemetryEnabled()).resolves.toBe(false);
    mod.track("cli_run");
    await mod.flushTelemetry();
    expect(sent).toHaveLength(0);
  });

  it("persists the off switch so the next run stays off", async () => {
    const { mod, dvloadDir } = await load({ DVLOAD_AI_CONNECTION_STRING: CONNECTION_STRING });
    writeConfig(dvloadDir, { enabled: true, notified: true });

    await mod.telemetryCommand("off");

    const cfg = JSON.parse(readFileSync(path.join(dvloadDir, "telemetry.json"), "utf8"));
    expect(cfg.enabled).toBe(false);
  });
});

/* -------------------------------------------------------------------------- */
/* 3. The first run sends nothing                                              */
/* -------------------------------------------------------------------------- */

describe("first-run notice", () => {
  it("shows the notice and sends nothing on that same run", async () => {
    // The GDPR posture stated in the module header. If this regressed, a
    // user would have data collected before they could possibly opt out.
    const { mod, sent } = await load({ DVLOAD_AI_CONNECTION_STRING: CONNECTION_STRING });
    const lines: string[] = [];

    await mod.telemetryNotice((m) => lines.push(m));
    mod.track("cli_run", { mode: "insert" });
    await mod.flushTelemetry();

    expect(lines.join("\n")).toMatch(/anonymous usage data/i);
    expect(lines.join("\n")).toMatch(/TELEMETRY\.md/);
    expect(sent).toHaveLength(0);
  });

  it("tells the user how to turn it off, in the notice itself", async () => {
    const { mod } = await load({ DVLOAD_AI_CONNECTION_STRING: CONNECTION_STRING });
    const lines: string[] = [];

    await mod.telemetryNotice((m) => lines.push(m));

    expect(lines.join("\n")).toMatch(/telemetry off/);
    expect(lines.join("\n")).toMatch(/DVLOAD_TELEMETRY=0/);
  });

  it("shows the notice only once", async () => {
    const { mod, dvloadDir } = await load({ DVLOAD_AI_CONNECTION_STRING: CONNECTION_STRING });
    const lines: string[] = [];

    await mod.telemetryNotice((m) => lines.push(m));
    await mod.telemetryNotice((m) => lines.push(m));

    expect(lines).toHaveLength(1);
    const cfg = JSON.parse(readFileSync(path.join(dvloadDir, "telemetry.json"), "utf8"));
    expect(cfg.notified).toBe(true);
  });

  it("routes the notice through the caller's logger so --json output stays parseable", async () => {
    // A stray line on stdout would break `dvload run --json | jq`.
    const { mod } = await load({ DVLOAD_AI_CONNECTION_STRING: CONNECTION_STRING });
    const spy = vi.spyOn(console, "log").mockImplementation(() => {});

    await mod.telemetryNotice(() => {});

    expect(spy).not.toHaveBeenCalled();
    spy.mockRestore();
  });
});

/* -------------------------------------------------------------------------- */
/* 4. Bucketing                                                                */
/* -------------------------------------------------------------------------- */

describe("bucket", () => {
  it("maps counts to the documented ranges, including the boundaries", async () => {
    const { mod } = await load();

    expect(mod.bucket(0)).toBe("0");
    expect(mod.bucket(-5)).toBe("0");
    expect(mod.bucket(1)).toBe("1-100");
    expect(mod.bucket(100)).toBe("1-100");
    expect(mod.bucket(101)).toBe("101-10k");
    expect(mod.bucket(10_000)).toBe("101-10k");
    expect(mod.bucket(10_001)).toBe("10k-100k");
    expect(mod.bucket(100_000)).toBe("10k-100k");
    expect(mod.bucket(100_001)).toBe(">100k");
  });

  it("never returns anything that could be read as an exact count", async () => {
    // The commitment is that volumes are coarse. A bucket that happened to
    // be a bare number would leak the row count it was meant to hide.
    const { mod } = await load();
    for (const n of [0, 1, 7, 99, 100, 101, 5000, 10_000, 99_999, 1_000_000]) {
      const b = mod.bucket(n);
      if (n === 0) continue; // "0" is not sensitive
      expect(b).not.toBe(String(n));
      expect(b).toMatch(/^(1-100|101-10k|10k-100k|>100k)$/);
    }
  });
});

/* -------------------------------------------------------------------------- */
/* 5. What actually goes over the wire                                         */
/* -------------------------------------------------------------------------- */

describe("payload contents", () => {
  /** Enabled, already notified — the steady state after the first run. */
  async function enabled(): Promise<Ctx> {
    const c = await load({ DVLOAD_AI_CONNECTION_STRING: CONNECTION_STRING });
    writeConfig(c.dvloadDir, { enabled: true, notified: true });
    return c;
  }

  it("sends an event once telemetry is on and the notice is behind us", async () => {
    const { mod, sent } = await enabled();

    mod.track("cli_run", { mode: "insert", outcome: "ok" });
    await mod.flushTelemetry();

    expect(sent).toHaveLength(1);
  });

  it("carries only the documented fields", async () => {
    const { mod, sent } = await enabled();

    mod.track("cli_run", { mode: "upsert", rows: "101-10k", outcome: "ok" });
    await mod.flushTelemetry();

    const envelope = (sent[0] as Array<Record<string, never>>)[0];
    const props = (envelope as unknown as { data: { baseData: { properties: Record<string, string> } } }).data
      .baseData.properties;

    expect(Object.keys(props).sort()).toEqual(["mode", "os", "outcome", "rows", "toolVersion"]);
    expect(props.os).toBe(process.platform);
  });

  it("contains no per-install identifier", async () => {
    // Stated in the module header: launches are counted in aggregate, so
    // events cannot be linked back to a machine or user. A GUID appearing
    // anywhere in the payload would break that.
    const { mod, sent } = await enabled();

    mod.track("cli_run", { mode: "insert" });
    await mod.flushTelemetry();

    const body = JSON.stringify(sent[0]);
    const guids = body.match(/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/gi) ?? [];
    // The instrumentation key is the only GUID that legitimately appears.
    expect(guids).toEqual(["00000000-0000-0000-0000-000000000001"]);
  });

  it("never leaks an environment URL, tenant id or mapping detail", async () => {
    // The single most important test in this file. Someone adding a
    // "helpful" diagnostic property is the realistic way this breaks, and
    // it would ship customer infrastructure names to a telemetry endpoint.
    const { mod, sent } = await enabled();

    mod.track("cli_run", {
      mode: "upsert",
      rows: "101-10k",
      outcome: "error",
      errorCode: "0x80040333",
    });
    await mod.flushTelemetry();

    const body = JSON.stringify(sent[0]);
    for (const forbidden of ["crm.dynamics.com", "contoso", "emailaddress1", ".dvmap.json", "C:\\", "@"]) {
      expect(body).not.toContain(forbidden);
    }
  });

  it("posts to the endpoint from the connection string, not a hardcoded one", async () => {
    const { mod, sent } = await enabled();
    const fetchMock = globalThis.fetch as unknown as ReturnType<typeof vi.fn>;

    mod.track("cli_run");
    await mod.flushTelemetry();

    expect(sent).toHaveLength(1);
    expect(fetchMock.mock.calls[0][0]).toBe("https://telemetry.invalid/v2/track");
  });

  it("cannot fail a command when the endpoint is unreachable", async () => {
    // Fire-and-forget: an offline machine, a blocked proxy, or a 500 must
    // never surface to the user or reject a flush.
    const { mod } = await enabled();
    vi.stubGlobal(
      "fetch",
      vi.fn(() => Promise.reject(new Error("getaddrinfo ENOTFOUND telemetry.invalid")))
    );

    expect(() => mod.track("cli_run")).not.toThrow();
    await expect(mod.flushTelemetry()).resolves.toBeUndefined();
  });

  it("returns immediately when there is nothing queued", async () => {
    const { mod } = await enabled();
    await expect(mod.flushTelemetry()).resolves.toBeUndefined();
  });
});
