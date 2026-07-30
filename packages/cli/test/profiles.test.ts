/**
 * Named environment profiles (~/.dvload/profiles.json).
 *
 * Small module, previously 0% covered, and the one that decides WHICH
 * Dataverse environment a command writes to. A resolution bug here doesn't
 * throw — it quietly loads production data into the wrong org.
 */

import { describe, it, expect, afterEach } from "vitest";
import { readFile, writeFile, mkdir } from "node:fs/promises";
import path from "node:path";
import { withFakeHome, type FakeHome } from "./support/fake-home.js";

type ProfilesModule = typeof import("../src/profiles.js");

let ctx: ({ mod: ProfilesModule } & FakeHome) | undefined;

async function load(): Promise<{ mod: ProfilesModule } & FakeHome> {
  ctx = await withFakeHome<ProfilesModule>(() => import("../src/profiles.js"));
  return ctx;
}

afterEach(() => {
  ctx?.cleanup();
  ctx = undefined;
});

describe("readProfiles", () => {
  it("returns an empty map when the file has never been created", async () => {
    // A first-run user has no ~/.dvload at all. That is not an error.
    const { mod } = await load();
    await expect(mod.readProfiles()).resolves.toEqual({});
  });

  it("refuses to treat malformed JSON as 'no profiles'", async () => {
    // The dangerous alternative: swallow the parse error, return {}, and tell
    // the user their profile "doesn't exist" — sending them off to re-add a
    // profile that is already there, over a stray comma.
    const { mod, dvloadDir } = await load();
    await mkdir(dvloadDir, { recursive: true });
    await writeFile(path.join(dvloadDir, "profiles.json"), '{ "prod": "https://x", }', "utf8");

    await expect(mod.readProfiles()).rejects.toThrow(/not valid JSON/i);
    // And it must say which file to fix.
    await expect(mod.readProfiles()).rejects.toThrow(/profiles\.json/);
  });
});

describe("writeProfiles", () => {
  it("creates ~/.dvload on demand and round-trips", async () => {
    const { mod } = await load();
    const profiles = { prod: "https://contoso.crm.dynamics.com", dev: "https://dev.crm.dynamics.com" };

    await mod.writeProfiles(profiles);

    expect(await mod.readProfiles()).toEqual(profiles);
  });

  it("writes human-editable JSON with a trailing newline", async () => {
    // Users do edit this file by hand; the docs tell them they can.
    const { mod, dvloadDir } = await load();
    await mod.writeProfiles({ prod: "https://contoso.crm.dynamics.com" });

    const raw = await readFile(path.join(dvloadDir, "profiles.json"), "utf8");
    expect(raw).toContain("\n  ");
    expect(raw.endsWith("\n")).toBe(true);
  });

  it("replaces the file rather than merging into it", async () => {
    // writeProfiles takes the whole map; the caller owns the merge. If this
    // ever started merging, `dvload profile remove` would silently no-op.
    const { mod } = await load();
    await mod.writeProfiles({ a: "https://a.crm.dynamics.com", b: "https://b.crm.dynamics.com" });
    await mod.writeProfiles({ a: "https://a.crm.dynamics.com" });

    expect(await mod.readProfiles()).toEqual({ a: "https://a.crm.dynamics.com" });
  });
});

describe("resolveEnv", () => {
  it("prefers an explicit --env over any profile", async () => {
    const { mod } = await load();
    await mod.writeProfiles({ prod: "https://profile.crm.dynamics.com" });

    const url = await mod.resolveEnv({
      env: "https://explicit.crm.dynamics.com",
      profile: "prod",
    });

    // Explicit beats saved. Otherwise `--env` would be a suggestion, and a
    // stale profile could redirect a one-off command somewhere unintended.
    expect(url).toBe("https://explicit.crm.dynamics.com");
  });

  it("resolves a saved profile by name", async () => {
    const { mod } = await load();
    await mod.writeProfiles({ prod: "https://contoso.crm.dynamics.com" });

    await expect(mod.resolveEnv({ profile: "prod" })).resolves.toBe("https://contoso.crm.dynamics.com");
  });

  it("lists the available names when the profile is unknown", async () => {
    const { mod } = await load();
    await mod.writeProfiles({ prod: "https://a", uat: "https://b" });

    const err = await mod.resolveEnv({ profile: "prodd" }).catch((e: Error) => e);

    expect(err).toBeInstanceOf(Error);
    // The typo case is the common one; showing the real names turns a dead
    // end into a one-second fix.
    expect((err as Error).message).toMatch(/prodd/);
    expect((err as Error).message).toMatch(/prod/);
    expect((err as Error).message).toMatch(/uat/);
  });

  it("points a first-time user at `profile add` instead of an empty list", async () => {
    const { mod } = await load();
    await expect(mod.resolveEnv({ profile: "prod" })).rejects.toThrow(/profile add/);
  });

  it("asks for one of the two flags when given neither", async () => {
    const { mod } = await load();
    await expect(mod.resolveEnv({})).rejects.toThrow(/--env|--profile/);
  });

  it("treats an empty --env as absent rather than as a valid URL", async () => {
    // `--env ""` from a shell variable that didn't expand is a real mistake;
    // resolving it to "" would produce requests against `/api/data/v9.2/…`.
    const { mod } = await load();
    await expect(mod.resolveEnv({ env: "" })).rejects.toThrow(/--env|--profile/);
  });
});
