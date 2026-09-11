// What a failed silent token refresh is allowed to do.
//
// The bug these tests pin down: `dvload serve` would serve the pane happily
// for an hour, then answer one request with "No response from the browser
// after 180s", and work again as soon as the user reloaded. Both halves came
// from the same line — a single transient `acquireTokenSilent` failure was
// escalated to a full interactive login, which inside an API handler means
// opening a browser and blocking the response until the auth timeout. The
// reload "fixed" it because the next silent refresh succeeded, which is
// exactly why the symptom looked like a flaky connection.

import { test, vi, beforeEach, afterEach } from "vitest";
import assert from "node:assert/strict";

const ENV = "https://contoso.crm.dynamics.com";
const HOST = "contoso.crm.dynamics.com";
const SHARED = "51f81489-12ee-4a9e-aaae-a2591f45987d";
const DVLOAD = "e6828b0f-9fde-43f8-85d0-602660d498bb";

/** Values the secure store hands back. Mutated per test. */
const stored = vi.hoisted(() => ({ map: new Map<string, string>() }));

vi.mock("./secure-store.js", () => ({
  getSecret: (k: string) => Promise.resolve(stored.map.get(k) ?? null),
  setSecret: (k: string, v: string) => {
    stored.map.set(k, v);
    return Promise.resolve();
  },
  deleteSecret: () => Promise.resolve(),
  listAccounts: () => Promise.resolve([]),
  protectBytes: (b: Buffer) => Promise.resolve(b),
  unprotectBytes: (b: Buffer) => Promise.resolve(b),
}));

/** Controls what the fake MSAL app does, and records what was called. */
const msal = vi.hoisted(() => ({
  silent: null as null | (() => Promise<unknown>),
  /** Per-attempt interactive behaviour; receives the client id being tried. */
  interactive: null as null | ((clientId: string) => Promise<unknown>),
  interactiveClientIds: [] as string[],
  clientIdsConstructed: [] as string[],
}));

vi.mock("@azure/msal-node", async (importOriginal) => {
  const actual = await importOriginal<typeof import("@azure/msal-node")>();
  class FakePublicClientApplication {
    private readonly clientId: string;
    constructor(config: { auth: { clientId: string } }) {
      this.clientId = config.auth.clientId;
      msal.clientIdsConstructed.push(this.clientId);
    }
    getTokenCache() {
      return {
        getAccountByHomeId: (id: string) =>
          Promise.resolve({ homeAccountId: id, username: "u@contoso.com", tenantId: "t" }),
      };
    }
    acquireTokenSilent() {
      return msal.silent!();
    }
    /** The real one opens a browser here. Never in a unit test. */
    acquireTokenInteractive() {
      msal.interactiveClientIds.push(this.clientId);
      return msal.interactive ? msal.interactive(this.clientId) : new Promise(() => {});
    }
    acquireTokenByDeviceCode() {
      msal.interactiveClientIds.push(this.clientId);
      return new Promise(() => {});
    }
  }
  return { ...actual, PublicClientApplication: FakePublicClientApplication };
});

const { getTokenProvider, needsInteractiveSignIn, isInteractiveSignInRequired, clearLoopbackProbeCache } =
  await import("./auth.js");
const { InteractionRequiredAuthError, ClientAuthError, ServerError } = await import("@azure/msal-node");

/** What a refresh that failed on the wire looks like coming out of MSAL. */
const networkFailure = (): Error => new ClientAuthError("network_error", "socket hang up");

/**
 * Stand in for Entra's answer to the loopback preflight.
 *
 * The shape is the signal, not the error code: a registered redirect URI gets
 * a 302 back to localhost, an unregistered one gets a rendered page whose
 * AADSTS code is about the missing session (50058) and says nothing about
 * redirects at all. See probeLoopbackRedirect.
 */
function stubProbe(verdicts: Record<string, "usable" | "unusable">): void {
  vi.stubGlobal("fetch", (url: string) => {
    const clientId = new URL(url).searchParams.get("client_id") ?? "";
    // Default unusable: a test that cares must say so.
    return Promise.resolve(
      (verdicts[clientId] ?? "unusable") === "usable"
        ? {
            status: 302,
            headers: { get: () => "http://localhost:53682/?error=login_required" },
            text: () => Promise.resolve(""),
          }
        : {
            status: 200,
            headers: { get: () => null },
            text: () => Promise.resolve("AADSTS50058: Silent sign-in request was sent."),
          }
    );
  });
}

beforeEach(() => {
  msal.silent = null;
  msal.interactive = null;
  msal.interactiveClientIds = [];
  msal.clientIdsConstructed = [];
  // Verdicts are cached per client id for the process lifetime, and these
  // tests need opposite answers for the same real app ids.
  clearLoopbackProbeCache();
  // No test may touch the network. The one caller that would is the loopback
  // preflight in acquireInteractive; a rejected fetch is its "inconclusive"
  // verdict, which is also the real-world state that let this bug through —
  // a blocked probe is why dvload opened a browser for an app that could
  // never complete the redirect.
  vi.stubGlobal("fetch", () => Promise.reject(new Error("offline")));
  stored.map = new Map([
    [`delegated:${HOST}`, "home-account-id"],
    [`delegatedTenant:${HOST}`, "tenant-guid"],
    // The shared Microsoft client: fine for silent refresh, but it has no
    // http://localhost redirect URI, so an interactive escalation pinned to
    // it can only ever time out. This is the state a device-code login
    // leaves behind, and it is what made the hang permanent rather than rare.
    [`delegatedClient:${HOST}`, SHARED],
  ]);
});

afterEach(() => {
  vi.unstubAllGlobals();
});

/* -------------------------------------------------------------------------- */
/* Classification                                                              */
/* -------------------------------------------------------------------------- */

test("needsInteractiveSignIn is true only for MSAL's interaction-required errors", () => {
  assert.equal(needsInteractiveSignIn(new InteractionRequiredAuthError("refresh_token_expired")), true);
  assert.equal(needsInteractiveSignIn(new InteractionRequiredAuthError("login_required")), true);

  // Transient / environmental: retrying is the right answer, not a browser.
  assert.equal(needsInteractiveSignIn(networkFailure()), false);
  assert.equal(needsInteractiveSignIn(new ServerError("server_error", "upstream 500")), false);
  assert.equal(needsInteractiveSignIn(new Error("socket hang up")), false);
  assert.equal(needsInteractiveSignIn(new Error("getaddrinfo EAI_AGAIN login.microsoftonline.com")), false);
});

/* -------------------------------------------------------------------------- */
/* Provider behaviour                                                          */
/* -------------------------------------------------------------------------- */

test("a transient refresh failure never opens a browser", async () => {
  msal.silent = () => Promise.reject(networkFailure());
  const getToken = await getTokenProvider({ environmentUrl: ENV, forceUser: true });

  await assert.rejects(getToken(), (e: Error) => {
    assert.match(e.message, /Couldn't refresh the Dataverse access token for contoso\.crm\.dynamics\.com/);
    // The remedy has to point at the real one. Telling someone to re-register
    // a redirect URI for a blip is how this cost an afternoon.
    assert.match(e.message, /transient/);
    return true;
  });

  assert.deepEqual(msal.interactiveClientIds, [], "a network blip must not escalate to interactive login");
});

test("the next call succeeds once the blip passes, with no re-login", async () => {
  msal.silent = () => Promise.reject(networkFailure());
  const getToken = await getTokenProvider({ environmentUrl: ENV, forceUser: true });
  await assert.rejects(getToken());

  // This is the "I refreshed the browser and it worked again" path — it must
  // resolve from the cached account, without any sign-in.
  msal.silent = () => Promise.resolve({ accessToken: "tok", account: {} });
  assert.equal(await getToken(), "tok");
  assert.deepEqual(msal.interactiveClientIds, []);
});

test("a genuinely expired session still escalates to sign-in", async () => {
  // The expired refresh token is only expired once: after the interactive
  // re-login the provider retries silently, and that retry is where the token
  // actually comes from.
  let calls = 0;
  msal.silent = () =>
    ++calls === 1
      ? Promise.reject(new InteractionRequiredAuthError("refresh_token_expired"))
      : Promise.resolve({ accessToken: "fresh", account: {} });
  msal.interactive = () => Promise.resolve({ accessToken: "fresh", account: { homeAccountId: "h" } });

  const getToken = await getTokenProvider({ environmentUrl: ENV, forceUser: true });
  assert.equal(await getToken(), "fresh");
  assert.deepEqual(msal.interactiveClientIds, [SHARED], "an expired refresh token must prompt");
});

/* -------------------------------------------------------------------------- */
/* silentOnly                                                                  */
/* -------------------------------------------------------------------------- */
//
// The tests above settle when escalation is *warranted*. These settle who is
// allowed to carry it out. A genuinely expired session is the CLI's cue to
// prompt — and the sidecar's cue to say "sign-in required" and stop, because
// its callers are background HTTP requests, and a browser opening over Excel
// in the middle of one is not something the user asked for.

test("silentOnly reports a dead session instead of signing in", async () => {
  msal.silent = () => Promise.reject(new InteractionRequiredAuthError("refresh_token_expired"));
  msal.interactive = () =>
    Promise.resolve({ accessToken: "should-never-happen", account: { homeAccountId: "h" } });

  const getToken = await getTokenProvider({ environmentUrl: ENV, forceUser: true, silentOnly: true });

  await assert.rejects(getToken(), (e: Error) => {
    // The type is the contract — serve.ts turns exactly this into a 401 with
    // needsSignIn, and the pane turns that into its Sign in button.
    assert.ok(isInteractiveSignInRequired(e), `expected a sign-in-required error, got ${e.name}`);
    assert.match(e.message, new RegExp(ENV));
    return true;
  });

  assert.deepEqual(
    msal.interactiveClientIds,
    [],
    "silentOnly must never open a browser, even for a genuinely expired session"
  );
});

test("silentOnly still reports a transient failure as transient", async () => {
  // Both refusals end in "no token", and conflating them would have the pane
  // telling someone to sign in because their VPN reconnected.
  msal.silent = () => Promise.reject(networkFailure());
  const getToken = await getTokenProvider({ environmentUrl: ENV, forceUser: true, silentOnly: true });

  await assert.rejects(getToken(), (e: Error) => {
    assert.equal(isInteractiveSignInRequired(e), false);
    assert.match(e.message, /transient/);
    return true;
  });
});

test("a fallback that cannot receive the redirect is never opened in a browser", async () => {
  // The user-visible bug: the shared client's window timed out, the chain
  // advanced, and a SECOND browser window opened on dvload's own app — which
  // has no http://localhost reply URL, so it could only ever end in
  // AADSTS900971 after a full sign-in. An unconfirmed fallback now costs
  // nothing instead of costing someone their credentials twice.
  stubProbe({ [SHARED]: "usable", [DVLOAD]: "unusable" });
  msal.silent = () => Promise.reject(new InteractionRequiredAuthError("refresh_token_expired"));
  msal.interactive = () => {
    const e = new Error("No response from the browser after 180s.") as Error & { errorCode: string };
    e.errorCode = "interactive_timeout";
    return Promise.reject(e);
  };

  const getToken = await getTokenProvider({ environmentUrl: ENV, forceUser: true });
  await getToken().catch(() => {});

  assert.deepEqual(
    msal.interactiveClientIds,
    [SHARED],
    "only the confirmed client may open a browser window"
  );
});

test("the remembered client id does not pin the sign-in chain", async () => {
  // Both apps can receive the redirect here, so the chain is genuinely two
  // long and the ordering question this test exists for is the live one.
  stubProbe({ [SHARED]: "usable", [DVLOAD]: "usable" });
  msal.silent = () => Promise.reject(new InteractionRequiredAuthError("refresh_token_expired"));
  // What the shared Microsoft client actually does in a browser: the loopback
  // listener hears nothing, because the app has no http://localhost redirect
  // URI, and the attempt dies on dvload's own timeout.
  msal.interactive = (clientId: string) => {
    if (clientId !== SHARED) return Promise.resolve({ accessToken: "t", account: { homeAccountId: "h" } });
    const e = new Error("No response from the browser after 180s.") as Error & { errorCode: string };
    e.errorCode = "interactive_timeout";
    return Promise.reject(e);
  };

  const getToken = await getTokenProvider({ environmentUrl: ENV, forceUser: true });
  await getToken().catch(() => {});

  // The regression: loginDelegated was handed the *stored* client id as an
  // explicit --client-id, so resolveClientIdChain collapsed to [SHARED] and
  // the chain had nowhere to go — every escalation ended in that 180s timeout,
  // permanently. It must still fall through to dvload's own app, which does
  // have a working localhost redirect.
  assert.deepEqual(
    msal.interactiveClientIds,
    [SHARED, DVLOAD],
    "the stored client is tried first, then the chain must continue past it"
  );
});
