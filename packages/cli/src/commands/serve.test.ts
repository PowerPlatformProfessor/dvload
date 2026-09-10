// Tests for the local sidecar.
//
// The emphasis is deliberately on the request guard rather than on the token
// routes: this process holds Dataverse access tokens and listens on a
// predictable loopback port, so the thing most worth pinning down is exactly
// which requests it refuses. Those rules are only meaningful when exercised
// over real HTTP — a hostile page's request differs from ours in headers the
// browser sets, not in anything a direct function call would model.
//
// The token routes themselves need a tenant and a signed-in user, so they
// belong in TEST-PROTOCOL.md, not here.

import { test, beforeAll as before, afterAll as after, vi } from "vitest";
import assert from "node:assert/strict";
import { mkdtemp, writeFile, mkdir } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { request } from "node:http";

// Stub the auth layer.
//
// Without this, POST /api/token reaches makeDelegatedProvider with an empty
// token cache, which escalates to an INTERACTIVE login: on Windows it opens a
// browser window and blocks for DVLOAD_AUTH_TIMEOUT_MS (3 minutes by default).
// A unit test must never do that — it hangs CI, pops a browser on a
// contributor's machine, and makes the result depend on the platform's
// default login flow.
//
// Stubbing here also makes the identity assertion below *stronger*: instead of
// inferring "it didn't go app-only" from the text of a network failure, the
// test can look at the arguments the route actually passed.
const getTokenProvider = vi.hoisted(() =>
  vi.fn(
    // Typed to match the real export so a signature change breaks this test
    // rather than silently making the assertions below meaningless.
    (_opts: {
      environmentUrl: string;
      forceUser?: boolean;
      silentOnly?: boolean;
    }): Promise<() => Promise<string>> =>
      Promise.resolve(() => Promise.reject(new Error("no cached token in tests")))
  )
);

/**
 * Stubbed for the same reason: the real one reads the secure store and then
 * rebuilds an MSAL app per environment to recover each username. Here it is
 * the route's own shaping that's under test — which fields are exposed, and
 * what happens to a session whose username can no longer be read.
 */
const listDelegatedSessions = vi.hoisted(() =>
  vi.fn(
    (): Promise<Array<{ host: string; environmentUrl: string; username: string | null }>> =>
      Promise.resolve([])
  )
);

/**
 * The sign-in trio, stubbed so /api/signin is testable at all: the real
 * loginDelegated opens a browser, and the real logoutDelegated deletes files
 * under ~/.dvload — neither belongs in a unit test. Defaults are the empty
 * cache the other routes' tests assume.
 */
const loginDelegated = vi.hoisted(() =>
  vi.fn(
    (_opts?: { environmentUrl?: string; onAuthorizeUrl?: (url: string) => void }): Promise<unknown> =>
      Promise.reject(new Error("loginDelegated not stubbed for this test"))
  )
);
const logoutDelegated = vi.hoisted(() => vi.fn((_env?: string): Promise<void> => Promise.resolve()));
const getSignedInAccount = vi.hoisted(() =>
  vi.fn((_env: string): Promise<{ username: string } | null> => Promise.resolve(null))
);

vi.mock("../auth.js", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../auth.js")>();
  return {
    ...actual,
    getTokenProvider,
    listDelegatedSessions,
    loginDelegated,
    logoutDelegated,
    getSignedInAccount,
  };
});

import { startServer, resolveWebSource, checkUi, type RunningServer } from "./serve.js";
import { InteractiveSignInRequiredError } from "../auth.js";

let server: RunningServer;
let webRoot: string;

before(async () => {
  webRoot = await mkdtemp(path.join(tmpdir(), "dvload-web-"));
  await writeFile(path.join(webRoot, "taskpane.html"), "<html>pane</html>");
  await mkdir(path.join(webRoot, "assets"), { recursive: true });
  await writeFile(path.join(webRoot, "assets", "icon-16.png"), "png-bytes");
  // A file the server must never serve, one level above the web root.
  await writeFile(path.join(webRoot, "..", "dvload-secret.txt"), "secret");

  server = await startServer({ port: 0, http: true, webRoot });
});

after(async () => {
  await server?.close();
});

interface Res {
  status: number;
  body: string;
  headers: Record<string, string | string[] | undefined>;
}

function call(
  method: string,
  urlPath: string,
  opts: { headers?: Record<string, string>; body?: string; host?: string } = {}
): Promise<Res> {
  return new Promise((resolve, reject) => {
    const req = request(
      {
        host: "127.0.0.1",
        port: server.port,
        method,
        path: urlPath,
        headers: {
          host: opts.host ?? `127.0.0.1:${server.port}`,
          ...(opts.headers ?? {}),
        },
      },
      (res) => {
        let body = "";
        res.setEncoding("utf8");
        res.on("data", (c) => (body += c));
        res.on("end", () => resolve({ status: res.statusCode ?? 0, body, headers: res.headers }));
      }
    );
    req.on("error", reject);
    if (opts.body) req.write(opts.body);
    req.end();
  });
}

/** The headers a legitimate request from our own page carries. */
function ours(): Record<string, string> {
  return {
    "content-type": "application/json",
    "x-dvload-client": "1",
    origin: `http://127.0.0.1:${server.port}`,
  };
}

/* -------------------------------------------------------------------------- */
/* Static serving                                                              */
/* -------------------------------------------------------------------------- */

test("serves the task pane at the root", async () => {
  const res = await call("GET", "/");
  assert.equal(res.status, 200);
  assert.match(res.body, /pane/);
});

test("serves nested assets with a sensible content type", async () => {
  const res = await call("GET", "/assets/icon-16.png");
  assert.equal(res.status, 200);
  assert.equal(res.headers["content-type"], "image/png");
});

test("unknown paths 404 rather than falling back to the pane", async () => {
  // A SPA-style catch-all would make a mistyped API route look like a
  // working page, which is a miserable thing to debug.
  const res = await call("GET", "/nope.js");
  assert.equal(res.status, 404);
});

test("never serves files outside the web root", async () => {
  for (const attempt of [
    "/../dvload-secret.txt",
    "/..%2Fdvload-secret.txt",
    "/assets/../../dvload-secret.txt",
    "/%2e%2e/dvload-secret.txt",
  ]) {
    const res = await call("GET", attempt);
    assert.ok(res.status === 403 || res.status === 404, `${attempt} → ${res.status}`);
    assert.doesNotMatch(res.body, /secret/, `${attempt} leaked the file`);
  }
});

test("a malformed percent-escape is a client error, not a crash", async () => {
  // decodeURIComponent throws on a lone '%'. Before the web source was
  // factored out this went through fs.readFile's try/catch and surfaced as a
  // 404; it is a bad request, and either way it must not take the server down.
  const res = await call("GET", "/%");
  assert.equal(res.status, 400);
});

/* -------------------------------------------------------------------------- */
/* Web source resolution                                                       */
/* -------------------------------------------------------------------------- */
//
// The exe embeds the UI as SEA assets and the npm install ships it as a
// directory, so resolution has two backends. Only the disk one is reachable
// from a test process — a SEA binary cannot be built in-process — so the
// embedded path is covered by the release smoke test (`dvload gui --check-ui`
// from a directory holding nothing but the exe). What is worth pinning here
// is that the precedence and the validation still behave.

test("an explicit web root wins over everything else", async () => {
  const web = await resolveWebSource(webRoot);
  assert.equal(web.label, path.resolve(webRoot));
  assert.match((await web.read("taskpane.html"))?.toString() ?? "", /pane/);
});

test("DVLOAD_WEB_ROOT is honoured when no explicit root is given", async () => {
  vi.stubEnv("DVLOAD_WEB_ROOT", webRoot);
  try {
    const web = await resolveWebSource();
    assert.equal(web.label, path.resolve(webRoot));
  } finally {
    vi.unstubAllEnvs();
  }
});

test("a web source refuses to read outside itself", async () => {
  const web = await resolveWebSource(webRoot);
  for (const attempt of ["../dvload-secret.txt", "assets/../../dvload-secret.txt"]) {
    assert.equal(await web.read(attempt), null, attempt);
  }
});

test("--check-ui accepts a complete UI and names the source", async () => {
  const complete = await mkdtemp(path.join(tmpdir(), "dvload-web-ok-"));
  for (const f of ["taskpane.html", "taskpane.js", "commands.html"]) {
    await writeFile(path.join(complete, f), "x");
  }

  // The release smoke test reads this line to tell "served from the embedded
  // assets" apart from "found a directory lying around next to the exe",
  // which is the distinction the whole check exists to make.
  const lines: string[] = [];
  const log = vi.spyOn(console, "log").mockImplementation((...args) => void lines.push(args.join(" ")));
  try {
    await assert.doesNotReject(() => checkUi(complete));
  } finally {
    log.mockRestore();
  }
  assert.ok(
    lines.some((l) => l.includes(path.resolve(complete))),
    lines.join("\n")
  );
});

test("--check-ui rejects a UI that is present but incomplete", async () => {
  // The failure this guards against: a bundle that stages taskpane.html but
  // drops the script, which starts fine and only breaks when a user opens
  // the pane. Naming the missing files is the whole point.
  const partial = await mkdtemp(path.join(tmpdir(), "dvload-web-partial-"));
  await writeFile(path.join(partial, "taskpane.html"), "<html>pane</html>");

  await assert.rejects(
    () => checkUi(partial),
    (err: Error) => {
      assert.match(err.message, /taskpane\.js/);
      assert.match(err.message, /commands\.html/);
      assert.doesNotMatch(err.message, /taskpane\.html,/);
      return true;
    }
  );
});

// Deliberately not tested here: the "no UI anywhere" throw. Resolution falls
// back to walking up for packages/addin/dist, which exists in a checkout
// whenever the add-in has been built — so the assertion would pass or fail
// depending on whether someone had run a build, which is worse than no test.

/* -------------------------------------------------------------------------- */
/* Guard                                                                       */
/* -------------------------------------------------------------------------- */

test("status probe answers a same-origin GET", async () => {
  const res = await call("GET", "/api/status", { headers: { "x-dvload-client": "1" } });
  assert.equal(res.status, 200);
  assert.equal(JSON.parse(res.body).ok, true);
});

test("rejects a foreign Host header (DNS rebinding)", async () => {
  // The attack: attacker.example resolves to 127.0.0.1, so the browser
  // considers their page same-origin with us and the Origin check passes.
  // The Host header is what still carries their domain.
  const res = await call("GET", "/api/status", {
    host: "attacker.example",
    headers: { "x-dvload-client": "1" },
  });
  assert.equal(res.status, 403);
  assert.match(res.body, /Host/);
});

test("rejects API calls without the client header", async () => {
  // Without this, a plain <form> POST from any page would reach the handler:
  // forms can't set custom headers, which is precisely why one is required.
  const res = await call("POST", "/api/token", {
    headers: { "content-type": "application/json", origin: `http://127.0.0.1:${server.port}` },
    body: JSON.stringify({ environmentUrl: "https://x.crm.dynamics.com" }),
  });
  assert.equal(res.status, 403);
});

test("rejects API calls from another origin", async () => {
  const res = await call("POST", "/api/token", {
    headers: { ...ours(), origin: "https://evil.example" },
    body: JSON.stringify({ environmentUrl: "https://x.crm.dynamics.com" }),
  });
  assert.equal(res.status, 403);
  assert.match(res.body, /Origin/);
});

test("rejects GET on API routes that mutate or return secrets", async () => {
  // GET is reachable from an <img> or a redirect; POST is not.
  const res = await call("GET", "/api/token", { headers: ours() });
  assert.equal(res.status, 403);
});

test("never sends CORS headers", async () => {
  // The last line of defence: even if a request slipped past the checks
  // above, the browser must not let the caller read the response.
  for (const [method, p] of [
    ["GET", "/api/status"],
    ["GET", "/"],
  ] as const) {
    const res = await call(method, p, { headers: { "x-dvload-client": "1" } });
    assert.equal(res.headers["access-control-allow-origin"], undefined);
  }
});

test("accepts a same-origin POST and reports unknown routes as 404", async () => {
  // Proves the guard lets our own requests through — otherwise every test
  // above would pass with a server that simply rejects everything.
  const res = await call("POST", "/api/not-a-route", {
    headers: ours(),
    body: "{}",
  });
  assert.equal(res.status, 404);
});

test("a malformed body is a client error, not a crash", async () => {
  const res = await call("POST", "/api/token", { headers: ours(), body: "{not json" });
  assert.equal(res.status, 500);
  assert.ok(res.body.length > 0);
});

test("missing environmentUrl is reported by name", async () => {
  const res = await call("POST", "/api/token", { headers: ours(), body: "{}" });
  assert.equal(res.status, 500);
  assert.match(res.body, /environmentUrl/);
});

/* -------------------------------------------------------------------------- */
/* Identity                                                                    */
/* -------------------------------------------------------------------------- */

test("the UI is delegated-only and a request cannot ask for app-only", async () => {
  // getTokenProvider prefers app-only whenever credentials are stored, which
  // is correct for `dvload run` and wrong for a task pane: someone clicking
  // "Run import" expects to write as themselves. `dvload app-login` is
  // recommended by `dvload schedule`, so without this the two features would
  // quietly interfere — setting up a nightly run would change who the
  // interactive UI writes as, with nothing on screen to say so.
  //
  getTokenProvider.mockClear();

  const res = await call("POST", "/api/token", {
    headers: ours(),
    body: JSON.stringify({
      environmentUrl: "https://example.invalid",
      forceUser: false,
      appOnly: true,
    }),
  });

  // The request asked for app-only in three different ways. The route must
  // have ignored all of them and asked for a delegated token regardless.
  assert.equal(getTokenProvider.mock.calls.length, 1);
  assert.deepEqual(getTokenProvider.mock.calls[0][0], {
    environmentUrl: "https://example.invalid",
    forceUser: true,
    // Non-negotiable for every provider this server builds — see the
    // sign-in tests below for what it buys.
    silentOnly: true,
  });

  // And the failure surfaces as a server error, not as a token.
  assert.equal(res.status, 500);
  assert.doesNotMatch(res.body, /client_credentials|clientSecret/i);
});

/* -------------------------------------------------------------------------- */
/* Sign-in required                                                            */
/* -------------------------------------------------------------------------- */
//
// The behaviour these pin down is a *refusal to act*. The sidecar can sign the
// user in — it owns the MSAL cache and can launch a browser — and for requests
// like these it must not, because they are background work the pane issued on
// its own. Escalating means a sign-in window appearing over Excel and an HTTP
// response blocked until somebody notices it. The routes answer 401 instead,
// and the pane's Sign in button stays the only thing that opens a browser.

test("an expired session on /api/token is a 401 the pane can act on", async () => {
  getTokenProvider.mockResolvedValueOnce(() =>
    Promise.reject(new InteractiveSignInRequiredError("https://expired.invalid"))
  );

  const res = await call("POST", "/api/token", {
    headers: ours(),
    body: JSON.stringify({ environmentUrl: "https://expired.invalid" }),
  });

  assert.equal(res.status, 401);
  const body = JSON.parse(res.body) as { error: string; needsSignIn?: boolean };
  // The machine-readable half. Without it the pane can only pattern-match on
  // message text to tell "sign in" apart from "something broke".
  assert.equal(body.needsSignIn, true);
  // Sign-in is per environment, so which one is the useful part of the message.
  assert.match(body.error, /expired\.invalid/);
  // And no token leaked out on the way.
  assert.doesNotMatch(res.body, /accessToken/);
});

test("an expired session on a data route is a 401, not a 500", async () => {
  // /api/dataflows reaches auth through DataverseClient rather than directly,
  // so it is worth proving the classification survives that trip.
  getTokenProvider.mockResolvedValueOnce(() =>
    Promise.reject(new InteractiveSignInRequiredError("https://stale.invalid"))
  );

  const res = await call("POST", "/api/dataflows", {
    headers: ours(),
    body: JSON.stringify({ environmentUrl: "https://stale.invalid" }),
  });

  assert.equal(res.status, 401);
  assert.equal((JSON.parse(res.body) as { needsSignIn?: boolean }).needsSignIn, true);
});

test("an ordinary token failure stays a 500 with no sign-in flag", async () => {
  // The distinction has to cut both ways: a transient failure that the user
  // cannot fix by signing in must not tell them to sign in.
  getTokenProvider.mockResolvedValueOnce(() =>
    Promise.reject(new Error("getaddrinfo ENOTFOUND login.microsoftonline.com"))
  );

  const res = await call("POST", "/api/token", {
    headers: ours(),
    body: JSON.stringify({ environmentUrl: "https://offline.invalid" }),
  });

  assert.equal(res.status, 500);
  assert.equal((JSON.parse(res.body) as { needsSignIn?: boolean }).needsSignIn, undefined);
});

test("account status reports delegated regardless of stored app-only creds", async () => {
  const res = await call("POST", "/api/account", {
    headers: ours(),
    body: JSON.stringify({ environmentUrl: "https://example.invalid" }),
  });
  assert.equal(res.status, 200);
  const status = JSON.parse(res.body);
  assert.equal(status.mode, "delegated");
  // Present so the pane can explain why `dvload run` may write as someone else.
  assert.ok("appOnlyConfigured" in status);
});

test("/api/accounts lists every signed-in identity, without an environment", async () => {
  listDelegatedSessions.mockResolvedValueOnce([
    {
      host: "contoso.crm.dynamics.com",
      environmentUrl: "https://contoso.crm.dynamics.com",
      username: "dan@contoso.com",
    },
    {
      host: "contoso-dev.crm.dynamics.com",
      environmentUrl: "https://contoso-dev.crm.dynamics.com",
      username: "dan@contoso.com",
    },
    // Store entry whose MSAL cache no longer holds the account: there is
    // nothing to show for it, and a nameless row in the pane's account picker
    // would be a choice that does nothing.
    { host: "gone.crm.dynamics.com", environmentUrl: "https://gone.crm.dynamics.com", username: null },
  ]);

  // No environmentUrl in the body: the point of this route is that the pane
  // can ask "who am I?" before it has one.
  const res = await call("POST", "/api/accounts", { headers: ours(), body: "{}" });
  assert.equal(res.status, 200);

  const { accounts } = JSON.parse(res.body) as {
    accounts: Array<{ username: string; environmentUrl: string; host: string }>;
  };
  assert.deepEqual(
    accounts.map((a) => `${a.username}@${a.host}`),
    ["dan@contoso.com@contoso.crm.dynamics.com", "dan@contoso.com@contoso-dev.crm.dynamics.com"]
  );
  // Tokens are never part of this answer, only who and where.
  assert.doesNotMatch(res.body, /token|secret/i);
});

test("signing in as a different user signs the previous user out everywhere", async () => {
  // One identity at a time: the pane's account picker must never offer two
  // live users, because which of them an import runs as would then depend on
  // the environment selected. Sessions belonging to anyone but the user who
  // just signed in are removed; the new user's other sessions survive, and a
  // store entry whose username is unreadable is left for its own cleanup.
  loginDelegated.mockResolvedValueOnce({ username: "new@contoso.com" });
  getSignedInAccount.mockImplementation((env) =>
    Promise.resolve(env === "https://new.crm.dynamics.com" ? { username: "new@contoso.com" } : null)
  );
  listDelegatedSessions.mockResolvedValueOnce([
    { host: "old.crm.dynamics.com", environmentUrl: "https://old.crm.dynamics.com", username: "Old@contoso.com" },
    { host: "old2.crm.dynamics.com", environmentUrl: "https://old2.crm.dynamics.com", username: "Old@contoso.com" },
    { host: "new.crm.dynamics.com", environmentUrl: "https://new.crm.dynamics.com", username: "NEW@contoso.com" },
    { host: "gone.crm.dynamics.com", environmentUrl: "https://gone.crm.dynamics.com", username: null },
  ]);

  try {
    const res = await call("POST", "/api/signin", {
      headers: ours(),
      body: JSON.stringify({ environmentUrl: "https://new.crm.dynamics.com" }),
    });
    assert.equal(res.status, 200);
    // The response names the account the pane is now acting as.
    assert.equal((JSON.parse(res.body) as { account: { username: string } }).account.username, "new@contoso.com");
    // Only the other user's sessions were signed out — comparison is
    // case-insensitive, so NEW@ was recognised as the same person as new@.
    assert.deepEqual(
      logoutDelegated.mock.calls.map(([env]) => env),
      ["https://old.crm.dynamics.com", "https://old2.crm.dynamics.com"]
    );
  } finally {
    getSignedInAccount.mockImplementation(() => Promise.resolve(null));
  }
});

test("/api/signin-url serves the pending authorize URL only while sign-in is in flight", async () => {
  // The pane polls this while its /api/signin request blocks, so a browser
  // that failed to open still yields a clickable link. The URL must appear
  // once the interactive flow produces it, be scoped to the environment that
  // asked, and vanish the moment the sign-in settles — after that its
  // loopback listener is gone and a click could never complete.
  const env = "https://pending.crm.dynamics.com";
  const authorizeUrl = "https://login.microsoftonline.com/organizations/oauth2/v2.0/authorize?test=1";
  let finishSignIn!: () => void;
  loginDelegated.mockImplementationOnce((opts) => {
    opts?.onAuthorizeUrl?.(authorizeUrl);
    return new Promise((resolve) => {
      finishSignIn = () => resolve({ username: "u@contoso.com" });
    });
  });

  const askUrl = async (forEnv: string): Promise<string | null> => {
    const res = await call("POST", "/api/signin-url", {
      headers: ours(),
      body: JSON.stringify({ environmentUrl: forEnv }),
    });
    assert.equal(res.status, 200);
    return (JSON.parse(res.body) as { url: string | null }).url;
  };

  assert.equal(await askUrl(env), null, "nothing pending before the sign-in starts");

  const signin = call("POST", "/api/signin", {
    headers: ours(),
    body: JSON.stringify({ environmentUrl: env, force: true }),
  });
  // The mock publishes the URL synchronously inside loginDelegated, so one
  // poll after the request is accepted must see it.
  await vi.waitFor(async () => {
    assert.equal(await askUrl(env), authorizeUrl);
  });
  // Scoped: another environment's poll gets nothing.
  assert.equal(await askUrl("https://other.crm.dynamics.com"), null);

  finishSignIn();
  const res = await signin;
  assert.equal(res.status, 200);
  assert.equal(await askUrl(env), null, "the URL is withdrawn once the sign-in settles");
});
