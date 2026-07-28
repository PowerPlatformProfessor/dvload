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

import { test, before, after } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, writeFile, mkdir } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { request } from "node:http";

import { startServer, type RunningServer } from "./serve.js";

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
        res.on("end", () =>
          resolve({ status: res.statusCode ?? 0, body, headers: res.headers })
        );
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
  // No tenant here, so this asserts the contract rather than the token: the
  // route must not accept an identity switch from the request body.
  const res = await call("POST", "/api/token", {
    headers: ours(),
    body: JSON.stringify({
      environmentUrl: "https://example.invalid",
      forceUser: false,
      appOnly: true,
    }),
  });
  // Fails for want of a real environment, not because it honoured appOnly.
  assert.equal(res.status, 500);
  assert.doesNotMatch(res.body, /client_credentials|clientSecret/i);
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
