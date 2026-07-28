// Local sidecar: serves the dvload web UI over loopback and hands it
// Dataverse access tokens acquired with the CLI's own auth stack.
//
// WHY THIS EXISTS
//
// The task pane cannot borrow Microsoft's pre-consented Dataverse client id
// the way the CLI does. A browser auth-code flow needs an `spa` redirect URI
// on the app registration — both for the redirect itself and for the token
// endpoint's Access-Control-Allow-Origin — and you cannot add one to a
// Microsoft-owned app. Device code hits the same CORS wall, and Entra
// rejects `spa` redirect URIs in non-SPA flows, so there is no configuration
// that makes browser-side auth work without our own registration. See
// docs/AUTH-NOTES.md.
//
// Moving token acquisition into *this* process removes the constraint
// entirely: Node is a native public client, it redirects to
// http://localhost, and no CORS is involved. The UI never talks to Entra at
// all — it asks the sidecar for a token over a same-origin fetch.
//
// The same server backs both front ends:
//   • Excel task pane — manifest SourceLocation points at this port
//   • `dvload gui`    — the same bundle in an ordinary browser
//
// SECURITY POSTURE
//
// This process holds Dataverse tokens and listens on a predictable port, so
// a hostile web page in any browser on this machine is the threat model.
// Four independent controls, all in `guard()`:
//   1. Bind 127.0.0.1 only — never reachable off-box.
//   2. Host header allowlist — blocks DNS rebinding, where an attacker's
//      domain resolves to 127.0.0.1 and defeats the origin checks below.
//   3. No CORS headers are ever sent, so a cross-origin caller cannot read
//      a response even if it manages to issue the request.
//   4. API routes are POST-only, require an `x-dvload-client` header (which
//      forces a preflight that step 3 then fails) and, when an Origin header
//      is present, require it to be one of ours.

import { createServer as createHttpServer, type IncomingMessage, type ServerResponse } from "node:http";
import { createServer as createHttpsServer } from "node:https";
import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";
import { fileURLToPath } from "node:url";
import kleur from "kleur";

import {
  getTokenProvider,
  loginDelegated,
  logoutDelegated,
  detectAuthMode,
  getStoredClientId,
  getSignedInAccount,
  describeClient,
  isSharedMicrosoftClient,
  openInBrowser,
} from "../auth.js";

/**
 * Fixed by design, not configurable in the manifest: an Office add-in's
 * SourceLocation is a literal URL baked into the XML, so the port has to be
 * knowable ahead of time. 44321 is high enough to avoid the usual dev-server
 * collisions (3000, 8080) and is the port Microsoft's own Office samples use.
 */
export const DEFAULT_PORT = 44321;

export interface ServeOptions {
  port?: number;
  /** Serve plain HTTP. Fine for `dvload gui`; Office rejects non-HTTPS panes. */
  http?: boolean;
  /** Open the UI in the system browser once listening. */
  open?: boolean;
  /** Override the directory containing taskpane.html. */
  webRoot?: string;
  cert?: string;
  key?: string;
}

/* -------------------------------------------------------------------------- */
/* Web root resolution                                                         */
/* -------------------------------------------------------------------------- */

const HERE = path.dirname(fileURLToPath(import.meta.url));

/**
 * Find the built add-in bundle. Ordered most-explicit-first so a developer
 * can always override, and so an installed CLI finds its shipped copy before
 * it goes looking for a repo checkout.
 */
async function resolveWebRoot(explicit?: string): Promise<string> {
  const candidates: string[] = [];
  if (explicit) candidates.push(path.resolve(explicit));
  if (process.env.DVLOAD_WEB_ROOT) candidates.push(path.resolve(process.env.DVLOAD_WEB_ROOT));
  // Installed layout: bundle.mjs copies the add-in dist to build/web,
  // alongside build/dvload.cjs.
  candidates.push(path.join(HERE, "web"));
  candidates.push(path.join(HERE, "..", "build", "web"));
  // Repo checkout: walk up looking for packages/addin/dist.
  let dir = HERE;
  for (let i = 0; i < 6; i++) {
    candidates.push(path.join(dir, "packages", "addin", "dist"));
    const parent = path.dirname(dir);
    if (parent === dir) break;
    dir = parent;
  }

  for (const c of candidates) {
    try {
      await fs.access(path.join(c, "taskpane.html"));
      return c;
    } catch {
      // keep looking
    }
  }

  throw new Error(
    "Could not find the built add-in UI (no taskpane.html).\n" +
      "Build it first:\n" +
      "  npm run build --workspace=@dvload/addin\n" +
      "or point at an existing build with --web-root <dir> or DVLOAD_WEB_ROOT."
  );
}

/* -------------------------------------------------------------------------- */
/* TLS                                                                         */
/* -------------------------------------------------------------------------- */

interface TlsMaterial {
  cert: Buffer | string;
  key: Buffer | string;
}

/**
 * Office requires HTTPS for a task pane's SourceLocation, even on localhost,
 * so `serve` needs a certificate that the machine already trusts.
 *
 * office-addin-dev-certs installs a CA into the *CurrentUser* store, which
 * matters here: it means the whole dvload install path stays admin-free,
 * which is the entire point of the sidecar. We read its output files
 * directly when they exist and only fall back to importing the package,
 * because the package is a devDependency of the add-in workspace and won't
 * be present next to an installed CLI.
 */
async function resolveTls(opts: ServeOptions): Promise<TlsMaterial> {
  if (opts.cert && opts.key) {
    return { cert: await fs.readFile(opts.cert), key: await fs.readFile(opts.key) };
  }

  const certDir = path.join(os.homedir(), ".office-addin-dev-certs");
  try {
    const [cert, key] = await Promise.all([
      fs.readFile(path.join(certDir, "localhost.crt")),
      fs.readFile(path.join(certDir, "localhost.key")),
    ]);
    return { cert, key };
  } catch {
    // fall through to the package
  }

  try {
    // Indirect specifier so bundlers leave this as a runtime lookup — the
    // package is optional and usually absent in an installed CLI.
    const spec = "office-addin-dev-certs";
    const mod = (await import(spec)) as {
      getHttpsServerOptions: () => Promise<{ cert: string; key: string }>;
    };
    const o = await mod.getHttpsServerOptions();
    return { cert: o.cert, key: o.key };
  } catch {
    throw new Error(
      "No trusted localhost certificate found.\n" +
        "Office only loads a task pane over HTTPS, so one is required.\n\n" +
        "Install one (no administrator rights needed — it goes in your user store):\n" +
        "  npx office-addin-dev-certs install\n\n" +
        "Then re-run `dvload serve`. To use your own certificate instead, pass\n" +
        "--cert <path> --key <path>. For browser-only use, `dvload gui --http`\n" +
        "skips TLS entirely (http://localhost is still a secure context)."
    );
  }
}

/* -------------------------------------------------------------------------- */
/* Token providers                                                             */
/* -------------------------------------------------------------------------- */

type Thunk = () => Promise<string>;

const providers = new Map<string, Promise<Thunk>>();

function providerFor(environmentUrl: string, forceUser: boolean): Promise<Thunk> {
  const key = `${environmentUrl}|${forceUser ? "user" : "auto"}`;
  let p = providers.get(key);
  if (!p) {
    p = getTokenProvider({ environmentUrl, forceUser });
    providers.set(key, p);
  }
  return p;
}

/**
 * Serialise token acquisition per environment.
 *
 * Without this, the pane's first render — which fires several metadata
 * requests at once — can trigger several concurrent interactive logins on a
 * cold cache, and the user gets a stack of browser windows. MSAL dedupes
 * silent refreshes internally but not the interactive escalation.
 */
const inFlight = new Map<string, Promise<string>>();

function acquireToken(environmentUrl: string, forceUser: boolean): Promise<string> {
  const key = `${environmentUrl}|${forceUser ? "user" : "auto"}`;
  const existing = inFlight.get(key);
  if (existing) return existing;
  const p = (async () => {
    const getToken = await providerFor(environmentUrl, forceUser);
    return getToken();
  })().finally(() => inFlight.delete(key));
  inFlight.set(key, p);
  return p;
}

/** `exp` out of a JWT, in epoch ms, so the UI can cache without re-asking. */
function tokenExpiry(accessToken: string): number | null {
  const seg = accessToken.split(".")[1];
  if (!seg) return null;
  try {
    const json = Buffer.from(seg.replace(/-/g, "+").replace(/_/g, "/"), "base64").toString("utf8");
    const exp = (JSON.parse(json) as { exp?: number }).exp;
    return typeof exp === "number" ? exp * 1000 : null;
  } catch {
    return null;
  }
}

/* -------------------------------------------------------------------------- */
/* Request plumbing                                                            */
/* -------------------------------------------------------------------------- */

const MIME: Record<string, string> = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".mjs": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".map": "application/json; charset=utf-8",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".svg": "image/svg+xml",
  ".ico": "image/x-icon",
  ".woff": "font/woff",
  ".woff2": "font/woff2",
  ".xml": "text/xml; charset=utf-8",
};

function send(res: ServerResponse, status: number, body: string | Buffer, type: string): void {
  res.writeHead(status, {
    "content-type": type,
    "content-length": Buffer.byteLength(body),
    // Never cache: `dvload serve` is also the dev loop, and a stale
    // taskpane.js inside Office's WebView cache is miserable to diagnose.
    "cache-control": "no-store",
    "x-content-type-options": "nosniff",
  });
  res.end(body);
}

function sendJson(res: ServerResponse, status: number, body: unknown): void {
  send(res, status, JSON.stringify(body), "application/json; charset=utf-8");
}

async function readJsonBody(req: IncomingMessage): Promise<Record<string, unknown>> {
  const chunks: Buffer[] = [];
  let size = 0;
  for await (const chunk of req) {
    size += (chunk as Buffer).length;
    if (size > 1_000_000) throw new Error("Request body too large.");
    chunks.push(chunk as Buffer);
  }
  if (chunks.length === 0) return {};
  return JSON.parse(Buffer.concat(chunks).toString("utf8")) as Record<string, unknown>;
}

function requiredString(body: Record<string, unknown>, field: string): string {
  const v = body[field];
  if (typeof v !== "string" || !v.trim()) throw new Error(`Missing "${field}".`);
  return v.trim();
}

/**
 * Everything that decides whether a request is allowed to reach a handler.
 * Returns an error string to reject with, or null to proceed.
 */
function guard(req: IncomingMessage, allowedOrigins: string[], allowedHosts: string[], isApi: boolean): string | null {
  // DNS rebinding: an attacker points evil.example at 127.0.0.1, so the
  // browser treats their origin as same-site with us. The Host header still
  // carries their domain, and that is what catches it.
  const hostHeader = (req.headers.host ?? "").toLowerCase();
  if (!allowedHosts.includes(hostHeader)) {
    return `Unexpected Host header "${hostHeader}". dvload only serves localhost.`;
  }

  if (!isApi) return null;

  if (req.method !== "POST") return "API routes require POST.";

  // A form or <img> cannot set a custom header, so requiring one forces any
  // cross-origin attempt through a preflight — which fails, because we never
  // emit Access-Control-Allow-Origin.
  if (req.headers["x-dvload-client"] !== "1") return "Missing x-dvload-client header.";

  // Same-origin POSTs do send Origin; a missing one means a non-browser
  // caller (curl, a test), which the header check above already gated.
  const origin = req.headers.origin;
  if (typeof origin === "string" && !allowedOrigins.includes(origin.toLowerCase())) {
    return `Origin "${origin}" is not allowed.`;
  }

  return null;
}

async function serveStatic(res: ServerResponse, webRoot: string, urlPath: string): Promise<void> {
  const rel = urlPath === "/" ? "taskpane.html" : decodeURIComponent(urlPath).replace(/^\/+/, "");
  const full = path.resolve(webRoot, rel);

  // Path traversal: resolve first, then confirm the result is still inside
  // the web root. Checking the raw string for ".." misses encoded forms.
  if (full !== webRoot && !full.startsWith(webRoot + path.sep)) {
    send(res, 403, "Forbidden", "text/plain; charset=utf-8");
    return;
  }

  try {
    const body = await fs.readFile(full);
    send(res, 200, body, MIME[path.extname(full).toLowerCase()] ?? "application/octet-stream");
  } catch {
    send(res, 404, `Not found: ${rel}`, "text/plain; charset=utf-8");
  }
}

/* -------------------------------------------------------------------------- */
/* API                                                                         */
/* -------------------------------------------------------------------------- */

/**
 * The UI is always delegated. Never app-only.
 *
 * `getTokenProvider` prefers app-only whenever credentials are stored for an
 * environment, which is right for `dvload run` — that is the unattended path.
 * It is wrong here. Someone clicking "Run import" in a task pane expects the
 * records to be created as themselves, and the old MSAL-popup add-in always
 * did exactly that.
 *
 * Without this, configuring a nightly schedule would silently change who the
 * *interactive* UI writes as, because `dvload app-login` stores credentials
 * per environment and `schedule` actively recommends running it. The two
 * features would quietly interfere. Ownership, audit fields and any
 * ownership-based security role would all shift to the Application User with
 * nothing on screen to say so.
 *
 * `forceUser: true` is the same switch as the CLI's `--user` flag.
 */
const UI_TOKEN_OPTS = { forceUser: true } as const;

async function describeAuth(environmentUrl: string): Promise<Record<string, unknown>> {
  const clientId = await getStoredClientId(environmentUrl);
  // Always the delegated account, regardless of what detectAuthMode says:
  // delegated is the only mode the UI uses.
  const account = await getSignedInAccount(environmentUrl);
  return {
    mode: "delegated",
    account: account ? { username: account.username } : null,
    clientId,
    // Reported so the pane can say so, since it changes nothing about the UI
    // but explains why `dvload run` on the same environment may write as a
    // different identity.
    appOnlyConfigured: (await detectAuthMode(environmentUrl)) === "appOnly",
    // Surfaced in the pane so the attribution trade-off is visible where the
    // user actually is: the consent screen and the tenant sign-in log will
    // say "Microsoft Dynamics CRM", not dvload.
    clientLabel: clientId ? describeClient(clientId) : null,
    sharedClient: clientId ? isSharedMicrosoftClient(clientId) : false,
  };
}

async function handleApi(route: string, body: Record<string, unknown>): Promise<unknown> {
  switch (route) {
    case "/api/account":
      return describeAuth(requiredString(body, "environmentUrl"));

    case "/api/signin": {
      const environmentUrl = requiredString(body, "environmentUrl");
      // Clear any memoised provider so the next token request picks up the
      // account we are about to create rather than a stale failed thunk.
      providers.clear();
      await loginDelegated({ environmentUrl });
      return describeAuth(environmentUrl);
    }

    case "/api/signout": {
      const environmentUrl = typeof body.environmentUrl === "string" ? body.environmentUrl : undefined;
      await logoutDelegated(environmentUrl);
      providers.clear();
      return { ok: true };
    }

    case "/api/token": {
      const environmentUrl = requiredString(body, "environmentUrl");
      // Not configurable from the request: a page must not be able to ask for
      // the Application User's token. See UI_TOKEN_OPTS.
      const accessToken = await acquireToken(environmentUrl, UI_TOKEN_OPTS.forceUser);
      return { accessToken, expiresOn: tokenExpiry(accessToken) };
    }

    default:
      return null;
  }
}

/* -------------------------------------------------------------------------- */
/* Entry points                                                                */
/* -------------------------------------------------------------------------- */

export interface RunningServer {
  server: import("node:http").Server;
  port: number;
  url: string;
  webRoot: string;
  close(): Promise<void>;
}

/**
 * Build and start the server. Split out from `serveCommand` so tests can
 * drive a real listener — the security rules in `guard()` are only worth
 * anything if they're exercised against actual HTTP, not called directly.
 */
export async function startServer(opts: ServeOptions): Promise<RunningServer> {
  const requestedPort = opts.port ?? DEFAULT_PORT;
  const webRoot = path.resolve(await resolveWebRoot(opts.webRoot));
  const scheme = opts.http ? "http" : "https";

  // Filled in after listen(), because port 0 (used by tests) isn't known
  // until the OS assigns one — and the guard lists are port-specific.
  let allowedHosts: string[] = [];
  let allowedOrigins: string[] = [];

  const handler = async (req: IncomingMessage, res: ServerResponse): Promise<void> => {
    const urlPath = (req.url ?? "/").split("?")[0];
    const isApi = urlPath.startsWith("/api/");
    // The liveness probe is a GET and therefore cannot satisfy the POST-only
    // rule the other API routes are held to. Exempting it is safe because it
    // returns nothing but "dvload is running" — no environment, no account,
    // no token — and it still passes the Host check below. It has to be a
    // GET so the pane can call it before the user has chosen an environment.
    const isProbe = urlPath === "/api/status";

    const denied = guard(req, allowedOrigins, allowedHosts, isApi && !isProbe);
    if (denied) {
      if (isApi) sendJson(res, 403, { error: denied });
      else send(res, 403, denied, "text/plain; charset=utf-8");
      return;
    }

    if (isProbe) {
      sendJson(res, 200, { ok: true, product: "dvload", version: "0.1.0" });
      return;
    }

    if (isApi) {
      try {
        const result = await handleApi(urlPath, await readJsonBody(req));
        if (result === null) sendJson(res, 404, { error: `Unknown route ${urlPath}` });
        else sendJson(res, 200, result);
      } catch (e) {
        const message = e instanceof Error ? e.message : String(e);
        console.error(kleur.red(`  ${urlPath} failed: ${message}`));
        sendJson(res, 500, { error: message });
      }
      return;
    }

    await serveStatic(res, webRoot, urlPath);
  };

  const server = opts.http
    ? createHttpServer((req, res) => void handler(req, res))
    : createHttpsServer(await resolveTls(opts), (req, res) => void handler(req, res));

  await new Promise<void>((resolve, reject) => {
    server.once("error", (err: NodeJS.ErrnoException) => {
      if (err.code === "EADDRINUSE") {
        reject(
          new Error(
            `Port ${requestedPort} is already in use.\n` +
              `Another \`dvload serve\` may already be running — if so, you're set.\n` +
              `Otherwise free the port, or use --port <n> (and update the add-in ` +
              `manifest's SourceLocation to match).`
          )
        );
      } else {
        reject(err);
      }
    });
    // 127.0.0.1, not 0.0.0.0: this server hands out access tokens and has no
    // business being reachable from anywhere but this machine.
    server.listen(requestedPort, "127.0.0.1", resolve);
  });

  const address = server.address();
  const port = typeof address === "object" && address ? address.port : requestedPort;
  allowedHosts = [`localhost:${port}`, `127.0.0.1:${port}`];
  allowedOrigins = allowedHosts.map((h) => `${scheme}://${h}`);

  return {
    server,
    port,
    webRoot,
    url: `${scheme}://localhost:${port}/taskpane.html`,
    close: () => new Promise<void>((resolve) => server.close(() => resolve())),
  };
}

export async function serveCommand(opts: ServeOptions): Promise<void> {
  const { server, port, webRoot, url } = await startServer(opts);
  const scheme = opts.http ? "http" : "https";

  console.log(kleur.green(`dvload serve — listening on ${scheme}://localhost:${port}`));
  console.log(kleur.gray(`  UI:       ${webRoot}`));
  console.log(kleur.gray(`  Excel:    open the Dataverse Load pane (manifest points here)`));
  console.log(kleur.gray(`  Browser:  ${url}`));
  console.log(kleur.gray("  Sign-in uses this process, so no Entra app registration is needed."));
  console.log(kleur.gray("  Ctrl+C to stop."));

  if (opts.open) await openInBrowser(url);

  await new Promise<void>((resolve) => {
    const stop = (): void => {
      server.close(() => resolve());
    };
    process.once("SIGINT", stop);
    process.once("SIGTERM", stop);
  });
}

/** `dvload gui` — same server, opens a browser at it. */
export async function guiCommand(opts: ServeOptions): Promise<void> {
  await serveCommand({ ...opts, open: true });
}
