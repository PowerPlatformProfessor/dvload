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
//
// Deliberately NOT defended: another process running as the same OS user. It
// can call the API directly (curl with the header) and receive tokens — but
// it could equally decrypt the DPAPI secret store, which is CurrentUser-
// scoped. Same-user local processes are inside the trust boundary; SECURITY.md
// rules them out of scope for the same reason.

import { createServer as createHttpServer, type IncomingMessage, type ServerResponse } from "node:http";
import { createServer as createHttpsServer } from "node:https";
import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";
import kleur from "kleur";

import {
  DataverseClient,
  buildWorkbookWithQueries,
  createMetadataResolver,
  getDataflow,
  listDataflows,
  mappingsFromDataflow,
  validateMapping,
} from "@dvload/core";

import {
  getTokenProvider,
  loginDelegated,
  logoutDelegated,
  detectAuthMode,
  getStoredClientId,
  getSignedInAccount,
  listDelegatedSessions,
  describeClient,
  isSharedMicrosoftClient,
  isInteractiveSignInRequired,
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
  /** Resolve and validate the UI, print where it came from, then exit. */
  checkUi?: boolean;
  cert?: string;
  key?: string;
}

/* -------------------------------------------------------------------------- */
/* Web root resolution                                                         */
/* -------------------------------------------------------------------------- */

const HERE = path.dirname(fileURLToPath(import.meta.url));

/**
 * Where the UI's files come from.
 *
 * There are two genuinely different backing stores — a directory on disk and
 * assets embedded in a single-file executable — and the request handler
 * should not care which it got. `label` is for logs only; never join paths
 * onto it, because in the SEA case it does not name a real directory.
 */
export interface WebSource {
  label: string;
  /** Read a file relative to the web root. `null` means "not found". */
  read(rel: string): Promise<Buffer | null>;
}

/** Reject traversal before it reaches any store. Encoded `..` is already
 *  decoded by the caller, so a segment check is sufficient and works for the
 *  SEA case too, where there is no filesystem to resolve against. */
function safeSegments(rel: string): string[] | null {
  const parts = rel.split(/[/\\]+/).filter((p) => p.length > 0 && p !== ".");
  if (parts.some((p) => p === "..")) return null;
  return parts;
}

function diskWebSource(dir: string): WebSource {
  const root = path.resolve(dir);
  return {
    label: root,
    async read(rel) {
      const parts = safeSegments(rel);
      if (!parts) return null;
      const full = path.resolve(root, ...parts);
      if (full !== root && !full.startsWith(root + path.sep)) return null;
      try {
        // path.resolve is lexical, so a symlink inside the root pointing out
        // of it passes the check above. Resolve the real path (of both file
        // and root — the root itself may be reached via a link) and re-check
        // containment before reading.
        const [real, realRoot] = await Promise.all([fs.realpath(full), fs.realpath(root)]);
        if (real !== realRoot && !real.startsWith(realRoot + path.sep)) return null;
        return await fs.readFile(real);
      } catch {
        return null;
      }
    },
  };
}

/**
 * Node's SEA API, or null when unavailable.
 *
 * `node:sea` landed in 20.12, and the CLI supports Node 20.0, so a bare
 * import would break the floor it advertises. It also must not be a static
 * import: esbuild would resolve it at bundle time and the require would then
 * run on every startup, including on Node versions that lack the module.
 */
function seaApi(): { isSea(): boolean; getRawAsset(key: string): ArrayBuffer } | null {
  try {
    const req = createRequire(import.meta.url);
    const sea = req("node:sea") as { isSea?: () => boolean; getRawAsset?: (k: string) => ArrayBuffer };
    if (typeof sea.isSea !== "function" || typeof sea.getRawAsset !== "function") return null;
    if (!sea.isSea()) return null;
    return sea as { isSea(): boolean; getRawAsset(key: string): ArrayBuffer };
  } catch {
    return null;
  }
}

/**
 * The UI as embedded in `dvload.exe`.
 *
 * `scripts/bundle.mjs` writes every file under `build/web` into the SEA
 * config as an asset keyed `web/<posix-relative-path>`. Without this the exe
 * would ship the CLI alone and `serve`/`gui` would die on a missing
 * taskpane.html — the single-file release exists precisely so users don't
 * have to assemble a directory next to it.
 */
function seaWebSource(sea: { getRawAsset(key: string): ArrayBuffer }): WebSource {
  return {
    label: "embedded in this executable",
    read(rel) {
      const parts = safeSegments(rel);
      if (!parts) return Promise.resolve(null);
      try {
        // getRawAsset throws when the key is absent; there is no has().
        return Promise.resolve(Buffer.from(sea.getRawAsset(`web/${parts.join("/")}`)));
      } catch {
        return Promise.resolve(null);
      }
    },
  };
}

/**
 * Find the built add-in bundle.
 *
 * Order matters, and it is not the obvious one. A repo checkout's live
 * `packages/addin/dist` is preferred over the bundled `build/web` copy,
 * because otherwise running `npm run bundle` once would pin the UI forever:
 * every later `npm run build --workspace=@dvload/addin` would rebuild a
 * directory the server had stopped reading, and the pane would silently keep
 * serving the stale copy. Losing an afternoon to that is easy.
 *
 * Embedded SEA assets come after the explicit overrides and the repo walk
 * (both are deliberate acts by someone who wants that copy served) but
 * before the `build/web` directory probes, because a stray extracted `web/`
 * next to an upgraded exe would otherwise silently outrank the UI that
 * actually matches the binary.
 *
 * This costs installed users nothing — an installed CLI has no
 * `packages/addin/dist` anywhere above it, so the walk falls through
 * immediately.
 */
export async function resolveWebSource(explicit?: string): Promise<WebSource> {
  const dirs: string[] = [];
  if (explicit) dirs.push(path.resolve(explicit));
  if (process.env.DVLOAD_WEB_ROOT) dirs.push(path.resolve(process.env.DVLOAD_WEB_ROOT));

  // Repo checkout: walk up looking for packages/addin/dist.
  let dir = HERE;
  for (let i = 0; i < 6; i++) {
    dirs.push(path.join(dir, "packages", "addin", "dist"));
    const parent = path.dirname(dir);
    if (parent === dir) break;
    dir = parent;
  }

  const found = await firstWithPane(dirs);
  if (found) return found;

  const sea = seaApi();
  if (sea) {
    const embedded = seaWebSource(sea);
    if (await embedded.read("taskpane.html")) return embedded;
  }

  // Installed npm layout: bundle.mjs copies the add-in dist to build/web,
  // alongside build/dvload.cjs.
  const installed = await firstWithPane([path.join(HERE, "web"), path.join(HERE, "..", "build", "web")]);
  if (installed) return installed;

  throw new Error(
    "Could not find the built add-in UI (no taskpane.html).\n" +
      "Build it first:\n" +
      "  npm run build --workspace=@dvload/addin\n" +
      "or point at an existing build with --web-root <dir> or DVLOAD_WEB_ROOT."
  );
}

async function firstWithPane(dirs: string[]): Promise<WebSource | null> {
  for (const d of dirs) {
    const source = diskWebSource(d);
    if (await source.read("taskpane.html")) return source;
  }
  return null;
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

/**
 * The authorize URL of the sign-in currently waiting on a browser, if any.
 *
 * One slot, not a map: the loopback flow binds one listener per process, so
 * there is never more than one live interactive sign-in. Set when MSAL hands
 * the interactive flow its URL, cleared the moment /api/signin settles —
 * after that the loopback listener is gone and the URL cannot complete.
 */
let pendingAuthorizeUrl: { environmentUrl: string; url: string } | null = null;

/**
 * Every provider this server builds is silent-only. Nothing here may open a
 * browser.
 *
 * The default provider escalates to `loginDelegated()` when the delegated
 * session is dead, which is right in a terminal and wrong in a sidecar. The
 * requests that reach these routes are background work — the pane listing
 * dataflows on first render, a token refresh in the middle of a long import —
 * and escalating means a system browser opening over Excel, plus an HTTP
 * response blocked until the user finishes signing in or DVLOAD_AUTH_TIMEOUT_MS
 * expires. The user did not ask for either.
 *
 * The pane already has an explicit sign-in path, `/api/signin`, driven by a
 * button. That is the only route allowed to put someone in front of Entra; a
 * dead session on any other route becomes a 401 with `needsSignIn: true` and
 * the pane asks.
 */
function providerFor(environmentUrl: string, forceUser: boolean): Promise<Thunk> {
  const key = `${environmentUrl}|${forceUser ? "user" : "auto"}`;
  let p = providers.get(key);
  if (!p) {
    p = getTokenProvider({ environmentUrl, forceUser, silentOnly: true });
    providers.set(key, p);
  }
  return p;
}

/**
 * The body of a 401 for a request that failed only because nobody is signed in.
 *
 * `needsSignIn` is the machine-readable half: the pane branches on it to show
 * a "click Sign in" prompt rather than treating the failure as an error it
 * should report. The message is the human half, and it names the environment
 * because sign-in is per environment — being signed in to one org says nothing
 * about another.
 */
function signInRequiredBody(environmentUrl: unknown): { error: string; needsSignIn: true } {
  let where = "this environment";
  if (typeof environmentUrl === "string" && environmentUrl) {
    try {
      where = new URL(environmentUrl).host;
    } catch {
      where = environmentUrl;
    }
  }
  return {
    error: `Not signed in to ${where}, or the session has expired. Click Sign in to continue.`,
    needsSignIn: true,
  };
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

async function serveStatic(res: ServerResponse, web: WebSource, urlPath: string): Promise<void> {
  let rel: string;
  try {
    rel = urlPath === "/" ? "taskpane.html" : decodeURIComponent(urlPath).replace(/^\/+/, "");
  } catch {
    // decodeURIComponent throws on a malformed escape (e.g. "/%").
    send(res, 400, "Bad request", "text/plain; charset=utf-8");
    return;
  }

  // Path traversal is rejected by the source itself, which is where the
  // knowledge of what "inside the root" means lives — a directory and a set
  // of embedded asset keys answer that question differently.
  if (safeSegments(rel) === null) {
    send(res, 403, "Forbidden", "text/plain; charset=utf-8");
    return;
  }

  const body = await web.read(rel);
  if (!body) {
    send(res, 404, `Not found: ${rel}`, "text/plain; charset=utf-8");
    return;
  }
  send(res, 200, body, MIME[path.extname(rel).toLowerCase()] ?? "application/octet-stream");
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

/**
 * A Dataverse client for the signed-in user of `environmentUrl`.
 *
 * Goes through UI_TOKEN_OPTS like every other UI path: reading dataflows is
 * harmless, but a request that could silently run as the Application User is
 * the thing that rule exists to prevent, and there is no reason for this
 * route to be the exception.
 */
async function uiClient(environmentUrl: string): Promise<DataverseClient> {
  const getToken = await providerFor(environmentUrl, UI_TOKEN_OPTS.forceUser);
  return new DataverseClient({ environmentUrl, getToken });
}

async function handleApi(route: string, body: Record<string, unknown>): Promise<unknown> {
  switch (route) {
    case "/api/account":
      return describeAuth(requiredString(body, "environmentUrl"));

    /*
     * Who this machine is already signed in as, across every environment.
     *
     * The pane asks this before an environment has been picked, so that
     * "sign in" and "choose where to write" can be two separate steps
     * instead of one. `/api/account` cannot answer it: sign-in state is
     * keyed by environment, and the pane has no environment yet.
     *
     * Only usernames and hosts leave the process — never tokens.
     */
    case "/api/accounts": {
      const sessions = await listDelegatedSessions();
      return {
        accounts: sessions
          .filter((s) => s.username)
          .map((s) => ({
            username: s.username as string,
            environmentUrl: s.environmentUrl,
            host: s.host,
          })),
      };
    }

    /* --- Live dataflows ---------------------------------------------------
     * The pane's counterpart to picking a .pqt off disk. Everything the
     * conversion needs lives in core/src/dataflow.ts; these two routes only
     * move bytes.
     */
    case "/api/dataflows": {
      const environmentUrl = requiredString(body, "environmentUrl");
      const client = await uiClient(environmentUrl);
      return { dataflows: await listDataflows(client, { includeDrafts: body.drafts === true }) };
    }

    case "/api/dataflow-import": {
      const environmentUrl = requiredString(body, "environmentUrl");
      const dataflowId = requiredString(body, "dataflowId");
      // Absent means "yes": the pane always sends both, but a caller that
      // omits them should get the same thing the checkboxes default to.
      const wantXlsx = body.xlsx !== false;
      const wantMapping = body.mapping !== false;

      const client = await uiClient(environmentUrl);
      const detail = await getDataflow(client, dataflowId);

      // Base64 rather than a binary response: this route already returns
      // JSON, and the pane hands the bytes to a Blob download either way.
      const workbook = wantXlsx
        ? Buffer.from(await buildWorkbookWithQueries(detail.archive)).toString("base64")
        : null;

      // Metadata reads (alternate keys, lookup targets) cost several round
      // trips, so they only happen when mappings were actually asked for.
      const mappings = wantMapping
        ? await mappingsFromDataflow(detail, {
            environmentUrl,
            resolver: createMetadataResolver(client),
          })
        : {};

      return {
        name: detail.name,
        queryNames: detail.queryNames,
        workbook,
        mappings: Object.entries(mappings).map(([queryName, mapping]) => ({
          queryName,
          mapping,
          // Sent alongside so the pane can warn in place instead of the user
          // finding out at load time.
          problems: validateMapping(mapping),
        })),
      };
    }

    case "/api/signin": {
      const environmentUrl = requiredString(body, "environmentUrl");
      // Which user the pane thinks it is signed in as, from an earlier
      // sign-in to a different environment. A hint only: it pre-fills the
      // Entra page so switching environment is usually a silent redirect
      // rather than another account picker. Entra decides what it's worth.
      const loginHint = typeof body.loginHint === "string" ? body.loginHint : undefined;
      // Try the session that already exists before putting anyone in front of
      // Entra. The pane asks to sign in whenever it *believes* it isn't — and
      // that belief can be stale (a transient /api/account failure, a race
      // during an environment switch). Acting on it opened a browser the user
      // had no reason to look at, whose loopback listener then timed out and
      // printed an alarming failure while imports kept working on the cached
      // session. A silent token acquisition is the proof either way: if it
      // succeeds for the hinted user, the sign-in is already done.
      // `force: true` (Sign in again / switch account) skips this — those
      // clicks exist precisely to replace the current session.
      if (body.force !== true) {
        try {
          const existing = await getSignedInAccount(environmentUrl);
          if (
            existing?.username &&
            (!loginHint || existing.username.toLowerCase() === loginHint.toLowerCase())
          ) {
            // The cache listing alone can name an account whose refresh token
            // is dead — only a real silent acquisition validates the session.
            await acquireToken(environmentUrl, UI_TOKEN_OPTS.forceUser);
            return describeAuth(environmentUrl);
          }
        } catch {
          // Dead or missing session — fall through to the interactive path.
        }
      }
      // Clear any memoised provider so the next token request picks up the
      // account we are about to create rather than a stale failed thunk.
      providers.clear();
      try {
        await loginDelegated({
          environmentUrl,
          loginHint,
          // The pane polls /api/signin-url while this request is in flight,
          // so a browser that failed to open (or opened behind Excel) still
          // puts a clickable sign-in link in front of the user instead of a
          // URL in a console they aren't watching.
          onAuthorizeUrl: (url) => {
            pendingAuthorizeUrl = { environmentUrl, url };
          },
        });
      } finally {
        // Whatever happened, the URL is dead now: the loopback listener for
        // it is gone, so offering it any longer would send a click into a
        // sign-in that can never complete.
        pendingAuthorizeUrl = null;
      }
      // One identity at a time: a sign-in as a different user signs the
      // previous user out of every environment. Two live accounts would make
      // "who will this import run as?" depend on which environment is
      // selected — a wrong answer to that question writes rows as the wrong
      // user. Runs only after a *successful* sign-in, so cancelling the
      // Entra page never strands the user signed out of everything.
      const signedIn = await getSignedInAccount(environmentUrl);
      if (signedIn?.username) {
        let removed = false;
        for (const session of await listDelegatedSessions()) {
          if (session.username && session.username.toLowerCase() !== signedIn.username.toLowerCase()) {
            await logoutDelegated(session.environmentUrl);
            removed = true;
          }
        }
        if (removed) providers.clear();
      }
      return describeAuth(environmentUrl);
    }

    case "/api/signin-url": {
      // Polled by the pane while its /api/signin request is in flight. Only
      // the URL for the environment being asked about is handed out — not
      // that a second environment's sign-in can be pending (see the single
      // slot above), but so a stale poll from an earlier attempt gets null
      // rather than a URL for somewhere else.
      const environmentUrl = requiredString(body, "environmentUrl");
      return pendingAuthorizeUrl && pendingAuthorizeUrl.environmentUrl === environmentUrl
        ? { url: pendingAuthorizeUrl.url }
        : { url: null };
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
  /** Where the UI is being served from. A directory, or a note that it is
   *  embedded in the executable — for display, not for path joining. */
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
  const web = await resolveWebSource(opts.webRoot);
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
      // Read outside the handler call so the error path can still name the
      // environment the request was about.
      let body: Record<string, unknown> = {};
      try {
        body = await readJsonBody(req);
        const result = await handleApi(urlPath, body);
        if (result === null) sendJson(res, 404, { error: `Unknown route ${urlPath}` });
        else sendJson(res, 200, result);
      } catch (e) {
        // Not a failure: an expired session is an ordinary thing to find, and
        // the pane's answer is to offer the Sign in button. Reported as 401
        // rather than 500 so it is distinguishable without reading the text.
        if (isInteractiveSignInRequired(e)) {
          console.log(kleur.yellow(`  ${urlPath}: sign-in required`));
          sendJson(res, 401, signInRequiredBody(body.environmentUrl));
          return;
        }
        const message = e instanceof Error ? e.message : String(e);
        console.error(kleur.red(`  ${urlPath} failed: ${message}`));
        sendJson(res, 500, { error: message });
      }
      return;
    }

    await serveStatic(res, web, urlPath);
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
    webRoot: web.label,
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
  if (opts.checkUi) {
    await checkUi(opts.webRoot);
    return;
  }
  await serveCommand({ ...opts, open: true });
}

/**
 * Resolve the UI and exit — no listener, no browser, no certificate.
 *
 * Exists for the release smoke test. The single-file exe embeds the task
 * pane as SEA assets, and the only way that can break is silently: the exe
 * runs, `--version` prints, and the failure surfaces the first time a user
 * opens the pane. This makes it a build-time error instead.
 */
export async function checkUi(webRoot?: string): Promise<void> {
  const web = await resolveWebSource(webRoot);
  const required = ["taskpane.html", "taskpane.js", "commands.html"];
  const missing: string[] = [];
  for (const f of required) {
    if (!(await web.read(f))) missing.push(f);
  }
  if (missing.length > 0) {
    throw new Error(
      `UI found at "${web.label}" but incomplete — missing: ${missing.join(", ")}.\n` +
        "Rebuild the add-in and re-bundle before releasing."
    );
  }
  console.log(kleur.green("UI OK"));
  console.log(kleur.gray(`  source: ${web.label}`));
}
