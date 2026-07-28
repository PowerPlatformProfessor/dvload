// Dataverse auth for the CLI. Supports two flows:
//
//   1. Delegated (user signs in)
//      - PublicClientApplication + device-code flow
//      - Refresh token cached in an encrypted MSAL cache file (DPAPI on
//        Windows, 0600 on POSIX)
//      - Best for interactive use
//
//   2. App-only (client credentials)
//      - ConfidentialClientApplication, no user
//      - Secret OR certificate; stored in the dvload secure store
//        (DPAPI-protected on Windows — see secure-store.ts)
//      - Best for scheduled/unattended runs (no 90-day refresh expiry,
//        no MFA / conditional-access surprises). Certificates avoid
//        secret-rotation churn entirely.
//
// `getTokenProvider()` picks app-only if it's configured for the
// environment, falling back to delegated. Pass `forceUser: true` to
// override and always use delegated (handy when developing).

import {
  PublicClientApplication,
  ConfidentialClientApplication,
  LogLevel,
  type Configuration,
  type AccountInfo,
  type AuthenticationResult,
  type ICachePlugin,
  type TokenCacheContext,
  type ILoopbackClient,
  type ServerAuthorizationCodeResponse,
} from "@azure/msal-node";
import type { Server } from "node:http";
import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";
import {
  getSecret,
  setSecret,
  deleteSecret,
  listAccounts,
  protectBytes,
  unprotectBytes,
} from "./secure-store.js";

/**
 * Microsoft-owned multi-tenant public clients that already have Dataverse
 * delegated access consented in every tenant. Because the consent record
 * exists tenant-wide out of the box, signing in through one of these never
 * hits the "Need admin approval" wall — which is how XrmToolBox, the XRM
 * Tooling SDK and the Power Platform CLI all connect without asking the
 * tenant admin for anything.
 *
 * This is not a privilege bypass: the flow is still delegated, the user
 * still authenticates as themselves, and Dataverse still enforces their
 * security roles. What you give up is branding and auditability — the
 * consent screen and the tenant's sign-in logs show the Microsoft app
 * name, not dvload.
 */
const SHARED_MS_CLIENTS: Record<string, string> = {
  // The Dataverse / XRM Tooling sample client. Same id XrmToolBox uses.
  "51f81489-12ee-4a9e-aaae-a2591f45987d": "Microsoft Dynamics CRM",
  "2ad88395-b77d-4561-9441-d0e40824f9bc": "Microsoft PowerApps",
};

/** Tried first: works in any tenant with no consent prompt. */
const SHARED_DATAVERSE_CLIENT_ID = "51f81489-12ee-4a9e-aaae-a2591f45987d";

/** Fallback: dvload's own registered multi-tenant public client. */
const DVLOAD_CLIENT_ID = "e6828b0f-9fde-43f8-85d0-602660d498bb";

/**
 * Public-client app ids to try, in order.
 *
 * Default is [shared Microsoft client, dvload's own app]: the shared client
 * needs no consent anywhere, and the own-app fallback covers tenants that
 * have blocked it (Conditional Access, or Power Platform's "allowed client
 * apps" control).
 *
 * UNRESOLVED: whether the shared client accepts an `http://localhost:<port>`
 * redirect for the interactive flow. One report of AADSTS900971 ("No reply
 * address provided") suggests not; against that, MSAL.NET now refuses
 * `app://`-style redirects on desktop and demands loopback, so tools like
 * XrmToolBox must be using loopback with *some* app id. Until that's settled
 * empirically (see docs/AUTH-NOTES.md) the chain is flow-agnostic: trying the
 * shared client first costs one clear error message and, if it works, saves
 * an admin-consent round trip entirely.
 *
 * An explicit `--client-id` or `DATAVERSE_LOAD_CLIENT_ID` disables the
 * chain entirely — if you named an app, that's the app we use. Set
 * `DVLOAD_NO_SHARED_CLIENT=1` to keep the default app without pinning one.
 */
export function resolveClientIdChain(explicit?: string): string[] {
  if (explicit) return [explicit];
  const pinned = process.env.DATAVERSE_LOAD_CLIENT_ID;
  if (pinned) return [pinned];
  if (process.env.DVLOAD_NO_SHARED_CLIENT === "1") return [DVLOAD_CLIENT_ID];
  return [SHARED_DATAVERSE_CLIENT_ID, DVLOAD_CLIENT_ID];
}

export function isSharedMicrosoftClient(clientId: string): boolean {
  return clientId in SHARED_MS_CLIENTS;
}

/** Human-readable name for log lines: the Microsoft app name, or the raw id. */
export function describeClient(clientId: string): string {
  const name = SHARED_MS_CLIENTS[clientId];
  if (name) return `"${name}" (${clientId})`;
  if (clientId === DVLOAD_CLIENT_ID) return `dvload's own app (${clientId})`;
  return clientId;
}

/**
 * Note on stderr which identity the user is actually signing in as. Not a
 * blocker — borrowing the shared client is a supported mode for the CLI —
 * but the consent screen will name Microsoft rather than dvload, and that
 * should never be a surprise.
 */
export function noteSharedClient(clientId: string): void {
  const name = SHARED_MS_CLIENTS[clientId];
  if (!name) return;
  process.stderr.write(
    `  No admin consent needed, but the consent screen and your tenant's\n` +
      `  sign-in logs will show "${name}", not dvload. To sign in as dvload\n` +
      `  instead, set DVLOAD_NO_SHARED_CLIENT=1 (an admin may then need to\n` +
      `  approve the app once).\n`
  );
}

/**
 * Say which identity we're about to sign in as, before doing it.
 *
 * This exists because of a genuinely painful debugging session: a stray
 * `DATAVERSE_LOAD_CLIENT_ID` in the environment silently pinned the client id,
 * the shared-client fallback never ran, and every symptom pointed at the wrong
 * app. The CLI knew exactly which client it was using and said nothing.
 *
 * Announcing the client id and where it came from costs two lines of output
 * and removes a whole class of misdiagnosis, so it is not optional or
 * debug-gated.
 */
export function announceLoginAttempt(clientId: string, flow: LoginFlow): void {
  const source = process.env.DATAVERSE_LOAD_CLIENT_ID
    ? " (pinned by DATAVERSE_LOAD_CLIENT_ID)"
    : process.env.DVLOAD_NO_SHARED_CLIENT === "1"
      ? " (shared client disabled by DVLOAD_NO_SHARED_CLIENT)"
      : "";
  const flowName = flow === "interactive" ? "browser" : "device code";
  process.stderr.write(`\nSigning in via ${flowName} as ${describeClient(clientId)}${source}.\n`);
  noteSharedClient(clientId);
  process.stderr.write("\n");
}

/* -------------------------------------------------------------------------- */
/* Fallback classification                                                     */
/* -------------------------------------------------------------------------- */

/** Pull the AADSTS error code out of whatever MSAL threw. */
export function aadstsCode(err: unknown): string | null {
  const text = [
    (err as { errorMessage?: string })?.errorMessage,
    (err as { errorCode?: string })?.errorCode,
    (err as Error)?.message,
  ]
    .filter(Boolean)
    .join(" ");
  return /AADSTS\d+/.exec(text)?.[0] ?? null;
}

/**
 * Client-level failures: this app registration cannot work in this tenant,
 * so a different one is worth trying.
 */
const RETRY_WITH_NEXT_CLIENT = new Set([
  "AADSTS65001", // no consent recorded for this client
  "AADSTS90094", // grant requires admin permission
  "AADSTS700016", // app not found in this directory
  "AADSTS7000218", // "Allow public client flows" is disabled
  "AADSTS7000112", // app disabled in this tenant
  "AADSTS50194", // single-tenant app reached via /organizations
  "AADSTS500011", // resource principal not found in tenant
  "AADSTS650057", // invalid resource for this client
  "AADSTS900971", // no reply address / client misconfiguration
  // http://localhost isn't a registered redirect URI on this app. Only the
  // interactive flow can hit this, and it's the empirical answer to "does
  // the shared Microsoft client support loopback?" — so advance the chain
  // to an app we know does.
  "AADSTS50011",
]);

/**
 * Failures about the *user* or an explicit policy decision. Retrying with a
 * different client id would just make the person do the device-code dance
 * again for nothing — or, worse, re-prompt someone who already said no.
 */
const DO_NOT_RETRY = new Set([
  "AADSTS65004", // user declined consent
  "AADSTS50105", // user not assigned to the app
  "AADSTS53000",
  "AADSTS53001",
  "AADSTS53002",
  "AADSTS53003", // blocked by Conditional Access
  "AADSTS50076",
  "AADSTS50079", // MFA required
  "AADSTS50158", // external security challenge not satisfied
]);

/** True when the next client id in the chain is worth attempting. */
export function shouldTryNextClient(err: unknown): boolean {
  const code = aadstsCode(err);
  if (code) {
    if (DO_NOT_RETRY.has(code)) return false;
    if (RETRY_WITH_NEXT_CLIENT.has(code)) return true;
  }
  const errorCode = (err as { errorCode?: string })?.errorCode ?? "";
  // The user walked away or cancelled — a second prompt helps nobody.
  if (/user_timeout_reached|device_code_expired|user_cancelled/.test(errorCode)) {
    return false;
  }
  return /invalid_client|unauthorized_client/.test(errorCode);
}

/** Conditional Access blocks. Worth explaining rather than dumping raw MSAL output. */
const CONDITIONAL_ACCESS_CODES = new Set([
  "AADSTS53003",
  "AADSTS53000",
  "AADSTS53001",
  "AADSTS53002",
  "AADSTS50005",
  "AADSTS50097",
  "AADSTS50158",
]);

export function isConditionalAccessBlock(err: unknown): boolean {
  const code = aadstsCode(err);
  return code ? CONDITIONAL_ACCESS_CODES.has(code) : false;
}

/**
 * Explain a Conditional Access rejection.
 *
 * The most common cause for a CLI is the "Authentication Flows" condition
 * blocking device code flow — Microsoft's own guidance recommends getting
 * "as close as possible to a unilateral block" on it, since an attacker can
 * generate a code and phish someone into entering it. Interactive browser
 * sign-in is not a way around that policy: it's the flow the policy steers
 * you to, and it works because a real browser can present device state
 * (primary refresh token, compliant-device claim) that device code
 * structurally cannot.
 *
 * If the blocking condition is device compliance or location instead, no
 * client-side change will or should help.
 */
export function conditionalAccessHint(err: unknown, flow: LoginFlow): string {
  const code = aadstsCode(err) ?? "AADSTS53003";
  const lines = [
    ``,
    `Conditional Access blocked this sign-in (${code}).`,
    `Authentication succeeded; your tenant then refused to issue the token.`,
    ``,
    `Find out which policy and condition: Microsoft Entra admin center ->`,
    `Monitoring -> Sign-in logs -> this attempt -> Conditional Access tab.`,
    `Look for the policy with Grant = Block and check its failed condition.`,
    ``,
  ];
  if (flow === "deviceCode") {
    lines.push(
      `Condition "Authentication Flows" / device code flow:`,
      `  Retry with a real browser instead — dvload login --interactive`,
      `  Many tenants block device code flow specifically; it's phishable,`,
      `  and Microsoft recommends blocking it.`,
      ``
    );
  }
  lines.push(
    `Condition "Require compliant device", "Require hybrid joined device"`,
    `or a location/IP rule:`,
    `  No dvload setting can satisfy these, by design. Either run from a`,
    `  managed device on an allowed network, or ask your admin to scope the`,
    `  policy so this app is usable. For unattended runs, app-only auth`,
    `  (dvload app-login) is evaluated separately from user CA policy.`,
    ``
  );
  return lines.join("\n");
}

/**
 * How the user proves who they are.
 *
 * - `interactive`: authorization code + PKCE in the system browser, with a
 *   loopback listener on 127.0.0.1. Default. The browser carries device
 *   state, so Conditional Access can evaluate the machine.
 * - `deviceCode`: print a code, user enters it elsewhere. The only option
 *   when there's no local browser (SSH, container, headless server), but
 *   widely blocked by Conditional Access because it's phishable.
 */
export type LoginFlow = "interactive" | "deviceCode";

export interface DelegatedAuthOptions {
  environmentUrl: string;
  /** Tenant for delegated flow. Default 'organizations' (multi-tenant). */
  tenantId?: string;
  /** Override the public-client app id. */
  clientId?: string;
  /** Override the sign-in flow. Default: see `defaultLoginFlow()`. */
  flow?: LoginFlow;
}

export interface AppOnlyCredentials {
  clientId: string;
  tenantId: string;
  /** Plain client secret (mutually exclusive with certificatePem). */
  secret?: string;
  /** PEM containing the certificate + private key (mutually exclusive with secret). */
  certificatePem?: string;
}

export interface TokenProviderOptions extends DelegatedAuthOptions {
  /** If true, always use delegated even when app-only is configured. */
  forceUser?: boolean;
  /** If true, don't fall back to interactive login on cache miss. */
  silentOnly?: boolean;
}

export type AuthMode = "appOnly" | "delegated" | "none";

/* -------------------------------------------------------------------------- */
/* Storage helpers                                                             */
/* -------------------------------------------------------------------------- */

function host(envUrl: string): string {
  return new URL(envUrl).host;
}

const keys = {
  delegatedAccount: (env: string) => `delegated:${host(env)}`,
  delegatedTenant: (env: string) => `delegatedTenant:${host(env)}`,
  // Which client id the successful login used. MSAL keys its cache entries
  // by client id, so silent acquisition has to reuse the same one or it
  // finds nothing and drops back to an interactive prompt.
  delegatedClient: (env: string) => `delegatedClient:${host(env)}`,
  appOnly: (env: string) => `appOnly:${host(env)}`,
};

/** Client id the last successful delegated login for `env` used, if any. */
export async function getStoredClientId(env: string): Promise<string | null> {
  return await getSecret(keys.delegatedClient(env));
}

/**
 * The delegated account currently cached for `env`, or null.
 *
 * Only the home account id is kept in the secure store, so the username has
 * to come back out of the MSAL cache — which means rebuilding the app with
 * the same tenant and client id the login used, exactly as
 * `makeDelegatedProvider` does. Purely a read: never prompts, and returns
 * null rather than throwing when the cache is missing or unreadable, because
 * every caller is rendering a status line.
 */
export async function getSignedInAccount(env: string): Promise<AccountInfo | null> {
  try {
    const accountId = await getSecret(keys.delegatedAccount(env));
    if (!accountId) return null;
    const app = makePublicApp({
      environmentUrl: env,
      tenantId: (await getSecret(keys.delegatedTenant(env))) ?? undefined,
      clientId: (await getStoredClientId(env)) ?? undefined,
    });
    return await app.getTokenCache().getAccountByHomeId(accountId);
  } catch {
    return null;
  }
}

/* -------------------------------------------------------------------------- */
/* Persistent MSAL cache                                                       */
/* -------------------------------------------------------------------------- */

function msalCachePath(env: string): string {
  // .bin: DPAPI-protected on Windows, plain 0600 JSON elsewhere.
  return path.join(os.homedir(), ".dvload", `msal-cache-${host(env)}.bin`);
}

function legacyMsalCachePath(env: string): string {
  return path.join(os.homedir(), ".dvload", `msal-cache-${host(env)}.json`);
}

/**
 * File-based ICachePlugin so MSAL's token cache (containing refresh
 * tokens) survives across process invocations. The cache holds refresh
 * tokens, so at rest it is DPAPI-protected (CurrentUser) on Windows and
 * written with mode 0600 on POSIX. Older plaintext .json caches are read
 * once for migration, then replaced and deleted.
 */
function makeCachePlugin(env: string): ICachePlugin {
  const filePath = msalCachePath(env);
  const legacyPath = legacyMsalCachePath(env);
  return {
    async beforeCacheAccess(ctx: TokenCacheContext) {
      try {
        const raw = await fs.readFile(filePath);
        ctx.tokenCache.deserialize((await unprotectBytes(raw)).toString("utf8"));
        return;
      } catch (e: unknown) {
        if ((e as NodeJS.ErrnoException).code !== "ENOENT") throw e;
      }
      // Migration: read the pre-encryption plaintext cache if present.
      try {
        const data = await fs.readFile(legacyPath, "utf8");
        ctx.tokenCache.deserialize(data);
      } catch (e: unknown) {
        if ((e as NodeJS.ErrnoException).code !== "ENOENT") throw e;
      }
    },
    async afterCacheAccess(ctx: TokenCacheContext) {
      if (ctx.cacheHasChanged) {
        await fs.mkdir(path.dirname(filePath), { recursive: true });
        const data = await protectBytes(Buffer.from(ctx.tokenCache.serialize(), "utf8"));
        await fs.writeFile(filePath, data, { mode: 0o600 });
        await fs.unlink(legacyPath).catch(() => {}); // remove plaintext copy
      }
    },
  };
}

export async function loadAppOnlyCredentials(env: string): Promise<AppOnlyCredentials | null> {
  const raw = await getSecret(keys.appOnly(env));
  if (!raw) return null;
  try {
    const creds = JSON.parse(raw) as AppOnlyCredentials;
    if (!creds.clientId || !creds.tenantId || (!creds.secret && !creds.certificatePem)) return null;
    return creds;
  } catch {
    return null;
  }
}

export async function saveAppOnlyCredentials(
  env: string,
  creds: AppOnlyCredentials
): Promise<void> {
  await setSecret(keys.appOnly(env), JSON.stringify(creds));
}

export async function clearAppOnlyCredentials(env: string): Promise<void> {
  await deleteSecret(keys.appOnly(env));
}

export async function detectAuthMode(env: string): Promise<AuthMode> {
  const appOnly = await loadAppOnlyCredentials(env);
  if (appOnly) return "appOnly";
  const delegated = await getSecret(keys.delegatedAccount(env));
  return delegated ? "delegated" : "none";
}

/* -------------------------------------------------------------------------- */
/* Scope                                                                       */
/* -------------------------------------------------------------------------- */

export function dataverseScope(environmentUrl: string): string {
  const u = new URL(environmentUrl);
  return `${u.origin}/.default`;
}

/* -------------------------------------------------------------------------- */
/* Certificates                                                                */
/* -------------------------------------------------------------------------- */

/**
 * Extract the certificate + private key from a PEM bundle and compute the
 * SHA-1 thumbprint MSAL needs for the client-assertion header (x5t).
 */
export function parseClientCertificate(pem: string): {
  thumbprint: string;
  privateKey: string;
} {
  const certMatch = /-----BEGIN CERTIFICATE-----([\s\S]*?)-----END CERTIFICATE-----/.exec(pem);
  if (!certMatch) {
    throw new Error("PEM file does not contain a CERTIFICATE block.");
  }
  if (!/-----BEGIN (?:RSA |EC )?PRIVATE KEY-----/.test(pem)) {
    throw new Error(
      "PEM file does not contain a PRIVATE KEY block. Provide a single PEM " +
        "with both the certificate and its (unencrypted) private key."
    );
  }
  const der = Buffer.from(certMatch[1].replace(/\s+/g, ""), "base64");
  const thumbprint = crypto.createHash("sha1").update(der).digest("hex").toUpperCase();
  return { thumbprint, privateKey: pem };
}

/* -------------------------------------------------------------------------- */
/* Flow selection                                                              */
/* -------------------------------------------------------------------------- */

/** Env snapshot, so this stays testable without mutating the real process. */
export interface FlowEnv {
  platform?: string;
  DVLOAD_AUTH_FLOW?: string;
  DISPLAY?: string;
  WAYLAND_DISPLAY?: string;
  SSH_CONNECTION?: string;
  SSH_TTY?: string;
}

/**
 * Interactive unless there's clearly no browser to open.
 *
 * A wrong guess is recoverable in both directions: interactive in a headless
 * shell prints a URL you can't click, and device code in a blocked tenant
 * fails with a CA error that tells you to pass --interactive. Erring toward
 * interactive is better because it's the flow tenants actually permit.
 */
export function defaultLoginFlow(env: FlowEnv = { ...process.env, platform: process.platform }): LoginFlow {
  const forced = env.DVLOAD_AUTH_FLOW?.toLowerCase();
  if (forced === "device-code" || forced === "devicecode") return "deviceCode";
  if (forced === "interactive") return "interactive";

  // Remote shell: the browser would open on the wrong machine.
  if (env.SSH_CONNECTION || env.SSH_TTY) return "deviceCode";

  // Unix without a display server has nothing to launch.
  const platform = env.platform ?? process.platform;
  if (platform !== "win32" && platform !== "darwin" && !env.DISPLAY && !env.WAYLAND_DISPLAY) {
    return "deviceCode";
  }
  return "interactive";
}

/** Thrown when we couldn't get a browser open — the one case worth falling back for. */
class BrowserLaunchError extends Error {
  readonly errorCode = "browser_launch_failed";
}

/**
 * Fail locally, with detail, rather than in the browser with a riddle.
 *
 * `AADSTS900971` ("No reply address provided") means the authorize request
 * carried no `redirect_uri`. Entra only reveals that after the user has typed
 * their password, as an HTML page the CLI never sees — so the same symptom
 * covers a library bug, a mangled URL, and a misconfigured app, with no way
 * to tell them apart. Checking the URL before we open it collapses that into
 * one precise error.
 */
export function assertUsableAuthorizeUrl(url: string): void {
  let parsed: URL;
  try {
    parsed = new URL(url);
  } catch {
    throw new Error(`MSAL produced an unparseable sign-in URL: ${url}`);
  }
  const redirectUri = parsed.searchParams.get("redirect_uri");
  if (!redirectUri) {
    const names = [...parsed.searchParams.keys()].join(", ") || "(none)";
    throw new Error(
      `MSAL built a sign-in URL with no redirect_uri. Entra would reject this ` +
        `with AADSTS900971 ("No reply address provided") only after you signed in.\n` +
        `  Parameters present: ${names}\n` +
        `  Full URL: ${url}`
    );
  }
  if (process.env.DATAVERSE_LOAD_DEBUG === "1") {
    process.stderr.write(`[auth] redirect_uri=${redirectUri}\n`);
    process.stderr.write(`[auth] authorize params: ${[...parsed.searchParams.keys()].join(", ")}\n`);
  }
}

/**
 * Open a URL in the user's default browser without pulling in a dependency.
 *
 * Windows is the fiddly one, and the wrong choice here corrupts the URL
 * rather than failing loudly — which produces a baffling AADSTS900971 ("No
 * reply address provided") from Entra, because the mangled request arrives
 * missing its redirect_uri.
 *
 *   - `cmd /c start` treats `&` as a command separator, and Entra authorize
 *     URLs are nothing but `&`.
 *   - `rundll32 url.dll,FileProtocolHandler` avoids the shell but mangles
 *     long, heavily-parameterised URLs. Do not use it here.
 *   - `powershell Start-Process` works but can be blocked by execution
 *     policy, which is likely in exactly the locked-down tenants that force
 *     us onto the browser flow in the first place.
 *   - `explorer.exe` invokes the shell's URL handler directly, takes the URL
 *     as a single argv element with no reparsing, and isn't subject to
 *     execution policy. It's the reliable option.
 *
 * Note that explorer.exe habitually exits non-zero even on success, so its
 * exit code is deliberately ignored; we only care that the spawn worked.
 */
export async function openInBrowser(url: string): Promise<void> {
  const { spawn } = await import("node:child_process");
  const platform = process.platform;
  const [cmd, args] =
    platform === "win32"
      ? ["explorer.exe", [url]]
      : platform === "darwin"
        ? ["open", [url]]
        : ["xdg-open", [url]];

  await new Promise<void>((resolve, reject) => {
    const child = spawn(cmd as string, args as string[], {
      stdio: "ignore",
      detached: true,
      shell: false,
    });
    child.once("error", (e) => reject(new BrowserLaunchError(`Could not launch a browser: ${e.message}`)));
    child.once("spawn", () => {
      child.unref();
      resolve();
    });
  });
}

/**
 * Loopback listener pinned to a specific port.
 *
 * msal-node's built-in client binds an ephemeral port, so the redirect URI is
 * `http://localhost:<random>` on every run. Entra is documented to ignore the
 * port when matching a registered `http://localhost` entry for public clients,
 * but that leniency depends on how the entry was registered — an app with a
 * specific `http://localhost:53682` entry would match one port and reject the
 * rest.
 *
 * Set `DVLOAD_LOOPBACK_PORT` to pin it. 53682 is a useful value to try: it's
 * the port Azure CLI uses, so it's the one most likely to be registered
 * explicitly on Microsoft-owned client apps.
 */
class FixedPortLoopbackClient implements ILoopbackClient {
  private server: Server | undefined;

  constructor(private readonly port: number) {}

  async listenForAuthCode(
    successTemplate?: string,
    errorTemplate?: string
  ): Promise<ServerAuthorizationCodeResponse> {
    if (this.server) {
      throw new Error("Loopback server already exists. Cannot create another.");
    }
    const { createServer } = await import("node:http");
    return await new Promise<ServerAuthorizationCodeResponse>((resolve, reject) => {
      this.server = createServer((req, res) => {
        if (!req.url) {
          res.end(errorTemplate ?? "Error: no url provided.");
          return;
        }
        const parsed = new URL(req.url, `http://localhost:${this.port}`);
        const params = Object.fromEntries(
          parsed.searchParams.entries()
        ) as ServerAuthorizationCodeResponse;
        // Ignore incidental requests (favicon and friends) so a browser
        // prefetch can't resolve the promise with an empty response.
        if (!params.code && !params.error) {
          res.end();
          return;
        }
        res.writeHead(200, { "content-type": "text/html" });
        res.end(
          params.code
            ? (successTemplate ?? "Signed in to dvload. You can close this tab.")
            : (errorTemplate ?? `Sign-in failed: ${params.error}`)
        );
        resolve(params);
      });
      this.server.on("error", reject);
      this.server.listen(this.port, "127.0.0.1");
    });
  }

  getRedirectUri(): string {
    if (!this.server?.listening) {
      throw new Error("Loopback server is not listening yet.");
    }
    return `http://localhost:${this.port}`;
  }

  closeServer(): void {
    if (this.server) {
      this.server.closeAllConnections?.();
      this.server.close();
      this.server = undefined;
    }
  }
}

/** Returns a pinned loopback client when DVLOAD_LOOPBACK_PORT is set. */
function makeLoopbackClient(): ILoopbackClient | undefined {
  const raw = process.env.DVLOAD_LOOPBACK_PORT;
  if (!raw) return undefined;
  const port = Number(raw);
  if (!Number.isInteger(port) || port < 1 || port > 65535) {
    throw new Error(`DVLOAD_LOOPBACK_PORT must be a port number, got "${raw}".`);
  }
  return new FixedPortLoopbackClient(port);
}

/**
 * True when interactive sign-in failed for mechanical reasons — no browser,
 * or nothing able to bind the loopback port. Device code is a genuine
 * alternative in those cases.
 *
 * A Conditional Access rejection is explicitly NOT one of these: device code
 * is the more restricted flow, so retrying with it would turn a clear error
 * into a confusing one.
 */
export function shouldFallBackToDeviceCode(err: unknown): boolean {
  if (isConditionalAccessBlock(err)) return false;
  const code =
    (err as { errorCode?: string })?.errorCode ?? (err as NodeJS.ErrnoException)?.code ?? "";
  return /browser_launch_failed|EADDRINUSE|EACCES|EPERM|ENOENT/.test(String(code));
}

/* -------------------------------------------------------------------------- */
/* Delegated flow                                                              */
/* -------------------------------------------------------------------------- */

function makePublicApp(opts: DelegatedAuthOptions): PublicClientApplication {
  const clientId = opts.clientId ?? resolveClientIdChain()[0];
  if (!clientId) {
    throw new Error(
      "No public-client app id set. Either pass --client-id, set DATAVERSE_LOAD_CLIENT_ID, " +
        "or pin a default in packages/cli/src/auth.ts."
    );
  }
  const tenant = opts.tenantId ?? "organizations";
  const verbose = process.env.DATAVERSE_LOAD_DEBUG === "1";
  const config: Configuration = {
    auth: { clientId, authority: `https://login.microsoftonline.com/${tenant}` },
    cache: { cachePlugin: makeCachePlugin(opts.environmentUrl) },
    ...(verbose && {
      system: {
        loggerOptions: {
          loggerCallback: (level, message, containsPii) => {
            if (containsPii) return;
            const levelName = LogLevel[level] ?? String(level);
            console.error(`[msal ${levelName}] ${message}`);
          },
          piiLoggingEnabled: false,
          logLevel: LogLevel.Verbose,
        },
      },
    }),
  };
  return new PublicClientApplication(config);
}

/**
 * Authorization code + PKCE in the system browser, with msal-node's loopback
 * listener catching the redirect on 127.0.0.1.
 *
 * Requires `http://localhost` as a registered redirect URI on the app. Entra
 * treats that entry specially for public clients and accepts any port, per
 * RFC 8252.
 *
 * Failures here are awkward: Entra validates the redirect URI only AFTER
 * authentication, and renders the result as a page in the user's browser.
 * The loopback listener never receives anything, so MSAL sees a hang rather
 * than an error and the client-id chain can't react. Hence the timeout below,
 * whose message tells the user to go read the browser tab.
 */
async function acquireInteractive(opts: DelegatedAuthOptions): Promise<AuthenticationResult> {
  const app = makePublicApp(opts);
  const timeoutMs = Number(process.env.DVLOAD_AUTH_TIMEOUT_MS ?? 180_000);

  // Errors at the /authorize endpoint (bad redirect URI, unconsented app,
  // Conditional Access) render as a page in the browser and never come back
  // to the loopback listener. Without a timeout the CLI would sit there
  // forever while the answer sits on screen, so say that out loud.
  const clientId = opts.clientId ?? resolveClientIdChain()[0];
  let timer: NodeJS.Timeout | undefined;
  const timeout = new Promise<never>((_, reject) => {
    timer = setTimeout(() => {
      reject(
        new Error(
          `No response from the browser after ${Math.round(timeoutMs / 1000)}s.\n` +
            `  Sign-in never completed, so whatever the Entra page showed is the\n` +
            `  real cause. Entra only validates the redirect URI AFTER you enter\n` +
            `  credentials, and reports it in the browser, which is why dvload\n` +
            `  can only report a timeout here.\n\n` +
            `  AADSTS900971 ("No reply address provided") means app ${clientId}\n` +
            `  has no redirect URI matching http://localhost registered under\n` +
            `  Authentication -> Mobile and desktop applications.\n` +
            `  Either register it, or unset DATAVERSE_LOAD_CLIENT_ID to use the\n` +
            `  pre-consented shared Microsoft client, which already accepts it.\n\n` +
            `  Set DVLOAD_AUTH_TIMEOUT_MS to wait longer.`
        )
      );
    }, timeoutMs);
    timer.unref();
  });

  const loopbackClient = makeLoopbackClient();
  const attempt = app.acquireTokenInteractive({
    scopes: [dataverseScope(opts.environmentUrl)],
    ...(loopbackClient && { loopbackClient }),
    openBrowser: async (url: string) => {
      assertUsableAuthorizeUrl(url);
      console.log("");
      console.log("Opening your browser to sign in...");
      console.log("If nothing opens, paste this URL yourself:");
      console.log(`  ${url}`);
      console.log("");
      await openInBrowser(url);
    },
    successTemplate:
      "<html><body style=\"font-family:system-ui;padding:3rem\">" +
      "<h2>Signed in to dvload</h2><p>You can close this tab.</p></body></html>",
    errorTemplate:
      "<html><body style=\"font-family:system-ui;padding:3rem\">" +
      "<h2>Sign-in failed</h2><p>Check the dvload output in your terminal.</p></body></html>",
  });

  try {
    return await Promise.race([attempt, timeout]);
  } finally {
    if (timer) clearTimeout(timer);
    // On the timeout path msal-node never gets to tear this down itself.
    loopbackClient?.closeServer();
  }
}

/** Device-code flow: prints a code for the user to enter elsewhere. */
async function acquireByDeviceCode(opts: DelegatedAuthOptions): Promise<AuthenticationResult> {
  const app = makePublicApp(opts);
  const result = await app.acquireTokenByDeviceCode({
    scopes: [dataverseScope(opts.environmentUrl)],
    deviceCodeCallback: (info) => {
      // Print the URL and code ourselves rather than trusting info.message,
      // which is empty in some msal-node versions.
      console.log("");
      console.log("=========================================================");
      console.log(`  Open a browser to: ${info.verificationUri || "https://microsoft.com/devicelogin"}`);
      console.log(`  Enter code:        ${info.userCode || "(msal returned no code — see JSON below)"}`);
      if (info.expiresIn) console.log(`  Code expires in:   ${info.expiresIn} seconds`);
      console.log("=========================================================");
      if (!info.userCode || !info.verificationUri) {
        console.log("Raw device-code response:", JSON.stringify(info, null, 2));
      }
      console.log("");
    },
  });
  if (!result) throw new Error("Device-code flow returned no result.");
  return result;
}

/**
 * Interactive login.
 *
 * Walks the client-id chain from `resolveClientIdChain()`: the shared
 * Microsoft Dataverse client first (no consent prompt in any tenant), then
 * dvload's own app if that client is unusable here. Only client-level
 * failures advance the chain — see `shouldTryNextClient`.
 *
 * Uses the browser flow by default; see `defaultLoginFlow()`.
 */
export async function loginDelegated(opts: DelegatedAuthOptions): Promise<AccountInfo> {
  const flow = opts.flow ?? defaultLoginFlow();
  const chain = resolveClientIdChain(opts.clientId);
  let lastErr: unknown;

  for (let i = 0; i < chain.length; i++) {
    const clientId = chain[i];
    const next = chain[i + 1];
    try {
      announceLoginAttempt(clientId, flow);
      return await loginWithClient({ ...opts, clientId, flow });
    } catch (e) {
      lastErr = e;
      if (isConditionalAccessBlock(e)) {
        // Don't churn through client ids: a CA block is a policy decision,
        // and the user needs to know what to do about it.
        process.stderr.write(conditionalAccessHint(e, flow));
        throw e;
      }
      if (!next || !shouldTryNextClient(e)) throw e;
      process.stderr.write(
        `\nSign-in with ${describeClient(clientId)} failed` +
          `${aadstsCode(e) ? ` (${aadstsCode(e)})` : ""}.\n` +
          `  Retrying with ${describeClient(next)}.\n\n`
      );
    }
  }
  throw lastErr;
}

/**
 * One sign-in attempt against one client id.
 *
 * Runs the requested flow, and drops from interactive to device code only
 * when the browser or loopback listener couldn't be started at all.
 */
async function loginWithClient(opts: DelegatedAuthOptions): Promise<AccountInfo> {
  const flow = opts.flow ?? defaultLoginFlow();
  let result: AuthenticationResult | null = null;

  if (flow === "interactive") {
    try {
      result = await acquireInteractive(opts);
    } catch (e) {
      if (!shouldFallBackToDeviceCode(e)) throw e;
      process.stderr.write(
        `\nCouldn't complete browser sign-in (${(e as Error).message}).\n` +
          `  Falling back to device code. If your tenant blocks device code\n` +
          `  flow, fix the browser issue instead — see --interactive.\n\n`
      );
      result = await acquireByDeviceCode(opts);
    }
  } else {
    result = await acquireByDeviceCode(opts);
  }

  if (!result?.account) throw new Error("Sign-in returned no account.");
  await setSecret(keys.delegatedAccount(opts.environmentUrl), result.account.homeAccountId);
  // Remember which tenant this login used so subsequent commands
  // (whoami, run, validate) can hit the same authority without --tenant.
  await setSecret(
    keys.delegatedTenant(opts.environmentUrl),
    opts.tenantId ?? result.account.tenantId ?? "organizations"
  );
  // Remember the winning client id too — acquireTokenSilent must use it.
  if (opts.clientId) {
    await setSecret(keys.delegatedClient(opts.environmentUrl), opts.clientId);
  }
  return result.account;
}

export async function logoutDelegated(env?: string): Promise<void> {
  if (env) {
    await deleteSecret(keys.delegatedAccount(env));
    await deleteSecret(keys.delegatedTenant(env));
    await deleteSecret(keys.delegatedClient(env));
    await fs.unlink(msalCachePath(env)).catch(() => {});
    await fs.unlink(legacyMsalCachePath(env)).catch(() => {});
    return;
  }
  // Clear every delegated:* entry. Leave app-only creds alone.
  for (const account of await listAccounts("delegated")) {
    await deleteSecret(account);
  }
  // Cache files are per-env; delete every one we can find.
  const dir = path.join(os.homedir(), ".dvload");
  const files = await fs.readdir(dir).catch(() => [] as string[]);
  for (const f of files) {
    if (/^msal-cache-.*\.(bin|json)$/.test(f)) {
      await fs.unlink(path.join(dir, f)).catch(() => {});
    }
  }
}

function makeDelegatedProvider(opts: TokenProviderOptions): () => Promise<string> {
  let app: PublicClientApplication | null = null;
  let cachedAccount: AccountInfo | null = null;
  let warmedUp = false;
  let effectiveOpts: DelegatedAuthOptions = opts;

  return async () => {
    const scopes = [dataverseScope(opts.environmentUrl)];

    if (!warmedUp) {
      // Reconstruct the app with the same tenant AND client id the login
      // used, unless the caller passed them explicitly. Without this,
      // silent acquisition fails because MSAL keys cache entries by both
      // authority and client id — a mismatch looks like an empty cache and
      // silently escalates to a fresh device-code prompt.
      const overrides: Partial<DelegatedAuthOptions> = {};
      if (!opts.tenantId) {
        const storedTenant = await getSecret(keys.delegatedTenant(opts.environmentUrl));
        if (storedTenant) overrides.tenantId = storedTenant;
      }
      if (!opts.clientId) {
        const storedClient = await getStoredClientId(opts.environmentUrl);
        if (storedClient) overrides.clientId = storedClient;
      }
      effectiveOpts = { ...opts, ...overrides };
      app = makePublicApp(effectiveOpts);

      const accountId = await getSecret(keys.delegatedAccount(opts.environmentUrl));
      if (accountId) {
        cachedAccount = await app.getTokenCache().getAccountByHomeId(accountId);
      }
      warmedUp = true;
    }
    if (!app) throw new Error("Unreachable: MSAL app not initialized");

    let result: AuthenticationResult | null = null;
    if (cachedAccount) {
      try {
        result = await app.acquireTokenSilent({ account: cachedAccount, scopes });
      } catch {
        result = null;
      }
    }
    if (!result) {
      if (opts.silentOnly) {
        throw new Error(
          `No valid cached token for ${opts.environmentUrl}. ` +
            `Run \`dvload login --env ${opts.environmentUrl}\` again.`
        );
      }
      const account = await loginDelegated(effectiveOpts);
      cachedAccount = await app.getTokenCache().getAccountByHomeId(account.homeAccountId);
      result = await app.acquireTokenSilent({ account: cachedAccount!, scopes });
    }
    if (!result?.accessToken) throw new Error("Failed to acquire delegated access token.");
    return result.accessToken;
  };
}

/* -------------------------------------------------------------------------- */
/* App-only flow                                                               */
/* -------------------------------------------------------------------------- */

function makeAppOnlyProvider(env: string, creds: AppOnlyCredentials): () => Promise<string> {
  const authority = `https://login.microsoftonline.com/${creds.tenantId}`;
  let auth: Configuration["auth"];
  if (creds.certificatePem) {
    const { thumbprint, privateKey } = parseClientCertificate(creds.certificatePem);
    auth = { clientId: creds.clientId, authority, clientCertificate: { thumbprint, privateKey } };
  } else {
    auth = { clientId: creds.clientId, authority, clientSecret: creds.secret! };
  }
  const app = new ConfidentialClientApplication({ auth });
  return async () => {
    const r = await app.acquireTokenByClientCredential({ scopes: [dataverseScope(env)] });
    if (!r?.accessToken) throw new Error("client_credentials grant failed (no token returned).");
    return r.accessToken;
  };
}

/* -------------------------------------------------------------------------- */
/* Public entry point                                                          */
/* -------------------------------------------------------------------------- */

/**
 * Returns a getToken() thunk for DataverseClient. Picks app-only if
 * credentials are stored for this environment; otherwise delegated.
 *
 * Set forceUser=true to always use delegated (handy when iterating in a
 * shell where the schedule has stored an app secret).
 */
export async function getTokenProvider(opts: TokenProviderOptions): Promise<() => Promise<string>> {
  if (!opts.forceUser) {
    const creds = await loadAppOnlyCredentials(opts.environmentUrl);
    if (creds) return makeAppOnlyProvider(opts.environmentUrl, creds);
  }
  return makeDelegatedProvider(opts);
}
