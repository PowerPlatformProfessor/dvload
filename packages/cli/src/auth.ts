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
} from "@azure/msal-node";
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
 * have blocked the shared client (Conditional Access, or Power Platform's
 * "allowed client apps" control).
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
    `\nSigning in via the shared Microsoft Dataverse client (${clientId}).\n` +
      `  No admin consent needed, but the consent screen and your tenant's\n` +
      `  sign-in logs will show "${name}", not dvload. To sign in as dvload\n` +
      `  instead, set DVLOAD_NO_SHARED_CLIENT=1 (an admin may then need to\n` +
      `  approve the app once).\n\n`
  );
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

export interface DelegatedAuthOptions {
  environmentUrl: string;
  /** Tenant for delegated flow. Default 'organizations' (multi-tenant). */
  tenantId?: string;
  /** Override the public-client app id. */
  clientId?: string;
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
 * Interactive device-code login.
 *
 * Walks the client-id chain from `resolveClientIdChain()`: the shared
 * Microsoft Dataverse client first (no consent prompt in any tenant), then
 * dvload's own app if that client is unusable here. Only client-level
 * failures advance the chain — see `shouldTryNextClient`.
 */
export async function loginDelegated(opts: DelegatedAuthOptions): Promise<AccountInfo> {
  const chain = resolveClientIdChain(opts.clientId);
  let lastErr: unknown;

  for (let i = 0; i < chain.length; i++) {
    const clientId = chain[i];
    const next = chain[i + 1];
    try {
      noteSharedClient(clientId);
      return await loginWithClient({ ...opts, clientId });
    } catch (e) {
      lastErr = e;
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

/** Single device-code attempt against one specific client id. */
async function loginWithClient(opts: DelegatedAuthOptions): Promise<AccountInfo> {
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
  if (!result?.account) throw new Error("Device-code flow returned no account.");
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
