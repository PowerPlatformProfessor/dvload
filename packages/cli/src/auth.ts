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
 * Microsoft-published multi-tenant public clients with Dataverse access
 * pre-consented. Convenient for local development — any work/school
 * account can sign in to them without anyone having to register a new
 * Entra ID app first.
 *
 * !!! DO NOT SHIP A TOOL WITH ONE OF THESE AS THE DEFAULT.  See
 * PRE-RELEASE-CHECKLIST.md at the repo root. Users would be signing in
 * to "Microsoft PowerApps" on the consent screen rather than to your
 * tool, which is confusing for them and against Microsoft's terms.
 */
const WELL_KNOWN_DEV_CLIENT_IDS: Record<string, string> = {
  "2ad88395-b77d-4561-9441-d0e40824f9bc": "Microsoft PowerApps",
  "51f81489-12ee-4a9e-aaae-a2591f45987d": "Microsoft Power Query",
};

/**
 * Default public-client app id used by `dvload login`. Pulled from
 * env var first; otherwise falls back to the PowerApps client id so
 * everything works out of the box for development without an Entra
 * registration. Replace with your own app id before shipping.
 */
const DEFAULT_PUBLIC_CLIENT_ID =
  process.env.DATAVERSE_LOAD_CLIENT_ID ?? "2ad88395-b77d-4561-9441-d0e40824f9bc";

export function isWellKnownDevClient(clientId: string): boolean {
  return clientId in WELL_KNOWN_DEV_CLIENT_IDS;
}

/**
 * Print a one-time warning to stderr if `clientId` is a borrowed public
 * client (PowerApps, Power Query, etc.). Call this from any command that
 * takes a user-facing action with the auth client (login, whoami, …).
 */
export function warnIfWellKnown(clientId: string): void {
  if (!isWellKnownDevClient(clientId)) return;
  const name = WELL_KNOWN_DEV_CLIENT_IDS[clientId];
  process.stderr.write(
    `\n[dev mode] Using "${name}" public client (${clientId}).\n` +
      `  This works for local testing but the consent screen shows "${name}",\n` +
      `  not your tool. Register your own Entra ID app and set\n` +
      `  DATAVERSE_LOAD_CLIENT_ID before sharing this with anyone.\n` +
      `  See PRE-RELEASE-CHECKLIST.md.\n\n`
  );
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
  appOnly: (env: string) => `appOnly:${host(env)}`,
};

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
  const clientId = opts.clientId ?? DEFAULT_PUBLIC_CLIENT_ID;
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

/** Interactive device-code login. Caches an account marker in the secure store. */
export async function loginDelegated(opts: DelegatedAuthOptions): Promise<AccountInfo> {
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
  return result.account;
}

export async function logoutDelegated(env?: string): Promise<void> {
  if (env) {
    await deleteSecret(keys.delegatedAccount(env));
    await deleteSecret(keys.delegatedTenant(env));
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
      // Reconstruct the app with the same tenant the login used, unless
      // the caller passed one explicitly. Without this, silent acquisition
      // fails because MSAL keys cache entries by authority.
      if (!opts.tenantId) {
        const storedTenant = await getSecret(keys.delegatedTenant(opts.environmentUrl));
        if (storedTenant) effectiveOpts = { ...opts, tenantId: storedTenant };
      }
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
