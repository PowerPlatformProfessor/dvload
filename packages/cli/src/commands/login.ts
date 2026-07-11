import kleur from "kleur";

import {
  loginDelegated,
  logoutDelegated,
  saveAppOnlyCredentials,
  clearAppOnlyCredentials,
  loadAppOnlyCredentials,
  detectAuthMode,
  getTokenProvider,
  dataverseScope,
  warnIfWellKnown,
  isWellKnownDevClient,
} from "../auth.js";
import { promptSecret } from "../prompt.js";
import { resolveEnv } from "../profiles.js";

const DEFAULT_PUBLIC_CLIENT_ID =
  process.env.DATAVERSE_LOAD_CLIENT_ID ?? "2ad88395-b77d-4561-9441-d0e40824f9bc";

interface LoginOpts {
  env?: string;
  profile?: string;
  tenant?: string;
  clientId?: string;
}

export async function loginCommand(opts: LoginOpts): Promise<void> {
  const envUrl = await resolveEnv(opts);
  warnIfWellKnown(opts.clientId ?? DEFAULT_PUBLIC_CLIENT_ID);
  const account = await loginDelegated({
    environmentUrl: envUrl,
    tenantId: opts.tenant,
    clientId: opts.clientId,
  });
  console.log(kleur.green(`Signed in as ${account.username} (${envUrl}).`));
  console.log(
    kleur.gray(
      "Account reference stored in the OS keychain; MSAL token cache written to ~/.dvload/."
    )
  );
}

export async function logoutCommand(opts: { env?: string; profile?: string }): Promise<void> {
  // If neither flag is given, clear all (existing behaviour).
  const envUrl = (opts.env || opts.profile) ? await resolveEnv(opts) : undefined;
  await logoutDelegated(envUrl);
  console.log(
    kleur.green(envUrl ? `Cleared delegated token for ${envUrl}.` : "Cleared all delegated tokens.")
  );
}

interface AppLoginOpts {
  env?: string;
  profile?: string;
  clientId: string;
  tenantId: string;
  /** If set, read the secret from this env var instead of prompting. */
  secretEnv?: string;
}

export async function appLoginCommand(opts: AppLoginOpts): Promise<void> {
  const envUrl = await resolveEnv(opts);
  let secret = opts.secretEnv ? process.env[opts.secretEnv] : undefined;
  if (!secret) {
    secret = await promptSecret("Client secret (input hidden): ");
  }
  if (!secret) throw new Error("Client secret was empty.");

  await saveAppOnlyCredentials(envUrl, {
    clientId: opts.clientId,
    tenantId: opts.tenantId,
    secret,
  });

  // Probe the credentials by acquiring a token. Cleaner UX than failing
  // later during a real run.
  try {
    const getToken = await getTokenProvider({ environmentUrl: envUrl });
    await getToken();
    console.log(kleur.green(`App-only credentials saved for ${envUrl}. Token acquired successfully.`));
    console.log(kleur.gray(`Scope: ${dataverseScope(envUrl)}`));
  } catch (e) {
    // Roll back so we don't leave broken creds in the keychain.
    await clearAppOnlyCredentials(envUrl);
    throw new Error(
      `Saved credentials but token acquisition failed (rolled back): ${(e as Error).message}\n` +
        "Common causes: wrong tenant id, wrong client secret, or the Application User has not " +
        "been created in Dataverse yet. See README → \"App-only auth setup\"."
    );
  }
}

export async function appLogoutCommand(opts: { env?: string; profile?: string }): Promise<void> {
  const envUrl = await resolveEnv(opts);
  await clearAppOnlyCredentials(envUrl);
  console.log(kleur.green(`Cleared app-only credentials for ${envUrl}.`));
}

interface WhoamiOpts {
  env?: string;
  profile?: string;
}

export async function whoamiCommand(opts: WhoamiOpts): Promise<void> {
  const envUrl = await resolveEnv(opts);
  const mode = await detectAuthMode(envUrl);
  console.log(kleur.bold(`Auth mode for ${envUrl}: ${mode}`));

  if (mode === "delegated" && isWellKnownDevClient(DEFAULT_PUBLIC_CLIENT_ID)) {
    console.log(
      kleur.yellow(
        `  [dev mode] borrowed public client (${DEFAULT_PUBLIC_CLIENT_ID})`
      )
    );
    console.log(kleur.gray("  Register your own app and set DATAVERSE_LOAD_CLIENT_ID before shipping."));
  }

  if (mode === "appOnly") {
    const creds = await loadAppOnlyCredentials(envUrl);
    console.log(`  client id: ${creds!.clientId}`);
    console.log(`  tenant id: ${creds!.tenantId}`);
    console.log(`  secret:    ${kleur.gray("(stored in Credential Manager)")}`);
  } else if (mode === "delegated") {
    console.log(kleur.gray("  refresh token cached for the last user that signed in."));
  } else {
    console.log(kleur.yellow("  Not configured. Run `dvload login` or `dvload app-login`."));
    return;
  }

  // Try to acquire a token to confirm things still work. silentOnly=true
  // so a cache miss reports failure instead of triggering a device-code
  // prompt (whoami is a status command, not an auth command).
  try {
    const getToken = await getTokenProvider({ environmentUrl: envUrl, silentOnly: true });
    await getToken();
    console.log(kleur.green("  Token acquisition: OK"));
  } catch (e) {
    console.log(kleur.red(`  Token acquisition FAILED: ${(e as Error).message}`));
    process.exitCode = 1;
  }
}
