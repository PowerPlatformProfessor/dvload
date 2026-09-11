import kleur from "kleur";
import { readFile } from "node:fs/promises";

import {
  loginDelegated,
  logoutDelegated,
  saveAppOnlyCredentials,
  clearAppOnlyCredentials,
  loadAppOnlyCredentials,
  detectAuthMode,
  getTokenProvider,
  dataverseScope,
  describeClient,
  isSharedMicrosoftClient,
  getStoredClientId,
  parseClientCertificate,
  type LoginFlow,
} from "../auth.js";
import { promptSecret } from "../prompt.js";
import { resolveEnv } from "../profiles.js";

interface LoginOpts {
  env?: string;
  profile?: string;
  tenant?: string;
  clientId?: string;
  interactive?: boolean;
  deviceCode?: boolean;
}

export async function loginCommand(opts: LoginOpts): Promise<void> {
  const envUrl = await resolveEnv(opts);
  if (opts.interactive && opts.deviceCode) {
    throw new Error("Pass either --interactive or --device-code, not both.");
  }
  const flow: LoginFlow | undefined = opts.interactive
    ? "interactive"
    : opts.deviceCode
      ? "deviceCode"
      : undefined; // let defaultLoginFlow() decide
  // loginDelegated walks the client-id chain and reports which one it used.
  const account = await loginDelegated({
    environmentUrl: envUrl,
    tenantId: opts.tenant,
    clientId: opts.clientId,
    flow,
  });
  console.log(kleur.green(`Signed in as ${account.username} (${envUrl}).`));
  const used = await getStoredClientId(envUrl);
  if (used) console.log(kleur.gray(`Signed in via ${describeClient(used)}.`));
  console.log(
    kleur.gray(
      "Account reference stored in the dvload secure store; encrypted MSAL token cache written to ~/.dvload/."
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
  /** Path to a PEM containing the certificate + private key (instead of a secret). */
  cert?: string;
}

export async function appLoginCommand(opts: AppLoginOpts): Promise<void> {
  const envUrl = await resolveEnv(opts);

  if (opts.cert) {
    const pem = await readFile(opts.cert, "utf8");
    parseClientCertificate(pem); // validate before saving
    await saveAppOnlyCredentials(envUrl, {
      clientId: opts.clientId,
      tenantId: opts.tenantId,
      certificatePem: pem,
    });
  } else {
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
  }

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
        "been created in Dataverse yet. See README → \"App-only auth setup\".",
      { cause: e }
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

  if (mode === "delegated") {
    const used = await getStoredClientId(envUrl);
    if (used) {
      console.log(`  client id: ${describeClient(used)}`);
      if (isSharedMicrosoftClient(used)) {
        console.log(
          kleur.gray(
            "  Shared Microsoft client — no admin consent required, but sign-in\n" +
              "  logs attribute this to Microsoft rather than dvload."
          )
        );
      }
    }
  }

  if (mode === "appOnly") {
    const creds = await loadAppOnlyCredentials(envUrl);
    console.log(`  client id: ${creds!.clientId}`);
    console.log(`  tenant id: ${creds!.tenantId}`);
    console.log(
      `  credential: ${kleur.gray(
        creds!.certificatePem ? "certificate (stored in secure store)" : "secret (stored in secure store)"
      )}`
    );
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
