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

const DEFAULT_PUBLIC_CLIENT_ID =
  process.env.DATAVERSE_LOAD_CLIENT_ID ?? "2ad88395-b77d-4561-9441-d0e40824f9bc";

interface LoginOpts {
  env: string;
  tenant?: string;
  clientId?: string;
}

export async function loginCommand(opts: LoginOpts): Promise<void> {
  warnIfWellKnown(opts.clientId ?? DEFAULT_PUBLIC_CLIENT_ID);
  const account = await loginDelegated({
    environmentUrl: opts.env,
    tenantId: opts.tenant,
    clientId: opts.clientId,
  });
  console.log(kleur.green(`Signed in as ${account.username} (${opts.env}).`));
  console.log(kleur.gray("Refresh token cached in Windows Credential Manager."));
}

export async function logoutCommand(opts: { env?: string }): Promise<void> {
  await logoutDelegated(opts.env);
  console.log(
    kleur.green(opts.env ? `Cleared delegated token for ${opts.env}.` : "Cleared all delegated tokens.")
  );
}

interface AppLoginOpts {
  env: string;
  clientId: string;
  tenantId: string;
  /** If set, read the secret from this env var instead of prompting. */
  secretEnv?: string;
}

export async function appLoginCommand(opts: AppLoginOpts): Promise<void> {
  let secret = opts.secretEnv ? process.env[opts.secretEnv] : undefined;
  if (!secret) {
    secret = await promptSecret("Client secret (input hidden): ");
  }
  if (!secret) throw new Error("Client secret was empty.");

  await saveAppOnlyCredentials(opts.env, {
    clientId: opts.clientId,
    tenantId: opts.tenantId,
    secret,
  });

  // Probe the credentials by acquiring a token. Cleaner UX than failing
  // later during a real run.
  try {
    const getToken = await getTokenProvider({ environmentUrl: opts.env });
    await getToken();
    console.log(kleur.green(`App-only credentials saved for ${opts.env}. Token acquired successfully.`));
    console.log(kleur.gray(`Scope: ${dataverseScope(opts.env)}`));
  } catch (e) {
    // Roll back so we don't leave broken creds in the keychain.
    await clearAppOnlyCredentials(opts.env);
    throw new Error(
      `Saved credentials but token acquisition failed (rolled back): ${(e as Error).message}\n` +
        "Common causes: wrong tenant id, wrong client secret, or the Application User has not " +
        "been created in Dataverse yet. See README → \"App-only auth setup\"."
    );
  }
}

export async function appLogoutCommand(opts: { env: string }): Promise<void> {
  await clearAppOnlyCredentials(opts.env);
  console.log(kleur.green(`Cleared app-only credentials for ${opts.env}.`));
}

interface WhoamiOpts {
  env: string;
}

export async function whoamiCommand(opts: WhoamiOpts): Promise<void> {
  const mode = await detectAuthMode(opts.env);
  console.log(kleur.bold(`Auth mode for ${opts.env}: ${mode}`));

  if (mode === "delegated" && isWellKnownDevClient(DEFAULT_PUBLIC_CLIENT_ID)) {
    console.log(
      kleur.yellow(
        `  [dev mode] borrowed public client (${DEFAULT_PUBLIC_CLIENT_ID})`
      )
    );
    console.log(kleur.gray("  Register your own app and set DATAVERSE_LOAD_CLIENT_ID before shipping."));
  }

  if (mode === "appOnly") {
    const creds = await loadAppOnlyCredentials(opts.env);
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
    const getToken = await getTokenProvider({ environmentUrl: opts.env, silentOnly: true });
    await getToken();
    console.log(kleur.green("  Token acquisition: OK"));
  } catch (e) {
    console.log(kleur.red(`  Token acquisition FAILED: ${(e as Error).message}`));
    process.exitCode = 1;
  }
}
