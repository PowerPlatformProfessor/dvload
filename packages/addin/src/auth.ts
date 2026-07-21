// Browser-side auth for the Office.js add-in. Uses MSAL.js with a popup
// flow: simplest path that doesn't require a backend. The downside is the
// user sees a popup on first sign-in and after token expiry; for personal
// productivity that's fine.
//
// If you later want SSO (no popup), swap to OfficeRuntime.auth.getAccessToken
// and add a tiny on-behalf-of token-exchange backend.

import {
  PublicClientApplication,
  type AccountInfo,
  type Configuration,
} from "@azure/msal-browser";

/**
 * Microsoft-published multi-tenant public clients with Dataverse access
 * pre-consented. Used as a dev-mode fallback. NEVER ship a tool with one
 * of these as the default — the consent screen would say "Microsoft
 * PowerApps" instead of your tool's name. See PRE-RELEASE-CHECKLIST.md.
 */
const WELL_KNOWN_DEV_CLIENT_IDS: Record<string, string> = {
  "2ad88395-b77d-4561-9441-d0e40824f9bc": "Microsoft PowerApps",
  "51f81489-12ee-4a9e-aaae-a2591f45987d": "Microsoft Power Query",
};

declare const ADDIN_CLIENT_ID: string; // injected by webpack DefinePlugin at build time
const CLIENT_ID = ADDIN_CLIENT_ID;

let msal: PublicClientApplication | null = null;

function warnIfDevModeClient(): void {
  const name = WELL_KNOWN_DEV_CLIENT_IDS[CLIENT_ID];
  if (!name) return;
  // eslint-disable-next-line no-console
  console.warn(
    `[dvload dev mode] Using "${name}" public client (${CLIENT_ID}).\n` +
      `Consent screen will show "${name}", not your add-in. Register your own ` +
      `Entra ID app and set DATAVERSE_LOAD_CLIENT_ID before publishing.`
  );
}

/** Returns the dev-mode banner text, or null when a custom client id is in use. */
export function devModeBanner(): string | null {
  const name = WELL_KNOWN_DEV_CLIENT_IDS[CLIENT_ID];
  return name ? `Dev mode — signing in as "${name}"` : null;
}

function ensure(): PublicClientApplication {
  if (msal) return msal;
  const config: Configuration = {
    auth: {
      clientId: CLIENT_ID,
      authority: "https://login.microsoftonline.com/organizations",
      // origin + pathname (not a hardcoded "/taskpane.html") so this works
      // both at https://localhost:3000/taskpane.html and when hosted under a
      // subpath, e.g. https://<user>.github.io/dvload/taskpane.html.
      redirectUri: window.location.origin + window.location.pathname,
    },
    cache: {
      cacheLocation: "localStorage",
      storeAuthStateInCookie: false,
    },
  };
  msal = new PublicClientApplication(config);
  return msal;
}

export function dataverseScope(environmentUrl: string): string {
  return new URL(environmentUrl).origin + "/.default";
}

export async function initAuth(): Promise<void> {
  warnIfDevModeClient();
  // Office WebView (WebView2 / Trident) can lose the `window` binding on `fetch`,
  // causing "Illegal invocation" inside MSAL. Rebind it before MSAL initialises.
  if (typeof window !== "undefined" && window.fetch) {
    window.fetch = window.fetch.bind(window);
  }
  const app = ensure();
  await app.initialize();
  await app.handleRedirectPromise();
}

export async function getAccount(): Promise<AccountInfo | null> {
  const app = ensure();
  const accounts = app.getAllAccounts();
  return accounts[0] ?? null;
}

export async function signIn(environmentUrl: string): Promise<AccountInfo> {
  const app = ensure();
  const result = await app.loginPopup({
    scopes: [dataverseScope(environmentUrl)],
    prompt: "select_account",
  });
  return result.account;
}

export async function signOut(): Promise<void> {
  const app = ensure();
  await app.logoutPopup();
}

/** Returns a thunk suitable for DataverseClient.getToken. */
export function makeTokenProvider(environmentUrl: string): () => Promise<string> {
  return async () => {
    const app = ensure();
    const account = (await getAccount()) ?? (await signIn(environmentUrl));
    const scopes = [dataverseScope(environmentUrl)];

    try {
      const r = await app.acquireTokenSilent({ account, scopes });
      return r.accessToken;
    } catch {
      const r = await app.acquireTokenPopup({ account, scopes });
      return r.accessToken;
    }
  };
}
