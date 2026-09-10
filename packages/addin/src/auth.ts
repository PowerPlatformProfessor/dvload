// Auth client for the dvload UI. There is no MSAL here and no Entra traffic
// from the browser at all: every token comes from the local sidecar
// (`dvload serve`) over a same-origin fetch.
//
// The previous version of this file used @azure/msal-browser against our own
// app registration, which meant every tenant needed a one-time admin consent
// before anyone could use the pane. That was never fixable in the browser:
//
//   1. A browser auth-code flow needs a redirect URI of type `spa` on the app
//      registration, and you cannot add one to a Microsoft-owned app — so the
//      pane could not borrow the pre-consented Dataverse client the CLI uses.
//   2. Entra only returns Access-Control-Allow-Origin from the token endpoint
//      when the caller's origin matches an `spa` redirect URI, so the token
//      response was unreadable from the WebView regardless of HTTP client.
//   3. Device code avoids the redirect URI but hits (2) just the same, and
//      Entra rejects `spa` redirect URIs in non-SPA flows.
//
// Moving token acquisition into a Node process sidesteps all three: it is a
// native public client, it redirects to http://localhost, and CORS does not
// apply. See docs/AUTH-NOTES.md and packages/cli/src/commands/serve.ts.
//
// The consequence to keep in mind is that the UI cannot function without the
// sidecar. `SidecarUnavailableError` exists so callers can say so plainly
// instead of surfacing a bare "Failed to fetch".

export interface Account {
  username: string;
}

export interface AuthStatus {
  /** Always "delegated": the UI never uses app-only. See serve.ts UI_TOKEN_OPTS. */
  mode: "delegated";
  account: Account | null;
  clientId: string | null;
  clientLabel: string | null;
  sharedClient: boolean;
  /**
   * App-only credentials exist for this environment, so `dvload run` and any
   * scheduled task will write as the Application User while this UI writes as
   * the signed-in user. Doesn't change what the UI does — but it means the
   * same mapping lands under two identities depending on how it's run, which
   * is worth saying out loud.
   */
  appOnlyConfigured: boolean;
}

/** Thrown when `dvload serve` isn't reachable. Distinct so the UI can advise. */
export class SidecarUnavailableError extends Error {
  constructor(cause?: unknown) {
    super(
      "Can't reach the dvload background service.\n\n" +
        "Start it from a terminal:\n" +
        "    dvload serve\n\n" +
        "It handles sign-in and Dataverse access for this pane."
    );
    this.name = "SidecarUnavailableError";
    this.cause = cause;
  }
}

/**
 * Thrown when the sidecar has no usable session for an environment.
 *
 * The sidecar could sign the user in itself — it owns the MSAL cache and can
 * open a browser — but deliberately doesn't, because the requests that hit
 * this are background ones: listing dataflows on first render, refreshing a
 * token mid-import. A browser window appearing over Excel in response to work
 * nobody asked for is worse than being told to click Sign in. So the sidecar
 * answers 401 `needsSignIn`, and this type carries that up to the UI, where
 * the existing Sign in button is the thing that opens a browser.
 */
export class SignInRequiredError extends Error {
  /** Mirrors the sidecar's flag; also survives a failed `instanceof`. */
  readonly needsSignIn = true;

  constructor(message: string) {
    super(message);
    this.name = "SignInRequiredError";
  }
}

/** True for `SignInRequiredError`, without depending on `instanceof`. */
export function isSignInRequired(e: unknown): e is SignInRequiredError {
  return e instanceof SignInRequiredError || (e as { needsSignIn?: boolean } | null)?.needsSignIn === true;
}

/**
 * Same-origin by construction: the sidecar serves this page, so a relative
 * URL always lands on it. Nothing to configure, and no CORS to negotiate.
 */
const API = "/api";

/** Populated by getAuthStatus(); read synchronously when rendering the banner. */
let lastStatus: AuthStatus | null = null;

async function post<T>(route: string, body: Record<string, unknown> = {}): Promise<T> {
  let res: Response;
  try {
    res = await fetch(`${API}${route}`, {
      method: "POST",
      headers: {
        "content-type": "application/json",
        // Not a secret — its job is to force a preflight on any cross-origin
        // attempt, which the sidecar then fails by never sending CORS
        // headers. See guard() in packages/cli/src/commands/serve.ts.
        "x-dvload-client": "1",
      },
      body: JSON.stringify(body),
    });
  } catch (e) {
    throw new SidecarUnavailableError(e);
  }

  const text = await res.text();
  let payload: unknown = null;
  try {
    payload = text ? JSON.parse(text) : null;
  } catch {
    // non-JSON error page; fall through to the status-based message
  }

  if (!res.ok) {
    const failure = payload as { error?: string; needsSignIn?: boolean } | null;
    const message = failure?.error ?? `${res.status} ${res.statusText}`;
    // Distinguished from any other failure so callers can offer the Sign in
    // button instead of reporting an error the user can do nothing with.
    if (failure?.needsSignIn) throw new SignInRequiredError(message);
    throw new Error(message);
  }
  return payload as T;
}

/**
 * Call a non-auth sidecar route (the dataflow reads) with the same guard
 * headers and error handling. Exported rather than duplicated so there is one
 * place that knows about `x-dvload-client` and how the sidecar reports errors.
 */
export function sidecarPost<T>(route: string, body: Record<string, unknown> = {}): Promise<T> {
  return post<T>(route, body);
}

/**
 * Confirm the sidecar is up before the UI starts making requests it can't
 * explain. Deliberately a cheap GET with no environment in it, so it works
 * before the user has picked one.
 */
export async function initAuth(): Promise<void> {
  // The Office WebView (WebView2 / Trident) can lose the `window` binding on
  // `fetch`, and a bare `fetch(...)` call then throws "Illegal invocation".
  // This used to be needed before MSAL initialised; MSAL is gone but every
  // request in this module is still a bare fetch, so the guard stays.
  if (typeof window !== "undefined" && window.fetch) {
    window.fetch = window.fetch.bind(window);
  }

  try {
    const res = await fetch(`${API}/status`, { headers: { "x-dvload-client": "1" } });
    if (!res.ok) throw new Error(`status ${res.status}`);
  } catch (e) {
    throw new SidecarUnavailableError(e);
  }
}

/** An environment this machine already has a session for, and whose it is. */
export interface KnownAccount {
  username: string;
  environmentUrl: string;
  host: string;
}

/**
 * Who the sidecar is already signed in as, across every environment.
 *
 * The pane's first step is "which account", which happens before an
 * environment exists — and `/account` can't answer that, because sign-in
 * state is keyed by environment. Returns an empty list rather than throwing:
 * on a first run there genuinely are no accounts, and that is the same UI as
 * a sidecar that couldn't read its store.
 */
export async function listKnownAccounts(): Promise<KnownAccount[]> {
  try {
    const { accounts } = await post<{ accounts: KnownAccount[] }>("/accounts");
    return accounts;
  } catch {
    return [];
  }
}

export async function getAuthStatus(environmentUrl: string): Promise<AuthStatus> {
  const status = await post<AuthStatus>("/account", { environmentUrl });
  lastStatus = status;
  return status;
}

export async function getAccount(environmentUrl?: string): Promise<Account | null> {
  if (!environmentUrl) return lastStatus?.account ?? null;
  return (await getAuthStatus(environmentUrl)).account;
}

/**
 * Interactive sign-in. The browser window opens from the *sidecar* process,
 * not from here, so this resolves only once the user has finished in it —
 * which can be a while. There's no popup for Office to block, and no
 * redirect back into the pane.
 *
 * `loginHint` is the username the pane already signed in as somewhere else.
 * Every environment needs its own sign-in — the token scope is the
 * environment's own origin — so passing it is what makes switching
 * environment feel like staying signed in rather than starting over.
 *
 * Without `force`, the sidecar reuses a live session for this environment
 * (and hint) instead of opening a browser — a sign-in the pane merely
 * *thinks* is needed resolves instantly. Pass `force` for the clicks whose
 * whole point is a fresh Entra prompt: "Sign in again", switching account.
 *
 * `onAuthorizeUrl` fires (at most once) with the Entra sign-in URL once the
 * sidecar has opened — or failed to open — the system browser with it. The
 * sidecar can only print that URL to its own console; this callback is what
 * lets the pane put it somewhere the user is actually looking, as a link to
 * click when no browser window appeared.
 */
export async function signIn(
  environmentUrl: string,
  loginHint?: string,
  opts: { force?: boolean; onAuthorizeUrl?: (url: string) => void } = {}
): Promise<Account> {
  const request = post<AuthStatus>("/signin", {
    environmentUrl,
    ...(loginHint ? { loginHint } : {}),
    ...(opts.force ? { force: true } : {}),
  });

  // While the sign-in request is in flight, watch for the authorize URL it
  // produces. Polling, not part of the response: /api/signin blocks until the
  // whole sign-in finishes, and the URL exists — and matters — in the middle.
  if (opts.onAuthorizeUrl) {
    const report = opts.onAuthorizeUrl;
    let settled = false;
    request.then(() => (settled = true)).catch(() => (settled = true));
    void (async () => {
      // Keeps polling after the first hit: a fallback client id opens a NEW
      // browser attempt with a new URL, and the old one's listener is dead.
      let reported: string | null = null;
      while (!settled) {
        await new Promise((r) => setTimeout(r, 1000));
        if (settled) return;
        try {
          const { url } = await post<{ url: string | null }>("/signin-url", { environmentUrl });
          if (url && url !== reported && !settled) {
            reported = url;
            report(url);
          }
        } catch {
          // Sidecar hiccup — keep waiting; the sign-in itself will report.
        }
      }
    })();
  }

  const status = await request;
  lastStatus = status;
  if (!status.account) throw new Error("Sign-in completed but no account was returned.");
  tokens.delete(environmentUrl);
  return status.account;
}

export async function signOut(environmentUrl?: string): Promise<void> {
  await post("/signout", environmentUrl ? { environmentUrl } : {});
  lastStatus = null;
  tokens.clear();
}

/**
 * Notes about *which identity* this session is using. Informational, not
 * warnings — both cases below are the intended behaviour. They're shown
 * because each one makes something appear under a name the user didn't
 * choose, and finding that out from an audit log later is worse.
 */
export function devModeBanner(): string | null {
  if (!lastStatus) return null;
  const notes: string[] = [];

  if (lastStatus.sharedClient) {
    // Borrowing the pre-consented Dataverse client is the supported default
    // and is what removes the admin-consent requirement. The cost is
    // attribution: the tenant's sign-in logs name Microsoft, not dvload.
    notes.push(
      `Signing in via ${lastStatus.clientLabel ?? "a shared Microsoft client"} — ` +
        `your tenant's sign-in logs will show that name, not dvload`
    );
  }

  if (lastStatus.appOnlyConfigured) {
    notes.push(
      "This environment also has app-only credentials, used by `dvload run` " +
        "and scheduled imports. Records created here are owned by you; " +
        "records created by a scheduled run are owned by the Application User"
    );
  }

  return notes.length ? notes.join(". ") : null;
}

export function dataverseScope(environmentUrl: string): string {
  return new URL(environmentUrl).origin + "/.default";
}

/* -------------------------------------------------------------------------- */
/* Token cache                                                                 */
/* -------------------------------------------------------------------------- */

interface CachedToken {
  accessToken: string;
  /** Epoch ms, already reduced by the safety skew. */
  goodUntil: number;
}

const tokens = new Map<string, CachedToken>();
const pending = new Map<string, Promise<string>>();

/** Renew this far ahead of expiry so a long batch never dies mid-flight. */
const EXPIRY_SKEW_MS = 5 * 60_000;
/** Used when the sidecar can't read an `exp` claim: short, but not per-call. */
const FALLBACK_TTL_MS = 60_000;

/**
 * Returns a thunk suitable for DataverseClient.getToken.
 *
 * A load issues thousands of requests, so this caches rather than crossing to
 * the sidecar every time, and de-duplicates concurrent misses — the pane's
 * first render fires several metadata calls at once, and without this they
 * would each independently ask for a token.
 */
export function makeTokenProvider(environmentUrl: string): () => Promise<string> {
  return async () => {
    const hit = tokens.get(environmentUrl);
    if (hit && hit.goodUntil > Date.now()) return hit.accessToken;

    const inFlight = pending.get(environmentUrl);
    if (inFlight) return inFlight;

    const p = (async () => {
      const { accessToken, expiresOn } = await post<{
        accessToken: string;
        expiresOn: number | null;
      }>("/token", { environmentUrl });

      const goodUntil =
        typeof expiresOn === "number"
          ? expiresOn - EXPIRY_SKEW_MS
          : Date.now() + FALLBACK_TTL_MS;
      tokens.set(environmentUrl, { accessToken, goodUntil });
      return accessToken;
    })().finally(() => pending.delete(environmentUrl));

    pending.set(environmentUrl, p);
    return p;
  };
}
