# Auth notes: what's verified, what isn't

Dataverse sign-in from a CLI runs into three independent Entra controls that
interact badly. This file records what has actually been tested, so the next
person (or the next AI) doesn't re-derive it from vibes.

## The three controls

| Control | Symptom | Escape |
|---|---|---|
| Consent — tenant restricts user consent for new apps | "Need admin approval" | Use a Microsoft-owned client that's pre-consented tenant-wide |
| Conditional Access — Authentication Flows condition | `AADSTS53003`, "…or an authentication flow that is restricted by your admin" | Use the browser flow instead of device code |
| Redirect URI registration | `AADSTS900971` "No reply address provided", or `AADSTS50011` mismatch | Use an app that registers `http://localhost` |

Good news: the first two escapes compose. The shared Microsoft client is both
pre-consented *and* loopback-capable, so browser sign-in through it needs
neither admin consent nor device code flow. See "Resolved" below — an earlier
version of this file claimed otherwise, incorrectly.

## Verified

- **`51f81489-12ee-4a9e-aaae-a2591f45987d`** is Microsoft's documented sample
  Dataverse/XRM Tooling client. Its Dataverse delegated permission is
  pre-consented in every tenant, so device-code sign-in through it never
  triggers admin approval. Documented sample redirect URI is
  `app://58145B91-0C36-4500-8554-080854F2AC97`.
- **Conditional Access can block device code flow specifically**, via
  Conditions -> Authentication Flows. Microsoft recommends getting "as close
  as possible to a unilateral block" on it, because an attacker can generate
  a code and phish a victim into entering it. This is why device code fails in
  hardened tenants.
- **Entra validates redirect URIs only after authentication.** An
  unregistered `redirect_uri` still returns the normal sign-in page; the error
  appears once credentials are submitted. Confirmed by requesting
  `/authorize` with a deliberately bogus redirect URI — sign-in page, no error.
  Consequence: a bad redirect URI is invisible to MSAL and to us. The loopback
  listener just never fires.
- **msal-node 2.16.3 does send a real redirect URI** for
  `acquireTokenInteractive`. Read from
  `dist/client/PublicClientApplication.mjs`: it awaits
  `waitForRedirectUri(loopbackClient)` and passes the result as
  `redirectUri`. So "MSAL sent an empty `redirect_uri`" is *not* a possible
  cause of `AADSTS900971` here.
- **MSAL.NET no longer accepts `app://` redirects on desktop.** It errors with
  `loopback_redirect_uri`: "Only loopback redirect uri is supported, but
  `app://…` was found. Configure `http://localhost` or `http://localhost:port`
  both during app registration and when you create the
  PublicClientApplication object."
  (microsoft/PowerPlatform-DataverseServiceClient discussion #456)
- **The add-in cannot borrow a client id *from the browser*.** A browser
  auth-code flow needs a redirect URI of type `spa` on the app registration —
  both for the redirect and for the token endpoint's
  `Access-Control-Allow-Origin` — and you can't add one to a Microsoft-owned
  app. Entra also rejects `spa` redirect URIs in non-SPA flows, so
  device-code-from-the-taskpane is closed too. Not a library problem;
  swapping MSAL.js changes nothing.

  Still true, and still worth knowing — but no longer a limitation of the
  product, because the pane stopped authenticating. See "Resolved: the pane
  doesn't authenticate" below.

## Resolved: the pane doesn't authenticate

Every constraint above is about *where the token request originates*. A
native client redirects to `http://localhost` and posts to the token endpoint
from a socket, so CORS never applies. A WebView cannot do either. That
framing makes the fix obvious in hindsight: stop asking the WebView to
authenticate.

`dvload serve` runs a loopback HTTP server that serves the pane's own bundle
and exposes `POST /api/token`, backed by the same `getTokenProvider` the CLI
uses. The pane fetches a token from its own origin. There is no MSAL in the
browser bundle, no app registration for the UI, no `spa` redirect URI, and no
admin consent.

What this bought, beyond the consent fix:

- One auth implementation instead of two (msal-node and msal-browser were
  both being maintained, with different failure modes).
- No public web host for the UI, so no deployment to keep in sync with the
  manifest's `SourceLocation`. Dev and production are the same setup.
- The same UI runs in an ordinary browser (`dvload gui`), because nothing in
  it depends on Office except reading the open workbook.

What it cost:

- The pane is dead without the sidecar running. Mitigated with an explicit
  "dvload isn't running" screen rather than a blank pane; a logon-triggered
  Scheduled Task is the obvious next step.
- Office requires HTTPS even on loopback, so a trusted localhost certificate
  is now a prerequisite. `office-addin-dev-certs` installs one into the
  **CurrentUser** store, which keeps the whole install admin-free — the point
  of the exercise.
- The port is fixed (44321), because `SourceLocation` is a literal URL in the
  manifest. A conflict is a hard failure with a clear message.

Threat model, since this process holds tokens and listens on a predictable
port: bind `127.0.0.1`; check the `Host` header (DNS rebinding is the attack
that defeats origin checks); require POST plus a custom header on API routes,
which forces a preflight; never send CORS headers, so a cross-origin caller
cannot read a response even if it gets one. Tested in
`packages/cli/src/commands/serve.test.ts`.

## Resolved: the shared client DOES support loopback

Tested by hand against a real tenant:

```
https://login.microsoftonline.com/organizations/oauth2/v2.0/authorize
  ?client_id=51f81489-12ee-4a9e-aaae-a2591f45987d
  &response_type=code
  &redirect_uri=http%3A%2F%2Flocalhost%3A53682
  &response_mode=query
  &scope=https%3A%2F%2FYOURORG.crm.dynamics.com%2F.default
```

(join the lines; replace `YOURORG`)

Result: sign-in completed and the browser landed on
`http://localhost:53682/?code=1.Aa4A…` with `ERR_CONNECTION_REFUSED`, which is
the expected outcome when nothing is listening. **`51f81489…` accepts
`http://localhost:<port>` as a public-client redirect URI.**

So the pre-consented Microsoft client and the browser flow *can* be combined,
which is presumably how XrmToolBox connects without admin consent: shared
client id, interactive loopback flow, no consent prompt, no device code for
Conditional Access to block.

### What the AADSTS900971 actually was

`DATAVERSE_LOAD_CLIENT_ID` was set in the user's environment, left over from
following the README's `setx` instruction. `resolveClientIdChain` treats a
pinned client id as "disable the chain", so `51f81489…` was never attempted —
every request went to dvload's own app `e6828b0f…`, which has no
`http://localhost` redirect URI registered. `AADSTS900971` means exactly that:
no reply address on *that* app.

The error was correct and specific the whole time. It took four wrong
hypotheses (client capability, browser launcher, loopback port, platform type)
because the CLI printed the sign-in URL but never printed **which client id it
was using or why**, and the client id was the one thing that mattered.

**Diagnostic rule:** before theorising about any Entra error, read the
authorize URL's `client_id` and confirm it's the app you think it is. dvload
now announces the client id, the flow, and which env var pinned it on every
login attempt (`announceLoginAttempt`) — unconditionally, not debug-gated.

Corollary worth remembering: an env var that silently removes a fallback path
turns a recoverable failure into an unexplainable one. If a setting disables
a retry chain, say so out loud at the point of use.

### It came back, with no env var this time

Same `AADSTS900971`, same app `e6828b0f…`, but `DATAVERSE_LOAD_CLIENT_ID` was
unset in every scope. The chain was working exactly as designed: the shared
client got the first browser window, that window timed out, `interactive_timeout`
advanced the chain, and window two opened on dvload's own app — which cannot
complete a browser sign-in. The user, already holding a valid session from an
earlier successful sign-in, saw an error page and a connected pane at once.

The preflight was supposed to make that second window impossible, and it was
silently failing open. `probeLoopbackRedirect` looked for AADSTS50011/500113/900971
in the response body — the codes that mean "bad reply address". Entra never
sends them here: it validates the redirect URI **before** it evaluates
`prompt=none`, so the page it renders complains about the missing session
(`AADSTS50058`) and never mentions the redirect. No match, verdict
`inconclusive`, browser opens.

Measured against the live tenant, holding client id and tenant fixed and
varying only `redirect_uri`:

| client | redirect_uri | Entra |
|---|---|---|
| `51f81489…` | `http://localhost:53682` | 302 → localhost |
| `51f81489…` | `http://localhost:61610` | 302 → localhost |
| `51f81489…` | `https://not-registered.example/cb` | 200 page, AADSTS50058 |
| `e6828b0f…` | `http://localhost:53682` | 200 page, AADSTS50058 |
| `e6828b0f…` | `http://localhost:61610` | 200 page, AADSTS50058 |

Row 3 is the control: same client that 302s for a good URI renders for a bad
one. So the discriminator is the **shape of the response**, not the code in it
— Entra redirects an error only to a target it has validated, and renders when
it has nowhere safe to send it. The probe now reads shape, and treats a body
with no AADSTS code at all as inconclusive so a proxy block page cannot
condemn a working app.

**Diagnostic rule, second edition:** when a check that exists to prevent a
failure doesn't fire, test the check against a known-bad input before trusting
it. This one had been returning "inconclusive" for every app it was pointed
at, which is indistinguishable from "not consulted" and equally useless.

### Windows browser launching

Unrelated to the above, but correct regardless: launch URLs with
`explorer.exe <url>`. Not `cmd /c start` (splits on `&`, and Entra URLs are
nothing but `&`), not `rundll32 url.dll,FileProtocolHandler` (mangles long
parameterised URLs), not `powershell Start-Process` (execution policy can
block it in exactly the hardened tenants that force the browser flow).

## Current behaviour

- `dvload login` defaults to the browser flow, falling back to device code
  only when no local browser exists (SSH, no `DISPLAY`).
- The client-id chain starts shared Microsoft client, then dvload's own app.
  For a **browser** sign-in the first entry is always attempted and a later
  one joins only if `probeLoopbackRedirect` confirms it can receive the
  redirect — which today leaves the shared client alone in the default
  interactive chain. Device code needs no redirect URI, so both stay live
  there.
- Consequence, and the intended one: unless you set `DATAVERSE_LOAD_CLIENT_ID`
  or `--client-id`, a browser sign-in uses the shared Microsoft client and
  nothing else. No consent prompt, no second window, nothing to explain.
- Chain advancement only happens on errors MSAL actually receives — i.e.
  token-endpoint errors. `/authorize`-time errors surface as a browser page
  and a 3-minute timeout, not as a retry.

## Who is allowed to open a browser

Escalating from "no valid token" to "sign in" is a decision about the *user's
attention*, so it belongs to whoever the user is currently talking to.

- **The CLI escalates.** `dvload run`, `dvload dataflows`, `dvload login`:
  someone typed a command and is watching a terminal, so prompting is the
  whole point. `--non-interactive` opts out.
- **The sidecar never escalates.** Every provider built in
  `commands/serve.ts` is `silentOnly`. The requests reaching those routes are
  background work the pane issued by itself — listing dataflows on first
  render, refreshing a token mid-import — and a browser window appearing over
  Excel in response to that is not something anyone asked for. Worse, the
  HTTP response stays blocked until sign-in completes or
  `DVLOAD_AUTH_TIMEOUT_MS` (3 min) expires.

A dead session therefore becomes `InteractiveSignInRequiredError` in
`auth.ts`, which the sidecar translates to **401 with `{ needsSignIn: true }`**.
The pane raises that as `SignInRequiredError` and points at its Sign in
button, which calls `POST /api/signin` — the one route that does open a
browser, because a click is what got it there.

This composes with the earlier rule about *what* may escalate.
`needsInteractiveSignIn()` decides whether a failed refresh is a dead session
at all — only MSAL's `InteractionRequiredAuthError` counts, so a network blip
leaves as `refreshFailed` and stays a 500. `InteractiveSignInRequiredError`
then decides who acts on the ones that are genuine: the CLI prompts, the
sidecar reports. Both halves are exercised in
`packages/cli/src/auth-refresh.test.ts`.
