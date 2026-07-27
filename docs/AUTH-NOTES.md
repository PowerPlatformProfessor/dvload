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
- **The add-in cannot borrow a client id at all.** A browser auth-code flow
  needs a redirect URI of type `spa` on the app registration — both for the
  redirect and for the token endpoint's `Access-Control-Allow-Origin` — and
  you can't add one to a Microsoft-owned app. Entra also rejects `spa`
  redirect URIs in non-SPA flows, so device-code-from-the-taskpane is closed
  too. Not a library problem; swapping MSAL.js changes nothing.

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

### Windows browser launching

Unrelated to the above, but correct regardless: launch URLs with
`explorer.exe <url>`. Not `cmd /c start` (splits on `&`, and Entra URLs are
nothing but `&`), not `rundll32 url.dll,FileProtocolHandler` (mangles long
parameterised URLs), not `powershell Start-Process` (execution policy can
block it in exactly the hardened tenants that force the browser flow).

## Current behaviour

- `dvload login` defaults to the browser flow, falling back to device code
  only when no local browser exists (SSH, no `DISPLAY`).
- The client-id chain is flow-agnostic: shared Microsoft client first, then
  dvload's own app. Trying the shared client costs one clear error if it
  can't do loopback, and saves an admin-consent round trip if it can.
- Chain advancement only happens on errors MSAL actually receives — i.e.
  token-endpoint errors. `/authorize`-time errors surface as a browser page
  and a 3-minute timeout, not as a retry.
