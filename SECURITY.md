# Security policy

## Reporting a vulnerability

**Please do not open a public issue.**

Report privately through GitHub Security Advisories:
<https://github.com/PowerPlatformProfessor/dvload/security/advisories/new>

You should get an acknowledgement within a few days. If a fix is needed, we'll
agree a disclosure timeline with you before publishing.

## What is in scope

dvload holds credentials for, and writes to, production Dataverse
environments. The areas where a defect has real consequences:

| Area | Why it matters |
|---|---|
| Token and secret storage (`packages/cli/src/secure-store.ts`, `auth.ts`) | Refresh tokens and client secrets are stored on disk, DPAPI-encrypted on Windows. Anything that weakens or bypasses that is in scope. |
| OData request construction (`packages/core/src/dataverse.ts`) | Batch bodies are assembled by string concatenation from spreadsheet cells. A value that can inject a header or an extra operation is a request-smuggling vulnerability. |
| The local web UI (`packages/cli/src/commands/serve.ts`) | Binds a local HTTP server. CSRF, unauthenticated access, or a path traversal in the static file handler are in scope. |
| The add-in's auth flow (`packages/addin/src/auth.ts`) | Token acquisition inside the Office host. |
| Telemetry (`TELEMETRY.md`) | Any path where environment URLs, tenant ids, record data or credentials could leave the machine. |
| The release pipeline | A way to get code into a published artifact without review. |

## Out of scope

- Findings that require an attacker to already have write access to the user's
  machine or to their Dataverse environment.
- Missing hardening headers on the local-only `serve` UI where no
  authentication boundary is crossed.
- Vulnerabilities in Dataverse or Microsoft Entra themselves — report those to
  Microsoft.
- Automated scanner output with no demonstrated impact.

## Handling secrets in contributions

The test suite never needs real credentials. If you're adding a test that
seems to require one, it belongs in the live E2E suite
(`packages/core/test/e2e/`), which reads from environment variables and
self-skips when they're absent.

Never commit: environment URLs, tenant ids, client ids, client secrets,
certificates, tokens, or exports of real records. `tests/dummy-data/` is
generated — regenerate it rather than committing anything derived from a real
environment.
