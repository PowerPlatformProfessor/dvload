// Thin OData Web API client for Microsoft Dataverse. Designed to work in both Node
// and the browser (Office.js add-in). Auth is delegated: callers provide a
// `getToken()` thunk so this module doesn't pin a specific MSAL flow.

export interface DataverseClientOptions {
  /** Environment URL like "https://contoso.crm.dynamics.com" (no trailing slash). */
  environmentUrl: string;
  /** Returns a Bearer token valid for the Dataverse scope. */
  getToken: () => Promise<string>;
  /** Optional fetch override (for testing or non-DOM Node < 18). */
  fetch?: typeof fetch;
  /** API version. Default "v9.2". */
  apiVersion?: string;
  /** Retry policy for throttle / transient failures. */
  retry?: RetryOptions;
  /** Called after every HTTP response (including retried ones). */
  onRequest?: (entry: RequestLogEntry) => void;
}

/** One HTTP request/response pair recorded by the client. */
export interface RequestLogEntry {
  method: string;
  url: string;
  status: number;
  ok: boolean;
  /** ISO 8601 timestamp — when the request was sent. */
  startedAt: string;
  /** ISO 8601 timestamp — when the response was received. */
  finishedAt: string;
  /** Response body for non-ok responses (parsed as JSON if possible, otherwise raw text). */
  errorBody?: unknown;
}

/** Retry policy for 429/503/504 responses. */
export interface RetryOptions {
  /** Max attempts INCLUDING the first call. Default 5. */
  maxAttempts?: number;
  /** Initial backoff in ms when Retry-After is absent. Default 500. */
  baseDelayMs?: number;
  /** Hard cap on any single sleep. Default 60_000 (60 seconds). */
  maxDelayMs?: number;
  /** Status codes that trigger a retry. Default [429, 503, 504]. */
  retryableStatuses?: number[];
  /** Called every time we sleep and retry. Useful for logging. */
  onRetry?: (info: RetryInfo) => void;
}

export interface RetryInfo {
  attempt: number; // 1-based; the attempt that just failed
  status: number;
  delayMs: number;
  url: string;
  /** Source of the delay: "retry-after-seconds", "retry-after-date", or "backoff". */
  source: "retry-after-seconds" | "retry-after-date" | "backoff";
}

export interface CreateResult {
  /** GUID of the created record (read from the OData-EntityId header). */
  id: string;
  /** Full entity URL returned by the server. */
  entityUrl: string;
}

export interface BatchOperation {
  /** Logical operation name; surfaced in error messages. */
  contentId: number;
  method: "POST" | "PATCH" | "DELETE";
  /** Entity-set-relative URL, e.g. "contacts" or "contacts(<guid>)". */
  url: string;
  body?: unknown;
  /** Extra request headers (e.g. If-Match for upsert). */
  headers?: Record<string, string>;
}

export interface BatchResultItem {
  contentId: number;
  status: number;
  ok: boolean;
  /** Created entity GUID for POSTs (parsed from OData-EntityId). */
  id?: string;
  body?: unknown;
  errorMessage?: string;
}

const DEFAULT_HEADERS: Record<string, string> = {
  "OData-MaxVersion": "4.0",
  "OData-Version": "4.0",
  Accept: "application/json",
  "Content-Type": "application/json; charset=utf-8",
  Prefer: 'return=representation,odata.include-annotations="*"',
};

export class DataverseError extends Error {
  constructor(public status: number, public code: string | undefined, message: string, public raw?: unknown) {
    super(message);
    this.name = "DataverseError";
  }
}

/* -------------------------------------------------------------------------- */
/* Input validation                                                            */
/*                                                                             */
/* Attribute / entity names and key values flow in from mapping files and     */
/* spreadsheet cells — untrusted input. Because $batch bodies are built by    */
/* string concatenation (multipart/mixed with embedded request lines), any    */
/* unvalidated value is a request-smuggling vector: a cell containing CRLF    */
/* could inject headers or whole extra operations into the changeset.        */
/* -------------------------------------------------------------------------- */

const LOGICAL_NAME_RE = /^[A-Za-z_][A-Za-z0-9_]*$/;
const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/** Validate a Dataverse logical name (entity, entity set, or attribute). */
export function assertLogicalName(name: string, what: string): string {
  if (!LOGICAL_NAME_RE.test(name)) {
    throw new Error(`${what} "${name}" is not a valid Dataverse logical name`);
  }
  return name;
}

/** Validate a GUID (as used in entity keys and @odata.bind paths). */
export function assertGuid(value: string, what: string): string {
  if (!GUID_RE.test(value)) {
    throw new Error(`${what} "${value}" is not a valid GUID`);
  }
  return value;
}

/**
 * Format a value as an OData key literal, safe for URL embedding: quotes are
 * doubled per OData rules, then the value is percent-encoded so it cannot
 * carry CRLF, parens, or other structure into a request line.
 */
export function formatKeyLiteral(value: unknown): string {
  if (typeof value === "number") {
    if (!Number.isFinite(value)) throw new Error(`Key value ${value} is not a finite number`);
    return String(value);
  }
  if (typeof value === "boolean") return value ? "true" : "false";
  // encodeURIComponent leaves ' ( ) ! * ~ raw; escape those too so the
  // encoded literal contains no characters with structural meaning in an
  // OData key expression or an HTTP request line.
  const encoded = encodeURIComponent(String(value).replace(/'/g, "''")).replace(
    /[()'!*~]/g,
    (c) => "%" + c.charCodeAt(0).toString(16).toUpperCase().padStart(2, "0")
  );
  return `'${encoded}'`;
}

/** Last line of defense: no entity-set-relative URL may contain CR/LF or spaces. */
function assertSafeBatchUrl(url: string): string {
  if (/[\r\n\s]/.test(url)) {
    throw new Error(`Refusing to build batch request: URL contains whitespace/CRLF: ${JSON.stringify(url)}`);
  }
  return url;
}

export class DataverseClient {
  private readonly base: string;
  private readonly fetchFn: typeof fetch;

  constructor(private readonly opts: DataverseClientOptions) {
    const apiVersion = opts.apiVersion ?? "v9.2";
    this.base = `${opts.environmentUrl.replace(/\/+$/, "")}/api/data/${apiVersion}`;
    this.fetchFn = opts.fetch ?? fetch;
  }

  /** Build full URL for an entity-set-relative path. */
  url(path: string): string {
    return `${this.base}/${path.replace(/^\/+/, "")}`;
  }

  private async authHeaders(): Promise<Record<string, string>> {
    const token = await this.opts.getToken();
    return { Authorization: `Bearer ${token}` };
  }

  /**
   * Fetch with throttle-aware retry. Honors the Retry-After header (either
   * seconds or HTTP-date), falls back to exponential backoff with a hard
   * cap. Only retries on the configured status codes (default 429/503/504);
   * other failures are returned to the caller verbatim.
   */
  private async fetchWithRetry(input: string | URL, init?: RequestInit): Promise<Response> {
    const policy = this.opts.retry ?? {};
    const maxAttempts = policy.maxAttempts ?? 5;
    const baseDelay = policy.baseDelayMs ?? 500;
    const maxDelay = policy.maxDelayMs ?? 60_000;
    const retryable = new Set(policy.retryableStatuses ?? [429, 503, 504]);

    const url = typeof input === "string" ? input : (input as URL).toString();
    const method = ((init?.method) ?? "GET").toUpperCase();
    let lastResponse: Response | undefined;
    for (let attempt = 1; attempt <= maxAttempts; attempt++) {
      const startedAt = new Date().toISOString();
      const res = await this.fetchFn(input as unknown as Parameters<typeof fetch>[0], init);
      const finishedAt = new Date().toISOString();

      // For non-ok responses that won't be retried, capture the body before
      // returning it to the caller (clone so the caller can still read it).
      let errorBody: unknown;
      if (!res.ok && !retryable.has(res.status) && this.opts.onRequest) {
        try {
          const text = await res.clone().text();
          if (text) {
            try { errorBody = JSON.parse(text); } catch { errorBody = text; }
          }
        } catch { /* ignore */ }
      }

      this.opts.onRequest?.({ method, url, status: res.status, ok: res.ok, startedAt, finishedAt, ...(errorBody !== undefined ? { errorBody } : {}) });
      if (!retryable.has(res.status)) return res;
      lastResponse = res;
      if (attempt === maxAttempts) return res;

      const retryAfter = res.headers.get("Retry-After");
      const parsed = parseRetryAfter(retryAfter);
      const delayMs = parsed
        ? parsed.value
        : Math.min(baseDelay * 2 ** (attempt - 1), maxDelay);

      policy.onRetry?.({
        attempt,
        status: res.status,
        delayMs,
        url: typeof input === "string" ? input : (input as URL).toString(),
        source: parsed?.source ?? "backoff",
      });

      // Drain the body so we don't leak the underlying connection.
      try {
        await res.arrayBuffer();
      } catch {
        /* ignore */
      }
      await sleep(delayMs);
    }
    return lastResponse!;
  }

  /** Fetch entity metadata for one entity (LogicalName -> definition). */
  async getEntityDefinition(logicalName: string): Promise<Record<string, unknown>> {
    assertLogicalName(logicalName, "Entity logical name");
    const res = await this.fetchWithRetry(
      this.url(`EntityDefinitions(LogicalName='${logicalName}')?$expand=Attributes`),
      { headers: { ...DEFAULT_HEADERS, ...(await this.authHeaders()) } }
    );
    if (!res.ok) await throwForResponse(res);
    return (await res.json()) as Record<string, unknown>;
  }

  /** List entity definitions, returning [{ LogicalName, EntitySetName, DisplayName }, ...]. */
  async listEntities(): Promise<Array<{ LogicalName: string; EntitySetName: string; DisplayName: string }>> {
    const res = await this.fetchWithRetry(
      this.url("EntityDefinitions?$select=LogicalName,EntitySetName,DisplayName"),
      { headers: { ...DEFAULT_HEADERS, ...(await this.authHeaders()) } }
    );
    if (!res.ok) await throwForResponse(res);
    const json = (await res.json()) as { value: Array<Record<string, unknown>> };
    return json.value.map((e) => ({
      LogicalName: String(e.LogicalName),
      EntitySetName: String(e.EntitySetName),
      DisplayName: extractLocalizedLabel(e.DisplayName),
    }));
  }

  /**
   * Return the logical names of entities a lookup attribute can target.
   * Uses the LookupAttributeMetadata cast on the Attributes navigation property.
   */
  async getLookupTargets(entityLogicalName: string, attrLogicalName: string): Promise<string[]> {
    assertLogicalName(entityLogicalName, "Entity logical name");
    assertLogicalName(attrLogicalName, "Attribute logical name");
    const res = await this.fetchWithRetry(
      this.url(
        `EntityDefinitions(LogicalName='${entityLogicalName}')/Attributes` +
        `/Microsoft.Dynamics.CRM.LookupAttributeMetadata` +
        `?$select=LogicalName,Targets&$filter=LogicalName eq '${attrLogicalName}'`
      ),
      { headers: { ...DEFAULT_HEADERS, ...(await this.authHeaders()) } }
    );
    if (!res.ok) await throwForResponse(res);
    const json = (await res.json()) as { value?: Array<{ Targets?: string[] }> };
    return json.value?.[0]?.Targets ?? [];
  }

  /** Create a single record. Slow for many rows — prefer batch(). */
  async create(entitySet: string, body: Record<string, unknown>): Promise<CreateResult> {
    const res = await this.fetchWithRetry(this.url(entitySet), {
      method: "POST",
      headers: { ...DEFAULT_HEADERS, ...(await this.authHeaders()) },
      body: JSON.stringify(body),
    });
    if (!res.ok) await throwForResponse(res);
    return parseCreateResult(res);
  }

  /** Resolve a record id by alternate key. Returns null if not found. */
  async resolveByKey(entitySet: string, keyAttribute: string, value: unknown): Promise<string | null> {
    assertLogicalName(entitySet, "Entity set");
    assertLogicalName(keyAttribute, "Key attribute");
    const literal = formatKeyLiteral(value);
    // We don't know the primary-id field name (and it doesn't always pluralize
    // cleanly — e.g. "people" → "systemuserid"). Read @odata.id instead,
    // which is the canonical URL containing the GUID.
    const path = `${entitySet}(${keyAttribute}=${literal})?$select=createdon`;
    const res = await this.fetchWithRetry(this.url(path), {
      headers: { ...DEFAULT_HEADERS, ...(await this.authHeaders()) },
    });
    if (res.status === 404) return null;
    if (!res.ok) await throwForResponse(res);
    const json = (await res.json()) as Record<string, unknown>;
    const odataId = String(json["@odata.id"] ?? "");
    const m = /\(([0-9a-f-]{36})\)/i.exec(odataId);
    return m?.[1] ?? null;
  }

  /**
   * Execute a $batch where EACH operation is wrapped in its own changeset.
   * Operations therefore execute independently: one row failing does not
   * roll back its neighbors, and per-row success/failure accounting is
   * accurate. (A single shared changeset is atomic — Dataverse aborts it on
   * the first failure and rolls everything back, which silently invalidated
   * the "created/updated" counts of every other op in the batch, and made
   * skip-if-exists abort the batch on the first existing record.)
   * Dataverse caps batches at 1000 operations.
   */
  async batch(operations: BatchOperation[]): Promise<BatchResultItem[]> {
    if (operations.length === 0) return [];
    const batchId = `batch_${cryptoUuid()}`;

    const body = buildBatchBody(operations, batchId, this.base);

    const res = await this.fetchWithRetry(this.url("$batch"), {
      method: "POST",
      headers: {
        ...(await this.authHeaders()),
        "OData-MaxVersion": "4.0",
        "OData-Version": "4.0",
        Accept: "application/json",
        "Content-Type": `multipart/mixed; boundary=${batchId}`,
      },
      body,
    });

    if (!res.ok) await throwForResponse(res);
    const text = await res.text();
    return parseBatchResponse(text, operations);
  }
}

/* -------------------------------------------------------------------------- */
/* Helpers                                                                     */
/* -------------------------------------------------------------------------- */

/** Parse the Retry-After header per RFC 7231: integer seconds OR HTTP-date. */
function parseRetryAfter(
  header: string | null
): { value: number; source: "retry-after-seconds" | "retry-after-date" } | null {
  if (!header) return null;
  const trimmed = header.trim();
  // Integer seconds form.
  if (/^\d+(\.\d+)?$/.test(trimmed)) {
    const seconds = Number(trimmed);
    return { value: Math.max(0, seconds * 1000), source: "retry-after-seconds" };
  }
  // HTTP-date form.
  const epoch = Date.parse(trimmed);
  if (!Number.isNaN(epoch)) {
    return { value: Math.max(0, epoch - Date.now()), source: "retry-after-date" };
  }
  return null;
}

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function throwForResponse(res: Response): Promise<never> {
  let raw: unknown;
  let message = `${res.status} ${res.statusText}`;
  let code: string | undefined;
  try {
    raw = await res.json();
    const err = (raw as { error?: { code?: string; message?: string } }).error;
    if (err?.message) message = err.message;
    if (err?.code) code = err.code;
  } catch {
    // body wasn't JSON
  }
  throw new DataverseError(res.status, code, message, raw);
}

function parseCreateResult(res: Response): CreateResult {
  const entityUrl = res.headers.get("OData-EntityId") ?? "";
  const m = /\(([0-9a-f-]{36})\)/i.exec(entityUrl);
  return { id: m?.[1] ?? "", entityUrl };
}

function extractLocalizedLabel(displayName: unknown): string {
  if (!displayName || typeof displayName !== "object") return "";
  const dn = displayName as { UserLocalizedLabel?: { Label?: string } };
  return dn.UserLocalizedLabel?.Label ?? "";
}

function cryptoUuid(): string {
  // Cross-runtime UUID without bringing in a dependency.
  if (typeof crypto !== "undefined" && "randomUUID" in crypto) {
    return crypto.randomUUID();
  }
  // Fallback: not crypto-strong but fine for batch boundaries.
  return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, (c) => {
    const r = (Math.random() * 16) | 0;
    const v = c === "x" ? r : (r & 0x3) | 0x8;
    return v.toString(16);
  });
}

function buildBatchBody(ops: BatchOperation[], batchId: string, base: string): string {
  const lines: string[] = [];

  for (const op of ops) {
    // One changeset per operation → independent execution (see batch()).
    const changesetId = `changeset_${op.contentId}_${cryptoUuid()}`;
    const url = assertSafeBatchUrl(op.url);

    lines.push(`--${batchId}`);
    lines.push(`Content-Type: multipart/mixed; boundary=${changesetId}`);
    lines.push("");
    lines.push(`--${changesetId}`);
    lines.push("Content-Type: application/http");
    lines.push("Content-Transfer-Encoding: binary");
    lines.push(`Content-ID: ${op.contentId}`);
    lines.push("");
    lines.push(`${op.method} ${base}/${url} HTTP/1.1`);
    lines.push("Content-Type: application/json; charset=utf-8");
    lines.push("OData-MaxVersion: 4.0");
    lines.push("OData-Version: 4.0");
    lines.push('Prefer: return=representation,odata.include-annotations="*"');
    if (op.headers) {
      for (const [k, v] of Object.entries(op.headers)) {
        if (/[\r\n]/.test(k) || /[\r\n]/.test(v)) {
          throw new Error(`Refusing to build batch request: header contains CRLF (${k})`);
        }
        lines.push(`${k}: ${v}`);
      }
    }
    lines.push("");
    lines.push(op.body !== undefined ? JSON.stringify(op.body) : "");
    lines.push(`--${changesetId}--`);
  }

  lines.push(`--${batchId}--`);
  lines.push("");
  return lines.join("\r\n");
}

/**
 * Parse a multipart/mixed batch response. We map each "Content-ID" back to
 * the originating BatchOperation so callers can correlate failures with rows.
 *
 * The response now contains one changesetresponse per operation (each with
 * its own server-generated boundary), so instead of extracting a single
 * boundary we split on every boundary line (lines starting with "--") and
 * keep the parts that contain an HTTP status line. Response bodies are
 * single-line JSON, so no body line can start with "--".
 */
function parseBatchResponse(text: string, ops: BatchOperation[]): BatchResultItem[] {
  const results: BatchResultItem[] = [];
  const parts = text
    .split(/^--[^\r\n]*/m) // every boundary line, batch- and changeset-level
    .filter((p) => /HTTP\/1\.1/.test(p));

  for (const part of parts) {
    const cid = /Content-ID:\s*(\d+)/i.exec(part)?.[1];
    const statusMatch = /HTTP\/1\.1\s+(\d{3})/.exec(part);
    if (!cid || !statusMatch) continue;
    const contentId = Number(cid);
    const status = Number(statusMatch[1]);
    const ok = status >= 200 && status < 300;
    const op = ops.find((o) => o.contentId === contentId);
    if (!op) continue;

    let body: unknown;
    let errorMessage: string | undefined;
    let id: string | undefined;

    const bodyStart = part.indexOf("\r\n\r\n", part.indexOf("HTTP/1.1"));
    const rawBody = bodyStart >= 0 ? part.slice(bodyStart + 4).trim() : "";
    if (rawBody) {
      try {
        body = JSON.parse(rawBody);
        if (!ok) {
          const err = (body as { error?: { message?: string } }).error;
          errorMessage = err?.message ?? `HTTP ${status}`;
        } else if (op.method === "POST") {
          id = (body as { [k: string]: unknown })[`${op.url.split("(")[0].replace(/s$/, "")}id`] as string | undefined;
        }
      } catch {
        if (!ok) errorMessage = rawBody;
      }
    } else if (!ok) {
      errorMessage = `HTTP ${status}`;
    }

    // Also try to read OData-EntityId for created records when body is empty.
    if (op.method === "POST" && !id) {
      const eid = /OData-EntityId:\s*[^\r\n]*\(([0-9a-f-]{36})\)/i.exec(part)?.[1];
      if (eid) id = eid;
    }

    results.push({ contentId, status, ok, body, errorMessage, id });
  }

  // Preserve original order.
  results.sort((a, b) => a.contentId - b.contentId);
  return results;
}
