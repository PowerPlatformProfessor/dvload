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

/** Retry policy for throttled/unavailable responses and dropped connections. */
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
  /** HTTP status, or 0 when the request never produced a response. */
  status: number;
  delayMs: number;
  url: string;
  /** Source of the delay: "retry-after-seconds", "retry-after-date", or "backoff". */
  source: "retry-after-seconds" | "retry-after-date" | "backoff";
  /** Set when the attempt threw instead of responding (e.g. "fetch failed (ECONNRESET)"). */
  error?: string;
}

/**
 * How a thrown fetch error should be treated:
 *  - "abort": the caller cancelled. Never retry.
 *  - "not-executed": the connection was never established (DNS, refused,
 *    connect timeout), so the server cannot have run anything. Safe to
 *    replay even for non-idempotent POSTs.
 *  - "ambiguous": the connection died mid-flight. The server may or may not
 *    have executed the request, so only replay it when the caller says the
 *    operation is idempotent.
 */
type NetworkFailureKind = "abort" | "not-executed" | "ambiguous";

/** Node/undici error codes meaning we never reached the server. */
const NOT_EXECUTED_CODES = new Set([
  "ENOTFOUND",
  "EAI_AGAIN",
  "ECONNREFUSED",
  "UND_ERR_CONNECT_TIMEOUT",
  "ERR_SOCKET_CONNECTION_TIMEOUT",
]);

/** Walk the cause chain — undici wraps the real cause inside "fetch failed". */
function errorCodes(e: unknown): string[] {
  const out: string[] = [];
  let cur: unknown = e;
  for (let depth = 0; cur instanceof Error && depth < 5; depth++) {
    const code = (cur as { code?: unknown }).code;
    if (typeof code === "string") out.push(code);
    if (cur.name) out.push(cur.name);
    cur = (cur as { cause?: unknown }).cause;
  }
  return out;
}

export function classifyNetworkError(e: unknown): NetworkFailureKind {
  const codes = errorCodes(e);
  if (codes.includes("AbortError") || codes.includes("ABORT_ERR")) return "abort";
  if (codes.some((c) => NOT_EXECUTED_CODES.has(c))) return "not-executed";
  return "ambiguous";
}

/** One-line description of a thrown fetch error, including the underlying code. */
export function describeNetworkError(e: unknown): string {
  const message = e instanceof Error ? e.message : String(e);
  const codes = errorCodes(e).filter((c) => c !== "Error" && c !== "TypeError");
  return codes.length > 0 ? `${message} (${codes[0]})` : message;
}

export interface CreateResult {
  /** GUID of the created record. Empty only if the server returned neither a header nor a parseable body. */
  id: string;
  /** Full entity URL, when the server sent one. */
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

/** Shape of one option in Dataverse option-set metadata. */
interface OptionMetadata {
  Value?: number;
  Label?: { UserLocalizedLabel?: { Label?: string } };
}

export class DataverseClient {
  private readonly base: string;
  private readonly fetchFn: typeof fetch;
  private readonly entitySetInfoCache = new Map<
    string,
    { logicalName: string; primaryIdAttribute: string; primaryNameAttribute?: string }
  >();

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
  private async fetchWithRetry(
    input: string | URL,
    init?: RequestInit,
    retryableOverride?: number[],
    /**
     * Whether replaying this request is safe when a connection dies mid-flight.
     * Defaults to true: every caller except plain POST creates is idempotent.
     */
    idempotent = true
  ): Promise<Response> {
    const policy = this.opts.retry ?? {};
    const maxAttempts = policy.maxAttempts ?? 5;
    const baseDelay = policy.baseDelayMs ?? 500;
    const maxDelay = policy.maxDelayMs ?? 60_000;
    const retryable = new Set(retryableOverride ?? policy.retryableStatuses ?? [429, 503, 504]);

    const url = typeof input === "string" ? input : (input as URL).toString();
    const method = ((init?.method) ?? "GET").toUpperCase();
    let lastResponse: Response | undefined;
    for (let attempt = 1; attempt <= maxAttempts; attempt++) {
      const startedAt = new Date().toISOString();

      // A dropped connection THROWS rather than returning a status — a laptop
      // sleeping mid-run, a VPN reconnecting, a flaky link. Without this the
      // error escapes the whole retry loop and loses an entire batch.
      let res: Response;
      try {
        res = await this.fetchFn(input as unknown as Parameters<typeof fetch>[0], init);
      } catch (e) {
        const kind = classifyNetworkError(e);
        const reason = describeNetworkError(e);
        this.opts.onRequest?.({
          method, url, status: 0, ok: false,
          startedAt, finishedAt: new Date().toISOString(),
          errorBody: reason,
        });
        // Cancellation is deliberate; a possibly-executed POST must not be
        // replayed or it duplicates records.
        const replayable = kind === "not-executed" || (kind === "ambiguous" && idempotent);
        if (kind === "abort" || !replayable || attempt === maxAttempts) throw e;

        const delayMs = Math.min(baseDelay * 2 ** (attempt - 1), maxDelay);
        policy.onRetry?.({ attempt, status: 0, delayMs, url, source: "backoff", error: reason });
        await sleep(delayMs);
        continue;
      }
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
      // maxDelayMs is documented as a hard cap on ANY single sleep, so it has
      // to bound the server-supplied Retry-After too. Dataverse can legally
      // ask for several minutes, and a misbehaving proxy can ask for hours —
      // either way, honouring it unbounded turns a throttle into a hang with
      // no output and no way to tell it apart from a wedged process.
      const delayMs = Math.min(
        parsed ? parsed.value : baseDelay * 2 ** (attempt - 1),
        maxDelay
      );

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

  /** List entity definitions, returning [{ LogicalName, EntitySetName, DisplayName, MetadataId }, ...]. */
  async listEntities(): Promise<
    Array<{ LogicalName: string; EntitySetName: string; DisplayName: string; MetadataId: string }>
  > {
    const res = await this.fetchWithRetry(
      this.url("EntityDefinitions?$select=LogicalName,EntitySetName,DisplayName,MetadataId"),
      { headers: { ...DEFAULT_HEADERS, ...(await this.authHeaders()) } }
    );
    if (!res.ok) await throwForResponse(res);
    const json = (await res.json()) as { value: Array<Record<string, unknown>> };
    return json.value.map((e) => ({
      LogicalName: String(e.LogicalName),
      EntitySetName: String(e.EntitySetName),
      DisplayName: extractLocalizedLabel(e.DisplayName),
      MetadataId: String(e.MetadataId ?? ""),
    }));
  }

  /**
   * List visible solutions (for filtering the entity picker in the UI).
   * Sorted by friendly name.
   */
  async listSolutions(): Promise<Array<{ id: string; uniqueName: string; friendlyName: string }>> {
    const rows = await this.queryAll(
      "solutions?$select=solutionid,uniquename,friendlyname&$filter=isvisible eq true&$orderby=friendlyname"
    );
    return rows
      .map((s) => ({
        id: String(s.solutionid ?? "").toLowerCase(),
        uniqueName: String(s.uniquename ?? ""),
        friendlyName: String(s.friendlyname ?? ""),
      }))
      .filter((s) => GUID_RE.test(s.id));
  }

  /**
   * MetadataIds of the entities that are components of a solution
   * (solutioncomponent componenttype 1 = Entity). Lowercased GUIDs.
   */
  async getSolutionEntityIds(solutionId: string): Promise<Set<string>> {
    assertGuid(solutionId, "Solution id");
    const rows = await this.queryAll(
      `solutioncomponents?$select=objectid&$filter=_solutionid_value eq ${solutionId} and componenttype eq 1`
    );
    return new Set(
      rows.map((r) => String(r.objectid ?? "").toLowerCase()).filter((id) => GUID_RE.test(id))
    );
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

  /**
   * The writable single-valued navigation property for a lookup attribute
   * bound to a specific referenced entity. Usually equal to the attribute
   * logical name, but NOT for polymorphic lookups (parentcustomerid →
   * "parentcustomerid_account") — and some lookup-typed attributes (e.g.
   * contact.accountid) have no writable navigation property at all, in which
   * case this returns undefined. Binding `attr@odata.bind` with the wrong
   * name fails the whole payload with an "undeclared property" error.
   */
  async getLookupNavigationProperty(
    entityLogicalName: string,
    attrLogicalName: string,
    referencedEntityLogicalName: string
  ): Promise<string | undefined> {
    assertLogicalName(entityLogicalName, "Entity logical name");
    assertLogicalName(attrLogicalName, "Attribute logical name");
    assertLogicalName(referencedEntityLogicalName, "Referenced entity logical name");
    const res = await this.fetchWithRetry(
      this.url(
        `EntityDefinitions(LogicalName='${entityLogicalName}')/ManyToOneRelationships` +
          `?$select=ReferencingAttribute,ReferencedEntity,ReferencingEntityNavigationPropertyName` +
          `&$filter=ReferencingAttribute eq '${attrLogicalName}' and ReferencedEntity eq '${referencedEntityLogicalName}'`
      ),
      { headers: { ...DEFAULT_HEADERS, ...(await this.authHeaders()) } }
    );
    if (!res.ok) await throwForResponse(res);
    const json = (await res.json()) as {
      value?: Array<{ ReferencingEntityNavigationPropertyName?: string }>;
    };
    return json.value?.[0]?.ReferencingEntityNavigationPropertyName || undefined;
  }

  /**
   * Alternate keys defined on an entity, with the attributes each one spans.
   *
   * Dataflows record their upsert key by the *key's* logical name, not by the
   * attributes it covers (`asker_actid` for a key over `asker_act_id`), so a
   * dataflow's mapping can only be turned into a dvload `upsertKey` by way of
   * this lookup.
   */
  async getEntityKeys(
    entityLogicalName: string
  ): Promise<Array<{ LogicalName: string; KeyAttributes: string[] }>> {
    assertLogicalName(entityLogicalName, "Entity logical name");
    const res = await this.fetchWithRetry(
      this.url(
        `EntityDefinitions(LogicalName='${entityLogicalName}')/Keys` +
          `?$select=LogicalName,KeyAttributes`
      ),
      { headers: { ...DEFAULT_HEADERS, ...(await this.authHeaders()) } }
    );
    if (!res.ok) await throwForResponse(res);
    const json = (await res.json()) as {
      value?: Array<{ LogicalName?: string; KeyAttributes?: string[] }>;
    };
    return (json.value ?? []).map((k) => ({
      LogicalName: String(k.LogicalName ?? ""),
      KeyAttributes: Array.isArray(k.KeyAttributes) ? k.KeyAttributes.map(String) : [],
    }));
  }

  /** Create a new table (entity). Payload from tablegen's buildEntityPayload. */
  async createEntity(payload: Record<string, unknown>): Promise<void> {
    const res = await this.fetchWithRetry(this.url("EntityDefinitions"), {
      method: "POST",
      headers: { ...DEFAULT_HEADERS, ...(await this.authHeaders()) },
      body: JSON.stringify(payload),
    });
    if (!res.ok) await throwForResponse(res);
  }

  /** Add an attribute to an existing table. Payload from buildAttributePayload. */
  async createAttribute(
    entityLogicalName: string,
    payload: Record<string, unknown>
  ): Promise<void> {
    assertLogicalName(entityLogicalName, "Entity logical name");
    const res = await this.fetchWithRetry(
      this.url(`EntityDefinitions(LogicalName='${entityLogicalName}')/Attributes`),
      {
        method: "POST",
        headers: { ...DEFAULT_HEADERS, ...(await this.authHeaders()) },
        body: JSON.stringify(payload),
      }
    );
    if (!res.ok) await throwForResponse(res);
  }

  /**
   * Create an alternate key. Payload from buildKeyPayload. Note the backing
   * index activates asynchronously — upserts against the key may not work
   * for a short while after creation; plain inserts are unaffected.
   */
  async createEntityKey(
    entityLogicalName: string,
    payload: Record<string, unknown>
  ): Promise<void> {
    assertLogicalName(entityLogicalName, "Entity logical name");
    const res = await this.fetchWithRetry(
      this.url(`EntityDefinitions(LogicalName='${entityLogicalName}')/Keys`),
      {
        method: "POST",
        headers: { ...DEFAULT_HEADERS, ...(await this.authHeaders()) },
        body: JSON.stringify(payload),
      }
    );
    if (!res.ok) await throwForResponse(res);
  }

  /** Create a single record. Slow for many rows — prefer batch(). */
  /**
   * @param primaryIdAttribute Primary key attribute name from entity metadata.
   *   Needed to read the id back when the server answers `Prefer:
   *   return=representation` with a body instead of an `OData-EntityId`
   *   header — which, given DEFAULT_HEADERS, is the normal case.
   */
  async create(
    entitySet: string,
    body: Record<string, unknown>,
    extraHeaders?: Record<string, string>,
    primaryIdAttribute?: string
  ): Promise<CreateResult> {
    const res = await this.fetchWithRetry(
      this.url(entitySet),
      {
        method: "POST",
        headers: { ...DEFAULT_HEADERS, ...(await this.authHeaders()), ...(extraHeaders ?? {}) },
        body: JSON.stringify(body),
      },
      undefined,
      // A plain create: if the connection dies after the server accepted it,
      // replaying makes a second record. Only failures that provably never
      // reached Dataverse are retried.
      false
    );
    if (!res.ok) await throwForResponse(res);
    return await parseCreateResult(res, primaryIdAttribute);
  }

  /**
   * GET a collection URL and follow @odata.nextLink until exhausted.
   * `path` is entity-set-relative and already query-encoded by the caller.
   */
  async queryAll(path: string): Promise<Array<Record<string, unknown>>> {
    const out: Array<Record<string, unknown>> = [];
    let next: string | null = this.url(path);
    while (next) {
      const res = await this.fetchWithRetry(next, {
        headers: {
          ...DEFAULT_HEADERS,
          ...(await this.authHeaders()),
          // Cap page size; Dataverse default is 5000 which is fine, but be explicit.
          Prefer: 'odata.maxpagesize=5000,odata.include-annotations="*"',
        },
      });
      if (!res.ok) await throwForResponse(res);
      const json = (await res.json()) as {
        value?: Array<Record<string, unknown>>;
        "@odata.nextLink"?: string;
      };
      out.push(...(json.value ?? []));
      next = json["@odata.nextLink"] ?? null;
    }
    return out;
  }

  /**
   * Look up { logicalName, primaryIdAttribute } for an entity SET name
   * (e.g. "accounts" → { account, accountid }). Cached per client instance.
   */
  async getEntitySetInfo(
    entitySetName: string
  ): Promise<{ logicalName: string; primaryIdAttribute: string; primaryNameAttribute?: string }> {
    assertLogicalName(entitySetName, "Entity set");
    const cached = this.entitySetInfoCache.get(entitySetName);
    if (cached) return cached;
    const res = await this.fetchWithRetry(
      this.url(
        `EntityDefinitions?$select=LogicalName,PrimaryIdAttribute,PrimaryNameAttribute&$filter=EntitySetName eq '${entitySetName}'`
      ),
      { headers: { ...DEFAULT_HEADERS, ...(await this.authHeaders()) } }
    );
    if (!res.ok) await throwForResponse(res);
    const json = (await res.json()) as {
      value?: Array<{ LogicalName?: string; PrimaryIdAttribute?: string; PrimaryNameAttribute?: string }>;
    };
    const e = json.value?.[0];
    if (!e?.LogicalName || !e?.PrimaryIdAttribute) {
      throw new Error(`No entity found with EntitySetName "${entitySetName}"`);
    }
    const info = {
      logicalName: e.LogicalName,
      primaryIdAttribute: e.PrimaryIdAttribute,
      primaryNameAttribute: e.PrimaryNameAttribute || undefined,
    };
    this.entitySetInfoCache.set(entitySetName, info);
    return info;
  }

  /**
   * Text lookup: find records where `attribute` equals `value` (exact match).
   * Returns up to 2 ids so the caller can detect ambiguity cheaply.
   */
  async resolveByText(
    entitySet: string,
    attribute: string,
    value: unknown,
    primaryIdAttribute: string
  ): Promise<string[]> {
    assertLogicalName(entitySet, "Entity set");
    assertLogicalName(attribute, "Attribute");
    assertLogicalName(primaryIdAttribute, "Primary id attribute");
    const literal = formatKeyLiteral(value);
    const res = await this.fetchWithRetry(
      this.url(
        `${entitySet}?$select=${primaryIdAttribute}&$filter=${attribute} eq ${literal}&$top=2`
      ),
      { headers: { ...DEFAULT_HEADERS, ...(await this.authHeaders()) } }
    );
    if (!res.ok) await throwForResponse(res);
    const json = (await res.json()) as { value?: Array<Record<string, unknown>> };
    return (json.value ?? [])
      .map((r) => String(r[primaryIdAttribute] ?? ""))
      .filter(Boolean);
  }

  /**
   * Batched text lookup: resolve MANY distinct values against one attribute
   * in chunked `or`-filter queries instead of one request per value (the
   * N+1 pattern resolveByText produces on large lookup columns).
   *
   * Returns a map keyed by the LOWERCASED source value (Dataverse text
   * comparison is case-insensitive) → all matching record ids. Values with
   * no entry in the map had no match.
   */
  async resolveManyByText(
    entitySet: string,
    attribute: string,
    values: unknown[],
    primaryIdAttribute: string
  ): Promise<Map<string, string[]>> {
    assertLogicalName(entitySet, "Entity set");
    assertLogicalName(attribute, "Attribute");
    assertLogicalName(primaryIdAttribute, "Primary id attribute");

    const out = new Map<string, string[]>();
    if (values.length === 0) return out;

    // Chunk small enough to keep the URL well under common limits even for
    // long-ish values (each is percent-encoded by formatKeyLiteral).
    const CHUNK = 15;
    for (let start = 0; start < values.length; start += CHUNK) {
      const chunk = values.slice(start, start + CHUNK);
      const filter = chunk
        .map((v) => `${attribute} eq ${formatKeyLiteral(v)}`)
        .join(" or ");
      const rows = await this.queryAll(
        `${entitySet}?$select=${primaryIdAttribute},${attribute}&$filter=${filter}`
      );
      for (const r of rows) {
        const key = String(r[attribute] ?? "").toLowerCase();
        const id = String(r[primaryIdAttribute] ?? "");
        if (!id) continue;
        const list = out.get(key) ?? [];
        list.push(id);
        out.set(key, list);
      }
    }
    return out;
  }

  /**
   * Read one record by alternate-key expression with a $select list.
   * Returns null on 404. keyExpr must be built from validated parts
   * (see buildKeyExpression in load.ts).
   */
  async getRecord(
    entitySet: string,
    keyExpr: string,
    select: string[]
  ): Promise<Record<string, unknown> | null> {
    assertLogicalName(entitySet, "Entity set");
    for (const s of select) assertLogicalName(s, "Select attribute");
    const res = await this.fetchWithRetry(
      this.url(`${entitySet}(${keyExpr})?$select=${select.join(",")}`),
      { headers: { ...DEFAULT_HEADERS, ...(await this.authHeaders()) } }
    );
    if (res.status === 404) return null;
    if (!res.ok) await throwForResponse(res);
    return (await res.json()) as Record<string, unknown>;
  }

  /**
   * Fetch option-set label → integer value for a picklist or multi-select
   * picklist attribute. Handles both local and global option sets.
   */
  async getOptionSetLabels(
    entityLogicalName: string,
    attrLogicalName: string
  ): Promise<Record<string, number>> {
    assertLogicalName(entityLogicalName, "Entity logical name");
    assertLogicalName(attrLogicalName, "Attribute logical name");
    const casts = [
      "Microsoft.Dynamics.CRM.PicklistAttributeMetadata",
      "Microsoft.Dynamics.CRM.MultiSelectPicklistAttributeMetadata",
      "Microsoft.Dynamics.CRM.StatusAttributeMetadata",
      "Microsoft.Dynamics.CRM.StateAttributeMetadata",
    ];
    for (const cast of casts) {
      const res = await this.fetchWithRetry(
        this.url(
          `EntityDefinitions(LogicalName='${entityLogicalName}')/Attributes/${cast}` +
            `?$select=LogicalName&$filter=LogicalName eq '${attrLogicalName}'` +
            `&$expand=OptionSet($select=Options),GlobalOptionSet($select=Options)`
        ),
        { headers: { ...DEFAULT_HEADERS, ...(await this.authHeaders()) } }
      );
      if (!res.ok) continue; // wrong cast for this attribute; try the next
      const json = (await res.json()) as {
        value?: Array<{
          OptionSet?: { Options?: OptionMetadata[] };
          GlobalOptionSet?: { Options?: OptionMetadata[] };
        }>;
      };
      const attr = json.value?.[0];
      const options = attr?.OptionSet?.Options ?? attr?.GlobalOptionSet?.Options;
      if (!options) continue;
      const map: Record<string, number> = {};
      for (const o of options) {
        const label = o.Label?.UserLocalizedLabel?.Label;
        if (label && typeof o.Value === "number") map[label] = o.Value;
      }
      if (Object.keys(map).length > 0) return map;
    }
    throw new Error(
      `No option-set metadata found for ${entityLogicalName}.${attrLogicalName} ` +
        `(is it a choice/multichoice/status/state attribute?)`
    );
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
  async batch(
    operations: BatchOperation[],
    opts?: {
      /**
       * Whether replaying the whole batch is safe. PATCH-by-key upserts and
       * DELETEs are; plain POST creates are NOT — a 503/504 can arrive after
       * some operations already executed, and a replay would duplicate them.
       * When false, only 429 (guaranteed not-executed) is retried.
       */
      idempotent?: boolean;
    }
  ): Promise<BatchResultItem[]> {
    if (operations.length === 0) return [];
    const batchId = `batch_${cryptoUuid()}`;

    const body = buildBatchBody(operations, batchId, this.base);

    const retryableOverride = opts?.idempotent === false ? [429] : undefined;
    const res = await this.fetchWithRetry(
      this.url("$batch"),
      {
        method: "POST",
        headers: {
          ...(await this.authHeaders()),
          "OData-MaxVersion": "4.0",
          "OData-Version": "4.0",
          Accept: "application/json",
          "Content-Type": `multipart/mixed; boundary=${batchId}`,
          // Without this, Dataverse stops at the first failed changeset and
          // returns the REMAINING operations unexecuted with an outer 400 —
          // even though earlier changesets already committed. With it, every
          // changeset is attempted and per-operation statuses come back.
          Prefer: "odata.continue-on-error",
        },
        body,
      },
      retryableOverride,
      // A mid-flight connection drop is ambiguous: the changesets may already
      // have committed. Replaying is only safe when the ops are idempotent.
      opts?.idempotent !== false
    );

    const text = await res.text();
    // Even with continue-on-error, some failures (and older endpoints or
    // proxies) surface as an outer non-2xx WITH a multipart body describing
    // per-operation outcomes — including operations that DID succeed. Parse
    // that body rather than throwing, or rows that were actually created
    // would be reported failed (and a naive re-run would duplicate them).
    // Only treat it as a whole-batch failure when there's no batch body.
    if (!res.ok && !/^--batchresponse/m.test(text)) {
      // Synchronous (returns never) — unlike throwForResponse, the body has
      // already been read into `text`.
      throwForResponseText(res, text);
    }
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
  let text = "";
  try {
    text = await res.text();
  } catch {
    // body unreadable
  }
  return throwForResponseText(res, text);
}

/** Like throwForResponse, but for callers that already consumed the body. */
function throwForResponseText(res: Response, text: string): never {
  let raw: unknown;
  let message = `${res.status} ${res.statusText}`;
  let code: string | undefined;
  try {
    raw = JSON.parse(text);
    const err = (raw as { error?: { code?: string; message?: string } }).error;
    if (err?.message) message = err.message;
    if (err?.code) code = err.code;
  } catch {
    // body wasn't JSON
  }
  throw new DataverseError(res.status, code, message, raw);
}

const GUID_IN_PARENS = /\(([0-9a-f-]{36})\)/i;
const BARE_GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/**
 * Pull the new record's id out of a create response.
 *
 * There are two response shapes, and which one you get depends on a header
 * this client always sends:
 *
 *   204 No Content  →  id is in the `OData-EntityId` header
 *   201 Created     →  body IS the record, and Dataverse omits
 *                      `OData-EntityId` entirely
 *
 * `DEFAULT_HEADERS` sets `Prefer: return=representation`, so in practice we
 * always get the second shape. Reading only the header therefore returned
 * `id: ""` for *every* create. That surfaced far downstream as
 * "create-if-missing returned no id" — while the records were being created
 * perfectly well, which is the worst version of this bug: it looks like a
 * failure and leaves data behind.
 *
 * Note what this deliberately does NOT do: guess the id by scanning the body
 * for a property named like `*id`. A Dataverse record is full of them —
 * `ownerid`, `businessunitid`, `_createdby_value` — and binding a lookup to
 * the wrong GUID is far worse than failing loudly.
 */
async function parseCreateResult(res: Response, primaryIdAttribute?: string): Promise<CreateResult> {
  const headerUrl = res.headers.get("OData-EntityId") ?? res.headers.get("Location") ?? "";
  const fromHeader = GUID_IN_PARENS.exec(headerUrl)?.[1];
  if (fromHeader) return { id: fromHeader, entityUrl: headerUrl };

  // 204 with no header: nothing else to read.
  if (res.status === 204) return { id: "", entityUrl: headerUrl };

  let payload: Record<string, unknown>;
  try {
    payload = (await res.json()) as Record<string, unknown>;
  } catch {
    return { id: "", entityUrl: headerUrl };
  }
  if (!payload || typeof payload !== "object") return { id: "", entityUrl: headerUrl };

  // Preferred: the caller knows the primary key's name from entity metadata.
  if (primaryIdAttribute) {
    const value = payload[primaryIdAttribute];
    if (typeof value === "string" && BARE_GUID.test(value)) {
      const odataId = typeof payload["@odata.id"] === "string" ? payload["@odata.id"] : headerUrl;
      return { id: value, entityUrl: odataId };
    }
  }

  // Fallback: the @odata.id annotation carries the canonical record URL.
  const odataId = typeof payload["@odata.id"] === "string" ? payload["@odata.id"] : "";
  const fromOdataId = odataId ? GUID_IN_PARENS.exec(odataId)?.[1] : undefined;
  if (fromOdataId) return { id: fromOdataId, entityUrl: odataId };

  return { id: "", entityUrl: headerUrl };
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
        }
      } catch {
        if (!ok) errorMessage = rawBody;
      }
    } else if (!ok) {
      errorMessage = `HTTP ${status}`;
    }

    // Created-record id: read the OData-EntityId header (canonical entity
    // URL) rather than guessing the primary-id attribute name from the
    // entity-set name — naive de-pluralization breaks on "opportunities",
    // "addresses", etc.
    if (op.method === "POST" || op.method === "PATCH") {
      const eid = /OData-EntityId:\s*[^\r\n]*\(([0-9a-f-]{36})\)/i.exec(part)?.[1];
      if (eid) id = eid;
      // With Prefer: return=representation the body has @odata.id instead.
      if (!id && body && typeof body === "object") {
        const odataId = String((body as Record<string, unknown>)["@odata.id"] ?? "");
        const m = /\(([0-9a-f-]{36})\)/i.exec(odataId);
        if (m) id = m[1];
      }
    }

    results.push({ contentId, status, ok, body, errorMessage, id });
  }

  // Preserve original order.
  results.sort((a, b) => a.contentId - b.contentId);
  return results;
}
