/**
 * An in-process fake of the Dataverse OData Web API.
 *
 * WHY THIS EXISTS
 * ---------------
 * The existing unit tests stub `DataverseClient` itself, which means the most
 * dangerous code in this repo — batch body construction, multipart response
 * parsing, retry classification, upsert status accounting — is never actually
 * exercised. A hand-written fake client will happily agree with a buggy
 * serializer.
 *
 * This module goes one layer lower: it implements the *wire protocol*, so
 * tests drive the real `DataverseClient` over a real (or realistic) HTTP
 * boundary. If `buildBatchBody` emits something Dataverse would reject, these
 * tests fail.
 *
 * TWO MODES
 * ---------
 *   createFakeDataverse().fetch   — a `fetch`-compatible function. Fast, no
 *                                   sockets, works identically on every OS.
 *   await createFakeDataverse().listen()
 *                                 — a real `node:http` server. Slower, but the
 *                                   only way to exercise undici behaviour such
 *                                   as mid-flight connection drops.
 *
 * FIDELITY NOTES (deliberate simplifications, documented so nobody assumes
 * more than is here):
 *   - `$filter` supports `<attr> eq <literal>` joined by ` or `/` and `. That
 *     covers every filter the client actually builds today.
 *   - Text comparison is case-insensitive, matching Dataverse's default
 *     collation, because `resolveManyByText` depends on that behaviour.
 *   - No relationship traversal, no `$expand` on records, no FetchXML.
 *   - Option-set metadata is served from whatever the fixture declares.
 *
 * Anything a test needs that isn't here should be added here rather than
 * mocked at the client level — that's the whole point.
 */

import { createServer, type Server } from "node:http";
import { AddressInfo } from "node:net";

/* -------------------------------------------------------------------------- */
/* Fixture shapes                                                              */
/* -------------------------------------------------------------------------- */

export interface FakeEntityDef {
  /** Entity set name as used in URLs, e.g. "contacts". */
  entitySet: string;
  /** Logical name, e.g. "contact". */
  logicalName: string;
  /** Primary id attribute, e.g. "contactid". */
  primaryIdAttribute: string;
  /** Primary name attribute, e.g. "fullname". */
  primaryNameAttribute?: string;
  /**
   * Alternate keys, each an ordered list of attributes. Only used to decide
   * whether a keyed PATCH is legal; lookups match on attribute equality
   * regardless, so a missing declaration will not silently pass a test that
   * should fail against a real environment.
   */
  alternateKeys?: string[][];
  /** Attribute metadata surfaced through EntityDefinitions. */
  attributes?: FakeAttributeDef[];
  /** Records present before the test starts. */
  seed?: Array<Record<string, unknown>>;
}

export interface FakeAttributeDef {
  LogicalName: string;
  AttributeType: string;
  /** For picklist attributes: label → value. */
  options?: Record<string, number>;
  /** Lookup targets, for `getLookupTargets`. */
  targets?: string[];
}

/** One request the fake observed. Assert against this instead of spying. */
export interface ObservedRequest {
  method: string;
  /** Path + query, relative to the API root. */
  path: string;
  headers: Record<string, string>;
  body?: string;
  status: number;
  /** For $batch requests: the parsed sub-operations. */
  batchOps?: Array<{ contentId: number; method: string; url: string; headers: Record<string, string> }>;
}

/** A programmed failure. Return `null` to let the operation proceed normally. */
export type OperationFault = (op: {
  method: string;
  /** Entity-set-relative URL of the sub-operation. */
  url: string;
  entitySet: string;
  body: Record<string, unknown> | undefined;
  /** 0-based index across every batch operation the fake has seen. */
  seq: number;
}) => { status: number; code?: string; message: string } | null;

export interface FakeDataverseOptions {
  entities?: FakeEntityDef[];
  /** Page size for collection GETs. Small values exercise @odata.nextLink. */
  pageSize?: number;
  /** Base URL reported in @odata.id and OData-EntityId. Overridden by listen(). */
  environmentUrl?: string;
  apiVersion?: string;
}

/* -------------------------------------------------------------------------- */
/* Deterministic ids                                                           */
/* -------------------------------------------------------------------------- */

/**
 * Sequential, valid-looking GUIDs. Deterministic on purpose: a failing test
 * prints the same ids on every run, and snapshots stay stable.
 */
function makeGuidFactory(): () => string {
  let n = 0;
  return () => {
    n++;
    const hex = n.toString(16).padStart(12, "0");
    return `00000000-0000-4000-8000-${hex}`;
  };
}

/* -------------------------------------------------------------------------- */
/* OData literal / key-expression handling                                     */
/* -------------------------------------------------------------------------- */

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/**
 * Inverse of `formatKeyLiteral` in src/dataverse.ts. Kept deliberately
 * independent of that implementation: if the encoder changes in a way the
 * decoder can't undo, that's a real protocol change and the tests should
 * notice rather than silently agree with themselves.
 */
export function parseODataLiteral(raw: string): unknown {
  const s = raw.trim();
  if (s === "true") return true;
  if (s === "false") return false;
  if (s === "null") return null;
  if (/^-?\d+(\.\d+)?$/.test(s)) return Number(s);
  if (s.startsWith("'") && s.endsWith("'")) {
    const inner = s.slice(1, -1);
    return decodeURIComponent(inner).replace(/''/g, "'");
  }
  return decodeURIComponent(s);
}

/** Parse `attr='a',attr2=3` or a bare GUID into a predicate over records. */
export function parseKeyExpression(
  expr: string
): { kind: "id"; id: string } | { kind: "alternate"; parts: Array<[string, unknown]> } {
  const trimmed = expr.trim();
  if (GUID_RE.test(trimmed)) return { kind: "id", id: trimmed.toLowerCase() };

  const parts: Array<[string, unknown]> = [];
  // Split on commas that are not inside a quoted literal. Values are
  // percent-encoded by the client, so a raw comma can only be a separator —
  // but be strict anyway so a future encoder change surfaces here.
  for (const seg of splitTopLevel(trimmed, ",")) {
    const eq = seg.indexOf("=");
    if (eq < 0) throw new Error(`Malformed key expression segment: ${seg}`);
    parts.push([seg.slice(0, eq).trim(), parseODataLiteral(seg.slice(eq + 1))]);
  }
  return { kind: "alternate", parts };
}

function splitTopLevel(s: string, sep: string): string[] {
  const out: string[] = [];
  let depth = 0;
  let inQuote = false;
  let cur = "";
  for (let i = 0; i < s.length; i++) {
    const c = s[i];
    if (c === "'") inQuote = !inQuote;
    else if (!inQuote && c === "(") depth++;
    else if (!inQuote && c === ")") depth--;
    if (!inQuote && depth === 0 && c === sep) {
      out.push(cur);
      cur = "";
      continue;
    }
    cur += c;
  }
  if (cur) out.push(cur);
  return out;
}

/** Dataverse text comparison is case-insensitive; numbers compare numerically. */
function looseEquals(a: unknown, b: unknown): boolean {
  if (a === null || a === undefined) return b === null || b === undefined;
  if (typeof a === "string" || typeof b === "string") {
    return String(a).toLowerCase() === String(b).toLowerCase();
  }
  return a === b;
}

/**
 * Evaluate the narrow `$filter` grammar the client emits:
 *   `attr eq <literal>` combined with ` or ` / ` and `.
 */
export function evaluateFilter(filter: string, record: Record<string, unknown>): boolean {
  const orTerms = filter.split(/\s+or\s+/i);
  return orTerms.some((orTerm) =>
    orTerm.split(/\s+and\s+/i).every((term) => {
      const m = /^\s*\(?\s*([A-Za-z_][A-Za-z0-9_]*)\s+eq\s+(.+?)\s*\)?\s*$/.exec(term);
      if (!m) throw new Error(`fake-dataverse: unsupported $filter term: ${JSON.stringify(term)}`);
      return looseEquals(record[m[1]], parseODataLiteral(m[2]));
    })
  );
}

/* -------------------------------------------------------------------------- */
/* The fake                                                                    */
/* -------------------------------------------------------------------------- */

export interface FakeDataverse {
  /** Drop-in `fetch` for DataverseClientOptions.fetch. */
  fetch: typeof fetch;
  /** Every request observed, in order. */
  readonly requests: ObservedRequest[];
  /** Live record store, keyed by entity set then primary id. */
  records(entitySet: string): Array<Record<string, unknown>>;
  /** Insert records without going through the API (test setup). */
  seed(entitySet: string, rows: Array<Record<string, unknown>>): void;
  /** Base URL to hand to DataverseClient. */
  readonly environmentUrl: string;

  /* --- fault injection ---------------------------------------------------- */

  /**
   * Make the next `count` HTTP responses a throttle. `retryAfter` is emitted
   * verbatim as the Retry-After header (a number of seconds or an HTTP-date).
   */
  throttleNext(count: number, opts?: { status?: number; retryAfter?: string | number }): void;
  /**
   * Make the next `count` requests throw instead of responding, simulating a
   * dropped connection. `code` drives the client's retry classification:
   * "ECONNRESET" is ambiguous, "ECONNREFUSED" provably never executed.
   */
  failNextNetwork(count: number, code?: string): void;
  /** Programmatic per-operation failures inside $batch. */
  setOperationFault(fault: OperationFault | null): void;
  /** Reset counters and fault programming (not the records). */
  resetFaults(): void;

  /* --- real socket mode --------------------------------------------------- */

  /** Start a real HTTP server; returns its base URL. */
  listen(): Promise<string>;
  close(): Promise<void>;
}

export function createFakeDataverse(opts: FakeDataverseOptions = {}): FakeDataverse {
  const apiVersion = opts.apiVersion ?? "v9.2";
  let environmentUrl = opts.environmentUrl ?? "https://fake.crm.dynamics.com";
  const pageSize = opts.pageSize ?? 5000;
  const newGuid = makeGuidFactory();

  const defs = new Map<string, FakeEntityDef>();
  const store = new Map<string, Map<string, Record<string, unknown>>>();
  const requests: ObservedRequest[] = [];

  let throttleRemaining = 0;
  let throttleStatus = 429;
  let throttleRetryAfter: string | number | undefined = 1;
  let networkFailRemaining = 0;
  let networkFailCode = "ECONNRESET";
  let operationFault: OperationFault | null = null;
  let opSeq = 0;

  for (const e of opts.entities ?? []) {
    defs.set(e.entitySet, e);
    const m = new Map<string, Record<string, unknown>>();
    for (const r of e.seed ?? []) {
      const id = String(r[e.primaryIdAttribute] ?? newGuid());
      m.set(id, { statecode: 0, ...r, [e.primaryIdAttribute]: id });
    }
    store.set(e.entitySet, m);
  }

  const apiRoot = (): string => `${environmentUrl.replace(/\/+$/, "")}/api/data/${apiVersion}`;

  function defFor(entitySet: string): FakeEntityDef {
    const d = defs.get(entitySet);
    if (!d) throw new HttpError(404, "0x80060888", `Resource not found for the segment '${entitySet}'.`);
    return d;
  }

  function rowsOf(entitySet: string): Map<string, Record<string, unknown>> {
    let m = store.get(entitySet);
    if (!m) {
      m = new Map();
      store.set(entitySet, m);
    }
    return m;
  }

  /* --- record operations -------------------------------------------------- */

  function findByKey(entitySet: string, expr: string): Record<string, unknown> | undefined {
    defFor(entitySet); // 404s for an unknown entity set, as Dataverse would
    const key = parseKeyExpression(expr);
    const rows = rowsOf(entitySet);
    if (key.kind === "id") return rows.get(key.id);
    for (const r of rows.values()) {
      if (key.parts.every(([attr, val]) => looseEquals(r[attr], val))) return r;
    }
    return undefined;
  }

  function createRecord(
    entitySet: string,
    body: Record<string, unknown>,
    forcedId?: string
  ): Record<string, unknown> {
    const def = defFor(entitySet);
    const id = forcedId ?? String(body[def.primaryIdAttribute] ?? newGuid());
    const rec: Record<string, unknown> = {
      statecode: 0,
      ...stripAnnotations(body),
      [def.primaryIdAttribute]: id,
    };
    rowsOf(entitySet).set(id, rec);
    return rec;
  }

  /**
   * `foo@odata.bind: "/accounts(guid)"` becomes the flattened lookup value
   * `_foo_value`, which is how Dataverse surfaces it on read. Tests that
   * assert on lookup binding therefore see something realistic.
   */
  function stripAnnotations(body: Record<string, unknown>): Record<string, unknown> {
    const out: Record<string, unknown> = {};
    for (const [k, v] of Object.entries(body)) {
      const bind = /^(.+)@odata\.bind$/.exec(k);
      if (bind) {
        const guid = /\(([0-9a-f-]{36})\)/i.exec(String(v))?.[1];
        out[`_${bind[1]}_value`] = guid ?? null;
        continue;
      }
      if (k.includes("@")) continue;
      out[k] = v;
    }
    return out;
  }

  class HttpError extends Error {
    constructor(
      public status: number,
      public code: string | undefined,
      message: string
    ) {
      super(message);
    }
  }

  /* --- routing ------------------------------------------------------------ */

  interface RawRequest {
    method: string;
    /** Path + query relative to the API root. */
    path: string;
    headers: Record<string, string>;
    body: string;
  }

  interface RawResponse {
    status: number;
    headers: Record<string, string>;
    body: string;
  }

  function handle(req: RawRequest): RawResponse {
    try {
      return route(req);
    } catch (e) {
      if (e instanceof HttpError) {
        return json(e.status, { error: { code: e.code ?? "0x80040265", message: e.message } });
      }
      return json(500, {
        error: { code: "0x80040265", message: e instanceof Error ? e.message : String(e) },
      });
    }
  }

  function json(status: number, body: unknown, extra: Record<string, string> = {}): RawResponse {
    return {
      status,
      headers: { "Content-Type": "application/json; charset=utf-8", ...extra },
      body: JSON.stringify(body),
    };
  }

  function route(req: RawRequest): RawResponse {
    const [rawPath, rawQuery = ""] = req.path.split("?");
    const path = decodeURIComponent(rawPath);
    const query = parseQuery(rawQuery);

    if (path === "$batch" && req.method === "POST") return handleBatch(req);
    if (path.startsWith("EntityDefinitions")) return handleMetadata(path, query);

    // <entitySet>(<keyExpr>)
    const keyed = /^([A-Za-z_][A-Za-z0-9_]*)\((.+)\)$/.exec(path);
    if (keyed) return handleKeyed(req.method, keyed[1], keyed[2], query, req);

    // <entitySet>
    const collection = /^([A-Za-z_][A-Za-z0-9_]*)$/.exec(path);
    if (collection) return handleCollection(req.method, collection[1], query, req);

    throw new HttpError(404, "0x80060888", `Resource not found for the segment '${path}'.`);
  }

  function parseQuery(raw: string): Record<string, string> {
    const out: Record<string, string> = {};
    for (const pair of raw.split("&")) {
      if (!pair) continue;
      const i = pair.indexOf("=");
      const key = decodeURIComponent(i < 0 ? pair : pair.slice(0, i));
      const v = i < 0 ? "" : pair.slice(i + 1);
      // $filter values keep their percent-encoding until parseODataLiteral
      // handles them — decoding here would undo the very escaping we want to
      // verify. `+` still means space, because an @odata.nextLink is built
      // with URLSearchParams and encodes it that way.
      out[key] = key === "$filter" ? v.replace(/\+/g, " ") : decodeURIComponent(v);
    }
    return out;
  }

  function handleMetadata(path: string, query: Record<string, string>): RawResponse {
    // EntityDefinitions?$filter=EntitySetName eq 'contacts'
    const filter = query["$filter"] ?? "";
    const bySet = /EntitySetName\s+eq\s+'([^']+)'/.exec(decodeURIComponent(filter));
    if (bySet) {
      const d = defs.get(bySet[1]);
      return json(200, {
        value: d
          ? [
              {
                LogicalName: d.logicalName,
                PrimaryIdAttribute: d.primaryIdAttribute,
                PrimaryNameAttribute: d.primaryNameAttribute ?? null,
              },
            ]
          : [],
      });
    }

    // EntityDefinitions(LogicalName='contact')?$expand=Attributes
    const byLogical = /LogicalName='([^']+)'/.exec(path);
    if (byLogical) {
      const d = [...defs.values()].find((x) => x.logicalName === byLogical[1]);
      if (!d) throw new HttpError(404, "0x80060888", `Entity '${byLogical[1]}' not found.`);

      // .../Attributes/<cast>?$filter=LogicalName eq 'x' — option-set metadata
      if (path.includes("/Attributes/")) {
        const attrName = /LogicalName\s+eq\s+'([^']+)'/.exec(decodeURIComponent(query["$filter"] ?? ""))?.[1];
        const attr = d.attributes?.find((a) => a.LogicalName === attrName);
        const cast = /\/Attributes\/([^?]+)/.exec(path)?.[1] ?? "";
        const isPicklistCast = /Picklist|Status|State/.test(cast);
        if (!attr?.options || !isPicklistCast) return json(200, { value: [] });
        return json(200, {
          value: [
            {
              LogicalName: attr.LogicalName,
              OptionSet: {
                Options: Object.entries(attr.options).map(([label, value]) => ({
                  Value: value,
                  Label: { UserLocalizedLabel: { Label: label } },
                })),
              },
            },
          ],
        });
      }

      return json(200, {
        LogicalName: d.logicalName,
        EntitySetName: d.entitySet,
        PrimaryIdAttribute: d.primaryIdAttribute,
        PrimaryNameAttribute: d.primaryNameAttribute ?? null,
        Attributes: d.attributes ?? [],
      });
    }

    return json(200, {
      value: [...defs.values()].map((d) => ({
        LogicalName: d.logicalName,
        EntitySetName: d.entitySet,
        PrimaryIdAttribute: d.primaryIdAttribute,
        PrimaryNameAttribute: d.primaryNameAttribute ?? null,
      })),
    });
  }

  function handleKeyed(
    method: string,
    entitySet: string,
    keyExpr: string,
    query: Record<string, string>,
    req: RawRequest
  ): RawResponse {
    const def = defFor(entitySet);
    const existing = findByKey(entitySet, keyExpr);

    if (method === "GET") {
      if (!existing) throw new HttpError(404, "0x80040217", "Record does not exist.");
      return json(200, {
        "@odata.id": `${apiRoot()}/${entitySet}(${existing[def.primaryIdAttribute]})`,
        ...project(existing, query["$select"]),
      });
    }

    if (method === "DELETE") {
      if (!existing) throw new HttpError(404, "0x80040217", "Record does not exist.");
      rowsOf(entitySet).delete(String(existing[def.primaryIdAttribute]));
      return { status: 204, headers: {}, body: "" };
    }

    if (method === "PATCH") {
      const body = req.body ? (JSON.parse(req.body) as Record<string, unknown>) : {};

      // If-None-Match: * is "create only" — 412 when the record already exists.
      if (headerOf(req.headers, "if-none-match") === "*" && existing) {
        throw new HttpError(412, "0x80060893", "A record with matching key values already exists.");
      }
      // If-Match: * is "update only" — 404 when it does not.
      if (headerOf(req.headers, "if-match") === "*" && !existing) {
        throw new HttpError(404, "0x80040217", "Record does not exist.");
      }

      if (existing) {
        Object.assign(existing, stripAnnotations(body));
        return json(
          200,
          {
            "@odata.id": `${apiRoot()}/${entitySet}(${existing[def.primaryIdAttribute]})`,
            ...existing,
          },
          { "OData-EntityId": `${apiRoot()}/${entitySet}(${existing[def.primaryIdAttribute]})` }
        );
      }

      // Upsert-create. A keyed PATCH seeds the key attributes into the new
      // record, exactly as Dataverse does.
      const key = parseKeyExpression(keyExpr);
      const seeded: Record<string, unknown> = { ...body };
      let forcedId: string | undefined;
      if (key.kind === "id") forcedId = key.id;
      else for (const [attr, val] of key.parts) seeded[attr] ??= val;

      const rec = createRecord(entitySet, seeded, forcedId);
      const url = `${apiRoot()}/${entitySet}(${rec[def.primaryIdAttribute]})`;
      return json(201, { "@odata.id": url, ...rec }, { "OData-EntityId": url });
    }

    throw new HttpError(405, undefined, `Method ${method} not allowed on a keyed URL.`);
  }

  function handleCollection(
    method: string,
    entitySet: string,
    query: Record<string, string>,
    req: RawRequest
  ): RawResponse {
    const def = defFor(entitySet);

    if (method === "POST") {
      const body = req.body ? (JSON.parse(req.body) as Record<string, unknown>) : {};
      const rec = createRecord(entitySet, body);
      const url = `${apiRoot()}/${entitySet}(${rec[def.primaryIdAttribute]})`;
      return json(201, { "@odata.id": url, ...rec }, { "OData-EntityId": url });
    }

    if (method !== "GET") throw new HttpError(405, undefined, `Method ${method} not allowed.`);

    let rows = [...rowsOf(entitySet).values()];
    if (query["$filter"]) rows = rows.filter((r) => evaluateFilter(query["$filter"], r));

    const skip = Number(query["$skiptoken"] ?? 0);
    const top = query["$top"] ? Number(query["$top"]) : undefined;
    if (top !== undefined) rows = rows.slice(0, top);

    const page = rows.slice(skip, skip + pageSize);
    const value = page.map((r) => ({
      "@odata.id": `${apiRoot()}/${entitySet}(${r[def.primaryIdAttribute]})`,
      ...project(r, query["$select"]),
    }));

    const body: Record<string, unknown> = { value };
    if (skip + pageSize < rows.length) {
      const next = new URLSearchParams(query as Record<string, string>);
      next.set("$skiptoken", String(skip + pageSize));
      body["@odata.nextLink"] = `${apiRoot()}/${entitySet}?${next.toString()}`;
    }
    return json(200, body);
  }

  function project(rec: Record<string, unknown>, select?: string): Record<string, unknown> {
    if (!select) return { ...rec };
    const keep = new Set(
      select
        .split(",")
        .map((s) => s.trim())
        .filter(Boolean)
    );
    const out: Record<string, unknown> = {};
    for (const k of keep) if (k in rec) out[k] = rec[k];
    return out;
  }

  /* --- $batch ------------------------------------------------------------- */

  interface ParsedBatchOp {
    contentId: number;
    method: string;
    /** Entity-set-relative URL. */
    url: string;
    headers: Record<string, string>;
    body: string;
  }

  /**
   * Parse the multipart/mixed body produced by buildBatchBody. Intentionally
   * strict: a malformed request line or a missing Content-ID is a protocol
   * bug in the client and should fail the test, not be papered over.
   */
  function parseBatchRequest(body: string, apiBase: string): ParsedBatchOp[] {
    const ops: ParsedBatchOp[] = [];
    const segments = body.split(/^--[^\r\n]*\r?\n?/m);

    for (const seg of segments) {
      if (!/HTTP\/1\.1/.test(seg)) continue;

      const contentId = Number(/Content-ID:\s*(\d+)/i.exec(seg)?.[1] ?? NaN);
      if (!Number.isFinite(contentId)) {
        throw new Error(`fake-dataverse: batch part missing Content-ID:\n${seg}`);
      }

      const reqLine = /^(GET|POST|PATCH|DELETE|PUT)\s+(\S+)\s+HTTP\/1\.1\s*$/m.exec(seg);
      if (!reqLine) throw new Error(`fake-dataverse: batch part has no request line:\n${seg}`);

      const [, method, absUrl] = reqLine;
      if (!absUrl.startsWith(apiBase)) {
        throw new Error(`fake-dataverse: batch op URL is not under the API root: ${absUrl}`);
      }
      const url = absUrl.slice(apiBase.length).replace(/^\/+/, "");

      // Headers run from the line after the request line to the blank line.
      const afterReqLine = seg.slice(reqLine.index + reqLine[0].length);
      const blank = afterReqLine.search(/\r?\n\r?\n/);
      const headerBlock = blank >= 0 ? afterReqLine.slice(0, blank) : afterReqLine;
      const rawBody = blank >= 0 ? afterReqLine.slice(blank).replace(/^\r?\n\r?\n/, "") : "";

      const headers: Record<string, string> = {};
      for (const line of headerBlock.split(/\r?\n/)) {
        const m = /^([A-Za-z0-9-]+):\s*(.*)$/.exec(line.trim());
        if (m) headers[m[1].toLowerCase()] = m[2];
      }

      ops.push({ contentId, method, url, headers, body: rawBody.trim() });
    }
    return ops;
  }

  function handleBatch(req: RawRequest): RawResponse {
    const apiBase = apiRoot();
    const ops = parseBatchRequest(req.body, apiBase);

    const boundary = "batchresponse_00000000-0000-0000-0000-000000000001";
    const lines: string[] = [];

    for (const op of ops) {
      const changeset = `changesetresponse_${op.contentId}`;
      const entitySet = /^([A-Za-z_][A-Za-z0-9_]*)/.exec(op.url)?.[1] ?? "";

      let res: RawResponse;
      const fault = operationFault?.({
        method: op.method,
        url: op.url,
        entitySet,
        body: op.body ? (JSON.parse(op.body) as Record<string, unknown>) : undefined,
        seq: opSeq,
      });
      opSeq++;

      if (fault) {
        res = json(fault.status, {
          error: { code: fault.code ?? "0x80040265", message: fault.message },
        });
      } else {
        res = handle({ method: op.method, path: op.url, headers: op.headers, body: op.body });
      }

      lines.push(`--${boundary}`);
      lines.push(`Content-Type: multipart/mixed; boundary=${changeset}`);
      lines.push("");
      lines.push(`--${changeset}`);
      lines.push("Content-Type: application/http");
      lines.push("Content-Transfer-Encoding: binary");
      lines.push(`Content-ID: ${op.contentId}`);
      lines.push("");
      lines.push(`HTTP/1.1 ${res.status} ${statusText(res.status)}`);
      for (const [k, v] of Object.entries(res.headers)) lines.push(`${k}: ${v}`);
      lines.push("");
      lines.push(res.body);
      lines.push(`--${changeset}--`);
    }

    lines.push(`--${boundary}--`);
    lines.push("");

    return {
      status: 200,
      headers: { "Content-Type": `multipart/mixed; boundary=${boundary}` },
      body: lines.join("\r\n"),
    };
  }

  function statusText(status: number): string {
    return (
      {
        200: "OK",
        201: "Created",
        204: "No Content",
        400: "Bad Request",
        401: "Unauthorized",
        403: "Forbidden",
        404: "Not Found",
        405: "Method Not Allowed",
        412: "Precondition Failed",
        429: "Too Many Requests",
        500: "Internal Server Error",
        503: "Service Unavailable",
        504: "Gateway Timeout",
      }[status] ?? "Unknown"
    );
  }

  function headerOf(headers: Record<string, string>, name: string): string | undefined {
    return headers[name.toLowerCase()];
  }

  /* --- fetch shim --------------------------------------------------------- */

  const fetchImpl = (async (input: RequestInfo | URL, init?: RequestInit): Promise<Response> => {
    const url = typeof input === "string" ? input : input instanceof URL ? input.toString() : String(input);

    if (init?.signal?.aborted) {
      throw Object.assign(new Error("The operation was aborted"), { name: "AbortError" });
    }

    if (networkFailRemaining > 0) {
      networkFailRemaining--;
      const cause = Object.assign(new Error(`socket hang up`), { code: networkFailCode });
      throw Object.assign(new TypeError("fetch failed"), { cause });
    }

    if (throttleRemaining > 0) {
      throttleRemaining--;
      const headers: Record<string, string> = { "Content-Type": "application/json" };
      if (throttleRetryAfter !== undefined) headers["Retry-After"] = String(throttleRetryAfter);
      recordRequest(url, init, throttleStatus, undefined);
      return new Response(
        JSON.stringify({ error: { code: "0x80072322", message: "Rate limit exceeded." } }),
        { status: throttleStatus, headers }
      );
    }

    const apiBase = `${apiRoot()}/`;
    if (!url.startsWith(apiBase)) {
      return new Response(JSON.stringify({ error: { message: `Unexpected URL ${url}` } }), {
        status: 404,
        headers: { "Content-Type": "application/json" },
      });
    }

    const headers = normalizeHeaders(init?.headers);
    if (!/^Bearer\s+.+/.test(headers["authorization"] ?? "")) {
      recordRequest(url, init, 401, undefined);
      return new Response(
        JSON.stringify({ error: { code: "0x80040001", message: "Authorization header missing." } }),
        { status: 401, headers: { "Content-Type": "application/json" } }
      );
    }

    const req: RawRequest = {
      method: (init?.method ?? "GET").toUpperCase(),
      path: url.slice(apiBase.length),
      headers,
      body: typeof init?.body === "string" ? init.body : "",
    };
    const res = handle(req);
    recordRequest(
      url,
      init,
      res.status,
      req.method === "POST" && req.path === "$batch" ? req.body : undefined
    );

    return new Response(res.body || null, { status: res.status, headers: res.headers });
  }) as typeof fetch;

  function normalizeHeaders(h: HeadersInit | undefined): Record<string, string> {
    const out: Record<string, string> = {};
    if (!h) return out;
    if (h instanceof Headers) {
      h.forEach((v, k) => (out[k.toLowerCase()] = v));
      return out;
    }
    if (Array.isArray(h)) {
      for (const [k, v] of h) out[k.toLowerCase()] = v;
      return out;
    }
    for (const [k, v] of Object.entries(h)) out[k.toLowerCase()] = String(v);
    return out;
  }

  function recordRequest(
    url: string,
    init: RequestInit | undefined,
    status: number,
    batchBody: string | undefined
  ): void {
    const apiBase = `${apiRoot()}/`;
    const entry: ObservedRequest = {
      method: (init?.method ?? "GET").toUpperCase(),
      path: url.startsWith(apiBase) ? url.slice(apiBase.length) : url,
      headers: normalizeHeaders(init?.headers),
      body: typeof init?.body === "string" ? init.body : undefined,
      status,
    };
    if (batchBody) {
      try {
        entry.batchOps = parseBatchRequest(batchBody, apiRoot()).map((o) => ({
          contentId: o.contentId,
          method: o.method,
          url: o.url,
          headers: o.headers,
        }));
      } catch {
        /* recorded requests are for assertions only; a parse failure will
           already have surfaced through the response path */
      }
    }
    requests.push(entry);
  }

  /* --- real HTTP server --------------------------------------------------- */

  let server: Server | undefined;

  return {
    fetch: fetchImpl,
    requests,
    get environmentUrl() {
      return environmentUrl;
    },

    records(entitySet) {
      return [...rowsOf(entitySet).values()];
    },
    seed(entitySet, rows) {
      const def = defFor(entitySet);
      for (const r of rows) {
        const id = String(r[def.primaryIdAttribute] ?? newGuid());
        rowsOf(entitySet).set(id, { statecode: 0, ...r, [def.primaryIdAttribute]: id });
      }
    },

    throttleNext(count, o) {
      throttleRemaining = count;
      throttleStatus = o?.status ?? 429;
      throttleRetryAfter = o?.retryAfter ?? 1;
    },
    failNextNetwork(count, code = "ECONNRESET") {
      networkFailRemaining = count;
      networkFailCode = code;
    },
    setOperationFault(fault) {
      operationFault = fault;
    },
    resetFaults() {
      throttleRemaining = 0;
      networkFailRemaining = 0;
      operationFault = null;
      opSeq = 0;
      requests.length = 0;
    },

    async listen() {
      server = createServer((req, res) => {
        const chunks: Buffer[] = [];
        req.on("data", (c: Buffer) => chunks.push(c));
        req.on("end", () => {
          const prefix = `/api/${"data"}/${apiVersion}/`;
          const path = (req.url ?? "").startsWith(prefix)
            ? (req.url ?? "").slice(prefix.length)
            : (req.url ?? "");
          const headers: Record<string, string> = {};
          for (const [k, v] of Object.entries(req.headers)) {
            headers[k.toLowerCase()] = Array.isArray(v) ? v.join(",") : String(v ?? "");
          }
          const out = handle({
            method: (req.method ?? "GET").toUpperCase(),
            path,
            headers,
            body: Buffer.concat(chunks).toString("utf8"),
          });
          res.writeHead(out.status, out.headers);
          res.end(out.body);
        });
      });

      await new Promise<void>((resolve) => server!.listen(0, "127.0.0.1", resolve));
      const { port } = server!.address() as AddressInfo;
      environmentUrl = `http://127.0.0.1:${port}`;
      return environmentUrl;
    },

    async close() {
      if (!server) return;
      await new Promise<void>((resolve, reject) => server!.close((e) => (e ? reject(e) : resolve())));
      server = undefined;
    },
  };
}

/* -------------------------------------------------------------------------- */
/* Ready-made fixtures                                                         */
/* -------------------------------------------------------------------------- */

/**
 * The `contact` / `account` pair the manual TEST-PROTOCOL exercises, with the
 * same alternate key (`emailaddress1`) the protocol asks you to create by hand
 * in §1.2. Integration tests use this so the automated and manual suites are
 * describing the same environment.
 */
export function contactsFixture(): FakeEntityDef[] {
  return [
    {
      entitySet: "contacts",
      logicalName: "contact",
      primaryIdAttribute: "contactid",
      primaryNameAttribute: "fullname",
      alternateKeys: [["emailaddress1"]],
      attributes: [
        { LogicalName: "emailaddress1", AttributeType: "String" },
        { LogicalName: "firstname", AttributeType: "String" },
        { LogicalName: "lastname", AttributeType: "String" },
        { LogicalName: "birthdate", AttributeType: "DateTime" },
        { LogicalName: "creditlimit", AttributeType: "Money" },
        { LogicalName: "donotemail", AttributeType: "Boolean" },
        {
          LogicalName: "dvlt_tier",
          AttributeType: "Picklist",
          options: { Bronze: 1, Silver: 2, Gold: 3 },
        },
        {
          LogicalName: "parentcustomerid",
          AttributeType: "Lookup",
          targets: ["account"],
        },
      ],
    },
    {
      entitySet: "accounts",
      logicalName: "account",
      primaryIdAttribute: "accountid",
      primaryNameAttribute: "name",
      attributes: [{ LogicalName: "name", AttributeType: "String" }],
    },
  ];
}
