// DataverseGateway implementation for the Power Platform ToolBox (PPTB).
//
// PPTB tools run in a sandboxed iframe and are deliberately never given an
// access token — every Dataverse operation goes through the ToolBox's
// window.dataverseAPI bridge, which holds the connection server-side. So the
// direct OData client (core's DataverseClient) cannot run here at all; this
// adapter maps the same gateway surface onto the bridge instead, and the
// engine (loadRows) runs unchanged on top of it.
//
// The semantic mapping worth knowing about:
//
//  - batch(): core's $batch wraps each operation in its own changeset, so
//    operations already execute independently. The bridge has no $batch, so
//    this adapter executes the operations individually (bounded parallelism)
//    and synthesises BatchResultItems with the statuses loadRows accounts on:
//    201 created / 204 updated / 412 for a matched If-None-Match.
//  - PATCH upserts are emulated: probe for the record (one $filter read),
//    then update or create. A true server-side upsert is atomic; this is not,
//    so a record created between probe and write surfaces as a per-row error
//    rather than being silently overwritten.
//  - bypassCustomLogic / impersonateUserId need per-request headers the
//    bridge cannot send. Operations carrying them are refused loudly.
//  - createEntityKey (alternate keys) has no bridge equivalent and throws.
//
// Everything here is bridge-in, gateway-out: no DOM, no toolboxAPI, so the
// whole class is unit-testable with a fake PptbDataverseApi.

import {
  assertGuid,
  assertLogicalName,
  formatKeyLiteral,
  type BatchOperation,
  type BatchResultItem,
  type CreateResult,
  type DataverseGateway,
  type RequestLogEntry,
} from "@dvload/core";
import type { PptbDataverseApi } from "./pptb-bridge.js";

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/** How many bridge operations one batch() call keeps in flight at once. */
const OP_CONCURRENCY = 4;

/** Per-operation headers the bridge can honour. Everything else is refused. */
const SUPPORTED_OP_HEADERS = new Set(["if-none-match"]);

export interface PptbClientOptions {
  dataverse: PptbDataverseApi;
  /** Called after every bridge call, mirroring DataverseClient's onRequest. */
  onRequest?: (entry: RequestLogEntry) => void;
}

interface EntityInfo {
  logicalName: string;
  entitySetName: string;
  primaryIdAttribute: string;
  primaryNameAttribute?: string;
  displayName: string;
  metadataId: string;
}

/** One parsed entity-set-relative operation URL, e.g. "contacts(<guid>)". */
type ParsedKey =
  | { kind: "none" }
  | { kind: "guid"; guid: string }
  | { kind: "alternate"; pairs: Array<{ attribute: string; value: string | number | boolean }> };

export class PptbDataverseClient implements DataverseGateway {
  private readonly api: PptbDataverseApi;
  private readonly onRequest?: (entry: RequestLogEntry) => void;
  private entityInfoPromise: Promise<EntityInfo[]> | null = null;

  constructor(opts: PptbClientOptions) {
    this.api = opts.dataverse;
    this.onRequest = opts.onRequest;
  }

  /** Run one bridge call, reporting it to onRequest like an HTTP request. */
  private async call<T>(method: string, what: string, okStatus: number, fn: () => Promise<T>): Promise<T> {
    const startedAt = new Date().toISOString();
    try {
      const result = await fn();
      this.onRequest?.({
        method,
        url: `pptb:${what}`,
        status: okStatus,
        ok: true,
        startedAt,
        finishedAt: new Date().toISOString(),
      });
      return result;
    } catch (e) {
      this.onRequest?.({
        method,
        url: `pptb:${what}`,
        status: statusFromError(e),
        ok: false,
        startedAt,
        finishedAt: new Date().toISOString(),
        errorBody: e instanceof Error ? e.message : String(e),
      });
      throw e;
    }
  }

  /* ------------------------------------------------------------------------ */
  /* Entity map: one metadata listing serves listEntities + getEntitySetInfo   */
  /* ------------------------------------------------------------------------ */

  private entityInfos(): Promise<EntityInfo[]> {
    this.entityInfoPromise ??= this.call("GET", "EntityDefinitions", 200, () =>
      this.api.getAllEntitiesMetadata([
        "LogicalName",
        "EntitySetName",
        "DisplayName",
        "MetadataId",
        "PrimaryIdAttribute",
        "PrimaryNameAttribute",
      ])
    ).then(
      (res) =>
        res.value
          .map((e) => ({
            logicalName: String(e.LogicalName ?? ""),
            entitySetName: String(e.EntitySetName ?? ""),
            primaryIdAttribute: String(e.PrimaryIdAttribute ?? ""),
            primaryNameAttribute: e.PrimaryNameAttribute ? String(e.PrimaryNameAttribute) : undefined,
            displayName: extractLocalizedLabel(e.DisplayName),
            metadataId: String(e.MetadataId ?? ""),
          }))
          .filter((e) => e.logicalName && e.entitySetName),
      (e) => {
        // A failed listing must not be memoised as "no entities".
        this.entityInfoPromise = null;
        throw e;
      }
    );
    return this.entityInfoPromise;
  }

  private async infoForSet(entitySetName: string): Promise<EntityInfo> {
    assertLogicalName(entitySetName, "Entity set");
    const info = (await this.entityInfos()).find((e) => e.entitySetName === entitySetName);
    if (!info || !info.primaryIdAttribute) {
      throw new Error(`No entity found with EntitySetName "${entitySetName}"`);
    }
    return info;
  }

  /* ------------------------------------------------------------------------ */
  /* Metadata reads                                                            */
  /* ------------------------------------------------------------------------ */

  async listEntities(): Promise<
    Array<{ LogicalName: string; EntitySetName: string; DisplayName: string; MetadataId: string }>
  > {
    return (await this.entityInfos()).map((e) => ({
      LogicalName: e.logicalName,
      EntitySetName: e.entitySetName,
      DisplayName: e.displayName,
      MetadataId: e.metadataId,
    }));
  }

  async getEntitySetInfo(
    entitySetName: string
  ): Promise<{ logicalName: string; primaryIdAttribute: string; primaryNameAttribute?: string }> {
    const info = await this.infoForSet(entitySetName);
    return {
      logicalName: info.logicalName,
      primaryIdAttribute: info.primaryIdAttribute,
      primaryNameAttribute: info.primaryNameAttribute,
    };
  }

  async getEntityDefinition(logicalName: string): Promise<Record<string, unknown>> {
    assertLogicalName(logicalName, "Entity logical name");
    // The direct client asks for $expand=Attributes in one request; the
    // bridge splits entity and related metadata into two calls.
    const [def, attrs] = await Promise.all([
      this.call("GET", `EntityDefinitions(${logicalName})`, 200, () =>
        this.api.getEntityMetadata(logicalName, true)
      ),
      this.call("GET", `EntityDefinitions(${logicalName})/Attributes`, 200, () =>
        this.api.getEntityRelatedMetadata(logicalName, "Attributes")
      ),
    ]);
    return { ...def, Attributes: (attrs as { value?: unknown }).value ?? [] };
  }

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

  async getSolutionEntityIds(solutionId: string): Promise<Set<string>> {
    assertGuid(solutionId, "Solution id");
    const rows = await this.queryAll(
      `solutioncomponents?$select=objectid&$filter=_solutionid_value eq ${solutionId} and componenttype eq 1`
    );
    return new Set(
      rows.map((r) => String(r.objectid ?? "").toLowerCase()).filter((id) => GUID_RE.test(id))
    );
  }

  async getLookupTargets(entityLogicalName: string, attrLogicalName: string): Promise<string[]> {
    assertLogicalName(entityLogicalName, "Entity logical name");
    assertLogicalName(attrLogicalName, "Attribute logical name");
    // The bridge's related-metadata path takes no $filter, so fetch the
    // lookup-cast collection and match client-side.
    const res = await this.call(
      "GET",
      `EntityDefinitions(${entityLogicalName})/Attributes/Lookup`,
      200,
      () =>
        this.api.getEntityRelatedMetadata(
          entityLogicalName,
          "Attributes/Microsoft.Dynamics.CRM.LookupAttributeMetadata",
          ["LogicalName", "Targets"]
        )
    );
    const rows = ((res as { value?: Array<Record<string, unknown>> }).value ?? []).find(
      (a) => String(a.LogicalName ?? "") === attrLogicalName
    );
    return Array.isArray(rows?.Targets) ? (rows.Targets as string[]).map(String) : [];
  }

  async getLookupNavigationProperty(
    entityLogicalName: string,
    attrLogicalName: string,
    referencedEntityLogicalName: string
  ): Promise<string | undefined> {
    assertLogicalName(entityLogicalName, "Entity logical name");
    assertLogicalName(attrLogicalName, "Attribute logical name");
    assertLogicalName(referencedEntityLogicalName, "Referenced entity logical name");
    const res = await this.call(
      "GET",
      `EntityDefinitions(${entityLogicalName})/ManyToOneRelationships`,
      200,
      () =>
        this.api.getEntityRelatedMetadata(entityLogicalName, "ManyToOneRelationships", [
          "ReferencingAttribute",
          "ReferencedEntity",
          "ReferencingEntityNavigationPropertyName",
        ])
    );
    const match = ((res as { value?: Array<Record<string, unknown>> }).value ?? []).find(
      (r) =>
        String(r.ReferencingAttribute ?? "") === attrLogicalName &&
        String(r.ReferencedEntity ?? "") === referencedEntityLogicalName
    );
    const prop = match ? String(match.ReferencingEntityNavigationPropertyName ?? "") : "";
    return prop || undefined;
  }

  async getEntityKeys(
    entityLogicalName: string
  ): Promise<Array<{ LogicalName: string; KeyAttributes: string[] }>> {
    assertLogicalName(entityLogicalName, "Entity logical name");
    const res = await this.call("GET", `EntityDefinitions(${entityLogicalName})/Keys`, 200, () =>
      this.api.getEntityRelatedMetadata(entityLogicalName, "Keys", ["LogicalName", "KeyAttributes"])
    );
    return ((res as { value?: Array<Record<string, unknown>> }).value ?? []).map((k) => ({
      LogicalName: String(k.LogicalName ?? ""),
      KeyAttributes: Array.isArray(k.KeyAttributes) ? (k.KeyAttributes as unknown[]).map(String) : [],
    }));
  }

  async getOptionSetLabels(
    entityLogicalName: string,
    attrLogicalName: string
  ): Promise<Record<string, number>> {
    assertLogicalName(entityLogicalName, "Entity logical name");
    assertLogicalName(attrLogicalName, "Attribute logical name");
    // Local option set first, then the global one — same order of preference
    // as the direct client's $expand of both.
    for (const nav of ["OptionSet", "GlobalOptionSet"]) {
      let res: Record<string, unknown>;
      try {
        res = await this.call(
          "GET",
          `EntityDefinitions(${entityLogicalName})/Attributes(${attrLogicalName})/${nav}`,
          200,
          () =>
            this.api.getEntityRelatedMetadata(
              entityLogicalName,
              `Attributes(LogicalName='${attrLogicalName}')/${nav}`
            )
        );
      } catch {
        continue; // not this option-set flavour; try the next
      }
      const options = (res as { Options?: Array<Record<string, unknown>> }).Options;
      if (!Array.isArray(options)) continue;
      const map: Record<string, number> = {};
      for (const o of options) {
        const label = (o.Label as { UserLocalizedLabel?: { Label?: string } } | undefined)
          ?.UserLocalizedLabel?.Label;
        if (label && typeof o.Value === "number") map[label] = o.Value;
      }
      if (Object.keys(map).length > 0) return map;
    }
    throw new Error(
      `No option-set metadata found for ${entityLogicalName}.${attrLogicalName} ` +
        `(is it a choice/multichoice/status/state attribute?)`
    );
  }

  /* ------------------------------------------------------------------------ */
  /* Table generation                                                          */
  /* ------------------------------------------------------------------------ */

  async createEntity(payload: Record<string, unknown>): Promise<void> {
    await this.call("POST", "EntityDefinitions", 201, () => this.api.createEntityDefinition(payload));
  }

  async createAttribute(entityLogicalName: string, payload: Record<string, unknown>): Promise<void> {
    assertLogicalName(entityLogicalName, "Entity logical name");
    await this.call("POST", `EntityDefinitions(${entityLogicalName})/Attributes`, 201, () =>
      this.api.createAttribute(entityLogicalName, payload)
    );
  }

  async createEntityKey(): Promise<void> {
    throw new Error(
      "Creating alternate keys isn't supported by the Power Platform ToolBox bridge. " +
        "Create the key in Power Apps (Table → Keys), or use the Excel add-in / CLI."
    );
  }

  /* ------------------------------------------------------------------------ */
  /* Queries                                                                   */
  /* ------------------------------------------------------------------------ */

  async queryAll(path: string): Promise<Array<Record<string, unknown>>> {
    const out: Array<Record<string, unknown>> = [];
    let query: string | null = path;
    // Hard stop so a bridge that echoes the same nextLink forever can't hang.
    for (let page = 0; query && page < 1000; page++) {
      const q: string = query;
      const res = await this.call("GET", q, 200, () => this.api.queryData(q));
      out.push(...(res.value ?? []));
      const next: unknown = res["@odata.nextLink"];
      // The bridge returns the absolute nextLink; queryData wants the
      // entity-set-relative form, so strip everything through /api/data/vX.Y/.
      query = typeof next === "string" ? relativizeODataUrl(next) : null;
    }
    return out;
  }

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

    // Same chunking as the direct client: keeps URLs comfortably short.
    const CHUNK = 15;
    for (let start = 0; start < values.length; start += CHUNK) {
      const chunk = values.slice(start, start + CHUNK);
      const filter = chunk.map((v) => `${attribute} eq ${formatKeyLiteral(v)}`).join(" or ");
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

  async resolveByKey(entitySet: string, keyAttribute: string, value: unknown): Promise<string | null> {
    assertLogicalName(entitySet, "Entity set");
    assertLogicalName(keyAttribute, "Key attribute");
    const info = await this.infoForSet(entitySet);
    const rows = await this.queryAll(
      `${entitySet}?$select=${info.primaryIdAttribute}&$filter=${keyAttribute} eq ${formatKeyLiteral(value)}`
    );
    if (rows.length === 0) return null;
    if (rows.length > 1) {
      // A real alternate key is unique; more than one match means the chosen
      // attribute isn't one. Failing beats binding to an arbitrary record.
      throw new Error(
        `${entitySet}.${keyAttribute} = ${JSON.stringify(String(value))} matched ${rows.length} records — not a unique key`
      );
    }
    const id = String(rows[0][info.primaryIdAttribute] ?? "");
    return GUID_RE.test(id) ? id.toLowerCase() : null;
  }

  async getRecord(
    entitySet: string,
    keyExpr: string,
    select: string[]
  ): Promise<Record<string, unknown> | null> {
    assertLogicalName(entitySet, "Entity set");
    for (const s of select) assertLogicalName(s, "Select attribute");
    const info = await this.infoForSet(entitySet);
    const filter = this.filterForKey(info, parseKeyExpression(keyExpr));
    const rows = await this.queryAll(
      `${entitySet}?$select=${select.join(",")}&$filter=${filter}`
    );
    // 0 → not found; 2+ → ambiguous. Both fall back to "send everything",
    // which is the safe direction for skipUnchanged.
    return rows.length === 1 ? rows[0] : null;
  }

  /* ------------------------------------------------------------------------ */
  /* Writes                                                                    */
  /* ------------------------------------------------------------------------ */

  async create(
    entitySet: string,
    body: Record<string, unknown>,
    extraHeaders?: Record<string, string>
  ): Promise<CreateResult> {
    assertSupportedHeaders(extraHeaders);
    const info = await this.infoForSet(entitySet);
    const res = await this.call("POST", entitySet, 201, () =>
      this.api.create(info.logicalName, body)
    );
    return { id: String(res.id ?? ""), entityUrl: "" };
  }

  async batch(operations: BatchOperation[]): Promise<BatchResultItem[]> {
    if (operations.length === 0) return [];

    const results = new Array<BatchResultItem>(operations.length);
    let next = 0;
    const worker = async (): Promise<void> => {
      for (;;) {
        const i = next++;
        if (i >= operations.length) return;
        results[i] = await this.executeOperation(operations[i]);
      }
    };
    const poolSize = Math.min(OP_CONCURRENCY, operations.length);
    await Promise.all(Array.from({ length: poolSize }, worker));

    return results.sort((a, b) => a.contentId - b.contentId);
  }

  /** One batch operation → one bridge call sequence → one result item. */
  private async executeOperation(op: BatchOperation): Promise<BatchResultItem> {
    try {
      const skipIfExists = hasHeader(op.headers, "if-none-match");
      assertSupportedHeaders(op.headers);

      const { entitySet, key } = parseOperationUrl(op.url);
      const info = await this.infoForSet(entitySet);
      const body = (op.body ?? {}) as Record<string, unknown>;

      switch (op.method) {
        case "POST": {
          const res = await this.call("POST", op.url, 201, () =>
            this.api.create(info.logicalName, body)
          );
          return ok(op, 201, String(res.id ?? "") || undefined);
        }

        case "DELETE": {
          if (key.kind !== "guid") {
            throw new Error(`DELETE needs a GUID key, got "${op.url}"`);
          }
          const guid = key.guid;
          await this.call("DELETE", op.url, 204, () => this.api.delete(info.logicalName, guid));
          return ok(op, 204);
        }

        case "PATCH": {
          // Probe first: the bridge has no atomic upsert and no way to send
          // If-None-Match, so existence decides update / create / 412.
          const existingId = await this.findExisting(entitySet, info, key);

          if (existingId && skipIfExists) {
            // The outcome If-None-Match: * requests — loadRows counts a 412
            // under skip-if-exists as the skip the user asked for.
            return {
              contentId: op.contentId,
              status: 412,
              ok: false,
              errorMessage: "record already exists",
            };
          }

          if (existingId) {
            await this.call("PATCH", op.url, 204, () =>
              this.api.update(info.logicalName, existingId, body)
            );
            return ok(op, 204, existingId);
          }

          // Not found → create, like a real PATCH-by-key upsert would.
          // A primary-id upsert carries the id only in the URL (the engine
          // strips it from the body), so put it back for the create.
          const createBody =
            key.kind === "guid" ? { ...body, [info.primaryIdAttribute]: key.guid } : body;
          const res = await this.call("POST", op.url, 201, () =>
            this.api.create(info.logicalName, stripLookupClears(createBody))
          );
          return ok(op, 201, String(res.id ?? "") || undefined);
        }

        default:
          throw new Error(`Unsupported batch method: ${String(op.method)}`);
      }
    } catch (e) {
      return {
        contentId: op.contentId,
        status: statusFromError(e),
        ok: false,
        errorMessage: e instanceof Error ? e.message : String(e),
      };
    }
  }

  /** The existing record's id for a PATCH target, or null. */
  private async findExisting(
    entitySet: string,
    info: EntityInfo,
    key: ParsedKey
  ): Promise<string | null> {
    if (key.kind === "none") throw new Error("PATCH needs a key expression in the URL");
    const filter = this.filterForKey(info, key);
    const rows = await this.queryAll(
      `${entitySet}?$select=${info.primaryIdAttribute}&$filter=${filter}`
    );
    if (rows.length > 1) {
      throw new Error(`upsert key matched ${rows.length} records — the key is not unique`);
    }
    const id = String(rows[0]?.[info.primaryIdAttribute] ?? "");
    return GUID_RE.test(id) ? id.toLowerCase() : null;
  }

  private filterForKey(info: EntityInfo, key: ParsedKey): string {
    if (key.kind === "guid") return `${info.primaryIdAttribute} eq ${key.guid}`;
    if (key.kind === "alternate") {
      return key.pairs
        .map((p) => {
          assertLogicalName(p.attribute, "Key attribute");
          return `${p.attribute} eq ${formatKeyLiteral(p.value)}`;
        })
        .join(" and ");
    }
    throw new Error("missing key expression");
  }
}

/* -------------------------------------------------------------------------- */
/* Helpers                                                                     */
/* -------------------------------------------------------------------------- */

function ok(op: BatchOperation, status: number, id?: string): BatchResultItem {
  return { contentId: op.contentId, status, ok: true, ...(id ? { id } : {}) };
}

function hasHeader(headers: Record<string, string> | undefined, name: string): boolean {
  return Object.keys(headers ?? {}).some((k) => k.toLowerCase() === name);
}

/**
 * Refuse per-operation headers the bridge cannot send, instead of silently
 * running without them: MSCRM.BypassCustomPluginExecution or MSCRMCallerID
 * being dropped would change what the import DOES, not just how it's logged.
 */
function assertSupportedHeaders(headers: Record<string, string> | undefined): void {
  for (const k of Object.keys(headers ?? {})) {
    if (!SUPPORTED_OP_HEADERS.has(k.toLowerCase())) {
      throw new Error(
        `The "${k}" request header (bypass custom logic / run as user) isn't supported ` +
          `in the Power Platform ToolBox — the ToolBox bridge can't send per-request headers. ` +
          `Use the Excel add-in or the CLI for this option.`
      );
    }
  }
}

/** "contacts", "contacts(<guid>)" or "contacts(attr='v',attr2=2)" → parts. */
export function parseOperationUrl(url: string): { entitySet: string; key: ParsedKey } {
  const m = /^([A-Za-z_][A-Za-z0-9_]*)(?:\((.*)\))?$/.exec(url);
  if (!m) throw new Error(`Unrecognised operation URL: "${url}"`);
  const entitySet = m[1];
  if (m[2] === undefined) return { entitySet, key: { kind: "none" } };
  return { entitySet, key: parseKeyExpression(m[2]) };
}

/**
 * Parse a key expression the engine built with buildKeyExpression: a bare
 * GUID, or comma-separated `attr=<literal>` pairs whose string literals were
 * produced by formatKeyLiteral (quoted, quote-doubled, percent-encoded — so
 * a literal can never contain a comma or an equals sign).
 */
export function parseKeyExpression(expr: string): ParsedKey {
  const trimmed = expr.trim();
  if (GUID_RE.test(trimmed)) return { kind: "guid", guid: trimmed.toLowerCase() };

  const pairs = trimmed.split(",").map((part) => {
    const eq = part.indexOf("=");
    if (eq <= 0) throw new Error(`Unrecognised key expression part: "${part}"`);
    const attribute = part.slice(0, eq).trim();
    const literal = part.slice(eq + 1).trim();
    return { attribute, value: decodeKeyLiteral(literal) };
  });
  if (pairs.length === 0) throw new Error(`Unrecognised key expression: "${expr}"`);
  return { kind: "alternate", pairs };
}

/** Reverse of formatKeyLiteral, for re-encoding a key value into a $filter. */
function decodeKeyLiteral(literal: string): string | number | boolean {
  if (literal === "true") return true;
  if (literal === "false") return false;
  if (/^-?\d+(\.\d+)?$/.test(literal)) return Number(literal);
  if (literal.startsWith("'") && literal.endsWith("'") && literal.length >= 2) {
    return decodeURIComponent(literal.slice(1, -1)).replace(/''/g, "'");
  }
  throw new Error(`Unrecognised key literal: "${literal}"`);
}

/** Absolute @odata.nextLink → the entity-set-relative query queryData wants. */
function relativizeODataUrl(url: string): string | null {
  const m = /\/api\/data\/v\d+\.\d+\/(.+)$/.exec(url);
  return m ? m[1] : null;
}

/** DisplayName metadata → the user-localised label, mirroring core. */
function extractLocalizedLabel(displayName: unknown): string {
  if (!displayName || typeof displayName !== "object") return "";
  const dn = displayName as { UserLocalizedLabel?: { Label?: string } };
  return dn.UserLocalizedLabel?.Label ?? "";
}

/** Best-effort HTTP status from a bridge error message; 0 when unknowable. */
function statusFromError(e: unknown): number {
  const message = e instanceof Error ? e.message : String(e);
  const m = /\b(4\d{2}|5\d{2})\b/.exec(message);
  return m ? Number(m[1]) : 0;
}

/**
 * A create has nothing to clear: the engine only adds `lookup: null` clears
 * for operations that may update an existing record, but an emulated upsert
 * that falls through to create would send them — and Dataverse rejects a
 * null single-valued navigation property on create.
 */
function stripLookupClears(body: Record<string, unknown>): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(body)) {
    if (v === null && !k.includes("@")) continue;
    out[k] = v;
  }
  return out;
}
