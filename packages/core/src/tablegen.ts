// Generate a new Dataverse table from an Excel source table: infer column
// types from sample values, and build the Web API metadata payloads
// (EntityMetadata, AttributeMetadata, EntityKeyMetadata). Pure functions —
// the add-in drives them and DataverseClient POSTs the results.

/** Subset of field kinds a generated table can contain. */
export type GeneratedKind =
  | "string"
  | "memo"
  | "integer"
  | "decimal"
  | "boolean"
  | "datetime"
  | "dateonly";

export interface GeneratedColumn {
  /** Source column header in the Excel table. */
  source: string;
  /** Label shown in Dataverse. Defaults to the source header. */
  displayName: string;
  /** SchemaName suffix: `<prefix>_<schemaSuffix>`. Letters/digits, starts with a letter. */
  schemaSuffix: string;
  kind: GeneratedKind;
  /** Exactly one included string column must be the primary name attribute. */
  isPrimaryName: boolean;
  /** Part of the table's alternate key (string/integer/decimal only). */
  inAlternateKey: boolean;
  /** Whether to create this column at all. */
  include: boolean;
  /** For strings: MaxLength to create the attribute with. */
  maxLength?: number;
  /**
   * Map this source column to an EXISTING system attribute instead of
   * creating a new one (e.g. "overriddencreatedon" to backdate createdon
   * during the initial load, or "ownerid" for per-row owners). Excluded
   * from attribute creation, primary name, and alternate keys.
   */
  systemAttribute?: string;
  /** For systemAttribute="ownerid": which entity set the GUIDs point at ("systemusers" or "teams"). */
  systemBindEntitySet?: string;
}

/** Kinds Dataverse supports as alternate-key members. */
export const KEYABLE_KINDS: readonly GeneratedKind[] = ["string", "integer", "decimal"];

const LANG = 1033;

function label(text: string): Record<string, unknown> {
  return {
    "@odata.type": "Microsoft.Dynamics.CRM.Label",
    LocalizedLabels: [
      { "@odata.type": "Microsoft.Dynamics.CRM.LocalizedLabel", Label: text, LanguageCode: LANG },
    ],
  };
}

/** "First Name (home)" → "FirstNameHome"; guarantees a letter start, non-empty. */
export function sanitizeSchemaSuffix(header: string): string {
  const words = header.split(/[^A-Za-z0-9]+/).filter(Boolean);
  let s = words.map((w) => w[0].toUpperCase() + w.slice(1)).join("");
  s = s.replace(/^[^A-Za-z]+/, "");
  return s || "Column";
}

const TRUEISH = new Set(["true", "yes", "y"]);
const FALSEISH = new Set(["false", "no", "n"]);
const ISO_DATE_RE = /^\d{4}-\d{2}-\d{2}([T ].+)?$/;

/** Infer a column's kind from its non-blank sample values. */
export function inferColumnKind(values: unknown[]): { kind: GeneratedKind; maxLength?: number } {
  const sample = values.filter(
    (v) => v !== null && v !== undefined && !(typeof v === "string" && v.trim() === "")
  );
  if (sample.length === 0) return { kind: "string", maxLength: 100 };

  const asBool = sample.every((v) => {
    if (typeof v === "boolean") return true;
    const s = String(v).trim().toLowerCase();
    return TRUEISH.has(s) || FALSEISH.has(s);
  });
  if (asBool) return { kind: "boolean" };

  const nums = sample.map((v) =>
    typeof v === "number" ? v : Number(String(v).trim())
  );
  if (nums.every((n) => Number.isFinite(n))) {
    const allInt = nums.every((n) => Number.isInteger(n) && Math.abs(n) <= 2147483647);
    return { kind: allInt ? "integer" : "decimal" };
  }

  const dates = sample.map((v) => {
    if (v instanceof Date) return { ok: true, hasTime: v.getHours() + v.getMinutes() + v.getSeconds() !== 0 };
    const s = String(v).trim();
    if (!ISO_DATE_RE.test(s) || Number.isNaN(Date.parse(s))) return { ok: false, hasTime: false };
    return { ok: true, hasTime: s.length > 10 };
  });
  if (dates.every((d) => d.ok)) {
    return { kind: dates.some((d) => d.hasTime) ? "datetime" : "dateonly" };
  }

  const maxLen = Math.max(...sample.map((v) => String(v).length));
  if (maxLen > 400) return { kind: "memo" };
  return { kind: "string", maxLength: Math.min(Math.max(100, maxLen * 2), 4000) };
}

/**
 * Suggest a full column set for a new table. The first string column becomes
 * the primary name attribute (every Dataverse table needs one).
 */
export function suggestColumns(
  headers: string[],
  rows: Array<Record<string, unknown>>
): GeneratedColumn[] {
  const seen = new Set<string>();
  const cols = headers.map((h): GeneratedColumn => {
    const { kind, maxLength } = inferColumnKind(rows.map((r) => r[h]));
    let suffix = sanitizeSchemaSuffix(h);
    while (seen.has(suffix.toLowerCase())) suffix += "1";
    seen.add(suffix.toLowerCase());
    return {
      source: h,
      displayName: h,
      schemaSuffix: suffix,
      kind,
      isPrimaryName: false,
      inAlternateKey: false,
      include: true,
      maxLength,
    };
  });
  const firstString = cols.find((c) => c.kind === "string");
  if (firstString) firstString.isPrimaryName = true;
  return cols;
}

export function attributeSchemaName(prefix: string, col: GeneratedColumn): string {
  return `${prefix}_${col.schemaSuffix}`;
}

export function attributeLogicalName(prefix: string, col: GeneratedColumn): string {
  return attributeSchemaName(prefix, col).toLowerCase();
}

/** Attribute metadata payload for one generated column. */
export function buildAttributePayload(
  prefix: string,
  col: GeneratedColumn
): Record<string, unknown> {
  const base = {
    SchemaName: attributeSchemaName(prefix, col),
    DisplayName: label(col.displayName),
    RequiredLevel: { Value: "None" },
  };
  switch (col.kind) {
    case "string":
      return {
        ...base,
        "@odata.type": "Microsoft.Dynamics.CRM.StringAttributeMetadata",
        FormatName: { Value: "Text" },
        MaxLength: col.maxLength ?? 200,
        ...(col.isPrimaryName ? { IsPrimaryName: true } : {}),
      };
    case "memo":
      return {
        ...base,
        "@odata.type": "Microsoft.Dynamics.CRM.MemoAttributeMetadata",
        MaxLength: 100000,
      };
    case "integer":
      return {
        ...base,
        "@odata.type": "Microsoft.Dynamics.CRM.IntegerAttributeMetadata",
        Format: "None",
        MinValue: -2147483648,
        MaxValue: 2147483647,
      };
    case "decimal":
      return {
        ...base,
        "@odata.type": "Microsoft.Dynamics.CRM.DecimalAttributeMetadata",
        Precision: 2,
        MinValue: -100000000000,
        MaxValue: 100000000000,
      };
    case "boolean":
      return {
        ...base,
        "@odata.type": "Microsoft.Dynamics.CRM.BooleanAttributeMetadata",
        OptionSet: {
          "@odata.type": "Microsoft.Dynamics.CRM.BooleanOptionSetMetadata",
          TrueOption: { Value: 1, Label: label("Yes") },
          FalseOption: { Value: 0, Label: label("No") },
        },
      };
    case "datetime":
      return {
        ...base,
        "@odata.type": "Microsoft.Dynamics.CRM.DateTimeAttributeMetadata",
        Format: "DateAndTime",
      };
    case "dateonly":
      return {
        ...base,
        "@odata.type": "Microsoft.Dynamics.CRM.DateTimeAttributeMetadata",
        Format: "DateOnly",
      };
  }
}

/**
 * Entity metadata payload. Includes ONLY the primary name attribute — the
 * remaining columns are created afterwards with createAttribute (attribute
 * creation inside the entity POST is limited to the primary name).
 */
export function buildEntityPayload(opts: {
  prefix: string;
  schemaSuffix: string;
  displayName: string;
  primaryNameColumn: GeneratedColumn;
}): Record<string, unknown> {
  return {
    "@odata.type": "Microsoft.Dynamics.CRM.EntityMetadata",
    SchemaName: `${opts.prefix}_${opts.schemaSuffix}`,
    DisplayName: label(opts.displayName),
    DisplayCollectionName: label(opts.displayName),
    Description: label(`Created by dvload from Excel table.`),
    OwnershipType: "UserOwned",
    HasNotes: false,
    HasActivities: false,
    Attributes: [buildAttributePayload(opts.prefix, opts.primaryNameColumn)],
  };
}

/** Alternate-key payload from the columns flagged inAlternateKey. */
export function buildKeyPayload(
  prefix: string,
  entitySchemaSuffix: string,
  keyColumns: GeneratedColumn[]
): Record<string, unknown> {
  const attrs = keyColumns.map((c) => attributeLogicalName(prefix, c));
  return {
    SchemaName: `${prefix}_key_${entitySchemaSuffix.toLowerCase()}`,
    DisplayName: label(`${keyColumns.map((c) => c.displayName).join(" + ")} (alternate key)`),
    KeyAttributes: attrs,
  };
}
