import { describe, expect, it } from "vitest";
import {
  dataflowToPqtArchive,
  detailFromRow,
  listDataflows,
  mappingFromDataflow,
  mappingsFromDataflow,
  parseMashupSettings,
  readDataflowState,
  readPrimaryKeyLogicalName,
  summarizeDataflow,
  type DataflowMetadataResolver,
} from "./dataflow.js";
import { validateMapping } from "./mapping.js";
import { buildWorkbookWithQueries } from "./pqt.js";
import {
  AI_MATCH_ACTIVE,
  AI_MATCH_DRAFT,
  ALL_ROWS,
  COMPANIES_ACTIVE,
  NEVER_PUBLISHED,
} from "../test/support/dataflow-rows.js";

const ENV = "https://contoso.crm4.dynamics.com";

/** Stands in for the Dataverse metadata endpoints. */
const resolver: DataflowMetadataResolver = {
  async entitySetName(logical) {
    return { account: "accounts", crfea_asker_customertype: "crfea_asker_customertypes" }[logical];
  },
  async keyAttributes(entity, key) {
    // The whole point of this lookup: the key is named asker_actid but
    // spans an attribute called asker_act_id.
    if (entity === "account" && key === "asker_actid") return ["asker_act_id"];
    return [];
  },
  async lookupEntitySet(entity, attr) {
    if (entity === "account" && attr === "parentaccountid") return "accounts";
    return undefined;
  },
};

/** Minimal stand-in for DataverseClient.queryAll. */
function fakeClient(rows: unknown[]) {
  return { queryAll: async () => rows } as never;
}

describe("parseMashupSettings", () => {
  it("reads QueriesMetadata out of the settings blob", () => {
    const s = parseMashupSettings(AI_MATCH_ACTIVE.msdyn_mashupsettings);
    expect(Object.keys(s.QueriesMetadata)).toEqual(["Full_Match_Table"]);
    expect(s.QueriesMetadata.Full_Match_Table.EntityName).toBe("Account");
  });

  it("degrades to empty rather than throwing on junk", () => {
    for (const bad of [null, undefined, "", "not json", "[1,2,3]", '"a string"']) {
      expect(parseMashupSettings(bad).QueriesMetadata).toEqual({});
    }
  });
});

describe("doubly-encoded JSON", () => {
  it("reads DataflowState out of the nested DataflowMetadata string", () => {
    expect(readDataflowState(parseMashupSettings(AI_MATCH_ACTIVE.msdyn_mashupsettings))).toBe("Active");
    expect(readDataflowState(parseMashupSettings(AI_MATCH_DRAFT.msdyn_mashupsettings))).toBe("Source");
  });

  it("reports Unknown when DataflowMetadata is absent", () => {
    expect(readDataflowState(parseMashupSettings("{}"))).toBe("Unknown");
  });

  it("reads primaryKeyLogicalName out of the nested MappingEditorProperties string", () => {
    const s = parseMashupSettings(AI_MATCH_ACTIVE.msdyn_mashupsettings);
    expect(readPrimaryKeyLogicalName(s.QueriesMetadata.Full_Match_Table)).toBe("asker_actid");
  });

  it("returns undefined when the editor recorded no key", () => {
    const s = parseMashupSettings(COMPANIES_ACTIVE.msdyn_mashupsettings);
    expect(readPrimaryKeyLogicalName(s.QueriesMetadata.asker_customertype)).toBeUndefined();
  });
});

describe("summarizeDataflow", () => {
  it("separates load targets from staging queries", () => {
    const s = summarizeDataflow(COMPANIES_ACTIVE);
    expect(s.queryNames).toHaveLength(3);
    expect(s.loadTargets).toEqual([{ queryName: "Companies", entityName: "Account", fieldCount: 4 }]);
  });
});

describe("listDataflows", () => {
  it("hides editing drafts that duplicate a published dataflow", async () => {
    const list = await listDataflows(fakeClient(ALL_ROWS));
    const names = list.map((d) => d.name);
    // Both AI_MATCH rows share a name; only the Active one survives.
    expect(names.filter((n) => n === AI_MATCH_ACTIVE.msdyn_name)).toHaveLength(1);
    expect(list.find((d) => d.name === AI_MATCH_ACTIVE.msdyn_name)?.id).toBe(
      AI_MATCH_ACTIVE.msdyn_dataflowid
    );
  });

  it("keeps a draft that has never been published", async () => {
    const list = await listDataflows(fakeClient(ALL_ROWS));
    expect(list.map((d) => d.id)).toContain(NEVER_PUBLISHED.msdyn_dataflowid);
  });

  it("returns every row when drafts are requested", async () => {
    const list = await listDataflows(fakeClient(ALL_ROWS), { includeDrafts: true });
    expect(list).toHaveLength(ALL_ROWS.length);
  });
});

describe("dataflowToPqtArchive", () => {
  it("produces an archive the existing .pqt writers accept", async () => {
    const archive = dataflowToPqtArchive(COMPANIES_ACTIVE);
    expect(archive.queriesMetadataShape).toBe("object");
    expect(archive.metadata.Name).toBe(COMPANIES_ACTIVE.msdyn_name);
    // The real assertion: the workbook writer runs against it unchanged.
    const bytes = await buildWorkbookWithQueries(archive);
    expect(bytes.byteLength).toBeGreaterThan(0);
    // .xlsx is a ZIP — check the local file header rather than just a length.
    expect([bytes[0], bytes[1]]).toEqual([0x50, 0x4b]);
  });
});

describe("mappingFromDataflow", () => {
  it("infers upsert from the alternate key, resolving key name to attributes", async () => {
    const detail = detailFromRow(AI_MATCH_ACTIVE);
    const m = await mappingFromDataflow(detail, "Full_Match_Table", {
      environmentUrl: ENV,
      resolver,
    });
    expect(m.conflictMode).toBe("upsert");
    // Not "asker_actid" — that is the key, not the attribute it spans.
    expect(m.upsertKey).toEqual(["asker_act_id"]);
    expect(m.targetEntitySet).toBe("accounts");
    expect(validateMapping(m)).toEqual([]);
  });

  it("falls back to insert, with a note, when there is no resolver", async () => {
    const detail = detailFromRow(AI_MATCH_ACTIVE);
    const m = await mappingFromDataflow(detail, "Full_Match_Table", { environmentUrl: ENV });
    expect(m.conflictMode).toBe("insert");
    expect(m.upsertKey).toBeUndefined();
    expect(m.description).toContain("asker_actid");
  });

  it("falls back to insert when the key's attributes are not all mapped", async () => {
    const detail = detailFromRow(AI_MATCH_ACTIVE);
    const partial: DataflowMetadataResolver = {
      ...resolver,
      async keyAttributes() {
        return ["asker_act_id", "asker_unmapped"];
      },
    };
    const m = await mappingFromDataflow(detail, "Full_Match_Table", {
      environmentUrl: ENV,
      resolver: partial,
    });
    expect(m.conflictMode).toBe("insert");
    expect(m.description).toContain("asker_unmapped");
  });

  it("splits a dotted target into a lookup bound by alternate key", async () => {
    const detail = detailFromRow(COMPANIES_ACTIVE);
    const m = await mappingFromDataflow(detail, "Companies", {
      environmentUrl: ENV,
      resolver,
    });
    const lookup = m.columns.find((c) => c.target === "ParentAccountId");
    expect(lookup).toMatchObject({
      source: "[-Company ID-]",
      kind: "lookup",
      lookupResolution: "alternateKey",
      keyAttribute: "asker_act_id",
      bindEntitySet: "accounts",
    });
    // Not a memo column, which is what DestinationFieldType alone would say.
    expect(m.columns.some((c) => c.target.includes("."))).toBe(false);
    expect(validateMapping(m)).toEqual([]);
  });

  it("flags an unresolvable lookup instead of emitting a broken column silently", async () => {
    const detail = detailFromRow(COMPANIES_ACTIVE);
    const blind: DataflowMetadataResolver = {
      ...resolver,
      async lookupEntitySet() {
        return undefined;
      },
    };
    const m = await mappingFromDataflow(detail, "Companies", {
      environmentUrl: ENV,
      resolver: blind,
    });
    expect(m.description).toContain("ParentAccountId.asker_act_id");
    // validateMapping is the backstop; it must not pass quietly.
    expect(validateMapping(m).join(" ")).toContain("bindEntitySet");
  });

  it("maps plain columns by DestinationFieldType", async () => {
    const detail = detailFromRow(AI_MATCH_ACTIVE);
    const m = await mappingFromDataflow(detail, "Full_Match_Table", {
      environmentUrl: ENV,
      resolver,
    });
    expect(m.columns).toContainEqual({
      source: "Account Van",
      target: "asker_companyorigin",
      kind: "integer",
      treatEmptyAsNull: true,
    });
  });

  it("rejects a query that is not in the dataflow", async () => {
    const detail = detailFromRow(AI_MATCH_ACTIVE);
    await expect(mappingFromDataflow(detail, "Nope", { environmentUrl: ENV })).rejects.toThrow(/not found/i);
  });
});

describe("mappingsFromDataflow", () => {
  it("emits one mapping per load-enabled query and skips staging queries", async () => {
    const detail = detailFromRow(COMPANIES_ACTIVE);
    const all = await mappingsFromDataflow(detail, { environmentUrl: ENV, resolver });
    expect(Object.keys(all)).toEqual(["Companies"]);
    expect(all.Companies.sourceTable).toBe("Companies");
  });
});
