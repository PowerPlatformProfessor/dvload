// msdyn_dataflow rows, reduced from a real Dataverse environment.
//
// The M documents are truncated (they run to tens of KB and none of the
// parsing under test reads past the `shared` declarations), but every
// structural feature is preserved verbatim, because those are exactly the
// things that were surprising:
//
//   - QueriesMetadata as a name-keyed object with per-query FieldsMetadata
//   - HostProperties.MappingEditorProperties as JSON *inside* a JSON string
//   - DataflowMetadata likewise doubly-encoded, carrying DataflowState
//   - the same dataflow present twice, once "Source" and once "Active"
//   - a dotted lookup target ("ParentAccountId.asker_act_id")
//   - staging queries with LoadEnabled=false and empty FieldsMetadata

import type { DataflowRow } from "../../src/dataflow.js";

const AI_MATCH_M =
  'section Section1;\r\nshared Full_Match_Table = let\n  Source = Excel.Workbook(Web.Contents("https://example-my.sharepoint.com/personal/x/Documents/AI Matched Accounts.xlsx"), null, true),\n  #"Navigation 1" = Source{[Item = "AI_Matched___Import", Kind = "Table"]}[Data]\nin\n  #"Navigation 1";\r\n';

const COMPANIES_M =
  'section Section1;\r\nshared Companies = let\n  Source = Excel.Workbook(Web.Contents("https://example-my.sharepoint.com/personal/x/Companies.xlsx"), null, true)\nin\n  Source;\r\nshared asker_customertype = let\n  Source = 1\nin\n  Source;\r\nshared AI_Matched___Import = let\n  Source = 2\nin\n  Source;\r\n';

/** MappingEditorProperties is stored as an escaped JSON string, not an object. */
const mappingEditorProps = (primaryKeyLogicalName?: string): string =>
  JSON.stringify(primaryKeyLogicalName ? { primaryKeyLogicalName } : {});

const dataflowMetadata = (state: "Active" | "Source", id: string): string =>
  JSON.stringify({
    DataflowType: "Cdst",
    DataflowId: id,
    DataflowState: state,
    HostContext: { Type: "CdsT", Details: { EnvironmentId: "env-1" } },
  });

function settings(queries: Record<string, unknown>, state: "Active" | "Source", id: string): string {
  return JSON.stringify({
    QueryGroups: [],
    DocumentLocale: "en-US",
    QueriesMetadata: queries,
    FastCombine: false,
    AllowNativeQueries: false,
    HostContext: { Type: "CdsT", Details: { EnvironmentId: "env-1" } },
    DataflowMetadata: dataflowMetadata(state, id),
  });
}

const fullMatchQuery = {
  QueryId: "8173fbbe-bf1f-4721-a764-caa836a7121d",
  QueryName: "Full_Match_Table",
  QueryGroupId: null,
  EntityName: "Account",
  LastKnownIsParameter: false,
  IsHidden: null,
  LoadEnabled: true,
  DeleteExistingDataOnLoad: false,
  FieldsMetadata: {
    asker_act_id: {
      SourceColumnName: "[-ID-] (A)",
      DestinationFieldType: "Memo",
      AutoNumberSettings: null,
    },
    asker_companyorigin: {
      SourceColumnName: "Account Van",
      DestinationFieldType: "Integer",
      AutoNumberSettings: null,
    },
    Name: {
      SourceColumnName: "Company Name (A)",
      DestinationFieldType: "Memo",
      AutoNumberSettings: null,
    },
  },
  HostProperties: { MappingEditorProperties: mappingEditorProps("asker_actid") },
};

const companiesQuery = {
  QueryId: "e35817e9-1ad4-4604-b01c-4947115fec98",
  QueryName: "Companies",
  QueryGroupId: null,
  EntityName: "Account",
  LoadEnabled: true,
  DeleteExistingDataOnLoad: false,
  FieldsMetadata: {
    asker_act_id: {
      SourceColumnName: "[-ID-]",
      DestinationFieldType: "Memo",
      AutoNumberSettings: null,
    },
    Name: {
      SourceColumnName: "Instellingsnaam",
      DestinationFieldType: "Memo",
      AutoNumberSettings: null,
    },
    asker_attitude: {
      SourceColumnName: "asker_attitude.Value",
      DestinationFieldType: "Integer",
      AutoNumberSettings: null,
    },
    // The lookup case: navigation property + alternate key, dot-separated.
    "ParentAccountId.asker_act_id": {
      SourceColumnName: "[-Company ID-]",
      DestinationFieldType: "Memo",
      AutoNumberSettings: null,
    },
  },
  HostProperties: { MappingEditorProperties: mappingEditorProps("asker_actid") },
};

/** A staging query: joined against by Companies, never written to Dataverse. */
const customerTypeQuery = {
  QueryId: "5d817cc4-ced2-4789-9d20-06eb2deffc09",
  QueryName: "asker_customertype",
  QueryGroupId: null,
  EntityName: "crfea_asker_customertype",
  LoadEnabled: false,
  DeleteExistingDataOnLoad: false,
  FieldsMetadata: {},
  HostProperties: {
    NewEntityPrimaryNameField: "crfea_Name",
    MappingEditorProperties: mappingEditorProps(),
  },
};

const importQuery = {
  QueryId: "c34de1c3-21b7-4dc0-b107-e5697fbb3590",
  QueryName: "AI_Matched___Import",
  QueryGroupId: null,
  EntityName: null,
  LoadEnabled: false,
  DeleteExistingDataOnLoad: false,
  FieldsMetadata: {},
  HostProperties: { MappingEditorProperties: mappingEditorProps() },
};

export const AI_MATCH_ACTIVE: DataflowRow = {
  msdyn_dataflowid: "339b8c59-6e23-f111-8342-7ced8d2f40b2",
  msdyn_name: "1. Import AI match Accounts",
  msdyn_mashupdocument: AI_MATCH_M,
  msdyn_mashupsettings: settings(
    { Full_Match_Table: fullMatchQuery },
    "Active",
    "566e43d6-ae94-47b4-b2d9-519bad3721e6"
  ),
  modifiedon: "2026-06-25T07:22:19Z",
};

/** Same dataflow, editing draft. Identical name — this is the dedupe case. */
export const AI_MATCH_DRAFT: DataflowRow = {
  ...AI_MATCH_ACTIVE,
  msdyn_dataflowid: "84b6a395-6d23-f111-8342-7c1e5276f4a9",
  msdyn_mashupsettings: settings(
    { Full_Match_Table: fullMatchQuery },
    "Source",
    "c8b3c6bf-2685-4239-a73a-14d7d7bce401"
  ),
  modifiedon: "2026-07-30T11:39:54Z",
};

export const COMPANIES_ACTIVE: DataflowRow = {
  msdyn_dataflowid: "70a5f2f3-d664-f111-a826-7ced8d2f40b2",
  msdyn_name: "1b. Import ACT extract Accounts",
  msdyn_mashupdocument: COMPANIES_M,
  msdyn_mashupsettings: settings(
    {
      Companies: companiesQuery,
      asker_customertype: customerTypeQuery,
      AI_Matched___Import: importQuery,
    },
    "Active",
    "84eb1c0a-da34-4c9e-b3f2-0f2d87f26e61"
  ),
  modifiedon: "2026-06-25T09:19:32Z",
};

/** Never published: only a draft exists, so the picker must still show it. */
export const NEVER_PUBLISHED: DataflowRow = {
  msdyn_dataflowid: "11111111-1111-1111-1111-111111111111",
  msdyn_name: "3. Work in progress",
  msdyn_mashupdocument: AI_MATCH_M,
  msdyn_mashupsettings: settings({ Full_Match_Table: fullMatchQuery }, "Source", "wip"),
  modifiedon: "2026-07-01T00:00:00Z",
};

export const ALL_ROWS: DataflowRow[] = [AI_MATCH_DRAFT, AI_MATCH_ACTIVE, COMPANIES_ACTIVE, NEVER_PUBLISHED];
