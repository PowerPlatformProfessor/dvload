// Task pane controller. Wires up the HTML in taskpane.html: list tables,
// fetch entities, build a column mapping, run the import. Persists the most
// recent mapping per workbook in Office's Settings store.

import {
  DataverseClient,
  loadRows,
  parseMapping,
  parseRunPlan,
  serializeRunPlan,
  serializeMapping,
  validateMapping,
  validateColumn,
  validateRunPlan,
  mappingWarnings,
  type Mapping,
  type ColumnMapping,
  type DataverseFieldKind,
  type RequestLogEntry,
  type RowSuccess,
  type LoadResult,
  type PqtArchive,
  readPqt,
  writePqt,
  extractPqtFromXlsx,
  injectMappingIntoPqt,
  mappingFromPqt,
  mappingsFromPqtAll,
  readTableFromBuffer,
  readTableFromCsvString,
  writeRowsToBuffer,
  SCHEMA_VERSION,
  suggestColumns,
  buildEntityPayload,
  buildAttributePayload,
  buildKeyPayload,
  attributeLogicalName,
  KEYABLE_KINDS,
  type GeneratedColumn,
  type GeneratedKind,
  type RunPlan,
  type RunPlanStep,
  type RunPlanStepOverrides,
  RUN_PLAN_SCHEMA_VERSION,
} from "@dvload/core";
import { initAuth, getAccount, signIn, makeTokenProvider, devModeBanner } from "../auth.js";
import { listTables, readTable, getWorkbookBytes, type TableInfo } from "../excel.js";
import { suggestMappings, suggestionsToMappings } from "../suggest.js";
import { enhanceSelect } from "../combobox.js";
import {
  track as trackTelemetry,
  bucket as telemetryBucket,
  telemetryAvailable,
  telemetryEnabled,
  setTelemetryEnabled,
} from "../telemetry.js";

const SETTINGS_KEY = "dvload:lastMapping";
const PLAN_SETTINGS_KEY = "dvload:lastRunPlan";
const CHECKPOINT_KEY = "dvload:runCheckpoint";
const PROFILES_KEY = "dvload:profiles";

interface EntityAttribute {
  logicalName: string;
  attributeType: string;
  format?: string;
  /** IsValidForCreate || IsValidForUpdate — read-only attributes (e.g. contact.accountid, fullname) are not mappable. */
  writable: boolean;
}

interface Profile {
  name: string;
  url: string;
}

/* -------------------------------------------------------------------------- */
/* Profile storage (localStorage)                                              */
/* -------------------------------------------------------------------------- */

function loadProfiles(): Profile[] {
  try {
    return JSON.parse(localStorage.getItem(PROFILES_KEY) ?? "[]") as Profile[];
  } catch {
    return [];
  }
}

function saveProfiles(profiles: Profile[]): void {
  localStorage.setItem(PROFILES_KEY, JSON.stringify(profiles));
}

function renderProfilePicker(): void {
  const sel = el<HTMLSelectElement>("profile");
  const profiles = loadProfiles();
  const current = sel.value; // preserve selection if re-rendering
  sel.innerHTML = "";
  sel.appendChild(new Option("Select a saved environment…", ""));
  for (const p of profiles) {
    sel.appendChild(new Option(`${p.name}  —  ${p.url}`, p.url));
  }
  if (current && [...sel.options].some((o) => o.value === current)) {
    sel.value = current;
  }
  el<HTMLButtonElement>("profileDelete").disabled = !sel.value;
}

/** Rows loaded from a picked file rather than a table in the open workbook. */
interface FileSource {
  fileName: string;
  /** Table/sheet name recorded in the mapping's sourceTable. */
  tableName: string;
  headers: string[];
  rows: Array<Record<string, unknown>>;
}

/** A cancelled or partly-failed run, so it can be picked up where it stopped. */
interface Checkpoint {
  /** Row index to restart from (rows before it were already attempted). */
  offset: number;
  total: number;
  sourceTable: string;
  savedAt: string;
}

interface AppState {
  account: { username: string } | null;
  environmentUrl: string;
  tables: TableInfo[];
  selectedTable: TableInfo | null;
  fileSource: FileSource | null;
  entitySet: string;
  entityLogicalName: string;
  entityAttributes: EntityAttribute[];
  entities: Array<{ logicalName: string; entitySetName: string; displayName: string; metadataId: string }>;
  mappings: ColumnMapping[];
  planSteps: RunPlanStep[];
  /**
   * Mapping-level fields the pane doesn't own a control for but must not
   * destroy. buildMapping() spreads this in, so loading a mapping written by
   * hand or by the CLI and re-saving it round-trips instead of silently
   * dropping createdAt, logDir and anything added to the schema later.
   */
  mappingExtras: Partial<Mapping>;
  planMeta: { name: string; description?: string; stopOnError: boolean; createdAt?: string };
  checkpoint: Checkpoint | null;
}

const state: AppState = {
  account: null,
  environmentUrl: "",
  tables: [],
  selectedTable: null,
  fileSource: null,
  entitySet: "",
  entityLogicalName: "",
  entityAttributes: [],
  entities: [],
  mappings: [],
  planSteps: [],
  mappingExtras: {},
  planMeta: { name: "Run plan", stopOnError: true },
  checkpoint: null,
};

const lookupTargetsCache = new Map<string, string[]>();
const optionLabelsCache = new Map<string, Record<string, number>>();
const solutionEntityIdsCache = new Map<string, Set<string>>();

const CHOICE_KINDS: readonly DataverseFieldKind[] = ["choice", "multichoice", "status", "state"];

/**
 * Fetch option-set labels from metadata and store them as the column's
 * optionMap, so users can keep human-readable labels in the spreadsheet
 * instead of typing integer values.
 */
async function applyOptionLabels(attrLogical: string, i: number): Promise<void> {
  if (!attrLogical || !state.entityLogicalName || !state.environmentUrl) return;
  const cacheKey = `${state.entityLogicalName}/${attrLogical}`;
  let map = optionLabelsCache.get(cacheKey);
  if (!map) {
    try {
      const client = new DataverseClient({
        environmentUrl: state.environmentUrl,
        getToken: makeTokenProvider(state.environmentUrl),
      });
      map = await client.getOptionSetLabels(state.entityLogicalName, attrLogical);
      optionLabelsCache.set(cacheKey, map);
    } catch {
      return; // stick with whatever optionMap the user typed
    }
  }
  // Only apply if the row still points at this attribute (async race).
  if (state.mappings[i]?.target === attrLogical) {
    state.mappings[i].optionMap = map;
    setStatus(
      "info",
      `Loaded ${Object.keys(map).length} option label(s) for ${attrLogical} — ` +
        `spreadsheet cells can use labels or integer values.`
    );
  }
}

Office.onReady(async () => {
  await initAuth();
  renderDevModeBanner();
  await refreshAccountUI();

  state.tables = await listTables();
  populateTablePicker();
  restoreMapping();
  restoreRunPlan();

  // If there is a cached MSAL account and a saved environment URL, auto-load
  // entities and restore the entity selection without requiring a Sign-in click.
  if (state.account && state.environmentUrl) {
    triggerLoadEntities();
  }

  // Profiles
  renderProfilePicker();
  el<HTMLSelectElement>("profile").addEventListener("change", (e) => {
    const url = (e.target as HTMLSelectElement).value;
    if (!url) return;
    state.environmentUrl = url;
    el<HTMLInputElement>("env").value = url;
    el<HTMLButtonElement>("profileDelete").disabled = false;
    if (state.account) triggerLoadEntities();
  });
  el<HTMLButtonElement>("profileDelete").addEventListener("click", () => {
    const url = el<HTMLSelectElement>("profile").value;
    if (!url) return;
    const profiles = loadProfiles().filter((p) => p.url !== url);
    saveProfiles(profiles);
    renderProfilePicker();
  });
  el<HTMLButtonElement>("profileSaveBtn").addEventListener("click", () => {
    el<HTMLDivElement>("profileSaveRow").style.display = "";
    el<HTMLInputElement>("profileName").focus();
  });
  el<HTMLButtonElement>("profileSaveCancel").addEventListener("click", () => {
    el<HTMLDivElement>("profileSaveRow").style.display = "none";
    el<HTMLInputElement>("profileName").value = "";
  });
  el<HTMLButtonElement>("profileSaveConfirm").addEventListener("click", () => {
    const name = el<HTMLInputElement>("profileName").value.trim();
    const url = el<HTMLInputElement>("env").value.trim();
    if (!name || !url) { setStatus("error", "Enter both a name and an environment URL."); return; }
    const profiles = loadProfiles().filter((p) => p.name !== name);
    profiles.push({ name, url });
    saveProfiles(profiles);
    el<HTMLDivElement>("profileSaveRow").style.display = "none";
    el<HTMLInputElement>("profileName").value = "";
    renderProfilePicker();
    el<HTMLSelectElement>("profile").value = url;
    el<HTMLButtonElement>("profileDelete").disabled = false;
    setStatus("success", `Profile "${name}" saved.`);
  });

  el<HTMLInputElement>("env").addEventListener("change", (e) => {
    state.environmentUrl = (e.target as HTMLInputElement).value.trim();
  });

  el<HTMLButtonElement>("signin").addEventListener("click", onSignIn);
  el<HTMLSelectElement>("table").addEventListener("change", onPickTable);
  el<HTMLSelectElement>("solution").addEventListener("change", onPickSolution);
  el<HTMLSelectElement>("entity").addEventListener("change", onPickEntity);
  el<HTMLButtonElement>("addMap").addEventListener("click", () => addMapping());
  el<HTMLButtonElement>("addConst").addEventListener("click", () => addConstant());
  el<HTMLButtonElement>("suggest").addEventListener("click", onSuggest);
  el<HTMLButtonElement>("run").addEventListener("click", () => void onRun());
  el<HTMLButtonElement>("cancelRun").addEventListener("click", onCancelRun);
  el<HTMLButtonElement>("resetAll").addEventListener("click", onResetAll);
  el<HTMLButtonElement>("save").addEventListener("click", onSave);
  el<HTMLButtonElement>("load").addEventListener("click", onLoad);
  el<HTMLButtonElement>("importPqt").addEventListener("click", onImportPqt);
  el<HTMLButtonElement>("extractPqt").addEventListener("click", () => void onExtractPqt());
  el<HTMLButtonElement>("pqtUse").addEventListener("click", onUsePqtMapping);
  el<HTMLButtonElement>("pqtUseAll").addEventListener("click", onUseAllPqtMappings);
  el<HTMLButtonElement>("pqtCopyM").addEventListener("click", onCopyPqtM);
  el<HTMLButtonElement>("pqtExport").addEventListener("click", onExportPqt);
  el<HTMLButtonElement>("planAddStep").addEventListener("click", onPlanAddStep);
  el<HTMLButtonElement>("planFromCurrent").addEventListener("click", onPlanAddFromCurrent);
  el<HTMLButtonElement>("planSave").addEventListener("click", onPlanSave);
  el<HTMLButtonElement>("planLoad").addEventListener("click", onPlanLoad);
  el<HTMLButtonElement>("pickFile").addEventListener("click", onPickFileSource);
  el<HTMLButtonElement>("clearFileSource").addEventListener("click", clearFileSource);
  el<HTMLButtonElement>("resumeRun").addEventListener("click", () => void onRun({ resume: true }));
  el<HTMLButtonElement>("discardResume").addEventListener("click", () => {
    state.checkpoint = null;
    persistCheckpoint();
    renderCheckpoint();
  });
  el<HTMLButtonElement>("downloadFailed").addEventListener("click", () => void onDownloadFailedRows());
  el<HTMLSelectElement>("conflictMode").addEventListener("change", updateOptionsVisibility);

  // Plan metadata — kept in state so load → save round-trips it.
  el<HTMLInputElement>("planName").addEventListener("change", () => {
    state.planMeta.name = el<HTMLInputElement>("planName").value.trim() || "Run plan";
    persistRunPlan();
  });
  el<HTMLInputElement>("planDescription").addEventListener("change", () => {
    state.planMeta.description = el<HTMLInputElement>("planDescription").value.trim() || undefined;
    persistRunPlan();
  });
  el<HTMLInputElement>("planStopOnError").addEventListener("change", () => {
    state.planMeta.stopOnError = el<HTMLInputElement>("planStopOnError").checked;
    persistRunPlan();
  });

  updateOptionsVisibility();
  restoreCheckpoint();
  rerenderRunPlan();
  writePlanMetaToForm();

  // Type-to-filter on the big pickers. The native selects stay in the DOM as
  // the source of truth; the combobox is a UI layer over them.
  enhanceSelect(el<HTMLSelectElement>("solution"));
  enhanceSelect(el<HTMLSelectElement>("entity"));

  initCtOwnerSearch();
  initImpersonateSearch();

  // Telemetry checkbox — only shown when this build can actually send
  // (a connection string was baked in). See TELEMETRY.md.
  if (telemetryAvailable()) {
    el<HTMLLabelElement>("telemetryRow").style.display = "";
    const box = el<HTMLInputElement>("telemetryOptIn");
    box.checked = telemetryEnabled();
    box.addEventListener("change", () => setTelemetryEnabled(box.checked));
    trackTelemetry("addin_open");
  }

  // External links: prefer Office's openBrowserWindow (guaranteed system
  // browser) over the anchor's target=_blank, which some hosts ignore.
  document.body.addEventListener("click", (e) => {
    const a = (e.target as HTMLElement).closest?.("a[target=_blank]") as HTMLAnchorElement | null;
    if (!a?.href) return;
    if (Office.context.ui && "openBrowserWindow" in Office.context.ui) {
      e.preventDefault();
      (Office.context.ui as unknown as { openBrowserWindow: (url: string) => void }).openBrowserWindow(a.href);
    }
  });
});

/* -------------------------------------------------------------------------- */
/* .pqt import (Dataverse Dataflows / Power Query Online export)               */
/* -------------------------------------------------------------------------- */

let currentPqt: PqtArchive | null = null;
/** The bytes it was read from — re-parsed on export so repeated exports don't stack edits. */
let currentPqtBytes: ArrayBuffer | null = null;

function onImportPqt(): void {
  const input = document.createElement("input");
  input.type = "file";
  input.accept = ".pqt,application/zip";
  input.onchange = async () => {
    const file = input.files?.[0];
    if (!file) return;
    try {
      currentPqtBytes = await file.arrayBuffer();
      currentPqt = await readPqt(currentPqtBytes);
      const names = Object.keys(currentPqt.mashupMetadata.QueriesMetadata);
      const sel = el<HTMLSelectElement>("pqtQuery");
      sel.innerHTML = "";
      for (const name of names) {
        const q = currentPqt.mashupMetadata.QueriesMetadata[name];
        const nFields = Object.keys(q.FieldsMetadata ?? {}).length;
        const suffix = nFields > 0 ? ` — ${nFields} field mapping(s)` : " — no mappings";
        sel.appendChild(new Option(`${name}${suffix}`, name));
      }
      el<HTMLDivElement>("pqtRow").style.display = "";
      setStatus(
        "info",
        `Read "${currentPqt.metadata.Name || file.name}": ${names.length} quer${names.length === 1 ? "y" : "ies"}. ` +
          `Pick one and click "Use mapping", or "Copy M" to paste the queries into Excel's Advanced Editor.`
      );
    } catch (e) {
      currentPqt = null;
      currentPqtBytes = null;
      el<HTMLDivElement>("pqtRow").style.display = "none";
      setStatus("error", `Couldn't read .pqt: ${(e as Error).message}`);
    }
  };
  input.click();
}

function onUsePqtMapping(): void {
  if (!currentPqt) return;
  const queryName = el<HTMLSelectElement>("pqtQuery").value;
  if (!queryName) return;
  try {
    const m = mappingFromPqt(currentPqt, queryName, {
      environmentUrl: state.environmentUrl || "",
    });
    state.mappings = m.columns;
    // Try to select the matching target entity if entities are loaded.
    const entSel = el<HTMLSelectElement>("entity");
    const match = [...entSel.options].find(
      (o) => o.value === m.targetEntitySet || o.dataset.logical === m.targetEntitySet
    );
    if (match) {
      entSel.value = match.value;
      state.entitySet = match.value;
      const logical = match.dataset.logical;
      if (logical) loadEntityAttributes(logical).catch(() => {});
    }
    rerenderMappings();
    setStatus(
      "success",
      `Loaded ${m.columns.length} column mapping(s) from query "${queryName}"` +
        (match ? ` (target: ${m.targetEntitySet}).` : `. Target "${m.targetEntitySet}" — sign in and pick the entity to verify attributes.`)
    );
  } catch (e) {
    setStatus("error", (e as Error).message);
  }
}

/**
 * Every query in the .pqt at once (the CLI's `pqt --all-queries`): each one is
 * downloaded as its own .dvmap.json and added to the run plan, which is the
 * only way to actually execute more than one of them.
 */
function onUseAllPqtMappings(): void {
  if (!currentPqt) return;
  try {
    const mappings = mappingsFromPqtAll(currentPqt, { environmentUrl: state.environmentUrl || "" });
    const names = Object.keys(mappings);
    if (names.length === 0) {
      setStatus("info", "No queries with field mappings in this .pqt.");
      return;
    }
    for (const name of names) {
      const stem = name.replace(/\W+/g, "-").replace(/^-|-$/g, "").toLowerCase() || "mapping";
      downloadBlob(
        new Blob([serializeMapping(mappings[name])], { type: "application/json" }),
        `${stem}.dvmap.json`
      );
      if (!state.planSteps.some((s) => s.id === stem)) {
        state.planSteps.push({
          id: stem,
          mapping: `./${stem}.dvmap.json`,
          workbook: "./source.xlsx",
          stage: 1,
        });
      }
    }
    persistRunPlan();
    rerenderRunPlan();
    setStatus(
      "success",
      `Downloaded ${names.length} mapping${names.length === 1 ? "" : "s"} and added ` +
        `${names.length} run-plan step${names.length === 1 ? "" : "s"}. Set each step's workbook path before running.`
    );
  } catch (e) {
    setStatus("error", (e as Error).message);
  }
}

/**
 * The reverse of importing: write the current column mapping into the .pqt's
 * FieldsMetadata and hand back the archive, so it imports into a Dataverse
 * Dataflow with the column-to-attribute mapping already filled in.
 */
async function onExportPqt(): Promise<void> {
  if (!currentPqt || !currentPqtBytes) return;
  const queryName = el<HTMLSelectElement>("pqtQuery").value;
  if (!queryName) return;
  if (state.mappings.length === 0) {
    setStatus("error", "Build a column mapping first — there's nothing to write into the .pqt.");
    return;
  }
  try {
    const mapping = buildMapping();
    const errs = validateMapping(mapping);
    if (errs.length > 0) {
      setStatus("error", "Fix the mapping first: " + errs.join("; "));
      renderRowErrors();
      return;
    }
    // readPqt again from the original bytes so repeated exports don't stack
    // edits onto an already-modified archive.
    const fresh = await readPqt(currentPqtBytes);
    injectMappingIntoPqt(fresh, mapping);
    downloadBlob(
      new Blob([toArrayBuffer(await writePqt(fresh))], { type: "application/zip" }),
      `${(mapping.name || queryName).replace(/\W+/g, "-").toLowerCase()}.pqt`
    );
    setStatus("success", `Wrote ${mapping.columns.length} field mapping(s) into the .pqt.`);
  } catch (e) {
    setStatus("error", `Couldn't write the .pqt: ${(e as Error).message}`);
  }
}

/**
 * Package this workbook's own Power Query as a .pqt, with the current mapping
 * injected — the CLI's `extract-pqt --mapping`. Unlike "Export .pqt" this
 * needs no imported archive, so it's the path from "I built a query in Excel"
 * to "I have a Dataverse Dataflow".
 */
async function onExtractPqt(): Promise<void> {
  try {
    setStatus("info", "Reading the workbook's Power Query…");
    const bytes = await getWorkbookBytes();
    const mapping = state.mappings.length > 0 ? buildMapping() : null;
    if (mapping) {
      const errs = validateMapping(mapping);
      if (errs.length > 0) {
        setStatus("error", "Fix the mapping first: " + errs.join("; "));
        renderRowErrors();
        return;
      }
    }
    const archive = await extractPqtFromXlsx(bytes, {
      name: mapping?.name || "Extracted from Excel",
      description: mapping?.description,
    });
    if (mapping) injectMappingIntoPqt(archive, mapping);
    const stem = (mapping?.name || "workbook").replace(/\W+/g, "-").toLowerCase();
    downloadBlob(new Blob([toArrayBuffer(await writePqt(archive))], { type: "application/zip" }), `${stem}.pqt`);
    setStatus(
      "success",
      mapping
        ? `Exported .pqt with ${mapping.columns.length} field mapping(s) — import it as a Dataverse Dataflow.`
        : "Exported .pqt (queries only — build a column mapping first if you want the field mappings in it)."
    );
  } catch (e) {
    setStatus("error", (e as Error).message);
  }
}

async function onCopyPqtM(): Promise<void> {
  if (!currentPqt) return;
  try {
    await navigator.clipboard.writeText(currentPqt.mashupDocument);
    setStatus(
      "success",
      "M code copied. In Excel: Data → Get Data → From Other Sources → Blank Query → " +
        "Advanced Editor → paste. Each shared member becomes a query."
    );
  } catch {
    setStatus("error", "Clipboard access denied — save the mapping instead and use the CLI (--emit-m).");
  }
}

/* -------------------------------------------------------------------------- */
/* File source (another workbook, or a .csv/.tsv)                              */
/* -------------------------------------------------------------------------- */

/**
 * Read rows from a picked file instead of a table in the open workbook. Core
 * already ships buffer-based readers for exactly this — the pane just never
 * used them, so a CSV or a second workbook meant dropping to the CLI.
 */
function onPickFileSource(): void {
  const input = document.createElement("input");
  input.type = "file";
  input.accept = ".xlsx,.xlsm,.csv,.tsv,text/csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  input.onchange = async () => {
    const file = input.files?.[0];
    if (!file) return;
    try {
      const lower = file.name.toLowerCase();
      let table: { headers: string[]; rows: Array<Record<string, unknown>> };
      let tableName: string;
      if (lower.endsWith(".csv") || lower.endsWith(".tsv")) {
        const text = await file.text();
        table = readTableFromCsvString(text, { delimiter: lower.endsWith(".tsv") ? "\t" : "," });
        tableName = file.name.replace(/\.(csv|tsv)$/i, "");
      } else {
        // No table name: core takes the workbook's first table, or the first
        // sheet's used range when there isn't one.
        table = await readTableFromBuffer(await file.arrayBuffer(), {});
        tableName = file.name.replace(/\.(xlsx|xlsm)$/i, "");
      }
      if (table.rows.length === 0) {
        setStatus("error", `${file.name} has no data rows.`);
        return;
      }
      state.fileSource = { fileName: file.name, tableName, headers: table.headers, rows: table.rows };
      // The source column pickers read selectedTable.columns, so present the
      // file as if it were a table in the workbook.
      state.selectedTable = {
        name: tableName,
        worksheetName: "",
        columns: table.headers,
      } as TableInfo;
      el<HTMLSelectElement>("table").value = "";
      renderFileSource();
      rerenderMappings();
      setStatus(
        "success",
        `Using ${file.name} — ${table.rows.length} row${table.rows.length === 1 ? "" : "s"}, ` +
          `${table.headers.length} column${table.headers.length === 1 ? "" : "s"}.`
      );
    } catch (e) {
      setStatus("error", `Couldn't read ${file.name}: ${(e as Error).message}`);
    }
  };
  input.click();
}

function clearFileSource(): void {
  state.fileSource = null;
  state.selectedTable = null;
  renderFileSource();
  rerenderMappings();
  setStatus("info", "Back to reading a table in this workbook — pick one above.");
}

function renderFileSource(): void {
  const info = el<HTMLSpanElement>("fileSourceInfo");
  const clear = el<HTMLButtonElement>("clearFileSource");
  const tableSel = el<HTMLSelectElement>("table");
  if (state.fileSource) {
    info.textContent = `Using file: ${state.fileSource.fileName} (${state.fileSource.rows.length} rows)`;
    clear.style.display = "";
    tableSel.disabled = true;
  } else {
    info.textContent = "";
    clear.style.display = "none";
    tableSel.disabled = false;
  }
}

/* -------------------------------------------------------------------------- */
/* Import options                                                              */
/* -------------------------------------------------------------------------- */

function updateOptionsVisibility(): void {
  const mode = el<HTMLSelectElement>("conflictMode").value;
  const needsKey = mode === "upsert" || mode === "skip-if-exists" || mode === "sync";
  el<HTMLLabelElement>("upsertKeyLabel").style.display = needsKey ? "" : "none";
  el<HTMLInputElement>("upsertKey").style.display = needsKey ? "" : "none";
  const isSync = mode === "sync";
  el<HTMLLabelElement>("syncActionLabel").style.display = isSync ? "" : "none";
  el<HTMLSelectElement>("syncAction").style.display = isSync ? "" : "none";
  el<HTMLInputElement>("skipUnchanged").disabled = !(mode === "upsert" || mode === "sync");
}

function readOptionsIntoMapping(m: Mapping): void {
  const mode = el<HTMLSelectElement>("conflictMode").value as Mapping["conflictMode"];
  m.conflictMode = mode;
  const keyRaw = el<HTMLInputElement>("upsertKey").value.trim();
  m.upsertKey = keyRaw ? keyRaw.split(",").map((s) => s.trim()).filter(Boolean) : undefined;
  if (mode === "sync") {
    m.syncAction = el<HTMLSelectElement>("syncAction").value as Mapping["syncAction"];
  }
  const batchSize = Number(el<HTMLInputElement>("batchSize").value);
  if (Number.isInteger(batchSize) && batchSize >= 1 && batchSize <= 1000) m.batchSize = batchSize;
  const concurrency = Number(el<HTMLInputElement>("concurrency").value);
  if (Number.isInteger(concurrency) && concurrency >= 1 && concurrency <= 8) m.concurrency = concurrency;
  m.bypassCustomLogic = el<HTMLInputElement>("bypassCustomLogic").checked;
  m.skipUnchanged =
    el<HTMLInputElement>("skipUnchanged").checked && (mode === "upsert" || mode === "sync");

  const description = el<HTMLInputElement>("mappingDescription").value.trim();
  m.description = description || undefined;

  const maxErrors = Number(el<HTMLInputElement>("maxErrors").value);
  m.maxErrors = Number.isInteger(maxErrors) && maxErrors >= 0 ? maxErrors : 0;

  const notifyUrl = el<HTMLInputElement>("notifyUrl").value.trim();
  m.notifyUrl = notifyUrl || undefined;

  // The visible text is a display name; the GUID lives in a data attribute so
  // a picked user survives without re-searching. A pasted GUID works too.
  const impersonateInput = el<HTMLInputElement>("impersonateUser");
  const typed = impersonateInput.value.trim();
  const picked = impersonateInput.dataset.guid;
  m.impersonateUserId = GUID_INPUT_RE.test(typed) ? typed.toLowerCase() : (typed ? picked : undefined);
}

function writeOptionsFromMapping(m: Mapping): void {
  el<HTMLSelectElement>("conflictMode").value = m.conflictMode;
  el<HTMLInputElement>("upsertKey").value = (m.upsertKey ?? []).join(", ");
  el<HTMLSelectElement>("syncAction").value = m.syncAction ?? "deactivate";
  el<HTMLInputElement>("batchSize").value = String(m.batchSize);
  el<HTMLInputElement>("concurrency").value = String(m.concurrency ?? 1);
  el<HTMLInputElement>("bypassCustomLogic").checked = m.bypassCustomLogic ?? false;
  el<HTMLInputElement>("skipUnchanged").checked = m.skipUnchanged ?? false;
  el<HTMLInputElement>("mappingName").value = m.name ?? "";
  el<HTMLInputElement>("mappingDescription").value = m.description ?? "";
  el<HTMLInputElement>("maxErrors").value = String(m.maxErrors ?? 0);
  el<HTMLInputElement>("notifyUrl").value = m.notifyUrl ?? "";
  const impersonate = el<HTMLInputElement>("impersonateUser");
  impersonate.value = m.impersonateUserId ?? "";
  impersonate.dataset.guid = m.impersonateUserId ?? "";
  // Anything set in the file that the pane can't otherwise show gets opened,
  // so a loaded mapping never has hidden behaviour.
  if (m.description || m.notifyUrl || m.impersonateUserId || (m.maxErrors ?? 0) > 0) {
    el<HTMLDetailsElement>("advancedOptions").open = true;
  }
  updateOptionsVisibility();
}

/* -------------------------------------------------------------------------- */
/* UI helpers                                                                  */
/* -------------------------------------------------------------------------- */

function el<T extends HTMLElement>(id: string): T {
  const e = document.getElementById(id);
  if (!e) throw new Error(`#${id} not found`);
  return e as T;
}

function setStatus(kind: "info" | "error" | "success", message: string): void {
  const s = el<HTMLDivElement>("status");
  s.className = `status ${kind}`;
  s.textContent = message;
  s.style.display = "";
}

function setProgress(
  processed: number,
  total: number,
  counts: { created: number; updated: number; failed: number; skipped: number }
): void {
  el<HTMLDivElement>("progressWrap").style.display = "";
  const pct = total > 0 ? Math.round((processed / total) * 100) : 0;
  el<HTMLDivElement>("progressBar").style.width = `${pct}%`;
  el<HTMLDivElement>("progressText").textContent =
    `${pct}%  ${processed}/${total}  ` +
    `created=${counts.created}  updated=${counts.updated}  ` +
    `skipped=${counts.skipped}  failed=${counts.failed}`;
}

function renderDevModeBanner(): void {
  const text = devModeBanner();
  if (!text) return;
  let banner = document.getElementById("devModeBanner") as HTMLDivElement | null;
  if (!banner) {
    banner = document.createElement("div");
    banner.id = "devModeBanner";
    banner.style.cssText =
      "background:#fff4ce;color:#5a4500;border:1px solid #e0c769;" +
      "border-radius:4px;padding:6px 8px;font-size:11px;margin-bottom:8px;";
    document.body.insertBefore(banner, document.body.firstChild);
  }
  banner.textContent = text + " — replace before sharing this tool.";
}

async function refreshAccountUI(): Promise<void> {
  const acc = await getAccount();
  state.account = acc ? { username: acc.username } : null;
  el<HTMLSpanElement>("authDot").style.color = acc ? "#107c10" : "#d13438";
  el<HTMLSpanElement>("who").textContent = acc ? acc.username : "Not signed in";
  const entSel = el<HTMLSelectElement>("entity");
  entSel.disabled = !acc;
  // Clear the "Sign in to load entities" placeholder once signed in
  if (acc && entSel.options.length === 1 && !entSel.options[0].value) {
    entSel.options[0].text = "Select an entity…";
  }
}

/* -------------------------------------------------------------------------- */
/* Sign-in / entity loading                                                    */
/* -------------------------------------------------------------------------- */

async function onSignIn(): Promise<void> {
  if (!state.environmentUrl) {
    setStatus("error", "Enter an environment URL first.");
    return;
  }
  try {
    const acc = await signIn(state.environmentUrl);
    state.account = { username: acc.username };
    await refreshAccountUI();
    await loadEntities();
  } catch (e) {
    setStatus("error", `Sign-in failed: ${(e as Error).message}`);
  }
}

function dvClient(): DataverseClient {
  return new DataverseClient({
    environmentUrl: state.environmentUrl,
    getToken: makeTokenProvider(state.environmentUrl),
  });
}

async function loadEntities(): Promise<void> {
  const client = dvClient();
  const entities = await client.listEntities();
  const sorted = entities.sort((a, b) => a.LogicalName.localeCompare(b.LogicalName));
  state.entities = sorted.map(e => ({
    logicalName: e.LogicalName,
    entitySetName: e.EntitySetName,
    displayName: e.DisplayName,
    metadataId: (e.MetadataId ?? "").toLowerCase(),
  }));
  renderEntityOptions(null);
  el<HTMLSelectElement>("entity").disabled = false;
  await loadSolutions(client);
}

/** Sentinel value for the "create a new table" entity option. */
const CREATE_NEW = "__create_new__";

/** Populate the entity select, optionally restricted to a set of MetadataIds. */
function renderEntityOptions(filter: Set<string> | null): void {
  const sel = el<HTMLSelectElement>("entity");
  const current = sel.value;
  sel.innerHTML = `<option value="">Select an entity…</option>`;
  sel.appendChild(new Option("＋ Create new table from source…", CREATE_NEW));
  let shown = 0;
  for (const e of state.entities) {
    if (filter && !filter.has(e.metadataId)) continue;
    const opt = document.createElement("option");
    opt.value = e.entitySetName;
    opt.textContent = `${e.displayName || e.logicalName} (${e.entitySetName})`;
    opt.dataset.logical = e.logicalName;
    sel.appendChild(opt);
    shown++;
  }
  if (filter && shown === 0) {
    sel.innerHTML = `<option value="">No entities in this solution</option>`;
    return;
  }
  // Keep the current selection when it survives the filter.
  if (current && [...sel.options].some((o) => o.value === current)) {
    sel.value = current;
  }
}

/**
 * Fill the solution picker. Defaults to "All entities" (equivalent to the
 * Default solution, which contains every entity); picking a solution
 * filters the entity list to that solution's tables.
 */
async function loadSolutions(client: DataverseClient): Promise<void> {
  const sel = el<HTMLSelectElement>("solution");
  try {
    const solutions = await client.listSolutions();
    sel.innerHTML = "";
    sel.appendChild(new Option("All entities (Default solution)", ""));
    for (const s of solutions) {
      // The Default solution contains everything — same as "All entities".
      if (s.uniqueName.toLowerCase() === "default") continue;
      sel.appendChild(new Option(s.friendlyName || s.uniqueName, s.id));
    }
    sel.disabled = false;
  } catch {
    // Solution filtering is a nicety; a user without read privilege on the
    // solution table still gets the full entity list.
    sel.innerHTML = `<option value="">All entities (solution list unavailable)</option>`;
    sel.disabled = true;
  }
}

async function onPickSolution(e: Event): Promise<void> {
  const sel = e.target as HTMLSelectElement;
  const solutionId = sel.value;
  if (!solutionId) {
    renderEntityOptions(null);
    return;
  }
  try {
    let ids = solutionEntityIdsCache.get(solutionId);
    if (!ids) {
      ids = await dvClient().getSolutionEntityIds(solutionId);
      solutionEntityIdsCache.set(solutionId, ids);
    }
    renderEntityOptions(ids);
  } catch (err) {
    setStatus("error", `Could not load solution components: ${(err as Error).message}`);
    renderEntityOptions(null);
  }
}

async function onPickEntity(e: Event): Promise<void> {
  const sel = e.target as HTMLSelectElement;
  if (sel.value === CREATE_NEW) {
    await enterCreateTableMode();
    return;
  }
  exitCreateTableMode();
  state.entitySet = sel.value;
  if (!state.entitySet) return;
  const logical = sel.selectedOptions[0]?.dataset.logical ?? state.entitySet.replace(/s$/, "");
  await loadEntityAttributes(logical);
}

/* -------------------------------------------------------------------------- */
/* Create-table mode                                                           */
/* -------------------------------------------------------------------------- */

let createMode = false;
let ctCols: GeneratedColumn[] = [];
let ctFixedOwner: { guid: string; label: string } | null = null;

/** Wire the panel's fixed-owner search (systemusers by fullname, or a pasted GUID). */
/**
 * Type-ahead systemuser search on a plain input. Shared by the create-table
 * owner field and the "run as user" (impersonation) field. A pasted GUID is
 * accepted without a round-trip; `current()` supplies the label to restore on
 * blur so a half-typed search doesn't clobber the picked user.
 */
function initUserSearch(
  inputId: string,
  listId: string,
  onPick: (picked: { guid: string; label: string } | null) => void,
  current: () => { guid: string; label: string } | null
): void {
  const input = el<HTMLInputElement>(inputId);
  const list = el<HTMLDivElement>(listId);
  const close = (): void => { list.style.display = "none"; };
  let timer: ReturnType<typeof setTimeout> | undefined;
  let seq = 0;

  input.addEventListener("input", () => {
    const term = input.value.trim();
    if (!term) { onPick(null); close(); return; }
    if (GUID_INPUT_RE.test(term)) {
      onPick({ guid: term.toLowerCase(), label: term.toLowerCase() });
      close();
      return;
    }
    if (timer) clearTimeout(timer);
    if (term.length < 2) { close(); return; }
    timer = setTimeout(async () => {
      const mySeq = ++seq;
      try {
        const safe = term.replace(/'/g, "''");
        const rows = await dvClient().queryAll(
          `systemusers?$select=systemuserid,fullname&$filter=contains(fullname,'${safe}')&$orderby=fullname&$top=10`
        );
        if (mySeq !== seq) return;
        list.innerHTML = "";
        if (rows.length === 0) {
          list.innerHTML = `<div style="padding:4px 8px;color:#605e5c;">No matching users</div>`;
        }
        for (const r of rows) {
          const guid = String(r["systemuserid"] ?? "");
          const label = String(r["fullname"] ?? guid);
          const item = document.createElement("div");
          item.textContent = label;
          item.style.cssText = "padding:4px 8px;cursor:pointer;";
          item.addEventListener("mousedown", (e) => {
            e.preventDefault();
            onPick({ guid, label });
            input.value = label;
            close();
          });
          list.appendChild(item);
        }
        list.style.display = "";
      } catch {
        if (mySeq === seq) close();
      }
    }, 300);
  });
  input.addEventListener("blur", () => {
    close();
    input.value = current()?.label ?? "";
  });
}

function initCtOwnerSearch(): void {
  initUserSearch(
    "ctOwner",
    "ctOwnerList",
    (picked) => { ctFixedOwner = picked; },
    () => ctFixedOwner
  );
}

/** "Run as user" — the picked GUID is stashed on the input's dataset. */
function initImpersonateSearch(): void {
  const input = el<HTMLInputElement>("impersonateUser");
  initUserSearch(
    "impersonateUser",
    "impersonateUserList",
    (picked) => { input.dataset.guid = picked?.guid ?? ""; input.dataset.label = picked?.label ?? ""; },
    () => (input.dataset.guid ? { guid: input.dataset.guid, label: input.dataset.label || input.dataset.guid } : null)
  );
}

async function enterCreateTableMode(): Promise<void> {
  if (!state.selectedTable) {
    setStatus("error", "Pick a source table first — its columns seed the new table.");
    el<HTMLSelectElement>("entity").value = "";
    return;
  }
  createMode = true;
  el<HTMLButtonElement>("run").textContent = "Create table and run import";
  el<HTMLDivElement>("createTablePanel").style.display = "";
  // The regular mapping grid doesn't apply here — the mapping is generated
  // from the panel when the table is created.
  el<HTMLDivElement>("mappingSection").style.display = "none";
  el<HTMLInputElement>("ctDisplayName").value = state.selectedTable.name.replace(/^tbl/i, "");
  ctFixedOwner = null;
  el<HTMLInputElement>("ctOwner").value = "";
  setStatus("info", "Reading source rows to infer column types…");
  const { rows } = await readTable(state.selectedTable.name);
  ctCols = suggestColumns(state.selectedTable.columns, rows.slice(0, 200));
  renderCtColumns();
  setStatus(
    "info",
    `Suggested ${ctCols.length} column(s) from ${Math.min(rows.length, 200)} sample row(s). ` +
      `Review names/types, pick the primary name and any key columns, then click "Create table and run import".`
  );
}

function exitCreateTableMode(): void {
  if (!createMode) return;
  createMode = false;
  el<HTMLButtonElement>("run").textContent = "Run import";
  el<HTMLDivElement>("createTablePanel").style.display = "none";
  el<HTMLDivElement>("mappingSection").style.display = "";
}

function renderCtColumns(): void {
  const root = el<HTMLDivElement>("ctColumns");
  root.innerHTML = "";
  const KINDS: GeneratedKind[] = ["string", "memo", "integer", "decimal", "boolean", "datetime", "dateonly"];
  const header = document.createElement("div");
  header.style.cssText = "display:grid;grid-template-columns:24px 1fr 1fr 90px 56px 36px;gap:4px;font-size:10px;color:#605e5c;";
  for (const t of ["", "Source", "Column name", "Type", "Primary", "Key"]) {
    const s = document.createElement("span");
    s.textContent = t;
    header.appendChild(s);
  }
  root.appendChild(header);

  ctCols.forEach((c, i) => {
    const row = document.createElement("div");
    row.style.cssText = "display:grid;grid-template-columns:24px 1fr 1fr 90px 56px 36px;gap:4px;align-items:center;margin-top:2px;";

    const inc = document.createElement("input");
    inc.type = "checkbox";
    inc.checked = c.include;
    inc.addEventListener("change", () => { ctCols[i].include = inc.checked; });

    const src = document.createElement("span");
    src.style.cssText = "font-size:11px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;";
    src.textContent = c.source;
    src.title = c.source;

    const name = document.createElement("input");
    name.value = c.systemAttribute ?? c.schemaSuffix;
    name.disabled = !!c.systemAttribute; // fixed system attribute, not a new column
    name.addEventListener("change", () => {
      ctCols[i].schemaSuffix = name.value.replace(/[^A-Za-z0-9]/g, "") || c.schemaSuffix;
      ctCols[i].displayName = c.source;
      name.value = ctCols[i].schemaSuffix;
    });

    const OVERRIDE_CREATEDON = "__overriddencreatedon__";
    const OWNER_USER = "__ownerid_user__";
    const OWNER_TEAM = "__ownerid_team__";
    const kind = document.createElement("select");
    const isSys = (attr: string, bind?: string): boolean =>
      c.systemAttribute === attr && (bind === undefined || c.systemBindEntitySet === bind);
    for (const k of KINDS) {
      const sel = !c.systemAttribute && c.kind === k;
      kind.appendChild(new Option(k, k, sel, sel));
    }
    // System-attribute targets: these don't create a column, they map the
    // source onto an attribute every table already has.
    kind.appendChild(new Option("Created On (backdate)", OVERRIDE_CREATEDON, isSys("overriddencreatedon"), isSys("overriddencreatedon")));
    kind.appendChild(new Option("Owner (user GUIDs)", OWNER_USER, isSys("ownerid", "systemusers"), isSys("ownerid", "systemusers")));
    kind.appendChild(new Option("Owner (team GUIDs)", OWNER_TEAM, isSys("ownerid", "teams"), isSys("ownerid", "teams")));
    kind.title =
      "System targets don't create a column: \"Created On (backdate)\" maps this " +
      "source column to the record's Created On via overriddencreatedon (needs the " +
      "Override Created On privilege); \"Owner\" maps per-row user/team GUIDs to ownerid.";
    kind.addEventListener("change", () => {
      if (kind.value === OVERRIDE_CREATEDON) {
        ctCols[i].systemAttribute = "overriddencreatedon";
        ctCols[i].systemBindEntitySet = undefined;
        ctCols[i].kind = "datetime";
        ctCols[i].isPrimaryName = false;
        ctCols[i].inAlternateKey = false;
      } else if (kind.value === OWNER_USER || kind.value === OWNER_TEAM) {
        ctCols[i].systemAttribute = "ownerid";
        ctCols[i].systemBindEntitySet = kind.value === OWNER_USER ? "systemusers" : "teams";
        ctCols[i].isPrimaryName = false;
        ctCols[i].inAlternateKey = false;
      } else {
        ctCols[i].systemAttribute = undefined;
        ctCols[i].systemBindEntitySet = undefined;
        ctCols[i].kind = kind.value as GeneratedKind;
        if (kind.value !== "string" && ctCols[i].isPrimaryName) ctCols[i].isPrimaryName = false;
        if (!KEYABLE_KINDS.includes(ctCols[i].kind) && ctCols[i].inAlternateKey) ctCols[i].inAlternateKey = false;
      }
      renderCtColumns(); // primary/key eligibility changed
    });

    const primary = document.createElement("input");
    primary.type = "radio";
    primary.name = "ctPrimary";
    primary.disabled = c.kind !== "string";
    primary.checked = c.isPrimaryName;
    primary.addEventListener("change", () => {
      ctCols.forEach((x, j) => (x.isPrimaryName = j === i));
    });

    const key = document.createElement("input");
    key.type = "checkbox";
    key.disabled = !KEYABLE_KINDS.includes(c.kind);
    key.checked = c.inAlternateKey;
    key.addEventListener("change", () => { ctCols[i].inAlternateKey = key.checked; });

    row.append(inc, src, name, kind, primary, key);
    root.appendChild(row);
  });
}

/** Create the table + columns (+ key), then hand over to the normal import. */
async function createTableThenImport(): Promise<void> {
  const prefix = el<HTMLInputElement>("ctPrefix").value.trim().toLowerCase();
  if (!/^[a-z][a-z0-9]{1,7}$/.test(prefix)) {
    setStatus("error", "Prefix must be 2-8 characters, letters/digits, starting with a letter (e.g. \"new\", \"contoso\").");
    return;
  }
  const displayName = el<HTMLInputElement>("ctDisplayName").value.trim();
  if (!displayName) { setStatus("error", "Enter a display name for the new table."); return; }
  const included = ctCols.filter((c) => c.include);
  if (included.length === 0) { setStatus("error", "Include at least one column."); return; }
  const primary = included.find((c) => c.isPrimaryName);
  if (!primary) { setStatus("error", "Pick a string column as the table's primary name."); return; }
  const keyCols = included.filter((c) => c.inAlternateKey);
  const entitySuffix = sanitize(displayName);
  const entityLogical = `${prefix}_${entitySuffix}`.toLowerCase();

  const client = dvClient();
  setStatus("info", `Creating table ${prefix}_${entitySuffix}…`);
  await client.createEntity(
    buildEntityPayload({ prefix, schemaSuffix: entitySuffix, displayName, primaryNameColumn: primary })
  );

  // System-attribute rows (e.g. overriddencreatedon) map to attributes every
  // table already has — nothing to create for them.
  const rest = included.filter((c) => !c.isPrimaryName && !c.systemAttribute);
  for (let i = 0; i < rest.length; i++) {
    setStatus("info", `Creating column ${i + 1}/${rest.length}: ${rest[i].displayName}…`);
    await client.createAttribute(entityLogical, buildAttributePayload(prefix, rest[i]));
  }

  if (keyCols.length > 0) {
    setStatus("info", "Creating alternate key…");
    await client.createEntityKey(entityLogical, buildKeyPayload(prefix, entitySuffix, keyCols));
  }

  // Resolve the server-assigned entity set name and refresh state.
  const def = await client.getEntityDefinition(entityLogical);
  const entitySetName = String(def["EntitySetName"] ?? `${entityLogical}s`);
  state.entities.push({
    logicalName: entityLogical,
    entitySetName,
    displayName,
    metadataId: String(def["MetadataId"] ?? "").toLowerCase(),
  });
  exitCreateTableMode();
  renderEntityOptions(null);
  el<HTMLSelectElement>("entity").value = entitySetName;
  state.entitySet = entitySetName;
  await loadEntityAttributes(entityLogical);

  // Mapping: source column → generated attribute (or system attribute), 1:1 kinds.
  state.mappings = included.map((c): ColumnMapping => {
    if (c.systemAttribute === "ownerid") {
      return {
        source: c.source,
        target: "ownerid",
        kind: "lookup",
        bindEntitySet: c.systemBindEntitySet ?? "systemusers",
        lookupResolution: "guid",
        treatEmptyAsNull: true,
      };
    }
    return {
      source: c.source,
      target: c.systemAttribute ?? attributeLogicalName(prefix, c),
      kind: c.kind,
      treatEmptyAsNull: true,
    };
  });
  // Fixed owner for every record, picked in the panel's owner search.
  if (ctFixedOwner) {
    state.mappings.push({
      constant: ctFixedOwner.guid,
      target: "ownerid",
      kind: "lookup",
      bindEntitySet: "systemusers",
      lookupResolution: "guid",
      treatEmptyAsNull: true,
      notes: ctFixedOwner.label,
    });
  }
  if (keyCols.length > 0) {
    el<HTMLInputElement>("upsertKey").value = keyCols.map((c) => attributeLogicalName(prefix, c)).join(", ");
  }
  rerenderMappings();
  trackTelemetry("addin_create_table", {
    columns: telemetryBucket(included.length),
    hasKey: String(keyCols.length > 0),
  });
  setStatus("success", `Table ${displayName} created (${entitySetName}). Starting import…`);
}

function sanitize(s: string): string {
  const words = s.split(/[^A-Za-z0-9]+/).filter(Boolean);
  return (words.map((w) => w[0].toUpperCase() + w.slice(1)).join("").replace(/^[^A-Za-z]+/, "") || "Table");
}

/**
 * Normalise an EntityDefinition's expanded Attributes into our shape.
 * Dataverse may return them as an inline array (OData v4) or wrapped in
 * { value: [] }, so both forms are handled.
 */
function attributesFromDefinition(def: Record<string, unknown>): EntityAttribute[] {
  const raw: Array<Record<string, unknown>> = Array.isArray(def.Attributes)
    ? (def.Attributes as Array<Record<string, unknown>>)
    : ((def.Attributes as { value?: Array<Record<string, unknown>> })?.value ?? []);
  return raw
    .map((a) => ({
      logicalName: String(a["LogicalName"] ?? ""),
      attributeType: String(a["AttributeType"] ?? "").toLowerCase(),
      format: a["Format"] != null ? String(a["Format"]) : undefined,
      writable: a["IsValidForCreate"] === true || a["IsValidForUpdate"] === true,
    }))
    .filter((a) => a.logicalName);
}

interface BoundEntityMeta {
  attributes: EntityAttribute[];
  /** Single-attribute alternate keys defined on the entity, by logical name. */
  singleKeyAttributes: string[];
  primaryNameAttribute?: string;
}

const boundEntityMetaCache = new Map<string, BoundEntityMeta>();

/**
 * Metadata for the entity a lookup binds TO (e.g. systemusers for ownerid) —
 * used to populate the key-attribute picker instead of asking the user to
 * type a logical name from memory. Alternate keys come from the Keys
 * collection so "by alt. key" can only offer keys Dataverse will accept.
 */
async function getBoundEntityMeta(entitySetName: string): Promise<BoundEntityMeta> {
  const cached = boundEntityMetaCache.get(entitySetName);
  if (cached) return cached;
  const ref = state.entities.find((e) => e.entitySetName === entitySetName);
  if (!ref || !state.environmentUrl) return { attributes: [], singleKeyAttributes: [] };
  const client = dvClient();
  const meta: BoundEntityMeta = {
    attributes: attributesFromDefinition(await client.getEntityDefinition(ref.logicalName)),
    singleKeyAttributes: [],
  };
  try {
    const info = await client.getEntitySetInfo(entitySetName);
    meta.primaryNameAttribute = info.primaryNameAttribute;
  } catch {
    /* primary name is a nicety — not worth failing the picker over */
  }
  try {
    const keys = await client.queryAll(
      `EntityDefinitions(LogicalName='${ref.logicalName}')/Keys?$select=LogicalName,KeyAttributes`
    );
    meta.singleKeyAttributes = keys
      .map((k) => (Array.isArray(k.KeyAttributes) ? (k.KeyAttributes as string[]) : []))
      .filter((attrs) => attrs.length === 1)
      .map((attrs) => String(attrs[0]));
  } catch {
    /* no permission on Keys, or none defined — fall back to all attributes */
  }
  boundEntityMetaCache.set(entitySetName, meta);
  return meta;
}

async function loadEntityAttributes(logical: string): Promise<void> {
  state.entityLogicalName = logical;
  lookupTargetsCache.clear();
  optionLabelsCache.clear();
  const client = new DataverseClient({
    environmentUrl: state.environmentUrl,
    getToken: makeTokenProvider(state.environmentUrl),
  });
  state.entityAttributes = attributesFromDefinition(await client.getEntityDefinition(logical));

  // If the user hasn't started mapping yet, auto-suggest. Otherwise leave
  // existing mappings alone — they can click "Suggest" to fill in the rest.
  if (state.mappings.length === 0 && state.selectedTable) {
    const writable = state.entityAttributes.filter((a) => a.writable);
    const suggestions = suggestMappings(state.selectedTable.columns, writable.map(a => a.logicalName));
    if (suggestions.length > 0) {
      state.mappings = suggestionsToMappings(suggestions, writable);
      setStatus(
        "info",
        `Auto-filled ${suggestions.length} mapping${suggestions.length === 1 ? "" : "s"} ` +
          `based on name similarity. Review and adjust before running.`
      );
    }
  }
  rerenderMappings();
}

/** Fill in suggestions for any source columns not already mapped. */
function onSuggest(): void {
  if (!state.selectedTable) {
    setStatus("error", "Pick a source table first.");
    return;
  }
  if (state.entityAttributes.length === 0) {
    setStatus("error", "Pick a target entity first.");
    return;
  }
  const writable = state.entityAttributes.filter((a) => a.writable);
  const suggestions = suggestMappings(state.selectedTable.columns, writable.map(a => a.logicalName), {
    excludeSources: state.mappings.map((m) => m.source).filter((s): s is string => !!s),
    excludeTargets: state.mappings.map((m) => m.target).filter(Boolean),
  });
  if (suggestions.length === 0) {
    setStatus("info", "No additional confident suggestions.");
    return;
  }
  state.mappings.push(...suggestionsToMappings(suggestions, writable));
  rerenderMappings();
  setStatus("success", `Added ${suggestions.length} suggestion${suggestions.length === 1 ? "" : "s"}.`);
}

/* -------------------------------------------------------------------------- */
/* Table + mapping UI                                                          */
/* -------------------------------------------------------------------------- */

function populateTablePicker(): void {
  const sel = el<HTMLSelectElement>("table");
  sel.innerHTML = "";
  if (state.tables.length === 0) {
    sel.innerHTML = `<option value="">No tables found in this workbook</option>`;
    return;
  }
  sel.appendChild(new Option("Select a table…", ""));
  for (const t of state.tables) {
    sel.appendChild(new Option(`${t.name} (${t.worksheetName}, ${t.rowCount} rows)`, t.name));
  }
}

function onPickTable(e: Event): void {
  const name = (e.target as HTMLSelectElement).value;
  // Picking a workbook table supersedes any file that was loaded.
  state.fileSource = null;
  renderFileSource();
  state.selectedTable = state.tables.find((t) => t.name === name) ?? null;
  // The create-table panel is seeded from the source table — a stale panel
  // for a different table would create the wrong columns.
  if (createMode) {
    exitCreateTableMode();
    el<HTMLSelectElement>("entity").value = "";
  }
  rerenderMappings();
}

function addMapping(seed?: ColumnMapping): void {
  state.mappings.push(
    seed ?? { source: "", target: "", kind: "string", treatEmptyAsNull: true }
  );
  rerenderMappings();
}

/** A fixed-value row: no source column; `constant` is applied to every record. */
function addConstant(): void {
  state.mappings.push({ constant: "", target: "", kind: "string", treatEmptyAsNull: true });
  rerenderMappings();
}

/** Constant rows are the ones without a source column. */
function isConstantRow(m: ColumnMapping): boolean {
  return m.source === undefined;
}

function matchesKind(kind: DataverseFieldKind): (a: EntityAttribute) => boolean {
  // attributeType is stored lowercase; Format comparison is also case-insensitive
  switch (kind) {
    case "string":           return a => a.attributeType === "string";
    case "memo":             return a => a.attributeType === "memo";
    case "integer":          return a => a.attributeType === "integer" || a.attributeType === "bigint";
    case "decimal":          return a => a.attributeType === "decimal";
    case "money":            return a => a.attributeType === "money";
    case "double":           return a => a.attributeType === "double";
    case "boolean":          return a => a.attributeType === "boolean";
    case "datetime":         return a => a.attributeType === "datetime" && a.format?.toLowerCase() !== "dateonly";
    case "dateonly":         return a => a.attributeType === "datetime" && a.format?.toLowerCase() === "dateonly";
    case "uniqueidentifier": return a => a.attributeType === "uniqueidentifier";
    case "lookup":           return a => ["lookup", "owner", "customer"].includes(a.attributeType);
    case "choice":           return a => a.attributeType === "picklist";
    case "multichoice":      return a => a.attributeType === "multiselectpicklist";
    case "status":           return a => a.attributeType === "status";
    case "state":            return a => a.attributeType === "state";
    default:                 return () => true;
  }
}

type EntityRecord = {
  logicalName: string;
  entitySetName: string;
  displayName: string;
  metadataId: string;
};

async function applyLookupTargets(
  entitySetSel: HTMLSelectElement,
  attrLogical: string,
  i: number
): Promise<void> {
  if (!attrLogical || !state.entityLogicalName || !state.environmentUrl || state.entities.length === 0) return;
  const cacheKey = `${state.entityLogicalName}/${attrLogical}`;
  let targets = lookupTargetsCache.get(cacheKey);
  if (!targets) {
    const client = new DataverseClient({
      environmentUrl: state.environmentUrl,
      getToken: makeTokenProvider(state.environmentUrl),
    });
    targets = await client.getLookupTargets(state.entityLogicalName, attrLogical);
    lookupTargetsCache.set(cacheKey, targets);
  }
  const validEntities = targets
    .map(ln => state.entities.find(e => e.logicalName === ln))
    .filter((e): e is EntityRecord => e != null);
  const prev = state.mappings[i]?.bindEntitySet ?? entitySetSel.value;
  const entitiesToShow = validEntities.length > 0 ? validEntities : state.entities;
  entitySetSel.innerHTML = "";
  if (validEntities.length !== 1) entitySetSel.appendChild(new Option("(entity set)…", ""));
  for (const e of entitiesToShow) {
    entitySetSel.appendChild(new Option(
      `${e.displayName || e.logicalName} (${e.entitySetName})`,
      e.entitySetName
    ));
  }
  if (validEntities.length === 1) {
    entitySetSel.value = validEntities[0].entitySetName;
    if (state.mappings[i]) {
      state.mappings[i].bindEntitySet = validEntities[0].entitySetName;
      await resolveLookupNavProp(i, attrLogical, validEntities[0].entitySetName);
    }
  } else if (prev && [...entitySetSel.options].some(o => o.value === prev)) {
    entitySetSel.value = prev;
    await resolveLookupNavProp(i, attrLogical, prev);
  }
}

function populateTargetSelect(sel: HTMLSelectElement, kind: DataverseFieldKind, currentValue: string): void {
  sel.innerHTML = "";
  sel.appendChild(new Option("(target)…", ""));
  const filtered = state.entityAttributes
    .filter(matchesKind(kind))
    // Read-only attributes (contact.accountid, fullname, createdon, …) can't
    // be loaded — offering them produces payloads Dataverse rejects.
    .filter((a) => a.writable)
    .sort((x, y) => x.logicalName.localeCompare(y.logicalName));
  for (const a of filtered) sel.appendChild(new Option(a.logicalName, a.logicalName));
  if (currentValue) {
    // Exact match, or — for lookups whose stored target is a resolved
    // navigation property like "parentcustomerid_account" — match the
    // attribute it was derived from ("parentcustomerid").
    const opt = [...sel.options].find(
      (o) => o.value === currentValue || (o.value && currentValue.startsWith(o.value + "_"))
    );
    if (opt) sel.value = opt.value;
  }
}

const navPropCache = new Map<string, string | undefined>();

/**
 * Swap a lookup row's target from the picked attribute logical name to the
 * writable navigation property for the chosen entity set. For most lookups
 * they're identical; for polymorphic ones (customer, regarding) the nav
 * property is suffixed ("parentcustomerid_account"), and binding the bare
 * attribute name fails the whole payload.
 */
async function resolveLookupNavProp(i: number, attrLogical: string, entitySetName: string): Promise<void> {
  const ref = state.entities.find((e) => e.entitySetName === entitySetName);
  if (!ref || !attrLogical || !state.entityLogicalName) return;
  const cacheKey = `${state.entityLogicalName}/${attrLogical}/${ref.logicalName}`;
  if (!navPropCache.has(cacheKey)) {
    try {
      navPropCache.set(
        cacheKey,
        await dvClient().getLookupNavigationProperty(state.entityLogicalName, attrLogical, ref.logicalName)
      );
    } catch {
      return; // metadata unavailable — leave the target as picked
    }
  }
  const navProp = navPropCache.get(cacheKey);
  const m = state.mappings[i];
  // Re-check the row still points at this attribute (async race).
  if (!m || m.kind !== "lookup" || m.bindEntitySet !== entitySetName) return;
  if (navProp && navProp !== m.target) {
    m.target = navProp;
    setStatus("info", `Using navigation property "${navProp}" for ${attrLogical} → ${ref.logicalName}.`);
  } else if (!navProp) {
    setStatus(
      "error",
      `${attrLogical} has no writable relationship to ${ref.logicalName} — ` +
        `this column would fail. Pick a different target field or entity set.`
    );
  }
}

const GUID_INPUT_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/** A constant editor plus a hook to rebuild it when the lookup config changes. */
interface ConstantEditor {
  el: HTMLElement;
  refresh: () => void;
}

/**
 * Editor for a fixed-value row's value cell. What it shows depends on the row:
 *
 *  - non-lookup kinds: a plain input, coerced by the engine like a cell value.
 *  - lookup + "by GUID": a record search against the bound entity set. Type ≥2
 *    characters, pick a record, its GUID is stored. Until an entity set is
 *    chosen there is nothing to search, so the field is disabled and says so
 *    rather than silently returning nothing.
 *  - lookup + "by alt. key"/"by text": the constant is a KEY VALUE, not a GUID,
 *    so a plain input is shown instead of a record search.
 *
 * `refresh()` rebuilds the editor in place — called when the entity set or
 * resolution below it changes, so the cell never contradicts its own config.
 */
function makeConstantEditor(
  i: number,
  getBindEntitySet: () => string | undefined
): ConstantEditor {
  const wrap = document.createElement("span");
  wrap.style.cssText = "position:relative;display:inline-block;width:100%;";

  const render = (): void => {
    wrap.innerHTML = "";
    const m = state.mappings[i];
    if (!m) return;

    if (m.kind !== "lookup") {
      wrap.title = "Fixed value applied to every record";
      wrap.appendChild(plainConstantInput(i, "fixed value"));
      return;
    }

    if (m.lookupResolution && m.lookupResolution !== "guid") {
      const what = m.lookupResolution === "alternateKey" ? "alternate key" : "text";
      wrap.title = `Fixed ${what} value matched against ${m.keyAttribute || "the key field"} on every record`;
      wrap.appendChild(plainConstantInput(i, m.keyAttribute ? `fixed ${m.keyAttribute}` : "fixed key value"));
      return;
    }

    // "by GUID": pick a record rather than making the user paste a GUID.
    wrap.title = "Fixed record applied to every record";
    const entitySet = getBindEntitySet();

    // Already picked: show it as a chip with a clear button, so it's obvious
    // a value is set (an empty-looking input previously read as "not set",
    // which is what silently produced owner-clearing mappings).
    if (m.constant !== undefined && String(m.constant) !== "") {
      wrap.appendChild(pickedRecordChip(i, render));
      return;
    }

    const input = document.createElement("input");
    input.autocomplete = "off";
    input.spellcheck = false;
    input.style.background = "#f3f9f1"; // subtle tint: this cell is not a source column
    if (!entitySet) {
      input.disabled = true;
      input.placeholder = "pick “Binds to” below first";
      wrap.appendChild(input);
      return;
    }
    input.placeholder = "search record or paste GUID…";
    wrap.appendChild(input);
    attachRecordSearch(input, wrap, i, entitySet, render);
  };

  render();
  return { el: wrap, refresh: render };
}

/** Plain text input bound to a row's `constant`. */
function plainConstantInput(i: number, placeholder: string): HTMLInputElement {
  const input = document.createElement("input");
  input.autocomplete = "off";
  input.spellcheck = false;
  input.style.background = "#f3f9f1";
  input.placeholder = placeholder;
  const cur = state.mappings[i].constant;
  input.value = cur === undefined ? "" : String(cur);
  input.addEventListener("change", () => {
    state.mappings[i].constant = input.value;
    renderRowErrors();
  });
  return input;
}

/** The "record is chosen" state: name (or GUID) plus a clear button. */
function pickedRecordChip(i: number, refresh: () => void): HTMLElement {
  const m = state.mappings[i];
  const chip = document.createElement("span");
  chip.className = "const-chip";
  const txt = document.createElement("span");
  txt.className = "txt";
  txt.textContent = m.notes ?? String(m.constant);
  txt.title = m.notes ? `${m.notes} (${String(m.constant)})` : String(m.constant);
  const clear = document.createElement("button");
  clear.type = "button";
  clear.className = "secondary";
  clear.textContent = "✕";
  clear.title = "Clear the fixed record";
  clear.addEventListener("click", () => {
    state.mappings[i].constant = "";
    state.mappings[i].notes = undefined;
    refresh();
    renderRowErrors();
  });
  chip.append(txt, clear);
  return chip;
}

/** Type-ahead record search against `entitySet`, writing the picked GUID. */
function attachRecordSearch(
  input: HTMLInputElement,
  wrap: HTMLElement,
  i: number,
  entitySet: string,
  onPicked: () => void
): void {
  const list = document.createElement("div");
  list.style.cssText =
    "position:absolute;left:0;right:0;top:100%;z-index:1000;display:none;" +
    "max-height:180px;overflow-y:auto;background:#fff;border:1px solid #8a8886;" +
    "box-shadow:0 4px 8px rgba(0,0,0,.15);font-size:12px;";
  wrap.appendChild(list);

  const close = (): void => {
    list.style.display = "none";
  };
  const message = (text: string, colour: string): void => {
    list.innerHTML = "";
    const div = document.createElement("div");
    div.style.cssText = `padding:4px 8px;color:${colour};`;
    div.textContent = text;
    list.appendChild(div);
    list.style.display = "";
  };

  let timer: ReturnType<typeof setTimeout> | undefined;
  let seq = 0;

  const search = async (term: string): Promise<void> => {
    const mySeq = ++seq;
    message("Searching…", "#605e5c");
    try {
      const client = dvClient();
      const info = await client.getEntitySetInfo(entitySet);
      const name = info.primaryNameAttribute ?? "name";
      const safe = term.replace(/'/g, "''");
      const rows = await client.queryAll(
        `${entitySet}?$select=${info.primaryIdAttribute},${name}` +
          `&$filter=contains(${name},'${safe}')&$orderby=${name}&$top=10`
      );
      if (mySeq !== seq) return; // stale response
      if (rows.length === 0) {
        message("No matches", "#605e5c");
        return;
      }
      list.innerHTML = "";
      for (const r of rows) {
        const guid = String(r[info.primaryIdAttribute] ?? "");
        const label = String(r[name] ?? guid);
        const item = document.createElement("div");
        item.textContent = label;
        item.style.cssText = "padding:4px 8px;cursor:pointer;";
        item.addEventListener("mouseenter", () => (item.style.background = "#f3f2f1"));
        item.addEventListener("mouseleave", () => (item.style.background = ""));
        item.addEventListener("mousedown", (e) => {
          e.preventDefault();
          state.mappings[i].constant = guid;
          state.mappings[i].lookupResolution = "guid";
          state.mappings[i].notes = label;
          close();
          onPicked(); // swap the input for the picked-record chip
          renderRowErrors();
        });
        list.appendChild(item);
      }
      list.style.display = "";
    } catch (err) {
      if (mySeq !== seq) return;
      message((err as Error).message, "#a4262c");
    }
  };

  input.addEventListener("input", () => {
    const term = input.value.trim();
    if (timer) clearTimeout(timer);
    // A pasted GUID is accepted directly — no search round-trip needed.
    if (GUID_INPUT_RE.test(term)) {
      state.mappings[i].constant = term.toLowerCase();
      state.mappings[i].lookupResolution = "guid";
      state.mappings[i].notes = undefined;
      close();
      onPicked();
      renderRowErrors();
      return;
    }
    if (term.length < 2) {
      close();
      return;
    }
    timer = setTimeout(() => void search(term), 300);
  });
  input.addEventListener("blur", () => setTimeout(close, 120));
}

function isDateKind(kind: DataverseFieldKind): boolean {
  return kind === "datetime" || kind === "dateonly";
}

/**
 * Offer the common unambiguous date layouts for `ColumnMapping.format`, which
 * the engine parses with `parseWithFormat`. A format loaded from a file that
 * isn't in the list is kept as an extra option rather than silently dropped.
 */
function populateFormatSelect(
  sel: HTMLSelectElement,
  kind: DataverseFieldKind,
  current: string | undefined
): void {
  const presets = kind === "dateonly"
    ? ["", "dd/MM/yyyy", "MM/dd/yyyy", "yyyy-MM-dd", "dd.MM.yyyy"]
    : ["", "dd/MM/yyyy HH:mm", "MM/dd/yyyy HH:mm", "yyyy-MM-dd HH:mm", "yyyy-MM-dd HH:mm:ss"];
  sel.innerHTML = "";
  for (const p of presets) sel.appendChild(new Option(p === "" ? "auto-detect" : p, p));
  if (current && !presets.includes(current)) sel.appendChild(new Option(current, current));
  sel.value = current ?? "";
}

/** Small caption for a lookup-config control. */
function cfgLabel(text: string, title: string): HTMLElement {
  const span = document.createElement("span");
  span.className = "lbl";
  span.textContent = text;
  if (title) span.title = title;
  return span;
}

/**
 * Fill the key-field picker from the bound entity's metadata:
 *  - alternateKey → only single-attribute alternate keys actually defined on
 *    the entity, because those are the only ones Dataverse can resolve.
 *  - text → text-ish attributes (a $filter eq against a number or a lookup is
 *    rarely what anyone means), primary name first.
 */
async function populateKeyAttributes(
  sel: HTMLSelectElement,
  i: number,
  res: "alternateKey" | "text"
): Promise<void> {
  const entitySet = state.mappings[i]?.bindEntitySet;
  const current = state.mappings[i]?.keyAttribute ?? "";
  const fill = (options: string[], empty: string): void => {
    sel.innerHTML = "";
    sel.appendChild(new Option(empty, ""));
    for (const o of options) sel.appendChild(new Option(o, o));
    // Keep a hand-edited value selectable even if it isn't in the list.
    if (current && !options.includes(current)) sel.appendChild(new Option(`${current} (custom)`, current));
    sel.value = current;
  };

  if (!entitySet) {
    fill([], "pick “Binds to” first…");
    return;
  }
  fill([], "loading fields…");
  try {
    const meta = await getBoundEntityMeta(entitySet);
    // The row may have moved on while the metadata request was in flight.
    if (state.mappings[i]?.bindEntitySet !== entitySet) return;

    if (res === "alternateKey") {
      if (meta.singleKeyAttributes.length === 0) {
        fill([], "no single-field alternate keys on this entity");
        return;
      }
      fill([...meta.singleKeyAttributes].sort(), "(alternate key)…");
      return;
    }

    const textish = meta.attributes
      .filter((a) => a.attributeType === "string" || a.attributeType === "memo")
      .map((a) => a.logicalName)
      .sort((x, y) => {
        if (x === meta.primaryNameAttribute) return -1;
        if (y === meta.primaryNameAttribute) return 1;
        return x.localeCompare(y);
      });
    fill(textish, "(text field)…");
  } catch {
    // Metadata unavailable (permissions, offline) — leave whatever is set and
    // let the combobox be used as a free-text box.
    fill(current ? [current] : [], "(key field)…");
  }
}

/**
 * Paint core's per-column validation under the row that caused it, so a
 * half-configured lookup is visible while it's being built rather than at
 * run time. Whole-mapping checks (duplicate targets, upsertKey) stay on the
 * status banner where they belong.
 */
function renderRowErrors(): void {
  const root = el<HTMLDivElement>("mappings");
  for (const node of root.querySelectorAll<HTMLDivElement>("[data-row-error]")) {
    const i = Number(node.dataset.rowError);
    const col = state.mappings[i];
    // A brand-new row has no target yet; flagging it immediately is noise.
    const errs = col && col.target ? validateColumn(col) : [];
    node.textContent = errs.join(" · ");
    node.style.display = errs.length > 0 ? "" : "none";
  }
}

function rerenderMappings(): void {
  const root = el<HTMLDivElement>("mappings");
  root.innerHTML = "";

  state.mappings.forEach((m, i) => {
    const row = document.createElement("div");
    row.className = "mapping-row";

    // A lookup row without an explicit resolution defaults to "by GUID" — and
    // that default has to be written to state here, not just displayed. The
    // dropdown below only fires on change, so an untouched row used to save
    // with no lookupResolution at all and failed validation at run time.
    if (m.kind === "lookup" && !m.lookupResolution) m.lookupResolution = "guid";

    // First cell: source column picker, or — for fixed-value rows — the
    // constant editor (a plain input, a record search, or a picked-record chip).
    let src: HTMLElement;
    let constEditor: ConstantEditor | undefined;
    if (isConstantRow(m)) {
      // Reads state, not the select below it — the select is created further
      // down and the editor renders immediately. Every path that changes the
      // bound entity writes state first, then calls refresh().
      constEditor = makeConstantEditor(i, () => state.mappings[i]?.bindEntitySet);
      src = constEditor.el;
    } else {
      const srcSel = document.createElement("select");
      srcSel.appendChild(new Option("(source)…", ""));
      const cols = state.selectedTable?.columns ?? [];
      for (const c of cols) srcSel.appendChild(new Option(c, c, m.source === c, m.source === c));
      srcSel.value = m.source ?? "";
      srcSel.addEventListener("change", () => {
        state.mappings[i].source = srcSel.value;
      });
      src = srcSel;
    }

    const kind = document.createElement("select");
    const kinds: DataverseFieldKind[] = [
      "string", "memo", "integer", "decimal", "money", "double",
      "boolean", "datetime", "dateonly", "uniqueidentifier",
      "lookup", "choice", "multichoice", "status", "state",
    ];
    for (const k of kinds) kind.appendChild(new Option(k, k, m.kind === k, m.kind === k));

    const tgt = document.createElement("select");
    populateTargetSelect(tgt, m.kind, m.target);
    const tgtCell = enhanceSelect(tgt);
    tgt.addEventListener("change", () => {
      state.mappings[i].target = tgt.value;
      renderRowErrors();
      if (state.mappings[i].kind === "lookup" && tgt.value) {
        applyLookupTargets(entitySetSel, tgt.value, i)
          .then(() => {
            // The entity set may have been auto-picked — the value editor and
            // key-field picker depend on it.
            constEditor?.refresh();
            applyResolutionVisibility();
            renderRowErrors();
          })
          .catch(() => {});
      } else if (CHOICE_KINDS.includes(state.mappings[i].kind) && tgt.value) {
        applyOptionLabels(tgt.value, i).catch(() => {});
      }
    });

    // Lookup config sub-row: labelled, and ordered the way it has to be filled
    // in — what it binds to, how the value is matched, then which field.
    const lookupCfg = document.createElement("div");
    lookupCfg.className = "lookup-cfg";
    lookupCfg.style.display = m.kind === "lookup" ? "" : "none";

    const entitySetSel = document.createElement("select");
    entitySetSel.appendChild(new Option("(entity set)…", ""));
    for (const e of state.entities) {
      entitySetSel.appendChild(new Option(
        `${e.displayName || e.logicalName} (${e.entitySetName})`,
        e.entitySetName
      ));
    }
    if (m.bindEntitySet) entitySetSel.value = m.bindEntitySet;
    const entitySetCell = enhanceSelect(entitySetSel);

    const resSel = document.createElement("select");
    resSel.appendChild(new Option("by record (GUID)", "guid"));
    resSel.appendChild(new Option("by alternate key", "alternateKey"));
    resSel.appendChild(new Option("by text match", "text"));
    resSel.value = m.lookupResolution ?? "guid";

    // Key field: a picker over the bound entity's attributes rather than a
    // free-text logical name. Populated async; the raw select doubles as a
    // text box via enhanceSelect, so a value can still be typed if metadata
    // is unavailable.
    const keyAttrSel = document.createElement("select");
    keyAttrSel.appendChild(new Option("(key field)…", ""));
    if (m.keyAttribute) {
      keyAttrSel.appendChild(new Option(m.keyAttribute, m.keyAttribute));
      keyAttrSel.value = m.keyAttribute;
    }
    const keyAttrCell = enhanceSelect(keyAttrSel);
    keyAttrSel.addEventListener("change", () => {
      state.mappings[i].keyAttribute = keyAttrSel.value.trim() || undefined;
      constEditor?.refresh();
      renderRowErrors();
    });

    const createBox = document.createElement("input");
    createBox.type = "checkbox";
    createBox.checked = m.createIfMissing === true;
    createBox.addEventListener("change", () => {
      state.mappings[i].createIfMissing = createBox.checked ? true : undefined;
      renderRowErrors();
    });
    const createLabel = document.createElement("label");
    createLabel.className = "check";
    createLabel.style.cssText = "margin:0;font-weight:400;font-size:11px;";
    createLabel.append(createBox, document.createTextNode("create the record if no match"));

    const dupSel = document.createElement("select");
    dupSel.appendChild(new Option("fail the row", "error"));
    dupSel.appendChild(new Option("take the first match", "first"));
    dupSel.value = m.duplicateBehavior ?? "error";
    dupSel.addEventListener("change", () => {
      state.mappings[i].duplicateBehavior = dupSel.value === "first" ? "first" : undefined;
    });

    const hint = document.createElement("div");
    hint.className = "hint";

    // Label/control pairs, shown or hidden per resolution mode.
    const bindLbl = cfgLabel("Binds to", "The entity this lookup points at");
    const matchLbl = cfgLabel("Match by", "How a source value is turned into a record id");
    const keyLbl = cfgLabel("Key field", "Attribute on the bound entity the value is matched against");
    const missLbl = cfgLabel("If missing", "");
    const dupLbl = cfgLabel("If ambiguous", "What to do when a text match hits more than one record");

    lookupCfg.append(
      bindLbl, entitySetCell,
      matchLbl, resSel,
      keyLbl, keyAttrCell,
      missLbl, createLabel,
      dupLbl, dupSel,
      hint
    );

    /** Show only the controls the chosen resolution actually uses. */
    const applyResolutionVisibility = (): void => {
      const res = state.mappings[i]?.lookupResolution ?? "guid";
      const needsKey = res === "alternateKey" || res === "text";
      keyLbl.style.display = needsKey ? "" : "none";
      // enhanceSelect sets display:inline-block inline, so restore it
      // explicitly rather than clearing to the (inline) span default.
      keyAttrCell.style.display = needsKey ? "inline-block" : "none";
      for (const elm of [missLbl, createLabel, dupLbl, dupSel]) {
        elm.style.display = res === "text" ? "" : "none";
      }
      hint.textContent =
        res === "guid"
          ? "The value is already a record GUID."
          : res === "alternateKey"
            ? "The value is matched against a Dataverse alternate key — fastest, but the key must exist on the target entity."
            : "The value is matched against any text field with a filter query. Slower, and it must match exactly one record.";
      if (needsKey) void populateKeyAttributes(keyAttrSel, i, res);
    };
    applyResolutionVisibility();

    entitySetSel.addEventListener("change", () => {
      state.mappings[i].bindEntitySet = entitySetSel.value || undefined;
      // Key fields belong to the previous entity — drop them and repopulate.
      state.mappings[i].keyAttribute = undefined;
      keyAttrSel.value = "";
      applyResolutionVisibility();
      constEditor?.refresh();
      renderRowErrors();
      // The bound entity determines the writable navigation property name.
      const attr = tgt.value || state.mappings[i].target;
      if (entitySetSel.value && attr) {
        resolveLookupNavProp(i, attr, entitySetSel.value).catch(() => {});
      }
    });

    resSel.addEventListener("change", () => {
      const res = resSel.value as "guid" | "alternateKey" | "text";
      state.mappings[i].lookupResolution = res;
      // createIfMissing/duplicateBehavior are text-only; carrying them over
      // to another mode is a validation error, so clear them on the way out.
      if (res !== "text") {
        state.mappings[i].createIfMissing = undefined;
        state.mappings[i].duplicateBehavior = undefined;
        createBox.checked = false;
        dupSel.value = "error";
      }
      // A GUID and a key value are different things — don't silently keep one
      // as the other when switching modes.
      if (isConstantRow(state.mappings[i])) {
        state.mappings[i].constant = "";
        state.mappings[i].notes = undefined;
      }
      applyResolutionVisibility();
      constEditor?.refresh();
      renderRowErrors();
    });

    // Date columns get an optional explicit parse format. Excel hands us real
    // Date objects most of the time, but a text column like "03/04/2025" is
    // ambiguous and the engine has to be told which way round it is.
    const fmtRow = document.createElement("div");
    fmtRow.className = "lookup-cfg";
    fmtRow.style.display = isDateKind(m.kind) ? "" : "none";
    const fmtSel = document.createElement("select");
    populateFormatSelect(fmtSel, m.kind, m.format);
    fmtSel.addEventListener("change", () => {
      state.mappings[i].format = fmtSel.value || undefined;
    });
    fmtRow.append(
      cfgLabel("Date format", "Tokens: yyyy MM M dd d HH H mm ss — anything else is a literal"),
      fmtSel
    );

    kind.addEventListener("change", () => {
      const nextKind = kind.value as DataverseFieldKind;
      state.mappings[i].kind = nextKind;
      // Lookup-only settings would fail validation on a non-lookup column.
      if (nextKind !== "lookup") {
        state.mappings[i].lookupResolution = undefined;
        state.mappings[i].keyAttribute = undefined;
        state.mappings[i].bindEntitySet = undefined;
        state.mappings[i].createIfMissing = undefined;
        state.mappings[i].duplicateBehavior = undefined;
      } else if (!state.mappings[i].lookupResolution) {
        state.mappings[i].lookupResolution = "guid";
      }
      // A format only means anything on a date column.
      if (!isDateKind(nextKind)) state.mappings[i].format = undefined;
      // A fixed-value row's value editor depends on the kind (plain input vs
      // record search), so rebuild the grid.
      if (isConstantRow(state.mappings[i])) {
        state.mappings[i].constant = "";
        state.mappings[i].notes = undefined;
        rerenderMappings();
        return;
      }
      populateTargetSelect(tgt, nextKind, state.mappings[i].target);
      state.mappings[i].target = tgt.value;
      lookupCfg.style.display = nextKind === "lookup" ? "" : "none";
      fmtRow.style.display = isDateKind(nextKind) ? "" : "none";
      populateFormatSelect(fmtSel, nextKind, state.mappings[i].format);
      applyResolutionVisibility();
      renderRowErrors();
      if (nextKind === "lookup" && tgt.value) {
        applyLookupTargets(entitySetSel, tgt.value, i).catch(() => {});
      } else if (CHOICE_KINDS.includes(nextKind) && tgt.value) {
        applyOptionLabels(tgt.value, i).catch(() => {});
      }
    });

    const remove = document.createElement("button");
    remove.className = "secondary";
    remove.textContent = "✕";
    remove.title = "Remove";
    remove.addEventListener("click", () => {
      state.mappings.splice(i, 1);
      rerenderMappings();
    });

    row.append(src, kind, tgtCell, remove);
    root.appendChild(row);
    root.appendChild(lookupCfg);
    root.appendChild(fmtRow);

    // Per-row validation messages, filled in by renderRowErrors().
    const err = document.createElement("div");
    err.className = "row-error";
    err.dataset.rowError = String(i);
    err.style.display = "none";
    root.appendChild(err);

    // Pre-resolve targets for already-configured lookup rows
    if (m.kind === "lookup" && m.target) {
      applyLookupTargets(entitySetSel, m.target, i)
        .then(() => {
          constEditor?.refresh();
          applyResolutionVisibility();
          renderRowErrors();
        })
        .catch(() => {});
    } else if (CHOICE_KINDS.includes(m.kind) && m.target && !m.optionMap) {
      // Auto-suggested choice rows arrive without an optionMap — fetch the
      // labels so spreadsheet cells can contain "Warm" instead of 2.
      applyOptionLabels(m.target, i).catch(() => {});
    }
  });

  renderRowErrors();
}

/* -------------------------------------------------------------------------- */
/* Run / save / load                                                           */
/* -------------------------------------------------------------------------- */

function rerenderRunPlan(): void {
  const root = el<HTMLDivElement>("planSteps");
  root.innerHTML = "";
  if (state.planSteps.length === 0) {
    root.textContent = "No steps yet.";
    root.style.fontSize = "11px";
    root.style.color = "#605e5c";
    renderPlanErrors(); // clears any stale message from a previous plan
    return;
  }
  root.style.fontSize = "";
  root.style.color = "";
  for (let i = 0; i < state.planSteps.length; i++) {
    const step = state.planSteps[i];
    const row = document.createElement("div");
    row.className = "plan-row";

    const id = document.createElement("input");
    id.placeholder = "step id";
    id.value = step.id;
    id.title = "Unique step id";
    id.addEventListener("change", () => {
      state.planSteps[i].id = id.value.trim();
      persistRunPlan();
    });

    const mapping = document.createElement("input");
    mapping.placeholder = "./contacts.dvmap.json";
    mapping.value = step.mapping;
    mapping.title = "Mapping path";
    mapping.addEventListener("change", () => {
      state.planSteps[i].mapping = mapping.value.trim();
      persistRunPlan();
    });

    const workbook = document.createElement("input");
    workbook.placeholder = "./customers.xlsx";
    workbook.value = step.workbook;
    workbook.title = "Workbook path";
    workbook.addEventListener("change", () => {
      state.planSteps[i].workbook = workbook.value.trim();
      persistRunPlan();
    });

    const stage = document.createElement("input");
    stage.type = "number";
    stage.min = "1";
    stage.value = String(step.stage ?? 1);
    stage.title = "Stage (same stage runs in parallel)";
    stage.addEventListener("change", () => {
      const n = Number(stage.value);
      state.planSteps[i].stage = Number.isInteger(n) && n > 0 ? n : 1;
      persistRunPlan();
    });

    const dependsOn = document.createElement("input");
    dependsOn.placeholder = "step-a,step-b";
    dependsOn.value = (step.dependsOn ?? []).join(",");
    dependsOn.title = "Optional explicit dependencies";
    dependsOn.addEventListener("change", () => {
      const parts = dependsOn.value
        .split(",")
        .map((x) => x.trim())
        .filter(Boolean);
      state.planSteps[i].dependsOn = parts.length > 0 ? parts : undefined;
      persistRunPlan();
    });

    const links = document.createElement("input");
    links.placeholder = "step1:parentcustomerid_account:accountnumber";
    links.value = formatAlternateKeyLinks(step.alternateKeyLinks);
    links.title =
      "Alternate-key links (comma-separated fromStep:lookupTarget:keyAttribute)";
    links.addEventListener("change", () => {
      try {
        state.planSteps[i].alternateKeyLinks = parseAlternateKeyLinksText(links.value);
        persistRunPlan();
      } catch (e) {
        setStatus("error", (e as Error).message);
        links.focus();
      }
    });

    // Per-step overrides. These round-tripped through the editor before but
    // were invisible and uneditable — a plan authored by hand could not be
    // touched here without losing track of them.
    const overrides = document.createElement("input");
    overrides.placeholder = "dryRun,maxErrors=5";
    overrides.value = formatStepOverrides(step.overrides);
    overrides.title =
      "Per-step overrides, comma-separated: maxErrors=N, concurrency=N, " +
      "notifyUrl=<url>, dryRun, user, noFailedRows";
    overrides.addEventListener("change", () => {
      try {
        state.planSteps[i].overrides = parseStepOverridesText(overrides.value);
        persistRunPlan();
        renderPlanErrors();
      } catch (e) {
        setStatus("error", (e as Error).message);
        overrides.focus();
      }
    });

    const refreshWrap = document.createElement("label");
    refreshWrap.style.display = "flex";
    refreshWrap.style.alignItems = "center";
    refreshWrap.style.gap = "4px";
    refreshWrap.style.fontSize = "11px";
    const refresh = document.createElement("input");
    refresh.type = "checkbox";
    refresh.style.width = "auto";
    refresh.checked = step.refresh === true;
    refresh.addEventListener("change", () => {
      state.planSteps[i].refresh = refresh.checked ? true : undefined;
      persistRunPlan();
    });
    const refreshText = document.createElement("span");
    refreshText.textContent = "refresh";
    refreshWrap.append(refresh, refreshText);

    const remove = document.createElement("button");
    remove.className = "secondary";
    remove.textContent = "✕";
    remove.title = "Remove step";
    remove.addEventListener("click", () => {
      state.planSteps.splice(i, 1);
      persistRunPlan();
      rerenderRunPlan();
    });

    row.append(id, mapping, workbook, stage, dependsOn, links, overrides, refreshWrap, remove);
    root.appendChild(row);
  }

  renderPlanErrors();
}

/** Render `validateRunPlan` output under the step grid. */
function renderPlanErrors(): void {
  const box = el<HTMLDivElement>("planErrors");
  const errs = state.planSteps.length > 0 ? validateRunPlan(buildRunPlan()) : [];
  box.textContent = errs.join(" · ");
  box.style.display = errs.length > 0 ? "" : "none";
}

function writePlanMetaToForm(): void {
  el<HTMLInputElement>("planName").value = state.planMeta.name;
  el<HTMLInputElement>("planDescription").value = state.planMeta.description ?? "";
  el<HTMLInputElement>("planStopOnError").checked = state.planMeta.stopOnError;
}

const OVERRIDE_INT_KEYS = ["maxErrors", "concurrency"] as const;

/** "dryRun,maxErrors=5" → RunPlanStepOverrides (or undefined when empty). */
function parseStepOverridesText(text: string): RunPlanStepOverrides | undefined {
  const out: RunPlanStepOverrides = {};
  for (const part of text.split(",").map((s) => s.trim()).filter(Boolean)) {
    const [rawKey, rawValue] = part.split("=").map((s) => s.trim());
    const key = rawKey;
    if ((OVERRIDE_INT_KEYS as readonly string[]).includes(key)) {
      const n = Number(rawValue);
      if (!Number.isInteger(n) || n < 0) throw new Error(`${key} must be a whole number, got "${rawValue}"`);
      out[key as (typeof OVERRIDE_INT_KEYS)[number]] = n;
    } else if (key === "notifyUrl") {
      if (!rawValue) throw new Error("notifyUrl needs a value, e.g. notifyUrl=https://…");
      out.notifyUrl = rawValue;
    } else if (key === "dryRun") {
      out.dryRun = true;
    } else if (key === "user") {
      out.user = true;
    } else if (key === "noFailedRows") {
      out.failedRows = false;
    } else {
      throw new Error(
        `Unknown override "${key}". Use maxErrors=N, concurrency=N, notifyUrl=<url>, dryRun, user or noFailedRows.`
      );
    }
  }
  return Object.keys(out).length > 0 ? out : undefined;
}

function formatStepOverrides(o: RunPlanStepOverrides | undefined): string {
  if (!o) return "";
  const parts: string[] = [];
  if (o.maxErrors !== undefined) parts.push(`maxErrors=${o.maxErrors}`);
  if (o.concurrency !== undefined) parts.push(`concurrency=${o.concurrency}`);
  if (o.notifyUrl) parts.push(`notifyUrl=${o.notifyUrl}`);
  if (o.dryRun) parts.push("dryRun");
  if (o.user) parts.push("user");
  if (o.failedRows === false) parts.push("noFailedRows");
  return parts.join(",");
}

function formatAlternateKeyLinks(links: RunPlanStep["alternateKeyLinks"]): string {
  if (!links || links.length === 0) return "";
  return links.map((x) => `${x.fromStep}:${x.lookupTarget}:${x.keyAttribute}`).join(",");
}

function parseAlternateKeyLinksText(text: string): RunPlanStep["alternateKeyLinks"] {
  const raw = text.trim();
  if (!raw) return undefined;
  return raw.split(",").map((item) => {
    const parts = item.split(":").map((x) => x.trim());
    if (parts.length !== 3) {
      throw new Error(
        `Invalid alternate-key link "${item}". Use fromStep:lookupTarget:keyAttribute`
      );
    }
    const [fromStep, lookupTarget, keyAttribute] = parts;
    if (!fromStep || !lookupTarget || !keyAttribute) {
      throw new Error(
        `Invalid alternate-key link "${item}". Use fromStep:lookupTarget:keyAttribute`
      );
    }
    return { fromStep, lookupTarget, keyAttribute };
  });
}

function onPlanAddStep(): void {
  const next = state.planSteps.length + 1;
  state.planSteps.push({
    id: `step-${next}`,
    mapping: "",
    workbook: "",
    stage: next,
  });
  persistRunPlan();
  rerenderRunPlan();
}

function onPlanAddFromCurrent(): void {
  const mapping = buildMapping();
  const next = state.planSteps.length + 1;
  state.planSteps.push({
    id: `step-${next}`,
    mapping: `./${mapping.name.replace(/\W+/g, "-").toLowerCase()}.dvmap.json`,
    workbook: "./workbook.xlsx",
    stage: next,
  });
  persistRunPlan();
  rerenderRunPlan();
  setStatus("info", "Added a step from the current mapping. Adjust paths and dependencies.");
}

function buildRunPlan(): RunPlan {
  return {
    schemaVersion: RUN_PLAN_SCHEMA_VERSION,
    name: state.planMeta.name || "Run plan",
    description: state.planMeta.description,
    createdAt: state.planMeta.createdAt,
    stopOnError: state.planMeta.stopOnError,
    steps: state.planSteps,
  };
}

function persistRunPlan(): void {
  Office.context.document.settings.set(PLAN_SETTINGS_KEY, JSON.stringify(buildRunPlan()));
  Office.context.document.settings.saveAsync();
}

function restoreRunPlan(): void {
  const raw = Office.context.document.settings.get(PLAN_SETTINGS_KEY);
  if (!raw || typeof raw !== "string") return;
  try {
    applyRunPlan(parseRunPlan(JSON.parse(raw)));
  } catch {
    state.planSteps = [];
  }
}

/** Adopt a parsed plan wholesale — steps and the metadata around them. */
function applyRunPlan(plan: RunPlan): void {
  state.planSteps = plan.steps;
  state.planMeta = {
    name: plan.name,
    description: plan.description,
    stopOnError: plan.stopOnError,
    createdAt: plan.createdAt,
  };
  writePlanMetaToForm();
  rerenderRunPlan();
}

function buildMapping(): Mapping {
  const sourceName = state.fileSource?.tableName ?? state.selectedTable?.name ?? "";
  const m: Mapping = {
    // Anything the pane has no control for (createdAt, logDir, future schema
    // additions) is carried over from the mapping that was loaded.
    ...state.mappingExtras,
    schemaVersion: SCHEMA_VERSION,
    name: el<HTMLInputElement>("mappingName").value.trim() || sourceName || "Mapping",
    environmentUrl: state.environmentUrl,
    targetEntitySet: state.entitySet,
    sourceTable: sourceName,
    sourceSheet: state.fileSource ? undefined : state.selectedTable?.worksheetName,
    columns: state.mappings,
    conflictMode: "insert",
    batchSize: 100,
    maxErrors: 0,
    logDir: state.mappingExtras.logDir ?? "./logs",
    concurrency: 1,
    bypassCustomLogic: false,
    skipUnchanged: false,
  };
  readOptionsIntoMapping(m);
  return m;
}

/**
 * In-pane replacement for window.confirm, which Office task pane webviews
 * don't support (it throws a script error). Renders a modal overlay with
 * OK/Cancel and resolves with the choice.
 */
function confirmDialog(message: string, okLabel = "OK"): Promise<boolean> {
  return new Promise((resolve) => {
    const overlay = document.createElement("div");
    overlay.style.cssText =
      "position:fixed;inset:0;background:rgba(0,0,0,.35);z-index:2000;" +
      "display:flex;align-items:center;justify-content:center;padding:16px;";
    const box = document.createElement("div");
    box.style.cssText =
      "background:#fff;border-radius:4px;box-shadow:0 8px 24px rgba(0,0,0,.3);" +
      "max-width:320px;width:100%;padding:14px;font-size:12px;";
    const text = document.createElement("div");
    text.style.cssText = "white-space:pre-wrap;margin-bottom:12px;";
    text.textContent = message;
    const row = document.createElement("div");
    row.style.cssText = "display:flex;gap:8px;justify-content:flex-end;";
    const done = (v: boolean): void => {
      overlay.remove();
      resolve(v);
    };
    const cancelBtn = document.createElement("button");
    cancelBtn.type = "button";
    cancelBtn.className = "secondary";
    cancelBtn.style.width = "auto";
    cancelBtn.textContent = "Cancel";
    cancelBtn.addEventListener("click", () => done(false));
    const okBtn = document.createElement("button");
    okBtn.type = "button";
    okBtn.style.width = "auto";
    okBtn.textContent = okLabel;
    okBtn.addEventListener("click", () => done(true));
    row.append(cancelBtn, okBtn);
    box.append(text, row);
    overlay.appendChild(box);
    overlay.addEventListener("click", (e) => {
      if (e.target === overlay) done(false);
    });
    document.body.appendChild(overlay);
    okBtn.focus();
  });
}

/** Non-null while an import is running; aborting it cancels the run. */
let runController: AbortController | null = null;

/**
 * Closing the pane mid-run kills the JS context: batches in flight complete
 * server-side, but nothing further is sent and no summary is shown — the
 * import is effectively aborted without accounting. Warn before unload.
 * (Best effort: browsers/webviews show their own generic dialog, and Office
 * can close the pane without firing it. The Cancel button is the clean path.)
 */
function onBeforeUnload(e: BeforeUnloadEvent): void {
  e.preventDefault();
  e.returnValue = "An import is still running — closing now abandons it mid-way.";
}

function setRunning(running: boolean): void {
  el<HTMLButtonElement>("run").disabled = running;
  el<HTMLButtonElement>("run").style.display = running ? "none" : "";
  el<HTMLButtonElement>("cancelRun").style.display = running ? "" : "none";
  el<HTMLButtonElement>("cancelRun").disabled = false;
  el<HTMLButtonElement>("cancelRun").textContent = "Cancel import";
  if (running) window.addEventListener("beforeunload", onBeforeUnload);
  else window.removeEventListener("beforeunload", onBeforeUnload);
}

async function onCancelRun(): Promise<void> {
  if (!runController) return;
  const ok = await confirmDialog(
    "Really cancel this import?\n\n" +
      "Rows already sent to Dataverse stay there — cancelling only stops the remaining rows. " +
      "The summary will show how far it got.",
    "Cancel import"
  );
  // The run may have finished while the dialog was open.
  if (!ok || !runController) return;
  runController.abort();
  const btn = el<HTMLButtonElement>("cancelRun");
  btn.disabled = true;
  btn.textContent = "Cancelling…"; // in-flight batches are finishing
  setStatus("info", "Cancelling — waiting for in-flight batches to finish…");
}

/** Start over: clear source, target, mappings, options, and panel state. Keeps sign-in, environment, and profiles. */
async function onResetAll(): Promise<void> {
  if (runController) {
    setStatus("error", "An import is running — cancel it before resetting.");
    return;
  }
  const ok = await confirmDialog(
    "Start over?\n\n" +
      "This clears the source table, target entity, all column mappings, import options, " +
      "and the mapping remembered for this workbook. Your sign-in, environment URL, and " +
      "saved profiles are kept.",
    "Reset"
  );
  if (!ok) return;

  exitCreateTableMode();
  ctCols = [];
  ctFixedOwner = null;
  el<HTMLInputElement>("ctOwner").value = "";

  state.selectedTable = null;
  el<HTMLSelectElement>("table").value = "";
  state.entitySet = "";
  state.entityLogicalName = "";
  state.entityAttributes = [];
  state.mappings = [];
  lookupTargetsCache.clear();
  optionLabelsCache.clear();
  navPropCache.clear();

  el<HTMLSelectElement>("solution").value = "";
  if (state.entities.length > 0) {
    // Signed in: keep the loaded entity list, just drop the selection.
    renderEntityOptions(null);
  }
  el<HTMLSelectElement>("entity").value = "";

  el<HTMLSelectElement>("conflictMode").value = "insert";
  el<HTMLInputElement>("upsertKey").value = "";
  el<HTMLSelectElement>("syncAction").value = "deactivate";
  el<HTMLInputElement>("batchSize").value = "100";
  el<HTMLInputElement>("concurrency").value = "1";
  el<HTMLInputElement>("bypassCustomLogic").checked = false;
  el<HTMLInputElement>("skipUnchanged").checked = false;
  el<HTMLInputElement>("dryRun").checked = false;
  el<HTMLInputElement>("mappingName").value = "";
  el<HTMLInputElement>("mappingDescription").value = "";
  el<HTMLInputElement>("maxErrors").value = "0";
  el<HTMLInputElement>("notifyUrl").value = "";
  const impersonate = el<HTMLInputElement>("impersonateUser");
  impersonate.value = "";
  impersonate.dataset.guid = "";
  impersonate.dataset.label = "";
  state.mappingExtras = {};
  state.fileSource = null;
  renderFileSource();
  updateOptionsVisibility();

  // Forget the per-workbook remembered mapping so it doesn't restore on reload.
  Office.context.document.settings.remove(SETTINGS_KEY);
  Office.context.document.settings.remove(PLAN_SETTINGS_KEY);
  Office.context.document.settings.saveAsync();
  state.planSteps = [];
  state.planMeta = { name: "Run plan", stopOnError: true };
  writePlanMetaToForm();
  rerenderRunPlan();
  state.checkpoint = null;
  persistCheckpoint();
  renderCheckpoint();
  lastFailedRows = null;

  el<HTMLDivElement>("progressWrap").style.display = "none";
  el<HTMLDivElement>("logWrap").style.display = "none";
  rerenderMappings();
  setStatus("info", "Reset. Pick a source table and target entity to start again.");
}

async function onRun(opts: { resume?: boolean } = {}): Promise<void> {
  if (runController) return; // already running
  try {
    if (createMode) {
      await createTableThenImport();
      // Still in create mode → validation failed inside the panel; stay put.
      if (createMode) return;
    }
    const mapping = buildMapping();
    const errs = validateMapping(mapping);
    if (errs.length > 0) {
      setStatus("error", "Mapping has errors: " + errs.join("; "));
      renderRowErrors();
      return;
    }
    // Advisories (e.g. overriddencreatedon with upsert) — confirm, don't block.
    const warnings = mappingWarnings(mapping);
    if (warnings.length > 0) {
      const proceed = await confirmDialog(
        "Heads up:\n\n" + warnings.join("\n\n"),
        "Run anyway"
      );
      if (!proceed) {
        setStatus("info", "Run cancelled.");
        return;
      }
    }
    const dryRun = el<HTMLInputElement>("dryRun").checked;

    setStatus("info", state.fileSource ? "Reading file…" : "Reading table…");
    const { rows, headers } = await readSourceRows(mapping.sourceTable);

    // Resume picks up at the checkpoint the last interrupted run left behind.
    // It's only valid against the same source — row N means nothing once the
    // data underneath has changed shape.
    let startOffset = 0;
    if (opts.resume && state.checkpoint) {
      if (state.checkpoint.sourceTable !== mapping.sourceTable || state.checkpoint.total !== rows.length) {
        setStatus(
          "error",
          "The checkpoint was taken against a different source (or the row count changed) — " +
            "discard it and run from the start."
        );
        return;
      }
      startOffset = state.checkpoint.offset;
    }

    setStatus(
      "info",
      `${dryRun ? "Dry run: planning" : "Loading"} ${rows.length - startOffset} rows…` +
        (startOffset > 0 ? ` (resuming at row ${startOffset + 1})` : "")
    );
    runController = new AbortController();
    setRunning(true);
    const requestLog: RequestLogEntry[] = [];
    const successLog: RowSuccess[] = [];
    let retries = 0;
    const client = new DataverseClient({
      environmentUrl: mapping.environmentUrl,
      getToken: makeTokenProvider(mapping.environmentUrl),
      onRequest: (entry) => requestLog.push(entry),
      retry: {
        onRetry: (info) => {
          retries++;
          // status 0 = the request threw before responding — a dropped
          // connection, usually the machine sleeping mid-run.
          setStatus(
            "info",
            `Retrying (${info.attempt}): ${info.status === 0 ? (info.error ?? "network error") : info.status} — ` +
              `waiting ${(info.delayMs / 1000).toFixed(1)}s…`
          );
        },
      },
    });
    // Track how far we got so a cancel can be resumed.
    let processed = startOffset;
    const result = await loadRows({
      mapping,
      rows,
      client,
      dryRun,
      startOffset,
      signal: runController.signal,
      onProgress: (e) => {
        if (e.type === "row-success") successLog.push(e.success);
        if (e.type === "batch") {
          processed = e.processed;
          setProgress(e.processed, e.total, {
            created: e.created,
            updated: e.updated,
            failed: e.failed,
            skipped: e.skipped,
          });
        }
      },
    });

    const summary =
      `${result.created} created, ${result.updated} updated, ` +
      `${result.skipped} skipped` +
      (result.unchanged > 0 ? ` (${result.unchanged} unchanged)` : "") +
      (result.removed > 0 ? `, ${result.removed} removed` : "") +
      `, ${result.failed} failed`;
    const prefix = dryRun ? "Dry run — nothing was written. " : "";
    const retryNote = retries > 0 ? ` (recovered from ${retries} dropped or throttled request${retries === 1 ? "" : "s"})` : "";
    if (result.cancelled) {
      setStatus("info", `${prefix}Cancelled. ${summary} — skipped rows were not attempted.`);
    } else if (result.failed === 0) {
      setStatus("success", `${prefix}Done. ${summary}${retryNote}.`);
    } else {
      setStatus(
        "error",
        `${prefix}Done with errors. ${summary}${retryNote}. First: ${result.errors[0]?.message ?? ""}`
      );
    }

    // A cancelled real run is the only case worth resuming: a completed run
    // has nothing left, and a dry run wrote nothing to pick up from.
    state.checkpoint =
      result.cancelled && !dryRun
        ? {
            offset: processed,
            total: rows.length,
            sourceTable: mapping.sourceTable,
            savedAt: new Date().toISOString(),
          }
        : null;
    persistCheckpoint();
    renderCheckpoint();

    persistMapping(mapping);
    lastFailedRows = collectFailedRows(rows, headers, result);
    showRunLog(requestLog, successLog, result, mapping);
    // Anonymous usage event — fields documented in TELEMETRY.md.
    trackTelemetry("addin_run", {
      mode: mapping.conflictMode,
      rows: telemetryBucket(result.total),
      failed: telemetryBucket(result.failed),
      outcome: result.cancelled ? "cancelled" : result.failed > 0 ? "partial" : "ok",
      dryRun: String(dryRun),
    });
  } catch (e) {
    setStatus("error", (e as Error).message);
  } finally {
    runController = null;
    setRunning(false);
  }
}

/** Rows from the picked file if there is one, else the open workbook's table. */
async function readSourceRows(
  tableName: string
): Promise<{ rows: Array<Record<string, unknown>>; headers: string[] }> {
  if (state.fileSource) {
    return { rows: state.fileSource.rows, headers: state.fileSource.headers };
  }
  const { rows } = await readTable(tableName);
  const headers = state.selectedTable?.columns ?? Object.keys(rows[0] ?? {});
  return { rows, headers };
}

/** The source rows behind a failed run, for the re-run workbook. */
let lastFailedRows: { headers: string[]; rows: Array<Record<string, unknown>> } | null = null;

function collectFailedRows(
  rows: Array<Record<string, unknown>>,
  headers: string[],
  result: LoadResult
): { headers: string[]; rows: Array<Record<string, unknown>> } | null {
  if (result.errors.length === 0) return null;
  // rowIndex is 0-based into the source rows; dedupe because a row can carry
  // more than one error.
  const indexes = [...new Set(result.errors.map((e) => e.rowIndex))].sort((a, b) => a - b);
  const failed = indexes.map((i) => rows[i]).filter((r): r is Record<string, unknown> => r != null);
  return failed.length > 0 ? { headers, rows: failed } : null;
}

async function onDownloadFailedRows(): Promise<void> {
  if (!lastFailedRows) return;
  try {
    const buf = await writeRowsToBuffer(lastFailedRows.headers, lastFailedRows.rows, "FailedRows");
    downloadBlob(
      new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
      `failed-rows-${new Date().toISOString().slice(0, 19).replace(/[:T]/g, "")}.xlsx`
    );
  } catch (e) {
    setStatus("error", `Couldn't build the failed-rows workbook: ${(e as Error).message}`);
  }
}

/** Copy a view's bytes into a standalone ArrayBuffer that Blob will accept. */
function toArrayBuffer(view: Uint8Array): ArrayBuffer {
  return view.buffer.slice(view.byteOffset, view.byteOffset + view.byteLength) as ArrayBuffer;
}

/** Trigger a browser download for a generated file. */
function downloadBlob(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

/* -------------------------------------------------------------------------- */
/* Checkpoint (resume after cancel)                                            */
/* -------------------------------------------------------------------------- */

function persistCheckpoint(): void {
  if (state.checkpoint) {
    Office.context.document.settings.set(CHECKPOINT_KEY, JSON.stringify(state.checkpoint));
  } else {
    Office.context.document.settings.remove(CHECKPOINT_KEY);
  }
  Office.context.document.settings.saveAsync();
}

function restoreCheckpoint(): void {
  const raw = Office.context.document.settings.get(CHECKPOINT_KEY);
  if (!raw || typeof raw !== "string") return;
  try {
    const c = JSON.parse(raw) as Checkpoint;
    if (typeof c.offset === "number" && typeof c.total === "number" && c.sourceTable) {
      state.checkpoint = c;
    }
  } catch {
    state.checkpoint = null;
  }
  renderCheckpoint();
}

function renderCheckpoint(): void {
  const row = el<HTMLDivElement>("resumeRow");
  const c = state.checkpoint;
  if (!c) {
    row.style.display = "none";
    return;
  }
  row.style.display = "";
  el<HTMLSpanElement>("resumeText").textContent =
    `Interrupted run on "${c.sourceTable}" stopped at row ${c.offset} of ${c.total} ` +
    `(${new Date(c.savedAt).toLocaleString()}).`;
}

function persistMapping(m: Mapping): void {
  Office.context.document.settings.set(SETTINGS_KEY, JSON.stringify(m));
  Office.context.document.settings.saveAsync();
}

/** Start loading entities into the entity select, then restore any saved entity selection. */
function triggerLoadEntities(): void {
  const entSel = el<HTMLSelectElement>("entity");
  entSel.innerHTML = `<option value="">Loading entities…</option>`;
  entSel.disabled = false;
  loadEntities()
    .then(async () => {
      const saved = state.entitySet;
      if (saved && [...entSel.options].some(o => o.value === saved)) {
        entSel.value = saved;
        const logical = entSel.selectedOptions[0]?.dataset.logical;
        if (logical) await loadEntityAttributes(logical);
      }
    })
    .catch(() => {
      entSel.innerHTML = `<option value="">Could not load — click Sign in to retry</option>`;
      entSel.disabled = true;
    });
}

function showRunLog(
  requestLog: RequestLogEntry[],
  successLog: RowSuccess[],
  result: LoadResult,
  mapping: Mapping
): void {
  el<HTMLDivElement>("logWrap").style.display = "";
  const errorDiv = el<HTMLDivElement>("errorDetails");
  if (result.errors.length > 0) {
    errorDiv.style.display = "";
    errorDiv.textContent = result.errors
      .map(e => `row ${e.rowIndex}: [${e.code ?? e.httpStatus ?? ""}] ${e.message}`)
      .join("\n");
  } else {
    errorDiv.style.display = "none";
  }
  el<HTMLButtonElement>("downloadLog").onclick = () =>
    downloadRunLog(requestLog, successLog, result, mapping);
  // The CLI writes a failed-rows workbook next to its logs; the pane can only
  // hand it over as a download.
  const failedBtn = el<HTMLButtonElement>("downloadFailed");
  failedBtn.style.display = lastFailedRows ? "" : "none";
  failedBtn.textContent = lastFailedRows
    ? `Download ${lastFailedRows.rows.length} failed row${lastFailedRows.rows.length === 1 ? "" : "s"}…`
    : "Download failed rows…";
}

function downloadRunLog(
  requestLog: RequestLogEntry[],
  successLog: RowSuccess[],
  result: LoadResult,
  mapping: Mapping
): void {
  const { errors, ...summary } = result;
  const lines: string[] = [];
  lines.push(JSON.stringify({ event: "summary", ...summary }));
  for (const r of requestLog) lines.push(JSON.stringify(r));
  for (const s of successLog) lines.push(JSON.stringify({ event: "success", ...s }));
  for (const e of errors) lines.push(JSON.stringify({ event: "error", ...e }));
  const blob = new Blob([lines.join("\n") + "\n"], { type: "application/x-ndjson" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  const startedAt = new Date(result.startedAt);
  const datePart = startedAt.toISOString().slice(0, 10);
  const timePart = startedAt.toISOString().slice(11, 19).replace(/:/g, "-");
  const stem = (mapping.name || mapping.sourceTable).replace(/[^\w.-]/g, "_");
  a.download = `${stem}_${datePart}_${timePart}.jsonl`;
  a.click();
  URL.revokeObjectURL(url);
}

function restoreMapping(): void {
  const raw = Office.context.document.settings.get(SETTINGS_KEY);
  if (!raw || typeof raw !== "string") return;
  try {
    const m = parseMapping(JSON.parse(raw));
    state.environmentUrl = m.environmentUrl;
    el<HTMLInputElement>("env").value = m.environmentUrl;
    state.entitySet = m.targetEntitySet;
    state.mappings = m.columns;
    state.mappingExtras = { createdAt: m.createdAt, logDir: m.logDir };
    const tableSel = el<HTMLSelectElement>("table");
    if ([...tableSel.options].some((o) => o.value === m.sourceTable)) {
      tableSel.value = m.sourceTable;
      state.selectedTable = state.tables.find((t) => t.name === m.sourceTable) ?? null;
    }
    writeOptionsFromMapping(m);
    rerenderMappings();
  } catch {
    // ignore: stored mapping was a different schema version
  }
}

async function onSave(): Promise<void> {
  const m = buildMapping();
  // Saving a mapping the CLI will reject is the worst place to find out, so
  // flag it here — but let it be saved anyway, since a work-in-progress
  // mapping is still worth keeping.
  const errs = validateMapping(m);
  renderRowErrors();
  if (errs.length > 0) {
    const proceed = await confirmDialog(
      "This mapping has errors and `dvload run` will refuse it:\n\n" +
        errs.map((e) => `• ${e}`).join("\n") +
        "\n\nSave it anyway?",
      "Save anyway"
    );
    if (!proceed) {
      setStatus("error", "Save cancelled — " + errs.join("; "));
      return;
    }
  }
  const blob = new Blob([serializeMapping(m)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `${m.name.replace(/\W+/g, "-").toLowerCase()}.dvmap.json`;
  a.click();
  URL.revokeObjectURL(url);
}

async function onPlanSave(): Promise<void> {
  const plan = buildRunPlan();
  // Same contract as mappings: warn about what run-all will reject, but don't
  // block saving work in progress.
  const errs = validateRunPlan(plan);
  renderPlanErrors();
  if (errs.length > 0) {
    const proceed = await confirmDialog(
      "This run plan has errors and `dvload run-all` will refuse it:\n\n" +
        errs.map((e) => `• ${e}`).join("\n") +
        "\n\nSave it anyway?",
      "Save anyway"
    );
    if (!proceed) {
      setStatus("error", "Save cancelled — " + errs.join("; "));
      return;
    }
  }
  const blob = new Blob([serializeRunPlan(plan)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  const stem = (plan.name || "run-plan")
    .replace(/\W+/g, "-")
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "")
    .toLowerCase();
  a.download = `${stem}.dvplan.json`;
  a.click();
  URL.revokeObjectURL(url);
}

function onLoad(): void {
  const input = document.createElement("input");
  input.type = "file";
  input.accept = ".json,.dvmap.json,application/json";
  input.onchange = async () => {
    const file = input.files?.[0];
    if (!file) return;
    const text = await file.text();
    try {
      const m = parseMapping(JSON.parse(text));
      state.environmentUrl = m.environmentUrl;
      el<HTMLInputElement>("env").value = m.environmentUrl;
      state.entitySet = m.targetEntitySet;
      state.mappings = m.columns;
      // Keep the fields the pane has no control for so re-saving doesn't
      // quietly rewrite someone's hand-authored mapping.
      state.mappingExtras = { createdAt: m.createdAt, logDir: m.logDir };
      writeOptionsFromMapping(m);
      rerenderMappings();
      setStatus("success", `Loaded ${m.name}.`);
    } catch (e) {
      setStatus("error", `Couldn't parse mapping: ${(e as Error).message}`);
    }
  };
  input.click();
}

function onPlanLoad(): void {
  const input = document.createElement("input");
  input.type = "file";
  input.accept = ".json,.dvplan.json,application/json";
  input.onchange = async () => {
    const file = input.files?.[0];
    if (!file) return;
    const text = await file.text();
    try {
      const plan = parseRunPlan(JSON.parse(text), file.name);
      applyRunPlan(plan);
      persistRunPlan();
      setStatus("success", `Loaded run plan (${plan.steps.length} step${plan.steps.length === 1 ? "" : "s"}).`);
    } catch (e) {
      setStatus("error", `Couldn't parse run plan: ${(e as Error).message}`);
    }
  };
  input.click();
}
