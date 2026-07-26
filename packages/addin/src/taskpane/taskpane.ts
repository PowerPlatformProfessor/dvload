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
  mappingWarnings,
  type Mapping,
  type ColumnMapping,
  type DataverseFieldKind,
  type RequestLogEntry,
  type RowSuccess,
  type LoadResult,
  type PqtArchive,
  readPqt,
  mappingFromPqt,
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
  RUN_PLAN_SCHEMA_VERSION,
} from "@dvload/core";
import { initAuth, getAccount, signIn, makeTokenProvider, devModeBanner } from "../auth.js";
import { listTables, readTable, type TableInfo } from "../excel.js";
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

interface AppState {
  account: { username: string } | null;
  environmentUrl: string;
  tables: TableInfo[];
  selectedTable: TableInfo | null;
  entitySet: string;
  entityLogicalName: string;
  entityAttributes: EntityAttribute[];
  entities: Array<{ logicalName: string; entitySetName: string; displayName: string; metadataId: string }>;
  mappings: ColumnMapping[];
  planSteps: RunPlanStep[];
}

const state: AppState = {
  account: null,
  environmentUrl: "",
  tables: [],
  selectedTable: null,
  entitySet: "",
  entityLogicalName: "",
  entityAttributes: [],
  entities: [],
  mappings: [],
  planSteps: [],
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
  el<HTMLButtonElement>("run").addEventListener("click", onRun);
  el<HTMLButtonElement>("cancelRun").addEventListener("click", onCancelRun);
  el<HTMLButtonElement>("resetAll").addEventListener("click", onResetAll);
  el<HTMLButtonElement>("save").addEventListener("click", onSave);
  el<HTMLButtonElement>("load").addEventListener("click", onLoad);
  el<HTMLButtonElement>("importPqt").addEventListener("click", onImportPqt);
  el<HTMLButtonElement>("pqtUse").addEventListener("click", onUsePqtMapping);
  el<HTMLButtonElement>("pqtCopyM").addEventListener("click", onCopyPqtM);
  el<HTMLButtonElement>("planAddStep").addEventListener("click", onPlanAddStep);
  el<HTMLButtonElement>("planFromCurrent").addEventListener("click", onPlanAddFromCurrent);
  el<HTMLButtonElement>("planSave").addEventListener("click", onPlanSave);
  el<HTMLButtonElement>("planLoad").addEventListener("click", onPlanLoad);
  el<HTMLSelectElement>("conflictMode").addEventListener("change", updateOptionsVisibility);
  updateOptionsVisibility();
  rerenderRunPlan();

  // Type-to-filter on the big pickers. The native selects stay in the DOM as
  // the source of truth; the combobox is a UI layer over them.
  enhanceSelect(el<HTMLSelectElement>("solution"));
  enhanceSelect(el<HTMLSelectElement>("entity"));

  initCtOwnerSearch();

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

function onImportPqt(): void {
  const input = document.createElement("input");
  input.type = "file";
  input.accept = ".pqt,application/zip";
  input.onchange = async () => {
    const file = input.files?.[0];
    if (!file) return;
    try {
      currentPqt = await readPqt(await file.arrayBuffer());
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
}

function writeOptionsFromMapping(m: Mapping): void {
  el<HTMLSelectElement>("conflictMode").value = m.conflictMode;
  el<HTMLInputElement>("upsertKey").value = (m.upsertKey ?? []).join(", ");
  el<HTMLSelectElement>("syncAction").value = m.syncAction ?? "deactivate";
  el<HTMLInputElement>("batchSize").value = String(m.batchSize);
  el<HTMLInputElement>("concurrency").value = String(m.concurrency ?? 1);
  el<HTMLInputElement>("bypassCustomLogic").checked = m.bypassCustomLogic ?? false;
  el<HTMLInputElement>("skipUnchanged").checked = m.skipUnchanged ?? false;
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
function initCtOwnerSearch(): void {
  const input = el<HTMLInputElement>("ctOwner");
  const list = el<HTMLDivElement>("ctOwnerList");
  const close = (): void => { list.style.display = "none"; };
  let timer: ReturnType<typeof setTimeout> | undefined;
  let seq = 0;

  input.addEventListener("input", () => {
    const term = input.value.trim();
    if (!term) { ctFixedOwner = null; close(); return; }
    if (/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(term)) {
      ctFixedOwner = { guid: term.toLowerCase(), label: term.toLowerCase() };
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
            ctFixedOwner = { guid, label };
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
    input.value = ctFixedOwner?.label ?? "";
  });
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

async function loadEntityAttributes(logical: string): Promise<void> {
  state.entityLogicalName = logical;
  lookupTargetsCache.clear();
  optionLabelsCache.clear();
  const client = new DataverseClient({
    environmentUrl: state.environmentUrl,
    getToken: makeTokenProvider(state.environmentUrl),
  });
  const def = await client.getEntityDefinition(logical);
  // Dataverse may return Attributes as an inline array (OData v4) or wrapped in { value: [] }
  const rawAttrs: Array<Record<string, unknown>> = Array.isArray(def.Attributes)
    ? (def.Attributes as Array<Record<string, unknown>>)
    : ((def.Attributes as { value?: Array<Record<string, unknown>> })?.value ?? []);
  state.entityAttributes = rawAttrs
    .map((a) => ({
      logicalName: String(a["LogicalName"] ?? ""),
      attributeType: String(a["AttributeType"] ?? "").toLowerCase(),
      format: a["Format"] != null ? String(a["Format"]) : undefined,
      writable: a["IsValidForCreate"] === true || a["IsValidForUpdate"] === true,
    }))
    .filter(a => a.logicalName);

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

/**
 * Editor for a fixed-value row's value cell. Non-lookup kinds get a plain
 * input (coerced by the engine exactly like a cell value). Lookup kinds get
 * a record search: type ≥2 characters, matching records from the bound
 * entity set (e.g. Users or Teams for ownerid) are fetched and picking one
 * stores its GUID as the constant.
 */
function makeConstantEditor(
  m: ColumnMapping,
  i: number,
  getBindEntitySet: () => string | undefined
): HTMLElement {
  const wrap = document.createElement("span");
  wrap.style.cssText = "position:relative;display:inline-block;width:100%;";
  wrap.title = "Fixed value applied to every record";

  const input = document.createElement("input");
  input.autocomplete = "off";
  input.spellcheck = false;
  input.style.background = "#f3f9f1"; // subtle tint: this cell is not a source column
  wrap.appendChild(input);

  if (m.kind !== "lookup") {
    input.placeholder = "fixed value";
    input.value = m.constant === undefined ? "" : String(m.constant);
    input.addEventListener("change", () => {
      state.mappings[i].constant = input.value;
    });
    return wrap;
  }

  // Lookup: async record search against the bound entity set.
  input.placeholder = "search record… (pick entity set below)";
  // Redisplay the picked record's name (kept in notes), else the raw GUID.
  input.value = m.notes ?? (m.constant === undefined ? "" : String(m.constant));

  const list = document.createElement("div");
  list.style.cssText =
    "position:absolute;left:0;right:0;top:100%;z-index:1000;display:none;" +
    "max-height:180px;overflow-y:auto;background:#fff;border:1px solid #8a8886;" +
    "box-shadow:0 4px 8px rgba(0,0,0,.15);font-size:12px;";
  wrap.appendChild(list);

  const close = (): void => {
    list.style.display = "none";
  };

  let timer: ReturnType<typeof setTimeout> | undefined;
  let seq = 0;

  const search = async (term: string): Promise<void> => {
    const entitySet = getBindEntitySet();
    if (!entitySet) {
      list.innerHTML = `<div style="padding:4px 8px;color:#605e5c;">Pick the entity set first (below)</div>`;
      list.style.display = "";
      return;
    }
    const mySeq = ++seq;
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
      list.innerHTML = "";
      if (rows.length === 0) {
        list.innerHTML = `<div style="padding:4px 8px;color:#605e5c;">No matches</div>`;
      }
      for (const r of rows) {
        const guid = String(r[info.primaryIdAttribute] ?? "");
        const label = String(r[name] ?? guid);
        const item = document.createElement("div");
        item.textContent = label;
        item.style.cssText = "padding:4px 8px;cursor:pointer;";
        item.addEventListener("mousedown", (e) => {
          e.preventDefault();
          state.mappings[i].constant = guid;
          state.mappings[i].lookupResolution = "guid";
          state.mappings[i].notes = label;
          input.value = label;
          close();
        });
        list.appendChild(item);
      }
      list.style.display = "";
    } catch (err) {
      if (mySeq !== seq) return;
      list.innerHTML = `<div style="padding:4px 8px;color:#a4262c;">${(err as Error).message}</div>`;
      list.style.display = "";
    }
  };

  input.addEventListener("input", () => {
    const term = input.value.trim();
    // A pasted GUID is accepted directly — no search round-trip needed.
    if (/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(term)) {
      state.mappings[i].constant = term.toLowerCase();
      state.mappings[i].lookupResolution = "guid";
      state.mappings[i].notes = undefined;
      close();
      return;
    }
    if (timer) clearTimeout(timer);
    if (term.length < 2) {
      close();
      return;
    }
    timer = setTimeout(() => void search(term), 300);
  });
  input.addEventListener("blur", () => {
    close();
    // Revert half-typed searches to the last picked record.
    input.value = state.mappings[i].notes ?? String(state.mappings[i].constant ?? "");
  });

  return wrap;
}

function rerenderMappings(): void {
  const root = el<HTMLDivElement>("mappings");
  root.innerHTML = "";

  state.mappings.forEach((m, i) => {
    const row = document.createElement("div");
    row.className = "mapping-row";

    // First cell: source column picker, or — for fixed-value rows — the
    // constant editor (a plain input, or a record search for lookups).
    let src: HTMLElement;
    if (isConstantRow(m)) {
      src = makeConstantEditor(m, i, () => entitySetSel.value || m.bindEntitySet);
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
      if (state.mappings[i].kind === "lookup" && tgt.value) {
        applyLookupTargets(entitySetSel, tgt.value, i).catch(() => {});
      } else if (CHOICE_KINDS.includes(state.mappings[i].kind) && tgt.value) {
        applyOptionLabels(tgt.value, i).catch(() => {});
      }
    });

    // Lookup config sub-row
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
    entitySetSel.addEventListener("change", () => {
      state.mappings[i].bindEntitySet = entitySetSel.value || undefined;
      // The bound entity determines the writable navigation property name.
      const attr = tgt.value || state.mappings[i].target;
      if (entitySetSel.value && attr) {
        resolveLookupNavProp(i, attr, entitySetSel.value).catch(() => {});
      }
    });

    const resSel = document.createElement("select");
    resSel.appendChild(new Option("by GUID", "guid"));
    resSel.appendChild(new Option("by alt. key", "alternateKey"));
    resSel.value = m.lookupResolution ?? "guid";

    const keyAttrInput = document.createElement("input");
    keyAttrInput.placeholder = "key attribute";
    keyAttrInput.value = m.keyAttribute ?? "";
    keyAttrInput.style.display = m.lookupResolution === "alternateKey" ? "" : "none";
    keyAttrInput.addEventListener("change", () => {
      state.mappings[i].keyAttribute = keyAttrInput.value.trim() || undefined;
    });

    resSel.addEventListener("change", () => {
      state.mappings[i].lookupResolution = resSel.value as "guid" | "alternateKey";
      keyAttrInput.style.display = resSel.value === "alternateKey" ? "" : "none";
    });

    lookupCfg.append(entitySetCell, resSel, keyAttrInput);

    kind.addEventListener("change", () => {
      state.mappings[i].kind = kind.value as DataverseFieldKind;
      // A fixed-value row's value editor depends on the kind (plain input vs
      // record search), so rebuild the grid.
      if (isConstantRow(state.mappings[i])) {
        state.mappings[i].constant = "";
        state.mappings[i].notes = undefined;
        rerenderMappings();
        return;
      }
      populateTargetSelect(tgt, kind.value as DataverseFieldKind, state.mappings[i].target);
      state.mappings[i].target = tgt.value;
      lookupCfg.style.display = kind.value === "lookup" ? "" : "none";
      if (kind.value === "lookup" && tgt.value) {
        applyLookupTargets(entitySetSel, tgt.value, i).catch(() => {});
      } else if (CHOICE_KINDS.includes(kind.value as DataverseFieldKind) && tgt.value) {
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

    // Pre-resolve targets for already-configured lookup rows
    if (m.kind === "lookup" && m.target) {
      applyLookupTargets(entitySetSel, m.target, i).catch(() => {});
    } else if (CHOICE_KINDS.includes(m.kind) && m.target && !m.optionMap) {
      // Auto-suggested choice rows arrive without an optionMap — fetch the
      // labels so spreadsheet cells can contain "Warm" instead of 2.
      applyOptionLabels(m.target, i).catch(() => {});
    }
  });
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
      state.planSteps[i].refresh = refresh.checked || undefined;
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

    row.append(id, mapping, workbook, stage, dependsOn, links, refreshWrap, remove);
    root.appendChild(row);
  }
}

function formatAlternateKeyLinks(links: RunPlanStep["alternateKeyLinks"]): string {
  if (!links || links.length === 0) return "";
  return links.map((x) => `${x.fromStep}:${x.lookupTarget}:${x.keyAttribute}`).join(",");
}

function parseAlternateKeyLinksText(text: string): RunPlanStep["alternateKeyLinks"] {
  const raw = text.trim();
  if (!raw) return undefined;
  return raw.split(",").map((item) => {
    const [fromStep, lookupTarget, keyAttribute] = item.split(":").map((x) => x.trim());
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
    name: "Run plan",
    stopOnError: true,
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
    const plan = parseRunPlan(JSON.parse(raw));
    state.planSteps = plan.steps;
    rerenderRunPlan();
  } catch {
    state.planSteps = [];
  }
}

function buildMapping(): Mapping {
  const m: Mapping = {
    schemaVersion: SCHEMA_VERSION,
    name: state.selectedTable?.name ?? "Mapping",
    environmentUrl: state.environmentUrl,
    targetEntitySet: state.entitySet,
    sourceTable: state.selectedTable?.name ?? "",
    sourceSheet: state.selectedTable?.worksheetName,
    columns: state.mappings,
    conflictMode: "insert",
    batchSize: 100,
    maxErrors: 0,
    logDir: "./logs",
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
  updateOptionsVisibility();

  // Forget the per-workbook remembered mapping so it doesn't restore on reload.
  Office.context.document.settings.remove(SETTINGS_KEY);
  Office.context.document.settings.remove(PLAN_SETTINGS_KEY);
  Office.context.document.settings.saveAsync();
  state.planSteps = [];
  rerenderRunPlan();

  el<HTMLDivElement>("progressWrap").style.display = "none";
  el<HTMLDivElement>("logWrap").style.display = "none";
  rerenderMappings();
  setStatus("info", "Reset. Pick a source table and target entity to start again.");
}

async function onRun(): Promise<void> {
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
    setStatus("info", "Reading table…");
    const { rows } = await readTable(mapping.sourceTable);

    setStatus("info", `Loading ${rows.length} rows…`);
    runController = new AbortController();
    setRunning(true);
    const requestLog: RequestLogEntry[] = [];
    const successLog: RowSuccess[] = [];
    const client = new DataverseClient({
      environmentUrl: mapping.environmentUrl,
      getToken: makeTokenProvider(mapping.environmentUrl),
      onRequest: (entry) => requestLog.push(entry),
    });
    const result = await loadRows({
      mapping,
      rows,
      client,
      signal: runController.signal,
      onProgress: (e) => {
        if (e.type === "row-success") successLog.push(e.success);
        if (e.type === "batch") {
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
    if (result.cancelled) {
      setStatus("info", `Cancelled. ${summary} — skipped rows were not attempted.`);
    } else if (result.failed === 0) {
      setStatus("success", `Done. ${summary}.`);
    } else {
      setStatus(
        "error",
        `Done with errors. ${summary}. First: ${result.errors[0]?.message ?? ""}`
      );
    }

    persistMapping(mapping);
    showRunLog(requestLog, successLog, result, mapping);
    // Anonymous usage event — fields documented in TELEMETRY.md.
    trackTelemetry("addin_run", {
      mode: mapping.conflictMode,
      rows: telemetryBucket(result.total),
      failed: telemetryBucket(result.failed),
      outcome: result.cancelled ? "cancelled" : result.failed > 0 ? "partial" : "ok",
    });
  } catch (e) {
    setStatus("error", (e as Error).message);
  } finally {
    runController = null;
    setRunning(false);
  }
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
  const blob = new Blob([serializeMapping(m)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `${m.name.replace(/\W+/g, "-").toLowerCase()}.dvmap.json`;
  a.click();
  URL.revokeObjectURL(url);
}

function onPlanSave(): void {
  const plan = buildRunPlan();
  const blob = new Blob([serializeRunPlan(plan)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  const stem = (plan.name || "run-plan").replace(/\W+/g, "-").toLowerCase();
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
      state.planSteps = plan.steps;
      persistRunPlan();
      rerenderRunPlan();
      setStatus("success", `Loaded run plan (${plan.steps.length} step${plan.steps.length === 1 ? "" : "s"}).`);
    } catch (e) {
      setStatus("error", `Couldn't parse run plan: ${(e as Error).message}`);
    }
  };
  input.click();
}
