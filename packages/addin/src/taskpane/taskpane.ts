// Task pane controller. Wires up the HTML in taskpane.html: list tables,
// fetch entities, build a column mapping, run the import. Persists the most
// recent mapping per workbook in Office's Settings store.

import {
  DataverseClient,
  loadRows,
  parseMapping,
  serializeMapping,
  validateMapping,
  type Mapping,
  type ColumnMapping,
  type DataverseFieldKind,
  type RequestLogEntry,
  type RowSuccess,
  type LoadResult,
  SCHEMA_VERSION,
} from "@dvload/core";
import { initAuth, getAccount, signIn, makeTokenProvider, devModeBanner } from "../auth.js";
import { listTables, readTable, type TableInfo } from "../excel.js";
import { suggestMappings, suggestionsToMappings } from "../suggest.js";

const SETTINGS_KEY = "dvload:lastMapping";
const PROFILES_KEY = "dvload:profiles";

interface EntityAttribute {
  logicalName: string;
  attributeType: string;
  format?: string;
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
  entities: Array<{ logicalName: string; entitySetName: string; displayName: string }>;
  mappings: ColumnMapping[];
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
};

const lookupTargetsCache = new Map<string, string[]>();
const optionLabelsCache = new Map<string, Record<string, number>>();

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
  el<HTMLSelectElement>("entity").addEventListener("change", onPickEntity);
  el<HTMLButtonElement>("addMap").addEventListener("click", () => addMapping());
  el<HTMLButtonElement>("suggest").addEventListener("click", onSuggest);
  el<HTMLButtonElement>("run").addEventListener("click", onRun);
  el<HTMLButtonElement>("save").addEventListener("click", onSave);
  el<HTMLButtonElement>("load").addEventListener("click", onLoad);
  el<HTMLSelectElement>("conflictMode").addEventListener("change", updateOptionsVisibility);
  updateOptionsVisibility();
});

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

async function loadEntities(): Promise<void> {
  const client = new DataverseClient({
    environmentUrl: state.environmentUrl,
    getToken: makeTokenProvider(state.environmentUrl),
  });
  const entities = await client.listEntities();
  const sorted = entities.sort((a, b) => a.LogicalName.localeCompare(b.LogicalName));
  state.entities = sorted.map(e => ({
    logicalName: e.LogicalName,
    entitySetName: e.EntitySetName,
    displayName: e.DisplayName,
  }));
  const sel = el<HTMLSelectElement>("entity");
  sel.innerHTML = `<option value="">Select an entity…</option>`;
  for (const e of sorted) {
    const opt = document.createElement("option");
    opt.value = e.EntitySetName;
    opt.textContent = `${e.DisplayName || e.LogicalName} (${e.EntitySetName})`;
    opt.dataset.logical = e.LogicalName;
    sel.appendChild(opt);
  }
  sel.disabled = false;
}

async function onPickEntity(e: Event): Promise<void> {
  const sel = e.target as HTMLSelectElement;
  state.entitySet = sel.value;
  if (!state.entitySet) return;
  const logical = sel.selectedOptions[0]?.dataset.logical ?? state.entitySet.replace(/s$/, "");
  await loadEntityAttributes(logical);
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
    }))
    .filter(a => a.logicalName);

  // If the user hasn't started mapping yet, auto-suggest. Otherwise leave
  // existing mappings alone — they can click "Suggest" to fill in the rest.
  if (state.mappings.length === 0 && state.selectedTable) {
    const suggestions = suggestMappings(state.selectedTable.columns, state.entityAttributes.map(a => a.logicalName));
    if (suggestions.length > 0) {
      state.mappings = suggestionsToMappings(suggestions);
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
  const suggestions = suggestMappings(state.selectedTable.columns, state.entityAttributes.map(a => a.logicalName), {
    excludeSources: state.mappings.map((m) => m.source).filter(Boolean),
    excludeTargets: state.mappings.map((m) => m.target).filter(Boolean),
  });
  if (suggestions.length === 0) {
    setStatus("info", "No additional confident suggestions.");
    return;
  }
  state.mappings.push(...suggestionsToMappings(suggestions));
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
  rerenderMappings();
}

function addMapping(seed?: ColumnMapping): void {
  state.mappings.push(
    seed ?? { source: "", target: "", kind: "string", treatEmptyAsNull: true }
  );
  rerenderMappings();
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

type EntityRecord = { logicalName: string; entitySetName: string; displayName: string };

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
    if (state.mappings[i]) state.mappings[i].bindEntitySet = validEntities[0].entitySetName;
  } else if (prev && [...entitySetSel.options].some(o => o.value === prev)) {
    entitySetSel.value = prev;
  }
}

function populateTargetSelect(sel: HTMLSelectElement, kind: DataverseFieldKind, currentValue: string): void {
  sel.innerHTML = "";
  sel.appendChild(new Option("(target)…", ""));
  const filtered = state.entityAttributes
    .filter(matchesKind(kind))
    .sort((x, y) => x.logicalName.localeCompare(y.logicalName));
  for (const a of filtered) sel.appendChild(new Option(a.logicalName, a.logicalName));
  if (currentValue && [...sel.options].some(o => o.value === currentValue)) {
    sel.value = currentValue;
  }
}

function rerenderMappings(): void {
  const root = el<HTMLDivElement>("mappings");
  root.innerHTML = "";

  state.mappings.forEach((m, i) => {
    const row = document.createElement("div");
    row.className = "mapping-row";

    const src = document.createElement("select");
    src.appendChild(new Option("(source)…", ""));
    const cols = state.selectedTable?.columns ?? [];
    for (const c of cols) src.appendChild(new Option(c, c, m.source === c, m.source === c));
    src.value = m.source;
    src.addEventListener("change", () => {
      state.mappings[i].source = src.value;
    });

    const kind = document.createElement("select");
    const kinds: DataverseFieldKind[] = [
      "string", "memo", "integer", "decimal", "money", "double",
      "boolean", "datetime", "dateonly", "uniqueidentifier",
      "lookup", "choice", "multichoice", "status", "state",
    ];
    for (const k of kinds) kind.appendChild(new Option(k, k, m.kind === k, m.kind === k));

    const tgt = document.createElement("select");
    populateTargetSelect(tgt, m.kind, m.target);
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
    entitySetSel.addEventListener("change", () => {
      state.mappings[i].bindEntitySet = entitySetSel.value || undefined;
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

    lookupCfg.append(entitySetSel, resSel, keyAttrInput);

    kind.addEventListener("change", () => {
      state.mappings[i].kind = kind.value as DataverseFieldKind;
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

    row.append(src, kind, tgt, remove);
    root.appendChild(row);
    root.appendChild(lookupCfg);

    // Pre-resolve targets for already-configured lookup rows
    if (m.kind === "lookup" && m.target) {
      applyLookupTargets(entitySetSel, m.target, i).catch(() => {});
    }
  });
}

/* -------------------------------------------------------------------------- */
/* Run / save / load                                                           */
/* -------------------------------------------------------------------------- */

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

async function onRun(): Promise<void> {
  try {
    const mapping = buildMapping();
    const errs = validateMapping(mapping);
    if (errs.length > 0) {
      setStatus("error", "Mapping has errors: " + errs.join("; "));
      return;
    }
    setStatus("info", "Reading table…");
    const { rows } = await readTable(mapping.sourceTable);

    setStatus("info", `Loading ${rows.length} rows…`);
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
    if (result.failed === 0) {
      setStatus("success", `Done. ${summary}.`);
    } else {
      setStatus(
        "error",
        `Done with errors. ${summary}. First: ${result.errors[0]?.message ?? ""}`
      );
    }

    persistMapping(mapping);
    showRunLog(requestLog, successLog, result, mapping);
  } catch (e) {
    setStatus("error", (e as Error).message);
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
