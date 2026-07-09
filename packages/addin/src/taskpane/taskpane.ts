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
  SCHEMA_VERSION,
} from "@dvload/core";
import { initAuth, getAccount, signIn, makeTokenProvider, devModeBanner } from "../auth.js";
import { listTables, readTable, type TableInfo } from "../excel.js";
import { suggestMappings, suggestionsToMappings } from "../suggest.js";

const SETTINGS_KEY = "dvload:lastMapping";
const PROFILES_KEY = "dvload:profiles";

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
  entityAttributes: string[];
  mappings: ColumnMapping[];
}

const state: AppState = {
  account: null,
  environmentUrl: "",
  tables: [],
  selectedTable: null,
  entitySet: "",
  entityAttributes: [],
  mappings: [],
};

Office.onReady(async () => {
  await initAuth();
  renderDevModeBanner();
  await refreshAccountUI();

  state.tables = await listTables();
  populateTablePicker();
  restoreMapping();

  // Profiles
  renderProfilePicker();
  el<HTMLSelectElement>("profile").addEventListener("change", (e) => {
    const url = (e.target as HTMLSelectElement).value;
    if (!url) return;
    state.environmentUrl = url;
    el<HTMLInputElement>("env").value = url;
    el<HTMLButtonElement>("profileDelete").disabled = false;
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
});

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
  el<HTMLSpanElement>("who").textContent = acc ? acc.username : "Not signed in";
  el<HTMLSelectElement>("entity").disabled = !acc;
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
  const sel = el<HTMLSelectElement>("entity");
  sel.innerHTML = `<option value="">Select an entity…</option>`;
  for (const e of entities.sort((a, b) => a.LogicalName.localeCompare(b.LogicalName))) {
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
  const client = new DataverseClient({
    environmentUrl: state.environmentUrl,
    getToken: makeTokenProvider(state.environmentUrl),
  });
  const def = (await client.getEntityDefinition(logical)) as {
    Attributes?: { value?: Array<{ LogicalName: string }> };
  };
  state.entityAttributes = (def.Attributes?.value ?? []).map((a) => a.LogicalName);

  // If the user hasn't started mapping yet, auto-suggest. Otherwise leave
  // existing mappings alone — they can click "Suggest" to fill in the rest.
  if (state.mappings.length === 0 && state.selectedTable) {
    const suggestions = suggestMappings(state.selectedTable.columns, state.entityAttributes);
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
  const suggestions = suggestMappings(state.selectedTable.columns, state.entityAttributes, {
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

    const tgt = document.createElement("input");
    tgt.placeholder = "logical name";
    tgt.value = m.target;
    tgt.setAttribute("list", "attrs");
    tgt.addEventListener("change", () => {
      state.mappings[i].target = tgt.value.trim();
    });

    const kind = document.createElement("select");
    const kinds: DataverseFieldKind[] = [
      "string", "memo", "integer", "decimal", "money", "double",
      "boolean", "datetime", "dateonly", "uniqueidentifier",
      "lookup", "choice", "multichoice", "status", "state",
    ];
    for (const k of kinds) kind.appendChild(new Option(k, k, m.kind === k, m.kind === k));
    kind.addEventListener("change", () => {
      state.mappings[i].kind = kind.value as DataverseFieldKind;
    });

    const remove = document.createElement("button");
    remove.className = "secondary";
    remove.textContent = "✕";
    remove.title = "Remove";
    remove.addEventListener("click", () => {
      state.mappings.splice(i, 1);
      rerenderMappings();
    });

    row.append(src, tgt, kind, remove);
    root.appendChild(row);
  });

  // datalist of known target attributes
  let dl = document.getElementById("attrs") as HTMLDataListElement | null;
  if (!dl) {
    dl = document.createElement("datalist");
    dl.id = "attrs";
    document.body.appendChild(dl);
  }
  dl.innerHTML = "";
  for (const a of state.entityAttributes) dl.appendChild(new Option(a));
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
  };
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
    const client = new DataverseClient({
      environmentUrl: mapping.environmentUrl,
      getToken: makeTokenProvider(mapping.environmentUrl),
    });
    const result = await loadRows({
      mapping,
      rows,
      client,
      onProgress: (e) => {
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
      `${result.skipped} skipped, ${result.failed} failed`;
    if (result.failed === 0) {
      setStatus("success", `Done. ${summary}.`);
    } else {
      setStatus(
        "error",
        `Done with errors. ${summary}. First: ${result.errors[0]?.message ?? ""}`
      );
    }

    persistMapping(mapping);
  } catch (e) {
    setStatus("error", (e as Error).message);
  }
}

function persistMapping(m: Mapping): void {
  Office.context.document.settings.set(SETTINGS_KEY, JSON.stringify(m));
  Office.context.document.settings.saveAsync();
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
      rerenderMappings();
      setStatus("success", `Loaded ${m.name}.`);
    } catch (e) {
      setStatus("error", `Couldn't parse mapping: ${(e as Error).message}`);
    }
  };
  input.click();
}
