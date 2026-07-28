// Host adapter: the only place that knows whether the UI is running inside
// Excel or in an ordinary browser tab.
//
// Both front ends are the same bundle served by `dvload serve`. Everything
// downstream of this module — entity picker, mapping builder, validation,
// the run itself — operates on table metadata and rows, not on Excel, so it
// needs no idea which host it is in. Four things genuinely differ:
//
//   settings      Office stores them in the workbook (so a mapping travels
//                 with the file); a browser has no workbook, so localStorage.
//   workbook      Office can enumerate and read the open workbook's tables;
//                 a browser can only read a file the user picks.
//   openExternal  Office hosts often swallow target=_blank.
//   ready         Office.onReady vs DOMContentLoaded.
//
// Keeping the split this narrow is deliberate: if it starts to sprawl, that
// is the signal to extract a proper packages/ui, not to widen this file.

import type { SourceRow } from "@dvload/core";
import { listTables, readTable, getWorkbookBytes, type TableInfo } from "./excel.js";

export type HostKind = "office" | "browser";

export interface SettingsStore {
  get(key: string): string | null;
  set(key: string, value: string): void;
  remove(key: string): void;
  /** Office needs an explicit round trip; localStorage doesn't. */
  save(): void;
}

export interface WorkbookSource {
  /**
   * Whether the host exposes a live workbook. False in a browser, where the
   * only source is a file the user picks — which the pane already supports.
   */
  readonly canReadOpenWorkbook: boolean;
  listTables(): Promise<TableInfo[]>;
  readTable(name: string): Promise<{ headers: string[]; rows: SourceRow[] }>;
  /** Raw .xlsx bytes, for pulling the embedded Power Query out. */
  getWorkbookBytes(): Promise<Uint8Array>;
}

export interface Host {
  readonly kind: HostKind;
  readonly settings: SettingsStore;
  readonly workbook: WorkbookSource;
  openExternal(url: string): void;
}

/* -------------------------------------------------------------------------- */
/* Office host                                                                 */
/* -------------------------------------------------------------------------- */

const officeHost: Host = {
  kind: "office",

  settings: {
    get: (key) => {
      const v = Office.context.document.settings.get(key) as unknown;
      return typeof v === "string" ? v : null;
    },
    set: (key, value) => Office.context.document.settings.set(key, value),
    remove: (key) => Office.context.document.settings.remove(key),
    // Fire-and-forget on purpose: every caller is a "user changed something"
    // handler, and a failed persist must not block the edit itself. The
    // in-memory state stays correct either way.
    save: () => Office.context.document.settings.saveAsync(),
  },

  workbook: {
    canReadOpenWorkbook: true,
    listTables,
    readTable,
    getWorkbookBytes,
  },

  openExternal: (url) => {
    // Prefer Office's own opener: some hosts ignore target=_blank entirely.
    if (Office.context.ui && "openBrowserWindow" in Office.context.ui) {
      (Office.context.ui as unknown as { openBrowserWindow: (u: string) => void }).openBrowserWindow(url);
      return;
    }
    window.open(url, "_blank", "noopener");
  },
};

/* -------------------------------------------------------------------------- */
/* Browser host                                                                */
/* -------------------------------------------------------------------------- */

const LS_PREFIX = "dvload:settings:";

function unavailable(what: string): never {
  throw new Error(
    `${what} needs the workbook open in Excel. In the browser UI, use "Use a file…" to pick an .xlsx, .csv or .tsv instead.`
  );
}

const browserHost: Host = {
  kind: "browser",

  settings: {
    // Wrapped because localStorage throws outright in private modes and when
    // the quota is exceeded, and losing a saved mapping should not take the
    // whole pane down with it.
    get: (key) => {
      try {
        return localStorage.getItem(LS_PREFIX + key);
      } catch {
        return null;
      }
    },
    set: (key, value) => {
      try {
        localStorage.setItem(LS_PREFIX + key, value);
      } catch {
        /* ignore */
      }
    },
    remove: (key) => {
      try {
        localStorage.removeItem(LS_PREFIX + key);
      } catch {
        /* ignore */
      }
    },
    save: () => {
      /* localStorage writes are already durable */
    },
  },

  workbook: {
    canReadOpenWorkbook: false,
    listTables: async () => [],
    readTable: async () => unavailable("Reading a workbook table"),
    getWorkbookBytes: async () => unavailable("Extracting Power Query from the workbook"),
  },

  openExternal: (url) => {
    window.open(url, "_blank", "noopener");
  },
};

/* -------------------------------------------------------------------------- */
/* Detection                                                                   */
/* -------------------------------------------------------------------------- */

let current: Host | null = null;

/**
 * Office.js loads from a CDN and, in a plain browser tab, resolves onReady
 * with a null `host`. So presence of the global proves nothing — the host
 * field is the actual signal.
 *
 * The timeout covers the third case: office.js failing to load at all
 * (offline, or a network that blocks the CDN). Without it, onReady never
 * settles and the UI hangs on a blank pane with no explanation, which is a
 * much worse failure than falling back to browser mode.
 */
async function detect(): Promise<Host> {
  const officeGlobal = (globalThis as { Office?: { onReady?: unknown } }).Office;
  if (!officeGlobal || typeof officeGlobal.onReady !== "function") return browserHost;

  const info = await Promise.race([
    Office.onReady(),
    new Promise<null>((resolve) => setTimeout(() => resolve(null), 3000)),
  ]);

  return info && info.host ? officeHost : browserHost;
}

/**
 * Resolve the host and wait until the DOM is usable. Call once at startup;
 * `host()` is the accessor everywhere else.
 */
export async function initHost(): Promise<Host> {
  current = await detect();
  if (document.readyState === "loading") {
    await new Promise<void>((resolve) =>
      document.addEventListener("DOMContentLoaded", () => resolve(), { once: true })
    );
  }
  return current;
}

export function host(): Host {
  if (!current) throw new Error("initHost() must be awaited before host() is used.");
  return current;
}

export type { TableInfo };
