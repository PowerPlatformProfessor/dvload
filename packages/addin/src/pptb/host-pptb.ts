// The Power Platform ToolBox host: how the shared UI's host seam maps onto a
// PPTB tool iframe. See host.ts for what a Host owes the pane.
//
// Two impedance mismatches worth noting:
//
//  - Settings. Host.settings is synchronous (the pane reads it during
//    render), PPTB's settings API is async IPC. Bridged the same way Office
//    is: all settings load once at init into an in-memory map, reads are
//    served from it, and every write is mirrored to the ToolBox
//    fire-and-forget — a failed persist must not block the edit itself.
//
//  - Saving files. A sandboxed iframe can't trigger downloads (<a download>
//    is inert without allow-downloads), so saveFile goes through the
//    ToolBox's native save dialog instead.

import type { SourceRow } from "@dvload/core";
import type { Host } from "../host.js";
import type { PptbToolboxApi, PptbConnection } from "./pptb-bridge.js";

/** Where the pane's settings live inside the tool's PPTB settings object. */
const SETTINGS_PREFIX = "dvload:";

function unavailable(what: string): never {
  throw new Error(
    `${what} needs the workbook open in Excel. In the ToolBox, use "Add files…" to pick an .xlsx, .csv or .tsv instead.`
  );
}

export interface PptbHost extends Host {
  readonly kind: "pptb";
  /** The connection the ToolBox handed us at init. */
  readonly connection: PptbConnection;
}

export async function createPptbHost(toolbox: PptbToolboxApi): Promise<PptbHost> {
  const connection = await toolbox.connections.getActiveConnection();
  if (!connection?.url) {
    throw new Error(
      "No active Dataverse connection. Pick a connection in the Power Platform ToolBox, then reopen this tool."
    );
  }

  // Load once, serve synchronously, write back fire-and-forget.
  const cache = new Map<string, string>();
  try {
    const all = await toolbox.settings.getAll();
    for (const [k, v] of Object.entries(all)) {
      if (k.startsWith(SETTINGS_PREFIX) && typeof v === "string") {
        cache.set(k.slice(SETTINGS_PREFIX.length), v);
      }
    }
  } catch {
    // First run, or settings unavailable — start empty rather than failing
    // the whole pane over remembered state.
  }

  return {
    kind: "pptb",
    connection,

    settings: {
      get: (key) => cache.get(key) ?? null,
      set: (key, value) => {
        cache.set(key, value);
        void toolbox.settings.set(SETTINGS_PREFIX + key, value).catch(() => {});
      },
      remove: (key) => {
        cache.delete(key);
        void toolbox.settings.set(SETTINGS_PREFIX + key, null).catch(() => {});
      },
      save: () => {
        // Each set() already persisted; nothing to flush.
      },
    },

    workbook: {
      canReadOpenWorkbook: false,
      listTables: async () => [],
      readTable: async (): Promise<{ headers: string[]; rows: SourceRow[] }> =>
        unavailable("Reading a workbook table"),
      getWorkbookBytes: async () => unavailable("Extracting Power Query from the workbook"),
    },

    openExternal: (url) => {
      // The connection's browser profile, so links into the environment are
      // already signed in. Falls back to the system browser inside PPTB.
      void toolbox.utils.openInConnectionBrowser(url).catch(() => {});
    },

    saveFile: async (content, filename, _mime) => {
      const bytes = typeof content === "string" ? content : toUint8(content);
      const saved = await toolbox.fileSystem.saveFile(filename, bytes);
      return saved !== null;
    },

    copyText: (text) => toolbox.utils.copyToClipboard(text),
  };
}

function toUint8(content: Uint8Array | ArrayBuffer): Uint8Array {
  return content instanceof Uint8Array ? content : new Uint8Array(content);
}
