// Minimal typings for the Power Platform ToolBox (PPTB) bridge objects a
// tool page receives: window.toolboxAPI and window.dataverseAPI.
//
// Deliberately hand-written rather than depending on @pptb/types:
//  - @pptb/types declares Node's Buffer (fileSystem.readBinary), and this
//    package's browser tsconfig has no Node types by design.
//  - Only the slice dvload actually calls is declared, so a bridge change
//    that matters fails the build instead of hiding in an `any`.
//
// Shapes verified against @pptb/types 1.2.4.

/** One Dataverse connection as PPTB models it. `url` is the environment URL. */
export interface PptbConnection {
  id: string;
  name: string;
  url: string;
  environment: string;
}

export interface PptbEventPayload {
  event: string;
  data: unknown;
  timestamp: string;
}

export interface PptbFileFilter {
  name: string;
  extensions: string[];
}

/** The slice of window.toolboxAPI dvload uses. */
export interface PptbToolboxApi {
  connections: {
    getActiveConnection(): Promise<PptbConnection | null>;
  };
  utils: {
    showNotification(options: {
      title: string;
      body: string;
      type?: "info" | "success" | "warning" | "error";
      duration?: number;
    }): Promise<void>;
    copyToClipboard(text: string): Promise<void>;
    openInConnectionBrowser(url: string): Promise<void>;
  };
  fileSystem: {
    saveFile(
      defaultPath: string,
      content: unknown,
      filters?: PptbFileFilter[]
    ): Promise<string | null>;
  };
  settings: {
    getAll(): Promise<Record<string, unknown>>;
    get(key: string): Promise<unknown>;
    set(key: string, value: unknown): Promise<void>;
  };
  events: {
    on(callback: (event: unknown, payload: PptbEventPayload) => void): void;
  };
}

/** The slice of window.dataverseAPI dvload's gateway adapter uses. */
export interface PptbDataverseApi {
  create(entityLogicalName: string, record: Record<string, unknown>): Promise<{ id: string }>;
  update(entityLogicalName: string, id: string, record: Record<string, unknown>): Promise<void>;
  delete(entityLogicalName: string, id: string): Promise<void>;
  /** Full OData query string, e.g. "accounts?$select=name&$filter=…". */
  queryData(odataQuery: string): Promise<Record<string, unknown> & { value: Array<Record<string, unknown>> }>;
  getEntityMetadata(
    entityLogicalName: string,
    searchByLogicalName: boolean,
    entityProperties?: string[]
  ): Promise<Record<string, unknown>>;
  getAllEntitiesMetadata(entityProperties?: string[]): Promise<{ value: Array<Record<string, unknown>> }>;
  /** relatedPath like "Attributes", "Keys", "Attributes(LogicalName='x')/OptionSet". */
  getEntityRelatedMetadata(
    entityLogicalName: string,
    relatedPath: string,
    relatedProperties?: string[]
  ): Promise<Record<string, unknown>>;
  createEntityDefinition(entityDefinition: Record<string, unknown>): Promise<{ id: string }>;
  createAttribute(
    entityLogicalName: string,
    attributeDefinition: Record<string, unknown>
  ): Promise<{ id: string }>;
}

/** The PPTB globals, when this page is running inside the ToolBox. */
export function pptbGlobals(): { toolbox: PptbToolboxApi; dataverse: PptbDataverseApi } | null {
  const g = globalThis as {
    toolboxAPI?: PptbToolboxApi;
    dataverseAPI?: PptbDataverseApi;
  };
  if (!g.toolboxAPI || typeof g.toolboxAPI.connections?.getActiveConnection !== "function") {
    return null;
  }
  if (!g.dataverseAPI || typeof g.dataverseAPI.queryData !== "function") return null;
  return { toolbox: g.toolboxAPI, dataverse: g.dataverseAPI };
}
