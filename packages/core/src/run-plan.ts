// Shared run-plan schema for orchestrating multiple .dvmap runs.
// Used by both CLI and add-in. Keeps the single-table Mapping contract intact.

import type { Mapping } from "./mapping.js";

export const RUN_PLAN_SCHEMA_VERSION = 1 as const;

export interface AlternateKeyLink {
  /** Upstream step id that writes the parent table first. */
  fromStep: string;
  /** Lookup target logical name in this step's mapping (column target). */
  lookupTarget: string;
  /** Alternate-key attribute used to resolve the lookup. */
  keyAttribute: string;
}

export interface RunPlanStepOverrides {
  maxErrors?: number;
  concurrency?: number;
  notifyUrl?: string;
  dryRun?: boolean;
  user?: boolean;
  failedRows?: boolean;
}

export interface RunPlanStep {
  id: string;
  mapping: string;
  workbook: string;
  refresh?: boolean;
  stage?: number;
  dependsOn?: string[];
  overrides?: RunPlanStepOverrides;
  alternateKeyLinks?: AlternateKeyLink[];
}

export interface RunPlan {
  schemaVersion: typeof RUN_PLAN_SCHEMA_VERSION;
  name: string;
  description?: string;
  createdAt?: string;
  updatedAt?: string;
  stopOnError: boolean;
  steps: RunPlanStep[];
}

interface LegacyManifestRun {
  mapping: string;
  workbook: string;
  refresh?: boolean;
}

interface LegacyManifest {
  stopOnError?: boolean;
  runs: LegacyManifestRun[];
}

export class RunPlanParseError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "RunPlanParseError";
  }
}

function isPlainObject(v: unknown): v is Record<string, unknown> {
  return typeof v === "object" && v !== null && !Array.isArray(v);
}

function requireString(v: unknown, path: string): string {
  if (typeof v !== "string" || v.trim().length === 0) {
    throw new RunPlanParseError(`${path} must be a non-empty string`);
  }
  return v.trim();
}

function optionalString(v: unknown, path: string): string | undefined {
  if (v === undefined || v === null) return undefined;
  if (typeof v !== "string") throw new RunPlanParseError(`${path} must be a string if present`);
  return v;
}

function optionalBoolean(v: unknown, def: boolean, path: string): boolean {
  if (v === undefined || v === null) return def;
  if (typeof v !== "boolean") throw new RunPlanParseError(`${path} must be a boolean`);
  return v;
}

function optionalInteger(v: unknown, path: string, min: number, max: number): number | undefined {
  if (v === undefined || v === null) return undefined;
  if (typeof v !== "number" || !Number.isInteger(v)) {
    throw new RunPlanParseError(`${path} must be an integer`);
  }
  if (v < min || v > max) throw new RunPlanParseError(`${path} must be in [${min}, ${max}]`);
  return v;
}

function optionalStringArray(v: unknown, path: string): string[] | undefined {
  if (v === undefined || v === null) return undefined;
  if (!Array.isArray(v) || v.some((x) => typeof x !== "string" || x.trim().length === 0)) {
    throw new RunPlanParseError(`${path} must be an array of non-empty strings`);
  }
  return v.map((x) => x.trim());
}

function parseOverrides(v: unknown, path: string): RunPlanStepOverrides | undefined {
  if (v === undefined || v === null) return undefined;
  if (!isPlainObject(v)) throw new RunPlanParseError(`${path} must be an object`);
  const notifyUrl = optionalString(v.notifyUrl, `${path}.notifyUrl`);
  if (notifyUrl !== undefined && !/^(https:\/\/|http:\/\/localhost(:\d+)?\/)/i.test(notifyUrl)) {
    throw new RunPlanParseError(
      `${path}.notifyUrl must be an https:// URL (or http://localhost for testing)`
    );
  }
  return {
    maxErrors: optionalInteger(v.maxErrors, `${path}.maxErrors`, 0, Number.MAX_SAFE_INTEGER),
    concurrency: optionalInteger(v.concurrency, `${path}.concurrency`, 1, 8),
    notifyUrl,
    dryRun: v.dryRun === undefined ? undefined : optionalBoolean(v.dryRun, false, `${path}.dryRun`),
    user: v.user === undefined ? undefined : optionalBoolean(v.user, false, `${path}.user`),
    failedRows:
      v.failedRows === undefined ? undefined : optionalBoolean(v.failedRows, false, `${path}.failedRows`),
  };
}

function parseAlternateKeyLinks(v: unknown, path: string): AlternateKeyLink[] | undefined {
  if (v === undefined || v === null) return undefined;
  if (!Array.isArray(v)) throw new RunPlanParseError(`${path} must be an array`);
  return v.map((item, i) => {
    if (!isPlainObject(item)) throw new RunPlanParseError(`${path}[${i}] must be an object`);
    return {
      fromStep: requireString(item.fromStep, `${path}[${i}].fromStep`),
      lookupTarget: requireString(item.lookupTarget, `${path}[${i}].lookupTarget`),
      keyAttribute: requireString(item.keyAttribute, `${path}[${i}].keyAttribute`),
    };
  });
}

function parseStep(input: unknown, path: string): RunPlanStep {
  if (!isPlainObject(input)) throw new RunPlanParseError(`${path} must be an object`);
  return {
    id: requireString(input.id, `${path}.id`),
    mapping: requireString(input.mapping, `${path}.mapping`),
    workbook: requireString(input.workbook, `${path}.workbook`),
    refresh: input.refresh === undefined ? undefined : optionalBoolean(input.refresh, false, `${path}.refresh`),
    stage: optionalInteger(input.stage, `${path}.stage`, 1, Number.MAX_SAFE_INTEGER),
    dependsOn: optionalStringArray(input.dependsOn, `${path}.dependsOn`),
    overrides: parseOverrides(input.overrides, `${path}.overrides`),
    alternateKeyLinks: parseAlternateKeyLinks(input.alternateKeyLinks, `${path}.alternateKeyLinks`),
  };
}

function parseLegacyManifest(input: LegacyManifest, sourceName: string): RunPlan {
  if (!Array.isArray(input.runs) || input.runs.length === 0) {
    throw new RunPlanParseError("Manifest needs a non-empty runs[] array.");
  }
  const steps = input.runs.map((r, i): RunPlanStep => {
    if (!isPlainObject(r)) throw new RunPlanParseError(`runs[${i}] must be an object`);
    return {
      id: `run-${i + 1}`,
      mapping: requireString(r.mapping, `runs[${i}].mapping`),
      workbook: requireString(r.workbook, `runs[${i}].workbook`),
      refresh: r.refresh === true,
      stage: i + 1,
    };
  });
  return {
    schemaVersion: RUN_PLAN_SCHEMA_VERSION,
    name: sourceName,
    stopOnError: input.stopOnError !== false,
    steps,
  };
}

export function parseRunPlan(input: unknown, sourceName = "Run plan"): RunPlan {
  if (!isPlainObject(input)) throw new RunPlanParseError("Run plan must be a JSON object.");

  // Backward-compatible path for the existing run-all manifest shape.
  if (Array.isArray((input as { runs?: unknown }).runs) && !Array.isArray((input as { steps?: unknown }).steps)) {
    return parseLegacyManifest(input as unknown as LegacyManifest, sourceName);
  }

  if (input.schemaVersion !== RUN_PLAN_SCHEMA_VERSION) {
    throw new RunPlanParseError(
      `schemaVersion must be ${RUN_PLAN_SCHEMA_VERSION}, got ${JSON.stringify(input.schemaVersion)}`
    );
  }
  if (!Array.isArray(input.steps) || input.steps.length === 0) {
    throw new RunPlanParseError("steps must be a non-empty array");
  }

  const plan: RunPlan = {
    schemaVersion: RUN_PLAN_SCHEMA_VERSION,
    name: requireString(input.name, "name"),
    description: optionalString(input.description, "description"),
    createdAt: optionalString(input.createdAt, "createdAt"),
    updatedAt: optionalString(input.updatedAt, "updatedAt"),
    stopOnError: optionalBoolean(input.stopOnError, true, "stopOnError"),
    steps: input.steps.map((s, i) => parseStep(s, `steps[${i}]`)),
  };

  const errors = validateRunPlan(plan);
  if (errors.length > 0) {
    throw new RunPlanParseError(errors[0]);
  }
  return plan;
}

export function validateRunPlan(plan: RunPlan): string[] {
  const errors: string[] = [];
  if (!Array.isArray(plan.steps) || plan.steps.length === 0) {
    errors.push("steps must be a non-empty array");
    return errors;
  }

  const ids = new Set<string>();
  for (const step of plan.steps) {
    if (ids.has(step.id)) errors.push(`Duplicate step id: ${step.id}`);
    ids.add(step.id);
  }

  // parseRunPlan refuses blank paths on the way in, but a plan assembled in
  // an editor can still hold a half-filled step. Catching it here means the
  // editor flags it, the pane refuses to run it, and it never reaches a
  // "couldn't find file ''" further down.
  for (const step of plan.steps) {
    if (!step.mapping || step.mapping.trim() === "") {
      errors.push(`Step ${step.id} needs a mapping file`);
    }
    if (!step.workbook || step.workbook.trim() === "") {
      errors.push(`Step ${step.id} needs a workbook file`);
    }
  }

  for (const step of plan.steps) {
    for (const dep of step.dependsOn ?? []) {
      if (!ids.has(dep)) errors.push(`Step ${step.id} dependsOn unknown step "${dep}"`);
      if (dep === step.id) errors.push(`Step ${step.id} cannot depend on itself`);
    }
    for (const link of step.alternateKeyLinks ?? []) {
      if (!ids.has(link.fromStep)) {
        errors.push(`Step ${step.id} alternateKeyLinks references unknown step "${link.fromStep}"`);
      }
      if (link.fromStep === step.id) {
        errors.push(`Step ${step.id} alternateKeyLinks cannot reference itself`);
      }
    }
  }

  // Cycle detection on explicit dependencies.
  const byId = new Map(plan.steps.map((s) => [s.id, s]));
  const visiting = new Set<string>();
  const visited = new Set<string>();
  const dfs = (id: string): void => {
    if (visiting.has(id)) {
      errors.push(`Dependency cycle detected at step "${id}"`);
      return;
    }
    if (visited.has(id)) return;
    visiting.add(id);
    const step = byId.get(id);
    if (step) {
      for (const dep of step.dependsOn ?? []) {
        if (byId.has(dep)) dfs(dep);
      }
    }
    visiting.delete(id);
    visited.add(id);
  };
  for (const step of plan.steps) dfs(step.id);

  return errors;
}

export function serializeRunPlan(plan: RunPlan): string {
  return JSON.stringify({ ...plan, updatedAt: new Date().toISOString() }, null, 2) + "\n";
}

/**
 * Cross-mapping checks for a plan whose step mappings have been loaded.
 *
 * Callers own the loading (the CLI reads files, the pane resolves in-memory
 * mappings by name) and pass the results keyed by step id; a step absent from
 * the map is skipped here because the caller has already reported why it
 * couldn't be loaded.
 */
export function crossValidatePlanMappings(
  plan: RunPlan,
  mappings: ReadonlyMap<string, Mapping>
): string[] {
  const errors: string[] = [];
  for (const step of plan.steps) {
    for (const link of step.alternateKeyLinks ?? []) {
      const from = mappings.get(link.fromStep);
      const current = mappings.get(step.id);
      if (!from || !current) continue;

      if (!from.upsertKey?.includes(link.keyAttribute)) {
        errors.push(
          `step "${step.id}" alternateKeyLinks: upstream step "${link.fromStep}" must include ` +
            `"${link.keyAttribute}" in upsertKey`
        );
      }

      const col = current.columns.find((c) => c.target === link.lookupTarget);
      if (!col) {
        errors.push(
          `step "${step.id}" alternateKeyLinks: lookup target "${link.lookupTarget}" is not mapped`
        );
        continue;
      }
      if (col.kind !== "lookup") {
        errors.push(`step "${step.id}" alternateKeyLinks: "${link.lookupTarget}" must be a lookup mapping`);
        continue;
      }
      if (col.lookupResolution !== "alternateKey") {
        errors.push(
          `step "${step.id}" alternateKeyLinks: "${link.lookupTarget}" must set lookupResolution=alternateKey`
        );
      }
      if (col.keyAttribute !== link.keyAttribute) {
        errors.push(
          `step "${step.id}" alternateKeyLinks: "${link.lookupTarget}" keyAttribute must be ` +
            `"${link.keyAttribute}"`
        );
      }
    }
  }
  return errors;
}

/**
 * Order a plan's steps into sequential batches; steps inside one batch have
 * no ordering constraints between them, so an executor may run a batch in
 * parallel (the CLI does) or sequentially (the task pane does — the ToolBox
 * bridge serializes requests anyway).
 *
 * Constraints considered: explicit `dependsOn`, implied dependencies from
 * `alternateKeyLinks` (the parent rows must exist before a lookup can resolve
 * against them), and `stage` numbers — when any step declares a stage, every
 * lower-staged step precedes every higher-staged one.
 */
export function buildExecutionBatches(plan: RunPlan): RunPlanStep[][] {
  const byId = new Map(plan.steps.map((s) => [s.id, s]));
  const originalOrder = new Map(plan.steps.map((s, i) => [s.id, i]));
  const stagesPresent = plan.steps.some((s) => typeof s.stage === "number");
  const stepStage = (s: RunPlanStep): number => (stagesPresent ? (s.stage ?? 1) : 1);

  const deps = new Map<string, Set<string>>();
  for (const step of plan.steps) {
    const d = new Set<string>(step.dependsOn ?? []);
    for (const link of step.alternateKeyLinks ?? []) d.add(link.fromStep);
    deps.set(step.id, d);
  }

  if (stagesPresent) {
    for (const step of plan.steps) {
      const curStage = stepStage(step);
      for (const maybeDep of plan.steps) {
        if (step.id === maybeDep.id) continue;
        if (stepStage(maybeDep) < curStage) deps.get(step.id)?.add(maybeDep.id);
      }
    }
  }

  const pending = new Set(plan.steps.map((s) => s.id));
  const completed = new Set<string>();
  const batches: RunPlanStep[][] = [];

  while (pending.size > 0) {
    const ready = [...pending]
      .filter((id) => {
        const d = deps.get(id) ?? new Set<string>();
        for (const dep of d) {
          if (pending.has(dep) && !completed.has(dep)) return false;
        }
        return true;
      })
      .sort((a, b) => {
        const sa = stepStage(byId.get(a)!);
        const sb = stepStage(byId.get(b)!);
        if (sa !== sb) return sa - sb;
        return (originalOrder.get(a) ?? 0) - (originalOrder.get(b) ?? 0);
      });

    if (ready.length === 0) {
      throw new Error("Run plan has unsatisfied dependencies.");
    }

    const minStage = stepStage(byId.get(ready[0])!);
    const batchIds = ready.filter((id) => stepStage(byId.get(id)!) === minStage);
    const batch = batchIds.map((id) => byId.get(id)!);
    batches.push(batch);

    for (const id of batchIds) {
      pending.delete(id);
      completed.add(id);
    }
  }

  return batches;
}
