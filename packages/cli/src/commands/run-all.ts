import { readFile } from "node:fs/promises";
import path from "node:path";
import kleur from "kleur";
import {
  mappingWarnings,
  parseMapping,
  parseRunPlan,
  validateMapping,
  type Mapping,
  type RunPlan,
  type RunPlanStep,
} from "@dvload/core";
import { executeRun, type RunOpts } from "./run.js";

export interface RunAllOpts {
  dryRun?: boolean;
  user?: boolean;
  notifyUrl?: string;
  failedRows?: boolean;
}

export async function runAllCommand(manifestPath: string, opts: RunAllOpts): Promise<void> {
  const file = path.resolve(manifestPath);
  const dir = path.dirname(file);
  const raw = JSON.parse(await readFile(file, "utf8"));
  const plan = parseRunPlan(raw, path.basename(file));
  const planValidation = await validatePlanMappings(plan, dir);
  if (planValidation.errors.length > 0) {
    console.error(kleur.red("Run plan validation failed:"));
    for (const e of planValidation.errors) console.error(kleur.red(`  - ${e}`));
    process.exitCode = 2;
    return;
  }
  for (const w of planValidation.warnings) console.log(kleur.yellow(`! ${w}`));

  const batches = buildExecutionBatches(plan);
  const stopOnError = plan.stopOnError;
  const totalSteps = plan.steps.length;
  let attempted = 0;

  let failedRuns = 0;
  for (let i = 0; i < batches.length; i++) {
    const batch = batches[i];
    console.log("");
    console.log(kleur.bold(`Stage ${i + 1}/${batches.length} (${batch.length} run${batch.length === 1 ? "" : "s"})`));

    const results = await Promise.all(
      batch.map(async (step) => {
        attempted++;
        console.log(kleur.bold(`[${attempted}/${totalSteps}] ${step.id} → ${step.mapping}`));
        const mappingPath = path.resolve(dir, step.mapping);
        const runOpts = buildRunOpts(step, opts, dir);
        try {
          const result = await executeRun(mappingPath, runOpts);
          const ok = result !== null && result.failed === 0;
          return { step, ok, error: null as string | null };
        } catch (e) {
          return { step, ok: false, error: (e as Error).message };
        }
      })
    );

    const failedThisBatch = results.filter((r) => !r.ok);
    for (const failed of failedThisBatch) {
      if (failed.error) {
        console.error(kleur.red(`${failed.step.id}: ${failed.error}`));
      }
    }
    failedRuns += failedThisBatch.length;

    if (failedThisBatch.length > 0 && stopOnError) {
      console.error(
        kleur.red(
          `Stopping: stage ${i + 1} had failure(s) and stopOnError is set. ` +
            `${totalSteps - attempted} run(s) not attempted.`
        )
      );
      break;
    }
  }

  if (failedRuns > 0) process.exitCode = 1;
  console.log("");
  console.log(
    failedRuns === 0
      ? kleur.green(`All ${totalSteps} runs completed successfully.`)
      : kleur.red(`${failedRuns} run(s) had failures.`)
  );
}

function buildRunOpts(step: RunPlanStep, opts: RunAllOpts, dir: string): RunOpts {
  const overrides = step.overrides ?? {};
  return {
    workbook: path.resolve(dir, step.workbook),
    refresh: step.refresh,
    dryRun: overrides.dryRun ?? opts.dryRun,
    user: overrides.user ?? opts.user,
    notifyUrl: overrides.notifyUrl ?? opts.notifyUrl,
    failedRows: overrides.failedRows ?? opts.failedRows,
    maxErrors: overrides.maxErrors,
    concurrency: overrides.concurrency,
  };
}

export async function validatePlanMappings(
  plan: RunPlan,
  dir: string
): Promise<{ errors: string[]; warnings: string[] }> {
  const errors: string[] = [];
  const warnings: string[] = [];
  const mappings = new Map<string, Mapping>();

  for (const step of plan.steps) {
    try {
      const mappingPath = path.resolve(dir, step.mapping);
      const raw = await readFile(mappingPath, "utf8");
      const mapping = parseMapping(JSON.parse(raw));
      const schemaErrs = validateMapping(mapping);
      for (const e of schemaErrs) errors.push(`step "${step.id}" (${step.mapping}): ${e}`);
      for (const w of mappingWarnings(mapping)) warnings.push(`step "${step.id}": ${w}`);
      mappings.set(step.id, mapping);
    } catch (e) {
      errors.push(`step "${step.id}" (${step.mapping}): ${(e as Error).message}`);
    }
  }

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
        errors.push(
          `step "${step.id}" alternateKeyLinks: "${link.lookupTarget}" must be a lookup mapping`
        );
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

  return { errors, warnings };
}

export function buildExecutionBatches(plan: RunPlan): RunPlanStep[][] {
  const byId = new Map(plan.steps.map((s) => [s.id, s]));
  const originalOrder = new Map(plan.steps.map((s, i) => [s.id, i]));
  const stagesPresent = plan.steps.some((s) => typeof s.stage === "number");
  const stepStage = (s: RunPlanStep): number => (stagesPresent ? s.stage ?? 1 : 1);

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
