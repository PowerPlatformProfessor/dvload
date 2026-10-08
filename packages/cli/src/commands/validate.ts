import { readFile } from "node:fs/promises";
import path from "node:path";
import kleur from "kleur";
import { parseMapping, validateMapping, mappingWarnings, DataverseClient } from "@dvload/core";
import { getTokenProvider } from "../auth.js";

interface ValidateOpts {
  remote: boolean;
  user?: boolean;
}

export interface CheckMappingOpts extends ValidateOpts {
  /** Never escalate to an interactive login for the metadata probe. */
  silentOnly?: boolean;
}

export interface MappingCheck {
  schemaErrors: string[];
  warnings: string[];
  /**
   * Null when the Dataverse-side check didn't run (remote=false, or schema
   * errors stopped it first).
   */
  remote: { entity: string; missingAttributes: string[] } | null;
}

/**
 * The validation itself, with no output — shared by `validate` and the MCP
 * server. Throws when the file is unreadable or fails parseMapping.
 */
export async function checkMapping(mappingPath: string, opts: CheckMappingOpts): Promise<MappingCheck> {
  const raw = await readFile(path.resolve(mappingPath), "utf8");
  const mapping = parseMapping(JSON.parse(raw));

  const { environmentUrl } = mapping;
  const schemaErrors = validateMapping(mapping);
  if (schemaErrors.length > 0) return { schemaErrors, warnings: [], remote: null };

  const warnings = mappingWarnings(mapping);
  if (!opts.remote) return { schemaErrors, warnings, remote: null };

  const getToken = await getTokenProvider({
    environmentUrl,
    forceUser: opts.user,
    silentOnly: opts.silentOnly,
  });
  const client = new DataverseClient({ environmentUrl, getToken });

  // Pull the entity definition; verify each target attribute exists.
  const entitySetSingular = mapping.targetEntitySet.replace(/s$/, ""); // best-effort
  const def = (await client.getEntityDefinition(entitySetSingular)) as {
    Attributes?: Array<{ LogicalName: string }>;
  };
  const known = new Set((def.Attributes ?? []).map((a) => a.LogicalName));

  const missingAttributes = mapping.columns
    .filter((c) => c.kind !== "lookup")
    .map((c) => c.target)
    .filter((t) => !known.has(t));

  return { schemaErrors, warnings, remote: { entity: entitySetSingular, missingAttributes } };
}

export async function validateCommand(mappingPath: string, opts: ValidateOpts): Promise<void> {
  // Printed as each phase completes, so "Schema OK." still shows when the
  // remote probe then fails on auth.
  const local = await checkMapping(mappingPath, { remote: false });

  if (local.schemaErrors.length > 0) {
    console.log(kleur.red(`Schema errors (${local.schemaErrors.length}):`));
    for (const e of local.schemaErrors) console.log("  - " + e);
    process.exitCode = 1;
    return;
  }
  console.log(kleur.green("Schema OK."));

  if (local.warnings.length > 0) {
    console.log(kleur.yellow(`Warnings (${local.warnings.length}):`));
    for (const w of local.warnings) console.log(kleur.yellow("  ! " + w));
  }

  if (!opts.remote) return;

  const { remote } = await checkMapping(mappingPath, opts);
  if (remote && remote.missingAttributes.length > 0) {
    console.log(
      kleur.red(`Unknown attributes on ${remote.entity}: ${remote.missingAttributes.join(", ")}`)
    );
    process.exitCode = 1;
    return;
  }
  console.log(kleur.green("Remote attribute check OK."));
}
