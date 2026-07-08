import { readFile } from "node:fs/promises";
import path from "node:path";
import kleur from "kleur";
import { parseMapping, validateMapping, DataverseClient } from "@dvload/core";
import { getTokenProvider } from "../auth.js";

interface ValidateOpts {
  remote: boolean;
  user?: boolean;
}

export async function validateCommand(mappingPath: string, opts: ValidateOpts): Promise<void> {
  const raw = await readFile(path.resolve(mappingPath), "utf8");
  const mapping = parseMapping(JSON.parse(raw));

  const localErrors = validateMapping(mapping);
  if (localErrors.length > 0) {
    console.log(kleur.red(`Schema errors (${localErrors.length}):`));
    for (const e of localErrors) console.log("  - " + e);
    process.exitCode = 1;
    return;
  }
  console.log(kleur.green("Schema OK."));

  if (!opts.remote) return;

  const getToken = await getTokenProvider({
    environmentUrl: mapping.environmentUrl,
    forceUser: opts.user,
  });
  const client = new DataverseClient({ environmentUrl: mapping.environmentUrl, getToken });

  // Pull the entity definition; verify each target attribute exists.
  const entitySetSingular = mapping.targetEntitySet.replace(/s$/, ""); // best-effort
  const def = (await client.getEntityDefinition(entitySetSingular)) as {
    Attributes?: Array<{ LogicalName: string }>;
  };
  const known = new Set((def.Attributes ?? []).map((a) => a.LogicalName));

  const missing = mapping.columns
    .filter((c) => c.kind !== "lookup")
    .map((c) => c.target)
    .filter((t) => !known.has(t));

  if (missing.length > 0) {
    console.log(kleur.red(`Unknown attributes on ${entitySetSingular}: ${missing.join(", ")}`));
    process.exitCode = 1;
    return;
  }
  console.log(kleur.green("Remote attribute check OK."));
}
