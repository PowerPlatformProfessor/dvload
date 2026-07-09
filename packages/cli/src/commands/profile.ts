import kleur from "kleur";
import { readProfiles, writeProfiles } from "../profiles.js";

export async function profileAddCommand(name: string, url: string): Promise<void> {
  const profiles = await readProfiles();
  const normalised = url.replace(/\/$/, "");
  profiles[name] = normalised;
  await writeProfiles(profiles);
  console.log(kleur.green(`Profile "${name}" saved → ${normalised}`));
}

export async function profileRemoveCommand(name: string): Promise<void> {
  const profiles = await readProfiles();
  if (!(name in profiles)) {
    console.error(kleur.red(`Profile "${name}" not found.`));
    process.exitCode = 1;
    return;
  }
  delete profiles[name];
  await writeProfiles(profiles);
  console.log(kleur.green(`Profile "${name}" removed.`));
}

export async function profileListCommand(): Promise<void> {
  const profiles = await readProfiles();
  const entries = Object.entries(profiles);
  if (entries.length === 0) {
    console.log(kleur.gray('No profiles saved. Run `dvload profile add <name> <url>` to create one.'));
    return;
  }
  const maxLen = Math.max(...entries.map(([n]) => n.length));
  for (const [name, url] of entries) {
    console.log(`  ${kleur.cyan(name.padEnd(maxLen))}  ${url}`);
  }
}
