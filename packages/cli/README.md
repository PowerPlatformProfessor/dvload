# dvload

Load Excel tables (typically Power Query output) into Microsoft Dataverse
from the command line. Pairs with Windows Task Scheduler for daily
unattended imports. The mapping (`.dvmap.json`) is built visually in the
companion [Excel add-in](https://github.com/PowerPlatformProfessor/dvload)
and run here.

## Install

```bash
npm i -g dvload
```

Requires Node 20+. Windows 10/11 for scheduling and Power Query refresh.

## Quick start

```bash
# one-time sign-in (opens your browser; no Entra setup, no admin consent)
dvload login --env https://yourorg.crm.dynamics.com

# sanity-check a mapping against live Dataverse metadata
dvload validate ./contacts.dvmap.json

# dry run, then the real thing
dvload run ./contacts.dvmap.json -w ./customers.xlsx --dry-run
dvload run ./contacts.dvmap.json -w ./customers.xlsx --refresh

# nightly at 03:30 (see docs for unattended app-only auth)
dvload schedule ./contacts.dvmap.json -w ./customers.xlsx --time 03:30
```

## Documentation

- [Full README](https://github.com/PowerPlatformProfessor/dvload#readme) —
  auth model, all commands, mapping format, .pqt bridge
- [Unattended scheduled runs](https://github.com/PowerPlatformProfessor/dvload/blob/main/docs/SCHEDULED-RUNS.md) —
  app-only auth setup for schedules

## License

MIT
