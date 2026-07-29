# dvload recipes

Worked examples for the things people actually need to do. Each recipe has
the task pane steps first, then the equivalent `.dvmap.json` and CLI command,
so you can stop at whichever level suits you.

Every mapping shown is complete and valid — copy one, change the
`environmentUrl`, `sourceTable` and column names, and run it.

- [Before you start](#before-you-start)
- [1. Load a simple table](#1-load-a-simple-table)
- [2. Update existing records instead of duplicating them](#2-update-existing-records-instead-of-duplicating-them)
- [3. Link to another record: lookups](#3-link-to-another-record-lookups)
  - [3a. You already have the GUID](#3a-you-already-have-the-guid)
  - [3b. Match on an alternate key](#3b-match-on-an-alternate-key)
  - [3c. Match on a name](#3c-match-on-a-name)
  - [3d. Customer, Owner, Regarding: polymorphic lookups](#3d-customer-owner-regarding-polymorphic-lookups)
- [4. Choice columns with labels instead of numbers](#4-choice-columns-with-labels-instead-of-numbers)
- [5. Set a field to the same value on every row](#5-set-a-field-to-the-same-value-on-every-row)
- [6. Load parents and children in one go](#6-load-parents-and-children-in-one-go)
- [7. Mirror a source system (delete or deactivate what's gone)](#7-mirror-a-source-system-delete-or-deactivate-whats-gone)
- [8. Load from CSV, or from several files](#8-load-from-csv-or-from-several-files)
- [9. Run it every night](#9-run-it-every-night)
- [10. Make a big load faster](#10-make-a-big-load-faster)
- [11. When a run goes wrong](#11-when-a-run-goes-wrong)
- [Glossary](#glossary)

---

## Before you start

Two front ends, one engine. Anything you build in the UI can be saved and run
by the CLI, unchanged.

```bash
dvload serve      # then open the pane in Excel
dvload gui        # or the same UI in a browser, no Excel needed
dvload run mapping.dvmap.json -w data.xlsx
```

**Always dry-run first.** It does everything except write:

```bash
dvload run contacts.dvmap.json -w contacts.xlsx --dry-run
```

In the pane, tick **Dry run** before clicking Run import. Coercion errors,
unresolved lookups and bad column names all surface here, at no cost.

---

## 1. Load a simple table

Three columns from a spreadsheet into new contact records.

**In the UI**

1. Pick the source table (or **Add files…** and pick one from a workbook).
2. Pick the target entity — *Contact*.
3. Click **Suggest mappings** to match columns by name, then fix anything it
   got wrong.
4. Tick **Dry run**, click **Run import**, check the summary.
5. Untick **Dry run**, run for real.

**The mapping**

```json
{
  "schemaVersion": 1,
  "name": "Contacts from marketing list",
  "environmentUrl": "https://contoso.crm.dynamics.com",
  "targetEntitySet": "contacts",
  "sourceTable": "tblContacts",
  "conflictMode": "insert",
  "batchSize": 100,
  "maxErrors": 0,
  "logDir": "./logs",
  "concurrency": 1,
  "bypassCustomLogic": false,
  "skipUnchanged": false,
  "columns": [
    { "source": "First Name", "target": "firstname",      "kind": "string", "treatEmptyAsNull": true },
    { "source": "Last Name",  "target": "lastname",       "kind": "string", "treatEmptyAsNull": true },
    { "source": "Email",      "target": "emailaddress1",  "kind": "string", "treatEmptyAsNull": true }
  ]
}
```

`maxErrors: 0` means unlimited — the run continues past bad rows and reports
them at the end. Set it to `1` if you'd rather stop at the first problem.

---

## 2. Update existing records instead of duplicating them

`conflictMode: "insert"` always creates. To match existing records you need
an **alternate key** on the target table — a Dataverse-enforced uniqueness
constraint, set up under *Table → Keys*. Without one, Dataverse has no way to
know that "the contact with this email" already exists.

| Mode | What it does |
|---|---|
| `insert` | Always creates. Duplicates on re-run. |
| `upsert` | Updates when the key matches, creates when it doesn't. |
| `skip-if-exists` | Creates only when the key doesn't match. Never updates. |
| `sync` | Upsert, then deal with records that vanished from the source. See [recipe 7](#7-mirror-a-source-system-delete-or-deactivate-whats-gone). |

**In the UI:** set **On conflict** to *Upsert*, then fill in **Match on** with
the alternate key attribute.

```json
{
  "conflictMode": "upsert",
  "upsertKey": ["emailaddress1"],
  "skipUnchanged": true
}
```

> **The key attribute must also be mapped as a column.** `upsertKey:
> ["emailaddress1"]` with no column writing `emailaddress1` fails validation
> with *upsertKey attribute "emailaddress1" is not mapped in columns* — there
> would be no value to match on. The same applies to `sync`.

`skipUnchanged` reads each target record first and drops attributes that
already match, skipping the row entirely if nothing changed. It costs one GET
per row, and buys you a quiet audit history and far fewer plugin executions.
Worth it on a nightly job where most rows don't change; not worth it on a
first load, where nothing exists yet.

> Lookup columns are always sent, even with `skipUnchanged` — navigation
> property values can't be compared without extra reads.

---

## 3. Link to another record: lookups

A lookup column needs three answers:

1. **Which table does it point at?** → `bindEntitySet`, e.g. `"accounts"`
2. **How do I find the record?** → `lookupResolution`
3. **Using which column?** → `keyAttribute` (not needed for GUIDs)

In the UI these are the **Binds to → Match by → Key field** row that appears
under a lookup column, filled in that order.

### 3a. You already have the GUID

The fastest and most reliable option — no lookup queries at all.

```json
{
  "source": "Account ID",
  "target": "parentcustomerid_account",
  "kind": "lookup",
  "bindEntitySet": "accounts",
  "lookupResolution": "guid",
  "treatEmptyAsNull": true
}
```

The cell must contain a real GUID. Anything else fails the row with
`not a valid GUID` — deliberately, since the value goes straight into a URL.

### 3b. Match on an alternate key

Best choice when your source has a stable business identifier and the target
table has a matching alternate key defined.

```json
{
  "source": "Account Number",
  "target": "parentcustomerid_account",
  "kind": "lookup",
  "bindEntitySet": "accounts",
  "lookupResolution": "alternateKey",
  "keyAttribute": "accountnumber",
  "treatEmptyAsNull": true
}
```

### 3c. Match on a name

When there's no key, match on any text attribute. Values are looked up in
batches, so this is not one query per row.

```json
{
  "source": "Company",
  "target": "parentcustomerid_account",
  "kind": "lookup",
  "bindEntitySet": "accounts",
  "lookupResolution": "text",
  "keyAttribute": "name",
  "duplicateBehavior": "error",
  "createIfMissing": false,
  "treatEmptyAsNull": true
}
```

Two settings decide what happens when the match isn't clean:

- `duplicateBehavior`: `"error"` (default) fails the row when two accounts
  share a name; `"first"` takes whichever comes back first. Use `"first"`
  only when you genuinely don't care which one you get.
- `createIfMissing`: creates a record in `bindEntitySet` with
  `keyAttribute` set to the source value when nothing matches.

> **`createIfMissing` creates real records.** It makes a bare record with one
> field populated — nothing else. That's reasonable for a lookup table of
> categories; it is almost never right for accounts, contacts, or owners.
> Leave it off unless you've thought about what a half-populated record means
> in your data.

### 3d. Customer, Owner, Regarding: polymorphic lookups

Some lookups can point at more than one table: *Customer* is account **or**
contact, *Owner* is user **or** team, *Regarding* can be almost anything.

These cannot be bound by their attribute name. `customerid@odata.bind` fails
the entire row with an "undeclared property" error. Dataverse wants the
**navigation property**, which encodes the target table:

| Attribute | Pointing at | Use as `target` |
|---|---|---|
| `parentcustomerid` | account | `parentcustomerid_account` |
| `parentcustomerid` | contact | `parentcustomerid_contact` |
| `customerid` | account | `customerid_account` |
| `ownerid` | user | `ownerid` |

**The UI does this for you.** Pick the lookup column, choose the entity set
under **Binds to**, and dvload reads the relationship metadata and rewrites
the target to the right navigation property. The dropdown only offers targets
that lookup actually allows.

Writing the mapping by hand, set `target` to the navigation property
yourself:

```json
{
  "source": "Account ID",
  "target": "customerid_account",
  "kind": "lookup",
  "bindEntitySet": "accounts",
  "lookupResolution": "guid",
  "treatEmptyAsNull": true
}
```

**One column, one target table.** `bindEntitySet` is fixed per column, so a
single source column holding a mix of account and contact GUIDs can't be
mapped as-is. Split it into two source columns, each row filling only one:

```json
"columns": [
  {
    "source": "Account ID", "target": "customerid_account", "kind": "lookup",
    "bindEntitySet": "accounts", "lookupResolution": "guid", "treatEmptyAsNull": false
  },
  {
    "source": "Contact ID", "target": "customerid_contact", "kind": "lookup",
    "bindEntitySet": "contacts", "lookupResolution": "guid", "treatEmptyAsNull": false
  }
]
```

An empty lookup value is omitted from the payload, so each row binds only the
column that's filled.

> Note `treatEmptyAsNull: false` here. With it on, an empty cell writes an
> explicit null — and nulling `customerid_account` while setting
> `customerid_contact` on the same underlying attribute is asking for
> trouble under `upsert` or `sync`. On plain `insert` it makes no difference.

---

## 4. Choice columns with labels instead of numbers

Dataverse stores choices as integers. Your spreadsheet almost certainly has
labels. `optionMap` translates.

**In the UI:** pick a choice column and dvload fills the map from metadata
automatically — you'll see a confirmation of how many labels it loaded.

```json
{
  "source": "Status",
  "target": "prioritycode",
  "kind": "choice",
  "optionMap": { "Low": 1, "Normal": 2, "High": 3 },
  "treatEmptyAsNull": true
}
```

Cells may contain either the label or the integer. Anything not in the map
and not a valid integer fails the row.

Use `kind: "multichoice"` for multi-select columns; the cell should hold
several labels separated by semicolons.

---

## 5. Set a field to the same value on every row

A column with `constant` instead of `source` — no spreadsheet column needed.

```json
"columns": [
  { "source": "Name", "target": "name", "kind": "string", "treatEmptyAsNull": true },
  { "constant": "Import 2026-07", "target": "description", "kind": "string", "treatEmptyAsNull": false },
  {
    "constant": "6fd0a17e-6b3f-4c21-9a44-0e51f2b39f2c", "target": "ownerid", "kind": "lookup",
    "bindEntitySet": "systemusers", "lookupResolution": "guid", "treatEmptyAsNull": false
  }
]
```

A constant on a `guid` lookup must be a complete GUID — validation rejects
anything else before the run starts.

**In the UI:** **Add fixed value**. For lookups you get a record search
against the bound table, so you pick a user by name rather than pasting a
GUID.

---

## 6. Load parents and children in one go

A run plan (`.dvplan.json`) runs several mappings in dependency order — load
accounts, then contacts that point at them.

The problem it solves: the child mapping needs the parents' GUIDs, which
don't exist until the parent step has run. `alternateKeyLinks` wires the two
together by key instead.

```bash
dvload run-all nightly.dvplan.json
```

**In the UI:** the **Run plan** section — add steps, set `stage` or
`dependsOn`, save the plan next to your mappings.

Steps sharing a `stage` run in parallel; `dependsOn` forces explicit
ordering. Full schema: [docs/schema/dvplan.schema.json](./schema/dvplan.schema.json).

---

## 7. Mirror a source system (delete or deactivate what's gone)

`conflictMode: "sync"` is upsert plus a removal pass: anything in Dataverse
whose key isn't in the source file gets dealt with.

```json
{
  "conflictMode": "sync",
  "upsertKey": ["accountnumber"],
  "syncAction": "deactivate",
  "skipUnchanged": true
}
```

`syncAction` is `"deactivate"` (default) or `"delete"`.

> Read this one twice before running it. A source file that is short, stale,
> or filtered will deactivate or delete every record it doesn't mention.
> Dry-run it, read the summary line for removals, and prefer `"deactivate"`
> until you trust the pipeline — deactivation is reversible.

---

## 8. Load from CSV, or from several files

CSV and TSV are first-class sources — no Excel required:

```bash
dvload run contacts.dvmap.json -w exported.csv
```

In the UI, **Add files…** takes several files at once, and each table inside
each workbook becomes a separate entry in the source picker, labelled with
the file it came from. Files are read on demand, so adding several large
workbooks costs their file size rather than their row count.

---

## 9. Run it every night

```bash
dvload schedule contacts.dvmap.json -w "C:\data\contacts.xlsx" --time 03:30
```

This registers a Windows Task Scheduler job that refreshes the workbook's
Power Query and then runs the import.

Set up app-only auth first, or the job will fail whenever the cached user
token expires:

```bash
dvload app-login --env https://contoso.crm.dynamics.com \
  --client-id <app-id> --tenant-id <tenant-id>
```

Full walkthrough, including the Entra app registration and the Application
User: [docs/SCHEDULED-RUNS.md](./SCHEDULED-RUNS.md).

> **Records created by a scheduled run are owned by the Application User;
> records created from the task pane are owned by you.** Same mapping, two
> identities, depending on how it ran. That's intended — an unattended job
> has no signed-in user to act as — but it surprises people looking at
> *Created By*.

---

## 10. Make a big load faster

In rough order of impact:

| Setting | Effect | Watch out for |
|---|---|---|
| `concurrency` | Parallel `$batch` requests, 1–8. | Dataverse throttles; 4 is a sane ceiling on most environments. |
| `batchSize` | Rows per changeset, max 1000. | A whole batch fails together, so bigger batches mean coarser error reporting. |
| `bypassCustomLogic` | Skips synchronous plugins and Power Automate triggers. | Needs the `prvBypassCustomPlugins` privilege, and your business logic genuinely won't run. |
| `skipUnchanged: false` | Avoids one GET per row. | Only relevant on upsert/sync. |

```json
{ "concurrency": 4, "batchSize": 500 }
```

Use GUID lookups over text lookups where you can — text resolution costs
extra queries before the load starts.

---

## 11. When a run goes wrong

**Read the first error, not the last.** One bad mapping produces thousands of
identical failures; the first line tells you which column.

**Failed rows come back as a spreadsheet.** Both front ends produce a file
with the failing source rows in their original shape — fix the cells and
re-run just that file:

```bash
dvload run contacts.dvmap.json -w logs/failed_contacts_2026-07-29_03-30-11.xlsx
```

**Common failures**

| Message | Cause |
|---|---|
| `not a valid GUID` | `lookupResolution: "guid"` on a column that holds a name. Use `"text"` or `"alternateKey"`. |
| `undeclared property` | A polymorphic lookup bound by attribute name. See [3d](#3d-customer-owner-regarding-polymorphic-lookups). |
| `lookup unresolved` | No match for the key value. Check `keyAttribute` and whether the referenced record exists. |
| `ambiguous text lookup` | Two records share the text. Set `duplicateBehavior` or use a key. |
| `conflictMode=upsert requires at least one upsertKey attribute` | Upsert with no key configured. |

**Cancel and resume.** Cancelling mid-run saves a checkpoint; the pane offers
to resume from the row it stopped at, and the CLI takes `--resume`.

**Validate without running anything:**

```bash
dvload validate contacts.dvmap.json
```

Checks the mapping against the schema *and* against live Dataverse metadata —
wrong attribute names, missing entity sets, bad lookup targets.

---

## Glossary

| Term | Meaning |
|---|---|
| **Entity set** | The plural name Dataverse uses in URLs: `accounts`, `contacts`, `new_projects`. Not the display name. |
| **Logical name** | The internal field name: `emailaddress1`, not "Email". Shown next to every field in the pane. |
| **Alternate key** | A uniqueness constraint you define on a table (*Table → Keys*). Required for upsert and sync. |
| **Navigation property** | The name used to write a lookup. Same as the attribute name, except for polymorphic lookups. |
| **Application User** | A Dataverse identity for an app registration rather than a person. Used by scheduled runs. |
| **`$batch`** | The OData mechanism dvload uses to send many rows in one request. |

---

Field-by-field reference: [README](../README.md#mapping-json-shape).
Machine-readable schema: [docs/schema/dvmap.schema.json](./schema/dvmap.schema.json) —
point your editor at it for completion and inline validation.
