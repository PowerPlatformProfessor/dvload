import { test } from "node:test";
import assert from "node:assert/strict";
import {
  inferColumnKind,
  suggestColumns,
  sanitizeSchemaSuffix,
  buildAttributePayload,
  buildEntityPayload,
  buildKeyPayload,
  attributeLogicalName,
} from "./tablegen.js";

test("inferColumnKind: booleans, ints, decimals, dates, strings, memo", () => {
  assert.equal(inferColumnKind(["TRUE", "no", "Yes"]).kind, "boolean");
  assert.equal(inferColumnKind([1, 2, "3"]).kind, "integer");
  assert.equal(inferColumnKind([1.5, "2"]).kind, "decimal");
  assert.equal(inferColumnKind([2147483648]).kind, "decimal"); // > int32
  assert.equal(inferColumnKind(["1980-01-01", "1990-12-31"]).kind, "dateonly");
  assert.equal(inferColumnKind(["1980-01-01T10:30:00"]).kind, "datetime");
  assert.equal(inferColumnKind(["hello", "world"]).kind, "string");
  assert.equal(inferColumnKind(["x".repeat(500)]).kind, "memo");
  assert.equal(inferColumnKind([null, "", undefined]).kind, "string"); // all blank
  // blanks are ignored when inferring
  assert.equal(inferColumnKind(["", 5, null, 7]).kind, "integer");
});

test("sanitizeSchemaSuffix strips punctuation and forces a letter start", () => {
  assert.equal(sanitizeSchemaSuffix("First Name"), "FirstName");
  assert.equal(sanitizeSchemaSuffix("credit-limit (USD)"), "CreditLimitUSD");
  assert.equal(sanitizeSchemaSuffix("123 weird!"), "Weird");
  assert.equal(sanitizeSchemaSuffix("!!!"), "Column");
});

test("suggestColumns marks the first string column as primary name and dedupes suffixes", () => {
  const rows = [
    { "First Name": "A", Age: 3, "first name": "B" },
    { "First Name": "C", Age: 4, "first name": "D" },
  ];
  const cols = suggestColumns(["First Name", "Age", "first name"], rows);
  assert.equal(cols[0].isPrimaryName, true);
  assert.equal(cols[1].kind, "integer");
  assert.equal(cols[1].isPrimaryName, false);
  // duplicate sanitized suffix gets uniquified
  assert.notEqual(cols[0].schemaSuffix.toLowerCase(), cols[2].schemaSuffix.toLowerCase());
});

test("payloads carry the right OData types and names", () => {
  const cols = suggestColumns(
    ["Email", "Credit Limit"],
    [{ Email: "a@x", "Credit Limit": 1.5 }]
  );
  const [email, credit] = cols;

  const entity = buildEntityPayload({
    prefix: "new",
    schemaSuffix: "Contacts",
    displayName: "Contacts",
    primaryNameColumn: email,
  });
  assert.equal(entity["SchemaName"], "new_Contacts");
  const attrs = entity["Attributes"] as Array<Record<string, unknown>>;
  assert.equal(attrs.length, 1);
  assert.equal(attrs[0]["IsPrimaryName"], true);
  assert.equal(attrs[0]["@odata.type"], "Microsoft.Dynamics.CRM.StringAttributeMetadata");

  const dec = buildAttributePayload("new", credit);
  assert.equal(dec["@odata.type"], "Microsoft.Dynamics.CRM.DecimalAttributeMetadata");
  assert.equal(dec["SchemaName"], "new_CreditLimit");
  assert.equal(attributeLogicalName("new", credit), "new_creditlimit");

  email.inAlternateKey = true;
  const key = buildKeyPayload("new", "Contacts", [email]);
  assert.deepEqual(key["KeyAttributes"], ["new_email"]);
  assert.equal(key["SchemaName"], "new_key_contacts");
});
