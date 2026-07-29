/**
 * Property tests for the OData encoding layer.
 *
 * `formatKeyLiteral` is the single most security-sensitive function in the
 * codebase: it takes an arbitrary spreadsheet cell and puts it inside an HTTP
 * request line in a multipart batch body. Example tests can only cover the
 * injection payloads someone imagined. These properties state the rules that
 * must hold for every possible string.
 */

import { describe, it, expect } from "vitest";
import { fc } from "./setup.js";
import { formatKeyLiteral, assertLogicalName, assertGuid } from "../../src/dataverse.js";
import { parseODataLiteral } from "../support/fake-dataverse.js";

describe("formatKeyLiteral", () => {
  it("never emits a character with structural meaning in a request line", () => {
    // CR, LF, space and tab would break out of the request line or inject a
    // header. Parens, quotes and commas would break out of the key expression.
    fc.assert(
      fc.property(fc.string(), (s) => {
        const out = formatKeyLiteral(s);
        const inner = out.slice(1, -1); // strip the enclosing quotes
        expect(inner).not.toMatch(/[\r\n\t ]/);
        expect(inner).not.toMatch(/['"()!*~,]/);
      })
    );
  });

  it("is injective: distinct inputs never collide", () => {
    // A collision would let one record's key address another's — the kind of
    // bug that quietly overwrites the wrong row during an upsert.
    fc.assert(
      fc.property(fc.string(), fc.string(), (a, b) => {
        fc.pre(a !== b);
        expect(formatKeyLiteral(a)).not.toBe(formatKeyLiteral(b));
      })
    );
  });

  it("round-trips through an independent decoder", () => {
    // The decoder lives in the fake server and was written from the OData
    // spec, not from this implementation — so agreement means something.
    fc.assert(
      fc.property(fc.string(), (s) => {
        expect(parseODataLiteral(formatKeyLiteral(s))).toBe(s);
      })
    );
  });

  it("survives a URL round-trip unchanged", () => {
    fc.assert(
      fc.property(fc.string(), (s) => {
        const literal = formatKeyLiteral(s);
        const url = new URL(`https://x.crm.dynamics.com/api/data/v9.2/contacts(email=${literal})`);
        // Nothing in the literal may cause the URL parser to re-interpret the
        // path, add a query string, or start a fragment.
        expect(url.search).toBe("");
        expect(url.hash).toBe("");
        expect(url.pathname).toContain(literal);
      })
    );
  });

  it("emits numbers and booleans bare, never quoted", () => {
    fc.assert(
      fc.property(fc.integer(), (n) => {
        expect(formatKeyLiteral(n)).toBe(String(n));
      })
    );
    expect(formatKeyLiteral(true)).toBe("true");
    expect(formatKeyLiteral(false)).toBe("false");
  });

  it("refuses non-finite numbers instead of emitting NaN into a URL", () => {
    for (const n of [NaN, Infinity, -Infinity]) {
      expect(() => formatKeyLiteral(n)).toThrow();
    }
  });
});

describe("assertLogicalName", () => {
  it("accepts exactly the Dataverse identifier grammar", () => {
    fc.assert(
      fc.property(fc.stringMatching(/^[A-Za-z_][A-Za-z0-9_]{0,40}$/), (name) => {
        expect(assertLogicalName(name, "test")).toBe(name);
      })
    );
  });

  it("rejects anything containing a character that could alter a URL or header", () => {
    fc.assert(
      fc.property(
        fc.string({ minLength: 1 }).filter((s) => !/^[A-Za-z_][A-Za-z0-9_]*$/.test(s)),
        (name) => {
          expect(() => assertLogicalName(name, "test")).toThrow();
        }
      )
    );
  });
});

describe("assertGuid", () => {
  it("accepts canonical GUIDs in any casing", () => {
    fc.assert(
      fc.property(fc.uuid(), fc.boolean(), (id, upper) => {
        const v = upper ? id.toUpperCase() : id;
        expect(assertGuid(v, "test")).toBe(v);
      })
    );
  });

  it("rejects anything else, including a GUID with trailing content", () => {
    fc.assert(
      fc.property(fc.uuid(), fc.string({ minLength: 1 }), (id, junk) => {
        fc.pre(!/^[\s]*$/.test(junk));
        expect(() => assertGuid(`${id}${junk}`, "test")).toThrow();
      })
    );
  });
});
