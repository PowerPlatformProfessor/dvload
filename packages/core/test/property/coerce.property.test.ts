/**
 * Property tests for cell coercion.
 *
 * Example-based tests pin the cases someone thought of. These pin the rules
 * that must hold for *every* input — which is what you actually want from a
 * layer that turns arbitrary spreadsheet cells into a production write.
 *
 * The invariants below are chosen so that a violation is a real bug, not a
 * restatement of the implementation.
 */

import { describe, it, expect } from "vitest";
import { fc } from "./setup.js";
import { coerceValue, coerceRow, parseWithFormat, CoerceError } from "../../src/coerce.js";
import type { ColumnMapping } from "../../src/mapping.js";
import type { DataverseFieldKind } from "../../src/types.js";

const ALL_KINDS: DataverseFieldKind[] = [
  "string",
  "memo",
  "integer",
  "decimal",
  "money",
  "double",
  "boolean",
  "datetime",
  "dateonly",
  "uniqueidentifier",
  "lookup",
  "choice",
  "multichoice",
  "status",
  "state",
];

const column = (over: Partial<ColumnMapping> = {}): ColumnMapping => ({
  source: "C",
  target: "attr",
  kind: "string",
  treatEmptyAsNull: true,
  ...over,
});

/** Any value a spreadsheet reader can realistically hand us. */
const anyCellValue = () =>
  fc.oneof(
    fc.string(),
    fc.integer(),
    fc.double({ noDefaultInfinity: true, noNaN: true }),
    fc.boolean(),
    fc.constant(null),
    fc.constant(undefined),
    fc.date({ noInvalidDate: true }),
    fc.array(fc.string(), { maxLength: 5 })
  );

describe("coerceValue never fails in an uncontrolled way", () => {
  it("throws only CoerceError, whatever the cell contains", () => {
    fc.assert(
      fc.property(
        anyCellValue(),
        fc.constantFrom(...ALL_KINDS),
        fc.boolean(),
        (value, kind, treatEmptyAsNull) => {
          try {
            coerceValue(value, column({ kind, treatEmptyAsNull }));
          } catch (e) {
            // A TypeError or RangeError escaping here would surface to the
            // user as an unhandled crash mid-load instead of a row error.
            expect(e).toBeInstanceOf(CoerceError);
          }
        }
      )
    );
  });

  it("produces only JSON-serialisable output", () => {
    // Anything not JSON-serialisable would be silently dropped or mangled by
    // JSON.stringify when the batch body is built.
    fc.assert(
      fc.property(anyCellValue(), fc.constantFrom(...ALL_KINDS), (value, kind) => {
        let out: unknown;
        try {
          out = coerceValue(value, column({ kind }));
        } catch {
          return; // rejected inputs are fine
        }
        if (out === undefined) return;
        // "lookup" passes values through untouched by design; it is resolved
        // later in load.ts and never reaches JSON.stringify as-is.
        if (kind === "lookup") return;
        expect(() => JSON.stringify(out)).not.toThrow();
        expect(["string", "number", "boolean", "object"]).toContain(typeof out);
      })
    );
  });
});

describe("blank handling", () => {
  const blanks = () => fc.constantFrom(null, undefined, "", " ", "\t", "\n", "   \t  ");

  it("treatEmptyAsNull decides between clearing the field and omitting it", () => {
    fc.assert(
      fc.property(blanks(), fc.constantFrom(...ALL_KINDS), (blank, kind) => {
        expect(coerceValue(blank, column({ kind, treatEmptyAsNull: true }))).toBeNull();
        expect(coerceValue(blank, column({ kind, treatEmptyAsNull: false }))).toBeUndefined();
      })
    );
  });

  it("omitted attributes never reach the payload", () => {
    fc.assert(
      fc.property(blanks(), fc.constantFrom(...ALL_KINDS), (blank, kind) => {
        if (kind === "lookup") return;
        const { payload } = coerceRow({ C: blank }, [column({ kind, treatEmptyAsNull: false })]);
        expect(Object.prototype.hasOwnProperty.call(payload, "attr")).toBe(false);
      })
    );
  });
});

describe("numeric kinds", () => {
  it("integers are truncated, never rounded", () => {
    fc.assert(
      fc.property(fc.double({ min: -1e12, max: 1e12, noNaN: true, noDefaultInfinity: true }), (n) => {
        const out = coerceValue(n, column({ kind: "integer" }));
        expect(out).toBe(Math.trunc(n));
        expect(Number.isInteger(out)).toBe(true);
      })
    );
  });

  it("decimal/money/double round-trip a finite number unchanged", () => {
    fc.assert(
      fc.property(
        fc.double({ noNaN: true, noDefaultInfinity: true }),
        fc.constantFrom("decimal" as const, "money" as const, "double" as const),
        (n, kind) => {
          expect(coerceValue(n, column({ kind }))).toBe(n);
        }
      )
    );
  });

  it("a numeric string and its number are coerced identically", () => {
    fc.assert(
      fc.property(
        fc.integer({ min: -1e9, max: 1e9 }),
        fc.constantFrom("integer" as const, "decimal" as const, "money" as const, "double" as const),
        (n, kind) => {
          expect(coerceValue(String(n), column({ kind }))).toBe(coerceValue(n, column({ kind })));
        }
      )
    );
  });

  it("rejects non-numeric text rather than sending NaN to Dataverse", () => {
    fc.assert(
      fc.property(
        fc.string().filter((s) => s.trim() !== "" && !Number.isFinite(Number(s.trim()))),
        fc.constantFrom("integer" as const, "decimal" as const, "money" as const),
        (s, kind) => {
          expect(() => coerceValue(s, column({ kind }))).toThrow(CoerceError);
        }
      )
    );
  });
});

describe("boolean kind", () => {
  it("accepts the documented vocabulary in any casing or padding", () => {
    const truthy = ["true", "yes", "y", "1"];
    const falsy = ["false", "no", "n", "0"];
    fc.assert(
      fc.property(
        fc.constantFrom(...truthy, ...falsy),
        fc.string({ unit: fc.constantFrom(" ", "\t"), maxLength: 3 }),
        fc.boolean(),
        (word, pad, upper) => {
          const input = `${pad}${upper ? word.toUpperCase() : word}${pad}`;
          expect(coerceValue(input, column({ kind: "boolean" }))).toBe(truthy.includes(word));
        }
      )
    );
  });

  it("rejects anything outside that vocabulary", () => {
    const known = new Set(["true", "yes", "y", "1", "false", "no", "n", "0"]);
    fc.assert(
      fc.property(
        fc.string().filter((s) => s.trim() !== "" && !known.has(s.trim().toLowerCase())),
        (s) => {
          expect(() => coerceValue(s, column({ kind: "boolean" }))).toThrow(CoerceError);
        }
      )
    );
  });
});

describe("date kinds", () => {
  it("dateonly never shifts the calendar day for an ISO-shaped string", () => {
    // The classic bug: parsing "2025-01-01 00:00" as local time and then
    // reading UTC getters moves the date back a day west of Greenwich.
    fc.assert(
      fc.property(
        fc.date({
          min: new Date("1900-01-01T00:00:00Z"),
          max: new Date("2100-01-01T00:00:00Z"),
          noInvalidDate: true,
        }),
        fc.constantFrom("", "T00:00:00", " 00:00", "T23:59:59Z"),
        (d, suffix) => {
          const ymd = d.toISOString().slice(0, 10);
          expect(coerceValue(`${ymd}${suffix}`, column({ kind: "dateonly" }))).toBe(ymd);
        }
      )
    );
  });

  it("dateonly output always matches yyyy-MM-dd", () => {
    fc.assert(
      // Bounded to Dataverse's own DateTime range (min 1753-01-01). Outside
      // it the question is moot: neither Excel nor Dataverse can represent
      // the value, and a proleptic year like -0001 has no yyyy form at all.
      fc.property(
        fc.date({
          min: new Date("1753-01-01T00:00:00Z"),
          max: new Date("9999-12-31T00:00:00Z"),
          noInvalidDate: true,
        }),
        (d) => {
          const out = coerceValue(d, column({ kind: "dateonly" }));
          expect(String(out)).toMatch(/^\d{4}-\d{2}-\d{2}$/);
        }
      )
    );
  });

  it("datetime output is always a valid ISO instant", () => {
    fc.assert(
      fc.property(fc.date({ noInvalidDate: true }), (d) => {
        const out = coerceValue(d, column({ kind: "datetime" }));
        expect(new Date(String(out)).getTime()).toBe(d.getTime());
      })
    );
  });

  it("parseWithFormat round-trips any date it formats", () => {
    fc.assert(
      fc.property(
        fc.integer({ min: 1000, max: 9999 }),
        fc.integer({ min: 1, max: 12 }),
        fc.integer({ min: 1, max: 28 }), // 28 is valid in every month
        fc.constantFrom("dd/MM/yyyy", "yyyy-MM-dd", "M/d/yyyy", "dd.MM.yyyy"),
        (y, mo, d, format) => {
          const rendered = format
            .replace("yyyy", String(y).padStart(4, "0"))
            .replace("MM", String(mo).padStart(2, "0"))
            .replace("dd", String(d).padStart(2, "0"))
            .replace(/\bM\b/, String(mo))
            .replace(/\bd\b/, String(d));
          const parsed = parseWithFormat(rendered, format);
          expect(parsed).not.toBeNull();
          expect(parsed).toMatchObject({ y, mo, d });
        }
      )
    );
  });

  it("parseWithFormat rejects impossible dates instead of rolling them over", () => {
    // Date.UTC turns Feb 30 into Mar 2 without complaint; that silently
    // writes the wrong date to Dataverse.
    fc.assert(
      fc.property(fc.constantFrom([2, 30], [2, 31], [4, 31], [6, 31], [9, 31], [11, 31]), ([mo, d]) => {
        expect(parseWithFormat(`${String(d)}/${String(mo).padStart(2, "0")}/2024`, "d/MM/yyyy")).toBeNull();
      })
    );
  });
});

describe("multichoice", () => {
  it("emits a comma-separated integer list for any separator mix", () => {
    fc.assert(
      fc.property(
        fc.array(fc.integer({ min: 1, max: 999 }), { minLength: 1, maxLength: 8 }),
        fc.constantFrom(",", ";", ", ", " ; "),
        (values, sep) => {
          const out = coerceValue(values.join(sep), column({ kind: "multichoice" }));
          expect(out).toBe(values.join(","));
        }
      )
    );
  });

  it("maps labels through optionMap and rejects unknown labels", () => {
    // Labels are constrained the way real option-set labels are: no comma or
    // semicolon (those are the separators), and no surrounding whitespace
    // (the splitter trims, so a label that depends on it can never match —
    // see the explicit test below).
    const label = () =>
      fc
        .stringMatching(/^[A-Za-z][A-Za-z0-9 _-]{0,15}$/)
        .map((s) => s.trim())
        .filter((s) => s.length > 0);

    fc.assert(
      fc.property(fc.uniqueArray(label(), { minLength: 1, maxLength: 5 }), (labels) => {
        const optionMap = Object.fromEntries(labels.map((l, i) => [l, i + 1]));
        expect(coerceValue(labels.join(","), column({ kind: "multichoice", optionMap }))).toBe(
          labels.map((_, i) => i + 1).join(",")
        );
        expect(() =>
          coerceValue("definitely-not-a-label", column({ kind: "multichoice", optionMap }))
        ).toThrow(CoerceError);
      })
    );
  });

  it("cannot match an optionMap label that has surrounding whitespace", () => {
    // Documented consequence of trimming each part: if someone pastes option
    // labels with a leading space into a mapping, every row fails with
    // "unknown option" rather than silently writing the wrong choice. Pinned
    // so the behaviour can't drift into silent mismatching.
    const optionMap = { " Gold": 3 };
    expect(() => coerceValue(" Gold", column({ kind: "multichoice", optionMap }))).toThrow(CoerceError);
  });
});

describe("uniqueidentifier", () => {
  it("normalises to lowercase and is idempotent", () => {
    fc.assert(
      fc.property(fc.uuid(), fc.boolean(), (id, upper) => {
        const input = upper ? id.toUpperCase() : id;
        const once = coerceValue(input, column({ kind: "uniqueidentifier" }));
        expect(once).toBe(id.toLowerCase());
        expect(coerceValue(once, column({ kind: "uniqueidentifier" }))).toBe(once);
      })
    );
  });
});

describe("coerceRow", () => {
  it("keeps lookups out of the payload and reports them separately", () => {
    fc.assert(
      fc.property(fc.array(fc.constantFrom(...ALL_KINDS), { minLength: 1, maxLength: 10 }), (kinds) => {
        const columns = kinds.map((kind, i) => column({ kind, source: `C${i}`, target: `attr${i}` }));
        const row = Object.fromEntries(kinds.map((_, i) => [`C${i}`, "1"]));
        let out: ReturnType<typeof coerceRow>;
        try {
          out = coerceRow(row, columns);
        } catch {
          return; // some kinds legitimately reject "1"
        }
        expect(out.lookups.every((c) => c.kind === "lookup")).toBe(true);
        for (const c of out.lookups) {
          expect(Object.prototype.hasOwnProperty.call(out.payload, c.target)).toBe(false);
        }
      })
    );
  });
});
