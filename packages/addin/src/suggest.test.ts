import { describe, it } from "node:test";
import assert from "node:assert/strict";
import {
  normalize,
  levenshtein,
  similarity,
  suggestMappings,
  suggestionsToMappings,
} from "./suggest.js";

// ─── normalize ────────────────────────────────────────────────────────────────

describe("normalize", () => {
  it("lowercases and strips non-alphanumeric characters", () => {
    assert.equal(normalize("Email Address"), "emailaddress");
    assert.equal(normalize("First Name"), "firstname");
    assert.equal(normalize("ZIP Code!"), "zipcode");
  });

  it("preserves digits", () => {
    assert.equal(normalize("emailaddress1"), "emailaddress1");
    assert.equal(normalize("address1line1"), "address1line1");
  });

  it("returns empty string for blank input", () => {
    assert.equal(normalize(""), "");
    assert.equal(normalize("---"), "");
  });
});

// ─── levenshtein ─────────────────────────────────────────────────────────────

describe("levenshtein", () => {
  it("returns 0 for identical strings", () => {
    assert.equal(levenshtein("hello", "hello"), 0);
    assert.equal(levenshtein("", ""), 0);
  });

  it("returns the length of the other string when one is empty", () => {
    assert.equal(levenshtein("", "abc"), 3);
    assert.equal(levenshtein("abc", ""), 3);
  });

  it("computes single-character substitution", () => {
    assert.equal(levenshtein("cat", "bat"), 1);
    assert.equal(levenshtein("cat", "car"), 1);
  });

  it("computes insertion and deletion", () => {
    assert.equal(levenshtein("abc", "ab"), 1);
    assert.equal(levenshtein("ab", "abc"), 1);
  });

  it("kitten → sitting = 3", () => {
    assert.equal(levenshtein("kitten", "sitting"), 3);
  });

  it("is symmetric", () => {
    assert.equal(levenshtein("hello", "world"), levenshtein("world", "hello"));
  });
});

// ─── similarity ──────────────────────────────────────────────────────────────

describe("similarity", () => {
  it("returns 1.0 for identical strings (after normalize)", () => {
    assert.equal(similarity("email", "email"), 1.0);
    assert.equal(similarity("First Name", "firstname"), 1.0);
  });

  it("returns 0 for empty strings", () => {
    assert.equal(similarity("", "emailaddress1"), 0);
    assert.equal(similarity("emailaddress1", ""), 0);
  });

  it("scores synonyms at 0.9", () => {
    assert.equal(similarity("firstname", "givenname"), 0.9);
    assert.equal(similarity("email", "emailaddress1"), 0.9);
    assert.equal(similarity("birthday", "birthdate"), 0.9);
    assert.equal(similarity("zip", "address1postalcode"), 0.9);
  });

  it("scores substring containment above 0.7 and below 0.9", () => {
    // "account" is contained in "accountname" but they are not synonyms,
    // so the substring branch fires: 0.7 + 0.1 * (7/11) ≈ 0.764.
    const s = similarity("account", "accountname");
    assert.ok(s > 0.7, `expected > 0.7, got ${s}`);
    assert.ok(s < 0.9, `expected < 0.9, got ${s}`);
  });

  it("scores unrelated strings below 0.7", () => {
    const s = similarity("birthday", "accountnumber");
    assert.ok(s < 0.7, `expected < 0.7, got ${s}`);
  });
});

// ─── suggestMappings ─────────────────────────────────────────────────────────

describe("suggestMappings", () => {
  const sources = ["Email", "First Name", "Last Name", "Birthday", "Phone"];
  const targets = [
    "emailaddress1",
    "firstname",
    "lastname",
    "birthdate",
    "telephone1",
    "accountid",
  ];

  it("matches common contact columns to Dataverse attributes", () => {
    const suggestions = suggestMappings(sources, targets);
    const pairs = Object.fromEntries(suggestions.map((s) => [s.source, s.target]));
    assert.equal(pairs["Email"], "emailaddress1");
    assert.equal(pairs["First Name"], "firstname");
    assert.equal(pairs["Last Name"], "lastname");
    assert.equal(pairs["Birthday"], "birthdate");
  });

  it("returns at most one suggestion per source and per target", () => {
    const suggestions = suggestMappings(sources, targets);
    const usedSources = suggestions.map((s) => s.source);
    const usedTargets = suggestions.map((s) => s.target);
    assert.equal(usedSources.length, new Set(usedSources).size, "sources should be unique");
    assert.equal(usedTargets.length, new Set(usedTargets).size, "targets should be unique");
  });

  it("respects the threshold — nothing below it is returned", () => {
    const suggestions = suggestMappings(["xyz123"], targets, { threshold: 0.7 });
    assert.equal(suggestions.length, 0);
  });

  it("excludeTargets prevents those targets from appearing", () => {
    const suggestions = suggestMappings(sources, targets, {
      excludeTargets: ["emailaddress1", "firstname"],
    });
    const mappedTargets = suggestions.map((s) => s.target);
    assert.ok(!mappedTargets.includes("emailaddress1"));
    assert.ok(!mappedTargets.includes("firstname"));
  });

  it("excludeSources prevents those sources from appearing", () => {
    const suggestions = suggestMappings(sources, targets, {
      excludeSources: ["Email", "First Name"],
    });
    const mappedSources = suggestions.map((s) => s.source);
    assert.ok(!mappedSources.includes("Email"));
    assert.ok(!mappedSources.includes("First Name"));
  });

  it("returns empty array when sources or targets are empty", () => {
    assert.deepEqual(suggestMappings([], targets), []);
    assert.deepEqual(suggestMappings(sources, []), []);
  });

  it("returns suggestions sorted by score descending", () => {
    const suggestions = suggestMappings(sources, targets);
    for (let i = 1; i < suggestions.length; i++) {
      assert.ok(
        suggestions[i - 1].score >= suggestions[i].score,
        "suggestions should be sorted by descending score"
      );
    }
  });
});

// ─── suggestionsToMappings ───────────────────────────────────────────────────

describe("suggestionsToMappings", () => {
  it("converts suggestions to ColumnMapping skeletons", () => {
    const suggestions = [{ source: "Email", target: "emailaddress1", score: 1.0 }];
    const mappings = suggestionsToMappings(suggestions);
    assert.equal(mappings.length, 1);
    assert.equal(mappings[0].source, "Email");
    assert.equal(mappings[0].target, "emailaddress1");
    assert.equal(mappings[0].kind, "string");
    assert.equal(mappings[0].treatEmptyAsNull, true);
  });

  it("includes the score in the notes field", () => {
    const suggestions = [{ source: "Email", target: "emailaddress1", score: 0.95 }];
    const mappings = suggestionsToMappings(suggestions);
    assert.ok(mappings[0].notes?.includes("0.95"));
  });

  it("returns an empty array for empty input", () => {
    assert.deepEqual(suggestionsToMappings([]), []);
  });
});
