// Name-based mapping suggestions. Given a list of source column headers and
// a list of target Dataverse attribute logical names, produce a best-effort
// pairing. The algorithm is deliberately simple:
//
//   1. Normalize both sides: lowercase, strip non-alphanumeric. So
//      "Email Address" and "emailaddress1" both become "emailaddress" /
//      "emailaddress1" — close enough to be picked up by step 2.
//
//   2. Score every (source, target) pair:
//        normalized exact match           → 1.0
//        normalized common-synonym match  → 0.9   (e.g. firstname ↔ givenname)
//        one is a prefix/suffix/contained → 0.8
//        else                             → 1 - levenshtein / max(len)
//
//   3. Sort all pairs by score descending; greedy-match — each source and
//      each target can only appear in one suggestion. Stop below the
//      threshold.
//
// This catches the common cases ("Email" → "emailaddress1", "First Name" →
// "firstname") without trying to be clever about semantics.

import type { ColumnMapping } from "@dvload/core";

export interface Suggestion {
  source: string;
  target: string;
  score: number;
}

/** Pairs of synonyms we want to boost above the Levenshtein floor. */
const SYNONYMS: Record<string, string[]> = {
  // Source-side normalized form → target normalized forms it commonly maps to.
  firstname: ["givenname", "firstname"],
  givenname: ["firstname"],
  lastname: ["surname", "familyname", "lastname"],
  surname: ["lastname"],
  email: ["emailaddress1", "emailaddress"],
  phone: ["telephone1", "mobilephone", "phone1"],
  mobile: ["mobilephone", "telephone2"],
  company: ["parentcustomerid", "companyname", "accountname"],
  account: ["parentcustomerid", "accountid"],
  zip: ["address1postalcode", "postalcode"],
  zipcode: ["address1postalcode", "postalcode"],
  postcode: ["address1postalcode", "postalcode"],
  city: ["address1city"],
  state: ["address1stateorprovince"],
  country: ["address1country"],
  street: ["address1line1"],
  address: ["address1line1", "address1composite"],
  birthday: ["birthdate"],
  dob: ["birthdate"],
  fullname: ["fullname", "name"],
  title: ["jobtitle"],
};

export function normalize(s: string): string {
  return s.toLowerCase().replace(/[^a-z0-9]/g, "");
}

export function levenshtein(a: string, b: string): number {
  if (a === b) return 0;
  if (a.length === 0) return b.length;
  if (b.length === 0) return a.length;

  let prev = new Array(b.length + 1);
  let curr = new Array(b.length + 1);
  for (let j = 0; j <= b.length; j++) prev[j] = j;

  for (let i = 1; i <= a.length; i++) {
    curr[0] = i;
    for (let j = 1; j <= b.length; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      curr[j] = Math.min(prev[j] + 1, curr[j - 1] + 1, prev[j - 1] + cost);
    }
    [prev, curr] = [curr, prev];
  }
  return prev[b.length];
}

export function similarity(source: string, target: string): number {
  const ns = normalize(source);
  const nt = normalize(target);
  if (!ns || !nt) return 0;
  if (ns === nt) return 1;

  // Synonym match: if a synonym for `ns` lists `nt` (or vice versa).
  const synSource = SYNONYMS[ns];
  if (synSource?.includes(nt)) return 0.9;
  const synTarget = SYNONYMS[nt];
  if (synTarget?.includes(ns)) return 0.9;

  // Substring containment: catches "Email" ↔ "emailaddress1".
  if (nt.includes(ns) || ns.includes(nt)) {
    const shorter = Math.min(ns.length, nt.length);
    const longer = Math.max(ns.length, nt.length);
    // Bias toward 0.8 but penalize length disparity slightly.
    return 0.7 + 0.1 * (shorter / longer);
  }

  // Levenshtein fallback.
  const dist = levenshtein(ns, nt);
  const maxLen = Math.max(ns.length, nt.length);
  return 1 - dist / maxLen;
}

export interface SuggestOptions {
  /** Minimum similarity to consider a suggestion. Default 0.7. */
  threshold?: number;
  /** Exclude these targets from suggestions (e.g. already-mapped ones). */
  excludeTargets?: Iterable<string>;
  /** Exclude these sources from suggestions. */
  excludeSources?: Iterable<string>;
}

export function suggestMappings(
  sources: string[],
  targets: string[],
  opts: SuggestOptions = {}
): Suggestion[] {
  const threshold = opts.threshold ?? 0.7;
  const excludeTargets = new Set(opts.excludeTargets ?? []);
  const excludeSources = new Set(opts.excludeSources ?? []);

  const pairs: Suggestion[] = [];
  for (const s of sources) {
    if (excludeSources.has(s)) continue;
    for (const t of targets) {
      if (excludeTargets.has(t)) continue;
      const score = similarity(s, t);
      if (score >= threshold) pairs.push({ source: s, target: t, score });
    }
  }
  pairs.sort((a, b) => b.score - a.score);

  const usedSources = new Set<string>();
  const usedTargets = new Set<string>();
  const out: Suggestion[] = [];
  for (const p of pairs) {
    if (usedSources.has(p.source) || usedTargets.has(p.target)) continue;
    out.push(p);
    usedSources.add(p.source);
    usedTargets.add(p.target);
  }
  return out;
}

/**
 * Convenience: given suggestions, build ColumnMapping skeletons. Type
 * defaults to "string"; the user can refine in the UI.
 */
export function suggestionsToMappings(suggestions: Suggestion[]): ColumnMapping[] {
  return suggestions.map((s) => ({
    source: s.source,
    target: s.target,
    kind: "string",
    treatEmptyAsNull: true,
    notes: `Auto-suggested (score ${s.score.toFixed(2)})`,
  }));
}
