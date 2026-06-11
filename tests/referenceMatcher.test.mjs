import assert from "node:assert/strict";
import test from "node:test";
import {
  buildLookup,
  citationBase,
  findReferenceMatch,
  findWorkDocMatch,
  firstAuthorBase,
  recordsFromCsv
} from "../src/referenceMatcher.js";

const lookup = buildLookup([
  {
    title: "A software package for assessing terrestrial planetary boundaries",
    doi: "10.1016/j.oneear.2025.101341",
    author: "gerten et al 2025a"
  },
  {
    title: "Planetary boundaries: Guiding human development on a changing planet",
    doi: "10.1126/science.1259855",
    author: "steffen et al 2015a"
  },
  {
    title: "Planetary Boundaries: Exploring the Safe Operating Space for Humanity",
    doi: "10.5751/ES-03180-140232",
    author: "rockstrom et al 2009a"
  },
  {
    title: "A safe operating space for humanity",
    doi: "10.1038/461472a",
    author: "rockstrom et al 2009b"
  }
]);

test("citationBase strips numbered working suffixes", () => {
  assert.equal(citationBase("Gerten et al 2025_1"), "gerten et al 2025");
  assert.equal(citationBase("gerten et al 2025_7"), "gerten et al 2025");
});

test("citationBase strips lookup letter suffixes", () => {
  assert.equal(citationBase("gerten et al 2025a"), "gerten et al 2025");
});

test("single lookup citation is returned without assuming number-letter order", () => {
  const match = findReferenceMatch(lookup, "gerten et al 2025_9");
  assert.equal(match.matched, true);
  assert.equal(match.ambiguous, false);
  assert.deepEqual(match.citations, ["gerten et al 2025a"]);
});

test("multiple lookup citations are flagged as ambiguous", () => {
  const match = findReferenceMatch(lookup, "rockstrom et al 2009_1");
  assert.equal(match.matched, true);
  assert.equal(match.ambiguous, true);
  assert.deepEqual(match.citations, ["rockstrom et al 2009a", "rockstrom et al 2009b"]);
});

test("missing citation leaves original value unmatched", () => {
  const match = findReferenceMatch(lookup, "unknown et al 2026_1");
  assert.equal(match.matched, false);
  assert.deepEqual(match.citations, []);
});

test("work doc uses DOI first", () => {
  const match = findWorkDocMatch(lookup, {
    doi: "https://doi.org/10.1016/j.oneear.2025.101341",
    title: "Different title",
    authors: "Someone Else",
    year: 2025
  });
  assert.equal(match.matched, true);
  assert.equal(match.kind, "doi");
  assert.equal(match.citations[0], "gerten et al 2025a");
});

test("work doc normalizes DOI prefixes", () => {
  const match = findWorkDocMatch(lookup, {
    doi: "doi: 10.1126/science.1259855",
    title: "",
    authors: "",
    year: ""
  });
  assert.equal(match.matched, true);
  assert.equal(match.kind, "doi");
  assert.equal(match.citations[0], "steffen et al 2015a");
});

test("work doc can fall back to title and year", () => {
  const match = findWorkDocMatch(lookup, {
    doi: "",
    title: "Planetary boundaries: Guiding human development on a changing planet",
    authors: "Someone Else",
    year: 2015
  });
  assert.equal(match.matched, true);
  assert.equal(match.kind, "title-year");
  assert.equal(match.citations[0], "steffen et al 2015a");
});

test("work doc can fall back to first-author year", () => {
  assert.equal(firstAuthorBase("Gerten, Braun, Breier et al.", 2025), "gerten et al 2025");
  const match = findWorkDocMatch(lookup, {
    doi: "",
    title: "Different title",
    authors: "Gerten, Braun, Breier et al.",
    year: 2025
  });
  assert.equal(match.matched, true);
  assert.equal(match.kind, "author-year");
});

test("CSV parser handles quoted commas", () => {
  const records = recordsFromCsv('title,doi,author\n"Paper, with comma",10.1/test,"name et al 2025a"\n');
  assert.deepEqual(records, [
    {
      title: "Paper, with comma",
      doi: "10.1/test",
      author: "name et al 2025a"
    }
  ]);
});
