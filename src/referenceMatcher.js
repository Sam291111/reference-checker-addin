const AUTHOR_YEAR_WITH_SUFFIX = /^(.+?\b\d{4})([a-z])$/i;
const AUTHOR_YEAR_WITH_NUMBER = /^(.+?\b\d{4})_\d+$/i;

export function normalizeWhitespace(value) {
  return String(value ?? "")
    .replace(/\s+/g, " ")
    .trim();
}

export function normalizeKey(value) {
  return normalizeWhitespace(value)
    .toLowerCase()
    .replace(/[’']/g, "")
    .replace(/[.,;:()[\]{}]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

export function normalizeDoi(value) {
  return normalizeWhitespace(value)
    .toLowerCase()
    .replace(/^https?:\/\/(dx\.)?doi\.org\//, "")
    .replace(/^doi:\s*/, "")
    .trim();
}

export function normalizeTitle(value) {
  return normalizeKey(value)
    .replace(/[-–—]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

export function citationBase(value) {
  const normalized = normalizeKey(value);
  const numbered = normalized.match(AUTHOR_YEAR_WITH_NUMBER);
  if (numbered) {
    return numbered[1];
  }

  const lookup = normalized.match(AUTHOR_YEAR_WITH_SUFFIX);
  if (lookup) {
    return lookup[1];
  }

  return normalized;
}

export function parseCsv(text) {
  const rows = [];
  let row = [];
  let cell = "";
  let inQuotes = false;

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    const next = text[index + 1];

    if (char === '"' && inQuotes && next === '"') {
      cell += '"';
      index += 1;
      continue;
    }

    if (char === '"') {
      inQuotes = !inQuotes;
      continue;
    }

    if (char === "," && !inQuotes) {
      row.push(cell);
      cell = "";
      continue;
    }

    if ((char === "\n" || char === "\r") && !inQuotes) {
      if (char === "\r" && next === "\n") {
        index += 1;
      }
      row.push(cell);
      if (row.some((value) => normalizeWhitespace(value) !== "")) {
        rows.push(row);
      }
      row = [];
      cell = "";
      continue;
    }

    cell += char;
  }

  row.push(cell);
  if (row.some((value) => normalizeWhitespace(value) !== "")) {
    rows.push(row);
  }

  return rows;
}

export function recordsFromCsv(text) {
  const rows = parseCsv(text);
  if (rows.length === 0) {
    return [];
  }

  const headers = rows[0].map((header) => normalizeKey(header));
  return rows.slice(1).map((row) => {
    const record = {};
    headers.forEach((header, index) => {
      record[header] = normalizeWhitespace(row[index] ?? "");
    });
    return record;
  });
}

export function firstAuthorBase(authors, year) {
  const cleanYear = String(year ?? "").match(/\d{4}/)?.[0];
  const cleanAuthors = normalizeWhitespace(authors);
  if (!cleanAuthors || !cleanYear) {
    return "";
  }

  const firstChunk = cleanAuthors.split(/,|;| and |\bet al\.?/i)[0] || "";
  const surname = normalizeKey(firstChunk).split(" ")[0] || "";
  return surname ? `${surname} et al ${cleanYear}` : "";
}

export function buildLookup(records) {
  const byDoi = new Map();
  const byCitationBase = new Map();
  const byTitleYear = new Map();

  const entries = records
    .map((record, index) => {
      const citation = normalizeWhitespace(record.author);
      const doi = normalizeDoi(record.doi);
      const title = normalizeWhitespace(record.title);
      const titleKey = normalizeTitle(title);
      const year = citation.match(/\b\d{4}(?=[a-z]?\b)/i)?.[0] || "";
      const base = citationBase(citation);
      return {
        index,
        citation,
        doi,
        title,
        titleKey,
        year,
        base
      };
    })
    .filter((entry) => entry.citation || entry.doi || entry.title);

  for (const entry of entries) {
    if (entry.doi) {
      addToMap(byDoi, entry.doi, entry);
    }
    if (entry.base) {
      addToMap(byCitationBase, entry.base, entry);
    }
    if (entry.titleKey && entry.year) {
      addToMap(byTitleYear, `${entry.titleKey}|${entry.year}`, entry);
    }
  }

  return {
    entries,
    byDoi,
    byCitationBase,
    byTitleYear
  };
}

function addToMap(map, key, entry) {
  if (!map.has(key)) {
    map.set(key, []);
  }
  map.get(key).push(entry);
}

export function findReferenceMatch(lookup, typedReference) {
  const base = citationBase(typedReference);
  const matches = lookup.byCitationBase.get(base) || [];
  return toMatchResult("reference", typedReference, matches);
}

export function findWorkDocMatch(lookup, row) {
  const doi = normalizeDoi(row.doi);
  if (doi) {
    const matches = lookup.byDoi.get(doi) || [];
    if (matches.length > 0) {
      return toMatchResult("doi", row.doi, matches);
    }
  }

  const titleKey = normalizeTitle(row.title);
  const year = String(row.year ?? "").match(/\d{4}/)?.[0] || "";
  if (titleKey && year) {
    const matches = lookup.byTitleYear.get(`${titleKey}|${year}`) || [];
    if (matches.length > 0) {
      return toMatchResult("title-year", row.title, matches);
    }
  }

  const authorBase = firstAuthorBase(row.authors || row.author, year);
  if (authorBase) {
    const matches = lookup.byCitationBase.get(authorBase) || [];
    if (matches.length > 0) {
      return toMatchResult("author-year", authorBase, matches);
    }
  }

  return toMatchResult("work-doc", row.doi || row.title || "", []);
}

function toMatchResult(kind, value, matches) {
  const uniqueMatches = uniqueEntries(matches);
  return {
    kind,
    value: normalizeWhitespace(value),
    matched: uniqueMatches.length > 0,
    ambiguous: uniqueMatches.length > 1,
    entries: uniqueMatches,
    citations: uniqueMatches.map((entry) => entry.citation).filter(Boolean)
  };
}

function uniqueEntries(entries) {
  const seen = new Set();
  const unique = [];

  for (const entry of entries) {
    const key = `${entry.doi}|${entry.citation}|${entry.title}`;
    if (!seen.has(key)) {
      seen.add(key);
      unique.push(entry);
    }
  }

  return unique;
}

export function isReferenceHeader(value) {
  return /^reference_\d+$/i.test(normalizeWhitespace(value));
}

export function detectHeaderRow(values) {
  const maxRows = Math.min(values.length, 10);
  for (let rowIndex = 0; rowIndex < maxRows; rowIndex += 1) {
    const normalized = values[rowIndex].map((value) => normalizeKey(value));
    const hasReferenceColumns = normalized.some(isReferenceHeader);
    const hasWorkDocColumns = normalized.includes("doi") && normalized.includes("authors") && normalized.includes("year");
    if (hasReferenceColumns || hasWorkDocColumns) {
      return rowIndex;
    }
  }
  return 0;
}
