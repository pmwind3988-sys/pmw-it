// What the file NAME says the sheet is about.
//
// A survey export is almost never called `export.xlsx`. It is called
// `Pain Points & Issues Across Departments (Responses).xlsx`, and that
// title is the only statement of intent anybody wrote down. Reading it
// costs nothing, needs no model, and is the difference between landing
// on "here are six charts about departments and the issues they raised"
// and landing on "here are six charts about whatever had the most
// distinct values".
//
// Deliberately lexical. A title is three to eight words, half of them
// export scaffolding; embedding it would be precision theatre over a
// sample far too small to earn it. The keywords come straight back out
// as a relevance nudge for chart ranking (`suggestCharts`) and as the
// tie-break for which written-answer column gets analysed first.

const EXTENSION = /\.(xlsx|xlsm|xls|csv|tsv)$/i;

// Everything a tool adds to a name that says nothing about the subject.
// Order matters only in that the date patterns must run before the
// separator collapse, while `-` and `_` still hold the parts apart.
const SCAFFOLDING = [
  /\bcopy of\b/gi,
  /\bform responses?\b/gi,
  /\bresponses?\b/gi,
  /\bsubmissions?\b/gi,
  /\b(final|latest|updated|new|old|draft|raw|clean(ed)?)\b/gi,
  /\bv\d+(\.\d+)?\b/gi,
  /\bexports?\b/gi,
  /\bsheet\s*\d*\b/gi,
  /\bdata\s*\d*\b/gi,
  // 2026-08-23, 23-08-2026, 23.08.2026, 20260823
  /\b\d{4}[-_.]\d{1,2}[-_.]\d{1,2}\b/g,
  /\b\d{1,2}[-_.]\d{1,2}[-_.]\d{2,4}\b/g,
  /\b\d{8}\b/g,
  // The "(1)" a second download picks up.
  /\(\s*\d+\s*\)/g,
];

// A file called `book1` or `untitled` has told us nothing, and the sheet
// tab is the next best thing to read.
const EMPTY_TITLES = new Set(['', 'book1', 'book', 'untitled', 'workbook', 'report', 'file']);

const STOPWORDS = new Set([
  'the', 'and', 'for', 'with', 'from', 'across', 'per', 'all', 'our', 'their',
  'into', 'onto', 'about', 'that', 'this', 'these', 'those', 'each', 'any',
  'list', 'form', 'survey', 'sheet', 'table', 'report', 'summary', 'analysis',
]);

/**
 * What kind of sheet this is, in the order the tests are applied.
 *
 * `textFirst` is the consequential field: it is what decides whether the
 * written answers are analysed on arrival or merely offered. The
 * analysis pulls a 23MB model on first use, so it fires when the title
 * says the writing IS the data -- a pain-point survey -- and waits to be
 * asked on a stock list that happens to carry a notes column.
 */
const KINDS = [
  {
    kind: 'issues',
    textFirst: true,
    label: 'issues raised',
    words: [
      'pain', 'point', 'issue', 'problem', 'complaint', 'challenge', 'blocker',
      'bottleneck', 'friction', 'gap', 'difficulty', 'frustration', 'obstacle',
    ],
  },
  {
    kind: 'feedback',
    textFirst: true,
    label: 'written feedback',
    words: [
      'feedback', 'satisfaction', 'opinion', 'sentiment', 'review', 'comment',
      'suggestion', 'idea', 'voice', 'engagement', 'poll', 'questionnaire',
    ],
  },
  {
    kind: 'tickets',
    textFirst: false,
    label: 'requests and tickets',
    words: ['ticket', 'incident', 'helpdesk', 'servicedesk', 'sla', 'escalation', 'case'],
  },
  {
    kind: 'inventory',
    textFirst: false,
    label: 'an inventory',
    words: [
      'asset', 'inventory', 'device', 'equipment', 'stock', 'hardware',
      'licence', 'license', 'register', 'machine', 'laptop',
    ],
  },
  {
    kind: 'finance',
    textFirst: false,
    label: 'money',
    words: ['cost', 'budget', 'spend', 'invoice', 'revenue', 'expense', 'payment', 'billing'],
  },
];

function stripScaffolding(name) {
  let out = String(name ?? '').replace(EXTENSION, '');
  for (const pattern of SCAFFOLDING) out = out.replace(pattern, ' ');
  return out
    .replace(/[_\-.]+/g, ' ')
    .replace(/[()[\]{}]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

// A word already carrying a capital is left exactly as typed. That is
// the only way `SAP`, `IT` and `HR` survive a title pass, and getting
// those wrong is more noticeable than any casing it fixes.
function titleCase(text) {
  return text
    .split(' ')
    .map((word) => (word === word.toLowerCase()
      ? word.charAt(0).toUpperCase() + word.slice(1)
      : word))
    .join(' ');
}

// Crude on purpose: "departments" and "department" must land on the same
// token for the relevance nudge to fire, and a real stemmer is a
// dependency for one rule.
function singular(word) {
  if (word.length > 4 && word.endsWith('ies')) return `${word.slice(0, -3)}y`;
  if (word.length > 4 && word.endsWith('sses')) return word.slice(0, -2);
  if (word.length > 3 && word.endsWith('s') && !word.endsWith('ss')) return word.slice(0, -1);
  return word;
}

export function keywordsOf(text) {
  const seen = new Set();
  for (const raw of String(text ?? '').toLowerCase().split(/[^a-z0-9]+/)) {
    if (raw.length < 3) continue;
    if (STOPWORDS.has(raw)) continue;
    seen.add(singular(raw));
  }
  return [...seen];
}

/**
 * Read a file name (and the sheet tab, when the name says nothing) into
 * a title, a keyword set and a guess at what the sheet is for.
 */
export function readFileIntent(fileName, sheetName = '') {
  const fromFile = stripScaffolding(fileName);
  const fromSheet = stripScaffolding(sheetName);

  // The sheet tab is a fallback, not an addition: on a Google Forms
  // export the tab is called "Form Responses 1" and folding it in would
  // dilute a perfectly good title with nothing.
  const source = EMPTY_TITLES.has(fromFile.toLowerCase()) && fromSheet ? fromSheet : fromFile;
  const title = titleCase(source);
  const keywords = keywordsOf(source);

  const words = new Set(keywords);
  const matched = KINDS.find((candidate) => candidate.words.some((w) => words.has(singular(w))));

  return {
    title,
    keywords,
    kind: matched?.kind ?? 'generic',
    label: matched?.label ?? 'this data',
    textFirst: matched?.textFirst ?? false,
  };
}

/**
 * How strongly a column name answers the question the title asks.
 *
 * Returns 0 when nothing overlaps, which is the common case and must
 * cost a chart nothing -- a nudge that penalises every unmatched column
 * is a re-ranking, not a nudge.
 */
export function relevanceOf(columnName, keywords = []) {
  if (keywords.length === 0) return 0;
  const words = new Set(keywordsOf(columnName));
  if (words.size === 0) return 0;
  let hits = 0;
  for (const keyword of keywords) if (words.has(keyword)) hits += 1;
  return hits / keywords.length;
}
