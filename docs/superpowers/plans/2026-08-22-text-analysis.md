# Data Studio Text Analysis Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Give Data Studio a Text Analysis tab that splits free-text survey answers into separate issues, files them into editable categories, discovers themes, ranks them, and writes the results back as ordinary dataset columns — entirely in the browser.

**Architecture:** Two independent halves. Part A (Tasks 1–5) adds a `multi` column type so semicolon-joined multi-select answers become chartable — no AI, useful on its own, and it supplies the "breadth" signal Part B needs. Part B (Tasks 6–25) adds the `text/` feature directory: pure functions for splitting, scoring, matching, clustering, ranking and overriding, with a single impure `embed.js` wrapping a local sentence-embedding model, driven from a dedicated worker.

**Tech Stack:** React 19, Vite 8, vitest, `@huggingface/transformers` (ONNX on WebAssembly), `Xenova/all-MiniLM-L6-v2` quantized.

**Spec:** `docs/superpowers/specs/2026-08-22-text-analysis-design.md`

## Global Constraints

- **No network call ever carries response text.** No SharePoint, no Graph, no API, no telemetry. Model and ONNX runtime are served from this app's own origin (`env.allowRemoteModels = false`, `env.localModelPath = '/models/'`, `env.backends.onnx.wasm.wasmPaths = '/ort/'`). Local IndexedDB persistence is required and permitted.
- **Pure layers stay pure.** `text/` modules other than `embed.js`, `text.worker.js` and `*.jsx` are pure functions of their arguments. `ingest/` → `profile/` → `clean/` → `engine/` → `text/`; only `.jsx` imports React.
- **A module exporting a component exports nothing else** — or it drops out of Fast Refresh and fails `npm run lint`.
- **Invisible characters are written as escapes**, never as literals (`\u200B`, `\uFEFF`).
- **Null encodings are a contract** (`engine/dataset.js`): numeric/temporal `NaN`, categorical `-1`, boolean `2`, text `null`.
- **No new npm dependency** beyond `@huggingface/transformers`.
- Tests: `npm test`. Lint: `npm run lint` (pre-existing errors in FormPage, AssetChecklistPage, SignatureDialog, ThemeContext are expected — do not fix, do not add new ones).

---

## File Structure

**Part A — multi-value columns**

| File | Responsibility |
|---|---|
| `src/features/datastudio/profile/inferType.js` (modify) | detect a multi-select column |
| `src/features/datastudio/profile/profileColumn.js` (modify) | rank the *options* of a multi column, not its raw strings |
| `src/features/datastudio/engine/dataset.js` (modify) | CSR encoding for multi; rowCount derived from raw input |
| `src/features/datastudio/engine/aggregate.js` (modify) | one row can land in several categories |
| `src/features/datastudio/engine/filterMask.js` (modify) | a row matches if any of its options match |
| `src/features/datastudio/clean/cleanOps.js` (modify) | `castType` normalises a multi column |
| `src/features/datastudio/clean/proposeCleanPlan.js` (modify) | offer that cast |
| `src/pages/DataStudioPage.jsx` (modify) | `multi` in the type dropdown |

**Part B — text analysis**

| File | Responsibility |
|---|---|
| `src/features/datastudio/text/boilerplate.js` | label prefixes + non-answer lexicon |
| `src/features/datastudio/text/splitIssues.js` | one answer → separate issue fragments |
| `src/features/datastudio/text/detectTextColumns.js` | which columns qualify as free text |
| `src/features/datastudio/text/buckets.js` | the starter bucket set |
| `src/features/datastudio/text/severity.js` | wording/length/breadth/emphasis → 0–1 |
| `src/features/datastudio/text/similarity.js` | cosine; fragment → bucket or Unsorted |
| `src/features/datastudio/text/cluster.js` | agglomerative clustering → themes |
| `src/features/datastudio/text/labelCluster.js` | c-TF-IDF naming |
| `src/features/datastudio/text/rankIssues.js` | groups → priority order |
| `src/features/datastudio/text/overrides.js` | user edits applied over a raw analysis |
| `src/features/datastudio/text/analysis.js` | orchestration |
| `src/features/datastudio/text/deriveColumns.js` | analysis → new dataset columns |
| `src/features/datastudio/text/embed.js` | THE ONLY IMPURE MODULE |
| `src/features/datastudio/worker/text.worker.js` | model lifetime + progress |
| `src/features/datastudio/text/TextAnalysis.jsx` | the stage |
| `src/features/datastudio/text/BucketEditor.jsx` | |
| `src/features/datastudio/text/IssueTable.jsx` | |
| `src/features/datastudio/text/ThemeList.jsx` | |
| `src/features/datastudio/text/PriorityBoard.jsx` | |
| `scripts/fetch-model.mjs` | pinned model + ORT fetch into `public/` |
| `src/features/datastudio/store/db.js` (modify) | `analyses` store, DB_VERSION 2 |
| `src/features/datastudio/dataStudioStore.js` (modify) | `text` stage + state |
| `src/features/datastudio/DataStudioContext.jsx` (modify) | text worker + actions |
| `src/styles/datastudio.css` (modify) | the tab's styles |

---

# PART A — Multi-value columns

### Task 1: Detect a multi-select column

**Files:**
- Modify: `src/features/datastudio/profile/inferType.js`
- Modify: `src/features/datastudio/profile/profileColumn.js` (`ROLE_BY_TYPE`)
- Test: `src/features/datastudio/profile/inferType.test.js`

**Interfaces:**
- Produces: `inferType(values, name)` may now return `{ type: 'multi', role: 'dimension', separator: ';', ... }`. `ROLE_BY_TYPE.multi === 'dimension'`.

- [ ] **Step 1: Write the failing tests**

Append to `src/features/datastudio/profile/inferType.test.js`:

```js
describe('multi-select detection', () => {
  const survey = [
    'Data Collection;Data Cleaning;Report Generation;',
    'Data Collection;Approval Tracking;',
    'Report Generation;Data Collection;',
    'Approval Tracking;',
    'Data Cleaning;Report Generation;',
    'Data Collection;Data Cleaning;',
  ];

  it('reads a semicolon-joined multi-select as multi', () => {
    const verdict = inferType(survey, 'Which challenges');
    expect(verdict.type).toBe('multi');
    expect(verdict.role).toBe('dimension');
    expect(verdict.separator).toBe(';');
  });

  it('leaves an ordinary categorical column alone', () => {
    const verdict = inferType(['IT', 'Finance', 'IT', 'Logistics', 'IT'], 'Department');
    expect(verdict.type).toBe('categorical');
  });

  it('does not read prose containing semicolons as multi', () => {
    // Parts that never repeat are sentences, not options.
    const prose = [
      'The process is manual; it takes hours to reconcile every month.',
      'We chase approvals by email; nobody knows the current status.',
      'Reports are rebuilt from scratch; version control is guesswork.',
      'Files live in five places; finding the latest one is luck.',
      'Data is retyped between systems; typos are common.',
      'Updates arrive on WhatsApp; important ones get missed.',
    ];
    expect(inferType(prose, 'Describe').type).not.toBe('multi');
  });

  it('does not read a single-option column as multi', () => {
    const single = ['IT;', 'Finance;', 'IT;', 'Logistics;', 'IT;', 'Finance;'];
    expect(inferType(single, 'Department').type).not.toBe('multi');
  });
});
```

- [ ] **Step 2: Run to verify they fail**

```bash
npm test -- inferType
```

Expected: FAIL — the first test reports `'categorical'` where `'multi'` is expected.

- [ ] **Step 3: Implement the detection**

In `src/features/datastudio/profile/inferType.js`, add above the `inferType` export:

```js
export const MULTI_SEPARATORS = [';', '|'];

// A multi-select answer is many options in one cell. Three conditions
// have to hold together, and each one rules out a different impostor:
//
//   * most values carry the separator          -- or it is prose that
//                                                 happens to contain one
//   * the average cell holds more than one     -- or it is a plain
//     option                                      category with a stray
//                                                 trailing separator
//   * the options repeat across rows           -- or the "options" are
//                                                 sentences, and every
//                                                 one is unique
//
// The last is the load-bearing one. Free text split on ';' produces a
// distinct part for almost every row; a real multi-select reuses a small
// fixed menu.
const MULTI_MIN_SEPARATED_RATIO = 0.6;
const MULTI_MIN_PARTS_PER_VALUE = 1.2;
const MULTI_MAX_DISTINCT_RATIO = 0.5;
const MULTI_MAX_DISTINCT_PARTS = 60;

function detectMultiSeparator(nonNull) {
  const strings = nonNull.filter((v) => !(v instanceof Date)).map(normalizeToString);
  if (strings.length === 0) return null;

  for (const separator of MULTI_SEPARATORS) {
    let separated = 0;
    let parts = 0;
    const distinct = new Set();

    for (const s of strings) {
      if (s.includes(separator)) separated++;
      for (const part of s.split(separator)) {
        const trimmed = part.trim();
        if (trimmed === '') continue;
        parts++;
        distinct.add(trimmed.toLowerCase());
      }
    }

    if (parts === 0) continue;
    if (separated / strings.length < MULTI_MIN_SEPARATED_RATIO) continue;
    if (parts / strings.length < MULTI_MIN_PARTS_PER_VALUE) continue;
    if (distinct.size > MULTI_MAX_DISTINCT_PARTS) continue;
    if (distinct.size / parts > MULTI_MAX_DISTINCT_RATIO) continue;

    return separator;
  }

  return null;
}
```

Then inside `inferType`, in the `else` branch that currently chooses between categorical and text, replace its opening so multi is tried first:

```js
  let verdict;
  if (numericVerdict) {
    verdict = numericVerdict;
  } else {
    // --- multi-select --------------------------------------------------
    const separator = detectMultiSeparator(nonNull);
    if (separator) {
      verdict = {
        type: 'multi',
        role: 'dimension',
        confidence: 1,
        dateOrder: null,
        isPercent: false,
        separator,
        nullCount,
        distinctCount,
        casualties: [],
        casualtyCount: 0,
      };
    } else {
    // --- categorical vs text -------------------------------------------
    const isCategorical = distinctCount <= 50 || distinctCount / nonNullCount < 0.05;
    verdict = isCategorical
      ? { /* ...unchanged categorical verdict... */ }
      : { /* ...unchanged text verdict... */ };
    }
  }
```

(Keep the existing categorical and text verdict object literals exactly as they are; only the surrounding `if (separator) { ... } else { ... }` is new.)

The identifier override below must not fire on multi — it already lists only `numeric`/`text`/`categorical`, so leave it untouched.

In `src/features/datastudio/profile/profileColumn.js`, add to `ROLE_BY_TYPE`:

```js
  multi: 'dimension',
```

- [ ] **Step 4: Run to verify they pass**

```bash
npm test -- inferType
```

Expected: PASS, and every pre-existing test in that file still passes.

- [ ] **Step 5: Commit**

```bash
git add src/features/datastudio/profile/inferType.js src/features/datastudio/profile/profileColumn.js src/features/datastudio/profile/inferType.test.js && git commit -m "Read semicolon-joined multi-select answers as their own type"
```

---

### Task 2: Rank a multi column by its options

**Files:**
- Modify: `src/features/datastudio/profile/profileColumn.js`
- Test: `src/features/datastudio/profile/profileDataset.test.js`

**Interfaces:**
- Consumes: `inferType` returning `{ type: 'multi', separator }` (Task 1).
- Produces: `profileColumn(...)` on a multi column returns `topValues` counting individual options.

- [ ] **Step 1: Write the failing test**

Append to `src/features/datastudio/profile/profileDataset.test.js`:

```js
import { profileColumn } from './profileColumn.js';

describe('profileColumn on a multi column', () => {
  it('counts options, not whole answers', () => {
    const values = [
      'Data Collection;Report Generation;',
      'Data Collection;Approval Tracking;',
      'Data Collection;',
      'Report Generation;Approval Tracking;',
      'Data Collection;Report Generation;',
      'Approval Tracking;',
    ];
    const column = profileColumn(values, 'Challenges', 0);

    expect(column.type).toBe('multi');
    expect(column.topValues[0]).toEqual({ value: 'Data Collection', count: 4 });
    // Whole-answer counting would have made every row its own value.
    expect(column.topValues.some((t) => t.value.includes(';'))).toBe(false);
  });
});
```

- [ ] **Step 2: Run to verify it fails**

```bash
npm test -- profileDataset
```

Expected: FAIL — `topValues[0]` is the whole joined string with a count of 1.

- [ ] **Step 3: Implement**

In `src/features/datastudio/profile/profileColumn.js`, add beside `rankTopValues`:

```js
// A multi column's frequent values are its OPTIONS, not its cells. A cell
// reading "A;B;C" is three answers; counting it whole makes every
// respondent their own category and the filter picker useless.
function rankTopOptions(values, separator) {
  const counts = new Map();
  for (const v of values) {
    if (isNullish(v)) continue;
    for (const part of String(v).split(separator)) {
      const label = part.trim();
      if (label === '') continue;
      counts.set(label, (counts.get(label) ?? 0) + 1);
    }
  }
  return [...counts.entries()]
    .map(([value, count], firstSeen) => ({ value, count, firstSeen }))
    .sort((a, b) => b.count - a.count || a.firstSeen - b.firstSeen)
    .slice(0, TOP_VALUE_LIMIT)
    .map(({ value, count }) => ({ value, count }));
}
```

In the returned object, replace the `topValues` line:

```js
    topValues: verdict.type === 'multi'
      ? rankTopOptions(values, verdict.separator ?? ';')
      : (isDimensionLike ? rankTopValues(values) : []),
```

- [ ] **Step 4: Run to verify it passes**

```bash
npm test -- profileDataset
```

Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/features/datastudio/profile/profileColumn.js src/features/datastudio/profile/profileDataset.test.js && git commit -m "Profile a multi column by its options"
```

---

### Task 3: Encode multi columns in the columnar store

**Files:**
- Modify: `src/features/datastudio/engine/dataset.js`
- Test: `src/features/datastudio/engine/dataset.test.js` (create)

**Interfaces:**
- Consumes: profile verdicts of `{ type: 'multi', separator }` (Task 1).
- Produces: a built column `{ type: 'multi', values: Int32Array, offsets: Int32Array, dictionary: string[] }` where row `r`'s option codes are `values[offsets[r] .. offsets[r+1])`. `dataset.rowCount` is now derived from the raw input length, not from `built[0].values.length`.

- [ ] **Step 1: Write the failing test**

Create `src/features/datastudio/engine/dataset.test.js`:

```js
import { describe, it, expect } from 'vitest';
import { buildDataset } from './dataset.js';

const profileOf = (columns) => ({ columns });

describe('buildDataset with a multi column', () => {
  const grid = {
    headers: ['Dept', 'Challenges'],
    columns: [
      ['IT', 'Finance', 'Logistics'],
      ['A;B;', 'B;', ''],
    ],
    profile: profileOf([
      { name: 'Dept', type: 'categorical', role: 'dimension' },
      { name: 'Challenges', type: 'multi', role: 'dimension', separator: ';' },
    ]),
  };

  it('stores option codes with row offsets', () => {
    const dataset = buildDataset(grid);
    const column = dataset.columns[1];

    expect(column.type).toBe('multi');
    expect(column.dictionary).toEqual(['A', 'B']);
    expect(Array.from(column.values)).toEqual([0, 1, 1]);
    expect(Array.from(column.offsets)).toEqual([0, 2, 3, 3]);
  });

  it('keeps rowCount the number of rows, not the number of options', () => {
    // The flat option array is longer than the grid. Deriving rowCount
    // from the first column's values would report 3 here by luck and 5
    // if the multi column happened to come first.
    const dataset = buildDataset({
      ...grid,
      headers: ['Challenges', 'Dept'],
      columns: [grid.columns[1], grid.columns[0]],
      profile: profileOf([grid.profile.columns[1], grid.profile.columns[0]]),
    });
    expect(dataset.rowCount).toBe(3);
  });
});
```

- [ ] **Step 2: Run to verify it fails**

```bash
npm test -- engine/dataset
```

Expected: FAIL — the multi column is encoded as text, and `rowCount` is 3 only by accident.

- [ ] **Step 3: Implement**

In `src/features/datastudio/engine/dataset.js`, add after `encodeCategorical`:

```js
// Multi-select: one row holds several options, so a flat code array plus
// per-row offsets (compressed-sparse-row) rather than one code per row.
// Row r's options are values[offsets[r] .. offsets[r + 1]). Both arrays
// are typed, so this survives structured clone and IndexedDB untouched
// the same way every other column does.
//
// A row with no options is offsets[r] === offsets[r + 1] -- an empty
// range, which is the null encoding for this type. There is no sentinel
// code, because there is no single slot to put one in.
function encodeMulti(values, separator = ';') {
  const dictionary = [];
  const codes = new Map();
  const flat = [];
  const offsets = new Int32Array(values.length + 1);

  for (let i = 0; i < values.length; i++) {
    offsets[i] = flat.length;
    const v = values[i];
    if (isMissing(v)) continue;
    for (const part of String(v).split(separator)) {
      const label = part.trim();
      if (label === '') continue;
      let code = codes.get(label);
      if (code === undefined) {
        code = dictionary.length;
        dictionary.push(label);
        codes.set(label, code);
      }
      flat.push(code);
    }
  }
  offsets[values.length] = flat.length;

  return { values: Int32Array.from(flat), offsets, dictionary };
}
```

Add the type set beside the others:

```js
const MULTI_TYPES = new Set(['multi']);
```

In `buildDataset`, add `let offsets = null;` beside `let dictionary = null;`, add the branch **before** the categorical branch:

```js
    } else if (MULTI_TYPES.has(type)) {
      ({ values, offsets, dictionary } = encodeMulti(raw, meta.separator ?? ';'));
    } else if (CATEGORICAL_TYPES.has(type)) {
```

Add `offsets` to the returned column object, and fix the row count:

```js
    return {
      name,
      type,
      role,
      values,
      dictionary,
      offsets,
      isPercent: Boolean(meta.isPercent),
      dateOnly: type === 'date',
      sourceZone: meta.sourceZone ?? 'local',
    };
  });

  // Derived from the RAW input, not from the first built column. A multi
  // column's `values` is the flat option array and is longer than the
  // grid, so reading the length off it reports the wrong row count for
  // the whole dataset -- and every mask allocated from it would be the
  // wrong size.
  const rowCount = headers.length > 0 ? (columns[0]?.length ?? 0) : 0;

  return {
    rowCount,
    columns: built,
    byName: new Map(built.map((c, i) => [c.name, i])),
  };
```

- [ ] **Step 4: Run to verify it passes**

```bash
npm test -- engine
```

Expected: PASS, including the existing aggregate and filterMask suites.

- [ ] **Step 5: Commit**

```bash
git add src/features/datastudio/engine/dataset.js src/features/datastudio/engine/dataset.test.js && git commit -m "Encode multi-select columns as options with row offsets"
```

---

### Task 4: Chart and filter multi columns

**Files:**
- Modify: `src/features/datastudio/engine/aggregate.js`
- Modify: `src/features/datastudio/engine/filterMask.js`
- Modify: `src/features/datastudio/clean/cleanOps.js`
- Modify: `src/features/datastudio/clean/proposeCleanPlan.js`
- Modify: `src/pages/DataStudioPage.jsx`
- Test: `src/features/datastudio/engine/aggregate.test.js`, `src/features/datastudio/engine/filterMask.test.js`, `src/features/datastudio/clean/cleanOps.test.js`

**Interfaces:**
- Consumes: multi columns from Task 3.
- Produces: `aggregate` counts a row once per option; `buildMask` keeps a row when any option matches; `castType(values, { type: 'multi', separator })` normalises spacing and drops empty options.

- [ ] **Step 1: Write the failing tests**

Append to `src/features/datastudio/engine/aggregate.test.js`:

```js
describe('aggregate over a multi column', () => {
  const dataset = buildDataset({
    headers: ['Challenges'],
    columns: [['A;B;', 'B;', 'A;B;C;', '']],
    profile: { columns: [{ name: 'Challenges', type: 'multi', role: 'dimension', separator: ';' }] },
  });

  it('counts each option a row picked', () => {
    const result = aggregate(dataset, null, {
      encoding: { x: { column: 'Challenges' }, y: [{ column: null, agg: 'count' }] },
      sort: { by: 'y', dir: 'desc' },
    });
    const counts = Object.fromEntries(
      result.categories.map((c, i) => [c, result.series[0].data[i]]),
    );
    expect(counts).toEqual({ B: 3, A: 2, C: 1 });
  });
});
```

Append to `src/features/datastudio/engine/filterMask.test.js`:

```js
describe('filtering a multi column', () => {
  const dataset = buildDataset({
    headers: ['Challenges'],
    columns: [['A;B;', 'B;', 'C;', '']],
    profile: { columns: [{ name: 'Challenges', type: 'multi', role: 'dimension', separator: ';' }] },
  });

  it('keeps a row when any of its options match', () => {
    const mask = buildMask(dataset, [{ column: 'Challenges', kind: 'in', values: ['A'] }]);
    expect(Array.from(mask)).toEqual([1, 0, 0, 0]);
  });

  it('drops a row with no options at all', () => {
    const mask = buildMask(dataset, [{ column: 'Challenges', kind: 'in', values: ['A', 'B', 'C'] }]);
    expect(Array.from(mask)).toEqual([1, 1, 1, 0]);
  });
});
```

Append to `src/features/datastudio/clean/cleanOps.test.js`:

```js
describe('castType to multi', () => {
  it('normalises spacing and drops empty options', () => {
    const out = castType(['A ; B ;', ';;', 'C'], { type: 'multi', separator: ';' });
    expect(out).toEqual(['A;B', null, 'C']);
  });
});
```

- [ ] **Step 2: Run to verify they fail**

```bash
npm test -- engine clean/cleanOps
```

Expected: FAIL — aggregate reports one category `"A;B;"`, the mask keeps nothing, `castType` returns the string untouched.

- [ ] **Step 3: Implement**

In `src/features/datastudio/engine/aggregate.js`, add to `makeXResolver` **before** the `if (column.dictionary)` branch:

```js
  // A multi column puts one row in several categories, so the resolver
  // returns an array. The grouping loop below normalises to an array
  // either way, which is why every other resolver can keep returning a
  // single object.
  if (column.type === 'multi') {
    return (row) => {
      const start = column.offsets[row];
      const end = column.offsets[row + 1];
      if (end <= start) return null;
      const out = [];
      for (let j = start; j < end; j++) {
        const code = column.values[j];
        out.push({ key: code, label: column.dictionary[code] });
      }
      return out;
    };
  }
```

In `makeSeriesResolver`, add at the top after the `if (!column) return null;`:

```js
  // Deliberately unsupported as a series. A row belonging to several
  // series at once would be counted once per series and every stacked
  // total would exceed the row count -- silently.
  if (column.type === 'multi') return null;
```

Replace the `seriesName` handling in `aggregate` so a null series resolver falls back rather than dropping every row:

```js
export function aggregate(dataset, mask, spec) {
  const measure = spec?.encoding?.y?.[0] ?? {};
  const agg = measure.agg ?? 'count';
  const measureColumn = columnOf(dataset, measure.column);

  const resolveX = makeXResolver(dataset, spec?.encoding?.x, mask);
  const resolveSeries = makeSeriesResolver(dataset, spec?.encoding?.series);
  // Used whenever no series resolver exists -- because none was asked
  // for, because the column is missing, or because it is a multi column.
  // Deriving this from the SPEC instead would name a series nothing can
  // resolve and drop every row.
  const fallbackSeriesName = measure.column ?? 'Count';

  const empty = { categories: [], series: [{ name: fallbackSeriesName, data: [] }] };
  if (!resolveX) return empty;
```

Delete the old `const seriesName = ...` line. In the row loop, replace the body from `const x = resolveX(row);` down to the end of the bucket accumulation with:

```js
    const x = resolveX(row);
    if (x === null) continue;

    // One entry for an ordinary column, several for a multi column.
    for (const one of (Array.isArray(x) ? x : [x])) {
      let group = groups.get(one.key);
      if (!group) {
        group = { label: one.label, sortKey: one.key, bySeries: new Map(), total: 0 };
        groups.set(one.key, group);
      }

      const name = resolveSeries ? resolveSeries(row) : fallbackSeriesName;
      if (name === null) continue;
      if (!seenSeries.has(name)) {
        seenSeries.add(name);
        seriesNames.push(name);
      }

      let bucket = group.bySeries.get(name);
      if (!bucket) {
        bucket = newBucket();
        group.bySeries.set(name, bucket);
      }

      bucket.count++;
      if (measureColumn) {
        const v = measureColumn.values[row];
        if (typeof v === 'number' && !Number.isNaN(v)) {
          bucket.sum += v;
          bucket.values.push(v);
        }
      }
    }
```

And at the bottom, replace `[seriesName ?? 'Count']` with `[fallbackSeriesName]`.

In `src/features/datastudio/engine/filterMask.js`, inside `applyFilter`, add **before** the `if (column.dictionary)` branch:

```js
  // Multi columns also carry a dictionary, so this has to come first.
  // A row is kept when ANY of its options was asked for -- "show me
  // everyone who mentioned approvals" must not exclude the people who
  // mentioned approvals and four other things.
  if (column.type === 'multi') {
    const codes = new Set();
    for (const label of wanted) {
      const code = column.dictionary.indexOf(label);
      if (code !== -1) codes.add(code);
    }
    for (let i = 0; i < column.offsets.length - 1; i++) {
      let keep = false;
      for (let j = column.offsets[i]; j < column.offsets[i + 1]; j++) {
        if (codes.has(values[j])) { keep = true; break; }
      }
      if (!keep) mask[i] = 0;
    }
    return;
  }
```

In `src/features/datastudio/clean/cleanOps.js`, add a `multi` case to `castType`:

```js
    case 'multi': {
      const separator = params.separator ?? ';';
      return values.map((v) => {
        if (isNullish(v)) return null;
        const options = String(v)
          .split(separator)
          .map((part) => part.trim())
          .filter(Boolean);
        // A cell of nothing but separators held no options at all, so it
        // is empty -- not an option named "".
        return options.length > 0 ? options.join(separator) : null;
      });
    }
```

In `src/features/datastudio/clean/proposeCleanPlan.js`, extend `proposeCast` to accept multi:

```js
  if (!['numeric', 'boolean', 'date', 'datetime', 'multi'].includes(type)) return null;
```

and, in the same function, treat a multi value as needing work only when normalising would change it:

```js
function alreadyTyped(value, type, column) {
  if (type === 'numeric') return typeof value === 'number';
  if (type === 'boolean') return typeof value === 'boolean';
  if (type === 'date' || type === 'datetime') return value instanceof Date;
  if (type === 'multi') {
    const separator = column?.separator ?? ';';
    const options = String(value).split(separator).map((p) => p.trim()).filter(Boolean);
    return options.join(separator) === String(value);
  }
  return true;
}
```

Update its two call sites in `proposeCast` to `alreadyTyped(v, type, column)`, and add the separator to the params:

```js
  const params = { type };
  if (type === 'multi') params.separator = column.separator ?? ';';
  if (type === 'date' || type === 'datetime') {
```

In `src/pages/DataStudioPage.jsx`, add `'multi'` to `TYPE_OPTIONS`:

```js
const TYPE_OPTIONS = ['numeric', 'categorical', 'multi', 'boolean', 'date', 'datetime', 'text', 'identifier'];
```

- [ ] **Step 4: Run to verify they pass**

```bash
npm test && npm run lint
```

Expected: all tests PASS. Lint reports only the four pre-existing files.

- [ ] **Step 5: Commit**

```bash
git add -A src/features/datastudio/engine src/features/datastudio/clean src/pages/DataStudioPage.jsx && git commit -m "Chart and filter multi-select columns by option"
```

---

### Task 5: Verify Part A against the real survey

**Files:**
- Create: `src/features/datastudio/engine/multiValue.integration.test.js`

**Interfaces:**
- Consumes: everything from Tasks 1–4.
- Produces: nothing — this is the proof that the multi-value path works end to end on real shapes.

- [ ] **Step 1: Write the test**

```js
import { describe, it, expect } from 'vitest';
import { profileDataset } from '../profile/profileDataset.js';
import { proposeCleanPlan } from '../clean/proposeCleanPlan.js';
import { applyCleanPlan } from '../clean/applyCleanPlan.js';
import { aggregate } from './aggregate.js';

// The shape of the real survey export: a multi-select question whose
// answers are semicolon-joined with a trailing separator.
const HEADERS = ['Department', 'Which challenges', 'Describe'];
const ROWS = [
  ['IT', 'no issue;', 'no issue from IT '],
  ['Finance', 'Data Collection;Data Consolidation;Report Generation;', 'Financial data is collected from many files.'],
  ['Logistics', 'Data Collection;Approval Tracking;', 'Updates arrive on WhatsApp and get missed.'],
  ['Finance', 'Data Collection;Report Generation;', 'Reports are rebuilt by hand each month.'],
  ['Sales', 'Approval Tracking;', 'Approvals sit with nobody chasing them.'],
  ['QAQC', 'Data Collection;Data Consolidation;', 'Numbers are retyped between two systems.'],
];

describe('the multi-select path, end to end', () => {
  it('ranks the options a survey offered', () => {
    const grid = { headers: HEADERS, rows: ROWS };
    const profile = profileDataset(grid);

    const challenges = profile.columns.find((c) => c.name === 'Which challenges');
    expect(challenges.type).toBe('multi');

    const plan = proposeCleanPlan(profile, grid);
    const dataset = applyCleanPlan(grid, plan, profile);

    const result = aggregate(dataset, null, {
      encoding: { x: { column: 'Which challenges' }, y: [{ column: null, agg: 'count' }] },
      sort: { by: 'y', dir: 'desc' },
      limit: 10,
    });

    expect(result.categories[0]).toBe('Data Collection');
    expect(result.series[0].data[0]).toBe(4);
    // Six respondents, but more than six option-picks: that is the whole
    // point of the type.
    expect(result.series[0].data.reduce((a, b) => a + b, 0)).toBeGreaterThan(ROWS.length);
  });
});
```

- [ ] **Step 2: Run it**

```bash
npm test -- multiValue.integration
```

Expected: PASS. If the profile does not report `multi`, the thresholds in Task 1 need revisiting against this fixture — fix them there, not here.

- [ ] **Step 3: Commit**

```bash
git add src/features/datastudio/engine/multiValue.integration.test.js && git commit -m "Prove the multi-select path on a real survey shape"
```

---

# PART B — Text analysis

### Task 6: Boilerplate and non-answers

**Files:**
- Create: `src/features/datastudio/text/boilerplate.js`
- Test: `src/features/datastudio/text/boilerplate.test.js`

**Interfaces:**
- Produces:
  - `stripLabelPrefix(line: string) -> string`
  - `isNonAnswer(fragment: string) -> boolean`
  - `normalizeText(value: unknown) -> string`

- [ ] **Step 1: Write the failing test**

Create `src/features/datastudio/text/boilerplate.test.js`:

```js
import { describe, it, expect } from 'vitest';
import { stripLabelPrefix, isNonAnswer, normalizeText } from './boilerplate.js';

describe('stripLabelPrefix', () => {
  it('removes the labels respondents copy out of the question', () => {
    expect(stripLabelPrefix('[Selected Challenge]: Data Collection')).toBe('Data Collection');
    expect(stripLabelPrefix('[Detailed Description]:')).toBe('');
    expect(stripLabelPrefix('Description: we retype everything')).toBe('we retype everything');
    // The real export is missing the opening bracket on some rows.
    expect(stripLabelPrefix('Selected Challenge]: Data Collection')).toBe('Data Collection');
  });

  it('leaves an ordinary sentence alone', () => {
    expect(stripLabelPrefix('The problem: nobody owns the report'))
      .toBe('The problem: nobody owns the report');
  });
});

describe('isNonAnswer', () => {
  it('drops the ways people say nothing is wrong', () => {
    expect(isNonAnswer('no issue from IT')).toBe(true);
    expect(isNonAnswer('N/A')).toBe(true);
    expect(isNonAnswer('-')).toBe(true);
    expect(isNonAnswer('none')).toBe(true);
    expect(isNonAnswer('')).toBe(true);
  });

  it('keeps a real complaint that starts with "no"', () => {
    // This pair is the whole rule. A prefix match alone deletes it.
    expect(isNonAnswer('No proper system exists for tracking approvals')).toBe(false);
    expect(isNonAnswer('Nothing is documented, so every handover starts over')).toBe(false);
  });
});

describe('normalizeText', () => {
  it('strips zero-width characters and collapses whitespace', () => {
    expect(normalizeText('a\u200Bb   c\uFEFF')).toBe('ab c');
  });

  it('survives a non-string', () => {
    expect(normalizeText(null)).toBe('');
    expect(normalizeText(42)).toBe('42');
  });
});
```

- [ ] **Step 2: Run to verify it fails**

```bash
npm test -- text/boilerplate
```

Expected: FAIL — `Failed to resolve import "./boilerplate.js"`.

- [ ] **Step 3: Implement**

Create `src/features/datastudio/text/boilerplate.js`:

```js
// What a survey answer says that is not an answer -- spec §6.2.
//
// Two rules, and the second is the one that can quietly destroy data.
// "no issue from IT" is somebody saying nothing is wrong; "No proper
// system exists for tracking approvals" is somebody reporting a real
// problem that happens to start with the same word. A prefix match
// alone deletes the second, so the rule needs BOTH a leading
// non-answer word and a short body -- a real complaint is never four
// words long.

// Escape sequences, never literals: an invisible character in source
// does not survive being retyped or diffed and rots into a no-op.
const ZERO_WIDTH_RE = /[\u200B-\u200D\uFEFF]/g;

export function normalizeText(value) {
  return String(value ?? '')
    .normalize('NFKC')
    .replace(ZERO_WIDTH_RE, '')
    .replace(/\s+/g, ' ')
    .trim();
}

// Respondents paste the question into the answer box. The bracket is
// optional on both sides because the real export has rows missing one.
const LABEL_RE = /^\s*\[?\s*(selected\s+challenge|detailed\s+description|challenge|description|issue|problem)s?\s*\]?\s*:\s*/i;

export function stripLabelPrefix(line) {
  return normalizeText(line).replace(LABEL_RE, '').trim();
}

const NON_ANSWER_WORDS = new Set([
  'no', 'none', 'nil', 'na', 'nothing', 'nope', 'n/a', '-', '–', '—', '.',
]);

// Short enough that it cannot be a report of anything. Twenty letters is
// "no issue from IT" (13) with room to spare, and well under the
// shortest real complaint in the source data.
const NON_ANSWER_MAX_LETTERS = 20;

export function isNonAnswer(fragment) {
  const text = normalizeText(fragment);
  if (text === '') return true;

  const letters = text.replace(/[^A-Za-z]/g, '').length;
  if (letters === 0) return true;

  const first = text.toLowerCase().split(/[\s,.;:!?]+/)[0];
  return NON_ANSWER_WORDS.has(first) && letters <= NON_ANSWER_MAX_LETTERS;
}
```

- [ ] **Step 4: Run to verify it passes**

```bash
npm test -- text/boilerplate
```

Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/features/datastudio/text/boilerplate.js src/features/datastudio/text/boilerplate.test.js && git commit -m "Tell a non-answer from a complaint that starts with no"
```

---

### Task 7: Split an answer into separate issues

**Files:**
- Create: `src/features/datastudio/text/splitIssues.js`
- Test: `src/features/datastudio/text/splitIssues.test.js`

**Interfaces:**
- Consumes: `stripLabelPrefix`, `isNonAnswer`, `normalizeText` from Task 6.
- Produces: `splitIssues(text: unknown) -> string[]`

- [ ] **Step 1: Write the failing test**

Create `src/features/datastudio/text/splitIssues.test.js`:

```js
import { describe, it, expect } from 'vitest';
import { splitIssues } from './splitIssues.js';

describe('splitIssues', () => {
  it('splits a paragraph into its separate complaints', () => {
    const text = 'Financial data is collected from multiple Excel files. '
      + 'The process involves extensive manual consolidation. '
      + 'Automating extraction would reduce turnaround time.';
    expect(splitIssues(text)).toHaveLength(3);
  });

  it('strips the labels respondents paste in', () => {
    const text = 'Selected Challenge]: Data Collection\n'
      + '[Detailed Description]:\n'
      + 'I collect information from multiple WhatsApp groups and Excel files.';
    const parts = splitIssues(text);
    expect(parts.some((p) => p.includes('Detailed Description'))).toBe(false);
    expect(parts.some((p) => p.includes('WhatsApp'))).toBe(true);
  });

  it('returns nothing for a non-answer', () => {
    expect(splitIssues('no issue from IT ')).toEqual([]);
    expect(splitIssues('')).toEqual([]);
    expect(splitIssues(null)).toEqual([]);
  });

  it('does not split on an abbreviation', () => {
    const text = 'We reconcile by hand, e.g. matching invoices to receipts, every month.';
    expect(splitIssues(text)).toHaveLength(1);
  });

  it('splits on bullets and newlines', () => {
    const text = '- Approvals are chased by email\n- Nobody knows the current status\n- Reports are rebuilt from scratch';
    expect(splitIssues(text)).toHaveLength(3);
  });

  it('caps a very long answer', () => {
    const text = Array.from({ length: 30 }, (_, i) => `Problem number ${i} wastes a lot of time here.`).join(' ');
    expect(splitIssues(text).length).toBeLessThanOrEqual(12);
  });
});
```

- [ ] **Step 2: Run to verify it fails**

```bash
npm test -- text/splitIssues
```

Expected: FAIL — module not found.

- [ ] **Step 3: Implement**

Create `src/features/datastudio/text/splitIssues.js`:

```js
// One written answer -> the separate issues inside it (spec §6.2).
//
// People do not answer "describe your challenge" with one challenge.
// They write a paragraph containing three, and counting the paragraph
// once undercounts every problem in it but the first.
//
// The splitting is rules-based on purpose. The model that follows can
// tell whether two fragments mean the same thing; it cannot tell where
// one thought ends -- and a sentence boundary is a perfectly good
// answer to that question in this data, where respondents write in
// clean sentences.

import { normalizeText, stripLabelPrefix, isNonAnswer } from './boilerplate.js';

// A fragment shorter than this is a split that went wrong -- an
// abbreviation, an initial, a stray "etc." -- and is folded back into
// the fragment before it rather than becoming an issue of its own.
const MIN_FRAGMENT_LENGTH = 25;

// A ceiling, not a target. A pathological answer must not produce a
// hundred rows the user has to read.
export const MAX_FRAGMENTS = 12;

const BULLET_RE = /^\s*(?:[-–—•*]|\d+[.)])\s+/;

// A sentence ends at .!? followed by space and something that starts a
// new sentence. The lookbehind excludes the common abbreviations that
// otherwise cut a sentence in half.
const SENTENCE_SPLIT_RE = /(?<!\b(?:e\.g|i\.e|etc|vs|no|dr|mr|ms|mrs)\.)(?<=[.!?])\s+(?=[A-Z0-9"'(])/;

export function splitIssues(text) {
  const normalized = normalizeText(String(text ?? '').replace(/\r\n?/g, '\n'));
  if (normalized === '') return [];

  // Newlines survive normalizeText as spaces, so split the raw value on
  // them first and normalise each line.
  const lines = String(text ?? '')
    .replace(/\r\n?/g, '\n')
    .split('\n')
    .map(stripLabelPrefix)
    .filter((line) => line !== '');

  const pieces = [];
  for (const line of lines) {
    const withoutBullet = line.replace(BULLET_RE, '').trim();
    if (withoutBullet === '') continue;
    for (const sentence of withoutBullet.split(SENTENCE_SPLIT_RE)) {
      const trimmed = sentence.trim();
      if (trimmed !== '') pieces.push(trimmed);
    }
  }

  // Fold a too-short piece into the one before it. Nothing to fold into
  // means it is the first piece, and it stands on its own.
  const merged = [];
  for (const piece of pieces) {
    if (piece.length < MIN_FRAGMENT_LENGTH && merged.length > 0) {
      merged[merged.length - 1] = `${merged[merged.length - 1]} ${piece}`;
      continue;
    }
    merged.push(piece);
  }

  return merged.filter((piece) => !isNonAnswer(piece)).slice(0, MAX_FRAGMENTS);
}
```

- [ ] **Step 4: Run to verify it passes**

```bash
npm test -- text/splitIssues
```

Expected: PASS. If the abbreviation test fails, the lookbehind is the thing to fix — do not relax `MIN_FRAGMENT_LENGTH` to paper over it.

- [ ] **Step 5: Commit**

```bash
git add src/features/datastudio/text/splitIssues.js src/features/datastudio/text/splitIssues.test.js && git commit -m "Split one written answer into the issues inside it"
```

---

### Task 8: Decide which columns are worth analysing

**Files:**
- Create: `src/features/datastudio/text/detectTextColumns.js`
- Test: `src/features/datastudio/text/detectTextColumns.test.js`

**Interfaces:**
- Consumes: a `profile` from `profileDataset`, and the raw `grid`.
- Produces: `detectTextColumns(profile, grid) -> [{ name, index, meanLength }]`, longest mean length first.

- [ ] **Step 1: Write the failing test**

Create `src/features/datastudio/text/detectTextColumns.test.js`:

```js
import { describe, it, expect } from 'vitest';
import { profileDataset } from '../profile/profileDataset.js';
import { detectTextColumns } from './detectTextColumns.js';

const grid = {
  headers: ['ID', 'Email', 'Department', 'Describe'],
  rows: [
    [1, 'a@x.com', 'IT', 'Financial data is collected from multiple Excel files before reporting.'],
    [2, 'b@x.com', 'Finance', 'The consolidation is manual, repetitive and prone to human error every month.'],
    [3, 'c@x.com', 'Logistics', 'Updates arrive over WhatsApp so important information is regularly missed.'],
    [4, 'd@x.com', 'Finance', 'Reports are rebuilt from scratch and version control is guesswork.'],
    [5, 'e@x.com', 'Sales', 'Approvals sit for days because nobody is told they are waiting.'],
    [6, 'f@x.com', 'QAQC', 'Numbers are retyped between two systems and the typos are only found later.'],
  ],
};

describe('detectTextColumns', () => {
  const found = detectTextColumns(profileDataset(grid), grid);

  it('picks the column people wrote in', () => {
    expect(found.map((c) => c.name)).toContain('Describe');
  });

  it('rejects identifiers, categories and short unique values', () => {
    expect(found.map((c) => c.name)).not.toContain('ID');
    expect(found.map((c) => c.name)).not.toContain('Department');
    // Emails are unique and text-typed but far too short to be prose.
    expect(found.map((c) => c.name)).not.toContain('Email');
  });

  it('returns nothing when a sheet has no prose in it', () => {
    const plain = { headers: ['Dept'], rows: [['IT'], ['Finance'], ['IT']] };
    expect(detectTextColumns(profileDataset(plain), plain)).toEqual([]);
  });
});
```

- [ ] **Step 2: Run to verify it fails**

```bash
npm test -- text/detectTextColumns
```

Expected: FAIL — module not found.

- [ ] **Step 3: Implement**

Create `src/features/datastudio/text/detectTextColumns.js`:

```js
// Which columns hold prose worth reading -- spec §6.1.
//
// The profiler already says a column is `text`, but that is necessary
// and not sufficient: an email address, a filename and a reference code
// are all text and near-unique too. What separates prose is length. A
// column of 70-character sentences is somebody writing; a column of
// 15-character values is somebody's identifier, whatever its type says.

import { normalizeText } from './boilerplate.js';

export const MIN_MEAN_LENGTH = 40;
export const MIN_FILLED_RATIO = 0.6;
export const MIN_DISTINCT_RATIO = 0.8;

export function detectTextColumns(profile, grid) {
  const rows = grid?.rows ?? [];
  const found = [];

  for (const column of profile?.columns ?? []) {
    if (column.type !== 'text') continue;

    const values = rows.map((row) => normalizeText(row?.[column.index]));
    const filled = values.filter((v) => v !== '');
    if (filled.length === 0) continue;
    if (filled.length / values.length < MIN_FILLED_RATIO) continue;

    const distinct = new Set(filled).size;
    if (distinct / filled.length < MIN_DISTINCT_RATIO) continue;

    const meanLength = filled.reduce((sum, v) => sum + v.length, 0) / filled.length;
    if (meanLength < MIN_MEAN_LENGTH) continue;

    found.push({ name: column.name, index: column.index, meanLength });
  }

  // Longest first: on a survey with two free-text questions, the one
  // people wrote most in is the one they were asked to describe.
  return found.sort((a, b) => b.meanLength - a.meanLength);
}
```

- [ ] **Step 4: Run to verify it passes**

```bash
npm test -- text/detectTextColumns
```

Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/features/datastudio/text/detectTextColumns.js src/features/datastudio/text/detectTextColumns.test.js && git commit -m "Find the columns people actually wrote in"
```

---

### Task 9: The starter buckets

**Files:**
- Create: `src/features/datastudio/text/buckets.js`
- Test: `src/features/datastudio/text/buckets.test.js`

**Interfaces:**
- Produces:
  - `STARTER_BUCKETS: [{ id, label, description, hints: string[] }]`
  - `UNSORTED_ID = 'unsorted'`
  - `bucketPromptText(bucket) -> string[]` — the strings that get embedded for a bucket

- [ ] **Step 1: Write the failing test**

Create `src/features/datastudio/text/buckets.test.js`:

```js
import { describe, it, expect } from 'vitest';
import { STARTER_BUCKETS, UNSORTED_ID, bucketPromptText } from './buckets.js';

describe('STARTER_BUCKETS', () => {
  it('gives every bucket a unique id', () => {
    const ids = STARTER_BUCKETS.map((b) => b.id);
    expect(new Set(ids).size).toBe(ids.length);
  });

  it('never ships a bucket called unsorted', () => {
    // Unsorted is where the model declines to guess. A real bucket with
    // that id would make a refusal indistinguishable from a match.
    expect(STARTER_BUCKETS.some((b) => b.id === UNSORTED_ID)).toBe(false);
  });

  it('describes every bucket in a sentence, not a label', () => {
    for (const bucket of STARTER_BUCKETS) {
      expect(bucket.description.length).toBeGreaterThan(30);
      expect(bucket.hints.length).toBeGreaterThan(0);
    }
  });
});

describe('bucketPromptText', () => {
  it('embeds the description and the hints, not the name', () => {
    const bucket = { id: 'x', label: 'SAP', description: 'Problems with SAP transactions.', hints: ['posting errors'] };
    expect(bucketPromptText(bucket)).toEqual(['Problems with SAP transactions.', 'posting errors']);
  });

  it('falls back to the label when someone clears the description', () => {
    expect(bucketPromptText({ id: 'x', label: 'SAP', description: '', hints: [] })).toEqual(['SAP']);
  });
});
```

- [ ] **Step 2: Run to verify it fails**

```bash
npm test -- text/buckets
```

Expected: FAIL — module not found.

- [ ] **Step 3: Implement**

Create `src/features/datastudio/text/buckets.js`:

```js
// The categories a survey answer gets filed into -- spec §7.
//
// A bucket is matched by its DESCRIPTION, not its name. "SAP" as a
// two-letter string is almost pure noise to a sentence model; "Problems
// with SAP transactions, ERP modules, master data or postings" sits
// next to the answers that belong in it. That is why the editor puts
// the description first and why renaming a bucket changes nothing about
// what lands in it.
//
// `hints` are extra phrasings, averaged in with the description. They
// exist for the cases a single sentence cannot cover -- a bucket that
// legitimately spans "the VPN drops" and "the office wifi is slow".

export const UNSORTED_ID = 'unsorted';
export const UNSORTED_LABEL = 'Unsorted';

export const STARTER_BUCKETS = [
  {
    id: 'sap',
    label: 'SAP / ERP',
    description: 'Problems working in SAP or another ERP system: transactions, modules, master data, postings and system limitations.',
    hints: ['SAP transaction is slow', 'ERP master data is wrong', 'the module cannot do what we need'],
  },
  {
    id: 'consolidation',
    label: 'Data Consolidation & Reporting',
    description: 'Combining data from several files, systems or subsidiaries into one report, and preparing recurring reports or dashboards.',
    hints: ['consolidating multiple Excel files', 'preparing the monthly management report', 'building the same dashboard again'],
  },
  {
    id: 'entry',
    label: 'Manual Data Entry',
    description: 'Retyping, copying and pasting information between systems, and transcribing from paper or documents by hand.',
    hints: ['retyping numbers into another system', 'copy and paste between spreadsheets', 'keying in data from a printout'],
  },
  {
    id: 'approvals',
    label: 'Approvals & Workflow',
    description: 'Chasing sign-off, tracking the status of a request, and sending reminders to move work along.',
    hints: ['following up on approvals', 'nobody knows the current status', 'sending reminders to get sign-off'],
  },
  {
    id: 'forms',
    label: 'Forms & Paperwork',
    description: 'Paper forms, physical signatures, hardcopy documents and routing them between people for completion.',
    hints: ['the form has to be printed and signed', 'passing hardcopy around the office', 'filling in the same form twice'],
  },
  {
    id: 'retrieval',
    label: 'Information Retrieval',
    description: 'Searching for files, records, emails or historical information, and not being able to find the latest version.',
    hints: ['hunting for the right file', 'searching old emails for a record', 'nobody knows which version is current'],
  },
  {
    id: 'communication',
    label: 'Communication & Coordination',
    description: 'Information arriving through chat groups or email instead of a system, handovers between people, and updates getting missed.',
    hints: ['updates come through WhatsApp groups', 'important messages get overlooked', 'the handover loses information'],
  },
  {
    id: 'network',
    label: 'Network & Internet',
    description: 'Connectivity problems: internet speed, VPN, remote access, shared drives and the network dropping.',
    hints: ['the internet is slow', 'the VPN keeps disconnecting', 'cannot reach the shared drive from home'],
  },
  {
    id: 'itsupport',
    label: 'IT Support & Systems',
    description: 'Hardware faults, slow computers, software problems, accounts and access rights, and waiting for IT to fix something.',
    hints: ['my laptop is very slow', 'I do not have access to the system', 'waiting for IT to respond'],
  },
  {
    id: 'digitization',
    label: 'Digitization & Automation',
    description: 'Asking for a manual process to be replaced by a system, automated, or made digital end to end.',
    hints: ['this should be automated', 'we need a proper system instead of spreadsheets', 'move the whole process online'],
  },
  {
    id: 'ai',
    label: 'AI Opportunities',
    description: 'Explicit requests for artificial intelligence, machine learning or an intelligent assistant to help with the work.',
    hints: ['AI could read these documents', 'a chatbot could answer these questions', 'machine learning to predict demand'],
  },
  {
    id: 'training',
    label: 'Training & Knowledge',
    description: 'Not knowing how to do something, undocumented processes, and knowledge that lives only in one person.',
    hints: ['nobody documented the process', 'I was never trained on this', 'only one person knows how'],
  },
];

// What actually gets embedded for a bucket. The label is deliberately
// excluded unless nothing else is left -- see the note at the top.
export function bucketPromptText(bucket) {
  const parts = [];
  const description = String(bucket?.description ?? '').trim();
  if (description !== '') parts.push(description);
  for (const hint of bucket?.hints ?? []) {
    const trimmed = String(hint ?? '').trim();
    if (trimmed !== '') parts.push(trimmed);
  }
  if (parts.length === 0) parts.push(String(bucket?.label ?? '').trim());
  return parts;
}
```

- [ ] **Step 4: Run to verify it passes**

```bash
npm test -- text/buckets
```

Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/features/datastudio/text/buckets.js src/features/datastudio/text/buckets.test.js && git commit -m "Ship a starter bucket set described in sentences"
```

---

### Task 10: Score how severe an issue sounds

**Files:**
- Create: `src/features/datastudio/text/severity.js`
- Test: `src/features/datastudio/text/severity.test.js`

**Interfaces:**
- Consumes: `normalizeText` from Task 6.
- Produces:
  - `INTENSITY_TERMS: string[]`
  - `severityOf(text: string, { meanLength: number, breadth: number }) -> number` in 0–1, where `breadth` is 0–1 (how many multi-select options that respondent picked, normalised against the most anyone picked).

- [ ] **Step 1: Write the failing test**

Create `src/features/datastudio/text/severity.test.js`:

```js
import { describe, it, expect } from 'vitest';
import { severityOf } from './severity.js';

const context = { meanLength: 80, breadth: 0 };

describe('severityOf', () => {
  it('scores a strongly worded issue above a mild one', () => {
    const mild = severityOf('We update the sheet each week.', context);
    const strong = severityOf(
      'The manual reconciliation is time-consuming, repetitive and prone to error, and deadlines are constantly missed.',
      context,
    );
    expect(strong).toBeGreaterThan(mild);
  });

  it('rises with each intensity term, then saturates', () => {
    const one = severityOf('The process is manual.', context);
    const three = severityOf('The process is manual, repetitive and time-consuming.', context);
    const many = severityOf(
      'The manual, repetitive, time-consuming, tedious, error-prone rework causes constant delays and duplicate effort.',
      context,
    );
    expect(three).toBeGreaterThan(one);
    expect(many).toBeGreaterThanOrEqual(three);
  });

  it('counts how many challenges the respondent picked', () => {
    const text = 'Reports take a long time to prepare.';
    expect(severityOf(text, { meanLength: 80, breadth: 1 }))
      .toBeGreaterThan(severityOf(text, { meanLength: 80, breadth: 0 }));
  });

  it('stays inside 0 and 1 whatever it is given', () => {
    const extreme = severityOf(
      'MANUAL!!! repetitive time-consuming tedious duplicate rework chase constantly unable cannot difficult delay missed overlooked bottleneck!!!'.repeat(5),
      { meanLength: 10, breadth: 1 },
    );
    expect(extreme).toBeLessThanOrEqual(1);
    expect(severityOf('', context)).toBeGreaterThanOrEqual(0);
  });
});
```

- [ ] **Step 2: Run to verify it fails**

```bash
npm test -- text/severity
```

Expected: FAIL — module not found.

- [ ] **Step 3: Implement**

Create `src/features/datastudio/text/severity.js`:

```js
// How strongly an issue is expressed -- spec §6.7.
//
// This is a SIGNAL, not a judgement, and the UI says so. Nothing here
// knows whether a problem is important; it measures four things that
// correlate with someone being frustrated enough to write about it, and
// the priority ranking then weighs that behind how many people raised it
// at all.
//
// The breadth term is the only one not inferred from prose: it comes
// from the structured multi-select column. Somebody who ticked seven
// challenges is telling us something the wording cannot.

import { normalizeText } from './boilerplate.js';

export const INTENSITY_TERMS = [
  'time-consuming', 'time consuming', 'prone to error', 'error-prone', 'error prone',
  'manual', 'manually', 'repetitive', 'repeatedly', 'delay', 'delays', 'delayed',
  'missed', 'miss', 'overlooked', 'bottleneck', 'tedious', 'frustrating', 'frustration',
  'duplicate', 'duplication', 'rework', 'chase', 'chasing', 'constantly', 'always',
  'difficult', 'cannot', 'unable', 'no way to', 'waste', 'wasted', 'slow', 'stuck',
];

const WEIGHT_INTENSITY = 0.5;
const WEIGHT_LENGTH = 0.2;
const WEIGHT_BREADTH = 0.2;
const WEIGHT_EMPHASIS = 0.1;

// Four matches is as angry as this measure gets. Without a ceiling a
// single long answer that repeats itself outscores four different
// people, which is the ordering the ranking exists to prevent.
const INTENSITY_SATURATION = 4;

function countIntensity(lower) {
  let matches = 0;
  for (const term of INTENSITY_TERMS) {
    if (lower.includes(term)) matches++;
    if (matches >= INTENSITY_SATURATION) break;
  }
  return matches;
}

function emphasisOf(text) {
  const bangs = (text.match(/!/g) ?? []).length;
  const shouted = (text.match(/\b[A-Z]{3,}\b/g) ?? []).length;
  return Math.min(1, (bangs + shouted) / 3);
}

export function severityOf(text, { meanLength = 80, breadth = 0 } = {}) {
  const normalized = normalizeText(text);
  if (normalized === '') return 0;

  const lower = normalized.toLowerCase();

  const intensity = countIntensity(lower) / INTENSITY_SATURATION;
  // Relative to the corpus, capped: twice the average length is as much
  // as length is allowed to say.
  const length = Math.min(1, normalized.length / (Math.max(1, meanLength) * 2));
  const spread = Math.min(1, Math.max(0, breadth));
  const emphasis = emphasisOf(normalized);

  const score = intensity * WEIGHT_INTENSITY
    + length * WEIGHT_LENGTH
    + spread * WEIGHT_BREADTH
    + emphasis * WEIGHT_EMPHASIS;

  return Math.min(1, Math.max(0, score));
}
```

- [ ] **Step 4: Run to verify it passes**

```bash
npm test -- text/severity
```

Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/features/datastudio/text/severity.js src/features/datastudio/text/severity.test.js && git commit -m "Measure how strongly an issue is expressed"
```

---

### Task 11: Match a fragment to a bucket

**Files:**
- Create: `src/features/datastudio/text/similarity.js`
- Test: `src/features/datastudio/text/similarity.test.js`

**Interfaces:**
- Consumes: `UNSORTED_ID` from Task 9.
- Produces:
  - `cosine(a: Float32Array, b: Float32Array) -> number`
  - `meanVector(vectors: Float32Array[]) -> Float32Array` (L2-normalised)
  - `DEFAULT_THRESHOLD = 0.3`
  - `assignBuckets(fragmentVectors, bucketVectors: [{ id, vector }], threshold) -> [{ bucketId, score }]`

- [ ] **Step 1: Write the failing test**

Create `src/features/datastudio/text/similarity.test.js`:

```js
import { describe, it, expect } from 'vitest';
import { cosine, meanVector, assignBuckets, DEFAULT_THRESHOLD } from './similarity.js';
import { UNSORTED_ID } from './buckets.js';

const v = (...xs) => Float32Array.from(xs);

describe('cosine', () => {
  it('is 1 for the same direction and 0 for a right angle', () => {
    expect(cosine(v(1, 0), v(1, 0))).toBeCloseTo(1);
    expect(cosine(v(1, 0), v(0, 1))).toBeCloseTo(0);
  });

  it('is 0 rather than NaN for a zero vector', () => {
    expect(cosine(v(0, 0), v(1, 0))).toBe(0);
  });
});

describe('meanVector', () => {
  it('averages and re-normalises', () => {
    const mean = meanVector([v(1, 0), v(0, 1)]);
    expect(Math.hypot(mean[0], mean[1])).toBeCloseTo(1);
    expect(mean[0]).toBeCloseTo(mean[1]);
  });
});

describe('assignBuckets', () => {
  const buckets = [{ id: 'a', vector: v(1, 0) }, { id: 'b', vector: v(0, 1) }];

  it('picks the closest bucket', () => {
    const out = assignBuckets([v(0.9, 0.1), v(0.1, 0.9)], buckets, DEFAULT_THRESHOLD);
    expect(out.map((r) => r.bucketId)).toEqual(['a', 'b']);
  });

  it('declines rather than forcing a poor match', () => {
    // Equidistant and far from both: a confident wrong answer is worse
    // than an honest gap.
    const out = assignBuckets([v(0.7, 0.7)], buckets, 0.9);
    expect(out[0].bucketId).toBe(UNSORTED_ID);
  });

  it('treats the threshold as inclusive at the boundary', () => {
    const out = assignBuckets([v(1, 0)], buckets, 1);
    expect(out[0].bucketId).toBe('a');
  });

  it('returns Unsorted when there are no buckets at all', () => {
    expect(assignBuckets([v(1, 0)], [], DEFAULT_THRESHOLD)[0].bucketId).toBe(UNSORTED_ID);
  });
});
```

- [ ] **Step 2: Run to verify it fails**

```bash
npm test -- text/similarity
```

Expected: FAIL — module not found.

- [ ] **Step 3: Implement**

Create `src/features/datastudio/text/similarity.js`:

```js
// Fragment -> bucket, or an honest refusal -- spec §6.4.
//
// The refusal is the point. Every fragment has a nearest bucket, and
// filing it there regardless produces a screen where everything is
// categorised and some of it is wrong -- with nothing marking which.
// Below the threshold a fragment goes to Unsorted, which is a visible
// pile the user can act on.

import { UNSORTED_ID } from './buckets.js';

// Calibrated against real survey text during implementation; exposed in
// the UI because no single value is right for every survey.
export const DEFAULT_THRESHOLD = 0.3;

export function cosine(a, b) {
  let dot = 0;
  let na = 0;
  let nb = 0;
  const n = Math.min(a.length, b.length);
  for (let i = 0; i < n; i++) {
    dot += a[i] * b[i];
    na += a[i] * a[i];
    nb += b[i] * b[i];
  }
  // A zero vector has no direction, so it is not similar to anything --
  // 0, never NaN, which would poison every comparison downstream.
  if (na === 0 || nb === 0) return 0;
  return dot / (Math.sqrt(na) * Math.sqrt(nb));
}

export function meanVector(vectors) {
  const list = vectors ?? [];
  if (list.length === 0) return Float32Array.from([]);

  const out = new Float32Array(list[0].length);
  for (const vector of list) {
    for (let i = 0; i < out.length; i++) out[i] += vector[i] ?? 0;
  }

  let norm = 0;
  for (let i = 0; i < out.length; i++) {
    out[i] /= list.length;
    norm += out[i] * out[i];
  }
  norm = Math.sqrt(norm);
  if (norm > 0) {
    for (let i = 0; i < out.length; i++) out[i] /= norm;
  }
  return out;
}

export function assignBuckets(fragmentVectors, bucketVectors, threshold = DEFAULT_THRESHOLD) {
  const buckets = bucketVectors ?? [];

  return (fragmentVectors ?? []).map((vector) => {
    let bestId = UNSORTED_ID;
    let bestScore = -Infinity;

    for (const bucket of buckets) {
      const score = cosine(vector, bucket.vector);
      if (score > bestScore) {
        bestScore = score;
        bestId = bucket.id;
      }
    }

    if (buckets.length === 0 || bestScore < threshold) {
      // The score is still reported for Unsorted rows -- it is what the
      // "lower the threshold" prompt reads to say how close they were.
      return { bucketId: UNSORTED_ID, score: buckets.length === 0 ? 0 : bestScore };
    }
    return { bucketId: bestId, score: bestScore };
  });
}
```

- [ ] **Step 4: Run to verify it passes**

```bash
npm test -- text/similarity
```

Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/features/datastudio/text/similarity.js src/features/datastudio/text/similarity.test.js && git commit -m "Match an issue to a bucket, or decline to guess"
```

---

### Task 12: Discover themes by clustering

**Files:**
- Create: `src/features/datastudio/text/cluster.js`
- Test: `src/features/datastudio/text/cluster.test.js`

**Interfaces:**
- Consumes: `cosine` from Task 11.
- Produces:
  - `DEFAULT_GRANULARITY = 0.45`, `MAX_CLUSTERABLE = 5000`
  - `clusterVectors(vectors, granularity) -> { clusters: number[][], oneOffs: number[] }` where each cluster is an array of indices into `vectors`, ordered largest first.
  - Throws `RangeError` above `MAX_CLUSTERABLE`.

- [ ] **Step 1: Write the failing test**

Create `src/features/datastudio/text/cluster.test.js`:

```js
import { describe, it, expect } from 'vitest';
import { clusterVectors, DEFAULT_GRANULARITY, MAX_CLUSTERABLE } from './cluster.js';

const v = (...xs) => Float32Array.from(xs);

describe('clusterVectors', () => {
  it('groups vectors that point the same way', () => {
    const vectors = [
      v(1, 0), v(0.99, 0.1), v(0.98, 0.05),
      v(0, 1), v(0.1, 0.99),
    ];
    const { clusters } = clusterVectors(vectors, DEFAULT_GRANULARITY);
    expect(clusters).toHaveLength(2);
    expect(clusters[0]).toHaveLength(3);
    expect(clusters[0]).toEqual(expect.arrayContaining([0, 1, 2]));
  });

  it('keeps a lone vector out of the themes', () => {
    // A theme of one is a quote, not a pattern.
    const vectors = [v(1, 0), v(0.99, 0.1), v(0, 1)];
    const { clusters, oneOffs } = clusterVectors(vectors, DEFAULT_GRANULARITY);
    expect(clusters).toHaveLength(1);
    expect(oneOffs).toEqual([2]);
  });

  it('makes fewer, broader themes as granularity rises', () => {
    const vectors = [v(1, 0), v(0.9, 0.44), v(0.44, 0.9), v(0, 1)];
    const narrow = clusterVectors(vectors, 0.1).clusters.length;
    const broad = clusterVectors(vectors, 0.9).clusters.length;
    expect(broad).toBeLessThanOrEqual(narrow);
  });

  it('returns nothing for an empty input', () => {
    expect(clusterVectors([], DEFAULT_GRANULARITY)).toEqual({ clusters: [], oneOffs: [] });
  });

  it('refuses rather than freezing on an enormous input', () => {
    const huge = Array.from({ length: MAX_CLUSTERABLE + 1 }, () => v(1, 0));
    expect(() => clusterVectors(huge, DEFAULT_GRANULARITY)).toThrow(RangeError);
  });
});
```

- [ ] **Step 2: Run to verify it fails**

```bash
npm test -- text/cluster
```

Expected: FAIL — module not found.

- [ ] **Step 3: Implement**

Create `src/features/datastudio/text/cluster.js`:

```js
// The themes that are in the text whether or not anyone listed them --
// spec §6.5.
//
// Agglomerative with average linkage: start with every fragment its own
// cluster and repeatedly merge the two closest, stopping when the
// closest pair is further apart than the granularity setting. Average
// linkage rather than single linkage because single linkage chains --
// A near B, B near C, C near D -- and produces one sprawling cluster
// that means nothing.
//
// O(n^2) is fine at the scale this runs (a survey, not a corpus) and
// the guard below makes the failure loud rather than a frozen tab.

import { cosine } from './similarity.js';

export const DEFAULT_GRANULARITY = 0.45;
export const MAX_CLUSTERABLE = 5000;

export function clusterVectors(vectors, granularity = DEFAULT_GRANULARITY) {
  const list = vectors ?? [];
  if (list.length === 0) return { clusters: [], oneOffs: [] };
  if (list.length > MAX_CLUSTERABLE) {
    throw new RangeError(
      `Too many responses to group (${list.length}); the limit is ${MAX_CLUSTERABLE}.`,
    );
  }

  // Distance, not similarity: 0 is identical, 1 is unrelated.
  const distance = (a, b) => 1 - cosine(a, b);

  let groups = list.map((_, i) => [i]);

  const linkage = (left, right) => {
    let total = 0;
    for (const i of left) {
      for (const j of right) total += distance(list[i], list[j]);
    }
    return total / (left.length * right.length);
  };

  for (;;) {
    let bestDistance = Infinity;
    let bestA = -1;
    let bestB = -1;

    for (let a = 0; a < groups.length; a++) {
      for (let b = a + 1; b < groups.length; b++) {
        const d = linkage(groups[a], groups[b]);
        if (d < bestDistance) {
          bestDistance = d;
          bestA = a;
          bestB = b;
        }
      }
    }

    if (bestA === -1 || bestDistance >= granularity) break;

    const merged = [...groups[bestA], ...groups[bestB]];
    groups = groups.filter((_, i) => i !== bestA && i !== bestB);
    groups.push(merged);
  }

  const clusters = groups
    .filter((g) => g.length >= 2)
    .map((g) => g.slice().sort((a, b) => a - b))
    .sort((a, b) => b.length - a.length || a[0] - b[0]);

  const oneOffs = groups
    .filter((g) => g.length < 2)
    .flat()
    .sort((a, b) => a - b);

  return { clusters, oneOffs };
}
```

- [ ] **Step 4: Run to verify it passes**

```bash
npm test -- text/cluster
```

Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/features/datastudio/text/cluster.js src/features/datastudio/text/cluster.test.js && git commit -m "Discover the themes sitting in the responses"
```

---

### Task 13: Name a theme from its own words

**Files:**
- Create: `src/features/datastudio/text/labelCluster.js`
- Test: `src/features/datastudio/text/labelCluster.test.js`

**Interfaces:**
- Consumes: `normalizeText` from Task 6.
- Produces: `labelCluster(memberTexts: string[], allTexts: string[], termCount = 4) -> string`

- [ ] **Step 1: Write the failing test**

Create `src/features/datastudio/text/labelCluster.test.js`:

```js
import { describe, it, expect } from 'vitest';
import { labelCluster } from './labelCluster.js';

describe('labelCluster', () => {
  const corpus = [
    'approval requests wait for days with no reminder',
    'chasing approval status by email every week',
    'approval sign-off has no reminder or tracking',
    'reports are rebuilt from scratch every month',
    'consolidating spreadsheets from five subsidiaries',
    'the monthly report takes three days to build',
  ];

  it('names a theme after what makes it different', () => {
    const name = labelCluster(corpus.slice(0, 3), corpus);
    expect(name).toContain('approval');
    expect(name).not.toContain('report');
  });

  it('ignores a word that is in every fragment', () => {
    const flat = ['data is slow', 'data is missing', 'data is wrong'];
    // "data" appears everywhere, so it distinguishes nothing.
    const name = labelCluster(flat.slice(0, 2), flat);
    expect(name.split(' · ')).not.toContain('data');
  });

  it('drops stopwords and short noise', () => {
    const name = labelCluster(corpus.slice(0, 3), corpus);
    for (const stop of ['the', 'for', 'with', 'and', 'has', 'no']) {
      expect(name.split(' · ')).not.toContain(stop);
    }
  });

  it('says so when there is nothing distinctive to say', () => {
    expect(labelCluster([], corpus)).toBe('Unnamed theme');
  });
});
```

- [ ] **Step 2: Run to verify it fails**

```bash
npm test -- text/labelCluster
```

Expected: FAIL — module not found.

- [ ] **Step 3: Implement**

Create `src/features/datastudio/text/labelCluster.js`:

```js
// What to call a theme nobody named -- spec §6.6.
//
// c-TF-IDF: a term earns its place by being frequent INSIDE the cluster
// and rare outside it. Plain frequency would name every theme after the
// survey's own subject -- "data · process · work" on all of them -- which
// tells the reader nothing about which theme they are looking at.
//
// The result is a starting point. Renaming a theme is a first-class
// action in the UI precisely because four words are a label, not a
// sentence.

import { normalizeText } from './boilerplate.js';

const STOPWORDS = new Set([
  'the', 'a', 'an', 'and', 'or', 'but', 'if', 'then', 'so', 'because', 'as',
  'of', 'in', 'on', 'at', 'to', 'for', 'with', 'from', 'by', 'into', 'about',
  'is', 'are', 'was', 'were', 'be', 'been', 'being', 'am',
  'do', 'does', 'did', 'have', 'has', 'had', 'having',
  'it', 'its', 'this', 'that', 'these', 'those', 'there', 'here',
  'we', 'our', 'us', 'i', 'my', 'me', 'you', 'your', 'they', 'them', 'their',
  'he', 'she', 'his', 'her', 'which', 'who', 'whom', 'what', 'when', 'where',
  'not', 'no', 'nor', 'can', 'cannot', 'could', 'will', 'would', 'should',
  'may', 'might', 'must', 'need', 'needs', 'very', 'more', 'most', 'much',
  'many', 'some', 'any', 'all', 'every', 'each', 'other', 'than', 'also',
  'up', 'out', 'down', 'over', 'under', 'again', 'only', 'just', 'still',
]);

const MIN_TERM_LENGTH = 3;

export const SEPARATOR = ' · ';
export const UNNAMED = 'Unnamed theme';

function tokenize(text) {
  return normalizeText(text)
    .toLowerCase()
    .split(/[^a-z0-9-]+/)
    .filter((token) => token.length >= MIN_TERM_LENGTH && !STOPWORDS.has(token));
}

export function labelCluster(memberTexts, allTexts, termCount = 4) {
  const members = memberTexts ?? [];
  if (members.length === 0) return UNNAMED;

  const corpus = allTexts ?? [];
  const total = Math.max(1, corpus.length);

  // How many fragments in the WHOLE corpus contain each term.
  const documentFrequency = new Map();
  for (const text of corpus) {
    for (const token of new Set(tokenize(text))) {
      documentFrequency.set(token, (documentFrequency.get(token) ?? 0) + 1);
    }
  }

  const termFrequency = new Map();
  for (const text of members) {
    for (const token of tokenize(text)) {
      termFrequency.set(token, (termFrequency.get(token) ?? 0) + 1);
    }
  }

  const scored = [];
  for (const [term, tf] of termFrequency) {
    const df = documentFrequency.get(term) ?? 1;
    // A term in every fragment scores log(1) = 0 and drops out, which is
    // exactly the "data · process · work" problem this exists to solve.
    const idf = Math.log(total / df);
    if (idf <= 0) continue;
    scored.push({ term, score: tf * idf });
  }

  if (scored.length === 0) return UNNAMED;

  scored.sort((a, b) => b.score - a.score || a.term.localeCompare(b.term));
  return scored.slice(0, termCount).map((s) => s.term).join(SEPARATOR);
}
```

- [ ] **Step 4: Run to verify it passes**

```bash
npm test -- text/labelCluster
```

Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/features/datastudio/text/labelCluster.js src/features/datastudio/text/labelCluster.test.js && git commit -m "Name a theme after what makes it different"
```

---

### Task 14: Rank issues into a priority order

**Files:**
- Create: `src/features/datastudio/text/rankIssues.js`
- Test: `src/features/datastudio/text/rankIssues.test.js`

**Interfaces:**
- Produces: `rankIssues(groups, { pinned = [], suppressed = [] }) -> ranked[]` where a group is `{ kind: 'bucket' | 'theme', id, label, respondents, count, meanSeverity }` and each ranked entry adds `{ score, pinned: boolean, suppressed: boolean }`.

- [ ] **Step 1: Write the failing test**

Create `src/features/datastudio/text/rankIssues.test.js`:

```js
import { describe, it, expect } from 'vitest';
import { rankIssues } from './rankIssues.js';

const group = (id, respondents, count, meanSeverity) => ({
  kind: 'bucket', id, label: id, respondents, count, meanSeverity,
});

describe('rankIssues', () => {
  it('puts five mild people above one furious person', () => {
    // The decision this whole ranking exists to encode.
    const ranked = rankIssues([
      group('lonely', 1, 5, 1),
      group('common', 5, 5, 0.1),
    ], {});
    expect(ranked[0].id).toBe('common');
  });

  it('breaks a tie on severity', () => {
    const ranked = rankIssues([
      group('calm', 4, 4, 0.1),
      group('angry', 4, 4, 0.8),
    ], {});
    expect(ranked[0].id).toBe('angry');
  });

  it('breaks a remaining tie alphabetically, so the order is stable', () => {
    const ranked = rankIssues([group('zebra', 2, 2, 0.5), group('apple', 2, 2, 0.5)], {});
    expect(ranked.map((r) => r.id)).toEqual(['apple', 'zebra']);
  });

  it('lifts pinned items to the top in pin order', () => {
    const ranked = rankIssues([
      group('big', 9, 9, 0.9),
      group('small', 1, 1, 0),
      group('mid', 4, 4, 0.4),
    ], { pinned: ['small', 'mid'] });
    expect(ranked.map((r) => r.id)).toEqual(['small', 'mid', 'big']);
    expect(ranked[0].pinned).toBe(true);
  });

  it('sinks suppressed items to the bottom without deleting them', () => {
    const ranked = rankIssues([
      group('hidden', 9, 9, 0.9),
      group('kept', 1, 1, 0),
    ], { suppressed: ['hidden'] });
    expect(ranked.map((r) => r.id)).toEqual(['kept', 'hidden']);
    expect(ranked[1].suppressed).toBe(true);
  });

  it('handles an empty list', () => {
    expect(rankIssues([], {})).toEqual([]);
  });
});
```

- [ ] **Step 2: Run to verify it fails**

```bash
npm test -- text/rankIssues
```

Expected: FAIL — module not found.

- [ ] **Step 3: Implement**

Create `src/features/datastudio/text/rankIssues.js`:

```js
// The order to take into a meeting -- spec §6.8.
//
// Distinct respondents leads, severity only scales it. One person
// writing five furious sentences must not outrank five people each
// writing one calm one: the first is an individual's bad week, the
// second is a process problem. Scoring on fragment count would get that
// exactly backwards, which is why `count` is carried for display and
// never enters the score.

export function rankIssues(groups, { pinned = [], suppressed = [] } = {}) {
  const pinOrder = new Map(pinned.map((id, i) => [id, i]));
  const suppressedSet = new Set(suppressed);

  const scored = (groups ?? []).map((group) => ({
    ...group,
    score: group.respondents * (1 + (group.meanSeverity ?? 0)),
    pinned: pinOrder.has(group.id),
    suppressed: suppressedSet.has(group.id),
  }));

  return scored.sort((a, b) => {
    // Suppressed sinks, pinned floats, and the two never collide because
    // a pinned item the user also suppressed is a contradiction the UI
    // does not offer.
    if (a.suppressed !== b.suppressed) return a.suppressed ? 1 : -1;
    if (a.pinned !== b.pinned) return a.pinned ? -1 : 1;
    if (a.pinned && b.pinned) return pinOrder.get(a.id) - pinOrder.get(b.id);

    if (b.score !== a.score) return b.score - a.score;
    if ((b.meanSeverity ?? 0) !== (a.meanSeverity ?? 0)) {
      return (b.meanSeverity ?? 0) - (a.meanSeverity ?? 0);
    }
    // Alphabetical last, so two identical groups do not swap places
    // between renders for no reason.
    return String(a.label).localeCompare(String(b.label));
  });
}
```

- [ ] **Step 4: Run to verify it passes**

```bash
npm test -- text/rankIssues
```

Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/features/datastudio/text/rankIssues.js src/features/datastudio/text/rankIssues.test.js && git commit -m "Rank issues by people first, intensity second"
```

---

### Task 15: Apply the user's corrections over the model's answer

**Files:**
- Create: `src/features/datastudio/text/overrides.js`
- Test: `src/features/datastudio/text/overrides.test.js`

**Interfaces:**
- Consumes: `rankIssues` (Task 14), `UNSORTED_ID`/`UNSORTED_LABEL` (Task 9).
- Produces:
  - `EMPTY_OVERRIDES` — the zero value
  - `applyOverrides(raw, overrides) -> analysis`

  A **raw analysis** is:
  ```js
  {
    columnName, settings: { threshold, granularity },
    buckets: [{ id, label, description, hints }],
    fragments: [{ id, row, index, text, severity, bucketId, score, themeId }],
    themes: [{ id, name, fragmentIds }],
    oneOffIds: string[], noIssueRows: number[],
  }
  ```
  An **overrides** record is:
  ```js
  { retags: {}, noise: [], themeNames: {}, themeMerges: {}, pinned: [], suppressed: [] }
  ```
  The returned **analysis** is the raw one with `fragments` carrying final `bucketId`/`themeId`/`noise`, plus `buckets` and `themes` bearing `{ count, respondents, meanSeverity }`, plus `priority`.

- [ ] **Step 1: Write the failing test**

Create `src/features/datastudio/text/overrides.test.js`:

```js
import { describe, it, expect } from 'vitest';
import { applyOverrides, EMPTY_OVERRIDES } from './overrides.js';
import { UNSORTED_ID } from './buckets.js';

const raw = {
  columnName: 'Describe',
  settings: { threshold: 0.3, granularity: 0.45 },
  buckets: [
    { id: 'sap', label: 'SAP / ERP', description: 'd', hints: [] },
    { id: 'approvals', label: 'Approvals & Workflow', description: 'd', hints: [] },
  ],
  fragments: [
    { id: '0:0', row: 0, index: 0, text: 'SAP posting fails', severity: 0.6, bucketId: 'sap', score: 0.5, themeId: 't1' },
    { id: '1:0', row: 1, index: 0, text: 'Approvals are chased', severity: 0.4, bucketId: 'approvals', score: 0.6, themeId: 't1' },
    { id: '1:1', row: 1, index: 1, text: 'Nobody knows the status', severity: 0.2, bucketId: UNSORTED_ID, score: 0.1, themeId: 't2' },
  ],
  themes: [
    { id: 't1', name: 'sap · approval', fragmentIds: ['0:0', '1:0'] },
    { id: 't2', name: 'status', fragmentIds: ['1:1'] },
  ],
  oneOffIds: [],
  noIssueRows: [2],
};

describe('applyOverrides', () => {
  it('counts people, not fragments', () => {
    const analysis = applyOverrides(raw, EMPTY_OVERRIDES);
    const approvals = analysis.buckets.find((b) => b.id === 'approvals');
    expect(approvals.count).toBe(1);
    const theme = analysis.themes.find((t) => t.id === 't1');
    expect(theme.count).toBe(2);
    expect(theme.respondents).toBe(2);
  });

  it('always offers Unsorted, even when nothing is in it', () => {
    const clean = { ...raw, fragments: raw.fragments.map((f) => ({ ...f, bucketId: 'sap' })) };
    const analysis = applyOverrides(clean, EMPTY_OVERRIDES);
    expect(analysis.buckets.some((b) => b.id === UNSORTED_ID)).toBe(true);
  });

  it('honours a hand retag', () => {
    const analysis = applyOverrides(raw, { ...EMPTY_OVERRIDES, retags: { '1:1': 'approvals' } });
    expect(analysis.fragments.find((f) => f.id === '1:1').bucketId).toBe('approvals');
    expect(analysis.buckets.find((b) => b.id === 'approvals').count).toBe(2);
  });

  it('excludes noise from every count but keeps the row visible', () => {
    const analysis = applyOverrides(raw, { ...EMPTY_OVERRIDES, noise: ['0:0'] });
    expect(analysis.fragments.find((f) => f.id === '0:0').noise).toBe(true);
    expect(analysis.buckets.find((b) => b.id === 'sap').count).toBe(0);
  });

  it('renames and merges themes', () => {
    const analysis = applyOverrides(raw, {
      ...EMPTY_OVERRIDES,
      themeNames: { t1: 'Approval chasing' },
      themeMerges: { t2: 't1' },
    });
    const merged = analysis.themes.find((t) => t.id === 't1');
    expect(merged.name).toBe('Approval chasing');
    expect(merged.count).toBe(3);
    expect(analysis.themes.some((t) => t.id === 't2')).toBe(false);
  });

  it('drops an override pointing at a fragment that no longer exists', () => {
    // A re-import with different data must not corrupt the screen.
    const analysis = applyOverrides(raw, { ...EMPTY_OVERRIDES, retags: { '99:9': 'sap' }, noise: ['99:9'] });
    expect(analysis.fragments).toHaveLength(3);
    expect(analysis.buckets.find((b) => b.id === 'sap').count).toBe(1);
  });

  it('survives a re-score: the retag is applied to the new raw result', () => {
    const overrides = { ...EMPTY_OVERRIDES, retags: { '1:1': 'sap' } };
    const rescored = { ...raw, fragments: raw.fragments.map((f) => ({ ...f, bucketId: 'approvals' })) };
    expect(applyOverrides(rescored, overrides).fragments.find((f) => f.id === '1:1').bucketId)
      .toBe('sap');
  });

  it('produces a priority order', () => {
    const analysis = applyOverrides(raw, EMPTY_OVERRIDES);
    expect(analysis.priority.length).toBeGreaterThan(0);
    expect(analysis.priority[0]).toHaveProperty('score');
  });
});
```

- [ ] **Step 2: Run to verify it fails**

```bash
npm test -- text/overrides
```

Expected: FAIL — module not found.

- [ ] **Step 3: Implement**

Create `src/features/datastudio/text/overrides.js`:

```js
// The user's corrections, applied over the model's answer -- spec §8.
//
// The raw analysis is never mutated. Corrections live in their own
// record and are re-applied to whatever the model most recently
// produced, which is what makes three things true at once:
//
//   * re-running the model, moving the threshold, or editing a bucket
//     description never destroys a correction;
//   * "reset to what the model said" is discarding one object;
//   * an override naming a fragment that no longer exists is dropped
//     rather than corrupting the screen -- a re-import with different
//     data is a normal thing to do.

import { rankIssues } from './rankIssues.js';
import { UNSORTED_ID, UNSORTED_LABEL } from './buckets.js';

export const EMPTY_OVERRIDES = {
  retags: {},
  noise: [],
  themeNames: {},
  themeMerges: {},
  pinned: [],
  suppressed: [],
};

// Follows a chain of merges to whichever theme is still standing, and
// refuses to loop if two themes were somehow merged into each other.
function resolveTheme(themeId, merges) {
  let current = themeId;
  const seen = new Set();
  while (merges[current] && !seen.has(current)) {
    seen.add(current);
    current = merges[current];
  }
  return current;
}

function summarise(fragmentIds, byId) {
  const respondents = new Set();
  let severityTotal = 0;
  let counted = 0;

  for (const id of fragmentIds) {
    const fragment = byId.get(id);
    if (!fragment || fragment.noise) continue;
    respondents.add(fragment.row);
    severityTotal += fragment.severity ?? 0;
    counted++;
  }

  return {
    count: counted,
    respondents: respondents.size,
    meanSeverity: counted > 0 ? severityTotal / counted : 0,
  };
}

export function applyOverrides(raw, overrides = EMPTY_OVERRIDES) {
  const {
    retags = {}, noise = [], themeNames = {}, themeMerges = {},
    pinned = [], suppressed = [],
  } = overrides ?? {};

  const noiseSet = new Set(noise);
  const bucketIds = new Set([...(raw?.buckets ?? []).map((b) => b.id), UNSORTED_ID]);

  const fragments = (raw?.fragments ?? []).map((fragment) => {
    // A retag naming a bucket that has since been deleted falls back to
    // the model's answer rather than to a bucket nothing can render.
    const retag = retags[fragment.id];
    const bucketId = retag && bucketIds.has(retag) ? retag : fragment.bucketId;
    return {
      ...fragment,
      bucketId,
      themeId: resolveTheme(fragment.themeId, themeMerges),
      noise: noiseSet.has(fragment.id),
    };
  });

  const byId = new Map(fragments.map((f) => [f.id, f]));

  const byBucket = new Map();
  for (const fragment of fragments) {
    if (!byBucket.has(fragment.bucketId)) byBucket.set(fragment.bucketId, []);
    byBucket.get(fragment.bucketId).push(fragment.id);
  }

  const definitions = [
    ...(raw?.buckets ?? []),
    // Always present, even when empty: the pile where the model declined
    // to guess is information, and an absent Unsorted reads as "nothing
    // was ambiguous".
    { id: UNSORTED_ID, label: UNSORTED_LABEL, description: '', hints: [] },
  ];

  const buckets = definitions.map((definition) => ({
    ...definition,
    fragmentIds: byBucket.get(definition.id) ?? [],
    ...summarise(byBucket.get(definition.id) ?? [], byId),
  }));

  const byTheme = new Map();
  for (const fragment of fragments) {
    if (!fragment.themeId) continue;
    if (!byTheme.has(fragment.themeId)) byTheme.set(fragment.themeId, []);
    byTheme.get(fragment.themeId).push(fragment.id);
  }

  const themes = (raw?.themes ?? [])
    // A theme merged into another one no longer exists on its own.
    .filter((theme) => !themeMerges[theme.id])
    .map((theme) => {
      const fragmentIds = byTheme.get(theme.id) ?? [];
      return {
        ...theme,
        name: themeNames[theme.id] ?? theme.name,
        fragmentIds,
        ...summarise(fragmentIds, byId),
      };
    });

  const priority = rankIssues(
    [
      ...buckets
        .filter((b) => b.id !== UNSORTED_ID && b.count > 0)
        .map((b) => ({ kind: 'bucket', id: b.id, label: b.label, respondents: b.respondents, count: b.count, meanSeverity: b.meanSeverity })),
      ...themes
        .filter((t) => t.count > 0)
        .map((t) => ({ kind: 'theme', id: t.id, label: t.name, respondents: t.respondents, count: t.count, meanSeverity: t.meanSeverity })),
    ],
    { pinned, suppressed },
  );

  return { ...raw, fragments, buckets, themes, priority };
}
```

- [ ] **Step 4: Run to verify it passes**

```bash
npm test -- text/overrides
```

Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/features/datastudio/text/overrides.js src/features/datastudio/text/overrides.test.js && git commit -m "Keep hand corrections separate from the model's answer"
```

---

### Task 16: Turn the analysis into dataset columns

**Files:**
- Create: `src/features/datastudio/text/deriveColumns.js`
- Test: `src/features/datastudio/text/deriveColumns.test.js`

**Interfaces:**
- Consumes: an analysis from `applyOverrides` (Task 15).
- Produces: `deriveColumns(analysis, rowCount) -> { headers: string[], columns: unknown[][] }` — five columns, each exactly `rowCount` long, ready to append to a raw grid. Header names are exported as `DERIVED_HEADERS`.

- [ ] **Step 1: Write the failing test**

Create `src/features/datastudio/text/deriveColumns.test.js`:

```js
import { describe, it, expect } from 'vitest';
import { deriveColumns, DERIVED_HEADERS, NO_ISSUE_LABEL } from './deriveColumns.js';

const analysis = {
  buckets: [
    { id: 'sap', label: 'SAP / ERP' },
    { id: 'approvals', label: 'Approvals & Workflow' },
  ],
  themes: [{ id: 't1', name: 'approval · chase' }],
  fragments: [
    { id: '0:0', row: 0, text: 'a', severity: 0.6, bucketId: 'sap', themeId: 't1', noise: false },
    { id: '0:1', row: 0, text: 'b', severity: 0.2, bucketId: 'approvals', themeId: 't1', noise: false },
    { id: '1:0', row: 1, text: 'c', severity: 0.9, bucketId: 'approvals', themeId: 't1', noise: true },
  ],
};

describe('deriveColumns', () => {
  const { headers, columns } = deriveColumns(analysis, 3);
  const col = (name) => columns[headers.indexOf(name)];

  it('emits one value per row for every column', () => {
    expect(headers).toEqual(DERIVED_HEADERS);
    for (const column of columns) expect(column).toHaveLength(3);
  });

  it('gives a row the category of its worst issue', () => {
    expect(col('Issue category')[0]).toBe('SAP / ERP');
  });

  it('lists every category a row raised, semicolon-joined', () => {
    // Deliberately the multi-value shape, so the chart canvas can count
    // it by option instead of by combination.
    expect(col('Issue categories')[0]).toBe('SAP / ERP;Approvals & Workflow');
  });

  it('leaves a row with no issues clearly marked', () => {
    expect(col('Issue category')[2]).toBe(NO_ISSUE_LABEL);
    expect(col('Issue count')[2]).toBe(0);
    expect(col('Severity')[2]).toBe(0);
  });

  it('ignores fragments marked as noise', () => {
    // Row 1's only fragment is noise, so the row has no issues at all.
    expect(col('Issue category')[1]).toBe(NO_ISSUE_LABEL);
    expect(col('Issue count')[1]).toBe(0);
  });

  it('reports severity as a whole number out of a hundred', () => {
    expect(col('Severity')[0]).toBe(60);
  });
});
```

- [ ] **Step 2: Run to verify it fails**

```bash
npm test -- text/deriveColumns
```

Expected: FAIL — module not found.

- [ ] **Step 3: Implement**

Create `src/features/datastudio/text/deriveColumns.js`:

```js
// The analysis, as ordinary spreadsheet columns -- spec §9.
//
// This is the payoff of the whole design. The results are appended to
// the raw grid as five more columns and everything downstream -- the
// profiler, the cleaner, the chart canvas, the tile editor, the filter
// bar, cross-filtering, saved dashboards, PNG export -- consumes them
// with no change whatsoever. The analysis adds data; it does not add a
// second charting system.
//
// `Issue categories` is deliberately semicolon-joined rather than one
// column per bucket. That makes it a multi column (see the profiler),
// so a chart counts it by OPTION, and one respondent who raised three
// kinds of problem is counted under all three instead of becoming their
// own private category.

export const NO_ISSUE_LABEL = 'No issue raised';
export const MULTI_SEPARATOR = ';';

export const DERIVED_HEADERS = [
  'Issue category',
  'Issue categories',
  'Theme',
  'Issue count',
  'Severity',
];

export function deriveColumns(analysis, rowCount) {
  const bucketLabel = new Map((analysis?.buckets ?? []).map((b) => [b.id, b.label]));
  const themeName = new Map((analysis?.themes ?? []).map((t) => [t.id, t.name]));

  const byRow = new Map();
  for (const fragment of analysis?.fragments ?? []) {
    if (fragment.noise) continue;
    if (!byRow.has(fragment.row)) byRow.set(fragment.row, []);
    byRow.get(fragment.row).push(fragment);
  }

  const primary = [];
  const all = [];
  const theme = [];
  const counts = [];
  const severities = [];

  for (let row = 0; row < rowCount; row++) {
    const fragments = byRow.get(row) ?? [];

    if (fragments.length === 0) {
      primary.push(NO_ISSUE_LABEL);
      all.push(NO_ISSUE_LABEL);
      theme.push(NO_ISSUE_LABEL);
      counts.push(0);
      severities.push(0);
      continue;
    }

    // The worst one speaks for the row. Picking the first would make the
    // headline category depend on the order somebody wrote their
    // sentences in.
    let worst = fragments[0];
    for (const fragment of fragments) {
      if ((fragment.severity ?? 0) > (worst.severity ?? 0)) worst = fragment;
    }

    const labels = [];
    for (const fragment of fragments) {
      const label = bucketLabel.get(fragment.bucketId);
      if (label && !labels.includes(label)) labels.push(label);
    }

    primary.push(bucketLabel.get(worst.bucketId) ?? NO_ISSUE_LABEL);
    all.push(labels.join(MULTI_SEPARATOR));
    theme.push(themeName.get(worst.themeId) ?? NO_ISSUE_LABEL);
    counts.push(fragments.length);
    // Whole numbers out of a hundred: a 0-1 float axis reads as a
    // proportion of something, which severity is not.
    severities.push(Math.round((worst.severity ?? 0) * 100));
  }

  return {
    headers: DERIVED_HEADERS,
    columns: [primary, all, theme, counts, severities],
  };
}
```

- [ ] **Step 4: Run to verify it passes**

```bash
npm test -- text/deriveColumns
```

Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/features/datastudio/text/deriveColumns.js src/features/datastudio/text/deriveColumns.test.js && git commit -m "Write the analysis back as ordinary chartable columns"
```

---

### Task 17: Orchestrate the analysis

**Files:**
- Create: `src/features/datastudio/text/analysis.js`
- Test: `src/features/datastudio/text/analysis.test.js`

**Interfaces:**
- Consumes: `splitIssues` (7), `bucketPromptText`/`STARTER_BUCKETS` (9), `severityOf` (10), `assignBuckets`/`meanVector`/`DEFAULT_THRESHOLD` (11), `clusterVectors`/`DEFAULT_GRANULARITY` (12), `labelCluster` (13).
- Produces:
  - `MIN_FRAGMENTS_FOR_THEMES = 5`
  - `buildFragments(texts, breadths) -> [{ id, row, index, text, severity }]`
  - `async analyze({ texts, breadths, buckets, settings, embedAll, onProgress }) -> rawAnalysis`
  - `async rescore({ fragments, vectors, buckets, settings, embedAll }) -> rawAnalysis`

  **The embedder is injected, never imported.** That is what lets every test here run without a model, and it is the seam the worker fills with the real one.

- [ ] **Step 1: Write the failing test**

Create `src/features/datastudio/text/analysis.test.js`:

```js
import { describe, it, expect, vi } from 'vitest';
import { analyze, buildFragments, MIN_FRAGMENTS_FOR_THEMES } from './analysis.js';
import { UNSORTED_ID } from './buckets.js';

// A deterministic stand-in for the model: two dimensions, one per
// keyword. Nothing here loads or needs the real thing.
const fakeEmbed = vi.fn(async (texts) => texts.map((text) => {
  const lower = text.toLowerCase();
  const approval = /approv|sign-off|status|remind|chase/.test(lower) ? 1 : 0;
  const sap = /sap|erp|posting|master data/.test(lower) ? 1 : 0;
  return Float32Array.from([approval, sap]);
}));

const buckets = [
  { id: 'approvals', label: 'Approvals', description: 'approval sign-off status reminder chase', hints: [] },
  { id: 'sap', label: 'SAP', description: 'sap erp posting master data', hints: [] },
];

describe('buildFragments', () => {
  it('gives every fragment a stable id of row and position', () => {
    const fragments = buildFragments(
      ['One problem here that is long enough. And a second one, also long enough.'],
      [0],
    );
    expect(fragments.map((f) => f.id)).toEqual(['0:0', '0:1']);
  });

  it('skips a row that raised nothing', () => {
    expect(buildFragments(['no issue from IT', ''], [0, 0])).toEqual([]);
  });
});

describe('analyze', () => {
  const texts = [
    'Approvals sit for days and nobody sends a reminder about the status.',
    'We chase sign-off by email and the status is never visible to anyone.',
    'The SAP posting fails and master data has to be corrected by hand.',
    'ERP master data is wrong so every SAP posting needs manual repair.',
    'Chasing approval status wastes a whole afternoon every single week.',
    'no issue from IT',
  ];
  const breadths = [0.2, 0.4, 0.6, 0.2, 0.8, 0];

  it('files fragments into the buckets they belong to', async () => {
    const raw = await analyze({ texts, breadths, buckets, embedAll: fakeEmbed });
    const approvals = raw.fragments.filter((f) => f.bucketId === 'approvals');
    const sap = raw.fragments.filter((f) => f.bucketId === 'sap');
    expect(approvals.length).toBeGreaterThan(0);
    expect(sap.length).toBeGreaterThan(0);
  });

  it('records which rows raised nothing at all', async () => {
    const raw = await analyze({ texts, breadths, buckets, embedAll: fakeEmbed });
    expect(raw.noIssueRows).toContain(5);
  });

  it('discovers themes and names them', async () => {
    const raw = await analyze({ texts, breadths, buckets, embedAll: fakeEmbed });
    expect(raw.themes.length).toBeGreaterThan(0);
    for (const theme of raw.themes) expect(theme.name).not.toBe('');
  });

  it('skips theme discovery when there is almost nothing to group', async () => {
    const raw = await analyze({
      texts: ['Approvals sit for days and nobody sends a reminder.'],
      breadths: [0], buckets, embedAll: fakeEmbed,
    });
    expect(raw.fragments.length).toBeLessThan(MIN_FRAGMENTS_FOR_THEMES);
    expect(raw.themes).toEqual([]);
  });

  it('sends everything to Unsorted when no bucket matches', async () => {
    const raw = await analyze({
      texts, breadths, buckets, settings: { threshold: 0.99 }, embedAll: fakeEmbed,
    });
    expect(raw.fragments.every((f) => f.bucketId === UNSORTED_ID)).toBe(true);
  });

  it('reports progress as it goes', async () => {
    const onProgress = vi.fn();
    await analyze({ texts, breadths, buckets, embedAll: fakeEmbed, onProgress });
    expect(onProgress).toHaveBeenCalled();
    const stages = onProgress.mock.calls.map(([p]) => p.stage);
    expect(new Set(stages).size).toBeGreaterThan(1);
  });

  it('embeds the bucket descriptions, not the bucket names', async () => {
    fakeEmbed.mockClear();
    await analyze({ texts, breadths, buckets, embedAll: fakeEmbed });
    const embedded = fakeEmbed.mock.calls.flatMap(([list]) => list);
    expect(embedded).toContain('approval sign-off status reminder chase');
    expect(embedded).not.toContain('Approvals');
  });

  it('returns an empty result rather than throwing on an empty sheet', async () => {
    const raw = await analyze({ texts: [], breadths: [], buckets, embedAll: fakeEmbed });
    expect(raw.fragments).toEqual([]);
    expect(raw.themes).toEqual([]);
  });
});
```

- [ ] **Step 2: Run to verify it fails**

```bash
npm test -- text/analysis
```

Expected: FAIL — module not found.

- [ ] **Step 3: Implement**

Create `src/features/datastudio/text/analysis.js`:

```js
// The whole pipeline, in order -- spec §6.
//
// The embedder arrives as an argument and is never imported here. That
// single choice is what keeps this file pure enough to test: every case
// below runs against a two-dimensional stand-in, with no model, no
// WebAssembly and no network. The worker is the only place that
// supplies the real one.

import { splitIssues } from './splitIssues.js';
import { bucketPromptText, STARTER_BUCKETS } from './buckets.js';
import { severityOf } from './severity.js';
import { assignBuckets, meanVector, DEFAULT_THRESHOLD } from './similarity.js';
import { clusterVectors, DEFAULT_GRANULARITY } from './cluster.js';
import { labelCluster } from './labelCluster.js';
import { normalizeText } from './boilerplate.js';

// Below this, grouping is theatre. Two sentences do not have themes,
// and presenting a "theme" over four fragments invites the reader to
// draw a conclusion the data cannot support.
export const MIN_FRAGMENTS_FOR_THEMES = 5;

export function buildFragments(texts, breadths = []) {
  const rows = texts ?? [];

  // The mean is computed over the fragments, not the raw answers, so the
  // length term in severity compares like with like.
  const perRow = rows.map((text) => splitIssues(text));
  const flat = perRow.flat();
  const meanLength = flat.length > 0
    ? flat.reduce((sum, t) => sum + t.length, 0) / flat.length
    : 80;

  const fragments = [];
  for (let row = 0; row < perRow.length; row++) {
    perRow[row].forEach((text, index) => {
      fragments.push({
        id: `${row}:${index}`,
        row,
        index,
        text,
        severity: severityOf(text, { meanLength, breadth: breadths[row] ?? 0 }),
      });
    });
  }
  return fragments;
}

function noIssueRowsOf(texts, fragments) {
  const withIssues = new Set(fragments.map((f) => f.row));
  const rows = [];
  for (let row = 0; row < (texts ?? []).length; row++) {
    if (!withIssues.has(row)) rows.push(row);
  }
  return rows;
}

// One averaged vector per bucket, built from its description and hints.
async function embedBuckets(buckets, embedAll) {
  const prompts = buckets.map(bucketPromptText);
  const flat = prompts.flat();
  if (flat.length === 0) return [];

  const vectors = await embedAll(flat);

  let cursor = 0;
  return buckets.map((bucket, i) => {
    const slice = vectors.slice(cursor, cursor + prompts[i].length);
    cursor += prompts[i].length;
    return { id: bucket.id, vector: meanVector(slice) };
  });
}

function discoverThemes(fragments, vectors, granularity) {
  if (fragments.length < MIN_FRAGMENTS_FOR_THEMES) {
    return { themes: [], oneOffIds: fragments.map((f) => f.id), themeByFragment: new Map() };
  }

  const { clusters, oneOffs } = clusterVectors(vectors, granularity);
  const allTexts = fragments.map((f) => f.text);
  const themeByFragment = new Map();

  const themes = clusters.map((members, i) => {
    const id = `theme_${i}`;
    const fragmentIds = members.map((index) => fragments[index].id);
    for (const fragmentId of fragmentIds) themeByFragment.set(fragmentId, id);
    return {
      id,
      name: labelCluster(members.map((index) => allTexts[index]), allTexts),
      fragmentIds,
    };
  });

  return { themes, oneOffIds: oneOffs.map((index) => fragments[index].id), themeByFragment };
}

function assemble({ columnName, buckets, fragments, assignments, themes, oneOffIds, themeByFragment, noIssueRows, settings }) {
  return {
    columnName,
    settings,
    buckets,
    fragments: fragments.map((fragment, i) => ({
      ...fragment,
      bucketId: assignments[i].bucketId,
      score: assignments[i].score,
      themeId: themeByFragment.get(fragment.id) ?? null,
    })),
    themes,
    oneOffIds,
    noIssueRows,
  };
}

export async function analyze({
  texts, breadths = [], buckets = STARTER_BUCKETS, columnName = '',
  settings = {}, embedAll, onProgress = () => {},
}) {
  const threshold = settings.threshold ?? DEFAULT_THRESHOLD;
  const granularity = settings.granularity ?? DEFAULT_GRANULARITY;
  const resolved = { threshold, granularity };

  onProgress({ stage: 'Reading responses', pct: 42 });
  const fragments = buildFragments(texts, breadths);
  const noIssueRows = noIssueRowsOf(texts, fragments);

  if (fragments.length === 0) {
    return {
      columnName, settings: resolved, buckets,
      fragments: [], themes: [], oneOffIds: [], noIssueRows,
      vectors: [],
    };
  }

  onProgress({ stage: 'Understanding responses', pct: 45 });
  const vectors = await embedAll(
    fragments.map((f) => normalizeText(f.text)),
    { onProgress },
  );

  const bucketVectors = await embedBuckets(buckets, embedAll);
  const assignments = assignBuckets(vectors, bucketVectors, threshold);

  onProgress({ stage: 'Grouping', pct: 85 });
  const { themes, oneOffIds, themeByFragment } = discoverThemes(fragments, vectors, granularity);

  onProgress({ stage: 'Ranking', pct: 95 });
  return {
    ...assemble({
      columnName, buckets, fragments, assignments,
      themes, oneOffIds, themeByFragment, noIssueRows, settings: resolved,
    }),
    // Kept so a threshold or granularity change never re-embeds. This is
    // what the sub-second settings budget in spec §16 rests on.
    vectors,
  };
}

/**
 * Re-file and re-group WITHOUT re-embedding the fragments.
 *
 * Only the bucket descriptions are embedded here -- a dozen short
 * strings. Anything that re-embeds fragments on a settings change is a
 * regression against the performance budget, not an optimisation
 * opportunity missed.
 */
export async function rescore({
  columnName = '', fragments, vectors, buckets = STARTER_BUCKETS,
  settings = {}, noIssueRows = [], embedAll,
}) {
  const threshold = settings.threshold ?? DEFAULT_THRESHOLD;
  const granularity = settings.granularity ?? DEFAULT_GRANULARITY;
  const resolved = { threshold, granularity };

  if ((fragments ?? []).length === 0) {
    return {
      columnName, settings: resolved, buckets,
      fragments: [], themes: [], oneOffIds: [], noIssueRows, vectors: [],
    };
  }

  const bucketVectors = await embedBuckets(buckets, embedAll);
  const assignments = assignBuckets(vectors, bucketVectors, threshold);
  const { themes, oneOffIds, themeByFragment } = discoverThemes(fragments, vectors, granularity);

  return {
    ...assemble({
      columnName, buckets, fragments, assignments,
      themes, oneOffIds, themeByFragment, noIssueRows, settings: resolved,
    }),
    vectors,
  };
}
```

- [ ] **Step 4: Run to verify it passes**

```bash
npm test -- text/analysis
```

Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/features/datastudio/text/analysis.js src/features/datastudio/text/analysis.test.js && git commit -m "Run the text pipeline end to end with the model injected"
```

---

### Task 18: The local model

**Files:**
- Create: `scripts/fetch-model.mjs`
- Create: `src/features/datastudio/text/embed.js`
- Modify: `package.json`
- Modify: `.gitignore`
- Modify: `vercel.json`

**Interfaces:**
- Produces: `createEmbedder() -> { embedAll(texts, { onProgress }) -> Promise<Float32Array[]> }`, and `public/models/` + `public/ort/` populated by `npm run fetch:model`.

**This is the task that keeps the spec §2 promise.** Two settings do it: `localModelPath` for the model, `wasmPaths` for the ONNX runtime. Miss the second and the runtime is fetched from a public CDN at first use, silently.

- [ ] **Step 1: Add the dependency and the fetch script**

```bash
npm install @huggingface/transformers
```

Create `scripts/fetch-model.mjs`:

```js
// Puts the analysis model where the app can serve it itself.
//
// The model must be served from this application's own origin -- see
// spec §2. Fetching it from a public host at RUNTIME would send a
// request from every user's browser to a third party and would fail
// outright on a corporate network that blocks it. Fetching it at BUILD
// time and serving it ourselves has neither problem.
//
// The files are gitignored rather than committed, which keeps ~25MB out
// of every clone at the cost of a build-time network dependency. That
// trade is spec §12; if it ever becomes unacceptable, commit
// public/models and public/ort and delete the prebuild hook.
//
//   node scripts/fetch-model.mjs            verify against the manifest
//   node scripts/fetch-model.mjs --update   record new hashes

import { createHash } from 'node:crypto';
import { mkdir, readFile, writeFile, copyFile, readdir } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const MODEL_ID = 'Xenova/all-MiniLM-L6-v2';
const REVISION = 'main';

const MODEL_FILES = [
  'config.json',
  'tokenizer.json',
  'tokenizer_config.json',
  'onnx/model_quantized.onnx',
];

const MANIFEST = join(ROOT, 'scripts', 'model-manifest.json');
const MODEL_DIR = join(ROOT, 'public', 'models', MODEL_ID);
const ORT_DIR = join(ROOT, 'public', 'ort');
const ORT_SOURCE = join(ROOT, 'node_modules', '@huggingface', 'transformers', 'dist');

const update = process.argv.includes('--update');

function sha256(buffer) {
  return createHash('sha256').update(buffer).digest('hex');
}

async function loadManifest() {
  if (!existsSync(MANIFEST)) return {};
  return JSON.parse(await readFile(MANIFEST, 'utf8'));
}

async function fetchModelFiles(manifest) {
  for (const file of MODEL_FILES) {
    const target = join(MODEL_DIR, file);
    if (existsSync(target) && !update) {
      const existing = await readFile(target);
      if (manifest[file] && sha256(existing) === manifest[file]) continue;
    }

    const url = `https://huggingface.co/${MODEL_ID}/resolve/${REVISION}/${file}`;
    process.stdout.write(`fetching ${file} ... `);
    const response = await fetch(url);
    if (!response.ok) throw new Error(`${url} -> HTTP ${response.status}`);
    const buffer = Buffer.from(await response.arrayBuffer());

    const hash = sha256(buffer);
    if (!update && manifest[file] && manifest[file] !== hash) {
      throw new Error(
        `${file} does not match the recorded hash. Re-run with --update only if you `
        + 'intend to accept a different model.',
      );
    }
    manifest[file] = hash;

    await mkdir(dirname(target), { recursive: true });
    await writeFile(target, buffer);
    process.stdout.write(`${(buffer.length / 1e6).toFixed(1)}MB\n`);
  }
}

// The ONNX runtime ships inside the installed package. Copying it from
// node_modules rather than downloading it means the runtime half of the
// promise needs no network at all.
async function copyOnnxRuntime() {
  if (!existsSync(ORT_SOURCE)) {
    throw new Error('@huggingface/transformers is not installed; run npm install first.');
  }
  await mkdir(ORT_DIR, { recursive: true });
  let copied = 0;
  for (const entry of await readdir(ORT_SOURCE)) {
    if (!entry.endsWith('.wasm') && !entry.endsWith('.mjs')) continue;
    await copyFile(join(ORT_SOURCE, entry), join(ORT_DIR, entry));
    copied++;
  }
  if (copied === 0) throw new Error(`No runtime files found in ${ORT_SOURCE}.`);
  process.stdout.write(`copied ${copied} ONNX runtime files\n`);
}

const manifest = await loadManifest();
await fetchModelFiles(manifest);
await copyOnnxRuntime();
await writeFile(MANIFEST, `${JSON.stringify(manifest, null, 2)}\n`);
process.stdout.write('model ready under public/\n');
```

In `package.json`, add to `scripts`:

```json
    "fetch:model": "node scripts/fetch-model.mjs",
    "prebuild": "node scripts/fetch-model.mjs",
```

In `.gitignore`, add:

```
# Fetched at build time by scripts/fetch-model.mjs -- see spec §12.
public/models/
public/ort/
```

- [ ] **Step 2: Run it and confirm the files land**

```bash
npm run fetch:model -- --update
```

Expected: four model files downloaded (the ONNX file is roughly 23MB) and the runtime files copied. Then:

```bash
ls -la public/models/Xenova/all-MiniLM-L6-v2/onnx public/ort | head -20
```

Expected: `model_quantized.onnx` present, and at least one `.wasm` in `public/ort`.

- [ ] **Step 3: Write the embedder**

Create `src/features/datastudio/text/embed.js`:

```js
// The only module in this feature that is not a pure function.
//
// Everything here exists to keep one promise (spec §2): no response text
// ever leaves the browser. Two settings do that, and the second one is
// the one that is easy to miss -- without `wasmPaths` the ONNX runtime
// is fetched from a public CDN the first time anyone opens the tab, and
// nothing on screen says so.

import { pipeline, env } from '@huggingface/transformers';

export const MODEL_ID = 'Xenova/all-MiniLM-L6-v2';
export const BATCH_SIZE = 16;

// Served by this app, not by anyone else.
env.allowRemoteModels = false;
env.localModelPath = '/models/';
env.backends.onnx.wasm.wasmPaths = '/ort/';

// Loading the model dominates the first run. The 0-40 band is reserved
// for it in the worker's progress contract, so the tab shows movement
// rather than looking hung for ten seconds.
const LOAD_PCT = 40;

export function createEmbedder() {
  let extractor = null;

  async function ready(onProgress) {
    if (extractor) return extractor;
    extractor = await pipeline('feature-extraction', MODEL_ID, {
      dtype: 'q8',
      progress_callback: (event) => {
        if (event?.status !== 'progress') return;
        const fraction = (event.progress ?? 0) / 100;
        onProgress?.({ stage: 'Loading the model', pct: Math.round(fraction * LOAD_PCT) });
      },
    });
    return extractor;
  }

  async function embedAll(texts, { onProgress } = {}) {
    const list = texts ?? [];
    if (list.length === 0) return [];

    const run = await ready(onProgress);
    const out = [];

    for (let start = 0; start < list.length; start += BATCH_SIZE) {
      const batch = list.slice(start, start + BATCH_SIZE);
      // Mean pooling and L2 normalisation happen here so every consumer
      // can treat a vector as a direction and compare with plain cosine.
      const tensor = await run(batch, { pooling: 'mean', normalize: true });
      for (const row of tensor.tolist()) out.push(Float32Array.from(row));

      onProgress?.({
        stage: 'Understanding responses',
        pct: LOAD_PCT + Math.round(((start + batch.length) / list.length) * 45),
      });
    }

    return out;
  }

  return { embedAll };
}
```

In `vercel.json`, make sure the model is cached hard once fetched — add to `headers` (create the array if it is not there):

```json
  "headers": [
    {
      "source": "/models/(.*)",
      "headers": [{ "key": "Cache-Control", "value": "public, max-age=31536000, immutable" }]
    },
    {
      "source": "/ort/(.*)",
      "headers": [{ "key": "Cache-Control", "value": "public, max-age=31536000, immutable" }]
    }
  ]
```

- [ ] **Step 4: Verify the build still works**

```bash
npm run build
```

Expected: build succeeds; `dist/models/` and `dist/ort/` exist. Then confirm nothing broke:

```bash
npm test && npm run lint
```

Expected: PASS, no new lint errors.

- [ ] **Step 5: Commit**

```bash
git add scripts/fetch-model.mjs scripts/model-manifest.json src/features/datastudio/text/embed.js package.json package-lock.json .gitignore vercel.json && git commit -m "Serve the analysis model from our own origin"
```

---

### Task 19: Remember the user's corrections

**Files:**
- Modify: `src/features/datastudio/store/db.js`
- Test: `src/features/datastudio/store/db.test.js`

**Interfaces:**
- Produces:
  - `STORE_ANALYSES = 'analyses'`, `DB_VERSION = 2`
  - `saveAnalysis(record) -> Promise<void>` where record is `{ datasetId, columnName, buckets, overrides, settings, vectors? }`
  - `loadAnalysis(datasetId) -> Promise<record | undefined>`
  - `MAX_CACHED_VECTORS = 2000`
  - `deleteDataset` also removes the analysis.

- [ ] **Step 1: Write the failing test**

Append to `src/features/datastudio/store/db.test.js`:

```js
import { saveAnalysis, loadAnalysis, MAX_CACHED_VECTORS } from './db.js';

describe('analyses', () => {
  it('round-trips buckets, overrides and settings', async () => {
    await saveAnalysis({
      datasetId: 'ds_1',
      columnName: 'Describe',
      buckets: [{ id: 'sap', label: 'SAP', description: 'd', hints: [] }],
      overrides: { retags: { '0:0': 'sap' }, noise: ['1:0'], themeNames: {}, themeMerges: {}, pinned: [], suppressed: [] },
      settings: { threshold: 0.35, granularity: 0.5 },
    });

    const record = await loadAnalysis('ds_1');
    expect(record.columnName).toBe('Describe');
    expect(record.overrides.retags['0:0']).toBe('sap');
    expect(record.settings.threshold).toBe(0.35);
  });

  it('keeps vectors for a small analysis so reopening is instant', async () => {
    await saveAnalysis({
      datasetId: 'ds_2', columnName: 'D', buckets: [], overrides: {}, settings: {},
      vectors: [Float32Array.from([1, 0]), Float32Array.from([0, 1])],
    });
    const record = await loadAnalysis('ds_2');
    expect(record.vectors).toHaveLength(2);
    expect(Array.from(record.vectors[0])).toEqual([1, 0]);
  });

  it('drops vectors for a large analysis rather than filling the quota', async () => {
    const many = Array.from({ length: MAX_CACHED_VECTORS + 1 }, () => Float32Array.from([1]));
    await saveAnalysis({
      datasetId: 'ds_3', columnName: 'D', buckets: [], overrides: {}, settings: {}, vectors: many,
    });
    const record = await loadAnalysis('ds_3');
    expect(record.vectors).toBeNull();
  });

  it('returns undefined for a dataset that was never analysed', async () => {
    expect(await loadAnalysis('nope')).toBeUndefined();
  });
});
```

- [ ] **Step 2: Run to verify it fails**

```bash
npm test -- store/db
```

Expected: FAIL — `saveAnalysis` is not exported.

- [ ] **Step 3: Implement**

In `src/features/datastudio/store/db.js`:

```js
export const DB_VERSION = 2;
```

Add beside the other store names:

```js
// Text analysis: the bucket definitions in force, the user's
// corrections, and the settings they were produced under. Never the
// analysis itself -- that is derived, and re-deriving it is cheap
// compared with storing a copy that can disagree with the data.
export const STORE_ANALYSES = 'analyses';
```

In `openDb`'s `onupgradeneeded`, add (the existing `contains` guards mean a v1 database upgrades in place without losing anything):

```js
      if (!db.objectStoreNames.contains(STORE_ANALYSES)) {
        db.createObjectStore(STORE_ANALYSES, { keyPath: 'datasetId' });
      }
```

Add the section before `// --- quota ---`:

```js
// --- text analysis ----------------------------------------------------

// Above this, the fragment vectors are re-computed on reopen instead of
// stored. 2,000 x 384 floats is about 3MB, which is a reasonable thing
// to keep; ten times that is not, and a survey that large is rare enough
// that paying for the model run again is the better trade.
export const MAX_CACHED_VECTORS = 2000;

export async function saveAnalysis(record) {
  const vectors = record?.vectors ?? null;
  const value = {
    datasetId: record.datasetId,
    columnName: record.columnName ?? '',
    buckets: record.buckets ?? [],
    overrides: record.overrides ?? {},
    settings: record.settings ?? {},
    vectors: vectors && vectors.length <= MAX_CACHED_VECTORS ? vectors : null,
    updatedAt: Date.now(),
  };
  await run(STORE_ANALYSES, 'readwrite', (store) => {
    store.put(value);
  });
}

export async function loadAnalysis(datasetId) {
  const db = await openDb();
  const tx = db.transaction(STORE_ANALYSES, 'readonly');
  return promisify(tx.objectStore(STORE_ANALYSES).get(datasetId));
}
```

In `deleteDataset`, add `STORE_ANALYSES` to the transaction's store list and delete from it:

```js
    const tx = db.transaction(
      [STORE_DATASETS, STORE_COLUMNS, STORE_PLANS, STORE_DASHBOARDS, STORE_ANALYSES],
      'readwrite',
    );
    ...
    tx.objectStore(STORE_ANALYSES).delete(id);
```

- [ ] **Step 4: Run to verify it passes**

```bash
npm test -- store/db
```

Expected: PASS, including the existing dataset, plan and dashboard tests.

- [ ] **Step 5: Commit**

```bash
git add src/features/datastudio/store/db.js src/features/datastudio/store/db.test.js && git commit -m "Remember text analysis corrections between visits"
```

---

### Task 20: The analysis worker

**Files:**
- Create: `src/features/datastudio/worker/text.worker.js`

**Interfaces:**
- Consumes: `analyze`/`rescore` (Task 17), `createEmbedder` (Task 18).
- Produces the message contract the context depends on:
  - **In:** `{ type: 'analyze', texts, breadths, buckets, columnName, settings }` · `{ type: 'rescore', buckets, settings }`
  - **Out:** `{ type: 'progress', stage, pct }` · `{ type: 'analyzed', raw }` · `{ type: 'error', message }`

  `raw` on the wire carries `vectors` as an array of `Float32Array`, which structured clone handles natively.

- [ ] **Step 1: Write the worker**

There is no unit test for this file — it is glue over two tested modules and a browser API vitest does not provide. Its behaviour is verified in Task 23 against the real survey.

Create `src/features/datastudio/worker/text.worker.js`:

```js
// The model's home -- spec §11.
//
// A worker of its own, not the studio worker. That one holds the parsed
// grid and has to answer a re-clean message the instant the user ticks a
// checkbox; giving it a 23MB model and multi-second inference sessions
// as well would make every clean wait behind an embedding run.
//
// The fragments and their vectors STAY here after an analysis, the same
// way the grid stays in the studio worker. A threshold change then costs
// one small message instead of re-embedding everything -- which is the
// difference between a control that feels live and one that does not.

import { analyze, rescore } from '../text/analysis.js';
import { createEmbedder } from '../text/embed.js';

const embedder = createEmbedder();

let current = null;

function report(progress) {
  self.postMessage({ type: 'progress', stage: progress.stage, pct: progress.pct });
}

async function handleAnalyze(msg) {
  report({ stage: 'Loading the model', pct: 2 });

  const raw = await analyze({
    texts: msg.texts,
    breadths: msg.breadths,
    buckets: msg.buckets,
    columnName: msg.columnName,
    settings: msg.settings,
    embedAll: (texts, options) => embedder.embedAll(texts, options),
    onProgress: report,
  });

  current = {
    columnName: raw.columnName,
    fragments: raw.fragments,
    vectors: raw.vectors,
    noIssueRows: raw.noIssueRows,
  };

  self.postMessage({ type: 'analyzed', raw });
}

async function handleRescore(msg) {
  if (!current) throw new Error('There is nothing analysed to re-score.');

  report({ stage: 'Grouping', pct: 60 });
  const raw = await rescore({
    columnName: current.columnName,
    fragments: current.fragments,
    vectors: current.vectors,
    noIssueRows: current.noIssueRows,
    buckets: msg.buckets,
    settings: msg.settings,
    embedAll: (texts, options) => embedder.embedAll(texts, options),
  });

  self.postMessage({ type: 'analyzed', raw });
}

self.onmessage = async (e) => {
  const msg = e.data ?? {};
  try {
    if (msg.type === 'analyze') await handleAnalyze(msg);
    else if (msg.type === 'rescore') await handleRescore(msg);
  } catch (err) {
    // Never leave the tab on a spinner. A model that will not load is
    // the most likely failure here and the message has to say so, or the
    // user is left looking at a progress bar that stopped.
    self.postMessage({
      type: 'error',
      message: err?.message || 'The analysis stopped unexpectedly.',
    });
  }
};
```

- [ ] **Step 2: Confirm it compiles into the bundle**

```bash
npm run build
```

Expected: build succeeds and the output lists a separate chunk for the text worker. Confirm the model is not inlined into the main bundle:

```bash
ls dist/assets | head -20
```

Expected: several chunks; none the size of the model.

- [ ] **Step 3: Commit**

```bash
git add src/features/datastudio/worker/text.worker.js && git commit -m "Run the analysis in a worker that keeps the model"
```

---

### Task 21: Wire the tab into Data Studio

**Files:**
- Modify: `src/features/datastudio/dataStudioStore.js`
- Modify: `src/features/datastudio/DataStudioContext.jsx`

**Interfaces:**
- Consumes: `detectTextColumns` (8), `STARTER_BUCKETS`/`UNSORTED_ID` (9), `applyOverrides`/`EMPTY_OVERRIDES` (15), `deriveColumns` (16), `saveAnalysis`/`loadAnalysis` (19), the worker contract (20).
- Produces on the context, in addition to everything already there:
  - state: `textColumns`, `textColumnName`, `rawAnalysis`, `analysis`, `buckets`, `textOverrides`, `textSettings`, `analysing`, `textProgress`, `textError`
  - actions: `startAnalysis(columnName)`, `setTextSetting(key, value)`, `updateBucket(id, patch)`, `addBucket()`, `removeBucket(id)`, `retagFragment(id, bucketId)`, `toggleNoise(id)`, `renameTheme(id, name)`, `mergeThemes(fromId, intoId)`, `togglePin(id)`, `toggleSuppress(id)`, `resetOverrides()`, `applyAnalysisColumns()`

- [ ] **Step 1: Extend the state**

In `src/features/datastudio/dataStudioStore.js`, add `'text'` to `STAGES` and these keys to `IDLE_STATE`:

```js
export const STAGES = ['idle', 'parsing', 'profiled', 'cleaning', 'canvas', 'text'];
```

```js
  // --- text analysis ---------------------------------------------------
  // `rawAnalysis` is what the model said; `textOverrides` is what the
  // user said about it; `analysis` is the two combined and is the only
  // one anything renders. Keeping the first two apart is what lets a
  // re-score leave hand corrections standing (spec §8).
  textColumns: [],
  textColumnName: '',
  rawAnalysis: null,
  analysis: null,
  buckets: [],
  textOverrides: null,
  textSettings: { threshold: 0.3, granularity: 0.45 },
  analysing: false,
  textProgress: { stage: '', pct: 0 },
  textError: '',
```

- [ ] **Step 2: Wire the worker and the actions**

In `src/features/datastudio/DataStudioContext.jsx`, add the imports:

```js
import { detectTextColumns } from './text/detectTextColumns.js';
import { STARTER_BUCKETS } from './text/buckets.js';
import { applyOverrides, EMPTY_OVERRIDES } from './text/overrides.js';
import { deriveColumns } from './text/deriveColumns.js';
import { saveAnalysis, loadAnalysis } from './store/db.js';
import { profileDataset } from './profile/profileDataset.js';
```

Add a second worker ref beside the first, created lazily — the model chunk must not be downloaded by everyone who opens Data Studio, only by whoever opens the tab:

```js
  const textWorkerRef = useRef(null);

  // Created on first use, not on mount. Constructing it eagerly would
  // pull the model chunk into every Data Studio visit, including the
  // ones that never open this tab.
  const textWorker = useCallback(() => {
    if (textWorkerRef.current) return textWorkerRef.current;

    const worker = new Worker(new URL('./worker/text.worker.js', import.meta.url), {
      type: 'module',
    });

    worker.onmessage = (e) => {
      const msg = e.data ?? {};

      if (msg.type === 'progress') {
        setState((s) => ({ ...s, textProgress: { stage: msg.stage, pct: msg.pct } }));
        return;
      }

      if (msg.type === 'analyzed') {
        setState((s) => {
          const overrides = s.textOverrides ?? EMPTY_OVERRIDES;
          return {
            ...s,
            rawAnalysis: msg.raw,
            analysis: applyOverrides(msg.raw, overrides),
            textOverrides: overrides,
            analysing: false,
            textError: '',
            textProgress: { stage: '', pct: 100 },
          };
        });
        return;
      }

      if (msg.type === 'error') {
        setState((s) => ({
          ...s,
          analysing: false,
          textError: msg.message,
          textProgress: { stage: '', pct: 0 },
        }));
      }
    };

    worker.onerror = (event) => {
      setState((s) => ({
        ...s,
        analysing: false,
        textError: event?.message || 'The analysis worker stopped unexpectedly.',
        textProgress: { stage: '', pct: 0 },
      }));
    };

    textWorkerRef.current = worker;
    return worker;
  }, []);
```

Terminate it in the existing unmount cleanup:

```js
    return () => {
      worker.terminate();
      workerRef.current = null;
      textWorkerRef.current?.terminate();
      textWorkerRef.current = null;
    };
```

Detect the text columns whenever a profile arrives — add to the `parsed` branch of the studio worker's `onmessage`, inside the `setState`:

```js
          textColumns: detectTextColumns(msg.profile, { rows: msg.grid.rows }),
          textColumnName: '',
          rawAnalysis: null,
          analysis: null,
          textOverrides: null,
```

Add the actions before the `value` memo:

```js
  // --- text analysis ----------------------------------------------------

  // How many multi-select options each respondent picked, normalised
  // against the most anyone picked. This is the one severity input that
  // is measured rather than inferred (spec §6.7), and it comes from the
  // structured column, not the prose.
  const breadthsOf = useCallback((grid, profile) => {
    const multi = (profile?.columns ?? []).find((c) => c.type === 'multi');
    if (!multi) return grid.rows.map(() => 0);

    const separator = multi.separator ?? ';';
    const counts = grid.rows.map((row) => String(row?.[multi.index] ?? '')
      .split(separator)
      .map((part) => part.trim())
      .filter(Boolean).length);

    const most = Math.max(1, ...counts);
    return counts.map((n) => n / most);
  }, []);

  const startAnalysis = useCallback((columnName) => {
    setState((s) => {
      const column = s.textColumns.find((c) => c.name === columnName) ?? s.textColumns[0];
      if (!column || !s.grid) return s;

      const buckets = s.buckets.length > 0 ? s.buckets : STARTER_BUCKETS;
      textWorker().postMessage({
        type: 'analyze',
        columnName: column.name,
        texts: s.grid.rows.map((row) => row?.[column.index] ?? ''),
        breadths: breadthsOf(s.grid, s.profile),
        buckets,
        settings: s.textSettings,
      });

      return {
        ...s,
        stage: 'text',
        textColumnName: column.name,
        buckets,
        analysing: true,
        textError: '',
        textProgress: { stage: 'Loading the model', pct: 1 },
      };
    });
  }, [textWorker, breadthsOf]);

  // Re-file against cached vectors. Never re-embeds the fragments --
  // that is what keeps a slider live (spec §16).
  const rescoreNow = useCallback((buckets, settings) => {
    textWorkerRef.current?.postMessage({ type: 'rescore', buckets, settings });
  }, []);

  const setTextSetting = useCallback((key, value) => setState((s) => {
    const textSettings = { ...s.textSettings, [key]: value };
    if (s.rawAnalysis) rescoreNow(s.buckets, textSettings);
    return { ...s, textSettings, analysing: Boolean(s.rawAnalysis) };
  }), [rescoreNow]);

  const setBuckets = useCallback((next) => setState((s) => {
    if (s.rawAnalysis) rescoreNow(next, s.textSettings);
    return { ...s, buckets: next, analysing: Boolean(s.rawAnalysis) };
  }), [rescoreNow]);

  const updateBucket = useCallback((id, patch) => setState((s) => {
    const buckets = s.buckets.map((b) => (b.id === id ? { ...b, ...patch } : b));
    if (s.rawAnalysis) rescoreNow(buckets, s.textSettings);
    return { ...s, buckets, analysing: Boolean(s.rawAnalysis) };
  }), [rescoreNow]);

  const addBucket = useCallback(() => setState((s) => {
    const buckets = [...s.buckets, {
      id: `bucket_${Date.now()}`,
      label: 'New category',
      description: '',
      hints: [],
    }];
    return { ...s, buckets };
  }), []);

  const removeBucket = useCallback((id) => setState((s) => {
    const buckets = s.buckets.filter((b) => b.id !== id);
    if (s.rawAnalysis) rescoreNow(buckets, s.textSettings);
    return { ...s, buckets, analysing: Boolean(s.rawAnalysis) };
  }), [rescoreNow]);

  // Every correction is the same shape: change the overrides record and
  // re-apply it to the raw result. Nothing here touches `rawAnalysis`.
  const editOverrides = useCallback((edit) => setState((s) => {
    if (!s.rawAnalysis) return s;
    const overrides = edit(s.textOverrides ?? EMPTY_OVERRIDES);
    return { ...s, textOverrides: overrides, analysis: applyOverrides(s.rawAnalysis, overrides) };
  }), []);

  const retagFragment = useCallback((fragmentId, bucketId) => editOverrides((o) => ({
    ...o, retags: { ...o.retags, [fragmentId]: bucketId },
  })), [editOverrides]);

  const toggleNoise = useCallback((fragmentId) => editOverrides((o) => ({
    ...o,
    noise: o.noise.includes(fragmentId)
      ? o.noise.filter((id) => id !== fragmentId)
      : [...o.noise, fragmentId],
  })), [editOverrides]);

  const renameTheme = useCallback((themeId, name) => editOverrides((o) => ({
    ...o, themeNames: { ...o.themeNames, [themeId]: name },
  })), [editOverrides]);

  const mergeThemes = useCallback((fromId, intoId) => editOverrides((o) => ({
    ...o, themeMerges: { ...o.themeMerges, [fromId]: intoId },
  })), [editOverrides]);

  const togglePin = useCallback((id) => editOverrides((o) => ({
    ...o, pinned: o.pinned.includes(id) ? o.pinned.filter((x) => x !== id) : [...o.pinned, id],
  })), [editOverrides]);

  const toggleSuppress = useCallback((id) => editOverrides((o) => ({
    ...o,
    suppressed: o.suppressed.includes(id)
      ? o.suppressed.filter((x) => x !== id)
      : [...o.suppressed, id],
  })), [editOverrides]);

  const resetOverrides = useCallback(() => editOverrides(() => EMPTY_OVERRIDES), [editOverrides]);

  /**
   * Append the analysis to the sheet as five more columns.
   *
   * The grid is re-profiled afterwards rather than patched, so the new
   * columns go through exactly the same type inference as every other
   * column -- which is how `Issue categories` becomes a multi column and
   * charts by option without a line of special-case code.
   */
  const applyAnalysisColumns = useCallback(() => setState((s) => {
    if (!s.analysis || !s.grid) return s;

    const { headers, columns } = deriveColumns(s.analysis, s.grid.rows.length);
    // Replace rather than append on a second run, or a user who
    // re-analyses ends up with "Severity" twice.
    const keep = s.grid.headers
      .map((name, i) => ({ name, i }))
      .filter(({ name }) => !headers.includes(name));

    const nextHeaders = [...keep.map((k) => k.name), ...headers];
    const nextRows = s.grid.rows.map((row, r) => [
      ...keep.map(({ i }) => row?.[i] ?? null),
      ...columns.map((column) => column[r]),
    ]);

    const grid = { headers: nextHeaders, rows: nextRows };
    const profile = profileDataset(grid);

    return {
      ...s,
      grid,
      profile,
      plan: proposeCleanPlan(profile, grid),
      stage: 'canvas',
    };
  }), []);
```

Persist and restore. Add an effect after the existing re-clean effect:

```js
  // Corrections are saved against the dataset, not the file, so they
  // survive a reload but never travel anywhere. Only meaningful once the
  // dataset has been saved and therefore has an id.
  const { datasetId, buckets, textOverrides, textSettings, textColumnName, rawAnalysis } = state;
  useEffect(() => {
    if (!datasetId || !rawAnalysis) return;
    saveAnalysis({
      datasetId,
      columnName: textColumnName,
      buckets,
      overrides: textOverrides ?? EMPTY_OVERRIDES,
      settings: textSettings,
      vectors: rawAnalysis.vectors,
    }).catch(() => { /* a failed save must not take the tab down */ });
  }, [datasetId, rawAnalysis, buckets, textOverrides, textSettings, textColumnName]);
```

In `openSavedDataset`, after the existing `setState`, restore what was stored:

```js
      const analysis = await loadAnalysis(id);
      if (analysis) {
        setState((s) => ({
          ...s,
          buckets: analysis.buckets ?? [],
          textOverrides: analysis.overrides ?? EMPTY_OVERRIDES,
          textSettings: analysis.settings ?? s.textSettings,
          textColumnName: analysis.columnName ?? '',
        }));
      }
```

Add every new action to the `value` memo and its dependency array.

- [ ] **Step 3: Verify nothing regressed**

```bash
npm test && npm run lint
```

Expected: PASS, no new lint errors. The tab has no UI yet, so nothing is visible — that is Task 22.

- [ ] **Step 4: Commit**

```bash
git add src/features/datastudio/dataStudioStore.js src/features/datastudio/DataStudioContext.jsx && git commit -m "Give Data Studio the text analysis stage and its actions"
```

---

### Task 22: The tab shell and the bucket editor

**Files:**
- Create: `src/features/datastudio/text/TextAnalysis.jsx`
- Create: `src/features/datastudio/text/BucketEditor.jsx`

**Interfaces:**
- Consumes: the context from Task 21.
- Produces: default-exported components. **Each file exports exactly one thing** — a module exporting a component may export nothing else, or it drops out of Fast Refresh and fails lint.

- [ ] **Step 1: Write the bucket editor**

Create `src/features/datastudio/text/BucketEditor.jsx`:

```jsx
import { Card } from '../../../components/ui/Surfaces';
import Button from '../../../components/ui/Button';
import { Plus, Trash } from '../../../components/ui/Icons';
import { useDataStudio } from '../useDataStudio';

/**
 * Where the categories are defined.
 *
 * The DESCRIPTION is the field that does the work and is why it gets the
 * larger control. The model matches an answer against these sentences,
 * not against the names -- rename "SAP / ERP" to "The Big System" and
 * nothing about what lands in it changes. Saying so on screen saves
 * somebody an afternoon of renaming things and wondering why.
 */
export default function BucketEditor() {
  const { buckets, updateBucket, addBucket, removeBucket, textSettings, setTextSetting, analysis } = useDataStudio();

  const countOf = (id) => analysis?.buckets.find((b) => b.id === id)?.count ?? 0;

  return (
    <Card className="ds-text-card">
      <div className="ds-toolbar">
        <span className="ds-summary">
          Answers are matched against each category&apos;s description, not its name.
        </span>
        <span className="ds-toolbar-spacer" />
        <label className="ds-field">
          <span>Confidence</span>
          <input
            type="range"
            min="0.1"
            max="0.7"
            step="0.01"
            value={textSettings.threshold}
            onChange={(e) => setTextSetting('threshold', Number(e.target.value))}
          />
          <span className="ds-summary">{textSettings.threshold.toFixed(2)}</span>
        </label>
        <Button variant="secondary" size="sm" icon={Plus} onClick={addBucket}>
          Add a category
        </Button>
      </div>

      <ul className="ds-bucket-list">
        {buckets.map((bucket) => (
          <li key={bucket.id} className="ds-bucket">
            <div className="ds-bucket-head">
              <input
                className="ds-input ds-bucket-name"
                value={bucket.label}
                aria-label="Category name"
                onChange={(e) => updateBucket(bucket.id, { label: e.target.value })}
              />
              <span className="ds-summary">{countOf(bucket.id)} issues</span>
              <Button
                variant="ghost"
                size="sm"
                icon={Trash}
                aria-label={`Remove ${bucket.label}`}
                onClick={() => removeBucket(bucket.id)}
              />
            </div>
            <textarea
              className="ds-input ds-bucket-description"
              rows={2}
              aria-label={`What belongs in ${bucket.label}`}
              placeholder="Describe in a sentence what belongs here."
              value={bucket.description}
              onChange={(e) => updateBucket(bucket.id, { description: e.target.value })}
            />
          </li>
        ))}
      </ul>
    </Card>
  );
}
```

- [ ] **Step 2: Write the tab shell**

Create `src/features/datastudio/text/TextAnalysis.jsx`:

```jsx
import { useState } from 'react';
import { Card, EmptyState, ErrorBanner } from '../../../components/ui/Surfaces';
import Button from '../../../components/ui/Button';
import { RefreshCw, BarChart } from '../../../components/ui/Icons';
import { useDataStudio } from '../useDataStudio';
import BucketEditor from './BucketEditor';
import IssueTable from './IssueTable';
import ThemeList from './ThemeList';
import PriorityBoard from './PriorityBoard';

const VIEWS = [
  ['buckets', 'Categories'],
  ['issues', 'Issues'],
  ['themes', 'Themes'],
  ['priority', 'Priority'],
];

/**
 * The Text Analysis stage.
 *
 * Two things on this screen are deliberate and easy to "tidy away":
 *
 *   * the progress bar names the stage it is in. Loading the model takes
 *     the better part of ten seconds the first time, and a bar with no
 *     label reads as a hang.
 *   * "everything landed in Unsorted" gets its own message with the fix
 *     in it. The alternative is an empty screen that looks broken but is
 *     the model being honest.
 */
export default function TextAnalysis() {
  const {
    textColumns, textColumnName, analysis, analysing, textProgress, textError,
    startAnalysis, applyAnalysisColumns, resetOverrides, setStage,
  } = useDataStudio();
  const [view, setView] = useState('buckets');

  if (textColumns.length === 0) {
    return <EmptyState>Nothing in this sheet is long enough to read as written answers.</EmptyState>;
  }

  const unsorted = analysis?.buckets.find((b) => b.id === 'unsorted')?.count ?? 0;
  const filed = analysis?.fragments.filter((f) => !f.noise).length ?? 0;
  const allUnsorted = filed > 0 && unsorted === filed;

  return (
    <>
      {textError && <ErrorBanner message={textError} onRetry={() => startAnalysis(textColumnName)} />}

      <div className="ds-toolbar">
        <label className="ds-field">
          <span>Read</span>
          <select
            className="ds-select"
            value={textColumnName}
            onChange={(e) => startAnalysis(e.target.value)}
          >
            {textColumns.map((c) => <option key={c.name} value={c.name}>{c.name}</option>)}
          </select>
        </label>

        <span className="ds-toolbar-spacer" />

        {analysis && (
          <span className="ds-summary">
            {`${filed} issues from ${analysis.fragments.length === 0 ? 0 : new Set(analysis.fragments.map((f) => f.row)).size} people · `}
            {`${analysis.themes.length} themes · ${unsorted} unsorted`}
          </span>
        )}

        <Button variant="secondary" size="sm" onClick={resetOverrides} disabled={!analysis}>
          Reset my edits
        </Button>
        <Button
          variant="secondary"
          size="sm"
          icon={BarChart}
          onClick={applyAnalysisColumns}
          disabled={!analysis}
        >
          Add to my charts
        </Button>
        <Button variant="secondary" size="sm" icon={RefreshCw} onClick={() => setStage('canvas')}>
          Back to charts
        </Button>
      </div>

      {!analysis && !analysing && (
        <Card>
          <EmptyState>
            <p>Read the written answers and sort them into categories.</p>
            <p className="ds-drop-hint">
              This runs on your machine. Nothing is uploaded. The first run downloads the
              model once, which takes a few seconds.
            </p>
            <Button onClick={() => startAnalysis(textColumnName || textColumns[0].name)}>
              Analyse the answers
            </Button>
          </EmptyState>
        </Card>
      )}

      {analysing && (
        <Card>
          <p className="ds-summary">{textProgress.stage || 'Working'}</p>
          <div className="bar-track">
            <span className="bar-fill" style={{ transform: `scaleX(${(textProgress.pct ?? 0) / 100})` }} />
          </div>
        </Card>
      )}

      {analysis && (
        <>
          {allUnsorted && (
            <Card className="ds-text-notice">
              <p className="ds-summary">
                Nothing matched a category confidently. Lower the confidence setting on the
                Categories tab, or describe your categories in more detail — the descriptions
                are what answers are matched against.
              </p>
            </Card>
          )}

          <div className="ds-text-tabs" role="tablist">
            {VIEWS.map(([id, label]) => (
              <button
                key={id}
                type="button"
                role="tab"
                aria-selected={view === id}
                className={`ds-text-tab${view === id ? ' ds-text-tab-on' : ''}`}
                onClick={() => setView(id)}
              >
                {label}
              </button>
            ))}
          </div>

          {view === 'buckets' && <BucketEditor />}
          {view === 'issues' && <IssueTable />}
          {view === 'themes' && <ThemeList />}
          {view === 'priority' && <PriorityBoard />}
        </>
      )}
    </>
  );
}
```

- [ ] **Step 3: Commit (the tab does not render until Task 23 adds the three missing files)**

```bash
git add src/features/datastudio/text/TextAnalysis.jsx src/features/datastudio/text/BucketEditor.jsx && git commit -m "Add the text analysis tab shell and category editor"
```

---

### Task 23: Issues, themes and priority

**Files:**
- Create: `src/features/datastudio/text/IssueTable.jsx`
- Create: `src/features/datastudio/text/ThemeList.jsx`
- Create: `src/features/datastudio/text/PriorityBoard.jsx`

**Interfaces:**
- Consumes: `analysis`, `buckets`, `retagFragment`, `toggleNoise`, `renameTheme`, `mergeThemes`, `togglePin`, `toggleSuppress` from the context.
- Produces: one default-exported component per file.

- [ ] **Step 1: Write the issue table**

Create `src/features/datastudio/text/IssueTable.jsx`:

```jsx
import { Card, EmptyState } from '../../../components/ui/Surfaces';
import { useDataStudio } from '../useDataStudio';

/**
 * Every separated issue, with the category it was given and a dropdown
 * to disagree.
 *
 * A noise row stays on screen, struck through, rather than disappearing.
 * Removing it outright would leave the user unable to undo a misclick,
 * and unable to see how much they had excluded.
 */
export default function IssueTable() {
  const { analysis, buckets, retagFragment, toggleNoise } = useDataStudio();

  if (!analysis || analysis.fragments.length === 0) {
    return <EmptyState>No issues were found in these answers.</EmptyState>;
  }

  const options = [...buckets, { id: 'unsorted', label: 'Unsorted' }];

  return (
    <Card className="ds-table-card">
      <div className="ds-table-scroll">
        <table className="ds-table">
          <thead>
            <tr>
              <th className="ds-num">Row</th>
              <th>Issue</th>
              <th>Category</th>
              <th className="ds-num">Severity</th>
              <th>Use</th>
            </tr>
          </thead>
          <tbody>
            {analysis.fragments.map((fragment) => (
              <tr key={fragment.id} className={fragment.noise ? 'ds-issue-noise' : undefined}>
                <td className="ds-num">{fragment.row + 1}</td>
                <td className="ds-issue-text">{fragment.text}</td>
                <td>
                  <select
                    className="ds-select"
                    aria-label={`Category for row ${fragment.row + 1}`}
                    value={fragment.bucketId}
                    onChange={(e) => retagFragment(fragment.id, e.target.value)}
                  >
                    {options.map((b) => <option key={b.id} value={b.id}>{b.label}</option>)}
                  </select>
                </td>
                <td className="ds-num">{Math.round((fragment.severity ?? 0) * 100)}</td>
                <td>
                  <label className="ds-field">
                    <input
                      type="checkbox"
                      checked={!fragment.noise}
                      aria-label={`Count row ${fragment.row + 1}`}
                      onChange={() => toggleNoise(fragment.id)}
                    />
                    <span>{fragment.noise ? 'Excluded' : 'Counted'}</span>
                  </label>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </Card>
  );
}
```

- [ ] **Step 2: Write the theme list**

Create `src/features/datastudio/text/ThemeList.jsx`:

```jsx
import { Card, EmptyState } from '../../../components/ui/Surfaces';
import { useDataStudio } from '../useDataStudio';

/**
 * The groupings the model found on its own.
 *
 * A theme's name is four distinctive words, not a sentence -- the model
 * measures sameness, it does not write. The name field is editable
 * because that is the honest interface for a starting point, and the
 * example lines under each theme are what someone reads to decide what
 * to call it.
 */
export default function ThemeList() {
  const { analysis, renameTheme, mergeThemes } = useDataStudio();

  if (!analysis) return null;
  if (analysis.themes.length === 0) {
    return (
      <EmptyState>
        Too few answers to find themes in. Categories and the issue list still work.
      </EmptyState>
    );
  }

  const textOf = (id) => analysis.fragments.find((f) => f.id === id)?.text ?? '';

  return (
    <div className="ds-theme-list">
      {analysis.themes.map((theme) => (
        <Card key={theme.id} className="ds-theme">
          <div className="ds-bucket-head">
            <input
              className="ds-input ds-bucket-name"
              value={theme.name}
              aria-label="Theme name"
              onChange={(e) => renameTheme(theme.id, e.target.value)}
            />
            <span className="ds-summary">
              {`${theme.count} issues · ${theme.respondents} people`}
            </span>
            <select
              className="ds-select"
              aria-label={`Merge ${theme.name} into another theme`}
              value=""
              onChange={(e) => e.target.value && mergeThemes(theme.id, e.target.value)}
            >
              <option value="">Merge into…</option>
              {analysis.themes
                .filter((other) => other.id !== theme.id)
                .map((other) => (
                  <option key={other.id} value={other.id}>{other.name}</option>
                ))}
            </select>
          </div>
          <ul className="ds-theme-examples">
            {theme.fragmentIds.slice(0, 3).map((id) => <li key={id}>{textOf(id)}</li>)}
          </ul>
        </Card>
      ))}
    </div>
  );
}
```

- [ ] **Step 3: Write the priority board**

Create `src/features/datastudio/text/PriorityBoard.jsx`:

```jsx
import { Card, EmptyState } from '../../../components/ui/Surfaces';
import Button from '../../../components/ui/Button';
import { useDataStudio } from '../useDataStudio';

/**
 * The ranked list, and the one screen that has to explain itself.
 *
 * "People" leads the row because it leads the score: five people with one
 * mild complaint each outrank one person with five furious ones. Showing
 * the issue count first would suggest the opposite ordering and make the
 * list look wrong to anyone who read it carefully.
 */
export default function PriorityBoard() {
  const { analysis, togglePin, toggleSuppress } = useDataStudio();

  if (!analysis || analysis.priority.length === 0) {
    return <EmptyState>Nothing to rank yet.</EmptyState>;
  }

  return (
    <Card className="ds-table-card">
      <p className="ds-summary">
        Ranked by how many different people raised it, scaled by how strongly they wrote.
        Severity is a signal from the wording, not a judgement.
      </p>
      <div className="ds-table-scroll">
        <table className="ds-table">
          <thead>
            <tr>
              <th className="ds-num">#</th>
              <th>Issue</th>
              <th>From</th>
              <th className="ds-num">People</th>
              <th className="ds-num">Mentions</th>
              <th className="ds-num">Severity</th>
              <th />
            </tr>
          </thead>
          <tbody>
            {analysis.priority.map((item, i) => (
              <tr key={`${item.kind}:${item.id}`} className={item.suppressed ? 'ds-issue-noise' : undefined}>
                <td className="ds-num">{i + 1}</td>
                <td>{item.label}</td>
                <td><span className="ds-summary">{item.kind === 'bucket' ? 'Category' : 'Theme'}</span></td>
                <td className="ds-num">{item.respondents}</td>
                <td className="ds-num">{item.count}</td>
                <td className="ds-num">{Math.round(item.meanSeverity * 100)}</td>
                <td className="ds-priority-actions">
                  <Button variant="ghost" size="sm" onClick={() => togglePin(item.id)}>
                    {item.pinned ? 'Unpin' : 'Pin'}
                  </Button>
                  <Button variant="ghost" size="sm" onClick={() => toggleSuppress(item.id)}>
                    {item.suppressed ? 'Restore' : 'Hide'}
                  </Button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </Card>
  );
}
```

- [ ] **Step 4: Verify lint**

```bash
npm run lint
```

Expected: no new errors.

A note on button variants, verified while planning: `shell.css` defines `ui-btn-primary`, `ui-btn-ghost`, `ui-btn-subtle` and `ui-btn-danger`. It does NOT define `ui-btn-secondary`, even though the rest of Data Studio passes `variant="secondary"` — those buttons render unstyled today. Use `ghost` and `subtle` in new code. Do not rename the existing call sites as part of this plan; that is a separate change.

- [ ] **Step 5: Commit**

```bash
git add src/features/datastudio/text/IssueTable.jsx src/features/datastudio/text/ThemeList.jsx src/features/datastudio/text/PriorityBoard.jsx && git commit -m "Add the issue, theme and priority screens"
```

---

### Task 24: Style the tab and put a door on it

**Files:**
- Modify: `src/styles/datastudio.css`
- Modify: `src/pages/DataStudioPage.jsx`
- Modify: `src/features/datastudio/clean/CleanReview.jsx`

**Interfaces:**
- Consumes: the components from Tasks 22–23, `startAnalysis` and the `text` stage from Task 21.
- Produces: a reachable, styled tab.

- [ ] **Step 1: Add the styles**

Append to `src/styles/datastudio.css`. Every colour is a token — the section has to work in both themes, and `datastudio.css` loads after `shell.css` so the tokens are already re-pointed:

```css
/* --- text analysis ---------------------------------------------------
   The tab reuses .ds-table, .ds-toolbar, .ds-field and .ds-select
   wholesale. Only what has no equivalent is defined here. */

.ds-text-tabs {
  display: flex;
  gap: 4px;
  margin: 12px 0;
  border-bottom: 1px solid var(--it-line);
}

.ds-text-tab {
  padding: 8px 14px;
  border: none;
  border-bottom: 2px solid transparent;
  background: none;
  color: var(--it-ink-soft);
  font: inherit;
  cursor: pointer;
}

.ds-text-tab-on {
  color: var(--it-brand);
  border-bottom-color: var(--it-brand);
}

.ds-text-card,
.ds-text-notice {
  margin-bottom: 12px;
}

.ds-text-notice {
  border-left: 3px solid var(--it-accent);
  background: var(--it-accent-wash);
}

.ds-bucket-list {
  display: grid;
  gap: 10px;
  margin: 0;
  padding: 0;
  list-style: none;
}

.ds-bucket {
  display: grid;
  gap: 6px;
  padding: 10px;
  border: 1px solid var(--it-line);
  border-radius: var(--it-radius);
  background: var(--it-panel);
}

.ds-bucket-head {
  display: flex;
  align-items: center;
  gap: 8px;
}

.ds-input {
  width: 100%;
  padding: 6px 8px;
  border: 1px solid var(--it-line);
  border-radius: 8px;
  background: var(--it-panel);
  color: var(--it-ink);
  font: inherit;
}

.ds-bucket-name {
  flex: 1;
  font-weight: 600;
}

.ds-bucket-description {
  resize: vertical;
}

.ds-theme-list {
  display: grid;
  gap: 10px;
}

.ds-theme-examples {
  margin: 8px 0 0;
  padding-left: 18px;
  color: var(--it-ink-soft);
  font-size: 0.9em;
}

.ds-issue-text {
  max-width: 44rem;
}

/* Excluded rows stay on screen rather than vanishing: a user who
   mis-clicks needs to see what they removed in order to put it back. */
.ds-issue-noise {
  opacity: 0.5;
  text-decoration: line-through;
}

.ds-priority-actions {
  display: flex;
  gap: 4px;
  white-space: nowrap;
}
```

- [ ] **Step 2: Route the stage and add the entry points**

In `src/pages/DataStudioPage.jsx`, import the tab and route it:

```js
import TextAnalysis from '../features/datastudio/text/TextAnalysis';
```

```js
function DataStudioBody() {
  const { stage } = useDataStudio();
  if (stage === 'parsing') return <ParsingStage />;
  if (stage === 'idle') return <DropStage />;
  if (stage === 'cleaning') return <CleanReview />;
  if (stage === 'text') return <TextAnalysis />;
  if (stage === 'canvas') return <CanvasStage />;
  return <ProfileStage />;
}
```

In `CanvasStage`, pull `textColumns` and `setStage` from the context and add the button before "Save this data" — rendered only when the sheet has something to analyse, so it never appears on a sheet of numbers:

```jsx
        {textColumns.length > 0 && (
          <Button variant="ghost" size="sm" onClick={() => setStage('text')}>
            Text analysis
          </Button>
        )}
```

In `src/features/datastudio/clean/CleanReview.jsx`, add the same button to its toolbar beside "Build the dashboard", using `textColumns` and `setStage` from `useDataStudio()`.

- [ ] **Step 3: Verify in the browser**

Start the dev server through the preview tooling (never `npm run dev` in a shell) and drop the survey file on it.

Expected, in order:
1. The Data Studio import screen accepts the `.xlsx`.
2. On the columns screen, `Which of the following operational or reporting activities…` reads as type **multi**.
3. The clean screen shows a **Text analysis** button.
4. Opening it offers `Please describe your selected challenge(s)…` in the Read dropdown.
5. "Analyse the answers" shows *Loading the model*, then *Understanding responses*, then results.
6. The console shows no request to any host but this one.

That last point is the spec §2 check and is not optional. Read the network requests and confirm every entry is same-origin.

- [ ] **Step 4: Commit**

```bash
git add src/styles/datastudio.css src/pages/DataStudioPage.jsx src/features/datastudio/clean/CleanReview.jsx && git commit -m "Style the text analysis tab and make it reachable"
```

---

### Task 25: Prove it on the real survey, then write it down

**Files:**
- Create: `src/features/datastudio/text/pipeline.integration.test.js`
- Modify: `AGENTS.md`

**Interfaces:**
- Consumes: everything.
- Produces: the end-to-end proof and the project knowledge base entry.

- [ ] **Step 1: Write the integration test**

Create `src/features/datastudio/text/pipeline.integration.test.js`:

```js
import { describe, it, expect, vi } from 'vitest';
import { analyze } from './analysis.js';
import { applyOverrides, EMPTY_OVERRIDES } from './overrides.js';
import { deriveColumns } from './deriveColumns.js';
import { profileDataset } from '../profile/profileDataset.js';

// Verbatim shapes from the real export: the bracketed labels, the
// trailing semicolons, and the "no issue from IT" row.
const RESPONSES = [
  'no issue from IT ',
  'Financial data is currently collected from multiple Excel files and different subsidiaries. '
    + 'The process involves extensive manual consolidation, which is repetitive, time-consuming and prone to human error. '
    + 'Automating extraction and report generation would reduce turnaround time.',
  'Selected Challenge]: Data Collection\n[Detailed Description]:\n'
    + 'I need to collect and monitor information from multiple WhatsApp groups and Excel files. '
    + 'Because there are many different groups, important information can sometimes be missed.',
  'Approvals are chased by email and nobody knows the current status of a request. '
    + 'Reminders have to be sent manually every week.',
  'The monthly report is rebuilt from scratch each time and version control is guesswork.',
  'SAP postings fail when master data is wrong, and correcting it is a manual job.',
];

const fakeEmbed = vi.fn(async (texts) => texts.map((text) => {
  const lower = text.toLowerCase();
  return Float32Array.from([
    /approv|sign-off|status|remind|chase/.test(lower) ? 1 : 0,
    /sap|erp|posting|master data/.test(lower) ? 1 : 0,
    /consolidat|report|excel|file|version/.test(lower) ? 1 : 0,
    /whatsapp|group|message|missed|communicat/.test(lower) ? 1 : 0,
  ]);
}));

describe('the text analysis pipeline, end to end', () => {
  it('turns written answers into chartable columns', async () => {
    const raw = await analyze({
      texts: RESPONSES,
      breadths: [0, 0.9, 0.4, 0.3, 0.2, 0.2],
      buckets: [
        { id: 'approvals', label: 'Approvals & Workflow', description: 'approval sign-off status reminder chase', hints: [] },
        { id: 'sap', label: 'SAP / ERP', description: 'sap erp posting master data', hints: [] },
        { id: 'consolidation', label: 'Data Consolidation & Reporting', description: 'consolidating excel files into a report with version control', hints: [] },
        { id: 'communication', label: 'Communication & Coordination', description: 'whatsapp groups messages missed communication', hints: [] },
      ],
      columnName: 'Describe',
      embedAll: fakeEmbed,
    });

    // The non-answer is excluded, and the rest produced more issues than
    // there were respondents -- which is the whole point of splitting.
    expect(raw.noIssueRows).toContain(0);
    expect(raw.fragments.length).toBeGreaterThan(RESPONSES.length - 1);
    // Nothing carries the pasted-in label through.
    for (const fragment of raw.fragments) {
      expect(fragment.text).not.toContain('Detailed Description');
    }

    const analysis = applyOverrides(raw, EMPTY_OVERRIDES);
    const { headers, columns } = deriveColumns(analysis, RESPONSES.length);

    // The derived columns go through the ordinary profiler, and
    // "Issue categories" has to come out as a multi column or the
    // by-option chart in spec §9 does not exist.
    const grid = {
      headers,
      rows: RESPONSES.map((_, r) => columns.map((column) => column[r])),
    };
    const profile = profileDataset(grid);

    const categories = profile.columns.find((c) => c.name === 'Issue categories');
    expect(['multi', 'categorical']).toContain(categories.type);

    const severity = profile.columns.find((c) => c.name === 'Severity');
    expect(severity.type).toBe('numeric');
    expect(severity.role).toBe('measure');
  });
});
```

- [ ] **Step 2: Run the whole suite**

```bash
npm test && npm run lint && npm run build
```

Expected: every test PASSES, lint reports only the four pre-existing files, the build succeeds.

- [ ] **Step 3: Verify against the actual spreadsheet in the browser**

Drop `IT Operational Efficiency & Process Improvement Survey(1-42).xlsx` on the running dev server and check, with real numbers on screen:

- roughly 100–130 issues from 42 respondents
- the Unsorted pile is a minority, not everything and not nothing
- themes have names made of four words
- changing the confidence slider re-files visibly and in well under a second
- editing a bucket description re-files in well under half a second
- **every network request is same-origin** — this is the spec §2 check
- "Add to my charts" returns to the canvas with `Issue category`, `Theme` and `Severity` available in the tile editor

Record anything that misses the §16 budgets. A slider that re-embeds is a bug, not a slow machine.

- [ ] **Step 4: Document it**

In `AGENTS.md`, add to the **WHERE TO LOOK** table:

```
| Multi-select column encoding | `engine/dataset.js` (`encodeMulti`), detected in `profile/inferType.js` |
| Splitting an answer into issues | `src/features/datastudio/text/splitIssues.js` |
| The analysis categories | `src/features/datastudio/text/buckets.js` |
| The model, and where it is served from | `src/features/datastudio/text/embed.js`, `scripts/fetch-model.mjs` |
| Text analysis state and worker | `DataStudioContext.jsx`, `worker/text.worker.js` |
```

Add to **ANTI-PATTERNS**:

```
- Don't let the analysis model or its runtime come from anywhere but this
  app's own origin. `embed.js` sets `allowRemoteModels = false`,
  `localModelPath` AND `wasmPaths`. Drop the last one and the ONNX runtime is
  fetched from a public CDN the first time anyone opens the tab -- silently,
  because the feature still works. The tab tells the user nothing leaves
  their machine; that setting is what makes it true.
- Don't re-embed the fragments when a setting changes. The threshold and
  granularity controls re-score against cached vectors -- that is why they
  feel live. Re-embedding turns a 100ms control into a five-second one and
  nothing in the code will look wrong.
- Don't score the priority list on how many issues were raised. It counts
  distinct PEOPLE, scaled by severity: one person writing five furious
  sentences must not outrank five people writing one each. `rankIssues.js`
  carries the test that pins it.
- Don't drop a fragment that starts with "no". "no issue from IT" is a
  non-answer; "No proper system exists for tracking approvals" is a report.
  `boilerplate.js` requires BOTH a leading non-answer word and a short body,
  and the pair has its own test.
- Don't add `multi` handling after the `column.dictionary` branch in
  `aggregate.js` or `filterMask.js`. A multi column carries a dictionary too,
  so the dictionary branch catches it first and every option collapses into
  one meaningless category.
- Don't derive `dataset.rowCount` from the first column's `values.length`. A
  multi column's `values` is the flat option array and is longer than the
  grid, so every mask allocated from it would be the wrong size.
```

Add to **ROUTES**, in the `/data-studio` row's description: `Text analysis tab reads written answers locally.`

- [ ] **Step 5: Commit**

```bash
git add src/features/datastudio/text/pipeline.integration.test.js AGENTS.md && git commit -m "Prove the text analysis pipeline and document it"
```

---

## Self-Review

**Spec coverage.**

| Spec section | Task |
|---|---|
| §2 no network with response text | 18 (embed + fetch script), 24 step 3, 25 step 3 |
| §4 where it lives (the `text` stage) | 21, 24 |
| §5 module layout | 6–18, 22–23 |
| §6.1 detectTextColumns | 8 |
| §6.2 splitIssues | 6, 7 |
| §6.3 embed | 18 |
| §6.4 similarity + threshold | 11 |
| §6.5 cluster + granularity + guard | 12 |
| §6.6 labelCluster | 13 |
| §6.7 severity (incl. breadth from multi) | 10, 21 (`breadthsOf`) |
| §6.8 rankIssues | 14 |
| §7 starter buckets | 9 |
| §8 overrides | 15, 21 |
| §9 derived columns | 16, 21 (`applyAnalysisColumns`) |
| §10 multi-select splitting | 1–5 |
| §11 worker protocol | 20 |
| §12 model hosting and build | 18 |
| §13 persistence | 19, 21 |
| §14 errors and edge cases | 20 (worker catch), 22 (all-Unsorted, no columns, few fragments), 12 (size guard) |
| §15 testing | every task's step 1 |
| §16 performance budget | 17 (`rescore`), 20 (cached vectors), 25 step 3 |
| §17 later phases | out of scope, stated |

**Deviations from the spec, and why.**

1. **Multi-select splitting is a column type, not a clean-plan reshape.** The spec (§10) described it as a clean operation. Implementing it as a first-class `multi` type instead means the profiler, the encoder, the aggregator and the filter all understand it, so the "top challenges" chart is built the ordinary way rather than through a special case. Same user-visible outcome, less machinery.
2. **`analyze` takes its embedder as an argument** rather than importing `embed.js`. The spec left this open. Injecting it is what makes every pipeline test run with no model at all.
3. **`themeSplits` is not implemented.** The spec's override record listed it. Merging, renaming, retagging and noise cover the real editing need; splitting a theme by hand is a second interaction pattern for a case nobody has hit yet. Left out under YAGNI — the override record has room for it, and `applyOverrides` is where it would go.

**Placeholder scan:** none. Every step names its files, its command and its expected result.

**Type consistency:** `bucketId`, `themeId`, `fragmentIds`, `respondents`, `meanSeverity`, `count`, `noise`, `pinned`, `suppressed` are spelled identically in Tasks 14, 15, 16, 21, 22 and 23. `separator` is the property name in Tasks 1, 2, 3 and 4. `embedAll(texts, { onProgress })` has the same signature in Tasks 17, 18 and 20. `DERIVED_HEADERS` in Task 16 matches the names read in Task 25.
