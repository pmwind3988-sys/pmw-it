# Data Studio Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a `/data-studio` section where a user imports an Excel workbook and lands on a cross-filtered dashboard of charts built from their own data, all in the browser.

**Architecture:** A pure-function pipeline — parse → detect header → profile → clean plan → apply → columnar dataset — runs inside a Web Worker and hands the main thread typed arrays. Charts are ECharts tiles that are pure functions of `(dataset, rowMask, tileSpec)`; cross-filtering computes one shared `Uint8Array` mask that every tile reads. Nothing persists to SharePoint; datasets and dashboards live in IndexedDB.

**Tech Stack:** React 19, Vite 8, SheetJS (`xlsx`), Apache ECharts (tree-shaken from `echarts/core`), IndexedDB, Vitest.

**Spec:** `docs/superpowers/specs/2026-08-20-data-studio-design.md` — read it before starting. Every task below cites the spec sections it implements; where this plan and the spec disagree, the spec wins and you should stop and flag it.

## Global Constraints

These apply to **every** task. They come from the spec and from `AGENTS.md`, and an engineer new to this repo will get them wrong by default.

- **No SharePoint token is acquired anywhere in this feature.** Never import `useSharePointToken`, `sharePointRequest`, or anything from `src/services/sharePointService.js`. The section must work offline once loaded.
- **Pages do not gate themselves.** `src/components/AppShell.jsx` is the one auth gate. A page renders `<AppShell title subtitle actions>` plus its own body — never its own nav, theme toggle, or sign-out.
- **Never call React Router's `navigate()` inside `useEffect`** — it re-renders, re-runs the effect, and loops forever. Use `window.location.replace()` in effects; `navigate()` is correct in event handlers.
- **Never export a helper from a file that also exports a component.** It drops the file out of Fast Refresh and fails `npm run lint`. Helpers go in their own module (this is why `src/utils/initials.js` exists).
- **Design tokens, exact names:** `--it-brand`, `--it-brand-deep`, `--it-brand-mid`, `--it-brand-line`, `--it-brand-wash`, `--it-canvas`, `--it-panel`, `--it-ink`, `--it-ink-soft`, `--it-line`, `--it-accent`, `--it-good`, `--it-danger`, `--it-radius`, `--it-card-shadow`. Dark mode is `[data-theme='dark']` on the root element. Never hardcode a hex value.
- **Mobile-first CSS.** This project uses `min-width` breakpoints at `640px` and `1024px` only. Honour `@media (prefers-reduced-motion: reduce)`.
- **Stylesheet order in `src/main.jsx` is load-bearing** — `index.css` → `App.css` → `shell.css` → `auth.css`. Add `datastudio.css` **after** `shell.css`.
- **ESM only.** `package.json` has `"type": "module"`; there is no CommonJS in this repo.
- **Do not add `assetsInclude: ['**/*.html']` to `vite.config.js`** — it breaks the HTML entry and produces a bundle-less build.
- **Commit after every task.** Branch is `feat/data-studio`.
- `npm run lint` has **pre-existing** errors in `FormPage`, `AssetChecklistPage`, `SignatureDialog` and `ThemeContext`. Those are not yours. Do not fix them, and do not let them mask new errors in files you create — check the filenames in the output.

---

## File Structure

| File | Responsibility |
|---|---|
| `src/features/datastudio/time/malaysiaTime.js` | Parse any date representation to epoch ms; format as MYT |
| `src/features/datastudio/profile/inferType.js` | Value-level parsers + one column's type verdict |
| `src/features/datastudio/profile/profileColumn.js` | One column's stats, role, and coercion casualties |
| `src/features/datastudio/profile/profileDataset.js` | Orchestrates profiling across all columns |
| `src/features/datastudio/ingest/detectHeader.js` | Find the real header row in a raw grid |
| `src/features/datastudio/ingest/parseWorkbook.js` | SheetJS wrapper: File → sheets of raw grids |
| `src/features/datastudio/worker/studio.worker.js` | Runs parse + profile + clean off the main thread |
| `src/features/datastudio/clean/cleanOps.js` | The ten cleaning operations, one pure function each |
| `src/features/datastudio/clean/proposeCleanPlan.js` | Profile → suggested steps with counts and previews |
| `src/features/datastudio/clean/applyCleanPlan.js` | Raw columns + plan → typed columnar dataset |
| `src/features/datastudio/engine/dataset.js` | Columnar store: construction and typed accessors |
| `src/features/datastudio/engine/filterMask.js` | Filters → shared `Uint8Array` row mask |
| `src/features/datastudio/engine/aggregate.js` | Group-by, measures, binning, date truncation, top-N |
| `src/features/datastudio/canvas/echartsTheme.js` | CSS tokens → ECharts theme, rebuilt on theme flip |
| `src/features/datastudio/canvas/EChart.jsx` | Thin React wrapper around an ECharts instance |
| `src/features/datastudio/canvas/chartSpecs.js` | Tile spec + aggregate result → ECharts option |
| `src/features/datastudio/canvas/ChartTile.jsx` | One tile: aggregate, render, emit clicks |
| `src/features/datastudio/canvas/CanvasGrid.jsx` | 12-column grid, size presets, reordering |
| `src/features/datastudio/suggest/suggestCharts.js` | Profile → ranked starter tiles |
| `src/features/datastudio/store/db.js` | IndexedDB open/upgrade, CRUD, quota handling |
| `src/features/datastudio/store/useDashboards.js` | React hook over `db.js` |
| `src/features/datastudio/DataStudioContext.jsx` | Dataset + filters + tiles state for the section |
| `src/pages/DataStudioPage.jsx` | Route: AppShell + stage switch |
| `src/styles/datastudio.css` | All styling for the section |

---

# PHASE 1 — Ingest and profile

> **Deviation from the spec's phasing, deliberate:** the spec puts `malaysiaTime.js` in phase 2, but type inference (§7.5) needs date parsing to decide whether a column is temporal. So date *parsing* lands in Task 1 and the *UTC-shift toggle UI* lands in Phase 2. Nothing else moves.

---

## Task 1: Vitest and `malaysiaTime.js`

Implements spec §9. This is the foundation task — it establishes the test runner the rest of the plan depends on.

**Files:**
- Modify: `package.json` (devDependency + scripts)
- Modify: `vite.config.js` (test block)
- Create: `src/features/datastudio/time/malaysiaTime.js`
- Test: `src/features/datastudio/time/malaysiaTime.test.js`

**Interfaces:**
- Consumes: nothing.
- Produces:
  - `MYT_OFFSET_MIN = 480`
  - `excelSerialToEpochMs(serial: number) → number`
  - `detectDateOrder(values: string[]) → 'dmy' | 'mdy' | 'iso' | 'conflict' | 'ambiguous'`
  - `toEpochMs(value: unknown, opts: { order?: 'dmy'|'mdy'|'iso', sourceZone?: 'local'|'utc', dateOnly?: boolean }) → number` — returns `NaN` when unparseable
  - `formatMYT(epochMs: number, style?: 'datetime'|'date'|'time') → string`

- [ ] **Step 1: Install Vitest and add scripts**

```bash
npm install -D vitest@^3
```

Then add to the `"scripts"` block in `package.json`:

```json
"test": "vitest run",
"test:watch": "vitest"
```

- [ ] **Step 2: Configure Vitest in `vite.config.js`**

Add a `test` key to the existing `defineConfig` object. The environment is `node` because every unit in Phases 1–3 is a pure function — no DOM required, and `node` is markedly faster.

```js
test: {
  environment: 'node',
  include: ['src/**/*.test.js'],
},
```

- [ ] **Step 3: Write the failing tests**

Create `src/features/datastudio/time/malaysiaTime.test.js`:

```js
import { describe, it, expect } from 'vitest';
import {
  MYT_OFFSET_MIN,
  excelSerialToEpochMs,
  detectDateOrder,
  toEpochMs,
  formatMYT,
} from './malaysiaTime.js';

describe('MYT_OFFSET_MIN', () => {
  it('is a flat UTC+8 with no DST', () => {
    expect(MYT_OFFSET_MIN).toBe(480);
  });
});

describe('excelSerialToEpochMs', () => {
  // Excel's epoch is 1899-12-31, and it wrongly believes 1900 was a leap year.
  // Serial 60 is 1900-02-29 -- a date that never existed.
  it('converts serial 61 (1900-03-01) correctly', () => {
    expect(excelSerialToEpochMs(61)).toBe(Date.UTC(1900, 2, 1));
  });

  it('converts serials below 61 using the pre-bug offset', () => {
    expect(excelSerialToEpochMs(1)).toBe(Date.UTC(1900, 0, 1));
  });

  it('converts a modern serial', () => {
    expect(excelSerialToEpochMs(45000)).toBe(Date.UTC(2023, 2, 15));
  });

  it('carries the fractional part as time of day', () => {
    expect(excelSerialToEpochMs(45000.5)).toBe(Date.UTC(2023, 2, 15, 12));
  });
});

describe('detectDateOrder', () => {
  it('proves dmy when a first component exceeds 12', () => {
    expect(detectDateOrder(['13/01/2024', '05/02/2024'])).toBe('dmy');
  });

  it('proves mdy when a second component exceeds 12', () => {
    expect(detectDateOrder(['01/13/2024', '02/05/2024'])).toBe('mdy');
  });

  it('reports a conflict when both are proven', () => {
    expect(detectDateOrder(['13/01/2024', '01/13/2024'])).toBe('conflict');
  });

  it('reports ambiguous when nothing proves it', () => {
    expect(detectDateOrder(['01/02/2024', '03/04/2024'])).toBe('ambiguous');
  });

  it('recognises ISO as unambiguous', () => {
    expect(detectDateOrder(['2024-01-13', '2024-02-05'])).toBe('iso');
  });
});

describe('toEpochMs', () => {
  it('passes Date objects straight through', () => {
    const d = new Date(Date.UTC(2024, 0, 15, 3, 30));
    expect(toEpochMs(d)).toBe(d.getTime());
  });

  it('reads dmy strings when told to', () => {
    expect(toEpochMs('15/01/2024', { order: 'dmy' }))
      .toBe(Date.UTC(2024, 0, 15));
  });

  it('reads the same string differently as mdy', () => {
    expect(toEpochMs('05/01/2024', { order: 'mdy' }))
      .toBe(Date.UTC(2024, 4, 1));
  });

  it('does not shift when the source is local', () => {
    expect(toEpochMs('15/01/2024 08:00', { order: 'dmy', sourceZone: 'local' }))
      .toBe(Date.UTC(2024, 0, 15, 8, 0));
  });

  it('shifts by +8 when the source is UTC', () => {
    expect(toEpochMs('15/01/2024 08:00', { order: 'dmy', sourceZone: 'utc' }))
      .toBe(Date.UTC(2024, 0, 15, 16, 0));
  });

  // Spec §9 hard rule -- shifting a pure date moves it to the wrong day.
  it('never shifts a date-only column even when marked UTC', () => {
    expect(toEpochMs('15/01/2024', { order: 'dmy', sourceZone: 'utc', dateOnly: true }))
      .toBe(Date.UTC(2024, 0, 15));
  });

  it('returns NaN for unparseable input', () => {
    expect(toEpochMs('not a date')).toBeNaN();
    expect(toEpochMs(null)).toBeNaN();
    expect(toEpochMs('')).toBeNaN();
  });
});

describe('formatMYT', () => {
  it('formats as DD/MM/YYYY HH:mm in Malaysian local time', () => {
    // 2024-01-15T00:00Z is 08:00 on the 15th in KL.
    expect(formatMYT(Date.UTC(2024, 0, 15, 0, 0))).toBe('15/01/2024 08:00');
  });

  it('formats date-only style without a time', () => {
    expect(formatMYT(Date.UTC(2024, 0, 15, 0, 0), 'date')).toBe('15/01/2024');
  });

  it('renders NaN as an em dash rather than "Invalid Date"', () => {
    expect(formatMYT(NaN)).toBe('—');
  });
});
```

- [ ] **Step 4: Run the tests and confirm they fail**

```bash
npm test -- malaysiaTime
```

Expected: every test fails — `Failed to resolve import "./malaysiaTime.js"`.

- [ ] **Step 5: Implement `malaysiaTime.js`**

Create `src/features/datastudio/time/malaysiaTime.js`. Implementation notes that matter:

- Excel serials: the standard conversion is `(serial - 25569) * 86400000`. For `serial < 61`, add one day back (`+ 86400000`) to undo the phantom 1900-02-29.
- `detectDateOrder` scans **every** value, tracking two booleans (`provenDmy`, `provenMdy`). ISO is detected by the `^\d{4}-\d{2}-\d{2}` shape and short-circuits.
- `toEpochMs` builds the result with `Date.UTC(...)` so that "no shift" genuinely means no shift — using `new Date(y, m, d)` would silently apply the *host machine's* timezone, which is the exact bug this module exists to prevent.
- `sourceZone: 'utc'` **adds** `MYT_OFFSET_MIN * 60000` — the stored value is a UTC wall-clock reading, and we want the MYT wall-clock reading of that same instant, which is 8 hours later. The direction is easy to invert by accident; the test above pins it, so make the test pass rather than reasoning from memory.
- `dateOnly: true` skips the shift unconditionally, before any other branch.
- `formatMYT` uses `Intl.DateTimeFormat('en-GB', { timeZone: 'Asia/Kuala_Lumpur', ... })` with `hour12: false`. `en-GB` already yields `DD/MM/YYYY`. Guard `Number.isNaN(epochMs)` first and return `'—'`.

- [ ] **Step 6: Run the tests and confirm they pass**

```bash
npm test -- malaysiaTime
```

Expected: all tests pass.

- [ ] **Step 7: Commit**

```bash
git add package.json package-lock.json vite.config.js src/features/datastudio/time/
git commit -m "Add Vitest and Malaysia time parsing"
```

---

## Task 2: `inferType.js`

Implements spec §7.2–7.4 — the value-level parsers and the per-column verdict.

**Files:**
- Create: `src/features/datastudio/profile/inferType.js`
- Test: `src/features/datastudio/profile/inferType.test.js`

**Interfaces:**
- Consumes: `toEpochMs`, `detectDateOrder` from Task 1.
- Produces:
  - `NULL_TOKENS: Set<string>` — `''`, `'-'`, `'–'`, `'—'`, `'n/a'`, `'na'`, `'null'`, `'nil'`, `'tbd'`, `'#n/a'`, `'#div/0!'`, `'#ref!'`, `'#value!'`, `'#name?'`
  - `isNullish(value) → boolean`
  - `parseNumberLike(value) → { ok: boolean, value: number, isPercent: boolean }`
  - `parseBooleanLike(value) → { ok: boolean, value: boolean }`
  - `inferType(values: unknown[], columnName: string) → ColumnVerdict`

  where `ColumnVerdict` is:

```js
{
  type: 'numeric'|'date'|'datetime'|'boolean'|'categorical'|'text'|'identifier'|'empty',
  role: 'measure'|'dimension'|'temporal'|'ignored',
  confidence: number,        // share of non-null values matching `type`, 0..1
  dateOrder: 'dmy'|'mdy'|'iso'|'conflict'|'ambiguous'|null,
  isPercent: boolean,
  nullCount: number,
  distinctCount: number,
  casualties: string[],      // up to 5 raw values that failed the chosen type
  casualtyCount: number,
}
```

- [ ] **Step 1: Write the failing tests**

Create `src/features/datastudio/profile/inferType.test.js`:

```js
import { describe, it, expect } from 'vitest';
import { isNullish, parseNumberLike, parseBooleanLike, inferType } from './inferType.js';

describe('isNullish', () => {
  it.each([['', true], ['-', true], ['N/A', true], ['#DIV/0!', true],
          ['NIL', true], ['  ', true], ['0', false], ['abc', false]])(
    'treats %j as nullish=%s', (input, expected) => {
      expect(isNullish(input)).toBe(expected);
    });
});

describe('parseNumberLike', () => {
  it('strips thousands separators', () => {
    expect(parseNumberLike('1,234.50')).toMatchObject({ ok: true, value: 1234.5 });
  });

  it('strips a currency prefix', () => {
    expect(parseNumberLike('RM 1,234')).toMatchObject({ ok: true, value: 1234 });
    expect(parseNumberLike('$1,234')).toMatchObject({ ok: true, value: 1234 });
  });

  it('reads percentages as fractions and flags them', () => {
    expect(parseNumberLike('45%')).toMatchObject({
      ok: true, value: 0.45, isPercent: true,
    });
  });

  it('reads accounting negatives', () => {
    expect(parseNumberLike('(1,234)')).toMatchObject({ ok: true, value: -1234 });
  });

  it('strips non-breaking spaces', () => {
    expect(parseNumberLike('1 234')).toMatchObject({ ok: true, value: 1234 });
  });

  // Spec §7.4 -- this rule protects employee IDs and cost centres.
  it('refuses values with leading zeros', () => {
    expect(parseNumberLike('007').ok).toBe(false);
  });

  it('accepts a bare zero', () => {
    expect(parseNumberLike('0')).toMatchObject({ ok: true, value: 0 });
  });

  it('rejects non-numeric text', () => {
    expect(parseNumberLike('pending').ok).toBe(false);
  });
});

describe('parseBooleanLike', () => {
  it.each([['yes', true], ['NO', false], ['true', true], ['Y', true], ['n', false]])(
    'reads %j', (input, expected) => {
      expect(parseBooleanLike(input)).toMatchObject({ ok: true, value: expected });
    });

  // Spec §7.3 -- 0/1 is numeric far more often than it is boolean.
  it('refuses 0 and 1', () => {
    expect(parseBooleanLike('1').ok).toBe(false);
    expect(parseBooleanLike('0').ok).toBe(false);
  });
});

describe('inferType', () => {
  it('types a clean numeric column as a measure', () => {
    const v = inferType(['10', '20', '30', '40'], 'Amount');
    expect(v).toMatchObject({ type: 'numeric', role: 'measure' });
  });

  // Spec §7.2 -- the 95% rule, and its casualties must be reported.
  it('still types numeric at 95% and reports the casualties', () => {
    const values = [...Array(19).fill('10'), 'pending'];
    const v = inferType(values, 'Amount');
    expect(v.type).toBe('numeric');
    expect(v.casualtyCount).toBe(1);
    expect(v.casualties).toContain('pending');
  });

  it('falls back to categorical below 95%', () => {
    const values = [...Array(17).fill('10'), 'a', 'b', 'c'];
    expect(inferType(values, 'Amount').type).toBe('categorical');
  });

  it('ignores nulls when computing the ratio', () => {
    const v = inferType(['10', '20', 'N/A', '', '-'], 'Amount');
    expect(v).toMatchObject({ type: 'numeric', nullCount: 3 });
  });

  it('types a low-cardinality string column as a dimension', () => {
    const v = inferType(['HR', 'IT', 'HR', 'Finance', 'IT'], 'Department');
    expect(v).toMatchObject({ type: 'categorical', role: 'dimension', distinctCount: 3 });
  });

  it('types high-cardinality free text as ignored', () => {
    const values = Array.from({ length: 200 }, (_, i) => `remark number ${i}`);
    expect(inferType(values, 'Remarks')).toMatchObject({ type: 'text', role: 'ignored' });
  });

  it('types a known boolean pair as a dimension', () => {
    const v = inferType(['Yes', 'No', 'Yes', 'No'], 'Active');
    expect(v).toMatchObject({ type: 'boolean', role: 'dimension' });
  });

  // Spec §7.3 -- summing employee IDs is meaningless.
  it('detects an identifier by name and uniqueness, and refuses to call it a measure', () => {
    const values = Array.from({ length: 50 }, (_, i) => String(1000 + i));
    const v = inferType(values, 'Employee ID');
    expect(v).toMatchObject({ type: 'identifier', role: 'ignored' });
  });

  it('detects a monotonic integer sequence as an identifier even without a matching name', () => {
    const values = Array.from({ length: 50 }, (_, i) => String(i + 1));
    expect(inferType(values, 'Seq').type).toBe('identifier');
  });

  it('types Date objects as datetime', () => {
    const values = [new Date(Date.UTC(2024, 0, 1)), new Date(Date.UTC(2024, 0, 2))];
    expect(inferType(values, 'Created')).toMatchObject({
      type: 'datetime', role: 'temporal',
    });
  });

  it('carries the detected date order onto the verdict', () => {
    const v = inferType(['13/01/2024', '05/02/2024'], 'Join Date');
    expect(v).toMatchObject({ role: 'temporal', dateOrder: 'dmy' });
  });

  it('leaves a conflicting date column as text for the user to resolve', () => {
    const v = inferType(['13/01/2024', '01/13/2024'], 'Join Date');
    expect(v).toMatchObject({ type: 'text', dateOrder: 'conflict' });
  });

  it('types an all-null column as empty', () => {
    expect(inferType(['', 'N/A', '-'], 'Blank')).toMatchObject({
      type: 'empty', role: 'ignored',
    });
  });

  it('preserves leading-zero codes as categorical, not numeric', () => {
    const v = inferType(['007', '008', '007', '009'], 'Cost Centre');
    expect(v.type).toBe('categorical');
  });
});
```

- [ ] **Step 2: Run the tests and confirm they fail**

```bash
npm test -- inferType
```

Expected: all fail on the missing module.

- [ ] **Step 3: Implement `inferType.js`**

Resolution order is **`boolean → datetime → numeric → categorical → text`, then the identifier override** (spec §7.3). Implementation notes:

- Normalise every value to a trimmed string first, stripping ` ` and zero-width characters (`​-‍﻿`), except `Date` objects which pass through untouched.
- Count non-null values once; each candidate type's confidence is `matches / nonNullCount`. First type at `>= 0.95` in resolution order wins.
- `parseNumberLike` must reject `/^0\d/` **before** stripping anything — that is the leading-zero rule, and it has to run on the raw string.
- Percent handling: divide by 100 and set `isPercent`. A column is `isPercent` only if *every* numeric match was a percent.
- `distinctCount` uses a `Set` over normalised strings.
- Categorical threshold: `distinctCount <= 50 || distinctCount / nonNullCount < 0.05`.
- Identifier override runs last and only over `numeric`, `text` or `categorical` verdicts: `distinctCount / nonNullCount > 0.95` **and** (`/id|no|code|ref|serial/i.test(columnName)` **or** the values form a monotonic integer sequence).
- `date` vs `datetime`: if no value carries a time component, the type is `date` and consumers must treat it as date-only (Task 1's `dateOnly` flag).
- `casualties` collects at most 5 raw values that failed the winning type; `casualtyCount` is the full count.

- [ ] **Step 4: Run the tests and confirm they pass**

```bash
npm test -- inferType
```

- [ ] **Step 5: Commit**

```bash
git add src/features/datastudio/profile/
git commit -m "Add column type inference"
```

---

## Task 3: `profileColumn.js` and `profileDataset.js`

Implements spec §7.1 — attaching stats and roles to each verdict.

**Files:**
- Create: `src/features/datastudio/profile/profileColumn.js`
- Create: `src/features/datastudio/profile/profileDataset.js`
- Test: `src/features/datastudio/profile/profileDataset.test.js`

**Interfaces:**
- Consumes: `inferType` from Task 2.
- Produces:
  - `profileColumn(values, columnName, index) → ColumnProfile` — the Task 2 verdict plus `{ name, index, nonNullRatio, topValues: [{value, count}], min, max, mean }` (numeric stats are `null` for non-numeric columns; `topValues` holds the 10 most frequent for dimensions and is `[]` otherwise)
  - `profileDataset(grid: { headers: string[], rows: unknown[][] }) → DatasetProfile` — `{ columns: ColumnProfile[], rowCount, topMeasure: string|null, primaryTemporal: string|null }`

`topMeasure` and `primaryTemporal` are defined in spec §10.5: the measure/temporal column with the highest `nonNullRatio`, ties broken by original column order. Later tasks depend on this definition; do not re-derive it elsewhere.

- [ ] **Step 1: Write the failing tests**

```js
import { describe, it, expect } from 'vitest';
import { profileDataset } from './profileDataset.js';

const grid = {
  headers: ['Department', 'Amount', 'Created', 'Notes'],
  rows: [
    ['HR', '100', '13/01/2024', 'first'],
    ['IT', '200', '14/01/2024', 'second'],
    ['HR', '300', '15/01/2024', ''],
  ],
};

describe('profileDataset', () => {
  it('profiles every column and counts the rows', () => {
    const p = profileDataset(grid);
    expect(p.rowCount).toBe(3);
    expect(p.columns.map((c) => c.name))
      .toEqual(['Department', 'Amount', 'Created', 'Notes']);
  });

  it('assigns the roles the canvas depends on', () => {
    const roles = Object.fromEntries(
      profileDataset(grid).columns.map((c) => [c.name, c.role]));
    expect(roles).toMatchObject({
      Department: 'dimension', Amount: 'measure', Created: 'temporal',
    });
  });

  it('computes numeric stats for measures only', () => {
    const [dept, amount] = profileDataset(grid).columns;
    expect(amount).toMatchObject({ min: 100, max: 300, mean: 200 });
    expect(dept.min).toBeNull();
  });

  it('ranks the top values for dimensions', () => {
    const dept = profileDataset(grid).columns[0];
    expect(dept.topValues[0]).toMatchObject({ value: 'HR', count: 2 });
  });

  it('picks topMeasure and primaryTemporal by non-null ratio', () => {
    const p = profileDataset(grid);
    expect(p.topMeasure).toBe('Amount');
    expect(p.primaryTemporal).toBe('Created');
  });

  it('breaks non-null ratio ties by column order', () => {
    const tied = {
      headers: ['B', 'A'],
      rows: [['1', '2'], ['3', '4']],
    };
    expect(profileDataset(tied).topMeasure).toBe('B');
  });

  it('returns null for topMeasure when there are no measures', () => {
    const noMeasures = { headers: ['Dept'], rows: [['HR'], ['IT']] };
    expect(profileDataset(noMeasures).topMeasure).toBeNull();
  });
});
```

- [ ] **Step 2: Run the tests and confirm they fail**

```bash
npm test -- profileDataset
```

- [ ] **Step 3: Implement both modules**

`profileColumn` calls `inferType`, then adds stats. `nonNullRatio = (total - nullCount) / total`, guarding division by zero for an empty grid. `profileDataset` walks columns by index (the grid is row-major, so transpose once rather than per column — this matters at 100k rows).

- [ ] **Step 4: Run the tests and confirm they pass**

- [ ] **Step 5: Commit**

```bash
git add src/features/datastudio/profile/
git commit -m "Add dataset profiling with roles and stats"
```

---

## Task 4: `detectHeader.js` and `parseWorkbook.js`

Implements spec §2.1 and the header-detection half of §12.

**Files:**
- Create: `src/features/datastudio/ingest/detectHeader.js`
- Create: `src/features/datastudio/ingest/parseWorkbook.js`
- Test: `src/features/datastudio/ingest/detectHeader.test.js`
- Modify: `package.json`

**Interfaces:**
- Consumes: `isNullish` from Task 2.
- Produces:
  - `detectHeader(rows: unknown[][]) → { headerIndex: number, confidence: number }` — `headerIndex` is `-1` when no plausible header exists
  - `toGrid(rows, headerIndex) → { headers: string[], rows: unknown[][] }` — de-duplicates repeated header names as `Name`, `Name (2)`, `Name (3)`
  - `parseWorkbook(arrayBuffer) → { sheets: [{ name, rows }] }`

- [ ] **Step 1: Install SheetJS**

```bash
npm install xlsx@^0.18
```

- [ ] **Step 2: Write the failing tests**

```js
import { describe, it, expect } from 'vitest';
import { detectHeader, toGrid } from './detectHeader.js';

describe('detectHeader', () => {
  it('finds a header on the first row', () => {
    const rows = [['Name', 'Amount'], ['a', 1], ['b', 2]];
    expect(detectHeader(rows).headerIndex).toBe(0);
  });

  it('skips a title banner above the header', () => {
    const rows = [
      ['IT Request Report 2024', null],
      [null, null],
      ['Name', 'Amount'],
      ['a', 1],
    ];
    expect(detectHeader(rows).headerIndex).toBe(2);
  });

  it('skips leading blank rows', () => {
    const rows = [[null, null], ['', ''], ['Name', 'Amount'], ['a', 1]];
    expect(detectHeader(rows).headerIndex).toBe(2);
  });

  it('returns -1 when no row looks like a header', () => {
    expect(detectHeader([[1, 2], [3, 4]]).headerIndex).toBe(-1);
  });

  it('returns -1 for an empty sheet', () => {
    expect(detectHeader([]).headerIndex).toBe(-1);
  });
});

describe('toGrid', () => {
  it('splits headers from data rows', () => {
    const rows = [['Name', 'Amount'], ['a', 1]];
    expect(toGrid(rows, 0)).toEqual({ headers: ['Name', 'Amount'], rows: [['a', 1]] });
  });

  it('de-duplicates repeated header names', () => {
    const rows = [['Name', 'Name', 'Name'], ['a', 'b', 'c']];
    expect(toGrid(rows, 0).headers).toEqual(['Name', 'Name (2)', 'Name (3)']);
  });

  it('names blank header cells by position', () => {
    const rows = [['Name', ''], ['a', 'b']];
    expect(toGrid(rows, 0).headers).toEqual(['Name', 'Column 2']);
  });
});
```

- [ ] **Step 3: Run the tests and confirm they fail**

- [ ] **Step 4: Implement both modules**

`detectHeader` scores the first 20 rows. A row scores well when its cells are mostly non-null **strings**, it has more filled cells than the rows above it, and the row *below* it differs in type shape (a header of strings over a body of numbers is the strongest signal). Pick the highest score above a floor of `0.5`; otherwise `-1`.

`parseWorkbook` calls SheetJS with the options that matter:

```js
import { read, utils } from 'xlsx';

export function parseWorkbook(arrayBuffer) {
  const wb = read(arrayBuffer, { type: 'array', cellDates: true, cellNF: false, cellText: false });
  return {
    sheets: wb.SheetNames.map((name) => ({
      name,
      rows: utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null, blankrows: false }),
    })),
  };
}
```

`cellDates: true` is what makes real Excel date cells arrive as `Date` objects, removing all D/M/Y ambiguity for them (spec §7.5). `header: 1` yields a row-major array of arrays. Do not use `sheet_to_json`'s object mode — it silently drops duplicate columns.

- [ ] **Step 5: Run the tests and confirm they pass**

- [ ] **Step 6: Commit**

```bash
git add package.json package-lock.json src/features/datastudio/ingest/
git commit -m "Add workbook parsing and header detection"
```

---

## Task 5: Worker, route, and the profile screen

First user-visible deliverable: drop a file, see it profiled. Implements spec §6.2, §6.5.

**Files:**
- Create: `src/features/datastudio/worker/studio.worker.js`
- Create: `src/features/datastudio/DataStudioContext.jsx`
- Create: `src/pages/DataStudioPage.jsx`
- Create: `src/styles/datastudio.css`
- Modify: `src/App.jsx` (route)
- Modify: `src/components/AppShell.jsx` (`NAV_ITEMS`)
- Modify: `src/main.jsx` (stylesheet import)
- Modify: `src/components/ui/Icons.jsx` (a `BarChart3` glyph)

**Interfaces:**
- Consumes: `parseWorkbook`, `toGrid`, `detectHeader`, `profileDataset`.
- Produces:
  - Worker message in: `{ type: 'parse', arrayBuffer, sheetName? }`
  - Worker message out: `{ type: 'progress', stage, pct }` | `{ type: 'parsed', sheets, activeSheet, grid, profile }` | `{ type: 'error', message }`
  - `useDataStudio()` hook exposing `{ stage, profile, grid, sheets, activeSheet, error, progress, importFile, selectSheet, setHeaderIndex, reset }`
  - `stage` is one of `'idle' | 'parsing' | 'profiled' | 'cleaning' | 'canvas'`

- [ ] **Step 1: Add the nav icon**

Add a `BarChart3` glyph to `src/components/ui/Icons.jsx`, matching the existing 24px stroke grid used by the icons already there. No icon package — that is a project rule.

- [ ] **Step 2: Write the worker**

```js
import { parseWorkbook } from '../ingest/parseWorkbook.js';
import { detectHeader, toGrid } from '../ingest/detectHeader.js';
import { profileDataset } from '../profile/profileDataset.js';

self.onmessage = (e) => {
  const { type, arrayBuffer, sheetName } = e.data;
  if (type !== 'parse') return;
  try {
    self.postMessage({ type: 'progress', stage: 'Reading workbook', pct: 10 });
    const { sheets } = parseWorkbook(arrayBuffer);
    if (!sheets.length) throw new Error('This workbook has no sheets.');

    const active = sheets.find((s) => s.name === sheetName) ?? sheets[0];
    self.postMessage({ type: 'progress', stage: 'Finding the header row', pct: 40 });
    const { headerIndex } = detectHeader(active.rows);
    if (headerIndex === -1) {
      throw new Error(`No header row found in "${active.name}". Pick one manually.`);
    }
    const grid = toGrid(active.rows, headerIndex);

    self.postMessage({ type: 'progress', stage: 'Profiling columns', pct: 70 });
    const profile = profileDataset(grid);

    self.postMessage({
      type: 'parsed',
      sheets: sheets.map((s) => s.name),
      activeSheet: active.name,
      headerIndex,
      grid,
      profile,
    });
  } catch (err) {
    self.postMessage({ type: 'error', message: err.message });
  }
};
```

- [ ] **Step 3: Write `DataStudioContext.jsx`**

Owns the worker lifetime and the stage machine. The worker is created with Vite's native syntax — no plugin needed:

```js
const worker = new Worker(new URL('./worker/studio.worker.js', import.meta.url), { type: 'module' });
```

Create it in a `useEffect` and `terminate()` in the cleanup. On `error`, set `stage: 'idle'` and surface the message — never leave the UI on a spinner (spec §12). Because this file exports a component (`DataStudioProvider`), the `useDataStudio` hook must live in its own file, `src/features/datastudio/useDataStudio.js` — exporting both from one file breaks Fast Refresh and fails lint. This is the same rule that put `initialsOf` in `src/utils/initials.js`.

- [ ] **Step 4: Write `DataStudioPage.jsx`**

Renders `<AppShell title="Data Studio" subtitle=…>` and switches on `stage`:

- `idle` — a drop zone plus a file input accepting `.xlsx,.xlsm,.csv`. Must handle both drag-drop and click-to-browse, and read via `file.arrayBuffer()`.
- `parsing` — progress bar driven by the worker's `pct` and `stage` label.
- `profiled` — a sheet picker (when the workbook has more than one), a header-row override control, and the profile table: one row per column showing name, inferred type, role badge, non-null share, distinct count, and — where `casualtyCount > 0` — an inline warning naming the count and previewing the offending values (spec §7.2).

Every type and role must be overridable via a `<select>` on its row; the override re-runs `profileColumn` for that column only.

- [ ] **Step 5: Wire the route, nav, and stylesheet**

- `src/App.jsx`: `const DataStudioPage = lazy(() => import('./pages/DataStudioPage'));` and a `<Route path="/data-studio" …>`. Wrap the `<Routes>` in `<Suspense fallback={…}>` if one is not already present. The lazy import is what keeps SheetJS and (later) ECharts out of the main bundle (spec §6.3).
- `src/components/AppShell.jsx`: add `{ to: '/data-studio', label: 'Data Studio', icon: BarChart3 }` to `NAV_ITEMS`.
- `src/main.jsx`: import `./styles/datastudio.css` **after** `./styles/shell.css`.

- [ ] **Step 6: Verify in the browser**

Start the dev server via the preview tool (never `npm run dev` in Bash), sign in, open `/data-studio`, and import a real `.xlsx`. Confirm: the progress bar advances, the profile table lists every column with a sensible type and role, dark mode looks right, and the console is clean.

- [ ] **Step 7: Confirm the bundle split held**

```bash
npm run build
```

Expected: a separate chunk containing `xlsx`, and an `index` chunk no larger than before this task.

- [ ] **Step 8: Commit**

```bash
git add src/ package.json
git commit -m "Add Data Studio route with worker-backed import and profiling"
```

---

# PHASE 2 — Cleaning and Malaysia time

---

## Task 6: `cleanOps.js`

Implements spec §8.2.

**Files:**
- Create: `src/features/datastudio/clean/cleanOps.js`
- Test: `src/features/datastudio/clean/cleanOps.test.js`

**Interfaces:**
- Consumes: `isNullish`, `parseNumberLike`, `parseBooleanLike` (Task 2); `toEpochMs` (Task 1).
- Produces one pure function per op, each `(values: unknown[], params) → unknown[]` except where noted:
  `trimWhitespace`, `normalizeNulls`, `parseNumber`, `unifyCase`, `mergeCategories`, `parseDate`, `castType`, and three whole-grid ops `dropEmptyColumns`, `dropEmptyRows`, `dedupeRows` with signature `(grid) → grid`.
- Also produces `categoryKey(value) → string` and `clusterCategories(values) → [{ canonical, variants, count }]`.

- [ ] **Step 1: Write the failing tests**

```js
import { describe, it, expect } from 'vitest';
import {
  trimWhitespace, normalizeNulls, parseNumber, unifyCase,
  categoryKey, clusterCategories, mergeCategories, dedupeRows, dropEmptyColumns,
} from './cleanOps.js';

describe('trimWhitespace', () => {
  it('trims, collapses internal runs, and strips non-breaking spaces', () => {
    expect(trimWhitespace(['  a  b  ', 'c d']))
      .toEqual(['a b', 'c d']);
  });

  it('leaves non-strings alone', () => {
    const d = new Date();
    expect(trimWhitespace([1, d])).toEqual([1, d]);
  });
});

describe('normalizeNulls', () => {
  it('maps every placeholder token to null', () => {
    expect(normalizeNulls(['a', '-', 'N/A', '#REF!', '']))
      .toEqual(['a', null, null, null, null]);
  });
});

describe('parseNumber', () => {
  it('coerces number-like strings and nulls the rest', () => {
    expect(parseNumber(['1,234', 'RM 5', 'nope'])).toEqual([1234, 5, null]);
  });
});

describe('unifyCase', () => {
  it('maps case variants onto the most frequent spelling', () => {
    expect(unifyCase(['HR', 'hr', 'HR', 'Hr'])).toEqual(['HR', 'HR', 'HR', 'HR']);
  });
});

describe('categoryKey', () => {
  it('ignores case, padding, internal runs and punctuation', () => {
    expect(categoryKey('  Kuala  Lumpur ')).toBe(categoryKey('kuala lumpur'));
    expect(categoryKey('PMW-SS')).toBe(categoryKey('pmw ss'));
  });
});

describe('clusterCategories', () => {
  it('groups variants under the most frequent spelling', () => {
    const clusters = clusterCategories([
      'Kuala Lumpur', 'kuala lumpur', 'Kuala  Lumpur ', 'Kuala Lumpur', 'Penang',
    ]);
    const kl = clusters.find((c) => c.canonical === 'Kuala Lumpur');
    expect(kl.count).toBe(4);
    expect(kl.variants).toHaveLength(3);
  });

  // Spec §3 and §8.2 -- these must NOT be merged.
  it('never merges genuinely different values that are merely similar', () => {
    const clusters = clusterCategories(['Dept A', 'Dept B']);
    expect(clusters).toHaveLength(2);
  });
});

describe('mergeCategories', () => {
  it('rewrites values to their canonical spelling', () => {
    expect(mergeCategories(
      ['HR', 'hr', 'IT'],
      { map: { hr: 'HR' } },
    )).toEqual(['HR', 'HR', 'IT']);
  });
});

describe('dedupeRows', () => {
  it('removes exact duplicate rows and keeps the first', () => {
    const grid = { headers: ['a'], rows: [['x'], ['y'], ['x']] };
    expect(dedupeRows(grid).rows).toEqual([['x'], ['y']]);
  });
});

describe('dropEmptyColumns', () => {
  it('removes columns that are entirely null', () => {
    const grid = { headers: ['a', 'blank'], rows: [['x', null], ['y', null]] };
    expect(dropEmptyColumns(grid)).toEqual({ headers: ['a'], rows: [['x'], ['y']] });
  });
});
```

- [ ] **Step 2: Run the tests and confirm they fail**

- [ ] **Step 3: Implement `cleanOps.js`**

`categoryKey` is `lowercase → trim → collapse whitespace → strip non-alphanumerics → collapse again`. `clusterCategories` groups by that key, picks the highest-count original spelling as `canonical`, and returns clusters of size ≥ 2 only. Every op must be pure and must not mutate its input array.

- [ ] **Step 4: Run the tests and confirm they pass**

- [ ] **Step 5: Commit**

```bash
git add src/features/datastudio/clean/
git commit -m "Add cleaning operations"
```

---

## Task 7: `proposeCleanPlan.js`

Implements spec §8.1.

**Files:**
- Create: `src/features/datastudio/clean/proposeCleanPlan.js`
- Test: `src/features/datastudio/clean/proposeCleanPlan.test.js`

**Interfaces:**
- Consumes: `DatasetProfile` (Task 3), `clusterCategories` (Task 6).
- Produces: `proposeCleanPlan(profile, grid) → CleanStep[]` where

```js
{ id: string, column: string|null, op: string, params: object,
  confidence: 'high'|'medium'|'low', affectedCount: number,
  preview: string, enabled: boolean }
```

`enabled` starts `true` for `confidence: 'high'` and `false` otherwise (spec §8.1). `column` is `null` for whole-grid ops.

- [ ] **Step 1: Write the failing tests**

```js
import { describe, it, expect } from 'vitest';
import { proposeCleanPlan } from './proposeCleanPlan.js';
import { profileDataset } from '../profile/profileDataset.js';

function planFor(grid) {
  return proposeCleanPlan(profileDataset(grid), grid);
}

describe('proposeCleanPlan', () => {
  it('proposes trimming when padded values exist, with a real count', () => {
    const grid = { headers: ['Dept'], rows: [[' HR '], ['IT '], ['HR']] };
    const step = planFor(grid).find((s) => s.op === 'trimWhitespace');
    expect(step).toMatchObject({ column: 'Dept', affectedCount: 2, enabled: true });
  });

  it('proposes a category merge and says what it will merge', () => {
    const grid = {
      headers: ['City'],
      rows: [['Kuala Lumpur'], ['kuala lumpur'], ['Penang']],
    };
    const step = planFor(grid).find((s) => s.op === 'mergeCategories');
    expect(step.preview).toContain('Kuala Lumpur');
    expect(step.affectedCount).toBe(1);
  });

  it('proposes dropping an all-null column', () => {
    const grid = { headers: ['a', 'blank'], rows: [['x', null], ['y', null]] };
    expect(planFor(grid).some((s) => s.op === 'dropEmptyColumns')).toBe(true);
  });

  it('proposes deduping only when duplicates exist', () => {
    const dupes = { headers: ['a'], rows: [['x'], ['x']] };
    const clean = { headers: ['a'], rows: [['x'], ['y']] };
    expect(planFor(dupes).some((s) => s.op === 'dedupeRows')).toBe(true);
    expect(planFor(clean).some((s) => s.op === 'dedupeRows')).toBe(false);
  });

  it('proposes nothing for an already-clean grid', () => {
    const grid = { headers: ['Dept', 'Amount'], rows: [['HR', 1], ['IT', 2]] };
    expect(planFor(grid)).toEqual([]);
  });

  it('leaves low-confidence steps disabled by default', () => {
    const grid = { headers: ['City'], rows: [['KL'], ['kl'], ['Penang']] };
    const step = planFor(grid).find((s) => s.confidence !== 'high');
    if (step) expect(step.enabled).toBe(false);
  });

  it('gives every step a unique id', () => {
    const grid = { headers: ['A', 'B'], rows: [[' x ', ' y '], ['x', 'y']] };
    const ids = planFor(grid).map((s) => s.id);
    expect(new Set(ids).size).toBe(ids.length);
  });
});
```

- [ ] **Step 2: Run the tests and confirm they fail**

- [ ] **Step 3: Implement `proposeCleanPlan.js`**

Confidence rules: `trimWhitespace`, `normalizeNulls`, `dropEmptyColumns`, `dropEmptyRows` and `castType` are `high` (mechanical and safe). `mergeCategories` and `unifyCase` are `high` when the cluster's variants differ **only** by case and whitespace, `medium` when punctuation differs. `dedupeRows` is `medium` — dropping rows is not obviously safe. Nothing is `low` yet; the level exists for the manual merges added in Task 9.

Steps must be ordered so that later ops see earlier ops' output: whitespace and nulls first, then casing and merges, then numeric/date coercion, then whole-grid row/column ops.

- [ ] **Step 4: Run the tests and confirm they pass**

- [ ] **Step 5: Commit**

```bash
git add src/features/datastudio/clean/
git commit -m "Add clean plan proposal"
```

---

## Task 8: `applyCleanPlan.js` and `dataset.js`

Implements spec §6.1 and §8.1 — the plan produces the typed columnar dataset every later phase reads.

**Files:**
- Create: `src/features/datastudio/clean/applyCleanPlan.js`
- Create: `src/features/datastudio/engine/dataset.js`
- Test: `src/features/datastudio/clean/applyCleanPlan.test.js`

**Interfaces:**
- Consumes: `cleanOps` (Task 6), `toEpochMs` (Task 1), `DatasetProfile` (Task 3).
- Produces:
  - `buildDataset({ headers, columns, profile }) → Dataset`
  - `applyCleanPlan(grid, plan, profile) → Dataset`
  - `Dataset` shape — **every later task depends on this exactly**:

```js
{
  rowCount: number,
  columns: [{
    name, type, role,
    values,          // Float64Array | Int32Array | Uint8Array | string[]
    dictionary,      // string[] for categorical, else null
    isPercent, dateOnly, sourceZone,
  }],
  byName: Map<string, ColumnIndex>,
}
```

Null encodings, per spec §6.1: `NaN` for numeric and temporal, `-1` for categorical codes, `2` for boolean, `null` for text.

- [ ] **Step 1: Write the failing tests**

```js
import { describe, it, expect } from 'vitest';
import { applyCleanPlan } from './applyCleanPlan.js';
import { profileDataset } from '../profile/profileDataset.js';
import { proposeCleanPlan } from './proposeCleanPlan.js';

const grid = {
  headers: ['Dept', 'Amount', 'Created'],
  rows: [
    [' HR ', '1,000', '13/01/2024'],
    ['hr', '2,000', '14/01/2024'],
    ['IT', 'nope', '15/01/2024'],
  ],
};

function build(enabledOverride) {
  const profile = profileDataset(grid);
  let plan = proposeCleanPlan(profile, grid);
  if (enabledOverride) plan = plan.map(enabledOverride);
  return applyCleanPlan(grid, plan, profile);
}

describe('applyCleanPlan', () => {
  it('stores numerics as a Float64Array with NaN for failures', () => {
    const col = build().columns.find((c) => c.name === 'Amount');
    expect(col.values).toBeInstanceOf(Float64Array);
    expect(col.values[0]).toBe(1000);
    expect(col.values[2]).toBeNaN();
  });

  it('dictionary-encodes categoricals', () => {
    const col = build().columns.find((c) => c.name === 'Dept');
    expect(col.values).toBeInstanceOf(Int32Array);
    expect(col.dictionary).toContain('HR');
    // ' HR ' and 'hr' both normalise onto the same code.
    expect(col.values[0]).toBe(col.values[1]);
  });

  it('stores temporal columns as epoch ms', () => {
    const col = build().columns.find((c) => c.name === 'Created');
    expect(col.values).toBeInstanceOf(Float64Array);
    expect(col.values[0]).toBe(Date.UTC(2024, 0, 13));
  });

  it('exposes columns by name', () => {
    expect(build().byName.get('Amount')).toBeDefined();
  });

  // Spec §8.1 -- the plan is data, so disabling a step must change the output.
  it('respects disabled steps', () => {
    const withMerge = build();
    const withoutMerge = build((s) =>
      s.op === 'mergeCategories' || s.op === 'unifyCase' ? { ...s, enabled: false } : s);
    const a = withMerge.columns.find((c) => c.name === 'Dept');
    const b = withoutMerge.columns.find((c) => c.name === 'Dept');
    expect(b.dictionary.length).toBeGreaterThan(a.dictionary.length);
  });

  it('is idempotent -- applying the same plan twice gives the same result', () => {
    const a = build();
    const b = build();
    expect(Array.from(a.columns[1].values)).toEqual(Array.from(b.columns[1].values));
  });

  it('never mutates the input grid', () => {
    const snapshot = JSON.stringify(grid);
    build();
    expect(JSON.stringify(grid)).toBe(snapshot);
  });
});
```

- [ ] **Step 2: Run the tests and confirm they fail**

- [ ] **Step 3: Implement both modules**

`applyCleanPlan` transposes the grid to columns once, runs enabled steps in plan order, then calls `buildDataset` to encode into typed arrays. Whole-grid ops (`dropEmptyRows`, `dedupeRows`, `dropEmptyColumns`) run against the row-major form before transposition.

The input grid must never be mutated — the whole non-destructive model in spec §8.1 depends on being able to re-run from raw.

- [ ] **Step 4: Run the tests and confirm they pass**

- [ ] **Step 5: Commit**

```bash
git add src/features/datastudio/
git commit -m "Add clean plan application and the columnar dataset"
```

---

## Task 9: The clean review screen

Implements spec §8.1, §4.3, §7.5 — the checklist, the UTC toggle, the date-order banner, manual merges.

**Files:**
- Create: `src/features/datastudio/clean/CleanReview.jsx`
- Modify: `src/pages/DataStudioPage.jsx` (the `cleaning` stage)
- Modify: `src/features/datastudio/DataStudioContext.jsx`
- Modify: `src/styles/datastudio.css`

**Interfaces:**
- Consumes: `proposeCleanPlan`, `applyCleanPlan`, `clusterCategories`.
- Produces: context additions `{ plan, setStepEnabled, setColumnZone, setColumnDateOrder, addManualMerge, dataset, commitClean }`.

- [ ] **Step 1: Build the checklist**

One row per step: a checkbox bound to `enabled`, the step's `preview` text, and its `affectedCount` as a badge. Toggling re-runs `applyCleanPlan` from raw — never mutate in place. Group rows by column, with whole-grid ops in their own group.

- [ ] **Step 2: Build the date-order banner**

For any temporal column whose `dateOrder` is `'ambiguous'` or `'conflict'`, show a banner naming the column, rendering the **first five values parsed both ways** side by side, and offering a one-click flip. Ambiguous columns default to `dmy` (spec §7.5); conflicting columns start unresolved and the user must choose.

- [ ] **Step 3: Build the timezone control**

Per temporal column, a two-option control: "Already Malaysia time" (default) and "Stored as UTC — shift +8". Setting it re-runs the apply step with `sourceZone`. Columns whose type is `date` (not `datetime`) must show the control **disabled** with the explanation that date-only values are never shifted — that is the spec §9 hard rule, and hiding the control entirely would leave users wondering.

- [ ] **Step 4: Build the manual merge control**

For each categorical column, list its distinct values with counts and allow selecting two or more to merge, choosing which spelling survives. This adds a `mergeCategories` step with `confidence: 'low'` and `enabled: true` (the user asked for it explicitly). It must be removable.

- [ ] **Step 5: Add the commit action**

A primary button that freezes the dataset and moves `stage` to `'canvas'`. Show the resulting row and column counts next to it.

- [ ] **Step 6: Verify in the browser**

Import a deliberately messy file — padded values, mixed casing, an `N/A` column, a duplicate row, an ambiguous date column. Confirm every proposal appears with a correct count, toggling a step changes the preview, the date banner shows both readings, and the timezone control is disabled for date-only columns.

- [ ] **Step 7: Commit**

```bash
git add src/
git commit -m "Add the clean review screen"
```

---

# PHASE 3 — The aggregation engine

> This phase ships **no UI**. That is deliberate (spec §14): the engine is proven by tests before any chart exists that could render a wrong number attractively.

---

## Task 10: `filterMask.js`

Implements spec §10.3 — including the self-exclusion rule, which is the single most important behaviour in the feature.

**Files:**
- Create: `src/features/datastudio/engine/filterMask.js`
- Test: `src/features/datastudio/engine/filterMask.test.js`

**Interfaces:**
- Consumes: `Dataset` (Task 8).
- Produces:
  - `buildMask(dataset, filters) → Uint8Array` — `1` keeps a row
  - `maskFor(dataset, globalFilters, selection, tileId) → Uint8Array` where `selection` is `{ sourceTileId, column, values } | null`
  - `createMaskCache() → { get(dataset, globalFilters, selection, tileId) }` — memoised by filter signature

A filter is `{ column, kind: 'in'|'range', values?: string[], min?: number, max?: number }`.

- [ ] **Step 1: Write the failing tests**

```js
import { describe, it, expect } from 'vitest';
import { buildMask, maskFor, createMaskCache } from './filterMask.js';
import { applyCleanPlan } from '../clean/applyCleanPlan.js';
import { profileDataset } from '../profile/profileDataset.js';

const grid = {
  headers: ['Dept', 'Amount'],
  rows: [['HR', '10'], ['IT', '20'], ['HR', '30'], ['Finance', '40']],
};
const ds = applyCleanPlan(grid, [], profileDataset(grid));
const count = (m) => m.reduce((a, b) => a + b, 0);

describe('buildMask', () => {
  it('keeps every row when there are no filters', () => {
    expect(count(buildMask(ds, []))).toBe(4);
  });

  it('filters a categorical column by membership', () => {
    const m = buildMask(ds, [{ column: 'Dept', kind: 'in', values: ['HR'] }]);
    expect(count(m)).toBe(2);
    expect(Array.from(m)).toEqual([1, 0, 1, 0]);
  });

  it('accepts multiple values on one filter', () => {
    const m = buildMask(ds, [{ column: 'Dept', kind: 'in', values: ['HR', 'IT'] }]);
    expect(count(m)).toBe(3);
  });

  it('ANDs separate filters together', () => {
    const m = buildMask(ds, [
      { column: 'Dept', kind: 'in', values: ['HR'] },
      { column: 'Amount', kind: 'range', min: 20, max: 100 },
    ]);
    expect(count(m)).toBe(1);
  });

  it('excludes NaN from range filters', () => {
    const withNull = { headers: ['n'], rows: [['1'], ['nope'], ['3']] };
    const d2 = applyCleanPlan(withNull, [], profileDataset(withNull));
    expect(count(buildMask(d2, [{ column: 'n', kind: 'range', min: 0, max: 10 }])))
      .toBe(2);
  });
});

describe('maskFor -- the self-exclusion rule (spec §10.3)', () => {
  const selection = { sourceTileId: 'tile_1', column: 'Dept', values: ['HR'] };

  it('does not filter the tile that originated the selection', () => {
    expect(count(maskFor(ds, [], selection, 'tile_1'))).toBe(4);
  });

  it('filters every other tile', () => {
    expect(count(maskFor(ds, [], selection, 'tile_2'))).toBe(2);
  });

  it('still applies global filters to the source tile', () => {
    const globals = [{ column: 'Amount', kind: 'range', min: 0, max: 25 }];
    expect(count(maskFor(ds, globals, selection, 'tile_1'))).toBe(2);
  });

  it('behaves normally when there is no selection', () => {
    expect(count(maskFor(ds, [], null, 'tile_1'))).toBe(4);
  });
});

describe('createMaskCache', () => {
  it('returns the identical array for identical inputs', () => {
    const cache = createMaskCache();
    const a = cache.get(ds, [], null, 'tile_1');
    const b = cache.get(ds, [], null, 'tile_1');
    expect(a).toBe(b);
  });

  it('shares one array between tiles that are not the selection source', () => {
    const cache = createMaskCache();
    const sel = { sourceTileId: 'tile_1', column: 'Dept', values: ['HR'] };
    expect(cache.get(ds, [], sel, 'tile_2')).toBe(cache.get(ds, [], sel, 'tile_3'));
  });

  it('returns a different array once the filters change', () => {
    const cache = createMaskCache();
    const a = cache.get(ds, [], null, 'tile_1');
    const b = cache.get(ds, [{ column: 'Dept', kind: 'in', values: ['HR'] }], null, 'tile_1');
    expect(a).not.toBe(b);
  });
});
```

- [ ] **Step 2: Run the tests and confirm they fail**

- [ ] **Step 3: Implement `filterMask.js`**

`maskFor` is the rule from spec §10.3:

```js
export function maskFor(dataset, globalFilters, selection, tileId) {
  const applySelection = selection && selection.sourceTileId !== tileId;
  const filters = applySelection
    ? [...globalFilters, { column: selection.column, kind: 'in', values: selection.values }]
    : globalFilters;
  return buildMask(dataset, filters);
}
```

The cache keys on a signature string built from the dataset identity, the global filters, and — critically — whether the selection applies, **not** on the tile id. That is what makes every non-source tile share one array (the third test above pins it).

Categorical filtering resolves value strings to dictionary codes **once** before the row loop, then compares integers.

- [ ] **Step 4: Run the tests and confirm they pass**

- [ ] **Step 5: Commit**

```bash
git add src/features/datastudio/engine/
git commit -m "Add filter masks with source-tile self-exclusion"
```

---

## Task 11: `aggregate.js`

Implements spec §10.2.

**Files:**
- Create: `src/features/datastudio/engine/aggregate.js`
- Test: `src/features/datastudio/engine/aggregate.test.js`

**Interfaces:**
- Consumes: `Dataset` (Task 8), masks (Task 10).
- Produces:
  - `aggregate(dataset, mask, spec) → { categories: string[], series: [{ name: string, data: number[] }] }`
  - `truncateDate(epochMs, unit: 'day'|'month'|'quarter'|'year') → number`
  - `chooseTruncation(minMs, maxMs) → 'day'|'month'|'quarter'`
  - `binNumeric(values, mask) → { edges: number[], labels: string[] }` — Freedman–Diaconis

Aggregations: `sum`, `avg`, `count`, `countDistinct`, `min`, `max`, `median`.

- [ ] **Step 1: Write the failing tests**

```js
import { describe, it, expect } from 'vitest';
import { aggregate, truncateDate, chooseTruncation } from './aggregate.js';
import { applyCleanPlan } from '../clean/applyCleanPlan.js';
import { profileDataset } from '../profile/profileDataset.js';

const grid = {
  headers: ['Dept', 'Entity', 'Amount'],
  rows: [
    ['HR', 'pmw', '10'], ['IT', 'pmw', '20'], ['HR', 'pmw-ss', '30'],
    ['IT', 'pmw', '40'], ['HR', 'pmw', 'nope'],
  ],
};
const ds = applyCleanPlan(grid, [], profileDataset(grid));
const ALL = new Uint8Array(5).fill(1);

const spec = (over = {}) => ({
  chart: 'bar',
  encoding: { x: { column: 'Dept' }, y: [{ column: 'Amount', agg: 'sum' }], series: null },
  sort: { by: 'y', dir: 'desc' }, limit: 20, ...over,
});

describe('aggregate', () => {
  it('groups by a dimension and sums a measure', () => {
    const r = aggregate(ds, ALL, spec());
    expect(r.categories).toEqual(['IT', 'HR']);
    expect(r.series[0].data).toEqual([60, 40]);
  });

  it('skips NaN when summing rather than producing NaN', () => {
    const r = aggregate(ds, ALL, spec());
    expect(r.series[0].data.every(Number.isFinite)).toBe(true);
  });

  it('counts rows including those with a null measure', () => {
    const r = aggregate(ds, ALL, spec({
      encoding: { x: { column: 'Dept' }, y: [{ column: 'Amount', agg: 'count' }], series: null },
    }));
    expect(Object.fromEntries(r.categories.map((c, i) => [c, r.series[0].data[i]])))
      .toMatchObject({ HR: 3, IT: 2 });
  });

  it('averages over non-null values only', () => {
    const r = aggregate(ds, ALL, spec({
      encoding: { x: { column: 'Dept' }, y: [{ column: 'Amount', agg: 'avg' }], series: null },
    }));
    const byCat = Object.fromEntries(r.categories.map((c, i) => [c, r.series[0].data[i]]));
    expect(byCat.HR).toBe(20); // (10 + 30) / 2, not / 3
  });

  it('honours the row mask', () => {
    const mask = new Uint8Array([1, 0, 1, 0, 0]);
    const r = aggregate(ds, mask, spec());
    expect(r.categories).toEqual(['HR']);
    expect(r.series[0].data).toEqual([40]);
  });

  it('splits into one series per series-column value', () => {
    const r = aggregate(ds, ALL, spec({
      encoding: {
        x: { column: 'Dept' },
        y: [{ column: 'Amount', agg: 'sum' }],
        series: { column: 'Entity' },
      },
    }));
    expect(r.series.map((s) => s.name).sort()).toEqual(['pmw', 'pmw-ss']);
    expect(r.series[0].data).toHaveLength(r.categories.length);
  });

  it('pads absent series/category combinations with 0', () => {
    const r = aggregate(ds, ALL, spec({
      encoding: {
        x: { column: 'Dept' },
        y: [{ column: 'Amount', agg: 'sum' }],
        series: { column: 'Entity' },
      },
    }));
    const ss = r.series.find((s) => s.name === 'pmw-ss');
    expect(ss.data).toContain(0);
  });

  it('applies top-N and folds the remainder into Other', () => {
    const r = aggregate(ds, ALL, spec({ limit: 1 }));
    expect(r.categories).toEqual(['IT', 'Other']);
    expect(r.series[0].data).toEqual([60, 40]);
  });

  it('sorts ascending when told to', () => {
    const r = aggregate(ds, ALL, spec({ sort: { by: 'y', dir: 'asc' } }));
    expect(r.categories).toEqual(['HR', 'IT']);
  });

  it('returns empty results for an all-zero mask rather than throwing', () => {
    const r = aggregate(ds, new Uint8Array(5), spec());
    expect(r.categories).toEqual([]);
    expect(r.series[0].data).toEqual([]);
  });
});

describe('truncateDate', () => {
  const t = Date.UTC(2024, 4, 17, 13, 45); // 17 May 2024
  it('truncates to day', () => expect(truncateDate(t, 'day')).toBe(Date.UTC(2024, 4, 17)));
  it('truncates to month', () => expect(truncateDate(t, 'month')).toBe(Date.UTC(2024, 4, 1)));
  it('truncates to quarter', () => expect(truncateDate(t, 'quarter')).toBe(Date.UTC(2024, 3, 1)));
  it('truncates to year', () => expect(truncateDate(t, 'year')).toBe(Date.UTC(2024, 0, 1)));
});

describe('chooseTruncation', () => {
  const DAY = 86400000;
  it('uses day below 90 days', () => {
    expect(chooseTruncation(0, 60 * DAY)).toBe('day');
  });
  it('uses month below 3 years', () => {
    expect(chooseTruncation(0, 500 * DAY)).toBe('month');
  });
  it('uses quarter beyond 3 years', () => {
    expect(chooseTruncation(0, 2000 * DAY)).toBe('quarter');
  });
});
```

- [ ] **Step 2: Run the tests and confirm they fail**

- [ ] **Step 3: Implement `aggregate.js`**

Group-by walks rows once, keyed by the x column's dictionary code (or truncated date, or bin index). Accumulate per `(category, series)` pair into a `Map`, then materialise the dense matrix, padding missing pairs with `0`.

`avg` divides by the count of **non-null** values, not the group size — the test pins this, and it is the most common aggregation bug.

Top-N: sort by the first series' value, keep `limit`, sum the rest into a trailing `'Other'` category. `'Other'` must never be sorted into the middle.

- [ ] **Step 4: Run the tests and confirm they pass**

- [ ] **Step 5: Run the whole suite — everything so far must still be green**

```bash
npm test
```

- [ ] **Step 6: Commit**

```bash
git add src/features/datastudio/engine/
git commit -m "Add the aggregation engine"
```

---

# PHASE 4 — The canvas

---

## Task 12: `echartsTheme.js` and `EChart.jsx`

Implements spec §10.6, §10.7, §6.3.

**Files:**
- Create: `src/features/datastudio/canvas/echartsTheme.js`
- Create: `src/features/datastudio/canvas/EChart.jsx`
- Modify: `package.json`

**Interfaces:**
- Produces:
  - `buildTheme() → object` — reads live CSS custom properties
  - `registerStudioTheme() → string` — registers and returns the theme name
  - `<EChart option onEvents className />`

- [ ] **Step 1: Install ECharts**

```bash
npm install echarts@^5
```

- [ ] **Step 2: Write the tree-shaken registration**

Create `src/features/datastudio/canvas/echartsCore.js` — importing the `echarts` umbrella anywhere in this feature defeats the whole bundle strategy (spec §6.3):

```js
import * as echarts from 'echarts/core';
import { BarChart, LineChart, PieChart, ScatterChart, HeatmapChart } from 'echarts/charts';
import {
  GridComponent, TooltipComponent, LegendComponent,
  DataZoomComponent, TitleComponent, VisualMapComponent,
} from 'echarts/components';
import { CanvasRenderer } from 'echarts/renderers';

echarts.use([
  BarChart, LineChart, PieChart, ScatterChart, HeatmapChart,
  GridComponent, TooltipComponent, LegendComponent,
  DataZoomComponent, TitleComponent, VisualMapComponent,
  CanvasRenderer,
]);

export default echarts;
```

- [ ] **Step 3: Write `echartsTheme.js`**

`buildTheme()` reads `getComputedStyle(document.documentElement).getPropertyValue('--it-ink')` and friends, trims them, and maps them onto ECharts' theme keys — `backgroundColor` from `--it-panel`, axis lines from `--it-line`, labels from `--it-ink-soft`, title from `--it-ink`.

The categorical `color` array must be colour-blind-safe and legible on both `--it-panel` values. Derive it from the brand hue and verify it in Step 6 rather than assuming — this is the spec §10.7 requirement.

Animations must be disabled when `window.matchMedia('(prefers-reduced-motion: reduce)').matches`.

- [ ] **Step 4: Write `EChart.jsx`**

```jsx
import { useEffect, useRef } from 'react';
import echarts from './echartsCore.js';
import { registerStudioTheme } from './echartsTheme.js';
import { useTheme } from '../../../context/ThemeContext.jsx';

export default function EChart({ option, onEvents, className }) {
  const hostRef = useRef(null);
  const chartRef = useRef(null);
  const { isDarkMode } = useTheme();

  useEffect(() => {
    const name = registerStudioTheme();
    const chart = echarts.init(hostRef.current, name, { renderer: 'canvas' });
    chartRef.current = chart;
    const ro = new ResizeObserver(() => chart.resize());
    ro.observe(hostRef.current);
    return () => { ro.disconnect(); chart.dispose(); chartRef.current = null; };
  }, [isDarkMode]);   // theme flip rebuilds the instance -- themes are baked at init

  useEffect(() => {
    chartRef.current?.setOption(option, { notMerge: true });
  }, [option]);

  useEffect(() => {
    const chart = chartRef.current;
    if (!chart || !onEvents) return undefined;
    const entries = Object.entries(onEvents);
    entries.forEach(([evt, handler]) => chart.on(evt, handler));
    return () => entries.forEach(([evt, handler]) => chart.off(evt, handler));
  }, [onEvents]);

  return <div ref={hostRef} className={className} />;
}
```

`notMerge: true` is required — merging leaves stale series behind when a tile's series count shrinks. Re-initialising on theme flip is deliberate: ECharts bakes the theme at `init` and cannot swap it live.

`onEvents` must be memoised by callers with `useMemo`, or the effect re-binds every render.

- [ ] **Step 5: Verify the bundle**

```bash
npm run build
```

Expected: the `echarts` code sits in the lazy `DataStudioPage` chunk, not in `index`. If it landed in `index`, something imported the page eagerly — fix that before continuing.

- [ ] **Step 6: Verify theming in the browser**

Render a throwaway bar chart on the page, toggle dark mode, and confirm the axes, labels and background all follow the tokens with no hardcoded colours. Check the series palette against a colour-blindness simulator.

- [ ] **Step 7: Commit**

```bash
git add package.json package-lock.json src/features/datastudio/canvas/
git commit -m "Add themed ECharts wrapper"
```

---

## Task 13: `chartSpecs.js` and `ChartTile.jsx`

Implements spec §10.1, §10.2.

**Files:**
- Create: `src/features/datastudio/canvas/chartSpecs.js`
- Create: `src/features/datastudio/canvas/ChartTile.jsx`
- Test: `src/features/datastudio/canvas/chartSpecs.test.js`

**Interfaces:**
- Consumes: `aggregate` (Task 11), `EChart` (Task 12).
- Produces:
  - `CHART_TYPES: [{ id, label, needs: { x, y, series } }]` for `bar`, `line`, `area`, `pie`, `scatter`, `heatmap`, `kpi`, `table`
  - `toEChartsOption(chartType, aggResult, tileSpec) → object`
  - `validateTileSpec(spec, dataset) → { ok: boolean, reason?: string }`
  - `<ChartTile tile dataset mask onSelect selection />`

- [ ] **Step 1: Write the failing tests**

```js
import { describe, it, expect } from 'vitest';
import { toEChartsOption, validateTileSpec, CHART_TYPES } from './chartSpecs.js';

const agg = {
  categories: ['HR', 'IT'],
  series: [{ name: 'Amount', data: [40, 60] }],
};
const tile = { id: 't1', title: 'By dept', chart: 'bar',
  encoding: { x: { column: 'Dept' }, y: [{ column: 'Amount', agg: 'sum' }], series: null } };

describe('toEChartsOption', () => {
  it('maps categories onto the x axis for a bar chart', () => {
    const o = toEChartsOption('bar', agg, tile);
    expect(o.xAxis.data).toEqual(['HR', 'IT']);
    expect(o.series[0].type).toBe('bar');
    expect(o.series[0].data).toEqual([40, 60]);
  });

  it('emits a line series for line charts', () => {
    expect(toEChartsOption('line', agg, tile).series[0].type).toBe('line');
  });

  it('fills the area for area charts', () => {
    expect(toEChartsOption('area', agg, tile).series[0].areaStyle).toBeDefined();
  });

  it('reshapes categories and values into name/value pairs for pie', () => {
    const o = toEChartsOption('pie', agg, tile);
    expect(o.series[0].data).toEqual([
      { name: 'HR', value: 40 }, { name: 'IT', value: 60 },
    ]);
  });

  it('enables select and blur states so the source tile can highlight', () => {
    const o = toEChartsOption('bar', agg, tile);
    expect(o.series[0].selectedMode).toBeTruthy();
  });

  it('stacks series when the tile asks for it', () => {
    const stacked = { ...tile, stacked: true };
    const multi = { categories: ['HR'], series: [
      { name: 'a', data: [1] }, { name: 'b', data: [2] }] };
    const o = toEChartsOption('bar', multi, stacked);
    expect(o.series[0].stack).toBe(o.series[1].stack);
  });

  it('reduces a kpi tile to a single number', () => {
    const o = toEChartsOption('kpi', agg, { ...tile, chart: 'kpi' });
    expect(o.value).toBe(100);
  });
});

describe('validateTileSpec', () => {
  const dataset = { byName: new Map([['Dept', {}], ['Amount', {}]]) };

  it('accepts a tile whose columns all exist', () => {
    expect(validateTileSpec(tile, dataset).ok).toBe(true);
  });

  // Spec §12 -- a tile pointing at a deleted column must explain itself.
  it('rejects a tile referencing a missing column and names it', () => {
    const broken = { ...tile, encoding: { ...tile.encoding, x: { column: 'Gone' } } };
    const r = validateTileSpec(broken, dataset);
    expect(r.ok).toBe(false);
    expect(r.reason).toContain('Gone');
  });
});

describe('CHART_TYPES', () => {
  it('declares what each chart type needs', () => {
    expect(CHART_TYPES.find((c) => c.id === 'scatter').needs.y).toBe(2);
  });
});
```

- [ ] **Step 2: Run the tests and confirm they fail**

- [ ] **Step 3: Implement both modules**

`ChartTile` composes: `validateTileSpec` → render the "column missing" state if invalid, otherwise `aggregate(dataset, mask, tile)` in a `useMemo` keyed on `[dataset, mask, tile]`, then `toEChartsOption`, then `<EChart>`. The click handler emits `{ tileId, column, value }` upward.

- [ ] **Step 4: Run the tests and confirm they pass**

- [ ] **Step 5: Commit**

```bash
git add src/features/datastudio/canvas/
git commit -m "Add chart specs and the chart tile"
```

---

## Task 14: `CanvasGrid.jsx` and the tile editor

Implements spec §10.4.

**Files:**
- Create: `src/features/datastudio/canvas/CanvasGrid.jsx`
- Create: `src/features/datastudio/canvas/TileEditor.jsx`
- Modify: `src/styles/datastudio.css`
- Modify: `src/pages/DataStudioPage.jsx` (the `canvas` stage)

- [ ] **Step 1: Build the grid**

12-column CSS Grid. Sizes map `S→3`, `M→6`, `L→9`, `XL→12` columns. Below `640px` every tile spans all 12. Heights: `kpi` short, charts standard, `table` tall.

- [ ] **Step 2: Build tile chrome**

Per tile: title, a size cycle button, move-left / move-right, edit, duplicate, remove, and export-PNG. All must be real `<button>` elements with `aria-label`s and visible focus rings — keyboard reordering is the replacement for the drag engine we cut, so it has to actually work.

- [ ] **Step 3: Build `TileEditor.jsx`**

A panel for one tile: chart type (only types whose `needs` the dataset can satisfy), X column (dimensions and temporal only), Y measures with an aggregation `<select>` each, an optional series column, sort, limit, stacked toggle, and a `respondsToFilters` toggle. Changes preview live.

Columns must be offered **by role** (spec §7.1): measures for Y, dimensions and temporal for X. Never offer an `identifier` as a measure.

- [ ] **Step 4: Verify in the browser**

Add several tiles of different types and sizes, reorder by keyboard, resize, delete, and confirm the layout reflows at 640px and 1024px and looks right in dark mode.

- [ ] **Step 5: Commit**

```bash
git add src/
git commit -m "Add the canvas grid and tile editor"
```

---

# PHASE 5 — Cross-filtering and suggestions

---

## Task 15: Cross-filter wiring and the filter bar

Implements spec §10.3.

**Files:**
- Create: `src/features/datastudio/canvas/FilterBar.jsx`
- Modify: `src/features/datastudio/DataStudioContext.jsx`
- Modify: `src/features/datastudio/canvas/CanvasGrid.jsx`

**Interfaces:**
- Consumes: `maskFor`, `createMaskCache` (Task 10).
- Produces: context additions `{ globalFilters, addFilter, removeFilter, clearFilters, selection, selectMark, clearSelection }`.

- [ ] **Step 1: Hold one mask cache for the canvas**

Create it once with `useRef(createMaskCache())`. Each tile calls `cache.get(dataset, globalFilters, selection, tile.id)`. Because the cache keys on whether the selection *applies* rather than on the tile id (Task 10), N tiles share at most two arrays.

- [ ] **Step 2: Wire clicks to selection**

A tile click sets `selection = { sourceTileId, column, values: [value] }`. Clicking the same value again clears it. Shift-click adds to `values`. `Escape` anywhere on the canvas calls `clearSelection`.

Tiles with `respondsToFilters: false` always receive the unfiltered mask.

- [ ] **Step 3: Highlight the source tile**

Pass `selection` into `ChartTile`; when `tile.id === selection.sourceTileId`, set the matching data indices' `selected` state so ECharts blurs the rest. No recomputation — this is presentation only (spec §10.3).

- [ ] **Step 4: Build the filter bar**

Chips for each global filter with a remove affordance, an add-filter control listing dimension and temporal columns, and a "Clear all" button. When a cross-filter selection is active, show a distinct chip naming the source tile with a clear button — users must be able to see *why* everything is filtered.

- [ ] **Step 5: Verify the behaviour that matters**

In the browser, with at least four tiles: click a bar. Confirm **the clicked chart still shows every category** with the clicked one highlighted, and every other tile has filtered. This is the spec §10.3 rule and the single most important interaction in the feature — if the source tile collapses to one bar, stop and fix it.

Then confirm: clicking again clears, shift-click multi-selects, Escape clears, global filters combine with the selection, and a `respondsToFilters: false` tile ignores both.

- [ ] **Step 6: Commit**

```bash
git add src/
git commit -m "Add cross-filtering and the filter bar"
```

---

## Task 16: `suggestCharts.js`

Implements spec §10.5.

**Files:**
- Create: `src/features/datastudio/suggest/suggestCharts.js`
- Test: `src/features/datastudio/suggest/suggestCharts.test.js`

**Interfaces:**
- Consumes: `DatasetProfile` (Task 3) — specifically `topMeasure` and `primaryTemporal`, which are defined there and must not be re-derived.
- Produces: `suggestCharts(profile) → TileSpec[]` — at most 6, ranked.

- [ ] **Step 1: Write the failing tests**

```js
import { describe, it, expect } from 'vitest';
import { suggestCharts } from './suggestCharts.js';
import { profileDataset } from '../profile/profileDataset.js';

const grid = {
  headers: ['Dept', 'Created', 'Amount'],
  rows: Array.from({ length: 40 }, (_, i) => [
    ['HR', 'IT', 'Finance'][i % 3],
    `${String((i % 28) + 1).padStart(2, '0')}/01/2024`,
    String((i + 1) * 10),
  ]),
};

describe('suggestCharts', () => {
  const tiles = suggestCharts(profileDataset(grid));

  it('returns at most six tiles', () => {
    expect(tiles.length).toBeGreaterThan(0);
    expect(tiles.length).toBeLessThanOrEqual(6);
  });

  it('leads with a KPI row', () => {
    expect(tiles[0].chart).toBe('kpi');
  });

  it('includes a time series on the primary temporal column', () => {
    const line = tiles.find((t) => t.chart === 'line');
    expect(line.encoding.x.column).toBe('Created');
  });

  it('truncates a sub-90-day span to day', () => {
    expect(tiles.find((t) => t.chart === 'line').encoding.x.bin).toBe('day');
  });

  it('includes a bar chart for the low-cardinality dimension', () => {
    const bar = tiles.find((t) => t.chart === 'bar' && t.encoding.x.column === 'Dept');
    expect(bar.encoding.y[0].column).toBe('Amount');
  });

  it('gives every tile a unique id and a title', () => {
    expect(new Set(tiles.map((t) => t.id)).size).toBe(tiles.length);
    expect(tiles.every((t) => t.title)).toBe(true);
  });

  it('never suggests an identifier as a measure', () => {
    const withId = {
      headers: ['Employee ID', 'Dept'],
      rows: Array.from({ length: 30 }, (_, i) => [String(1000 + i), 'HR']),
    };
    const t = suggestCharts(profileDataset(withId));
    expect(t.every((x) => (x.encoding.y ?? []).every((y) => y.column !== 'Employee ID')))
      .toBe(true);
  });

  it('returns an empty array when nothing is chartable', () => {
    const junk = { headers: ['Notes'],
      rows: Array.from({ length: 50 }, (_, i) => [`free text ${i}`]) };
    expect(suggestCharts(profileDataset(junk))).toEqual([]);
  });
});
```

- [ ] **Step 2: Run the tests and confirm they fail**

- [ ] **Step 3: Implement `suggestCharts.js`**

Generate all candidates per spec §10.5, score each, sort, take 6. The interest score: dimensions with 3–12 distinct values score highest and fall off outside that band; multiply by `nonNullRatio`; for measures add a term for coefficient of variation so a constant column ranks last.

- [ ] **Step 4: Run the tests and confirm they pass**

- [ ] **Step 5: Wire it into the canvas stage**

When `commitClean` produces a dataset with no saved dashboard, seed the canvas with `suggestCharts(profile)`.

- [ ] **Step 6: Verify in the browser**

Import a real workbook and confirm you land on a populated canvas in about a second, with sensible charts.

- [ ] **Step 7: Commit**

```bash
git add src/
git commit -m "Add automatic chart suggestion"
```

---

# PHASE 6 — Persistence and export

---

## Task 17: `db.js`

Implements spec §11.

**Files:**
- Create: `src/features/datastudio/store/db.js`
- Test: `src/features/datastudio/store/db.test.js`
- Modify: `package.json`, `vite.config.js`

**Interfaces:**
- Produces: `openDb()`, `saveDataset(record)`, `loadDataset(id)`, `listDatasets()`, `deleteDataset(id)`, `saveCleanPlan(datasetId, steps)`, `loadCleanPlan(datasetId)`, `saveDashboard(record)`, `loadDashboard(id)`, `listDashboards(datasetId?)`, `deleteDashboard(id)`, `storageEstimate()`.

- [ ] **Step 1: Install a fake IndexedDB for tests**

```bash
npm install -D fake-indexeddb@^6
```

Add a second Vitest project entry, or simply set `environment: 'node'` with a setup file that imports `fake-indexeddb/auto`. Scope the setup file to `src/features/datastudio/store/*.test.js` so the pure-function suites stay fast.

- [ ] **Step 2: Write the failing tests**

```js
import { describe, it, expect, beforeEach } from 'vitest';
import 'fake-indexeddb/auto';
import {
  saveDataset, loadDataset, listDatasets, deleteDataset,
  saveDashboard, listDashboards, saveCleanPlan, loadCleanPlan,
} from './db.js';

const record = () => ({
  id: 'ds1', name: 'Requests', sourceFileName: 'r.xlsx', sheetName: 'Sheet1',
  importedAt: Date.now(), rowCount: 3,
  columns: [{ name: 'Amount', type: 'numeric', role: 'measure' }],
  rawColumns: [new Float64Array([1, 2, 3])],
});

describe('datasets', () => {
  beforeEach(async () => {
    for (const d of await listDatasets()) await deleteDataset(d.id);
  });

  it('round-trips a dataset', async () => {
    await saveDataset(record());
    expect((await loadDataset('ds1')).name).toBe('Requests');
  });

  it('preserves TypedArrays through structured clone', async () => {
    await saveDataset(record());
    const back = await loadDataset('ds1');
    expect(back.rawColumns[0]).toBeInstanceOf(Float64Array);
    expect(Array.from(back.rawColumns[0])).toEqual([1, 2, 3]);
  });

  it('lists datasets without their column payloads', async () => {
    await saveDataset(record());
    const list = await listDatasets();
    expect(list[0]).toMatchObject({ id: 'ds1', rowCount: 3 });
    expect(list[0].rawColumns).toBeUndefined();
  });

  it('deletes a dataset', async () => {
    await saveDataset(record());
    await deleteDataset('ds1');
    expect(await loadDataset('ds1')).toBeUndefined();
  });
});

describe('clean plans', () => {
  it('stores plans separately from the dataset', async () => {
    await saveDataset(record());
    await saveCleanPlan('ds1', [{ id: 's1', op: 'trimWhitespace', enabled: true }]);
    expect(await loadCleanPlan('ds1')).toHaveLength(1);
    // The dataset blob itself must be untouched by a plan edit (spec §11).
    expect((await loadDataset('ds1')).rowCount).toBe(3);
  });
});

describe('dashboards', () => {
  it('lists dashboards filtered by dataset', async () => {
    await saveDashboard({ id: 'd1', datasetId: 'ds1', name: 'A', tiles: [], globalFilters: [] });
    await saveDashboard({ id: 'd2', datasetId: 'ds2', name: 'B', tiles: [], globalFilters: [] });
    expect((await listDashboards('ds1')).map((d) => d.id)).toEqual(['d1']);
  });
});
```

- [ ] **Step 3: Run the tests and confirm they fail**

- [ ] **Step 4: Implement `db.js`**

Database `pmw-datastudio`, version 1, three object stores keyed as in spec §11, with an index on `datasetId` for `dashboards`. `listDatasets` reads from a lightweight `meta` projection stored alongside — or strips `rawColumns` before returning; either way the list must not deserialise every column.

Wrap every write in a `try/catch` that re-throws `QuotaExceededError` as a typed error the UI can recognise.

- [ ] **Step 5: Run the tests and confirm they pass**

- [ ] **Step 6: Commit**

```bash
git add package.json package-lock.json vite.config.js src/features/datastudio/store/
git commit -m "Add IndexedDB persistence"
```

---

## Task 18: Saved datasets and dashboards in the UI

**Files:**
- Create: `src/features/datastudio/store/useDashboards.js`
- Modify: `src/pages/DataStudioPage.jsx`, `src/features/datastudio/DataStudioContext.jsx`

- [ ] **Step 1: Add a library to the idle stage**

Below the drop zone, list saved datasets — name, source file, row count, import date via `formatMYT`, size — each opening straight to its canvas. Include a delete affordance with a confirmation.

- [ ] **Step 2: Add dashboard save and load**

A "Save dashboard" action naming the current tiles and global filters, and a picker listing saved dashboards for the current dataset. Loading one replaces the tiles.

- [ ] **Step 3: Add the storage meter**

Drive a usage bar from `storageEstimate()`. On `QuotaExceededError`, open a dialog listing datasets by size with delete buttons — never fail silently (spec §11).

- [ ] **Step 4: Verify persistence across a reload**

Import, clean, build a canvas, save, **reload the browser**, and confirm the dataset and dashboard both come back with charts identical to before.

- [ ] **Step 5: Commit**

```bash
git add src/
git commit -m "Add saved datasets and dashboards"
```

---

## Task 19: Exports

Implements spec §11.

**Files:**
- Create: `src/features/datastudio/store/exporters.js`
- Modify: `src/features/datastudio/canvas/CanvasGrid.jsx`, `src/pages/DataStudioPage.jsx`

**Interfaces:**
- Produces: `exportTilePng(chartInstance, title)`, `exportDatasetCsv(dataset, name)`, `exportDashboardJson(dashboard, name)`, `importDashboardJson(file) → dashboard`.

- [ ] **Step 1: Implement the exporters**

PNG via `chart.getDataURL({ type: 'png', pixelRatio: 2, backgroundColor: <--it-panel> })` — pass the token value explicitly or the export comes out transparent. CSV must quote fields containing commas, quotes or newlines, and render dates via `formatMYT`. Dashboard JSON carries `{ version: 1, name, tiles, globalFilters, datasetName, columns }` so an import can tell the user which file it expects.

- [ ] **Step 2: Wire the buttons**

Export PNG on each tile; export CSV and export dashboard JSON on the canvas toolbar. Import dashboard JSON in the dashboard picker, validating `version` and reporting missing columns by name rather than throwing.

- [ ] **Step 3: Verify each export**

Confirm the PNG has a solid background in both themes, the CSV opens in Excel with dates as `DD/MM/YYYY HH:mm`, and an exported dashboard re-imports onto the same dataset.

- [ ] **Step 4: Commit**

```bash
git add src/
git commit -m "Add PNG, CSV and dashboard exports"
```

---

## Task 20: Documentation and final verification

- [ ] **Step 1: Update `AGENTS.md`**

Add `/data-studio` to the ROUTES table. Add rows to WHERE TO LOOK for the pipeline, the aggregation engine and the IndexedDB store. Document the `src/features/` pattern under CONVENTIONS, saying why it exists. Add to ANTI-PATTERNS: never import the `echarts` umbrella (import from `echarts/core` via `echartsCore.js`), never shift a date-only column, and never filter the tile that originated a cross-filter selection.

- [ ] **Step 2: Run the full suite**

```bash
npm test
```

Expected: all green.

- [ ] **Step 3: Lint**

```bash
npm run lint
```

Expected: **no new** errors. `FormPage`, `AssetChecklistPage`, `SignatureDialog` and `ThemeContext` errors are pre-existing — check the filenames, and fix anything reported in files this plan created.

- [ ] **Step 4: Build and check the bundle**

```bash
npm run build
```

Expected: `xlsx` and `echarts` live in the lazy Data Studio chunk; the `index` chunk is within a few KB of its pre-feature size.

- [ ] **Step 5: Full manual pass**

Import a real messy workbook → review the profile → adjust a clean step → set a timezone → commit → land on suggested charts → click a bar and confirm the source tile does not filter itself → add a tile → save → reload → export a PNG. Repeat in dark mode and at a 375px viewport.

- [ ] **Step 6: Commit**

```bash
git add AGENTS.md
git commit -m "Document Data Studio"
```

---

## Self-Review Notes

Checked against the spec section by section:

- §4.1–4.6 decisions → Tasks 1, 4, 12, 17 (deps and persistence), Task 10 (engine choice)
- §6.1 columnar model → Task 8; §6.2 worker → Task 5; §6.3 bundle → Tasks 5, 12; §6.4 layout → File Structure; §6.5 shell integration → Task 5
- §7 inference → Tasks 2, 3 (§7.5 date order also in Task 9's banner)
- §8 cleaning → Tasks 6, 7, 8, 9
- §9 Malaysia time → Task 1, with the UI in Task 9
- §10.1–10.2 → Tasks 11, 13; §10.3 filters → Tasks 10, 15; §10.4 layout → Task 14; §10.5 suggestion → Task 16; §10.6–10.7 → Task 12
- §11 persistence and export → Tasks 17, 18, 19
- §12 error handling → distributed: parse errors Task 5, header failure Tasks 4 and 5, 95%-rule casualties Tasks 2 and 5, date conflict Task 9, missing column Task 13, quota Tasks 17 and 18
- §13 testing → every logic task; §14 phasing → the phase headings; §15 risks → bundle checks in Tasks 5, 12, 20

Type consistency verified across tasks: `ColumnVerdict` (Task 2) is extended by `ColumnProfile` (Task 3); `Dataset` (Task 8) is consumed unchanged by Tasks 10, 11, 13; `TileSpec` is produced by Task 16 and consumed by Task 13; `topMeasure` / `primaryTemporal` are defined once in Task 3 and referenced, not re-derived, in Task 16.

One deliberate deviation from the spec's phasing is flagged inline at the head of Phase 1: date parsing moves from phase 2 to Task 1 because type inference depends on it.
