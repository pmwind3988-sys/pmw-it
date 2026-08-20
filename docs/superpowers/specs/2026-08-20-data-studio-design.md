# Data Studio — Excel import, profiling and a cross-filtered chart canvas

**Date:** 2026-08-20
**Status:** Approved design, ready for implementation planning
**Route:** `/data-studio`

## 1. Purpose

A section of the PMW IT Service Portal where a user drops in an Excel workbook and,
within about a second, lands on a populated dashboard of charts built from their own
data — which they can then filter, re-cut, extend and save.

It has to beat PowerBI on three specific axes: it should identify what the data *is*
without being told, clean it without the user writing a single formula, and make
clicking a chart to slice everything else feel instant.

## 2. Goals

1. Import `.xlsx` / `.xlsm` / `.csv`, pick a sheet, find the real header row.
2. Identify each column's type and its analytical role — quantitative (measure),
   qualitative (dimension), temporal, or ignorable.
3. Propose a cleaning plan the user approves as a checklist, never a formula.
4. Parse datetimes correctly for Malaysian use, with the timezone assumption visible
   rather than silent.
5. Auto-generate a starting dashboard, then let the user add and edit chart tiles.
6. Cross-filter: clicking a mark in one tile filters every other tile.
7. Save datasets and dashboards so they survive a browser restart.

## 3. Non-goals (v1)

Explicitly excluded, each for a stated reason:

| Excluded | Reason |
|---|---|
| Uploading anything to SharePoint | Data stays browser-only by decision (§4.1) |
| Drag-to-resize tiles | Size presets deliver the value; a drag engine does not |
| Multi-sheet joins / relationships | One sheet at a time; a relational model is its own project |
| Undo/redo on the canvas | Clean plan is already non-destructive; canvas edits are cheap to redo |
| Calculated columns / formula language | An expression parser is a separate project |
| Reshaping: unpivot, wide→long crosstabs | Out of scope per §4.2. **Consequence: a sheet with months across the top cannot be plotted on a time axis.** |
| Statistical scaling (min-max, z-score, log) | Not requested; cheap to add later once numeric typing exists |
| Database-style (1NF/2NF/3NF) normalisation | Not what "normalize" meant here |
| Fuzzy / edit-distance category merging | Causes silent false merges — `Dept A` and `Dept B` are edit distance 1 (§8.2) |

## 4. Decisions

### 4.1 Persistence: browser-only

Datasets and dashboards live in IndexedDB. The workbook never leaves the machine.
No SharePoint provisioning, no upload governance, no token required.

**Consequence:** work is per-browser and not shareable. Mitigated by dashboard JSON
export (§11).

### 4.2 "Normalize" means clean and tidy values

Trim, unify case, merge near-duplicate categories, coerce numbers-stored-as-text,
standardise nulls, drop empty rows/columns, dedupe. Not reshaping, not scaling, not
relational modelling.

### 4.3 Malaysia time: per-column, default no shift

Every date column is assumed to already be local wall-clock. A per-column toggle marks
a column as UTC, which applies +8. The assumption is always visible in the UI.

### 4.4 Dependencies: SheetJS + Apache ECharts

Chosen deliberately for maximum capability. Mitigated by tree-shaking and route-level
lazy loading (§6.3).

### 4.5 Feel: cross-filtered dashboard canvas

Multiple tiles, a global filter bar, click-a-mark-to-filter-everything.

### 4.6 Engine: in-memory columnar store + shared filter mask

Rejected alternatives: DuckDB-WASM (~30MB WASM, COOP/COEP headers, overkill for
spreadsheet-sized data) and per-tile `.filter().reduce()` (stutters at 8 tiles × 50k
rows, which is precisely the interaction being built).

## 5. Pipeline

```
File → parse → raw grid → detect header → profile → propose clean plan
     → apply → typed columnar dataset → canvas (tiles + filter bus)
```

Every stage is a pure function of the previous stage's output. Nothing before `apply`
touches React state; nothing after it re-reads the file.

## 6. Architecture

### 6.1 Columnar dataset

| Role | Type | Representation | Null encoding |
|---|---|---|---|
| Measure | numeric | `Float64Array` | `NaN` |
| Temporal | date, datetime | `Float64Array` of epoch ms | `NaN` |
| Dimension | categorical | `Int32Array` codes + `string[]` dictionary | `-1` |
| Dimension | boolean | `Uint8Array` | `2` |
| Ignored | text, identifier | `string[]` | `null` |

Dictionary encoding is load-bearing: group-by on a categorical column counts small
integers instead of hashing strings, and a cross-filter selection is a set of integer
codes.

Filtering produces **one shared `Uint8Array` row mask** read by every tile — one pass
over the data per filter change, not one pass per tile.

### 6.2 Web Worker

Parse, profile and clean run in `studio.worker.js`. A 50k-row workbook parsed on the
main thread freezes the UI for seconds. Typed arrays transfer at zero copy cost, so the
boundary is cheap. The worker posts progress events that drive a real progress bar.

### 6.3 Bundle strategy

- Import from `echarts/core`; register only `BarChart`, `LineChart`, `PieChart`,
  `ScatterChart`, `HeatmapChart`, `CanvasRenderer` and the components actually used.
  Roughly 1MB → ~200KB gzipped.
- `DataStudioPage` is loaded via `React.lazy`. ECharts and SheetJS never enter the main
  bundle; every existing route loads exactly as fast as today.

### 6.4 File layout

Introduces `src/features/` — a new pattern for this repo, chosen because ~20 files
scattered into the existing flat `components/` and `utils/` would bury the structure
that is there.

```
src/pages/DataStudioPage.jsx          route: AppShell + stage switch
src/features/datastudio/
  ingest/   parseWorkbook.js  detectHeader.js
  profile/  inferType.js  profileColumn.js  profileDataset.js
  clean/    proposeCleanPlan.js  cleanOps.js  applyCleanPlan.js
  time/     malaysiaTime.js
  engine/   dataset.js  filterMask.js  aggregate.js
  canvas/   CanvasGrid.jsx  ChartTile.jsx  EChart.jsx
            echartsTheme.js  chartSpecs.js
  suggest/  suggestCharts.js
  store/    db.js  useDashboards.js
  worker/   studio.worker.js
  DataStudioContext.jsx
src/styles/datastudio.css
```

### 6.5 Integration with the existing shell

- Route added to `src/App.jsx`; nav entry added to `NAV_ITEMS` in `AppShell.jsx`
  (label "Data Studio").
- The page renders `<AppShell title subtitle actions>` and its own body, per the
  project's page-composition convention. It does not gate itself — `AppShell` is the
  one auth gate.
- **No SharePoint token is acquired anywhere in this feature.** The section works
  offline once loaded.

## 7. Type inference (`inferType.js`)

Full scan of every non-null cell. No sampling: typing is cheap per cell, we are in a
worker, and sampling produces "it guessed wrong because the odd rows are at the bottom".

### 7.1 Roles

| Role | Types | Used for |
|---|---|---|
| Measure | numeric | Y axes, aggregation |
| Dimension | categorical, boolean | X axes, series, filters |
| Temporal | date, datetime | time axes, date truncation |
| Ignored | text, identifier, empty | offered manually, never auto-suggested |

### 7.2 The 95% rule

A column takes a type when ≥95% of its non-null values parse as that type. Values that
fail become nulls **and are reported** with a count and a preview of the offending
values. Silent coercion is data corruption.

### 7.3 Resolution order

`boolean → datetime → numeric → categorical → text`, then identifier override.

- **Boolean** fires only on known pairs: `yes/no`, `true/false`, `y/n`. Deliberately
  **not** `0/1`, which is genuinely numeric more often than not.
- **Categorical vs text**: distinct count ≤ 50 **or** distinct ratio < 0.05 → categorical.
- **Identifier override**: distinct ratio > 0.95 **and** (name matches
  `/id|no|code|ref|serial/i` **or** the values are a monotonic integer sequence).
  Identifiers are never offered as a measure — summing employee IDs is meaningless, and
  PowerBI does it by default.

### 7.4 Excel traps handled explicitly

| Input | Result |
|---|---|
| `"1,234.50"`, `"RM 1,234"`, `"$1,234"` | numeric `1234.5` / `1234` |
| `"45%"` | `0.45`, flagged as percent for display formatting |
| `"(1,234)"` | `-1234` (accounting negative) |
| `#N/A` `#DIV/0!` `#REF!` `#VALUE!` | null |
| `""` `"-"` `"–"` `"N/A"` `"NA"` `"NIL"` `"TBD"` | null |
| non-breaking space `U+00A0`, zero-width chars | stripped |
| `"007"` | **stays text** — leading zeros mean it is a code, not the number 7 |

The leading-zero rule is what protects employee IDs and cost centres from being
silently mangled.

### 7.5 D/M/Y vs M/D/Y

True Excel date cells arrive from SheetJS (`cellDates: true`) as `Date` objects —
unambiguous, used directly.

For **string** dates the whole column is scanned:

1. Any first component > 12 → proves D/M/Y.
2. Any second component > 12 → proves M/D/Y.
3. Both present → conflict, flagged loudly, column left as text pending user choice.
4. Neither → ambiguous. **Default to D/M/Y (Malaysian convention)** and show a banner
   with the first five dates rendered both ways and a one-click flip.

ISO `YYYY-MM-DD` is detected first and is unambiguous.

## 8. Cleaning

### 8.1 The plan is data, not mutation

`proposeCleanPlan(profile)` returns steps:

```js
{ id, column, op, params, confidence, affectedCount, preview }
```

Rendered as a checklist with live counts ("trim 1,204 values", "merge 3 spellings of
Kuala Lumpur"). High-confidence steps arrive pre-checked.

**Raw columns are retained** in memory and in IndexedDB; cleaned columns are derived.
Toggling any step re-runs `applyCleanPlan` from raw. Nothing is ever destroyed, the
preview is free, and re-importing the same file can re-apply a saved plan.

Memory cost of holding raw + cleaned is acceptable: raw categoricals are already
dictionary-encoded.

### 8.2 Operations (`cleanOps.js`, one pure function each)

`trimWhitespace` · `normalizeNulls` · `parseNumber` · `unifyCase` · `mergeCategories` ·
`parseDate` · `dropEmptyColumns` · `dropEmptyRows` · `dedupeRows` · `castType`

**`mergeCategories` is key-based only.** Values cluster by
`lowercase → trim → collapse internal whitespace → strip punctuation`; the most frequent
original spelling becomes the canonical label. `"Kuala Lumpur"`, `"kuala lumpur"` and
`"Kuala  Lumpur "` collapse to one category.

Fuzzy/edit-distance merging is excluded (§3): it demos well and quietly corrupts data.
Instead a **manual merge UI** lets the user select two categories and merge them,
reversibly.

## 9. Malaysia time (`malaysiaTime.js`)

MYT has been a flat UTC+8 with no DST since 1982, so no timezone database is needed for
current data.

- `MYT_OFFSET_MIN = 480`
- `toEpochMs(value, { order, sourceZone })` — accepts `Date` objects, Excel serial
  numbers, ISO strings, and D/M/Y or M/D/Y strings. Applies +8 **only** when
  `sourceZone === 'utc'`.
- Excel serial → epoch: `(serial - 25569) * 86400000`, with the **1900 leap-year bug**
  handled (serial 60 is a date that never existed; serials below 61 need a different
  offset).
- `formatMYT(epochMs, style)` uses
  `Intl.DateTimeFormat('en-GB', { timeZone: 'Asia/Kuala_Lumpur' })` rather than manual
  arithmetic, so historical pre-1982 dates format correctly too. Default style
  `DD/MM/YYYY HH:mm`.

**Hard rule: date-only columns are never shifted**, even when the UTC toggle is on.
Adding 8 hours to a pure date moves it to the wrong day.

## 10. Canvas

### 10.1 Tile spec

```js
{
  id: 'tile_7',
  title: 'Requests by Department',
  size: 'M',                        // S=3 | M=6 | L=9 | XL=12 grid columns
  chart: 'bar',                     // bar|line|area|pie|scatter|heatmap|kpi|table
  encoding: {
    x:      { column: 'Department', bin: null },
    y:      [{ column: 'Amount', agg: 'sum' }],
    series: { column: 'Entity' },   // nullable
  },
  sort: { by: 'y', dir: 'desc' },
  limit: 20,                        // top-N, remainder grouped as "Other"
  respondsToFilters: true,
}
```

Aggregations: `sum`, `avg`, `count`, `countDistinct`, `min`, `max`, `median`.

### 10.2 One aggregator, N renderers

`aggregate(dataset, mask, spec)` returns a single shape for every chart type:

```js
{ categories: [...], series: [{ name, data: [...] }] }
```

Tests target the aggregator, which is where the bugs that matter live. A wrong sum
renders as a perfectly convincing chart.

### 10.3 Filter model

Two distinct concepts:

- **Global filters** — from the filter bar. Explicit, persistent, apply everywhere,
  saved with the dashboard.
- **Cross-filter selection** — from clicking a mark. Transient, one source tile at a
  time, not persisted.

**The source tile does not filter itself.**

```
maskFor(tileId) = globalMask ∩ (selectionMask if sourceTileId !== tileId else ALL)
```

Click "Finance" in the department bar chart: that chart keeps showing every department
with Finance highlighted; every other tile filters to Finance. Filtering the source too
would collapse it to a single bar with nothing left to click — the most common way
homemade BI tools feel broken.

Masks are memoised by filter signature, so N tiles share at most two mask computations.

Source-tile highlighting uses ECharts' native `select` / `blur` states — no extra
computation. Interactions: click to select, click again to clear, shift-click to
multi-select, `Escape` clears all.

### 10.4 Layout

12-column CSS Grid on the project's existing mobile-first breakpoints (640px, 1024px).
Size presets S/M/L/XL; heights vary by chart type (KPI short, charts standard, table
tall). Reordering via move-left/move-right buttons, keyboard accessible. Below 640px
every tile is full width.

### 10.5 Chart suggestion (`suggestCharts.js`)

**"Top measure"** is defined once and used throughout: the measure column with the
highest non-null ratio, ties broken by original column order. **"Primary temporal"** is
the temporal column with the highest non-null ratio, same tie-break.

On a fresh dataset, generate candidates:

1. A KPI row — row count plus one KPI per measure, up to three, ordered by non-null ratio.
2. Primary temporal + top measure → line chart, truncation chosen by span: day if under
   90 days, month if under 3 years, otherwise quarter.
3. Each dimension with 12 or fewer distinct values × top measure → bar chart, sorted
   descending, top 10 plus "Other".
4. Each measure → histogram, bins via Freedman–Diaconis.
5. Measure pairs with `|Pearson r| ≥ 0.3` → scatter. Correlation is computed on up to
   5,000 evenly-spaced sampled rows, since this is a ranking heuristic and not a
   reported statistic.

Score by interest — cardinality sweet spot (3–12), null-ratio penalty, variance — and
keep the top 6. The user lands on a populated canvas, not an empty "Add chart" page.

### 10.6 `EChart.jsx`

~60 lines: `echarts.init` on mount, `setOption` on spec change, `ResizeObserver` →
`resize()`, `dispose()` on unmount, `click` bound to the cross-filter dispatch. No
`echarts-for-react` dependency.

### 10.7 Theming (`echartsTheme.js`)

Reads the real tokens from `getComputedStyle(document.documentElement)`: `--it-brand`,
`--it-canvas`, `--it-panel`, `--it-ink`, `--it-ink-soft`, `--it-line`, `--it-accent`,
`--it-good`, `--it-danger`. Rebuilds and re-registers the theme when `ThemeContext`
flips `[data-theme='dark']`.

The categorical series palette must be colour-blind-safe and legible in both themes; it
is validated at implementation time rather than assumed.

Chart animations respect `prefers-reduced-motion`, matching `shell.css`.

## 11. Persistence

IndexedDB database `pmw-datastudio`, version 1. TypedArrays survive structured clone
natively, so columns persist as-is.

| Store | Key | Contents |
|---|---|---|
| `datasets` | `id` | name, sourceFileName, sheetName, importedAt, rowCount, column metadata, **raw** columns |
| `cleanPlans` | `datasetId` | ordered step array — stored separately so editing a plan does not rewrite the column blob |
| `dashboards` | `id` | name, datasetId, tiles, globalFilters, createdAt, updatedAt |

Cleaned columns are never persisted; they are derived from raw plus plan on load.

**Quota is handled explicitly.** `navigator.storage.estimate()` drives a usage bar, and
`QuotaExceededError` raises a real "storage full — here are your datasets by size"
dialog. The default behaviour is a silent failure that is miserable to diagnose.

**Exports**: chart PNG via `echarts.getDataURL()`; cleaned data → CSV; dashboard
definition → JSON. The last closes the sharing gap opened by the browser-only decision
at near-zero cost.

## 12. Error handling

| Failure | Behaviour |
|---|---|
| Unreadable / corrupt workbook | Named error with the file name; import stage stays open |
| Password-protected workbook | Explicit message — SheetJS cannot open it |
| Zero data rows after header detection | Offer manual header-row selection |
| Column fails the 95% rule | Falls back to text; reported in the profile panel |
| Date order conflict (§7.5 case 3) | Column left as text; banner asks the user to choose |
| Worker crash | Error surfaced with a retry; app is not left on a spinner |
| `QuotaExceededError` | Storage dialog (§11) |
| Tile references a deleted column | Tile renders an inline "column missing" state, not a blank card |

## 13. Testing

Adds **Vitest** as a dev dependency — no runtime cost, native Vite integration. The
project has no test runner today.

The justification is specific: `inferType`, `malaysiaTime`, `aggregate` and `filterMask`
are pure functions whose bugs are **invisible**. A mis-parsed date or a wrong group-by
renders as a perfectly convincing chart.

Written test-first, per the project's TDD workflow:

- `inferType` — a fixture table of nasty columns → expected type and role, covering
  every row of §7.4
- `malaysiaTime` — serial 1, serial 60 (leap-year bug), serial 61; D/M/Y vs M/D/Y vs
  ISO; UTC shift on and off; **date-only never shifts**
- `aggregate` — every aggregation, group-by with nulls, top-N plus "Other", numeric
  binning, date truncation at each span boundary
- `filterMask` — the self-exclusion rule (§10.3), multi-select, memoisation identity
- `cleanOps` — each op in isolation, plus plan-reapplication idempotency
- `detectHeader` — banner rows above the header, blank leading rows, no header at all

Post-implementation: a manual smoke pass through the browser preview covering import →
profile → clean → canvas → cross-filter → save → reload.

## 14. Phasing

Ordered so each phase is independently verifiable:

1. **Ingest + profile** — parse, header detection, type inference, profile panel.
   Worker in place.
2. **Clean + time** — clean plan proposal and apply, Malaysia time, the review
   checklist UI.
3. **Engine** — dataset, aggregate, filterMask, fully unit-tested with no UI.
4. **Canvas** — EChart wrapper, theming, tile rendering, grid layout, chart editor.
5. **Cross-filter + suggest** — filter bar, click-to-filter, auto-suggested dashboard.
6. **Persistence + export** — IndexedDB, saved dashboards, quota handling, exports.

Phase 3 before phase 4 is deliberate: the aggregation engine is proven by tests before
any chart can disguise a wrong number as a nice picture.

## 15. Risks

| Risk | Mitigation |
|---|---|
| ECharts bundle weight | Tree-shaken core build plus route-level lazy load (§6.3); verify against `npm run build` output |
| Type inference guessing wrong | Every verdict is overridable in the profile panel; the 95% rule reports its casualties |
| Browser-only data feels limiting later | Dashboard JSON export now; the store layer is isolated in `store/db.js` if a SharePoint backend is added |
| Large workbooks exhausting memory | Row-count warning above 200k rows; columnar plus dictionary encoding keeps footprint low |
| `src/features/` diverging from repo conventions | Documented here, and to be added to AGENTS.md on completion |
