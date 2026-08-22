# Data Studio — Text Analysis

**Date:** 2026-08-22
**Status:** design approved, spec under review

## 1. What this is

A side function of Data Studio that reads free-text survey answers and turns
them into structured, chartable, editable data — entirely inside the browser.

Given a sheet of survey responses, it:

1. splits each written answer into the separate issues it actually contains,
2. files each issue into a category,
3. discovers the groupings that exist in the text regardless of those
   categories,
4. ranks everything into a priority order,
5. writes the results back onto the dataset as ordinary columns, so the
   existing chart canvas, filters and cross-filtering work on them unchanged.

Every result is overridable by hand, and every override survives a reload.

The driving example is
`IT Operational Efficiency & Process Improvement Survey`, 42 responses,
columns: `Please Select Your Department` (categorical), `Which of the
following operational or reporting activities create the biggest challenges
or bottlenecks in your daily work?` (semicolon-joined multi-select), `Please
describe your selected challenge(s) in as much detail as possible.` (free
text, 1–5 sentences).

## 2. Hard constraints

**This function never sends data anywhere.** No SharePoint list, no Graph
call, no API, no telemetry, no analytics. The survey text, the derived
issues and the user's overrides exist only in the browser tab and in that
browser's own IndexedDB. The language model runs locally on WebAssembly and
is served from this application's own origin — including the ONNX runtime
binaries, which by default would be fetched from a public CDN and must be
self-hosted to hold this line.

This is a stated requirement, not an implementation preference. Any future
change that introduces a network call carrying response text breaks the
promise the tab makes on screen.

Local persistence (IndexedDB, this machine only) is in scope and required —
the user's overrides must survive a reload.

## 3. Non-goals

- **No generated prose.** The model produces similarity, not sentences.
  Themes are named from their most distinctive terms, not by a writer.
  Severity is a measurable signal, not a judgement. AI-written summaries
  and explanations would require a hosted LLM and are explicitly excluded.
- **No Microsoft Forms link.** Forms exposes no public response API. The
  achievable version — reading the response workbook out of OneDrive over
  Graph, using the sign-in this app already has — is real, wanted, and
  deferred to a later phase with its own design.
- **No languages other than English.** The model is English. The data is
  English.
- **No model training or fine-tuning.**

## 4. Where it lives

A new stage in the existing Data Studio flow:

```
idle → parsing → profile → cleaning → canvas
                                   ↘ text ↗
```

`stage === 'text'` renders `TextAnalysis`. It is reachable by a
**Text analysis** button in the toolbars of both the cleaning and canvas
stages, and that button is rendered only when `detectTextColumns` finds at
least one qualifying column. On a dataset with no free text the feature is
invisible.

Leaving the tab returns to the canvas. The derived columns it wrote stay on
the dataset.

## 5. Module layout

Follows the existing feature layering: pure functions over plain data
underneath, React only at the top, the model isolated behind one seam.

```
src/features/datastudio/text/
  detectTextColumns.js    which columns qualify as free text
  splitIssues.js          one answer → separate issue fragments
  boilerplate.js          label prefixes and non-answer lexicon
  buckets.js              the starter bucket set
  embed.js                THE ONLY IMPURE MODULE — wraps transformers.js
  similarity.js           cosine; fragment → bucket, or Unsorted
  cluster.js              agglomerative clustering → themes
  labelCluster.js         c-TF-IDF naming of a theme
  severity.js             wording/length/emphasis → 0–1 signal
  rankIssues.js           groups → priority order
  overrides.js            user edits applied over a raw analysis
  analysis.js             orchestration; returns the analysis object
  deriveColumns.js        analysis → new dataset columns
  TextAnalysis.jsx        the stage
  BucketEditor.jsx
  IssueTable.jsx
  ThemeList.jsx
  PriorityBoard.jsx
src/features/datastudio/worker/text.worker.js
```

Every module except `embed.js`, `text.worker.js` and the JSX is a pure
function of its arguments and is unit-tested. `embed.js` is mocked in tests
with deterministic vectors, so the whole pipeline is testable without the
model present.

## 6. The pipeline

### 6.1 detectTextColumns

The profiler already infers a `text` type. That is necessary but not
sufficient — an identifier column is also text. A column qualifies when:

- inferred type is `text`,
- mean non-empty length >= 40 characters,
- at least 60% of values are non-empty,
- distinct ratio >= 0.8 (near-unique values; a repeated 40-character value
  is a category, not prose).

Returns the qualifying columns ordered by mean length, longest first. That
column is the default selection; the user can switch columns in the tab.

### 6.2 splitIssues

Pure. `splitIssues(text) -> string[]`.

1. Normalize: NFKC, collapse whitespace, strip zero-width characters.
   (Per project convention, invisible characters in tests are written as
   escapes, never as literals.)
2. Strip label prefixes at the start of any line —
   `[Selected Challenge]:`, `[Detailed Description]:`, `Challenge:`,
   `Description:`, `Issue:`, `Problem:`, with or without brackets, colon
   optional, case-insensitive. These appear verbatim in the source data
   because respondents copy the question into the answer.
3. Split on hard newlines, then on bullet markers (`-`, `•`, `*`, `1.`,
   `1)`), then on sentence boundaries.
4. Merge any fragment shorter than 25 characters into the preceding one — a
   sentence split on `e.g.` or an abbreviation must not become an issue.
5. Drop non-answers: a fragment whose first word matches the non-answer
   lexicon (`no`, `none`, `nil`, `n/a`, `na`, `nothing`, `-`) and which
   contains fewer than 20 alphabetic characters. This removes
   `no issue from IT` without removing
   `No proper system exists for tracking approvals`.
6. Cap at 12 fragments per response.

A response yielding zero fragments is recorded as **no issue raised** and
excluded from every count — but stays visible in the issue table with that
status, so the count of respondents who reported nothing is itself a figure.

Expected on the driving dataset: 42 responses to roughly 100–130 fragments.

### 6.3 embed

The only module that touches the model.

```js
embedAll(texts, { onProgress }) -> Promise<Float32Array[]>
```

`@huggingface/transformers`, feature-extraction pipeline,
`Xenova/all-MiniLM-L6-v2`, quantized ONNX, mean pooling, L2-normalized,
384 dimensions. Batch size 16, progress reported per batch.

Configured for strict local operation:

```js
env.allowRemoteModels = false;
env.localModelPath = '/models/';
env.backends.onnx.wasm.wasmPaths = '/ort/';
```

The second line is what keeps the model local. The third is what keeps the
ONNX runtime local; without it the runtime binaries are fetched from a
public CDN at first use and §2 is quietly broken.

The pipeline instance is created once and held for the worker's lifetime.

### 6.4 similarity

Pure. Each bucket is embedded as its **description plus its example
phrases**, averaged and re-normalized — not as its name. A bucket named
"SAP" carries almost no signal; the sentence describing what belongs in it
carries a lot. This is why the description field is prominent in the editor.

`assignBuckets(fragmentVectors, bucketVectors, threshold)` returns, per
fragment, the best-scoring bucket and its cosine score, or `Unsorted` when
the best score falls below `threshold`.

`threshold` defaults to **0.30** and is exposed in the UI as a confidence
control. The default is a starting point requiring calibration against real
output during implementation; the control exists because no single value is
right for every survey.

### 6.5 cluster

Pure. Agglomerative, average linkage, cosine distance, merging while
distance < **0.45** (exposed as a granularity control: fewer, broader themes
versus more, narrower ones).

At the expected scale (~130 fragments) the O(n²) distance matrix is
negligible. A guard rejects clustering above 5,000 fragments with an
explanatory message rather than freezing.

Clusters with fewer than 2 members are collected into **one-offs** rather
than presented as themes — a theme of one is a quote, not a pattern. The
one-offs remain visible and countable.

### 6.6 labelCluster

Pure, c-TF-IDF. Tokenize to lowercase words, drop English stopwords and a
short domain stoplist. Score each term by its frequency inside the cluster
times `log(N / documentFrequency)` across all fragments. Take the top 4 and
join with ` · `, producing names like `approval · reminder · follow-up ·
status`.

The name is a starting point. Renaming a theme is a first-class action.

### 6.7 severity

Pure, and documented on screen as a signal rather than an assessment.
Score in 0–1, weighted:

- 0.50 — matches against an intensity lexicon (`time-consuming`,
  `prone to error`, `manual`, `repetitive`, `delay`, `missed`, `overlooked`,
  `bottleneck`, `tedious`, `duplicate`, `rework`, `chase`, `constantly`,
  `unable`, `cannot`, `difficult`), saturating at 4 matches
- 0.20 — fragment length, normalized against the corpus, capped
- 0.20 — how many challenges that respondent selected in the multi-select
  column (breadth of pain, taken from structured data rather than inferred)
- 0.10 — emphasis: exclamation marks, shouted words

### 6.8 rankIssues

Pure. For every group — each bucket and each theme — compute distinct
respondents, fragment count, and mean severity.

```
score = distinctRespondents × (1 + meanSeverity)
```

Distinct respondents leads deliberately: one person writing five furious
sentences must not outrank five people each writing one. Ties break on mean
severity, then alphabetically for stability. Pinned items are lifted to the
top in pin order; suppressed items sink to the bottom.

## 7. Starter buckets

Shipped in `buckets.js`, editable in full. Each entry is
`{ id, label, description, hints: string[] }`.

| Label | Covers |
|---|---|
| SAP / ERP | SAP transactions, ERP modules, master data, postings |
| Data Consolidation & Reporting | combining files, recurring reports, dashboards |
| Manual Data Entry | retyping, copy-paste between systems, transcription |
| Approvals & Workflow | sign-off chasing, status tracking, reminders |
| Forms & Paperwork | paper forms, physical signatures, hardcopy routing |
| Information Retrieval | hunting for files, records, emails, history |
| Communication & Coordination | WhatsApp/email as a system of record, handoffs |
| Network & Internet | connectivity, VPN, speed, remote access |
| IT Support & Systems | hardware, accounts, access, software faults |
| Digitization & Automation | requests to replace a manual process |
| AI Opportunities | explicit asks for AI or intelligent assistance |
| Training & Knowledge | not knowing how, undocumented process |

Plus the fixed non-bucket **Unsorted**, which cannot be deleted.

Editing a bucket re-embeds only the bucket descriptions — twelve short
strings — and re-scores against cached fragment vectors. This is what makes
the editor feel instant.

## 8. Overrides

The raw analysis is never mutated. Overrides are a separate record:

```js
{
  retags:      { [fragmentId]: bucketId },
  noise:       Set<fragmentId>,
  themeNames:  { [themeId]: string },
  themeMerges: { [themeId]: targetThemeId },
  themeSplits: { [themeId]: { [fragmentId]: newThemeId } },
  pinned:      fragmentOrGroupId[],
  suppressed:  Set<groupId>,
}
```

`applyOverrides(rawAnalysis, overrides) -> analysis` is pure and is what
every screen renders. Consequences that matter:

- re-running the model, changing the threshold, or editing a bucket never
  destroys hand corrections;
- "reset to what the model said" is discarding one object;
- an override referring to a fragment that no longer exists is dropped
  silently on load, so a re-import with different data cannot corrupt state.

Fragment identity is `${rowIndex}:${fragmentIndex}` — stable for the same
source data, and correctly invalidated when it changes.

## 9. Derived columns

`deriveColumns(analysis, rowCount)` returns columns appended to the dataset:

| Column | Type | Value |
|---|---|---|
| `Issue category` | categorical | bucket of the row's highest-severity fragment |
| `Issue categories` | categorical | all distinct buckets for the row, `;`-joined |
| `Theme` | categorical | theme of the row's highest-severity fragment |
| `Issue count` | numeric | fragments extracted from that row |
| `Severity` | numeric | highest fragment severity, 0–100 |

Rows with no issues get `No issue raised`, `0`, `0`.

These are ordinary dataset columns. The chart canvas, tile editor, filter
bar, cross-filtering, saved dashboards and PNG export all consume them with
no change whatsoever. This is the intended payoff of the design: the
analysis adds data, not a parallel charting system.

## 10. Multi-select splitting

Independent of the model, and worth having on its own: the challenges
column is a semicolon-joined multi-select and is currently unchartable —
every distinct combination reads as its own category, so 42 responses
produce ~35 meaningless bars.

Add **split multi-value column** as a clean operation in the existing
`clean/` layer: detect a text column whose values contain a consistent
separator and whose split parts repeat across rows, and offer to treat it as
multi-value. Counting then reports per-option totals.

This delivers a correct "top challenges, ranked" chart with no AI involved,
and gives §6.7 its breadth signal.

## 11. Worker protocol

A dedicated worker. The existing studio worker holds the parsed grid and
must stay responsive to re-clean messages; it should not also own a 23MB
model and multi-second inference sessions.

**In:**

- `{ type: 'analyze', texts, respondentBreadth, buckets, settings }`
- `{ type: 'rescore', buckets, settings }` — cached fragment vectors, no
  re-embedding of fragments
- `{ type: 'recluster', settings }` — cached vectors, clustering only

**Out:**

- `{ type: 'progress', stage, pct }` — `Loading the model` 0–40,
  `Reading responses` 40–45, `Understanding responses` 45–85 (per batch),
  `Grouping` 85–95, `Ranking` 95–100
- `{ type: 'analyzed', raw }`
- `{ type: 'error', message }`

Progress is reported before each stage starts, matching the studio worker's
existing convention. Model loading dominates the first run and must show its
own progress or the tab looks hung.

## 12. Model hosting and build

`public/models/Xenova/all-MiniLM-L6-v2/` — `config.json`,
`tokenizer.json`, `tokenizer_config.json`, `onnx/model_quantized.onnx`.
`public/ort/` — the ONNX runtime WASM binaries.

Both directories are **gitignored** and populated by a pinned prebuild
script, `scripts/fetch-model.mjs`, run from `prebuild` and available as
`npm run fetch:model` for local development. The script verifies a
recorded SHA-256 for every file and fails the build on mismatch.

Rationale, and its cost: this keeps ~25MB out of every clone while still
serving the model from this application's own origin at runtime, which is
what §2 requires. The precedent is the `xlsx` dependency, already pinned to
a CDN tarball with an integrity hash for comparable reasons. The cost is
that a production deploy needs the upstream host reachable at build time; a
release cut while it is down fails to build. If that becomes unacceptable,
the fallback is committing the files, and nothing else in the design
changes.

The dev server must serve `public/` before the tab is opened; a missing
model produces the §14 error, not a crash.

## 13. Persistence

`store/db.js` gains an `analyses` object store keyed by dataset id:

```js
{ datasetId, columnName, buckets, overrides, settings, updatedAt, vectors? }
```

`vectors` is cached only when fragment count <= 2,000 (about 3MB as
Float32). Above that, reopening re-embeds. Writes go through the existing
quota handling, so a full-storage failure surfaces in the same dialog Data
Studio already shows rather than failing silently.

Local to this browser. Never synced, never uploaded.

## 14. Errors and edge cases

| Situation | Behaviour |
|---|---|
| Model files missing or fail to load | Error banner: the analysis model could not load, with retry. The rest of Data Studio is unaffected. |
| No qualifying text column | The tab and its entry button are not rendered. |
| Fewer than 5 usable fragments | Bucketing and the issue table run; clustering and ranking are skipped with an explanation. Two sentences do not have themes. |
| Every fragment lands in Unsorted | Explicit prompt to lower the confidence threshold or edit bucket descriptions — not an empty screen. |
| More than 5,000 fragments | Clustering declines with a message; bucketing still runs. |
| Worker throws | Existing `ErrorBanner` pattern. |
| Overrides reference vanished fragments | Dropped silently on load. |

## 15. Testing

Vitest, matching the existing feature's practice — the pure layers carry the
weight, because they are the parts that can be wrong without looking wrong.

- `splitIssues` — real fixtures from the driving survey: the bracketed-label
  response, the seven-challenge response, `no issue from IT`, a
  single-sentence response, an empty cell.
- `boilerplate` — `no issue from IT` is dropped; `No proper system exists
  for tracking approvals` is kept. This pair is the whole point of the rule
  and gets its own test.
- `similarity` — with fixed synthetic vectors: correct best match, correct
  fall-through to `Unsorted` at the threshold boundary.
- `cluster` — known point sets with a known correct grouping; threshold
  boundary behaviour; the singleton rule.
- `labelCluster` — a term common to every fragment must not become a theme
  name; a term unique to the cluster must.
- `severity` — monotonic in lexicon matches; bounded to 0–1.
- `rankIssues` — five respondents with one mild fragment each outrank one
  respondent with five severe fragments. This encodes the §6.8 decision.
- `overrides` — a retag survives a re-score; a merge is idempotent; a stale
  fragment reference is dropped.
- `deriveColumns` — a no-issue row; a multi-category row; column types match
  what the chart engine expects.
- `detectTextColumns` — an identifier column is rejected; the survey's
  description column is selected.
- The multi-value clean op — round-trips through `applyCleanPlan`.

`embed.js` is mocked throughout. No test loads the model.

## 16. Performance budget

Measured on the driving dataset, 42 responses / ~130 fragments:

| Operation | Budget |
|---|---|
| Cold model load (WASM) | <= 10s, with visible progress |
| Warm load (browser cache) | <= 2s |
| Embedding 130 fragments | <= 5s |
| Editing a bucket description, re-scored | <= 400ms |
| Changing the confidence threshold | <= 100ms (no embedding) |
| Changing clustering granularity | <= 300ms |

The last three are what make the tab feel like a tool rather than a batch
job, and they are achievable only because fragment vectors are computed once
and cached. Any change that re-embeds fragments on a settings edit is a
regression against this table.

## 17. Later phases

1. **OneDrive workbook link.** Paste a link to the Forms response workbook;
   read it over Graph with the existing sign-in; refresh as responses come
   in. Needs its own design covering permissions and refresh semantics.
2. **Comparison across surveys.** Same buckets, two datasets, what moved.
3. **Hosted-model analysis.** Written summaries and stated reasoning, if the
   §2 constraint is ever revisited deliberately.
