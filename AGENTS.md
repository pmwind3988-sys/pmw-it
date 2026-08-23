# PROJECT KNOWLEDGE BASE

**Generated:** 2026-05-04
**Updated:** 2026-08-21
**Project:** PMW IT Service Portal (formerly "IT Onboarding Portal")

## OVERVIEW
React 19 + Vite 8 SPA with Azure AD MSAL authentication and SurveyJS forms. Deployed on Vercel.

The UI is the SI CMMS shell — a branded nav column, a sticky bar, a dashboard of
stat cards over a canvas — minus everything maintenance-specific (work orders,
machines, priorities, SLA, roles). The sign-in screen is the PMW OSHE portal's
split poster/card layout, with an idle animation of its own.

## STRUCTURE
```
pmw-it/
├── src/
│   ├── components/
│   │   ├── AppShell.jsx      # nav column + sticky bar + main; the one auth gate
│   │   ├── IdleAnimation.jsx # the sign-in screen's chip/packet animation
│   │   ├── Logo.jsx          # PMW mark (src/assets/logo-*.png)
│   │   ├── SessionDialog.jsx # what a timed-out session looks like being fixed
│   │   ├── SignInTransition.jsx # post-sign-in veil, fades into the dashboard
│   │   ├── SignatureDialog.jsx
│   │   └── ui/               # Icons, Button, Surfaces, StatCard, Badges
│   ├── features/
│   │   ├── datastudio/       # ingest/ profile/ clean/ engine/ canvas/
│   │   │                     # intent/ suggest/ store/ text/ time/ worker/
│   │   └── devices/          # parse/ derive/ sharepoint/ stats/ ui/
│   ├── hooks/
│   │   ├── useRequests.js    # the one SharePoint read + row helpers + token
│   │   └── useSession.js     # session context + phases (no component here)
│   ├── pages/                # Homepage, LoginPage, DashboardPage, ListPage,
│   │                         # FormPage, AssetChecklistPage, DevicesPage
│   ├── styles/
│   │   ├── shell.css         # tokens, brand surface, shell, UI, dashboard
│   │   ├── auth.css          # sign-in layout + the idle animation
│   │   ├── devices.css       # the device list section
│   │   └── datastudio.css    # Data Studio, incl. the chart series palette
│   ├── context/              # ThemeContext (dark/light), SessionContext (auto
│   │                         # re-sign-in + its dialog and entrance animation)
│   ├── services/             # sharePointService.js
│   ├── utils/                # timeout.js, initials.js, authErrors.js,
│   │                         # sessionKeys.js
│   ├── App.jsx               # Router setup
│   ├── main.jsx              # MSAL bootstrap + providers + stylesheet order
│   └── authConfig.js         # Azure AD + SharePoint scopes
├── public/                   # Static assets
├── vite.config.js
├── eslint.config.js
└── package.json
```

## ROUTES
| Path | Screen |
|------|--------|
| `/` | Redirects to `/dashboard` or `/login` |
| `/login` | Sign-in poster + card |
| `/dashboard` | Stat cards, charts, latest requests |
| `/requests` | The records table; filters live in the query string |
| `/list` | Legacy alias, redirects to `/requests` (keeps the query) |
| `/it-boarding-form` | SurveyJS request form (`?edit=<id>` opens a record) |
| `/asset-checklist` | Handover checklist (IN / OUT / individual) |
| `/devices` | Device list: fleet dashboard, register and scan-report import (`?view=`) |
| `/data-studio` | Drop a spreadsheet and land on a dashboard. It reads the file name for the subject, parks the form's bookkeeping columns, charts the rest and reads the written answers. Every chart drills down to the rows behind it and to one record in full. Lazy route. |

## WHERE TO LOOK
| Task | Location |
|------|----------|
| Auth logic | `src/main.jsx` (MSAL init), `src/authConfig.js` |
| Auth gate for a page | `src/components/AppShell.jsx` — pages do not gate themselves |
| Timed-out session / auto sign-in | `src/context/SessionContext.jsx` |
| What counts as a dead session | `isInteractionRequired` in `src/utils/authErrors.js` |
| Any SharePoint token | `useSharePointToken()` in `src/hooks/useRequests.js` |
| Routes | `src/App.jsx` |
| Nav items | `NAV_ITEMS` in `src/components/AppShell.jsx` |
| Design tokens / layout | `src/styles/shell.css` |
| Sign-in screen | `src/pages/LoginPage.jsx`, `src/styles/auth.css` |
| SharePoint reads | `src/hooks/useRequests.js` |
| Device report parsing | `src/features/devices/parse/` |
| Device derived fields and risk | `src/features/devices/derive/` |
| Drives that belong to IT, not to the machine | `src/features/devices/derive/itMedia.js` |
| Which values on a device page read red or green | `src/features/devices/fieldTone.js` |
| Device SharePoint schema | `src/features/devices/sharepoint/deviceSchema.js` |
| Device SharePoint list views | `src/features/devices/sharepoint/deviceViews.js` |
| Editing or removing one device row | `src/features/devices/sharepoint/updateDevice.js` |
| Device fleet statistics | `src/features/devices/stats/deviceStats.js` |
| Bar and column charts | `src/components/ui/Charts.jsx` (shared by both dashboards) |
| SharePoint writes | `src/services/sharePointService.js` |
| Theme | `src/context/ThemeContext.jsx`; toggle lives in the shell's bar |
| Spreadsheet parsing / header detection | `src/features/datastudio/ingest/` |
| Column type inference and stats | `src/features/datastudio/profile/` |
| The cleaning ops, proposals and apply | `src/features/datastudio/clean/` |
| Columnar store, filter masks, aggregation | `src/features/datastudio/engine/` |
| Chart tiles, theme, grid, tile editor | `src/features/datastudio/canvas/` |
| Starter chart suggestions | `src/features/datastudio/suggest/suggestCharts.js` |
| IndexedDB, exports, the saved library | `src/features/datastudio/store/` |
| Malaysia time parsing and formatting | `src/features/datastudio/time/malaysiaTime.js` |
| Data Studio state and worker lifetime | `src/features/datastudio/DataStudioContext.jsx` |
| Multi-select column detection and encoding | `profile/inferType.js` (`detectMultiSeparator`), `engine/dataset.js` (`encodeMulti`) |
| Splitting a written answer into issues | `src/features/datastudio/text/splitIssues.js` |
| The analysis categories and their descriptions | `src/features/datastudio/text/buckets.js` |
| The model, and where it is served from | `src/features/datastudio/text/embed.js`, `scripts/fetch-model.mjs` |
| Text analysis pipeline and its worker | `text/analysis.js`, `worker/text.worker.js` |
| The user's corrections to the analysis | `src/features/datastudio/text/overrides.js` |
| What the file NAME says the sheet is about | `src/features/datastudio/intent/fileIntent.js` |
| Which columns are form bookkeeping | `src/features/datastudio/intent/adminColumns.js` |
| The one decision taken per import | `src/features/datastudio/intent/planAutopilot.js` |
| The card that discloses that decision | `src/features/datastudio/intent/AutoBrief.jsx` |
| Decoding one stored cell into readable text | `src/features/datastudio/engine/formatCell.js` |
| The rows behind the charts, and one record in full | `engine/rows.js`, `canvas/RecordsPanel.jsx` |
| The dashboard built out of the text analysis | `src/features/datastudio/text/analysisTiles.js` |

## CONVENTIONS

**Page composition**: a screen renders `<AppShell title subtitle actions>` and
its own body. The bar, the nav, the theme toggle, sign-out and the sign-in gate
all belong to the shell — do not re-add per-page copies of them.

**Stylesheet order** (`src/main.jsx`): `index.css` → `App.css` → `styles/shell.css`
→ `styles/auth.css` → `styles/devices.css` → `styles/datastudio.css`.
shell.css re-points the older `--bg` / `--surface` /
`--border` / `--text-*` tokens at the new palette, so it must load last.
`--accent` is deliberately left alone: it fills `.ms-button`, whose text colour
is `--bg`.

**Dashboard ↔ records**: every dashboard figure links into `/requests` with a
query string (`?type=`, `?entity=`, `?department=`, `?range=`, `?equipment=`).
Both screens read the same `useRequests()` fetch, so a card and the list it opens
cannot disagree.

**Navigation**: use `window.location.replace()` instead of React Router
`navigate()` *inside `useEffect`*. WHY: navigate causes a state update →
re-render → effect runs again → infinite loop. In event handlers `navigate()` is
correct and is what the shell and pages use.

**MSAL redirect handling**: always await `handleRedirectPromise()` before
rendering. Silent `no_token_request_cache_error` is normal on fresh load.

**Session guard** (`src/context/SessionContext.jsx`): a timed-out session signs
itself back in. Two rules keep it from disturbing anyone who is still signed in:

1. Nothing is on a clock — no idle timer, no expiry watcher. A recovery starts
   only on proof (an `InteractionRequiredAuthError`, or MSAL's silent renewal
   coming back `timed_out`), or when the account has vanished from the cache on
   a page that needs one. Network failures and SharePoint errors stay errors.
2. A recovery tries `ssoSilent` first. A live Azure AD session comes back in
   about a second and nobody leaves the page; only a second refusal escalates to
   `loginRedirect`, guarded by a one-shot `sessionStorage` flag so a sign-in that
   keeps failing cannot bounce the browser in a loop.

`/login` is exempt, and sign-out clears the stored login hint *before* the
redirect — leaving it would let the guard read a deliberate sign-out as a
timeout and undo it. Sign out through `useSession().signOut()` for that reason.

**`src/features/<name>/`** is where a section with more than a handful of modules
lives — `datastudio/` and `devices/` both follow it. Layering inside a feature:
`parse/` knows nothing about the domain, `derive/` knows nothing about SharePoint,
`sharepoint/` imports no React. Each layer is testable without the one above it.
Data Studio layers the same way — `ingest/` → `profile/` → `clean/` → `engine/`,
each a set of pure functions over plain data, with `canvas/` the only part that
touches React. That is why `engine/aggregate.js` carries more tests than the
whole canvas does: it is the part that can be wrong without looking wrong.

**An import decides everything at once, in `intent/planAutopilot.js`.** Dropping
a sheet no longer parks the user on a profile screen: the file name is read for
the subject, the form's bookkeeping columns are set to role `ignored`, the
starter charts are ranked with the title's keywords as a nudge, and a
pain-point or feedback survey has its written answers read in the background.
The provider carries the plan out and takes no decisions of its own, so the
whole of the behaviour is testable without React — see `intent/*.test.js`.

Everything the plan decided is disclosed by `intent/AutoBrief.jsx` at the top of
the canvas, and every part of it is reversible from there in one click. That
card is not decoration: it is the reason an autopilot that reads header names
with a lexicon is safe to ship at all.

**The scan is run off a USB disk, and that disk is in the reports.**
`derive/itMedia.js` lists it by exact model (`WDC WD10 JPVX-60JC3T1`) and
`deriveStorage` keeps it out of the total, the drive count, the disk type and
therefore the risk score -- counted, it adds 932 GB the machine does not have,
turns an all-SSD laptop into "Mixed", and charges 10 points for a spinning disk
that is not in it. The drive is still reported, under `ignoredDrives`, so the
device page can say why the numbers are lower than the report. Match on the
exact model, never on "WDC" or on the size: a real 1 TB disk inside a desktop
has to keep counting. Rows already in SharePoint keep their old figures until
the reports are imported again.

**Every processor lands on one scale: the Intel generation it is contemporary
with** (`cpuGenerationRank`). A Ryzen badge does not carry its age -- a Ryzen 5
7530U and a Ryzen 5 7640U are both "7000" and are three years of architecture
apart -- so `deriveCpu` reads the architecture out of the model number and
places the Zen it finds against Intel: Zen 3 with 11th gen, Zen 4 with 13th,
Core Ultra 1 and 2 continuing the count at 14 and 15. Two rules do the work, and
both matter: on 2022-and-later MOBILE parts the THIRD digit is the architecture
(7*3*30U is Zen 3, 7*8*40U is Zen 4), and mobile parts before that run one
series behind their desktop namesakes (a 3500U is Zen+ where a desktop 3600 is
Zen 2). The digit rule is mobile-only -- applied to a desktop 7950X it would
read the 5 and call a Zen 4 chip Zen 5.

**A device page colours its values, and can be told not to.** `fieldTone.js`
holds the judgement -- red for what needs attention, green for what does not --
and only for fields with a settled right answer. A RAM discrepancy, a static IP
and a free memory slot stay the colour of the rest of the page on purpose:
colour that appears everywhere says nothing. The toggle sits in the page header
and is remembered per browser under `deviceValueTones`.

**Device report parsing keys off a known-label whitelist** (`parse/labels.js`). A
generic `^Word:` split reads `Total Slots: 2 | Used Slots: 2` and
`Y: | \\server\PMW\IT` as field names and moves those values out of the blocks they
belong to. An unknown label owns the lines beneath it, so a field the scan script
adds later surfaces in review rather than contaminating its predecessor.

**A hand-edited device field outranks the scan file.** The register lets the
three DERIVED fields be retyped (owner, department, device type) and records
which ones in `ManualFields`. `applyManualOverrides` in `syncDevices.js` then
holds those back on re-import — from the diff AND from the body, or updating
anything else would overwrite them as a side effect. Clearing a field is how
somebody hands it back to the scan file; without that it would stay frozen
against every future import for good.

**`Total RAM` in a scan report is usable RAM, not installed RAM.** Windows subtracts
the integrated GPU's reserved share, so a 16 GB laptop reports 15 GB and an 8 GB one
reports 7 GB. Sum `RAM Slot Info` for the real figure; ranking on the reported one
puts a 16 GB machine below an 8 GB machine.

**SharePoint column creation**, verified against the tenant on 2026-08-21 while
provisioning the device lists. All three of these fail silently or confusingly
if you get them wrong:

1. **Create each field as its concrete type**, not the base `SP.Field`.
   `SP.Field` does not declare `Choices`, so a choice column sent that way
   fails with *"The property 'Choices' does not exist on type 'SP.Field'"*.
   Use `SP.FieldChoice`, `SP.FieldNumber`, `SP.FieldDateTime`,
   `SP.FieldMultiLineText`.
2. **The internal name comes from the `Title` a field is CREATED with.**
   `StaticName` in the creation body does not control it. Create
   `Title: 'Device Type'` and the field is addressable only as
   `Device_x0020_Type`; every item write of `DeviceType` then fails with
   *"The property 'DeviceType' does not exist on type 'SP.Data...ListItem'"*.
   Create under the internal name, then MERGE the display `Title` on
   afterwards. This is what produced the hand-encoded `Calling_x0020_Name`
   in `sharePointService.js`.
3. **Read `InternalName`, never `StaticName`**, when checking which columns
   already exist. The two can disagree, and a column where they disagree is
   precisely the broken one.
4. **REST-created columns join no view.** A freshly provisioned list shows
   nothing but its Title until view fields are set explicitly, which is what
   `deviceViews.js` is for.
5. **`ViewQuery` is only honoured in the creation body.** A default view is
   never created, so a filter or sort declared on one has to be MERGEd on
   afterwards or it is silently dropped. Address the built-in view through
   `/defaultView`, not `getByTitle('All Items')` — that title is English-only.

`ensureAssetColumns` in `src/services/sharePointService.js` still has bug 1:
it sends `Choices` with `__metadata: SP.Field`. Its lists predate the bug, so
nothing is broken today, but the same code on a fresh site would fail.

**SharePoint scopes**: use ROOT domain only, never site paths.
- ✅ `https://pmwgroupcom.sharepoint.com/AllSites.Write`
- ❌ `https://pmwgroupcom.sharepoint.com/sites/IThelpdesk/AllSites.Write`

**Icons**: `src/components/ui/Icons.jsx`, transcribed on the 24px stroke grid.
No icon package is installed — add a glyph there rather than a dependency.

## ANTI-PATTERNS (THIS PROJECT)
- Don't use `navigate()` in useEffect — causes infinite loops
- Don't patch a Data Studio profile's `columns` without re-running
  `retopProfile`. `topMeasure` and `primaryTemporal` are derived from the
  roles, and a profile patched by hand goes on naming a column that has since
  been ignored — which is how the starter dashboard opened with
  "Time taken over Timestamp" on a survey where the autopilot had parked both.
- Don't let a tile with no x axis fall through `makeXResolver`'s missing-column
  guard. "No x asked for" (a KPI) and "the x column is gone" (a stale saved
  dashboard) are different answers: the first aggregates the whole dataset as
  one group, the second draws nothing. Conflating them made every starter
  dashboard's KPI row read a confident `0` over a full sheet.
- Don't have the autopilot DELETE a column it thinks is bookkeeping. It sets
  role `ignored` and lists what it did in `AutoBrief`, because the lexicon in
  `adminColumns.js` reads header names and will eventually be wrong about
  somebody's sheet. Being wrong must cost a click, not a re-import.
- Don't start the text analysis on every import that happens to have a prose
  column. The first run pulls a 23MB model, so `planAutopilot` only sets
  `autoAnalyse` when the file TITLE says the writing is the data (a pain-point
  or feedback survey). Every other sheet gets a button.
- Don't add SharePoint scopes to loginRequest — separate request required
- Don't call `acquireTokenSilent` / `acquireTokenPopup` from a page. Use
  `useSharePointToken()`, which routes through the session guard. The popup
  fallback this replaced is where the "it just stopped loading" reports came
  from: a popup opened from an expired timer is not a user gesture, so browsers
  block it and the page waits forever on a window nobody was shown.
- Don't widen `isInteractionRequired` to catch more errors. Signing someone back
  in who never lost their session is worse than the error they were going to
  see; ambiguous codes belong on the "not a timeout" side.
- Don't use `useNavigate` for auth redirects — use window.location
- Don't add `assetsInclude: ['**/*.html']` to vite.config.js. It matches
  index.html itself, so Vite stops treating it as the HTML entry and emits it as
  an asset — `npm run build` then produces a dist/index.html containing
  `export default "/assets/index-….html"` and no bundle at all.
- Don't create a SharePoint DateTime column with `DisplayFormat: 0` when the time
  matters — that is DateOnly and silently discards it. Device columns use `1`,
  confirmed round-tripping a real instant back out of the list.
- Don't give `.bar-fill` anything but `display: block` in `shell.css`. It is a
  `<span>` inside `.bar-track`, which is a plain block, so it is not blockified
  the way a flex or grid child would be — left inline it ignores width and
  height and every dashboard bar paints as an empty track.
- Don't create a SharePoint Note column without `RichText: false`; a rich-text Note
  wraps stored values in `<div>` markup and will not round-trip.
- Don't send `Choices` on a base `SP.Field`. A property exists only on the type
  that declares it, and the tenant answers "The property 'Choices' does not exist
  on type 'SP.Field'". Choice columns go out as `SP.FieldChoice`.
- Don't create a SharePoint column under its display name. SharePoint derives the
  internal name from the `Title` a field is created with — `StaticName` in the
  creation body does not control it — so "Device Type" becomes
  `Device_x0020_Type` and every item write then fails with "The property
  'DeviceType' does not exist". Create the column under its internal name and
  MERGE the display name onto it afterwards, as `provisionLists.js` does. The
  hand-encoded `Calling_x0020_Name` in `sharePointService.js` is the same trap,
  paid the other way.
- Don't add `hour12` beside `hourCycle` in `malaysiaTime.js` — per the Intl spec an
  explicit `hour12` nullifies `hourCycle` entirely. The 24-hour path pins `h23`, the
  AM/PM path pins `h12`, and neither passes `hour12`.
- Don't "fix" the `xlsx` dependency back to a registry version. It is pinned to a
  tarball URL on `https://cdn.sheetjs.com`, and that is deliberate: the npm registry
  copy is frozen at 0.18.5 by a known registry bug, and that version carries two
  open high-severity advisories (prototype pollution GHSA-4r6h-8v6p-xvw6, ReDoS
  GHSA-5pgg-2g8v-p4x9) with no registry fix. `npm install xlsx`, or accepting a
  tooling suggestion to "resolve" the URL, silently downgrades and brings both back
  — and `npm audit` will then say nothing, because the CDN version is not in the
  registry's advisory graph at all. Upgrade by changing the version in the URL. Note
  the install therefore reaches cdn.sheetjs.com at build time; the lockfile pins an
  integrity hash so the contents are verified, but if that dependency is ever
  unacceptable, SheetJS documents vendoring the tarball into the repo instead.
- Don't import the `echarts` umbrella. It pulls in every chart type and both
  renderers — about a megabyte, most of it for charts this app never draws — and
  it does so silently, because the code still works. Import the default export
  of `src/features/datastudio/canvas/echartsCore.js`, which registers exactly
  the pieces used.
- Don't shift a date-only column when the UTC toggle is on. Adding eight hours to
  a value that has no time of day moves it to the wrong day, and nothing on
  screen shows that it happened. `castType` keys this off `dateOnly`, and the
  control in the clean review is rendered DISABLED rather than hidden so someone
  who knows their export is UTC is not left hunting for a setting.
- Don't replace the grid in state without sending it to the worker. The
  worker KEEPS the parsed grid so a re-clean costs one small plan message
  however large the sheet is -- so a main thread that rebuilds the grid
  (which is exactly what adding the text analysis does) leaves the worker
  cleaning the PREVIOUS sheet. Nothing errors: the analysis dashboard
  simply renders six tiles all reporting `Column "Severity" is not in
  this dataset`. `worker/gridSync.js` decides when to carry it, by
  identity, and the provider keeps `workerGridRef` in step.
- Don't decode a cell by checking `column.dictionary` before
  `type === 'multi'`. A multi column carries a dictionary too, so the
  dictionary branch reads its FLAT option array as one code per row and
  prints a real label from the wrong row -- a confident wrong answer, in a
  CSV somebody has already emailed on. `engine/formatCell.js` is the one
  decoder; the row table, the record card and the CSV export all read it,
  so there is one place to get this right rather than three.
- Don't give the records panel its own query. It reads `maskFor` with no
  tile id, which is the same mask the tiles read, so the list cannot
  disagree with the chart above it. A table that disagrees with its chart
  is worse than no table.
- Don't reset the records panel's page from an effect. `npm run lint`
  fails it (`react-hooks/set-state-in-effect`), and it renders one frame
  of the empty page first. The page offset is CLAMPED during render
  instead.
- Don't let Escape inside the record card reach the window. The canvas
  clears the cross-filter on Escape, so closing the card would also throw
  away the selection the user drilled in from and return them to an
  unfiltered dashboard. The card calls `stopImmediatePropagation` on the
  native event.
- Don't build the analysis dashboard by running `suggestCharts` over the
  derived columns. The suggester ranks by SHAPE because on an unknown
  sheet shape is the only evidence; here every column's meaning is known,
  and guessing would sum severity instead of averaging it -- ranking a
  category twenty people mentioned mildly above one three people are
  furious about. `text/analysisTiles.js` names the six tiles outright.
- Don't filter the tile that originated a cross-filter selection. Click "HR" on a
  department bar chart and the other tiles narrow to HR — but that chart keeps
  all its bars, or it collapses to the one bar just clicked and deletes the
  context needed to click anything else. `maskFor` in `engine/filterMask.js`
  implements it, its cache keys on whether the selection *applies* rather than on
  the tile id, and the source tile dims its unselected marks as presentation only.
- Don't write an invisible character as a literal in source. A U+00A0 or a
  zero-width space does not survive being retyped or pasted, so a test that
  claims to cover it silently becomes a no-op, and a diff cannot show the
  difference. Use ` `, `​`, `﻿`. This is not hypothetical here:
  one such swap voided a whole test in `inferType.test.js` and was only caught
  by byte-diffing the file.
- Don't let the analysis model come from anywhere but this app's own origin.
  `embed.js` sets `allowLocalModels = true`, `allowRemoteModels = false` and
  `localModelPath`. The tab tells the user nothing leaves their machine;
  those three lines are what make it true, and the first is a trap —
  `allowLocalModels` defaults to FALSE in the browser, so switching remote
  models off alone disables both and the pipeline refuses to start.
  `embed.contract.test.js` pins all of it by reading the file as text.
- Don't set `wasmPaths`. The runtime reaches its `.wasm` through
  `new URL(…, import.meta.url)`, which the bundler already rewrites to a
  hashed asset on this origin — so the default keeps the promise, and an
  override made the build ship the same 22.5MB file twice. If it ever has to
  be set, name the ONE file rather than a directory prefix: a prefix makes
  the runtime treat its loader as external too and fetch it with a dynamic
  `import()`, which Vite's dev server refuses to do for anything in `public/`,
  so the prefix form works in a build and fails in dev. The file wanted is
  `ort-wasm-simd-threaded.asyncify.wasm` — transformers.js imports
  `onnxruntime-web/webgpu` and that is what its bundled loader names. Grep
  the bundle for `ort-wasm` if it changes; guessing costs an afternoon.
- Don't point runtime code at a path under `public/` that only exists because
  somebody copied it there. `public/models` is gitignored and recreated by
  `scripts/fetch-model.mjs`; anything else gitignored is not recreated by
  anything. A missing file does not 404 — the SPA rewrite answers with
  `index.html`, so the browser reports
  "expected magic word 00 61 73 6d, found 3c 21 64 6f", which is `<!do`.
  That cost a live bug; `embed.contract.test.js` is the guard.
- Don't re-embed on a settings change. The threshold and granularity controls
  re-score against cached fragment vectors, and bucket vectors are cached by
  prompt TEXT so renaming a bucket is free. Measured on the real survey:
  2143ms before the cache, 172ms after, against a 300ms budget for a control
  the user drags. Nothing about a regression here looks wrong in the code.
- Don't cluster by re-summing member distances on every merge. That is
  O(n^3) and took 3–10 seconds on the real survey's 134 fragments.
  `cluster.js` builds the distance matrix once and updates it with the
  Lance-Williams rule; `cluster.test.js` pins the cost.
- Don't score the priority list on how many issues were raised. It counts
  distinct PEOPLE, scaled by severity: one person writing five furious
  sentences must not outrank five people writing one each. `rankIssues.js`
  carries the test that pins it.
- Don't drop a fragment just because it starts with "no". "no issue from IT"
  is a non-answer; "No proper system exists for tracking approvals" is a
  report. `boilerplate.js` requires BOTH a leading non-answer word and a
  short body, and the pair has its own test.
- Don't gate free-text detection on the profiler's `text` verdict. The
  profiler calls any column with 50 or fewer distinct values categorical, so
  on a 42-response survey the written-answer column is never `text` and the
  whole feature would be invisible on exactly the files it was built for.
  `detectTextColumns.js` tests shape — length, fill, uniqueness — instead.
- Don't put `multi` handling after the `column.dictionary` branch in
  `aggregate.js` or `filterMask.js`. A multi column carries a dictionary too,
  so the dictionary branch catches it first and every option collapses into
  one meaningless category.
- Don't derive `dataset.rowCount` from the first column's `values.length`. A
  multi column's `values` is the flat option array and is longer than the
  grid, so every mask allocated from it would be the wrong size.
- Don't export a helper next to a component from the same file — it drops the
  file out of Fast Refresh (and eslint fails the build). `initialsOf` lives in
  `src/utils/initials.js` for exactly this reason.

## COMMANDS
```bash
npm run dev      # Start dev server on port 5173
npm run build    # Build for production (outputs to dist/)
npm run lint     # Run ESLint
npm run preview  # Preview production build
```

## NOTES
- Vite port 5173 is for local dev; Vercel ignores this
- MSAL handles Azure AD login flow + token caching
- SurveyJS drives `/it-boarding-form` and `/asset-checklist`
- `npm run lint` still reports pre-existing errors in FormPage,
  AssetChecklistPage, SignatureDialog and ThemeContext (unused imports, SurveyJS
  model mutation inside hooks). They predate the shell work and are untouched.
